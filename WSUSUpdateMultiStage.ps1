# See README.md for full documentation, usage instructions, and how to modify this script.

# ==============================================================================
#  SECTION 1 — CONFIGURATION
# ==============================================================================

# --- Script identity & persistence ---
$script:ScriptName   = "WSUSUpdateMultiStage"
$script:GitHubRawURL = "https://raw.githubusercontent.com/JaphethWun1/WinUpdatez/main/WSUSUpdateMultiStage.ps1"

# --- Paths ---
$script:LogRoot     = "$env:SystemDrive\WinUpdatez"
$script:LogFile     = "$script:LogRoot\UpdateLog.txt"
$script:LocalScript = "$script:LogRoot\$script:ScriptName.ps1"

# --- Registry ---
$script:StateRegPath = "HKLM:\SOFTWARE\WinUpdatez"
$script:RunOnceKey   = "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"

# --- Update stages ---
# Stages 0 to ($MaxStages - 1) run updates and reboot.
# Stage $MaxStages runs the final update pass and finalizes the machine.
$script:MaxStages = 4

# --- Windows Updates to skip ---
# Add KB numbers here to prevent them from being installed.
$script:ExcludedKBs = @(
    'KB5063878'
)

# --- Chrome ---
$script:ChromeZipUrl  = "https://files.getupdates.me/chrome.zip"
$script:ChromeZipPath = "$env:USERPROFILE\chrome.zip"
$script:ChromeExtPath = "$env:USERPROFILE\chrome"

# URLs opened in Chrome at the end of provisioning (for tech testing)
$script:ChromeTestUrls = @(
    "https://retest.us/laptop-no-keypad",
    "https://testmyscreen.com",
    "https://monkeytype.com"
)

# --- Desktop shortcuts ---
# "Filename on Desktop.lnk" = "URL to download .lnk from"
$script:DesktopShortcuts = @{
    "View Battery Info.lnk"          = "https://getupdates.me/BatteryInfo.lnk"
    "Intel 11th Gen+ Drivers.lnk"    = "https://getupdates.me/Intel_11th_Gen+_Drivers.lnk"
    "Intel 6th-10th Gen Drivers.lnk" = "https://getupdates.me/Intel_6th-10th_Gen_Drivers.lnk"
    "Intel 4th-5th Gen Drivers.lnk"  = "https://getupdates.me/Intel_4th-5th_Gen_Drivers.lnk"
    "Disable HP Absolute.lnk"        = "https://getupdates.me/DisableAbsoluteHP.lnk"
    "MAS - Activate Windows.lnk"     = "https://getupdates.me/MASActivateWindows.lnk"
    "Chrome.lnk"                     = "https://getupdates.me/Chrome.lnk"
}

# --- BatteryInfoView tool ---
$script:BatteryInfoZipUrl = "https://raw.githubusercontent.com/JaphethWun1/WinUpdatez/main/files/batteryinfoview.zip"

# --- Log upload (currently disabled — uncomment $script:LogUploadBaseUrl and Upload-Log call in Invoke-Finalization to enable) ---
# $script:LogUploadBaseUrl = "https://logs.getupdates.me"

# --- Current stage (set at runtime in Section 13) ---
$script:CurrentStage = 0


# ==============================================================================
#  SECTION 2 — ADMIN CHECK
# ==============================================================================

$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin     = ([Security.Principal.WindowsPrincipal] $currentUser).IsInRole(
                   [Security.Principal.WindowsBuiltinRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Not running as Administrator. Relaunching with elevation..."
    Start-Process powershell `
        -ArgumentList "-ExecutionPolicy Bypass -NoProfile -Command `"irm $script:GitHubRawURL | iex`"" `
        -Verb RunAs
    exit
}


# ==============================================================================
#  SECTION 3 — INITIALIZATION
# ==============================================================================

if (-not (Test-Path $script:LogRoot)) {
    New-Item -Path $script:LogRoot -ItemType Directory -Force | Out-Null
}

# Required for Invoke-WebRequest / Invoke-RestMethod to work on older Windows builds
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


# ==============================================================================
#  SECTION 4 — LOGGING & STATE
# ==============================================================================

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp] [Stage $script:CurrentStage] $Message"
    Write-Host $entry
    Add-Content -Path $script:LogFile -Value $entry
}

# Returns the saved stage number from the registry (0 if not set yet)
function Get-Stage {
    if (-not (Test-Path $script:StateRegPath)) { return 0 }
    try {
        return (Get-ItemProperty -Path $script:StateRegPath -Name Stage -ErrorAction Stop).Stage
    } catch {
        return 0
    }
}

# Saves the stage number to the registry so it survives a reboot
function Save-Stage {
    param([int]$StageNumber)
    if (-not (Test-Path $script:StateRegPath)) {
        New-Item -Path $script:StateRegPath -Force | Out-Null
    }
    New-ItemProperty -Path $script:StateRegPath -Name Stage -Value $StageNumber -PropertyType DWord -Force | Out-Null
}

# Removes the stage entry from the registry (called at the very end)
function Remove-Stage {
    if (Test-Path $script:StateRegPath) {
        Remove-Item -Path $script:StateRegPath -Recurse -Force
    }
}


# ==============================================================================
#  SECTION 5 — INTERNET CONNECTIVITY
# ==============================================================================

function Wait-ForInternet {
    param(
        [string]$TestUrl    = "https://www.msftconnecttest.com/connecttest.txt",
        [int]   $MaxRetries = 12,
        [int]   $RetryDelay = 10
    )

    Write-Log "Checking internet connectivity..."

    # Inner function: tries to reach the test URL up to $MaxRetries times.
    # Parameters are passed explicitly to avoid PowerShell closure scoping issues.
    function Test-Connection-Loop {
        param([string]$Url, [int]$Retries, [int]$Delay)
        for ($i = 1; $i -le $Retries; $i++) {
            try {
                $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
                if ($response.StatusCode -eq 200) {
                    Write-Log "Internet confirmed on attempt $i."
                    return $true
                }
            } catch {
                Write-Log "No internet yet (attempt $i of $Retries). Retrying in $Delay seconds..."
                Start-Sleep -Seconds $Delay
            }
        }
        return $false
    }

    if (Test-Connection-Loop -Url $TestUrl -Retries $MaxRetries -Delay $RetryDelay) {
        Write-Log "Internet is available. Continuing."
        return
    }

    # First pass failed — restart all active network adapters and try again
    Write-Log "No internet after $MaxRetries attempts. Restarting network adapters..."
    Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
        try {
            Write-Log "  Disabling adapter: $($_.Name)"
            Disable-NetAdapter -Name $_.Name -Confirm:$false -ErrorAction Stop
            Start-Sleep -Seconds 5
            Write-Log "  Enabling adapter:  $($_.Name)"
            Enable-NetAdapter -Name $_.Name -Confirm:$false -ErrorAction Stop
        } catch {
            Write-Log "  Could not restart adapter '$($_.Name)': $_"
        }
    }

    if (Test-Connection-Loop -Url $TestUrl -Retries $MaxRetries -Delay $RetryDelay) {
        Write-Log "Internet is available after adapter restart. Continuing."
        return
    }

    # Adapters didn't help — prompt the technician to fix it manually
    Write-Log "Internet still unavailable. Manual intervention required."
    while (-not (Test-Connection-Loop -Url $TestUrl -Retries $MaxRetries -Delay $RetryDelay)) {
        Read-Host ">>> Please check the network connection, then press ENTER to retry"
    }
    Write-Log "Internet is available after manual intervention. Continuing."
}


# ==============================================================================
#  SECTION 6 — WINDOWS UPDATE RESET
# ==============================================================================

function Reset-WindowsUpdate {
    Write-Log "Resetting Windows Update components..."

    $updateServices = @(
        "wuauserv",         # Windows Update
        "bits",             # Background Intelligent Transfer Service
        "cryptsvc",         # Cryptographic Services
        "msiserver",        # Windows Installer
        "trustedinstaller"  # Windows Modules Installer
    )

    # Stop services before deleting their cache folders
    foreach ($svc in $updateServices) {
        Get-Service $svc -ErrorAction SilentlyContinue |
            Where-Object { $_.Status -ne "Stopped" } |
            Stop-Service -Force -ErrorAction SilentlyContinue
    }

    # Delete cached update data — Windows Update will rebuild these fresh
    Remove-Item "C:\Windows\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "C:\Windows\System32\catroot2"     -Recurse -Force -ErrorAction SilentlyContinue

    # Reset network components that can interfere with update downloads
    netsh winsock reset        | Out-Null
    netsh winhttp reset proxy  | Out-Null
    bitsadmin /reset /allusers | Out-Null

    # Restart services
    foreach ($svc in $updateServices) {
        Start-Service $svc -ErrorAction SilentlyContinue
    }

    Write-Log "Windows Update reset complete."
}


# ==============================================================================
#  SECTION 7 — INSTALL WINDOWS UPDATES
# ==============================================================================

function Install-Updates {
    Write-Log "---------- Update Pass Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ----------"

    # Previously failed update titles are stored one-per-line so we can skip them
    $failedLogPath    = "$env:USERPROFILE\Desktop\FailedUpdates.txt"
    $previousFailures = if (Test-Path $failedLogPath) { Get-Content $failedLogPath } else { @() }

    # Install PSWindowsUpdate module if not already present
    try {
        if (-not (Get-Module -ListAvailable PSWindowsUpdate)) {
            Write-Log "PSWindowsUpdate not found. Installing from PSGallery..."
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null
            Install-Module PSWindowsUpdate -Force -Confirm:$false -Scope AllUsers -ErrorAction Stop | Out-Null
            Write-Log "PSWindowsUpdate installed."
        }
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue
    } catch {
        Write-Log "ERROR: Could not set up PSWindowsUpdate. Skipping update pass. Details: $_"
        return
    }

    # Scan for available updates
    Write-Log "Scanning for available updates..."
    $allUpdates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false

    # Filter out any KBs listed in $ExcludedKBs
    $filteredUpdates = $allUpdates | Where-Object {
        $kbArticleId = $_.KBArticleIDs
        $title       = $_.Title
        $shouldSkip  = $false

        foreach ($kb in $script:ExcludedKBs) {
            # Match against both the KB article ID and the title for robustness
            if ($kbArticleId -contains $kb -or $title -match [regex]::Escape($kb)) {
                Write-Log "  Skipping excluded KB: $title"
                $shouldSkip = $true
                break
            }
        }
        -not $shouldSkip
    }

    if (-not $filteredUpdates -or $filteredUpdates.Count -eq 0) {
        Write-Log "No updates found (or all were excluded). Nothing to install."
        Write-Log "---------- Update Pass End: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ----------"
        return
    }

    Write-Log "$($filteredUpdates.Count) update(s) to install."

    # Install each update individually for accurate per-update error logging.
    # We pass the KB article ID to Install-WindowsUpdate instead of the title string
    # to avoid partial-match issues with PSWindowsUpdate's -Title parameter.
    foreach ($update in $filteredUpdates) {
        $title = $update.Title
        $kb    = ($update.KBArticleIDs | Select-Object -First 1)

        # Skip if this update title was logged as a failure in a previous stage
        if ($previousFailures -contains $title) {
            Write-Log "  Skipping previously failed update: $title"
            continue
        }

        Write-Log "  Installing: $title"
        try {
            if ($kb) {
                Install-WindowsUpdate -KBArticleID $kb -AcceptAll -IgnoreReboot -Confirm:$false -ErrorAction Stop | Out-Null
            } else {
                # No KB ID available — fall back to title match with escaped string
                Install-WindowsUpdate -Title ([regex]::Escape($title)) -AcceptAll -IgnoreReboot -Confirm:$false -ErrorAction Stop | Out-Null
            }
            Write-Log "  OK: $title"
        } catch {
            $errMsg = $_.Exception.Message
            Write-Log "  FAILED: $title — $errMsg"
            # Store just the title so the -contains check above works correctly next time
            $title | Out-File -FilePath $failedLogPath -Append -Encoding UTF8
        }
    }

    Write-Log "---------- Update Pass End: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ----------"
}


# ==============================================================================
#  SECTION 8 — RUNONCE PERSISTENCE
# ==============================================================================

function Register-RunOnce {
    # Download a local copy of the script so RunOnce doesn't pull from GitHub on every boot
    if (-not (Test-Path $script:LocalScript)) {
        try {
            Write-Log "Saving local script copy to: $script:LocalScript"
            Invoke-RestMethod -Uri $script:GitHubRawURL -OutFile $script:LocalScript -UseBasicParsing -ErrorAction Stop
        } catch {
            Write-Log "WARNING: Could not save script locally. Will pull from GitHub on next boot. Error: $_"
            $script:LocalScript = $null  # $script: scope ensures this is visible outside the function
        }
    }

    # Prefer running from local file; fall back to GitHub if unavailable
    if ($script:LocalScript -and (Test-Path $script:LocalScript)) {
        $runOnceCmd = "cmd.exe /c start powershell.exe -ExecutionPolicy Bypass -NoProfile -File `"$script:LocalScript`""
    } else {
        $runOnceCmd = "powershell.exe -ExecutionPolicy Bypass -NoProfile -Command `"irm $script:GitHubRawURL | iex`""
    }

    New-ItemProperty -Path $script:RunOnceKey -Name $script:ScriptName -Value $runOnceCmd -PropertyType String -Force | Out-Null
    Write-Log "RunOnce registered. Script will resume automatically after reboot."
}

function Unregister-RunOnce {
    Remove-ItemProperty -Path $script:RunOnceKey -Name $script:ScriptName -ErrorAction SilentlyContinue
    Write-Log "RunOnce entry removed."
}


# ==============================================================================
#  SECTION 9 — CHROME
# ==============================================================================

# Checks whether Chrome has already been extracted to disk
function Test-ChromeExtracted {
    return (Test-Path "$script:ChromeExtPath\chrome.exe") -or
           (Test-Path "$script:ChromeExtPath\chrome\chrome.exe")
}

# Returns the path to the Chrome executable, or $null if not found
function Get-ChromeExePath {
    if (Test-Path "$script:ChromeExtPath\chrome.exe")         { return "$script:ChromeExtPath\chrome.exe" }
    if (Test-Path "$script:ChromeExtPath\chrome\chrome.exe")  { return "$script:ChromeExtPath\chrome\chrome.exe" }
    return $null
}

# Kicks off a background job to download and extract Chrome during Stage 0,
# so it is ready by the time finalization runs at the final stage
function Start-ChromeBackgroundDownload {
    if (Test-ChromeExtracted) {
        Write-Log "Chrome already extracted. Background download skipped."
        return
    }

    Write-Log "Starting Chrome download in the background..."

    Start-Job -ScriptBlock {
        param($url, $zipPath, $extractPath, $logPath)

        function Write-BgLog ($msg) {
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [Chrome BG] $msg" | Add-Content -Path $logPath
        }

        try {
            Write-BgLog "Downloading Chrome from $url ..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop

            if (Test-Path $extractPath) {
                Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }

            Write-BgLog "Extracting Chrome to $extractPath ..."
            Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
            Write-BgLog "Chrome is ready at $extractPath."
        } catch {
            Write-BgLog "ERROR: $_"
        }

    } -ArgumentList $script:ChromeZipUrl, $script:ChromeZipPath, $script:ChromeExtPath, $script:LogFile | Out-Null
}

# Foreground Chrome download — used as a fallback at finalization if background job didn't finish
function Install-Chrome {
    if (Test-ChromeExtracted) {
        Write-Log "Chrome already extracted. Skipping foreground download."
        return
    }

    Write-Log "Downloading Chrome (foreground)..."
    try {
        Invoke-WebRequest -Uri $script:ChromeZipUrl -OutFile $script:ChromeZipPath -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop

        if (Test-Path $script:ChromeExtPath) {
            Remove-Item $script:ChromeExtPath -Recurse -Force
        }
        Expand-Archive -Path $script:ChromeZipPath -DestinationPath $script:ChromeExtPath -Force
        Remove-Item $script:ChromeZipPath -Force -ErrorAction SilentlyContinue
        Write-Log "Chrome extracted to $script:ChromeExtPath"
    } catch {
        Write-Log "ERROR: Chrome download failed. $_"
    }
}

# Finds Chrome, downloads it if missing, then launches it with the configured test URLs
function Open-Chrome {
    $exePath = Get-ChromeExePath

    if (-not $exePath) {
        Write-Log "Chrome not found. Attempting download now..."
        Install-Chrome
        $exePath = Get-ChromeExePath
    }

    if ($exePath) {
        try {
            # Note: avoid assigning to $args — it is a PowerShell automatic variable
            $chromeArgs = @("-no-default-browser-check") + $script:ChromeTestUrls
            Start-Process -FilePath $exePath -ArgumentList $chromeArgs
            Write-Log "Chrome launched from: $exePath"
        } catch {
            Write-Log "ERROR: Could not launch Chrome. $_"
        }
    } else {
        Write-Log "ERROR: Chrome executable not found even after download attempt."
    }
}


# ==============================================================================
#  SECTION 10 — FINALIZATION TASKS
# ==============================================================================

function Install-DesktopShortcuts {
    Write-Log "Installing desktop shortcuts..."
    $desktopPath = [Environment]::GetFolderPath('Desktop')

    foreach ($shortcut in $script:DesktopShortcuts.GetEnumerator()) {
        $dest = Join-Path $desktopPath $shortcut.Key
        try {
            Invoke-WebRequest -Uri $shortcut.Value -OutFile $dest -UseBasicParsing -ErrorAction Stop
            Write-Log "  Placed: $($shortcut.Key)"
        } catch {
            Write-Log "  FAILED to place '$($shortcut.Key)': $_"
        }
    }
}

function Open-BatteryInfoView {
    $zipPath    = Join-Path $env:TEMP "batteryinfoview.zip"
    $extractDir = Join-Path $env:TEMP "batteryinfoview"
    $exePath    = Join-Path $extractDir "BatteryInfoView.exe"

    Write-Log "Downloading and launching BatteryInfoView..."
    try {
        Invoke-WebRequest -Uri $script:BatteryInfoZipUrl -OutFile $zipPath -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop

        if (-not (Test-Path $extractDir)) {
            New-Item -ItemType Directory -Path $extractDir | Out-Null
        }
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force

        if (Test-Path $exePath) {
            Start-Process -FilePath $exePath
            Write-Log "BatteryInfoView launched."
        } else {
            Write-Log "ERROR: BatteryInfoView.exe not found after extraction."
        }
    } catch {
        Write-Log "ERROR: BatteryInfoView failed. $_"
    }
}

# Reads the OA3 product key embedded in the machine's firmware and activates Windows with it
function Invoke-WindowsActivation {
    Write-Log "Attempting Windows activation using firmware (OA3) key..."
    try {
        $firmwareKey = (Get-CimInstance -Class SoftwareLicensingService).OA3xOriginalProductKey

        if (-not $firmwareKey) {
            Write-Log "No OA3 key found in firmware. Skipping activation."
            return
        }

        cscript /b C:\Windows\System32\slmgr.vbs /upk | Out-Null
        cscript /b C:\Windows\System32\slmgr.vbs /ipk $firmwareKey | Out-Null
        cscript /b C:\Windows\System32\slmgr.vbs /ato | Out-Null

        $licenseStatus = Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" |
            Where-Object { $_.PartialProductKey } |
            Select-Object -ExpandProperty LicenseStatus

        $statusLabel = switch ($licenseStatus) {
            0 { "Unlicensed" }
            1 { "Licensed" }
            2 { "OOBGrace" }
            3 { "OOTGrace" }
            4 { "NonGenuineGrace" }
            5 { "Not Activated" }
            6 { "ExtendedGrace" }
            default { "Unknown (code: $licenseStatus)" }
        }

        Write-Log "Activation result: $statusLabel"

        if ($licenseStatus -ne 1) {
            Write-Log "NOTE: Windows is not fully activated. Manual activation may be required."
        }
    } catch {
        Write-Log "ERROR: Activation failed. $_"
    }
}

function Copy-LogToDocuments {
    try {
        $dest = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "WSUSUpdateLog.txt"
        Copy-Item -Path $script:LogFile -Destination $dest -Force
        Write-Log "Log copied to: $dest"
    } catch {
        Write-Log "WARNING: Could not copy log to Documents. $_"
    }
}

# Uploads the log to Cloudflare R2, named by serial number + timestamp for easy identification.
# Currently disabled — to re-enable: uncomment the function body below, uncomment the
# Upload-Log call in Invoke-Finalization, and uncomment $script:LogUploadBaseUrl in Section 1.
function Upload-Log {
    # Write-Log "Uploading log to cloud..."
    # try {
    #     $serial    = (Get-CimInstance Win32_BIOS).SerialNumber -replace '[^a-zA-Z0-9\-]', ''
    #     $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    #     $fileName  = if ($serial) { "$serial-$timestamp.txt" } else { "$timestamp.txt" }
    #     $logBytes  = [IO.File]::ReadAllBytes($script:LogFile)
    #
    #     Invoke-RestMethod -Uri "$script:LogUploadBaseUrl/$fileName" -Method Put -Body $logBytes -ContentType "application/octet-stream" | Out-Null
    #     Write-Log "Log uploaded as: $fileName"
    # } catch {
    #     Write-Log "WARNING: Log upload failed (non-critical). $_"
    # }
}


# ==============================================================================
#  SECTION 11 — SYSTEM FINALIZATION
# ==============================================================================

function Invoke-Finalization {
    Write-Log "========== FINALIZATION PHASE =========="

    Invoke-WindowsActivation    # Activate Windows using the OA3 firmware key
    Install-DesktopShortcuts    # Place .lnk shortcuts on the desktop
    Open-BatteryInfoView        # Download and launch battery health tool

    Write-Log "TIP: If Intel GPU drivers are missing, use the matching Intel driver shortcut on the desktop."

    Copy-LogToDocuments         # Save a copy of the log to the Documents folder
    # Upload-Log                # Upload the log to Cloudflare R2 (currently disabled)
    Open-Chrome                 # Launch Chrome with tech test pages

    Write-Log "========== FINALIZATION COMPLETE =========="
}


# ==============================================================================
#  SECTION 12 — SYSPREP / OOBE
# ==============================================================================

function Start-OOBE {
    Write-Log "Launching Sysprep to enter OOBE..."
    $sysprepExe = "$env:SystemRoot\System32\Sysprep\sysprep.exe"

    if (Test-Path $sysprepExe) {
        Start-Process $sysprepExe -ArgumentList "/oobe /reboot" -Wait
    } else {
        Write-Log "ERROR: Sysprep not found at $sysprepExe"
    }
}


# ==============================================================================
#  SECTION 13 — MAIN EXECUTION
# ==============================================================================

# Read the saved stage from the registry (returns 0 on first run)
$script:CurrentStage = Get-Stage
Write-Log "Script started. Stage $script:CurrentStage of $script:MaxStages."

# Stage 0 only: tasks that should run exactly once on the very first execution
if ($script:CurrentStage -eq 0) {

    # Rename the machine to PC-<SerialNumber> for easy identification on the network
    try {
        $serial = (Get-CimInstance Win32_BIOS).SerialNumber
        Rename-Computer -NewName "PC-$serial" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        Write-Log "Computer renamed to: PC-$serial"
    } catch {
        Write-Log "WARNING: Could not rename computer (non-critical). $_"
    }

    # Start Chrome download in the background now so it's ready by the final stage
    Start-ChromeBackgroundDownload
}

# Stages 0 through (MaxStages - 1): run an update pass, then reboot
if ($script:CurrentStage -lt $script:MaxStages) {

    Save-Stage ($script:CurrentStage + 1)   # Advance the stage counter before rebooting
    Register-RunOnce                         # Ensure the script resumes after reboot

    Wait-ForInternet
    Reset-WindowsUpdate
    Install-Updates

    Write-Log "Rebooting in 15 seconds..."
    Start-Sleep -Seconds 15
    Restart-Computer -Force

# Final stage: last update pass, finalize the machine, enter OOBE
} else {

    Write-Log "All update stages complete. Running final update pass..."

    Wait-ForInternet
    Reset-WindowsUpdate
    Install-Updates

    Invoke-Finalization     # Activate Windows, shortcuts, Chrome, battery tool, log upload
    Unregister-RunOnce      # Remove RunOnce — provisioning is complete
    Remove-Stage            # Clean up the registry stage entry

    Write-Log "Provisioning complete. Entering OOBE..."
    Start-OOBE
}