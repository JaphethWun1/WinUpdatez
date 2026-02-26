# Windows Update Provisioning Script

Automates the full Windows Update process across multiple reboots, then finalizes
the machine for handoff — activating Windows, placing desktop shortcuts, running
diagnostic tools, and entering OOBE.

---

## What It Does

1. **Renames the computer** to `PC-<SerialNumber>` on first run
2. **Runs Windows Update** across multiple reboot cycles to ensure all updates are applied
3. **Finalizes the machine** after all updates are done:
   - Activates Windows using the firmware (OA3) embedded product key
   - Places desktop shortcuts (driver links, tools, Chrome)
   - Downloads and launches BatteryInfoView
   - Launches Chrome with test pages open
   - Uploads the log to Cloudflare R2 for record-keeping
4. **Runs Sysprep** to prepare the machine for OOBE (first-user setup)

---

## How the Stages Work

The script uses a stage number stored in the registry to track progress across reboots.

| Stage   | What happens                                          |
|---------|-------------------------------------------------------|
| 0       | First run — rename PC, start Chrome download, run updates, reboot |
| 1       | Run updates, reboot                                   |
| 2       | Run updates, reboot                                   |
| 3       | Run updates, reboot                                   |
| 4       | Final update pass → finalize → Sysprep OOBE           |

Before each reboot, a **RunOnce** registry entry is written so the script
automatically resumes after the machine comes back up. The entry is removed
at the very end.

---

## How to Run It

Right-click PowerShell → **Run as Administrator**, then:

```powershell
irm https://raw.githubusercontent.com/JaphethWun1/WinUpdatez/main/WSUSUpdateMultiStage.ps1 | iex
```

The script will self-elevate if not already running as Administrator.

---

## Files & Locations

| Path | Description |
|------|-------------|
| `C:\WinUpdatez\UpdateLog.txt` | Main log file — written throughout the process |
| `C:\WinUpdatez\WSUSUpdateMultiStage.ps1` | Local copy of this script used by RunOnce |
| `HKLM:\SOFTWARE\WinUpdatez` | Registry key storing the current stage number |
| `HKLM:\...\RunOnce\WSUSUpdateMultiStage` | RunOnce entry (present only between reboots) |
| `%USERPROFILE%\Desktop\FailedUpdates.txt` | Log of any updates that failed to install |
| `Documents\WSUSUpdateLog.txt` | Copy of the log placed in Documents at the end |

---

## Common Modifications

### Add or remove update stages
Open the script and change `$MaxStages` in the Configuration section.
The loop handles everything else automatically — no other changes needed.

```powershell
$MaxStages = 4   # Change to 5 for an extra reboot cycle, 3 for one fewer, etc.
```

### Skip a specific Windows Update (KB)
Add the KB number to `$ExcludedKBs` in the Configuration section.

```powershell
$ExcludedKBs = @(
    'KB5063878',
    'KB1234567'   # Add as many as needed
)
```

### Add or remove desktop shortcuts
Edit the `$DesktopShortcuts` hashtable in the Configuration section.
Each entry is: `"Filename on Desktop.lnk" = "URL to download .lnk from"`

```powershell
$DesktopShortcuts = @{
    "My Tool.lnk"   = "https://example.com/mytool.lnk"
    "Chrome.lnk"    = "https://getupdates.me/Chrome.lnk"
    # Remove a line to remove that shortcut
}
```

### Change the Chrome test URLs
Edit `$ChromeTestUrls` in the Configuration section.

```powershell
$ChromeTestUrls = @(
    "https://retest.us/laptop-no-keypad",
    "https://testmyscreen.com",
    "https://monkeytype.com"
)
```

### Point the script to a different GitHub URL
Change `$GitHubRawURL` in the Configuration section. This is used both for
the initial download and as a fallback if the local script copy is missing.

```powershell
$GitHubRawURL = "https://raw.githubusercontent.com/YourUser/YourRepo/main/WSUSUpdateMultiStage.ps1"
```

---

## Script Structure (for developers)

The script is organized into numbered sections. Each section is self-contained
and clearly labelled in the file.

| Section | Name | Description |
|---------|------|-------------|
| 1 | Configuration | All tunable values — paths, URLs, lists, stage count |
| 2 | Admin Check | Re-launches as Administrator if needed |
| 3 | Initialization | Creates log folder, enforces TLS 1.2 |
| 4 | Logging & State | `Write-Log`, `Get-Stage`, `Save-Stage`, `Remove-Stage` |
| 5 | Internet Check | Waits for connectivity, restarts adapters if needed |
| 6 | WU Reset | Stops services, clears cache, resets winsock/proxy/BITS |
| 7 | Install Updates | PSWindowsUpdate setup, KB filtering, per-update install |
| 8 | RunOnce | `Register-RunOnce` / `Unregister-RunOnce` |
| 9 | Chrome | Background download, foreground fallback, launcher |
| 10 | Finalization Tasks | Shortcuts, BatteryInfoView, activation, log copy/upload |
| 11 | Finalization | `Invoke-Finalization` — calls all Section 10 tasks in order |
| 12 | OOBE | Sysprep `/oobe /reboot` |
| 13 | Main Execution | Entry point — reads stage, runs the appropriate path |

---

## Troubleshooting

**Script doesn't resume after reboot**
Check that the RunOnce entry was written:
```powershell
Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
```
If missing, re-run the script manually — it will read the saved stage and continue.

**Stuck on a specific stage**
Check `C:\WinUpdatez\UpdateLog.txt` for errors. You can also check the current stage:
```powershell
(Get-ItemProperty "HKLM:\SOFTWARE\WinUpdatez").Stage
```
To manually set the stage (e.g. skip back to stage 2):
```powershell
Set-ItemProperty "HKLM:\SOFTWARE\WinUpdatez" -Name Stage -Value 2
```

**Windows not activating**
The script uses the OA3 key embedded in the machine's firmware (the key the OEM put there).
If the machine doesn't have one, activation will be skipped and a note is logged.
Use the `MAS - Activate Windows` desktop shortcut as a fallback.

**Chrome not launching**
Check `C:\WinUpdatez\UpdateLog.txt` for Chrome-related errors.
The script tries a background download at Stage 0 and a foreground fallback at Stage 4.
You can also use the `Chrome.lnk` desktop shortcut.

**An update keeps failing**
Its title will be logged to `%USERPROFILE%\Desktop\FailedUpdates.txt`.
The script automatically skips updates that appear in this file on future stages.
To permanently skip it, add its KB number to `$ExcludedKBs` in the script.
