
# PowerShell version of Scan-UpdatesOffline.bat logic, with update scan
$LogFile = "C:\scripts\ScanReport.txt"
$WsusCab = "C:\scripts\wsusscn2.cab"
$WsusUrl = "http://go.microsoft.com/fwlink/p/?LinkID=74689"

# Fail fast if not running elevated
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    $isAdmin = $false
}
if (-not $isAdmin) {
    Write-Host "This script must be run as Administrator. Please open an elevated PowerShell (Run as Administrator) and re-run the script." -ForegroundColor Yellow
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm') - Script exited: requires Administrator elevation."
    exit 2
}

# --- Progress helpers ---
function Format-TimeSpan {
    param($seconds)
    if ($seconds -is [TimeSpan]) { $ts = $seconds } else { $ts = [TimeSpan]::FromSeconds([double]$seconds) }
    return ('{0:00}:{1:00}:{2:00}' -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds)
}

function Download-File {
    param(
        [string]$Url,
        [string]$OutFile
    )
    # Run Invoke-WebRequest in a background job and show a spinner while it downloads.
    $job = Start-Job -ScriptBlock { param($u,$o) Invoke-WebRequest -Uri $u -OutFile $o } -ArgumentList $Url, $OutFile
    $spinner = "|/-\"
    $i = 0
    while ($job.State -eq 'Running') {
        $char = $spinner[$i % $spinner.Length]
        Write-Progress -Activity "Downloading wsusscn2.cab" -Status "Downloading... $char" -PercentComplete 0
        Start-Sleep -Milliseconds 200
        $i++
    }
    # Capture any errors and complete progress
    Receive-Job $job -ErrorAction SilentlyContinue | Out-Null
    $jobState = $job.State
    Remove-Job $job -Force
    Write-Progress -Activity "Downloading wsusscn2.cab" -Completed
    return (Test-Path $OutFile)
}


Write-Host "Scan running. Check the log file at $LogFile once this window closes....."


# Always prompt to download a fresh wsusscn2.cab
if (Test-Path $WsusCab) {
    Write-Host "wsusscn2.cab found at $WsusCab."
    $download = Read-Host "Do you want to download a fresh copy now? (Y/N)"
    if ($download -eq 'Y' -or $download -eq 'y') {
        Write-Host "Downloading wsusscn2.cab from $WsusUrl ..."
            try {
            Invoke-WebRequest -Uri $WsusUrl -OutFile $WsusCab
            if (Test-Path $WsusCab) {
                Write-Host "Download successful."
            } else {
                Write-Error "Download failed. Please download manually from: $WsusUrl"
                exit 4
            }
        } catch {
            Write-Error "Download failed. Please download manually from: $WsusUrl"
            exit 4
        }
    }
} else {
    Write-Warning "wsusscn2.cab not found at $WsusCab."
    $download = Read-Host "Do you want to download a new copy now? (Y/N)"
    if ($download -eq 'Y' -or $download -eq 'y') {
        Write-Host "Downloading wsusscn2.cab from $WsusUrl ..."
            try {
            Invoke-WebRequest -Uri $WsusUrl -OutFile $WsusCab
            if (Test-Path $WsusCab) {
                Write-Host "Download successful."
            } else {
                Write-Error "Download failed. Please download manually from: $WsusUrl"
                exit 4
            }
        } catch {
            Write-Error "Download failed. Please download manually from: $WsusUrl"
            exit 4
        }
    } else {
        Write-Host "wsusscn2.cab is required. Exiting."
        exit 5
    }
}

# Get start timestamp
$startTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
Add-Content -Path $LogFile -Value ('*' * 99)
Add-Content -Path $LogFile -Value "Scan started at $startTimestamp"
Add-Content -Path $LogFile -Value ('*' * 99)
Add-Content -Path $LogFile -Value ""

# === Begin update scan logic ===
try {
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateServiceManager  = New-Object -ComObject Microsoft.Update.ServiceManager
    $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $WsusCab, 1)
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    Write-Output "Searching for updates... `r`n" | Tee-Object -FilePath $LogFile -Append
    $UpdateSearcher.ServerSelection = 3 #ssOthers
    $UpdateSearcher.IncludePotentiallySupersededUpdates = $true
    $UpdateSearcher.ServiceID = $UpdateService.ServiceID
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
    $Updates = $SearchResult.Updates
    if($Updates.Count -eq 0){
        Write-Output "There are no applicable updates." | Tee-Object -FilePath $LogFile -Append
    } else {
        Write-Output "List of applicable items on the machine when using wsusscn2.cab: `r`n" | Tee-Object -FilePath $LogFile -Append
        $i = 0
        $total = $Updates.Count
        foreach($Update in $Updates){
            $percent = [int]((($i+1)/[double]$total)*100)
            Write-Progress -Activity "Scanning for updates" -Status "Processing update $($i+1) of $total" -PercentComplete $percent
            Write-Host ("{0}> {1}" -f ($i+1), $Update.Title)
            Write-Output ("{0}> {1}" -f ($i+1), $Update.Title) | Tee-Object -FilePath $LogFile -Append
            $i++
        }
        Write-Progress -Activity "Scanning for updates" -Completed
    }
    $exitCode = 0
} catch {
    Write-Error "Update scan failed: $_"
    Add-Content -Path $LogFile -Value "Update scan failed: $_"
    $exitCode = 1
}
# === End update scan logic ===

# Get end timestamp
$endTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
Add-Content -Path $LogFile -Value ('*' * 99)
if ($exitCode -eq 0) {
    Add-Content -Path $LogFile -Value "Scan completed at $endTimestamp"
    Write-Host "Scan completed successfully."
} else {
    Add-Content -Path $LogFile -Value "Scan FAILED at $endTimestamp (exit code $exitCode)"
    Write-Host "Scan failed with exit code $exitCode."
}
Add-Content -Path $LogFile -Value ('*' * 99)
Add-Content -Path $LogFile -Value ""

exit $exitCode
