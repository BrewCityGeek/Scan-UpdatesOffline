
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


# CAB hash validation helper
function Get-FileHashString {
    param([string]$Path)
    if (Test-Path $Path) {
        return (Get-FileHash -Path $Path -Algorithm SHA256).Hash.ToUpper()
    } else {
        return $null
    }
}

# Optionally set this to the expected SHA256 hash of wsusscn2.cab
$ExpectedCabHash = $null  # e.g. 'ABCDEF123456...'



Write-Host "Scan running. Check the log file at $LogFile once this window closes....."


# Always prompt to download a fresh wsusscn2.cab, with strict Y/N validation
function Read-YesNo($Prompt) {
    while ($true) {
        $resp = Read-Host $Prompt
        if ($resp -match '^[YyNn]$') { return $resp.ToUpper() }
        Write-Host "Please enter Y or N."
    }
}
if (Test-Path $WsusCab) {
    Write-Host "wsusscn2.cab found at $WsusCab."
    $download = Read-YesNo "Do you want to download a fresh copy now? (Y/N)"
    if ($download -eq 'Y') {
        Write-Host "Downloading wsusscn2.cab from $WsusUrl ..."
        try {
            Invoke-WebRequest -Uri $WsusUrl -OutFile $WsusCab
            if (Test-Path $WsusCab) {
                $actualHash = Get-FileHashString $WsusCab
                if ($ExpectedCabHash) {
                    if ($actualHash -eq $ExpectedCabHash.ToUpper()) {
                        Write-Host "Download successful. Hash matches: $actualHash"
                    } else {
                        Write-Error "Download hash mismatch! Expected: $ExpectedCabHash, Got: $actualHash"
                        Remove-Item $WsusCab -Force -ErrorAction SilentlyContinue
                        exit 6
                    }
                } else {
                    Write-Host "Download successful. SHA256: $actualHash"
                    Write-Host "(Set $ExpectedCabHash in the script to enforce hash validation.)"
                }
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
    $download = Read-YesNo "Do you want to download a new copy now? (Y/N)"
    if ($download -eq 'Y') {
        Write-Host "Downloading wsusscn2.cab from $WsusUrl ..."
        try {
            Invoke-WebRequest -Uri $WsusUrl -OutFile $WsusCab
            if (Test-Path $WsusCab) {
                $actualHash = Get-FileHashString $WsusCab
                if ($ExpectedCabHash) {
                    if ($actualHash -eq $ExpectedCabHash.ToUpper()) {
                        Write-Host "Download successful. Hash matches: $actualHash"
                    } else {
                        Write-Error "Download hash mismatch! Expected: $ExpectedCabHash, Got: $actualHash"
                        Remove-Item $WsusCab -Force -ErrorAction SilentlyContinue
                        exit 6
                    }
                } else {
                    Write-Host "Download successful. SHA256: $actualHash"
                    Write-Host "(Set $ExpectedCabHash in the script to enforce hash validation.)"
                }
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
    $scanJob = $null
    try {
        # Start the scan in a background job
        $scanJob = Start-Job -ScriptBlock {
            $UpdateSession = $null
            $UpdateServiceManager = $null
            $UpdateService = $null
            $UpdateSearcher = $null
            try {
                $UpdateSession = New-Object -ComObject Microsoft.Update.Session
                $UpdateServiceManager  = New-Object -ComObject Microsoft.Update.ServiceManager
                $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $using:WsusCab, 1)
                $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
                $UpdateSearcher.ServerSelection = 3 #ssOthers
                $UpdateSearcher.IncludePotentiallySupersededUpdates = $true
                $UpdateSearcher.ServiceID = $UpdateService.ServiceID
                $result = $UpdateSearcher.Search("IsInstalled=0")
            } finally {
                foreach ($obj in @($UpdateSearcher, $UpdateService, $UpdateServiceManager, $UpdateSession)) {
                    if ($null -ne $obj) {
                        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
                    }
                }
            }
            return $result
        }
        Write-Output "Searching for updates... `r`n" | Tee-Object -FilePath $LogFile -Append
        # Show a busy progress bar while the scan runs
        $i = 0
        while ($scanJob.State -eq 'Running') {
            Write-Progress -Activity "Scanning for updates..." -Status "Please wait" -PercentComplete (($i % 100))
            Start-Sleep -Milliseconds 300
            $i += 5
        }
        Write-Progress -Activity "Scanning for updates..." -Completed
        if ($scanJob.State -ne 'Completed') {
            $jobError = $scanJob.ChildJobs[0].JobStateInfo.Reason
            throw "Scan job failed: $jobError"
        }
        $SearchResult = Receive-Job $scanJob
        $Updates = $SearchResult.Updates
        if($Updates.Count -eq 0){
            Write-Output "There are no applicable updates." | Tee-Object -FilePath $LogFile -Append
        } else {
            Write-Output "List of applicable items on the machine when using wsusscn2.cab: `r`n" | Tee-Object -FilePath $LogFile -Append
            $total = $Updates.Count
            for ($i = 0; $i -lt $total; $i++) {
                $Update = $Updates.Item($i)
                $percent = [int]((($i+1)/[double]$total)*100)
                Write-Progress -Activity "Listing updates" -Status "$($i+1) of $total" -PercentComplete $percent
                Write-Host ("{0}> {1}" -f ($i+1), $Update.Title)
                Write-Output ("{0}> {1}" -f ($i+1), $Update.Title) | Tee-Object -FilePath $LogFile -Append
            }
            Write-Progress -Activity "Listing updates" -Completed
        }
        $exitCode = 0
    } finally {
        if ($scanJob) { Remove-Job $scanJob -Force -ErrorAction SilentlyContinue }
    }
} catch {
    Write-Error "Update scan failed: $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value "Update scan failed: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
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
