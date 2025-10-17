
# PowerShell version of Scan-UpdatesOffline.bat logic, with update scan
# Get the directory where this script/exe is located
if ($MyInvocation.MyCommand.Path) {
    # Running as a script
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    # Running as a compiled executable
    $ScriptDir = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}
$LogFile = Join-Path $ScriptDir "ScanReport.txt"
$WsusCab = Join-Path $ScriptDir "wsusscn2.cab"
$WsusUrl = "http://go.microsoft.com/fwlink/p/?LinkID=74689"
$WsusUrlBackup = "https://download.microsoft.com/download/9/3/9/939A4A46-91B6-4276-BC5F-9C9FF69B7DA2/wsusscn2.cab"
$WsusUrlBackup2 = "http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab"

# Helper function for reliable log writing
function Write-LogFile {
    param([string]$Message, [string]$FilePath)
    try {
        # Use StreamWriter for consistent, reliable output
        $sw = New-Object System.IO.StreamWriter($FilePath, $true, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($Message)
        $sw.Close()
    } catch {
        # Fallback method using Out-File
        try {
            $Message | Out-File -FilePath $FilePath -Append -Encoding UTF8
        } catch {
            # Last resort fallback
            [System.IO.File]::AppendAllText($FilePath, "$Message`r`n", [System.Text.Encoding]::UTF8)
        }
    }
}
# --- IGNORE ---

# Fail fast if not running elevated
try {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    $isAdmin = $false
}
if (-not $isAdmin) {
    Write-Host "This script must be run as Administrator. Please open an elevated PowerShell (Run as Administrator) and re-run the script." -ForegroundColor Yellow
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm') - Script exited: requires Administrator elevation." -Encoding UTF8
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

# Download helper with fallback URLs - optimized for speed
function Get-WsusCab {
    param(
        [string[]]$Urls,
        [string]$OutputPath,
        [string]$ExpectedHash = $null
    )
    
    foreach ($url in $Urls) {
        Write-Host "Attempting download from: $url"
        
        # Try multiple download methods for each URL
        $downloadMethods = @(
            { 
                # Check if BITS is available and service is running
                if (-not (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue)) {
                    throw "BITS cmdlets not available"
                }
                
                $bitsService = Get-Service -Name "BITS" -ErrorAction SilentlyContinue
                if (-not $bitsService -or $bitsService.Status -ne "Running") {
                    throw "BITS service not running (Status: $($bitsService.Status))"
                }
                
                Write-Host "Trying BITS transfer with progress..."
                
                # Test BITS functionality with a quick job creation test
                try {
                    $testJob = Start-BitsTransfer -Source $url -Destination $OutputPath -DisplayName "Downloading wsusscn2.cab" -Description "Windows Update Offline Scan Package" -Asynchronous -Priority High -ErrorAction Stop
                    $bitsJob = $testJob
                } catch {
                    # Check for specific BITS errors and provide clearer messages
                    $errorCode = $_.Exception.HResult
                    if ($errorCode -eq -2147024675 -or $_.Exception.Message -match "0x800704DD") {
                        throw "BITS network authentication error (service/network issue)"
                    } elseif ($_.Exception.Message -match "Access is denied") {
                        throw "BITS access denied (insufficient permissions)"
                    } else {
                        throw "BITS job creation failed: $($_.Exception.Message)"
                    }
                }
                
                # Monitor progress until complete (optimized for compiled executables)
                $lastPercent = -1
                Write-Host "Starting BITS transfer..." -NoNewline
                
                do {
                    Start-Sleep -Milliseconds 1000
                    $job = Get-BitsTransfer -JobId $bitsJob.JobId -ErrorAction SilentlyContinue
                    if ($job) {
                        $percent = if ($job.BytesTotal -gt 0) { [math]::Round(($job.BytesTransferred / $job.BytesTotal) * 100, 0) } else { 0 }
                        $transferredMB = [math]::Round($job.BytesTransferred / 1MB, 2)
                        $totalMB = [math]::Round($job.BytesTotal / 1MB, 2)
                        
                        # Show progress every 10% or every 10 seconds for compiled executables
                        if ($percent -ne $lastPercent -and ($percent % 10 -eq 0 -or $percent -gt $lastPercent + 5)) {
                            Write-Host "`r$percent% ($transferredMB MB / $totalMB MB)" -NoNewline
                            $lastPercent = $percent
                        }
                        
                        # Also try Write-Progress (may work in some compiled scenarios)
                        try {
                            Write-Progress -Activity "BITS Download: wsusscn2.cab" -Status "$transferredMB MB / $totalMB MB" -PercentComplete $percent
                        } catch {
                            # Ignore Write-Progress errors in compiled executables
                        }
                    }
                } while ($job -and ($job.JobState -eq "Transferring" -or $job.JobState -eq "Connecting"))
                
                Write-Host "`rDownload completed!                    "
                try { Write-Progress -Activity "BITS Download: wsusscn2.cab" -Completed } catch { }
                
                # Complete the job
                if ($job -and $job.JobState -eq "Transferred") {
                    Complete-BitsTransfer -BitsJob $job
                } elseif ($job) {
                    Remove-BitsTransfer -BitsJob $job
                    throw "BITS transfer failed with state: $($job.JobState)"
                } else {
                    throw "BITS job not found"
                }
            },
            { 
                Write-Host "Trying WebClient with progress..."
                $webClient = New-Object System.Net.WebClient
                
                # Add optimized progress event handler (throttled updates)
                $Global:downloadComplete = $false
                $Global:lastProgressUpdate = 0
                Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
                    $now = (Get-Date).Ticks
                    # Only update progress every 250ms to reduce overhead
                    if (($now - $Global:lastProgressUpdate) -gt 2500000) {
                        $percent = $Event.SourceEventArgs.ProgressPercentage
                        $received = [math]::Round($Event.SourceEventArgs.BytesReceived / 1MB, 2)
                        $total = [math]::Round($Event.SourceEventArgs.TotalBytesToReceive / 1MB, 2)
                        Write-Progress -Activity "Downloading wsusscn2.cab" -Status "$received MB / $total MB" -PercentComplete $percent
                        $Global:lastProgressUpdate = $now
                    }
                } | Out-Null
                
                Register-ObjectEvent -InputObject $webClient -EventName DownloadFileCompleted -Action {
                    $Global:downloadComplete = $true
                    Write-Progress -Activity "Downloading wsusscn2.cab" -Completed
                } | Out-Null
                
                try {
                    $webClient.DownloadFileAsync($url, $OutputPath)
                    while (-not $Global:downloadComplete) {
                        Start-Sleep -Milliseconds 200
                    }
                } finally {
                    Get-EventSubscriber | Where-Object {$_.SourceObject -eq $webClient} | Unregister-Event
                    $webClient.Dispose()
                    Remove-Variable -Name downloadComplete, lastProgressUpdate -Scope Global -ErrorAction SilentlyContinue
                }
            },
            { 
                Write-Host "Trying Invoke-WebRequest with progress..."
                # Use Invoke-WebRequest with progress (PowerShell 7+ has built-in progress)
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    Invoke-WebRequest -Uri $url -OutFile $OutputPath -TimeoutSec 60 -UseBasicParsing
                } else {
                    # For older PowerShell, use basic method
                    $ProgressPreference = 'Continue'
                    Invoke-WebRequest -Uri $url -OutFile $OutputPath -TimeoutSec 60 -UseBasicParsing
                }
            }
        )
        
        $downloadSucceeded = $false
        foreach ($method in $downloadMethods) {
            try {
                & $method
                $downloadSucceeded = $true
                break
            } catch {
                Write-Warning "Method failed: $($_.Exception.Message)"
                Remove-Item $OutputPath -Force -ErrorAction SilentlyContinue
                continue
            }
        }
        
        if (-not $downloadSucceeded) {
            Write-Warning "All download methods failed for $url"
            continue
        }
        
        # Verify the download
        if (Test-Path $OutputPath) {
            $actualHash = Get-FileHashString $OutputPath
            if ($ExpectedHash) {
                if ($actualHash -eq $ExpectedHash.ToUpper()) {
                    Write-Host "Download successful. Hash matches: $actualHash"
                    return $true
                } else {
                    Write-Warning "Download hash mismatch from $url! Expected: $ExpectedHash, Got: $actualHash"
                    Remove-Item $OutputPath -Force -ErrorAction SilentlyContinue
                    continue  # Try next URL
                }
            } else {
                Write-Host "Download successful from $url. SHA256: $actualHash"
                Write-Host "(Set `$ExpectedCabHash in the script to enforce hash validation.)"
                return $true
            }
        } else {
            Write-Warning "Download completed but file not found at $OutputPath"
            continue
        }
    }
    
    Write-Error "All download URLs failed. Please download manually from one of these URLs:"
    foreach ($url in $Urls) {
        Write-Host "  - $url"
    }
    return $false
}



Write-Host "Scan running. Check the log file at $LogFile once this window closes....."


# Always prompt to download a fresh wsusscn2.cab, with strict Y/N validation
function Read-YesNo($Prompt) {
    while ($true) {
        $resp = Read-Host $Prompt
        if ($resp -match '^[YyNn]$') { return $resp.ToUpper() }
        Write-Host "Please enter Y, or N."
    }
}
if (Test-Path $WsusCab) {
    Write-Host "wsusscn2.cab found at $WsusCab."
    $download = Read-YesNo "Do you want to download a fresh copy now? (Y/N)"
    if ($download -eq 'Y') {
        Write-Host "Downloading wsusscn2.cab with fallback URLs..."
        $urls = @($WsusUrl, $WsusUrlBackup, $WsusUrlBackup2)
        if (-not (Get-WsusCab -Urls $urls -OutputPath $WsusCab -ExpectedHash $ExpectedCabHash)) {
            exit 4
        }
    }
} else {
    Write-Warning "wsusscn2.cab not found at $WsusCab."
    $download = Read-YesNo "Do you want to download a new copy now? (Y/N)"
    if ($download -eq 'Y') {
        Write-Host "Downloading wsusscn2.cab with fallback URLs..."
        $urls = @($WsusUrl, $WsusUrlBackup, $WsusUrlBackup2)
        if (-not (Get-WsusCab -Urls $urls -OutputPath $WsusCab -ExpectedHash $ExpectedCabHash)) {
            exit 4
        }
    } else {
        Write-Host "wsusscn2.cab is required. Exiting."
        exit 5
    }
}

# Get start timestamp
$startTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
Write-LogFile -Message ('*' * 99) -FilePath $LogFile
Write-LogFile -Message "Scan started at $startTimestamp" -FilePath $LogFile
Write-LogFile -Message ('*' * 99) -FilePath $LogFile
Write-LogFile -Message "" -FilePath $LogFile

# === Begin update scan logic ===
try {
    Write-Host "Searching for updates..."
    Write-LogFile -Message "Searching for updates..." -FilePath $LogFile
    Write-LogFile -Message "" -FilePath $LogFile
    
    $UpdateSession = $null
    $UpdateServiceManager = $null
    $UpdateService = $null
    $UpdateSearcher = $null
    $SearchResult = $null
    
    try {
        # Create COM objects and perform the scan
        Write-Host "Creating Windows Update Session..."
        Write-LogFile -Message "Creating Windows Update Session..." -FilePath $LogFile
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        
        Write-Host "Creating Update Service Manager..."
        Write-LogFile -Message "Creating Update Service Manager..." -FilePath $LogFile
        $UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
        
        Write-Host "Adding scan package service..."
        Write-LogFile -Message "Adding scan package service using: $WsusCab" -FilePath $LogFile
        $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $WsusCab, 1)
        
        Write-Host "Creating Update Searcher..."
        Write-LogFile -Message "Creating Update Searcher..." -FilePath $LogFile
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        $UpdateSearcher.ServerSelection = 3 # ssOthers
        $UpdateSearcher.IncludePotentiallySupersededUpdates = $true
        $UpdateSearcher.ServiceID = $UpdateService.ServiceID
        
        Write-Host "Performing update search (this may take several minutes)..."
        Write-LogFile -Message "Performing update search (this may take several minutes)..." -FilePath $LogFile
        
        # Perform the search with timeout handling
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
        $Updates = $SearchResult.Updates
        
        Write-Host "Search completed. Found $($Updates.Count) updates."
        Write-LogFile -Message "Search completed. Found $($Updates.Count) updates." -FilePath $LogFile
        
        if ($Updates.Count -eq 0) {
            Write-Host "There are no applicable updates."
            Write-LogFile -Message "There are no applicable updates." -FilePath $LogFile
        } else {
            Write-Host "List of applicable items on the machine when using wsusscn2.cab:"
            Write-LogFile -Message "List of applicable items on the machine when using wsusscn2.cab:" -FilePath $LogFile
            Write-LogFile -Message "" -FilePath $LogFile
            $total = $Updates.Count
            for ($i = 0; $i -lt $total; $i++) {
                $Update = $Updates.Item($i)
                $percent = [int]((($i+1)/[double]$total)*100)
                Write-Progress -Activity "Listing updates" -Status "$($i+1) of $total" -PercentComplete $percent
                Write-Host ("{0}> {1}" -f ($i+1), $Update.Title)
                Write-LogFile -Message ("{0}> {1}" -f ($i+1), $Update.Title) -FilePath $LogFile
            }
            Write-Progress -Activity "Listing updates" -Completed
        }
        $exitCode = 0
    } finally {
        # Clean up COM objects
        foreach ($obj in @($SearchResult, $UpdateSearcher, $UpdateService, $UpdateServiceManager, $UpdateSession)) {
            if ($null -ne $obj) {
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
            }
        }
    }
} catch {
    $errorMsg = "Update scan failed: $($_.Exception.Message)"
    $errorDetails = "Full error: $($_.Exception.ToString())"
    $stackTrace = "Stack trace: $($_.ScriptStackTrace)"
    
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host $errorDetails -ForegroundColor Red
    
    Write-LogFile -Message $errorMsg -FilePath $LogFile
    Write-LogFile -Message $errorDetails -FilePath $LogFile
    Write-LogFile -Message $stackTrace -FilePath $LogFile
    
    # Check for common error scenarios
    if ($_.Exception.Message -match "0x80240024") {
        $msg = "Error 0x80240024: This typically means the wsusscn2.cab file is corrupted or invalid."
        Write-Host $msg -ForegroundColor Yellow
        Write-LogFile -Message $msg -FilePath $LogFile
    } elseif ($_.Exception.Message -match "0x8024001E") {
        $msg = "Error 0x8024001E: Windows Update service is not running or accessible."
        Write-Host $msg -ForegroundColor Yellow
        Write-LogFile -Message $msg -FilePath $LogFile
    } elseif ($_.Exception.Message -match "access.*denied") {
        $msg = "Access denied: Make sure you're running as Administrator."
        Write-Host $msg -ForegroundColor Yellow
        Write-LogFile -Message $msg -FilePath $LogFile
    }
    
    $exitCode = 1
}
# === End update scan logic ===

# Get end timestamp
$endTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
Write-LogFile -Message ('*' * 99) -FilePath $LogFile
if ($exitCode -eq 0) {
    Write-LogFile -Message "Scan completed at $endTimestamp" -FilePath $LogFile
    Write-Host "Scan completed successfully."
} else {
    Write-LogFile -Message "Scan FAILED at $endTimestamp (exit code $exitCode)" -FilePath $LogFile
    Write-Host "Scan failed with exit code $exitCode."
}
Write-LogFile -Message ('*' * 99) -FilePath $LogFile
Write-LogFile -Message "" -FilePath $LogFile

exit $exitCode
