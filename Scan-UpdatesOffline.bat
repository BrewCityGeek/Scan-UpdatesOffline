@echo off
REM Improved Scan-UpdatesOffline.bat with error handling, accurate timestamps, and configurability
setlocal enabledelayedexpansion


REM === CONFIGURABLE VARIABLES ===
set "LOGFILE=c:\scripts\ScanReport.txt"
set "VBSFILE=c:\scripts\Scan-UpdatesOffline.vbs"
set "WSUSCAB=c:\scripts\wsusscn2.cab"
set "WSUS_URL=http://go.microsoft.com/fwlink/p/?LinkID=74689"
set "WSUS_URL_BACKUP=https://download.microsoft.com/download/9/3/9/939A4A46-91B6-4276-BC5F-9C9FF69B7DA2/wsusscn2.cab"
set "WSUS_URL_BACKUP2=http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab"


REM === DOWNLOAD FUNCTION WITH FALLBACK ===
:download_wsus_cab
echo Attempting download from primary URL: %WSUS_URL%
powershell -NoProfile -ep bypass -Command "try { Invoke-WebRequest -Uri '%WSUS_URL%' -OutFile '%WSUSCAB%' -UseBasicParsing -TimeoutSec 60 } catch { Write-Host 'Primary URL failed: ' + $_.Exception.Message; }"
if exist "%WSUSCAB%" (
    echo Download successful from primary URL.
    goto :eof
) else (
    echo PowerShell download failed from primary URL.
)
echo Attempting download from backup URL: %WSUS_URL_BACKUP%
powershell -NoProfile -ep bypass -Command "try { Invoke-WebRequest -Uri '%WSUS_URL_BACKUP%' -OutFile '%WSUSCAB%' -UseBasicParsing -TimeoutSec 60 } catch { Write-Host 'Backup URL failed: ' + $_.Exception.Message; exit 1 }; if (!(Test-Path '%WSUSCAB%')) { exit 1 }"
if exist "%WSUSCAB%" (
    echo Download successful from backup URL.
    goto :eof
)
echo Attempting download from backup URL 2: %WSUS_URL_BACKUP2%
powershell -NoProfile -ep bypass -Command "try { Invoke-WebRequest -Uri '%WSUS_URL_BACKUP2%' -OutFile '%WSUSCAB%' -UseBasicParsing -TimeoutSec 60 } catch { Write-Host 'Backup URL 2 failed: ' + $_.Exception.Message; exit 1 }; if (!(Test-Path '%WSUSCAB%')) { exit 1 }"
if exist "%WSUSCAB%" (
    echo Download successful from backup URL 2.
    goto :eof
)

echo ERROR: All download URLs failed. Please download manually from one of these URLs:
echo   Primary: %WSUS_URL%
echo   Backup:  %WSUS_URL_BACKUP%
exit /b 4


REM === CHECK FOR WSUSSCN2.CAB ===
if exist "%WSUSCAB%" (
      echo wsusscn2.cab found at %WSUSCAB%.
      set /p DLOAD="Do you want to download a fresh copy now? (Y/N): "
      if /i "!DLOAD!"=="Y" (
            echo Downloading wsusscn2.cab with fallback URLs...
            call :download_wsus_cab
            set "_dlerr=!errorlevel!"
            if !_dlerr! neq 0 exit /b !_dlerr!
      )
) else (
      echo WARNING: wsusscn2.cab not found at %WSUSCAB%.
      set /p DLOAD="Do you want to download a new copy now? (Y/N): "
      if /i "!DLOAD!"=="Y" (
            echo Downloading wsusscn2.cab with fallback URLs...
            call :download_wsus_cab
            if !errorlevel! neq 0 exit /b !errorlevel!
      ) else (
            echo wsusscn2.cab is required. Exiting.
            exit /b 5
      )
)


REM === ENSURE Scan-UpdatesOffline.vbs EXISTS ===
echo Creating VBS file: %VBSFILE%
> "%VBSFILE%" echo Set UpdateSession = CreateObject("Microsoft.Update.Session")
>> "%VBSFILE%" echo Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
>> "%VBSFILE%" echo Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", "c:\scripts\wsusscn2.cab", 1)
>> "%VBSFILE%" echo Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo WScript.Echo "Searching for updates..." ^& vbCRLF
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo UpdateSearcher.ServerSelection = 3 ' ssOthers
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo UpdateSearcher.IncludePotentiallySupersededUpdates = True 'good for older OSes, to include Security-Only or superseded updates in the result list, otherwise these are pruned out and not returned as part of the final result list
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo UpdateSearcher.ServiceID = UpdateService.ServiceID
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo Set SearchResult = UpdateSearcher.Search("IsInstalled=0") 'or "IsInstalled=0 or IsInstalled=1" to also list the installed updates as MBSA did
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo Set Updates = SearchResult.Updates
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo If SearchResult.Updates.Count = 0 Then
>> "%VBSFILE%" echo       WScript.Echo "There are no applicable updates."
>> "%VBSFILE%" echo       WScript.Quit
>> "%VBSFILE%" echo End If
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo WScript.Echo "List of applicable items on the machine when using wssuscan.cab:" ^& vbCRLF
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo For I = 0 to SearchResult.Updates.Count-1
>> "%VBSFILE%" echo       Set update = SearchResult.Updates.Item(I)
>> "%VBSFILE%" echo       WScript.Echo I + 1 ^& "^> " ^& update.Title
>> "%VBSFILE%" echo Next
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo WScript.Quit

echo Scan is running. Results will be saved to "%LOGFILE%" after completion. Please review that file for the update scan report.

REM === CHECK DEPENDENCIES ===
if not exist "%VBSFILE%" (
      echo ERROR: VBS script not found: %VBSFILE%
      exit /b 2
)
where cscript >nul 2>&1
if errorlevel 1 (
      echo ERROR: cscript.exe not found in PATH.
      exit /b 3
)

REM === GET START TIMESTAMP (using PowerShell) ===
for /f "delims=" %%a in ('powershell -NoProfile -ep bypass -Command "Get-Date -Format yyyy-MM-dd HH:mm"') do set "_datetimestamp=%%a"

echo ******************************************************************************************* >> "%LOGFILE%"
echo Scan started at !_datetimestamp! >> "%LOGFILE%"
echo ******************************************************************************************* >> "%LOGFILE%"
echo. >> "%LOGFILE%"

REM === CHECK IF LOG FILE IS WRITABLE ===
echo. >> "%LOGFILE%" 2>nul
if errorlevel 1 (
    echo ERROR: Cannot write to log file: %LOGFILE%
    exit /b 6
)

REM === RUN THE VBS SCRIPT AND CAPTURE EXIT CODE ===
cscript.exe "%VBSFILE%" >> "%LOGFILE%" 2>&1
set "_exitcode=!errorlevel!"

REM === GET END TIMESTAMP (using PowerShell) ===
for /f "delims=" %%a in ('powershell -NoProfile -ep bypass -Command "Get-Date -Format yyyy-MM-dd HH:mm"') do set "_datetimestamp=%%a"

echo ******************************************************************************************* >> "%LOGFILE%"
if !_exitcode! equ 0 (
echo. >> "%LOGFILE%"
REM Ensure delayed expansion is still enabled before exit
setlocal enabledelayedexpansion
exit /b !_exitcode!pleted successfully.
) else (
      echo Scan FAILED at !_datetimestamp! (exit code !_exitcode!) >> "%LOGFILE%"
      echo Scan failed with exit code !_exitcode!.
)
echo ******************************************************************************************* >> "%LOGFILE%"
echo. >> "%LOGFILE%"
exit /b !_exitcode!
