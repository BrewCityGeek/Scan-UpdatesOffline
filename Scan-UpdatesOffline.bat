@echo off
REM Improved Scan-UpdatesOffline.bat with error handling, accurate timestamps, and configurability
setlocal enabledelayedexpansion


REM === CONFIGURABLE VARIABLES ===
set "LOGFILE=c:\scripts\ScanReport.txt"
set "VBSFILE=c:\scripts\Scan-UpdatesOffline.vbs"
set "WSUSCAB=c:\scripts\wsusscn2.cab"
set "WSUS_URL=http://go.microsoft.com/fwlink/p/?LinkID=74689"


REM === CHECK FOR WSUSSCN2.CAB ===
if exist "%WSUSCAB%" (
      echo wsusscn2.cab found at %WSUSCAB%.
      set /p DLOAD="Do you want to download a fresh copy now? (Y/N): "
      if /i "!DLOAD!"=="Y" (
            echo Downloading wsusscn2.cab from %WSUS_URL% ...
            powershell -NoProfile -ep bypass -Command "try { Invoke-WebRequest -Uri '%WSUS_URL%' -OutFile '%WSUSCAB%' -UseBasicParsing } catch { Write-Host 'Download failed.'; exit 1 }"
            if exist "%WSUSCAB%" (
                  echo Download successful.
            ) else (
                  echo ERROR: Download failed. Please download manually from:
                  echo %WSUS_URL%
                  exit /b 4
            )
      )
) else (
      echo WARNING: wsusscn2.cab not found at %WSUSCAB%.
      set /p DLOAD="Do you want to download a new copy now? (Y/N): "
      if /i "!DLOAD!"=="Y" (
            echo Downloading wsusscn2.cab from %WSUS_URL% ...
            powershell -NoProfile -ep bypass -Command "try { Invoke-WebRequest -Uri '%WSUS_URL%' -OutFile '%WSUSCAB%' -UseBasicParsing } catch { Write-Host 'Download failed.'; exit 1 }"
            if exist "%WSUSCAB%" (
                  echo Download successful.
            ) else (
                  echo ERROR: Download failed. Please download manually from:
                  echo %WSUS_URL%
                  exit /b 4
            )
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
>> "%VBSFILE%" echo If searchResult.Updates.Count = 0 Then
>> "%VBSFILE%" echo       WScript.Echo "There are no applicable updates."
>> "%VBSFILE%" echo       WScript.Quit
>> "%VBSFILE%" echo End If
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo WScript.Echo "List of applicable items on the machine when using wssuscan.cab:" ^& vbCRLF
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo For I = 0 to searchResult.Updates.Count-1
>> "%VBSFILE%" echo       Set update = searchResult.Updates.Item(I)
>> "%VBSFILE%" echo       WScript.Echo I + 1 ^& "^> " ^& update.Title
>> "%VBSFILE%" echo Next
>> "%VBSFILE%" echo.
>> "%VBSFILE%" echo WScript.Quit

echo Scan running. Check the log file at %LOGFILE% once this window closes.....

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

REM === RUN THE VBS SCRIPT AND CAPTURE EXIT CODE ===
cscript.exe "%VBSFILE%" >> "%LOGFILE%" 2>&1
set "_exitcode=!errorlevel!"

REM === GET END TIMESTAMP (using PowerShell) ===
for /f "delims=" %%a in ('powershell -NoProfile -ep bypass -Command "Get-Date -Format yyyy-MM-dd HH:mm"') do set "_datetimestamp=%%a"

echo ******************************************************************************************* >> "%LOGFILE%"
if !_exitcode! equ 0 (
      echo Scan completed at !_datetimestamp! >> "%LOGFILE%"
      echo Scan completed successfully.
) else (
      echo Scan FAILED at !_datetimestamp! (exit code !_exitcode!) >> "%LOGFILE%"
      echo Scan failed with exit code !_exitcode!.
)
echo ******************************************************************************************* >> "%LOGFILE%"
echo. >> "%LOGFILE%"
exit /b !_exitcode!
