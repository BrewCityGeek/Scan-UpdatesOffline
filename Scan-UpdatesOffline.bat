@echo Off

Echo Scan running. Check the log file at c:\Scripts\scanreport_{date}_{time}.txt once this window closes.....

:: Use WMIC to retrieve date and time
FOR /F "skip=1 tokens=1-6" %%G IN ('WMIC Path Win32_LocalTime Get Day^,Hour^,Minute^,Month^,Second^,Year /Format:table') DO (
   IF "%%~L"=="" goto s_done
      Set _yyyy=%%L
      Set _mm=00%%J
      Set _dd=00%%G
      Set _hour=00%%H
      SET _minute=00%%I
)
:s_done

:: Pad digits with leading zeros
      Set _mm=%_mm:~-2%
      Set _dd=%_dd:~-2%
      Set _hour=%_hour:~-2%
      Set _minute=%_minute:~-2%

:: Display the date/time in ISO 8601 format:
Set _datetimestamp=%_yyyy%-%_mm%-%_dd% %_hour%:%_minute%
Set _datestamp=%_yyyy%%_mm%%_dd%_%_hour%%_minute%

Echo ******************************************************************************************* >> c:\scripts\ScanReport_%_datestamp%.txt
Echo Scan started at %_datetimestamp% >> c:\scripts\scanreport_%_datestamp%.txt
Echo ******************************************************************************************* >> c:\scripts\ScanReport_%_datestamp%.txt
Echo. >> c:\scripts\scanreport_%_datestamp%.txt

cscript.exe c:\scripts\Scan-UpdatesOffline.vbs >> c:\scripts\scanreport_%_datestamp%.txt

:: Use WMIC to retrieve date and time
FOR /F "skip=1 tokens=1-6" %%G IN ('WMIC Path Win32_LocalTime Get Day^,Hour^,Minute^,Month^,Second^,Year /Format:table') DO (
   IF "%%~L"=="" goto s_done
      Set _yyyy=%%L
      Set _mm=00%%J
      Set _dd=00%%G
      Set _hour=00%%H
      SET _minute=00%%I
)
:s_done

:: Pad digits with leading zeros
      Set _mm=%_mm:~-2%
      Set _dd=%_dd:~-2%
      Set _hour=%_hour:~-2%
      Set _minute=%_minute:~-2%

:: Display the date/time in ISO 8601 format:
Set _datetimestamp=%_yyyy%-%_mm%-%_dd% %_hour%:%_minute%

Echo ******************************************************************************************* >> c:\scripts\scanreport_%_datestamp%.txt
Echo Scan completed at %_datetimestamp% >> c:\scripts\scanreport_%_datestamp%.txt
Echo ******************************************************************************************* >> c:\scripts\scanreport_%_datestamp%.txt
Echo. >> c:\scripts\scanreport_%_datestamp%.txt