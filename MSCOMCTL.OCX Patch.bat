@echo off
REM Author: Adrian van den Houten
REM Date: 2013/11/13 - 05:24 PM
REM *********************************************************
REM Disclaimer:
REM Do not edit/decompile or extract this program.
REM I am not responsible for any damage or loss of data.
REM It is strictly forbidden to sell this program without
REM permission of the author!
REM *********************************************************
setlocal enableextensions enabledelayedexpansion
title XBC Error 339 Patch
color b
reg query "HKU\S-1-5-19" >nul 2>&1 && (
goto start
) || (
mode 45,10
echo Run me as Administrator
ping localhost -n 3 >nul
exit )

:start
set "FileName=%~d0%~p0"
set "Reg1=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\SharedDLLs"
set "Reg2=HKEY_CLASSES_ROOT\TypeLib\{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}\2.0\0\win32"

:mainmenu
mode 45,10
echo\
echo         ---- XBC Error 339 Patch ----
echo          - Version 0.1.5 By Adrian -
echo\
echo                 1. Run patch
echo                 2. Restore
echo                 3. Download
echo\
set "errorlevel="
<nul set /p "=Make your selection: " &call :choice 123
if "%errorlevel%"=="1" goto PatchMain
if "%errorlevel%"=="2" goto restore
if "%errorlevel%"=="3" goto download

::Patch File START
:PatchMain
if exist "%FileName%MSCOMCTL.OCX" (set "FileName=%FileName%MSCOMCTL.OCX" & goto definedfile) else (goto enterOCXfiledir)
:enterOCXfiledir
set "FileName="
mode 65,10
for %%i in (powershell.exe) do if "%%~$path:i"=="" (goto powershellexplorer ) else (goto oldOCXfileexplorer )

:powershellexplorer
set "FileName="
echo Please select the MSCOMCTL.OCX file...
echo Processing...
set "ps=Add-Type -AssemblyName System.windows.forms | Out-Null;"
set "ps=%ps% $f=New-Object System.Windows.Forms.OpenFileDialog;"
set "ps=%ps% $f.Filter='MSCOMCTL.OCX|MSCOMCTL.OCX';"
set "ps=%ps% $f.showHelp=$true;"
set "ps=%ps% $f.ShowDialog() | Out-Null;"
set "ps=%ps% $f.FileName"
for /f "delims=" %%I in ('powershell "%ps%"') do set "FileName=%%I"
if "%FileName%"=="" goto notdefinedfile
if defined FileName goto definedfile
if not defined FileName goto notdefinedfile

:notdefinedfile
echo\
echo You did not select the MSCOMCTL.OCX file...
ping localhost -n 3 >nul
cls
goto mainmenu
:definedfile
mode 65,10
if exist "%FileName%" (
echo Found file...
echo\ %FileName%
ping localhost -n 3 >nul
set "OCXuserinpdir=%FileName%"
cls
goto patchfileOSget
) else (
echo ERROR - Defined address file not found...
echo\ %FileName%
ping localhost -n 2 >nul
echo Try again...
ping localhost -n 2 >nul
cls & goto mainmenu )

:oldOCXfileexplorer
echo Please enter the directory of MSCOMCTL.OCX
echo example... %homedrive%\Users\%username%\Desktop
echo b/back
set /p patchfiledir=
if "%patchfiledir%"=="b" goto mainmenu
if exist "%patchfiledir%\MSCOMCTL.OCX" (
echo Found file in directory...
echo %patchfiledir%\MSCOMCTL.OCX
ping localhost -n 3 >nul
set OCXuserinpdir=%patchfiledir%\MSCOMCTL.OCX
cls
goto patchfileOSget
) else (
echo ERROR - Defined file address not found...
echo\"%patchfiledir%"
ping localhost -n 2 >nul
echo\
echo Place the MSCOMCTL.OCX in a directory
ping localhost -n 4 >nul
cls & goto mainmenu )


:patchfileOSget
mode 45,10
echo Scanning...
ping localhost -n 2 >nul
if "%PROCESSOR_ARCHITECTURE%"=="x86" goto 32BITpatch
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" goto 64BITpatch
goto Unknown

::32-Bit Patch START
:32BITpatch
set "OCXwindfiledir=%homedrive%\Windows\System32"
echo PC is 32-bit
ping localhost -n 2 >nul
echo Patching File...
ping localhost -n 2 >nul
copy "%OCXuserinpdir%" "%OCXwindfiledir%" >nul
if "%errorlevel%"=="0" goto run32patch
if "%errorlevel%"=="1" goto patchfail
if "%errorlevel%"=="9009" goto patchfail
:run32patch
Regsvr32 /s "%OCXwindfiledir%\MSCOMCTL.OCX"
if "%errorlevel%"=="0" (
echo Done Patching...
ping localhost -n 2 >nul
cls
goto restart )
if "%errorlevel%"=="1" goto patchfail
if "%errorlevel%"=="9009" goto patchfail
::32-Bit Patch END

reg query "%Reg1%" /v "%homedrive%\Windows\system32\MSCOMCTL.OCX" >nul 2>nul
if "%errorlevel%"=="0" for /f "tokens=2*" %%i in ('reg query "%Reg1%" /v "C:\Windows\system32\MSCOMCTL.OCX"') do (set "Reg1RegString=%%j"
) else (
reg query "%Reg2%" /v "^(Default)" >nul 2>nul
if "%errorlevel%"=="0" for /f "tokens=2*" %%i in ('reg query "%Reg1%" /v "^(Default)"') do set "Reg2RegString=%%j" )


if "%Reg1RegString%"=="2" exit /b
if "%Reg2RegString%"=="C:\Windows\SysWOW64\MSCOMCTL.OCX" exit /b


::64-Bit Patch START
:64BITpatch
set "OCXwindfiledir=%homedrive%\Windows\SysWOW64"
echo PC is 64-bit
ping localhost -n 2 >nul
echo Patching File...
ping localhost -n 2 >nul
copy "%OCXuserinpdir%" "%OCXwindfiledir%" >nul
if "%errorlevel%"=="0" goto run64patch
if "%errorlevel%"=="1" goto patchfail
if "%errorlevel%"=="9009" goto patchfail
:run64patch
Regsvr32 /s "%OCXwindfiledir%\MSCOMCTL.OCX"
if "%errorlevel%"=="0" (
echo Done Patching...
ping localhost -n 2 >nul
cls
goto restart )
if "%errorlevel%"=="1" goto patchfail
if "%errorlevel%"=="9009" goto patchfail
::64-Bit Patch END
:patchfail
echo Patching Failed!
ping localhost -n 2 >nul
echo Try Again Later...
ping localhost -n 3 >nul
cls
goto help
::Patch END


::Restore File START
:restore
cls
echo Scanning...
ping localhost -n 2 >nul
if "%PROCESSOR_ARCHITECTURE%"=="x86" goto 32BITrestore
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" goto 64BITrestore
goto Unknown

::32-Bit Restore START
:32BITrestore
set "OCXwindfiledir=%HOMEDRIVE%\Windows\System32\MSCOMCTL.OCX"
echo PC is 32-bit
ping localhost -n 2 >nul
echo Restoring File...
ping localhost -n 2 >nul
Regsvr32 /u /s "%OCXwindfiledir%" >nul
if "%errorlevel%"=="0" goto run32restore
if "%errorlevel%"=="1" goto restorefail
if "%errorlevel%"=="9009" goto restorefail
:run32restore
if exist "%OCXwindfiledir%" del "%OCXwindfiledir%" >nul
if "%errorlevel%"=="0" (
echo Done Restoring...
ping localhost -n 2 >nul
cls
goto restart )
if "%errorlevel%"=="1" goto restorefail
if "%errorlevel%"=="9009" goto restorefail
::32-Bit Patch END

::64-Bit Restore START
:64BITrestore
set "OCXwindfiledir=%HOMEDRIVE%\Windows\SysWOW64\MSCOMCTL.OCX"
echo PC is 64-bit
ping localhost -n 2 >nul
echo Restoring File...
ping localhost -n 2 >nul
Regsvr32 /u /s "%OCXwindfiledir%" >nul
if "%errorlevel%"=="0" goto run64restore
if "%errorlevel%"=="1" goto restorefail
if "%errorlevel%"=="9009" goto restorefail
:run64restore
if exist "%OCXwindfiledir%" del "%OCXwindfiledir%" >nul
if "%errorlevel%"=="0" (
echo Done Restoring...
ping localhost -n 2 >nul
cls
goto restart )
if "%errorlevel%"=="1" goto restorefail
if "%errorlevel%"=="9009" goto restorefail
::64-Bit Patch END
:restorefail
echo Restoring Failed!
ping localhost -n 2 >nul
echo Try Again Later...
ping localhost -n 3 >nul
cls
goto mainmenu
::Restore END

::Restart Option START
:restart
set restartUI=restartUI
echo\
echo Would you like to restart your PC?
echo 1) Yes (recommended)
echo 2) No restart later
echo\
set /p restartUI=Choose:
cls
if "%restartUI%"=="1" goto restartnow
if "%restartUI%"=="2" goto mainmenu
if "%restartUI%"=="restartUI" goto restart
goto restart
:restartnow
echo Restarting PC...
ping localhost -n 2 >nul
shutdown -r -t 0
exit
::Restart Option END

::Help START
:help
cls
mode 58,15
color c
echo ACCESS DENIED - UAC setting causing program not to work...
ping localhost -n 2 >nul
echo ...ADVICE... (at your own risk)
echo 1) Hit start button
echo 2) Type UAC
echo 3) Click Change User Account Control settings
echo 4) Change mode to bottom one ("Never notify me when:")
echo 5) Restart your PC
echo 6) Run XBC Error 339 Patch
echo\
echo -NOTE change UAC settings back to default every day or so
echo\
pause
ping localhost -n 1 >nul
cls
exit
::Help END

::Unknown CPU START
:Unknown
echo Error - You CPU is not reconised with this program...
ping localhost -n 2
echo\
echo Pleas try again later...
ping localhost -n 2
cls
exit
::Unknown CPU END


:PSWebClient
setlocal
set "ps="
set "ps=%ps%try {"
set "ps=%ps%  $filename = \"!PSClientFileName!\";"
set "ps=%ps%  $url = \"!PSClientURL!\";"
set "ps=%ps%  $client = new-object System.Net.WebClient;"
set "ps=%ps%  $client.DownloadFile($url, $filename);"
set "ps=%ps%  Exit 1"
set "ps=%ps%}"
set "ps=%ps%catch [System.Net.WebException] {"
set "ps=%ps%  Exit 2"
set "ps=%ps%}"
set "ps=%ps%catch [System.IO.IOException] {"
set "ps=%ps%  Exit 3"
set "ps=%ps%}"
set "ps=%ps%catch {"
set "ps=%ps%  Exit 4"
set "ps=%ps%}"
set "ps=%ps%Exit 0"
powershell Set-ExecutionPolicy Unrestricted
powershell -ExecutionPolicy RemoteSigned -Command "%ps%"  >nul 2>nul
call :title & exit /b
endlocal


:choice
setlocal DisableDelayedExpansion
set "n=0" &set "c=" &set "e=" &set "map=%~1"
if not defined map endlocal &exit /b 0
for /f "delims=" %%i in ('2^>nul xcopy /lw "%~f0" "%~f0"') do if not defined c set "c=%%i"
set "c=%c:~-1%"
if defined c (
  for /f delims^=^ eol^= %%i in ('cmd /von /u /c "echo(!map!"^|find /v ""^|findstr .') do (
    set /a "n += 1" &set "e=%%i"
    setlocal EnableDelayedExpansion
    if /i "!e!"=="!c!" (
      echo(!c!
      for /f %%j in ("!n!") do endlocal &endlocal &exit /b %%j
    )
    endlocal
  )
)
endlocal &goto choice