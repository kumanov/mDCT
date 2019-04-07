:: Filename: mDCT.cmd - mini Data Collection Tool script + ext
@if "%_ECHO%" == "" ECHO OFF
setlocal enableDelayedExpansion

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
:: mini DCT +ext batch script file  (krasimir.kumanov@gmail.com)
::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
::  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
::  IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

:: initialize variables
set _ScriptVersion=v1.05
:: Last-Update by krasimir.kumanov@gmail.com: 2019-04-03

:: change the cmd prompt environment to English
chcp 437 >NUL

:: Adding a Window Title
REM.-- Set the title
SET title=%~nx0 - version %_ScriptVersion%
TITLE %title%

::@::if defined _DbgOut ( echo.%time% : Start of mDCT ^(%_ScriptVersion% - krasimir.kumanov@gmail.com^))
call :WriteHostNoLog white black %date% %time% : Start of mDCT [%_ScriptVersion% - krasimir.kumanov@gmail.com]

:: handle /?
if "%~1"=="/?" (
	call :usage
	@exit /b 0
)

@rem parsed args - off
@set _Usage=
@set _noCabZip=

call :ArgsParse %*

if /i "%_Usage%" EQU "1" ( call :usage
					@exit /b 0 )

set _DirScript=%~dp0
if defined _DbgOut ( echo. %time% _DirScript: %_DirScript% )
if not "%_DirScript%"=="%_DirScript: =%" ( echo.
	echo  *** Your script execution path '%_DirScript%' contains one or more space characters, please use a different local path without space characters.
	exit /b 1
	)

:: Change Directory to the location of the batch script file (%0)
CD /d "%_DirScript%"
@echo. .. starting '%_DirScript%%~n0 %*'


:: _OSVER* will be set in function getWinVer
call :getWinVer
call :getDateTime
for /f "delims=" %%a in ('PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { $PSVersionTable.PSVersion.Major }}"') do set _PSVer=%%a
set _ProcArch=%PROCESSOR_ARCHITECTURE%
if "!_ProcArch!" equ "AMD64" Set _ProcArch=x64

set _Comp_Time=%COMPUTERNAME%_%_CurDateTime%
if defined _DbgOut ( echo. %time% _Comp_Time: %_Comp_Time% )
:: set work folder
set _DirWork=%_DirScript%%_Comp_Time%

if defined _DbgOut ( echo. %time% _DirWork: !_DirWork! )


:: check current user account permissions
call :check_Permissions
if "%errorlevel%" neq "0" goto :end


:Configuration parameters ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Settings for quick testing: in CMD use: 'set _DbgOut=1', use SET _Echo=1 to debug this script

:: regional settings ^(choose EN-US, DE-DE^) for localized Perfmon counter names, hardcode to EN-US, or choose _GetLocale=1
	@set _locale=EN-US
	@set _GetLocale=1
:end_of_Configuration parameters ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


goto :mDCT_Begin


:ArgsParse
	:: count number of arguments
	set _argCount=0
	for %%x in (%*) do ( set /A _argCount+=1 )

	:: end parse, if no more args
	IF "%~1"=="" exit /b
	
	IF "%~1"=="-?" set _Usage=1

	for %%i in (help /help -help -h /h) do (
		if /i "%~1" equ "%%i" (set _Usage=1)
	)

	IF /i "%~1"=="all" set _ALL=1
	IF /i "%~1"=="noblg" set _NoBlg=1
	if /i "%~1"=="nocabzip" (set _noCabZip=1)

	SHIFT
	GOTO ArgsParse


:check_Permissions
	@echo.  Administrative permissions required. Detecting permissions...
	net session >NUL 2>&1
	if %errorLevel% == 0 (
		call :WriteHostNoLog green black Success: Administrative permissions confirmed.
		@exit /b 0
	) else (
		color 0e
		@echo.
		@echo.  Failure: Current permissions inadequate.
		@echo.  Please relaunch the command prompt with administrator privileges in elevated CMD.
		@pause
		@exit /b 1
	)



:getWinVer - UTILITY to get Windows Version
	:: #########################
	:: OS-specific checks...
	:: #########################
	for /f "tokens=2 delims=[]" %%o in ('ver')    do @set _OSVERTEMP=%%o
	for /f "tokens=2" %%o in ('echo %_OSVERTEMP%') do @set _OSVER=%%o
	for /f "tokens=1 delims=." %%o in ('echo %_OSVER%') do @set _OSVER1=%%o
	for /f "tokens=2 delims=." %%o in ('echo %_OSVER%') do @set _OSVER2=%%o
	for /f "tokens=3 delims=." %%o in ('echo %_OSVER%') do @set _OSVER3=%%o
	for /f "tokens=4 delims=." %%o in ('echo %_OSVER%') do @set _OSVER4=%%o
	for /f "tokens=4-8 delims=[.] " %%i in ('ver') do (if %%i==Version (set v=%%j.%%k.%%l.%%m) else (set v=%%i.%%j.%%k.%%l))
	if defined _DbgOut ( echo. %time% ###getWinVer OS: %_OSVER1% %_OSVER2% %_OSVER3% %_OSVER4% Version %v% )
	:: echo Windows Version: %v%
	:: 10.0 - Windows 10		10240 RTM, 10586 TH2 v1511, 14393 RS1 v1607, 15063 RS2 v1703, 16299 RS3 1709, 17134 RS4 1803, 17692 RS5 1809
	::  6.3 - Windows 8.1 and Windows Server 2012R2 9600
	::  6.2 - Windows 8			9200
	::  6.1 - Windows 7			7601
	::  6.0 - Windows Vista		6002
	::  5.2 - Windows XP x64	2600
	::  5.1 - Windows XP		2600
	::  5.0 - Windows 2000		2195
	::  4.10 -Windows 98
	@exit /b
	@goto :eof

:getDateTime - UTILITY to get current Date and Time on localized OS
	For /f "skip=1 tokens=1-2 delims=. " %%a in ('wmic os get LocalDateTime') do (set _CurDateTime=%%a&goto :nextline)
	:nextline
	For /f "tokens=1-2 delims=/: " %%a in ("%TIME%") do (if %%a LSS 10 (set _CurTime=0%%a%%b) else (set _CurTime=%%a%%b))
	set _CurDateTime=%_CurDateTime:~0,8%_%_CurDateTime:~8,6%
	@goto :eof

:getLocale - UTILITY to get System locale
	echo . get System locale
	FOR /F "delims==" %%G IN ('systeminfo.exe') Do  (
		set input=%%G
		for /l %%a in (1,1,100) do @if "!input:~-1!"==" " set input=!input:~0,-1!
		IF "!input:~0,13!"=="System Locale" (
			set answer=!input:~15!
			set answer=!answer: =!
			set VERBOSE_SYSTEM_LOCALE=!answer:*;=!
			call set SYSTEM_LOCALE_WITH_SEMICOLON=%%answer:!VERBOSE_SYSTEM_LOCALE!=%%
			set SYSTEM_LOCALE=!SYSTEM_LOCALE_WITH_SEMICOLON:~0,-1!
			@rem echo locale: !SYSTEM_LOCALE!
	   )
	)
	@goto :eof

:SleepX - UTILITY to wait/sleep x seconds
	timeout /t %1 >NUL
	::@:: set "sleep1=PING -n 2 127.0.0.1 >NUL 2>&1 || PING -n 2 ::1 >NUL 2>&1"
	@exit /b 0
	@goto :eof


:usage
@echo.
@echo. Usage example: %~n0 - mini DCT +ext batch script file  (krasimir.kumanov@gmail.com)
@echo.                %~n0 all - run all data collection commamds
@echo.                %~n0 noBlg - skip Experion Performance counters (*.blg) collection
@echo.                %~n0 noCabZip - You can use param noCabZip to suppress data compresssion at end stage
@echo.
@echo. mDCT updates on: https://github.com/kumanov/mDCT
@echo. -^> see '%~n0 /help' for more detailed help info
@echo. -^> Looking for help on specific keywords^? Try: mDCT /help ^|findstr /i /c:noblg
@goto :eof

:InitLog [LogFileName]
	@if not exist %~1 (
		@echo.%date% %time% . INITIALIZE file %~1 by %USERNAME% on %COMPUTERNAME% in Domain %USERDOMAIN% > %~1
		@echo.mDCT [%_ScriptVersion%] 'krasimir.kumanov@gmail.com' >> %~1
		@echo.>> %~1
	)
	@goto :eof

:logitem  - UTILITY to write a message to the log file (no indent) and screen
	@echo %date% %time% : %* >> !_LogFile!
	@echo %time% : %*
	@goto :eof

:logOnlyItem  - UTILITY to write a message to the log file (no indent)
	@echo %date% %time% : %* >> !_LogFile!
	@goto :eof

:logNoTimeItem  - UTILITY to write a message to the log file (no indent)
	@echo. %* >> !_LogFile!
	@goto :eof

:showlogitem  - UTILITY to write a message to the log file (no time indent) and screen
	@echo. %* >> !_LogFile!
	@echo. %*
	@goto :eof

:doCmd  - UTILITY log execution and output of a command to the current log file
	@echo ================================================================================== >> !_LogFile!
	@echo ===== %time% : %* >> !_LogFile!
	@echo ================================================================================== >> !_LogFile!
	%* >> %_LogFile% 2>&1
	@echo. >> !_LogFile!
	call :SleepX 1
	@goto :eof

:doCmdNoLog  - UTILITY log execution of a command to the current log file
	@echo ================================================================================== >>!_LogFile!
	@echo ===== %time% : %* >> !_LogFile!
	@echo ================================================================================== >> !_LogFile!
	%*
	@echo. >> !_LogFile!
	@goto :eof

:LogCmd [filename; command] - UTILITY to log command header and output in filename
	for /f "tokens=1* delims=; " %%a in ("%*") do (
		set _LogFileName=%%a
		@echo ================================================================================== >> !_LogFileName!
		@echo ===== %time% : %%b >> !_LogFileName!
		@echo ================================================================================== >> !_LogFileName!
		%%b >> !_LogFileName! 2>&1
	)
	@echo. >> !_LogFileName!
	call :SleepX 1
	@goto :eof

:mkNewDir
	set _MS_NewDir=%*
	if not exist "%_MS_NewDir%" mkdir "%_MS_NewDir%"
	@goto :eof

:WriteHost [ forground background message]
	for /f "tokens=1,2* delims=; " %%a in ("%*") do (
		call :logOnlyItem  %%c
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -command write-host -foreground %%a -background %%b "%%c"
		)
	@goto :eof

:WriteHostNoLog [ forground background message]
	for /f "tokens=1,2* delims=; " %%a in ("%*") do (
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -command write-host -foreground %%a -background %%b "%%c"
		)
	@goto :eof

:play_attention_sound - play sound  .beep.
	if NOT "%_noSound%" equ "1" (
		if !_OSVER3! GEQ 9600 (rundll32.exe cmdext.dll,MessageBeepStub) else ( echo  )
	)
	@goto :eof

:GetReg
	@echo ================================================================================== >> %_RegFile%
	@echo ===== %time% : REG.EXE %* >> %_RegFile%
	@echo ================================================================================== >> %_RegFile%
	%SYSTEMROOT%\SYSTEM32\REG.EXE %* >> %_RegFile% 2>&1
	@goto :eof

:DoSqlCmd [database query] - run sql query
	@echo ================================================================================== >> %_SqlFile%
	@echo ===== %time% : sqlcmd DB:%~1 Q:%~2 >> %_SqlFile%
	@echo ================================================================================== >> %_SqlFile%
	sqlcmd -E -w 10000 -d "%~1" -Q "%~2" >> %_SqlFile% 2>&1
	@echo. >> %_SqlFile%
	@goto :eof

:DoGetSVC [comment] - UTILITY to dump Service information into log file
	call :logitem .. collecting Services info at %~1
	set _ServicesFile=!_DirWork!\GeneralSystemInfo\serviceslist.txt
	call :InitLog !_ServicesFile!
	call :LogCmd !_ServicesFile! SC.exe query type= all state= all
	call :SleepX 1
	@goto :eof

:DoNltestDomInfo [comment] - UTILITY to dump NLTEST Domain infos into log file
	call :logitem .. collecting NLTEST Domain information at %~1
	set _NltestInfoFile=!_DirWork!\GeneralSystemInfo\NltestDomInfo.txt
	call :InitLog !_NltestInfoFile!
	call :LogCmd !_NltestInfoFile! nltest /dsgetsite
	call :LogCmd !_NltestInfoFile! nltest /dsgetdc: /kdc /force
	call :LogCmd !_NltestInfoFile! nltest /dclist:
	call :LogCmd !_NltestInfoFile! nltest /trusted_domains
	@goto :eof

:mkCab sourceFolder cabFolder cabName	-- make cab file
::										-- sourceFolder [in] - source folder
::										-- cabFolder    [in] - destination, cab file folder
::										-- cabName      [out] - cab file name
SETLOCAL ENABLEDELAYEDEXPANSION
:: change working directory
pushd "%temp%"
:: define variables in .ddf file
>directives.ddf echo ; Makecab Directive File
>>directives.ddf echo ; Created by mDCT script tool
>>directives.ddf echo ; %date% %time%
>>directives.ddf echo .Option Explicit
>>directives.ddf echo .Set DiskDirectoryTemplate="%~2"
>>directives.ddf echo .Set CabinetNameTemplate="%~3"
>>directives.ddf echo .Set MaxDiskSize=0
>>directives.ddf echo .Set CabinetFileCountThreshold=0
>>directives.ddf echo .Set UniqueFiles=OFF
>>directives.ddf echo .Set Cabinet=ON
>>directives.ddf echo .Set Compress=ON
>>directives.ddf echo .Set CompressionType=MSZIP
::>>directives.ddf echo .Set CompressionType=LZX
:: save current ASCII code page
for /f "tokens=2 delims=:" %%i in ('chcp') do set /a _oemcp=%%~ni
:: change code page to ANSI
chcp 1252>nul
:: append all file names of the source folder
set _sourceFolder=%~1
for /f "delims=" %%i in ('dir /a-d /b /s "%~1"') do (
  set _fileName=%%i
  call :MakeRelative _fileName "!_sourceFolder!"
  >>directives.ddf echo "%%i" 	"!_fileName!"
)
:: change back to ASCII
chcp %_oemcp%>NUL
:: call makecab
makecab /f directives.ddf >NUL
:: clean up
del setup.inf
del setup.rpt
del directives.ddf
:: return to the previous working directory
popd
goto :eof


:MakeRelative file base -- makes a file name relative to a base path
::                      -- file [in,out] - variable with file name to be converted, or file name itself for result in stdout
::                      -- base [in,opt] - base path, leave blank for current directory
:$source https://www.dostips.com
SETLOCAL ENABLEDELAYEDEXPANSION
set src=%~1
if defined %1 set src=!%~1!
set bas=%~2
if not defined bas set bas=%cd%
for /f "tokens=*" %%a in ("%src%") do set src=%%~fa
for /f "tokens=*" %%a in ("%bas%") do set bas=%%~fa
set mat=&rem variable to store matching part of the name
set upp=&rem variable to reference a parent
for /f "tokens=*" %%a in ('echo.%bas:\=^&echo.%') do (
    set sub=!sub!%%a\
    call set tmp=%%src:!sub!=%%
    if "!tmp!" NEQ "!src!" (set mat=!sub!)ELSE (set upp=!upp!..\)
)
set src=%upp%!src:%mat%=!
( ENDLOCAL & REM RETURN VALUES
    IF defined %1 (	SET %~1=%src%) ELSE ECHO.%src%
)
exit /b

:myFunctionName    -- function description here
::                 -- %~1 [in,out,opt]: argument description here
SETLOCAL
REM.--function body here
set LocalVar1=...
set LocalVar2=...
(ENDLOCAL & REM -- RETURN VALUES
	IF "%~1" NEQ "" SET %~1=%LocalVar1%
	IF "%~2" NEQ "" SET %~2=%LocalVar2%
)
exit /b

:doit
	%* >nul
	if errorlevel 1 ( call :logitem failed: %* )
	@goto :eof


::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:mDCT_Begin section ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
call :WriteHostNoLog blue black *** %_ScriptVersion% Dont click inside the script window while processing as it will cause the script to pause. ***

:: init working dir
call :mkNewDir !_DirWork!
:: init LogFile
if not defined _LogFile set _LogFile=!_DirWork!\mDCTlog.txt
call :InitLog !_LogFile!
if defined _GetLocale (
	call :getLocale
	set _locale=!SYSTEM_LOCALE!)

call :logOnlyItem  mDCT (krasimir.kumanov@gmail.com) -%_ScriptVersion% start invocation: '%_DirScript%%~n0 %*'
call :logNoTimeItem  Windows version:  !v! Minor: !_OSVER4!
call :showlogitem   ScriptVersion: %~n0 %_ScriptVersion% - DateTime: !_CurDateTime! Locale: !_locale! PSversion: %_PSVer%

:::::::::: debug call/goto :::::::::::::::::::::::
::::::::::::::::::goto :CrashDumps

:: change priority to idle - this & all child commands
call :logitem change priority to IDLE
wmic process where name="cmd.exe" CALL setpriority "idle"  >NUL 2>&1

:: GeneralSystemInfo folder
call :mkNewDir  !_DirWork!\GeneralSystemInfo

:whoami
call :logitem whoami - currently logged on user
call :LogCmd !_DirWork!\GeneralSystemInfo\whoami.txt whoami /all

:EnvVariables
call :logitem Windows Environment Variables
call :LogCmd !_DirWork!\GeneralSystemInfo\EnvVariables.txt set

:: SIDs
:: skip it - take too long time on some computers !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
:: https://www.maketecheasier.com/find-user-security-identifier-windows/
::@::call :SleepX 1
::@::echo --Users and Groups SID
::@::wmic useraccount get domain,name,sid > !_DirWork!\GeneralSystemInfo\UsersGroupsSID.txt
::@::wmic group get domain,name,sid >> !_DirWork!\GeneralSystemInfo\UsersGroupsSID.txt

:MSInfo32 /report
call :logitem MSInfo32 export
call :doCmd msinfo32 /report !_DirWork!\GeneralSystemInfo\MSInfo32.txt

:hosts file
call :logitem Windows hosts file copy
call :doCmd copy /y %windir%\System32\drivers\etc\hosts !_DirWork!\GeneralSystemInfo\

:ipconfig.output
call :logitem ipconfig output
call :LogCmd !_DirWork!\GeneralSystemInfo\ipconfig.txt ipconfig /all

:timezone
call :logitem get time zone information
call :doCmd wmic /output:!_DirWork!\GeneralSystemInfo\timezone.output.txt timezone get Bias, Description, StandardName

:: export Windows Events
call :logitem export Windows Events
call :doCmd wevtutil epl Application !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_Application.evtx /overwrite:true
call :doCmd wevtutil epl FTE !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_FTE.evtx /overwrite:true
call :doCmd wevtutil epl HwSnmp !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_HwSnmp.evtx /overwrite:true
call :doCmd wevtutil epl HwSysEvt !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_HwSysEvt.evtx /overwrite:true
call :doCmd wevtutil epl Security !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_Security.evtx /overwrite:true
call :doCmd wevtutil epl System !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_System.evtx /overwrite:true

:: get Experion PKS Product Version file
call :logitem get Experion PKS Product Version file
call :doCmd copy /y "%HwInstallPath%\Experion PKS\ProductVersion.txt" !_DirWork!\GeneralSystemInfo\

:: query services
call :logitem query services
call :DoGetSVC %time%
call :DoNltestDomInfo %time%


:schtasks
call :logitem scheduled task - query
schtasks /query /xml ONE >!_DirWork!\GeneralSystemInfo\scheduled_tasks.xml

:gpresult
call :mkNewDir  !_DirWork!\GeneralSystemInfo
set _GPresultFile=!_DirWork!\GeneralSystemInfo\GPresult.htm
call :logitem collecting gpresult output
call :doCmd gpresult /h !_GPresultFile! /f

:powercfg
call :logitem get power configuration settings - powercfg
call :LogCmd !_DirWork!\GeneralSystemInfo\powercfg.txt powercfg -Q
:: reg query power settings
call :logitem reg query power settings
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile=!_DirWork!\GeneralSystemInfo\powercfg.txt
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power"  /s /t reg_dword
::  fast reboot - "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled


:SystemInfo
call :logitem collecting SystemInfo
call :LogCmd !_DirWork!\GeneralSystemInfo\SystemInfo.txt systeminfo.exe

:WmicQfeList
call :logitem collecting Quick Fix Engineering information (Hotfixes)
call :doCmd  wmic /output:!_DirWork!\GeneralSystemInfo\Hotfixes.txt qfe list

:WindowsUpdate.log
call :logitem WindowsUpdate.log
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :doCmd copy /y "%windir%\WindowsUpdate.log" "!_DirWork!\GeneralSystemInfo\"
if exist %windir%\Logs\WindowsUpdate (
	call :logitem .. get Windows Update ETL Logs
	call :mkNewDir  !_DirWork!\GeneralSystemInfo\WindowsUpdateEtlLogs
	call :doCmd copy /y "%windir%\Logs\WindowsUpdate\*.etl" "!_DirWork!\GeneralSystemInfo\WindowsUpdateEtlLogs\"
)

:Honeywell_MsPatches
if exist "%windir%\Honeywell_MsPatches.txt" (
	call :logitem get Honeywell_MsPatches.txt
	call :doCmd copy /y "%windir%\Honeywell_MsPatches.txt" "!_DirWork!\GeneralSystemInfo\"
)

:WmicProductList
::@:: too slow and this information is available in DCT
::@:: call :logitem MS Installation package task management - get name, version
::@:: call :doCmd  wmic /output:!_DirWork!\GeneralSystemInfo\InstallList.txt product get Description,Version,InstallDate

:WmiRootSecurityDescriptor
call :logitem Wmi Root Security Descriptor
call :doCmd  wmic /output:!_DirWork!\GeneralSystemInfo\WmiRootSecurityDescriptor.txt /namespace:\\root path __systemsecurity call GetSecurityDescriptor


:exportExperionRegistrySettings
call :logitem export Experion registry settings
call :mkNewDir  !_DirWork!\RegistryInfo
call :doCmd %windir%\SysWOW64\reg.exe EXPORT HKEY_CURRENT_USER\Software\Honeywell !_DirWork!\RegistryInfo\HKEY_CURRENT_USER_Software_Honeywell.txt
call :doCmd %windir%\SysWOW64\reg.exe EXPORT HKEY_LOCAL_MACHINE\SOFTWARE\Honeywell !_DirWork!\RegistryInfo\HKEY_LOCAL_MACHINE_Software_Honeywell.txt
call :doCmd %windir%\SysWOW64\reg.exe EXPORT HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall !_DirWork!\RegistryInfo\HKEY_LOCAL_MACHINE_Software_Microsoft_Uninstall.txt
call :doCmd REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\McAfee\SystemCore\VSCore\On Access Scanner" !_DirWork!\RegistryInfo\HKLM_McAfee_OnAccessScanner.txt

:getPerformanceLogs
if /i "%_NoBlg%" EQU "1" goto :NoBlg
call :logitem get Experion Performance Logs
  ::@::xcopy /s/e/i/q/y/H "%HwProgramData%\Experion PKS\Perfmon\*.blg" "!_DirWork!\Perfmon Logs\"  1>NUL
if not defined HWPERFLOGPATH set HWPERFLOGPATH=%HwProgramData%\Experion PKS\Perfmon
if Not Exist "%HWPERFLOGPATH%" (
	call :logOnlyItem perfmon folder Not Exist "%HWPERFLOGPATH%"
	goto :NoBlg
)
call :mkNewDir !_DirWork!\Perfmon Logs
PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci $env:HWPERFLOGPATH -filt *.blg | where{$_.LastWriteTime -gt (get-date).AddDays(-10)} | foreach{copy $_.fullName -dest '!_DirWork!\Perfmon Logs\'; sleep 1} }}
if defined _DbgOut ( echo. .. ** ERRORLEVEL: %errorlevel% - 'at Copy blg files with PowerShell'. )
if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy blg files with PowerShell' failed.)
:NoBlg

:FteLogs
call :logitem get FTE logs
call :doCmd xcopy /s/e/i/q/y/H "%HwProgramData%\ProductConfig\FTE\*.log" "!_DirWork!\FTELogs\"


:HMIWebLog
call :logitem get HMIWeb log files
call :doCmd xcopy /i/q/y/H "%HwProgramData%\HMIWebLog\*.txt" "!_DirWork!\Station-logs\"
call :doCmd xcopy /i/q/y/H "%HwProgramData%\HMIWebLog\Archived Logfiles\*.txt" "!_DirWork!\Station-logs\Rollover-logs\"

:tasklist_svc
call :logitem task list /services
call :mkNewDir  !_DirWork!\ServerDataDirectory
call :LogCmd !_DirWork!\ServerDataDirectory\TaskList.txt tasklist /fo csv /svc



:: Network :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
call :mkNewDir !_DirWork!\Network

:netstat connections
call :logitem netstat connections
call :LogCmd !_DirWork!\Network\netstat-nato.txt netstat -nato

:ipconfig_displaydns
call :logitem ipconfig /displaydns
call :LogCmd !_DirWork!\Network\ipconfig.displaydns.txt ipconfig /displaydns

:NetSh_ConfigAndStats
call :logitem netsh config/stats
call :LogCmd !_DirWork!\Network\netsh.ipstats.txt      netsh interface ipv4 show ipstats
call :LogCmd !_DirWork!\Network\netsh.tcpstats.txt     netsh interface ipv4 show tcpstats
call :LogCmd !_DirWork!\Network\netsh.dynamicport.txt  netsh int ipv4 show dynamicport tcp
call :LogCmd !_DirWork!\Network\netsh.global.txt       netsh int tcp show global
call :LogCmd !_DirWork!\Network\netsh.ipv4.offload.txt netsh int ipv4 show offload

:nslookup
call :logitem nslookup
call :SleepX 1
:: NS LookUp - Forward
echo ======================================== > !_DirWork!\Network\nslookup.txt
echo %date% %time% >> !_DirWork!\Network\nslookup.txt
echo cmd: nslookup %computername% >> !_DirWork!\Network\nslookup.txt
nslookup %computername% >> !_DirWork!\Network\nslookup.txt  2>>&1
:: NS LookUp - Reverse
set ip_address_string="IPv4 Address"
for /f "usebackq tokens=2 delims=:" %%f in (`ipconfig ^| findstr /c:%ip_address_string%`) do (
	echo ======================================== >> !_DirWork!\Network\nslookup.txt
	echo %date% %time% >> !_DirWork!\Network\nslookup.txt
    echo Your IP Address is: %%f  >> !_DirWork!\Network\nslookup.txt
	echo cmd: nslookup %%f >> !_DirWork!\Network\nslookup.txt
	nslookup %%f >> !_DirWork!\Network\nslookup.txt  2>>&1
)
echo ======================================== >> !_DirWork!\Network\nslookup.txt
echo %date% %time% - done>> !_DirWork!\Network\nslookup.txt

:route_arp
call :logitem route / arp
call :LogCmd !_DirWork!\Network\arp.txt  arp -a -v
call :LogCmd !_DirWork!\Network\route.print.txt route print

:nbtstat
call :logitem nbtstat-n
call :InitLog !_DirWork!\Network\nbtstat.txt
call :LogCmd !_DirWork!\Network\nbtstat.txt  nbtstat -n

:advfirewall
call :logitem firewall rules
call :LogCmd !_DirWork!\Network\firewall_rules.txt netsh advfirewall firewall show rule name=all

:net.commands
call :logitem net commands
call :LogCmd !_DirWork!\Network\netcmd.txt NET CONFIG SERVER
call :LogCmd !_DirWork!\Network\netcmd.txt NET SESSION
call :LogCmd !_DirWork!\Network\netcmd.txt NET SHARE
call :LogCmd !_DirWork!\Network\netcmd.txt NET USER
call :LogCmd !_DirWork!\Network\netcmd.txt NET USE
call :LogCmd !_DirWork!\Network\netcmd.txt NET ACCOUNTS
call :LogCmd !_DirWork!\Network\netcmd.txt NET CONFIG WKSTA
call :LogCmd !_DirWork!\Network\netcmd.txt NET STATISTICS Workstation
call :LogCmd !_DirWork!\Network\netcmd.txt NET STATISTICS SERVER

:TcpIpParameters
set _RegFile=!_DirWork!\Network\TcpIpParameters.txt
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp\Parameters" /v ArpRetryCount
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp\Parameters" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\Tcpip6\Parameters" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\tcpipreg" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\iphlpsvc" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}" /s


:driverquery
call :logitem query drivers information
call :LogCmd !_DirWork!\GeneralSystemInfo\driverquery.output.csv driverquery /fo csv /v


::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::  Experion console station & server node
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:setpar /active
where setpar >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion active log paranoids
	call :mkNewDir !_DirWork!\SloggerLogs
	call :LogCmd !_DirWork!\SloggerLogs\setpar.active.txt setpar /active
)

:SloggerLogs
:: R5xx
if exist "%HwProgramData%\Experion PKS\logfiles\logServer.txt" (
	call :logitem get Experion log files
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\logfiles\log*.txt" "!_DirWork!\SloggerLogs\"
	:: copy server log archives
	if exist "%HwProgramData%\Experion PKS\logfiles\00-Server\" (
		call :mkNewDir !_DirWork!\SloggerLogs\Archives
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '%HwProgramData%\Experion PKS\logfiles\00-Server\' -filt logServerY*.txt | where{$_.LastWriteTime -gt (get-date).AddDays(-14)} | foreach{copy $_.fullName -dest '!_DirWork!\SloggerLogs\Archives\'; sleep 1} }}
		if defined _DbgOut ( echo. .. ** ERRORLEVEL: %errorlevel% - 'at Copy server log archived files with PowerShell'. )
		if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy server log archived files with PowerShell' failed.)
	)
)
:: R4xx
if exist "%HwProgramData%\Experion PKS\Server\data\log.txt" (
	call :logitem get Experion log files
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\*log*.txt" "!_DirWork!\SloggerLogs\"
)

:CSLog - Configuration Studio.log
if exist "%HWCONFIGSTUDIOLOGPATH%\Configuration Studio.log" (
	call :mkNewDir !_DirWork!\Configuration Studio
	call :logitem get Configuration Studio log file
	call :doCmd copy /y "%HWCONFIGSTUDIOLOGPATH%\Configuration Studio.log*" "!_DirWork!\Configuration Studio\"
)

if exist "%HwProgramData%\Experion PKS\Server\data\Report\" (
	call :logitem File Replication logs
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :mkNewDir !_DirWork!\ServerDataDirectory\File-Replication
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\Report\filrep*.*" "!_DirWork!\ServerDataDirectory\File-Replication\" /s
)

:almdmp
where almdmp >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion alarm/event dump
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :LogCmd !_DirWork!\ServerDataDirectory\almdmp.output.txt almdmp A 32000 S
	call :LogCmd !_DirWork!\ServerDataDirectory\eventdmp.output.txt almdmp E 32000 S
)

:shheap
where shheap >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion shheap output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\shheap.1.output.txt shheap 1 struct
    call :logCmd !_DirWork!\ServerDataDirectory\shheap.1.dump.output.txt shheap 1 dump
    call :logCmd !_DirWork!\ServerDataDirectory\shheap.4.struct.output.txt shheap 4 struct
)

:lisscn
where lisscn >NUL 2>&1
if %errorlevel%==0 (
	call :logitem lisscn output
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd lisscn -all_ref -OUT !_DirWork!\ServerDataDirectory\lisscn_all.txt
	call :doCmd lisscn -OUT !_DirWork!\ServerDataDirectory\lisscn.txt
)


:bckbld
where bckbld >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion point back build
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd bckbld -out !_DirWork!\ServerDataDirectory\back_build.output.txt
)

:hdwbckbld
where hdwbckbld >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion hardware back build
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd hdwbckbld -out !_DirWork!\ServerDataDirectory\hardware_back_build.output.txt
)

:hstdiag
where hstdiag >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion history diagnostic
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\hstdiag.output.txt hstdiag
)

:embckbuilder
where embckbuilder >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion embckbuilder output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.alarmgroup.output.txt  -ALARMGROUP
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.asset.output.txt  -ASSET
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.network.output.txt  -NETWORK
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.system.output.txt  -SYSTEM
)

:fildmp
where fildmp >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion System Flags Table output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd fildmp -DUMP -FILE !_DirWork!\ServerDataDirectory\sysflg.output.txt -FILENUM 8 -RECORDS 1 -FORMAT HEX
	call :logitem Experion Area Asignmnt Table output
    call :doCmd fildmp -DUMP -FILE !_DirWork!\ServerDataDirectory\areaasignmnt.output.txt -FILENUM 7 -RECORDS 1,1001 -FORMAT HEX
)

:OPCIntegrator
if exist "%HwProgramData%\Experion PKS\Server\data\OPCIntegrator\" (
	call :logitem Experion OPC Integrator
	call :mkNewDir !_DirWork!\OPCIntegrator
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\OPCIntegrator\*.tsv" "!_DirWork!\OPCIntegrator\"
)

:listag
where listag >NUL 2>&1
if %errorlevel%==0 (
	call :logitem listag output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\listag.output.txt listag -ALL
)

:filfrag
where filfrag >NUL 2>&1
if %errorlevel%==0 (
	call :logitem filfrag output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\filfrag.output.txt filfrag
)

:system.build
if exist "%HwProgramData%\Experion PKS\Server\data\system.build" (
	call :logitem copy system.build file
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\system.build" !_DirWork!\ServerDataDirectory\
)

:BadFiles
if exist "%HwProgramData%\Experion PKS\Server\data\" (
	call :logitem collect bad files .\server\data\*.bad
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\*.bad" !_DirWork!\ServerDataDirectory\
)

:TPNServer.log
if exist "%HwProgramData%\TPNServer\TPNServer.log" (
	call :logitem copy TPNServer.log file
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\TPNServer\TPNServer.log" !_DirWork!\ServerDataDirectory\
)

:mapping.tps.xml
if exist "%HwProgramData%\Experion PKS\Server\data\mapping\tps.xml" (
	call :logitem copy .\mapping\tps.xml
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\mapping\tps.xml" !_DirWork!\ServerDataDirectory\mapping.tps.xml
)

:dsasublist
where dsasublist >NUL 2>&1
if %errorlevel%==0 (
	call :logitem dsasublist
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\dsasublist.txt dsasublist
)

:winsxs
	call :logitem list C:\Windows\winsxs\
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\winsxs.txt dir %windir%\winsxs

:liclist
where liclist >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion license list - liclist
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\liclist.output.txt
    call :logCmd !_DirWork!\ServerRunDirectory\liclist.output.txt liclist
)

:hwlictool
where hwlictool >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion license list - hwlictool
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\hwlictool.output.txt
    call :logCmd !_DirWork!\ServerRunDirectory\hwlictool.output.txt hwlictool export -format:xml
)

:usrlrn
where usrlrn >NUL 2>&1
if %errorlevel%==0 (
	call :logitem usrlrn usrlrn -p -a
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\usrlrn.txt
    call :logCmd !_DirWork!\ServerRunDirectory\usrlrn.txt usrlrn -p -a
)

:WhatOutput
where what >NUL 2>&1
if %errorlevel%==0 (
	call :logitem What - Getting Experion exe/dll and source file information
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\what.output.txt
	for /r "%HwInstallPath%\Experion PKS\Server\run" %%a in (*.exe *.dll) do what "%%a" >>!_DirWork!\ServerRunDirectory\what.output.txt
)

:CrashDumps
	call :logitem create crash dumps list
	call :mkNewDir !_DirWork!\CrashDumps
	call :logCmd !_DirWork!\CrashDumps\CrashDumpsList.txt dir %windir%\memory.dmp
	call :logCmd !_DirWork!\CrashDumps\CrashDumpsList.txt dir /o:-d %windir%\Minidump
	call :logCmd !_DirWork!\CrashDumps\CrashDumpsList.txt dir /o:-d "%HwProgramData%\Experion PKS\CrashDump"
	call :logCmd !_DirWork!\CrashDumps\CrashDumpsList.txt dir /o:-d "%HwProgramData%\HMIWebLog\DumpFiles"
	call :logCmd !_DirWork!\CrashDumps\CrashDumpsList.txt dir /o:-d "%HwProgramData%\Experion PKS\server\data\*.dmp"
	call :logCmd !_DirWork!\CrashDumps\CrashDumpsList.txt dir  /o-d /s c:\users\*.dmp
	
	call :logitem crash control registry settings
	set _RegFile=!_DirWork!\CrashDumps\RegCrashControl.txt
	call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\CrashControl" /s
	call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\AeDebug" /s
	call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug" /s
	call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\Windows Error Reporting\LocalDumps" /s
	call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps" /s
	call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options" /s

	call :logitem Recover OS settings
	call :DoCmd  wmic /output:!_DirWork!\CrashDumps\recoveros.txt RECOVEROS

:MSSQL
where sqlcmd >NUL 2>&1
if %errorlevel% NEQ 0 (
	call :logOnlyItem no sqlcmd utility - skip sql queries
	goto :NoMSSQL
)
	call :logitem *** query MS SQL ***
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem select count^(*^) from NON_ERDB_POINTS_PARAMS
	set _SqlFile=!_DirWork!\MSSQL-Logs\erdb.p2p.txt
	call :InitLog !_SqlFile!
	call :DoSqlCmd ps_erdb "select count(*) from NON_ERDB_POINTS_PARAMS"
	sqlcmd -E -w 10000 -d ps_erdb -Q "select s.StrategyName as NonCEEStrategy, containingStrat.StrategyName +'.'+ strat_cont.StrategyName as ReferencedBlock, CASE WHEN s.StrategyID  & 0x80000000 = 0 THEN 'Project' ELSE 'Monitoring' END 'Avatar' from STRATEGY S inner join NON_ERDB_POINTS_PARAMS N on n.StrategyID = S.StrategyID and n.ReferenceCount > 0 inner join CONNECTION Conn on conn.PassiveParamID = N.ParamID and conn.passivecontrolid = n.strategyid  INNER JOIN STRATEGY strat_cont ON strat_cont.StrategyID = conn.ActiveControlID INNER JOIN RELATIONSHIP rel ON rel.TargetID = strat_cont.StrategyID AND rel.RelationshipID = 3 INNER JOIN STRATEGY containingStrat ON rel.SourceID = containingStrat.StrategyID " >>!_SqlFile!

	:: SQL Loggins
	set _SqlFile=!_DirWork!\MSSQL-Logs\SqlLogins.txt
	call :InitLog !_SqlFile!
	call :logitem EXEC sp_helplogins
	call :DoSqlCmd master "EXEC sp_helplogins"
	call :logitem SELECT name, type_desc, is_disabled FROM sys.server_principals
	call :DoSqlCmd master "SELECT name, type_desc, is_disabled FROM sys.server_principals"
	call :logitem EXEC sp_who2
	call :DoSqlCmd "master" "EXEC sp_who2"

:emsqueries
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem emsevents sql queries
	set _SqlFile=!_DirWork!\MSSQL-Logs\emsqueries.output.txt
	call :InitLog !_SqlFile!
	call :DoSqlCmd EMSEvents "SELECT @@VERSION"
	call :DoSqlCmd EMSEvents "SELECT SERVERPROPERTY('MachineName')"
	call :DoSqlCmd EMSEvents "SELECT @@SERVERNAME"
	call :DoSqlCmd EMSEvents "SELECT * FROM sys.databases"
	call :DoSqlCmd EMSEvents "SELECT * FROM sys.assemblies"
	call :DoSqlCmd EMSEvents "sp_helpdb EMSEvents"
	call :DoSqlCmd EMSEvents "SELECT * FROM sys.sysobjects ORDER BY [name]"
	call :DoSqlCmd EMSEvents "SELECT * FROM dbo.EventConfig"
	call :DoSqlCmd EMSEvents "SELECT * FROM EMSEvents.dbo.KnownArchives"
	call :DoSqlCmd EMSEvents "SELECT MIN(EventID) AS MinEventID, MAX(EventID) AS MaxEventID, COUNT(*) AS EventCount FROM dbo.Events"
	call :DoSqlCmd EMSEvents "SELECT EventSourceID, MIN(EventID) AS MinEventID, MAX(EventID) AS MaxEventID, COUNT(*) AS RestoredEventCount FROM dbo.RestoredEvents GROUP BY EventSourceID"
	call :DoSqlCmd EMSEvents "SELECT TOP 5 * FROM dbo.ems_vw_Events ORDER BY Time DESC"
	call :DoSqlCmd EMSEvents "SELECT TOP 5 * FROM dbo.ems_vw_comments"
	call :DoSqlCmd EMSEvents "SELECT TOP 5 * FROM dbo.EventComments"
	call :DoSqlCmd EMSEvents "SELECT TOP 5 * FROM dbo.ems_vw_OnlineEvents"
	call :DoSqlCmd EMSEvents "SELECT count(*) FROM EMSEvents.dbo.ems_vw_OnlineEvents"
	call :DoSqlCmd EMSEvents "SELECT TOP 100 * FROM EMSEvents.dbo.ems_vw_EMSEvents ORDER BY EventID DESC"
	call :DoSqlCmd EMSEvents "SELECT MIN(EventID) AS MinEventID, MAX(EventID) AS MaxEventID, EMSEvents.dbo.UTCFILETIMEToDateTime(MIN(LocalTime)) AS MinLocalTime, EMSEvents.dbo.UTCFILETIMEToDateTime(MAX(LocalTime)) AS MaxLocalTime, COUNT(*) AS EventCount FROM EMSEvents.dbo.ems_vw_EMSEvents"
	call :DoSqlCmd EMSEvents "SELECT MIN(EventID) AS MinEventID, MAX(EventID) AS MaxEventID, EMSEvents.dbo.UTCFILETIMEToDateTime(MIN(LocalTime)) AS MinLocalTime, EMSEvents.dbo.UTCFILETIMEToDateTime(MAX(LocalTime)) AS MaxLocalTime, COUNT(*) AS EventCount FROM EMSEvents.dbo.ems_vw_OnlineEvents"
	call :DoSqlCmd EMSEvents "sp_helptext ems_vw_EMSEvents"
	call :DoSqlCmd EMSEvents "sp_helptext ems_vw_OnlineEvents"
	call :DoSqlCmd EMSEvents "sp_helptext ems_vw_Events"
	call :DoSqlCmd EMSEvents "sp_help 'Events'"
	call :DoSqlCmd EMSEvents "sp_help 'tempEventTable'"
	call :DoSqlCmd EMSEvents "sp_help 'RestoredEvents'"
	call :DoSqlCmd EMSEvents "sp_help 'ems_vw_Events EMSEvents'"
	call :DoSqlCmd EMSEvents "sp_help 'ems_vw_RestoredEvents EMSEvents'"
	call :DoSqlCmd EMSEvents "SELECT * FROM msdb.dbo.sysjobs"
	call :DoSqlCmd EMSEvents "SELECT * FROM msdb.dbo.sysjobactivity"
	call :DoSqlCmd EMSEvents "SELECT * FROM msdb.dbo.sysjobhistory"
	call :DoSqlCmd EMSEvents "SELECT * FROM EMSEvents.dbo.EventDelivery"

:CheckSQLDBLogs
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem Check SQL DB Logs
	set _SqlFile=!_DirWork!\MSSQL-Logs\CheckSQLDBLogs.txt
	call :InitLog !_SqlFile!
	call :DoSqlCmd master "select [name] AS 'Database Name', DATABASEPROPERTYEX([name],'recovery') AS 'Recovery Model' from master.dbo.SysDatabases"
	call :DoSqlCmd master "DBCC SQLPERF(LOGSPACE)"
	call :DoSqlCmd master "select name AS 'Database name',log_reuse_wait_desc AS 'Log  Reuse' from sys.databases"

:NoMSSQL

if defined _noCabZip goto :off_nocab

:COMPRESS
@rem check for makecab.exe
for %%i in (makecab.exe) do (set exe=%%~$PATH:i)
if "!exe!" equ "" (
    @echo.
    @echo.WARNING: makecab.exe not found. Proceeding as if 'nocab' was specified.
    goto :off_nocab )

@rem construct cab directive file
set cabName=%_Comp_Time%.cab

:: ## remove space and comma in cab name
set cabName=%cabName:,=-%
set cabName=%cabName: =_%
if !_PSVer! LEQ 2 (
	call :logitem .. Your PowerShell version !_PSVer! is outdated, using MakeCab
	call :logitem .. MakeCab: compressing data %_DirRepro% - please be patient...

	echo --- !time! ---
	call :MkCab "!_DirWork!" "%_DirScript%" "%cabName%
	if !errorlevel! neq 0 (
		@echo.ERROR: failed to compress trace files
		@echo.  * MakeCab can only compress data less than 2GB, you can use '%~n0 off nocab' to avoid compressing phase.
		goto :off_nocab
		@exit /b 1)
	echo --- !time! ---

) else (
	set cabName=%cabName:cab=zip%
	:: next line only works for PS version greater v2
	call :logitem .. PowerShell: compressing data %_DirWork% - please be patient...
	PowerShell.exe -NonInteractive  -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { try { $ErrorActionPreference = 'continue'; Add-Type -Assembly 'System.IO.Compression.FileSystem'; [System.IO.Compression.ZipFile]::CreateFromDirectory('%_DirWork%', '%_DirWork%\..\!cabName!'); Exit 0 } catch { Write-host -ForegroundColor red 'Compress Failed'; Throw $error[0].Exception.Message; Exit 23 } }}"
		if defined _DbgOut ( echo. .. ** ERRORLEVEL: %errorlevel% - 'at Compress with PowerShell'. )
		if "%errorlevel%" neq "0" (
			call :logItem %time% .. ERROR: %errorlevel% - 'Compress with PowerShell' failed.
			set _DoCleanup=0
			echo. _DoCleanup: !_DoCleanup!
			goto :off_nocab
			@exit /b 1)
)
	
@echo.

call :showlogitem  *** %time% : %~n0 diagnostic files are in:
call :showlogitem  ***  %_DirScript%!cabName!
call :WriteHost white black *** [Note] Please upload data %_DirScript%!cabName! onto given GTAC workspace.
call :play_attention_sound

goto :end


@rem nocab or cab failure: print working directory location
:off_nocab
@echo.
@echo.
@echo. Diagnostic data have NOT been compressed!
@echo. Data located in: %_DirWork%
call :WriteHost white black *** Please compress all files in %_DirWork% and upload zip file to GTAC ftp site
call :play_attention_sound

:: Info:
:: 	- v1.xx see Revision History in SCN file
::  - v1.05 Add NltestDomInfo; mkCab


:: ToDo:
:: - [] Delete source files, if compressed

:: - [] lisscn - change
::    lisscn -chn n -all_ref > lisscn_all.txt
::    lisscn -chn n > lisscn.txt

:: - [] McAfee - check reg key before
::    reg query "HKLM\SOFTWARE\Wow6432Node\McAfee\SystemCore\VSCore\On Access Scanner" >NUL 2>&1
::    if %ERRORLEVEL% EQU 0 goto :noMcAfeeScanner

:: - [] Collect DCOM & TCP settings
::	if /i "%~1" equ "DCOM" 	(
::		set _RegFile=!_PrefixT!Reg_%~1_%mode%.txt
::		call :InitLog !_RegFile!
::		call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\LSA" /s
::		call :GetReg QUERY "HKLM\Software\Policies\Microsoft\Windows NT\DCOM" /s
::		call :GetReg QUERY "HKLM\Software\Microsoft\COM3" /s
::		call :GetReg QUERY "HKLM\Software\Microsoft\Rpc" /s
::		call :GetReg QUERY "HKLM\Software\Microsoft\OLE" /s
::	)
::	if /i "%~1" equ "Tcp" 	(
::		call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp\Parameters" /v ArpRetryCount
::		call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp\Parameters" /s
::		call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\Tcpip6\Parameters" /s
::		call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\tcpipreg" /s
::		call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\iphlpsvc" /s
::		call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}" /s
::	)

:: - [x] get WindowsUpdate.log
::		"C:\Windows\WindowsUpdate.log"
::			Windows Update logs are now generated using ETW (Event Tracing for Windows).
::			Please run the Get-WindowsUpdateLog PowerShell command to convert ETW traces into a readable WindowsUpdate.log.
::			
::			For more information, please visit http://go.microsoft.com/fwlink/?LinkId=518345

:: - [x] Reg query power settings -> Turn off fast startup (/v HiberbootEnabled)
::		reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /s
:: - [x] Add cabzip input parameter
:: - [x] Add Experion license output infiormation (hwlictool.output.txt & liclist.output.txt)
::      hwlictool list -format:xml >"c:\temp\test\hwlictool.output.txt" 2>&1
:: - [x] Add usrlrn output
:: - [x] xcopy command fixes
:: - [x] fix mkCab


:END
call :logitem done.
endlocal
@echo.&goto:eof

