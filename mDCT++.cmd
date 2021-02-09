:: Filename: mDCT++.cmd - mini Data Collection Tool script ++ extensions
@if "%_ECHO%" == "" ECHO OFF
setlocal enableDelayedExpansion


set _ScriptVersion=v1.42

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
:: mini DCT +ext batch script file  (krasimir.kumanov@gmail.com)
::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::
::  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
::  IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

@if defined _Debug ( echo [#Debug#] !time!: Start of mDCT++)
:: handle /?
if "%~1"=="/?" (
	call :usage
	exit /b 0
)

::set _DirScript=%~dp0
REM change to use current dirctory,not he script directory
set _DirScript=%CD%\
call :preRequisites
if /i "%errorlevel%" NEQ "0" (@goto :eof)
call :Initialize %*
if /i "!_Usage!" EQU "1" (@goto :eof)
call :CollectDctData
call :getPerformanceLogs
call :CollectAdditionalData
call :compress
:: delete source files, if compressed
if "%errorlevel%"=="0" (
	if exist %_DirScript%!cabName! (
	call :logitem *** data archived - delete working folder '!_DirWork!'
	RD /S /Q "!_DirWork!")
)
echo.done.
:: restore cmd title
title Command Prompt
@echo. & @goto :eof

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:region check preRequisites
:preRequisites
	@echo ..check preRequisites
	:: check current user account permissions - Admin required
	call :check_Permissions
	if "%errorlevel%" neq "0" @goto :eof
	
	:: require no spaces in full path
	if not "%_DirScript%"=="%_DirScript: =%" ( echo.
		call :WriteHostNoLog yellow black  *** Your script execution path '%_DirScript%' contains one or more space characters
		call :WriteHostNoLog yellow black  Please use a different local path without space characters.
		exit /b 1
	) else (
		call :WriteHostNoLog green black Success: no spaces in folder full path
	)
	
	:: user account membership : "Product Administrators" , "Local Engineers" , test
	whoami /all 2>NUL | findstr /irc:"Product Administrators" /c:"Local Engineers" /c:E402276 >NUL 2>&1
	if "%errorlevel%" NEQ "0" (
		call :WriteHostNoLog yellow black  *** User Account Membership `nUser account should be member on `'Product Administrators`' or `'Local Engineers`' `nPlease use user account member on the above groups
		::call :WriteHostNoLog yellow black  Please use user account member on the above groups
		exit /b 1
	) else (
		call :WriteHostNoLog green black Success: user account group membership
	)

	@goto :eof
:endregion

:region initialize
:initialize
:: initialize variables

:: change the cmd prompt environment to English
chcp 437 >NUL

:: Adding a Window Title
SET _title=%~nx0 - version %_ScriptVersion%
TITLE %_title% & set _title=

@if defined _Debug ( echo [#Debug#] !time!: _DirScript: %_DirScript% )

:: Change Directory to the location of the batch script file (%0)
CD /d "%_DirScript%"
@echo. .. starting '%_DirScript%%~n0 %*'

::@::if defined _Debug ( echo [#Debug#] !time!: Start of mDCT++ ^(%_ScriptVersion% - krasimir.kumanov@gmail.com^))
call :WriteHostNoLog white black %date% %time% : Start of mDCT++ [%_ScriptVersion% - krasimir.kumanov@gmail.com]

:region Configuration parameters ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Settings for quick testing: in CMD use: 'set _DbgOut=1', use SET _Echo=1 to debug this script

:: init variables
	@set _Comp_Time=
	@set _DirWork=
	@set _LogFile=

:: parsed args - off
	@set _Usage=
	@set _noDctData=
	@set _noAddData=
	@set _noCabZip=
	@set _noPerfMon=

:: regional settings ^(choose EN-US, DE-DE^) for localized Perfmon counter names, hardcode to EN-US, or choose _GetLocale=1
	@set _locale=EN-US
	@set _GetLocale=

:: VEP
	@set _VEP=

:: HSCServerType
	@set _HSCServerType=
	(for /f "tokens=2,* delims= " %%h in ('reg query "HKLM\SOFTWARE\Wow6432Node\Honeywell" /v HSCServerType ^| find /i "Server"') do @set _HSCServerType=%%i) 2>NUL
	:: is Server
	@if defined _Debug ( echo [#Debug#] !time!: _HSCServerType=%_HSCServerType%)
	if NOT "%_HSCServerType%"=="%_HSCServerType:Server=%" (
		@set _isServer=1
	) else (
		@set _isServer=
	)

:: EXperion Release
	@set _EPKS_Release=
	call :GetRegValue "HKLM\SOFTWARE\Wow6432Node\Honeywell\Experion PKS Server" Release _EPKS_Release
	@set _EPKS_MajorRelease=
	@if defined _EPKS_Release (
		set _EPKS_Release=%_EPKS_Release:"=%
		set _EPKS_MajorRelease=!_EPKS_Release:~0,3!
	)


:: TPSNodeInstallation
	reg query "HKLM\SOFTWARE\WOW6432Node\Honeywell" /v TPSNodeInstallation >NUL 2>&1
	if errorlevel 1 (
		@set _isTPS=
	) else (
		@set _isTPS=1
	)


:endregion Configuration parameters ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:region parse args

call:ArgsParse %*

if /i "%_Usage%" EQU "1" ( call :usage
					@exit /b 0 )
:endregion

:: _OSVER* will be set in function getWinVer
call :getWinVer
call :getDateTime
for /f "delims=" %%a in ('PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { $PSVersionTable.PSVersion.Major }}"') do set _PSVer=%%a
:: Win OS 32/64 bit version
Set _xOS=64
if "%PROCESSOR_ARCHITECTURE%" equ "x86" (
	if not defined PROCESSOR_ARCHITECTURE6432 ( set _xOS=32 )
)

set _Comp_Time=%COMPUTERNAME%_!_CurDateTime!
@if defined _Debug ( echo [#Debug#] !time!: _Comp_Time: %_Comp_Time% )
:: set work folder
set _DirWork=%_DirScript%%_Comp_Time%

@if defined _Debug ( echo [#Debug#] !time!: _DirWork: !_DirWork! )

:: init working dir
call :mkNewDir !_DirWork!
:: init LogFile
if not defined _LogFile set _LogFile=!_DirWork!\mDCTlog.txt
@if defined _Debug ( echo [#Debug#] !time!: _LogFile: !_LogFile! )
call :InitLog "!_LogFile!"

:: change priority to idle - this & all child commands
call :logitem change priority to IDLE
wmic process where name="cmd.exe" CALL setpriority "idle"  >NUL 2>&1

call :WriteHostNoLog blue gray *** Dont click inside the script window while processing as it will cause the script to pause. ***

:: VEP detect
wmic path win32_computersystem get Manufacturer /value | findstr /ic:"VMware" >NUL 2>&1
if "%errorlevel%"=="0" (set _VEP=1)

if defined _GetLocale ( call :getLocale _locale )

call :logOnlyItem  mDCT++ (krasimir.kumanov@gmail.com) -%_ScriptVersion% start invocation: '%_DirScript%%~n0 %*'
call :logNoTimeItem  Windows version:  !_v! Minor: !_OSVER4!
call :showlogitem   ScriptVersion: %~n0 %_ScriptVersion% - DateTime: !_CurDateTime! Locale: !_locale! PSversion: %_PSVer%

goto :eof
:endregion

:region functions
:ArgsParse
	:: end parse, if no more args
	IF "%~1"=="" exit /b

	set _KnownArg=0

	IF "%~1"=="-?" (set _Usage=1& exit /b)

	for %%i in (help /help -help -h /h) do (
		if /i "%~1" equ "%%i" (set _Usage=1& exit /b)
	)

	IF /i "%~1"=="noDctData" (set _NoDctData=1
		set _KnownArg=1)
	IF /i "%~1"=="noAddData" (set _NoAddData=1
		set _KnownArg=1)
	IF /i "%~1"=="noPerfMon" (set _noPerfMon=1
		set _KnownArg=1)
	if /i "%~1"=="noCabZip"  (set _noCabZip=1
		set _KnownArg=1)
	if /i "%~1"=="Debug"  (set _Debug=1
		set _KnownArg=1)

	if /i "!_KnownArg!"=="0" (
		@echo.
		call :WriteHostNoLog yellow black  *** Unknown input argument: "%~1"
		set _Usage=1
		exit /b 1)
	
	SHIFT
	GOTO ArgsParse


:check_Permissions
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
	for /f "tokens=2 delims=[]" %%o in ('ver')    do @set _OSVERTEMP=%%o
	for /f "tokens=2" %%o in ('echo %_OSVERTEMP%') do @set _OSVER=%%o
	for /f "tokens=1 delims=." %%o in ('echo %_OSVER%') do @set _OSVER1=%%o
	for /f "tokens=2 delims=." %%o in ('echo %_OSVER%') do @set _OSVER2=%%o
	for /f "tokens=3 delims=." %%o in ('echo %_OSVER%') do @set _OSVER3=%%o
	for /f "tokens=4 delims=." %%o in ('echo %_OSVER%') do @set _OSVER4=%%o
	for /f "tokens=4-8 delims=[.] " %%i in ('ver') do (if %%i==Version (set _v=%%j.%%k.%%l.%%m) else (set _v=%%i.%%j.%%k.%%l))
	@if defined _Debug ( echo [#Debug#] !time!: ###getWinVer OS: %_OSVER1% %_OSVER2% %_OSVER3% %_OSVER4% Version %_v% )
	:: echo Windows Version: %_v%
	:: 10.0 - Windows 10		10240 RTM, 10586 TH2 v1511, 14393 RS1 v1607, 15063 RS2 v1703, 16299 RS3 1709, 17134 RS4 1803, 17692 RS5 1809
	::  6.3 - Windows 8.1 and Windows Server 2012R2 9600
	::  6.2 - Windows 8			9200
	::  6.1 - Windows 7			7601
	::  6.0 - Windows Vista		6002
	::  5.2 - Windows XP x64	2600
	::  5.1 - Windows XP		2600
	::  5.0 - Windows 2000		2195
	::  4.10 -Windows 98
	@goto :eof

:getDateTime - UTILITY to get current Date and Time on localized OS
	For /f "skip=1 tokens=1-2 delims=. " %%a in ('wmic os get LocalDateTime') do (set _CurDateTime=%%a&goto :nextline)
	:nextline
	For /f "tokens=1-2 delims=/: " %%a in ("%TIME%") do (if %%a LSS 10 (set _CurTime=0%%a%%b) else (set _CurTime=%%a%%b))
	set _CurDateTime=!_CurDateTime:~0,8!_!_CurDateTime:~8,6!
	rem fix - remove commas, if any
	set _CurDateTime=!_CurDateTime:,=!
	@goto :eof

:getLocale	-- UTILITY to get System locale
::			-- %~1 [out]: output variable
	SETLOCAL
	@echo . get System locale
	FOR /F "usebackq delims==" %%G IN (`systeminfo.exe 2^>NUL ^| find /i "System Locale"`) Do  (
		set input=%%G
		for /l %%a in (1,1,100) do @if "!input:~-1!"==" " set input=!input:~0,-1!
		IF "!input:~0,13!"=="System Locale" (
			set answer=!input:~15!
			set answer=!answer: =!
			set VERBOSE_SYSTEM_LOCALE=!answer:*;=!
			call set SYSTEM_LOCALE_WITH_SEMICOLON=%%answer:!VERBOSE_SYSTEM_LOCALE!=%%
			set SYSTEM_LOCALE=!SYSTEM_LOCALE_WITH_SEMICOLON:~0,-1!
			@rem echo locale: !SYSTEM_LOCALE! ::
	   )
	)
	(ENDLOCAL & REM -- RETURN VALUES
		IF "%~1" NEQ "" SET %~1=%SYSTEM_LOCALE%
	)
	@goto :eof

:SleepX - UTILITY to wait/sleep x seconds
	timeout /t %1 >NUL
	::@:: set "sleep1=PING -n 2 127.0.0.1 >NUL 2>&1 || PING -n 2 ::1 >NUL 2>&1"
	@exit /b 0
	@goto :eof


:usage
@echo.
@echo.mini Data Collection Tool script ++ extensions
@echo.The script will colect DCT data (not all), PerfMon Logs (last 10 days), list of crash dumps (only list, no dmp files) and additional diagnostic data
@echo.The data will be archived in .cab or .zip file with name [hostName]_[Date]_[Time].cab/zip and working files will be deleted
@echo.This is default behaviour when running without paramaters. To change it you can run script with any of below parameters
@echo.  noDctData - skip collection of DCT data
@echo.  noPerfMon - skip Performance Counter colection - *.blg files
@echo.  noAddData - skip colection of the addtional diagnostic data
@echo.  noCabZip  - the data collected will not be compressed
@echo.Usage examples:
@echo. Example 1 - collect all data and create archive - default run without parameters
@echo. c:\Temp\^> %~nx0
@echo.
@echo. Example 2 - Do not collect Perfromance counters
@echo. c:\Temp\^> %~nx0  noPerfMon
@echo.
@echo. Example 3 - No additional data - only DCT data, PerfMon logs and crash dump list
@echo. c:\Temp\^> %~nx0  noAddData
@echo.
@echo. Example 4 - small DCT data collection - no PerMOn Logs, no addtional data
@echo. c:\Temp\^> %~nx0  noPerfMon  noAddData
@echo.
@echo. Example 5 - collect only extended diagnostic data
@echo. c:\Temp\^> %~nx0  noDctData noPerfMon
@echo.
@echo.** mDCT++ updates, more details and list of extended data collected on: https://github.com/kumanov/mDCT
@goto :eof

:InitLog [LogFileName]
	@if not exist %~1 (
		@echo.%date% %time% . INITIALIZE file %~1 by %USERNAME% on %COMPUTERNAME% in Domain %USERDOMAIN% > "%~1"
		@echo.mDCT++ [%_ScriptVersion%] ^(krasimir.kumanov@gmail.com^)^, git: https://github.com/kumanov/mDCT >> "%~1"
		@echo.>> "%~1"
	)
	@goto :eof

:logLine [_LogFile]
	@echo ================================================================================== >> %1
	@goto :eof

:logitem  - UTILITY to write a message to the log file (no indent) and screen
	@echo %date% %time% : %* >> "!_LogFile!"
	@echo %time% : %*
	@goto :eof

:logOnlyItem  - UTILITY to write a message to the log file (no indent)
	@echo %date% %time% : %* >> "!_LogFile!"
	@goto :eof

:logNoTimeItem  - UTILITY to write a message to the log file (no indent)
	@echo. %* >> "!_LogFile!"
	@goto :eof

:showlogitem  - UTILITY to write a message to the log file (no time indent) and screen
	@echo. %* >> "!_LogFile!"
	@echo. %*
	@goto :eof

:doCmd  - UTILITY log execution and output of a command to the current log file
	call :logLine "!_LogFile!"
	@echo ===== %time% : %* >> "!_LogFile!"
	call :logLine "!_LogFile!"
	%* >> "!_LogFile!" 2>&1
	@echo. >> "!_LogFile!"
	call :SleepX 1
	@goto :eof

:doCmdNoLog  - UTILITY log execution of a command to the current log file
	call :logLine "!_LogFile!"
	@echo ===== %time% : %* >> "!_LogFile!"
	call :logLine "!_LogFile!"
	%*
	@echo. >> "!_LogFile!"
	@goto :eof

:LogCmd [filename; command] - UTILITY to log command header and output in filename
	SETLOCAL
	for /f "tokens=1* delims=; " %%a in ("%*") do (
		set _LogFileName=%%a
		call :logLine "!_LogFileName!"
		@echo ===== %time% : %%b >> "!_LogFileName!"
		call :logLine "!_LogFileName!"
		%%b >> "!_LogFileName!" 2>&1
	)
	@echo. >> "!_LogFileName!"
	call :SleepX 1
	ENDLOCAL
	@goto :eof

:LogCmdNoSleep [filename; command] - UTILITY to log command header and output in filename
	SETLOCAL
	for /f "tokens=1* delims=; " %%a in ("%*") do (
		set _LogFileName=%%a
		call :logLine "!_LogFileName!"
		@echo ===== %time% : %%b >> "!_LogFileName!"
		call :logLine "!_LogFileName!"
		%%b >> "!_LogFileName!" 2>&1
	)
	@echo. >> "!_LogFileName!"
	ENDLOCAL
	@goto :eof

:LogWmicCmd [filename; command] - UTILITY to log command header and output in filename
	SETLOCAL
	for /f "tokens=1* delims=; " %%a in ("%*") do (
		set _LogFileName=%%a
		call :logLine "!_LogFileName!"
		@echo ===== %time% : %%b >> "!_LogFileName!"
		call :logLine "!_LogFileName!"
		::%%b /Format:Texttable | more /s >> "!_LogFileName!" 2>&1
		For /F "tokens=* delims=" %%h in ('%%b 2^>^&1') do (
			set "_line=%%h"
			set "_line=!_line:~0,-1!"
			echo.!_line!>>"!_LogFileName!"
		)
	)
	@echo.>> "!_LogFileName!"
	call :SleepX 1
	ENDLOCAL
	@goto :eof

:mkNewDir
	SETLOCAL
	set _NewDir=%*
	set _NewDir=!_NewDir:"=!!"!
	if not exist "%_NewDir%" mkdir "%_NewDir%"
	ENDLOCAL & REM -- RETURN VALUES
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
	@echo.>> %_RegFile%
	call :logLine %_RegFile%
	@echo ===== %time% : REG.EXE %* >> %_RegFile%
	call :logLine %_RegFile%
	%SYSTEMROOT%\SYSTEM32\REG.EXE %* >> %_RegFile% 2>&1
	@goto :eof

:GetReg_async [ filename "key"]
	for /f "tokens=1* delims= " %%a in ("%*") do (
		@echo.>> %%a
		call :logLine %%a
		@echo ===== %time% : REG QUERY %%b >> %%a
		call :logLine %%a
		START /WAIT "Please wait for REG QUERY to finish.." /MIN CMD /C "%SYSTEMROOT%\SYSTEM32\REG.EXE QUERY %%b >> %%a 2>&1"
	)
	@goto :eof

:GetRegValue Key Value Data Type -- returns a registry value
::                               -- Key    [in]  - registry key
::                               -- Value  [in]  - registry value
::                               -- Data   [out] - return variable for Data
::                               -- Type   [out] - return variable for Type, i.e.: REG_SZ, REG_MULTI_SZ, REG_DWORD_BIG_ENDIAN, REG_DWORD, REG_BINARY, REG_DWORD_LITTLE_ENDIAN, REG_NONE, REG_EXPAND_SZ
:$created 20060101 :$changed 20080219 :$categories Registry
:$source https://www.dostips.com
:$lastmodified 20190505 - krasimir.kumanov@gmail.com
SETLOCAL ENABLEDELAYEDEXPANSION
set Key=%~1
set Val=%~2
if "%Val%"=="" (set v=/ve) ELSE set v=/v "%Val%"
set Data=
set Type=
for /f "tokens=2,* delims= " %%a in ('reg query "%Key%" %v%^|findstr /b "....%match%"') do (
    set Type=%%a
	if /i "!Type!"=="REG_SZ" (
		set Data="%%b"
	) else (
		set Data=%%b
	)
)
( ENDLOCAL & REM RETURN VALUES
    IF "%~3" NEQ "" (SET %~3=%Data%) ELSE echo.%Data%
    IF "%~4" NEQ "" (SET %~4=%Type%)
)
EXIT /b

:DoSqlCmd [database query] - run sql query
	call :logLine %_SqlFile%
	@echo ===== %time% : sqlcmd DB:%~1 Q:%~2 >> %_SqlFile%
	call :logLine %_SqlFile%
	sqlcmd -E -w 10000 -d "%~1" -Q "%~2" >> %_SqlFile% 2>&1
	@echo. >> %_SqlFile%
	@goto :eof

:getSVC [comment] - UTILITY to dump Service information into log file
	call :logitem collecting Services info at %~1
	set _ServicesFile=!_DirWork!\GeneralSystemInfo\serviceslist.txt
	call :InitLog !_ServicesFile!
	call :LogCmd !_ServicesFile! SC.exe query type= all state= all
	call :SleepX 2
	@goto :eof

:DoNltestDomInfo [comment] - UTILITY to dump NLTEST Domain infos into log file
	call :logitem . collecting NLTEST Domain information
	set _NltestInfoFile=!_DirWork!\GeneralSystemInfo\_NltestDomInfo.txt
	call :InitLog !_NltestInfoFile!
	call :LogCmd !_NltestInfoFile! nltest /dsgetsite
	call :LogCmd !_NltestInfoFile! nltest /dsgetdc: /kdc /force
	call :LogCmd !_NltestInfoFile! nltest /dclist:
	call :LogCmd !_NltestInfoFile! nltest /trusted_domains
	@goto :eof

:getStationFiles   -- collect station configuration files
	SETLOCAL
	call :mkNewDir !_DirWork!\Station-logs

	set _HMIWebLog=%HwProgramData%\HMIWebLog\
	@set _stnFiles=!_DirWork!\Station-logs\_stnFiles.txt
	:: create stn file list
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci -Path '!_HMIWebLog!' -Include log.txt,hmiweblogY*.txt -Recurse | Select-String -Pattern 'Connecting using .stn file: (.*stn)$' | foreach{$_.Matches} | foreach {$_.Groups[1].Value} | sort -Unique | out-file '!_stnFiles!'  -Encoding unicode}}"
	:: copy files
	for /f "tokens=* delims=" %%h in ('type !_stnFiles!') DO (
		set _stnFile=_%%~nh%%~xh
		if exist "!_DirWork!\Station-logs\!_stnFile!" (set _stnFile=_%%~nh_!random!%%~xh)
		call :doCmd copy /y "%%h" "!_DirWork!\Station-logs\!_stnFile!"
	)
	call :SleepX 1
	
	:: get stb files
	@set _stbFiles=%temp%\_stbFiles.txt
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci -Path '!_DirWork!\Station-logs\' -Include *.stn -Recurse | Select-String -Pattern 'Toolbar_Settings=(.*stb)$' | foreach{$_.Matches} | foreach {$_.Groups[1].Value} | sort -Unique | out-file '!_stbFiles!'}}"
	:: copy files
	for /f "tokens=* delims=" %%h in ('type !_stbFiles!') DO (
		set _stbFile=_%%~nh%%~xh
		if exist "!_DirWork!\Station-logs\!_stbFile!" (set _stbFile=_%%~nh_!random!%%~xh)
		call :doCmd copy /y "%%h" "!_DirWork!\Station-logs\!_stbFile!"
	)
	del !_stbFiles!
	call :SleepX 1

	:: get Display Links files
	@set _dspLinksFiles=%temp%\_stbFiles.txt
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci -Path '!_DirWork!\Station-logs\' -Include *.stn -Recurse | Select-String -Pattern 'DisplayLinksPath=(.*xml)$' | foreach{$_.Matches} | foreach {$_.Groups[1].Value} | sort -Unique | out-file '!_dspLinksFiles!'}}"
	:: copy files
	for /f "tokens=* delims=" %%h in ('type !_dspLinksFiles!') DO (
		set _dspLinksFile=_%%~nh%%~xh
		if exist "!_DirWork!\Station-logs\!_dspLinksFile!" (set _dspLinksFile=_%%~nh_!random!%%~xh)
		call :doCmd copy /y "%%h" "!_DirWork!\Station-logs\!_dspLinksFile!"
	)
	del !_dspLinksFiles!
	call :SleepX 1

	(ENDLOCAL & REM -- RETURN VALUES
	)
	exit /b

:getGDIHandlesCount    -- get GDI Handles Count
SETLOCAL
call :logitem . get GDI Handles Count
set _psFile=%temp%\psFile.ps1
set _outFile=!_DirWork!\GeneralSystemInfo\_GDIHandlesCount.txt
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :InitLog !_outFile!
:: create ps1 file
>!_psFile!  echo Add-Type -Name NativeMethods -Namespace Win32 -MemberDefinition @'
>>!_psFile! echo [DllImport("User32.dll")]
>>!_psFile! echo public static extern int GetGuiResources(IntPtr hProcess, int uiFlags);
>>!_psFile! echo '@;
>>!_psFile! echo $allProcesses = Get-Process; $auxCountHandles = [int]0; $auxCountProcess = [int]0; $GuiResources = @();
>>!_psFile! echo ForEach ($p in $allProcesses) { if ( ($p.Handle -eq '') -or ($p.Handle -eq $null) ) { continue }; $auxCountProcess += 1; $auxGdiHandles = [Win32.NativeMethods]::GetGuiResources($p.Handle, 0);
>>!_psFile! echo If ($auxGdiHandles -eq 0) { continue }
>>!_psFile! echo $auxCountHandles += $auxGdiHandles; $auxDict = @{ PID = $p.Id; GDIHandles = $auxGdiHandles; ProcessName = $p.Name; };
>>!_psFile! echo $GuiResources += New-Object -TypeName psobject -Property $auxDict; };
>>!_psFile! echo $GuiResources ^| sort GDIHandles -Desc ^| select -First 10 ^| ft -a ^| out-file '!_outFile!' -Append -Encoding ascii
::@::>>!_psFile! echo '' ^| out-file !_outFile! -Append -Encoding ascii
>>!_psFile! echo $('{0} processes; {1}/{2} with/without GDI objects' -f $allProcesses.Count, $GuiResources.Count, ($allProcesses.Count - $GuiResources.Count)) ^| out-file '!_outFile!' -Append -Encoding ascii
>>!_psFile! echo "Total number of GDI handles: $auxCountHandles" ^| out-file '!_outFile!' -Append -Encoding ascii
PowerShell.exe -NonInteractive  -NoProfile -ExecutionPolicy Bypass %_psFile%
@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at getGDIHandlesCount with PowerShell'. )
if "%errorlevel%" neq "0" (
	call :logItem %time% .. ERROR: %errorlevel% - 'getGDIHandlesCount with PowerShell' failed.
	)
::del PS file
call :doit del "!_psFile!"
ENDLOCAL
call :SleepX 1
exit /b

:getGroupMembers    -- get group mebers
::                 -- %~1 [in]: group name
::                 -- %~2 [in]: out file
	SETLOCAL
	set _groupName=%~1
	if defined %1 set _groupName=!%~1!
	set _outFile=%~2
	if defined %2 set _outFile=!%~2!

	>>!_outFile! echo ========================================
	>>!_outFile! echo get '!_groupName!' members
	>>!_outFile! echo ========================================
	PowerShell.exe -NonInteractive  -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { @(try{@(([ADSI]'WinNT://./!_groupName!').Invoke('Members'))}catch{}) | foreach{$_.GetType().InvokeMember('Name', 'GetProperty', $null, $_, $null)} | Out-File '!_outFile!' -Append -Encoding ascii }}"
	@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - at get '!_groupName!' members with PowerShell. )
	if "%errorlevel%" neq "0" (
		call :logItem %time% .. ERROR: %errorlevel% - get '!_groupName!' members with PowerShell failed.
		)
	call :SleepX 1
	>>!_outFile! echo.

	ENDLOCAL
	call :SleepX 1
	exit /b

:GetHkeyUsersRegValues    -- GetHkeyUsersRegValues - output to _RegFile
::                 -- %~1 [in,out,opt]: argument description here
	SETLOCAL
	call :InitLog !_RegFile!
	for /f "tokens=2 delims=\" %%h in ('reg query hkey_users') do (
		set _SID=%%h
		reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\!_SID!" /v ProfileImagePath >NUL 2>&1
		if "!errorlevel!"=="0" (
			echo.>>!_RegFile!
			echo @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@>>!_RegFile!
			echo USER SID: !_SID!>>!_RegFile!
			call :GetReg QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\!_SID!" /v ProfileImagePath
			call :GetReg QUERY "HKEY_USERS\!_SID!\Control Panel\International" /v LocaleName
			call :GetReg QUERY "HKEY_USERS\!_SID!\Control Panel\International" /v sDecimal
			call :GetReg QUERY "HKEY_USERS\!_SID!\Control Panel\Desktop" /v ForegroundLockTimeout
			call :GetReg QUERY "HKEY_USERS\!_SID!\Control Panel\Desktop" /v WindowArrangementActive
			:: R4xx - Configure for best performance
			call :GetReg QUERY "HKEY_USERS\!_SID!\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" /v VisualFXSetting
			:: R5xx - Setting Windows ColorPrelevance (making active window title bar to change its color)
			call :GetReg QUERY "HKEY_USERS\!_SID!\Software\Microsoft\Windows\DWM" /v ColorPrevalence
			if NOT "!_VEP!"=="1" (
				if NOT "!_isServer!"=="1" (
					call :GetReg QUERY "HKEY_USERS\!_SID!\Software\Microsoft\Avalon.Graphics" /v DisableHWAcceleration
				)
			)

			)
		)
	(ENDLOCAL & REM -- RETURN VALUES
		call :SleepX 1
	)
	exit /b

:GetUserSID~   -- get user SID
::             -- %~1 [in]: user name
::             -- %~2 [out]: var SID
	SETLOCAL
	set _UserName=%~1
	for /f "tokens=1* delims==" %%i in (
	  'wmic useraccount where "name='!_UserName!'" get sid /value'
	) do for /f "delims=" %%k in ("%%j") do set "_SID=%%k"

	(ENDLOCAL & REM -- RETURN VALUES
		IF "%~2" NEQ "" SET %~2=%_SID%
	)
	exit /b

:GetUserSID    -- get user SID
::             -- %~1 [in]: user name
::             -- %~2 [out]: var SID
	SETLOCAL
	set __UserName=%~1
	for /f "usebackq delims=" %%h in (`PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { (New-Object System.Security.Principal.NTAccount '!__UserName!').Translate([System.Security.Principal.SecurityIdentifier]).Value }}"`) do set "__SID=%%h"

	(ENDLOCAL & REM -- RETURN VALUES
		IF "%~2" NEQ "" (SET %~2=%__SID%
		) else (echo %__SID%)
	)
	exit /b

:ExperionAclVerify    -- Experion Acl Verify
	SETLOCAL
	set _AclVerify="!_DirWork!\GeneralSystemInfo\_AclHwVerify.txt"
	call :mkNewDir !_DirWork!\GeneralSystemInfo
	call :InitLog !_AclVerify!
	call :logitem . Experion ACL Verify
	
	call :logOnlyItem . Experion ACL Verify - HKLM:\SOFTWARE\Honeywell\
	@echo.>>!_AclVerify!
	call :logLine !_AclVerify!
	@echo ===== !time! : Get-Acl HKLM:\SOFTWARE\Honeywell\ >>!_AclVerify!
	call :logLine !_AclVerify!
	PowerShell.exe -NonInteractive  -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { Get-ChildItem -Path HKLM:\SOFTWARE\Honeywell\ -Recurse -ea 0 | ForEach-Object {Get-Acl $_.PSPath -ea 0} | Where-Object { -not $_.AreAccessRulesCanonical} | out-file '!_AclVerify!' -Append -Encoding ascii }}"
	
	call :logOnlyItem . Experion ACL Verify - HKLM:\SOFTWARE\Wow6432Node\Honeywell\
	@echo.>>!_AclVerify!
	call :logLine !_AclVerify!
	@echo ===== !time! : Get-Acl HKLM:\SOFTWARE\Wow6432Node\Honeywell\ >>!_AclVerify!
	call :logLine !_AclVerify!
	PowerShell.exe -NonInteractive  -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { Get-ChildItem -Path HKLM:\SOFTWARE\Wow6432Node\Honeywell\ -Recurse -ea 0 | ForEach-Object {Get-Acl $_.PSPath -ea 0} | Where-Object { -not $_.AreAccessRulesCanonical} | out-file '!_AclVerify!' -Append -Encoding ascii }}"

	call :logOnlyItem . Experion ACL Verify - "%HwProgramData%\Experion PKS"
	@echo.>>!_AclVerify!
	call :LogCmd !_AclVerify! ICACLS "%HwProgramData%\Experion PKS" /verify /T /C /L /Q
	
	call :logOnlyItem . Experion ACL Verify - "%HwInstallPath%\Experion PKS"
	@echo.>>!_AclVerify!
	call :logLine !_AclVerify!
	@echo ===== !time! : ICACLS "%HwInstallPath%\Experion PKS" /verify /T /C /L /Q >>!_AclVerify!
	call :logLine !_AclVerify!
	rem too slow on R511 with "User Assistance" installed - skip
	if not exist "%HwInstallPath%\Experion PKS\User Assistance" (
		ICACLS "%HwInstallPath%\Experion PKS" /verify /T /C /L /Q >>!_AclVerify!
	)
	
	if exist "%HwProgramData%\Experion PKS\PatchDB" (
		call :logOnlyItem . VERIFY THE PATCH DB FOLDER SECURITY
		@echo.>>!_AclVerify!
		call :logLine !_AclVerify!
		@echo ===== !time! : ICACLS "%HwProgramData%\Experion PKS\PatchDB" >>!_AclVerify!
		call :logLine !_AclVerify!
		call :LogCmd !_AclVerify! ICACLS "%HwProgramData%\Experion PKS\PatchDB"
	)
	
	(ENDLOCAL & REM -- RETURN VALUES
	)
	exit /b

:isEmpty    -- function description here
::                 -- %~1 [in]: file to check
::                 -- %~1 [out,opt]: out var
SETLOCAL
@if defined _Debug ( echo [#Debug#] !time!: function isEmpty - file: %~f1 , size: %~z1 )
if %~z1 == 0 (
	set __isEmpty=1
) else (
	set __isEmpty=
)
(ENDLOCAL & REM -- RETURN VALUES
	IF "%~2" NEQ "" SET %~2=%__isEmpty%
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


:mkCab sourceFolder cabFolder cabName	-- make cab file
::										-- sourceFolder [in] - source folder
::										-- cabFolder    [in] - destination, cab file folder
::										-- cabName      [out] - cab file name
SETLOCAL ENABLEDELAYEDEXPANSION
:: change working directory
pushd "%temp%"
:: define variables in .ddf file
>directives.ddf echo ; Makecab Directive File
>>directives.ddf echo ; Created by mDCT++ script tool
>>directives.ddf echo ; %date% %time%
>>directives.ddf echo .Option Explicit
>>directives.ddf echo .Set DiskDirectoryTemplate="%~2"
>>directives.ddf echo .Set CabinetNameTemplate="%~3"
>>directives.ddf echo .Set MaxDiskSize=0
>>directives.ddf echo .Set CabinetFileCountThreshold=0
>>directives.ddf echo .Set UniqueFiles=OFF
>>directives.ddf echo .Set Cabinet=ON
>>directives.ddf echo .Set Compress=ON
if "_isServer"=="1" (
>>directives.ddf echo .Set CompressionType=MSZIP
) else (
>>directives.ddf echo .Set CompressionType=LZX
)
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
set mat=&rem variable to store matching part of the name ::
set upp=&rem variable to reference a parent ::
for /f "tokens=*" %%a in ('echo.%bas:\=^&echo.%') do (
    set sub=!sub!%%a\
    call set tmp=%%src:!sub!=%%
    if "!tmp!" NEQ "!src!" (set mat=!sub!)ELSE (set upp=!upp!..\)
)
set src=%upp%!src:%mat%=!
( ENDLOCAL & REM RETURN VALUES ::
    IF defined %1 (	SET %~1=%src%) ELSE ECHO.%src%
)
exit /b

:getTimeZoneInfo
::                 -- %~1 [out]: TimeZone Name
::                 -- %~1 [out]: TimeZone Bias
SETLOCAL
REM Obtain the ActiveBias value and convert to decimal
for /f "tokens=3" %%a in ('reg query HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation /v ActiveTimeBias') do set /a abias=%%a
for /f "tokens=2*" %%h in ('reg query HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation /v TimeZoneKeyName') do set tzn=%%i 2>NUL

REM Set the + or - sign variable to reflect the timezone offset
IF "%abias:~0,1%"=="-" (set si=+) ELSE (set si=-)
for /f "tokens=1 delims=-" %%t in ('echo %abias%') do set tzc=%%t

REM Calculate to obtain floating points (decimal values)
set /a tzd=100*%tzc%/60

REM Calculate the active bias to obtain the hour
set /a tze=%tzc%/60

REM Set the minutes based on the result of the floating point calculation
IF "%tzd%"=="0" (set en=00 && set si=)
IF "%tzd:~1%"=="00" (set en=00) ELSE IF "%tzd:~2%"=="00" (set en=00 && set tz=%tzd:~0,2%)
IF "%tzd:~1%"=="50" (set en=30) ELSE IF "%tzd:~2%"=="50" (set en=30 && set tz=%tzd:~0,2%)
IF "%tzd:~1%"=="75" (set en=45) ELSE IF "%tzd:~2%"=="75" (set en=45 && set tz=%tzd:~0,2%)

REM Adding a 0 to the beginning of a single digit hour value
IF %tze% LSS 10 (set tz=0%tze%)
(ENDLOCAL & REM -- RETURN VALUES
	IF "%~1" NEQ "" SET %~1=%tzn%
	IF "%~2" NEQ "" SET %~2=%si%%tz%%en%
)
exit /b

:export-evtx  -- function to export Windows Events in evtx file
	::            -- %~1 [in]: Log Name
	::            -- %~2 [in,opt]: output folder or file
	SETLOCAL
	set "_Channel=%~1"
	set "_folder=%~2"
	if /i "%_folder:~-1%"=="\" (
		call :mkNewDir %_folder%
		set "_evtxFile=!_Channel:/=-!.evtx"
		if /i "!_evtxFile:~0,18!" equ "Microsoft-Windows-" (set "_evtxFile=!_evtxFile:~18!")
		set "_evtxFile=!_folder!!_evtxFile!"
	) else (
		set "_evtxFile=!_folder!"
	)
	:: set events date/time limit
	for /f "usebackq delims=" %%h in (`PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { (Get-Date).AddDays(-60).toString('s')+'Z' }}"`) do set "_TimeLimit=%%h"
	:: export events
	wevtutil epl "!_Channel!" "!_evtxFile!" "/q:*[System[TimeCreated[@SystemTime>='!_TimeLimit!']]]" /overwrite:true
	ENDLOCAL
	call :SleepX 1
	@goto :eof

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

:ClearCaches - UTILITY to clear DNS, Netbios and Kerberos-Ticket Caches
	call :logitem .. deleting DNS, NetBIOS and Kerberos caches
	call :doCmd IPconfig /flushDNS
	call :doCmd NBTstat -RR
	call :doCmd %windir%\system32\KLIST.exe purge -li 0x3e7
	call :doCmd %windir%\system32\KLIST.exe purge
	::call :logitem .. deleting DFS cache
	::call :doCmd DFSutil /PKTflush
	@goto :eof

:endregion

:region DCT Data
:CollectDctData
:: return, if no DCT data required
if /i "!_NoDctData!" EQU "1" (@goto :eof)

call :logitem *** DCT data collection ... ***

:: GeneralSystemInfo folder
call :mkNewDir  !_DirWork!\GeneralSystemInfo

call :logitem MSInfo32 report
call :doCmd msinfo32 /report "!_DirWork!\GeneralSystemInfo\MSInfo32.txt"

call :logitem Windows hosts file copy
call :doCmd copy /y %windir%\System32\drivers\etc\hosts "!_DirWork!\GeneralSystemInfo\"

call :logitem ipconfig output
call :LogCmd !_DirWork!\GeneralSystemInfo\ipconfig.txt ipconfig /all

call :logitem netsh firewall show config
call :LogCmd !_DirWork!\GeneralSystemInfo\firewall.txt netsh firewall show config

call :logitem get time zone information
call :getTimeZoneInfo _tzName _tzBias
call :InitLog !_DirWork!\GeneralSystemInfo\timezone.output.txt
@echo TimeZone Name: %_tzName% >>"!_DirWork!\GeneralSystemInfo\timezone.output.txt"
@echo TimeZone Bias: %_tzBias% >>"!_DirWork!\GeneralSystemInfo\timezone.output.txt"
call :SleepX 1

call :logitem export Windows Events
call :export-evtx Application !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_Application.evtx
call :SleepX 1
call :export-evtx FTE         !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_FTE.evtx
call :export-evtx HwSnmp      !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_HwSnmp.evtx
call :export-evtx HwSysEvt    !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_HwSysEvt.evtx
call :export-evtx Security    !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_Security.evtx
call :SleepX 1
call :export-evtx System      !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_System.evtx
call :SleepX 1

call :logitem get Experion PKS Product Version file
call :doCmd copy /y "%HwInstallPath%\Experion PKS\ProductVersion.txt" "!_DirWork!\GeneralSystemInfo\"

call :logitem query services
call :getSVC %time%
:: one more second to sleep
call :SleepX 1

call :logitem export Experion registry settings
call :mkNewDir  "!_DirWork!\RegistryInfo"
if exist %windir%\SysWOW64\regedit.exe (set _regedit="%windir%\SysWOW64\regedit.exe") else (set _regedit=regedit.exe)
:: HKEY_CURRENT_USER\Software\Honeywell
set _RegFileOut=!_DirWork!\RegistryInfo\HKEY_CURRENT_USER_Software_Honeywell.txt
call :doCmd !_regedit! /E "!_RegFileOut!" "HKEY_CURRENT_USER\Software\Honeywell"
if not exist "!_RegFileOut!" (call :doCmd REG EXPORT "HKEY_CURRENT_USER\Software\Honeywell" "!_RegFileOut!")
::HKEY_LOCAL_MACHINE\SOFTWARE\Honeywell
set _RegFileOut=!_DirWork!\RegistryInfo\HKEY_LOCAL_MACHINE_Software_Honeywell.txt
call :doCmd !_regedit! /E "!_RegFileOut!" "HKEY_LOCAL_MACHINE\SOFTWARE\Honeywell"
if not exist "!_RegFileOut!" (call :doCmd REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Honeywell" "!_RegFileOut!")
::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set _RegFileOut=!_DirWork!\RegistryInfo\HKEY_LOCAL_MACHINE_Software_Microsoft_Uninstall.txt
call :doCmd !_regedit! /E "!_RegFileOut!" "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
if not exist "!_RegFileOut!" (call :doCmd REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" "!_RegFileOut!")


call :logitem get FTE logs
::@::call :doCmd xcopy /s/e/i/q/y/H "%HwProgramData%\ProductConfig\FTE\*.log" "!_DirWork!\FTELogs\"
call :mkNewDir  "!_DirWork!\FTELogs"
for /r "%HwProgramData%\ProductConfig\FTE" %%g in (*.log) do (type %%g >"!_DirWork!\FTELogs\%%~ng%%~xg")

call :logitem get HMIWeb log files
call :doCmd xcopy /i/q/y/H "%HwProgramData%\HMIWebLog\*log*.txt" "!_DirWork!\Station-logs\"
if NOT "_isServer"=="1" (
	::old cmd:: call :doCmd xcopy /i/q/y/H "%HwProgramData%\HMIWebLog\Archived Logfiles\*.txt" "!_DirWork!\Station-logs\Rollover-logs\"
	if exist "%HwProgramData%\HMIWebLog\Archived Logfiles\*.txt" (
		call :logOnlyItem get HMIWeb backup log files
		call :mkNewDir !_DirWork!\Station-logs\Rollover-logs
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '%HwProgramData%\HMIWebLog\Archived Logfiles\' -filt *.txt | where{$_.LastWriteTime -gt (get-date).AddDays(-14)} | foreach{copy $_.fullName -dest '!_DirWork!\Station-logs\Rollover-logs'; sleep -Milliseconds 500} }}
		@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Copy station log archived files with PowerShell'. )
		if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy station log archived files with PowerShell' failed.)
	)
)
call :doCmd copy /y "%HwProgramData%\HMIWebLog\PersistentDictionary.xml" "!_DirWork!\Station-logs\"

call :logitem tasklist /svc
call :mkNewDir  !_DirWork!\ServerDataDirectory
call :LogCmd !_DirWork!\ServerDataDirectory\TaskList.txt tasklist /svc

where setpar >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion active log paranoids
	call :mkNewDir !_DirWork!\SloggerLogs
	call :LogCmd !_DirWork!\SloggerLogs\setpar.active.txt setpar /active
)

::SloggerLogs
:: R5xx
if exist "%HwProgramData%\Experion PKS\logfiles\logServer.txt" (
	call :logitem get Experion log files
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\logfiles\log*.txt" "!_DirWork!\SloggerLogs\"
	:: copy server log archives
	set _LogArchiveDirectory=%HwProgramData%\Experion PKS\logfiles\00-Server\
	if exist "!_LogArchiveDirectory!" (
		call :logOnlyitem get Experion Server backup log files
		call :logOnlyitem _LogArchiveDirectory=!_LogArchiveDirectory!
		call :mkNewDir !_DirWork!\SloggerLogs\Archives
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '!_LogArchiveDirectory!' -filt log*.txt | where{$_.LastWriteTime -gt (get-date).AddDays(-14)} | foreach{copy $_.fullName -dest '!_DirWork!\SloggerLogs\Archives\'; sleep -Milliseconds 500} }}
		@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Copy server log archived files with PowerShell'. )
		if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy server log archived files with PowerShell' failed.)
	)
	:: copy CDA log archives
	set _LogArchiveDirectory=%HwProgramData%\Experion PKS\logfiles\09-CDA\
	if exist "!_LogArchiveDirectory!" (
		call :logOnlyitem get CDA backup log files
		call :logOnlyitem _LogArchiveDirectory=!_LogArchiveDirectory!
		call :mkNewDir !_DirWork!\SloggerLogs\Archives
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '!_LogArchiveDirectory!' -filt log*.txt | where{$_.LastWriteTime -gt (get-date).AddDays(-14)} | foreach{copy $_.fullName -dest '!_DirWork!\SloggerLogs\Archives\'; sleep -Milliseconds 500} }}
		@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Copy server log archived files with PowerShell'. )
		if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy server log archived files with PowerShell' failed.)
	)
	:: copy SR log archives
	set _LogArchiveDirectory=%HwProgramData%\Experion PKS\logfiles\11-SysRep\
	if exist "!_LogArchiveDirectory!" (
		call :logOnlyitem get SR backup log files
		call :logOnlyitem _LogArchiveDirectory=!_LogArchiveDirectory!
		call :mkNewDir !_DirWork!\SloggerLogs\Archives
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '!_LogArchiveDirectory!' -filt logS*.txt | where{$_.LastWriteTime -gt (get-date).AddDays(-14)} | foreach{copy $_.fullName -dest '!_DirWork!\SloggerLogs\Archives\'; sleep -Milliseconds 500} }}
		@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Copy server log archived files with PowerShell'. )
		if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy server log archived files with PowerShell' failed.)
	)
	:: copy Activity log archives
	set _LogArchiveDirectory=%HwProgramData%\Experion PKS\logfiles\12-Activity\
	if exist "!_LogArchiveDirectory!" (
		call :logOnlyitem get Activity backup log files
		call :logOnlyitem _LogArchiveDirectory=!_LogArchiveDirectory!
		call :mkNewDir !_DirWork!\SloggerLogs\Archives
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '!_LogArchiveDirectory!' -filt log*.txt | where{$_.LastWriteTime -gt (get-date).AddDays(-14)} | foreach{copy $_.fullName -dest '!_DirWork!\SloggerLogs\Archives\'; sleep -Milliseconds 500} }}
		@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Copy server log archived files with PowerShell'. )
		if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy server log archived files with PowerShell' failed.)
	)
	:: copy EnggTools log archives
	set _LogArchiveDirectory=%HwProgramData%\Experion PKS\logfiles\25-EnggTools\
	if exist "!_LogArchiveDirectory!" (
		call :logOnlyitem get EnggTools backup log files
		call :logOnlyitem _LogArchiveDirectory=!_LogArchiveDirectory!
		call :mkNewDir !_DirWork!\SloggerLogs\Archives
		PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '!_LogArchiveDirectory!' -filt log*.txt | where{$_.LastWriteTime -gt (get-date).AddDays(-14)} | foreach{copy $_.fullName -dest '!_DirWork!\SloggerLogs\Archives\'; sleep -Milliseconds 500} }}
		@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Copy server log archived files with PowerShell'. )
		if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy server log archived files with PowerShell' failed.)
	)
)
:: R4xx
if exist "%HwProgramData%\Experion PKS\Server\data\log.txt" (
	call :logitem get Experion log files
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\*log*.txt" "!_DirWork!\SloggerLogs\"
)


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

where actutil >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion point back build
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd actutil --dump -o "!_DirWork!\ServerDataDirectory\actutil.output.txt"
)

where almdmp >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion alarm/event dump
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :LogCmd !_DirWork!\ServerDataDirectory\almdmp.output.txt almdmp A 32000 S
	call :LogCmd !_DirWork!\ServerDataDirectory\eventdmp.output.txt almdmp E 32000 S
	call :LogCmd !_DirWork!\ServerDataDirectory\msgdmp.output.txt almdmp M 32000 S
	call :LogCmd !_DirWork!\ServerDataDirectory\alertdmp.output.txt almdmp T 32000 S
	call :LogCmd !_DirWork!\ServerDataDirectory\soedmp.output.txt almdmp S 32000 S
)

where shheap >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion shheap output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\shheap.1.output.txt shheap 1 struct
    call :logCmd !_DirWork!\ServerDataDirectory\shheap.1.dump.output.txt shheap 1 dump
    call :logCmd !_DirWork!\ServerDataDirectory\shheap.4.struct.output.txt shheap 4 struct
)

where bckbld >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion point back build
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd bckbld -out "!_DirWork!\ServerDataDirectory\back_build.output.txt"
)

where hdwbckbld >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion hardware back build
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd hdwbckbld -out "!_DirWork!\ServerDataDirectory\hardware_back_build.output.txt"
)

where hstdiag >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion history diagnostic
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\hstdiag.output.txt hstdiag
)

where embckbuilder >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion embckbuilder output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd embckbuilder  "!_DirWork!\ServerDataDirectory\embckbuilder.alarmgroup.output.txt"  -ALARMGROUP
    call :doCmd embckbuilder  "!_DirWork!\ServerDataDirectory\embckbuilder.asset.output.txt"       -ASSET
    call :doCmd embckbuilder  "!_DirWork!\ServerDataDirectory\embckbuilder.network.output.txt"     -NETWORK
    call :doCmd embckbuilder  "!_DirWork!\ServerDataDirectory\embckbuilder.system.output.txt"      -SYSTEM
)

if exist "%HwProgramData%\Experion PKS\Server\data\OPCIntegrator\" (
	call :logitem Experion OPC Integrator
	call :mkNewDir !_DirWork!\OPCIntegrator
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\OPCIntegrator\*.tsv" "!_DirWork!\OPCIntegrator\"
)

where listag >NUL 2>&1
if %errorlevel%==0 (
	call :logitem listag output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :logCmd !_DirWork!\ServerDataDirectory\listag.output.txt listag -ALL
)

where fildmp >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion System Flags Table output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd fildmp -DUMP -FILE "!_DirWork!\ServerDataDirectory\sysflg.output.txt" -FILENUM 8 -RECORDS 1 -FORMAT HEX
	call :logitem Experion Area Asignmnt Table output
    call :doCmd fildmp -DUMP -FILE "!_DirWork!\ServerDataDirectory\areaasignmnt.output.txt" -FILENUM 7 -RECORDS 1,1001 -FORMAT HEX
)

if exist "%HwProgramData%\Experion PKS\Server\data\system.build" (
	call :logitem copy system.build file
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\system.build" "!_DirWork!\ServerDataDirectory\"
)

if exist "%HwProgramData%\Experion PKS\Server\data\" (
	call :logitem collect bad files .\server\data\*.bad
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\*.bad" "!_DirWork!\ServerDataDirectory\"
)

if exist "%HwProgramData%\TPNServer\TPNServer.log" (
	call :logitem copy TPNServer.log file
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\TPNServer\TPNServer.log" "!_DirWork!\ServerDataDirectory\"
)

(::winsxs
call :logitem list C:\Windows\winsxs\
call :mkNewDir !_DirWork!\ServerDataDirectory
call :logCmd !_DirWork!\ServerDataDirectory\winsxs.txt dir %windir%\winsxs)

where liclist >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion license list - liclist
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\liclist.output.txt
    call :logCmd !_DirWork!\ServerRunDirectory\liclist.output.txt liclist
)

where hwlictool >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion license list - hwlictool
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\hwlictool.output.txt
    call :logCmd !_DirWork!\ServerRunDirectory\hwlictool.output.txt hwlictool export -format:xml
    call :logCmd !_DirWork!\ServerRunDirectory\hwlictool.output.txt hwlictool status -format:xml
    call :logCmd !_DirWork!\ServerRunDirectory\hwlictool.output.txt hwlictool list
)
where usrlrn >NUL 2>&1
if %errorlevel%==0 (
	call :logitem usrlrn -p -a
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\usrlrn.txt
    call :logCmd !_DirWork!\ServerRunDirectory\usrlrn.txt usrlrn -p -a
)

:: skip it - very slow and trigger AV scans
::where what >NUL 2>&1
::if %errorlevel%==0 (
::	call :logitem What - Getting Experion exe/dll and source file information
::	call :mkNewDir !_DirWork!\ServerRunDirectory
::	call :InitLog !_DirWork!\ServerRunDirectory\what.output.txt
::	for /r "%HwInstallPath%\Experion PKS\Server\run" %%a in (*.exe *.dll) do what "%%a" >>!_DirWork!\ServerRunDirectory\what.output.txt
::)

where notifdmp >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Notification Utility - Dump indexes
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :InitLog !_DirWork!\ServerDataDirectory\notifindexes.output.txt
    call :logCmd  !_DirWork!\ServerDataDirectory\notifindexes.output.txt notifdmp --dump-indexes
)

:: ErrorHandling
call :logitem get ErrorHandling log files
call :mkNewDir !_DirWork!\ErrorHandling
call :doCmd copy /y "%HwProgramData%\Experion PKS\BrowserLog_Current.txt" "!_DirWork!\ErrorHandling\"
call :doCmd copy /y "%HwProgramData%\Experion PKS\CurrentLogIndex.txt" "!_DirWork!\ErrorHandling\"
call :doCmd copy /y "%HwProgramData%\Experion PKS\ErrLog_*.txt" "!_DirWork!\ErrorHandling\"

:: get CreateSQLObject logs
if exist "%HwProgramData%\\Experion PKS\Server\data\CreateSQLObject*.txt" (
	call :logitem get CreateSQLObject logs
	call :mkNewDir !_DirWork!\CreateSQL-Logs
	call :doCmd copy /y "%HwProgramData%\\Experion PKS\Server\data\CreateSQLObject*.txt" "!_DirWork!\CreateSQL-Logs\"
)

:: get SQL errorlog files
where sqlcmd >NUL 2>&1
if %errorlevel% EQU 0 (
	call :logitem get SQL Error log files
	call :mkNewDir !_DirWork!\MSSQL-Logs
	FOR /F "usebackq tokens=2 delims='" %%h IN (`sqlcmd -E -w 10000 -d master -Q "xp_readerrorlog 0, 1, N'Logging SQL Server messages in file'" ^| find /i "Logging SQL Server messages in file"`) Do  ( @set _SqlLogFile=%%h)
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '!_SqlLogFile!*' | foreach{copy $_.fullName -dest '!_DirWork!\MSSQL-Logs\'; sleep -Milliseconds 250} }}
)


if exist "%HwProgramData%\Experion PKS\Server\data\das\DasConfig.xml" (
	call :logitem get DasConfig.xml
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\das\DasConfig.xml" "!_DirWork!\ServerDataDirectory\"
)


goto :eof
:endregion

:region Additional Data
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:CollectAdditionalData
:: return, if no Additional data required
if /i "!_NoAddData!" EQU "1" (@goto :eof)

call :logitem *** Additional data collection ... ***

call :ExperionAddData
call :NetworkAddData
call :WindowsAddData
call :SqlAddData

goto :eof

:region Windows
:WindowsAddData
call :logitem * Windows data *

call :mkNewDir  !_DirWork!\GeneralSystemInfo

call :logitem . SystemInfo.exe output
call :InitLog !_DirWork!\GeneralSystemInfo\_SystemInfo.txt
call :LogCmd !_DirWork!\GeneralSystemInfo\_SystemInfo.txt systeminfo.exe

call :logitem . whoami - currently logged in user
call :InitLog !_DirWork!\GeneralSystemInfo\_whoami.txt
START "Please wait on WhoAmI.exe to finish..." /MIN /D "!_DirWork!" CMD /C "whoami.exe -all >"!_DirWork!\GeneralSystemInfo\_whoami.txt""

call :logitem . scheduled task - query
schtasks /query /xml ONE >!_DirWork!\GeneralSystemInfo\_scheduled_tasks.xml
schtasks /query /fo list /v >!_DirWork!\GeneralSystemInfo\_scheduled_tasks.txt
call :SleepX 1

call :logitem . collecting GPResult output
set _GPresultFile=!_DirWork!\GeneralSystemInfo\_GPresult.htm
START "Please wait on GPresult.exe to finish..." /MIN /D "!_DirWork!" CMD /C "GPresult.exe /h "!_GPresultFile!" /f"
REM if "%errorlevel%" neq "0" call :LogCmd !_DirWork!\GeneralSystemInfo\_GPresultZ.txt gpresult /Z


call :DoNltestDomInfo

call :logitem . get power configuration settings
set _powerCfgFile="!_DirWork!\GeneralSystemInfo\_powercfg.txt"
call :InitLog !_powerCfgFile!
call :LogCmd !_powerCfgFile! powercfg -list
call :LogCmd !_powerCfgFile! powercfg -Q
:: reg query power settings
call :logOnlyItem . reg query power settings
set _RegFile=!_powerCfgFile!
::  fast reboot - "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power"  /s /t reg_dword
::  Hibernate - "HKLM\System\CurrentControlSet\Control\Power" /v HibernateEnabled
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\Power"  /s /t reg_dword

call :logitem . collecting Quick Fix Engineering information (Hotfixes)
call :doCmd  wmic /output:"!_DirWork!\GeneralSystemInfo\_Hotfixes.txt" qfe list full /format:table

if exist "%windir%\Honeywell_MsPatches.txt" (
	call :logitem . get Honeywell_MsPatches.txt
	call :doCmd copy /y "%windir%\Honeywell_MsPatches.txt" "!_DirWork!\GeneralSystemInfo\_Honeywell_MsPatches.txt"
)
if exist "%HwInstallPath%\\Experion PKS\Install\Honeywell_MsPatches.txt" (
	call :logitem . get Honeywell_MsPatches.txt
	call :doCmd copy /y "%HwInstallPath%\\Experion PKS\Install\Honeywell_MsPatches.txt" "!_DirWork!\GeneralSystemInfo\_Honeywell_MsPatches.txt"
)
if exist "%windir%\honeywell_installed_updates.txt" (
	call :logitem . get honeywell_installed_updates.txt
	call :doCmd copy /y "%windir%\honeywell_installed_updates.txt" "!_DirWork!\GeneralSystemInfo\_honeywell_installed_updates.txt"
)
if exist "%windir%\honeywell_required_patches.txt" (
	call :logitem . get honeywell_required_patches.txt
	call :doCmd copy /y "%windir%\honeywell_required_patches.txt" "!_DirWork!\GeneralSystemInfo\_honeywell_required_patches.txt"
)
if exist "%HwInstallPath%\Experion PKS\Install\honeywell_required_patches.log" (
	call :logitem . get honeywell_required_patches.log
	call :doCmd copy /y "%HwInstallPath%\\Experion PKS\Install\honeywell_required_patches.log" "!_DirWork!\GeneralSystemInfo\_honeywell_required_patches.log"
)

call :logitem . WindowsUpdate.log
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :doCmd copy /y "%windir%\WindowsUpdate.log" "!_DirWork!\GeneralSystemInfo\_WindowsUpdate.log"
:: ETL logs collection - skiped for now
::if exist %windir%\Logs\WindowsUpdate (
::	call :logitem . get Windows Update ETL Logs
::	call :mkNewDir  !_DirWork!\GeneralSystemInfo\_WindowsUpdateEtlLogs
::	call :doCmd copy /y "%windir%\Logs\WindowsUpdate\*.etl" "!_DirWork!\GeneralSystemInfo\_WindowsUpdateEtlLogs\"
::)

:WmiConfiguration
call :logitem . WMI Configuration
::WmiRootSecurityDescriptor
call :LogWmicCmd !_DirWork!\GeneralSystemInfo\_WmiRootSecurityDescriptor.txt wmic /namespace:\\root path __systemsecurity call GetSecurityDescriptor
:: WMI Provider Host Quota Configuration
call :LogWmicCmd !_DirWork!\GeneralSystemInfo\_WmiRootSecurityDescriptor.txt wmic /namespace:\\root path __ProviderHostQuotaConfiguration

:: McAfee on accesss scanner settings
reg query "HKLM\SOFTWARE\Wow6432Node\McAfee\SystemCore\VSCore\On Access Scanner" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
	call :logitem . McAfee On Access Scanner - reg settings
	call :mkNewDir  !_DirWork!\RegistryInfo
	call :doCmd REG EXPORT "HKLM\SOFTWARE\Wow6432Node\McAfee\SystemCore\VSCore" "!_DirWork!\RegistryInfo\_HKLM_McAfee_OnAccessScanner.txt"
) else (
	if exist "%SystemRoot%\System32\drivers\mfe*.*" (
		set _RegFile="!_DirWork!\RegistryInfo\_HKLM_McAfee_OnAccessScanner.txt"
		call :mkNewDir  !_DirWork!\RegistryInfo
		call :GetReg QUERY "HKLM\SOFTWARE"
		call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node"
	)
)
:: ToDo: test - where are the 'McAfee Agent' registry settings


:: Symantec\Symantec Endpoint Protection\AV
reg query "HKLM\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\AV" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
	call :logitem . SEP / AV - reg settings
	call :mkNewDir  !_DirWork!\RegistryInfo
	set _RegFile="!_DirWork!\RegistryInfo\_Symantec_SEP_AV.txt"
	call :InitLog !_RegFile!
	call :GetReg QUERY "HKLM\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\AV" /s
)

call :logOnlyItem . secedit /export /cfg
set _SecurityFile=!_DirWork!\GeneralSystemInfo\_SecurityCfg.txt
call :InitLog !_SecurityFile!
secedit /export /cfg !_SecurityFile! >> "!_LogFile!"
call :SleepX 1

::call :logOnlyItem . query drivers information
::call :LogCmd !_DirWork!\GeneralSystemInfo\_driverquery.output.csv driverquery /fo csv /v

call :logItem . reg query RPC settings
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_RPC_registry_settings.txt"
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\Software\Microsoft\Rpc" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\RpcEptMapper" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\RpcLocator" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\RpcSs" /s
call :GetReg QUERY "HKLM\Software\Policies\Microsoft\Windows NT\Rpc" /s


call :logItem . reg query Microsoft\OLE
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_OLE_registry_settings.txt"
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\Software\Microsoft\OLE" /s

call :logOnlyItem . reg query Graphics Drivers
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_GraphicsDrivers.txt"
call :InitLog !_RegFile!
call :GetReg QUERY  "HKCU\Control Panel\Desktop" /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg QUERY  "HKCU\Control Panel\Desktop\PerMonitorSettings" /s
call :GetReg QUERY  "HKLM\System\CurrentControlSet\Control\GraphicsDrivers" /s
call :GetReg QUERY  "HKLM\SYSTEM\CurrentControlSet\Control\Video" /s /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg QUERY  "HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\UnitedVideo" /s /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE

call :logOnlyItem . reg query services
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_HKLM_Services.txt"
call :InitLog !_RegFile!
call :GetReg QUERY  "HKLM\SYSTEM\CurrentControlSet\Services" /s /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE

@if defined _Debug ( echo [#Debug#] !time!: reg query misc)
call :logOnlyItem . reg query misc
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_reg_query_misc.txt"
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows" /v CEIPEnable
call :GetReg QUERY "HKLM\Software\Policies\Microsoft\Windows Defender" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\Session Manager\Memory Management" /s /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem" /s
:: Art.No: 000102530 - Disable "Updates available" notifications popping up on the operator screen when using WSUS
call :GetReg QUERY "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /s
:: AllowCortana in Windows Search
call :GetReg QUERY "HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /s
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\MusNotification.exe"
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\MusNotificationUx.exe"
:: acronis registry settings
reg query "HKLM\SOFTWARE\Wow6432Node\Acronis" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Acronis"
)
reg query "HKLM\SOFTWARE\Acronis" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
call :GetReg QUERY "HKLM\SOFTWARE\Acronis"
)
call :GetReg QUERY "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main" /v UseSWRender
:: security-enhanced channel protocols
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL" /s
:: KSM2020-035 (Art.No: 000115926) Microsoft Defect in IE11 causing high CPU on Experion R50x and R51x Station processes resolved in September’20 (or later) roll up
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Policies\Microsoft\FeatureManagement" /s
call :GetReg QUERY "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System" /s
:: Solution to enable RSLinx to be used with Experion R51x - reg key: .\Rockwell Software\FactoryTalk Diagnostics
reg query "HKLM\SOFTWARE\Wow6432Node\Rockwell Software\FactoryTalk Diagnostics" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
	call :logitem . Rockwell Software\FactoryTalk Diagnostics - reg settings
	call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Rockwell Software\FactoryTalk Diagnostics"
)


@if defined _Debug ( echo [#Debug#] !time!: Windows Time status/settings)
call :logItem . Windows Time status/settings
set _WindowsTimeFile="!_DirWork!\GeneralSystemInfo\_WindowsTime.txt"
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :InitLog !_WindowsTimeFile!
call :LogCmd !_WindowsTimeFile! w32tm /query /status /verbose
call :LogCmd !_WindowsTimeFile! w32tm /query /configuration
set _RegFile=!_WindowsTimeFile!
call :GetReg QUERY  "HKLM\SYSTEM\CurrentControlSet\Services\W32Time" /s  /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE

:: temperature
@if defined _Debug ( echo [#Debug#] !time!: _VEP=!_VEP!)
if not "!_VEP!"=="1" (
	wmic /namespace:\\root\wmi PATH MSAcpi_ThermalZoneTemperature get Active,CriticalTripPoint,CurrentTemperature 2>NUL | find/i "CurrentTemperature" >NUL 2>&1
	if !errorlevel!==0 (
		call :logItem . get Windows Thermal Zone Temperature information
		call :mkNewDir  !_DirWork!\GeneralSystemInfo
		set _ThermalZoneTemperature="!_DirWork!\GeneralSystemInfo\_ThermalZoneTemperature.txt"
		call :InitLog !_ThermalZoneTemperature!
		@echo Temperature at thermal zone in tenths of degrees Kelvin >>!_ThermalZoneTemperature!
		@echo Convert to Celsius: xxx / 10 - 273.15 >>!_ThermalZoneTemperature!
		@echo.>>!_ThermalZoneTemperature!
		wmic /namespace:\\root\wmi PATH MSAcpi_ThermalZoneTemperature get Active,CriticalTripPoint,CurrentTemperature>>!_ThermalZoneTemperature!
	)
)

:: get GDI Handles Count
@if defined _Debug ( echo [#Debug#] !time!: getGDIHandlesCount)
call :getGDIHandlesCount

:localgroups
call :logitem . get members of Experion groups
set _localgroups=!_DirWork!\GeneralSystemInfo\_localgroupsExperion.txt
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :InitLog !_localgroups!
call :LogCmd !_localgroups! net localgroup "Local Servers"
call :LogCmd !_localgroups! net localgroup "Product Administrators"
call :LogCmd !_localgroups! net localgroup "Local Ack View Only Users"
call :LogCmd !_localgroups! net localgroup "Local Engineers"
call :LogCmd !_localgroups! net localgroup "Local Operators"
call :LogCmd !_localgroups! net localgroup "Local SecureComms Administrators"
call :LogCmd !_localgroups! net localgroup "Local Supervisors"
call :LogCmd !_localgroups! net localgroup "Local View Only Users"
call :LogCmd !_localgroups! net localgroup "Distributed COM Users"
if "!_isServer!"=="1" (
	call :LogCmd !_localgroups! net localgroup "Local DSA Connections"
)

call :logitem . get HKEY_USERS Reg Values
set _RegFile="!_DirWork!\RegistryInfo\_HKEY_USERS.txt"
call :mkNewDir !_DirWork!\RegistryInfo
if exist !_RegFile! call :doit del "!_RegFile!"
call :GetHkeyUsersRegValues

:: get mngr account information - Local Group Memberships
call :logitem . get mngr account information - Local Group Memberships
set _mngrUserAccount=!_DirWork!\GeneralSystemInfo\_mngrUserAccount.txt
call :InitLog !_mngrUserAccount!
call :LogCmd !_mngrUserAccount! net user mngr

:: MiniFilter drivers
call :logitem . MiniFilter drivers
set _fltmc=!_DirWork!\GeneralSystemInfo\_fltmc.output.txt
call :InitLog !_fltmc!
call :LogCmd !_fltmc! FLTMC Filters
call :LogCmd !_fltmc! fltmc Instances
call :LogCmd !_fltmc! fltmc Volumes

:: Experion ACL Verify
call :ExperionAclVerify

:: diskdrive status
if not "!_VEP!"=="1" (
	call :logitem . diskdrive status
	call :mkNewDir  !_DirWork!\GeneralSystemInfo
	set _diskdrive=!_DirWork!\GeneralSystemInfo\_diskdrive.txt
	call :InitLog !_diskdrive!
	@echo wmic diskdrive get InterfaceType,MediaType,Model,Size,Status>>!_diskdrive!
	@echo The Status property will return "Pred Fail" if your drive's death is imminent, or "OK" if it thinks the drive is doing fine.>>!_diskdrive!
	@echo.>>!_diskdrive!
	wmic diskdrive get InterfaceType,MediaType,Model,Size,Status | more /s >>!_diskdrive!
)


call :logitem . tasklist /verbose
call :mkNewDir  !_DirWork!\GeneralSystemInfo
tasklist /v /fo csv >!_DirWork!\GeneralSystemInfo\_TaskList.csv

call :logitem . cmd query output
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :LogCmd !_DirWork!\GeneralSystemInfo\_query.output.txt QUERY USER
call :LogCmd !_DirWork!\GeneralSystemInfo\_query.output.txt QUERY SESSION
call :LogCmd !_DirWork!\GeneralSystemInfo\_query.output.txt QUERY TERMSERVER
call :LogCmd !_DirWork!\GeneralSystemInfo\_query.output.txt QUERY PROCESS


:: BranchCache
call :logitem . collecting branch cache status and settings
set _BranchcacheFile=!_DirWork!\GeneralSystemInfo\_BranchCache.txt
call :InitLog !_BranchcacheFile!
call :logCmd !_BranchcacheFile! netsh branchcache show hostedcache
call :logCmd !_BranchcacheFile! netsh branchcache show localcache		
call :logCmd !_BranchcacheFile! netsh branchcache show publicationcache
call :logCmd !_BranchcacheFile! netsh branchcache show status all
call :logCmd !_BranchcacheFile! netsh branchcache smb show latency
if !_OSVER3! GEQ 9200 ( call :logitem . fetching Branchcache infos using PowerShell
						call :logCmd !_BranchcacheFile! PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -command "Get-BCStatus")
call :logCmd !_BranchcacheFile! bitsadmin /list /AllUsers /verbose
call :logCmd !_BranchcacheFile! bitsadmin /util /version /verbose
call :logCmd !_BranchcacheFile! bitsadmin /PEERS /LIST
call :logCmd !_BranchcacheFile! DIR /A/B/S %windir%\ServiceProfiles\NetworkService\AppData\Local\PeerDistpub
call :logCmd !_BranchcacheFile! DIR /A/B/S %windir%\ServiceProfiles\NetworkService\AppData\Local\PeerDistRepub


:: additional Windows Event Logs
call :logitem . export additional Windows Event Logs
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :export-evtx Microsoft-Windows-TerminalServices-LocalSessionManager/Operational !_DirWork!\GeneralSystemInfo\
call :export-evtx Microsoft-Windows-TerminalServices-LocalSessionManager/Admin !_DirWork!\GeneralSystemInfo\
call :export-evtx Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational !_DirWork!\GeneralSystemInfo\
call :export-evtx Microsoft-Windows-TerminalServices-RemoteConnectionManager/Admin !_DirWork!\GeneralSystemInfo\
call :export-evtx Microsoft-Windows-TaskScheduler/Operational !_DirWork!\GeneralSystemInfo\
call :export-evtx "Microsoft-Windows-Windows Firewall With Advanced Security/Firewall" !_DirWork!\GeneralSystemInfo\
call :export-evtx "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin" !_DirWork!\GeneralSystemInfo\
:: System & Application event logs (errors & warnings) to csv
PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -Command "& {Get-WinEvent -LogName System -FilterXPath '*[System[(Level=1  or Level=2 or Level=3)]]' | select -Property TimeCreated,Id,LevelDisplayName,ProviderName,Message | Export-Csv '!_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_System.csv' -NoTypeInformation}"
PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -Command "& {Get-WinEvent -LogName Application -FilterXPath '*[System[(Level=1  or Level=2 or Level=3)]]' | select -Property TimeCreated,Id,LevelDisplayName,ProviderName,Message | Export-Csv '!_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_Application.csv' -NoTypeInformation}"



:: MSPower_DeviceEnable
call :logItem . MSPower_DeviceEnable query
call :mkNewDir  !_DirWork!\GeneralSystemInfo
set _MSPower_DeviceEnable=!_DirWork!\GeneralSystemInfo\_MSPower_DeviceEnable.txt
call :InitLog !_MSPower_DeviceEnable!
@echo.>>!_MSPower_DeviceEnable!
call :LogWmicCmd !_MSPower_DeviceEnable! wmic /namespace:\\root\wmi PATH MSPower_DeviceEnable get Active,Enable,InstanceName

:region PendingRebot
call :logitem . check pending reboot
set _outFile="!_DirWork!\GeneralSystemInfo\_PendingRebot.txt"
set _RegFile=!_outFile!
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :InitLog !_outFile!
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing" /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager" /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Services\Netlogon"
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName" /s

call :logLine "!_outFile!"
@echo ===== %time% :PS: Invoke-WmiMethod -Namespace root\ccm\clientsdk -Class CCM_ClientUtilities -Name DetermineIfRebootPending>> "!_outFile!"
call :logLine "!_outFile!"
PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -Command "& {@(try{Invoke-WmiMethod -Namespace root\ccm\clientsdk -Class CCM_ClientUtilities -Name DetermineIfRebootPending -ea stop | select IsHardRebootPending,RebootPending,ReturnValue} catch{$_.Exception.Message} ) | out-file '!_outFile!' -Append -Encoding ascii}"

@echo. >> "!_outFile!"
call :logLine "!_outFile!"
@echo ===== %time% : WMIC.EXE /NAMESPACE:\\root\ccm\clientsdk PATH CCM_ClientUtilities call DetermineIfRebootPending>> "!_outFile!"
call :logLine "!_outFile!"
WMIC.EXE /NAMESPACE:\\root\ccm\clientsdk PATH CCM_ClientUtilities call DetermineIfRebootPending 2>>&1| more /s | find /v "" >>!_outFile! 2>>&1
:endregion PendingRebot

:: DISM Check Health
if "%_OSVER1%" GTR "6" (
	call :logItem . DISM Check Health
	Dism /Online /Cleanup-Image /CheckHealth | find /i "No component store corruption detected" >NUL 2>&1
	if /i "%errorlevel%" NEQ "0" (
		call :mkNewDir  !_DirWork!\GeneralSystemInfo
		call :InitLog !_DirWork!\GeneralSystemInfo\_DISM_CheckHealth.txt
		call :logCmd !_DirWork!\GeneralSystemInfo\_DISM_CheckHealth.txt Dism /Online /Cleanup-Image /CheckHealth
	)
)

:: vssadmin list report
call :logitem . vssadmin list report
set _VssAdminListReport=!_DirWork!\GeneralSystemInfo\_VssAdminListReport.txt
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :InitLog !_VssAdminListReport!
call :logCmd !_VssAdminListReport! vssadmin List Providers
call :logCmd !_VssAdminListReport! vssadmin List Shadows
call :logCmd !_VssAdminListReport! vssadmin List ShadowStorage
call :logCmd !_VssAdminListReport! vssadmin List Volumes
call :logCmd !_VssAdminListReport! vssadmin List Writers
call :LogWmicCmd !_VssAdminListReport! wmic shadowcopy

:: omreport
where omreport >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . omreport info
	call :mkNewDir !_DirWork!\GeneralSystemInfo
	call :InitLog !_DirWork!\GeneralSystemInfo\_omreport.txt
	
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport  about details=true
	
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system operatingsystem
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system summary
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system version
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system thrmshutdown
	
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system alertaction
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system alertlog
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system esmlog
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system cmdlog
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport system postlog
	
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis info
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis firmware
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis leds
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis bios
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis biossetup

	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis pwrmonitoring
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis pwrmanagement
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis pwrsupplies
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis acswitch
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis fans
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis fancontrol
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis memory
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis nics
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis processors
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis temps
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis volts
	rem This command is no longer available through Server Administrator!!   call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis currents
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport chassis batteries
	
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport storage pdisk controller=0
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport storage vdisk
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport storage controller
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport storage battery
	call :logCmd !_DirWork!\GeneralSystemInfo\_omreport.txt omreport storage connector controller=0
)

:: logon scripts (operators)
call :mkNewDir  !_DirWork!\GeneralSystemInfo
if exist %windir%\SYSVOL\domain\scripts\*.bat set _domainScripts=1
if exist %windir%\SYSVOL\domain\scripts\*.cmd set _domainScripts=1
if defined _domainScripts (
	call :logOnlyItem operator logon scripts - %windir%\SYSVOL\domain\scripts\
	call :mkNewDir  !_DirWork!\GeneralSystemInfo\_scripts
	call :doCmd copy /y %windir%\SYSVOL\domain\scripts\*.bat "!_DirWork!\GeneralSystemInfo\_scripts\"
	call :doCmd copy /y %windir%\SYSVOL\domain\scripts\*.cmd "!_DirWork!\GeneralSystemInfo\_scripts\"
)
if exist %windir%\System32\repl\import\scripts\*.bat set _replImportScripts=1
if exist %windir%\System32\repl\import\scripts\*.cmd set _replImportScripts=1
if defined _replImportScripts (
	call :logOnlyItem operator logon scripts - %windir%\System32\repl\import\scripts\
	call :mkNewDir  !_DirWork!\GeneralSystemInfo\_scripts
	call :doCmd copy /y %windir%\System32\repl\import\scripts\*.bat "!_DirWork!\GeneralSystemInfo\_scripts\"
	call :doCmd copy /y %windir%\System32\repl\import\scripts\*.cmd "!_DirWork!\GeneralSystemInfo\_scripts\"
)

:: hdd defrag analysis
if NOT "!_VEP!"=="1" (
	call :logItem . hdd defrag analysis
	call :mkNewDir  !_DirWork!\GeneralSystemInfo
	set _defrag_analysis="!_DirWork!\GeneralSystemInfo\_defrag.analysis.txt"
	call :InitLog !_defrag_analysis!
	call :logCmd !_defrag_analysis! defrag /c /a /v
)

call :logOnlyItem . reg query HKCU\Control Panel
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_HKCU_Control_Panel.txt"
call :InitLog !_RegFile!
call :GetReg_async !_RegFile! "HKCU\Control Panel" /s /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE

call :logOnlyItem . reg query Internet Settings
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_InternetSettings.txt"
call :InitLog !_RegFile!
call :GetReg_async !_RegFile! "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /s /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg_async !_RegFile! "HKEY_USERS\S-1-5-18\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /s
call :GetReg_async !_RegFile! "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings" /s

call :logOnlyItem . reg query Installed Components
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_InstalledComponents.txt"
call :InitLog !_RegFile!
call :GetReg_async !_RegFile! "HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components" /s

call :logOnlyItem . Windows Policies
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile="!_DirWork!\RegistryInfo\_Windows_Policies.txt"
for %%h in ("HKLM\Software\Policies\Microsoft" "HKLM\SYSTEM\CurrentControlSet\Policies" "HKCU\Software\Policies") do ( call :GetReg_async !_RegFile! %%h /s)		


call :logitem . fetching environment Variables
call :InitLog !_DirWork!\GeneralSystemInfo\_EnvVariables.txt
call :LogCmd !_DirWork!\GeneralSystemInfo\_EnvVariables.txt set


:: next

goto :eof
:endregion Windows

:region Network data
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:NetworkAddData
call :logitem * Network configuration data *

call :mkNewDir !_DirWork!\_Network

call :logitem . netstat connections
set _NetStatFile=!_DirWork!\_Network\netstat-nato.txt
call :InitLog !_NetStatFile!
call :LogCmd !_NetStatFile! netstat -nato

call :logitem . ipconfig output
call :InitLog !_DirWork!\_Network\ipconfig.txt
call :LogCmd !_DirWork!\_Network\ipconfig.txt ipconfig /all
call :InitLog !_DirWork!\_Network\ipconfig.displaydns.txt
call :LogCmd !_DirWork!\_Network\ipconfig.displaydns.txt ipconfig /displaydns

:NetSh_ConfigAndStats
call :logitem . netsh config/stats
set _NetShFile=!_DirWork!\_Network\netsh.output.txt
call :InitLog !_NetShFile!
call :LogCmd !_NetShFile! netsh int tcp show global
call :LogCmd !_NetShFile! netsh int tcp show heuristics
call :LogCmd !_NetShFile! netsh int IP show config
call :LogCmd !_NetShFile! netsh int ipv4 show dynamicport tcp
call :LogCmd !_NetShFile! netsh int ipv4 show offload
call :LogCmd !_NetShFile! netsh int ipv4 show addresses
call :LogCmd !_NetShFile! netsh int ipv4 show ipstats
call :LogCmd !_NetShFile! netsh int ipv4 show udpstats
call :LogCmd !_NetShFile! netsh int ipv4 show tcpstats
call :LogCmd !_NetShFile! netsh http show urlacl
call :LogCmd !_NetShFile! netsh http show servicestate

:nslookup
call :logitem . nslookup
call :SleepX 1
:: NS LookUp - Forward
echo ======================================== > "!_DirWork!\_Network\nslookup.txt"
echo %date% %time% >> "!_DirWork!\_Network\nslookup.txt"
echo cmd: nslookup %computername% >> "!_DirWork!\_Network\nslookup.txt"
nslookup %computername% >> "!_DirWork!\_Network\nslookup.txt"  2>>&1
:: NS LookUp - Reverse
set ip_address_string="IPv4 Address"
for /f "usebackq tokens=2 delims=:" %%f in (`ipconfig ^| findstr /c:%ip_address_string%`) do (
	echo ======================================== >> "!_DirWork!\_Network\nslookup.txt"
	echo %date% %time% >> "!_DirWork!\_Network\nslookup.txt"
    echo Your IP Address is: %%f  >> "!_DirWork!\_Network\nslookup.txt"
	echo cmd: nslookup %%f >> "!_DirWork!\_Network\nslookup.txt"
	nslookup %%f >> "!_DirWork!\_Network\nslookup.txt"  2>>&1
)
echo ======================================== >> "!_DirWork!\_Network\nslookup.txt"
echo %date% %time% - done>> "!_DirWork!\_Network\nslookup.txt"

call :logitem . route / arp
call :InitLog !_DirWork!\_Network\arp.txt
call :LogCmd !_DirWork!\_Network\arp.txt  arp -a -v
call :InitLog !_DirWork!\_Network\route.print.txt
call :LogCmd !_DirWork!\_Network\route.print.txt route print

call :logitem . nbtstat
call :InitLog !_DirWork!\_Network\nbtstat.txt
call :LogCmd !_DirWork!\_Network\nbtstat.txt  nbtstat -n
call :LogCmd !_DirWork!\_Network\nbtstat.txt  nbtstat -c

:advfirewall
call :logitem . firewall rules
call :InitLog !_DirWork!\_Network\firewall_rules.txt
call :LogCmd !_DirWork!\_Network\firewall_rules.txt netsh advfirewall monitor show firewall verbose
if exist %windir%\system32\LogFiles\Firewall\pfirewall.log call :doCmd Copy /y %windir%\system32\LogFiles\Firewall\pfirewall.log "!_DirWork!\_Network\_pFirewall.log"

call :logitem . net commands
set _NetCmdFile=!_DirWork!\_Network\netcmd.txt
call :InitLog !_NetCmdFile!
call :LogCmd !_NetCmdFile! NET SHARE
if !_PSVer! GEQ 5 (
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { Get-SmbShare | select Name, ScopeName, FolderEnumerationMode, Path, Description |ft|out-string -width 200 }}" >>!_NetCmdFile!
)
call :LogCmd !_NetCmdFile! NET START
call :LogCmd !_NetCmdFile! NET CONFIG SERVER
call :LogCmd !_NetCmdFile! NET SESSION
call :LogCmd !_NetCmdFile! NET USER
call :LogCmd !_NetCmdFile! NET USE
call :LogCmd !_NetCmdFile! NET ACCOUNTS
call :LogCmd !_NetCmdFile! NET LOCALGROUP
call :LogCmd !_NetCmdFile! NET CONFIG WKSTA
call :LogCmd !_NetCmdFile! NET STATISTICS Workstation
call :LogCmd !_NetCmdFile! NET STATISTICS SERVER
call :LogCmd !_NetCmdFile! NET CONFIG RDR

set _RegFile="!_DirWork!\_Network\TcpIpParameters.txt"
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Services\FTEMUXMP" /s /t REG_SZ,REG_MULTI_SZ,REG_EXPAND_SZ,REG_DWORD,REG_QWORD,REG_NONE
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp\Parameters" /v ArpRetryCount
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\Tcpip6" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\tcpipreg" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\iphlpsvc" /s
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Network" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}" /s

call :logitem . wmic nic/nicconfig get
set _NicConfig=!_DirWork!\_Network\nicconfig.txt
call :InitLog !_NicConfig!
call :LogWmicCmd !_NicConfig! wmic nic get
call :LogWmicCmd !_NicConfig! wmic nicconfig get Description,DHCPEnabled,Index,InterfaceIndex,IPEnabled,MACAddress,SettingID
call :LogWmicCmd !_NicConfig! wmic nicconfig get


call :logitem . msft_providers
call :InitLog !_DirWork!\_Network\msft_providers.txt
call :LogWmicCmd !_DirWork!\_Network\msft_providers.txt wmic path msft_providers get Provider,HostProcessIdentifier

if exist "c:\Program Files\VMware\VMware Tools\VMwareToolboxCmd.exe" (
	set _vmtbcmd="C:\Program Files\VMware\VMware Tools\VMwareToolboxCmd.exe"
	call :logItem . VMWare status
	call :mkNewDir  !_DirWork!\_Network
	set _vmStatLog=!_DirWork!\_Network\VMware.stat.txt
	call :InitLog !_vmStatLog!
	call :LogCmd !_vmStatLog! !_vmtbcmd!  upgrade status
	call :LogCmd !_vmStatLog! !_vmtbcmd!  timesync status
	:: vSphere API/SDK Documentation - vSphere Guest and HA Application Monitoring SDK Documentation - Guest and HA Application Monitoring SDK Programming Guide - Tools for Extended Guest Statistics - Metrics Examples
	for /f "tokens=*" %%g in ('!_vmtbcmd! stat raw') do (
		call :logcmd !_vmStatLog! !_vmtbcmd! stat raw text %%g
	)
)

call :logitem . get NetTcpPortSharing config file
for /f "usebackq skip=1" %%h in (`wmic service where "name='NetTcpPortSharing'" get PathName`) do (
	if exist "%%h.config" (
		call :doCmd copy /y "%%h.config" "!_DirWork!\_Network\"
		)
	)

call :logitem . get clientaccesspolicy.xml for Silverlight
call :mkNewDir  "!_DirWork!\_Network"
::PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -Command "& {@(try{(Invoke-WebRequest -Uri http://localhost/clientaccesspolicy.xml -UseBasicParsing).Content} catch{$_.Exception.Message}) | out-file '!_DirWork!\_Network\clientaccesspolicy.xml' }"
PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -Command "& {@(try{(New-Object System.Net.WebClient).DownloadString('http://localhost/clientaccesspolicy.xml')} catch{$_.Exception.Message} ) | out-file '!_DirWork!\_Network\clientaccesspolicy.xml' }"
call :SleepX 1


call :logitem . get FTE config files
call :mkNewDir  "!_DirWork!\FTELogs"
call :doCmd copy /y "%HwProgramData%\ProductConfig\FTE\FTEinstall.inf" "!_DirWork!\FTELogs\_FTEinstall.inf"
call :doCmd copy /y "%HwProgramData%\ProductConfig\FTE\fteconfig.inf" "!_DirWork!\FTELogs\_fteConfig.inf"

goto :eof
:endregion NetworkAddData

:region Experion data
:ExperionAddData
call :logitem * Experion data *

::CrashDumps - file list & reg settings
call :logitem . create crash dumps list
call :mkNewDir !_DirWork!\CrashDumps
set _CrashDumpsList=!_DirWork!\CrashDumps\_CrashDumpsList.txt
call :InitLog !_CrashDumpsList!
::call :logCmd  !_CrashDumpsList! dir %windir%\memory.dmp
::call :logCmd  !_CrashDumpsList! dir /o:-d %windir%\Minidump\
::call :logCmd  !_CrashDumpsList! dir /o:-d "%HwProgramData%\Experion PKS\CrashDump"
::call :logCmd  !_CrashDumpsList! dir /o:-d "%HwProgramData%\HMIWebLog\DumpFiles"
::call :logCmd  !_CrashDumpsList! dir /o:-d "%HwProgramData%\Experion PKS\server\data\*.dmp"
::call :logCmd  !_CrashDumpsList! dir  /o-d /s c:\users\*.dmp
::call :logCmd  !_CrashDumpsList! dir  /o-d /s %windir%\System32\config\systemprofile\AppData\Local\CrashDumps\*.dmp
::call :logCmd  !_CrashDumpsList! dir  /o-d /s %windir%\SysWOW64\config\systemprofile\AppData\Local\CrashDumps\*.dmp
::call :logCmd  !_CrashDumpsList! dir  /o-d /s %windir%\ServiceProfiles\*.dmp
::call :logCmd  !_CrashDumpsList! dir  /o-d /s %windir%\LiveKernelReports\*.dmp

START "Crash dump files list" /MIN CMD /C "dir /o-d /s %SystemDrive%\*.dmp >> "!_CrashDumpsList!"" 2>&1
call :SleepX 3

if NOT "!_VEP!"=="1" (
	where filfrag >NUL 2>&1
	if %errorlevel%==0 (
		call :logitem . filfrag output
		call :mkNewDir !_DirWork!\ServerDataDirectory
		call :InitLog !_DirWork!\ServerDataDirectory\_filfrag.output.txt
		call :logCmd !_DirWork!\ServerDataDirectory\_filfrag.output.txt filfrag
	)
)

if exist "%HwProgramData%\Experion PKS\Server\data\mapping\tps.xml" (
	call :logitem . copy .\mapping\tps.xml
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\mapping\tps.xml" "!_DirWork!\ServerDataDirectory\_mapping.tps.xml"
)
if exist "%HwProgramData%\Experion PKS\Server\data\mapping\" (
	call :logItem . get mapping xml files
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :mkNewDir !_DirWork!\ServerDataDirectory\_mapping
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '%HwProgramData%\Experion PKS\Server\data\mapping\' -filt *.xml | foreach{copy $_.fullName -dest '!_DirWork!\ServerDataDirectory\_mapping'; sleep -Milliseconds 250} }}
	@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at get mapping xml files with PowerShell'. )
	if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'get mapping xml files with PowerShell' failed.)
)

if exist "%HwProgramData%\Experion PKS\Server\data\scripts\" (
	call :logItem . get server scripts xml files
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :mkNewDir !_DirWork!\ServerDataDirectory\_scripts
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci '%HwProgramData%\Experion PKS\Server\data\scripts\' -filt *.xml | foreach{copy $_.fullName -dest '!_DirWork!\ServerDataDirectory\_scripts'; sleep -Milliseconds 100} }}
	@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at get server scripts xml files with PowerShell'. )
	if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'get server scripts xml files with PowerShell' failed.)
)


if defined _isServer (
	where lisscn >NUL 2>&1
	if %errorlevel%==0 (
		call :logitem . lisscn output
		call :mkNewDir !_DirWork!\ServerDataDirectory
		for /f "tokens=2 delims= " %%h in ('hdwbckbld ^| find /i "DEF CHN"') do (
			@set _CHN=%%h
			REM remove CHN
			@set _CHN=!_CHN:CHN=!
			REM remove leading zeroes
			REM set /a num=1000%_x% %% 1000
			FOR /F "tokens=* delims=0" %%A IN ("!_CHN!") DO @set _CHN=%%A
			@set /a _CHN=!_CHN:CHN=!
			call :LogCmdNoSleep "!_DirWork!\ServerDataDirectory\_lisscn.txt" lisscn -CHN !_CHN!
			call :LogCmdNoSleep "!_DirWork!\ServerDataDirectory\_lisscn_all.txt" lisscn -CHN !_CHN! -all_ref
		)
		call :SleepX 1
	)

	where dsasublist >NUL 2>&1
	if %errorlevel%==0 (
		call :logitem . dsasublist
		call :mkNewDir !_DirWork!\ServerRunDirectory
		call :InitLog !_DirWork!\ServerRunDirectory\_dsasublist.txt
		call :logCmd !_DirWork!\ServerRunDirectory\_dsasublist.txt dsasublist
	)
)


call :logitem . crash control registry settings
set _RegFile="!_DirWork!\CrashDumps\_RegCrashControl.txt"
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\CrashControl" /s
call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\AeDebug" /s
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug" /s
call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\Windows Error Reporting" /s
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting" /s
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options" /s

call :logitem . Recover OS settings
call :DoCmd  wmic /output:"!_DirWork!\CrashDumps\_recoveros.txt" RECOVEROS

where dual_status >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . get server dual status
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\_dual_status.output.txt
    call :logCmd !_DirWork!\ServerRunDirectory\_dual_status.output.txt dual_status
)

where cstn_status >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . get console station status
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\_cstn_status.output.txt
    call :logCmd !_DirWork!\ServerRunDirectory\_cstn_status.output.txt cstn_status
)

where ps >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . ps output
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\_ps.output.txt
    call :logCmd !_DirWork!\ServerRunDirectory\_ps.output.txt ps
)

where shheap >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . Experion shheap 1 check
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :InitLog !_DirWork!\ServerDataDirectory\_shheap.1.check.output.txt
    call :logCmd !_DirWork!\ServerDataDirectory\_shheap.1.check.output.txt shheap 1 check

	call :logitem . Experion shheap 2 dump (dual_q heap)
    call :InitLog !_DirWork!\ServerDataDirectory\_shheap.2.dump.output.txt
    call :logCmd  !_DirWork!\ServerDataDirectory\_shheap.2.dump.output.txt shheap 2 dump
)

call :logitem . list disk resident heap files
call :mkNewDir !_DirWork!\ServerDataDirectory
if exist "%HwProgramData%\Experion PKS\Server\data\locks" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\locks"
)
if exist "%HwProgramData%\Experion PKS\Server\data\dual*q" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\dual*q"
)
if exist "%HwProgramData%\Experion PKS\Server\data\gda" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\gda"
)
if exist "%HwProgramData%\Experion PKS\Server\data\tagcache" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\tagcache"
)
if exist "%HwProgramData%\Experion PKS\Server\data\taskrequest" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\taskrequest"
)
if exist "%HwProgramData%\Experion PKS\Server\data\dbrepsrvup*" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\dbrepsrvup*"
)
if exist "%HwProgramData%\Experion PKS\Server\data\pntxmt_q*" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\pntxmt_q*"
)
if exist "%HwProgramData%\Experion PKS\Server\data\bacnetheap*" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\bacnetheap*"
)
if exist "%HwProgramData%\Experion PKS\Server\data\shheap*" (
	call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\shheap*"
)


if exist "%HwProgramData%\Experion PKS\Client\Station\station.ini" (
	call :logitem . copy station.ini file
	call :mkNewDir !_DirWork!\Station-logs
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Client\Station\station.ini" "!_DirWork!\Station-logs\_station.ini"
)

if exist "%HwProgramData%\HMIWebLog\Log.txt" (
	call :logitem . collect station configuration files
	call :getStationFiles
)

@if defined _Debug ( echo [#Debug#] !time!: _isTPS=!_isTPS!)
if "!_isTPS!"=="1" (
	@if defined _Debug ( echo [#Debug#] !time!: _isServer=!_isServer!)
	if "!_isServer!"=="1" (
		call :logitem . chkem /tpsmappings
		call :mkNewDir !_DirWork!\ServerRunDirectory
		call :logCmd !_DirWork!\ServerRunDirectory\_tpsmappings.txt chkem /tpsmappings
	)
)

where bckbld >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . Backbuild history assignments
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd bckbld -np -nt -ng -ns -out "!_DirWork!\ServerDataDirectory\_hist_scada.txt"
    call :doCmd bckbld -np -nt -ng -ns -tag CDA -out "!_DirWork!\ServerDataDirectory\_hist_cda.txt"
    call :doCmd bckbld -np -nt -ng -ns -tag RDA -out "!_DirWork!\ServerDataDirectory\_hist_dsa.txt"
    call :doCmd bckbld -np -nt -ng -ns -tag PSA -out "!_DirWork!\ServerDataDirectory\_hist_psa.txt"
)

:: HKLM\SOFTWARE\classes\Hw...
if defined _isServer (
	call :logOnlyItem . reg query HKEY_LOCAL_MACHINE\SOFTWARE\classes\Hw...
	call :mkNewDir  !_DirWork!\RegistryInfo
	set _RegFile="!_DirWork!\RegistryInfo\_HKLM_SOFTWARE_Classes_Experion.txt"
	call :InitLog !_RegFile!
	call :GetReg QUERY "HKLM\SOFTWARE\classes\HwHsc.OPCServer" /s
	call :GetReg QUERY "HKLM\SOFTWARE\classes\HwHsc.OPCServer2" /s
	call :GetReg QUERY "HKLM\SOFTWARE\classes\HwHsc.OPCServer3" /s
	call :GetReg QUERY "HKLM\SOFTWARE\classes\HwHsc.OPCServer4" /s
	call :GetReg QUERY "HKLM\SOFTWARE\classes\HwHsc.OPCServer5" /s
)

:: check files in Abstract folder for Zone.Identifier stream data
call :logItem . check files in Abstract folder for Zone.Identifier stream data
call :mkNewDir  !_DirWork!\Station-logs
call :logCmd %temp%\_Zone.Identifier.txt dir/s/r "%HwProgramData%\Experion PKS\Client\Abstract"
find/i "Zone.Identifier:$DATA" %temp%\_Zone.Identifier.txt >NUL 2>&1
if %errorlevel%==0 (
	call :doCmd move /y %temp%\_Zone.Identifier.txt "!_DirWork!\Station-logs\"
) else (
	call :doit del %temp%\_Zone.Identifier.txt
)


:: get system station configuration files
call :logItem . get system station configuration files  (Factory.stn, etc.)
call :mkNewDir  !_DirWork!\Station-logs
call :doCmd copy /y "%HwInstallPath%\Experion PKS\Client\Station\Default.stn" "!_DirWork!\Station-logs\@Default.stn"
call :doCmd copy /y "%HwInstallPath%\Experion PKS\Client\Station\Factory.stn" "!_DirWork!\Station-logs\@Factory.stn"
call :doCmd copy /y "%HwInstallPath%\Experion PKS\Client\Station\PanelStation_Default.stn" "!_DirWork!\Station-logs\@PanelStation_Default.stn"

if !_EPKS_MajorRelease! LSS 430 (
	:: get system flags settings in pre R43x releases
	where fildmp >NUL 2>&1
	if %errorlevel%==0 (
		call :logitem . Experion System Flags Table output
		call :mkNewDir !_DirWork!\ServerDataDirectory
		call :doCmd fildmp -DUMP -FILE "!_DirWork!\ServerDataDirectory\sysflg.output.txt" -FILENUM 8 -RECORDS 1 -FORMAT HEX
	)
)

where plexus >NUL 2>&1
if %errorlevel%==0 (
	call :logItem . plexus -printconfig
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\_plexus.printconfig.txt
    call :logCmd !_DirWork!\ServerRunDirectory\_plexus.printconfig.txt plexus -printconfig
)

if exist "%HwProgramData%\Experion PKS\Server\data\flbkup.def" (
	call :logitem . copy flbkup.def file
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\flbkup.def" "!_DirWork!\ServerDataDirectory\_flbkup.def"
)

:: search read only files in history archives
call :logitem . search read only files in history archives
if "!_xOS!"=="64" (set _regEPKS=HKLM\SOFTWARE\Wow6432Node\Honeywell\Experion PKS Server) ELSE (set _regEPKS=HKLM\SOFTWARE\Honeywell\Experion PKS Server)
@set _ArchiveDirectory=
call :GetRegValue "!_regEPKS!" ArchiveDirectory _ArchiveDirectory
@set _EPKS_MajorRelease=
@if defined _ArchiveDirectory (
	set _HistoryRO=!_DirWork!\ServerDataDirectory\_HistoryRO.txt
	set _ArchiveDirectory=%_ArchiveDirectory:;=" "%
	for %%h in (!_ArchiveDirectory!) do (
		call :logOnlyItem . search read only files in %%h
		dir /s/a:r %%h >NUL 2>&1
		if [!errorlevel!] EQU [0] (
			if not exist !_HistoryRO! (
				call :mkNewDir !_DirWork!\ServerDataDirectory
				call :InitLog !_HistoryRO!
			)
			call :logcmd !_HistoryRO! dir /s/a:r %%h
		)
		
		call :logOnlyItem . search hidden files in %%h
		dir /ah/s/b %%h >NUL 2>&1
		if [!errorlevel!] EQU [0] (
			if not exist !_HistoryRO! (
				call :mkNewDir !_DirWork!\ServerDataDirectory
				call :InitLog !_HistoryRO!
			)
			call :logcmd !_HistoryRO! dir /ah/s/b %%h
		)
		
		call :logOnlyItem . search system files in %%h
		dir /as/s/b %%h >NUL 2>&1
		if [!errorlevel!] EQU [0] (
			if not exist !_HistoryRO! (
				call :mkNewDir !_DirWork!\ServerDataDirectory
				call :InitLog !_HistoryRO!
			)
			call :logcmd !_HistoryRO! dir /as/s/b %%h
		)
		
	)
)

if defined _isServer (
	call :logitem . fixduplicates -report
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog "!_DirWork!\ServerRunDirectory\_duplicates.txt"
	call :doCmd fixduplicates -report "!_DirWork!\ServerRunDirectory\_duplicates.txt"
	call :isEmpty "!_DirWork!\ServerRunDirectory\_duplicates.txt" _isEmptyFixDuplicates
	@if defined _Debug ( echo [#Debug#] !time!: isEmpty return - _isEmptyFixDuplicates: !_isEmptyFixDuplicates! )
	if "!_isEmptyFixDuplicates!"=="1" (
		@if defined _Debug ( echo [#Debug#] !time!: delete empty file "!_DirWork!\ServerRunDirectory\_duplicates.txt" )
		del "!_DirWork!\ServerRunDirectory\_duplicates.txt"
	)
)

where fildmp >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . Experion GDA Table output
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd fildmp -DUMP -FILE "!_DirWork!\ServerDataDirectory\_gdatbldump.txt" -FILENUM 5 -RECORDS 1,1001 -FORMAT HEX
)

call :logItem . list HPSInstall temp folder
call :mkNewDir  !_DirWork!\ServerSetupDirectory
call :logCmd !_DirWork!\ServerSetupDirectory\_HPSInstallAppDataLocalTemp.txt dir /-C/S C:\Users\HPSInstall\AppData\Local\Temp

where databld >NUL 2>&1
if %errorlevel%==0 (
	call :logItem . export Experion operator settings
	call :mkNewDir  !_DirWork!\ServerRunDirectory
	call :doCmd databld -export -def OPERATORS -out "!_DirWork!\ServerRunDirectory\_OPERATORS.xml"
	REM call :logCmd !_DirWork!\ServerRunDirectory\_OPERATORS.dmp.txt operdmp
)

call :logitem . find system/hidden files in HwProgramData
call :mkNewDir !_DirWork!\ServerDataDirectory
REM dir /ah/s/b %HwProgramData% >"!_DirWork!\ServerDataDirectory\_HwProgramDataHiddenFiles.txt" 2>NUL
PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { $files=Get-ChildItem -Hidden '%HwProgramData%' -Exclude Thumbs.db -Recurse; if($files) {$files | out-file '!_DirWork!\ServerDataDirectory\_HwProgramDataHiddenFiles.txt' }}}"
REM dir /as/s/b %HwProgramData% >"!_DirWork!\ServerDataDirectory\_HwProgramDataSystemFiles.txt" 2>NUL
PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -ScriptBlock { $files=Get-ChildItem -System '%HwProgramData%' -Exclude Thumbs.db -Recurse; if($files) {$files | out-file '!_DirWork!\ServerDataDirectory\_HwProgramDataSystemFiles.txt' }}}"

if exist "%HwProgramData%\Quick Builder\LogFiles\Default\" (
	call :logitem . Quick Builder Log Files
	call :mkNewDir !_DirWork!\SloggerLogs
	call :mkNewDir !_DirWork!\SloggerLogs\_QBLogFiles
	call :mkNewDir !_DirWork!\SloggerLogs\_QBLogFiles\Default
	call :doCmd copy /y "%HwProgramData%\Quick Builder\LogFiles\Default\*.log" "!_DirWork!\SloggerLogs\_QBLogFiles\Default\"
)


goto :eof
:endregion ExperionAddData

:region SQL queries
:SqlAddData
where sqlcmd >NUL 2>&1
if %errorlevel% NEQ 0 (
	call :logOnlyItem no sqlcmd utility - skip sql queries
	@goto :eof
)
:: non erdb points
	call :logitem * MS SQL queries *
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem . select count^(*^) from NON_ERDB_POINTS_PARAMS
	set _SqlFile="!_DirWork!\MSSQL-Logs\_erdb.p2p.txt"
	call :InitLog !_SqlFile!
	call :DoSqlCmd ps_erdb "select count(*) from NON_ERDB_POINTS_PARAMS"
	@echo.>>!_SqlFile!
	call :logLine !_SqlFile!
	@echo ===== %time% : sqlcmd >>!_SqlFile!
	call :logLine !_SqlFile!
	sqlcmd -E -w 10000 -d ps_erdb -Q "select cast(s.StrategyName as varchar(40)) as NonCEEStrategy, cast(containingStrat.StrategyName +'.'+ strat_cont.StrategyName as varchar(40)) as ReferencedBlock, CASE WHEN s.StrategyID  & 0x80000000 = 0 THEN 'Project' ELSE 'Monitoring' END 'Avatar' from STRATEGY S inner join NON_ERDB_POINTS_PARAMS N on n.StrategyID = S.StrategyID and n.ReferenceCount > 0 inner join CONNECTION Conn on conn.PassiveParamID = N.ParamID and conn.passivecontrolid = n.strategyid  INNER JOIN STRATEGY strat_cont ON strat_cont.StrategyID = conn.ActiveControlID INNER JOIN RELATIONSHIP rel ON rel.TargetID = strat_cont.StrategyID AND rel.RelationshipID = 3 INNER JOIN STRATEGY containingStrat ON rel.SourceID = containingStrat.StrategyID " >>!_SqlFile!

:: ps_erdb stats
	set _SqlFile="!_DirWork!\MSSQL-Logs\_erdb-stats.txt"
	call :InitLog !_SqlFile!
	call :logitem . get ps_erdb database stats
	call :DoSqlCmd ps_erdb "SELECT a.object_id, cast(object_name(a.object_id) as varchar(60)) AS Object_Name, a.index_id, cast(name as varchar(60)) AS IndedxName, avg_fragmentation_in_percent FROM sys.dm_db_index_physical_stats (DB_ID (N'ps_erdb'), NULL, NULL, NULL, NULL) AS a INNER JOIN sys.indexes AS b ON a.object_id = b.object_id AND a.index_id = b.index_id where name is not null order by avg_fragmentation_in_percent desc"
	call :DoSqlCmd ps_erdb "SELECT cast(object_name(stats.object_id) as varchar(60)) AS Object_Name, cast(indexes.name as varchar(60)) AS Index_Name, STATS_DATE(stats.object_id, stats.stats_id) AS Stats_Last_Update FROM sys.stats JOIN sys.indexes ON stats.object_id = indexes.object_id AND stats.name = indexes.name order by 3 desc"

:: SQL Loggins
	set _SqlFile="!_DirWork!\MSSQL-Logs\_SqlLogins.txt"
	call :InitLog !_SqlFile!
	call :logitem . EXEC sp_helplogins
	call :DoSqlCmd master "EXEC sp_helplogins"
	call :logitem . SELECT name, type_desc, is_disabled FROM sys.server_principals
	call :DoSqlCmd master "SELECT name, type_desc, is_disabled FROM sys.server_principals"
	call :logitem . EXEC sp_who2
	call :DoSqlCmd "master" "EXEC sp_who2"

::emsqueries
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem . emsevents sql queries
	set _SqlFile="!_DirWork!\MSSQL-Logs\_emsqueries.output.txt"
	call :InitLog !_SqlFile!
	call :DoSqlCmd EMSEvents "SELECT @@VERSION"
	call :DoSqlCmd EMSEvents "SELECT CAST( SERVERPROPERTY('MachineName') as varchar(100)) AS 'MachineName'"
	call :DoSqlCmd EMSEvents "SELECT @@SERVERNAME"
	call :DoSqlCmd EMSEvents "SELECT * FROM sys.databases"
	call :DoSqlCmd EMSEvents "SELECT * FROM sys.assemblies"
	call :DoSqlCmd EMSEvents "DBCC SQLPERF(LOGSPACE)"
	call :DoSqlCmd EMSEvents "sp_helpdb EMSEvents"
	call :DoSqlCmd EMSEvents "sp_spaceused"
	call :DoSqlCmd EMSEvents "SELECT cast(DB_NAME() as varchar(25)) AS DbName, cast(name as varchar(50)) AS FileName, cast(type_desc as varchar(10)) as type_desc, size/128.0 AS CurrentSizeMB, size/128.0 - CAST(FILEPROPERTY(name, 'SpaceUsed') AS INT)/128.0 AS FreeSpaceMB FROM sys.database_files WHERE type IN (0,1);"
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

::CheckSQLDBLogs
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem . Check SQL DB Logs
	set _SqlFile="!_DirWork!\MSSQL-Logs\_CheckSQLDBLogs.txt"
	call :InitLog !_SqlFile!
	call :DoSqlCmd master "select [name] AS 'DBName', CAST(DATABASEPROPERTYEX([name],'recovery') as varchar(25)) AS 'Recovery Model' from master.dbo.SysDatabases"
	call :DoSqlCmd master "DBCC SQLPERF(LOGSPACE)"
	call :DoSqlCmd master "select CAST([name] as varchar(40)) AS 'DBName',log_reuse_wait_desc AS 'LogReuse' from sys.databases"

:: SQL Config & Status
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem . SQL Server Status
	set _SqlFile="!_DirWork!\MSSQL-Logs\_SqlStatus.txt"
	call :InitLog !_SqlFile!
	::call :DoSqlCmd master "SELECT * FROM sys.configurations"
	call :DoSqlCmd master "EXEC sp_configure"
	call :DoSqlCmd master "select * from sys.dm_os_process_memory"
	call :DoSqlCmd master "select * FROM sys.dm_os_sys_memory"
	call :DoSqlCmd msdb   "EXEC msdb.dbo.sp_help_jobhistory"

goto :eof
:endregion

:: end additional data collection
:endregion

(:getPerformanceLogs
	if /i "!_noPerfMon!" EQU "1" @goto :eof -- exit function
	call :logitem get Experion Performance Logs
	if not defined HWPERFLOGPATH set HWPERFLOGPATH=%HwProgramData%\Experion PKS\Perfmon
	if Not Exist "%HWPERFLOGPATH%" (
		call :logOnlyItem perfmon folder Not Exist "%HWPERFLOGPATH%"
	    @goto :eof
	)
	call :mkNewDir !_DirWork!\Perfmon Logs
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci $env:HWPERFLOGPATH -filt *.blg | where{$_.LastWriteTime -gt (get-date).AddDays(-10)} | foreach{copy $_.fullName -dest '!_DirWork!\Perfmon Logs\'; sleep 1} }}
	@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Copy blg files with PowerShell'. )
	if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy blg files with PowerShell' failed.)
    @goto :eof
)


:region compress data
:compress
if defined _noCabZip goto :off_nocab

@rem check for makecab.exe
for %%i in (makecab.exe) do (set exe=%%~$PATH:i)
if "!exe!" equ "" (
    @echo.
    @echo.WARNING: makecab.exe not found. Proceeding as if 'nocab' was specified.
    goto :off_nocab )

@rem construct cab directive file ::
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
		@if defined _Debug ( echo [#Debug#] !time!: ERRORLEVEL: %errorlevel% - 'at Compress with PowerShell'. )
		if "%errorlevel%" neq "0" (
			call :logItem %time% .. ERROR: %errorlevel% - 'Compress with PowerShell' failed.
			set _DoCleanup=0
			echo. _DoCleanup: !_DoCleanup!
			goto :off_nocab
			)
)

@echo.

call :showlogitem  *** %time% : %~n0 diagnostic files are in:
call :showlogitem  ***  %_DirScript%!cabName!
call :WriteHost white black *** [Note] Please upload data %_DirScript%!cabName! onto given GTAC workspace.
call :play_attention_sound

goto :eof


@rem nocab or cab failure: print working directory location ::
:off_nocab
@echo.
@echo.
@echo. Diagnostic data have NOT been compressed!!
@echo. Data located in: %_DirWork%
call :WriteHost white black *** [Note] Please compress all files in %_DirWork% and upload zip file to GTAC ftp site
call :play_attention_sound

exit /b 1 -- no cab, end compress
:endregion

:: Info:
:: 	- v1.xx see git Revision History
::  - v1.06 delete working files, if compressed
::  - v1.07 updates, fixes, ++ data
::  - v1.08 validate input arguments
::  - v1.09 - collect station configuration files
::    copy station.ini file
::    station configuration files (*.stn)
::    station toolbar files (*.stb)
::    Display Links files
::  - v1.10 - get GDI Handles Count
::  - v1.11 - Backbuild history assignments
::  - v1.12 - get NetTcpPortSharing config file
::  - v1.13 groups membership
::    net localgroup
::    get members of Experion groups
::    get mngr account information - Local Group Memberships
::  - v1.14 groups membership changed to use 'net localgroup <GrpName>'
::  - v1.15 clientaccesspolicy.xml for Silverlight
::  - v1.16 get HKEY_USERS Reg Values
::  - v1.17 function GetUserSID
::  - v1.18 Notification Utility - Dump indexes
::  - v1.19 PersistentDictionary.xml
::  - v1.20 list mini filter drivers
::  - v1.21 reg query
::    HKLM\System\CurrentControlSet\Control\GraphicsDrivers
::    HKLM\System\CurrentControlSet\Control\Power" /v HibernateEnabled
::  - v1.22
::    _HSCServerType
::    _isServer
::    _isTPS
::    .\Avalon.Graphics\DisableHWAcceleration - skip on VEP and servers
::    chkem /tpspoints
::    chkem /tpsmappings
::    McAfee_OnAccessScanner - verify reg exists before export
::  - v1.23 - Experion ACL Verify
::  - v1.24 :
::    skip filfrag on VEP
::    "HKEY_USERS\!_SID!\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" /v VisualFXSetting
::    fix copy command - destination in quotes
::    diskdrive status
::  - v1.25 :
::    hwlictool status -format:xml
::    tasklist /v /fo csv
::    query information about processes, sessions, and RD Session Host servers
::  - v1.26 :
::    reg query "HKLM\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\AV"
::    fif ECHO is off. in wmic output
::    added -UseBasicParsing in Invoke-WebRequest
::    - The response content cannot be parsed because the Internet Explorer engine is not available, or Internet Explorer's first-launch configuration is not complete. Specify the UseBasicParsing parameter and try again.
::  - v1.27 :
::    reg query HKLM\SOFTWARE\classes\Hw...
::    collecting branch cache status and settings
::  - v1.28 : export "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" Events
::  - v1.29 fix 'wmic /output:"path"' command
::  - v1.30
::    extended search for crash files
::    fix Windows version output in log file
::    get FTE config files - FTEinstall.inf & fteconfig.inf
::    get Display Links file - [TouchPanel] section in stn file
::    wmic query MSPower_DeviceEnable
::    get ErrorHandling log files
::    check files in Abstract folder for Zone.Identifier stream data
::    function :export-evtx
::    get system station configuration files
::  - v1.31
::    SQL Config & Memory Status
::    sleep delay in export-evtx
::    SQL queries use CAST(field as varchar(25)) to reduce column width
::    shheap 2 dump (dual_q heap)
::    AV reg settings export - MkNewDir
::    move Defender policies query in reg misc file
::    use System.Net.WebClient instead of Invoke-WebRequest (not available on PS2.0)
::    check pending reboot
::    fix Zone.Identifier search
::    search %SystemDrive% for crash dump files
::  - v1.32
::    skip collection of station backup logs on servers
::    station backup logs copy - add sleep 500 ms
::    export "Microsoft-Windows-TaskScheduler/Operational" Events
::  - v1.33
::    get Windows Thermal Zone Temperature information - chek if information exists
::    get EPKS Release number
::    get system flags settings in pre R43x releases
::    Dism /Online /Cleanup-Image /CheckHealth
::    get "%HwProgramData%\Experion PKS\Server\data\mapping\*.xml" files
::    vssadmin list report
::  - v1.34
::    net localgroup "Distributed COM Users"
::    actutil dump output
::    plexus -printconfig
::    gpresult /z , if /h failed
::    search for Experion resident based heap files extended
::  - v1.35
::    remove "chkem /tpspoints" - took a too long time for some nodes
::    get Windows Thermal Zone Temperature information - chek if information exists - fix
::    search for dual*q resident heap file, not only for dual_q
::    Reg QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
::    copy flbkup.def file
::  - v1.36
::    restore default command prompt title
::    get netlogon registry - check for not completed join/unjoin DC action
::    search read only files in history archives
::    get SQL jobs history
::    fixduplicates -report
::    Experion GDA Table output
::    omreport storage vdisk controller=0
::    list HPSInstall temp folder
::    EMSEvents "sp_spaceused"
::    export Experion operator settings
::  - v1.37
::    omreport verbose output
::    acronis registry settings
::    check resident based heap file exist before listitng it
::    get operator logon scripts
::    activeMonitorsReg - HKCU\Control Panel\Desktop\PerMonitorSettings
::    WMI Provider Host Quota Configuration
::  - v1.38
::    hdd defrag analysis
::    fixed VEP check: if NOT "!_VEP!"=="1"
::    reg query HKCU\Control Panel
::    reg query services
::    output paths fixes - special cases/chars
::    regedit export check and fix (EPKS reg export)
::    reg query Internet Settings
::    Reg QUERY "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main" /v UseSWRender
::    added "hwlictool list"
::    lisscn by channel
::    System & Application event logs (errors & warnings) to csv
::    get CreateSQLObject logs
::  - v1.39
::    Reg QUERY "HKLM\System\CurrentControlSet\Control\SecurityProviders\SCHANNEL" /s
::    copy CDA log archives
::    copy Activity log archives
::    copy EnggTools log archives
::    get SQL Error log files
::  - v1.40
::    netsh int tcp show heuristics
::    fix: if exist "%HwProgramData%\\Experion PKS\Server\data\CreateSQLObject*.txt"
::    get server scripts xml files
::  - v1.41
::    Reg QUERY "HKLM\SYSTEM\CurrentControlSet\Policies\Microsoft\FeatureManagement"
::    reg query Installed Components
::    find system/hidden files in HwProgramData
::    Reg QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search" /s  (AllowCortana)
::    Quick Builder Log Files
::    fixed lisscn channel number parsing
::    export-evtx "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin"
::    fixed xcopy of station logs
::    changed search for hidden/system files to use PS
::    Reg QUERY "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System
::    get output of: net localgroup "Local DSA Connections"
::    flbkup.def copy fix
::    get DasConfig.xml
::    get reg Windows Policies
::    Rockwell Software\FactoryTalk Diagnostics - reg settings
::    export-evtx - fixed(using !! instead of %%)
::    VERIFY THE PATCH DB FOLDER SECURITY
::  - v1.42
::    Windows Event Logs
::    - TerminalServicesLocalSessionManager --> Admin
::    - TerminalServicesRemoteConnectionManager --> Operational
::    - TerminalServicesRemoteConnectionManager --> Admin
::    wmic shadowcopy output
::    fix LogWmicCmd to capture errors
::    get ps_erdb database stats
::    HKLM\\SYSTEM\CurrentControlSet\Control\Video
::    HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\UnitedVideo
::    change script to use current dirctory,not he script directory