:: Filename: mDCT++.cmd - mini Data Collection Tool script ++ extensions
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

if defined _DbgOut ( echo. %time% : Start of mDCT++)
:: handle /?
if "%~1"=="/?" (
	call :usage
	exit /b 0
)

set _DirScript=%~dp0
call :preRequisites
if /i "%errorlevel%" NEQ "0" (goto :eof)
call :Initialize %*
if /i "!_Usage!" EQU "1" (goto :eof)
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
@echo. & goto :eof

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:region check preRequisites
:preRequisites
	@echo ..check preRequisites
	:: check current user account permissions - Admin required
	call :check_Permissions
	if "%errorlevel%" neq "0" goto :eof
	
	:: require no spaces in full path
	if not "%_DirScript%"=="%_DirScript: =%" ( echo.
		call :WriteHostNoLog yellow black  *** Your script execution path '%_DirScript%' contains one or more space characters
		call :WriteHostNoLog yellow black  Please use a different local path without space characters.
		exit /b 1
	) else (
		call :WriteHostNoLog green black Success: no spaces in script full path
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

	goto :eof
:endregion

:region initialize
:initialize
:: initialize variables
set _ScriptVersion=v1.17
:: Last-Update by krasimir.kumanov@gmail.com: 2019-05-31

:: change the cmd prompt environment to English
chcp 437 >NUL

:: Adding a Window Title
SET _title=%~nx0 - version %_ScriptVersion%
TITLE %_title% & set _title=

if defined _DbgOut ( echo. %time% _DirScript: %_DirScript% )

:: Change Directory to the location of the batch script file (%0)
CD /d "%_DirScript%"
@echo. .. starting '%_DirScript%%~n0 %*'

::@::if defined _DbgOut ( echo.%time% : Start of mDCT++ ^(%_ScriptVersion% - krasimir.kumanov@gmail.com^))
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
set _ProcArch=%PROCESSOR_ARCHITECTURE%
if "!_ProcArch!" equ "AMD64" Set _ProcArch=x64

set _Comp_Time=%COMPUTERNAME%_%_CurDateTime%
if defined _DbgOut ( echo. %time% _Comp_Time: %_Comp_Time% )
:: set work folder
set _DirWork=%_DirScript%%_Comp_Time%

if defined _DbgOut ( echo. %time% _DirWork: !_DirWork! )

:: init working dir
call :mkNewDir !_DirWork!
:: init LogFile
if not defined _LogFile set _LogFile=!_DirWork!\mDCTlog.txt
if defined _DbgOut ( echo. %time% _LogFile: !_LogFile! )
call :InitLog !_LogFile!

:: change priority to idle - this & all child commands
call :logitem change priority to IDLE
wmic process where name="cmd.exe" CALL setpriority "idle"  >NUL 2>&1

call :WriteHostNoLog blue black *** %_ScriptVersion% Dont click inside the script window while processing as it will cause the script to pause. ***

:: VEP detect
wmic path win32_computersystem get Manufacturer /value | findstr /ic:"VMware" >NUL 2>&1
if "%errorlevel%"=="0" (set _VEP=1)

if defined _GetLocale ( call :getLocale _locale )

call :logOnlyItem  mDCT++ (krasimir.kumanov@gmail.com) -%_ScriptVersion% start invocation: '%_DirScript%%~n0 %*'
call :logNoTimeItem  Windows version:  !v! Minor: !_OSVER4!
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
	:: #########################
	:: OS-specific checks...
	:: #########################
	for /f "tokens=2 delims=[]" %%o in ('ver')    do @set _OSVERTEMP=%%o
	for /f "tokens=2" %%o in ('echo %_OSVERTEMP%') do @set _OSVER=%%o
	for /f "tokens=1 delims=." %%o in ('echo %_OSVER%') do @set _OSVER1=%%o
	for /f "tokens=2 delims=." %%o in ('echo %_OSVER%') do @set _OSVER2=%%o
	for /f "tokens=3 delims=." %%o in ('echo %_OSVER%') do @set _OSVER3=%%o
	for /f "tokens=4 delims=." %%o in ('echo %_OSVER%') do @set _OSVER4=%%o
	for /f "tokens=4-8 delims=[.] " %%i in ('ver') do (if %%i==Version (set _v=%%j.%%k.%%l.%%m) else (set _v=%%i.%%j.%%k.%%l))
	if defined _DbgOut ( echo. %time% ###getWinVer OS: %_OSVER1% %_OSVER2% %_OSVER3% %_OSVER4% Version %_v% )
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
	set _CurDateTime=%_CurDateTime:~0,8%_%_CurDateTime:~8,6%
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
	@echo ================================================================================== >> "!_LogFile!"
	@echo ===== %time% : %* >> "!_LogFile!"
	@echo ================================================================================== >> "!_LogFile!"
	%* >> %_LogFile% 2>&1
	@echo. >> "!_LogFile!"
	call :SleepX 1
	@goto :eof

:doCmdNoLog  - UTILITY log execution of a command to the current log file
	@echo ================================================================================== >> "!_LogFile!"
	@echo ===== %time% : %* >> !_LogFile!
	@echo ================================================================================== >> "!_LogFile!"
	%*
	@echo. >> "!_LogFile!"
	@goto :eof

:LogCmd [filename; command] - UTILITY to log command header and output in filename
	SETLOCAL
	for /f "tokens=1* delims=; " %%a in ("%*") do (
		set _LogFileName=%%a
		@echo ================================================================================== >> "!_LogFileName!"
		@echo ===== %time% : %%b >> "!_LogFileName!"
		@echo ================================================================================== >> "!_LogFileName!"
		%%b >> "!_LogFileName!" 2>&1
	)
	@echo. >> "!_LogFileName!"
	call :SleepX 1
	ENDLOCAL
	@goto :eof

:LogWmicCmd [filename; command] - UTILITY to log command header and output in filename
	SETLOCAL
	for /f "tokens=1* delims=; " %%a in ("%*") do (
		set _LogFileName=%%a
		@echo ================================================================================== >> "!_LogFileName!"
		@echo ===== %time% : %%b >> "!_LogFileName!"
		@echo ================================================================================== >> "!_LogFileName!"
		::%%b /Format:Texttable | more /s >> "!_LogFileName!" 2>&1
		For /F "tokens=* delims=" %%h in ('%%b') do (
			set "_line=%%h"
			set "_line=!_line:~0,-1!"
			echo !_line!>>"!_LogFileName!"
		)
	)
	@echo. >> "!_LogFileName!"
	call :SleepX 1
	ENDLOCAL
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
	@echo.>> %_RegFile%
	@echo ================================================================================== >> %_RegFile%
	@echo ===== %time% : REG.EXE %* >> %_RegFile%
	@echo ================================================================================== >> %_RegFile%
	%SYSTEMROOT%\SYSTEM32\REG.EXE %* >> %_RegFile% 2>&1
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
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci -Path !_HMIWebLog! -Include log.txt,hmiweblogY*.txt -Recurse | Select-String -Pattern 'Connecting using .stn file: (.*stn)$' | foreach{$_.Matches} | foreach {$_.Groups[1].Value} | sort -Unique | out-file !_stnFiles!}}
	:: copy files
	for /f "tokens=* delims=" %%h in ('type !_stnFiles!') DO (
		set _stnFile=_%%~nh%%~xh
		if exist "!_DirWork!\Station-logs\!_stnFile!" (set _stnFile=_%%~nh_!random!%%~xh)
		call :doCmd copy /y "%%h" "!_DirWork!\Station-logs\!_stnFile!"
	)
	call :SleepX 1
	
	:: get stb files
	@set _stbFiles=%temp%\_stbFiles.txt
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci -Path !_DirWork!\Station-logs\ -Include *.stn -Recurse | Select-String -Pattern 'Toolbar_Settings=(.*stb)$' | foreach{$_.Matches} | foreach {$_.Groups[1].Value} | sort -Unique | out-file !_stbFiles!}}
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
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci -Path !_DirWork!\Station-logs\ -Include *.stn -Recurse | Select-String -Pattern 'DisplayLinksPath=(.*xml)$' | foreach{$_.Matches} | foreach {$_.Groups[1].Value} | sort -Unique | out-file !_dspLinksFiles!}}
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
>>!_psFile! echo $GuiResources ^| sort GDIHandles -Desc ^| select -First 10 ^| ft -a ^| out-file !_outFile! -Append -Encoding ascii
::@::>>!_psFile! echo '' ^| out-file !_outFile! -Append -Encoding ascii
>>!_psFile! echo $('{0} processes; {1}/{2} with/without GDI objects' -f $allProcesses.Count, $GuiResources.Count, ($allProcesses.Count - $GuiResources.Count)) ^| out-file !_outFile! -Append -Encoding ascii
>>!_psFile! echo "Total number of GDI handles: $auxCountHandles" ^| out-file !_outFile! -Append -Encoding ascii
PowerShell.exe -NonInteractive  -NoProfile -ExecutionPolicy Bypass %_psFile%
if defined _DbgOut ( echo. .. ** ERRORLEVEL: %errorlevel% - 'at getGDIHandlesCount with PowerShell'. )
if "%errorlevel%" neq "0" (
	call :logItem %time% .. ERROR: %errorlevel% - 'getGDIHandlesCount with PowerShell' failed.
	)
::del PS file
call :doit del "!_psFile!"
ENDLOCAL
call :SleepX 1
exit /b

:getGropuMembers    -- get group mebers
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
if defined _DbgOut ( echo. .. ** ERRORLEVEL: %errorlevel% - at get '!_groupName!' members with PowerShell. )
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
			call :GetReg QUERY "HKEY_USERS\!_SID!\Software\Microsoft\Windows\DWM" /v ColorPrevalence
			call :GetReg QUERY "HKEY_USERS\!_SID!\Software\Microsoft\Avalon.Graphics" /v DisableHWAcceleration

			)
		)
	(ENDLOCAL & REM -- RETURN VALUES
		call :SleepX 1
	)
	exit /b

:GetUserSID    -- get user SID
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

:endregion

:region DCT Data
:CollectDctData
:: return, if no DCT data required
if /i "!_NoDctData!" EQU "1" (goto :eof)

call :logitem *** DCT data collection ... ***

:: GeneralSystemInfo folder
call :mkNewDir  !_DirWork!\GeneralSystemInfo

call :logitem MSInfo32 report
call :doCmd msinfo32 /report !_DirWork!\GeneralSystemInfo\MSInfo32.txt

call :logitem Windows hosts file copy
call :doCmd copy /y %windir%\System32\drivers\etc\hosts !_DirWork!\GeneralSystemInfo\

call :logitem ipconfig output
call :LogCmd !_DirWork!\GeneralSystemInfo\ipconfig.txt ipconfig /all

call :logitem netsh firewall show config
call :LogCmd !_DirWork!\GeneralSystemInfo\firewall.txt netsh firewall show config

call :logitem get time zone information
call :getTimeZoneInfo _tzName _tzBias
call :InitLog !_DirWork!\GeneralSystemInfo\timezone.output.txt
@echo TimeZone Name: %_tzName% >>!_DirWork!\GeneralSystemInfo\timezone.output.txt
@echo TimeZone Bias: %_tzBias% >>!_DirWork!\GeneralSystemInfo\timezone.output.txt
call :SleepX 1

(call :logitem export Windows Events
call :doCmd wevtutil epl Application !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_Application.evtx /overwrite:true
call :doCmd wevtutil epl FTE !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_FTE.evtx /overwrite:true
call :doCmd wevtutil epl HwSnmp !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_HwSnmp.evtx /overwrite:true
call :doCmd wevtutil epl HwSysEvt !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_HwSysEvt.evtx /overwrite:true
call :doCmd wevtutil epl Security !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_Security.evtx /overwrite:true
call :doCmd wevtutil epl System !_DirWork!\GeneralSystemInfo\%COMPUTERNAME%_System.evtx /overwrite:true)

call :logitem get Experion PKS Product Version file
call :doCmd copy /y "%HwInstallPath%\Experion PKS\ProductVersion.txt" !_DirWork!\GeneralSystemInfo\

call :logitem query services
call :DoGetSVC %time%

(call :logitem export Experion registry settings
call :mkNewDir  "!_DirWork!\RegistryInfo"
if exist %windir%\SysWOW64\regedit.exe (set _regedit="%windir%\SysWOW64\regedit.exe") else (set _regedit=regedit.exe)
call :doCmd !_regedit! /E !_DirWork!\RegistryInfo\HKEY_CURRENT_USER_Software_Honeywell.txt "HKEY_CURRENT_USER\Software\Honeywell"
call :doCmd !_regedit! /E !_DirWork!\RegistryInfo\HKEY_LOCAL_MACHINE_Software_Honeywell.txt "HKEY_LOCAL_MACHINE\SOFTWARE\Honeywell"
call :doCmd !_regedit! /E !_DirWork!\RegistryInfo\HKEY_LOCAL_MACHINE_Software_Microsoft_Uninstall.txt "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
)

call :logitem get FTE logs
::@::call :doCmd xcopy /s/e/i/q/y/H "%HwProgramData%\ProductConfig\FTE\*.log" "!_DirWork!\FTELogs\"
call :mkNewDir  "!_DirWork!\FTELogs"
for /r "%HwProgramData%\ProductConfig\FTE" %%g in (*.log) do (type %%g >"!_DirWork!\FTELogs\%%~ng%%~xg")

call :logitem get HMIWeb log files
call :doCmd xcopy /i/q/y/H "%HwProgramData%\HMIWebLog\*.txt" "!_DirWork!\Station-logs\"
call :doCmd xcopy /i/q/y/H "%HwProgramData%\HMIWebLog\Archived Logfiles\*.txt" "!_DirWork!\Station-logs\Rollover-logs\"

call :logitem task list /services
call :mkNewDir  !_DirWork!\ServerDataDirectory
call :LogCmd !_DirWork!\ServerDataDirectory\TaskList.txt tasklist /fo csv /svc

where setpar >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion active log paranoids
	call :mkNewDir !_DirWork!\SloggerLogs
	call :LogCmd !_DirWork!\SloggerLogs\setpar.active.txt setpar /active
)

(::SloggerLogs
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
    call :doCmd bckbld -out !_DirWork!\ServerDataDirectory\back_build.output.txt
)

where hdwbckbld >NUL 2>&1
if %errorlevel%==0 (
	call :logitem Experion hardware back build
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd hdwbckbld -out !_DirWork!\ServerDataDirectory\hardware_back_build.output.txt
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
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.alarmgroup.output.txt  -ALARMGROUP
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.asset.output.txt  -ASSET
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.network.output.txt  -NETWORK
    call :doCmd embckbuilder  !_DirWork!\ServerDataDirectory\embckbuilder.system.output.txt  -SYSTEM
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
    call :doCmd fildmp -DUMP -FILE !_DirWork!\ServerDataDirectory\sysflg.output.txt -FILENUM 8 -RECORDS 1 -FORMAT HEX
	call :logitem Experion Area Asignmnt Table output
    call :doCmd fildmp -DUMP -FILE !_DirWork!\ServerDataDirectory\areaasignmnt.output.txt -FILENUM 7 -RECORDS 1,1001 -FORMAT HEX
)

if exist "%HwProgramData%\Experion PKS\Server\data\system.build" (
	call :logitem copy system.build file
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\system.build" !_DirWork!\ServerDataDirectory\
)

if exist "%HwProgramData%\Experion PKS\Server\data\" (
	call :logitem collect bad files .\server\data\*.bad
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd xcopy /i/q/y/H "%HwProgramData%\Experion PKS\Server\data\*.bad" !_DirWork!\ServerDataDirectory\
)

if exist "%HwProgramData%\TPNServer\TPNServer.log" (
	call :logitem copy TPNServer.log file
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\TPNServer\TPNServer.log" !_DirWork!\ServerDataDirectory\
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
)

where usrlrn >NUL 2>&1
if %errorlevel%==0 (
	call :logitem usrlrn usrlrn -p -a
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\usrlrn.txt
    call :logCmd !_DirWork!\ServerRunDirectory\usrlrn.txt usrlrn -p -a
)

where what >NUL 2>&1
if %errorlevel%==0 (
	call :logitem What - Getting Experion exe/dll and source file information
	call :mkNewDir !_DirWork!\ServerRunDirectory
	call :InitLog !_DirWork!\ServerRunDirectory\what.output.txt
	for /r "%HwInstallPath%\Experion PKS\Server\run" %%a in (*.exe *.dll) do what "%%a" >>!_DirWork!\ServerRunDirectory\what.output.txt
)


goto :eof
:endregion

:region Additional Data
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:CollectAdditionalData
:: return, if no Additional data required
if /i "!_NoAddData!" EQU "1" (goto :eof)

call :logitem *** Additional data collection ... ***

call :WindowsAddData
call :NetworkAddData
call :ExperionAddData
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
call :LogCmd !_DirWork!\GeneralSystemInfo\_whoami.txt whoami /all

call :logitem . fetching environment Variables
call :InitLog !_DirWork!\GeneralSystemInfo\_EnvVariables.txt
call :LogCmd !_DirWork!\GeneralSystemInfo\_EnvVariables.txt set

call :logitem . scheduled task - query
schtasks /query /xml ONE >!_DirWork!\GeneralSystemInfo\_scheduled_tasks.xml
call :SleepX 1

call :logitem . collecting GPResult output
set _GPresultFile=!_DirWork!\GeneralSystemInfo\_GPresult.htm
call :doCmd gpresult /h !_GPresultFile! /f

call :DoNltestDomInfo

call :logitem . get power configuration settings
set _powerCfgFile=!_DirWork!\GeneralSystemInfo\_powercfg.txt
call :InitLog !_powerCfgFile!
call :LogCmd !_powerCfgFile! powercfg -Q
:: reg query power settings
call :logOnlyItem . reg query power settings
set _RegFile=!_powerCfgFile!
call :GetReg QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power"  /s /t reg_dword
::  fast reboot - "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled

call :logitem . collecting Quick Fix Engineering information (Hotfixes)
call :doCmd  wmic /output:!_DirWork!\GeneralSystemInfo\_Hotfixes.txt qfe list

if exist "%windir%\Honeywell_MsPatches.txt" (
	call :logitem . get Honeywell_MsPatches.txt
	call :doCmd copy /y "%windir%\Honeywell_MsPatches.txt" "!_DirWork!\GeneralSystemInfo\_Honeywell_MsPatches.txt"
)

call :logitem . WindowsUpdate.log
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :doCmd copy /y "%windir%\WindowsUpdate.log" "!_DirWork!\GeneralSystemInfo\_WindowsUpdate.log"
if exist %windir%\Logs\WindowsUpdate (
	call :logitem . get Windows Update ETL Logs
	call :mkNewDir  !_DirWork!\GeneralSystemInfo\_WindowsUpdateEtlLogs
	call :doCmd copy /y "%windir%\Logs\WindowsUpdate\*.etl" "!_DirWork!\GeneralSystemInfo\_WindowsUpdateEtlLogs\"
)

:WmiRootSecurityDescriptor
call :logitem . WMI Root Security Descriptor
call :doCmd  wmic /output:!_DirWork!\GeneralSystemInfo\_WmiRootSecurityDescriptor.txt /namespace:\\root path __systemsecurity call GetSecurityDescriptor

:: AV on accesss scanner settings
call :doCmd REG EXPORT "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\McAfee\SystemCore\VSCore\On Access Scanner" !_DirWork!\RegistryInfo\HKLM_McAfee_OnAccessScanner.txt

set _SecurityFile=!_DirWork!\GeneralSystemInfo\_SecurityCfg.txt
call :InitLog !_SecurityFile!
call :logitem . secedit /export /cfg
secedit /export /cfg !_SecurityFile! >> !_LogFile!
call :SleepX 1

call :logitem . query drivers information
call :LogCmd !_DirWork!\GeneralSystemInfo\_driverquery.output.csv driverquery /fo csv /v

call :logOnlyItem . reg query Policies Windows Defender
set _RegFile=!_DirWork!\GeneralSystemInfo\_PoliciesWindowsDefender.txt
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\Software\Policies\Microsoft\Windows Defender" /s

call :logOnlyItem . reg query RPC settings
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile=!_DirWork!\RegistryInfo\_RPC_registry_settings.txt
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\RPC" /s
call :GetReg QUERY "HKLM\Software\Microsoft\Rpc" /s

call :logOnlyItem . reg query Microsoft\OLE
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile=!_DirWork!\RegistryInfo\_OLE_registry_settings.txt
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\Software\Microsoft\OLE" /s

call :logOnlyItem . reg query misc
call :mkNewDir  !_DirWork!\RegistryInfo
set _RegFile=!_DirWork!\RegistryInfo\_reg_query_misc.txt
call :InitLog !_RegFile!
call :GetReg QUERY  "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows" /v CEIPEnable

call :logOnlyItem . Windows Time status/settings
set _WindowsTimeFile=!_DirWork!\GeneralSystemInfo\_WindowsTime.txt
call :mkNewDir  !_DirWork!\GeneralSystemInfo
call :InitLog !_WindowsTimeFile!
call :LogCmd !_WindowsTimeFile! w32tm /query /status /verbose
call :LogCmd !_WindowsTimeFile! w32tm /query /configuration
set _RegFile=!_WindowsTimeFile!
call :GetReg QUERY  "HKLM\SYSTEM\CurrentControlSet\Services\W32Time" /s

(:: temperature
if not "%_VEP%"=="1" (
	call :mkNewDir  !_DirWork!\GeneralSystemInfo
	set _ThermalZoneTemperature=!_DirWork!\GeneralSystemInfo\_ThermalZoneTemperature.txt
	call :InitLog !_ThermalZoneTemperature!
	@echo Temperature at thermal zone in tenths of degrees Kelvin >>!_ThermalZoneTemperature!
	@echo Convert to Celsius: xxx / 10 - 273.15 >>!_ThermalZoneTemperature!
	@echo.>>!_ThermalZoneTemperature!
	wmic /namespace:\\root\wmi PATH MSAcpi_ThermalZoneTemperature get Active,CriticalTripPoint,CurrentTemperature | more /s >>!_ThermalZoneTemperature!
)
)

:: get GDI Handles Count
call :getGDIHandlesCount


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

call :logitem . get HKEY_USERS Reg Values
set _RegFile=!_DirWork!\RegistryInfo\_HKEY_USERS.txt
call :mkNewDir !_DirWork!\RegistryInfo
if exist !_RegFile! call :doit del "!_RegFile!"
call :GetHkeyUsersRegValues


:: get mngr account information - Local Group Memberships
call :logitem . get mngr account information - Local Group Memberships
set _mngrUserAccount=!_DirWork!\GeneralSystemInfo\_mngrUserAccount.txt
call :InitLog !_mngrUserAccount!
call :LogCmd !_mngrUserAccount! net user mngr


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
call :LogCmd !_NetShFile! netsh int ipv4 show dynamicport tcp
call :LogCmd !_NetShFile! netsh int tcp show global
call :LogCmd !_NetShFile! netsh int ipv4 show offload
call :LogCmd !_NetShFile! netsh interface ipv4 show ipstats
call :LogCmd !_NetShFile! netsh interface ipv4 show udpstats
call :LogCmd !_NetShFile! netsh interface ipv4 show tcpstats
call :LogCmd !_NetShFile! netsh http show urlacl

:nslookup
call :logitem . nslookup
call :SleepX 1
:: NS LookUp - Forward
echo ======================================== > !_DirWork!\_Network\nslookup.txt
echo %date% %time% >> !_DirWork!\_Network\nslookup.txt
echo cmd: nslookup %computername% >> !_DirWork!\_Network\nslookup.txt
nslookup %computername% >> !_DirWork!\_Network\nslookup.txt  2>>&1
:: NS LookUp - Reverse
set ip_address_string="IPv4 Address"
for /f "usebackq tokens=2 delims=:" %%f in (`ipconfig ^| findstr /c:%ip_address_string%`) do (
	echo ======================================== >> !_DirWork!\_Network\nslookup.txt
	echo %date% %time% >> !_DirWork!\_Network\nslookup.txt
    echo Your IP Address is: %%f  >> !_DirWork!\_Network\nslookup.txt
	echo cmd: nslookup %%f >> !_DirWork!\_Network\nslookup.txt
	nslookup %%f >> !_DirWork!\_Network\nslookup.txt  2>>&1
)
echo ======================================== >> !_DirWork!\_Network\nslookup.txt
echo %date% %time% - done>> !_DirWork!\_Network\nslookup.txt

call :logitem . route / arp
call :InitLog !_DirWork!\_Network\arp.txt
call :LogCmd !_DirWork!\_Network\arp.txt  arp -a -v
call :InitLog !_DirWork!\_Network\route.print.txt
call :LogCmd !_DirWork!\_Network\route.print.txt route print

call :logitem . nbtstat-n
call :InitLog !_DirWork!\_Network\nbtstat.txt
call :LogCmd !_DirWork!\_Network\nbtstat.txt  nbtstat -n

:advfirewall
call :logitem . firewall rules
call :InitLog !_DirWork!\_Network\firewall_rules.txt
call :LogCmd !_DirWork!\_Network\firewall_rules.txt netsh advfirewall monitor show firewall verbose

call :logitem . net commands
set _NetCmdFile=!_DirWork!\_Network\netcmd.txt
call :InitLog !_NetCmdFile!
call :LogCmd !_NetCmdFile! NET SHARE
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

set _RegFile=!_DirWork!\_Network\TcpIpParameters.txt
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp\Parameters" /v ArpRetryCount
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\TcpIp\Parameters" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\Tcpip6\Parameters" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\tcpipreg" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Services\iphlpsvc" /s
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}" /s

call :logitem . wmic nic/nicconfig get
set _NicConfig=!_DirWork!\_Network\nicconfig.txt
call :InitLog !_NicConfig!
call :LogWmicCmd !_NicConfig! wmic nic get
call :LogWmicCmd !_NicConfig! wmic nicconfig get Description,DHCPEnabled,Index,InterfaceIndex,IPEnabled,MACAddress
call :LogWmicCmd !_NicConfig! wmic nicconfig get


call :logitem . msft_providers get Provider,HostProcessIdentifier 
call :InitLog !_DirWork!\_Network\msft_providers.txt
call :LogWmicCmd !_DirWork!\_Network\msft_providers.txt wmic path msft_providers get Provider,HostProcessIdentifier

if exist "c:\Program Files\VMware\VMware Tools\VMwareToolboxCmd.exe" (
	set _vmtbcmd="C:\Program Files\VMware\VMware Tools\VMwareToolboxCmd.exe"
	call :logOnlyItem . VMWare status
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
powershell -Command "& {@(try{(Invoke-WebRequest -Uri http://localhost/clientaccesspolicy.xml).Content} catch{$_.Exception.Message}) | out-file !_DirWork!\_Network\clientaccesspolicy.xml }"
call :SleepX 1

goto :eof
:endregion NetworkAddData

:region Experion data
:ExperionAddData
call :logitem * Experion data *

where lisscn >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . lisscn output
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd lisscn -all_ref -OUT !_DirWork!\ServerDataDirectory\_lisscn_all.txt
	call :doCmd lisscn -OUT !_DirWork!\ServerDataDirectory\_lisscn.txt
)

where filfrag >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . filfrag output
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :InitLog !_DirWork!\ServerDataDirectory\_filfrag.output.txt
    call :logCmd !_DirWork!\ServerDataDirectory\_filfrag.output.txt filfrag
)

if exist "%HwProgramData%\Experion PKS\Server\data\mapping\tps.xml" (
	call :logitem . copy .\mapping\tps.xml
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Server\data\mapping\tps.xml" !_DirWork!\ServerDataDirectory\_mapping.tps.xml
)

where dsasublist >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . dsasublist
	call :mkNewDir !_DirWork!\ServerDataDirectory
	call :InitLog !_DirWork!\ServerDataDirectory\_dsasublist.txt
    call :logCmd !_DirWork!\ServerDataDirectory\_dsasublist.txt dsasublist
)

(::CrashDumps - file list & reg settings
call :logitem . create crash dumps list
call :mkNewDir !_DirWork!\CrashDumps
set _CrashDumpsList=!_DirWork!\CrashDumps\_CrashDumpsList.txt
call :InitLog !_CrashDumpsList!
call :logCmd  !_CrashDumpsList! dir %windir%\memory.dmp
call :logCmd  !_CrashDumpsList! dir /o:-d %windir%\Minidump
call :logCmd  !_CrashDumpsList! dir /o:-d "%HwProgramData%\Experion PKS\CrashDump"
call :logCmd  !_CrashDumpsList! dir /o:-d "%HwProgramData%\HMIWebLog\DumpFiles"
call :logCmd  !_CrashDumpsList! dir /o:-d "%HwProgramData%\Experion PKS\server\data\*.dmp"
call :logCmd  !_CrashDumpsList! dir  /o-d /s c:\users\*.dmp

call :logitem . crash control registry settings
set _RegFile=!_DirWork!\CrashDumps\_RegCrashControl.txt
call :InitLog !_RegFile!
call :GetReg QUERY "HKLM\System\CurrentControlSet\Control\CrashControl" /s
call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\AeDebug" /s
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug" /s
call :GetReg QUERY "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\Windows Error Reporting\LocalDumps" /s
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps" /s
call :GetReg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options" /s

call :logitem . Recover OS settings
call :DoCmd  wmic /output:!_DirWork!\CrashDumps\_recoveros.txt RECOVEROS
)

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
)

(call :logitem . list disk resident heap files
call :mkNewDir !_DirWork!\ServerDataDirectory
call :logCmd !_DirWork!\ServerDataDirectory\_DiskResidentHeaps.txt dir "%HwProgramData%\Experion PKS\Server\data\locks" "%HwProgramData%\Experion PKS\Server\data\dual_q" "%HwProgramData%\Experion PKS\Server\data\gda" "%HwProgramData%\Experion PKS\Server\data\tagcache" "%HwProgramData%\Experion PKS\Server\data\taskrequest" "%HwProgramData%\Experion PKS\Server\data\dbrepsrvup*"
)

if exist "%HwProgramData%\Experion PKS\Client\Station\station.ini" (
	call :logitem . copy station.ini file
	call :mkNewDir !_DirWork!\Station-logs
	call :doCmd copy /y "%HwProgramData%\Experion PKS\Client\Station\station.ini" !_DirWork!\Station-logs\_station.ini
)

if exist "%HwProgramData%\HMIWebLog\Log.txt" (
	call :logitem . collect station configuration files
	call :getStationFiles
)

where bckbld >NUL 2>&1
if %errorlevel%==0 (
	call :logitem . Backbuild history assignments
	call :mkNewDir !_DirWork!\ServerDataDirectory
    call :doCmd bckbld -np -nt -ng -ns -out !_DirWork!\ServerDataDirectory\_hist_scada.txt
    call :doCmd bckbld -np -nt -ng -ns -tag CDA -out !_DirWork!\ServerDataDirectory\_hist_cda.txt
    call :doCmd bckbld -np -nt -ng -ns -tag RDA -out !_DirWork!\ServerDataDirectory\_hist_dsa.txt
    call :doCmd bckbld -np -nt -ng -ns -tag PSA -out !_DirWork!\ServerDataDirectory\_hist_psa.txt
)

goto :eof
:endregion ExperionAddData

:region SQL queries
:SqlAddData
where sqlcmd >NUL 2>&1
if %errorlevel% NEQ 0 (
	call :logOnlyItem no sqlcmd utility - skip sql queries
	goto :eof
)
	call :logitem * MS SQL queries *
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem . select count^(*^) from NON_ERDB_POINTS_PARAMS
	set _SqlFile=!_DirWork!\MSSQL-Logs\_erdb.p2p.txt
	call :InitLog !_SqlFile!
	call :DoSqlCmd ps_erdb "select count(*) from NON_ERDB_POINTS_PARAMS"
	sqlcmd -E -w 10000 -d ps_erdb -Q "select s.StrategyName as NonCEEStrategy, containingStrat.StrategyName +'.'+ strat_cont.StrategyName as ReferencedBlock, CASE WHEN s.StrategyID  & 0x80000000 = 0 THEN 'Project' ELSE 'Monitoring' END 'Avatar' from STRATEGY S inner join NON_ERDB_POINTS_PARAMS N on n.StrategyID = S.StrategyID and n.ReferenceCount > 0 inner join CONNECTION Conn on conn.PassiveParamID = N.ParamID and conn.passivecontrolid = n.strategyid  INNER JOIN STRATEGY strat_cont ON strat_cont.StrategyID = conn.ActiveControlID INNER JOIN RELATIONSHIP rel ON rel.TargetID = strat_cont.StrategyID AND rel.RelationshipID = 3 INNER JOIN STRATEGY containingStrat ON rel.SourceID = containingStrat.StrategyID " >>!_SqlFile!

	:: SQL Loggins
	set _SqlFile=!_DirWork!\MSSQL-Logs\_SqlLogins.txt
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
	set _SqlFile=!_DirWork!\MSSQL-Logs\_emsqueries.output.txt
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

::CheckSQLDBLogs
	call :mkNewDir !_DirWork!\MSSQL-Logs
	call :logitem . Check SQL DB Logs
	set _SqlFile=!_DirWork!\MSSQL-Logs\_CheckSQLDBLogs.txt
	call :InitLog !_SqlFile!
	call :DoSqlCmd master "select [name] AS 'Database Name', DATABASEPROPERTYEX([name],'recovery') AS 'Recovery Model' from master.dbo.SysDatabases"
	call :DoSqlCmd master "DBCC SQLPERF(LOGSPACE)"
	call :DoSqlCmd master "select name AS 'Database name',log_reuse_wait_desc AS 'Log  Reuse' from sys.databases"

goto :eof
:endregion

:: end additional data collection
:endregion

(:getPerformanceLogs
	if /i "!_noPerfMon!" EQU "1" goto :eof -- exit function
	call :logitem get Experion Performance Logs
	if not defined HWPERFLOGPATH set HWPERFLOGPATH=%HwProgramData%\Experion PKS\Perfmon
	if Not Exist "%HWPERFLOGPATH%" (
		call :logOnlyItem perfmon folder Not Exist "%HWPERFLOGPATH%"
		goto :eof
	)
	call :mkNewDir !_DirWork!\Perfmon Logs
	PowerShell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass "&{Invoke-Command -Script{ gci $env:HWPERFLOGPATH -filt *.blg | where{$_.LastWriteTime -gt (get-date).AddDays(-10)} | foreach{copy $_.fullName -dest '!_DirWork!\Perfmon Logs\'; sleep 1} }}
	if defined _DbgOut ( echo. .. ** ERRORLEVEL: %errorlevel% - 'at Copy blg files with PowerShell'. )
	if "%errorlevel%" neq "0" ( call :logItem %time% .. ERROR: %errorlevel% - 'Copy blg files with PowerShell' failed.)
	goto :eof
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
		if defined _DbgOut ( echo. .. ** ERRORLEVEL: %errorlevel% - 'at Compress with PowerShell'. )
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
:: 	- v1.xx see Revision History in SCN file
::  - v1.05 Add NltestDomInfo; mkCab
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

:: ToDo:
:: - [] McAfee - check reg key before query
::    reg query "HKLM\SOFTWARE\Wow6432Node\McAfee\SystemCore\VSCore\On Access Scanner" >NUL 2>&1
::    if %ERRORLEVEL% EQU 0 goto :noMcAfeeScanner

:: - [] Log and XML files from C:\Program Files\Honeywell\Experion PKS\Engineering Tools\temp\EMB.
::    The XML files contain the last asset, alarm group and system models that were downloaded to the server from the Enterprise Model Builder.
::    The log files contain any errors or warnings from these downloads

