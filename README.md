# mDCT++
mDCT++ : Windows CMD based Data Collection Script toolset v.`2019.04.18`
### Quick Overview of Data Collection Script mDCT++.cmd
Purpose: Multi-purpose Troubleshooting tool to simplify just-in-time rapid data collection for standard and sporadic issues in in complex Experion environments or collecting data when standard DCT tool have problems.

Please start the script in the C:\temp (or any other folder path without spaces) folder in **elevated CMD window**.

For help, just run: `mDCT++ /help`

` C:\temp>mDCT++ [parameter list] `
Please invoke the mDCT++ command with necessary/appropriate parameters from here.
```
Parameters:
  noDctData - skip collection of DCT data
  noPerfMon - skip Performance Counter colection - *.blg files
  noAddData - skip colection of the addtional diagnostic data
  noCabZip  - the data collected will not be compressed

Usage examples:
 Example 1 - collect all data and create archive - default run without parameters
 c:\Temp\> mDCT++.cmd

 Example 2 - Do not collect Perfromance counters
 c:\Temp\> mDCT++.cmd  noPerfMon

 Example 3 - No additional data - only DCT data, PerfMon logs and crash dump list
 c:\Temp\> mDCT++.cmd  noAddData

 Example 4 - small DCT data collection - no PerMOn Logs, no addtional data
 c:\Temp\> mDCT++.cmd  noPerfMon  noAddData

 Example 5 - collect only extended diagnostic data
 c:\Temp\> mDCT++.cmd  noDctData noPerfMon
```
** mDCT++ updates and more details on: https://github.com/kumanov/mDCT

## Extended diagnostic data Collection
- Windows general
  - Enviroment variables output
  - whoami /all
  - gpresult /h
  - Driver query
  - Hotfixes
  - Power settings
    - powercfg /query
    - "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
  - WindowsUpdate.log
    - collect %windir%\Logs\WindowsUpdate\*.etl
  - SystemInfo.exe output
  - WMI root:/ security descriptor
  - Scheduled tasks query output
  - NlTestDomInfo
  - registry query
    - HKLM\Software\Policies\Microsoft\Windows Defender
    - HKLM\SOFTWARE\Policies\Microsoft\Windows NT\RPC
    - HKLM\Software\Microsoft\RPC
    - HKLM\Software\Microsoft\OLE
    - reg query misc
      - "HKLM\SOFTWARE\Policies\Microsoft\SQMClient\Windows" /v CEIPEnable
  - Windows Time status/settings
    - w32tm /query /configuration /verbose
    - w32tm /query /configuration
    - HKLM\SYSTEM\CurrentControlSet\Services\W32Time
  - get GDI Handles Count
- Network information
  - netstat -nato
  - ipconfig /displaydns
  - netsh config/stats
    - netsh int ipv4 show dynamicport tcp
    - netsh int tcp show global
    - netsh int ipv4 show offload
    - netsh interface ipv4 show ipstats
    - netsh interface ipv4 show tcpstats
    - netsh http show urlacl
  - nslookup - forward & revers
  - arp -a -v
  - route print
  - nbtstat -n
  - firewall rules
  - TcpIp/Parameters registry export
  - net commands
    - net config server
    - net session
    - net share
    - net user
    - net accounts
    - net config wksta
    - net use
    - net statistics workstation
    - net statistics server
  - wmic nic get
  - wmic nicconfig get
  - msft_providers get Provider,HostProcessIdentifier
  - VMware stat
    - upgrade status
    - timesync status
    - stat raw text
  - get NetTcpPortSharing config file
- Crash Dumps
  - crash dump files list
  - Recover OS settings
  - registry Crash Control settings
= SQL queries
  - EMSEvents queries
  - SQLDBLogs status
  = SQL Loggins
  - NON_ERDB_POINTS_PARAMS
- Experion
  - dsasublist output
  - filfrag output
  - lisscn output
  - mapping tps.xml
  - dual_status
  - cstn_status
  - ps output
  - shheap 1 check output
  - list disk resident heap files
  - copy station.ini file
  = collect station configuration files
    - station configuration files (*.stn)
    - station toolbar files (*.stb)
    - Display Links files
  - Backbuild history assignments
