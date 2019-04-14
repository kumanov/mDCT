# mDCT
mDCT Windows CMD based Data Collection Script toolset v`2019.04.07
### Quick Overview of Data Collection Script mDCT.cmd
Purpose: Multi-purpose Troubleshooting tool to simplify just-in-time rapid data collection for standard and sporadic issues in in complex Experion environments or collecting data when standard DCT tool have problems.

Please start the script in the C:\temp (or any other folder path without spaces) folder in **elevated CMD window**.

For help, just run: `mdct /help`

` C:\temp>dct [parameter list] `
Please invoke the mDCT command with necessary/appropriate parameters from here.
``` Usage example: mDCT - mini DCT +ext batch script file  (krasimir.kumanov@gmail.com)
                mDCT all - run all data collection commamds
                mDCT noblg - skip Experion Performance counters (*.blg) collection
mDCT updates on: https://github.com/
 --> see 'mDCT /help' for more detailed help info
 --> Looking for help on specific keywords? Try: mDCT /help |findstr /i /c:noblg
```
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
  - reg query Policies Windows Defender
- Network information
  - netstat -nato
  - ipconfig /displaydns
  - netsh int ipv4 show dynamicport tcp
  - netsh int tcp show global
  - netsh interface ipv4 show ipstats
  - netsh interface ipv4 show tcpstats
  - netsh int ipv4 show offload
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
  - listag output
  - lisscn output
  - mapping tps.xml
