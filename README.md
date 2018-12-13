# SharpCOM

SharpCOM is a c# port of [Invoke-DCOM](https://github.com/rvrsh3ll/Misc-Powershell-Scripts/blob/master/Invoke-DCOM.ps1)


Major credit to @cobbr_io for the initial conversion of Invoke-DCOM to c# in [SharpSploit](https://github.com/cobbr/SharpSploit/blob/master/SharpSploit/LateralMovement/DCOM.cs)


This version is meant to be a "weaponized" port of the SharpSploit DCOM lateral movement module.


As an example, one could execute SharpCOM.exe through Cobalt Strike's Beacon "execute-assembly" module.


#### Example usage
beacon>execute-assembly /root/SharpCOM/SharpCOM.exe --Method ShellWindows --ComputerName hosta.example.local --Command "calc.exe"
