# UPSMonitor

## What is UPSMonitor

UPSMonitor is a utility written and tested in Powershell script (v 7.1) that taps into the Microsoft OS *Win32_Battery* class in order to provide the following  functions:

* Email alerts 
* Logging functions
* Action script 

UPSMonitor monitors the UPS Battery state, the percentage capacity remaining and the estimatd run time remaining (in minutes).  It raised email alerts, logs to the log file or invoke the shutdown script based on these settings.

All options are customisable and optional. UPSMonitor can be setup so that no email alerts are generated, or have logging switched off. The Action script is optional as well. At least one of the three options must be activated for the utility to run (let's save CPU cycles if nothing useful is coming out of this.)

For more information type
    get-help .\UPSMonitor.ps1
or
    .\UPSMonitor.ps1 -help

The most uptodate version of this utility can be downloaded from https://github.com/chribonn/UPSMonitor

## Action Script

The purpose of the action script is normally to shutdown the computer. It is invoked when the power supply is off and trigger points (percentage capacity and remainign time) fall below specified values.

The Action script is seperate from UPSMonitor in order to allow users the freedom to set it up to their needs without having to mess with the UPSMonitor core code.  Use cases could be to shut down multiple computers, virtual machines and send alerts (other than built in email).

## Configure Powershell execution policy if you get a PSSecurityException error

If you get an error when you execute the script similar to the one herunder you need to change the execution policy.

    .\UPSMonitor.ps1 : File .\UPSMonitor.ps1 cannot be loaded because running scripts is disabled on this system. For more information, see about_Execution_Policies at https:/go.microsoft.com/fwlink/?LinkID=135170.

    At line:1 char:1
    + .\UPSMonitor.ps1 -help
    + ~~~~~~~~~~~~~~~~
        + CategoryInfo          : SecurityError: (:) [], PSSecurityException
        + FullyQualifiedErrorId : UnauthorizedAccess
	
Open Powershell as administrator and execute the following

    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
