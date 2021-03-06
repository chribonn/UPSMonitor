<#  
.SYNOPSIS 
    Reset-Logs takes a mandatory parameter:
        * Logfile = Name of the file to reset

    It takes two optional parameters:
        * BackupLogFile = Name of the back log file. If specified the original log file is renamed to this file before it is reset. If the backup log file already exists it will be deleted.
        * HeaderRow = If specified the reset log file will have this text written as a header record
        * Help = Describes this utility

    This script may be invoked via a task schedular task. The frequency of execution (weekly, monthly, quarterly, anually) dictates the the amount of data stored in each file

    << NOT Implemented >>: The email function in Watch-Win32_UPS can be adopted to email the BackupLogFile before it is deleted (for permanent archival)

 For more information type
     get-help .\Reset-Logs.ps1
 or
     .\Reset-Logs.ps1 -help
 
 The most uptodate version of this utility can be downloaded from https://github.com/chribonn/Watch-Win32_UPS

.NOTES 
    File Name  : Reset-Logs.ps1 
    Tested on  : PowerShell Version 7.1

.PARAMETER CurrentLogFile
 The name of the in-use log file

.PARAMETER BackupLogFile
 The name of the backup log file

.PARAMETER HeaderRow
 An optional string that is to be written to the newly recreated CurrentLogFile

.PARAMETER Help
 Describes this utility

.LINK
 Latest version and documentation: https://github.com/chribonn/Watch-Win32_UPS

.EXAMPLE
 * Reset E:\UPSMonitor\Watch-Win32_UPS.log backuping it up to E:\UPSMonitor\Watch-Win32_UPS.log before. 
 * The newly created E:\UPSMonitor\Watch-Win32_UPS.log will have the text specified in HeaderRow written to it.

 .\Reset-Logs.ps1 -Logfile "E:\UPSMonitor\Watch-Win32_UPS.log" -BackupLogFile E:\UPSMonitor\Watch-Win32_UPS.log" -HeaderRow "DateTime`tEventLog`tEventID`tEventMsg`tBattSysemName`tBattName`r`n"

.EXAMPLE
 * Reset E:\UPSMonitor\Watch-Win32_UPS.log. Do not back it up or write a header record

 .\Reset-Logs.ps1 -Logfile "E:\UPSMonitor\Watch-Win32_UPS.log"

.EXAMPLE
 * Reset E:\UPSMonitor\Watch-Win32_UPS.log without backing up. Write a header record. 
 
 .\Reset-Logs.ps1 -Logfile "E:\UPSMonitor\Watch-Win32_UPS.log" -HeaderRow "DateTime`tEventLog`tEventID`tEventMsg`tBattSysemName`tBattName`r`n"

#>

param (
    [Parameter(Mandatory=$true)] [string] $CurrentLogFile,
    [Parameter(Mandatory=$false)] [string] $BackupLogFile = $null,
    [Parameter(Mandatory=$false)] [string] $HeaderRow = $null,
    [switch] $help
)

if ($help) {
    write-host "Reset-Logs is a utility written and tested in Powershell script (v 7.1) that recycles a log file."
    write-host "If a backup log file name is specified the current log file is renamed rather than deleted."
    write-host "An optional header record may be written to the newly created log fileThe current log file."
}


function BackupCurrFile {
    Param(
        [Parameter(Mandatory=$true)] [string] $CurrentLogFile,
        [Parameter(Mandatory=$true)] [string] $BackupLogFile
    )

    # Delete the backup log file if it exists.
    # Add Code functionality: Add code to email the log file before deleting it
    if ((Test-Path -Path $BackupLogFile)) {
        Remove-Item "$BackupLogFile"
        
        Move-Item -Path "$CurrentLogFile" -Destination "$BackupLogFile"
    }
}


# ************************** Debug
# $DebugPreference = "Continue"

<#
    ********************** Main 
#>

New-Variable -Name LogDir
New-Variable -Name LogFile


# If the backup file has been specified and there is an existing file to backup process it
if (($PSBoundParameters.ContainsKey('BackupLogFile')) -and (Test-Path -Path "$CurrentLogFile")) {
{
    BackupCurrFile -BackupLogFile $BackupLogFile -CurrentLogFile $CurrentLogFile
}

$LogDir = Split-Path -Path "$CurrentLogFile"
$LogFile = Split-Path -Path "$CurrentLogFile" -Leaf

if (-not (Test-Path -Path $LogDir)) {
    New-Item -ItemType "directory" -Path "$LogDir"
}

if ($PSBoundParameters.ContainsKey('HeaderRow')) {
    New-Item -Path "$LogDir" -Name "$LogFile" -ItemType "file" -Value "$HeaderRow"
}
else {
    New-Item -Path "$LogDir" -Name "$LogFile" -ItemType "file" -Value ""
}

