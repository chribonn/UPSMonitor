<#  
.SYNOPSIS 
    Reset-Logs takes a mandatory parameter:
        * Logfile = Name of the file to reset

    It takes two optional parameters:
        * BackupLogFile = Name of the back log file. If specified the original log file is renamed to this file before it is reset. If the backup log file already exists it will be deleted.
        * HeaderRow = If specified the reset log file will have this text written as a header record
        * Email credentials and server information - If specified, the old log file will be emailed before it is renamed
        * Help = Describes this utility

    This script may be invoked via a task schedular task. The frequency of execution (weekly, monthly, quarterly, anually) dictates the the amount of data stored in each file

 For more information type
     get-help .\Reset-Logs.ps1
 or
     .\Reset-Logs.ps1 -help
 
 The most uptodate version of this utility can be downloaded from https://github.com/chribonn/Reset-Logs

.NOTES 
    File Name  : Reset-Logs.ps1 
    Tested on  : PowerShell Version 7.1

.PARAMETER CurrentLogFile
 The name of the in-use log file

.PARAMETER BackupLogFile
 The name of the backup log file

.PARAMETER HeaderRow
 An optional string that is to be written to the newly recreated CurrentLogFile

.PARAMETER EmailTo
 The email address to receive notifications

.PARAMETER EmailFromUn
 The SMTP server sender Username

.PARAMETER EmailFromPw
 The SMTP server sender Password

.PARAMETER EmailSMTP
 The SMTP server address (default - smtp.gmail.com)

.PARAMETER EmailSMTPPort
 The SMTP server port (default - 587)

.PARAMETER EmailSMTPUseSSL
 Specifies if the SMTP client requires SSL

.PARAMETER Help
 Describes this utility

.LINK
 Latest version and documentation: https://github.com/chribonn/Watch-Win32_UPS

.EXAMPLE
 * Reset E:\UPSMonitor\Watch-Win32_UPS.log backuping it up to E:\UPSMonitor\Watch-Win32_UPS.log before. 
 * The newly created E:\UPSMonitor\Watch-Win32_UPS.log will have the text specified in HeaderRow written to it.

 .\Reset-Logs.ps1 -Logfile "E:\UPSMonitor\Watch-Win32_UPS.log" -BackupLogFile E:\UPSMonitor\Watch-Win32_UPS.log" -HeaderRow "DateTime`tEventLog`tEventID`tEventMsg`tBattSysemName`tBattName`r`n"

.EXAMPLE
 * Reset E:\UPSMonitor\Watch-Win32_UPS.log. Do not back it up or write a header record. Email the log file.

 .\Reset-Logs.ps1 -Logfile "E:\UPSMonitor\Watch-Win32_UPS.log" -EmailTo "<notification email>" -EmailFromUn "<send email username>" -EmailFromPw "<send email password>" -EmailSMTP "smtp.gmail.com" -EmailSMTPPort 587 -EmailSMTPUseSSL

.EXAMPLE
 * Reset E:\UPSMonitor\Watch-Win32_UPS.log without backing up. Write a header record. 
 
 .\Reset-Logs.ps1 -Logfile "E:\UPSMonitor\Watch-Win32_UPS.log" -HeaderRow "DateTime`tEventLog`tEventID`tEventMsg`tBattSysemName`tBattName`r`n"

#>

param (
    [Parameter(Mandatory=$true)] [string] $CurrentLogFile,
    [Parameter(Mandatory=$false)] [string] $BackupLogFile = $null,
    [Parameter(Mandatory=$false)] [string] $HeaderRow = $null,
    [Parameter(Mandatory=$false)] [string] $EmailTo = $null,
    [Parameter(Mandatory=$false)] [string] $EmailFromUn = $null,
    [Parameter(Mandatory=$false)] [string] $EmailFromPw = $null,
    [Parameter(Mandatory=$false)] [string] $EmailSMTP = "smtp.gmail.com",
    [Parameter(Mandatory=$false)] [ValidateRange(0, 65535)] [int] $EmailSMTPPort = 587,
    [switch] $EmailSMTPUseSSL,
    [switch] $help
)

if ($help) {
    write-host "Reset-Logs is a utility written and tested in Powershell script (v 7.1) that recycles a log file."
    write-host "If a backup log file name is specified the current log file is renamed rather than deleted."
    write-host "An optional header record may be written to the newly created log fileThe current log file."
}

if ($EmailSMTPUseSSL) {
    $EmailSMTPUseSSL = $true
}
else {
    $EmailSMTPUseSSL = $false
}

New-Variable -Name CodeRef -Value "Reset-Logs" -Option Constant
New-Variable -Name Code_Version -Value "0.1" -Option Constant

function EmailCurrFile {
    Param (
        [Parameter(Mandatory=$true)] [string] $CurrentLogFile,
        [Parameter(Mandatory=$false)] [string] $EmailTo = $null,
        [Parameter(Mandatory=$false)] [string] $EmailFromUn = $null,
        [Parameter(Mandatory=$false)] [string] $EmailFromPw = $null,
        [Parameter(Mandatory=$false)] [string] $EmailSMTP = $null,
        [Parameter(Mandatory=$false)] [ValidateRange(0, 65535)] [int] $EmailSMTPPort,
        [parameter(Mandatory=$false)] [bool] $EmailSMTPUseSSL
    )

    New-Variable -Name SecStr -Value $null -Option private
    New-Variable -Name Cred -Value $null -Option private
    New-Variable -Name EmailSubject -Value $null -Option private
    New-Variable -Name EmailBody -Value $null -Option private

    # Email the attachment only if it exits and if there are email credentials
    if ((-not $EmailTo) -or (-not $EmailSMTP) -or (-not $EmailSMTPPort)) {
        exit
    }
        
    # If there should be an attachment check that it exists
    if ($PSBoundParameters.ContainsKey('CurrentLogFile')) {
        if (-not (Test-Path -Path "$CurrentLogFile")) {
            exit
        }
    }

    if (($EmailFromUn) -and ($EmailFromPw)) {
        $SecStr = $(ConvertTo-SecureString -string $EmailFromPw -AsPlainText -Force)
        $Cred = $(New-Object System.Management.Automation.PSCredential -argumentlist $EmailFromUn, $SecStr)
    }
    
    $EmailSubject = $CodeRef + $(Get-Date -Format 'yyyyMMdd HHmmss K') + ": Emailing Log File"
    $EmailBody = "Attached to this email is the log file " + $CurrentLogFile +". This file has been reset"

    if (($EmailFromUn) -and ($EmailFromPw)) {
        try {
            if ($EmailSMTPUseSSL) {
                Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "$EmailSubject" -Body "$EmailBody" -Credential $Cred -SmtpServer $EmailSMTP -Port $EmailSMTPPort -UseSsl -Attachments "$EmailAttachment"
            }
            else {
                Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "$EmailSubject" -Body "$EmailBody" -Credential $Cred -SmtpServer $EmailSMTP -Port $EmailSMTPPort -Attachments "$EmailAttachment"
            }
        }
        catch {
        }
    }
    else {
        # This code block is called if the email SMTP server does not require credentials
        try {
            if ($EmailSMTPUseSSL) {
                Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "$EmailSubject" -Body "$EmailBody" -SmtpServer $EmailSMTP -Port $EmailSMTPPort -UseSsl -Attachments "$EmailAttachment"
            }
            else {
                Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "$EmailSubject" -Body "$EmailBody" -SmtpServer $EmailSMTP -Port $EmailSMTPPort -Attachments "$EmailAttachment"
            }
        }
        catch {
        }
    }
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

New-Variable -Name LogDir -Value $null -Option Private
New-Variable -Name LogFile -Value $null -Option Private

# If the backup file has been specified and there is an existing file to backup process it
if (($PSBoundParameters.ContainsKey('BackupLogFile')) -and (Test-Path -Path "$CurrentLogFile")) {
    EmailCurrFile -CurrentLogFile $CurrentLogFile -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL
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

