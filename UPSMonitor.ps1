<#
.SYNOPSIS 
    This script monitors a UPS that registers as a win32-battery device and calls an external script if conditions are met. The external script is beyond the scope of this powerscript although the sample included will shut downt he current computer.

    Other actions of the called script would be to send an alert or shutdown a group of machines.

    This script write to the event viewer and will need administrative privalages to create the entry the first time.  This can be aceived by running the script with administrative rights the first time or create it before executing the script.
    Ref: https://stackoverflow.com/questions/9564420/the-source-was-not-found-but-some-or-all-event-logs-could-not-be-searched
    
.DESCRIPTION 
    This script periodically polls the win32-battery for status and availability conditions. If triggered it will report to the user and shut down the computer

    Battery Reference: https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-battery

    **************************
    Event Log

    The script writes to the log and (optionally) sends an email. The log name is defined by 'EventLogName'. If the EventLog does not exists it needs to be created.
 
    The IDs are documented under the EventType ennumeration.
         
.NOTES 
    File Name  : UPSMonitor.ps1 
    Requires   : PowerShell Version 7.1
    
.PARAMETER TriggerShutdownPerc
The Percentage remaining charge before triggering a shutdown
.PARAMETER TriggerShutDownRunTime
The remaining battery capacity in minutes before triggering a shutdown
.PARAMETER EmailTo
The email address to receive notifications
.PARAMETER EmailFromUn
The SMTP server sender Username
.PARAMETER EmailFromPw
The SMTP server sender Password
.PARAMETER EmailSMTP
The SMTP server address
.PARAMETER EmailSMTPPort
The SMTP server port
.PARAMETER EmailSMTPUseSSL
(1 <default> / or 0)
.PARAMETER LogDir
The directory where the log file will be retained
.PARAMETER LogFile
The name of the log file. Do not specify to disable logging
.PARAMETER PollFrequency
The time in seconds between script iterations
Default is 5 seconds
.PARAMETER ShutdownScript
This is the name of the Shutdown Script that will be called when computers are to be shut down.
The full path to the script must be specified. If no file is specified the file will not be called.
.PARAMETER Help
Show this help text.
#>

param (
    [Parameter(Mandatory=$false)] [ValidateRange(1, 100)] [int] $TriggerShutdownPerc = 50,
    [Parameter(Mandatory=$false)] [ValidateRange(1, 999)] [int] $TriggerShutDownRunTime = 20,
    [Parameter(Mandatory=$false)] [string] $EmailTo = $null,
    [Parameter(Mandatory=$false)] [string] $EmailFromUn = $null,
    [Parameter(Mandatory=$false)] [string] $EmailFromPw = $null,
    [Parameter(Mandatory=$false)] [string] $EmailSMTP = "smtp.gmail.com",
    [Parameter(Mandatory=$false)] [ValidateRange(0, 65535)] [int] $EmailSMTPPort = 587,
    [Parameter(Mandatory=$false)] [bool] $EmailSMTPUseSSL = 1,
    [Parameter(Mandatory=$false)] [string] $LogDir = $null,
    [Parameter(Mandatory=$false)] [string] $LogFile = $null,
    [Parameter(Mandatory=$false)] [string] $PollFrequency = 5,
    [Parameter(Mandatory=$false)] [string] $ShutdownScript = $null,
    [switch] $help
)

if ($help) {
    write-host 'This script is used to monitor your windows installed UPS and allows you to programatically set it to shutdown your computers. '
    Write-Host '<<add parameter information>>'
}

New-Variable -Name CodeRef -Value "UPSMonitor" -Option Constant
New-Variable -Name UPSMonitor_Version -Value "0.1" -Option Constant
New-Variable -Name EventLogName -Value "UPSMonitor" -Option Constant
New-Variable -Name NL -Value "`r`n" -Option Constant 

<#
    The existance of the file below will cause the script to loop. 
    To stop the script gracefully delete this file. 
#>
New-Variable -Name RunStateFile -Value "UPSMonitor.run" -Option Constant

[Flags()]
enum AlertType {
    NoAlert = 0             # Nothing to report
    Information = 1         # Reporting a state such as the service running, and minor secondary events
    Warning = 2             # Reproting a situation that may lead to an action. This is typically a power failure
    Error = 4               # Reporting an Error situation.
    Action = 8              # Reporting a situation that is tied to an activity. The most common action would be to initiate power down
}

enum EventType {
    ProgramStart = 1000 
    Information = 1001
    Warning = 1002
    Action = 1004
    Error = 1006
    ProgramStop = 1050 
    UnsupportedOS = 1096
    InvalidShutdownScript = 1097
    NoBattery = 1098
    EmailFailed = 1099
}

function WriteEventLog {
    Param(
        [parameter(Mandatory=$true)] [String] $EventLog,
        [parameter(Mandatory=$true)] [String] $EventID,
        [parameter(Mandatory=$true)] [String] $EventMsg,
        [parameter(Mandatory=$false)] [String] $BattSystemName,
        [parameter(Mandatory=$false)] [String] $BattName,
        [parameter(Mandatory=$true)] [String] $LogDir,
        [parameter(Mandatory=$true)] [String] $LogFile
    )

    #   If no log file is specified exit
    if (!$LogFile) {
        exit
    }

    if (-not (Test-Path -Path $LogDir)) {
        New-Item -ItemType "directory" -Path "$LogDir"
    }

    <# 
        Write to the log. 
        If it doesn't exist create it with the header record
    #>
    if (-not (Test-Path -Path "$LogPath")) {
        New-Item -Path "$LogDir" -Name "$LogFile" -ItemType "file" -Value "DateTime`tEventLog`tEventID`tEventMsg`tBattSysemName`tBattName$NL"
    }

    if (-not $PSBoundParameters.ContainsKey('BattSystemName')) {
        $BattSystemName = ""
    }

    if (-not $PSBoundParameters.ContainsKey('BattName')) {
        $BattName = ""
    }

    Add-Content -Path "$LogPath" -Value "$(Get-Date -Format 'yyyyMMdd HH:mm:ss K')`t$EventLog`t$EventID`t$EventMsg`t$BattSystemName`t$BattName$NL"
}

function EmailAlert {
    Param(
        [parameter(Mandatory=$true)] [String] $EmailSubject,
        [parameter(Mandatory=$true)] [String] $EmailBody,
        [parameter(Mandatory=$false)][String] $EmailTo,
        [parameter(Mandatory=$true)] [String] $EmailFromUn,
        [parameter(Mandatory=$true)] [String] $EmailFromPw,
        [parameter(Mandatory=$true)] [String] $EmailSMTP, 
        [parameter(Mandatory=$true)] [String] $EmailSMTPPort,
        [parameter(Mandatory=$true)] [String] $EmailSMTPUseSSL,
        [parameter(Mandatory=$false)] [String] $EventLogName = $EventLogName,
        [parameter(Mandatory=$true)] [String] $LogDir,
        [parameter(Mandatory=$true)] [String] $LogFile
    )

    if (-not $PSBoundParameters.ContainsKey('EmailTo')) {
        exit
    }
    
    # Do not use the -Computerhere as this function since this is more associated with email functionality rather than UPS (Battery operation)

    New-Variable -Name SecStr -Value (ConvertTo-SecureString -string $EmailFromPw -AsPlainText -Force) -Option private
    New-Variable -Name Cred -Value (New-Object System.Management.Automation.PSCredential -argumentlist $EmailFromUn, $SecStr) -Option private

    try {
        if ($EmailSMTPUseSSL) {
            Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "$CodeRef $(Get-Date -Format 'yyyyMMdd HHmmss K'): $EmailSubject" -Body "$EmailBody" -Credential $Cred -SmtpServer $EmailSMTP -Port $EmailSMTPPort -UseSsl
        }
        else {
            Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "$CodeRef $(Get-Date -Format 'yyyyMMdd HHmmss K'): $EmailSubject" -Body "$EmailBody" -Credential $Cred -SmtpServer EmailSMTP -Port EmailSMTPPort
        }
    }
    catch {
        WriteEventLog -EventLog $EventLogName -EventID $([EventType]::EmailFailed) -EventMsg "Failed to send $CodeRef email to $EmailTo." -LogDir $LogDir -LogFile $LogFile
    }

}

function InvokeShutdownScript{
    Param(
        [parameter(Mandatory=$true)] [String] $ShutdownScript,
        [parameter(Mandatory=$true)] [String] $EventLogName,
        [parameter(Mandatory=$true)] [String] $EventMsg,
        [parameter(Mandatory=$false)] [String] $BattSystemName = $false,
        [parameter(Mandatory=$false)] [String] $BattName = $false,
        [parameter(Mandatory=$true)] [String] $LogDir,
        [parameter(Mandatory=$true)] [String] $LogFile
    )

	Write-Debug 'Calling the shutdown script'
	#   If the Shutdown script is valid invoke it
	if ($ShutdownScript) {
		WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg "Invoking Shutdown Script @($ShutdownScript)" -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
        &$ShutdownScript

        # if the script performs an actual shutdown the script will not return here
		WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg "Shutdown Script @($ShutdownScript) finished executing" -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
	}
	else {
		WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg "Function InvokeShutdownScript but no script called." -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
	}
	
	return
}

function ValidDateShutdownScript {
    <#
        .SYNOPSIS
        Verifies that the passed ShitDownScript external file exists and logs the event.

        .DESCRIPTION
        if the parameter ShutdownScript is not null the function verifies that the files does exist (no internal content inspection).
        If the file does not exist and alert is emailed and logged and the variable is nulified (thereby ensuring that the process continues)
    #>
    
    New-Variable -Name EmailSubject -Value $null -Option private
    New-Variable -Name EmailBody -Value $null -Option private
    New-Variable -Name EventLogMsg -Value $false -Option private

    #   If no log file is specified exit
    if ($ShutdownScript) {
        if (-not (Test-Path -Path "$ShutdownScript")) {
            $EmailSubject = "[InvalidShutdownScript] ShutdownScript " + $ShutdownScript + " not found"
            $EmailBody = "Shutdown action will not take place."
            $EventLogMsg = "ShutdownScript " + $ShutdownScript + " not found"
    
            EmailAlert -EmailSubject $EmailSubject -EmailBody $EmailBody -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile
            WriteEventLog -EventLog $EventLogName -EventID $([EventType]::InvalidShutdownScript) -EventMsg $EventLogMsg -LogDir $LogDir -LogFile $LogFile
            $ShutdownScript = $null

            Write-Debug -Message "ShutdownScript $ShutdownScript not found"
        }
    }
    
    return $ShutdownScript
}

function BattNotFound {
    <#
        .SYNOPSIS
        Returns whether the battery has been detected.

        .DESCRIPTION
        Check the various conditions that may indicate that a battery has not been found / or has been unplugged. 
        Report the conditions (multiple) that will trigger this event.

        Each condition will trigger its own event.
    #>

    New-Variable -Name EmailSubject -Value $null -Option private
    New-Variable -Name EmailBody -Value $null -Option private
    New-Variable -Name NoBatteryFound -Value $false -Option private
    New-Variable -Name EventLogMsg -Value $false -Option private

    $Battery = Get-CimInstance -ClassName win32_battery

    if ($Battery.BatteryStatus -eq 10) {
        $EmailSubject = "[Error] Battery Not Found - Script Terminating"
        $EmailBody = "Battery.SystemName: " + $Battery.SystemName + $NL + "Battery Name: " + $Battery.Name + $NL + "Battery Status: " + $Battery.BatteryStatus + $NL + "Script Terminating"
        $EventLogMsg = "Battery Not Found (Battery Status: " + $Battery.BatteryStatus + ")"

        Write-Debug "$EmailSubject $NL $EmailBody $NL Script Terminating"

        EmailAlert -EmailSubject "$EmailSubject" -EmailBody "$EmailBody" -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile

        WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Error) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
        $NoBatteryFound = $true
    }

    if ($Battery.Count -eq 0) {
        $EmailSubject = "[Error] Battery Not Found - Script Terminating"
        $EmailBody = "Battery.SystemName: " + $Battery.SystemName + $NL + "Battery Name: " + $Battery.Name + $NL + "Battery Count: " + $Battery.Count + $NL + "Script Terminating"
        $EventLogMsg = "Battery Not Found (Battery Count: " + $Battery.Count + ")"

        Write-Debug "$EmailSubject $NL $EmailBody $NL Script Terminating"

        EmailAlert -EmailSubject "$EmailSubject" -EmailBody "$EmailBody" -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile

        WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Error) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
        $NoBatteryFound = $true
    }

    if ($Battery.Availability -eq 11) {
        $EmailSubject = "[Error] Battery Not Found - Script Terminating"
        $EmailBody = "Battery.SystemName: " + $Battery.SystemName + $NL + "Battery Name: " + $Battery.Name + $NL + "Battery Availability: " + $Battery.Availability + $NL + "Script Terminating"
        $EventLogMsg = "Battery Not Found (Battery Availability: " + $Battery.Availability + ")"

        Write-Debug "$EmailSubject $NL $EmailBody $NL Script Terminating"

        EmailAlert -EmailSubject "$EmailSubject" -EmailBody "$EmailBody" -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile

        WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Error) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
        $NoBatteryFound = $true
    }

    return $NoBatteryFound
}

# ************************** Debug
$DebugPreference = "Continue"

<#
    ********************** Main 
#>

New-Variable -Name Details
New-Variable -Name Subject
New-Variable -Name EventLogMsg

# Script currently only works with Windows OS
Write-Debug -Message "Is this OS suppored?"
if (-not $IsWindows) {
	$EventLogMsg = "Runing on a Non-Windows OS." + $NL + "Script Terminating"
    $EmailSubject = "[Error] Unsupported OS - Script Terminating"
    $EmailDetails = $EventLogMsg

	EmailAlert -EmailSubject $EmailSubject -EmailBody $EmailDetails -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile
    WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Error) -EventMsg $EventLogMsg -LogDir $LogDir -LogFile $LogFile

    Write-Debug $EventLogMsg
	
	exit
}

Write-Debug -Message "Program starting"

#   If no log file is specified exit
if (-not $PSBoundParameters.ContainsKey('LogFile')) {
    $LogFile = $null
}
else {
    #   If the directory for the log file has not been specified, assume it is current directory
    if (-not $PSBoundParameters.ContainsKey('LogDir')) {
        $LogDir = Get-Location
    }

	New-Variable -Name LogPath -Value @(Join-Path -Path "$LogDir" -ChildPath "$LogFile") -Option Constant 
}

# Create the run file. While this file exists the script will loop
$RunFilePath = Join-Path -Path "$(Get-Location)" -ChildPath "$RunStateFile"
if (-not (Test-Path -Path "$RunFilePath")) {
    New-Item -Path "$(Get-Location)" -Name "$RunStateFile" -ItemType "file" -Value "$NL"
}

$Battery = Get-CimInstance -ClassName win32_battery

$EventLogMsg = "UPS Monitor Started (Name: " + $Battery.Name + " / Device ID: " + $Battery.DeviceID
$EmailSubject = "[ProgramStart] Script has started"
$EmailDetails = "Script Name: " + $CodeRef + " (version " + $UPSMonitor_Version + ")" +$NL + 
    "Battery Name: " + $Battery.Name + $NL +
    "Device ID: " + $Battery.DeviceID + $NL +
    "To cause the script to exit delete the file " + $RunStateFile
if ($LogDir) {
    $EmailDetails = $EmailDetails + $NL + "Log file is at: " + $LogDir
}
if ($ShutdownScript) {
    $EmailDetails = $EmailDetails + $NL + "Shutdown script resides at: " + $LogDir + $NL + "ShutdownScript name: " + $ShutdownScript
}
else {
    $EmailDetails = $EmailDetails + $NL + "No shutdown script defined."
    
}

WriteEventLog -EventLog $EventLogName -EventID $([EventType]::ProgramStart) -EventMsg $EventLogMsg -LogDir $LogDir -LogFile $LogFile
EmailAlert -EmailSubject $EmailSubject -EmailBody $EmailDetails -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile

$ShutdownScript = ValidDateShutdownScript

if (BattNotFound) {
    exit
}

[bool] $UPSonAC = ($Battery.BatteryStatus -eq 2)                # Is the UPS plugged into the power socket?
[bool] $UPSFullycharged = ($Battery.BatteryStatus -eq 3)        # Is the UPS Fully charged?

do {
    $EmailSubject = ''
    $EmailDetails = ''
    [AlertType] $Notification = [AlertType]::NoAlert

    Write-Debug -Message "Battery BatteryStatus : $Battery.BatteryStatus"

    Switch ($Battery.BatteryStatus) {
        # The battery is discharging
        1 {
            $UPSFullycharged = ($Battery.BatteryStatus -eq 3)
                
            # If the UPS was previosuly plugged in than alert about a state change
            if ($UPSonAC) {
                $UPSonAC = 0
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }
                
                $EmailSubject = "$EmailSubject Powerfailure".Trim()
                $EventLogMsg = "The computer is running on battery following power failure."
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Warning) {
                    $Notification = [AlertType]::Warning
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Warning) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }
            else {
                # Log the current state of the discharing battery to the event log
                $EventLogMsg = "Battery Estimated Remaing Charge: $Battery.EstimatedChargeRemaining % (Minutes: $Battery.EstimatedRunTime)."
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }

            # Has a shutdown trigger point been reached?
            if (($Battery.EstimatedChargeRemaining -le $TriggerShutdownPerc) -or ($Battery.EstimatedRunTime -le $TriggerShutDownRunTime)) {
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }
                
                $EmailSubject = "$EmailSubject Shutdown started".Trim()
                $EventLogMsg = "System is being shut down."
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Action) {
                    $Notification = [AlertType]::Action
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile

                Write-Debug 'Calling the shutdown script'
                #   If the Shutdown script is valid invoke it
                if ($ShutdownScript) {
                    $EventLogMsg = "Invoking Shutdown Script " + $ShutdownScript
                    WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
                    &$ShutdownScript
                    $EventLogMsg = "Shutdown Script " + $ShutdownScript + " finished executing"
                    WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
                    Remove-Item -Path "$RunFilePath" -Force
                    break
                }

            }
            break;
        }
        2 {
            # The system is being powered by the AC.
            if (-not $UPSonAC) {
                $UPSonAC = 1;
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }

                $EmailSubject = "$EmailSubject AC restored".Trim()
                $EventLogMsg = "AC Restored."
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }

            # Is the UPS fully charged
            if ((-not $UPSFullycharged) -and ($Battery.EstimatedChargeRemaining -eq 100)) {
                $UPSFullycharged = 1
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }

                $EmailSubject = "$EmailSubject Battery Fully Charged".Trim()
                $EventLogMsg = "The battery is fully charged."
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }
            break;
        }
        3 {
            # Is the UPS reporting a state of fully charged
            if (-not $UPSFullycharged) {
                if ($Battery.EstimatedChargeRemaining -eq 100) {
                    $UPSFullycharged = 1
                    if ($EmailSubject -ne '') {
                        $EmailSubject = "$EmailSubject /"
                        $EmailDetails = "$EmailDetails $NL"
                    }
                    $EmailSubject = "$EmailSubject UPS Fully Charged".Trim()
                    $EventLogMsg = "The battery is fully charged."
                                    
                    $EmailDetails = "$EmailDetails $EventLogMsg"
                    if ($Notification -lt [AlertType]::Information) {
                        $Notification = [AlertType]::Information
                    }

                    WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
                }
                else {
                    # This condition should not occur: the Battery reports that it is fully charged while the EstimatedChargeRemaining is not 100
                    # If it happens report it    
                    if ($EmailSubject -ne '') {
                        $EmailSubject = "$EmailSubject /"
                        $EmailDetails = "$EmailDetails $NL"
                    }

                    $EmailSubject = "$EmailSubject Conflict in UPS charge state".Trim()
                    $EventLogMsg = "The battery charge cannot be accuratly determined."
                    $EmailDetails = "$EmailDetails $EventLogMsg"
                    if ($Notification -lt [AlertType]::Warning) {
                        $Notification = [AlertType]::Warning
                    }

                    WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Warning) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
                }
   
            }
            break;
        }
        4 {
            # Low Status - Initiate shutdown.
            if ($EmailSubject -ne '') {
                $EmailSubject = "$EmailSubject /"
                $EmailDetails = "$EmailDetails $NL"
            }

            $EmailSubject = "$EmailSubject Shutdown started - Battery Status Low".Trim()
            $EventLogMsg = "System is being shut down due to status of LOW."
            $EmailDetails = "$EmailDetails $EventLogMsg"
            if ($Notification -lt [AlertType]::Action) {
                $Notification = [AlertType]::Action
            }

            WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            
			InvokeShutdownScript -ShutdownScript $ShutdownScript -EventLogName $EventLogName -EventMsg $EventMsg -BattSystemName $BattSystemName -BattName $BattName -LogDir $LogDir -LogFile $LogFile
			break;
        }
        5 {
            # Critical Status - Initiate shutdown.
            if ($EmailSubject -ne '') {
                $EmailSubject = "$EmailSubject /"
                $EmailDetails = "$EmailDetails $NL"
            }

            $EmailSubject = "$EmailSubject Shutdown started - Battery Status Critical".Trim()
            $EventLogMsg = "System is being shut down due to status of CRITICAL."
            $EmailDetails = "$EmailDetails $EventLogMsg"
            if ($Notification -lt [AlertType]::Action) {
                $Notification = [AlertType]::Action
            }

            WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            
			# Initiate Shutdown
			InvokeShutdownScript -ShutdownScript $ShutdownScript -EventLogName $EventLogName -EventMsg $EventMsg -BattSystemName $BattSystemName -BattName $BattName -LogDir $LogDir -LogFile $LogFile
            break;
        }
        6 {
            # Charging. Alert the first time the UPS switches from on battery to on charge state
            if (-not $UPSonAC) {
                $UPSonAC = ($Battery.BatteryStatus -eq 2)                # Is the UPS plugged into the power socket?
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }

                $EmailSubject = "$EmailSubject UPS Charging".Trim()
                $EventLogMsg = "The battery is charging"
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }
            break;
        }
        7 {
            # Charging and High
            if (-not $UPSonAC) {
                $UPSonAC = ($Battery.BatteryStatus -eq 2)                # Is the UPS plugged into the power socket?
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }

                $EmailSubject = "$EmailSubject UPS Charging".Trim()
                $EventLogMsg = "The battery is charging"
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }
            break;
        }
        8 {
            # Charging and High
            if (-not $UPSonAC) {
                $UPSonAC = ($Battery.BatteryStatus -eq 2)                # Is the UPS plugged into the power socket?
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }

                $EmailSubject = "$EmailSubject UPS Charging".Trim()
                $EventLogMsg = "The battery is charging"
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }
            break;
        }
        9 {
            # Charging and Critical
            if (-not $UPSonAC) {
                $UPSonAC = ($Battery.BatteryStatus -eq 2)                # Is the UPS plugged into the power socket?
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }

                $EmailSubject = "$EmailSubject UPS Charging".Trim()
                $EventLogMsg = "The battery is charging"
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }
            break;
        }
        11 {
            # Partially Charged
            if (-not $UPSonAC) {
                $UPSonAC = ($Battery.BatteryStatus -eq 2)                # Is the UPS plugged into the power socket?
                if ($EmailSubject -ne '') {
                    $EmailSubject = "$EmailSubject /"
                    $EmailDetails = "$EmailDetails $NL"
                }

                $EmailSubject = "$EmailSubject UPS Charging".Trim()
                $EventLogMsg = "The battery is charging"
                $EmailDetails = "$EmailDetails $EventLogMsg"
                if ($Notification -lt [AlertType]::Information) {
                    $Notification = [AlertType]::Information
                }

                WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Information) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            }
            break;
        }
    }

    Write-Debug -Message "Battery Availability : $Battery.Availability"
    
    # Checks based on battery availability status
    Switch ($Battery.Availability) {
        4 {
            # Warning
            if ($EmailSubject -ne '') {
                $EmailSubject = "$EmailSubject /"
                $EmailDetails = "$EmailDetails $NL"
            }

            $EmailSubject = "$EmailSubject Check battery".Trim()
            $EventLogMsg = "The Battery is in a state of warning."
            $EmailDetails = "$EmailDetails $EventLogMsg"
            if ($Notification -lt [AlertType]::Warning) {
                $Notification = [AlertType]::Warning
            }

            WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Warning) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            break
        }
        10 {
            # Degraded
            if ($EmailSubject -ne '') {
                $EmailSubject = "$EmailSubject /"
                $EmailDetails = "$EmailDetails $NL"
            }
            $EmailSubject = "$EmailSubject Battery is in a Degraded state"
            $EventLogMsg = "The Battery is in a degraded state."
                
            $EmailDetails = "$EmailDetails $EventLogMsg"
            if ($Notification -lt [AlertType]::Action) {
                $Notification = [AlertType]::Action
            }

            WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Action) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            break;
        }
        12 {
            # Install Error
            if ($EmailSubject -ne '') {
                $EmailSubject = "$EmailSubject /"
                $EmailDetails = "$EmailDetails $NL"
            }

            $EmailSubject = "$EmailSubject Battery Installation Error".Trim()
            $EventLogMsg = "The Battery is not installed properly."
            $EmailDetails = "$EmailDetails $EventLogMsg"
            if ($Notification -lt [AlertType]::Action) {
                $Notification = [AlertType]::Action
            }

            WriteEventLog -EventLog $EventLogName -EventID $([EventType]::Error) -EventMsg $EventLogMsg -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
            break
        }
    }

    if ($Notification -ne [AlertType]::NoAlert) {
        EmailAlert -EmailSubject "[$Notification] $Battery.SystemName/$Battery.Name : $EmailSubject" -EmailBody $EmailDetails  -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile
    }

    Start-Sleep -Seconds $PollFrequency

    Write-Debug -Message "Restarting Loop"
    # Refresh the battery instance
    $Battery = Get-CimInstance -ClassName win32_battery

    if (BattNotFound) {
        exit
    }
    
} while (Test-Path -Path "$RunFilePath")

Write-Debug -Message "Program exiting"

EmailAlert -EmailSubject "[Information] $CodeRef : Script has stopped" -EmailBody "Terminated. GoodBye."  -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL -LogDir $LogDir -LogFile $LogFile
WriteEventLog -EventLog $EventLogName -EventID $([EventType]::ProgramStop) -EventMsg "Script Exiting" -BattSystemName $Battery.SystemName -BattName $Battery.Name -LogDir $LogDir -LogFile $LogFile
