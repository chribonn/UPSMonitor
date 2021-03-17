# Sends out an email to ensure that service is operational.
# It references the EmailAlert function of Watch-Win32_UPS as it is bundled with this tool but can be used to test if there is connectivity to an SMTP server

param (
    [Parameter(Mandatory=$true)] [string] $EmailTo = $null,
    [Parameter(Mandatory=$false)] [string] $EmailFromUn = $null,
    [Parameter(Mandatory=$false)] [string] $EmailFromPw = $null,
    [Parameter(Mandatory=$true)] [string] $EmailSMTP = "smtp.gmail.com",
    [Parameter(Mandatory=$false)] [ValidateRange(0, 65535)] [int] $EmailSMTPPort = 587,
    [switch] $EmailSMTPUseSSL
)


if ($EmailSMTPUseSSL) {
    $EmailSMTPUseSSL = $true
}
else {
    $EmailSMTPUseSSL = $false
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
        [parameter(Mandatory=$true)] [bool] $EmailSMTPUseSSL
    )

    if ((-not $PSBoundParameters.ContainsKey('EmailTo')) -or (-not $EmailSMTP) -or (-not $EmailSMTPPort)) {
        Write-Host "Email function failed: insufficient parameters"
    }
    
    # Do not use the -Computerhere as this function since this is more associated with email functionality rather than UPS (Battery operation)

    New-Variable -Name SecStr -Value (ConvertTo-SecureString -string $EmailFromPw -AsPlainText -Force) -Option private
    New-Variable -Name Cred -Value (New-Object System.Management.Automation.PSCredential -argumentlist $EmailFromUn, $SecStr) -Option private

    try {
        if ($EmailSMTPUseSSL) {
            Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "Test-Email $(Get-Date -Format 'yyyyMMdd HHmmss K'): $EmailSubject" -Body "$EmailBody" -Credential $Cred -SmtpServer $EmailSMTP -Port $EmailSMTPPort -UseSsl
        }
        else {
            Send-MailMessage -To $EmailTo -From $EmailFromUn -Subject "Test-Email $(Get-Date -Format 'yyyyMMdd HHmmss K'): $EmailSubject" -Body "$EmailBody" -Credential $Cred -SmtpServer EmailSMTP -Port EmailSMTPPort
        }
    }
    catch {
        Write-Host "Email function failed"
    }

}

<#
    ********************** Main 
#>

New-Variable -Name EmailDetails -Value "Testing the Email Service"
New-Variable -Name EmailSubject -Value "Email from script: Test-Email.ps1"

EmailAlert -EmailSubject $EmailSubject -EmailBody $EmailDetails -EmailTo $EmailTo -EmailFromUn $EmailFromUn -EmailFromPw $EmailFromPw -EmailSMTP $EmailSMTP -EmailSMTPPort $EmailSMTPPort -EmailSMTPUseSSL $EmailSMTPUseSSL


