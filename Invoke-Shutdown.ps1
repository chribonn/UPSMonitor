
New-Variable -Name i -Option private
New-Variable -Name j -Option private

Write-Debug 'Do some action'

# Most UPSs sound an alarm when they are on battery. If not generate your own
for ($j=1; $j -le 10; $j++) {
    for ($i=1; $i -le 10; $i++) {
        [console]::beep((800 * $i),400)
    }        
    [console]::beep(1000,500)
}

# Information on shutdown command: https://ss64.com/nt/shutdown.html or type shutdown at a command prompt

#shut down the computer in 5 seconds
#shutdown.exe /s /f /t 5 /d 6:11 /c "Watch-Win32_UPS detected power failure"

#shutdown.exe /s /f /t 5 /d 6:11 /c "Watch-Win32_UPS detected power failure"

Start-Sleep -Seconds 5
