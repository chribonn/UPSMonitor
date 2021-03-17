# Dumps the content of this object
New-Variable -Name ObjName

$Battery = Get-CimInstance -ClassName win32_battery
$Battery.PSObject.Properties | ForEach-Object {
    $ObjName = $_.Name
    Write-Host ([string]$ObjName).PadLeft(25,' ') :`t $_.Value
}