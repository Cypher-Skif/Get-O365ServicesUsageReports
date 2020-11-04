$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
write-host "Enter your ID and Secret"
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$Credentials = Get-Credential       # Set your id and secret
$Credentials | Export-Clixml -Path "$scriptPath\Secret.xml" -Confirm:$false
 