param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$User,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$DC
)

try {
  Import-Module ActiveDirectory
  $address = Get-AdUser -Identity $User -Server $DC -Properties proxyAddresses | Select-Object -Expand proxyAddresses | Where-Object {$_ -clike "SMTP:*"}
  $address = $address.SubString(5)
  Set-ADUser -Identity $User -UserPrincipalName $address -Server $DC
}
catch {
  Import-Module ".\Powershell\Logging.psm1" -Force
  Add-LogContent $_
}