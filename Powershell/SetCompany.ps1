param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$User,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$DC
)

Import-Module ".\Powershell\Logging.psm1" -Force

$depObj = Get-Content ".\Powershell\Config\Companies.json" | ConvertFrom-Json

if ($User -notlike "*-*") {
  $Unit = "Unknown"
}
else {
  $Unit = $User.Split('-')[1]
  $Unit = $Unit.ToUpper()
  $Unit = $depObj.$Unit
}

try {
  Import-Module ActiveDirectory -ErrorAction Stop
  Set-ADUser -Identity $User -Company $Unit -Server $DC -ErrorAction Stop
}
catch {
  Add-LogContent $_
}