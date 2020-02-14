param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$User,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$DC
)

Add-PSSnapin *ex*
Import-Module ".\Powershell\Logging.psm1" -Force

try {
  Disable-Mailbox -Identity $User -DomainController $DC -Confirm:$false -ErrorAction Stop
}
catch {
  Add-LogContent $_
}