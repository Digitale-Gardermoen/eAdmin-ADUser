param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$User,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$DC
)

Import-Module ".\Powershell\Modules\Logging.psm1" -Force
Import-Module ".\Powershell\Modules\GetConfig.psm1" -Force
[PSCustomObject]$Config = Get-ConfigFile

try {
  $UserCredential = Import-Clixml -Path $Config.ExchangeCredentialsPath -ErrorAction Stop
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Config.ExchangeUri -Authentication Kerberos -Credential $UserCredential -ErrorAction Stop
  Import-PSSession $Session -DisableNameChecking -ErrorAction Stop
  
  Disable-Mailbox -Identity $User -DomainController $DC -Confirm:$false -ErrorAction Stop
}
catch {
  Add-LogContent $_
}
finally {
  Remove-PSSession $Session
}