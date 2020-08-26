param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$User,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$MailDB,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$DC,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string]$Ident
)

Import-Module ".\Powershell\Modules\Logging.psm1" -Force
Import-Module ".\Powershell\Modules\GetConfig.psm1" -Force
[PSCustomObject]$Config = Get-ConfigFile

try {
  $UserCredential = Import-Clixml -Path $Config.ExchangeCredentialsPath -ErrorAction Stop
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Config.ExchangeUri -Authentication Kerberos -Credential $UserCredential -ErrorAction Stop
  Import-PSSession $Session -DisableNameChecking -ErrorAction Stop

  $migrated = Get-RemoteMailbox -Identity $User
  if ($migrated) {
    throw "The user has a migrated mailbox, the script will exit"
  }

  $exist = Get-MailboxDatabase | Get-MailboxStatistics |`
    Where-Object { ($_.DisconnectReason -eq "Disabled") -and ($_.DisplayName -eq "$Ident") }
  if (!$exist) {
    Enable-Mailbox -Identity $User -Alias $User -Database $MailDB -DomainController $DC -ErrorAction Stop
    Set-Mailbox -Identity $User -CustomAttribute2 $Config.CustomAttribute2 -DomainController $DC
    Set-MailboxRegionalConfiguration -Identity $User -Language "nb-no" -DateFormat "dd.MM.yyyy"`
      -TimeFormat "HH.mm" -TimeZone "W. Europe Standard Time" -DomainController $DC -ErrorAction Stop
  }
  else {
    Connect-Mailbox -Database $exist.Database -Identity $exist.Identity -User $User -DomainController $DC`
      -Confirm:$false -ErrorAction Stop
  }
}
catch {
  Add-LogContent $_
}
finally {
  Remove-PSSession $Session
}