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

# This is for fixing a bug where running the Set-MailboxRegionalConfiguration CmdLet would throw an error:
# Set-MailboxRegionalConfiguration : Operation is not valid due to the current state of the object.
# Not entierly sure why this happens, but it seems to be related to the exhchange scripting Agent
# and how remote exchange works.
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -Auto -AllowClobber

Import-Module ".\Powershell\Logging.psm1" -Force

$migrated = Get-RemoteMailbox -Identity $User
if ($migrated) {
  try {
    throw "The user has a migrated mailbox, the script will exit"
  }
  catch {
    Add-LogContent $_
  }
  finally {
    Exit
  }
}

$exist = Get-MailboxDatabase $MailDB | Get-MailboxStatistics |`
  Where-Object { ($_.DisconnectReason -eq "Disabled") -and ($_.DisplayName -eq "$Ident") }

try {
  if (!$exist) {
    Enable-Mailbox -Identity $User -Alias $User -Database $MailDB -DomainController $DC -ErrorAction Stop
    Set-MailboxRegionalConfiguration -Identity $User -Language "nb-no" -DateFormat "dd.MM.yyyy"`
      -TimeFormat "HH:mm" -TimeZone "W. Europe Standard Time" -ErrorAction Stop
  }
  else {
    Connect-Mailbox -Database $MailDB -Identity $Ident -User $User -DomainController $DC`
      -Confirm:$false -ErrorAction Stop
  }
}
catch {
  Add-LogContent $_
}