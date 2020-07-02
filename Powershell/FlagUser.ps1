<#
  .SYNOPSIS
    Send a POST request to the cleanup API
  .DESCRIPTION
    Send a user to add to the cleanup app.
  .PARAMETER Username
    The user to add
  .EXAMPLE
    FlagUser.ps1 -Username "testuser"
  #>
param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [String]$Username
)

try {
  $cred = Import-Clixml -Path ".\Powershell\Config\Credentials_$($env:USERNAME)_$($env:COMPUTERNAME).xml" -ErrorAction Stop
}
catch {
  Import-Module ".\Powershell\Modules\ErrorHandling.psm1" -Force
  New-ErrorEvent -action "flag" -user $Username -errorObj $_
  Exit
}

Import-Module ".\Powershell\Modules\GetConfig.psm1" -Force
[PSCustomObject]$Config = Get-ConfigFile
$uri = "$($Config.ApiUrl)$Username"

try {
  Invoke-RestMethod `
    -uri $uri `
    -Credential $cred `
    -Method 'POST' `
    -ErrorAction Stop
}
catch {
  Import-Module ".\Powershell\Modules\ErrorHandling.psm1" -Force
  if ($_.Exception.Response.StatusCode -eq "Conflict") {
    Import-Module ".\Powershell\Modules\Logging.psm1" -Force
    $Log = [PSCustomObject]@{
      Exception = @{
        HResult = $_.Exception.HResult
        Message = "Got a Conflig error: The user '$Username' is already flagged."
      }
      ScriptStackTrace = $_.ScriptStackTrace
    }
    Add-LogContent -ErrorObj $Log -Type "WARN"
  }
  else {
    New-ErrorEvent -action "flag" -user $Username -errorObj $_
  }
}
