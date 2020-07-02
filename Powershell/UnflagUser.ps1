<#
  .SYNOPSIS
    Send a DELETE request to the cleanup API
  .DESCRIPTION
    Send a user to delete from the cleanup app.
  .PARAMETER Username
    The user to delete
  .EXAMPLE
    UnflagUser.ps1 -Username "testuser"
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
  New-ErrorEvent -action "unflag" -user $Username -errorObj $_
  Exit
}

Import-Module ".\Powershell\Modules\GetConfig.psm1" -Force
[PSCustomObject]$Config = Get-ConfigFile
$uri = "$($Config.ApiUrl)$Username"

try {
  Invoke-RestMethod `
    -uri $uri `
    -Credential $cred `
    -Method 'DELETE' `
    -ErrorAction Stop
}
catch {
  Import-Module ".\Powershell\Modules\ErrorHandling.psm1" -Force
  if ($_.Exception.Response.StatusCode -eq "NotFound") {
    Import-Module ".\Powershell\Modules\Logging.psm1" -Force
    $Log = [PSCustomObject]@{
      Exception = @{
        HResult = $_.Exception.HResult
        Message = "Got a NotFound error: The user '$Username' is not flagged, ignoring this error."
      }
      ScriptStackTrace = $_.ScriptStackTrace
    }
    Add-LogContent -ErrorObj $Log -Type "WARN"
  }
  else {
    New-ErrorEvent -action "unflag" -user $Username -errorObj $_
  }
}
