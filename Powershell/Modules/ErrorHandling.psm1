# errorhandling for the cleanup API
# this module will store the user and the action it took
# in the case the flagUser and unflagUser scripts doesn't get contact with the API
# We store the user and action in a JSON file and run it later.
function New-ErrorEvent {
  <#
  .SYNOPSIS
  Create a new error event
  .DESCRIPTION
  Adds the user to the offline file for sending to the API later. This is meant to be ran when sending to the API triggers an error.
  .PARAMETER Action
      The action tried to be performed. Usually "Flag" or "Unflag".
    .PARAMETER User
      The username that got the error
      .PARAMETER ErrorObj
      The object received when an Error is cast. Use the $Error object or $_ or the ErrorVariable if it is defined.
    .PARAMETER LogPath
    Optional path to save the log. If the path doesn't exist, it creates the path.
  #>
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string]$User,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [System.Object]$ErrorObj,
    [Parameter(Mandatory = $false)]
    [string]$LogPath
  )
    
  # import the logging module
  Import-Module ".\Powershell\Modules\Logging.psm1" -Force
  Import-Module ".\Powershell\Modules\UserStorage.psm1" -Force
    
  Add-LogContent -ErrorObj $ErrorObj
  Add-StoreUser -Action $Action -User $User
}

Export-ModuleMember -Function New-ErrorEvent