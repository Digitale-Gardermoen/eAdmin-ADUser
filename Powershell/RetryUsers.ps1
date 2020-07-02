Import-Module ".\Powershell\Modules\UserStorage.psm1" -Force
Import-Module ".\Powershell\Modules\Logging.psm1" -Force


$Users = Get-StoreUsers

$Users."flag" | Get-Member -MemberType NoteProperty | ForEach-Object {
  try {
    $Username = $_.Name
    Remove-StoreUser -Action "flag" -User $Username
    . .\Powershell\FlagUser.ps1 -Username $Username
  }
  catch {
    Add-LogContent -ErrorObj $_
  }
}

$Users."unflag" | Get-Member -MemberType NoteProperty | ForEach-Object {
  try {
    $Username = $_.Name
    Remove-StoreUser -Action "unflag" -User $Username
    . .\Powershell\UnflagUser.ps1 -Username $Username
  }
  catch {
    Add-LogContent -ErrorObj $_
  }
}