Import-Module ".\Powershell\TestStoreFile.psm1" -Force
Import-Module ".\Powershell\Logging.psm1" -Force

function Add-StoreUser {
  param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string]$User
  )
  
  # get the storefile
  $File = Test-StoreFile

  try {
    $ErroredUsers = Get-Content -Path $File -ErrorAction Ignore
    if ($ErroredUsers) {
      [PSCustomObject]$ErroredUsers = ($ErroredUsers | ConvertFrom-Json -Depth 3)
    }

    if (!$ErroredUsers) {
      [PSCustomObject]$ErroredUsers = @{ }
    }

    if ($ErroredUsers.$($Action)) {
      [PSCustomObject]$UserObject = $ErroredUsers.$($Action)
    }
    else {
      [PSCustomObject]$UserObject = @{ }
    }

    if (!($UserObject.$User)) {
      $UserObject | Add-Member -MemberType "NoteProperty" -Name $User -Value ([int][double]::Parse((Get-Date -UFormat %s))) -Force
    }

    $ErroredUsers | Add-Member -MemberType "NoteProperty" -Name $($Action) -Value $UserObject -Force
    $ErroredUsers | ConvertTo-Json -Compress | Out-File -FilePath $File
  }
  catch { Add-LogContent -ErrorObj $_ }
}

function Remove-StoreUser {
  param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string]$Action,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string]$User
  )

  $File = Test-StoreFile

  try {
    $ErroredUsers = Get-Content -Path $File -ErrorAction Ignore
    if ($ErroredUsers) {
      [PSCustomObject]$ErroredUsers = ($ErroredUsers | ConvertFrom-Json -Depth 3)
    }
    else {
      throw "No data from file, does it exist?"
    }

    $ErroredUsers.$Action.PSObject.Members.Remove($User)
    $ErroredUsers | ConvertTo-Json -Compress | Out-File -FilePath $File
  }
  catch { Add-LogContent -ErrorObj $_ }
}

function Get-StoreUsers {
  $File = Test-StoreFile
  $ErroredUsers = Get-Content -Path $File -ErrorAction Ignore
  if ($ErroredUsers) {
    [PSCustomObject]$ErroredUsers = ($ErroredUsers | ConvertFrom-Json -Depth 3)
  }
  else {
    throw "No data from file, does it exist?"
  }

  return $ErroredUsers
}

Export-ModuleMember -Function Add-StoreUser
Export-ModuleMember -Function Remove-StoreUser
Export-ModuleMember -Function Get-StoreUsers