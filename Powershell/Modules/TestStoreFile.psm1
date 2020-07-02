Import-Module ".\Powershell\Modules\GetConfig.psm1" -Force
Import-Module ".\Powershell\Modules\Logging.psm1" -Force

function Test-StoreFile {
  # get the config
  $Config = Get-ConfigFile
  $StorePath = $Config.storePath
  $File = $Config.storeFile

  # check if the file exists
  # if it doesnt exists check the path
  # create the file or path if its missing
  if (!(Test-Path "$($StorePath)\$($File)")) {
    if (!(Test-Path $StorePath)) {
      try {
        # Out-Null this because powershell takes the returned
        # object and appends it to the function output.
        New-Item $StorePath -ItemType Directory | Out-Null
      }
      catch { Add-LogContent -ErrorObj $Error[0] }
    }

    try {
      $File = New-Item "$StorePath\$File" -ItemType File
    }
    catch { Add-LogContent -ErrorObj $Error[0] }
  }
  else {
    $File = Get-Item "$($StorePath)\$($File)"
  }

  return $File
}

Export-ModuleMember -Function Test-StoreFile