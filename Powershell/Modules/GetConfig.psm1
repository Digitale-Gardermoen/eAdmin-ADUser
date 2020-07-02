function Get-ConfigFile {
  <#
  .SYNOPSIS
    Get the config.
  .DESCRIPTION
    Get the config file parse the JSON object.
  .OUTPUTS
    PSCustomObject. Configuration properties.
  .EXAMPLE
    Get-ConfigFile
  #>
  try {
    $config = Get-Content -Path ".\Powershell\Config\Config.json" -ErrorAction Stop
    [PSCustomObject]$config = ($config | ConvertFrom-Json)
    return $config
  }
  catch {
    Import-Module ".\Powershell\Modules\Logging.psm1" -Force
    Add-LogContent -ErrorObj $_
  }
}

Export-ModuleMember -Function Get-ConfigFile