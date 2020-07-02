function Add-LogContent {
  <#
    .SYNOPSIS
      Add a new line to the flagging log.
    .DESCRIPTION
      Logs the provided error object to a log file, creates the log file if it doesn't exist.
    .PARAMETER ErrorObj
      The object received when an Error is cast. Use the $Error object or an ErrorVariable if it is defined.
    .PARAMETER Type
      This is the type of message to add to the file.
    .PARAMETER LogPath
      Optional path to save the log. If the path doesn't exist, it creates the path.
    .EXAMPLE
      Try { throw } Catch { Add-LogContent -errorObj $Error; $Error.Clear() }
    .NOTES
      The string in the log file is of the type PSCustomObject.
  #>
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [System.Object]$ErrorObj,
    [Parameter(Mandatory = $false)]
    [String]$Type = 'ERROR',
    [Parameter(Mandatory = $false)] 
    [string]$LogPath = ".\Powershell\Log\"
  )
  
  if (!(Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory | Out-Null
  }

  [String]$logFile = "POSH_$(Get-Date -Format "yyyy-MM-dd").log"
  [String]$logFilePath = $LogPath + $logFile

  $out = [PSCustomObject]@{
    time = (Get-Date -Format "HH:mm:ss")
    type = $Type
    code = $ErrorObj.Exception.HResult
    message = $ErrorObj.Exception.Message
    stack = ($ErrorObj.ScriptStackTrace.replace("`r`n", " "))
  }

  Add-Content -Path $logFilePath -Value $out
}

Export-ModuleMember -Function Add-LogContent
