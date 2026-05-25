<#
    Purpose: install multiple npm packages, echoing the command prior to executing it via PowerShell's built-in Write-Verbose command,
             with care taken to verify the npm cache after each install, a lesson learned the hard way.
    Author: Fred Morrison  
    Location: C:\Users\frede\OneDrive\PowerShell\Install-npm-packages.ps1

    MODIFY THE FOLLOWING COMMA-SEPARATED LIST WITH PACKAGE NAMES THAT NEED TO BE UPDATED TO THE LATEST VERSION
#>
param(
    [string[]] $cmds = @('webpack')
)
$previousVerbosePreference = $VerbosePreference
$VerbosePreference = 'Continue'
# set the date/time display format to be similar to ISO 8601
Set-Variable -Name 'dateFormat' -Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue
$startDateTime = Get-Date 
$displayDateTime = $startDateTime.ToString($dateFormat)
Write-Verbose -Message "Started at $displayDateTime"

Get-Location -Verbose

[string] $cmd = ''
$cmds | ForEach-Object  -Process {    
  $cmd = 'npm install --location=global {0}@latest' -f $_
  Write-Verbose -Message $cmd
  Invoke-Command -ScriptBlock ([ScriptBlock]::Create($cmd))
  
  Get-Date -Format $dateFormat
  
  $cmd = 'npm cache verify'
  Write-Verbose -Message $cmd
  Invoke-Command -ScriptBlock ([ScriptBlock]::Create($cmd))
  
  Get-Date -Format $dateFormat
}

$endDateTime = Get-Date
$displayDateTime = $endDateTime.ToString($dateFormat)
Write-Verbose -Message "Ended at $displayDateTime"

[TimeSpan] $elapsed = $endDateTime - $startDateTime
$elapsedFormatted = '{0:0000}-{1:00}:{2:00}:{3:00}' -f `
    $elapsed.Days,
    $elapsed.Hours,
    $elapsed.Minutes,
    $elapsed.Seconds
	
Write-Verbose -Message "Elapsed time (days-hh:mm:ss) $elapsedFormatted"

# restore user's verbose preference
$VerbosePreference = $previousVerbosePreference
