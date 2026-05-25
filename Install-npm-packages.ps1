<#
    Purpose: install multiple npm packages, echoing the command prior to executing it via PowerShell's built-in Write-Verbose command,
             with care taken to verify the npm cache after each install, a lesson learned the hard way.
    Author: Fred Morrison  
    Location: C:\Users\frede\OneDrive\PowerShell\Install-npm-packages.ps1
#>
$previousVerbosePreference = $VerbosePreference
if ($VerbosePreference -ne 'Continue') { $VerbosePreference = 'Continue' }

# set the date/time display format to be similar to ISO 8601
Set-Variable -Name 'dateFormat' -Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue
Get-Date -Format $dateFormat

Get-Location -Verbose
# MODIFY THE FOLLOWING COMMA-SEPARATED LIST WITH PACKAGE NAMES THAT NEED TO BE UPDATED TO THE LATEST VERSION
[string[]] $cmds = @('webpack')
[string] $cmd = ''
$cmds | ForEach-Object  -Process {    
  $cmd = 'npm install --location=global {0}@latest' -f $_
  Write-Verbose -Message $cmd
  Invoke-Command -ScriptBlock ([ScriptBlock]::Create($cmd)) -Verbose
  
  Get-Date -Format $dateFormat
  
  $cmd = 'npm cache verify'
  Write-Verbose -Message $cmd
  Invoke-Command -ScriptBlock ([ScriptBlock]::Create($cmd)) -Verbose
  
  Get-Date -Format $dateFormat
}

$VerbosePreference = $previousVerbosePreference
