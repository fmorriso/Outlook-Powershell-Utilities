<#
Remove old Deleted Items from Outlook using Microsoft Graph API
and a configurable number of days old as the cutoff. 
This script will hard delete items older than the cutoff date.
#>
Set-Variable -Name 'dateFormat'-Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue

$PSVersionTable.DotNetVersion = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
Write-Verbose -Message "Using PowerShell version $($PSVersionTable.PSVersion) on $($PSVersionTable.Platform) with  $($PSVersionTable.DotNetVersion)"

$startDateTime = Get-Date
Write-Verbose -Message "Started at: $($startDateTime.ToString($dateFormat))"

# Requires Microsoft.Graph module
# Install-Module Microsoft.Graph -Scope CurrentUser
# Import-Module Microsoft.Graph <--- exceeds 4096 limit - do not use
$vpref = $VerbosePreference
$VerbosePreference = 'SilentlyContinue'

[string] $moduleName = 'Microsoft.Graph.Mail'
if (-not (Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue)) {
    Install-Module -Name $moduleName -Scope CurrentUser -Force
}
Import-Module -Name $moduleName


# temporarily change verbose preference so we can see what's happening
if ($VerbosePreference -ne 'Continue') {
    $VerbosePreference = 'Continue'
}

Disconnect-MgGraph -ErrorAction SilentlyContinue

# Connect to Microsoft Graph
Connect-MgGraph -Scopes 'Mail.ReadWrite', 'Mail.ReadWrite.Shared', 'User.Read' -NoWelcome

# Get Deleted Items folder info
$folderUri = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me/mailFolders/DeletedItems'
if (-not $folderUri) {
    Write-Error "Deleted Items folder not found."
    return
}
$folderUriId = $folderUri.id
Write-Verbose -Message "Deleted Items folder ID: $folderUriId"

# Calculate cutoff date
# Number of days old before hard delete
$daysOld = 30
$cutoffDate = (Get-Date).AddDays(-$daysOld).ToString("o") # ISO 8601 format
Write-Verbose -Message "cut off date: $cutOffDate"

# Batch size for deletions
$batchSize = 50

# Get all messages in Deleted Items folder (pagination needed if many)
$messagesUri = "https://graph.microsoft.com/v1.0/me/mailFolders/$folderUriId/messages?`$filter=receivedDateTime lt $cutoffDate&`$top=$batchSize"

[int] $count = 0
do {
    $httpGetResponse = Invoke-MgGraphRequest -Method GET -Uri $messagesUri -Verbose
    $messages = $httpGetResponse.value

    foreach ($msg in $messages) {
        $count++
        Write-Verbose -Message "Deleting message ID: $($msg.id) - Subject: $($msg.subject)"
        $httpDeleteResponse = Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/me/messages/$($msg.id)?$deleteType=hardDelete" -OutputType HttpResponseMessage -Verbose
        Write-Verbose -Message "HTTP Delete StatusCode: $httpDeleteResponse.StatusCode"
        Write-Verbose -Message "HTTP Delete attempt True/False was $httpDeleteResponse.IsSuccessStatusCode"
    }

    $messagesUri = $httpGetResponse.'@odata.nextLink'
} while ($messagesUri)

Write-Verbose -Message "Cleanup of $count messages complete."

# restore user's original verbose preference
$VerbosePreference = $vpref

$endDateTime = Get-Date 
Write-Verbose -Message "Ended at:  $($endDateTime.ToString($dateFormat))"
$elapsed = $endDateTime - $startDateTime  # This creates a TimeSpan object

# Display the TimeSpan
$elapsedTimeDisplay = "Elapsed time: $($elapsed.Hours.ToString("D2")):$($elapsed.Minutes.ToString("D2")):$($elapsed.Seconds.ToString("D2"))"
Write-Verbose -Message $elapsedTimeDisplay
