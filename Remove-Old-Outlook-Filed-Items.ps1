function Get-MailFolderRecursively {param ([string]$parentFolderId)

    $prevSetting = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'

    $folders = @()
    $uri = "https://graph.microsoft.com/v1.0/me/mailFolders/$parentFolderId/childFolders"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $folders += $response.value
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    foreach ($folder in $folders) {
        $folders += Get-MailFolderRecursively -parentFolderId $folder.id
    }

    $VerbosePreference = $prevSetting
    return $folders
}

<#
 START OF MAIN PROGRAM 
#>

Set-Variable -Name 'dateFormat'-Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue

$PSVersionTable.DotNetVersion = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
Write-Verbose -Message "Using PowerShell version $($PSVersionTable.PSVersion) on $($PSVersionTable.Platform) with $($PSVersionTable.DotNetVersion)"

$startDateTime = Get-Date
Write-Verbose -Message "Started at: $($startDateTime.ToString($dateFormat))"

# Requires Microsoft.Graph module
# Install-Module Microsoft.Graph -Scope CurrentUser
# Import-Module Microsoft.Graph <--- exceeds 4096 limit - do not use
$vpref = $VerbosePreference
# temporarily change verbose preference so we can see what's happening
if ($VerbosePreference -ne 'Continue') {
    $VerbosePreference = 'Continue'
}

[string] $moduleName = 'Microsoft.Graph.Mail'
if (-not (Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue)) {
    Install-Module -Name $moduleName -Scope CurrentUser -Force
}
Import-Module -Name $moduleName



Disconnect-MgGraph -ErrorAction SilentlyContinue

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Mail.ReadWrite","Mail.ReadWrite.Shared","User.Read" -NoWelcome

# Find "Filed" folder
$folder = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me/mailFolders' | 
    Select-Object -ExpandProperty value | 
    Where-Object { $_.displayName -eq "Filed" }

if (-not $folder) {
    Write-Error "Filed folder not found."
    return
}

Write-Verbose -Message 'Found Filed folder'

# Calculate cutoff date
# Number of days old before hard delete
$daysOld = 270
$cutoffDate = (Get-Date).AddDays(-$daysOld).ToString("o") # ISO 8601 format
Write-Verbose -Message "cut off date: $cutOffDate"

# Batch size for deletions
$batchSize = 50

# Get all subfolders recursively
$allFolders = @($folder)
$allFolders += Get-MailFolderRecursively -parentFolderId $folder.id
$folderCount = $allFolders.Count
Write-Verbose -Message "all folders contains $folderCount entries"

[int]$totalDeleted = 0
foreach ($folder in $allFolders) {
    Write-Verbose -Message "Processing folder: $($folder.displayName)"
    $messagesUri = "https://graph.microsoft.com/v1.0/me/mailFolders/$($folder.id)/messages?`$filter=receivedDateTime lt $cutoffDate&`$top=$batchSize"

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $messagesUri
        $messages = $response.value

        foreach ($msg in $messages) {
            Write-Verbose "Deleting message ID: $($msg.id) - Subject: $($msg.subject)"
            $deleteUri = "https://graph.microsoft.com/v1.0/me/messages/$($msg.id)?$deleteType=hardDelete"
            $deleteResponse = Invoke-MgGraphRequest -Method DELETE -Uri $deleteUri -OutputType HttpResponseMessage
            if ($deleteResponse.IsSuccessStatusCode) {
                $totalDeleted++
            }
        }

        $messagesUri = $response.'@odata.nextLink'
    } while ($messagesUri)
}

Write-Verbose -Message "Cleanup of $totalDeleted messages complete."

# restore user's original verbose preference
$VerbosePreference = $vpref

$endDateTime = Get-Date 
Write-Verbose -Message "Ended at:  $($endDateTime.ToString($dateFormat))"
$elapsed = $endDateTime - $startDateTime  # This creates a TimeSpan object

# Display the TimeSpan
$elapsedTimeDisplay = "Elapsed time: $($elapsed.Hours.ToString("D2")):$($elapsed.Minutes.ToString("D2")):$($elapsed.Seconds.ToString("D2"))"
Write-Verbose -Message $elapsedTimeDisplay
