<#
Remove old Outlook Filed items
#>
Set-Variable -Name 'dateFormat' -Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue

$startDateTime = Get-Date
Write-Verbose "Started at: $($startDateTime.ToString($dateFormat))"

# -----------------------------
# CONFIGURATION
# -----------------------------
$rootFolderName = "Filed"       # <--- Start here
$yearsOld = 1.5                 # 1.5 is the number of years old, accounting for (most) leap years.
$daysOld = [int](((365*3 + 366) / 4.0) * $yearsOld) 
$batchSize      = 100
# -----------------------------

# Ensure required modules
[string[]] $modules = @('Microsoft.Graph', 'Microsoft.Graph.Mail')
$modules | ForEach-Object {
    if (-not (Get-InstalledModule -Name $_ -ErrorAction SilentlyContinue)) {
        Install-Module -Name $_ -Scope CurrentUser -Force -Verbose
    }
}

# Temporarily enable verbose
$vpref = $VerbosePreference
if ($VerbosePreference -ne 'Continue') { $VerbosePreference = 'Continue' }

Disconnect-MgGraph -ErrorAction SilentlyContinue -Verbose

Connect-MgGraph -Scopes 'Mail.ReadWrite','Mail.ReadWrite.Shared','User.Read' -NoWelcome -Verbose

# -----------------------------
# Resolve root folder
# -----------------------------
Write-Verbose "Resolving root folder: $rootFolderName"

$folderList = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/mailFolders?`$top=200"
$rootFolder = $folderList.value | Where-Object { $_.displayName -eq $rootFolderName }

if (-not $rootFolder) {
    Write-Error "Folder '$rootFolderName' not found."
    $VerbosePreference = $vpref
    return
}

Write-Verbose "Resolved '$rootFolderName' → ID: $($rootFolder.id)"

# -----------------------------
# Build cutoff timestamp
# -----------------------------
$cutOffIso = (Get-Date).AddDays(-$daysOld).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Verbose "Cutoff date (UTC): $cutOffIso"

[int]$globalDeleted = 0

# -----------------------------
# FUNCTION: Delete old messages in a folder
# -----------------------------
function Remove-OldMessagesFromFolder {
    param(
        [string]$FolderId,
        [string]$FolderDisplayName
    )

    Write-Verbose "Processing folder: $FolderDisplayName (ID: $FolderId)"

    $baseMessagesUri =
        "https://graph.microsoft.com/v1.0/me/mailFolders/$FolderId/messages?" +
        "`$filter=receivedDateTime lt $cutOffIso&" +
        "`$orderby=receivedDateTime asc&" +
        "`$top=$batchSize"

    do {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $baseMessagesUri
        }
        catch {
            Write-Warning "GET failed in folder '$FolderDisplayName': $($_.Exception.Message)"
            break
        }

        $messages = $response.value
        if (-not $messages -or $messages.Count -eq 0) { break }

        Write-Verbose "Found $($messages.Count) messages in '$FolderDisplayName'"

        foreach ($msg in $messages) {
            $encodedId = [System.Web.HttpUtility]::UrlEncode($msg.id)
            $deleteUrl = "https://graph.microsoft.com/v1.0/me/mailFolders/$FolderId/messages/$encodedId"

            try {
                Invoke-MgGraphRequest -Method DELETE -Uri $deleteUrl
                $globalDeleted++
            }
            catch {
                Write-Warning "Failed to delete message $($msg.id) in '$FolderDisplayName': $($_.Exception.Message)"
            }
        }

    } while ($true)
}

# -----------------------------
# FUNCTION: Recursively walk folders
# -----------------------------
function Process-FolderRecursively {
    param(
        [string]$FolderId,
        [string]$FolderDisplayName
    )

    # 1. Clean this folder
    Remove-OldMessagesFromFolder -FolderId $FolderId -FolderDisplayName $FolderDisplayName

    # 2. Get child folders
    $childUri = "https://graph.microsoft.com/v1.0/me/mailFolders/$FolderId/childFolders?`$top=200"

    try {
        $childResponse = Invoke-MgGraphRequest -Method GET -Uri $childUri
    }
    catch {
        Write-Warning "Failed to get child folders for '$FolderDisplayName': $($_.Exception.Message)"
        return
    }

    foreach ($child in $childResponse.value) {
        Process-FolderRecursively -FolderId $child.id -FolderDisplayName $child.displayName
    }
}

# -----------------------------
# START RECURSION
# -----------------------------
Process-FolderRecursively -FolderId $rootFolder.id -FolderDisplayName $rootFolder.displayName

Write-Verbose "Cleanup complete. Deleted $globalDeleted messages."

# Restore verbose preference
$VerbosePreference = $vpref

$endDateTime = Get-Date
Write-Verbose "Ended at: $($endDateTime.ToString($dateFormat))"

$elapsed = $endDateTime - $startDateTime
Write-Verbose ("Elapsed time: {0:hh\:mm\:ss}" -f $elapsed)
