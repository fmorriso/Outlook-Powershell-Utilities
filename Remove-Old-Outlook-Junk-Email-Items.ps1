<#
Remove old Outlook Junk Email items
#>
Set-Variable -Name 'dateFormat' -Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue

$startDateTime = Get-Date
Write-Verbose -Message "Started at: $($startDateTime.ToString($dateFormat))"

# -----------------------------
# CONFIGURATION
# -----------------------------
$folderName = "Junk Email"      # <--- CHANGE THIS to any folder name
$daysOld    = 7                 # <--- CHANGE THIS to desired age cutoff
$batchSize  = 100               # number of messages to fetch per loop
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
if ($VerbosePreference -ne 'Continue') {
     $VerbosePreference = 'Continue' 
}

Disconnect-MgGraph -ErrorAction SilentlyContinue -Verbose

# Connect to Graph
Connect-MgGraph -Scopes 'Mail.ReadWrite','Mail.ReadWrite.Shared','User.Read' -NoWelcome -Verbose

# -----------------------------
# Resolve folder by display name
# -----------------------------
Write-Verbose -Message "Resolving folder: $folderName"

$folderList = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/mailFolders?`$top=200"
$folder = $folderList.value | Where-Object { $_.displayName -eq $folderName }

if (-not $folder) {
    Write-Error -Message "Folder '$folderName' not found."
    $VerbosePreference = $vpref
    return
}

$folderId = $folder.id
Write-Verbose -Message "Resolved folder '$folderName' → ID: $folderId"

# -----------------------------
# Build cutoff timestamp (Graph-safe)
# -----------------------------
$cutOffIso = (Get-Date).AddDays(-$daysOld).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Verbose -Message "Cutoff date (UTC): $cutOffIso"

# -----------------------------
# Build base query
# -----------------------------
$baseMessagesUri =
    "https://graph.microsoft.com/v1.0/me/mailFolders/$folderId/messages?" +
    "`$filter=receivedDateTime lt $cutOffIso&" +
    "`$orderby=receivedDateTime asc&" +
    "`$top=$batchSize"

[int] $totalDeleted = 0

# -----------------------------
# MAIN LOOP
# -----------------------------
do {
    Write-Verbose -Message "Querying up to $batchSize messages older than cutoff..."

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $baseMessagesUri
    }
    catch {
        Write-Warning -Message "GET request failed: $($_.Exception.Message)"
        break
    }

    $messages = $response.value

    if (-not $messages -or $messages.Count -eq 0) {
        Write-Verbose "No more messages found before cutoff."
        break
    }

    Write-Verbose -Message "Found $($messages.Count) messages in this batch."

    foreach ($msg in $messages) {
        $encodedId = [System.Web.HttpUtility]::UrlEncode($msg.id)
        $deleteUrl = "https://graph.microsoft.com/v1.0/me/mailFolders/$folderId/messages/$encodedId"

        Write-Verbose -Message "Deleting message id=$($msg.id)..."

        try {
            Invoke-MgGraphRequest -Method DELETE -Uri $deleteUrl
            $totalDeleted++
        }
        catch {
            Write-Warning -Message "Failed to delete message id=$($msg.id): $($_.Exception.Message)"
        }
    }

} while ($true)

Write-Verbose "Deleted $totalDeleted messages older than $cutOffIso from $folderName"

# Restore verbose preference
$VerbosePreference = $vpref

$endDateTime = Get-Date
Write-Verbose "Ended at: $($endDateTime.ToString($dateFormat))"

$elapsed = $endDateTime - $startDateTime
$elapsedTimeDisplay =
    "Elapsed time: $($elapsed.Hours.ToString('D2')):" +
    "$($elapsed.Minutes.ToString('D2')):" +
    "$($elapsed.Seconds.ToString('D2'))"

Write-Verbose $elapsedTimeDisplay
