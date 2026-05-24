<#
List unread Outlook messages (recursively) and generate Outlook.com links
Includes: Mark All As Read (PowerShell-side REST PATCH)
#>

Set-Variable -Name 'dateFormat' -Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue

$startDateTime = Get-Date
Write-Verbose "Started at: $($startDateTime.ToString($dateFormat))"

# -----------------------------
# CONFIGURATION
# -----------------------------
$rootFolderNames = @('Filed', 'Inbox', 'Junk Email')
$batchSize      = 100
# -----------------------------

# -----------------------------
# HTML BUFFER + COLLAPSIBLE UI
# -----------------------------
$Html = @()
$Html += "<html><head><meta charset='UTF-8'>"
$Html += "<style>
body {
    font-family: Arial, sans-serif;
    font-size: 1rem;
    line-height: 1.4;
}
h1 {
    margin-bottom: 1.25rem;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin: 0.6rem 0;
}
th, td {
    border: 1px solid #ccc;
    padding: 0.45rem;
}
th {
    background: #f0f0f0;
}
.collapsible {
    background-color: #0078D4;
    color: white;
    cursor: pointer;
    padding: 0.65rem;
    width: 100%;
    border: none;
    text-align: left;
    outline: none;
    font-size: 1rem;
    margin-top: 0.6rem;
    border-radius: 0.2rem;
}
.active, .collapsible:hover {
    background-color: #005A9E;
}
.content {
    padding: 0 0.6rem;
    display: none;
    overflow: hidden;
    background-color: #f9f9f9;
}
</style>

<script>
document.addEventListener('DOMContentLoaded', function() {
    var coll = document.getElementsByClassName('collapsible');
    for (var i = 0; i < coll.length; i++) {
        coll[i].addEventListener('click', function() {
            this.classList.toggle('active');
            var content = this.nextElementSibling;
            if (content.style.display === 'block') {
                content.style.display = 'none';
            } else {
                content.style.display = 'block';
            }
        });
    }
});
</script>

<title>Unread Outlook Messages</title></head><body>
<h1>Unread Outlook Messages</h1>
"

# -----------------------------
# Ensure required modules
# -----------------------------
[string[]] $modules = @('Microsoft.Graph', 'Microsoft.Graph.Mail')
$modules | ForEach-Object {
    if (-not (Get-InstalledModule -Name $_ -ErrorAction SilentlyContinue)) {
        Install-Module -Name $_ -Scope CurrentUser -Force -Verbose
    }
}

# Temporarily enable verbose
$vpref = $VerbosePreference
if ($VerbosePreference -ne 'Continue') { $VerbosePreference = 'Continue' }

# Connect to Graph
Disconnect-MgGraph -ErrorAction SilentlyContinue -Verbose
Connect-MgGraph -Scopes 'Mail.Read','Mail.Read.Shared','Mail.ReadWrite','User.Read' -NoWelcome -Verbose

# Force authentication to complete
Get-MgContext | Out-Null

# -----------------------------
# FUNCTION: Mark messages as read
# -----------------------------
function Set-MessagesRead {
    param(
        [string[]]$MessageIds
    )

    foreach ($id in $MessageIds) {
        $uri = "https://graph.microsoft.com/v1.0/me/messages/$id"
        try {
            Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body @{ isRead = $true } | Out-Null
            Write-Verbose "Marked message $id as read"
        }
        catch {
            Write-Warning "Failed to mark $id as read: $($_.Exception.Message)"
        }
    }
}

# -----------------------------
# FUNCTION: Show unread messages
# -----------------------------
function Show-UnreadMessagesFromFolder {
    param(
        [string]$FolderId,
        [string]$FolderDisplayName,
        [ref]$Html
    )

    Write-Verbose -Message "Checking unread messages in: $FolderDisplayName"

    $uri =
        "https://graph.microsoft.com/v1.0/me/mailFolders/$FolderId/messages?" +
        "`$filter=isRead eq false&" +
        "`$orderby=receivedDateTime desc&" +
        "`$top=$batchSize&" +
        "`$select=receivedDateTime,subject,webLink,id,from"

    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    }
    catch {
        Write-Warning -Message "Failed to query unread messages in '$FolderDisplayName': $($_.Exception.Message)"
        return
    }

    $messages = $response.value
    if (-not $messages -or $messages.Count -eq 0) {
        Write-Verbose "No unread messages in '$FolderDisplayName'"
        return
    }

    # Collect IDs for Mark All As Read
    $messageIds = @()

    # Console output
    Write-Host ""
    Write-Host "📁 Folder: $FolderDisplayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------"

    # Collapsible section
    $safeFolder = [System.Web.HttpUtility]::HtmlEncode($FolderDisplayName)
    $Html.Value += "<button class='collapsible'>$safeFolder</button>"
    $Html.Value += "<div class='content'>"

    # Mark All As Read button
    $Html.Value += "<button onclick=""location.href='markread://$FolderId'"" 
                    style='margin:0.5rem 0; padding:0.4rem; background:#d9534f; color:white; border:none; border-radius:4px;'>
                    Mark All As Read
                    </button>"

    # Table start
    $Html.Value += "<table>"
    $Html.Value += "<tr><th>Received</th><th>From</th><th>Subject</th><th>Open</th></tr>"

    foreach ($msg in $messages) {

        $messageIds += $msg.id

        # Convert Graph datetime to local timezone
        $receivedRaw   = [datetimeoffset]$msg.receivedDateTime
        $receivedLocal = $receivedRaw.ToLocalTime().DateTime
        $receivedText  = $receivedLocal.ToString('yyyy-MM-dd HH:mm:ss')

        # Sender info
        $fromName  = $msg.from.emailAddress.name
        $fromEmail = $msg.from.emailAddress.address

        if ([string]::IsNullOrWhiteSpace($fromName)) {
            $fromDisplay = $fromEmail
        } else {
            $fromDisplay = "$fromName <$fromEmail>"
        }

        $safeFrom     = [System.Web.HttpUtility]::HtmlEncode($fromDisplay)
        $safeSubject  = [System.Web.HttpUtility]::HtmlEncode($msg.subject)
        $safeDate     = [System.Web.HttpUtility]::HtmlEncode($receivedText)
        $webLink      = $msg.webLink

        # Console output
        Write-Host "• $receivedText — $fromDisplay — $($msg.subject)"
        Write-Host "  $webLink" -ForegroundColor Yellow
        Write-Host ""

        # HTML row
        $Html.Value += "<tr>"
        $Html.Value += "<td>$safeDate</td>"
        $Html.Value += "<td>$safeFrom</td>"
        $Html.Value += "<td>$safeSubject</td>"
        $Html.Value += "<td><a href='$webLink' target='_blank' rel='noopener noreferrer'>Open</a></td>"
        $Html.Value += "</tr>"
    }

    # Close table + collapsible content
    $Html.Value += "</table></div>"
}

# -----------------------------
# FUNCTION: Get all child folders
# -----------------------------
function Get-ChildFoldersPaged {
    param(
        [string]$ParentFolderId
    )

    $allChildren = @()
    $childUri = "https://graph.microsoft.com/v1.0/me/mailFolders/$ParentFolderId/childFolders?`$top=200"

    while ($childUri) {
        try {
            $childResponse = Invoke-MgGraphRequest -Method GET -Uri $childUri
        }
        catch {
            Write-Warning -Message "Failed to get child folders for '$ParentFolderId': $($_.Exception.Message)"
            break
        }

        if ($childResponse.value) {
            $allChildren += $childResponse.value
        }

        $childUri = $childResponse.'@odata.nextLink'
    }

    return $allChildren
}

# -----------------------------
# HANDLE markread:// PROTOCOL
# -----------------------------
if ($args.Count -gt 0 -and $args[0].StartsWith("markread://")) {

    $folderId = $args[0].Substring("markread://".Length)

    Write-Host "Re-querying unread messages for folder: $folderId"

    $uri = "https://graph.microsoft.com/v1.0/me/mailFolders/$folderId/messages?`$filter=isRead eq false&`$select=id&`$top=999"
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri

    $ids = $resp.value.id
    if ($ids.Count -gt 0) {
        Write-Host "Marking $($ids.Count) messages as read..."
        Set-MessagesRead -MessageIds $ids
    } else {
        Write-Host "No unread messages found."
    }

    exit
}

# -----------------------------
# PAGE THROUGH ALL ROOT FOLDERS
# -----------------------------
Write-Verbose -Message "Resolving ALL root folders (paged)..."

$folderList = @()
$next = "https://graph.microsoft.com/v1.0/me/mailFolders?`$top=999"

while ($next) {
    $resp = Invoke-MgGraphRequest -Method GET -Uri $next
    $folderList += $resp.value
    $next = $resp.'@odata.nextLink'
}

Write-Verbose -Message "Total root-level folders retrieved: $($folderList.Count)"

# -----------------------------
# PROCESS ROOT + CHILD FOLDERS
# -----------------------------
foreach ($rootName in $rootFolderNames) {
    $root = $folderList | Where-Object { $_.displayName -eq $rootName }
    if (-not $root) { continue }

    Show-UnreadMessagesFromFolder -FolderId $root.id -FolderDisplayName $root.displayName -Html ([ref]$Html)

    $children = Get-ChildFoldersPaged -ParentFolderId $root.id
    foreach ($child in $children) {
        Show-UnreadMessagesFromFolder -FolderId $child.id -FolderDisplayName $child.displayName -Html ([ref]$Html)
    }
}

# -----------------------------
# WRITE HTML OUTPUT
# -----------------------------
$Html += "</body></html>"

$outFile = Join-Path (Get-Location) "UnreadMessages.html"
Set-Content -Path $outFile -Value $Html -Encoding UTF8

Write-Host ""
Write-Host "Output written to: $outFile" -ForegroundColor Green
Start-Process $outFile
