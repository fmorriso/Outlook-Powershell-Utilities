<#
List unread Outlook messages (recursively) and generate Outlook.com links
#>

Set-Variable -Name 'dateFormat' -Value 'yyyy-MM-dd HH:mm:ss' -ErrorAction SilentlyContinue

$startDateTime = Get-Date
Write-Verbose "Started at: $($startDateTime.ToString($dateFormat))"

# -----------------------------
# CONFIGURATION
# -----------------------------
$rootFolderNames = @("Filed", "Inbox", "Junk Email")
$batchSize      = 100
# -----------------------------

# -----------------------------
# HTML BUFFER + COLLAPSIBLE UI
# -----------------------------
$global:Html = @()
$global:Html += "<html><head><meta charset='UTF-8'>"
$global:Html += "<style>
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
Connect-MgGraph -Scopes 'Mail.Read','Mail.Read.Shared','User.Read' -NoWelcome -Verbose

# Force authentication to complete
Get-MgContext | Out-Null

# -----------------------------
# Resolve all root folders
# -----------------------------
Write-Verbose "Resolving root folders..."

$folderList = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/mailFolders?`$top=200"

# -----------------------------
# FUNCTION: Show unread messages
# -----------------------------
function Show-UnreadMessagesFromFolder {
    param(
        [string]$FolderId,
        [string]$FolderDisplayName   # full breadcrumb path
    )

    Write-Verbose "Checking unread messages in: $FolderDisplayName"

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
        Write-Warning "Failed to query unread messages in '$FolderDisplayName': $($_.Exception.Message)"
        return
    }

    $messages = $response.value
    if (-not $messages -or $messages.Count -eq 0) {
        Write-Verbose "No unread messages in '$FolderDisplayName'"
        return
    }

    # Console output
    Write-Host ""
    Write-Host "📁 Folder: $FolderDisplayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------"

    # Collapsible section
    $safeFolder = [System.Web.HttpUtility]::HtmlEncode($FolderDisplayName)
    $global:Html += "<button class='collapsible'>$safeFolder</button>"
    $global:Html += "<div class='content'>"

    # Table start
    $global:Html += "<table>"
    $global:Html += "<tr><th>Received</th><th>From</th><th>Subject</th><th>Open</th></tr>"

    foreach ($msg in $messages) {

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
        $global:Html += "<tr>"
        $global:Html += "<td>$safeDate</td>"
        $global:Html += "<td>$safeFrom</td>"
        $global:Html += "<td>$safeSubject</td>"
        $global:Html += "<td><a href='$webLink' target='_blank' rel='noopener noreferrer'>Open</a></td>"
        $global:Html += "</tr>"
    }

    # Close table + collapsible content
    $global:Html += "</table></div>"
}

# -----------------------------
# FUNCTION: Recursively walk folders (with breadcrumb path)
# -----------------------------
function Get-FolderRecursively {
    param(
        [string]$FolderId,
        [string]$FolderDisplayName,
        [string]$BreadcrumbPath
    )

    Show-UnreadMessagesFromFolder -FolderId $FolderId -FolderDisplayName $BreadcrumbPath

    $childUri = "https://graph.microsoft.com/v1.0/me/mailFolders/$FolderId/childFolders?`$top=200"

    try {
        $childResponse = Invoke-MgGraphRequest -Method GET -Uri $childUri
    }
    catch {
        Write-Warning "Failed to get child folders for '$FolderDisplayName': $($_.Exception.Message)"
        return
    }

    foreach ($child in $childResponse.value) {

        $childPath = "$BreadcrumbPath → $($child.displayName)"

        Get-FolderRecursively `
            -FolderId $child.id `
            -FolderDisplayName $child.displayName `
            -BreadcrumbPath $childPath
    }
}

# -----------------------------
# START RECURSION FOR ALL ROOTS
# -----------------------------
foreach ($rootName in $rootFolderNames) {

    $rootFolder = $folderList.value | Where-Object { $_.displayName -eq $rootName }

    if (-not $rootFolder) {
        Write-Warning "Folder '$rootName' not found — skipping."
        continue
    }

    Write-Verbose "Resolved '$rootName' → ID: $($rootFolder.id)"

    Get-FolderRecursively `
        -FolderId $rootFolder.id `
        -FolderDisplayName $rootFolder.displayName `
        -BreadcrumbPath $rootFolder.displayName
}

Write-Verbose "Unread message scan complete."

# Restore verbose preference
$VerbosePreference = $vpref

$endDateTime = Get-Date
Write-Verbose "Ended at: $($endDateTime.ToString($dateFormat))"

$elapsed = $endDateTime - $startDateTime
Write-Verbose ("Elapsed time: {0:hh\:mm\:ss}" -f $elapsed)

# -----------------------------
# WRITE HTML FILE
# -----------------------------
$global:Html += "</body></html>"

$outFile = Join-Path $PSScriptRoot "UnreadMessages.html"

if (Test-Path $outFile) {
    Remove-Item $outFile -Force
}

$global:Html -join "`n" | Set-Content -Path $outFile -Encoding UTF8

Start-Process $outFile
