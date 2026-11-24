<#
Remove attachments based on whether messages are from me or to me.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('FromMe', 'ToMe', 'CcMe')]
    [string[]]$Direction = @('FromMe'),

    [Parameter()]
    [string]$FolderName = 'Inbox',

    [Parameter()]
    [datetime]$ReceivedBefore = (Get-Date).AddDays((-7),

    [Parameter()]
    [long]$MinimumAttachmentBytes = 1MB,

    [Parameter()]
    [string]$ReportPath = (
        Join-Path $HOME 'Documents/Outlook-Attachment-Report.csv'
    ),

    [Parameter()]
    [switch]$DeleteAttachments,

    [Parameter()]
    [switch]$Commit
)

enum MessageDirection {
    FromMe
    ToMe
    CcMe
}

[MessageDirection[]]$SelectedDirections = $Direction
