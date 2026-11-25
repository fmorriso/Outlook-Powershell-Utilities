#requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('FromMe', 'ToMe', 'CcMe')]
    [string[]]$Direction = 'FromMe',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string[]]$FolderPath = @('Filed', 'YeBu'),

    [Parameter()]
    [datetime]$ReceivedBefore = (Get-Date).AddYears(-3),

    [Parameter()]
    [ValidateRange(0, [long]::MaxValue)]
    [long]$MinimumAttachmentBytes = 1MB,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath = (Join-Path $HOME 'Documents/Outlook-Attachment-Report.csv'),

    [Parameter()]
    [switch]$DeleteAttachments,

    [Parameter()]
    [switch]$Commit
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

enum MessageDirection {
    FromMe
    ToMe
    CcMe
}

[MessageDirection[]]$SelectedDirections = $Direction

function Escape-ODataLiteral {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)
    $Value.Replace("'", "''")
}

function Test-EmailAddressEqual {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Address1,
        [AllowNull()][string]$Address2
    )

    if ([string]::IsNullOrWhiteSpace($Address1) -or
        [string]::IsNullOrWhiteSpace($Address2)) {
        return $false
    }

    $Address1.Equals($Address2, [StringComparison]::OrdinalIgnoreCase)
}

function Get-MessageDirectionMatches {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Message,
        [Parameter(Mandatory)][string]$MyAddress,
        [Parameter(Mandatory)][MessageDirection[]]$SelectedDirection
    )

    $matches = [System.Collections.Generic.List[string]]::new()

    $fromAddress = if ($null -ne $Message.From -and
                       $null -ne $Message.From.EmailAddress) {
        $Message.From.EmailAddress.Address
    }
    else {
        $null
    }

    $fromMe = Test-EmailAddressEqual -Address1 $fromAddress -Address2 $MyAddress

    $toMe = $false
    foreach ($recipient in @($Message.ToRecipients)) {
        if ($null -ne $recipient -and
            $null -ne $recipient.EmailAddress -and
            (Test-EmailAddressEqual -Address1 $recipient.EmailAddress.Address -Address2 $MyAddress)) {
            $toMe = $true
            break
        }
    }

    $ccMe = $false
    foreach ($recipient in @($Message.CcRecipients)) {
        if ($null -ne $recipient -and
            $null -ne $recipient.EmailAddress -and
            (Test-EmailAddressEqual -Address1 $recipient.EmailAddress.Address -Address2 $MyAddress)) {
            $ccMe = $true
            break
        }
    }

    foreach ($selected in $SelectedDirection) {
        if ($selected -eq [MessageDirection]::FromMe -and $fromMe) {
            $matches.Add('FromMe')
        }
        elseif ($selected -eq [MessageDirection]::ToMe -and $toMe) {
            $matches.Add('ToMe')
        }
        elseif ($selected -eq [MessageDirection]::CcMe -and $ccMe) {
            $matches.Add('CcMe')
        }
    }

    $matches.ToArray()
}

function Get-MailFolderByPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$Path
    )

    $currentFolder = $null
    $resolvedSegments = [System.Collections.Generic.List[string]]::new()

    for ($index = 0; $index -lt $Path.Count; $index++) {
        $segment = $Path[$index]

        if ([string]::IsNullOrWhiteSpace($segment)) {
            throw "FolderPath contains an empty segment at index $index."
        }

        $escapedSegment = Escape-ODataLiteral -Value $segment
        $folderFilter = "displayName eq '$escapedSegment'"

        if ($index -eq 0) {
            $folderCandidates = @(
                Get-MgUserMailFolder -UserId $UserId -Filter $folderFilter `
                    -Property 'id', 'displayName', 'parentFolderId', 'childFolderCount' -All
            )
        }
        else {
            $folderCandidates = @(
                Get-MgUserMailFolderChildFolder -UserId $UserId `
                    -MailFolderId $currentFolder.Id -Filter $folderFilter `
                    -Property 'id', 'displayName', 'parentFolderId', 'childFolderCount' -All
            )
        }

        $nameMatches = @(
            $folderCandidates | Where-Object {
                $_.DisplayName.Equals($segment, [StringComparison]::OrdinalIgnoreCase)
            }
        )

        if ($nameMatches.Count -eq 0) {
            $parentPath = if ($resolvedSegments.Count -eq 0) {
                '<mailbox root>'
            }
            else {
                $resolvedSegments -join ' > '
            }
            throw "Folder '$segment' was not found directly under '$parentPath'."
        }

        if ($nameMatches.Count -gt 1) {
            $ambiguousPath = (@($resolvedSegments.ToArray()) + $segment) -join ' > '
            throw "More than one folder matched '$ambiguousPath'."
        }

        $currentFolder = $nameMatches[0]
        $resolvedSegments.Add($currentFolder.DisplayName)
    }

    $currentFolder
}

function Get-AttachmentProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Attachment,
        [Parameter(Mandatory)][string]$PropertyName,
        [AllowNull()][object]$DefaultValue
    )

    $directProperty = $Attachment.PSObject.Properties[$PropertyName]
    if ($null -ne $directProperty -and $null -ne $directProperty.Value) {
        return $directProperty.Value
    }

    $additionalProperty = $Attachment.PSObject.Properties['AdditionalProperties']
    if ($null -ne $additionalProperty -and
        $null -ne $additionalProperty.Value -and
        $additionalProperty.Value.ContainsKey($PropertyName)) {
        return $additionalProperty.Value[$PropertyName]
    }

    $DefaultValue
}

# Import-Module Microsoft.Graph.Authentication
# Import-Module Microsoft.Graph.Mail
# Ensure required modules
[string[]] $modules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Mail')
$modules | ForEach-Object {
    if (-not (Get-InstalledModule -Name $_ -ErrorAction SilentlyContinue)) {
        Install-Module -Name $_ -Scope CurrentUser -Force -Verbose
    }
}

Connect-MgGraph -Scopes 'Mail.ReadWrite' -NoWelcome

try {
    $context = Get-MgContext
    if ($null -eq $context -or [string]::IsNullOrWhiteSpace($context.Account)) {
        throw 'Microsoft Graph sign-in did not complete.'
    }

    $userId = $context.Account
    $folderPathText = $FolderPath -join ' > '
    $directionText = $SelectedDirections -join ', '

    Write-Host "Signed in as: $userId"
    Write-Host "Resolving folder: $folderPathText"

    $folder = Get-MailFolderByPath -UserId $userId -Path $FolderPath

    $cutoffUtc = $ReceivedBefore.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

    Write-Host "Resolved folder: $folderPathText"
    Write-Host "Message direction: $directionText"
    Write-Host "Messages received before: $cutoffUtc"
    Write-Host ('Minimum attachment size: {0:N1} MB' -f ($MinimumAttachmentBytes / 1MB))

    $messageProperties = @(
        'id'
        'subject'
        'receivedDateTime'
        'sentDateTime'
        'from'
        'toRecipients'
        'ccRecipients'
        'hasAttachments'
    )

    $messageFilter = "receivedDateTime lt $cutoffUtc and hasAttachments eq true"
    $messages = @(
        Get-MgUserMailFolderMessage -UserId $userId -MailFolderId $folder.Id `
            -Filter $messageFilter -Property $messageProperties -All
    )

    Write-Host "Messages returned by Graph: $($messages.Count)"

    $report = [System.Collections.Generic.List[object]]::new()
    $matchingMessageCount = 0
    $matchingAttachmentCount = 0
    $deletedAttachmentCount = 0

    foreach ($message in $messages) {
        $matchedDirections = @(
            Get-MessageDirectionMatches -Message $message -MyAddress $userId `
                -SelectedDirection $SelectedDirections
        )

        if ($matchedDirections.Count -eq 0) {
            continue
        }

        $matchingMessageCount++

        $attachments = @(
            Get-MgUserMessageAttachment -UserId $userId -MessageId $message.Id -All
        )

        foreach ($attachment in $attachments) {
            $name = [string](Get-AttachmentProperty -Attachment $attachment `
                -PropertyName 'Name' -DefaultValue '(unnamed attachment)')
            $size = [long](Get-AttachmentProperty -Attachment $attachment `
                -PropertyName 'Size' -DefaultValue 0L)
            $isInline = [bool](Get-AttachmentProperty -Attachment $attachment `
                -PropertyName 'IsInline' -DefaultValue $false)
            $odataType = [string](Get-AttachmentProperty -Attachment $attachment `
                -PropertyName '@odata.type' -DefaultValue '')

            if ($isInline -or $size -lt $MinimumAttachmentBytes) {
                continue
            }

            $matchingAttachmentCount++
            $action = 'Report only'

            if ($DeleteAttachments) {
                $removeParameters = @{
                    UserId       = $userId
                    MessageId    = $message.Id
                    AttachmentId = $attachment.Id
                    Confirm      = $false
                    PassThru     = $true
                }

                if (-not $Commit) {
                    $removeParameters.WhatIf = $true
                    $action = 'WhatIf deletion'
                    Remove-MgUserMessageAttachment @removeParameters | Out-Null
                }
                else {
                    $removed = Remove-MgUserMessageAttachment @removeParameters
                    $action = if ($removed) { 'Deleted' } else { 'Deletion submitted' }
                    $deletedAttachmentCount++
                }
            }

            $sender = if ($null -ne $message.From -and
                          $null -ne $message.From.EmailAddress) {
                $message.From.EmailAddress.Address
            }
            else {
                $null
            }

            $report.Add([pscustomobject]@{
                FolderPath      = $folderPathText
                Received        = $message.ReceivedDateTime
                Sent            = $message.SentDateTime
                Direction       = $matchedDirections -join ','
                Sender          = $sender
                Subject         = $message.Subject
                AttachmentName  = $name
                AttachmentBytes = $size
                AttachmentMB    = [math]::Round($size / 1MB, 2)
                AttachmentType  = $odataType
                MessageId       = $message.Id
                AttachmentId    = $attachment.Id
                Action          = $action
            })

            Write-Host ('{0}: [{1}] {2} | {3} | {4:N1} MB' -f `
                $action, ($matchedDirections -join ','), $message.Subject, `
                $name, ($size / 1MB))
        }
    }

    $reportDirectory = Split-Path -Parent -Path $ReportPath
    if (-not [string]::IsNullOrWhiteSpace($reportDirectory) -and
        -not (Test-Path -LiteralPath $reportDirectory)) {
        New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
    }

    if ($report.Count -gt 0) {
        $report |
            Sort-Object Received, Subject, AttachmentName |
            Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding utf8BOM
    }
    else {
        $columns = @(
            'FolderPath', 'Received', 'Sent', 'Direction', 'Sender', 'Subject',
            'AttachmentName', 'AttachmentBytes', 'AttachmentMB', 'AttachmentType',
            'MessageId', 'AttachmentId', 'Action'
        )
        ('"' + ($columns -join '","') + '"') |
            Set-Content -LiteralPath $ReportPath -Encoding utf8BOM
    }

    Write-Host
    Write-Host "Report written to: $ReportPath"
    Write-Host "Direction-matching messages: $matchingMessageCount"
    Write-Host "Matching attachments: $matchingAttachmentCount"

    if (-not $DeleteAttachments) {
        Write-Host 'No attachments were changed.'
    }
    elseif (-not $Commit) {
        Write-Host 'WhatIf mode was used; no attachments were changed.'
    }
    else {
        Write-Host "Attachment deletions submitted: $deletedAttachmentCount"
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
