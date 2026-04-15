# Outlook PowerShell Utilities

Powershell 7 scripts that replace lost functionality when Microsoft
shifted from full-functionality Outlook Desktop to the Outlook "app" and Outlook.com, including:

1. show all unread email in a nice display, regardless of how deeply nested (e.g., Filed > Companies > Oracle) the unread email may sit within (perhaps due to Rules routing). Includes logic to compensate for known bug in Outlook
   that prevents it from seeing nested folder names longer than 32 characters.

1. remove old emails from the Outlook Junk folder based on a configurable number of days old.

1. remove old emails from the Outlook InBox folder based on a configurable
number of days old.

## Developer Notes

My Outlook folder structure is as follows, skipping many of the standard,
built-in folders except where necessary to convey my specific preferences:

* Inbox
* Junk Email
* ... the usual built-in folders
* Filed
  * aaSpamByRule
  * Companies
    * (many subfolders but no addtional subfolder levels)
  * Government
    * (many subfolders but no addtional subfolder levels)
  * Organizations
    * (many subfolders but no addtional subfolder levels)
  * People
    * (many subfolders but no addtional subfolder levels)

Reasons for the above structure are:
1. keep the top-level folder structure as ***simple*** as possible.
1. Use *Rules* to route emails to folders underneath ***Filed*** as much as possible.
1. Use ***Filed*** as a way to show which folders contain emails that originally
arrived (briefly, in some cases) in the ***InBox*** before being routed by ***Rules*** to somewhere else.
1. Keep even the ***Filed*** folder as shallow as possible, no more than three levels deep.

## Tools Used

| Tool       | Version |
| :--------- | ------: |
| Powershell |   7.6.0 |
| VSCode     | 1.116.0 |

## Change History

| Date       | Description                       |
| :--------- | :-------------------------------- |
| 2026-04-15 | Add remove old InBox items script |
| 2026-04-14 | Initial creation                  |

## References

### Microsoft Graph Commands Used

This script relies on several Microsoft Graph PowerShell SDK commands:

- [`Connect-MgGraph`][connect] — establishes authentication with Microsoft Graph  
- [`Disconnect-MgGraph`][disconnect] — closes the Graph session  
- [`Invoke-MgGraphRequest`][invoke] — used for GET and DELETE operations  
- [`Microsoft.Graph`][mod-core] and [`Microsoft.Graph.Mail`][mod-mail] modules

[connect]: https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
[disconnect]: https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/disconnect-mggraph
[invoke]: https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/invoke-mggraphrequest
[mod-core]: https://learn.microsoft.com/powershell/microsoftgraph/overview
[mod-mail]: https://learn.microsoft.com/powershell/module/microsoft.graph.mail
