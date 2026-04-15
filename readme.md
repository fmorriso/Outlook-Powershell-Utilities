# Outlook PowerShell Utilities

Powershell 7 scripts that replace lost functionality when Microsoft
shifted from full-functionality Outlook Desktop to the Outlook "app" and Outlook.com, including:

1. show all unread email in a nice display, regardless of how deeply nested (e.g., Filed > Companies > Oracle) the unread email may sit within (perhaps due to Rules routing). Includes logic to compensate for known bug in Outlook
   that prevents it from seeing nested folder names longer than 32 characters.

1. remove old emails from the Outlook Junk folder based on a configurable number of days old.

1. remove old emails from the Outlook InBox folder based on a configurable
number of days old.

## Tools Used

| Tool       | Version |
| :--------- | ------: |
| Powershell |   7.6.0 |
| VSCode     | 1.115.0 |

## Change History

| Date       | Description                       |
| :--------- | :-------------------------------- |
| 2026-04-15 | Add remove old InBox items script |
| 2026-04-14 | Initial creation                  |

## References

-
