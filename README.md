# Report-MailboxStatistics.ps1

The script sends a report to Exchange users with the overview of the largest folders and the folders with the most items.

## Usage

Download and copy this script to an Exchange Server.
Run this script interactive or with task scheduler.
Change the first lines of the script :

```
$Group = "Domain Users"
$CountTopFolder = 10

$SMTPServer = "smtp.domain.tld"
$From = "postfachbericht@domain.tld"
$Subject = "Postfach Übersicht"
```

## Tested Exchange / Windows Server Versions

- Exchange Server 2019
- Windows Server 2022

## Visit my Blog for an example

 [FrankysWeb](https://www.frankysweb.de/exchange-server-bericht-ueber-postfachgroesse-an-benutzer-senden/)
