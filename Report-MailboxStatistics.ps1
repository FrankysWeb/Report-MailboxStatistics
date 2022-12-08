$Group = "Domain Users"
$CountTopFolder = 10

$SMTPServer = "smtp.domain.tld"
$From = "postfachbericht@domain.tld"
$Subject = "Postfach Übersicht"

[System.Collections.ArrayList]$MailboxStatistics = @()
$GroupMembers = Get-ADGroup $Group | Get-ADGroupMember -Recursive | Get-ADUser -Properties msExchMailboxGuid | where {$_.msExchMailboxGuid -ne $Null}
foreach ($GroupMember in $GroupMembers) {
 $Mailbox = get-mailbox $GroupMember.SamAccountName
 $EMail = $Mailbox.PrimarySmtpAddress.Address
 $Stats = $Mailbox | Get-MailboxStatistics | select displayname, @{label="Size"; expression={$_.TotalItemSize.Value.ToMB()}}
 $Displayname = $Stats.Displayname
 $MailboxSize = $Stats.Size
 $MailboxFolderStatistics = Get-MailboxFolderStatistics $mailbox | select FolderPath,FolderSize,ItemsInFolder
 $TopFoldersBySize = $MailboxFolderStatistics | Select-Object FolderPath,@{Name="Foldersize";Expression={$r=$_.FolderSize; [long]$a = ($r.Substring($r.IndexOf("(")+1,($r.Length - 2 - $r.IndexOf("("))) -replace " bytes","" -replace ",","") ; [math]::Round($a/1048576,2) } } | sort foldersize -Descending | select -first $CountTopFolder
 $TopFoldersByItems = $MailboxFolderStatistics | sort ItemsInFolder -Descending | select -first $CountTopFolder
 
 $Statistic = [PSCustomObject]@{
	 DisplayName = $Displayname
	 EMail = $EMail
	 MailboxSize = $MailboxSize
	 TopFoldersBySize = $TopFoldersBySize
	 TopFoldersByItems = $TopFoldersByItems
	}
 $MailboxStatistics.Add($Statistic) | out-null
}

foreach ($MailboxStatistic in $MailboxStatistics) {
 $MailBody = '<!DOCTYPE html>
 <html lang="de">
  <head>
   <title>Mailbox Report</title>
   <style>
    body {font-family: Calibri;}
    td {width:100px; max-width:300px; background-color:white;}
    table {width:100%;}
    th {text-align:left; font-size:12pt; background-color:lightgrey;}
   </style>
  </head>
 <body>
  <h2>Mailbox Übersicht</h2>'
 
 $MailboxSize = $MailboxStatistic.MailboxSize
 $MailBody += '<div><p>Ihr Postfach ist '
 $MailBody += $MailboxSize
 $MailBody += ' MB groß, bitte löschen Sie nicht mehr benötigte Daten aus Ihrem Postfach.</p></div>'
 
 $TopFoldersBySize = $MailboxStatistic.TopFoldersBySize | select @{label="Ordnerpfad"; expression={$_.Folderpath}}, @{label="Größe"; expression={$str = $_.Foldersize; [string]$str + " MB"}} | ConvertTo-Html -Fragment
 $MailBody += '<div><p>Dies ist eine Übersicht ihrer '
 $MailBody += $CountTopFolder
 $MailBody += ' größten Ordner in ihrem Postfach:</p></div>'
 $MailBody += $TopFoldersBySize
 
 $TopFoldersByItems = $MailboxStatistic.TopFoldersByItems | select @{label="Ordnerpfad"; expression={$_.Folderpath}}, @{label="Anzahl Elemente"; expression={$_.ItemsInFolder}} | ConvertTo-Html -Fragment
 $MailBody += '<div><p>Ordner mit vielen Elementen beeinträchtigen die Outlook Geschwindigkeit, löschen Sie nicht mehr benötigte Elemente um Outlook nicht zu verlangsamen. Dies sind Ihre '
 $MailBody += $CountTopFolder
 $MailBody +=' Ordner mit den meisten Elementen:</p></div>'
 $MailBody += $TopFoldersByItems
 
 $MailBody += '<div><p>Hier finden Sie weitere Informationen: https://www.frankysweb.de</p></div>'
 $MailBody += '<div><p>Vielen Dank für Ihre Mithilfe</p></div>'
 
 $MailBody += '</body>
  </html>'
 
 $To =  $MailboxStatistic.EMail
 Send-MailMessage -SmtpServer $SMTPServer -From $From -To $To -Body $MailBody -BodyAsHtml -Encoding UTF8 -Subject $Subject
}