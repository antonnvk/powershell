$m = Get-mailbox | Get-MailboxStatistics | Sort-Object TotalItemSize -descending | ft Database, DisplayName, TotalItemSize



$a = "<style>"
$a = $a + "<BODY>"
$a = $a + "<H2>Информация по почтовым ящикам НХС</H2>"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$a = $a + "</style>"

#Get-mailbox | Get-MailboxStatistics | Sort-Object TotalItemSize -descending | ConvertTo-HTML Database, DisplayName, TotalItemSize, DatabaseIssueWarningQuota, DatabaseProhibitSendQuota, DatabaseProhibitSendReceiveQuota -head $a | Out-File C:\MailboxStatistics.htm

$l = ConvertTo-HTML $m -head $a 
Out-File $l C:\MailboxStatistics.htm

Invoke-Expression C:\MailboxStatistics.htm