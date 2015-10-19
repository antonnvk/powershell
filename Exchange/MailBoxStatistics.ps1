Get-mailbox | Get-MailboxStatistics | Sort-Object TotalItemSize -descending |ft Database, DisplayName, TotalItemSize | Out-File C:\Test.htm

Invoke-Expression C:\Test.htm