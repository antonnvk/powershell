Get-mailbox | Get-MailboxStatistics | Sort-Object TotalItemSize -descending | ConvertTo-HTML Database, DisplayName, TotalItemSize -body "<H2>Test</H2>" | Out-File C:\MailboxStatistics.htm

Invoke-Expression C:\MailboxStatistics.htm