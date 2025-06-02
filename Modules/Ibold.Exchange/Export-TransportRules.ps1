$file = Export-TransportRuleCollection

[System.IO.File]::WriteAllBytes('C:\MailFlowRuleCollections\BackupRuleCollection.xml', $file.FileData)