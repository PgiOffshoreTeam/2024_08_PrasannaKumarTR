$exclude = @("venv", "GmailAccess.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "GmailAccess.zip" -Force