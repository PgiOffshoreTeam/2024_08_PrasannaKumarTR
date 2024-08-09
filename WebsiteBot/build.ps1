$exclude = @("venv", "WebsiteBot.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "WebsiteBot.zip" -Force