$exclude = @("venv", "FileHandling.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "FileHandling.zip" -Force