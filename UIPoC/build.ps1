$exclude = @("venv", "UIPoC.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "UIPoC.zip" -Force