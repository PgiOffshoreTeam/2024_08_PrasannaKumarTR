$exclude = @("venv", "APIIntergrations.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "APIIntergrations.zip" -Force