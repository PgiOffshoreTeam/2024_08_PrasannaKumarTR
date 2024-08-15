$exclude = @("venv", "InvoiceDataExtractor.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "InvoiceDataExtractor.zip" -Force