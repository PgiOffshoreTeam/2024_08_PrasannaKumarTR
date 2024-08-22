$exclude = @("venv", "InvoiceExtractorAPI.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "InvoiceExtractorAPI.zip" -Force