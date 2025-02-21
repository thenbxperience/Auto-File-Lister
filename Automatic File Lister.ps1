# Define the path to the directory and the output Word document
$directoryPath = "YOUR FILE PATH HERE"
$outputWordDocument = "YOUR WORD DOC FILE PATH HERE"

# Load Word interop assembly
Add-Type -AssemblyName "Microsoft.Office.Interop.Word"

# Create a new instance of Word application
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

# Add a new document
$document = $wordApp.Documents.Add()

# Add a title to the document
$selection = $wordApp.Selection
$selection.TypeText("File List")
$selection.TypeParagraph()

# Get the list of files in the directory and all subdirectories
$files = Get-ChildItem -Path $directoryPath -File -Recurse

# Add file names to the document (without paths)
foreach ($file in $files) {
    $selection.TypeText($file.Name)
    $selection.TypeParagraph()
}

# Save the document
$document.SaveAs([ref] $outputWordDocument)
$document.Close()
$wordApp.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($selection) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "File names have been written to $outputWordDocument"