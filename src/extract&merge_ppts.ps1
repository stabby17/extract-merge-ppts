$source = "$env:USERPROFILE\Desktop\source"
$destination = "$env:USERPROFILE\Desktop\merged"

# Create destination folder if it doesn't exist
If (!(Test-Path -Path $destination)) {
    New-Item -ItemType Directory -Path $destination
}

# Get all PowerPoint files and copy them
Get-ChildItem -Path $source -Recurse -Include *.ppt, *.pptx | ForEach-Object {
    Copy-Item $_.FullName -Destination $destination -Force
}

# Initialize PowerPoint Application
$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Create a new presentation for merged slides
$mergedPresentation = $powerPoint.Presentations.Add()

# Get all PowerPoint files in destination folder
$presentationFiles = Get-ChildItem -Path $destination -Recurse -Include *.ppt, *.pptx

foreach ($file in $presentationFiles) {
    $presentation = $powerPoint.Presentations.Open($file.FullName, $false, $false, $false)
    foreach ($slide in $presentation.Slides) {
        $slide.Copy()
        $mergedPresentation.Slides.Paste()
    }
    $presentation.Close()
}

# Save the merged presentation
$mergedPath = Join-Path $destination "MergedPresentation.pptx"
$mergedPresentation.SaveAs($mergedPath)

# Save the merged presentation as PDF
$pdfPath = Join-Path $destination "MergedPresentation.pdf"
$mergedPresentation.SaveAs($pdfPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)

$mergedPresentation.Close()
$powerPoint.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null