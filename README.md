
# Extract and Merge PowerPoint Presentations

## Description
A PowerShell script to extract slides from multiple PowerPoint files and merge them into a single presentation. This tool automates the consolidation of presentations, saving time and ensuring consistency.

## Prerequisites
- Windows Operating System
- PowerPoint installed on the system
- PowerShell version 5.0 or higher

## Usage
1. Place all the source `.ppt` or `.pptx` files in the `source` folder on your Desktop.
2. Run the `extract&merge_ppts.ps1` script:
   ```powershell
   ./extract&merge_ppts.ps1
   ```
3. The merged presentation will be saved in the `merged` folder on your Desktop as `MergedPresentation.pptx` and `MergedPresentation.pdf`.

## Script Details
- **Source Directory:** `C:\Users\Username\Desktop\source`
- **Destination Directory:** `C:\Users\Username\Desktop\merged`
- **Functionality:**
  - Copies all PowerPoint files from the source to the destination.
  - Opens each presentation and copies all slides to a new merged presentation.
  - Saves the merged presentation in both `.pptx` and `.pdf` formats.

## Example
After placing your PowerPoint files in the `source` folder, executing the script will result in:
- `C:\Users\Username\Desktop\merged\MergedPresentation.pptx`
- `C:\Users\Username\Desktop\merged\MergedPresentation.pdf`

## License
This project is licensed under the MIT License.
