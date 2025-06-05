# PowerShell Script to Convert .doc files to .docx from file_trios paths
# Purpose: Batch convert agenda and minutes .doc files to .docx format for officer package compatibility
# Usage: Run this script in PowerShell, then update file_trios paths with find/replace .doc -> .docx

param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = "file_trios_paths.csv"  # CSV export of file paths from R
)

# Function to safely quit Word application
function Close-WordApplication {
    param($WordApp)
    try {
        if ($WordApp -ne $null) {
            $WordApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordApp) | Out-Null
        }
    }
    catch {
        Write-Warning "Error closing Word application: $($_.Exception.Message)"
    }
}

# Function to convert a single .doc file to .docx
function Convert-DocToDocx {
    param(
        [string]$SourcePath,
        [object]$WordApp
    )

    if (-not (Test-Path $SourcePath)) {
        Write-Warning "File not found: $SourcePath"
        return $false
    }

    # Only process .doc files (not .docx)
    if ([System.IO.Path]::GetExtension($SourcePath).ToLower() -ne ".doc") {
        Write-Host "Skipping non-.doc file: $SourcePath" -ForegroundColor Yellow
        return $false
    }

    # Generate output path (same directory, .docx extension)
    $OutputPath = [System.IO.Path]::ChangeExtension($SourcePath, ".docx")

    # Skip if .docx already exists
    if (Test-Path $OutputPath) {
        Write-Host "Already exists, skipping: $OutputPath" -ForegroundColor Green
        return $true
    }

    try {
        Write-Host "Converting: $([System.IO.Path]::GetFileName($SourcePath))" -ForegroundColor Cyan

        # Open the .doc file
        $Document = $WordApp.Documents.Open($SourcePath, $false, $true)  # ReadOnly = true

        # Save as .docx (format 16 = wdFormatXMLDocument)
        $Document.SaveAs2($OutputPath, 16)

        # Close the document
        $Document.Close($false)  # Don't save changes to original

        Write-Host "  → Created: $([System.IO.Path]::GetFileName($OutputPath))" -ForegroundColor Green
        return $true

    }
    catch {
        Write-Error "Failed to convert $SourcePath`: $($_.Exception.Message)"
        return $false
    }
}

# Main execution
Write-Host "=== .doc to .docx Batch Converter ===" -ForegroundColor Magenta
Write-Host "Starting conversion process..." -ForegroundColor White

# Initialize counters
$TotalFiles = 0
$ConvertedFiles = 0
$SkippedFiles = 0
$ErrorFiles = 0

try {
    # Start Word application
    Write-Host "Starting Microsoft Word..." -ForegroundColor Yellow
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0  # Suppress alerts

    # Option 1: If you export file paths from R as CSV
    if (Test-Path $CsvPath) {
        Write-Host "Reading file paths from: $CsvPath" -ForegroundColor Cyan
        $FilePaths = Import-Csv $CsvPath

        # Process agenda files
        if ($FilePaths[0].PSObject.Properties.Name -contains "agenda") {
            Write-Host "`nProcessing AGENDA files..." -ForegroundColor Magenta
            foreach ($Row in $FilePaths) {
                if (![string]::IsNullOrWhiteSpace($Row.agenda)) {
                    $TotalFiles++
                    $Result = Convert-DocToDocx -SourcePath $Row.agenda -WordApp $Word
                    if ($Result -eq $true) { $ConvertedFiles++
                    } elseif ($Result -eq $false -and (Test-Path ([System.IO.Path]::ChangeExtension($Row.agenda, ".docx")))) {
                        $SkippedFiles++
                    } else {
                        $ErrorFiles++
                    }
                }
            }
        }

        # Process minutes files
        if ($FilePaths[0].PSObject.Properties.Name -contains "minutes") {
            Write-Host "`nProcessing MINUTES files..." -ForegroundColor Magenta
            foreach ($Row in $FilePaths) {
                if (![string]::IsNullOrWhiteSpace($Row.minutes)) {
                    $TotalFiles++
                    $Result = Convert-DocToDocx -SourcePath $Row.minutes -WordApp $Word
                    if ($Result -eq $true) { $ConvertedFiles++
                    } elseif ($Result -eq $false -and (Test-Path ([System.IO.Path]::ChangeExtension($Row.agenda, ".docx")))) {
                        $SkippedFiles++
                    } else {
                        $ErrorFiles++
                    }
                }
            }
        }
    }
    else {
        Write-Host "CSV file not found: $CsvPath" -ForegroundColor Red
        Write-Host "Please export your file paths from R using:" -ForegroundColor Yellow
        Write-Host "  fwrite(file_trios[, .(agenda, minutes)], 'file_trios_paths.csv')" -ForegroundColor White
        exit 1
    }
}
catch {
    Write-Error "Critical error: $($_.Exception.Message)"
    $ErrorFiles++
}
finally {
    # Always close Word application
    Write-Host "`nClosing Microsoft Word..." -ForegroundColor Yellow
    Close-WordApplication -WordApp $Word

    # Cleanup COM objects
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Summary
Write-Host "`n=== Conversion Summary ===" -ForegroundColor Magenta
Write-Host "Total files processed: $TotalFiles" -ForegroundColor White
Write-Host "Successfully converted: $ConvertedFiles" -ForegroundColor Green
Write-Host "Skipped (already .docx): $SkippedFiles" -ForegroundColor Yellow
Write-Host "Errors: $ErrorFiles" -ForegroundColor Red

if ($ConvertedFiles -gt 0) {
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "1. In R, update file_trios paths:" -ForegroundColor White
    Write-Host "   file_trios[, agenda := gsub('\\.doc$', '.docx', agenda)]" -ForegroundColor Gray
    Write-Host "   file_trios[, minutes := gsub('\\.doc$', '.docx', minutes)]" -ForegroundColor Gray
    Write-Host "2. Run your fine-tuning script with updated paths" -ForegroundColor White
}

Write-Host "`nConversion complete!" -ForegroundColor Green
