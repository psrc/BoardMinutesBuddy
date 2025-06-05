# PowerShell script to find Microsoft Word files with "minutes" or "agenda" in filename
# Searches recursively through specified directories
# Author: Generated for Word document search task

# Define the directories to search
$SearchDirectories = @(
    "\\FILE\dept\Exec\PSRCcommittees\Operations\Agendas",
    "\\FILE\dept\Exec\PSRCcommittees\Executive Board\Agendas",
    "\\FILE\dept\Grow\Committees\GMPB",
    "\\FILE\dept\Trans\Comm\TPB",
    "\\FILE\dept\EDD\EDD Board of Directors\EDD Board Meetings\Meetings"
)

# Define the regex pattern for matching filenames
$RegexPattern = "([Mm]inutes|[Aa]genda)"

# Initialize array to store results
$FoundFiles = @()

Write-Host "Starting search for Word documents containing 'minutes' or 'agenda'..." -ForegroundColor Green
Write-Host "Search directories:" -ForegroundColor Yellow
$SearchDirectories | ForEach-Object { Write-Host "  - $_" -ForegroundColor Cyan }
Write-Host ""

# Loop through each specified directory
foreach ($Directory in $SearchDirectories) {
    Write-Host "Searching in: $Directory" -ForegroundColor Yellow
    
    # Check if directory exists
    if (Test-Path $Directory) {
        try {
            # Get all .doc and .docx files recursively
            $WordFiles = Get-ChildItem -Path $Directory -Recurse -Include "*.doc", "*.docx" -File -ErrorAction SilentlyContinue
            
            # Filter files that match the regex pattern
            $MatchingFiles = $WordFiles | Where-Object { $_.Name -match $RegexPattern }
            
            # Add matching files to results array
            foreach ($File in $MatchingFiles) {
                $FoundFiles += [PSCustomObject]@{
                    FullPath = $File.FullName
                    FileName = $File.Name
                    Directory = $File.DirectoryName
                    Size = [math]::Round($File.Length / 1KB, 2)  # Size in KB
                    LastModified = $File.LastWriteTime
                }
                Write-Host "  Found: $($File.FullName)" -ForegroundColor Green
            }
            
            Write-Host "  Found $($MatchingFiles.Count) matching files in this directory" -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Error accessing directory '$Directory': $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning "Directory not found: $Directory"
    }
    Write-Host ""
}

# Display summary results
Write-Host "=== SEARCH COMPLETE ===" -ForegroundColor Magenta
Write-Host "Total files found: $($FoundFiles.Count)" -ForegroundColor Green
Write-Host ""

if ($FoundFiles.Count -gt 0) {
    Write-Host "Full file paths:" -ForegroundColor Yellow
    $FoundFiles | ForEach-Object { 
        Write-Host "$($_.FullPath)" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "Detailed results:" -ForegroundColor Yellow
    $FoundFiles | Format-Table -Property FileName, Directory, Size, LastModified -AutoSize
    
    # Export results to CSV file
    $OutputFile = "agenda_minutes_paths.csv"
    $FoundFiles | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported to: $OutputFile" -ForegroundColor Green
}
else {
    Write-Host "No matching files found." -ForegroundColor Red
}

# Optional: Open the results directory
$Continue = Read-Host "Press Enter to continue or type 'open' to open current directory"
if ($Continue -eq "open") {
    Start-Process explorer.exe -ArgumentList (Get-Location).Path
}
