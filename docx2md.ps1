#Check if called with Source and Target already specified
param (
    [string]$Source,
    [string]$Target
)

# Set the output encoding to UTF-8 to support any language
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::UTF8

Write-Host "This script uses Pandoc to convert a .docx file to markdown, removing empty lines."
Write-Host "It's intended for importing Word notes into markdown-based programs like Obsidian or Notion."
Write-Host "Ensure that the command line program Pandoc is installed." "`n"

# Get source and target paths if not provided as arguments
if (-not $Source) {
    $Source = (Read-Host -Prompt "Filepath of initial .docx file").Trim("'",'"')
} else {
    $Source = $Source.Trim("'",'"')
}

if (-not $Target) {
    $Target = (Read-Host -Prompt "Desired filepath of your .md file").Trim("'",'"')
} else {
    $Target = $Target.Trim("'",'"')
}

# Check if source file exists
if (!(Test-Path -Path $Source -PathType Leaf)) {
    Write-Host "Error: The specified source file does not exist." -ForegroundColor Red
    exit
}

# Check if the source file has a .docx extension
if ([System.IO.Path]::GetExtension($Source).ToLower() -ne ".docx") {
    Write-Host "Error: The source file is not a .docx file." -ForegroundColor Red
    exit
}

# Check if Pandoc is available
if (!(Get-Command pandoc -ErrorAction SilentlyContinue)) {
    Write-Host "Error: Pandoc is not installed or not found in the system path." -ForegroundColor Red
    exit
}

# Check if the target markdown file already exists
if (Test-Path -Path $Target) {
    $overwrite = Read-Host "The file $Target already exists. Do you want to overwrite it? (Y/N)"
    if ($overwrite -ne 'Y' -and $overwrite -ne 'y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit
    }
}

# Run Pandoc with empty line removal and UTF-8 encoding for output, capturing any errors
Try {
    Write-Host "`nConverting .docx to markdown and removing empty lines..."

    # Run Pandoc and capture any errors
    $pandocOutput = pandoc -f docx -t markdown $Source 2>&1
    if ($pandocOutput -match "couldn't unpack docx container") {
        Write-Host "Error: The source file is not a valid .docx file or is corrupted." -ForegroundColor Red
        exit
    }

    # Continue with filtering empty lines and writing to the target file
    $pandocOutput | Where-Object { $_ -ne "" } | Out-File -FilePath $Target -Encoding UTF8
    Write-Host "`nConversion complete! Markdown file saved to $Target."
}
Catch {
    Write-Host "Conversion failed. Please check your permissions and file paths." -ForegroundColor Red
}
