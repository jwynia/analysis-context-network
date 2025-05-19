#!/usr/bin/env pwsh
<#
.SYNOPSIS
    md2docx - Convert Markdown to DOCX using pandoc
    
.DESCRIPTION
    A wrapper around pandoc that mimics the md2docx interface.
    This script is designed to be a drop-in replacement for the md2docx CLI tool,
    using pandoc as the underlying conversion engine.
    
.PARAMETER i
    Input markdown file
    
.PARAMETER r
    Optional reference DOCX file for styling
    
.PARAMETER o
    Optional output file name. If not specified, uses the same name as the input file with .docx extension.
    
.PARAMETER toc
    Include table of contents (optional)
    
.PARAMETER tocDepth
    Table of contents depth (optional, default is 3)
    
.PARAMETER extraArgs
    Additional arguments to pass to pandoc (optional)
    
.EXAMPLE
    md2docx -i document.md
    
.EXAMPLE
    md2docx -i document.md -r template.docx
    
.EXAMPLE
    md2docx -i document.md -o custom_output.docx
    
.EXAMPLE
    md2docx -i document.md -toc -tocDepth 2
    
.EXAMPLE
    md2docx -i document.md -extraArgs "--number-sections"
#>

param (
    [Parameter(Mandatory=$true)]
    [Alias("i")]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false)]
    [Alias("r")]
    [string]$ReferenceDoc,
    
    [Parameter(Mandatory=$false)]
    [Alias("o")]
    [string]$OutputFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$toc,
    
    [Parameter(Mandatory=$false)]
    [int]$tocDepth = 3,
    
    [Parameter(Mandatory=$false)]
    [string]$extraArgs
)

# Verify input file exists
if (-not (Test-Path $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    exit 1
}

# Generate output filename if not specified
if (-not $OutputFile) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, "docx")
}

# Build the pandoc command
$pandocArgs = @(
    "-f", "markdown",
    "-t", "docx",
    "-o", "`"$OutputFile`""
)

# Add reference doc if specified
if ($ReferenceDoc) {
    if (Test-Path $ReferenceDoc) {
        $pandocArgs += "--reference-doc=`"$ReferenceDoc`""
    } else {
        Write-Warning "Reference document not found: $ReferenceDoc"
    }
}

# Add table of contents if specified
if ($toc) {
    $pandocArgs += "--toc"
    $pandocArgs += "--toc-depth=$tocDepth"
}

# Add any extra arguments
if ($extraArgs) {
    $pandocArgs += $extraArgs
}

# Add input file as the last argument
$pandocArgs += "`"$InputFile`""

# Build the final command
$pandocCommand = "pandoc " + ($pandocArgs -join " ")

# Execute pandoc
try {
    Write-Verbose "Executing: $pandocCommand"
    Invoke-Expression $pandocCommand
    
    # Check if conversion was successful
    if (Test-Path $OutputFile) {
        Write-Host "Successfully converted $InputFile to $OutputFile"
        exit 0
    } else {
        Write-Error "Failed to convert $InputFile to $OutputFile"
        exit 1
    }
} catch {
    Write-Error "Error executing pandoc: $_"
    exit 1
}
