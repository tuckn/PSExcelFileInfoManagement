using module ".\WriteExcelFileMetadata.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [String] $SourcePath,

    [Parameter(Position = 1)]
    [String] $FilteredName = "*.xls*",

    [Parameter(Position = 2)]
    [Boolean] $IncludesSubdir = $False,

    [Parameter(Position = 3)]
    [String] $OutFilePath,

    [Parameter(Position = 4)]
    [String] $OutFileEncoding = "utf8"
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

Write-ExcelFileMetadata `
    -SourcePath "$SourcePath" `
    -FilteredName $FilteredName `
    -IncludesSubdir $IncludesSubdir `
    -OutFilePath $OutFilePath `
    -OutFileEncoding $OutFileEncoding