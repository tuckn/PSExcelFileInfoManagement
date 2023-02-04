using module "..\ExcelFileInfoManagement.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [String] $SourceDir,

    [Parameter(Position = 2)]
    [Boolean] $IncludesSubdir = $False,

    [Parameter(Position = 3)]
    [String] $FilterString = "*.xls*",

    [Parameter(Position = 4)]
    [String] $DbFilePath,

    [Parameter(Position = 5)]
    [bool] $AbsolutePath = $False,

    [Parameter(Position = 6)]
    [String] $DbFileEncoding = "utf8"
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

Update-ExcelFileInfoDb `
    -SourceDir $SourceDir `
    -IncludesSubdir $IncludesSubdir `
    -FilterString $FilterString `
    -OutFilePath $DbFilePath `
    -OutFileEncoding $DbFileEncoding