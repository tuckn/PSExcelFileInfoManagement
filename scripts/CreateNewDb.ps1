using module "..\ExcelFileInfoManagement.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [ValidateScript({ (Get-Item $_).PSIsContainer })]
    [String] $SourceDir,

    [Parameter(Position = 1)]
    [switch] $IncludesSubdir,

    [Parameter(Position = 2)]
    [String] $FilterString = "*.xls*",

    [Parameter(Position = 3)]
    [String] $DbFilePath,

    [Parameter(Position = 4)]
    [String] $DbFileEncoding = "utf8",

    [Parameter(Position = 5)]
    [switch] $Force,

    [Parameter(Position = 6)]
    [switch] $AbsolutePath
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

$params = @{
    SourceDir = $SourceDir
    FilterString = $FilterString
    DbFilePath = $DbFilePath
    DbFileEncoding = $DbFileEncoding
}

if ($IncludesSubdir) { $params.IncludesSubdir = $True }
if ($Force) { $params.Force = $True }
if ($AbsolutePath) { $params.AbsolutePath = $True }

New-ExcelFileInfoDb @params
