using module "..\ExcelFileInfoManagement.psm1"

Param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [ValidateScript({ (Get-Item $_).PSIsContainer })]
    [String] $SourceDir,

    [Parameter(Position = 1, Mandatory = $true)]
    [String[]] $FilePaths=@(),

    [Parameter(Position = 2)]
    [switch] $IncludesSubdir,

    [Parameter(Position = 3)]
    [String] $DbFilePath,

    [Parameter(Position = 4)]
    [String] $DbFileEncoding = "utf8",

    [Parameter(Position = 5)]
    [bool] $AbsolutePath = $False
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version 3.0

$params = @{
    SourceDir = $SourceDir
    FilePaths = $FilePaths
    DbFilePath = $DbFilePath
    DbFileEncoding = $DbFileEncoding
}

if ($IncludesSubdir) { $params.IncludesSubdir = $True }
if ($AbsolutePath) { $params.AbsolutePath = $True }

Update-ExcelFileInfoDb @params