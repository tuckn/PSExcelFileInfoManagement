<#
.Synopsis
Read metadata from Excel files and save it in a database file.

.Description
Read Excel file metadata from a directory and save it in a database file.
The required module is below.
[dfinke/ImportExcel: PowerShell module to import/export Excel spreadsheets, without Excel](https://github.com/dfinke/ImportExcel)

If a Excel file is open when an error will occur. -> Failed retrieving Excel workbook information

.Parameter SourceDir
The directory including Excel files.

.Parameter FilterString
The file string to read. The default is "*.xls*"

.Parameter IncludesSubdir
Whether to include subfolders or not. The default is `$False`.

.Parameter DbFilePath
The database file path to writing metadata. The default is `$SourceDir\.metadata.json`.

.Parameter DbFileEncoding
The default is "utf8" (UTF8 with BOM).
If your PowerShell is less than 6, when you can only use "unknown, string, unicode, bigendianunicode, utf8, utf7, utf32, ascii, default, oem".

[Out-File (Microsoft.PowerShell.Utility) - PowerShell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/out-file?view=powershell-7.3#-encoding)

.Parameter Force
When the out file exists, an error occurs.
If you specify this switch, the existing file will be cleared with non-error. (current records will be removed).

.Parameter AbsolutePath
Whether to write the Excel file path as absolute.

.Example
PS> New-ExcelFileInfoDb -SourceDir "C:\Excel\notes" -DbFilePath "C:\Excel_note_metadata.json"
#>
using namespace System.Collections.Generic # PowerShell 5
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

. (Join-Path -Path $PSScriptRoot -ChildPath "./Common.ps1")

function New-ExcelFileInfoDb {
    [CmdletBinding()]
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
    Process {
        # Checking the arguments
        . {
            Write-Host "`$SourceDir: $($SourceDir)"
            Write-Host "`$FilterString: $($FilterString)"
            Write-Host "`$IncludesSubdir: $($IncludesSubdir)"
            Write-Host "`$DbFilePath: $($DbFilePath)"
            Write-Host "`$DbFileEncoding: $($DbFileEncoding)"
            Write-Host "`$Force: $($Force)"
            Write-Host "`$AbsolutePath: $($AbsolutePath)"
        }

        # Setting the path of database file writing metadata
        if([String]::IsNullOrEmpty($DbFilePath)) {
            $DbFilePath = Join-Path -Path $SourceDir -ChildPath ".metadata.json"
        }

        Write-Host "[info] The path of database file writing metadata is `"$($DbFilePath)`""

        if (Test-Path -LiteralPath $DbFilePath) {
            if ($Force) {
                Write-Warning "[warn] `The existing database file will be cleared: `"$($DbFilePath)`""
            }
            else {
                Write-Error "[error] `The metadfata file is existing: `"$($DbFilePath)`". If you want to overwrite this, remove it or use -Force option."
            }
        }

        # Setting a list of Excel files in the directory
        [System.IO.FileInfo[]] $fileInfo = ReadFileInfoInSourceDir -SourceDir $SourceDir $IncludesSubdir

        # Convert System.IO.FileInfo[] to System.Collections.Generic.List
        [List[PSCustomObject]] $booksList = ConvertFileInfoToGenericList -SourceDir $SourceDir -FileInfo $fileInfo

        # Initializing the database
        $db = [List[PSCustomObject]]::new()

        foreach ($b in $booksList) {
            # Get a info of the Excel file
            [PSCustomObject] $xlInfoObj = ReadExcelInfoAsPSCustomObject -BookPath $b.FullName

            Write-Host "[info] Creating a new record FilePath: `"$($b.FilePath)`""

            $db.Add([PSCustomObject]@{
                Id = New-Guid
                FilePath = $b.FilePath
                ExcelInfo = $xlInfoObj
            })
        }

        # Writing the metadata database file
        Write-Host "[info] Writing the metadata to the file: `"$($DbFilePath)`""
        try {
            ConvertTo-Json -InputObject $db | Out-File -LiteralPath $DbFilePath -Encoding $DbFileEncoding
        }
        catch {
            Write-Host "[error] $($_.Exception.Message)"
        }
    }
}

Export-ModuleMember -Function New-ExcelFileInfoDb