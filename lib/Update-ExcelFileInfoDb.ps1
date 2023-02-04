<#
.Synopsis
Read info on Excel files from a directory and update records in a database file.

.Description
Read info on Excel files from a directory and update records in a database file.

The required module is below.
[dfinke/ImportExcel: PowerShell module to import/export Excel spreadsheets, without Excel](https://github.com/dfinke/ImportExcel)

If a Excel file is open when an error will occur. -> Failed retrieving Excel workbook information

.Parameter SourceDir
The directory including Excel files.

.Parameter FilePaths
An array of FilePath string

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

.Parameter AbsolutePath
Whether to write the Excel file path as absolute.

.Example
PS> Update-ExcelFileInfoDb -SourceDir "C:\Excel\notes" -DbFilePath "C:\Excel_note_metadata.json"
#>
using namespace System.Collections.Generic # PowerShell 5
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

. (Join-Path -Path $PSScriptRoot -ChildPath "./Common.ps1")

function Update-ExcelFileInfoDb {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [ValidateScript({ (Get-Item $_).PSIsContainer })]
        [String] $SourceDir,

        [Parameter(Position = 1, Mandatory = $true)]
        [String[]] $FilePaths,

        [Parameter(Position = 2)]
        [switch] $IncludesSubdir,

        [Parameter(Position = 3)]
        [String] $FilterString = "*.xls*",

        [Parameter(Position = 4)]
        [String] $DbFilePath,

        [Parameter(Position = 5)]
        [String] $DbFileEncoding = "utf8",

        [Parameter(Position = 7)]
        [bool] $AbsolutePath = $False
    )
    Process {
        # Checking the arguments
        . {
            Write-Host "`$SourceDir: $($SourceDir)"
            Write-Host "`$FilePaths: $($FilePaths)"
            Write-Host "`$FilterString: $($FilterString)"
            Write-Host "`$IncludesSubdir: $($IncludesSubdir)"
            Write-Host "`$DbFilePath: $($DbFilePath)"
            Write-Host "`$DbFileEncoding: $($DbFileEncoding)"
            Write-Host "`$AbsolutePath: $($AbsolutePath)"
        }

        # Setting the path of database file writing metadata
        if([String]::IsNullOrEmpty($DbFilePath)) {
            $DbFilePath = Join-Path -Path $SourceDir -ChildPath ".metadata.json"
        }

        Write-Host "[info] The path of database file writing metadata is `"$($DbFilePath)`""

        # Reading the existing database
        if (-not(Test-Path -LiteralPath $DbFilePath)) {
            Write-Error "[error] The database file is not existing: `"$($DbFilePath)`""
        }

        # Setting a list of Excel files in the directory
        [System.IO.FileInfo[]] $fileInfo = ReadFileInfoInSourceDir -SourceDir $SourceDir $IncludesSubdir

        # Convert System.IO.FileInfo[] to System.Collections.Generic.List
        [List[PSCustomObject]] $booksList = ConvertFileInfoToGenericList -SourceDir $SourceDir -FileInfo $fileInfo

        # Reading the database
        [List[PSCustomObject]] $db = ReadDatabase -DbPath $DbFilePath

        # Loop the specified file paths for updating
        foreach ($fp in $FilePaths) {
            [Int] $idx = $booksList.FindIndex({ param($item) $item.FilePath -eq $fp })

            if ($idx -eq -1) {
                Write-Warning "[warn] Skipped the non-existing file: `"$($fp)`""
            }

            # Get a info of the Excel file
            [PSCustomObject] $xlInfoObj = ReadExcelInfoAsPSCustomObject -BookPath $booksList[$idx].FullName

            # Finding the FilePath from the database
            [Int] $dbIdx = $db.FindIndex({ param($item) $item.FilePath -eq $fp })
            Write-Host "[info] matched DB index: $($dbIdx)"

            # Creating a new record
            if ($dbIdx -eq -1) {
                Write-Host "[info] Creating a new record FilePath: `"$($fp)`""

                $db.Add([PSCustomObject]@{
                    Id = New-Guid
                    FilePath = $fp
                    ExcelInfo = $xlInfoObj
                })
            }
            # Updating the existing record
            elseif ($null -ne $xlInfo) {
                Write-Host "[info] Updating the existing record Id: $($db[$dbIdx].Id)"
                $db[$dbIdx].ExcelInfo = $xlInfoObj
            }
        }

        # Writing the metadata database file
        Write-Host "[info] Writing the database: `"$($DbFilePath)`""
        try {
            ConvertTo-Json -InputObject $db | Out-File -LiteralPath $DbFilePath -Encoding $DbFileEncoding
        }
        catch {
            Write-Host "[error] $($_.Exception.Message)"
        }
    }
}

Export-ModuleMember -Function Update-ExcelFileInfoDb