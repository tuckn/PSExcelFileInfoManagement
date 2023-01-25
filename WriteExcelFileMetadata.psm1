<#
.Synopsis
Read metadata from Excel files and save it in a JSON file.

.Description
Read metadata from Excel files and save it in a JSON file.
The required module is below.
[dfinke/ImportExcel: PowerShell module to import/export Excel spreadsheets, without Excel](https://github.com/dfinke/ImportExcel)

If a Excel file is open when an error will occur. -> Failed retrieving Excel workbook information

.Parameter SourcePath
The path of an Excel file or directory.

.Parameter FilteredName
The file name to read. The default is "*.xls*"

.Parameter IncludesSubdir
Whether to include subfolders or not. The default is `$False`.

.Parameter OutFilePath
The JSON file path to writing metadata. The default is `$SourcePath\.metadata.json`.

.Parameter OutFileEncoding
The default is "utf8" (UTF8 with BOM).
If your PowerShell is less than 6, when you can only use "unknown, string, unicode, bigendianunicode, utf8, utf7, utf32, ascii, default, oem".

[Out-File (Microsoft.PowerShell.Utility) - PowerShell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/out-file?view=powershell-7.3#-encoding)

.Example
PS> Write-ExcelFileMetadata -SourcePath "C:\Excel\notes" -OutFilePath "C:\Excel_note_metadata.json"
#>
using namespace System.Collections.Generic # PowerShell 5
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

function Write-ExcelFileMetadata {
    [CmdletBinding()]
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
    Process {
        # Checking the arguments
        . {
            Write-Host "`$SourcePath: $($SourcePath)"
            Write-Host "`$FilteredName: $($FilteredName)"
            Write-Host "`$IncludesSubdir: $($IncludesSubdir)"
            Write-Host "`$OutFilePath: $($OutFilePath)"
            Write-Host "`$OutFileEncoding: $($OutFileEncoding)"
        }

        $fileList = [List[PSObject]]::new()
        Function Convert-OOXmlPropsToPSObj ([OfficeOpenXml.OfficeProperties] $x) {
            return [PSCustomObject]@{
                Application = $x.Application
                Title = $x.Title
                Subject = $x.Subject
                Author = $x.Author
                Comments = $x.Comments
                Keywords = $x.Keywords
                LastModifiedBy = $x.LastModifiedBy
                # LastPrinted = $x.LastPrinted.ToString("yyyy/MM/dd HH:mm:ss")
                Created = $x.Created.ToString("yyyy-MM-ddTHH:mm:ss.fffK") # ISO 8601
                Category = $x.Category
                Status = $x.Status
                AppVersion = $x.AppVersion
                Company = $x.Company
                Manager = $x.Manager
                Modified = $x.Modified.ToString("yyyy-MM-ddTHH:mm:ss.fffK")
            }
        }

        # SourcePath is a directory
        if ((Get-Item -LiteralPath $SourcePath).PSIsContainer) {
            # Set the path of output JSON file
            if([String]::IsNullOrEmpty($OutFilePath)) {
                $OutFilePath = Join-Path -Path $SourcePath -ChildPath ".metadata.json"
            }

            [String] $childPath = Join-Path -Path $SourcePath -ChildPath $FilteredName

            foreach ($f in Get-ChildItem $childPath) {
                try {
                    # dfinke/ImportExcel: PowerShell module to import/export Excel spreadsheets, without Excel
                    # https://github.com/dfinke/ImportExcel
                    [OfficeOpenXml.OfficeProperties] $info = Get-ExcelWorkbookInfo -Path "$($f.FullName)"
                    $fileList.Add((Convert-OOXmlPropsToPSObj $info))
                }
                catch {
                    # If the Excel file is open when an error will occur. -> Failed retrieving Excel workbook information
                    Write-Error $_
                }
            }
        }
        # SourcePath is not a directory
        elseif (Test-Path -LiteralPath $SourcePath) {
            # Set the path of output JSON file
            if([String]::IsNullOrEmpty($OutFilePath)) {
                $OutFilePath = Join-Path -Path (Get-Item $SourcePath).DirectoryName -ChildPath ".metadata.json"
            }

            try {
                [OfficeOpenXml.OfficeProperties] $info = Get-ExcelWorkbookInfo -Path "$($SourcePath)"
                $fileList.Add((Convert-OOXmlPropsToPSObj $info))
            }
            catch {
                Write-Error $_
            }
        }
        else {
            Write-Error "`$SourcePath is not existing. $($SourcePath)"
            exit 1
        }

        # Write the metadata in the JSON file
        Write-Host "[info] The path of output JSON file is `"$($OutFilePath)`""

        try {
            ConvertTo-Json -InputObject $fileList | Out-File -LiteralPath $OutFilePath -Encoding $OutFileEncoding
        }
        catch {
            Write-Host "[error] $($_.Exception.Message)"
        }
    }
}

Export-ModuleMember -Function Write-ExcelFileMetadata