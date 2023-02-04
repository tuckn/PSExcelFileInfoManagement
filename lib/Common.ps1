using namespace System.Collections.Generic # PowerShell 5
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

function ReadFileInfoInSourceDir {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [ValidateScript({ (Get-Item $_).PSIsContainer })]
        [String] $SourceDir,

        [Parameter(Position = 1)]
        [switch] $IncludesSubdir
    )
    Process {
        [String] $childPath = Join-Path -Path $SourceDir -ChildPath $FilterString

        [System.IO.FileInfo[]] $fileInfo = . {
            if ($IncludesSubdir) {
                return Get-ChildItem -Path $childPath -Recurse
            }
            else {
                return Get-ChildItem -Path $childPath
            }
        }

        return $fileInfo
    }
}

function ConvertFileInfoToGenericList {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [ValidateScript({ (Get-Item $_).PSIsContainer })]
        [String] $SourceDir,

        [Parameter(Position = 1, Mandatory = $true)]
        [System.IO.FileInfo[]] $FileInfo
    )
    Process {
        $ls = [List[PSCustomObject]]::new()

        # Save the current location
        $cwd = Get-Location
        Set-Location $SourceDir

        foreach ($f in $FileInfo) {
            $fInfo = [PSCustomObject]@{
                FullName = $f.FullName
                Name = $f.Name
                FilePath = ""
            }

            # Setting a FilePath property
            if ($AbsolutePath) {
                $fInfo.FilePath = $f.FullName
            }
            else {
                # Setting a relative path from the $SourceDir
                $fInfo.FilePath = Resolve-Path -Relative $f.FullName
            }

            $ls.Add($fInfo)
        }

        Set-Location $cwd

        return $ls
    }
}

function ReadExcelInfoAsPSCustomObject {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $BookPath
    )
    Process {
        # Get a info of the Excel file
        [OfficeOpenXml.OfficeProperties] $xlInfo = . {
            try {
                # dfinke/ImportExcel: PowerShell module to import/export Excel spreadsheets, without Excel
                # https://github.com/dfinke/ImportExcel
                return Get-ExcelWorkbookInfo -Path "$($BookPath)"
            }
            catch {
                # If the Excel file is open when an error will occur. -> Failed retrieving Excel workbook information
                Write-Warning "[error] $($_)"
            }
        }

        # Convert OfficeOpenXml.OfficeProperties to PSCustomObject
        [PSCustomObject] $xlInfoObj = . {
            if ($null -eq $xlInfo) {
                Write-Warning "[warn] Read Excel info is empty: `"$($BookPath)`""
                return [PSCustomObject]@{}
            }
            else {
                return [PSCustomObject]@{
                    Title = $xlInfo.Title
                    Subject = $xlInfo.Subject
                    Author = $xlInfo.Author
                    Comments = $xlInfo.Comments
                    Category = $xlInfo.Category
                    Keywords = $xlInfo.Keywords
                    Modified = $xlInfo.Modified.ToString("yyyy-MM-ddTHH:mm:ss.fffK")
                    Created = $xlInfo.Created.ToString("yyyy-MM-ddTHH:mm:ss.fffK") # ISO 8601
                    LastModifiedBy = $xlInfo.LastModifiedBy
                    # LastPrinted = $xlInfo.LastPrinted.ToString("yyyy/MM/dd HH:mm:ss")
                    Company = $xlInfo.Company
                    Manager = $xlInfo.Manager
                    Status = $xlInfo.Status
                    Application = $xlInfo.Application
                    AppVersion = $xlInfo.AppVersion
                }
            }
        }

        return $xlInfoObj
    }
}

function ReadDatabase {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [String] $DbPath
    )
    Process {
        # Read the database
        [PSCustomObject] $json = Get-Content -LiteralPath $DbPath | ConvertFrom-Json

        $db = [List[PSCustomObject]]::new()
        $json | ForEach-Object { $db.Add($_) }

        return $db
    }
}