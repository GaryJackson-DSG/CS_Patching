param(
    [parameter(Mandatory=$true)]
    [String]
    $ReportLogPath,

    [parameter(Mandatory=$true)]
    [String]
    $LogFileSearchPattern,

    [parameter(Mandatory=$true)]
    [String]
    $ExcelPathFileName
)

#   Example: .\ExportLogToExcel.ps1 -ReportLogPath "D:\" -LogFileSearchPattern "Full*Log" -ExcelPathFileName "D:\Results.xlsx"

#   Check for ImportExcel Installed
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -AllowClobber -Force
}

#   Check for ImportExcel Imported
if (!(Get-Module -Name ImportExcel)) {
    Import-Module ImportExcel -Force
}

#   Get List of Log Files
$FilesToMerge = Get-ChildItem -Path $ReportLogPath -Filter $LogFileSearchPattern -File

#   Process each Log File
foreach ($File in $FilesToMerge)
{
    Write-Output ("Processing " + $ReportLogPath + $File.Name);

    #   Read each log file, convert to csv, export to excel
    Get-Content "$file" | ConvertFrom-Csv -Delimiter "," | Export-Excel -Path $ExcelPathFileName -Append
}