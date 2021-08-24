

# PAQUETES REQUERIDOS
# Install-Module -Name SqlServer
# Install-Module -Name ImportExcel

$JsonConfigFile = "$PSScriptRoot\settings\settings.json"
$Global:oJsonSettings = Get-Content -Path $JsonConfigFile | ConvertFrom-Json

. "$PSScriptRoot\sql.ps1"
. "$PSScriptRoot\jsonData.ps1"
. "$PSScriptRoot\test.ps1"

$drummondSQLquery = @"
SELECT id,InvoiceNumber,JsonFact 
FROM [BotAbc].[dbo].[tfact_ApiProcesos] 
WHERE NitTercero = '800021308' 
AND Status = 'Recibido Cliente'
AND Created BETWEEN  
'2021-12-08' AND '2021-14-08'
"@

$invoicesList = SQL_Query $drummondSQLquery

# $invoicesList | Format-Table -AutoSize
# break

$reporteDrummond = ReporteDrummond $invoicesList


$ReportInvoicesCount = $reporteDrummond | Measure-Object
$ReportInvoicesCount = $ReportInvoicesCount.Count + 1


Remove-Item "$PSScriptRoot\data\*.json*"
$xlFile = "$PSScriptRoot\ImportExcelExample.xlsx"
Remove-Item $xlFile -ErrorAction Ignore
$reporteDrummond | Export-Excel -Path $xlFile -AutoNameRange -AutoSize -Show -CellStyleSB {
    
    param (
        $workSheet
    ) 

    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV1" -Bold
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV1" -FontName 'Arial' -FontSize 12
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)"  -AutoFit
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)" -Width 35 -WrapText
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)" -HorizontalAlignment Center
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)" -VerticalAlignment Center
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)" -BorderTop Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)" -BorderBottom Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)" -BorderLeft Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReportInvoicesCount)" -BorderRight Thin

    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:B1" -BackgroundColor White -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "C1:D1" -BackgroundColor Bisque -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "E1" -BackgroundColor SkyBlue -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "F1" -BackgroundColor SkyBlue -FontColor Red
    Set-ExcelRange -Worksheet $WorkSheet -Range "G1" -BackgroundColor Gold -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "H1:AZ1" -BackgroundColor Bisque -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "BA1" -BackgroundColor Gold -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "BB1:BH1" -BackgroundColor Bisque -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "BI1:BN1" -BackgroundColor Gold -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "BO1" -BackgroundColor SkyBlue -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "BP1" -BackgroundColor Gold -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "BQ1" -BackgroundColor SkyBlue -FontColor Black
    Set-ExcelRange -Worksheet $WorkSheet -Range "BR1:BS1" -BackgroundColor SkyBlue -FontColor Red
    Set-ExcelRange -Worksheet $WorkSheet -Range "BT1:BV1" -BackgroundColor SkyBlue -FontColor Black

    Set-ExcelRange -Worksheet $WorkSheet -Range "E2:BN$($ReportInvoicesCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "BT2:BT$($ReportInvoicesCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "BP2:BP$($ReportInvoicesCount)" -NumberFormat "0"
    Set-ExcelRange -Worksheet $WorkSheet -Range "D2:D$($ReportInvoicesCount)" -NumberFormat "Short Date"

}



