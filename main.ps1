# PAQUETES REQUERIDOS
# Install-Module -Name SqlServer
# Install-Module -Name ImportExcel

$JsonConfigFile = "$PSScriptRoot\settings\settings.json"
$Global:oJsonSettings = Get-Content -Path $JsonConfigFile | ConvertFrom-Json

. "$PSScriptRoot\sql.ps1"
. "$PSScriptRoot\lote.ps1"
. "$PSScriptRoot\operativo.ps1"

$initDate = "2021-03-09"
$endDate = "2021-06-09"

$ReporteLoteDrummond = ReporteLoteDrummond $initDate $endDate
$ReporteLoteCount = $ReporteLoteDrummond | Measure-Object
$ReporteLoteCount = $ReporteLoteCount.Count + 1

$ReporteOperativoDrummond = ReporteOperativoDrummond $initDate $endDate 
$ReporteOperativoCount = $ReporteOperativoDrummond | Measure-Object
$ReporteOperativoCount = $ReporteOperativoCount.Count + 1

$date = Get-Date -Format "ddMMyyyy"
$xlFile = "$PSScriptRoot\reports\ReporteDrummond_$($date).xlsx"

$ReporteLoteDrummond | Export-Excel -Path $xlFile -AutoNameRange -WorksheetName 'LOTE' -AutoSize -CellStyleSB {
    param (
        $workSheet
    ) 
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV1" -Bold
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV1" -FontName 'Arial' -FontSize 12
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)"  -AutoFit
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)" -Width 35 -WrapText
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)" -HorizontalAlignment Center
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)" -VerticalAlignment Center
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)" -BorderTop Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)" -BorderBottom Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)" -BorderLeft Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:BV$($ReporteLoteCount)" -BorderRight Thin
        
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
        
    Set-ExcelRange -Worksheet $WorkSheet -Range "E2:BN$($ReporteLoteCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "BT2:BT$($ReporteLoteCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "BP2:BP$($ReporteLoteCount)" -NumberFormat "0"
    Set-ExcelRange -Worksheet $WorkSheet -Range "D2:D$($ReporteLoteCount)" -NumberFormat "Short Date"
        
}
    
$ReporteOperativoDrummond | Select-Object -Property  * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors `
| Export-Excel -Path $xlFile -AutoNameRange -WorksheetName 'Operativo' -AutoSize -CellStyleSB {
    param (
        $workSheet
    ) 
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK1" -Bold
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK1" -BackgroundColor powderBlue
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK1" -FontName 'Arial' -FontSize 12
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)"  -AutoFit
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)" -Width 35 -WrapText
    Set-ExcelRange -Worksheet $WorkSheet -Range "A2:AK$($ReporteOperativoCount)" -BackgroundColor azure
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)" -HorizontalAlignment Center
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)" -VerticalAlignment Center
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)" -BorderTop Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)" -BorderBottom Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)" -BorderLeft Thin
    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:AK$($ReporteOperativoCount)" -BorderRight Thin

    Set-ExcelRange -Worksheet $WorkSheet -Range "A1:A$($ReporteOperativoCount)" -NumberFormat "0"
    Set-ExcelRange -Worksheet $WorkSheet -Range "B1:B$($ReporteOperativoCount)" -NumberFormat "0"
    Set-ExcelRange -Worksheet $WorkSheet -Range "C1:D$($ReporteOperativoCount)" -NumberFormat "0"
    Set-ExcelRange -Worksheet $WorkSheet -Range "H1:H$($ReporteOperativoCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "K1:K$($ReporteOperativoCount)" -NumberFormat "0"
    Set-ExcelRange -Worksheet $WorkSheet -Range "L1:L$($ReporteOperativoCount)" -NumberFormat "Short Date"
    Set-ExcelRange -Worksheet $WorkSheet -Range "M1:M$($ReporteOperativoCount)" -NumberFormat "0"
    Set-ExcelRange -Worksheet $WorkSheet -Range "Q1:Q$($ReporteOperativoCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "T1:V$($ReporteOperativoCount)" -NumberFormat "Short Date"
    Set-ExcelRange -Worksheet $WorkSheet -Range "W1:AA$($ReporteOperativoCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "AE:AE$($ReporteOperativoCount)" -NumberFormat "Currency"
    Set-ExcelRange -Worksheet $WorkSheet -Range "AF:AF$($ReporteOperativoCount)" -NumberFormat "Short Date"
    Set-ExcelRange -Worksheet $WorkSheet -Range "AI:AI$($ReporteOperativoCount)" -NumberFormat "Currency"
}
