# PAQUETES REQUERIDOS
# Install-Module -Name SqlServer
# Install-Module -Name ImportExcel


$JsonConfigFile = "$PSScriptRoot\settings\settings.json"
$Global:oJsonSettings = Get-Content -Path $JsonConfigFile | ConvertFrom-Json
. "$PSScriptRoot\sql.ps1"
. "$PSScriptRoot\jsonData.ps1"




$report = @()
$drummondSQLquery = @"
SELECT id,InvoiceNumber,JsonFact 
FROM [BotAbc].[dbo].[tfact_ApiProcesos] 
WHERE NitTercero = '800021308' 
AND Created BETWEEN  
'2021-01-08' AND '2021-30-08'
"@
$invoices = SQL_Query $drummondSQLquery
$i = 0
foreach ($invoice in $invoices) {
    JsonData $invoice.id $invoice.InvoiceNumber $invoice.InvoiceNumber.JsonFact 
    Write-Progress -Activity "Consultando facturas" -Status "Progress:" -PercentComplete ($i / $invoices.count * 100)
    # Extracción nacionalización data por id de factura
    $ordenNacIdFacturaQuery = "SELECT TOP(1) [OrdenNacID] FROM [BotAbc].[dbo].[tfact_FactNac] WHERE id_factura = $($invoice.Id)"
    $ordenNacIdFactura = SQL_Query $ordenNacIdFacturaQuery
    $ordenNacIdFactura = $ordenNacIdFactura.OrdenNacID
    # Generación de datros DRUMMOND por número de nacionalización
    $ReporteDrummondQuery = "[Repecev2005].[dbo].Reporte_Factuacion_Drummon $ordenNacIdFactura"
    $data = SQL_Query $ReporteDrummondQuery
    $objReport = New-Object -TypeName PSObject
    $objReport | Add-Member -MemberType NoteProperty -Name IdFactura -Value $invoice.Id
    $objReport | Add-Member -MemberType NoteProperty -Name OrdenNacID -Value $data.OrdenNacID
    $objReport | Add-Member -MemberType NoteProperty -Name Numero_Factura_RPC -Value $invoice.InvoiceNumber
    $objReport | Add-Member -MemberType NoteProperty -Name Fecha_Facturacion -value $data.Fecha_Facturacion
    $objReport | Add-Member -MemberType NoteProperty -Name Tasa_de_cambio -value $data.Tasa_de_cambio
    $objReport | Add-Member -MemberType NoteProperty -Name Valor_Cif -value $data.Valor_Cif
    $objReport | Add-Member -MemberType NoteProperty -Name Valor_Pesos -value $data.Valor_Pesos
    $objReport | Add-Member -MemberType NoteProperty -Name DocImpoNoDO -value $data.Do
    $objReport | Add-Member -MemberType NoteProperty -Name Documento_transporte -value $data.Documento_transporte
    $objReport | Add-Member -MemberType NoteProperty -Name Fob -value $data.Fob
    $objReport | Add-Member -MemberType NoteProperty -Name Pedido -value $data.Pedido
    $report += $objReport 
    $i = $i + 1
}
Write-Host "Terminado" -ForegroundColor green


# Format-Table -Autosize -Wrap

# $xlFile = "$PSScriptRoot\ImportExcelExample.xlsx"
# Remove-Item $xlFile -ErrorAction Ignore
# $report | Export-Excel -Path $xlFile -AutoFilter -AutoNameRange -AutoSize -Show 



