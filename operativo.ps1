

function ReporteOperativoDrummond() {
    param (
        [string]
        $initDate,
        [string]
        $endDate
    )  
    $drummondOperativoSQLquery = @"
    SELECT id,InvoiceNumber,JsonFact 
    FROM [BotAbc].[dbo].[tfact_ApiProcesos] 
    WHERE NitTercero = '800021308' 
    AND Created BETWEEN  
    '$initDate' AND '$endDate'
"@
    $invoicesList = SQL_Query $drummondOperativoSQLquery
    $reporteDrummondOperativo = @()
    foreach ($invoice in $invoicesList) {
        $VCosteoDrummondQuery = "SELECT * FROM [Repecev2005].[dbo].[VCosteoDrummond_fact] WHERE FACTURASERVICIOS LIKE '$($invoice.InvoiceNumber)'"
        $VCosteoDrummond = SQL_Query $VCosteoDrummondQuery
        foreach ($itemCosteoDrummond in $VCosteoDrummond) {
            $facturaServicios = $itemCosteoDrummond.FACTURASERVICIOS
            $facturaServiciosSqlQuery = @"
        SELECT [JsonFact] FROM [BotAbc].[dbo].[tfact_ApiProcesos]
        WHERE InvoiceNumber = '$facturaServicios'
"@
            $facturaServiciosSqlData = SQL_Query $facturaServiciosSqlQuery
            $JsonData = $facturaServiciosSqlData.JsonFact | ConvertFrom-Json
            $invoiceTotal = $JsonData.InvoiceTotal.PayableAmount - ($objInvoiceTotals.InvoiceRETEIVA + $objInvoiceTotals.InvoiceRETEICA + $objInvoiceTotals.InvoiceAnticipo)
            $itemCosteoDrummond.VALORFACTURA = $invoiceTotal
        }
        $reporteDrummondOperativo += $VCosteoDrummond
    }
    return $reporteDrummondOperativo
}

