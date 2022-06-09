

function ReporteOperativoDrummond() {
    param (
        [string]
        $initDate,
        [string]
        $endDate,
        $objInvoiceTotals
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


            $JsonFile = "$PSScriptRoot\data\$($facturaServicios).json"
            $JsonData = Get-Content $JsonFile | ConvertFrom-Json
            $RETEIVA = 0
            $RETEFUENTE = 0
            $RETEICA = 0
            foreach ($tax in $JsonData.InvoiceTaxTotal) {
                switch ($tax.Id) {
                    #RETEIVA
                    "05" { 
                        $RETEIVA = $tax.TaxAmount
                    }
                    #RETEFUENTE
                    "06" { 
                        $RETEFUENTE = $tax.TaxAmount
                    }
                    #RETEICA
                    "07" { 
                        $RETEICA = $tax.TaxAmount
                    }
                }
            }
            $invoiceTotal = $JsonData.InvoiceTotal.PayableAmount - ($RETEIVA + $RETEFUENTE + $RETEICA + $JsonData.InvoiceTotal.PrePaidAmount)
            $itemCosteoDrummond.VALORFACTURA = $invoiceTotal
            $itemCosteoDrummond.PEDIDO = "'$($itemCosteoDrummond.PEDIDO)"
        }
        $reporteDrummondOperativo += $VCosteoDrummond
    }
    
    

    return $reporteDrummondOperativo
}

