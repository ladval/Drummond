
function ReporteLoteDrummond() {
    param (
        [string]
        $initDate,
        [string]
        $endDate
    )  
    $drummondSQLquery = @"
    SELECT id,InvoiceNumber,JsonFact,FechaFact 
    FROM [BotAbc].[dbo].[tfact_ApiProcesos] 
    WHERE NitTercero = '800021308'
    AND Created BETWEEN  
    '$initDate' AND '$endDate'
"@

    $invoicesList = SQL_Query $drummondSQLquery
    $reporteDrummond = @()
    foreach ($invoice in $invoicesList) {
        Write-Host "=========== $($invoice.InvoiceNumber) ===========" 
        $JsonFile = "$PSScriptRoot\data\$($invoice.InvoiceNumber).json"
        Set-Content -Path $JsonFile -Value $invoice.JsonFact
        $JsonData = Get-Content $JsonFile | ConvertFrom-Json
        $IVA = 0
        $RETEIVA = 0
        $RETEFUENTE = 0
        $RETEICA = 0
        foreach ($tax in $JsonData.InvoiceTaxTotal) {
            switch ($tax.Id) {
                #IVA
                "01" { 
                    Write-Host "IVA " -ForegroundColor Red -NoNewline 
                    Write-Host $tax.TaxAmount -ForegroundColor Green
                    $IVA = $tax.TaxAmount
                }
                #RETEIVA
                "05" { 
                    Write-Host "RETEIVA " -ForegroundColor Red -NoNewline 
                    Write-Host $tax.TaxAmount -ForegroundColor Green
                    $RETEIVA = $tax.TaxAmount
                }
                #RETEFUENTE
                "06" { 
                    Write-Host "RETEFUENTE " -ForegroundColor Red -NoNewline
                    Write-Host $tax.TaxAmount -ForegroundColor Green
                    $RETEFUENTE = $tax.TaxAmount
                }
                #RETEICA
                "07" { 
                    Write-Host "RETEICA " -ForegroundColor Red -NoNewline
                    Write-Host $tax.TaxAmount -ForegroundColor Green
                    $RETEICA = $tax.TaxAmount
                }
            }
        }
        $objInvoiceTotals = New-Object -TypeName psobject
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceTotalOwn" -Value $JsonData.InvoiceTotalOwn.LineExtensionAmount
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceTotalOthers" -Value $JsonData.InvoiceTotalOthers.LineExtensionAmount
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceIVA" -Value $IVA
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceRETEIVA" -Value $RETEIVA
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceRETEFUENTE" -Value $RETEFUENTE
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceRETEICA" -Value $RETEICA
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceAnticipo" -Value $JsonData.InvoiceTotal.PrePaidAmount
        
        $invoiceTotal = $JsonData.InvoiceTotal.PayableAmount - ($objInvoiceTotals.InvoiceRETEIVA + $objInvoiceTotals.InvoiceRETEICA + $objInvoiceTotals.InvoiceAnticipo)
        Write-Host "Total factura $ $invoiceTotal" -ForegroundColor White
        $objInvoiceTotals | Add-Member -MemberType NoteProperty -Name "InvoiceTotal" -Value  $invoiceTotal

        $oItemInformation = @()
        foreach ($item in $JsonData.ItemInformation) {
            $objItemInformation = New-Object -TypeName psobject
            $objItemInformation | Add-Member -MemberType NoteProperty -Name "Code" -Value $item.ItemReference
            $objItemInformation | Add-Member -MemberType NoteProperty -Name "Valor" -Value $item.LineExtensionAmount
            $oItemInformation += $objItemInformation
        }
        
        $oItemInformation = $oItemInformation | Group-Object Code | ForEach-Object {
            New-Object PSObject -Property @{
                Code  = $_.Name
                Valor = ($_.Group | Measure-Object -Property Valor -Sum).Sum
            }
        }

        $reporteTRK = reporteFacturacionDrummond_TRK $invoice.id $invoice.InvoiceNumber
        $reporteAPI = reporteFacturacionDrummond_API $reporteTRK $oItemInformation $invoice.FechaFact
        $reporteDrummond += $reporteAPI
    }
    # Remove-Item "$PSScriptRoot\data\*.json*"
    return $reporteDrummond
}

function reporteFacturacionDrummond_API {
    param (
        [Object[]]
        $reporteFacturacionDrummond_TRK,
        [Object[]]
        $oItemInformation,
        [string]
        $dateFact
    )    

    $dateFact = $($dateFact.substring(0, 10))
  
 
    $report_Numero_Factura_RPC = $reporteFacturacionDrummond_TRK.report_Numero_Factura_RPC
    $objReport = New-Object -TypeName PSObject
    $objReport | Add-Member -MemberType NoteProperty -Name "LOTE" -Value $(ItemData $oItemInformation '0000')
    $objReport | Add-Member -MemberType NoteProperty -Name "SERVICIO" -Value $(ItemData $oItemInformation '0000')
    $objReport | Add-Member -MemberType NoteProperty -Name "INVOICE" -Value $report_Numero_Factura_RPC
    $objReport | Add-Member -MemberType NoteProperty -Name "INVOICE DATE" -Value $dateFact
    $objReport | Add-Member -MemberType NoteProperty -Name "CIF US $" -Value $reporteFacturacionDrummond_TRK.report_Valor_Cif
    $objReport | Add-Member -MemberType NoteProperty -Name "TCRM" -Value $reporteFacturacionDrummond_TRK.report_Tasa_de_cambio
    $objReport | Add-Member -MemberType NoteProperty -Name "CIF COP $" -Value $reporteFacturacionDrummond_TRK.report_Valor_Pesos
    $objReport | Add-Member -MemberType NoteProperty -Name "MANEJO NAVIERO" -Value $(ItemData $oItemInformation '1034')
    $objReport | Add-Member -MemberType NoteProperty -Name "USO INSTALACIONES" -Value $(ItemData $oItemInformation '1003')
    $objReport | Add-Member -MemberType NoteProperty -Name "BODEGAJE" -Value $(ItemData $oItemInformation '1004')
    $objReport | Add-Member -MemberType NoteProperty -Name "VUCE" -Value $(ItemData $oItemInformation '1016')
    $objReport | Add-Member -MemberType NoteProperty -Name "VACIADO" -Value $(ItemData $oItemInformation '1008')
    $objReport | Add-Member -MemberType NoteProperty -Name "INSPECCION DIAN" -Value $(ItemData $oItemInformation '1042')
    $objReport | Add-Member -MemberType NoteProperty -Name "DEPOSITO CONTENEDOR" -Value $(ItemData $oItemInformation '1010')
    $objReport | Add-Member -MemberType NoteProperty -Name "LIBERACION Y MANEJO GUIA" -Value $(ItemData $oItemInformation '1005')
    $objReport | Add-Member -MemberType NoteProperty -Name "DEMORAS CONTENEDOR" -Value $(ItemData $oItemInformation '1011')
    $objReport | Add-Member -MemberType NoteProperty -Name "MOVILIZACION INSPECCION" -Value $(ItemData $oItemInformation '1043')
    $objReport | Add-Member -MemberType NoteProperty -Name "ARANCEL" -Value $(ItemData $oItemInformation '1001')
    $objReport | Add-Member -MemberType NoteProperty -Name "IVA TER" -Value $(ItemData $oItemInformation '1002')
    $objReport | Add-Member -MemberType NoteProperty -Name "MAQUINA INTERNA" -Value $(ItemData $oItemInformation '1061')
    $objReport | Add-Member -MemberType NoteProperty -Name "VISTOS BUENOS" -Value $(ItemData $oItemInformation '1023')
    $objReport | Add-Member -MemberType NoteProperty -Name "CARGUE Y DESCARGUE" -Value $(ItemData $oItemInformation '1025')
    $objReport | Add-Member -MemberType NoteProperty -Name "PORTEO" -Value $(ItemData $oItemInformation '1024')
    $objReport | Add-Member -MemberType NoteProperty -Name "INSPECCION RECONOCIMIENTO" -Value $(ItemData $oItemInformation '1046')
    $objReport | Add-Member -MemberType NoteProperty -Name "TRASLADO" -Value $(ItemData $oItemInformation '1018')
    $objReport | Add-Member -MemberType NoteProperty -Name "STACKER" -Value $(ItemData $oItemInformation '1045')
    $objReport | Add-Member -MemberType NoteProperty -Name "DAÑOS" -Value $(ItemData $oItemInformation '1050')
    $objReport | Add-Member -MemberType NoteProperty -Name "SUCIEDAD" -Value $(ItemData $oItemInformation '1026')
    $objReport | Add-Member -MemberType NoteProperty -Name "EMISION BL" -Value $(ItemData $oItemInformation '1028')
    $objReport | Add-Member -MemberType NoteProperty -Name "APERTURA Y CIERRE" -Value $(ItemData $oItemInformation '1015')
    $objReport | Add-Member -MemberType NoteProperty -Name "CERTIFICADOS" -Value $(ItemData $oItemInformation '1020')
    $objReport | Add-Member -MemberType NoteProperty -Name "BASCULAS" -Value $(ItemData $oItemInformation '1029')
    $objReport | Add-Member -MemberType NoteProperty -Name "PAPELERIA" -Value $(ItemData $oItemInformation '1038')
    $objReport | Add-Member -MemberType NoteProperty -Name "CAMBIO MODALIDAD" -Value $(ItemData $oItemInformation '1048')
    $objReport | Add-Member -MemberType NoteProperty -Name "HORA ADICIONAL" -Value $(ItemData $oItemInformation '1052')
    $objReport | Add-Member -MemberType NoteProperty -Name "CARGO FIJO" -Value $(ItemData $oItemInformation '1054')
    $objReport | Add-Member -MemberType NoteProperty -Name "MANEJO DE DOCUMENTOS" -Value $(ItemData $oItemInformation '1006')
    $objReport | Add-Member -MemberType NoteProperty -Name "CAMA ALTA-BAJA" -Value $(ItemData $oItemInformation '1044')
    $objReport | Add-Member -MemberType NoteProperty -Name "SERVICIO DE CARPADO" -Value $(ItemData $oItemInformation '1060')
    $objReport | Add-Member -MemberType NoteProperty -Name "SERVICIO CONSOLIDACION" -Value $(ItemData $oItemInformation '1007')
    $objReport | Add-Member -MemberType NoteProperty -Name "COMODATOS" -Value $(ItemData $oItemInformation '1039')
    $objReport | Add-Member -MemberType NoteProperty -Name "TRABAJOS VARIOS HORAS/HOMBRES" -Value $(ItemData $oItemInformation '1049')
    $objReport | Add-Member -MemberType NoteProperty -Name "SELLOS DE CONTENEDOR" -Value $(ItemData $oItemInformation '1051')
    $objReport | Add-Member -MemberType NoteProperty -Name "ENVIO" -Value $(ItemData $oItemInformation '1031')
    $objReport | Add-Member -MemberType NoteProperty -Name "DISMOUNTING" -Value $(ItemData $oItemInformation '1012')
    $objReport | Add-Member -MemberType NoteProperty -Name "TRANSPORTE" -Value $(ItemData $oItemInformation '1027')
    $objReport | Add-Member -MemberType NoteProperty -Name "ADICIONAL 1" -Value $(ItemData $oItemInformation '1022')
    $objReport | Add-Member -MemberType NoteProperty -Name "ADICIONAL 2" -Value $(ItemData $oItemInformation '1032')
    $objReport | Add-Member -MemberType NoteProperty -Name "ADICIONAL 3" -Value $(ItemData $oItemInformation '1041')
    $objReport | Add-Member -MemberType NoteProperty -Name "ADICIONAL 4" -Value $(ItemData $oItemInformation '1009')
    $objReport | Add-Member -MemberType NoteProperty -Name "SERVICIO EXTRAORDINARIO" -Value $(ItemData $oItemInformation '1014')
    $objReport | Add-Member -MemberType NoteProperty -Name "VR 4XMIL" -Value $(ItemData $oItemInformation '1030')
    $objReport | Add-Member -MemberType NoteProperty -Name "SUBTOTAL TER" -Value $($objInvoiceTotals.InvoiceTotalOthers)
    $objReport | Add-Member -MemberType NoteProperty -Name "SERVICIO ADUANA" -Value $(ItemData $oItemInformation '2003')
    $objReport | Add-Member -MemberType NoteProperty -Name "RECONOCIMIENTO MCIA" -Value $(ItemData $oItemInformation '2052')
    $objReport | Add-Member -MemberType NoteProperty -Name "ELABORACION REGISTROS DE IMP." -Value $(ItemData $oItemInformation '2005')
    $objReport | Add-Member -MemberType NoteProperty -Name "ELABORACION DECLARACIONES DE IMPORTACION" -Value $(ItemData $oItemInformation '2009')
    $objReport | Add-Member -MemberType NoteProperty -Name "ELABORACION DECLARACIONES DE VALOR" -Value $(ItemData $oItemInformation '2042')
    $objReport | Add-Member -MemberType NoteProperty -Name "DESCARGUE DIRECTO" -Value $(ItemData $oItemInformation '2007')
    $objReport | Add-Member -MemberType NoteProperty -Name "VISTO BUENO" -Value $(ItemData $oItemInformation '2045')
    $objReport | Add-Member -MemberType NoteProperty -Name "SUBTOTAL PROP" -Value $($objInvoiceTotals.InvoiceTotalOwn)
    $objReport | Add-Member -MemberType NoteProperty -Name "IVA PROP" -Value $objInvoiceTotals.InvoiceIVA
    $objReport | Add-Member -MemberType NoteProperty -Name "RETE IVA" -Value $objInvoiceTotals.InvoiceRETEIVA
    $objReport | Add-Member -MemberType NoteProperty -Name "RETE ICA" -Value $objInvoiceTotals.InvoiceRETEICA
    $objReport | Add-Member -MemberType NoteProperty -Name "ANTICIPO" -Value $objInvoiceTotals.InvoiceAnticipo
    $objReport | Add-Member -MemberType NoteProperty -Name "TOTAL" -Value $objInvoiceTotals.InvoiceTotal
    $objReport | Add-Member -MemberType NoteProperty -Name "PO." $reporteFacturacionDrummond_TRK.report_Pedido
    $objReport | Add-Member -MemberType NoteProperty -Name "DO" -Value $reporteFacturacionDrummond_TRK.report_DocImpoNoDO
    $objReport | Add-Member -MemberType NoteProperty -Name "DOCUMENTO DE TRANSPORTE" -Value $reporteFacturacionDrummond_TRK.report_Documento_transporte
    $objReport | Add-Member -MemberType NoteProperty -Name "CONTENEDOR" -Value $reporteFacturacionDrummond_TRK.CONTENEDOR
    $objReport | Add-Member -MemberType NoteProperty -Name "OBSERVACIONES" -Value ""
    $objReport | Add-Member -MemberType NoteProperty -Name "VALOR.FOB US" -Value $reporteFacturacionDrummond_TRK.report_Fob
    $objReport | Add-Member -MemberType NoteProperty -Name "BOOKING" -Value $(CosteoDrummondGnral $report_Numero_Factura_RPC 'FACTURA')
    $objReport | Add-Member -MemberType NoteProperty -Name "LINEA ORDEN DE COMPRA" -Value $(CosteoDrummondGnral $report_Numero_Factura_RPC 'LINEA')
    return $objReport
}  

function CosteoDrummondGnral() {
    param (
        [string]$invoiceNumber,
        [string]$dbFieldName,
        [int]$invoiceTotal
    )
    $VCosteoDrummondQuery = "SELECT * FROM [Repecev2005].[dbo].[VCosteoDrummond_fact] WHERE FACTURASERVICIOS LIKE '$invoiceNumber'"
    $VCosteoDrummond = SQL_Query $VCosteoDrummondQuery
    if ($dbFieldName.Length -gt 0) {
        foreach ($itemCosteo in $VCosteoDrummond) {
            $dbFieldValue = $itemCosteo.$dbFieldName
            If ($dbFieldValue.Length -gt 0) {
                return $dbFieldValue
            }

        }
        return 0
    }
}


function reporteFacturacionDrummond_TRK {
    param (
        [int]$idFactura,
        [string]$invoiceNumber
    )
    # Extracción nacionalización data por id de factura
    $OrdenNacIDQuery = @"
    SELECT TOP(1) 
    [OrdenNacID]
    FROM [BotAbc].[dbo].[tfact_FactNac]
    WHERE id_factura = $($idFactura)
"@
    $OrdenNacID = SQL_Query $OrdenNacIDQuery
    $OrdenNacID = $OrdenNacID.OrdenNacID

    $DeclImpoIDQuery = @"
SELECT TOP (1) [DeclImpoID]
  FROM [Repecev2005].[dbo].[IMDeclImpoOtrosDatos]
  WHERE OrdenNacID = $OrdenNacID
"@

    $DeclImpoID = SQL_Query $DeclImpoIDQuery
    
    # Generación de datros DRUMMOND por número de nacionalización
    $ReporteDrummondQuery = "[Repecev2005].[dbo].Reporte_Factuacion_Drummon $OrdenNacID"
    $data = SQL_Query $ReporteDrummondQuery
    $objReportTrk = New-Object -TypeName PSObject
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_IdFactura -Value $invoice.id
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_OrdenNacID -Value $data.OrdenNacID
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Numero_Factura_RPC -Value $invoice.invoiceNumber
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Fecha_Facturacion -value $data.Fecha_Facturacion
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Tasa_de_cambio -value $data.Tasa_de_cambio
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Valor_Cif -value $data.Valor_Cif
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Valor_Pesos -value $data.Valor_Pesos
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_DocImpoNoDO -value $data.Do
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Documento_transporte -value $data.Documento_transporte
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Fob -value $data.Fob
    $objReportTrk | Add-Member -MemberType NoteProperty -Name report_Pedido -value $data.Pedido
    $objReportTrk | Add-Member -MemberType NoteProperty -Name CONTENEDOR -value $data.CONTENEDOR
    $objReportTrk | Add-Member -MemberType NoteProperty -Name DeclImpoId -value $DeclImpoID.DeclImpoID
    return $objReportTrk
}

function ItemData {
    param (
        
        [Object[]]
        $oItemInformation,
        [string]
        $Code
    )
    foreach ($item in $oItemInformation) {
        if ($item.Code -eq $Code) {
            Write-Host $Code -ForegroundColor Red -NoNewline
            Write-Host " - " -ForegroundColor White -NoNewline
            Write-Host  "($) $($item.Valor)" -ForegroundColor Green
            return $item.Valor
        }
    }
    return 0
}

