


function JsonData() {
    param (
        [int]
        $id,
        [string]
        $invoice
    )
    $drummondSQLquery = @"
    SELECT JsonFact 
    FROM [BotAbc].[dbo].[tfact_ApiProcesos] 
    WHERE id = $($id)
"@
    $jsonFact = SQL_Query $drummondSQLquery
    $jsonFact = $jsonFact.JsonFact
    $JsonDataFile = "$PSScriptRoot\data\$($invoice).json"
    Set-Content -Path $JsonDataFile -Value $jsonFact
    $JsonDataFile = Get-Content $JsonDataFile | ConvertFrom-Json
    $ItemInformation = $JsonDataFile.ItemInformation
    foreach ($item in $ItemInformation) {
        $item.Reference
    }
    
    break
}





# "MANEJO NAVIERO|1034"
# "USO INSTALACIONES|1003"
# "BODEGAJE|1004"
# "VUCE|1016"
# "VACIADO|1008"
# "INSPECCION DIAN|1042"
# "DEPOSITO CONTENEDOR|1010"
# "LIBERACION Y MANEJO GUIA|1005"
# "DEMORAS CONTENEDOR|1011"
# "MOVILIZACION INSPECCION|1043"
# "ARANCEL|1001"
# "IVA|1002"
# "MAQUINA INTERNA|1061"
# "VISTOS BUENOS|1023"
# "CARGUE Y DESCARGUE|1025"
# "PORTEO|1024"
# "INSPECCION RECONOCIMIENTO|1046"
# "TRASLADO|1018"
# "STACKER|1045"
# "DAÑOS|1050"
# "SUCIEDAD|1026"
# "EMISION BL|1028"
# "APERTURA Y CIERRE|1015"
# "CERTIFICADOS|1020"
# "BASCULAS|1029"
# "PAPELERIA|1038"
# "CAMBIO MODALIDAD|1048"
# "HORA ADICIONAL|1052"
# "CARGO FIJO|1054"
# "MANEJO DE DOCUMENTOS|1006"
# "CAMA ALTA-BAJA|1044"
# "SERVICIO DE CARPADO|1060"
# "SERVICIOCONSOLIDACION|1007"
# "COMODATOS|1039"
# "TRABAJOS VARIOS HORAS/HOMBRES|1049"
# "SELLOS DE CONTENEDOR|1051"
# "ENVIO|1031"
# "DISMOUNTING|1012"
# "TRANSPORTE|1027"
# "ADICIONAL 1|1022"
# "ADICIONAL 2|1032"
# "ADICIONAL 3|1041"
# "ADICIONAL 4|1009"
# "SERVICIO EXTRAORDINARIO|1014"
# "VR 4XMIL|1030"
# "SUBTOTAL"
# "SERVICIO ADUANA|2003"
# "RECONOCIMIENTO MCIA|2052"
# "ELABORACION REGISTROS DE IMP.|2005"
# "ELABORACIONDECLARACIONES DE IMPORTACION|2009"
# "ELABORACION DECLARACIONES DE VALOR|2042"
# "DESCARGUE DIRECTO|2007"
# "VISTO BUENO|2045"