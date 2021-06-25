#cs ----------------------------------------------------------------------------
 AutoIt Version: 3.3.14.5
 Author: Jesús Antonio Ladino Valbuena
 Script Function:
	Generación de reporte de Drummond
#ce ----------------------------------------------------------------------------

#include <File.au3>
#include <Date.au3>
#include <Array.au3>
#include <String.au3>
#include "modulo_json.au3"
#include "modulo_sql.au3"
#include "modulo_misc.au3"

;~ Local $sSQL_InvoiceNumber = "BQA102372"
Local $sSQL_InvoiceNumber = "SMR18947"

Local $sSQL_InvoiceNumber = "SMR18921"
Local $sCosteoDrummondQuery = "SELECT * FROM [Repecev2005].[dbo].[VCosteoDrummond_fact] WHERE FACTURASERVICIOS LIKE '" & $sSQL_InvoiceNumber & "'"
Local $aTRK_Data = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)

If UBound($aTRK_Data) < 2 Then
	ConsoleWrite('DATOS INSUFICIENTES' & @CRLF)
	Exit
EndIf

_JsonFactData($sSQL_InvoiceNumber)
_JsonCodedData()
_InformeLote()

Func _InformeLote()
	Local $aIDDIM_Data = _2dArray_UniqueElements($aTRK_Data, 0) ;Se define la columna 0 como criterio de elementos repetidos, dado que campo IDDIM es único para cálculo de valor CIF
	Local $sLote = "PorDefinir"
	Local $sServicio = "PorDefinir"
	Local $sInvoice = $aTRK_Data[1][29]
	Local $sInvoiceDate = $aTRK_Data[1][31]
	Local $sCIFUS = _CIF($aIDDIM_Data, 'USD')
	Local $sTCRM = Number($aTRK_Data[1][16])
	Local $sCIFCOP = _CIF($aIDDIM_Data, 'COP')
	Local $sSubtotal ;Subtotal ubicado en zona de información de items
	Local $sSubtotal
	Local $sIVA
	Local $sReteIVA
	Local $sReteICA
	Local $sAnticipo
	Local $sTotal
	Local $sPO
	Local $sDO
	Local $sDocumentoTransporte
	Local $sContenedor
	Local $sObservaciones
	Local $sValorFobUS
	Local $sBooking
	Local $sLineaOrdenCompra
	Local $aReporteData[3][66]
;############################################################
	$aReporteData[0][0] = "LOTE"
	$aReporteData[0][1] = "SERVICIO"
	$aReporteData[0][2] = "INVOICE "
	$aReporteData[0][3] = "INVOICE DATE"
	$aReporteData[0][4] = "CIF US $"
	$aReporteData[0][5] = "TCRM"
	$aReporteData[0][6] = "CIF COP $"
;############################################################
$aReporteData[0][7]= "MANEJO NAVIERO|1034"
$aReporteData[0][8]= "USO INSTALACIONES|1003"
$aReporteData[0][9]= "BODEGAJE|1004"
$aReporteData[0][10] = "VUCE|1016"
$aReporteData[0][11] = "VACIADO|1008"
$aReporteData[0][12] = "INSPECCION DIAN|1042"
$aReporteData[0][13] = "DEPOSITO CONTENEDOR|1010"
$aReporteData[0][14] = "LIBERACION Y MANEJO GUIA|1005"
$aReporteData[0][15] = "DEMORAS CONTENEDOR|1011"
$aReporteData[0][16] = "MOVILIZACION INSPECCION|1043"
$aReporteData[0][17] = "ARANCEL|1001"
$aReporteData[0][18] = "IVA|1002"
$aReporteData[0][19] = "MAQUINA INTERNA|1061"
$aReporteData[0][20] = "VISTOS BUENOS|1023"
$aReporteData[0][21] = "CARGUE Y DESCARGUE|1025"
$aReporteData[0][22] = "PORTEO|1024"
$aReporteData[0][23] = "INSPECCION RECONOCIMIENTO|1046"
$aReporteData[0][24] = "TRASLADO|1018"
$aReporteData[0][25] = "STACKER|1045"
$aReporteData[0][26] = "DAÑOS|1050"
$aReporteData[0][27] = "SUCIEDAD|1026"
$aReporteData[0][28] = "EMISION BL|1028"
$aReporteData[0][29] = "APERTURA Y CIERRE|1015"
$aReporteData[0][30] = "CERTIFICADOS|1020"
$aReporteData[0][31] = "BASCULAS|1029"
$aReporteData[0][32] = "PAPELERIA|1038"
$aReporteData[0][33] = "CAMBIO MODALIDAD|1048"
$aReporteData[0][34] = "HORA ADICIONAL|1052"
$aReporteData[0][35] = "CARGO FIJO|1054"
$aReporteData[0][36] = "MANEJO DE DOCUMENTOS|1006"
$aReporteData[0][37] = "CAMA ALTA-BAJA|1044"
$aReporteData[0][38] = "SERVICIO DE CARPADO|1060"
$aReporteData[0][39] = "SERVICIOCONSOLIDACION|1007"
$aReporteData[0][40] = "COMODATOS|1039"
$aReporteData[0][41] = "TRABAJOS VARIOS HORAS/HOMBRES|1049"
$aReporteData[0][42] = "SELLOS DE CONTENEDOR|1051"
$aReporteData[0][42] = "ENVIO|1031"
$aReporteData[0][43] = "DISMOUNTING|1012"
$aReporteData[0][44] = "TRANSPORTE|1027"
$aReporteData[0][45] = "ADICIONAL 1|1022"
$aReporteData[0][46] = "ADICIONAL 2|1032"
$aReporteData[0][47] = "ADICIONAL 3|1041"
$aReporteData[0][48] = "ADICIONAL 4|1009"
$aReporteData[0][49] = "SERVICIO EXTRAORDINARIO|1014"
$aReporteData[0][50] = "VR 4XMIL|1030"
;############################################################
$aReporteData[0][51] = "SUBTOTAL"
;############################################################
	$aReporteData[0][52] = "SERVICIO ADUANA|2003"
	$aReporteData[0][53] = "RECONOCIMIENTO MCIA|2052"
	$aReporteData[0][54] = "ELABORACION REGISTROS DE IMP.|2005"
	$aReporteData[0][55] = "ELABORACIONDECLARACIONES DE IMPORTACION|2009"
	$aReporteData[0][56] = "ELABORACION DECLARACIONES DE VALOR|2042"
	$aReporteData[0][57] = "DESCARGUE DIRECTO|2007"
	$aReporteData[0][58] = "VISTO BUENO|2045"
;############################################################
	$aReporteData[0][59] = "SUBTOTAL"
	$aReporteData[0][60] = "IVA"
	$aReporteData[0][61] = "RETE IVA"
	$aReporteData[0][62] = "RETE ICA"
	$aReporteData[0][63] = "ANTICIPO"
	$aReporteData[0][64] = "TOTAL"
	$aReporteData[0][65] = "PO."
	$aReporteData[0][66] = "DO"
	$aReporteData[0][67] = "DOCUMENTO  DE TRANSPORTE"
	$aReporteData[0][68] = "CONTENEDOR"
	$aReporteData[0][69]= "OBSERVACIONES"
	$aReporteData[0][70] = "VALOR.FOB US"
	$aReporteData[0][71] = "BOOKING"
	$aReporteData[0][72] = "LINEA ORDEN DE COMPRA"
	;######################################################################
	$aReporteData[1][0] = $sLote
	$aReporteData[1][1] = $sServicio
	$aReporteData[1][2] = $sInvoice
	$aReporteData[1][3] = $sInvoiceDate
	$aReporteData[1][4] = $sCIFUS
	$aReporteData[1][5] = $sTCRM
	$aReporteData[1][6] = $sCIFCOP
	_ArrayDisplay($aReporteData, '$aReporteData')
EndFunc ;==>_InformeLote

Func _JsonFactData($sSQL_InvoiceNumber)
	Local $sJsonFactQuery = "SELECT JsonFact FROM [BotAbc].[dbo].[tfact_ApiProcesos] where InvoiceNumber = '" & $sSQL_InvoiceNumber & "'"
	Local $sDrummondConceptosQuery = "SELECT * FROM [BotRepecev].[dbo].[tdrummond_Conceptos]"
	Local $aAPI_Data = _ModuloSQL_SQL_SELECT($sJsonFactQuery)
	Local $aDrummondConceptos = _ModuloSQL_SQL_SELECT($sDrummondConceptosQuery)
	Local $sJsonFact = $aAPI_Data[1][0]
	Local $sJsonFile_InvoiceData = @ScriptDir & '\data\' & $sSQL_InvoiceNumber & '.json'
	_SaveDataToFile($sJsonFile_InvoiceData, $sJsonFact)
	Local $sJsonInvoiceData = _ReadDataFromFile($sJsonFile_InvoiceData)
	StringReplace($sJsonInvoiceData, '"ItemReference":', "")
	Local $iItemsNumber = @extended - 1
	Local $aItemReferences[@extended][4]
	Local $oJsonInvoiceData = Json_Decode($sJsonInvoiceData)
	For $i = 0 To $iItemsNumber Step +1
		Local $iItemReferenceCode = Json_Get($oJsonInvoiceData, ".ItemInformation[" & $i & "].ItemReference")
		Local $iItemReferenceName = Json_Get($oJsonInvoiceData, ".ItemInformation[" & $i & "].Name")
		Local $iItemReferenceAmount = Json_Get($oJsonInvoiceData, ".ItemInformation[" & $i & "].LineExtensionAmount")
		Local $aIndexNoConcepto = _ArraySearch($aDrummondConceptos, $iItemReferenceCode, Default, Default, Default, Default, Default, 1, False)
		$aItemReferences[$i][0] = $iItemReferenceCode
		$aItemReferences[$i][1] = $iItemReferenceName
		$aItemReferences[$i][2] = $aDrummondConceptos[$aIndexNoConcepto][2]
		$aItemReferences[$i][3] = $iItemReferenceAmount
	Next
	Local $sCheckItems = _CheckItems($aItemReferences)
	If $sCheckItems Then
		ConsoleWrite('Fin funcion JsonFactData' & @CRLF)
	Else
		ConsoleWrite($sCheckItems & @CRLF)
	EndIf
EndFunc ;==>_JsonFactData

Func _CheckItems($aArray)
	Local $aUniqueItemReferences = _ArrayUnique($aArray, Default, Default, Default, 0)
	Local $iFactItems = UBound($aArray)
	Local $iUniqueFactItems = UBound($aUniqueItemReferences)
	Local $iDiffItems = ($iFactItems - $iUniqueFactItems)
	If $iDiffItems = 0 Then
		Return True
	Else
		Return 'Hay ' & $iDiffItems & ' elementos repetidos en esta factura. Por favor validar' & @CRLF
	EndIf
EndFunc ;==>_CheckItems

Func _JsonCodedData()
	Local $sManejoNaviero = '1034'
	Local $sUsoInstalaciones = '1003'
	Local $sBodegaje = '1004'
	Local $sVuce = '1016'
	Local $sVaciado = '1008'
	Local $sInspeccionDian = '1042'
	Local $sDepositoContenedor = '1010'
	Local $sLiberacionManejoGuia = '1005'
	Local $sDemorasContenedor = '1011'
	Local $sMovilizacionInspeccion = '1043'
	Local $sArancel = '1001'
	Local $sIVA = '1002'
	Local $sMaquinaInterna = '1061'
	Local $sVistosBuenos = '1023'
	Local $sCargueDescargue = '1025'
	Local $sPorteo = '1024'
	Local $sInspeccionReconocimiento = '1046'
	Local $sTraslado = '1018'
	Local $sStacker = '1045'
	Local $sDanos = '1050'
	Local $sSuciedad = '1026'
	Local $sEmisionBL = '1028'
	Local $sAperturaCierre = '1015'
	Local $sCertificados = '1020'
	Local $sBasculas = '1029'
	Local $sPapeleria = '1038'
	Local $sCambioModalidad = '1048'
	Local $sHoraAdicional = '1052'
	Local $sCargoFijo = '1054'
	Local $sManejoDocumentos = '1006'
	Local $sCamaAltaBaja = '1044'
	Local $sServicioCarpado = '1060'
	Local $sServicioConsolidacion = '1007'
	Local $sComodatos = '1039'
	Local $sTrabajosVariosHorasHombres = '1049'
	Local $sSellosContenedor = '1051'
	Local $sEnvio = '1031'
	Local $sDismounting = '1012'
	Local $sTransporte = '1027'
	Local $sReempaque = '1022'
	Local $sAdicional2 = '1032'
	Local $sAdicional3 = '1041'
	Local $sAdicional4 = '1009'
	Local $sServicioExtraordinario = '1014'
	Local $sVr4xmil = '1030'
	Local $sSubtotal = 0
	Local $sServicioAduana = '2003'
	Local $sReconocimientoMcia = '2052'
	Local $sElaboracionRegistrosImp = '2005'
	Local $sElaboracionDeclaracionesImportacion = '2009'
	Local $sElaboracionDeclaracionesValor = '2042'
	Local $sDescargueDirecto = '2007'
	Local $sVistoBueno = '2045'
	Local $aJsonReportData[52]
	$aJsonReportData[0] = $sManejoNaviero
	$aJsonReportData[1] = $sUsoInstalaciones
	$aJsonReportData[2] = $sBodegaje
	$aJsonReportData[3] = $sVuce
	$aJsonReportData[4] = $sVaciado
	$aJsonReportData[5] = $sInspeccionDian
	$aJsonReportData[6] = $sDepositoContenedor
	$aJsonReportData[7] = $sLiberacionManejoGuia
	$aJsonReportData[8] = $sDemorasContenedor
	$aJsonReportData[9] = $sMovilizacionInspeccion
	$aJsonReportData[10] = $sArancel
	$aJsonReportData[11] = $sIVA
	$aJsonReportData[12] = $sMaquinaInterna
	$aJsonReportData[13] = $sVistosBuenos
	$aJsonReportData[14] = $sCargueDescargue
	$aJsonReportData[15] = $sPorteo
	$aJsonReportData[16] = $sInspeccionReconocimiento
	$aJsonReportData[17] = $sTraslado
	$aJsonReportData[18] = $sStacker
	$aJsonReportData[19] = $sDanos
	$aJsonReportData[20] = $sSuciedad
	$aJsonReportData[21] = $sEmisionBL
	$aJsonReportData[22] = $sAperturaCierre
	$aJsonReportData[23] = $sCertificados
	$aJsonReportData[24] = $sBasculas
	$aJsonReportData[25] = $sPapeleria
	$aJsonReportData[26] = $sCambioModalidad
	$aJsonReportData[27] = $sHoraAdicional
	$aJsonReportData[28] = $sCargoFijo
	$aJsonReportData[29] = $sManejoDocumentos
	$aJsonReportData[30] = $sCamaAltaBaja
	$aJsonReportData[31] = $sServicioCarpado
	$aJsonReportData[32] = $sServicioConsolidacion
	$aJsonReportData[33] = $sComodatos
	$aJsonReportData[34] = $sTrabajosVariosHorasHombres
	$aJsonReportData[35] = $sSellosContenedor
	$aJsonReportData[36] = $sEnvio
	$aJsonReportData[37] = $sDismounting
	$aJsonReportData[38] = $sTransporte
	$aJsonReportData[39] = $sReempaque
	$aJsonReportData[40] = $sAdicional2
	$aJsonReportData[41] = $sAdicional3
	$aJsonReportData[42] = $sAdicional4
	$aJsonReportData[43] = $sServicioExtraordinario
	$aJsonReportData[44] = $sVr4xmil
	$aJsonReportData[45] = $sServicioAduana
	$aJsonReportData[46] = $sReconocimientoMcia
	$aJsonReportData[47] = $sElaboracionRegistrosImp
	$aJsonReportData[48] = $sElaboracionDeclaracionesImportacion
	$aJsonReportData[49] = $sElaboracionDeclaracionesValor
	$aJsonReportData[50] = $sDescargueDirecto
	$aJsonReportData[51] = $sVistoBueno
EndFunc ;==>_JsonCodedData

Func _2dArray_UniqueElements($aArray, $iCol)
	Local $aArrayUnique = _ArrayUnique($aArray, $iCol, 1, Default, 0) ;Configuración que permite seleccionar únicamente la información no repetida de la tabla sin header ni COUNT
	Local $iUniqueElements = (UBound($aArrayUnique) - 1)
	Local $aUniqueElements[0][UBound($aArray, 2)]
	For $i = 0 To $iUniqueElements Step +1
		Local $IdDimUnique = $aArrayUnique[$i]
		For $j = 0 To UBound($aArray) - 1 Step +1
			Local $IdDim = $aArray[$j][0]
			If $IdDim == $IdDimUnique Then
				Local $aExtracted_Data = _ArrayExtract($aArray, $j, $j)
				_ArrayAdd($aUniqueElements, $aExtracted_Data)
				ExitLoop
			EndIf
		Next
	Next
	Return $aUniqueElements
EndFunc ;==>_2dArray_UniqueElements

Func _CIF($aIDDIM_Data, $sCurrency)
	Local $iCurrency = 1
	Local $iCIF = 0
	For $i = 0 To UBound($aIDDIM_Data) - 1 Step +1
		If $sCurrency = 'COP' Then $iCurrency = Number($aIDDIM_Data[$i][16])
		$iCIF += (Number($aIDDIM_Data[$i][26]) * $iCurrency)
	Next
	$iCIF = Round($iCIF, 2)
	Return $iCIF
EndFunc ;==>_CIF





;~ Exit


;~ Local $aInvoices = _ExtractSingleInvoices($aTRK_Data)
;~ _ArrayDisplay($aInvoices, '$aInvoices')


;~ For $i = 1 To UBound($aInvoices) - 1 Step +1 ;Starts at 1. First position indicates invoices amount.
;~ 	Local $sInvoiceNumber = $aInvoices[$i]
;~ 	Local $aInvoiceCoincidences = _ArrayFindAll($aTRK_Data, $sInvoiceNumber, Default, Default, Default, Default, 29, False)
;~ 	_ArrayDisplay($aInvoiceCoincidences, $sInvoiceNumber)
;~ Next

;~ _ArrayDisplay($aInvoices, '$aTRK_Data')

Func _ExtractSingleInvoices($aArray)
	Local $aTempArray[1] = ['']
	For $i = 1 To UBound($aArray) - 1
		Local $sInvoiceNumber = $aArray[$i][29] ;Col 29: Where SIIGO Invoice number is lkocated in the array
		_ArraySearch($aTempArray, $sInvoiceNumber)
		If @error Then _ArrayAdd($aTempArray, $sInvoiceNumber)
	Next
	$aTempArray[0] = UBound($aTempArray) - 1
	Return $aTempArray
EndFunc ;==>_ExtractSingleInvoices


