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

Local $sInvoiceNumber = "SMR18921"
Local $sCosteoDrummondQuery = "SELECT * FROM [Repecev2005].[dbo].[VCosteoDrummond_fact] WHERE FACTURASERVICIOS LIKE '" & $sInvoiceNumber & "'"
Local $aTRK_Data = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)
If UBound($aTRK_Data) < 2 Then
	ConsoleWrite('DATOS INSUFICIENTES' & @CRLF)
	Exit
EndIf
Local $sJsonFile_InvoiceData = @ScriptDir & '\data\' & $sInvoiceNumber & '.json'
Local $sJsonFactQuery = "SELECT JsonFact FROM [BotAbc].[dbo].[tfact_ApiProcesos] where InvoiceNumber = '" & $sInvoiceNumber & "'"
Local $aAPI_Data = _ModuloSQL_SQL_SELECT($sJsonFactQuery)
_SaveDataToFile($sJsonFile_InvoiceData, $aAPI_Data[1][0])
Local $sJsonInvoiceData = _ReadDataFromFile($sJsonFile_InvoiceData) ;~ Json_Dump($sJsonInvoiceData)
Local $oJsonInvoiceAllData = Json_Decode($sJsonInvoiceData)
Local $aJsonFactItemsData = _JsonFactItemsData($oJsonInvoiceAllData)
Local $aJsonFactTaxesData = _JsonFactTaxesData($oJsonInvoiceAllData)
If IsArray($aJsonFactItemsData) And IsArray($aJsonFactTaxesData) Then
	Local $aInformeLote = _InformeLote($aTRK_Data, $aJsonFactItemsData, $aJsonFactTaxesData)
	_ArrayDisplay($aInformeLote, '$aInformeLote')
EndIf

Func _InformeLote($aTRK_Data, $aJsonFactItemsData, $aJsonFactTaxesData)
	Local $aIDDIM_Data = _2dArray_UniqueElements($aTRK_Data, 0) ;Se define la columna 0 como criterio de elementos repetidos, dado que campo IDDIM es único para cálculo de valor CIF
	Local $aReporteData = _ReporteData()
	Local $sLote = "PorDefinir"
	Local $sServicio = "PorDefinir"
	Local $sInvoice = $aTRK_Data[1][29]
	Local $sInvoiceDate = $aTRK_Data[1][31]
	Local $sCIFUS = _CIF($aIDDIM_Data, 'USD')
	Local $sTCRM = Number($aTRK_Data[1][16])
	Local $sCIFCOP = _CIF($aIDDIM_Data, 'COP')
	$aReporteData[2][0] = $sLote
	$aReporteData[2][1] = $sServicio
	$aReporteData[2][2] = $sInvoice
	$aReporteData[2][3] = $sInvoiceDate
	$aReporteData[2][4] = $sCIFUS
	$aReporteData[2][5] = $sTCRM
	$aReporteData[2][6] = $sCIFCOP
	For $j = 0 To UBound($aReporteData, 2) - 1 Step +1
		Local $sDescription = $aReporteData[0][$j]
		Local $aDescription = StringSplit($sDescription, '|', 3)
		If UBound($aDescription) = 2 Then
			$aReporteData[0][$j] = $aDescription[0]
			$aReporteData[1][$j] = 'SIIGO Code: ' & $aDescription[1]
			Local $iSIIGOCode = $aDescription[1]
			For $i = 0 To UBound($aJsonFactItemsData) - 1 Step +1
				Local $iJsonSIIGOCode = $aJsonFactItemsData[$i][1]
				Local $iJsonSIIGODesc = $aJsonFactItemsData[$i][2]
				Local $iJsonSIIGOValue = $aJsonFactItemsData[$i][4]
				If Number($iSIIGOCode) == Number($iJsonSIIGOCode) Then
					$aReporteData[2][$j] = $iJsonSIIGOValue
					$aReporteData[3][$j] = $iJsonSIIGODesc
				EndIf
			Next
		EndIf
	Next
	Local $iSubtotalOthers = $aJsonFactItemsData[1][5]
	Local $iSubtotalOwn = $aJsonFactItemsData[1][6]
	Local $iJsonFactIVA = $aJsonFactTaxesData[0]
	Local $iJsonFactReteIVA = $aJsonFactTaxesData[1]
	Local $iJsonFactReteICA = $aJsonFactTaxesData[2]
	Local $iJsonFactAnticipo
	Local $iJsonFactTotal
	$aReporteData[2][51] = $iSubtotalOthers
	$aReporteData[2][59] = $iSubtotalOwn
	$aReporteData[2][60] = $iJsonFactIVA
	$aReporteData[2][61] = $iJsonFactReteIVA
	$aReporteData[2][62] = $iJsonFactReteICA
	Return $aReporteData
EndFunc   ;==>_InformeLote

Func _JsonFactItemsData($oJsonInvoiceAllData)
	Local $sDrummondConceptosQuery = "SELECT * FROM [BotRepecev].[dbo].[tdrummond_Conceptos]"
	Local $aDrummondConceptos = _ModuloSQL_SQL_SELECT($sDrummondConceptosQuery)
	Local $iItemsNumber = UBound(Json_Get($oJsonInvoiceAllData, '.ItemInformation'))  ;Returns Items number with array lenght fron JSON_GET function
	Local $aItemReferences[$iItemsNumber][7]
	For $i = 0 To $iItemsNumber - 1 Step +1
		Local $iItemReferenceCode = Json_Get($oJsonInvoiceAllData, ".ItemInformation[" & $i & "].ItemReference")
		Local $iItemReferenceName = Json_Get($oJsonInvoiceAllData, ".ItemInformation[" & $i & "].Name")
		Local $iItemReferenceAmount = Json_Get($oJsonInvoiceAllData, ".ItemInformation[" & $i & "].LineExtensionAmount")
		Local $iSubtotalOwn = Json_Get($oJsonInvoiceAllData, ".InvoiceTotalOwn.LineExtensionAmount")
		Local $iSubtotalOthers = Json_Get($oJsonInvoiceAllData, ".InvoiceTotalOthers.LineExtensionAmount")
		;.InvoiceTotal.PrePaidAmount  =0
		;.InvoiceTotal.PayableAmount  =2082947
		Local $aIndexNoConcepto = _ArraySearch($aDrummondConceptos, $iItemReferenceCode, Default, Default, Default, Default, Default, 1, False)
		$aItemReferences[$i][0] = ''
		$aItemReferences[$i][1] = $iItemReferenceCode
		$aItemReferences[$i][2] = $iItemReferenceName
		$aItemReferences[$i][3] = $aDrummondConceptos[$aIndexNoConcepto][2]
		$aItemReferences[$i][4] = $iItemReferenceAmount
		$aItemReferences[$i][5] = $iSubtotalOthers
		$aItemReferences[$i][6] = $iSubtotalOwn
	Next
	Local $sCheckItems = _CheckItems($aItemReferences)
	If $sCheckItems Then
		Return $aItemReferences
	Else
		ConsoleWrite($sCheckItems & @CRLF)
		Return False
	EndIf
EndFunc   ;==>_JsonFactItemsData

Func _JsonFactTaxesData($oJsonInvoiceAllData)
	Local $aInvoiceTaxTotal = Json_Get($oJsonInvoiceAllData, ".InvoiceTaxTotal")
	Local $iTaxItems = UBound($aInvoiceTaxTotal) - 1
	Local $aTaxes[3] = [0, 0, 0]
	For $i = 0 To $iTaxItems Step +1
		Local $sJsonId = Json_Get($oJsonInvoiceAllData, '.InvoiceTaxTotal[' & $i & '].Id')
		Local $sJsonTaxAmount = Json_Get($oJsonInvoiceAllData, '.InvoiceTaxTotal[' & $i & '].TaxAmount')
		Select
			Case $sJsonId = '01' ;IVA
				$aTaxes[0] = $sJsonTaxAmount
			Case $sJsonId = '05' ;RETEIVA
				$aTaxes[1] = $sJsonTaxAmount
			Case $sJsonId = '07' ;RETEICA
				$aTaxes[2] = $sJsonTaxAmount
			Case Else
				ConsoleWrite('Id: ' & $sJsonId & '. Impuesto no parametrizado' & @CRLF)
		EndSelect
	Next
	Return $aTaxes
EndFunc   ;==>_JsonFactTaxesData

Func _CheckItems($aArray)
	Local $aUniqueItemReferences = _ArrayUnique($aArray, 1, Default, Default, 0)
	Local $iFactItems = UBound($aArray)
	Local $iUniqueFactItems = UBound($aUniqueItemReferences)
	If $iFactItems = $iUniqueFactItems Then
		Return True
	Else
		ConsoleWrite('Hay elementos repetidos en esta factura. Por favor validar' & @CRLF)
		Return False
	EndIf
EndFunc   ;==>_CheckItems

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
EndFunc   ;==>_2dArray_UniqueElements

Func _CIF($aIDDIM_Data, $sCurrency)
	Local $iCurrency = 1
	Local $iCIF = 0
	For $i = 0 To UBound($aIDDIM_Data) - 1 Step +1
		If $sCurrency = 'COP' Then $iCurrency = Number($aIDDIM_Data[$i][16])
		$iCIF += (Number($aIDDIM_Data[$i][26]) * $iCurrency)
	Next
	$iCIF = Round($iCIF, 2)
	Return $iCIF
EndFunc   ;==>_CIF

Func _ReporteData()
	Local $aReporteData[4][73]
	;############################################################
	$aReporteData[0][0] = "LOTE"
	$aReporteData[0][1] = "SERVICIO"
	$aReporteData[0][2] = "INVOICE "
	$aReporteData[0][3] = "INVOICE DATE"
	$aReporteData[0][4] = "CIF US $"
	$aReporteData[0][5] = "TCRM"
	$aReporteData[0][6] = "CIF COP $"
	;############################################################
	$aReporteData[0][7] = "MANEJO NAVIERO|1034"
	$aReporteData[0][8] = "USO INSTALACIONES|1003"
	$aReporteData[0][9] = "BODEGAJE|1004"
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
	$aReporteData[0][69] = "OBSERVACIONES"
	$aReporteData[0][70] = "VALOR.FOB US"
	$aReporteData[0][71] = "BOOKING"
	$aReporteData[0][72] = "LINEA ORDEN DE COMPRA"
	;######################################################################
	Return $aReporteData
EndFunc   ;==>_ReporteData



