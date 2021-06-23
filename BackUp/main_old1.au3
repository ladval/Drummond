#cs ----------------------------------------------------------------------------
 AutoIt Version: 3.3.14.5
 Author:         Jesús Antonio Ladino Valbuena
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

;~ Local $sStartDate = "16/06/2021"
;~ Local $sEndDate = "17/06/2021"
;~ Local $sCosteoDrummondQuery = "SELECT * FROM Repecev2005.dbo.VCosteoDrummond_fact WHERE FECHAFACTURA  BETWEEN  '" & $sStartDate & "' and '" & $sEndDate & "' ORDER BY FACTURASERVICIOS"

;~ Local $sSQL_InvoiceNumber = "BQA102372"
;~ Local $sSQL_InvoiceNumber = "SMR18947"
Local $sSQL_InvoiceNumber = "SMR18921"
Local $sCosteoDrummondQuery = "SELECT * FROM [Repecev2005].[dbo].[VCosteoDrummond_fact] WHERE FACTURASERVICIOS LIKE '" & $sSQL_InvoiceNumber & "'"
Local $aTRK_Data = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)
If UBound($aTRK_Data) < 2 Then
	ConsoleWrite('DATOS INSUFICIENTES' & @CRLF)
	Exit
EndIf
_JsonFactData($sSQL_InvoiceNumber)
;~ _JsonCodedData()
Exit

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
	Local $aItemReferences[@extended][3]
	Local $oJsonInvoiceData = Json_Decode($sJsonInvoiceData)
	For $i = 0 To $iItemsNumber Step +1
		Local $iItemReferenceCode = Json_Get($oJsonInvoiceData, ".ItemInformation[" & $i & "].ItemReference")
		Local $iItemReferenceName = Json_Get($oJsonInvoiceData, ".ItemInformation[" & $i & "].Name")
		Local $aIndexNoConcepto = _ArraySearch($aDrummondConceptos, $iItemReferenceCode, Default, Default, Default, Default, Default, 1, False)
		$aItemReferences[$i][0] = $iItemReferenceCode
		$aItemReferences[$i][1] = $iItemReferenceName
		$aItemReferences[$i][2] = $aDrummondConceptos[$aIndexNoConcepto][2]
	Next
	Local $sCheckItems = _CheckItems($aItemReferences)
	If $sCheckItems Then
_ArrayDisplay($aItemReferences,'Test')
	Else
		ConsoleWrite($sCheckItems & @CRLF)
	EndIf

;~ Json_Dump($sJsonInvoiceData)
EndFunc   ;==>_JsonFactData


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
EndFunc   ;==>_CheckItems

Local $aIDDIM_Data = _2dArray_UniqueElements($aTRK_Data, 0) ;Se define la columna 0 como criterio de elementos repetidos, dado que campo IDDIM es único para cálculo de valor CIF

Local $sLote = "PorDefinir"
Local $sServicio = "PorDefinir"
Local $sInvoice = $aTRK_Data[1][29]
Local $sInvoiceDate = $aTRK_Data[1][31]
Local $sCIFUS = _CIF($aIDDIM_Data, 'USD')
Local $sTCRM = Number($aTRK_Data[1][16])
Local $sCIFCOP = _CIF($aIDDIM_Data, 'COP')

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
	_ArrayDisplay($aJsonReportData, '$aJsonReportData')
EndFunc   ;==>_JsonCodedData



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
EndFunc   ;==>_ExtractSingleInvoices


