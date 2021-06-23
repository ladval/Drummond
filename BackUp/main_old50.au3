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

Local $sSQL_InvoiceNumber = "BQA102372"
Local $sCosteoDrummondQuery = "SELECT DISTINCT * FROM [Repecev2005].[dbo].[VCosteoDrummond_fact] WHERE FACTURASERVICIOS LIKE '" & $sSQL_InvoiceNumber & "'"
Local $aTRK_Data = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)
If UBound|$aTRK_Data) < 2 Then
	ConsoleWrite|'DATOS INSUFICIENTES' & @CRLF)
	Exit
EndIf
_JsonFactData($sSQL_InvoiceNumber)

Exit

Func _JsonFactData($sSQL_InvoiceNumber)
	Local $sJsonFactQuery = "SELECT JsonFact FROM [BotAbc].[dbo].[tfact_ApiProcesos] where InvoiceNumber = '" & $sSQL_InvoiceNumber & "'"
	Local $aAPI_Data = _ModuloSQL_SQL_SELECT($sJsonFactQuery)
	Local $sJsonFact = $aAPI_Data[1][0]
	Local $sJsonFile_InvoiceData = @ScriptDir & '\data\' & $sSQL_InvoiceNumber & '.json'
	_SaveDataToFile($sJsonFile_InvoiceData, $sJsonFact)
	Local $sJsonInvoiceData = _ReadDataFromFile($sJsonFile_InvoiceData)
	Json_Dump($sJsonInvoiceData)
EndFunc   ;==>_JsonFactData

Local $aIDDIM_Data = _2dArray_UniqueElements($aTRK_Data, 0) ;Se define la columna 0 como criterio de elementos repetidos, dado que campo IDDIM es único para cálculo de valor CIF
_ArrayDisplay|$aTRK_Data, '$aTRK_Data')




Local $sLote = "PorDefinir"
Local $sServicio = "PorDefinir"
Local $sInvoice = $aTRK_Data[1][29]
Local $sInvoiceDate = $aTRK_Data[1][31]
Local $sCIFUS = _CIF($aIDDIM_Data, 'USD')
Local $sTCRM = Number($aTRK_Data[1][16])
Local $sCIFCOP = _CIF($aIDDIM_Data, 'COP')

Local $sManejoNaviero = 'MANEJO NAVIERO|1034'
Local $sUsoInstalaciones = 'USO INSTALACIONES|1003'
Local $sBodegaje = 'BODEGAJE|1004'
Local $sVuce = 'VUCE|1016'
Local $sVaciado = 'VACIADO|1008'
Local $sInspeccionDian = 'INSPECCION DIAN|1042'
Local $sDepositoContenedor = 'DEPOSITO CONTENEDOR|1010'
Local $sLiberacionManejoGuia = 'LIBERACION Y MANEJO GUIA|1005'
Local $sDemorasContenedor = 'DEMORAS CONTENEDOR|1011'
Local $sMovilizacionInspeccion = 'MOVILIZACION INSPECCION|1043'
Local $sArancel = 'ARANCEL|1001'
Local $sIVA = 'IVA|1002'
Local $sMaquinaInterna = 'MAQUINA INTERNA|1061'
Local $sVistosBuenos = 'VISTOS BUENOS|1023'
Local $sCargueDescargue = 'CARGUE Y DESCARGUE|1025'
Local $sPorteo = 'PORTEO|1024'
Local $sInspeccionReconocimiento = 'INSPECCION RECONOCIMIENTO|1046'
Local $sTraslado = 'TRASLADO|1018'
Local $sStacker = 'STACKER|1045'
Local $sDanos = 'DAÑOS|1050'
Local $sSuciedad = 'SUCIEDAD|1026'
Local $sEmisionBL = 'EMISION BL|1028'
Local $sAperturaCierre = 'APERTURA Y CIERRE|1015'
Local $sCertificados = 'CERTIFICADOS|1020'
Local $sBasculas = 'BASCULAS|1029'
Local $sPapeleria = 'PAPELERIA|1038'
Local $sCambioModalidad = 'CAMBIO MODALIDAD|1048'
Local $sHoraAdicional = 'HORA ADICIONAL|1052'
Local $sCargoFijo = 'CARGO FIJO|1054'
Local $sManejoDocumentos = 'MANEJO DE DOCUMENTOS|1006'
Local $sCamaAltaBaja = 'CAMA ALTA-BAJA|1044'
Local $sServicioCarpado = 'SERVICIO DE CARPADO|1060'
Local $sServicioConsolidacion = 'SERVICIOCONSOLIDACION|1007'
Local $sComodatos = 'COMODATOS|1039'
Local $sTrabajosVariosHorasHombres = 'TRABAJOS VARIOS HORAS/HOMBRES|1049'
Local $sSellosContenedor = 'SELLOS DE CONTENEDOR|1051'
Local $sEnvio = 'ENVIO|1031'
Local $sDismounting = 'DISMOUNTING|1012'
Local $sTransporte = 'TRANSPORTE|1027'
Local $sReempaque = 'ADICIONAL 1|1022'
Local $sAdicional2 = 'ADICIONAL 2|1032'
Local $sAdicional3 = 'ADICIONAL 3|1041'
Local $sAdicional4 = 'ADICIONAL 4|1009'
Local $sServicioExtraordinario = 'SERVICIO EXTRAORDINARIO|1014'
Local $sVr4xmil = 'VR 4XMIL|1030'
Local $sSubtotal = 'SUBTOTAL'
Local $sServicioAduana = 'SERVICIO ADUANA|2003'
Local $sReconocimientoMcia = 'RECONOCIMIENTO MCIA|2052'
Local $sElaboracionRegistrosImp = 'ELABORACION REGISTROS DE IMP.|2005'
Local $sElaboracionDeclaracionesImportacion = 'ELABORACIONDECLARACIONES DE IMPORTACION|2009'
Local $sElaboracionDeclaracionesValor = 'ELABORACION DECLARACIONES DE VALOR|2042'
Local $sDescargueDirecto = 'DESCARGUE DIRECTO|2007'
Local $sVistoBueno = 'VISTO BUENO|2045'

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
	Local $iUniqueElements =(UBound($aArrayUnique) - 1)
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
		$iCIF +=(Number($aIDDIM_Data[$i][26]) * $iCurrency)
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


