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
Local $sJsonFactQuery = "SELECT JsonFact FROM [BotAbc].[dbo].[tfact_ApiProcesos] where InvoiceNumber = '" & $sSQL_InvoiceNumber & "'"
Local $aTRK_Data = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)

If UBound($aTRK_Data) < 2 Then
	ConsoleWrite('DATOS INSUFICIENTES' & @CRLF)
	Exit
EndIf
Local $aIDDIM_Data = _2dArray_UniqueElements($aTRK_Data, 0) ;Se define la columna 0 como criterio de elementos repetidos, dado que campo IDDIM es único para cálculo de valor CIF

Func _CIF($aIDDIM_Data, $sCurrency)
Select
Case $sCurrency = 'USD'
Case $sCurrency = 'COP'
EndSelect
EndFunc   ;==>_CIF


Local $sLote = "PorDefinir"
Local $sServicio = "PorDefinir"
Local $sInvoice = $aTRK_Data[1][29]
Local $sInvoiceDate = $aTRK_Data[1][31]
Local $sCIFUS = _CIF($aIDDIM_Data, 'USD')
Local $sTCRM = Number($aTRK_Data[1][31])
Local $sCIFCOP = _CIF($aIDDIM_Data, 'COP')
Local $sManejoNaviero
Local $sUsoInstalaciones
Local $sBodegaje
Local $sVuce
Local $sVaciado
Local $sInspeccionDian
Local $sDepositoContenedor
Local $sLiberacionManejoGuia
Local $sDemorasContenedor
Local $sMovilizacionInspeccion
Local $sArancel
Local $sIVA
Local $sMaquinaInterna
Local $sVistosBuenos
Local $sCargueDescargue
Local $sPorteo
Local $sInspeccionReconocimiento
Local $sTraslado
Local $sStacker
Local $sDanos
Local $sSuciedad
Local $sEmisionBL
Local $sAperturaCierre
Local $sCertificados
Local $sBasculas
Local $sPapeleria
Local $sCambioModalidad
Local $sHoraAdicional
Local $sCargoFijo
Local $sManejoDocumentos
Local $sCamaAltaBaja
Local $sServicioCarpado
Local $sServicioConsolidacion
Local $sComodatos
Local $sTrabajosVariosHorasHombres
Local $sSellosContenedor
Local $sEnvio
Local $sDismounting
Local $sTransporte
Local $sReempaque
Local $sAdicional2
Local $sAdicional3
Local $sAdicional4
Local $sServicioExtraordinario
Local $sVr4xmil
Local $sSubtotal
Local $sServicioAduana
Local $sReconocimientoMcia
Local $sElaboracionRegistrosImp
Local $sElaboracionDeclaracionesImportacion
Local $sElaboracionDeclaracionesValor
Local $sDescargueDirecto
Local $sVistoBueno
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




;~ Local $aAPI_Data = _ModuloSQL_SQL_SELECT($sJsonFactQuery)
;~ Local $sJsonFact = $aAPI_Data[1][0]
;~ Local $sJsonFile_InvoiceData = @ScriptDir & '\data\' & $sSQL_InvoiceNumber & '.json'

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


Exit

;~ _SaveDataToFile($sJsonFile_InvoiceData, $sJsonFact)
;~ Local $sJsonInvoiceData = _ReadDataFromFile($sJsonFile_InvoiceData)
;~ Json_Dump($sJsonInvoiceData)
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


