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

Local $sStartDate = "16/06/2021"
Local $sEndDate = "17/06/2021"
Local $sCosteoDrummondQuery = "SELECT * FROM Repecev2005.dbo.VCosteoDrummond_fact WHERE FECHAFACTURA  BETWEEN  '"&$sStartDate&"' and '"&$sEndDate&"'"
Local $aTRK_Data = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)
_ArrayDisplay($aTRK_Data, '$aTRK_Data')
Local $aInvoices = _ExtractSingleInvoices($aTRK_Data)
;~ _ArrayDisplay($aInvoices, '$aInvoices')


For $i = 1 To UBound($aInvoices) - 1 Step +1 ;Starts at 1. First position indicates invoices amount.
	Local $sInvoiceNumber = $aInvoices[$i]
	Local $aInvoiceCoincidences = _ArrayFindAll($aTRK_Data, $sInvoiceNumber, Default, Default, Default, Default, 29, False)
	_ArrayDisplay($aInvoiceCoincidences,$sInvoiceNumber)
Next

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


