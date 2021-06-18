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

Local $sCosteoDrummondQuery = "SELECT * FROM Repecev2005.dbo.VCosteoDrummond_fact WHERE FECHAFACTURA  BETWEEN  '16/06/2021' and '17/06/2021'"
Local $aTRK_Data = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)
_ArrayDisplay($aTRK_Data, '$aTRK_Data')
Local $aInvoices = _ExtractSingleInvoices($aTRK_Data)


For $i = 0 To Ubound($aInvoices)-1 Step +1
	$aInvoices[$i]
Next

_ArrayDisplay($aInvoices, '$aTRK_Data')

Func _ExtractSingleInvoices($aArray)
	Local $aTempArray[1] = ['']
	For $i = 1 To UBound($aArray) - 1
		Local $sInvoiceNumber = $aArray[$i][29]
		_ArraySearch($aTempArray, $sInvoiceNumber)
		If @error Then _ArrayAdd($aTempArray, $sInvoiceNumber)
	Next
	$aTempArray[0] = UBound($aTempArray) - 1
	Return $aTempArray
EndFunc   ;==>_ExtractSingleInvoices


