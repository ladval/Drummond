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


Local $a2D = _ModuloSQL_SQL_SELECT($sCosteoDrummondQuery)
_ArrayDisplay($a2D, '$a2D')
$a2D = _ExtractSingleInvoices($a2D)
_ArrayDisplay($a2D, '$a2D')

Func _ExtractSingleInvoices($aArray)
	Local $aTempArray[1] = ['']
	For $i = 1 To UBound($aArray) - 1
		_ArraySearch($aTempArray, $aArray[$i][0])
		If @error Then _ArrayAdd($aTempArray, $aArray[$i][0])
	Next
	$aTempArray[0] = UBound($aTempArray) - 1
	Return $aTempArray
EndFunc   ;==>dupecheckerthingy


