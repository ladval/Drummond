#include-once
#include <FileConstants.au3>

Func _SaveDataToFile($sFile, $sData)
	Local $hFile = FileOpen($sFile, $FO_OVERWRITE + $FO_CREATEPATH)
	FileWrite($hFile, $sData)
	FileClose($hFile)
	Return $sFile
EndFunc   ;==>_SaveDataToFile

Func _ReadDataFromFile($sFile)
	Local $hFile = FileOpen($sFile, $FO_READ)
	Local $sData = FileRead($hFile)
	FileClose($hFile)
	Return $sData
EndFunc   ;==>_ReadDataFromFile



;~ Func _ExtractSingle2DData($aArray)
;~ 	Local $aTempArray[1] = ['']
;~ 	For $i = 1 To UBound($aArray) - 1
;~ 		Local $sInvoiceNumber = $aArray[$i][29] ;Col 29: Where SIIGO Invoice number is lkocated in the array
;~ 		_ArraySearch($aTempArray, $sInvoiceNumber)
;~ 		If @error Then _ArrayAdd($aTempArray, $sInvoiceNumber)
;~ 	Next
;~ 	$aTempArray[0] = UBound($aTempArray) - 1
;~ 	Return $aTempArray
;~ EndFunc   ;==>_ExtractSingleInvoices
