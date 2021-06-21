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

