#include <File.au3>
#include <Date.au3>
#include <Array.au3>
#include <String.au3>
#include "modulo_json.au3"
#include "modulo_sql.au3"
#include "modulo_misc.au3"


Local $sJsonFileSettings = @ScriptDir & "\test.json"
Local $sJsonSettings = _ReadDataFromFile($sJsonFileSettings)
Local $oJsonSettings = Json_Decode($sJsonSettings)
Local $sInvoiceState_8 = Json_Get($oJsonSettings, '.AdditionalProperty[8].Value')
Local $sInvoiceState_7 = Json_Get($oJsonSettings, '.AdditionalProperty[7].Value')
Local $sInvoiceState_14 = Json_Get($oJsonSettings, '.AdditionalProperty[14].Value')
Local $sInvoiceState_3 = Json_Get($oJsonSettings, '.AdditionalProperty[3].Value')
Local $sInvoiceState_13 = Json_Get($oJsonSettings, '.AdditionalProperty[13].Value')


Local $a = ReadCSV(@ScriptDir & "\Libro2.csv")



For $i = 0 To UBound($a) - 1 Step +1
	Local $sInvoiceNumber = $a[$i][0]
	Local $sDO = $a[$i][1]
	Local $sPedido = $a[$i][2]
	Local $sJefeCuenta = $a[$i][3]
	Local $sProveedor = $a[$i][4]
	Local $sTMR = $a[$i][5]
	Local $sSQLInvoiceDataQuery = "SELECT JsonFact FROM [BotAbc].[dbo].[tfact_ApiProcesos] WHERE InvoiceNumber = '" & $sInvoiceNumber & "' AND Empresa = '860536003'"
	Local $aSQLInvoiceDataQuery = _ModuloSQL_SQL_SELECT($sSQLInvoiceDataQuery)
	Local $sJsonFactData = $aSQLInvoiceDataQuery[1][0]
	Local $oJsonFactData = Json_Decode($sJsonFactData)
	Local $sInvoiceState_8 = Json_Put($oJsonFactData, '.AdditionalProperty[8].Value', $sDO)
	Local $sInvoiceState_7 = Json_Put($oJsonFactData, '.AdditionalProperty[7].Value', $sPedido)
	Local $sInvoiceState_14 = Json_Put($oJsonFactData, '.AdditionalProperty[14].Value', $sJefeCuenta)
	Local $sInvoiceState_3 = Json_Put($oJsonFactData, '.AdditionalProperty[3].Value', $sProveedor)
	Local $sInvoiceState_13 = Json_Put($oJsonFactData, '.AdditionalProperty[13].Value', $sTMR)
	Local $sStringJson = Json_Encode($oJsonFactData)
	Local $sUpdateQuery = "UPDATE [BotAbc].[dbo].[tfact_ApiProcesos] SET  Status = 'SinProcesar', JsonFact = '" & $sStringJson & "' WHERE InvoiceNumber = '" & $sInvoiceNumber & "'"
	Local $aSQLInvoiceDataQuery = _ModuloSQL_SQL_EXEC($sUpdateQuery)
	ConsoleWrite($sInvoiceNumber & ': OK' & @CRLF)
Next



_ArrayDisplay($a)
;MsgBox (0, "", StringReplace ($a[4][3], """", ""))

Func ReadCSV($p_csv_file)
	Local $file = FileOpen($p_csv_file)
	If $file = -1 Then Return SetError(1, 0, 0)

	Local $s = FileRead($file)
	FileClose($file)

	Return StringToArray($s, False, ",")
EndFunc   ;==>ReadCSV

Func StringToArray($p_string, $p_transpose_1d = False, $p_delim_col = "|", $p_delim_row = @CRLF, $p_arr_delim_col = Chr(31), $p_arr_delim_row = Chr(30))
	Local $array, $rows, $cols, $last_row, $last_col

	$rows = StringSplit($p_string, $p_delim_row, 3)
	$last_row = UBound($rows) - 1

	If $last_row < 1 Then
		;Array rows = 0
		$cols = StringSplit($p_string, $p_delim_col, 3)
		$last_col = UBound($cols) - 1

		If $last_col < 1 Then
			;Array rows = 0, columns = 0
			Dim $array[1]
			If StringInStr($p_string, $p_arr_delim_col) Or StringInStr($p_string, $p_arr_delim_row) Then
				$array[0] = StringToArray($p_string, $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
			Else
				$array[0] = $p_string
			EndIf
		Else
			;Array rows = 0, columns > 0
			Dim $array[$last_col + 1]
			For $i = 0 To $last_col
				If StringInStr($cols[$i], $p_arr_delim_col) Or StringInStr($cols[$i], $p_arr_delim_row) Then
					$array[$i] = StringToArray($cols[$i], $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
				Else
					$array[$i] = $cols[$i]
				EndIf
			Next
		EndIf
		If $p_transpose_1d Then _ArrayTranspose($array)
	Else
		;Array rows > 0
		$last_col = 0
		For $i = 0 To $last_row
			StringReplace($rows[$i], $p_delim_col, "")
			If @extended > $last_col Then $last_col = @extended
		Next

		If $last_col < 1 Then
			;Array rows > 0, columns = 0
			Dim $array[$last_row + 1]
			For $i = 0 To $last_row
				If StringInStr($rows[$i], $p_arr_delim_col) Or StringInStr($rows[$i], $p_arr_delim_row) Then
					$array[$i] = StringToArray($rows[$i], $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
				Else
					$array[$i] = $rows[$i]
				EndIf
			Next
		Else
			;Array rows > 0, columns > 0
			Dim $array[$last_row + 1][$last_col + 1]
			For $i = 0 To $last_row
				$cols = StringSplit($rows[$i], $p_delim_col, 3)
				If $last_col > UBound($cols) - 1 Then ReDim $cols[$last_col + 1]
				For $n = 0 To $last_col
					If StringInStr($cols[$n], $p_arr_delim_col) Or StringInStr($cols[$n], $p_arr_delim_row) Then
						$array[$i][$n] = StringToArray($cols[$n], $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
					Else
						$array[$i][$n] = $cols[$n]
					EndIf
				Next
			Next
		EndIf
	EndIf

	If Not IsString($p_string) Then SetError(1)
	If $p_string = "" Then SetError(2)
	$array = _Array2DDeleteEmptyRows($array)
	Return $array
EndFunc   ;==>StringToArray



Func _Array2DDeleteEmptyRows(ByRef $iArray)
	Local $vEmpty = False
	$cols = UBound($iArray, 2)
	Local $iArrayOut[1][$cols], $a = 0
	For $i = 0 To (UBound($iArray, 1) - 1) Step 1
		If $vEmpty = True Then $a += 1
		$vEmpty = False
		For $j = 0 To (UBound($iArray, 2) - 1) Step 1
			If StringCompare($iArray[$i][0], "", 0) <> 0 Then
				ReDim $iArrayOut[$a + 1][$cols]
				$iArrayOut[$a][$j] = $iArray[$i][$j]
				$vEmpty = True
			EndIf
		Next
	Next
	Return $iArrayOut
EndFunc   ;==>_Array2DDeleteEmptyRows
