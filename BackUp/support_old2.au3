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



ConsoleWrite($sInvoiceState_8 & @CRLF)
ConsoleWrite($sInvoiceState_7 & @CRLF)
ConsoleWrite($sInvoiceState_14 & @CRLF)
ConsoleWrite($sInvoiceState_3 & @CRLF)
ConsoleWrite($sInvoiceState_13 & @CRLF)



#include <Array.au3>

Global $a = ReadCSV ("D:\Temp\Test.csv")
_ArrayDisplay ($a)
;MsgBox (0, "", StringReplace ($a[4][3], """", ""))

Func ReadCSV ($p_csv_file)
  Local $file = FileOpen ($p_csv_file)
  If $file = -1 Then  Return SetError (1, 0, 0)

  Local $s = FileRead ($file)
  FileClose ($file)

  Return StringToArray ($s, false, ",")
EndFunc

Func StringToArray ($p_string, $p_transpose_1d = false, $p_delim_col = "|", $p_delim_row = @CRLF, $p_arr_delim_col = Chr (31), $p_arr_delim_row = Chr (30))
  Local $array, $rows, $cols, $last_row, $last_col

  $rows = StringSplit ($p_string, $p_delim_row, 3)
  $last_row = UBound ($rows) - 1

  If $last_row < 1 Then
    ;Array rows = 0
    $cols = StringSplit ($p_string, $p_delim_col, 3)
    $last_col = UBound ($cols) - 1

    If $last_col < 1 Then
      ;Array rows = 0, columns = 0
      Dim $array[1]
      If StringInStr ($p_string, $p_arr_delim_col) or StringInStr ($p_string, $p_arr_delim_row) Then
        $array[0] = StringToArray ($p_string, $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
      Else
        $array[0] = $p_string
      EndIf
    Else
      ;Array rows = 0, columns > 0
      Dim $array[$last_col + 1]
      For $i = 0 To $last_col
        If StringInStr ($cols[$i], $p_arr_delim_col) or StringInStr ($cols[$i], $p_arr_delim_row) Then
          $array[$i] = StringToArray ($cols[$i], $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
        Else
          $array[$i] = $cols[$i]
        EndIf
      Next
    EndIf
    If $p_transpose_1d Then _ArrayTranspose ($array)
  Else
    ;Array rows > 0
    $last_col = 0
    For $i = 0 To $last_row
      StringReplace ($rows[$i], $p_delim_col, "")
      If @extended > $last_col Then $last_col = @extended
    Next

    If $last_col < 1 Then
      ;Array rows > 0, columns = 0
      Dim $array[$last_row + 1]
      For $i = 0 To $last_row
        If StringInStr ($rows[$i], $p_arr_delim_col) or StringInStr ($rows[$i], $p_arr_delim_row) Then
          $array[$i] = StringToArray ($rows[$i], $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
        Else
          $array[$i] = $rows[$i]
        EndIf
      Next
    Else
      ;Array rows > 0, columns > 0
      Dim $array[$last_row + 1][$last_col + 1]
      For $i = 0 To $last_row
        $cols = StringSplit ($rows[$i], $p_delim_col, 3)
        If $last_col > UBound ($cols) - 1 Then ReDim $cols[$last_col + 1]
        For $n = 0 To $last_col
          If StringInStr ($cols[$n], $p_arr_delim_col) or StringInStr ($cols[$n], $p_arr_delim_row) Then
            $array[$i][$n] = StringToArray ($cols[$n], $p_transpose_1d, $p_arr_delim_col, $p_arr_delim_row)
          Else
            $array[$i][$n] = $cols[$n]
          EndIf
        Next
      Next
    EndIf
  EndIf

  If not IsString ($p_string) Then SetError (1)
  If $p_string = "" Then SetError (2)

  Return $array
EndFunc