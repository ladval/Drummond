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



    ConsoleWrite($sInvoiceState_8&@CRLF)
    ConsoleWrite($sInvoiceState_7&@CRLF)
    ConsoleWrite($sInvoiceState_14&@CRLF)
    ConsoleWrite($sInvoiceState_3&@CRLF)
    ConsoleWrite($sInvoiceState_13&@CRLF)