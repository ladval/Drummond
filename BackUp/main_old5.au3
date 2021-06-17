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

Local $sCosteoDrummond
$sCosteoDrummond &= "SELECT VCosteoDrummond_fact.FACTURASERVICIOS, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.IDDIM, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.ACEPTACION, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.FACTURA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.PEDIDO, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.LINEA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.REFERENCIA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.CANTIDADREF, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.PRECIOUNITARIO, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.IMPORTADOR, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.DECLARANTE, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.DO, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHASTICKER, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.STICKER, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.RECIBOPAGOANT, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.TIPODECLARACION, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.ITEMFLETE, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.TASACAMBIO," & @CRLF
$sCosteoDrummond &= "FLETEDIM," & @CRLF
$sCosteoDrummond &= "REGISTRO,  " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHALEVANTE, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHAACEPTACIÓN, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.FECREA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.PRECIOTOTAL, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.TOTALLIQARA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.TOTALLIQIVA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.PAGOTOTAL, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.CIF, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.PROVEEDOR, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.MODALIDAD, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.VALORFACTURA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHAFACTURA, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.DCTOTRANSPORTE, " & @CRLF
$sCosteoDrummond &= "VCosteoDrummond_fact.ANTICIPO," & @CRLF
$sCosteoDrummond &= "VALORFOBDIM," & @CRLF
$sCosteoDrummond &= "CONTENEDORES," & @CRLF
$sCosteoDrummond &= "INCOTERM" & @CRLF
$sCosteoDrummond &= "FROM Repecev2005.dbo.VCosteoDrummond_fact" & @CRLF
$sCosteoDrummond &= "WHERE FECHAFACTURA  BETWEEN  '01/06/2021' and '10/06/2021'"
Local $aCosteoDrummond = _ModuloSQL_SQL_SELECT($sCosteoDrummond)
_ArrayDisplay($aCosteoDrummond, '$aCosteoDrummond')






Local $aMult = _ArrayNoSingles($aCosteoDrummond)

Func _ArrayNoSingles($aArray)
    ; Create dictionary
    Local $oDict = ObjCreate("Scripting.Dictionary")
    For $i = 0 To UBound($aArray) - 1
        $vElem = $aArray[$i][0]
        $vUnique = StringSplit($vElem, "|")[3] ; Get the unique element at the end of the string
        ; Check if already in dictionary
        If $oDict.Exists($vUnique) Then
            ; Add whole element to existing list
            $oDict($vUnique) = $oDict($vUnique) & "#" & $vElem ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        Else
            ; Add whole element to dictionary
            $oDict.Item($vUnique) = $vElem
        EndIf
    Next
    ; Create return array large enough for all elements
    Local $aRet[UBound($aArray) + 1], $iIndex = 0
    ; Loop through dictionary and transfer multiple elements to return array
    For $vKey in $oDict
        $vValue = $oDict($vKey)
        ; Check if more than one delimiter
        StringReplace($vValue, "#", "&")  ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        If @extended Then
            $aSplit = StringSplit($vValue, "#")  ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For $i = 1 To $aSplit[0]
                $iIndex += 1
                $aRet[$iIndex] = $aSplit[$i]
            Next
        EndIf
    Next
    ; Add count to [0]element
    $aRet[0] = $iIndex
    ; Remove ampty elements of return array
    ReDim $aRet[$iIndex + 1]
    Return $aRet
EndFunc