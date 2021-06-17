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
Local $a2D = _ModuloSQL_SQL_SELECT($sCosteoDrummond)
_ArrayDisplay($a2D,'$a2D')

$a2D = dupecheckerthingy($a2D)
_ArrayDisplay($a2D,'$a2D')



Func dupecheckerthingy($showmethearray)
    Local $tmparray[1] = ['']
    For $i = 1 To UBound($showmethearray) - 1
        _ArraySearch($tmparray, $showmethearray[$i][0])
        If @error Then _ArrayAdd($tmparray, $showmethearray[$i][0])
    Next
    $tmparray[0] = 'Number of elements = ' & UBound($tmparray) - 1
    Return $tmparray
EndFunc


_ArrayDisplay($a2D)