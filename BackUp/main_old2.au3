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
$sCosteoDrummond &= "SELECT VCosteoDrummond_fact.FACTURASERVICIOS, "
$sCosteoDrummond &= "VCosteoDrummond_fact.IDDIM, "
$sCosteoDrummond &= "VCosteoDrummond_fact.ACEPTACION, "
$sCosteoDrummond &= "VCosteoDrummond_fact.FACTURA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.PEDIDO, "
$sCosteoDrummond &= "VCosteoDrummond_fact.LINEA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.REFERENCIA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.CANTIDADREF, "
$sCosteoDrummond &= "VCosteoDrummond_fact.PRECIOUNITARIO, "
$sCosteoDrummond &= "VCosteoDrummond_fact.IMPORTADOR, "
$sCosteoDrummond &= "VCosteoDrummond_fact.DECLARANTE, "
$sCosteoDrummond &= "VCosteoDrummond_fact.DO, "
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHASTICKER, "
$sCosteoDrummond &= "VCosteoDrummond_fact.STICKER, "
$sCosteoDrummond &= "VCosteoDrummond_fact.RECIBOPAGOANT, "
$sCosteoDrummond &= "VCosteoDrummond_fact.TIPODECLARACION, "
$sCosteoDrummond &= "VCosteoDrummond_fact.ITEMFLETE, "
$sCosteoDrummond &= "VCosteoDrummond_fact.TASACAMBIO,"
$sCosteoDrummond &= "FLETEDIM,"
$sCosteoDrummond &= "REGISTRO,  "
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHALEVANTE, "
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHAACEPTACIÓN, "
$sCosteoDrummond &= "VCosteoDrummond_fact.FECREA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.PRECIOTOTAL, "
$sCosteoDrummond &= "VCosteoDrummond_fact.TOTALLIQARA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.TOTALLIQIVA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.PAGOTOTAL, "
$sCosteoDrummond &= "VCosteoDrummond_fact.CIF, "
$sCosteoDrummond &= "VCosteoDrummond_fact.PROVEEDOR, "
$sCosteoDrummond &= "VCosteoDrummond_fact.MODALIDAD, "
$sCosteoDrummond &= "VCosteoDrummond_fact.VALORFACTURA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.FECHAFACTURA, "
$sCosteoDrummond &= "VCosteoDrummond_fact.DCTOTRANSPORTE, "
$sCosteoDrummond &= "VCosteoDrummond_fact.ANTICIPO,"
$sCosteoDrummond &= "VALORFOBDIM,"
$sCosteoDrummond &= "CONTENEDORES,"
$sCosteoDrummond &= "INCOTERM"
$sCosteoDrummond &= "FROM Repecev2005.dbo.VCosteoDrummond_fact VCosteoDrummond_fact"
$sCosteoDrummond &= "WHERE FECHAFACTURA  BETWEEN  '22/04/2021' and '24/04/2021'"

ConsoleWrite($sCosteoDrummond&@CRLF)
