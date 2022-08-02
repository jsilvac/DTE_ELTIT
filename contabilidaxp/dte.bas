Attribute VB_Name = "DTE"
Public fechacertificacion As String
Public resolucion As String
Public email_cuenta_usuario As String
Public email_cuenta_clave As String
Public email_cuenta_server As String

Public librerias(20) As String
Public lib_generapdf As String
Public lib_base As String
Public lib_java As String
Public lib_enviarsii As String
Public lib_recepcionar As String
Public lib_Firmaenvio As String
Public lib_FirmaLibro As String
Public lib_imprime As String

Public clavecertificado As String
Public confi_activacentinelas As String
Public confi_archivopdfnormal As String
Public confi_archivopdfcedible As String
Public confi_servidor As String
Public confi_empresaactiva As String
Public confi_localempresa As String
Public confi_rutafae As String
Public confi_rutadte As String
Public confi_servermail As String
Public confi_mailsalida As String
Public confi_clavemail As String
Public confi_rutapdf As String
Public confi_java As String
Public dte_tipodte As String
Public tipoDoc(10) As Double
Public snum As Double
Public dte_rec(50, 10) As String
Public ncref As Boolean
Public dte_referencia As String
Public dte_caso As String
Public dte_tipo As String
Public dte_tiporef As String
Public dte_folio As String
Public dte_fecha As String
Public dte_indtraslado As String
Public dte_servicio As String
Public dte_e_rut As String
Public dte_e_nombre As String
Public dte_e_acti As String
Public dte_e_direccion As String
Public dte_e_giro As String
Public dte_e_comuna As String
Public dte_e_ciudad As String
Public dte_e_sucursal As String
Public dte_e_vendedor As String
Public dte_r_rut As String
Public dte_r_nombre As String
Public dte_r_giro As String
Public dte_r_direccion As String
Public dte_r_comuna As String
Public dte_r_ciudad As String
Public matrixdte(100, 10) As String
Public totallineas As Double
Public certificado  As String
Public caf As String
Public xml As New ChilkatXml
Public dte_rutenvia As String
Public dte_neto As Double
Public dte_iva As Double
Public dte_exento As Double
Public dte_total As Double
Public dte_ref As Double
Public dte_vin As Double
Public dte_lic As Double
Public dte_cer As Double
Public dte_car As Double
Public dte_har As Double
Public dte_descuento As Double
Public descuento1 As Double
Public descuento2 As Double



Public Sub GENERADTE(loc, caja, tipo, numero, fecha, rut, neto, iva, total, ref, vin, lic, iha, ica)

Dim dato(100) As String
Dim dato2(100) As String
Dim datos(100, 10) As String
Dim dato3(100) As String
Dim comi As String
Dim cabeza As String
Dim DETALLE As String
Dim K As Integer
Dim pasada As String
Dim i As Integer
Dim FOLIOS As String
Dim salida As String
Dim entrada As String
Dim FIRMA As String
Dim entradapdf As String
Dim salidapdf As String
Dim pdf As String
Dim enviasii As String
Dim firmaenvio As String
Dim entradafirma As String
Dim salidafirma As String
Dim entradaenvio As String
Dim SALIDAENVIO As String
Dim cer As String
Dim lin1 As Double
Dim cadena As String

xml.Encoding = "ISO-8859-1"


comi = Chr(34)

If tipo = "FV" Then dte_tipodte = "33": Rem FACTURA ELECTRONICA
If tipo = "FE" Then dte_tipodte = "34": Rem FACTURA NO AFECTA
If tipo = "FC" Then dte_tipodte = "46": Rem FACTURA DE COMPRA ELECTRONICA
If tipo = "GD" Then dte_tipodte = "52": Rem GUIA DE DESPACHO ELECTRONICA:
If tipo = "ND" Then dte_tipodte = "56": Rem NOTA DEBITO ELECTRONICA
If tipo = "NF" Or tipo = "NB" Then dte_tipodte = "61": Rem NOTA DE CREDITO ELECTRONICA
dte_fecha = Format(fecha, "dd-mm-yyyy")
empresa

Call empresadte(codigocontable)
Call clientedte(rut)
dte_folio = leerfoliodte(confi_empresaactiva, dte_tipodte)
 Rem dte_folio = 1503

    If leerfolioautorizado(confi_empresaactiva, dte_tipodte, dte_folio) = False Then
    Unload electro04
    Exit Sub
    End If
    
    For K = 1 To 100
    dato(K) = ""
    Next K

    dato(1) = "<DTE version=" + comi + "1.0" + comi + " >"
    dato(2) = "<Documento ID=" + comi + "ENVIOFOLIO_" & dte_folio & "T" + dte_tipodte + comi + ">"
    dato(3) = "<Encabezado>"

    dte_fecha = Format(dte_fecha, "yyyy-mm-dd")
    dato(4) = "<IdDoc>"
    dato(5) = "<TipoDTE>" + dte_tipodte + "</TipoDTE>"
    dato(6) = "<Folio>" & dte_folio & "</Folio>"
    dato(7) = "<FchEmis>" + dte_fecha + "</FchEmis>"
    Rem indicador no rebaja <IndNoRebaja>"
    Rem tipodespacho <TipoDespacho>"
    If dte_tipodte = "52" Then
    dato(8) = "<Indtraslado>" & dte_indtraslado & "</Indtraslado>"
    End If
Rem 1 = venta
Rem 2  = venta por efectuar
Rem 3 = consignaciones
Rem 4 = entrega gratuita
Rem 5 = traslados internos
Rem 6 = otros traslado no venta

Rem tipode impresion <TpoImpresion> n = normal t=ticket

Rem  "<FmaPago>" & PAGO & "</FmaPago>" solo servicios

Rem dato(10) = "<FchCancel>" + Format(fechaemision, "yyyy-mm-dd") + "</FchCancel>"
Rem dato(11) = "<MedioPago>EF</MedioPago>"
Rem dato(12) = "<FchVenc>" + Format(fechaemision, "yyyy-mm-dd") + "</FchVenc>"
    dato(10) = ""
    dato(11) = ""
    dato(12) = ""
    dato(13) = "</IdDoc>"
    dato(14) = "<Emisor>"
    dato(15) = "<RUTEmisor>" + dte_e_rut + "</RUTEmisor>"
    dato(16) = "<RznSoc>" + dte_e_nombre + "</RznSoc>"
    dato(17) = "<GiroEmis>" + Mid(dte_e_giro, 1, 80) + "</GiroEmis>"
    dato(18) = "<Acteco>" + dte_e_acti + "</Acteco>"
    dato(19) = "<DirOrigen>" + Mid(dte_e_direccion, 1, 60) + "</DirOrigen>"
    dato(20) = "<CmnaOrigen>" + dte_e_comuna + "</CmnaOrigen>"
    dato(21) = "<CiudadOrigen>" + dte_e_ciudad + "</CiudadOrigen>"
    If dte_e_sucursal <> "" Then
    dato(22) = "<cdgSIISucur>" + dte_e_sucursal + "</CdgSIISucur>"
    End If
    If dte_e_vendedor <> "" Then
    dato(23) = "<cdgVendedor>" + dte_e_vendedor + "</cdgVendedor>"
    End If

    dato(24) = "</Emisor>"
    dato(25) = "<Receptor>"
    dato(26) = "<RUTRecep>" + dte_r_rut + "</RUTRecep>"
    dato(27) = "<RznSocRecep>" + dte_r_nombre + "</RznSocRecep>"
    dato(28) = "<GiroRecep>" + Mid(dte_r_giro, 1, 40) + "</GiroRecep>"
    dato(29) = "<DirRecep>" + Mid(dte_r_direccion, 1, 70) + "</DirRecep>"
    dato(30) = "<CmnaRecep>" + Mid(dte_r_comuna, 1, 20) + "</CmnaRecep>"
    dato(31) = "<CiudadRecep>" + Mid(dte_r_ciudad, 1, 20) + "</CiudadRecep>"
    dato(33) = "</Receptor>"
    ncref = False

    Rem If dte_tipodte = 61 And dte_folio = 4 Then ncref = True
    
    Call leerdetalledte(tipo, numero, caja, Format(fecha, "yyyy-mm-dd"), total)


    dato(42) = "<Totales>"
    dato(43) = "<MntNeto>" & dte_neto & "</MntNeto>"
    If dte_exento <> 0 Then
    dato(44) = "<MntExe>" & dte_exento & "</Mntexe>"
    End If


    dato(45) = "<TasaIVA>19</TasaIVA>"
    dato(46) = "<IVA>" & dte_iva & "</IVA>"
    Rem impuesto

If dte_car + dte_har + dte_ref + dte_vin + dte_lic + dte_cer <> 0 Then
    If dte_car <> 0 Then
    dato(47) = "<ImptoReten>"
    dato(48) = "<TipoImp>18</tipoImp"
    dato(49) = "<TasaImp>5</TasaImp>"
    dato(50) = "<MontoImp>" & dte_car & "</MontoImp>"
    dato(51) = "</ImptoReten>"
    End If
    If dte_har <> 0 Then
    dato(52) = "<ImptoReten>"
    dato(53) = "<TipoImp>19</tipoImp"
    dato(54) = "<TasaImp>12</TasaImp>"
    dato(55) = "<MontoImp>" & dte_har & "</MontoImp>"
    dato(56) = "</ImptoReten>"
    End If
    If dte_lic <> 0 Then
    dato(57) = "<ImptoReten>"
    dato(58) = "<TipoImp>24</tipoImp"
    dato(59) = "<TasaImp>27</TasaImp>"
    dato(60) = "<MontoImp>" & dte_lic & "</MontoImp>"
    dato(61) = "</ImptoReten>"
    End If
    If dte_vin <> 0 Then
    dato(62) = "<ImptoReten>"
    dato(63) = "<TipoImp>25</tipoImp"
    dato(64) = "<TasaImp>15</TasaImp>"
    dato(65) = "<MontoImp>" & dte_vin & "</MontoImp>"
    dato(66) = "</ImptoReten>"
    End If
    cer = "0"
    If dte_cer <> 0 Then
    dato(67) = "<ImptoReten>"
    dato(68) = "<TipoImp>26</tipoImp"
    dato(69) = "<TasaImp>15</TasaImp>"
    dato(70) = "<MontoImp>" & dte_cer & "</MontoImp>"
    dato(71) = "</ImptoReten>"
    End If
    If dte_ref <> 0 Then
    dato(72) = "<ImptoReten>"
    dato(73) = "<TipoImp>27</tipoImp"
    dato(74) = "<TasaImp>13</TasaImp>"
    dato(75) = "<MontoImp>" & dte_ref & "</MontoImp>"
    dato(76) = "</ImptoReten>"
    End If
End If

dato(77) = "<MntTotal>" & dte_total & "</MntTotal>"
dato(78) = "</Totales>"
dato(79) = "</Encabezado>"

For K = 1 To 100
If dato(K) <> "" Then
cabeza = cabeza + dato(K)
End If
Next K
'datos(1, 1) = "1"
'datos(1, 2) = "INT1"
'datos(1, 3) = "011"
'datos(1, 4) = "Parlantes Multimedia 180 w"
'datos(1, 5) = "0"
'datos(1, 6) = "20"
'datos(1, 7) = "4500"
'datos(1, 8) = "90000"
DETALLE = ""
For K = 1 To 40
dato2(K) = ""
dato3(K) = ""

Next K

If ncref = False Then
    For K = 1 To totallineas
    dato2(1) = "<Detalle>"
    dato2(2) = "<NroLinDet>" & K & "</NroLinDet>"
    dato2(3) = "<CdgItem>"
    dato2(4) = "<TpoCodigo>" + "EAN13" + "</TpoCodigo>"
    dato2(5) = "<VlrCodigo>" + matrixdte(K, 1) + "</VlrCodigo>"
    dato2(6) = "</CdgItem>"
        If matrixdte(K, 0) = "1" Then
        dato2(7) = "<IndExe>1</IndExe>"
        End If
        dato2(8) = "<NmbItem>" + Replace(matrixdte(K, 3), "+CHR(34)+", "&quot") + "</NmbItem>"
        If matrixdte(K, 2) <> "0" Then
        dato2(9) = "<QtyItem>" + Replace(matrixdte(K, 2), ",", ".") + "</QtyItem>"
        dato2(10) = "<PrcItem>" + Replace(matrixdte(K, 4), ",", ".") + "</PrcItem>"
        End If
        
        If Val(matrixdte(K, 5)) <> 0 And Val(matrixdte(K, 8)) <> 0 Then
        dato2(11) = "<DescuentoPct>" + matrixdte(K, 5) + "</Descuentopct>"
        dato2(12) = "<DescuentoMonto>" & Replace(matrixdte(K, 8), ",", ".") & "</DescuentoMonto>"
        dato2(13) = "<SubDscto>"
        dato2(14) = "<TipoDscto>" + "%" + "</TipoDscto>"
        dato2(15) = "<ValorDscto>" & Replace(matrixdte(K, 5), ",", ".") & "</ValorDscto>"
        dato2(16) = "</SubDscto>"
Else
        dato2(11) = ""
        dato2(12) = ""
        dato2(13) = ""
        dato2(14) = ""
        dato2(15) = ""
        dato2(16) = ""

        End If
        
        If matrixdte(K, 7) <> "" Then
        dato2(17) = "<CodImpAdic>" + matrixdte(K, 7) + "</CodImpAdic>"
        Else
        dato2(17) = ""
        End If
  
    dato2(18) = "<MontoItem>" + matrixdte(K, 6) + "</MontoItem>"
       
       dato2(19) = "</Detalle>"


For i = 1 To 20
    If dato2(i) <> "" Then
    pasada = pasada + dato2(i)
    End If
Next i
Next K
End If
DETALLE = DETALLE & pasada
lin1 = 0

    If descuento1 + descuento2 <> 0 Then
        If descuento1 <> 0 Then
        lin1 = lin1 + 1
        dato3(2) = "<DscRcgGlobal>"
        dato3(3) = "<NroLinDR>" & lin1 & "</NroLinDR>"
        dato3(4) = "<TpoMov>D</TpoMov>"
        dato3(5) = "<TpoValor>%</tpoValor>"
        dato3(6) = "<ValorDR>" & dte_descuento & "</ValorDR>"
        dato3(7) = "</DscRcgGlobal>"
    End If

'    If descuento2 <> 0 Then
'        lin1 = lin1 + 1
'        dato3(8) = "<DscRcgGlobal>"
'        dato3(9) = "<NroLinDR>" & lin1 & "</NroLinDR>"
'        dato3(10) = "<TpoMov>D</TpoMov>"
'        dato3(11) = "<TpoValor>%</tpoValor>"
'        dato3(12) = "<ValorDR>" & dte_descuento & "</ValorDR>"
'        dato3(13) = "<IndExeDR>" & "1" & "</IndExeDR>"
'        dato3(14) = "</DscRcgGlobal>"
'    End If
    End If

Rem If dte_tipodte = 33 Then
dte_caso = dte_folio
dato3(15) = "<Referencia>"
dato3(16) = "<NroLinRef>1</NroLinRef>"
dato3(17) = "<TpoDocRef>SET</TpoDocRef>"
dato3(18) = "<FolioRef>" + dte_folio + "</FolioRef>"
dato3(19) = "<FchRef>" + Format(fecha, "yyyy-mm-dd") + "</FchRef>"
dato3(20) = "<CodRef>3</CodRef>"
dato3(21) = "<RazonRef>" + "CAJA:" + caja + " NI:" & numero & "</RazonRef>"
dato3(22) = "</Referencia>"


Rem End If



If dte_tipodte = 61 Or dte_tipodte = 56 Then
dato3(23) = "<Referencia>"
dato3(24) = "<NroLinRef>2</NroLinRef>"
dato3(25) = "<TpoDocRef>" + leerdatocaso("ref_tipo", tipo, numero, caja, fecha) + "</TpoDocRef>"
dato3(26) = "<FolioRef>" + leerdatocaso("ref_folio", tipo, numero, caja, fecha) + "</FolioRef>"
dato3(27) = "<FchRef>" + Format(leerdatocaso("ref_fecha", tipo, numero, caja, fecha), "yyyy-mm-dd") + "</FchRef>"
dato3(28) = "<CodRef>" + leerdatocaso("ref_codref", tipo, numero, caja, fecha) + "</CodRef>"
dato3(29) = "<RazonRef>" + leerdatocaso("ref_glosa", tipo, numero, caja, fecha) + "</RazonRef>"
dato3(30) = "</Referencia>"

End If

'If dte_tipodte = 56 And dte_folio = 1 Then
'dato3(15) = "<Referencia>"
'dato3(16) = "<NroLinRef>1</NroLinRef>"
'dato3(17) = "<TpoDocRef>61</TpoDocRef>"
'dato3(18) = "<FolioRef>18</FolioRef>"
'dato3(19) = "<FchRef>" + "2010-09-19" + "</FchRef>"
'dato3(20) = "<CodRef>1</CodRef>"
'
'dato3(21) = "<RazonRef>Anula Nota de Credito" + "</RazonRef>"
'dato3(22) = "</Referencia>"
'
'End If
'If dte_tipodte = 56 And dte_folio = 3 Then
'dato3(15) = "<Referencia>"
'dato3(16) = "<NroLinRef>1</NroLinRef>"
'dato3(17) = "<TpoDocRef>33</TpoDocRef>"
'dato3(18) = "<FolioRef>307</FolioRef>"
'dato3(19) = "<FchRef>" + "2010-09-15" + "</FchRef>"
'dato3(20) = "<CodRef>3</CodRef>"
'
'dato3(21) = "<RazonRef>Anula Descuentos" + "</RazonRef>"
'dato3(22) = "</Referencia>"
'
'End If
'
'
'If dte_tipodte = 61 And dte_folio = 13 Then
'dato3(15) = "<Referencia>"
'dato3(16) = "<NroLinRef>1</NroLinRef>"
'dato3(17) = "<TpoDocRef>33</TpoDocRef>"
'dato3(18) = "<FolioRef>302</FolioRef>"
'dato3(19) = "<FchRef>" + "2010-09-15" + "</FchRef>"
'dato3(20) = "<CodRef>3</CodRef>"
'
'dato3(21) = "<RazonRef>Devolucion de Productos" + "</RazonRef>"
'dato3(22) = "</Referencia>"
'
'End If
'
'If dte_tipodte = 61 And dte_folio = 15 Then
'dato3(15) = "<Referencia>"
'dato3(16) = "<NroLinRef>1</NroLinRef>"
'dato3(17) = "<TpoDocRef>33</TpoDocRef>"
'dato3(18) = "<FolioRef>303</FolioRef>"
'dato3(19) = "<FchRef>" + "2010-09-15" + "</FchRef>"
'dato3(20) = "<CodRef>1</CodRef>"
'
'dato3(21) = "<RazonRef>" + "Anula Documento" + "</RazonRef>"
'dato3(22) = "</Referencia>"
'
'End If
'If dte_tipodte = 61 And dte_folio = 16 Then
'dato3(15) = "<Referencia>"
'dato3(16) = "<NroLinRef>1</NroLinRef>"
'dato3(17) = "<TpoDocRef>33</TpoDocRef>"
'dato3(18) = "<FolioRef>305</FolioRef>"
'dato3(19) = "<FchRef>" + "2010-09-15" + "</FchRef>"
'dato3(20) = "<CodRef>3</CodRef>"
'
'dato3(21) = "<RazonRef>Devolucion de Productos" + "</RazonRef>"
'dato3(22) = "</Referencia>"
'
'End If
'
'

pasada = ""
For i = 1 To 40
If dato3(i) <> "" Then
pasada = pasada + dato3(i)
End If

Next i

DETALLE = DETALLE + pasada + "</Documento></DTE>"
DETALLE = cabeza + DETALLE
cadena = DETALLE

For K = 1 To Len(DETALLE)
If Asc(Mid(DETALLE, K, 1)) > 128 And Mid(DETALLE, K, 1) <> "—" Then
cadena = Replace(cadena, Mid(DETALLE, K, 1), "")
End If

Next K
DETALLE = cadena

DETALLE = Replace(DETALLE, "•", "N")
DETALLE = Replace(DETALLE, "—", "#209;")
DETALLE = Replace(DETALLE, "ß", " ")
DETALLE = Replace(DETALLE, "∫", " ")
DETALLE = Replace(DETALLE, "∞", " ")
DETALLE = Replace(DETALLE, "&", "&amp;")
DETALLE = Replace(DETALLE, "¯", " ")
DETALLE = Replace(DETALLE, ",", ".")
DETALLE = Replace(DETALLE, "*", "x")
DETALLE = Replace(DETALLE, "¥", "")
DETALLE = Replace(DETALLE, "«", "")
DETALLE = Replace(DETALLE, "Ô", "")


caf = leerrutacaf(dte_tipodte, dte_folio)
entrada = "C:\FAE\" + confi_empresaactiva + "\DTE\" & dte_tipodte & "_" & dte_folio & ".xml"
salida = "C:\FAE\" + confi_empresaactiva + "\DTE\firmado_" & dte_tipodte & "_" & dte_folio & ".xml"

xml.LoadXml DETALLE


Call xml.SaveXml(entrada)



FIRMA2 = entrada + " " + salida + " " + caf + " " + certificado
Rem FIRMA2 = "-a " + caf + " -p " + entrada + " -c " + CERTIFICADO + " -s 123 -o " + salida
Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + confi_empresaactiva + "\programas;C:\fae\" + confi_empresaactiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\jargs.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\itext-1.3.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.FirmaDTE " + FIRMA2
Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;C:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.FirmaDTE " + FIRMA2
FIRMA = "c:\fae\programas\firmadte.bat " + FIRMA2

Shell FIRMA
Call Sleep(10000)
DETALLE = leerxml(salida)
5 If DETALLE = "" Then Exit Sub
Rem Call GRABADTE(dte_tipodte, dte_folio, Format(fecha, "yyyy-mm-dd"), loc, tipo, numero, Format(fecha, "yyyy-mm-dd"), caja, detalle)

Rem enviasii = "c:\fae\programas\enviasii.bat " + entradaenvio + " " + certificado + " " + salidaenvio
Rem enviasii = "c:\fae\programas\enviasii.bat " + entradaenvio + " " + certificado + " " + salidaenvio


Rem datos para firmar envio
Rem detalle = timbrafactura(rutemisor, RUTENVIA, rutreceptor, Format(fechaemision, "dd-mm-yyyy"), "0", numero, TIPO) + leerxml(salida) + "</SetDTE></EnvioDTE>"
Rem Call xml.LoadXML(detalle)
Rem Call xml.SaveXml(entradafirma)

Rem Shell firmaenvio
Rem Shell enviasii
Rem Kill entrada
Rem Kill salida
descuento1 = 0
descuento2 = 0
End Sub
Public Function leerxml(ARCHIVO) As String

Dim ss As String
Dim AA As Double
Dim nombrearchivo As String
Dim contador As Double

nombrearchivo = ARCHIVO
Close 20
leerxml = ""


10 If ExisteArchivo(nombrearchivo) = True Then

Open ARCHIVO For Input As #20
AA = 0
20 If EOF(20) = False Then
Line Input #20, ss
If AA > 0 Then
leerxml = leerxml + ss
End If
AA = AA + 1
GoTo 20:

End If
Close 20

Else
contador = contador + 1
If contador = 1000000 Then MsgBox "ERROR NO ENCUENTRA EL ARCHIVO " + nombrearchivo


Exit Function
End If
Rem If leerxml = "" Then End


End Function

Public Function leerxmlrecibido(ARCHIVO) As String

Dim ss As String
Dim AA As Double
Dim nombrearchivo As String
Dim contador As Double

nombrearchivo = ARCHIVO
Close 20
leerxmlrecibido = ""


10 If ExisteArchivo(nombrearchivo) = True Then

Open ARCHIVO For Input As #20
20 If EOF(20) = False Then
Line Input #20, ss
leerxmlrecibido = leerxmlrecibido + ss



GoTo 20
End If
Close 20

End If
End Function


Public Sub GRABADTE(ARCHIVO)
    Dim campos(20, 3) As Variant
    
    Dim op As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "localdocumento"
    campos(4, 0) = "tipodocumento"
    campos(5, 0) = "numerodocumento"
    campos(6, 0) = "fechadocumento"
    campos(7, 0) = "cajadocumento"
    campos(8, 0) = "xml"
    campos(9, 0) = "xmlpdf"
    
    campos(10, 0) = ""
    
    tipo = leerdatoxml(xmlcorto(ARCHIVO), "TipoDTE", 2)
    numero = leerdatoxml(xmlcorto(ARCHIVO), "Folio", 2)
    fecha = "2011-02-08"
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = fecha
    campos(3, 1) = LOCALDO
    campos(4, 1) = TIPODO
    campos(5, 1) = numerodo
    campos(6, 1) = fechado
    campos(7, 1) = CAJADO
    firmado = FIRMADTE(xmlcorto(ARCHIVO), tipo, numero)
    
    campos(8, 1) = firmado
    
    campos(9, 1) = ARCHIVO
    
    
    LOCALDO = "02"
    campos(0, 2) = "admin_fae" + LOCALDO + ".sv_dte" + LOCALDO
    condicion = ""
    op = 2
    sqlventas.response = campos
    Set sqlventas.conexion = VENTAS
    Call sqlventas.sqlventas(op, condicion)
    Rem Call FIRMADTE(campos(8, 1), tipo, numero)
    Call enviarsii_express(tipo, numero)
    
    
    
    
End Sub

Public Function timbrafactura(EMISOR, envia, RECEPTOR, fecha, RESOLU, numero, tipo, INICIO, final) As String
Dim dato(100) As String
Dim comi As String
Dim i As Integer
empresa
Call empresadte(codigocontable)
rutreceptor = "60803000-K"
Call clientedte(rutreceptor)

comi = Chr(34)
tipo = "33"
If tipo = "NF" Then tipo = "61"
If tipo = "ND" Then tipo = "56"
If tipo = "ND" Then tipo = "56"

dato(1) = "<EnvioDTE xmlns=" + comi + "http://www.sii.cl/SiiDte" + comi + " xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " version=" + comi + "1.0" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiDte EnvioDTE_v10.xsd" + comi + ">"
dato(2) = "<SetDTE ID=" + comi + "EnvDte-" & INICIO & "-" & final & comi + "> "
dato(3) = "<Caratula version =" + comi + "1.0" + comi + "> "
dato(4) = "<RutEmisor>" + EMISOR + "</RutEmisor>"
dato(5) = "<RutEnvia>" + envia + "</RutEnvia>"
dato(6) = "<RutReceptor>" + rutreceptor + "</RutReceptor>"
dato(7) = "<FchResol>" + "2010-12-29" + "</FchResol>"
dato(8) = "<NroResol>" + "198" + "</NroResol>  "
dato(9) = "<TmstFirmaEnv>" & Format(Date, "yyyy-mm-dd") & "T" & Time & "</TmstFirmaEnv>"
If tipoDoc(1) <> 0 Then
dato(10) = "<SubTotDTE>"
dato(11) = "<TpoDTE>" + "33" + "</TpoDTE>"
dato(12) = "<NroDTE>" & tipoDoc(1) & "</NroDTE>"
dato(13) = "</SubTotDTE>"

End If
If tipoDoc(2) <> 0 Then
dato(14) = "<SubTotDTE>"
dato(15) = "<TpoDTE>" + "56" + "</TpoDTE>"
dato(16) = "<NroDTE>" & tipoDoc(2) & "</NroDTE>"
dato(17) = "</SubTotDTE>"
End If
If tipoDoc(3) <> 0 Then
dato(18) = "<SubTotDTE>"
dato(19) = "<TpoDTE>" + "61" + "</TpoDTE>"
dato(20) = "<NroDTE>" & tipoDoc(3) & "</NroDTE>"
dato(21) = "</SubTotDTE>"
End If

dato(22) = "</Caratula>"
For i = 1 To 22
If dato(i) <> "" Then
timbrafactura = timbrafactura + Chr(13) + dato(i)
End If
Next i

End Function
Public Function timbrafacturacliente(EMISOR, envia, RECEPTOR, fecha, RESOLU, numero, tipo, INICIO, final) As String
Dim dato(100) As String
Dim comi As String
Dim i As Integer
empresa
Call empresadte(codigocontable)
Rem Call clientedte(rutreceptor)

comi = Chr(34)

dato(1) = "<EnvioDTE xmlns=" + comi + "http://www.sii.cl/SiiDte" + comi + " xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " version=" + comi + "1.0" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiDte EnvioDTE_v10.xsd" + comi + ">"
dato(2) = "<SetDTE ID=" + comi + "EnvDte-" & INICIO & "-" & final & comi + "> "
dato(3) = "<Caratula version =" + comi + "1.0" + comi + "> "
dato(4) = "<RutEmisor>" + EMISOR + "</RutEmisor>"
dato(5) = "<RutEnvia>" + envia + "</RutEnvia>"
dato(6) = "<RutReceptor>" + RECEPTOR + "</RutReceptor>"
dato(7) = "<FchResol>" + Format(fechacertificacion, "yyyy-mm-dd") + "</FchResol>"
dato(8) = "<NroResol>" + resolucion + "</NroResol>  "
dato(9) = "<TmstFirmaEnv>" & Format(Date, "yyyy-mm-dd") & "T" & Time & "</TmstFirmaEnv>"
If tipoDoc(1) <> 0 Then
dato(10) = "<SubTotDTE>"
dato(11) = "<TpoDTE>" + "33" + "</TpoDTE>"
dato(12) = "<NroDTE>" & tipoDoc(1) & "</NroDTE>"
dato(13) = "</SubTotDTE>"
End If
If tipoDoc(2) <> 0 Then
dato(14) = "<SubTotDTE>"
dato(15) = "<TpoDTE>" + "52" + "</TpoDTE>"
dato(16) = "<NroDTE>" & tipoDoc(2) & "</NroDTE>"
dato(17) = "</SubTotDTE>"
End If
If tipoDoc(3) <> 0 Then
dato(18) = "<SubTotDTE>"
dato(19) = "<TpoDTE>" + "61" + "</TpoDTE>"
dato(20) = "<NroDTE>" & tipoDoc(3) & "</NroDTE>"
dato(21) = "</SubTotDTE>"
End If
If tipoDoc(4) <> 0 Then
dato(22) = "<SubTotDTE>"
dato(23) = "<TpoDTE>" + "34" + "</TpoDTE>"
dato(24) = "<NroDTE>" & tipoDoc(4) & "</NroDTE>"
dato(25) = "</SubTotDTE>"
End If

dato(26) = "</Caratula>"
For i = 1 To 26
If dato(i) <> "" Then
timbrafacturacliente = timbrafacturacliente + Chr(13) + dato(i)
End If
Next i

End Function

Public Function leerfoliodte(loc, tipo) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select numero from " + clientesistema + "fae" + loc + ".sv_dte" + loc
csql.sql = csql.sql & " where tipo='" + tipo + "' and numero<9999999999 order by numero desc limit 0,1"
csql.Execute
leerfoliodte = 1
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerfoliodte = resultados(0) + 1
End If
csql.Close
Set csql = Nothing
End Function
Public Function leerfolioautorizado(loc, tipo, numero) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select hasta from " + clientesistema + "fae" + loc + ".sv_caf" + loc
csql.sql = csql.sql & " where tipo='" + tipo + "' and desde<='" & numero & "' and hasta >= '" & numero & " ' "
csql.Execute
leerfolioautorizado = False

If csql.RowsAffected > 0 Then
    leerfolioautorizado = True
    
Else
MsgBox "FOLIO DE DOCUMENTO NO AUTORIZADO POR SII "

End If
csql.Close
Set csql = Nothing
End Function

Public Function existedte(loc, tipo, numero, fecha, caja, impre) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select numero from " + clientesistema + "fae" + loc + ".sv_dte" + loc
csql.sql = csql.sql & " where tipodocumento='" + tipo + "' and numerodocumento='" + numero + "' and fechadocumento='" + Format(fecha, "yyyy-mm-dd") + "' and cajadocumento='" + caja + "' "
If impre = "1" Then
csql.sql = csql.sql + "and impresa='0' "
End If
csql.Execute
existedte = False
If csql.RowsAffected > 0 Then
    existedte = True
End If
csql.Close
Set csql = Nothing
End Function

Public Function foliodte(loc, tipo, numero, fecha, caja) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select numero from " + clientesistema + "fae" + loc + ".sv_dte" + loc
csql.sql = csql.sql & " where tipodocumento='" + tipo + "' and numerodocumento='" + numero + "' and fechadocumento='" + Format(fecha, "yyyy-mm-dd") + "' and cajadocumento='" + caja + "' "
csql.Execute
foliodte = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    foliodte = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function
Public Function leerxmldte(loc, tipo, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = VENTAS
csql.sql = "select xml from " + clientesistema + "fae" + loc + ".sv_dte" + loc
csql.sql = csql.sql & " where tipo='" & tipo & "' and numero='" + numero + "' "
csql.Execute
leerxmldte = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerxmldte = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function
Public Function leerxmldtecliente(loc, tipo, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = VENTAS
csql.sql = "select xml from " + clientesistema + "fae" + loc + ".sv_dte" + loc
csql.sql = csql.sql & " where tipo='" + tipo + "' and numero='" + numero + "' "
csql.Execute
leerxmldtecliente = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerxmldtecliente = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function

Public Function leerxmldterecibido(loc, tipo, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = VENTAS
csql.sql = "select xml from " + clientesistema + "fae" + loc + ".sv_dte" + loc + "_recibidos "
csql.sql = csql.sql & " where tipo='" + tipo + "' and numero='" + numero + "' "
csql.Execute
leerxmldterecibido = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    Rem If IsNull(resultados(0)) = False Then
        leerxmldterecibido = resultados(0)
    Rem End If
End If
csql.Close
Set csql = Nothing
End Function
Public Function leerxmllibro(ruta, loc) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = VENTAS
csql.sql = "select respuesta from " + clientesistema + "fae" + loc + ".sv_envio_libros" + loc + " "
csql.sql = csql.sql & " where ruta='" + ruta + "' "
csql.Execute
leerxmllibro = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerxmllibro = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function

Public Function leerdatodte(loc, tipo, numero, dato) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = VENTAS
csql.sql = "select " + dato + " from " + clientesistema + "fae" + loc + ".sv_dte" + loc
csql.sql = csql.sql & " where tipo='" & tipo & "' and numero='" + numero + "' "
csql.Execute
leerdatodte = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    If IsNull(resultados(0)) = False Then
    leerdatodte = resultados(0)
    End If
End If
csql.Close
Set csql = Nothing
End Function

Public Function leerdatosfactura(tipo, numero, caja, fecha)

End Function
Public Sub empresadte(empresa)
   On Error GoTo pasa:
   Dim op As Integer
   Dim K As Integer
   Dim campos(20, 10) As String
   
    campos(0, 0) = "nombre"
    campos(1, 0) = "direccion"
    campos(2, 0) = "comuna"
    campos(3, 0) = "ciudad"
    campos(4, 0) = "rut"
    campos(5, 0) = "girodte"
    campos(6, 0) = "codactividadeconomica"
    campos(7, 0) = "certificado"
    campos(8, 0) = "rutenviasii"
    campos(9, 0) = "clave_certificado"
    campos(10, 0) = "fecharesolucion"
    campos(11, 0) = "numeroresolucion"
    campos(12, 0) = "servermail"
    campos(13, 0) = "mailsalida"
    campos(14, 0) = "clavemail"
    campos(15, 0) = "empresafae"
    campos(16, 0) = ""
    campos(0, 2) = clientesistema + "conta.maestroempresas"
  
    condicion = "codigoempresa = '" + confi_empresaactiva + "' "
    op = 5
    sqlventas.response = campos
    Set sqlventas.conexion = gestion
    
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
    dte_e_nombre = sqlventas.response(0, 3)
    dte_e_direccion = sqlventas.response(1, 3)
    dte_e_comuna = sqlventas.response(2, 3)
    dte_e_ciudad = sqlventas.response(3, 3)
    dte_e_rut = sqlventas.response(4, 3)
    dte_e_giro = sqlventas.response(5, 3)
    dte_e_acti = sqlventas.response(6, 3)
    certificado = sqlventas.response(7, 3)
    dte_rutenvia = sqlventas.response(8, 3)
    clavecertificado = sqlventas.response(9, 3)
    fechacertificacion = sqlventas.response(10, 3)
    resolucion = sqlventas.response(11, 3)
    confi_localempresa = sqlventas.response(15, 3)
    Call librerias_java
    
    
    End If
Exit Sub
pasa:
    MsgBox "empresa  no es electronica"
End Sub
Public Sub clientedte(rut)
   
   Dim op As Integer
   Dim K As Integer
   Dim campos(10, 10) As String
   
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = "giro"
    campos(3, 0) = "direccion"
    campos(4, 0) = "comuna"
    campos(5, 0) = "ciudad"
    campos(6, 0) = ""
    
    campos(0, 2) = clientesistema + "ventas" + ".sv_maestroclientes "
  
    condicion = "rut = '" & rut & "' "
    op = 5
    sqlventas.response = campos
    Set sqlventas.conexion = gestion
    
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
    
    dte_r_rut = Format(Mid(sqlventas.response(0, 3), 1, 9), "#########") + "-" + Mid(sqlventas.response(0, 3), 10, 1)
    dte_r_nombre = sqlventas.response(1, 3)
    dte_r_giro = sqlventas.response(2, 3)
    dte_r_direccion = sqlventas.response(3, 3)
    dte_r_comuna = sqlventas.response(4, 3)
    dte_r_ciudad = sqlventas.response(5, 3)
    
    End If
End Sub

Sub leerdetalledte(tipo, numero, caja, fecha, dcto)
   Dim csql As New rdoQuery
   Dim resultados As rdoResultset
   Dim tabla As String
   Dim linea As Double
   Dim taza As Double
   Dim Descuento As Double
   Dim montodescuento As Double
   Dim total As Double
    Dim DETALLENETO As Double
    
   taza = 1.19
   dte_descuento = dcto
   Set csql.ActiveConnection = ventasRubro
   tabla = ""
   tabla = "select codigo,descripcion,cantidad,precio,total,descuento,impuesto,porcentajeimpuesto "
   tabla = tabla & "from sv_documento_detalle_" & confi_empresaactiva
   tabla = tabla & " where caja='" + caja + "' and tipo='" & tipo & "' and numero='" & numero & "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' and local='" & confi_empresaactiva & "' order by linea"
   csql.sql = tabla
   csql.Execute
   linea = 0
   If csql.RowsAffected > 0 Then
   Set resultados = csql.OpenResultset
   
   dte_neto = 0
   dte_iva = 0
   dte_total = 0
   dte_ref = 0
   dte_vin = 0
   dte_lic = 0
   dte_cer = 0
   dte_har = 0
   dte_car = 0
   dte_exento = 0
   dte_descuentoglobal = 0

  While Not resultados.EOF
    linea = linea + 1
    taza = 1.19 + resultados("porcentajeimpuesto")
    tazaimpuesto = resultados("porcentajeimpuesto")
    matrixdte(linea, 1) = resultados("codigo")
    matrixdte(linea, 2) = resultados("cantidad")
    matrixdte(linea, 3) = resultados("descripcion")
    Rem DETALLENETO = Round(resultados("total") / taza / resultados("cantidad"), 4)
    DETALLENETO = Round((resultados("precio") / taza), 4)
    matrixdte(linea, 4) = DETALLENETO
    If dte_descuento = 0 Then
    matrixdte(linea, 5) = resultados("descuento")
    Else
    Rem matrixdte(linea, 5) = dte_descuento
    matrixdte(linea, 5) = "0"
    End If
    
    total = matrixdte(linea, 2) * matrixdte(linea, 4)
    If Val(matrixdte(linea, 5)) <> 0 Then
    montodescuento = Round((total * (matrixdte(linea, 5) / 100)))
    Else
    montodescuento = 0
    End If
    If dte_descuento <> 0 Then
    If Val(resultados("impuesto")) = 0 And dte_descuento <> 0 Then
    descuento1 = descuento1 + montodescuento
    End If
    
'    If Val(resultados("impuesto")) = 8 And dte_descuento <> 0 Then
'    descuento2 = descuento2 + montodescuento
'    End If
    End If
    matrixdte(linea, 8) = montodescuento
    total = total - montodescuento
    matrixdte(linea, 6) = Round(total, 0)
    matrixdte(linea, 7) = ""
    If Val(resultados("impuesto")) <> 0 And Val(resultados("impuesto")) <> 8 Then
    matrixdte(linea, 7) = leerimpuestofae(resultados("impuesto"))
    If Val(resultados("impuesto")) = 5 Then
    dte_car = dte_car + Round(total * tazaimpuesto)
    End If
    If Val(resultados("impuesto")) = 4 Then
    dte_har = dte_har + Round(total * tazaimpuesto)
    End If
    If Val(resultados("impuesto")) = 3 Then
    dte_lic = dte_lic + (total * tazaimpuesto)
    End If
    If Val(resultados("impuesto")) = 2 Then
    dte_vin = dte_vin + (total * tazaimpuesto)
    End If
    If Val(resultados("impuesto")) = 6 Then
    dte_cer = dte_cer + (total * tazaimpuesto)
    End If
    If Val(resultados("impuesto")) = 1 Then
    dte_ref = dte_ref + (total * tazaimpuesto)
    End If
    
    
    Else
    matrixdte(linea, 7) = ""
    
    
    End If
    matrixdte(linea, 0) = "0"
    If Val(resultados("impuesto")) <> 8 Then
    dte_neto = dte_neto + total
    
    Else
    dte_exento = dte_exento + total
    matrixdte(linea, 0) = "1"
    
    End If
    
    resultados.MoveNext
    
  Wend
   End If
   csql.Close
   Set resultados = Nothing
   Set csql = Nothing
   
   
   If dte_descuento <> 0 Then

   Descuento = Round(dte_neto * dte_descuento / 100)
   dte_neto = dte_neto - Descuento
   descuento1 = Descuento
   dte_iva = Round(dte_neto * 19 / 100, 0)

   Descuento = Round(dte_ref * dte_descuento / 100, 3)
   dte_ref = dte_ref - Descuento

   Descuento = Round(dte_vin * dte_descuento / 100, 3)
   dte_vin = dte_vin - Descuento

   Descuento = Round(dte_lic * dte_descuento / 100)
   dte_lic = dte_lic - Descuento

   Descuento = Round(dte_cer * dte_descuento / 100)
   dte_cer = dte_cer - Descuento

   Descuento = Round(dte_har * dte_descuento / 100)
   dte_har = dte_har - Descuento

   Descuento = Round(dte_car * dte_descuento / 100)
   dte_car = dte_car - Descuento

  Rem  Descuento = Round(dte_exento * dte_descuento / 100)
   dte_exento = dte_exento

'   If dte_exento <> 0 Then
'   descuento2 = Descuento
'
'   End If
'
   
   End If
   dte_iva = Round(dte_neto * 19 / 100, 0)
   dte_ref = Round(dte_ref)
   dte_vin = Round(dte_vin)
   dte_lic = Round(dte_lic)
   dte_cer = Round(dte_cer)
   dte_har = Round(dte_har)
   dte_car = Round(dte_car)
   dte_neto = Round(dte_neto)
   dte_exento = Round(dte_exento)
   dte_total = dte_neto + dte_iva + dte_ref + dte_vin + dte_cer + dte_lic + dte_har + dte_car + dte_exento
   dte_descuento = Round(dte_descuento)
   totallineas = linea
    
    
    
   End Sub
Public Function leerrutacaf(tipo, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select ubicacion,nombredelarchivo from " + clientesistema + "fae" + confi_empresaactiva + ".sv_caf" + confi_empresaactiva
csql.sql = csql.sql & " where tipo='" + tipo + "' and desde<='" & numero & "' and hasta >= '" & numero & " ' "
csql.Execute


If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerrutacaf = resultados(0) + resultados(1)
    
End If
csql.Close
Set csql = Nothing
End Function
Public Sub modificaimpresa(tipo, folio)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_empresaactiva + ".sv_dte" + confi_empresaactiva + " "
        csql.sql = csql.sql & "set impresa='1' WHERE tipo='" + tipo + "' and numero='" + folio + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
        Set csql = Nothing
    End Sub

Sub imprimelectronica(tipo, folio)
Dim entradapdf As String
Dim salidapdf As String
Dim salidapdf2 As String


DETALLE2 = leerxmldte(confi_empresaactiva, tipo, folio)

DETALLE2 = Replace(DETALLE2, "&amp;", "&")
DETALLE2 = Replace(DETALLE2, "#209;", "—")
DETALLE2 = Replace(DETALLE2, "#243;", "Û")
DETALLE2 = Replace(DETALLE2, "¯", " ")


xml.LoadXml DETALLE2

Call xml.SaveXml(confi_rutafae + confi_empresaactiva + "\DTE\" & tipo & "-" & folio & ".xml") '"E:\FAE_eltit\"
entradapdf = confi_rutafae + confi_empresaactiva + "\DTE\" & tipo & "-" & folio & ".xml"
salidapdf = confi_rutafae + confi_empresaactiva + "\PDF\" & tipo & "-" & folio & ".pdf"
salidapdf2 = confi_rutafae + confi_empresaactiva + "\PDF\" & tipo & "-" & folio & "cedi.pdf"

Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\dte\rodrigo;C:\DTE\RODRIGO\lib\OpenLibsDTE.jar;C:\DTE\RODRIGO\lib\jargs.jar;C:\DTE\RODRIGO\lib\itext-1.3.jar;C:\DTE\RODRIGO\lib\log4j-1.2.14.jar;C:\DTE\RODRIGO\lib\xercesImpl.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\" + confi_empresaactiva + "\impresion\FA_estandar.properties -f c:\fae\" + confi_empresaactiva + "\impresion\FA_estandar2.pdf -o " + salidapdf
Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\OpenLibsDTE.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\xercesImpl.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\" + confi_empresaactiva + "\impresion\FA_estandar.properties -f c:\fae\" + confi_empresaactiva + "\impresion\FA_estandar2.pdf -o " + salidapdf
Rem original
Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\00\programas;C:\fae\00\programas\lib\apache-mime4j-0.6.jar;C:\fae\00\programas\lib\commons-codec-1.3.jar;C:\fae\00\programas\lib\commons-httpclient-3.0.jar;C:\fae\00\programas\lib\commons-logging-1.0.4.jar;C:\fae\00\programas\lib\httpclient-4.0.jar;C:\fae\00\programas\lib\httpcore-4.0.1.jar;C:\fae\00\programas\lib\httpmime-4.0.jar;C:\fae\programas\00\lib\itext-1.3.jar;C:\fae\00\programas\lib\jargs.jar;C:\fae\00\programas\lib\jdom.jar;C:\fae\00\programas\lib\log4j-1.2.14.jar;C:\fae\00\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\00\programas\lib\OpenLibsDte.jar;C:\fae\00\programas\lib\xbean.jar;C:\fae\00\programas\lib\xercesImpl.jar;C:\fae\programas\00\lib\xfire-all-1.2.6.jar;C:\fae\00\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\%3\impresion\FA_estandar.properties -f c:\fae\%3\impresion\FA_estandar2.pdf -o " + salidapdf
Rem original
Rem esto si va


Rem pdf = "E:\fae\programas\generapdf.bat  " + entradapdf + " " + salidapdf + " " + confi_empresaactiva
Rem Shell "E:\fae\programas\generapdf.bat  " + entradapdf + " " + salidapdf + " " + confi_empresaactiva
pdf = lib_generapdf + " -d " + entradapdf + " -p " & confi_rutafae & "%3\impresion\FA_estandar.properties -f " & confi_rutafae & confi_empresaactiva + "\impresion\FA_estandar2.pdf -o " + salidapdf

Shell pdf
Call Sleep(2000)

11: If ExisteArchivo(salidapdf) = True Then
  If imprIMETIPO <> "CAJA98" Then
      RUTA_IMPRESORA = leerdatoxml(DETALLE2, "RutaImpresora>", 1)
Call PrintFile(salidapdf, RUTA_IMPRESORA)
End If
Else
GoTo 11
End If

Rem cedible
cedible = True

    If cedible = True Then

        pdf = confi_rutafae & "programas\generapdf2.bat  " + entradapdf + " " + salidapdf2 + " " + confi_empresaactiva
        Shell pdf

        Call Sleep(2000)
10:         If ExisteArchivo(salidapdf2) = True Then
            If imprIMETIPO <> "CAJA98" Then
            RUTA_IMPRESORA = leerdatoxml(DETALLE2, "RutaImpresora>", 1)
            Call PrintFile(salidapdf2, RUTA_IMPRESORA)
            End If
            Else
            GoTo 10
            End If
    End If

Call modificaimpresa(tipo, folio)


End Sub
Sub enviasii(tipo, folio, RUTEMPRESA, rutenvia, rutreceptor, fecha, entradafirma)
empresa
Call empresadte(codigocontable)
rutreceptor = "60803000-K"
Call clientedte(rutreceptor)


firmaenvio = "c:\fae\programas\firmaenvio.bat " + entradafirma + " " + certificado + " " + salidafirma
Rem firmaenvio = "c:\fae\" + confi_empresaactiva + "\programas\firmaenvio.bat " + entradafirma + " " + CERTIFICADO + " " + salidafirma

Rem firmaenvio = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;C:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.FirmaEnvio -p %1 -c %2 -s 123 -o %3"


Shell firmaenvio
Sleep (4000)
'
'
respuestasii = "c:\fae\" + confi_empresaactiva + "\respuesta_sii\res" + folio + ".xml"
enviarsii = "c:\fae\programas\enviasii.bat " + salidafirma + " " + certificado + " c:\fae\" + confi_empresaactiva + "\respuesta_sii\res" + folio + ".xml"
Rem enviarsii = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + confi_empresaactiva + "\programas;C:\fae\" + confi_empresaactiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\jargs.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\itext-1.3.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.EnviarSII -p " + salidafirma + " -c " + CERTIFICADO + " -s 123 -o " + respuestasii
Shell enviarsii







End Sub
Public Function leerimpuestofae(codigo) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select codigofae from " + clientesistema + "gestion.g_maestroimpuestos "
csql.sql = csql.sql & " where codigo='" + codigo + "' "
csql.Execute

leerimpuestofae = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerimpuestofae = resultados(0)
    
End If
csql.Close
Set csql = Nothing
End Function

Public Function leerdatocaso(dato, tipo, numero, caja, fecha) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select " + dato + " from sv_documento_cabeza_" + confi_empresaactiva + " "
csql.sql = csql.sql & " where tipo='" + tipo + "' and numero='" + numero + "' and caja='" + caja + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' "
csql.Execute

leerdatocaso = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerdatocaso = resultados(0)
    
End If
csql.Close
Set csql = Nothing
End Function


Public Sub grabar_recepcion(correo, ARCHIVO, fecha_recepcion, archivo_recepcion, archivo_respuesta, fecha_respuesta)
Dim recepciona As String
Dim estado As String


Dim campos(10, 10) As String

        Dim op As Integer
        
        Dim sss As Double
        
        campos(0, 0) = "correo"
        campos(1, 0) = "archivo"
        campos(2, 0) = "fecha_recepcion"
        campos(3, 0) = "archivo_recepcion"
        campos(4, 0) = "archivo_respuesta"
        campos(5, 0) = "rut"
        campos(6, 0) = "tipo"
        campos(0, 1) = correo
        archivo_recepcion = Replace(archivo_recepcion, "'", "~")
        campos(1, 1) = ARCHIVO
        campos(2, 1) = Format(fecha_recepcion, "yyyy-mm-dd")
        campos(3, 1) = archivo_recepcion
        campos(4, 1) = archivo_respuesta
        campos(5, 1) = leerdatoxml(archivo_recepcion, "RutEmisor>", 1)
        campos(6, 1) = "PROVEEDOR"
        If Mid(archivo_recepcion, 44, 13) = "<RespuestaDTE" Then
        campos(6, 1) = "CLIENTE"
        End If
       
       If InStr(1, correo, "sii") > 0 Then
      campos(6, 1) = "SII"
      End If
        campos(0, 2) = clientesistema + "fae" + confi_localempresa + ".sv_recepcion_dte" + confi_localempresa
    
        condicion = ""

        op = 2
        sqlventas.response = campos
        Set sqlventas.conexion = VENTAS
        Call sqlventas.sqlventas(op, condicion)
        
        
       If InStr(1, correo, "sii") > 0 Then
       track = Format(leerdatoxml(archivo_recepcion, "TRACKID>", 1), "0000000000")
       

       If Mid(correo, 1, 13) = "siidte_error@" Then
       estado = "3"
       End If
       If Mid(correo, 1, 7) = "siidte@" Then
       estado = "1"
       End If
       If Mid(correo, 1, 14) = "siidte_reparo@" Then
       estado = "2"
       End If
       If track <> "" Then
       If UCase(Mid(ARCHIVO, 1, 3)) = "LBR" Then
       Call modificarespuestaenvioLIBRO(archivo_recepcion, track, estado)
       Else
       Call modificarespuestaenvio(archivo_recepcion, track, estado)
       
       End If
       
       End If
       End If
       
       
End Sub
Public Sub recepcionar(ARCHIVO, rut, certificado, CLAVE, RUTEMPRESA)
Dim dato(10) As String
Dim respuesta As String
Dim datos As String
Dim cantidaddefacturas As Double
Dim estadodte As String
Dim glosadte As String
Dim correo As String

        Call empresadte(confi_empresaactiva)
        datos = leerdatorecepcion_recibidos(ARCHIVO, confi_localempresa)
        
        cantidaddefacturas = cargarecepciones(datos)
        
'        If cantidaddefacturas < 5 Then
'        dato(0) = RUTADTE + confi_empresaactiva + "\" + ARCHIVO
'        Close 20
'        Open dato(0) For Output As #20
'        Print #20, datos
'        Close 20
'        Else
'        dato(0) = RUTADTE + ARCHIVO
'        End If
        dato(0) = confi_rutafae + confi_localempresa + "\correos_recibidos\" + ARCHIVO
        
        dato(1) = confi_rutafae + confi_localempresa + "\dte_recibidos\"
        dato(2) = confi_rutafae + confi_localempresa + "\ACUSE\" + "ACUSE_" + ARCHIVO
        dato(3) = certificado
        dato(4) = clavecertificado
        dato(5) = dte_rutenvia
       
        dato(6) = dte_e_rut
        dato(7) = "1"
        dato(8) = "1"
       
        dato(9) = dte_rec(1, 4)
        
        recepciona = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;c:\fae\programas\lib\apache-mime4j-0.6.jar;c:\fae\programas\lib\commons-codec-1.3.jar;c:\fae\programas\lib\commons-httpclient-3.0.jar;c:\fae\programas\lib\commons-logging-1.0.4.jar;c:\fae\programas\lib\httpclient-4.0.jar;c:\fae\programas\lib\httpcore-4.0.1.jar;c:\fae\programas\lib\httpmime-4.0.jar;c:\fae\programas\lib\itext-1.3.jar;c:\fae\programas\lib\jargs.jar;c:\fae\programas\lib\jdom.jar;c:\fae\programas\lib\log4j-1.2.14.jar;c:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;c:\fae\programas\lib\OpenLibsDte.jar;c:\fae\programas\lib\xbean.jar;c:\fae\programas\lib\xercesImpl.jar;c:\fae\programas\lib\xfire-all-1.2.6.jar;c:\fae\programas\lib\xmlsec-1.4.0.jar;c:\fae\programas\lib\XmlSchema-1.1.jar;c:\fae\programas\lib\wsdl4j.jar cl.eltit.dte.Recepcionar "
        Rem recepciona = "e:\fae_eltit\programas\recepcionar.bat "
        recepciona = lib_recepcionar + dato(0) + " " + dato(1) + " " + dato(2) + " " + dato(3) + " " + dato(4) + " " + dato(5) + " " + dato(6) + " " + dato(7) + " " + dato(8) + " " + dato(9) + " validarsii=N "
        
        Shell recepciona
        
        
        Sleep (5000)
        
        
        respuesta = leerxmlrecibido(dato(2))
    
        
        estadodte = ""
        glosadte = ""
        If respuesta <> "" Then
        estadodte = leerdatoxml(respuesta, "EstadoRecepEnv>", 1)
        glosadte = leerdatoxml(respuesta, "RecepEnvGlosa>", 1)
        If estadodte = "" Then
        estadodte = leerdatoxml(respuesta, "EstadoDTE>", 1)
        glosadte = leerdatoxml(respuesta, "EstadoDTEGlosa>", 1)
        
        End If
        End If
        If respuesta = "" Then MsgBox "RESPUESTA EN BLANCO DEL " & dato(2)
        If respuesta <> "" Then
        
        correo = recuperacorreo(ARCHIVO)
        
        Call modificarespuesta(ARCHIVO, respuesta, cantidaddefacturas)
        
        Rem If cantidaddefacturas > 5 Then Stop
        For K = 1 To cantidaddefacturas
        Call GRABARECEPCION(dte_rec(K, 1), dte_rec(K, 2), Format(dte_rec(K, 3), "yyyy-mm-dd"), dte_rec(K, 4), dte_rec(K, 5), Format(fechasistema, "yyyy-mm-dd"), ARCHIVO, dte_rec(K, 6), respuesta, "", "", "", correo, "", estadodte, glosadte)
        Next K
        End If

End Sub

Public Function leer_rutcorreo(correo) As String
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
correo = Replace(correo, "<", "")
correo = Replace(correo, ">", "")

 Set csql.ActiveConnection = VENTAS
 csql.sql = "select rut from " & clientesistema + "fae" + ".sv_fae_proveedores "
 csql.sql = csql.sql & "where mailintercambio='" & correo & "' "
 csql.Execute
 
 If csql.RowsAffected > 0 Then
             Set resultados = csql.OpenResultset
    
    leer_rutcorreo = resultados(0)
 End If
 csql.Close
 Set csql = Nothing
 
End Function


Public Sub modificarespuesta(ARCHIVO, respuesta, canti)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        respuesta = Replace(respuesta, "'", "")
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_localempresa + ".sv_recepcion_dte" + confi_localempresa + " "
        csql.sql = csql.sql & "set documentos='" & canti & "',archivo_respuesta='" + respuesta + "',fecha_respuesta='" + Format(fechasistema, "yyyy-mm-dd") + "'  WHERE archivo='" + ARCHIVO + "' "
        csql.Execute
        
        
        
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub
Public Function existecorreo(ARCHIVO) As Boolean

        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "select correo from " + clientesistema + "fae" + confi_empresaactiva + ".sv_recepcion_dte" + confi_empresaactiva + " "
        csql.sql = csql.sql & "WHERE archivo='" + ARCHIVO + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
        existecorreo = True
        Else
        existecorreo = False
        
        End If
        
        
        
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Function
Public Function recuperacorreo(ARCHIVO) As String

        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "select correo from " + clientesistema + "fae" + confi_localempresa + ".sv_recepcion_dte" + confi_localempresa + " "
        csql.sql = csql.sql & "WHERE archivo='" + ARCHIVO + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
    
        recuperacorreo = resultados(0)
        Else
        recuperacorreo = False
        
        End If
        
        
        
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Function


Public Sub modificaenviocorreo(ARCHIVO)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_localempresa + ".sv_dte" + confi_localempresa + "_recibidos "
        csql.sql = csql.sql & "set fecha_respuesta_enviada='" + Format(fechasistema, "yyyy-mm-dd") + "'  WHERE nombrearchivo='" + ARCHIVO + "' "
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub
Public Sub modificaaceptacioncliente(tipo, numero, respuesta)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_empresaactiva + ".sv_dte" + confi_empresaactiva + " "
        csql.sql = csql.sql & "set cli_acepta='" + respuesta + "' where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub


Public Function ultimo_envio() As String
Dim resultados As rdoResultset
        
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "select max(folio) from " + clientesistema + "fae" + confi_localempresa + ".sv_envios_dte" + confi_localempresa + " "
        csql.Execute
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
    
        ultimo_envio = resultados(0) + 1
        
        Else
        ultimo_envio = "1"
        
        End If
        
        
        
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Function

Public Sub modificaenvio(tipo, folio, envio, track)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_empresaactiva + ".sv_dte" + confi_empresaactiva + " "
        csql.sql = csql.sql & "set track='" + track + "', nombreenvio='" + envio + "',fechaenviosii='" + Format(fechasistema, "yyyy-mm-dd") + "' WHERE tipo='" + tipo + "' and numero='" + folio + "' "
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub
Public Sub modificaenviolibro(ruta, track)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        ruta = Replace(ruta, "\", "+")
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_localempresa + ".sv_envio_libros" + confi_localempresa + " "
        csql.sql = csql.sql & "set track='" + track + "', fecha_envio='" + Format(fechasistema, "yyyy-mm-dd") + "' WHERE ruta='" + ruta + "'"
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub

Public Sub eliminaenviolibro(ruta, track)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        ruta = Replace(ruta, "\", "+")
        Set csql.ActiveConnection = VENTAS
        csql.sql = "delete from " + clientesistema + "fae" + confi_localempresa + ".sv_envio_libros" + confi_localempresa + " "
        csql.sql = csql.sql & " WHERE ruta='" + ruta + "'"
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub
Public Sub eliminaenviodte(tipo, numero)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        ruta = Replace(ruta, "\", "+")
        Set csql.ActiveConnection = VENTAS
        csql.sql = "delete from " + clientesistema + "fae" + confi_empresaactiva + ".sv_dte" + confi_empresaactiva + " "
        csql.sql = csql.sql & " WHERE tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub

Public Sub modificaenviocliente(tipo, folio, correo)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_localempresa + ".sv_dte" + confi_localempresa + " "
        csql.sql = csql.sql & "set cli_envia='1', correo_envio_cliente='" + correo + "',fechaenviocliente='" + Format(fechasistema, "yyyy-mm-dd") + "' WHERE tipo='" & Val(Mid(tipo, 1, 3)) & "' and numero='" + folio + "' "
        csql.Execute
       
        csql.Close
        Set csql = Nothing
    End Sub

Public Sub grabar_envio(folio, fechaenvio, respuestaenvio, track, track_respuesta)
Dim recepciona As String
Dim campos(10, 10) As String
Dim op As Integer
        
        
        campos(0, 0) = "folio"
        campos(1, 0) = "fechaenvio"
        campos(2, 0) = "respuestaenvio"
        campos(3, 0) = "track"
        campos(4, 0) = "track_respuesta"
        campos(5, 0) = ""
        campos(0, 1) = folio
        campos(1, 1) = Format(fechaenvio, "yyyy-mm-dd")
        campos(2, 1) = respuestaenvio
        campos(3, 1) = track
        campos(4, 1) = track_respuesta
        
        
        
        campos(0, 2) = clientesistema + "fae" + confi_localempresa + ".sv_envios_dte" + confi_localempresa
    
        condicion = ""

        op = 2
        sqlventas.response = campos
        Set sqlventas.conexion = VENTAS
        Call sqlventas.sqlventas(op, condicion)
        
        
       
End Sub

Public Function leertrack(ARCHIVO) As String

Dim ss As String
Dim AA As Double
Dim nombrearchivo As String
Dim contador As Double

nombrearchivo = ARCHIVO
Close 20
leertrack = ""
10 If ExisteArchivo(nombrearchivo) = True Then
Open ARCHIVO For Input As #20
20 If EOF(20) = False Then
Line Input #20, ss
leertrack = leertrack + ss
For K = 1 To Len(ss)
If Mid(ss, K, 7) = "TRACKID" Then
leertrack = Mid(ss, K + 8, 10)
Exit Function
End If

Next K

GoTo 20
End If
Close 20

End If
End Function
Public Sub modificarespuestaenvio(envio, track, estado)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_localempresa + ".sv_dte" + confi_localempresa + " "
        csql.sql = csql.sql & "set respuestasii='" + envio + "',aceptada='" + estado + "' WHERE track = '" + track + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
        Set csql = Nothing
         Call grabarlogenvio(leercampo("tipo", track), leercampo("numero", track), leercampo("fechaenviosii", track), track, estado)
    End Sub
Public Sub modificarespuestaenvioLIBRO(envio, track, estado)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + confi_localempresa + ".sv_envio_libros" + confi_localempresa + " "
        csql.sql = csql.sql & "set respuesta='" + envio + "',estado='" + estado + "' WHERE track = '" + track + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
        Set csql = Nothing
        Rem  Call grabarlogenvio(leercampo("tipo", track), leercampo("numero", track), leercampo("fechaenviosii", track), track, estado)
    End Sub

Public Sub modificafacturadepublicidad(tipo, fecha, numero, caja, empresa, folio)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".movimientoscontables "
        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='EF' and numero='" + numero + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "'"
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
    
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = VENTAS
    
        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".facturasdepublicidad "
        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='2' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
    
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = VENTAS
        
        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".facturasdepublicidad_glosa "
        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='2' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
    
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = VENTAS
    
        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".facturasdeventas "
        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='6' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = VENTAS
    
        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".facturasdeventas_detalle "
        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='6' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, VENTAS)
        csql.Close
    
    End Sub


Public Function xmlcorto(ARCHIVO) As String


For K = 1 To Len(ARCHIVO)
If Mid(ARCHIVO, K, 12) = "<Parametros>" Then
xmlcorto = Mid(ARCHIVO, 1, K - 1) + "</DTE>"
End If

Next K

End Function
Public Function FIRMADTE(DETALLE, tipo, folio) As String


cadena = DETALLE


For K = 1 To Len(DETALLE)
If Asc(Mid(DETALLE, K, 1)) > 128 And Mid(DETALLE, K, 1) <> "—" Then
cadena = Replace(cadena, Mid(DETALLE, K, 1), "")
End If

Next K
DETALLE = cadena

DETALLE = Replace(DETALLE, "•", "N")
DETALLE = Replace(DETALLE, "—", "#209;")
DETALLE = Replace(DETALLE, "ß", " ")
DETALLE = Replace(DETALLE, "∫", " ")
DETALLE = Replace(DETALLE, "∞", " ")
DETALLE = Replace(DETALLE, "&", "&amp;")
DETALLE = Replace(DETALLE, "¯", " ")
DETALLE = Replace(DETALLE, ",", ".")
DETALLE = Replace(DETALLE, "*", "x")
DETALLE = Replace(DETALLE, "¥", "")
DETALLE = Replace(DETALLE, "«", "")
DETALLE = Replace(DETALLE, "Ô", "")

certificado = "e:\evantec3.pfx"
CLAVE = "evantec"
caf = leerrutacaf(tipo, folio)
entrada = "e:\FAE_eltit\" + confi_empresaactiva + "\DTE\" & tipo & "_" & folio & ".xml"
salida = "e:\FAE_eltit\" + confi_empresaactiva + "\DTE\firmado_" & tipo & "_" & folio & ".xml"

xml.LoadXml DETALLE


Call xml.SaveXml(entrada)



FIRMA2 = entrada + " " + salida + " " + caf + " " + certificado + " " + CLAVE
Rem FIRMA2 = "-a " + caf + " -p " + entrada + " -c " + CERTIFICADO + " -s 123 -o " + salida
Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + confi_empresaactiva + "\programas;C:\fae\" + confi_empresaactiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\jargs.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\itext-1.3.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.FirmaDTE " + FIRMA2
Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;C:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.FirmaDTE " + FIRMA2
FIRMA = "e:\fae_eltit\programas\firmadte.bat " + FIRMA2

Shell FIRMA
Call Sleep(10000)
FIRMADTE = leerxml(salida)
Rem Call imprimelectronica(tipo, folio)

End Function
Public Function leerdatoxml(ARCHIVO, campo, numero) As String
Dim o As Double

Dim numero2 As Double

Dim ss As String
Dim AA As Double
Dim nombrearchivo As String
Dim contador As Double
numero2 = numero * 2
Dim INICIO As Double
Dim final As Double
Rem If numero = 2 Then Stop
contador = 0
For o = 1 To Len(ARCHIVO)
If UCase(Mid(ARCHIVO, o, Len(campo))) = UCase(campo) Then
contador = contador + 1
If contador = numero2 - 1 Then
INICIO = o + Len(campo)
End If

If contador = numero2 Then
final = o - INICIO - 2
End If
If contador = numero2 Then Exit For

End If

Next o
If INICIO <> 0 And final <> 0 Then
leerdatoxml = Mid(ARCHIVO, INICIO, final)
End If

End Function

Public Sub enviarsii_express(tipo, numero)
Dim K As Integer
Dim INICIO As Double
Dim final As Double
Dim TIPOENVIO As String
Dim folioen As String

Dim CUENTA As Double
Dim entradafirma As String
Dim salidafirma As String
Dim rutreceptor As String
Dim detalle3 As String
Dim firmaenvio As String
Dim folio As String
Dim respuestasii As String
Dim enviarsii As String
Dim respuestaenvio As String
Dim track As String
Dim track_respuesta As String

folio = "1"

detalleenviosii = ""
CUENTA = 0
tipoDoc(1) = 0
tipoDoc(2) = 0
tipoDoc(3) = 0
INICIO = numero
detalleenviosii = detalleenviosii + leerxmldte(confi_empresaactiva, tipo, numero)
final = numero
CUENTA = CUENTA + 1
If tipo = "33" Then
tipoDoc(1) = tipoDoc(1) + 1
End If
If tipo = "62" Then
tipoDoc(2) = tipoDoc(2) + 1
End If
If tipo = "61" Then
tipoDoc(3) = tipoDoc(3) + 1
End If





'Rem Call xml.LoadXML(detalleenviosii)
'Rem SALIDAENVIO = "c:\fae\" + confi_empresaactiva + "\paso.xml"
folioen = ultimo_envio


entradafirma = "e:\fae_eltit\" + confi_empresaactiva + "\envio_sii\envio_" + folioen + ".xml"
salidafirma = "e:\fae_eltit\" + confi_empresaactiva + "\envio_sii\timbrado_envio_" + folioen + ".xml"
'Rem Call xml.SaveXml(SALIDAENVIO)
''
'If Option1.Value = True Then TIPOENVIO = "FV"
'If Option2.Value = True Then TIPOENVIO = "ND"
'If Option3.Value = True Then TIPOENVIO = "NF"
rutreceptor = "60803000-K"
dte_e_rut = "78209000-3"
dte_e_rutenvia = "7762388-4"
detalle3 = timbrafactura(dte_e_rut, dte_rutenvia, rutreceptor, Format(fechasistema, "dd-mm-yyyy"), "0", CUENTA, TIPOENVIO, INICIO, final) + Chr(13) + detalleenviosii + Chr(13) + "</SetDTE></EnvioDTE>"
'Rem Call xml.LoadXML(detalle3)
'Rem Call xml.SaveXml(entradafirma)
'
'Rem firmaenvio = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + confi_empresaactiva + "\programas\;C:\fae\" + confi_empresaactiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\jargs.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\itext-1.3.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.FirmaEnvio -p " + entradafirma + " -c " + CERTIFICADO + " -s 123 -o " + salidafirma
detalle3 = Replace(detalle3, "•", "N")
detalle3 = Replace(detalle3, "—", "#209;")
detalle3 = Replace(detalle3, "ß", " ")
detalle3 = Replace(detalle3, "«", " ")

detalle3 = Replace(detalle3, "∫", " ")
detalle3 = Replace(detalle3, "∞", " ")
detalle3 = Replace(detalle3, "Û", "&#243;")
detalle3 = Replace(detalle3, ",", ".")
detalle3 = Replace(detalle3, "*", "x")
detalle3 = Replace(detalle3, "Å", " ")
detalle3 = Replace(detalle3, "Ô", " ")
detalle3 = Replace(detalle3, "¯", " ")

Close 22
If detalleenviosii = "" Then
MsgBox ("DEBE SELECCIONAR ENVIOS ")
Exit Sub
End If

Open entradafirma For Output As #22
Print #22, detalle3
Close 22
firmaenvio = "e:\fae_eltit\programas\firmaenvio.bat " + entradafirma + " " + certificado + " " + salidafirma
Rem firmaenvio = "c:\fae\" + confi_empresaactiva + "\programas\firmaenvio.bat " + entradafirma + " " + CERTIFICADO + " " + salidafirma
Rem firmaenvio = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;C:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.FirmaEnvio -p %1 -c %2 -s 123 -o %3"


Shell firmaenvio
Sleep (5000)
'
'
respuestasii = "e:\fae_eltit\" + confi_empresaactiva + "\respuesta_sii\res_" + folioen + ".xml"
Rem enviassii = "java -Dfile.encoding=ISO-8859-1 -classpath e:\fae_eltit\programas;e:\fae_eltit\programas\lib\apache-mime4j-0.6.jar;e:\fae_eltit\programas\lib\commons-codec-1.3.jar;e:\fae_eltit\programas\lib\commons-httpclient-3.0.jar;e:\fae_eltit\programas\lib\commons-logging-1.0.4.jar;e:\fae_eltit\programas\lib\httpclient-4.0.jar;e:\fae_eltit\programas\lib\httpcore-4.0.1.jar;e:\fae_eltit\programas\lib\httpmime-4.0.jar;e:\fae_eltit\programas\lib\itext-1.3.jar;e:\fae_eltit\programas\lib\jargs.jar;e:\fae_eltit\programas\lib\jdom.jar;e:\fae_eltit\programas\lib\log4j-1.2.14.jar;e:\fae_eltit\programas\lib\not-yet-commons-ssl-0.3.11.jar;e:\fae_eltit\programas\lib\OpenLibsDte.jar;e:\fae_eltit\programas\lib\xbean.jar;e:\fae_eltit\programas\lib\xercesImpl.jar;e:\fae_eltit\programas\lib\xfire-all-1.2.6.jar;e:\fae_eltit\programas\lib\xmlsec-1.4.0.jar cl.admin.dte.FirmaDTE -a " + salidafirma + " -p " + salidafirma + " -c %4 -s %5 -o %2"
enviarsii = "e:\fae_eltit\programas\enviasii.bat " + salidafirma + " " + certificado + " e:\fae_eltit\" + confi_empresaactiva + "\respuesta_sii\res_" + folioen + ".xml"
Rem enviarsii = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + confi_empresaactiva + "\programas;C:\fae\" + confi_empresaactiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\jargs.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\itext-1.3.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.EnviarSII -p " + salidafirma + " -c " + CERTIFICADO + " -s 123 -o " + respuestasii
Shell enviarsii
Call Sleep(10000)

track = leertrack(respuestasii)
respuestaenvio = leerxmlrecibido(respuestasii)
If track <> "" Then
    Call modificaenvio(tipo, numero, "res_" + folioen + ".xml", track)
   
    Call grabar_envio(folioen, Format(fechasistema, "yyyy-mm-dd"), respuestaenvio, track, track_respuesta)
End If



End Sub

Sub imprimelectronica2(tipo, folio)
Dim entradapdf As String
Dim salidapdf As String
Dim salidapdf2 As String


DETALLE2 = leerxmldte(confi_empresaactiva, tipo, folio)

DETALLE2 = Replace(DETALLE2, "&amp;", "&")
DETALLE2 = Replace(DETALLE2, "#209;", "—")
DETALLE2 = Replace(DETALLE2, "#243;", "Û")
DETALLE2 = Replace(DETALLE2, "¯", " ")


xml.LoadXml DETALLE2

Call xml.SaveXml(RUTADTE + confi_empresaactiva + "\DTE\" & tipo & "_" & folio & ".xml")
entradapdf = RUTADTE + confi_empresaactiva + "\DTE\" & tipo & "_" & folio & ".xml"
salidapdf = RUTADTE + confi_empresaactiva + "\PDF\" & tipo & "_" & folio & ".pdf"
salidapdf2 = RUTADTE + confi_empresaactiva + "\PDF\" & tipo & "_" & folio & "cedi.pdf"
Rem pdf = "E:\fae\programas\generapdf.bat  " + entradapdf + " " + salidapdf + " " + confi_empresaactiva
pdf = "java -Dfile.encoding=ISO-8859-1 -classpath " & RUTADTE & "programas;" & RUTADTE _
 & "programas\lib\apache-mime4j-0.6.jar;" & RUTADTE & "programas\lib\commons-codec-1.3.jar;" & RUTADTE & "programas\lib\commons-httpclient-3.0.jar;" & RUTADTE & _
 "programas\lib\commons-logging-1.0.4.jar;" & RUTADTE & "programas\lib\httpclient-4.0.jar;" & RUTADTE & "programas\lib\httpcore-4.0.1.jar;" & RUTADTE & "programas\lib\httpmime-4.0.jar;" & RUTADTE & _
 "programas\lib\itext-1.3.jar;" & RUTADTE & "programas\lib\jargs.jar;" & RUTADTE & "programas\lib\jdom.jar;" & RUTADTE & _
 "programas\lib\log4j-1.2.14.jar;" & RUTADTE & "programas\lib\not-yet-commons-ssl-0.3.11.jar;" & RUTADTE & "programas\lib\OpenLibsDte.jar;" & RUTADTE & _
 "programas\lib\xbean.jar;" & RUTADTE & "programas\lib\xercesImpl.jar;" & RUTADTE & "programas\lib\xfire-all-1.2.6.jar;" & RUTADTE & "programas\lib\xmlsec-1.4.0.jar cl.admin.dte.GeneraPDF -d " + entradapdf + " -p " & RUTADTE & "%3\impresion\FA_estandar.properties -f " & RUTADTE + confi_empresaactiva + "\impresion\FA_estandar2.pdf -o " + salidapdf
Shell pdf
Call Sleep(1000)

11: If ExisteArchivo(salidapdf) = True Then
Shell "C:\Archivos de programa\Adobe\Reader 10.0\Reader\acrord32.exe " + salidapdf
Else
GoTo 11
End If


Call modificaimpresa(tipo, folio)


End Sub

Public Function correo_cliente(rut) As String
Dim resultados As rdoResultset
        
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "select mailintercambio from " + clientesistema + "fae.sv_fae_proveedores where rut='" + rut + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
    
        correo_cliente = resultados(0)
        
        Else
        correo_cliente = ""
        
        End If
        
        
        csql.Close
        Set csql = Nothing
    End Function

Public Function leerdato_xml(palabra, ARCHIVO) As String

End Function

Public Function generapdf(tipo, folio, copia) As String
Dim entradapdf As String
Dim salidapdf As String
Dim salidapdf2 As String
tipo = Val(Mid(tipo, 1, 3))

DETALLE2 = leerxmldte(Mid(envio01.ComboLOCAL.text, 1, 2), Val(Mid(tipo, 1, 3)), folio)
Rem detalle2 = Replace(detalle2, "&amp;", "&")
DETALLE2 = Replace(DETALLE2, "#209;", "—")
DETALLE2 = Replace(DETALLE2, "#243;", "Û")
DETALLE2 = Replace(DETALLE2, "¯", " ")
xml.LoadXml DETALLE2

Call xml.SaveXml(confi_rutafae + confi_localempresa + "\DTE\" & tipo & "-" & folio & ".xml") '"E:\FAE_eltit\"
entradapdf = confi_rutafae + confi_localempresa + "\DTE\" & tipo & "-" & folio & ".xml"
salidapdf = confi_rutafae + confi_localempresa + "\PDF\" & tipo & "-" & folio & ".pdf"
salidapdf2 = confi_rutafae + confi_localempresa + "\PDF\" & tipo & "-" & folio & "cedi.pdf"
If tipo <> "41" Then
    If copia = 0 Then
        pdf = lib_generapdf + " -d " + entradapdf + " -p " & confi_rutafae & confi_localempresa + "\impresion\FA_estandar.properties -f " & confi_rutafae & confi_localempresa + "\impresion\" + confi_archivopdfnormal + " -o " + salidapdf
    Else
        pdf = lib_generapdf + " -d " + entradapdf + " -p " & confi_rutafae & confi_localempresa + "\impresion\FA_estandar.properties -f " & confi_rutafae & confi_localempresa + "\impresion\" + confi_archivopdfcedible + " -o " + salidapdf2
    End If
Else
    If copia = 0 Then
        pdf = lib_generapdf + " -d " + entradapdf + " -p " & confi_rutafae & confi_localempresa + "\impresion\BO_estandar.properties -f " & confi_rutafae & confi_localempresa + "\impresion\" + "bo_estandar.pdf" + " -o " + salidapdf
    Else
        pdf = lib_generapdf + " -d " + entradapdf + " -p " & confi_rutafae & confi_localempresa + "\impresion\BO_estandar.properties -f " & confi_rutafae & confi_localempresa + "\impresion\" + "bo_estandar.pdf" + " -o " + salidapdf2
    End If
End If



Call Shell(pdf, vbHide)


Call Sleep(5000)
If copia = 0 Then
generapdf = salidapdf
Else
generapdf = salidapdf2

End If

End Function


Public Sub EnviarMail(ByRef Asunto, ByRef mensaje, ByRef Servidor, ByRef MailDestinatario, ByVal NombreDestinatario As String, ByRef ArchivAdjunto, ByVal archivadjunto2 As String)
Dim enviados As String

Set email = New clsSendMail
Screen.MousePointer = vbHourglass
MailDestinatario = LCase(MailDestinatario)

 With email

     ' **************************************************************************
        .SMTPHostValidation = VALIDATE_NONE
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
     ' **************************************************************************
        .AsHTML = False                             'ENVIAR COMO HTML O TEXTO PLANO
        .SMTPHost = confi_servermail     'Servidor                        ' servidor smtp
        .From = confi_mailsalida 'usuario                  'email de quien envia
        .FromDisplayName = leerNombreEmpresa(confi_empresaactiva) 'nombreempresa            'NOMBRE DE QUIEN ENVIA
        .Recipient = MailDestinatario           'email para
        .RecipientDisplayName = NombreDestinatario  'NOMBRE PARA
        '.CcRecipient = usuario                           'EMAIL COPIA
        .CcDisplayName = ""                         'NOMBRE COPIA
        .BccRecipient = MailDestinatario                                     'EMAIL COPIA OCULTA
        .ReplyToAddress = ""                        'responder a otro email
        .Subject = Asunto                           'ASUNTO
        .Message = mensaje 'txtemail.text                    'CUERPO MAIL

If ArchivAdjunto <> "" And archivadjunto2 = "" Then enviados = ArchivAdjunto
If archivadjunto2 <> "" And ArchivAdjunto = "" Then enviados = archivadjunto2
If archivadjunto2 <> "" And ArchivAdjunto <> "" Then enviados = ArchivAdjunto + ";" + archivadjunto2

If enviados <> "" Then .Attachment = enviados 'RUTA ARCHIVO ADJUNTO Trim(txtAttach.Text)
     
     ' **************************************************************************
        .ContentBase = ""
        .EncodeType = 0 'CODIFICACION ARCHIVOS ADJUNTOS
        .Priority = HIGH_PRIORITY                   'PRIORIDAD
        .Receipt = False                            '
        .UseAuthentication = True                   'servidor requiere autenticacion
        .UsePopAuthentication = True           'usar autenticacion del pop
        .UserName = confi_mailsalida 'usuario                         'usuario cuenta correo
        .password = confi_clavemail    'CLAVE                           'clave cuenta correo
        .POP3Host = confi_servermail 'Servidor                        'servidor pop
        .MaxRecipients = 100                        '
         Rem Call Sleep(10000)
        .Send                                       'envia el correo
      
    End With
   
    Screen.MousePointer = vbDefault
End Sub

Public Function leerdatorecepcion(ARCHIVO, loc) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = VENTAS
csql.sql = "select respuesta from " + clientesistema + "fae" + loc + ".sv_envio_libros" + loc
csql.sql = csql.sql & " where ruta='" + Replace(ARCHIVO, "\", "+") + "' "
csql.Execute
leerdatorecepcion = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerdatorecepcion = resultados(0)
    

End If
csql.Close
Set csql = Nothing
End Function
Public Function leerdatorecepcion_recibidos(ARCHIVO, loc) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = VENTAS
csql.sql = "select archivo_recepcion from " + clientesistema + "fae" + loc + ".sv_recepcion_dte" + loc
csql.sql = csql.sql & " where archivo='" + ARCHIVO + "' "
csql.Execute
leerdatorecepcion_recibidos = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerdatorecepcion_recibidos = resultados(0)
    

End If
csql.Close
Set csql = Nothing
End Function


Public Function leerdatorecepcion_dte(track, tipo, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = VENTAS
csql.sql = "select respuestasii from " + clientesistema + "fae" + confi_localempresa + ".sv_dte" + confi_localempresa
csql.sql = csql.sql & " where track='" + track + "' limit 0,1 "
csql.Execute
leerdatorecepcion_dte = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerdatorecepcion_dte = resultados(0)
    

End If
csql.Close
Set csql = Nothing
End Function

Public Function cargarecepciones(ARCHIVO) As Double

Dim s As String
Dim K As Double

Dim cantidad As Double
Dim w As Double
cantidad = Val(leerdatoxml(ARCHIVO, "NroDTE>", 1))

For w = 1 To cantidad
dte_rec(w, 1) = leerdatoxml(ARCHIVO, "TipoDTE>", w)
dte_rec(w, 2) = leerdatoxml(ARCHIVO, "Folio>", w)
dte_rec(w, 3) = leerdatoxml(ARCHIVO, "FchEmis>", w)
dte_rec(w, 4) = leerdatoxml(ARCHIVO, "RUTEmisor>", w)
dte_rec(w, 5) = UCase(leerdatoxml(ARCHIVO, "RznSoc>", w))
dte_rec(w, 6) = leerdatoxml(ARCHIVO, "MntTotal>", w)
Next w
cargarecepciones = cantidad

End Function

Public Function cargaaceptaciones(ARCHIVO) As Double

Dim s As String
Dim K As Double

Dim cantidad As Double
Dim w As Double
cantidad = Val(leerdatoxml(ARCHIVO, "NroDTE>", 1))

For w = 1 To cantidad
dte_rec(w, 1) = leerdatoxml(ARCHIVO, "TipoDTE>", w)
dte_rec(w, 2) = leerdatoxml(ARCHIVO, "Folio>", w)
dte_rec(w, 3) = leerdatoxml(ARCHIVO, "FchEmis>", w)
dte_rec(w, 4) = leerdatoxml(ARCHIVO, "RUTEmisor>", w)
dte_rec(w, 5) = UCase(leerdatoxml(ARCHIVO, "RznSoc>", w))
dte_rec(w, 6) = leerdatoxml(ARCHIVO, "MntTotal>", w)
Next w


End Function






Public Sub GRABARECEPCION(tipo, numero, fecha, rut, nombre, fecharecepcion, nombrearchivo, monto, respuesta_enviada, fecha_respuesta_enviada, acepta_enviada, fecha_acepta_enviada, correo_proveedor, xml, estadodte, glosadte)
    Dim campos(20, 3) As Variant
    
    Dim op As Integer
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "rut"
    campos(4, 0) = "nombre"
    campos(5, 0) = "fecharecepcion"
    campos(6, 0) = "nombrearchivo"
    campos(7, 0) = "monto"
    campos(8, 0) = "respuesta_enviada"
    campos(9, 0) = "fecha_respuesta_enviada"
    campos(10, 0) = "acepta_enviada"
    campos(11, 0) = "fecha_acepta_enviada"
    campos(12, 0) = "correo_proveedor"
    campos(13, 0) = "xml"
    campos(14, 0) = "estadodte"
    campos(15, 0) = "glosadte"
    campos(16, 0) = ""
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = fecha
    campos(3, 1) = rut
    campos(4, 1) = nombre
    campos(5, 1) = fecharecepcion
    campos(6, 1) = nombrearchivo
    campos(7, 1) = monto
    campos(8, 1) = Replace(respuesta_enviada, "'", "")
    campos(9, 1) = fecha_respuesta_enviada
    campos(10, 1) = acepta_enviada
    campos(11, 1) = fecha_acepta_enviada
    campos(12, 1) = correo_proveedor
    campos(13, 1) = Replace(leerxml(confi_rutafae & confi_localempresa & "\dte_recibidos\" + "dte-" + rut + "-" + tipo + "-" + numero + ".xml"), "'", "")
    campos(14, 1) = estadodte
    campos(15, 1) = Replace(glosadte, "'", "")
  
    campos(0, 2) = clientesistema + "fae" & confi_localempresa & ".sv_dte" & confi_localempresa & "_recibidos"
    condicion = ""
    op = 2
    sqlventas.response = campos
    Set sqlventas.conexion = VENTAS
    Call sqlventas.sqlventas(op, condicion)
    Rem Call FIRMADTE(campos(8, 1), tipo, numero)
    Rem Call enviarsii_express(tipo, numero)
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    op = 3
    sqlventas.response = campos
    Set sqlventas.conexion = VENTAS
    Call sqlventas.sqlventas(op, condicion)
    
    
    
End Sub

Public Sub GRABARLIBRO(tipo, fecha, ARCHIVO, periodo)
    Dim campos(20, 3) As Variant
    
    Dim op As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "fecha"
    campos(2, 0) = "ruta"
    campos(3, 0) = "periodo"
    campos(4, 0) = "archivo"
    campos(5, 0) = "periodo2"
    campos(6, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = fecha
    campos(2, 1) = Replace(ARCHIVO, "\", "+")
    campos(3, 1) = periodo
    campos(4, 1) = ""
    campos(5, 1) = Format(periodo, "yyyymm")
    
    
    campos(0, 2) = clientesistema + "fae" + confi_localempresa + ".sv_envio_libros" + confi_localempresa
    condicion = ""
    op = 2
    sqlventas.response = campos
    Set sqlventas.conexion = VENTAS
    Call sqlventas.sqlventas(op, condicion)
    
    
    
    
End Sub


Public Sub enviar_libro(ARCHIVO)
Dim K As Integer
Dim INICIO As Double
Dim final As Double
Dim TIPOENVIO As String
Dim folioen As String

Dim CUENTA As Double
Dim entradafirma As String
Dim salidafirma As String
Dim rutreceptor As String
Dim detalle3 As String
Dim firmaenvio As String
Dim folio As String
Dim respuestasii As String
Dim enviarsii As String
Dim respuestaenvio As String
Dim track As String
Dim track_respuesta As String
Dim archivo2 As String
Dim firmalibro As String
Dim xtipo As String
Dim glosa_sii As String
Dim xtipolb As String

folioen = ultimo_envio

If Mid(ARCHIVO, 18, 2) = "LV" Then
  xtipolb = "5"
  xtipo = "VENTA " & Mid(ARCHIVO, 20, 2) & "-" & Mid(ARCHIVO, 22, 4)
Else
  xtipolb = "6"
  xtipo = "COMPRA " & Mid(ARCHIVO, 20, 2) & "-" & Mid(ARCHIVO, 22, 4)
End If

SALIDASII = confi_rutafae + confi_localempresa + "\libros\timbrado_" + Right(ARCHIVO, 11)

ultimo.COMANDO.text = lib_FirmaLibro & ARCHIVO + " " + certificado + " " + clavecertificado + " " + SALIDASII
firmalibro = lib_FirmaLibro & ARCHIVO + " " + certificado + " " + clavecertificado + " " + SALIDASII


Shell firmalibro

Call Sleep(10000)

respuestasii = confi_rutafae + confi_localempresa + "\respuesta_sii\res_" + folioen + ".xml"


enviarsii = lib_enviarsii & " -p " + SALIDASII + " -c " + certificado + " -s " + clavecertificado + " -o " + respuestasii
Rem enviarsii = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + confi_empresaactiva + "\programas;C:\fae\" + confi_empresaactiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\jargs.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\itext-1.3.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + confi_empresaactiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.EnviarSII "-p " + salidafirma + " -c " + CERTIFICADO + " -s 123 -o " + respuestasii
Shell enviarsii
Call Sleep(20000)




respuestaenvio = leerxmlrecibido(respuestasii)

track = leerdatoxml(respuestaenvio, "TRACKID>", 1)
'************************** ENVIO DE STATUS ***************************
glosa_sii = "ULTIMO LIBRO" & xtipo & " Enviado a las " & Time & " Track:" & track & " Status:(Esperando Correo del Status)"

If track <> "" Then
    Call modificaenviolibro(ARCHIVO, track)
    Call grabar_envio(folioen, Format(fechasistema, "yyyy-mm-dd"), respuestaenvio, track, track_respuesta)
    Call update_adminerp(clientesistema, "facturaelectronica.exe", xtipolb, Format(Now, "yyyy-mm-dd"), "NO", glosa_sii)
End If



End Sub

Sub grabarlogenvio(tipo, folio, fechaenvio, track, estado)
    
        Dim numfic, numficaux As Integer
        Dim cad, cadaux As String
        numfic = FreeFile
            If Len(App.path & "\fae_" & Format(fechasistema, "mmyyyy") & ".txt") = 0 Then
            Open App.path & "\" & ARCHIVO For Output As #numfic
            Close #numfic
            End If
        numfic = FreeFile
        
        Open App.path & "\" & "\fae_" & Format(fechasistema, "mmyyyy") & ".txt" For Append As #numfic
        Do While Not EOF(numfic)
            Line Input #numfic, cad
            cadaux = cad
            Print #numficaux, cad
        Loop
        
        Print #numfic, tipo & "," & folio & "," & fechaenvio & "," & estado & "," & track
    
        Close #numfic
        
End Sub
Public Sub librerias_java()
lib_base = "c:\jre6\bin\java -Dfile.encoding=ISO-8859-1 -classpath " & confi_rutafae & confi_localempresa + "\" + "programas;"
librerias(1) = confi_rutafae + confi_localempresa + "\" + "programas\lib\apache-mime4j-0.6.jar;"
librerias(2) = confi_rutafae + confi_localempresa + "\" + "programas\lib\commons-codec-1.3.jar;"
librerias(3) = confi_rutafae + confi_localempresa + "\" + "programas\lib\commons-httpclient-3.0.jar;"
librerias(4) = confi_rutafae + confi_localempresa + "\" + "programas\lib\commons-logging-1.0.4.jar;"
librerias(5) = confi_rutafae + confi_localempresa + "\" + "programas\lib\httpclient-4.0.jar;"
librerias(6) = confi_rutafae + confi_localempresa + "\" + "programas\lib\httpcore-4.0.1.jar;"
librerias(7) = confi_rutafae + confi_localempresa + "\" + "programas\lib\httpmime-4.0.jar;"
librerias(8) = confi_rutafae + confi_localempresa + "\" + "programas\lib\itext-1.3.jar;"
librerias(9) = confi_rutafae + confi_localempresa + "\" + "programas\lib\jargs.jar;"
librerias(10) = confi_rutafae + confi_localempresa + "\" + "programas\lib\jdom.jar;"
librerias(11) = confi_rutafae + confi_localempresa + "\" + "programas\lib\log4j-1.2.14.jar;"
librerias(12) = confi_rutafae + confi_localempresa + "\" + "programas\lib\not-yet-commons-ssl-0.3.11.jar;"
librerias(13) = confi_rutafae + confi_localempresa + "\" + "programas\lib\OpenLibsDte.jar;"
librerias(14) = confi_rutafae + confi_localempresa + "\" + "programas\lib\xbean.jar;"
librerias(15) = confi_rutafae + confi_localempresa + "\" + "programas\lib\xercesImpl.jar;"
librerias(16) = confi_rutafae + confi_localempresa + "\" + "programas\lib\xfire-all-1.2.6.jar;"
librerias(17) = confi_rutafae + confi_localempresa + "\" + "programas\lib\xmlsec-1.4.0.jar "
lib_java = lib_base
For K = 1 To 17
lib_java = lib_java + librerias(K)
Next K

lib_generapdf = lib_java + "cl." + confi_java + ".dte.GeneraPDF"
lib_enviarsii = lib_java + "cl." + confi_java + ".dte.EnviarSII "
lib_recepcionar = lib_java + "cl." + confi_java + ".dte.Recepcionar "
lib_Firmaenvio = lib_java + "cl." + confi_java + ".dte.FirmaEnvio "
lib_FirmaLibro = lib_java + "cl." + confi_java + ".dte.FirmaLibro "
lib_imprime = "java -jar " + confi_rutafae + confi_localempresa + "\impresion\imprimepdf.jar "
If ExisteArchivo("c:\jre6") = False Then
MsgBox "debe instalar java jre6 que esta en equipo desarrollo en carpeta c:\jre6 "
End
End If

End Sub

Public Sub consultarsii(loc, tipo, numero)
Dim contador As Double

Dim respuesta As String
Dim Pregunta As String


ARCHIVO = leerxmldtecliente(loc, Mid(tipo, 1, 2), numero)
If ARCHIVO = "0" Then
    MsgBox "NO HAY RESPUESTA PARA ESTE ARCHIVO", vbInformation, "ATENCION"
    Exit Sub
End If
comi = Chr(34)
Rem DETALLE = "<?xml-stylesheet type=" + comi + "text/xsl" + comi + " href=" + comi + "visualizador3.xsl" + comi + "?>" + ARCHIVO
DETALLE = ARCHIVO
cadena = DETALLE
For K = 1 To Len(DETALLE)
If Asc(Mid(DETALLE, K, 1)) > 128 And Mid(DETALLE, K, 1) <> "—" Then
cadena = Replace(cadena, Mid(DETALLE, K, 1), "")
End If

Next K
DETALLE = cadena
DETALLE = Replace(DETALLE, "•", "N")
DETALLE = Replace(DETALLE, "—", "#209;")
DETALLE = Replace(DETALLE, "ß", " ")
DETALLE = Replace(DETALLE, "∫", " ")
DETALLE = Replace(DETALLE, "∞", " ")
DETALLE = Replace(DETALLE, "&", "&amp;")
DETALLE = Replace(DETALLE, "¯", " ")
DETALLE = Replace(DETALLE, ",", ".")
DETALLE = Replace(DETALLE, "*", "x")
DETALLE = Replace(DETALLE, "¥", "")
DETALLE = Replace(DETALLE, "«", "")
DETALLE = Replace(DETALLE, "Ô", "")
DETALLE = Replace(DETALLE, "Ô", "")


Close 20
        Pregunta = "u:\FAE_ADMIN\consulta_sii.xml"
        respuesta = "u:\FAE_ADMIN\consulta_sii_" + loc + "_" + Mid(tipo, 1, 2) + "_" + numero + ".xml"
        Open Pregunta For Output As #20
        Print #20, DETALLE
        Close 20
     
consultasii = "U:\FAE_ADMIN\consultaEstado.bat " + certificado + " " + clavecertificado + " " + Pregunta + " " + respuesta

Shell consultasii
contador = 0
20:

If ExisteArchivo(respuesta) = True Then
Rem Sleep (1000)
respuestaenvio = leerxmlrecibido(respuesta)

glosa_estado = leerdatoxml(respuestaenvio, "SII:GLOSA_ESTADO>", 1)

If glosa_estado <> "" Then
    Call modificasii(loc, Mid(tipo, 1, 2), numero, "1", glosa_estado)
End If



Else
contador = contador + 1
        If contador > 100000 Then
        
        Exit Sub

        End If

GoTo 20:
End If







End Sub

Public Sub modificasii(loc, tipo, folio, Status, glosa)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = VENTAS
        csql.sql = "UPDATE " + clientesistema + "fae" + loc + ".sv_dte" + loc + " "
        csql.sql = csql.sql & "set status='" + Status + "', glosa_sii='" + glosa + "' WHERE tipo='" + tipo + "' and numero='" + folio + "' "
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub




Public Sub EnviarMail22(ByRef Asunto, ByRef mensaje, ByRef Servidor, ByRef MailDestinatario, ByVal NombreDestinatario As String, ByRef ArchivAdjunto, _
                        ByVal archivadjunto2 As String)

Dim enviados As String
On Error GoTo error
Screen.MousePointer = vbHourglass
MailDestinatario = LCase(MailDestinatario)
'Call empresadte(empresaactiva)
'rz 05-12-2017 '
'funcion adaptada para enviar con cuenta de eltit

If LeerCuentaAlternativa = True Then
'        confi_mailsalida = email_cuenta_usuario
'        confi_clavemail = email_cuenta_clave
'        confi_servermail = email_cuenta_server
End If
 
If ArchivAdjunto <> "" And archivadjunto2 = "" Then enviados = ArchivAdjunto
If archivadjunto2 <> "" And ArchivAdjunto = "" Then enviados = archivadjunto2
If archivadjunto2 <> "" And ArchivAdjunto <> "" Then enviados = ArchivAdjunto + ";" + archivadjunto2



Dim iMsg As Object
Dim iConf As Object
Dim strbody  As String
Dim Flds As Variant
Dim comi As String
Dim puertosalida As Double
Dim destinatario As String
'destinatario = "aalarcon@eltit.cl; rlzurita@gmail.com"
 
comi = Chr(34)
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
destinatario = Replace(destinatario, "<", "")
destinatario = Replace(destinatario, ">", "")
iConf.Load -1
 puertosalida = 465

Set Flds = iConf.Fields
With Flds
    .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = email_cuenta_usuario
    .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = email_cuenta_clave
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = email_cuenta_server
    .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = puertosalida
    .Update
End With

With iMsg
Set .Configuration = iConf
.To = MailDestinatario
.cc = ""
.BCC = "" 'aalarcon@eltit.cl; raulzurita@adminerp.cl"
 
.From = comi & comi & comi & nombreempresa & comi & " <" & email_cuenta_usuario & ">"
.Subject = Asunto
.TextBody = mensaje

If enviados <> "" Then .AddAttachment enviados


.Send
Call grabar_envio_correo("", email_cuenta_usuario, MailDestinatario)
End With

error:
End Sub


Public Function LeerCuentaAlternativa2() As Boolean
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "cuenta"
    campos(1, 0) = "servidor"
    campos(2, 0) = "contraseÒa"
    campos(3, 0) = ""
    campos(4, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.maestro_correos_cuentas"
    
    condicion = "empresa = '' "
    op = 5
    sqlventas.response = campos
    Set sqlventas.conexion = VENTAS
    Call sqlventas.sqlventas(op, condicion)
    
    If sqlventas.Status = 0 Then
        email_cuenta_usuario = sqlventas.response(0, 3)
        email_cuenta_server = sqlventas.response(1, 3)
        email_cuenta_clave = sqlventas.response(2, 3)
        
        confi_mailsalida = email_cuenta_usuario
        confi_clavemail = email_cuenta_clave
        confi_servermail = email_cuenta_server
        
        LeerCuentaAlternativa = True
        
    End If
End Function




Public Sub grabar_envio_correo(id, de, para)
Dim recepciona As String
        Dim BDcte As String
    
        campos(0, 0) = "empresa"
        campos(1, 0) = "fecha"
        campos(2, 0) = "hora"
        campos(3, 0) = "programa"
        campos(4, 0) = "cuenta"
        campos(5, 0) = "destino"
        campos(6, 0) = ""
        
        campos(0, 1) = empresaactiva
        campos(1, 1) = Format(Now, "yyyy-mm-dd")
        campos(2, 1) = Format(Now, "hh:mm:ss")
        campos(3, 1) = App.EXEName
        campos(4, 1) = de
        campos(5, 1) = para
        
        campos(0, 2) = clientesistema & "fae.sv_envios_correo"
    
        condicion = ""

        op = 2
        sqlventas.response = campos
        Set sqlventas.conexion = VENTAS
        Call sqlventas.sqlventas(op, condicion)
        
      
       
End Sub
