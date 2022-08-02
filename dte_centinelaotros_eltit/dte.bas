Attribute VB_Name = "DTE"
Public dte_tipodte As String
Public tipodoc(3) As Double
Public CE_PTO As Boolean
Public ncref As Boolean
Public dte_referencia As String
Public dte_caso As String
Public dte_tiporef As String
Public dte_folio As String
Public dte_fecha As String
Public dte_indtraslado As String
Public dte_servicio As String
Public dte_e_rut As String
Public dte_e_nombre As String
Public dte_e_acti As String
Public dte_e_direccion As String
Public dte_ivaretenido As Double
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
Public LOCAL_PROCESO As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 
 Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
    Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
    Public Type PROCESSENTRY32
        dwSize As Long
        cntUsage As Long
        th32ProcessID As Long
        th32DefaultHeapID As Long
        th32ModuleID As Long
        cntThreads As Long
        th32ParentProcessID As Long
        pcPriClassBase As Long
        dwFlags As Long
        szExeFile As String * 260
    End Type
 Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Const PROCESS_TERMINATE = &H1
    Const PROCESS_CREATE_THREAD = &H2
    Const PROCESS_VM_OPERATION = &H8
    Const PROCESS_VM_READ = &H10
    Const PROCESS_VM_WRITE = &H20
    Const PROCESS_DUP_HANDLE = &H40
    Const PROCESS_CREATE_PROCESS = &H80
    Const PROCESS_SET_QUOTA = &H100
    Const PROCESS_SET_INFORMATION = &H200
    Const PROCESS_QUERY_INFORMATION = &H400
    Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Const SYNCHRONIZE = &H100000
    Const PROCESS_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
     
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
 
 
Private Const INFINITE = -1&



'Public Sub GENERADTE(loc, caja, tipo, numero, fecha, rut, neto, iva, total, ref, vin, lic, iha, ica, sucursal, tipo_despacho, indicador_traslado, local_destino)
'
'Dim dato(100) As String
'Dim dato2(100) As String
'Dim datos(100, 10) As String
'Dim dato3(100) As String
'Dim comi As String
'Dim cabeza As String
'Dim detalle As String
'Dim K As Integer
'Dim pasada As String
'Dim i As Integer
'Dim FOLIOS As String
'Dim salida As String
'Dim entrada As String
'Dim FIRMA As String
'Dim entradapdf As String
'Dim salidapdf As String
'Dim pdf As String
'Dim enviasii As String
'Dim firmaenvio As String
'Dim entradafirma As String
'Dim salidafirma As String
'Dim entradaenvio As String
'Dim SALIDAENVIO As String
'Dim cer As String
'Dim lin1 As Double
'Dim cadena As String
'Dim tiponc As String
'Dim numeronc As String
'
'xml.Encoding = "ISO-8859-1"
'
'
'comi = Chr(34)
'
'If tipo = "FV" Then dte_tipodte = "33": Rem FACTURA ELECTRONICA
'If tipo = "FE" Then dte_tipodte = "34": Rem FACTURA NO AFECTA
'If tipo = "FC" Then dte_tipodte = "46": Rem FACTURA DE COMPRA ELECTRONICA
'If Mid(tipo, 1, 1) = "G" Then dte_tipodte = "52": Rem GUIA DE DESPACHO ELECTRONICA:
'If tipo = "ND" Then dte_tipodte = "56": Rem NOTA DEBITO ELECTRONICA
'If tipo = "NF" Or tipo = "NB" Then dte_tipodte = "61": Rem NOTA DE CREDITO ELECTRONICA
'dte_fecha = Format(fecha, "dd-mm-yyyy")
'empresa
'
'Call empresadte(codigoCONTABLE)
'Call clientedte(rut, sucursal, local_destino)
'
''TIPO G4 ES DEVOLUCION A PROVEEDOR Y NO GUIA INTERNA ENTONCES LEE DATOS DE ENVIO DESDE MAESTRO DE CLIENTE
'If tipo = "G4" Then Call clientedtePROVEEDOR(rut, "0")
'
'
'dte_folio = leerfoliodte(empresaActiva, dte_tipodte)
' Rem dte_folio = 1503
'' dte_folio = 36001
'
'    If leerfolioautorizado(empresaActiva, dte_tipodte, dte_folio) = False Then
'    Unload electro04
'    Exit Sub
'    End If
'
'    For K = 1 To 100
'    dato(K) = ""
'    Next K
'
'
'    dato(1) = "<DTE version=" + comi + "1.0" + comi + " >"
'    dato(2) = "<Documento ID=" + comi + "ENVIOFOLIO_" & dte_folio & "T" + dte_tipodte + comi + ">"
'    dato(3) = "<Encabezado>"
'
'    dte_fecha = Format(dte_fecha, "yyyy-mm-dd")
'    dato(4) = "<IdDoc>"
'    dato(5) = "<TipoDTE>" + dte_tipodte + "</TipoDTE>"
'    dato(6) = "<Folio>" & dte_folio & "</Folio>"
'    dato(7) = "<FchEmis>" + dte_fecha + "</FchEmis>"
'    Rem indicador no rebaja <IndNoRebaja>"
'    Rem tipodespacho <TipoDespacho>"
'    If dte_tipodte = "52" Then
'    Rem dato(8) = "<TipoDespacho>" & tipo_despacho & "</Tipo_despacho>"
'    dato(8) = ""
'    dato(9) = "<IndTraslado>" & indicador_traslado & "</IndTraslado>"
'    End If
'Rem 1 = venta
'Rem 2  = venta por efectuar
'Rem 3 = consignaciones
'Rem 4 = entrega gratuita
'Rem 5 = traslados internos
'Rem 6 = otros traslado no venta
'
'Rem tipode impresion <TpoImpresion> n = normal t=ticket
'
'Rem  "<FmaPago>" & PAGO & "</FmaPago>" solo servicios
'
'Rem dato(10) = "<FchCancel>" + Format(fechaemision, "yyyy-mm-dd") + "</FchCancel>"
'Rem dato(11) = "<MedioPago>EF</MedioPago>"
'Rem dato(12) = "<FchVenc>" + Format(fechaemision, "yyyy-mm-dd") + "</FchVenc>"
'    dato(10) = ""
'    dato(11) = ""
'    dato(12) = ""
'    dato(13) = "</IdDoc>"
'    dato(14) = "<Emisor>"
'    dato(15) = "<RUTEmisor>" + dte_e_rut + "</RUTEmisor>"
'    dato(16) = "<RznSoc>" + dte_e_nombre + "</RznSoc>"
'    dato(17) = "<GiroEmis>" + Mid(dte_e_giro, 1, 80) + "</GiroEmis>"
'    dato(18) = "<Acteco>" + dte_e_acti + "</Acteco>"
'    dato(19) = "<DirOrigen>" + Mid(dte_e_direccion, 1, 60) + "</DirOrigen>"
'    dato(20) = "<CmnaOrigen>" + dte_e_comuna + "</CmnaOrigen>"
'    dato(21) = "<CiudadOrigen>" + dte_e_ciudad + "</CiudadOrigen>"
'    If dte_e_sucursal <> "" Then
'    dato(22) = "<cdgSIISucur>" + dte_e_sucursal + "</CdgSIISucur>"
'    End If
'    dato(23) = "<CdgVendedor>" + "ORIGEN " + empresaActiva + " " + DIRECCIONEMPRESA + "</CdgVendedor>"
'
'If tipo = "GD" Then
'    dato(23) = "<CdgVendedor>" + "ORIGEN " + empresaActiva + " " + DIRECCIONEMPRESA + "</CdgVendedor>"
'End If
'If tipo = "G1" Then
'    dato(23) = "<CdgVendedor>TIPO DOCUMENTO ME </CdgVendedor>"
'End If
'If tipo = "G2" Then
'    dato(23) = "<CdgVendedor>TIPO DOCUMENTO HU </CdgVendedor>"
'End If
'If tipo = "G3" Then
'    dato(23) = "<CdgVendedor>TIPO DOCUMENTO CI </CdgVendedor>"
'End If
'If tipo = "G4" Then
'    dato(23) = "<CdgVendedor>TIPO DOCUMENTO DP </CdgVendedor>"
'End If
'
'    dato(24) = "</Emisor>"
'    dato(25) = "<Receptor>"
'    dato(26) = "<RUTRecep>" + dte_r_rut + "</RUTRecep>"
'    dato(27) = "<RznSocRecep>" + dte_r_nombre + "</RznSocRecep>"
'    dato(28) = "<GiroRecep>" + Mid(dte_r_giro, 1, 40) + "</GiroRecep>"
'    dato(29) = "<DirRecep>" + Mid(dte_r_direccion, 1, 70) + "</DirRecep>"
'    dato(30) = "<CmnaRecep>" + Mid(dte_r_comuna, 1, 20) + "</CmnaRecep>"
'    dato(31) = "<CiudadRecep>" + Mid(dte_r_ciudad, 1, 20) + "</CiudadRecep>"
'    dato(33) = "</Receptor>"
'    ncref = False
'
'    Rem If dte_tipodte = 61 And dte_folio = 4 Then ncref = True
'    autorizadte = True
'
'    Call leerdetalledte(tipo, numero, caja, Format(fecha, "yyyy-mm-dd"), total)
'If autorizadte = False Then
'Exit Sub
'End If
''DIFE = dte_neto - CDbl(neto)
''If DIFE > 100 Or DIFE < 100 Then
''MsgBox tipo + " " + numero + " " & dte_neto & " " + neto
''Exit Sub
''End If
'
'
'    dato(42) = "<Totales>"
'    dato(43) = "<MntNeto>" & dte_neto & "</MntNeto>"
'    If dte_exento <> 0 Then
'    dato(44) = "<MntExe>" & dte_exento & "</Mntexe>"
'    End If
'
'
'    dato(45) = "<TasaIVA>19</TasaIVA>"
'    dato(46) = "<IVA>" & dte_iva & "</IVA>"
'    Rem impuesto
'
'If dte_car + dte_har + dte_ref + dte_vin + dte_lic + dte_cer <> 0 Then
'    If dte_car <> 0 Then
'    dato(47) = "<ImptoReten>"
'    dato(48) = "<TipoImp>18</tipoImp"
'    dato(49) = "<TasaImp>5</TasaImp>"
'    dato(50) = "<MontoImp>" & dte_car & "</MontoImp>"
'    dato(51) = "</ImptoReten>"
'    End If
'    If dte_har <> 0 Then
'    dato(52) = "<ImptoReten>"
'    dato(53) = "<TipoImp>19</tipoImp"
'    dato(54) = "<TasaImp>12</TasaImp>"
'    dato(55) = "<MontoImp>" & dte_har & "</MontoImp>"
'    dato(56) = "</ImptoReten>"
'    End If
'    If dte_lic <> 0 Then
'    dato(57) = "<ImptoReten>"
'    dato(58) = "<TipoImp>24</tipoImp"
'    dato(59) = "<TasaImp>27</TasaImp>"
'    dato(60) = "<MontoImp>" & dte_lic & "</MontoImp>"
'    dato(61) = "</ImptoReten>"
'    End If
'    If dte_vin <> 0 Then
'    dato(62) = "<ImptoReten>"
'    dato(63) = "<TipoImp>25</tipoImp"
'    dato(64) = "<TasaImp>15</TasaImp>"
'    dato(65) = "<MontoImp>" & dte_vin & "</MontoImp>"
'    dato(66) = "</ImptoReten>"
'    End If
'    cer = "0"
'    If dte_cer <> 0 Then
'    dato(67) = "<ImptoReten>"
'    dato(68) = "<TipoImp>26</tipoImp"
'    dato(69) = "<TasaImp>15</TasaImp>"
'    dato(70) = "<MontoImp>" & dte_cer & "</MontoImp>"
'    dato(71) = "</ImptoReten>"
'    End If
'    If dte_ref <> 0 Then
'    dato(72) = "<ImptoReten>"
'    dato(73) = "<TipoImp>27</tipoImp"
'    dato(74) = "<TasaImp>13</TasaImp>"
'    dato(75) = "<MontoImp>" & dte_ref & "</MontoImp>"
'    dato(76) = "</ImptoReten>"
'    End If
'End If
'
'dato(77) = "<MntTotal>" & dte_total & "</MntTotal>"
'dato(78) = "</Totales>"
'dato(79) = "</Encabezado>"
'
'For K = 1 To 100
'If dato(K) <> "" Then
'cabeza = cabeza + dato(K)
'End If
'Next K
''datos(1, 1) = "1"
''datos(1, 2) = "INT1"
''datos(1, 3) = "011"
''datos(1, 4) = "Parlantes Multimedia 180 w"
''datos(1, 5) = "0"
''datos(1, 6) = "20"
''datos(1, 7) = "4500"
''datos(1, 8) = "90000"
'detalle = ""
'For K = 1 To 40
'dato2(K) = ""
'dato3(K) = ""
'
'Next K
'
'If ncref = False Then
'    For K = 1 To totallineas
'    dato2(1) = "<Detalle>"
'    dato2(2) = "<NroLinDet>" & K & "</NroLinDet>"
'    dato2(3) = "<CdgItem>"
'    dato2(4) = "<TpoCodigo>" + "EAN13" + "</TpoCodigo>"
'    dato2(5) = "<VlrCodigo>" + matrixdte(K, 1) + "</VlrCodigo>"
'    dato2(6) = "</CdgItem>"
'        If matrixdte(K, 0) = "1" Then
'        dato2(7) = "<IndExe>1</IndExe>"
'        End If
'        dato2(8) = "<NmbItem>" + Replace(matrixdte(K, 3), "+CHR(34)+", "&quot") + "</NmbItem>"
''        If matrixdte(K, 2) <> "0" Then
''        dato2(9) = "<QtyItem>" + Replace(matrixdte(K, 2), ",", ".") + "</QtyItem>"
''        dato2(10) = "<PrcItem>" + Replace(matrixdte(K, 4), ",", ".") + "</PrcItem>"
''        End If
'
'        If matrixdte(K, 2) <> "0" Then
'        dato2(9) = "<QtyItem>" + Replace(matrixdte(K, 2), ",", ".") + "</QtyItem>"
'        End If
'
'        If matrixdte(K, 4) <> "0" Then
'        dato2(10) = "<PrcItem>" + Replace(matrixdte(K, 4), ",", ".") + "</PrcItem>"
'        End If
'
'
'        If Val(matrixdte(K, 5)) <> 0 And Val(matrixdte(K, 8)) <> 0 Then
'        dato2(11) = "<DescuentoPct>" + matrixdte(K, 5) + "</Descuentopct>"
'        dato2(12) = "<DescuentoMonto>" & Replace(matrixdte(K, 8), ",", ".") & "</DescuentoMonto>"
'        dato2(13) = "<SubDscto>"
'        dato2(14) = "<TipoDscto>" + "%" + "</TipoDscto>"
'        dato2(15) = "<ValorDscto>" & Replace(matrixdte(K, 5), ",", ".") & "</ValorDscto>"
'        dato2(16) = "</SubDscto>"
'Else
'        dato2(11) = ""
'        dato2(12) = ""
'        dato2(13) = ""
'        dato2(14) = ""
'        dato2(15) = ""
'        dato2(16) = ""
'
'        End If
'
'        If matrixdte(K, 7) <> "" Then
'        dato2(17) = "<CodImpAdic>" + matrixdte(K, 7) + "</CodImpAdic>"
'        Else
'        dato2(17) = ""
'        End If
'
'    dato2(18) = "<MontoItem>" + matrixdte(K, 6) + "</MontoItem>"
'
'       dato2(19) = "</Detalle>"
'
'
'For i = 1 To 20
'    If dato2(i) <> "" Then
'    pasada = pasada + dato2(i)
'    End If
'Next i
'Next K
'End If
'detalle = detalle & pasada
'lin1 = 0
'
'    If descuento1 + descuento2 <> 0 Then
'        If descuento1 <> 0 Then
'        lin1 = lin1 + 1
'        dato3(2) = "<DscRcgGlobal>"
'        dato3(3) = "<NroLinDR>" & lin1 & "</NroLinDR>"
'        dato3(4) = "<TpoMov>D</TpoMov>"
'        dato3(5) = "<TpoValor>%</tpoValor>"
'        dato3(6) = "<ValorDR>" & dte_descuento & "</ValorDR>"
'        dato3(7) = "</DscRcgGlobal>"
'    End If
'
''    If descuento2 <> 0 Then
''        lin1 = lin1 + 1
''        dato3(8) = "<DscRcgGlobal>"
''        dato3(9) = "<NroLinDR>" & lin1 & "</NroLinDR>"
''        dato3(10) = "<TpoMov>D</TpoMov>"
''        dato3(11) = "<TpoValor>%</tpoValor>"
''        dato3(12) = "<ValorDR>" & dte_descuento & "</ValorDR>"
''        dato3(13) = "<IndExeDR>" & "1" & "</IndExeDR>"
''        dato3(14) = "</DscRcgGlobal>"
''    End If
'    End If
'
'If dte_tipodte = 52 Then
'dte_caso = dte_folio
'dato3(15) = "<Referencia>"
'dato3(16) = "<NroLinRef>1</NroLinRef>"
'dato3(17) = "<TpoDocRef>52</TpoDocRef>"
'dato3(18) = "<FolioRef>" + dte_folio + "</FolioRef>"
'dato3(19) = "<FchRef>" + Format(fecha, "yyyy-mm-dd") + "</FchRef>"
'dato3(20) = "<CodRef>3</CodRef>"
'dato3(21) = "<RazonRef>" + "DESP.:" + leerNombreEmpresa(local_destino) & "FOL:" & numero & "</RazonRef>"
'dato3(22) = "</Referencia>"
'
'
'End If
'
'
'
'pasada = ""
'For i = 1 To 40
'If dato3(i) <> "" Then
'pasada = pasada + dato3(i)
'End If
'
'Next i
'
'detalle = detalle + pasada + "</Documento></DTE>"
'detalle = cabeza + detalle
'cadena = detalle
'
'For K = 1 To Len(detalle)
'    If Asc(Mid(detalle, K, 1)) > 128 And Mid(detalle, K, 1) <> "Ñ" Then
'        cadena = Replace(cadena, Mid(detalle, K, 1), "")
'    End If
'Next K
'detalle = cadena
'
'For K = 1 To Len(detalle)
'    If Asc(Mid(detalle, K, 1)) < 32 Then
'        cadena = Replace(cadena, Mid(detalle, K, 1), "")
'    End If
'Next K
'detalle = cadena
'
'detalle = Replace(detalle, "¥", "N")
'detalle = Replace(detalle, "Ñ", "N")
'detalle = Replace(detalle, "§", " ")
'detalle = Replace(detalle, "º", " ")
'detalle = Replace(detalle, "°", " ")
'detalle = Replace(detalle, "&", "&amp;")
'detalle = Replace(detalle, "ø", " ")
'detalle = Replace(detalle, ",", ".")
'detalle = Replace(detalle, "*", "x")
'detalle = Replace(detalle, "´", "")
'detalle = Replace(detalle, "Ç", "")
'detalle = Replace(detalle, "ï", "")
'
'
'caf = leerrutacaf(dte_tipodte, dte_folio)
'If ExisteArchivo(caf) = False Then
'MsgBox "ruta " + caf + " no esta instalado"
'End
'End If
'
'entrada = "C:\FAE\" + empresaActiva + "\DTE\" & dte_tipodte & "_" & dte_folio & ".xml"
'salida = "C:\FAE\" + empresaActiva + "\DTE\firmado_" & dte_tipodte & "_" & dte_folio & ".xml"
'
'xml.LoadXml detalle
'
'
'Call xml.SaveXml(entrada)
'
'Close 20
'Close 30
'
'Open entrada For Input As #20
'
'Open entrada + "2" For Output As #30
'While EOF(20) = False
'
'Line Input #20, ss
'
'
'Rem ariel  quita espacio antes del sigono pregunta en el encabezado del xml
'
'If InStr(ss, " ?>") Then
'ss = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>"
'End If
'Rem hasta aqui
'
'If ss <> "" Then
'Print #30, ss
'End If
'
'Wend
'Close 20
'Close 30
'
'
'
'
'FIRMA2 = entrada + "2" + " " + salida + " " + caf + " " + certificado
'Rem FIRMA2 = "-a " + caf + " -p " + entrada + " -c " + CERTIFICADO + " -s 123 -o " + salida
'Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + empresaActiva + "\programas;C:\fae\" + empresaActiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + empresaActiva + "\programas\lib\jargs.jar;C:\fae\" + empresaActiva + "\programas\lib\itext-1.3.jar;C:\fae\" + empresaActiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + empresaActiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.FirmaDTE " + FIRMA2
'Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;C:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.FirmaDTE " + FIRMA2
'FIRMA = "c:\fae\programas\firmadte.bat " + FIRMA2
'
'Shell FIRMA
'Call Sleep(10000)
'detalle = leerxml(salida)
'5 If detalle = "" Then Exit Sub
'Call GRABADTE(dte_tipodte, dte_folio, Format(fecha, "yyyy-mm-dd"), loc, tipo, numero, Format(fecha, "yyyy-mm-dd"), caja, detalle, dte_r_rut, dte_r_nombre, dte_total)
'
'Rem enviasii = "c:\fae\programas\enviasii.bat " + entradaenvio + " " + certificado + " " + salidaenvio
'Rem enviasii = "c:\fae\programas\enviasii.bat " + entradaenvio + " " + certificado + " " + salidaenvio
'
'
'Rem datos para firmar envio
'Rem detalle = timbrafactura(rutemisor, RUTENVIA, rutreceptor, Format(fechaemision, "dd-mm-yyyy"), "0", numero, TIPO) + leerxml(salida) + "</SetDTE></EnvioDTE>"
'Rem Call xml.LoadXML(detalle)
'Rem Call xml.SaveXml(entradafirma)
'
'Rem Shell firmaenvio
'Rem Shell enviasii
'Rem Kill entrada
'Rem Kill salida
'descuento1 = 0
'descuento2 = 0
' Call imprimelectronica(dte_tipodte, dte_folio)
'
'
'End Sub
Public Sub GENERADTE(loc, caja, TIPO, numero, fecha, rut, neto, iva, total, ref, vin, lic, iha, ica, sucursal, contacto, exento)

Dim dato(100) As String
Dim dato2(100) As String
Dim datos(100, 10) As String
Dim dato3(100) As String
Dim comi As String
Dim cabeza As String
Dim detalle As String
Dim K As Integer
Dim pasada As String
Dim I As Integer
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
Dim tiponc As String
Dim numeronc As String
Dim datos2 As Variant
Dim glosafactura As String
Dim pago As String

xml.Encoding = "ISO-8859-1"


comi = Chr(34)

If TIPO = "FV" Then dte_tipodte = "33": Rem FACTURA ELECTRONICA
If TIPO = "EX" Then dte_tipodte = "34": Rem FACTURA NO AFECTA
If TIPO = "FC" Then dte_tipodte = "46": Rem FACTURA DE COMPRA ELECTRONICA
If TIPO = "GD" Then dte_tipodte = "52": Rem GUIA DE DESPACHO ELECTRONICA:
If TIPO = "ND" Then dte_tipodte = "56": Rem NOTA DEBITO ELECTRONICA
If TIPO = "NF" Or TIPO = "NB" Then dte_tipodte = "61": Rem NOTA DE CREDITO ELECTRONICA
'dte_tipodte = tipo
dte_fecha = Format(fecha, "dd-mm-yyyy")
empresa

Call empresadte(codigoCONTABLE)
Call clientedte(rut, sucursal)

dte_folio = leerfoliodte(LOCAL_PROCESO, dte_tipodte)
       'dte_folio = 5

    If leerfolioautorizado(LOCAL_PROCESO, dte_tipodte, dte_folio) = False Then
    'Unload electro04
    MsgBox "FOLIO DE DOCUMENTO TIPO " & dte_tipodte & " NO AUTORIZADO POR SII EN LOCAL " & loc & " numero " + dte_folio
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
    If esempresarelacionada(rut) = True Then
         pago = "1"
    Else
         pago = leerformadepago(loc, TIPO, numero, fecha, caja)
    End If

    dato(10) = "<FmaPago>" & pago & "</FmaPago>" ' 1 contado 2 credito 3 gratis

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
    If contacto <> "" Then
        dato(29) = "<Contacto>" + contacto + "</Contacto>"
    End If
    
    dato(30) = "<DirRecep>" + Mid(dte_r_direccion, 1, 70) + "</DirRecep>"
    dato(31) = "<CmnaRecep>" + Mid(dte_r_comuna, 1, 20) + "</CmnaRecep>"
    dato(32) = "<CiudadRecep>" + Mid(dte_r_ciudad, 1, 20) + "</CiudadRecep>"
    dato(33) = "</Receptor>"
    ncref = False

    Rem If dte_tipodte = 61 And dte_folio = 4 Then ncref = True
    autorizadte = True
    
    Call leerdetalledte(TIPO, numero, caja, Format(fecha, "yyyy-mm-dd"), 0, exento)
    
    
If autorizadte = False Then
Exit Sub
End If


    dato(42) = "<Totales>"
    dato(43) = "<MntNeto>" & dte_neto & "</MntNeto>"
    
   
    
    If dte_exento <> 0 Then
    dato(44) = "<MntExe>" & dte_exento & "</Mntexe>"
    End If

    If dte_iva <> 0 Then
    dato(45) = "<TasaIVA>19</TasaIVA>"
    dato(46) = "<IVA>" & dte_iva & "</IVA>"
    End If
   
    Rem impuesto

If dte_car + dte_har + dte_ref + dte_vin + dte_lic + dte_cer <> 0 + dte_ivaretenido <> 0 Then
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
    If dte_ivaretenido <> 0 Then ' agregado para factura de compra
        dato(77) = "<ImptoReten>"
        dato(78) = "<TipoImp>15</tipoImp"
        dato(79) = "<TasaImp>19</TasaImp>"
        dato(80) = "<MontoImp>" & dte_ivaretenido & "</MontoImp>"
        dato(81) = "</ImptoReten>"
    End If
End If

dato(82) = "<MntTotal>" & dte_total & "</MntTotal>"
dato(83) = "</Totales>"
dato(84) = "</Encabezado>"

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
detalle = ""
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
    dato2(9) = "<DscItem>" + Replace(Mid(matrixdte(K, 9), 1, 1000), "+CHR(34)+", "&quot") + "</DscItem>"
    
    If matrixdte(K, 2) <> "0" Then
        dato2(10) = "<QtyItem>" + Replace(matrixdte(K, 2), ",", ".") + "</QtyItem>"
        dato2(11) = "<PrcItem>" + Replace(matrixdte(K, 4), ",", ".") + "</PrcItem>"
    End If
    
    If Val(matrixdte(K, 5)) <> 0 And Val(matrixdte(K, 8)) <> 0 Then
        dato2(12) = "<DescuentoPct>" + matrixdte(K, 5) + "</Descuentopct>"
        dato2(13) = "<DescuentoMonto>" & Replace(matrixdte(K, 8), ",", ".") & "</DescuentoMonto>"
        dato2(14) = "<SubDscto>"
        dato2(15) = "<TipoDscto>" + "%" + "</TipoDscto>"
        dato2(16) = "<ValorDscto>" & Replace(matrixdte(K, 5), ",", ".") & "</ValorDscto>"
        dato2(17) = "</SubDscto>"
    Else
        dato2(12) = ""
        dato2(13) = ""
        dato2(14) = ""
        dato2(15) = ""
        dato2(16) = ""
        dato2(17) = ""

     End If
        
    If matrixdte(K, 7) <> "" Then
        dato2(18) = "<CodImpAdic>" + matrixdte(K, 7) + "</CodImpAdic>"
    Else
        dato2(18) = ""
    End If
  
    dato2(19) = "<MontoItem>" + matrixdte(K, 6) + "</MontoItem>"
       
    dato2(20) = "</Detalle>"


For I = 1 To 21
    If dato2(I) <> "" Then
    pasada = pasada + dato2(I)
    End If
Next I
Next K
End If
detalle = detalle & pasada
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
    'agregado el 09-06-2014 pedido por carozzi
'    If dte_tipodte = 33 Then
'        datos2 = Split(contacto, "/")
'        dte_caso = dte_folio
'        dato3(15) = "<Referencia>"
'        dato3(16) = "<NroLinRef>1</NroLinRef>"
'        dato3(17) = "<TpoDocRef>801</TpoDocRef>"
'        dato3(18) = "<FolioRef>" & datos2(0) & "</FolioRef>"
'        dato3(19) = "<FchRef>" & Format(fechasistema, "yyyy-mm-dd") & "</FchRef>"
'        dato3(20) = "</Referencia>"
'
'        dato3(21) = "<Referencia>"
'        dato3(22) = "<NroLinRef>2</NroLinRef>"
'        dato3(23) = "<TpoDocRef>802</TpoDocRef>"
'        dato3(24) = "<FolioRef>" & datos2(1) & "</FolioRef>"
'        dato3(25) = "<FchRef>" & Format(fechasistema, "yyyy-mm-dd") & "</FchRef>"
'        dato3(26) = "</Referencia>"
'    End If

    If dte_tipodte = 34 Or dte_tipodte = 46 Or dte_tipodte = 52 Then
        dte_caso = dte_folio
        dato3(15) = "<Referencia>"
        dato3(16) = "<NroLinRef>1</NroLinRef>"
        dato3(17) = "<TpoDocRef>" + dte_tipodte + "</TpoDocRef>"
        dato3(18) = "<FolioRef>" + dte_folio + "</FolioRef>"
        dato3(19) = "<FchRef>" + Format(fecha, "yyyy-mm-dd") + "</FchRef>"
        dato3(20) = "<CodRef>3</CodRef>"
        dato3(21) = "<RazonRef>" + "CAJA:" + caja + " NI:" & numero & "</RazonRef>"
        dato3(22) = "</Referencia>"
    End If
    
    
    'validar si no tiene
    
    
    
        If (dte_tipodte = 33) And (LeedatoOCproveedor(TIPO, numero, loc, "LPAD(ordencompra,10,'0')") <> "0000000000") Then
        dte_caso = dte_folio
        dato3(15) = "<Referencia>"
        dato3(16) = "<NroLinRef>1</NroLinRef>"
        dato3(17) = "<TpoDocRef>801</TpoDocRef>"
        dato3(18) = "<FolioRef>" + LeedatoOCproveedor(TIPO, numero, loc, "ordencompra") + "</FolioRef>"
        dato3(19) = "<FchRef>" + Format(LeedatoOCproveedor(TIPO, numero, loc, "fechaoc"), "yyyy-mm-dd") + "</FchRef>"
        dato3(20) = "<CodRef>3</CodRef>"
        dato3(21) = "<RazonRef>" + "CAJA:" + caja + " NI:" & numero & "</RazonRef>"
        dato3(22) = "</Referencia>"
    End If

    


Rem ZURITA ACA
If dte_tipodte = 61 Or dte_tipodte = 56 Then
dato3(23) = "<Referencia>"
dato3(24) = "<NroLinRef>1</NroLinRef>"
tiponc = LeerDatoCasoNc("tipodocumento", TIPO, numero, caja, fecha)
dato3(25) = tiponc
If dato3(25) = "FV" Then dato3(25) = "<TpoDocRef>33</TpoDocRef>"
If dato3(25) = "EX" Then dato3(25) = "<TpoDocRef>34</TpoDocRef>"
If dato3(25) = "BV" Then dato3(25) = "<TpoDocRef>38</TpoDocRef>"
If dato3(25) = "ND" Then dato3(25) = "<TpoDocRef>61</TpoDocRef>"
If dato3(25) = "NF" Or dato3(25) = "NB" Or dato3(25) = "NC" Then dato3(25) = "<TpoDocRef>61</TpoDocRef>"
folionc = LeerDatoCasoNc("numerodocumento", TIPO, numero, caja, fecha)

dato3(25) = dato3(25) & " <IndGlobal>1</IndGlobal>"

dato3(26) = "<FolioRef>" + folionc + "</FolioRef>"
dato3(27) = "<FchRef>" + Format(LeerfechaCasoNc("fechadocumento", TIPO, numero, CAJADO, fecha), "yyyy-mm-dd") + "</FchRef>"
    If dte_tipodte = 56 Then
         dato3(28) = "<CodRef>1</CodRef>"
    Else
        dato3(28) = "<CodRef>3</CodRef>"
    End If
    dato3(29) = "<RazonRef>DEVOLUCION DE MERCADERIA AUT:" & leerdatocaso("revisado", TIPO, numero, caja, fecha) & "</RazonRef>"

If dte_neto = 0 And dte_tipodte = 61 Then
    dato3(28) = "<CodRef>2</CodRef>"
    dato3(29) = "<RazonRef>NOTA CREDITO ADMINISTRATIVA</RazonRef>"
End If


dato3(30) = "</Referencia>"

End If


pasada = ""
For I = 1 To 40
If dato3(I) <> "" Then
pasada = pasada + dato3(I)
End If

Next I

detalle = detalle + pasada + "</Documento></DTE>"
detalle = cabeza + detalle
cadena = detalle

For K = 1 To Len(detalle)
If Asc(Mid(detalle, K, 1)) > 128 And Mid(detalle, K, 1) <> "Ñ" Then
cadena = Replace(cadena, Mid(detalle, K, 1), "")
End If

Next K
detalle = cadena

detalle = Replace(detalle, "¥", "N")
detalle = Replace(detalle, "Ñ", "N")
detalle = Replace(detalle, "§", " ")
detalle = Replace(detalle, "º", " ")
detalle = Replace(detalle, "°", " ")
detalle = Replace(detalle, "&", "&amp;")
detalle = Replace(detalle, "ø", " ")
detalle = Replace(detalle, ",", ".")
detalle = Replace(detalle, "*", "x")
detalle = Replace(detalle, "´", "")
detalle = Replace(detalle, "Ç", "")
detalle = Replace(detalle, "ï", "")


caf = leerrutacaf(dte_tipodte, dte_folio)

entrada = "C:\FAE\" + LOCAL_PROCESO + "\DTE\" & dte_tipodte & "_" & dte_folio & ".xml"
salida = "C:\FAE\" + LOCAL_PROCESO + "\DTE\firmado_" & dte_tipodte & "_" & dte_folio & ".xml"

xml.LoadXml detalle


Call xml.SaveXml(entrada)

Close 20
Close 30

Open entrada For Input As #20

Open entrada + "2" For Output As #30
While EOF(20) = False

Line Input #20, ss


Rem ariel  quita espacio antes del sigono pregunta en el encabezado del xml

If InStr(ss, " ?>") Then
ss = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>"
End If
Rem hasta aqui

If ss <> "" Then
Print #30, ss
End If

Wend
Close 20
Close 30




 FIRMA2 = entrada + "2" + " " + salida + " " + caf + " " + certificado
Rem FIRMA2 = "-a " + caf + " -p " + entrada + " -c " + CERTIFICADO + " -s 123 -o " + salida
Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + empresaActiva + "\programas;C:\fae\" + empresaActiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + empresaActiva + "\programas\lib\jargs.jar;C:\fae\" + empresaActiva + "\programas\lib\itext-1.3.jar;C:\fae\" + empresaActiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + empresaActiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.FirmaDTE " + FIRMA2
Rem FIRMA = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;C:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.FirmaDTE " + FIRMA2
FIRMA = "c:\fae\programas\firmadte.bat " + FIRMA2

'Shell FIRMA
Call ShellAndWait(FIRMA, vbNormalFocus)
Call Sleep(10000)
detalle = leerxml(salida)
 If detalle = "" Then Exit Sub
Call GRABADTE(dte_tipodte, dte_folio, Format(fecha, "yyyy-mm-dd"), loc, TIPO, numero, Format(fecha, "yyyy-mm-dd"), caja, detalle, dte_r_rut, dte_r_nombre, dte_total)
Call imprimelectronica(dte_tipodte, CDbl(dte_folio), Format(fecha, "yyyy-mm-dd"), dte_r_rut, numero, caja)
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
Public Sub ShellAndWait(ByVal program_name As String, _
ByVal window_style As VbAppWinStyle)
Dim process_id As Long
Dim process_handle As Long
'ariel

'ariel


    ' Start the program.
    On Error GoTo ShellError
    process_id = Shell(program_name, window_style)
    On Error GoTo 0

    ' Hide.
    'Me.Visible = False
    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    ' Reappear.
    'Me.Visible = True
    Exit Sub

ShellError:
    MsgBox "Error starting task " & _
        txtProgram.text & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub

Sub imprimelectronica(TIPO, folio, fecha, rutcliente, numerointerno, caja)
Dim entradapdf As String
Dim salidapdf As String
Dim salidapdf2 As String
Dim csql As New rdoQuery
Dim tipo2 As String
Dim linea As Double
' Function leerglosafactura(TIPO, numero, caja, fecha)


detalle2 = leerxmldte(LOCAL_PROCESO, TIPO, folio)

detalle2 = Replace(detalle2, "&amp;", "&")
detalle2 = Replace(detalle2, "#209;", "Ñ")
detalle2 = Replace(detalle2, "#243;", "ó")
detalle2 = Replace(detalle2, "ø", " ")


     
      tipo2 = TIPO
      If tipo2 = "33" Then tipo2 = "1"
      If tipo2 = "61" Then tipo2 = "61"
      
   Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select linea,glosa from "
    If caja = "98" Then
        If tipo2 = "1" Then tipo2 = "2"
        csql.sql = csql.sql & clientesistema & "conta" & codigoCONTABLE & ".facturasdepublicidad_glosa "
    End If
    If caja = "99" Then
        If tipo2 = "1" Then tipo2 = "33"
         csql.sql = csql.sql & clientesistema & "conta" & codigoCONTABLE & ".facturasvarias_glosa "
    End If
    
    csql.sql = csql.sql & "where numero='" & numerointerno & "' and tipo='" & tipo2 & "' "
    csql.sql = csql.sql & "order by linea  limit 0,10"
    csql.Execute
 
    
    If csql.RowsAffected > 0 Then
        ' agrega si hay glosa
        detalle2 = Replace(detalle2, "</DTE>", "")
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            If resultados(1) <> "" Then
            linea = linea + 1
                detalle2 = detalle2 & Chr(10) & "<AD_linea" & linea & ">" & resultados(1) & "</AD_linea" & linea & ">"
            End If
            resultados.MoveNext
        Wend
        detalle2 = detalle2 & Chr(10) & "</DTE>"
    End If
    
'   End Function



xml.LoadXml detalle2

Call xml.SaveXml("c:\FAE\" + LOCAL_PROCESO + "\DTE\" & TIPO & "-" & folio & ".xml")
entradapdf = "c:\FAE\" + LOCAL_PROCESO + "\DTE\" & TIPO & "-" & folio & ".xml"
salidapdf = "C:\FAE\" + LOCAL_PROCESO + "\PDF\" & TIPO & "-" & folio & ".pdf"
salidapdf2 = "C:\FAE\" + LOCAL_PROCESO + "\PDF\" & TIPO & "-" & folio & "cedi.pdf"

Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\dte\rodrigo;C:\DTE\RODRIGO\lib\OpenLibsDTE.jar;C:\DTE\RODRIGO\lib\jargs.jar;C:\DTE\RODRIGO\lib\itext-1.3.jar;C:\DTE\RODRIGO\lib\log4j-1.2.14.jar;C:\DTE\RODRIGO\lib\xercesImpl.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\" + empresaActiva + "\impresion\FA_estandar.properties -f c:\fae\" + empresaActiva + "\impresion\FA_estandar2.pdf -o " + salidapdf
Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\OpenLibsDTE.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\xercesImpl.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\" + empresaActiva + "\impresion\FA_estandar.properties -f c:\fae\" + empresaActiva + "\impresion\FA_estandar2.pdf -o " + salidapdf
Rem original
Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\00\programas;C:\fae\00\programas\lib\apache-mime4j-0.6.jar;C:\fae\00\programas\lib\commons-codec-1.3.jar;C:\fae\00\programas\lib\commons-httpclient-3.0.jar;C:\fae\00\programas\lib\commons-logging-1.0.4.jar;C:\fae\00\programas\lib\httpclient-4.0.jar;C:\fae\00\programas\lib\httpcore-4.0.1.jar;C:\fae\00\programas\lib\httpmime-4.0.jar;C:\fae\programas\00\lib\itext-1.3.jar;C:\fae\00\programas\lib\jargs.jar;C:\fae\00\programas\lib\jdom.jar;C:\fae\00\programas\lib\log4j-1.2.14.jar;C:\fae\00\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\00\programas\lib\OpenLibsDte.jar;C:\fae\00\programas\lib\xbean.jar;C:\fae\00\programas\lib\xercesImpl.jar;C:\fae\programas\00\lib\xfire-all-1.2.6.jar;C:\fae\00\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\%3\impresion\FA_estandar.properties -f c:\fae\%3\impresion\FA_estandar2.pdf -o " + salidapdf
Rem original
Rem esto si va


pdf = "c:\fae\programas\generapdf_fa.bat  " + entradapdf + " " + salidapdf + " " + LOCAL_PROCESO

If TIPO = "46" Then
pdf = "c:\fae\programas\generapdf_fc.bat  " + entradapdf + " " + salidapdf + " " + LOCAL_PROCESO
End If
If TIPO = "52" Then
pdf = "c:\fae\programas\generapdf_gd.bat  " + entradapdf + " " + salidapdf + " " + LOCAL_PROCESO
End If
Call ShellAndWait(pdf, vbNormalFocus)



11:  If ExisteArchivo(salidapdf) = True Then
Call grabarpdf(salidapdf, TIPO, folio, Format(fecha, "yyyy-mm-dd"), Replace(rutcliente, "-", ""), 0, LOCAL_PROCESO)
    Else
        GoTo 11
    End If

Rem cedible
cedible = True

    If cedible = True Then
        
        pdf = "c:\fae\programas\generapdf_fa_cedible.bat  " + entradapdf + " " + salidapdf2 + " " + LOCAL_PROCESO
        If TIPO = "46" Then
        pdf = "c:\fae\programas\generapdf_fc_cedible.bat  " + entradapdf + " " + salidapdf2 + " " + LOCAL_PROCESO
        End If
        If TIPO = "52" Then
        pdf = "c:\fae\programas\generapdf_gd_cedible.bat  " + entradapdf + " " + salidapdf2 + " " + LOCAL_PROCESO
        End If


Rem Shell pdf
        Call ShellAndWait(pdf, vbNormalFocus)

       Rem  Call Sleep(2000)
10:          If ExisteArchivo(salidapdf2) = True Then
'                If imprIMETIPO <> "CAJA98" Then
'                    Call PrintFile(salidapdf2)
'                End If
                Call grabarpdf(salidapdf2, TIPO, folio, Format(fecha, "yyyy-mm-dd"), Replace(rutcliente, "-", ""), 1, LOCAL_PROCESO)
            Else
                GoTo 10
            End If
    End If

Call modificaimpresa(TIPO, folio)


End Sub


Public Sub grabarpdf(Ruta, TIPO, folio, fecha, rutcliente, cedible, loc)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim pdfpath, pdfpath1 As String
    Dim pdffile As ADODB.Stream
    pdfpath = Ruta
'    On Error GoTo no:
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
   cn.CursorLocation = adUseClient
    cn.Open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & clientesistema & "ventas" & ";PWD=" & password & "; UID=" & usuario & ";OPTION=3"
    rs.Open " select * from " & clientesistema & "fae" & loc & ".sv_dtepdf_" & loc & " limit 0,1", cn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    Set pdffile = New ADODB.Stream
            pdffile.Type = adTypeBinary
            pdffile.Open
            pdffile.LoadFromFile pdfpath
            rs!TIPO = TIPO
            rs!numero = folio
            rs!rut = rutcliente
            rs!fecha = Format(fecha, "yyyy-mm-dd")
            rs!cedible = cedible
            rs.Fields("pdf") = pdffile.Read
            pdffile.Close
            Set pdffile = Nothing
            rs.Update
            Set rs = Nothing
'no:
End Sub
Public Function Cargarpdf(TIPO, numero, fecha, rutcliente) As String
Dim Tamaño As Double
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim mstream As ADODB.Stream
Dim pdfpath, pdfpath1 As String
Dim pdffile As ADODB.Stream

Dim ImgTemporal As String
ImgTemporal = "C:\tmp_pdf.pdf"
If ExisteArchivo(ImgTemporal) = True Then Kill ImgTemporal

Set cn = New ADODB.Connection
cn.Open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & clientesistema & "ventas" & ";PWD=" & password & "; UID=" & usuario & ";OPTION=3"
cn.CursorLocation = adUseClient
 

Set rs = New ADODB.Recordset
'Rs.Open " select * from pdf where pdfid='" & txtid.text & "' and pdfname='" & txtname.text & "'", cn, adOpenKeyset, adLockOptimistic
rs.Open "Select * from " & clientesistema & "fae" + LOCAL_PROCESO + ".sv_dtepdf_" + LOCAL_PROCESO & " where tipo='" & TIPO & "' and numero='" & numero & "' and rut = '" & rutcliente & "' and  fecha ='" & Format(fecha, "yyyy-mm-dd") & "' limit 0,1 ", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
Set pdffile = New ADODB.Stream
pdffile.Type = adTypeBinary
pdffile.Open
If IsNull(rs.Fields("pdf")) = False Then
pdffile.Write rs.Fields("pdf").Value
'Dim pdfnme As String
'pdfnme = txtid.text & txtname.text
'pdffile.SaveToFile "" & App.Path & "\reports\" & pdfnme & ".pdf", adSaveCreateOverWrite
pdffile.SaveToFile ImgTemporal, adSaveCreateOverWrite
pdffile.SaveToFile ImgTemporal, adSaveCreateOverWrite
pdffile.Close
Set pdffile = Nothing
ShellExecute tmplistado7.hwnd, "print", ImgTemporal, vbNullString, App.Path, 0

'Shell "C:\Archivos de programa\Adobe\Reader 10.0\Reader\AcroRd32.exe " & ImgTemporal
'MsgBox "pdf file downloaded"
Else
MsgBox "NO SE HA ENCONTRADO EL ARCHIVO", vbCritical, "ATENCION"
rs.Close
Set rs = Nothing
End If
End If

End Function


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

GoTo 10:
End If
If leerxml = "" Then End


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


Public Sub GRABADTE(TIPO, numero, fecha, LOCALDO, tipodo, numerodo, FECHADO, CAJADO, ARCHIVO, rut, nombre, monto)
    Dim CAMPOS(20, 3) As Variant
    
    Dim op As Integer
    CAMPOS(0, 0) = "tipo"
    CAMPOS(1, 0) = "numero"
    CAMPOS(2, 0) = "fecha"
    CAMPOS(3, 0) = "localdocumento"
    CAMPOS(4, 0) = "tipodocumento"
    CAMPOS(5, 0) = "numerodocumento"
    CAMPOS(6, 0) = "fechadocumento"
    CAMPOS(7, 0) = "cajadocumento"
    CAMPOS(8, 0) = "xml"
    CAMPOS(9, 0) = "rut"
    CAMPOS(10, 0) = "nombre"
    CAMPOS(11, 0) = "monto"
    CAMPOS(12, 0) = "administracion"
    'agregado para que no lo envie
     CAMPOS(13, 0) = ""
'    CAMPOS(13, 0) = "fechaenviosii"
'    CAMPOS(14, 0) = "track"
'    CAMPOS(15, 0) = "respuestasii"
'    CAMPOS(16, 0) = "status"
'    CAMPOS(17, 0) = "enviada"
'    CAMPOS(18, 0) = "aceptada"
'    CAMPOS(19, 0) = ""
    
'    'agregado para que no envie
    
    
    
    
    CAMPOS(0, 1) = TIPO
    CAMPOS(1, 1) = numero
    CAMPOS(2, 1) = fecha
    CAMPOS(3, 1) = LOCALDO
    CAMPOS(4, 1) = tipodo
    CAMPOS(5, 1) = numerodo
    CAMPOS(6, 1) = FECHADO
    CAMPOS(7, 1) = CAJADO
    CAMPOS(8, 1) = ARCHIVO
    CAMPOS(9, 1) = rut
    CAMPOS(10, 1) = nombre
    CAMPOS(11, 1) = monto
    CAMPOS(12, 1) = "AD"
    'AGREGADO PARA QUE NO ENVIE
'    CAMPOS(13, 1) = "2014-12-30"
'    CAMPOS(14, 1) = "9999999999"
'    CAMPOS(15, 1) = "PRUEBA"
'    CAMPOS(16, 1) = "1"
'    CAMPOS(17, 1) = "1"
'    CAMPOS(18, 1) = "1"
'    'AGREGADO PARA QUE NO ENVIE
    
    
    
    
    
    CAMPOS(0, 2) = clientesistema & "fae" + LOCALDO + ".sv_dte" + LOCALDO ' & "_prueba"
    condicion = ""
    op = 2
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    Call sincronizadatos(generacadena(sqlventas.response, op, condicion), ventas)
    
    Call actualizafoliosii(tipodo, numerodo, CAJADO, FECHADO, numero)
    If CAJADO = "98" Then
        Call modificafacturadepublicidad(TIPO, FECHADO, numerodo, CAJADO, "00", numero, rut)
    End If
    If CAJADO = "99" Then
        Call modificafacturaotros(TIPO, FECHADO, numerodo, CAJADO, "00", numero, rut)
    End If
    
End Sub

Public Function timbrafactura(EMISOR, envia, RECEPTOR, fecha, RESOLU, numero, TIPO, INICIO, final, sucursal) As String
Dim dato(100) As String
Dim comi As String
Dim I As Integer
empresa
Call empresadte(codigoCONTABLE)
rutreceptor = "60803000-K"
Call clientedte(rutreceptor, sucursal)

comi = Chr(34)
If TIPO = "FV" Then TIPO = "33"
If TIPO = "NF" Then TIPO = "61"
If TIPO = "NB" Then TIPO = "61"
If TIPO = "ND" Then TIPO = "56"

dato(1) = "<EnvioDTE xmlns=" + comi + "http://www.sii.cl/SiiDte" + comi + " xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " version=" + comi + "1.0" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiDte EnvioDTE_v10.xsd" + comi + ">"
dato(2) = "<SetDTE ID=" + comi + "EnvDte-" & INICIO & "-" & final & comi + "> "
dato(3) = "<Caratula version =" + comi + "1.0" + comi + "> "
dato(4) = "<RutEmisor>" + EMISOR + "</RutEmisor>"
dato(5) = "<RutEnvia>" + envia + "</RutEnvia>"
dato(6) = "<RutReceptor>" + rutreceptor + "</RutReceptor>"
dato(7) = "<FchResol>" + "2010-12-29" + "</FchResol>"
dato(8) = "<NroResol>" + "198" + "</NroResol>  "
dato(9) = "<TmstFirmaEnv>" & Format(Date, "yyyy-mm-dd") & "T" & Time & "</TmstFirmaEnv>"
If tipodoc(1) <> 0 Then
dato(10) = "<SubTotDTE>"
dato(11) = "<TpoDTE>" + "33" + "</TpoDTE>"
dato(12) = "<NroDTE>" & tipodoc(1) & "</NroDTE>"
dato(13) = "</SubTotDTE>"

End If
If tipodoc(2) <> 0 Then
dato(14) = "<SubTotDTE>"
dato(15) = "<TpoDTE>" + "56" + "</TpoDTE>"
dato(16) = "<NroDTE>" & tipodoc(2) & "</NroDTE>"
dato(17) = "</SubTotDTE>"
End If
If tipodoc(3) <> 0 Then
dato(18) = "<SubTotDTE>"
dato(19) = "<TpoDTE>" + "61" + "</TpoDTE>"
dato(20) = "<NroDTE>" & tipodoc(3) & "</NroDTE>"
dato(21) = "</SubTotDTE>"
End If

dato(22) = "</Caratula>"
For I = 1 To 22
If dato(I) <> "" Then
timbrafactura = timbrafactura + Chr(13) + dato(I)
End If
Next I

End Function

Public Function leerfoliodte(loc, TIPO) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select numero from " + clientesistema + "fae" + loc + ".sv_dte" + loc '& "_prueba"
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero<9999999999 and administracion='AD' order by numero desc limit 0,1"
csql.Execute
leerfoliodte = 1
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerfoliodte = resultados(0) + 1
End If
csql.Close
Set csql = Nothing
End Function
Public Function leerfolioautorizado(loc, TIPO, ByRef numero As String) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim cSql2 As New rdoQuery
Dim resultados2 As rdoResultset


Set csql.ActiveConnection = ventasRubro
csql.sql = "select hasta from " + clientesistema + "fae" + loc + ".sv_caf_administrador" + loc
csql.sql = csql.sql & " where tipo='" + TIPO + "' and desde<='" & numero & "' and hasta >= '" & numero & " ' and tipocaf='A' "
csql.Execute
leerfolioautorizado = False

If csql.RowsAffected > 0 Then
    leerfolioautorizado = True
    
Else
        Set cSql2.ActiveConnection = ventasRubro
    cSql2.sql = "select desde from " + clientesistema + "fae" + loc + ".sv_caf_administrador" + loc
    cSql2.sql = cSql2.sql & " where tipo='" + TIPO + "' and desde>='" & numero & "' and tipocaf='A' order by desde limit 0,1 "
    cSql2.Execute
    If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        leerfolioautorizado = True
        numero = resultados2(0)
    Else
      If electro04.sihayerror.Value = 1 Then MsgBox "FOLIO DE DOCUMENTO TIPO " & TIPO & " NO AUTORIZADO POR SII EN LOCAL " & loc
    End If
    cSql2.Close
    Set cSql2 = Nothing
    
End If
csql.Close
Set csql = Nothing
End Function

Public Function existedte(loc, TIPO, numero, fecha, caja, impre) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select numero from " + clientesistema + "fae" + loc + ".sv_dte" + loc '& "_prueba"
csql.sql = csql.sql & " where tipodocumento='" + TIPO + "' and numerodocumento='" + numero + "' and fechadocumento='" + Format(fecha, "yyyy-mm-dd") + "' and cajadocumento='" + caja + "' "
csql.Execute
existedte = False
If csql.RowsAffected > 0 Then
    existedte = True
End If
csql.Close
Set csql = Nothing
End Function

Public Function foliodte(loc, TIPO, numero, fecha, caja) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select numero from " + clientesistema + "fae" + loc + ".sv_dte" + loc '& "_prueba"
csql.sql = csql.sql & " where tipodocumento='" + TIPO + "' and numerodocumento='" + numero + "' and fechadocumento='" + Format(fecha, "yyyy-mm-dd") + "' and cajadocumento='" + caja + "' "
csql.Execute
foliodte = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    foliodte = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function
Public Function leerxmldte(loc, TIPO, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
If TIPO = "FV" Then TIPO = "33"
If TIPO = "NF" Then TIPO = "61"
If TIPO = "NB" Then TIPO = "61"
If TIPO = "ND" Then TIPO = "56"
If TIPO = "GD" Then TIPO = "52"
If TIPO = "FC" Then TIPO = "46"

Set csql.ActiveConnection = ventasRubro
csql.sql = "select xml from " + clientesistema + "fae" + loc + ".sv_dte" + loc '& "_prueba"
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero='" & numero & "' "
csql.Execute
leerxmldte = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerxmldte = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function

Public Function leerDATODTE(loc, TIPO, numero, dato) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
If TIPO = "FV" Then TIPO = "33"
Set csql.ActiveConnection = ventasRubro
csql.sql = "select " + dato + " from " + clientesistema + "fae" + loc + ".sv_dte" + loc '& "_prueba"
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero='" + numero + "' "
csql.Execute
leerDATODTE = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    If IsNull(resultados(0)) = False Then
    leerDATODTE = resultados(0)
    End If
End If
csql.Close
Set csql = Nothing
End Function

Public Function leerdatosfactura(TIPO, numero, caja, fecha)

End Function
Public Sub empresadte(empresa)
   
   Dim op As Integer
   Dim K As Integer
   Dim CAMPOS(10, 10) As String
   
    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = "direccion"
    CAMPOS(2, 0) = "comuna"
    CAMPOS(3, 0) = "ciudad"
    CAMPOS(4, 0) = "rut"
    CAMPOS(5, 0) = "girodte"
    CAMPOS(6, 0) = "codactividadeconomica"
    CAMPOS(7, 0) = "certificado"
    CAMPOS(8, 0) = "rutenviasii"
    CAMPOS(9, 0) = ""
    CAMPOS(0, 2) = clientesistema + "conta" + ".maestroempresas"
  
    condicion = "codigoempresa = '" & empresa & "' "
    op = 5
    sqlventas.response = CAMPOS
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
    End If
End Sub
Public Sub clientedte(rut, sucursal)
   Dim op As Integer
   Dim K As Integer
   Dim CAMPOS(10, 10) As String
   Dim cuentapublicidad As String
   
    CAMPOS(0, 0) = "rut"
    CAMPOS(1, 0) = "nombre"
    CAMPOS(2, 0) = "giro"
    CAMPOS(3, 0) = "direccion"
    CAMPOS(4, 0) = "comuna"
    CAMPOS(5, 0) = "ciudad"
    CAMPOS(6, 0) = ""
   ' cuentapublicidad = leerdatos(gestion, cliente_sql & "conta.maestroempresas", "cuentapublicidad", "codigoempresa='" + codigoCONTABLE + "' ")
    
    
   ' campos(0, 2) = cliente_sql & "conta" & codigoCONTABLE & ".cuentascorrientes "
    CAMPOS(0, 2) = clientesistema + "ventas" + ".sv_maestroclientes "
    condicion = "rut = '" & rut & "' and sucursal='" + sucursal + "' "
   ' condicion = "rut LIKE '%" & RUT & "%' AND tipo='" & cuentapublicidad & "' AND año='" & Format(fechasistema, "yyyy") & "' "
    op = 5
    sqlventas.response = CAMPOS
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


Public Function leerdatos(ByRef coneccion As rdoConnection, tabla, dato, CONSULTA) As String
    Dim CAMPOS(10, 10) As String
    CAMPOS(0, 0) = dato
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = tabla
    condicion = CONSULTA
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = coneccion
    Call sqlventas.sqlventas(op, condicion)
    
    If sqlventas.Status = 0 Then
    leerdatos = sqlventas.response(0, 3)
    Else
    leerdatos = ""
    
    End If
    
End Function

Public Sub clientedtePROVEEDOR(rut, sucursal)
   Dim op As Integer
   Dim K As Integer
   Dim CAMPOS(10, 10) As String
    CAMPOS(0, 0) = "rut"
    CAMPOS(1, 0) = "nombre"
    CAMPOS(2, 0) = "giro"
    CAMPOS(3, 0) = "direccion"
    CAMPOS(4, 0) = "comuna"
    CAMPOS(5, 0) = "ciudad"
    CAMPOS(6, 0) = ""
    CAMPOS(0, 2) = clientesistema + "ventas" + ".sv_maestroclientes "
    condicion = "rut = '" & rut & "' and sucursal='" + sucursal + "' "
    op = 5
    sqlventas.response = CAMPOS
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




'Sub leerdetalledte(tipo, numero, caja, fecha, dcto)
'   Dim csql As New rdoQuery
'   Dim resultados As rdoResultset
'   Dim tabla As String
'   Dim linea As Double
'   Dim taza As Double
'   Dim Descuento As Double
'   Dim montodescuento As Double
'   Dim total As Double
'    Dim DETALLENETO As Double
'
'   taza = 1.19
'   dte_descuento = dcto
'   Set csql.ActiveConnection = ventasRubro
'   tabla = ""
'   tabla = "select codigo,descripcion,cantidad,precio,total,descuento,impuesto,porcentajeimpuesto "
'   tabla = tabla & "from sv_otros_documento_detalle_" & empresaActiva
'   tabla = tabla & " where caja='" + caja + "' and tipo='" & tipo & "' and numero='" & numero & "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' and local='" & empresaActiva & "' order by linea"
'   csql.sql = tabla
'   csql.Execute
'   linea = 0
'   If csql.RowsAffected = 0 Then autorizadte = False
'   If csql.RowsAffected > 0 Then
'   Set resultados = csql.OpenResultset
'   autorizadte = True
'
'   dte_neto = 0
'   dte_iva = 0
'   dte_total = 0
'   dte_ref = 0
'   dte_vin = 0
'   dte_lic = 0
'   dte_cer = 0
'   dte_har = 0
'   dte_car = 0
'   dte_exento = 0
'   dte_descuentoglobal = 0
'
'  While Not resultados.EOF
'    linea = linea + 1
'    taza = 1.19 + resultados("porcentajeimpuesto")
'    tazaimpuesto = resultados("porcentajeimpuesto")
'    matrixdte(linea, 1) = resultados("codigo")
'    matrixdte(linea, 2) = resultados("cantidad")
'    matrixdte(linea, 3) = resultados("descripcion")
'    Rem DETALLENETO = Round(resultados("total") / taza / resultados("cantidad"), 4)
'    DETALLENETO = Round((resultados("precio") / taza), 4)
'    matrixdte(linea, 4) = DETALLENETO
'    If dte_descuento = 0 Then
'    matrixdte(linea, 5) = resultados("descuento")
'    Else
'    Rem matrixdte(linea, 5) = dte_descuento
'    matrixdte(linea, 5) = "0"
'    End If
'
'    total = matrixdte(linea, 2) * matrixdte(linea, 4)
'    If Val(matrixdte(linea, 5)) <> 0 Then
'    montodescuento = Round((total * (matrixdte(linea, 5) / 100)))
'    Else
'    montodescuento = 0
'    End If
'    If dte_descuento <> 0 Then
'    If Val(resultados("impuesto")) = 0 And dte_descuento <> 0 Then
'    descuento1 = descuento1 + montodescuento
'    End If
'
''    If Val(resultados("impuesto")) = 8 And dte_descuento <> 0 Then
''    descuento2 = descuento2 + montodescuento
''    End If
'    End If
'    matrixdte(linea, 8) = montodescuento
'    total = total - montodescuento
'    matrixdte(linea, 6) = Round(total, 0)
'    matrixdte(linea, 7) = ""
'    If Val(resultados("impuesto")) <> 0 And Val(resultados("impuesto")) <> 8 Then
'    matrixdte(linea, 7) = leerimpuestofae(resultados("impuesto"))
'    If Val(resultados("impuesto")) = 5 Then
'    dte_car = dte_car + Round(total * tazaimpuesto)
'    End If
'    If Val(resultados("impuesto")) = 4 Then
'    dte_har = dte_har + Round(total * tazaimpuesto)
'    End If
'    If Val(resultados("impuesto")) = 3 Then
'    dte_lic = dte_lic + (total * tazaimpuesto)
'    End If
'    If Val(resultados("impuesto")) = 2 Then
'    dte_vin = dte_vin + (total * tazaimpuesto)
'    End If
'    If Val(resultados("impuesto")) = 6 Then
'    dte_cer = dte_cer + (total * tazaimpuesto)
'    End If
'    If Val(resultados("impuesto")) = 1 Then
'    dte_ref = dte_ref + (total * tazaimpuesto)
'    End If
'
'
'    Else
'    matrixdte(linea, 7) = ""
'
'
'    End If
'    matrixdte(linea, 0) = "0"
'    If Val(resultados("impuesto")) <> 8 Then
'    dte_neto = dte_neto + total
'
'    Else
'    dte_exento = dte_exento + total
'    matrixdte(linea, 0) = "1"
'
'    End If
'
'    resultados.MoveNext
'
'  Wend
'   End If
'   csql.Close
'   Set resultados = Nothing
'   Set csql = Nothing
'
'
'   If dte_descuento <> 0 Then
'
'   Descuento = Round(dte_neto * dte_descuento / 100)
'   dte_neto = dte_neto - Descuento
'   descuento1 = Descuento
'   dte_iva = Round(dte_neto * 19 / 100, 0)
'
'   Descuento = Round(dte_ref * dte_descuento / 100, 3)
'   dte_ref = dte_ref - Descuento
'
'   Descuento = Round(dte_vin * dte_descuento / 100, 3)
'   dte_vin = dte_vin - Descuento
'
'   Descuento = Round(dte_lic * dte_descuento / 100)
'   dte_lic = dte_lic - Descuento
'
'   Descuento = Round(dte_cer * dte_descuento / 100)
'   dte_cer = dte_cer - Descuento
'
'   Descuento = Round(dte_har * dte_descuento / 100)
'   dte_har = dte_har - Descuento
'
'   Descuento = Round(dte_car * dte_descuento / 100)
'   dte_car = dte_car - Descuento
'
'  Rem  Descuento = Round(dte_exento * dte_descuento / 100)
'   dte_exento = dte_exento
'
''   If dte_exento <> 0 Then
''   descuento2 = Descuento
''
''   End If
''
'
'   End If
'   dte_iva = Round(dte_neto * 19 / 100, 0)
'   dte_ref = Round(dte_ref)
'   dte_vin = Round(dte_vin)
'   dte_lic = Round(dte_lic)
'   dte_cer = Round(dte_cer)
'   dte_har = Round(dte_har)
'   dte_car = Round(dte_car)
'   dte_neto = Round(dte_neto)
'   dte_exento = Round(dte_exento)
'   dte_total = dte_neto + dte_iva + dte_ref + dte_vin + dte_cer + dte_lic + dte_har + dte_car + dte_exento
'   dte_descuento = Round(dte_descuento)
'   totallineas = linea
'
'
'
'   End Sub
   
   
   Function leerglosafactura(TIPO, numero, caja, fecha)
      Dim csql As New rdoQuery
      Dim tipo2 As String
      tipo2 = TIPO
      If tipo2 = "FV" Then tipo2 = "1"
      If tipo2 = "NF" Then tipo2 = "61"
      
   Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select linea,glosa from "
    If caja = "98" Then
        If tipo2 = "1" Then tipo2 = "2"
        csql.sql = csql.sql & clientesistema & "conta" & codigoCONTABLE & ".facturasdepublicidad_glosa "
    End If
    If caja = "99" Then
        If tipo2 = "1" Then tipo2 = "33"
         csql.sql = csql.sql & clientesistema & "conta" & codigoCONTABLE & ".facturasvarias_glosa "
    End If
    
    csql.sql = csql.sql & "where numero='" & numero & "' and tipo='" & tipo2 & "' "
    csql.sql = csql.sql & "order by linea "
    csql.Execute
    leerglosafactura = ""
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            If resultados(1) <> "" Then
                leerglosafactura = leerglosafactura & " " & resultados(1)
            End If
            resultados.MoveNext
        Wend
    End If
    
   End Function
   Sub leerdetalledte(TIPO, numero, caja, fecha, dcto, exento)
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
   tabla = tabla & "from sv_otros_documento_detalle_" & LOCAL_PROCESO
   tabla = tabla & " where caja='" + caja + "' and tipo='" & TIPO & "' and numero='" & numero & "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' and local='" & LOCAL_PROCESO & "' order by linea"
   csql.sql = tabla
   csql.Execute
   linea = 0
   If csql.RowsAffected = 0 Then autorizadte = False
   If csql.RowsAffected > 0 Then
   Set resultados = csql.OpenResultset
   autorizadte = True
   
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
   dte_ivaretenido = 0
 
  While Not resultados.EOF
    linea = linea + 1
    taza = 1.19 + resultados("porcentajeimpuesto")
    tazaimpuesto = resultados("porcentajeimpuesto")
    If TIPO = "EX" Then taza = 1
    If resultados("impuesto") = "00008" Then taza = 1
    matrixdte(linea, 1) = resultados("codigo")
    matrixdte(linea, 2) = resultados("cantidad")
    
    'agregado para validar el largo de la descripcion
    matrixdte(linea, 3) = Mid(resultados("descripcion"), 1, 40)
    If Mid(matrixdte(linea, 3), 40, 1) = " " Then
       matrixdte(linea, 3) = Mid(matrixdte(linea, 3), 1, 39) & "_"
    End If
    'agregado para validar el largo de la descripcion
    
    If linea = 1 Then
     matrixdte(linea, 9) = leerglosafactura(TIPO, numero, caja, Format(fecha, "yyyy-mm-dd"))
     End If
    Rem DETALLENETO = Round(resultados("total") / taza / resultados("cantidad"), 4)
'      If tipo = "FV" And exento > 0 Then dte_exento = exento ' agregado para el exento de factura normales
      If TIPO = "FV" And exento > 0 And linea = 1 And resultados("impuesto") <> "00008" Then ' agregado para el exento de factura normales
        DETALLENETO = Round(((resultados("precio") - exento) / taza), 4)
      Else
        DETALLENETO = Round((resultados("precio") / taza), 4)
      End If
      
    
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
    
    If Val(resultados("impuesto")) = 15 Then
    dte_ivaretenido = dte_ivaretenido + (total * 0.19)
    End If
    
    Else
    matrixdte(linea, 7) = ""
    
    
    End If
    matrixdte(linea, 0) = "0"
    If TIPO <> "EX" And resultados("impuesto") <> "00008" Then
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
'   If tipo = "FV" And exento > 0 Then dte_exento = exento ' agregado para el exento de factura normales
   
   dte_iva = Round(dte_neto * 19 / 100, 0)
   dte_ref = Round(dte_ref)
   dte_vin = Round(dte_vin)
   dte_lic = Round(dte_lic)
   dte_cer = Round(dte_cer)
   dte_har = Round(dte_har)
   dte_car = Round(dte_car)
   dte_neto = Round(dte_neto)
   dte_exento = Round(dte_exento)
   If dte_ivaretenido <> 0 Then
        dte_ivaretenido = dte_iva
   End If
   
   dte_total = dte_neto + dte_iva + dte_ref + dte_vin + dte_cer + dte_lic + dte_har + dte_car + dte_exento - dte_ivaretenido
   dte_descuento = Round(dte_descuento)
   totallineas = linea
  
   
    
    
   End Sub
Public Function leerrutacaf(TIPO, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select ubicacion,nombredelarchivo from " + clientesistema + "fae" + LOCAL_PROCESO + ".sv_caf_administrador" + LOCAL_PROCESO
csql.sql = csql.sql & " where tipo='" + TIPO + "' and desde<='" & numero & "' and hasta >= '" & numero & " ' "
csql.Execute


If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerrutacaf = resultados(0) + resultados(1)
    
End If
csql.Close
Set csql = Nothing
End Function
Public Sub modificaimpresa(TIPO, folio)
'        Dim csql As rdoQuery
'        Set csql = New rdoQuery
'        Set csql.ActiveConnection = ventas
'        csql.sql = "UPDATE " + clientesistema + "fae" + empresaActiva + ".sv_dte" + empresaActiva + "_prueba "
'        csql.sql = csql.sql & "set impresa='1' WHERE tipo='" + tipo + "' and numero='" + folio + "' "
'        csql.Execute
'        Call sincronizadatos(csql.sql, ventas)
'        csql.Close
'        Set csql = Nothing
    End Sub

'Sub imprimelectronica(tipo, folio)
'Dim entradapdf As String
'Dim salidapdf As String
'Dim salidapdf2 As String
'folio = Val(folio)
'
'detalle2 = leerxmldte(empresaActiva, tipo, folio)
'
'detalle2 = Replace(detalle2, "&amp;", "&")
'detalle2 = Replace(detalle2, "#209;", "Ñ")
'detalle2 = Replace(detalle2, "#243;", "ó")
'detalle2 = Replace(detalle2, "ø", " ")
'
'
'xml.LoadXml detalle2
'
'Call xml.SaveXml("c:\FAE\" + empresaActiva + "\DTE\" & tipo & "-" & folio & ".xml")
'entradapdf = "c:\FAE\" + empresaActiva + "\DTE\" & tipo & "-" & folio & ".xml"
'salidapdf = "C:\FAE\" + empresaActiva + "\PDF\" & tipo & "-" & folio & ".pdf"
'salidapdf2 = "C:\FAE\" + empresaActiva + "\PDF\" & tipo & "-" & folio & "cedi.pdf"
'
'Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\dte\rodrigo;C:\DTE\RODRIGO\lib\OpenLibsDTE.jar;C:\DTE\RODRIGO\lib\jargs.jar;C:\DTE\RODRIGO\lib\itext-1.3.jar;C:\DTE\RODRIGO\lib\log4j-1.2.14.jar;C:\DTE\RODRIGO\lib\xercesImpl.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\" + empresaActiva + "\impresion\FA_estandar.properties -f c:\fae\" + empresaActiva + "\impresion\FA_estandar2.pdf -o " + salidapdf
'Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\OpenLibsDTE.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\xercesImpl.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\" + empresaActiva + "\impresion\FA_estandar.properties -f c:\fae\" + empresaActiva + "\impresion\FA_estandar2.pdf -o " + salidapdf
'Rem original
'Rem pdf = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\00\programas;C:\fae\00\programas\lib\apache-mime4j-0.6.jar;C:\fae\00\programas\lib\commons-codec-1.3.jar;C:\fae\00\programas\lib\commons-httpclient-3.0.jar;C:\fae\00\programas\lib\commons-logging-1.0.4.jar;C:\fae\00\programas\lib\httpclient-4.0.jar;C:\fae\00\programas\lib\httpcore-4.0.1.jar;C:\fae\00\programas\lib\httpmime-4.0.jar;C:\fae\programas\00\lib\itext-1.3.jar;C:\fae\00\programas\lib\jargs.jar;C:\fae\00\programas\lib\jdom.jar;C:\fae\00\programas\lib\log4j-1.2.14.jar;C:\fae\00\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\00\programas\lib\OpenLibsDte.jar;C:\fae\00\programas\lib\xbean.jar;C:\fae\00\programas\lib\xercesImpl.jar;C:\fae\programas\00\lib\xfire-all-1.2.6.jar;C:\fae\00\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.GeneraPDF -d " + entradapdf + " -p c:\fae\%3\impresion\FA_estandar.properties -f c:\fae\%3\impresion\FA_estandar2.pdf -o " + salidapdf
'Rem original
'Rem esto si va
'
'
'pdf = "c:\fae\programas\generapdfgdinterna.bat  " + entradapdf + " " + salidapdf + " " + empresaActiva
'Rem Shell "c:\fae\programas\generapdfgdinterna.bat  " + entradapdf + " " + salidapdf + " " + empresaActiva
'
'Shell pdf
''Call Sleep(2000)
''
''11: If ExisteArchivo(salidapdf) = True Then
''  If imprIMETIPO <> "CAJA98" Then
''
''Call PrintFile(salidapdf)
''End If
''Else
''GoTo 11
''End If
''
''Rem cedible
''cedible = True
''
''    If cedible = True Then
''
''        pdf = "c:\fae\programas\generapdf2.bat  " + entradapdf + " " + salidapdf2 + " " + empresaActiva
''        Shell pdf
''
''        Call Sleep(2000)
''10:         If ExisteArchivo(salidapdf2) = True Then
''            If imprIMETIPO <> "CAJA98" Then
''            Call PrintFile(salidapdf2)
''            End If
''            Else
''            GoTo 10
''            End If
''    End If
''
'Call modificaimpresa(tipo, folio)
'
'
'End Sub
Sub enviasii(TIPO, folio, rutempresa, rutenvia, rutreceptor, fecha, entradafirma, sucursal)
empresa
Call empresadte(codigoCONTABLE)

rutreceptor = "60803000-K"
Rem Call clientedte(rutreceptor, sucursal)


firmaenvio = "c:\fae\programas\firmaenvio.bat " + entradafirma + " " + certificado + " " + salidafirma
Rem firmaenvio = "c:\fae\" + empresaActiva + "\programas\firmaenvio.bat " + entradafirma + " " + CERTIFICADO + " " + salidafirma

Rem firmaenvio = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;C:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar cl.eltit.dte.FirmaEnvio -p %1 -c %2 -s 123 -o %3"


Shell firmaenvio
Sleep (4000)
'
'
respuestasii = "c:\fae\" + LOCAL_PROCESO + "\respuesta_sii\res" + folio + ".xml"
enviarsii = "c:\fae\programas\enviasii.bat " + salidafirma + " " + certificado + " c:\fae\" + LOCAL_PROCESO + "\respuesta_sii\res" + folio + ".xml"
Rem enviarsii = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\" + empresaActiva + "\programas;C:\fae\" + empresaActiva + "\programas\lib\OpenLibsDTE.jar;C:\fae\" + empresaActiva + "\programas\lib\jargs.jar;C:\fae\" + empresaActiva + "\programas\lib\itext-1.3.jar;C:\fae\" + empresaActiva + "\programas\lib\log4j-1.2.14.jar;C:\fae\" + empresaActiva + "\programas\lib\xercesImpl.jar cl.eltit.dte.EnviarSII -p " + salidafirma + " -c " + CERTIFICADO + " -s 123 -o " + respuestasii
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

Public Function leerdatocaso(dato, TIPO, numero, caja, fecha) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select " + dato + " from sv_documento_cabeza_" + LOCAL_PROCESO + " "
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero='" + numero + "' and caja='" + caja + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' "
csql.Execute

leerdatocaso = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerdatocaso = resultados(0)
    
End If
csql.Close
Set csql = Nothing
End Function

Public Function LeerDatoCasoNc(dato, TIPO, numero, caja, fecha) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
If caja = "99" And (TIPO = "NF" Or TIPO = "ND") Then
csql.sql = "select " + dato + " from sv_otros_documento_detalle_" + LOCAL_PROCESO + " "
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero='" + numero + "'and fecha='" + Format(fecha, "yyyy-mm-dd") + "' "
csql.sql = csql.sql & " group by tipo,fecha,numero "

Else
csql.sql = "select " + dato + " from sv_documento_detalle_" + LOCAL_PROCESO + " "
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero='" + numero + "'and fecha='" + Format(fecha, "yyyy-mm-dd") + "' "
csql.sql = csql.sql & " group by tipo,fecha,numero "

End If

csql.Execute

LeerDatoCasoNc = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    LeerDatoCasoNc = resultados(0)
    
End If
csql.Close
Set csql = Nothing
End Function
Public Function LeerfechaCasoNc(dato, TIPO, numero, caja, fecha) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select " + dato + " from sv_otros_documento_detalle_" + LOCAL_PROCESO + " "
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero='" + numero + "' "
csql.sql = csql.sql & " group by tipo,fecha,numero "
csql.Execute

LeerfechaCasoNc = fechasistema
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    LeerfechaCasoNc = resultados(0)
    
End If
csql.Close
Set csql = Nothing
End Function



Public Sub grabar_recepcion(correo, ARCHIVO, fecha_recepcion, archivo_recepcion, archivo_respuesta, fecha_respuesta)
Dim recepciona As String


Dim CAMPOS(10, 10) As String

        Dim op As Integer
        
        Dim sss As Double
        
        CAMPOS(0, 0) = "correo"
        CAMPOS(1, 0) = "archivo"
        CAMPOS(2, 0) = "fecha_recepcion"
        CAMPOS(3, 0) = "archivo_recepcion"
        CAMPOS(4, 0) = "archivo_respuesta"
        CAMPOS(5, 0) = "rut"
        CAMPOS(6, 0) = ""
        CAMPOS(0, 1) = correo
        archivo_recepcion = Replace(archivo_recepcion, "'", "")
        CAMPOS(1, 1) = ARCHIVO
        CAMPOS(2, 1) = Format(fecha_recepcion, "yyyy-mm-dd")
        CAMPOS(3, 1) = archivo_recepcion
        CAMPOS(4, 1) = archivo_respuesta
        CAMPOS(5, 1) = leer_rutcorreo(correo)
        
        
        
        CAMPOS(0, 2) = clientesistema + "fae" + LOCAL_PROCESO + ".sv_recepcion_dte00"
    
        condicion = ""

        op = 2
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        Call sincronizadatos(generacadena(sqlventas.response, op, condicion), ventas)
        
       If InStr(1, correo, "sii") > 0 Then
       For sss = 1 To Len(archivo_recepcion)
        If Mid(archivo_recepcion, sss, 7) = "TRACKID" Then
    track = "0" + Mid(archivo_recepcion, sss + 8, 9)
    Exit For
    End If
Next sss
       Call modificarespuestaenvio(archivo_recepcion, track)
       
       End If
       
       
End Sub
Public Sub recepcionar(ARCHIVO, rut)
Dim dato(10) As String
Dim respuesta As String

        Call empresadte(codigoCONTABLE)

        dato(0) = "c:\DTE_RECIBIDOS\" + ARCHIVO
        dato(1) = "c:\FAE\" + empresaActiva + "\RECEPCIONAR\"
        dato(2) = "c:\fae\" + empresaActiva + "\ACUSE\" + "ACUSE_" + ARCHIVO
        dato(3) = certificado
        dato(4) = "123"
        dato(5) = dte_rutenvia
        dato(6) = dte_e_rut
        dato(7) = "1"
        dato(8) = "1"
        dato(9) = rut
        recepciona = "java -Dfile.encoding=ISO-8859-1 -classpath c:\fae\programas;C:\fae\programas\lib\apache-mime4j-0.6.jar;C:\fae\programas\lib\commons-codec-1.3.jar;C:\fae\programas\lib\commons-httpclient-3.0.jar;C:\fae\programas\lib\commons-logging-1.0.4.jar;C:\fae\programas\lib\httpclient-4.0.jar;:\fae\programas\lib\httpcore-4.0.1.jar;C:\fae\programas\lib\httpmime-4.0.jar;C:\fae\programas\lib\itext-1.3.jar;C:\fae\programas\lib\jargs.jar;C:\fae\programas\lib\jdom.jar;C:\fae\programas\lib\log4j-1.2.14.jar;C:\fae\programas\lib\not-yet-commons-ssl-0.3.11.jar;C:\fae\programas\lib\OpenLibsDte.jar;C:\fae\programas\lib\xbean.jar;C:\fae\programas\lib\xercesImpl.jar;C:\fae\programas\lib\xfire-all-1.2.6.jar;C:\fae\programas\lib\xmlsec-1.4.0.jar;C:\fae\programas\lib\XmlSchema-1.1.jar;C:\fae\programas\lib\wsdl4j.jar cl.eltit.dte.Recepcionar "
        recepciona = recepciona + dato(0) + " " + dato(1) + " " + dato(2) + " " + dato(3) + " " + dato(4) + " " + dato(5) + " " + dato(6) + " " + dato(7) + " " + dato(8) + " " + dato(9) + " validarsii=S "
        Shell recepciona
        Sleep (3000)
        respuesta = leerxmlrecibido(dato(2))
        
        Call modificarespuesta(ARCHIVO, respuesta)
        
        

End Sub

Public Function leer_rutcorreo(correo) As String
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
    
 Set csql.ActiveConnection = ventas
 csql.sql = "select rut from " & clientesistema + "fae" + empresaActiva + ".sv_fae_proveedores "
 csql.sql = csql.sql & "where mailintercambio='" & correo & "' "
 csql.Execute
 
 If csql.RowsAffected > 0 Then
             Set resultados = csql.OpenResultset
    
    leer_rutcorreo = resultados(0)
 End If
 csql.Close
 Set csql = Nothing
 
End Function


Public Sub modificarespuesta(ARCHIVO, respuesta)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "UPDATE " + clientesistema + "fae" + empresaActiva + ".sv_recepcion_dte" + empresaActiva + " "
        csql.sql = csql.sql & "set archivo_respuesta='" + respuesta + "',fecha_respuesta='" + Format(fechasistema, "yyyy-mm-dd") + "'  WHERE archivo='" + ARCHIVO + "' "
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub
Public Function existecorreo(ARCHIVO) As Boolean

        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "select correo from " + clientesistema + "fae" + empresaActiva + ".sv_recepcion_dte" + empresaActiva + " "
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


Public Sub modificaenviocorreo(ARCHIVO)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "UPDATE " + clientesistema + "fae" + empresaActiva + ".sv_recepcion_dte" + empresaActiva + " "
        csql.sql = csql.sql & "set fecha_envio='" + Format(fechasistema, "yyyy-mm-dd") + "'  WHERE archivo='" + ARCHIVO + "' "
        csql.Execute
        Rem Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub


Public Function ultimo_envio() As String
Dim resultados As rdoResultset
        
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "select max(folio) from " + clientesistema + "fae" + empresaActiva + ".sv_envios_dte" + empresaActiva + " "
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

Public Sub modificaenvio(TIPO, folio, envio, track)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "UPDATE " + clientesistema + "fae" + empresaActiva + ".sv_dte" + empresaActiva '+ "_prueba"
        csql.sql = csql.sql & " set track='" + track + "', nombreenvio='" + envio + "',fechaenviosii='" + Format(fechasistema, "yyyy-mm-dd") + "' WHERE tipo='" + TIPO + "' and numero='" + folio + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub

Public Sub grabar_envio(folio, fechaenvio, respuestaenvio, track, track_respuesta)
Dim recepciona As String


Dim CAMPOS(10, 10) As String

        Dim op As Integer
        
        
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "fechaenvio"
        CAMPOS(2, 0) = "respuestaenvio"
        CAMPOS(3, 0) = "track"
        CAMPOS(4, 0) = "track_respuesta"
        CAMPOS(5, 0) = ""
        CAMPOS(0, 1) = folio
        CAMPOS(1, 1) = Format(fechaenvio, "yyyy-mm-dd")
        CAMPOS(2, 1) = respuestaenvio
        CAMPOS(3, 1) = track
        CAMPOS(4, 1) = track_respuesta
        
        
        
        CAMPOS(0, 2) = clientesistema + "fae" + empresaActiva + ".sv_envios_dte00"
    
        condicion = ""

        op = 2
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        Call sincronizadatos(generacadena(sqlventas.response, op, condicion), ventas)
        
       
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
Public Sub modificarespuestaenvio(envio, track)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "UPDATE " + clientesistema + "fae" + empresaActiva + ".sv_dte" + empresaActiva '+ "_prueba"
        csql.sql = csql.sql & " set respuestasii='" + envio + "' WHERE track = '" + track + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub
Public Sub modificafacturaotros(TIPO, fecha, numero, caja, empresa, folio, rutcliente)
        Dim csql As rdoQuery
        Dim tipodoc As String
        Dim tipoconta As String
        
         If TIPO = 33 Then ' factura
            tipoconta = "FA"
            tipodoc = 6
        End If
    
        If TIPO = 34 Then ' factura extenta
            tipoconta = "EX"
            tipodoc = 0
        End If
    
        If TIPO = 56 Then ' nota debito
            tipoconta = "ND"
            tipodoc = 7
        End If
        
        If TIPO = 61 Then ' nota credito
            tipoconta = "NF"
            tipodoc = 8
        End If
        If TIPO = 46 Then ' nota credito
            tipoconta = "FC"
            tipodoc = 10
        End If
        
        
        
'        Set csql = New rdoQuery
'        Set csql.ActiveConnection = ventas2
'        csql.sql = "UPDATE " + clientesistema + "conta" + codigoCONTABLE + ".movimientoscontables "
'        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='" & tipoconta & "' and numero='" + numero + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "'"
'        csql.Execute
'        csql.Close
        
       
        
    
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = ventas2
    
        csql.sql = "UPDATE " + clientesistema + "conta" + codigoCONTABLE + ".facturasvarias "
        csql.sql = csql.sql & "set foliosii='" + folio + "' WHERE tipo='" & TIPO & "' and numero='" + numero + "' and rut='" & Format(Replace(rutcliente, "-", ""), "0000000000") & "'"
        csql.Execute
        csql.Close
    
'        Set csql = New rdoQuery
'
'        Set csql.ActiveConnection = ventas2
'
'        csql.sql = "UPDATE " + clientesistema + "conta" + codigoCONTABLE + ".facturasvarias_glosa "
'        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='" & tipo & "' and numero='" + numero + "' and rut='" & rutcliente & "'"
'        csql.Execute
'        csql.Close
'
'        Set csql = New rdoQuery
'
'        Set csql.ActiveConnection = ventas2
'
'        csql.sql = "UPDATE " + clientesistema + "conta" + codigoCONTABLE + ".facturasdeventas "
'        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='" & tipodoc & "' and numero='" + numero + "'  and rut='" & rutcliente & "'"
'        csql.Execute
'        csql.Close
'        Set csql = New rdoQuery
'
'        Set csql.ActiveConnection = ventas2
'
'        csql.sql = "UPDATE " + clientesistema + "conta" + codigoCONTABLE + ".facturasdeventas_detalle "
'        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='" & tipodoc & "' and numero='" + numero + "'  and rut='" & rutcliente & "'"
'        csql.Execute
'        csql.Close
    
    End Sub
    
    Public Sub modificafacturadepublicidad(TIPO, fecha, numero, caja, empresa, folio, rutcliente)
        Dim csql As rdoQuery
         Dim tipodoc As String
        Dim tipoconta As String
        
         If TIPO = 33 Then ' factura
            tipoconta = "FA"
            tipodoc = 2
        End If
    
        If TIPO = 34 Then ' factura extenta
            tipoconta = "EX"
            tipodoc = 0
        End If
    
        If TIPO = 56 Then ' nota debito
            tipoconta = "ND"
            tipodoc = 7
        End If
        
        If TIPO = 61 Then ' nota credito
            tipoconta = "NF"
            tipodoc = 8
        End If
        
        
        
'        Set csql = New rdoQuery
'        Set csql.ActiveConnection = ventas2
'        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".movimientoscontables "
'        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='" & tipoconta & "' and numero='" + numero + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "'"
'        csql.Execute
'        csql.Close
        
       
        
    
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = ventas2
    
        csql.sql = "UPDATE " + clientesistema + "conta" + codigoCONTABLE + ".facturasdepublicidad "
        csql.sql = csql.sql & "set foliosii='" + folio + "' WHERE tipo='" & tipodoc & "' and numero='" + numero + "'  and rut='" & Format(Replace(rutcliente, "-", ""), "0000000000") & "'"
        csql.Execute
        csql.Close
    
        
        ' esto ya no va, por que se crea cuando se imprimime en contabilidad con el folio fiscal
        
'        Set csql = New rdoQuery
'
'        Set csql.ActiveConnection = ventas2
'
'        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".facturasdepublicidad_glosa "
'        csql.sql = csql.sql & "set foliosii='" + folio + "' WHERE tipo='" & tipodoc & "' and numero='" + numero + "'  and rut='" & rutcliente & "'"
'        csql.Execute
'        csql.Close
    
        
    
'        Set csql = New rdoQuery
'
'        Set csql.ActiveConnection = ventas2
'
'        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".facturasdeventas "
'        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='" & tipodoc & "' and numero='" + numero + "' and rut='" & rutcliente & "'"
'        csql.Execute
'        csql.Close
'        Set csql = New rdoQuery
'
'        Set csql.ActiveConnection = ventas2
'
'        csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".facturasdeventas_detalle "
'        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='" & tipodoc & "' and numero='" + numero + "' and rut='" & rutcliente & "' "
'        csql.Execute
'        csql.Close
    
    End Sub
Public Sub modificamercaderiaentrelocales(TIPO, fecha, numero, caja, empresa, folio)
        Dim csql As rdoQuery
    
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = ventas2
        
        csql.sql = "UPDATE " + clientesistema + "gestion" + empresa + ".l_movimientos_cabeza_" + empresa + " "
        If TIPO = "GD" Then
        csql.sql = csql.sql & "set numero='" + folio + "',foliosii='" + folio + "' WHERE tipo='EG' and numero='" + numero + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "'  "
        End If
        
        csql.Execute
        csql.Close
        
    
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = ventas2
    
        
        csql.sql = "UPDATE " + clientesistema + "gestion" + empresa + ".l_movimientos_detalle_" + empresa + " "
        If TIPO = "GD" Then
        
        csql.sql = csql.sql & "set numero='" + folio + "' WHERE tipo='EL' and numero='" + numero + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "'  "
        End If
        
        csql.Execute
        csql.Close
        
    
    End Sub

Function LeedatoOCproveedor(TIPO, numero, loc, campo) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventas
    csql.sql = "select " & campo & " from " & clientesistema & "ventas" & loc & ".sv_otros_documento_cabeza_" & loc
    csql.sql = csql.sql & " where tipo='" & TIPO & "' and numero='" & numero & "' "
    csql.Execute
    LeedatoOCproveedor = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LeedatoOCproveedor = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    If campo = "fechaoc" And LeedatoOCproveedor = "" Then LeedatoOCproveedor = Now()
End Function

Function esempresarelacionada(rutcli) As Boolean
     Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim rutconsulta As String
    rutconsulta = Mid(rutcli, 1, 9)
    rutconsulta = Val(rutconsulta)
    
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select rut from " & clientesistema & "conta.maestroempresas "
    csql.sql = csql.sql & " where rut like '" & rutconsulta & "%' "
    csql.Execute
    
    esempresarelacionada = False
    If csql.RowsAffected > 0 Then
         esempresarelacionada = True
         
        
    End If
    csql.Close
    Set csql = Nothing
End Function
Function leerformadepago(loc, TIPO, numero, fecha, caja) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select tipopago from " & clientesistema & "ventas" & loc & ".sv_otros_documento_pagos_" & loc & " "
    csql.sql = csql.sql & " where  local='" & loc & "' and tipo='" & TIPO & "' and numero='" & numero & "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' "
    csql.sql = csql.sql & " and caja='" & caja & "'"
    csql.Execute
    
    leerformadepago = "2"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        If resultados(0) = "1" Or resultados(0) = "4" Or resultados(0) = "17" Then
            leerformadepago = "1"
        Else
            leerformadepago = "2"
        End If
        
    End If
    csql.Close
    Set csql = Nothing
End Function

Public Function leerxmlcaf(ARCHIVO) As String

Dim ss As String
Dim AA As Double
Dim nombrearchivo As String
Dim contador As Double

nombrearchivo = ARCHIVO
Close 20
leerxmlcaf = ""


 If ExisteArchivo(nombrearchivo) = True Then

Open ARCHIVO For Input As #20
AA = 0
20 If EOF(20) = False Then
Line Input #20, ss
 
leerxmlcaf = leerxmlcaf + ss
 
AA = AA + 1
GoTo 20:

End If
Close 20

Else
contador = contador + 1

End If


End Function
Sub BORRARFOLIOSVENCIDOS(loc)
'    Dim csql As New rdoQuery
'    Dim resultados As rdoResultset
'    Dim dife As Double
'
'    Set csql.ActiveConnection = ventasRubro
'    csql.sql = "select ubicacion,nombredelarchivo,tipo,desde,hasta  from " + clientesistema + "fae" + loc + ".sv_caf_administrador" + loc
'    csql.Execute
'
'    If csql.RowsAffected > 0 Then
'        Set resultados = csql.OpenResultset
'        While Not resultados.EOF
'            ARCHIVO = leerxmlcaf(resultados(0) & resultados(1))
'            fechacaf = leerdatoxml(ARCHIVO, "FA>", 1)
'            dife = DateDiff("m", fechacaf, fechasistema)
'            If buscarultimofae(resultados("tipo"), resultados("hasta"), loc) = True Then
'                dife = 19
'            End If
'
'
'
'                If dife > 18 Then
'                   Call BORRARcaf(resultados(2), resultados(3), resultados(4), resultados(1), resultados(0), loc)
'                End If
'
'            resultados.MoveNext
'        Wend
'    End If

               
End Sub
Function buscarultimofae(TIPO, hasta, loc) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select numero  from " + clientesistema + "fae" + loc + ".sv_dte" + loc
    csql.sql = csql.sql & " where tipo='" & TIPO & "' and numero='" & hasta & "' and administracion='AD' "
    csql.Execute
        buscarultimofae = False
    If csql.RowsAffected > 0 Then
     buscarultimofae = True
    End If
  
    
               
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
Sub BORRARcaf(TIPO, desde, hasta, nombrearchivo, ubicacion, loc)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "delete  from " + clientesistema + "fae" + loc + ".sv_caf_administrador" + loc
    csql.sql = csql.sql & " where tipo='" & TIPO & "' and desde='" & desde & "' and hasta='" & hasta & "' and nombredelarchivo='" & nombrearchivo & "'  "
    csql.Execute
    csql.Close
    
    Set csql = Nothing
    
    If ExisteArchivo(ubicacion & nombrearchivo) = True Then
        Kill (ubicacion & nombrearchivo)
    End If
    
               
End Sub
