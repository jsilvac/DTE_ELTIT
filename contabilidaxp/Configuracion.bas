Attribute VB_Name = "Configuracion"
Option Explicit
Public NOESTAENELSII As Boolean
Public lc_exento As Double

Public mes_lc As String
Public año_lc As String

Public PASA_TIPO As String
Public PASA_RUT As String
Public PASA_NUMERO As String

Public apagar As Boolean
Public RutaArchivos As String
Public numerodte As Double
Public NUMERODOCUMENTO_DTE As Double
Public documento_dte_impreso As Boolean
Public dte_tipodte As String
Public ipusada As String
Public FECHALC As String
Public FECHALV As String
Public f3327 As Boolean
Public f3328 As Boolean
Public destino As String
Public ctadebe2(12) As Double
Public ctahaber2(12) As Double
Public ctadebemes As Double
Public ctahabermes As Double
Public glosaflujo As String
Public empresaflujo As String
Public localorden As String
Public numerodeorden As String
Public rubro As String
Public CONFI_EMPRESAFAE As String
Public nogenerar As Boolean

Public D35_cantidad As Double
Public D35_neto As Double
Public D35_iva As Double
Public D35_total As Double

Public D38_cantidad As Double
Public D38_neto As Double
Public D38_iva As Double
Public D38_total As Double

Public D48_cantidad As Double
Public D48_neto As Double
Public D48_iva As Double
Public D48_total As Double

Public UsuarioCorreo As String
Public ServerCorreo As String
Public ClaveCorreo As String
Public RUTACOMPROBANTE As String



Public diacierrecompra As String
Public conexionauditoria As rdoConnection
Public contemporal As rdoConnection
Public añotraspaso As String
Public cartola As String
Public clientesistema As String
Public ivaretenido As String
Public cuentadiferencia As String
Public rubrolocal As String
Public mercaderias As String
Public respuesta As String
Public scrut As String
Public CREANDO As String
Public controlgrilla As String
Public empresaactiva As String
Public EMPRESASPERMISO(4) As String
Public basebus As String
Public DATOSEMPRESA(30) As String
Public MENUPASO As MENU
Public DATOSARRENDATARIO(30) As String
Public FORMULARIOPASO As Form
Public permisos(100, 7)
Public PERMISOPROGRAMA(6)
Public ejecuta(100) As String
Public codigoempresa As String
Public nombreempresa As String
Public direccionempresa As String
Public comunaempresa As String
Public codigosii As String
Public giroempresa As String
Public cuentaperdida As String
Public cuentaganancia As String
Public cuentaiva As String
Public CUENTAPROVEEDOR As String
Public cuentaprovisiones As String
Public cuentagastoferia As String

Public PASO(15) As String
Public ctamadre As String
Public suma(10) As Variant ' para balance
Public sumas(10) As Variant ' para balance
Public difer(10) As Variant ' para balance
Public sumast(10) As Variant ' para balance
Public anterior(10) As Variant ' para balance
Public tipocue(10) As String ' para balance
Public anted As Variant ' para balance
Public anteh As Variant ' para balance
Public Titulos As String
Public iva As Integer
Public mescontabilizado As String
Public añocontabilizado As String
Public sw As String
Public dia As String
Public MES As String
Public año As String
Public SUMAR As Double
Public sumadebe As Double
Public sumahaber As Double
Public saldo As Double
Public varimonto As Double
Public VARINUM As Integer
Public fechasistema As Date
Public sumador As Double
Public LINEA As Integer
Public tipocentro As String
Public tipocuenta As String
Public varipaso As String
Public TITU(16) As String
Public titu1(13) As String
Public LINEAS As Double
Public largopagina As Integer
Public pagina As Integer
Public tituloinforme As String
Public rutempresa As String
Public totales As Double
Public descu As Double
Public DOCU2(10) As String
Public DOCU(10) As String
Public CANDO As Integer
Public CANDO2 As Integer
Public k As Integer ' contaDOR BUCLE
Public colu(20)
Public dato(20) As Variant
Public cancolu As Integer
Public palabra As String
Public tipodato(20) As String
Public db As rdoConnection
Public contadb As rdoConnection
Public conta As rdoConnection
Public conta2 As rdoConnection
Public teso As rdoConnection
Public ventaslocal As rdoConnection
Public gestion As rdoConnection
Public gestionrubro As rdoConnection
Public temporal As rdoConnection
Public tablaempresas(1000, 2)
Public Servidor As String
Public Usuario As String
Public password As String
Public basedatos As String
Public USUARIOSISTEMA As String
Public clavesistema As String
Public empresa As String
Public fecha As String
Public glosa As String
Public programa As String
Public largo2(10) As Integer
Public RETORNOPROGRAMA As String
Public retorno As String
Public formulario As Form
Public texto1 As TextBox
Public texto2 As TextBox
Public ayudatext As String
Public campos(49, 3) As Variant
Public EVENTO As Integer
Public nume As Double
Public sc As Integer
Public lar As Integer
Public condicion As String
Public largocero As Integer
Public valor As Double
Public snum As Integer
Public op As Integer
Public sl As Integer
Public status As Integer
Public procedimiento As Integer
Public MODIFI As Integer
Public cuentahonorarios As String
Public cuentacliente As String
Public retencion As String
Public ivadebito As String
Public ivacredito As String
Public ingresosporventa As String
Public cierrect As String
Public fechacierre As String
Public mensaje_nopermiso As String
Public rut_representante As String
Public nombre_representante As String
Public fechaflujo As String
Public tipoflujo As String
Public DIGITA_RUT_RUT As String
Public DIGITA_RUT_NOMBRE As String
Public DIGITA_CRCC_CODIGO As String
Public DIGITA_CRCC_NOMBRE As String
Public DIGITA_ANALISIS_CODIGO As String
Public DIGITA_ANALISIS_NOMBRE As String
Public DIGITA_CENTROS_CODIGO As String
Public DIGITA_CENTROS_NOMBRE As String
Public rut_enviasii As String
Public numeroresolucion As String
Public fecharesolucion As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal Milisegundos As Long)


'====================================================================================
'Rutina de conexion al servidor de bases de datos
'====================================================================================
Sub Conectarconta(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set conta = New rdoConnection
    conta.Connect = cadena_conexion
    conta.CursorDriver = rdUseServer
    conta.EstablishConnection
  
'    Call Conectarconta2(Servidor, clientesistema + "conta", Usuario, password)
End Sub
Sub Conectarteso(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set teso = New rdoConnection
    teso.Connect = cadena_conexion
    teso.CursorDriver = rdUseServer
    teso.EstablishConnection
    Call Conectarconta2(Servidor, clientesistema + "teso", Usuario, password)
End Sub

Sub Conectarconta2(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set conta2 = New rdoConnection
    conta2.Connect = cadena_conexion
    conta2.CursorDriver = rdUseServer
    conta2.EstablishConnection
End Sub

'Sub Conectartemporal(ByVal servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
'    Dim cadena_conexion As String
'    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
'    Set conta2 = New rdoConnection
'    contemporal.Connect = cadena_conexion
'    contemporal.CursorDriver = rdUseServer
'    contemporal.EstablishConnection
'
'End Sub

Sub Conectarventas(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set ventaslocal = New rdoConnection
    ventaslocal.Connect = cadena_conexion
    ventaslocal.CursorDriver = rdUseServer
    ventaslocal.EstablishConnection
End Sub
Sub Conectargestion(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set gestion = New rdoConnection
    gestion.Connect = cadena_conexion
    gestion.CursorDriver = rdUseServer
    gestion.EstablishConnection

End Sub
Sub Conectargestionrubro(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set gestionrubro = New rdoConnection
    gestionrubro.Connect = cadena_conexion
    gestionrubro.CursorDriver = rdUseServer
    gestionrubro.EstablishConnection
End Sub








'====================================================================================
'Rutina de conexion al servidor de bases de datos
'====================================================================================
Sub Conectar(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & ";PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set db = New rdoConnection
    db.Connect = cadena_conexion
    db.CursorDriver = rdUseServer
    db.EstablishConnection
    
End Sub

Sub Conectardb2(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & ";PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set contadb = New rdoConnection
    contadb.Connect = cadena_conexion
    contadb.CursorDriver = rdUseServer
    contadb.EstablishConnection
    
End Sub

Sub Conectartemporal(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & ";PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set temporal = New rdoConnection
    temporal.Connect = cadena_conexion
    temporal.CursorDriver = rdUseServer
    temporal.EstablishConnection
End Sub


'====================================================================================
'Rutina de conexion al servidor de bases de datos
'====================================================================================
Sub ConectarControlData(ByRef data As Adodc, ByVal Servidor As String, ByVal bd As String, ByVal Usuario As String, ByVal password As String, ByVal tabla As String)
        Dim cadena_conexion As String
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & ";PWD=" & password & "; UID=" & Usuario & ";OPTION=3"
        data.ConnectionString = cadena_conexion
        data.RecordSource = tabla
        data.Refresh
End Sub
Sub ConectarControlData2(ByRef Data2 As Adodc, ByVal Servidor As String, ByVal bd As String, ByVal Usuario As String, ByVal password As String, ByVal tabla As String)
        Dim cadena_conexion As String
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & password & "; UID=" & Usuario & "; OPTION=3"
        Data2.ConnectionString = cadena_conexion
        Data2.RecordSource = tabla
        Data2.Refresh
End Sub

Sub Conectar_BD()
''RUTINA PARA CONECTAR A LA BASE DE DATOS
'    basedatos = clientesistema + "conta" + empresaactiva
'    basebus = basedatos
'
'    'On Error GoTo controlerror
'        Call Conectar(Servidor, basebus, Usuario, password)
'    Exit Sub
'controlerror:
'
'    Resume Next
End Sub

Sub configurabasededatos()
'RUTINA PARA CONECTAR A LA BASE DE DATOS
    basedatos = clientesistema + "conta" + empresaactiva
    basebus = basedatos

    'On Error GoTo controlerror
        Call Conectardb2(Servidor, basebus, Usuario, password)
    Exit Sub
controlerror:

    Resume Next
End Sub
Sub CARGAPERMISO(programa)
For k = 1 To 100
If programa = permisos(k, 1) Then
PERMISOPROGRAMA(2) = permisos(k, 3)
PERMISOPROGRAMA(3) = permisos(k, 4)
PERMISOPROGRAMA(4) = permisos(k, 5)
PERMISOPROGRAMA(5) = permisos(k, 6)
End If
Next k
End Sub


Sub comas(monto, tipo As Integer)
If tipo = 12 Then varipaso = Format(monto, "###,###,###,###")
If tipo = 9 Then varipaso = Format(monto, "###,###,###")
If tipo = 6 Then varipaso = Format(monto, "###,###")
varipaso = Replace(varipaso, ".", ",")


End Sub
Sub ACEPTA(MENSAJE2)
preguntar.MENSAJE.Caption = MENSAJE2
respuesta = "N"
preguntar.Show vbModal
End Sub

Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub

Sub esfecha(ByRef dias As Integer, ByRef MES As Integer, ByRef ano As Integer)
If dias < 1 Or dias > 31 Then varipaso = "N": GoTo no
If MES < 1 Or MES > 12 Then varipaso = "N": GoTo no
If ano < 2005 Then varipaso = "N": GoTo no:
varipaso = "S"
no:

End Sub

Sub CONSULTAFECHAS(glosa)

fechas.Caption = glosa
fechas.Show vbModal


End Sub
Sub FORMATOCALENDARIO(ByRef calendario As Grid)
Dim FORMATOGRILLA(100, 20)
Rem DATOS DE LA COLUMNA
    calendario.DefaultFont.Size = 12
    calendario.DefaultFont.Bold = True
    
    
    
    FORMATOGRILLA(1, 1) = "DESDE"
    FORMATOGRILLA(1, 2) = "HASTA"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "D"
    
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    
    calendario.Cols = 4
    calendario.Rows = 2
    
    calendario.AllowUserResizing = False
    calendario.DisplayFocusRect = False
    calendario.ExtendLastCol = False
    calendario.BoldFixedCell = False
    
    calendario.DrawMode = cellOwnerDraw
    
    calendario.Appearance = Flat
    calendario.ScrollBarStyle = Flat
    calendario.FixedRowColStyle = Flat
    
   calendario.BackColorFixed = RGB(90, 158, 214)
   calendario.BackColorFixedSel = RGB(110, 180, 230)
   calendario.BackColorBkg = RGB(90, 158, 214)
   calendario.BackColorScrollBar = RGB(231, 235, 247)
   calendario.BackColor1 = RGB(231, 235, 247)
   calendario.BackColor2 = RGB(239, 243, 255)
   calendario.GridColor = RGB(148, 190, 231)
   calendario.Column(0).Width = 0
    
    For k = 1 To calendario.Cols - 1
        
        calendario.Cell(0, k).text = FORMATOGRILLA(1, k)
        calendario.Column(k).Width = Val(FORMATOGRILLA(2, k)) * calendario.DefaultFont.Size
        
        
        calendario.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        calendario.Column(k).FormatString = FORMATOGRILLA(4, k)
        calendario.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then calendario.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then calendario.Column(k).CellType = cellCalendar
        
    Next k
End Sub
Public Sub retornofecha(ByRef etiqueta1 As label, ByRef etiqueta2 As label)
    Load fechas
    Set fechas.fecha1 = etiqueta1
    Set fechas.fecha2 = etiqueta2
    fechas.calendario.Cell(1, 1).text = etiqueta1.Caption
    fechas.calendario.Cell(1, 2).text = etiqueta2.Caption
    fechas.Show vbModal
End Sub

Public Sub CENTRAR(ByRef FORMU As Form)
FORMU.Left = (Screen.Width - FORMU.Width) / 2
FORMU.Top = (Screen.Height - FORMU.Height) / 2 - 700

End Sub
Public Sub esfechareal(ByRef dia As TextBox, MES As TextBox, año As TextBox, tipo As String)
Dim maximo As String

If dia.text <> "" Then
        maximo = "31"
        If MES.text = "01" Then maximo = "31"
        If MES.text = "02" Then maximo = "28"
        
        If MES.text = "02" And año.text = "2008" Then maximo = "29"
        If MES.text = "02" And año.text = "2012" Then maximo = "29"
        If MES.text = "02" And año.text = "2016" Then maximo = "29"
        If MES.text = "02" And año.text = "2020" Then maximo = "29"
        If MES.text = "02" And año.text = "2024" Then maximo = "29"
        If MES.text = "02" And año.text = "2028" Then maximo = "29"
        If MES.text = "02" And año.text = "2032" Then maximo = "29"
        If MES.text = "02" And año.text = "2036" Then maximo = "29"
        If MES.text = "02" And año.text = "2040" Then maximo = "29"
        If MES.text = "02" And año.text = "2044" Then maximo = "29"
        If MES.text = "02" And año.text = "2048" Then maximo = "29"
        If MES.text = "02" And año.text = "2052" Then maximo = "29"
        
        If MES.text = "03" Then maximo = "31"
        If MES.text = "04" Then maximo = "30"
        If MES.text = "05" Then maximo = "31"
        If MES.text = "06" Then maximo = "30"
        If MES.text = "07" Then maximo = "31"
        If MES.text = "08" Then maximo = "30"
        If MES.text = "09" Then maximo = "30"
        If MES.text = "10" Then maximo = "31"
        If MES.text = "11" Then maximo = "30"
        If MES.text = "12" Then maximo = "31"
        
        If dia.text < "01" Or dia.text > maximo Then
        dia.text = ""
        dia.SetFocus
        
        End If

End If
If MES.text <> "" Then
        maximo = "12"
        If MES.text < "01" Or MES.text > maximo Then
        
        MES.text = ""
        MES.SetFocus
        
        End If

End If
If año.text <> "" Then
        maximo = "2100"
        If año.text < "2000" Or año.text > maximo Then
        año.text = ""
        año.SetFocus
        
        End If

End If


End Sub

Sub Conectar_Auditoria()
'RUTINA PARA CONECTAR A LA BASE DE DATOS DE AUDITORIA
    Dim bd As String
    bd = clientesistema + "auditoria"
    
    On Error GoTo controlerror
        Call Conectar2(Servidor, bd, Usuario, password)
    Exit Sub

controlerror:
    
    Resume Next
End Sub
Sub Conectar2(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & ";PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set conexionauditoria = New rdoConnection
   conexionauditoria.Connect = cadena_conexion
   conexionauditoria.CursorDriver = rdUseServer
   conexionauditoria.EstablishConnection
End Sub
Sub actualizamayor(EVENTO, codigo, monto, DH, tipo, rut, CRCC, MES, año)
    Dim SUMAVALOR As Double
    Dim tienerut As Boolean
    Dim TIENECRCC As Boolean
    sqlconta.audit = False
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = ""
    
    If DH = "D" Then campos(2, 0) = "debe" + MES
    If DH = "H" Then campos(2, 0) = "haber" + MES
    campos(3, 0) = ""
    
    condicion = "codigo=" + "'" + codigo + "' and año ='" + año + "' "
    
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
    If sqlconta.status = 4 Then Exit Sub
    tienerut = sqlconta.response(3, 3)
    TIENECRCC = sqlconta.response(4, 3)
    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    If EVENTO = "+" Then campos(2, 1) = Str(sqlconta.response(2, 3) + Val(monto))
    If EVENTO = "-" Then campos(2, 1) = Str(sqlconta.response(2, 3) - Val(monto))
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    tienerut = leertiene(codigo, "1")
    TIENECRCC = leertiene(codigo, "2")
    If rut <> "" And tienerut = True Then Call actualizactacte(EVENTO, codigo, rut, monto, DH, MES, año)
    If CRCC <> "" And TIENECRCC = True Then Call actualizacrcc(EVENTO, CRCC, codigo, monto, DH, MES, año)
    End If
Rem actualiza cuenta madre
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = ""
    If DH = "D" Then campos(2, 0) = "debe" + MES
    If DH = "H" Then campos(2, 0) = "haber" + MES
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + Mid(codigo, 1, 4) + "0000" + "' and año ='" + año + "' order by codigo"
    
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    
    
    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    If EVENTO = "+" Then campos(2, 1) = Str(sqlconta.response(2, 3) + Val(monto))
    If EVENTO = "-" Then campos(2, 1) = Str(sqlconta.response(2, 3) - Val(monto))
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
Rem actualiza cuenta principal
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = ""
    If DH = "D" Then campos(2, 0) = "debe" + MES
    If DH = "H" Then campos(2, 0) = "haber" + MES
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + Mid(codigo, 1, 2) + "000000" + "' and año ='" + año + "' order by codigo"
    
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    
    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    If EVENTO = "+" Then campos(2, 1) = Str(sqlconta.response(2, 3) + Val(monto))
    If EVENTO = "-" Then campos(2, 1) = Str(sqlconta.response(2, 3) - Val(monto))
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    
    sqlconta.audit = True
End Sub

Sub actualizactacte(EVENTO, tipo, rut, monto, DH, MES, año)

    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    If DH = "D" Then campos(2, 0) = "debe" + MES
    If DH = "H" Then campos(2, 0) = "haber" + MES
    campos(3, 0) = ""
    condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año ='" + año + "'"
    campos(0, 2) = "saldosctacte"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    If EVENTO = "+" Then campos(2, 1) = Str(sqlconta.response(2, 3) + Val(monto))
    If EVENTO = "-" Then campos(2, 1) = Str(sqlconta.response(2, 3) - Val(monto))

    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
End Sub

Sub actualizacrcc(EVENTO, CRCC, cuenta, monto, DH, MES, año)
    

    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If DH = "D" Then campos(2, 0) = "debe" + MES
    If DH = "H" Then campos(2, 0) = "haber" + MES
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + CRCC + "' and año ='" + año + "' and cuenta='" + cuenta + "' order by codigo"
    
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    
    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    If EVENTO = "+" Then campos(2, 1) = Str(sqlconta.response(2, 3) + Val(monto))
    If EVENTO = "-" Then campos(2, 1) = Str(sqlconta.response(2, 3) - Val(monto))
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    Rem actualiza cuenta madre
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If DH = "D" Then campos(2, 0) = "debe" + MES
    If DH = "H" Then campos(2, 0) = "haber" + MES
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + Mid(CRCC, 1, 2) + "00" + "' and año ='" + año + "' and cuenta='" + cuenta + "' order by codigo"
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    If sqlconta.status = 4 Then Exit Sub
    
    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    If EVENTO = "+" Then campos(2, 1) = Str(sqlconta.response(2, 3) + Val(monto))
    If EVENTO = "-" Then campos(2, 1) = Str(sqlconta.response(2, 3) - Val(monto))
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then Exit Sub
    
    'If sqlconta.status = 4 Then MsgBox (condicion + " no existe "): Exit Sub
    
End Sub

Sub grabaprincipal(programa)
    Dim cadena As String
    Dim csql2 As New rdoQuery
                    Set csql2.ActiveConnection = conAuditoria
                    
                    
                    cadena = "INSERT INTO auditoriacontabilidad ("
                    cadena = cadena + "programa,fecha,hora,usuario,evento,tabla) VALUES ( "
                    cadena = cadena & "'" & programa & "','" & Format(Date, "yyyy-mm-dd") & "','" & Time & "','" & usuarioauditoria & "','0','" + nombreempresa + "')"
                    csql2.sql = cadena
                    csql2.Execute
End Sub

Sub crearcajera(rut)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion

            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + "11250005" + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,'','' "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestrocajeras as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' "
            csql.sql = csql.sql & "on duplicate key update tipo=tipo"
            csql.Execute
            
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + "11250005" + "',mc.rut "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestrocajeras as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' "
            csql.sql = csql.sql & "on duplicate key update tipo=tipo "
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            


End Sub
Public Function leerrutctacte(tipo, numero, fecha) As String
Dim csql As New rdoQuery
Dim resultados  As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select rutctacte from movimientoscontables "
csql.sql = csql.sql & "where tipo='" & tipo & "' and numero='" & numero & "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "'  and rutctacte<>''  group by rutctacte"
csql.Execute

 leerrutctacte = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerrutctacte = resultados(0)
End If

End Function

Public Function leeripc(MES, año) As Double
Dim csql As New rdoQuery
Dim resultados  As rdoResultset

Set csql.ActiveConnection = conta

csql.sql = "select porcentaje from ipc where mes='" & MES & "' and año='" & año & "' "

csql.Execute

 leeripc = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leeripc = resultados(0)
End If

End Function


Public Function leevencimiento(ORDEN, dias, fecharecepcion) As String
Dim csql As New rdoQuery
Dim resultados  As rdoResultset
Dim FECHACONSULTA As String
FECHACONSULTA = Format(DateAdd("D", dias, fecharecepcion), "yyyy-mm-dd")



Set csql.ActiveConnection = conta

csql.sql = "select * from " + clientesistema + "conta.maximopagoproveedores where fecha>'" + FECHACONSULTA + "' and empresa='" + empresaactiva + "' "

csql.Execute

 
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    leevencimiento = Format(resultados(1), "yyyy-mm-dd")
    Exit Function
    resultados.MoveNext
    
    Wend
    
End If




End Function


Public Function TIENEANALISIS(cuenta) As Boolean
Dim csql As New rdoQuery
Dim resultados  As rdoResultset

Set csql.ActiveConnection = conta

csql.sql = "select * from presupuesto_detalle where cuenta ='" & cuenta & "' "
csql.Execute

TIENEANALISIS = False

If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
TIENEANALISIS = True

End If

End Function

Public Function leernombreanalisis(cuenta, codigo) As String
Dim csql As New rdoQuery
Dim resultados  As rdoResultset

Set csql.ActiveConnection = conta

csql.sql = "select nombre from presupuesto_detalle where cuenta ='" & cuenta & "' and codigo='" + codigo + "' "
csql.Execute

leernombreanalisis = ""

If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
leernombreanalisis = resultados(0)

End If

End Function


Public Function leerabonofactura(TIPORIGINAL, tipo, numero, rut, cuenta, DH, fechafac, Optional ByVal NoLeerFactorizada As Boolean) As Double
Dim csql As New rdoQuery
Dim csql2 As New rdoQuery
Dim cSql3 As New rdoQuery
Dim resultados  As rdoResultset
Dim resultados3  As rdoResultset
Dim tipo2 As String
Dim tipoDTE As String
Set csql.ActiveConnection = contadb
Set cSql3.ActiveConnection = contadb
If IsMissing(NoLeerFactorizada) = True Then NoLeerFactorizada = False


csql.sql = "select ifnull(sum(monto),'0') from movimientoscontables "
csql.sql = csql.sql & "where tipodocumento='" & tipo & "' and numerodocumento='" & numero & "' and codigocuenta='" & cuenta & "' and rutctacte='" + rut + "' and dh='" + DH + "' and fecha>='" & Format(fechafac, "yyyy-mm-dd") & "'"
csql.sql = csql.sql & "group by numerodocumento "
csql.Execute
'ariel  este pone lento
leerabonofactura = 0
tipoDTE = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerabonofactura = resultados(0)
Else
    'linea nueva consulta si esta factorizada la da por pagada
    
    If NoLeerFactorizada = False Then
        
        If tipo = "4" Or tipo = "FC" Then tipo = "4": tipoDTE = "33"
        
        If tipo = "5" Then tipo = "5": tipoDTE = "56"
        If tipo = "6" Then tipo = "6": tipoDTE = "61"
        If tipo = "0" Or tipo = "EE" Then tipo = "0": tipoDTE = "34"
        
        cSql3.sql = "select ifnull(sum(total),'0') from " & clientesistema & "fae" & CONFI_EMPRESAFAE & ".sv_dte_cedidos_" & CONFI_EMPRESAFAE
        cSql3.sql = cSql3.sql & " where tipo='" & tipoDTE & "'"
        cSql3.sql = cSql3.sql & " and numero ='" & Val(numero) & "' "
        cSql3.sql = cSql3.sql & " and cedente_rut='" & Val(Mid(rut, 1, 9)) & "-" & Right(rut, 1) & "'"
        cSql3.Execute
        
        leerabonofactura = 0
        If cSql3.RowsAffected > 0 Then
            Set resultados3 = cSql3.OpenResultset
            leerabonofactura = resultados3(0)
        End If
        cSql3.Close
    End If
End If

csql.Close



Set csql2.ActiveConnection = contadb

If tipo = "FC" Then tipo = "1": tipo2 = "4"
If tipo = "ND" Then tipo = "2": tipo2 = "5"
If tipo = "NC" Then tipo = "3": tipo2 = "6"
If tipo = "EE" Then tipo = "0": tipo2 = "0"

 
 
 
If tipo <> "BH" Then
    csql2.sql = "update facturasdecompras set abono='" & leerabonofactura & "' "
    csql2.sql = csql2.sql & "where (tipo='" & tipo & "' or tipo='" + tipo2 + "') and numero='" & numero & "' and  rut='" + rut + "' and fecha='" & Format(fechafac, "yyyy-mm-dd") & "'  "
    csql2.Execute
Else
   csql2.sql = "update boletasdehonorarios set abono='" & leerabonofactura & "' "
    csql2.sql = csql2.sql & "where (tipo='1') and numero='" & numero & "' and  rut='" + rut + "' and fecha='" & Format(fechafac, "yyyy-mm-dd") & "'  "
    csql2.Execute
End If
'Call sincronizadatos(csql2.sql, contadb, "")


csql2.Close
NOESTAENELSII = True
If tipo <> "BH" And (TIPORIGINAL = "4" Or TIPORIGINAL = "0") And Format(fechafac, "yyyy-mm-dd") > "2017-08-01" Then
NOESTAENELSII = False
    If USUARIOSISTEMA <> "RSCHICK" And USUARIOSISTEMA <> "JMONTECINOS" Then
        If Verifica_Permiso("Ingreso de Comprobantes Contables", "AUTORIZA") = False Then
            If ESTAENSII_todo(TIPORIGINAL, numero, rut, fechafac, empresaactiva) = False Then
                NOESTAENELSII = True
                leerabonofactura = 999999
            End If
        End If
    End If
End If

End Function


Public Function leerabonofacturaFactoring(tipo, numero, rut, cuenta, DH, fechafac, Optional ByVal NoLeerFactorizada As Boolean) As Double
Dim csql As New rdoQuery
Dim csql2 As New rdoQuery
Dim cSql3 As New rdoQuery
Dim resultados  As rdoResultset
Dim resultados3  As rdoResultset
Dim tipo2 As String
Dim tipoDTE As String
Set csql.ActiveConnection = contadb
Set cSql3.ActiveConnection = contadb
If IsMissing(NoLeerFactorizada) = True Then NoLeerFactorizada = False

csql.sql = "select ifnull(sum(monto),'0') from movimientoscontables "
csql.sql = csql.sql & "where tipodocumento='" & tipo & "' and numerodocumento='" & numero & "' and codigocuenta='" & cuenta & "' and rutctacte='" + rut + "' and dh='" + DH + "' and fecha>='" & Format(fechafac, "yyyy-mm-dd") & "'"
csql.sql = csql.sql & "group by numerodocumento "
csql.Execute
'ariel  este pone lento
leerabonofacturaFactoring = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerabonofacturaFactoring = resultados(0)
Else
''    'linea nueva consulta si esta factorizada la da por pagada
'    If NoLeerFactorizada = False Then
'        If tipo = "FC" Then tipo = "1": tipoDTE = "33"
'        If tipo = "ND" Then tipo = "2": tipoDTE = "56"
'        If tipo = "NC" Then tipo = "3": tipoDTE = "61"
'        If tipo = "EE" Then tipo = "0": tipoDTE = "34"
'
'        cSql3.sql = "select ifnull(sum(total),'0') from " & clientesistema & "fae" & CONFI_EMPRESAFAE & ".sv_dte_cedidos_" & CONFI_EMPRESAFAE
'        cSql3.sql = cSql3.sql & " where tipo='" & tipoDTE & "'"
'        cSql3.sql = cSql3.sql & " and numero ='" & Val(numero) & "' "
'        cSql3.sql = cSql3.sql & " and cedente_rut='" & Val(Mid(rut, 1, 9)) & "-" & Right(rut, 1) & "'"
'        cSql3.Execute
'
'        leerabonofactura = 0
'        If cSql3.RowsAffected > 0 Then
'            Set resultados3 = cSql3.OpenResultset
'            leerabonofactura = resultados3(0)
'        End If
'        cSql3.Close
'    End If
End If

csql.Close

Set csql2.ActiveConnection = contadb

If tipo = "FC" Then tipo = "1": tipo2 = "4"
If tipo = "ND" Then tipo = "2": tipo2 = "5"
If tipo = "NC" Then tipo = "3": tipo2 = "6"
If tipo = "EE" Then tipo = "0": tipo2 = "0"

 
If tipo <> "BH" Then
    csql2.sql = "update facturasdecompras set abono='" & leerabonofacturaFactoring & "' "
    csql2.sql = csql2.sql & "where (tipo='" & tipo & "' or tipo='" + tipo2 + "') and numero='" & numero & "' and  rut='" + rut + "' and fecha='" & Format(fechafac, "yyyy-mm-dd") & "'  "
    csql2.Execute
Else
   csql2.sql = "update boletasdehonorarios set abono='" & leerabonofacturaFactoring & "' "
    csql2.sql = csql2.sql & "where (tipo='1') and numero='" & numero & "' and  rut='" + rut + "' and fecha='" & Format(fechafac, "yyyy-mm-dd") & "'  "
    csql2.Execute
End If
'Call sincronizadatos(csql2.sql, contadb, "")


csql2.Close


End Function


Sub CARGAGRILLAtiponombre(frm As Form, Grid2 As Grid)
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 3
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
'    GRID2BackColor1 = RGB(231, 235, 247)
'    GRID2BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
    
    End Sub

Public Function leercodigolocal(empresacontable) As String
Dim csql As New rdoQuery
Dim resultados  As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select codigo from " + clientesistema + "gestion.g_maestroempresas where codigocontable ='" & empresacontable & "' "
csql.Execute

leercodigolocal = ""

If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
leercodigolocal = resultados(0)

End If

End Function

Public Function leertotalremuneraciones(empresacontable, codigo, MES, año, todo) As Double
Dim csql As New rdoQuery
Dim resultados  As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select ifnull(sum(monto),0) from " + clientesistema + "remu" + empresacontable + "."
csql.sql = csql.sql & "calculoliquidaciones where codigo='" + codigo + "' and mes='" + MES + "'  "
csql.sql = csql.sql & "and año='" + año + "' "
If todo = "" Then
    csql.sql = csql.sql & "AND origen='' "
End If
csql.Execute

leertotalremuneraciones = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
leertotalremuneraciones = resultados(0)

End If

End Function

Public Function leerdatostrabajador(ByVal buscados As String, _
                            tabla As String, _
                            condicion As String, _
                            conactiva As rdoConnection) As String
                            
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = conactiva
    
    csql.sql = "SELECT " & buscados
    csql.sql = csql.sql & " FROM " & tabla
    csql.sql = csql.sql & " WHERE " & condicion
    csql.Execute
    If csql.RowsAffected > 0 Then ' ENTONCES EXISTE
    Set resultados = csql.OpenResultset
        If IsNull(resultados(0)) = False Then
        leerdatostrabajador = resultados(0)
        Else
        leerdatostrabajador = "9999-99-99"
    End If
        Else
        leerdatostrabajador = "0"
    End If
    csql.Close
End Function


Public Function LEERSALDOSCTACTEmovi(tipoctacte, cuenta, empresa) As Double

   Dim resultados3 As rdoResultset
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = ""
    condicion = "tipo=" + "'" + tipoctacte + "' and rut='" + cuenta + "' and año='" + Format(fechasistema, "yyyy") + "'"
    campos(0, 2) = clientesistema + "conta" + empresa + ".saldosctacte"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
    saldo = sumador
    Rem acumula fecha
        fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    For k = 0 To 12
    ctadebe2(k) = 0
    ctahaber2(k) = 0
    Next k
    
        Set cSql3.ActiveConnection = contadb
        cSql3.sql = "SELECT SUM(monto),dh,mes "
        cSql3.sql = cSql3.sql + "FROM " + clientesistema + "conta" + empresa + ".movimientoscontables where codigocuenta='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + Format(fechasistema, "yyyy-mm-dd") + "' "
        cSql3.sql = cSql3.sql + "GROUP BY DH,mes"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
         If resultados3(1) = "D" Then
         saldo = saldo + resultados3(0)
         ctadebe2(0) = ctadebe2(0) + resultados3(0)
         ctadebe2(resultados3(2)) = ctadebe2(resultados3(2)) + resultados3(0)
         End If
         
         If resultados3(1) = "H" Then
         saldo = saldo - resultados3(0)
         ctahaber2(0) = ctahaber2(0) + resultados3(0)
         ctahaber2(resultados3(2)) = ctahaber2(resultados3(2)) + resultados3(0)
         
         End If
         
             
         resultados3.MoveNext
           
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If
        LEERSALDOSCTACTEmovi = saldo
        
End Function
Public Function LEERSALDOSCTACTEmovisolo(tipoctacte, cuenta, empresa, fecha2) As Double
    Dim resultados4 As rdoResultset
   Dim resultados3 As rdoResultset
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim cSql4 As New rdoQuery
    
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    
    Dim mesante As Integer
    
'    campos(0, 0) = "tipo"
'    campos(1, 0) = "rut"
'    campos(2, 0) = "año"
'    campos(3, 0) = "debeanterior"
'    campos(4, 0) = "haberanterior"
'    campos(5, 0) = ""
'    condicion = "tipo=" + "'" + tipoctacte + "' and rut='" + cuenta + "' and año='" + Format(fechasistema, "yyyy") + "'"
'    campos(0, 2) = clientesistema + "conta" + empresa + ".saldosctacte"
'    op = 5
'    sqlconta.response = campos
'    Set sqlconta.conexion = contadb
'    Call sqlconta.sqlconta(op, condicion)
'    SUMADOR = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
    saldo = sumador
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    Set cSql3.ActiveConnection = contadb
   
 
    cSql3.sql = "SELECT SUM(IF(dh='D',monto,0)),SUM(IF(dh='H',monto,0)),dh "
    cSql3.sql = cSql3.sql + "FROM " + clientesistema + "conta" + empresa + ".movimientoscontables where codigocuenta='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + Format(fecha2, "yyyy-mm-dd") + "' "
    cSql3.sql = cSql3.sql + "GROUP BY DH "
    cSql3.Execute
   
    If cSql3.RowsAffected > 0 Then
     
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
         saldo = saldo + resultados3(0) - resultados3(1)
         
         resultados3.MoveNext
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If
        LEERSALDOSCTACTEmovisolo = saldo
        
'    Set cSql4.ActiveConnection = contadb
'    cSql4.sql = "SELECT SUM(monto),dh "
'    cSql4.sql = cSql4.sql + "FROM " + clientesistema + "conta" + empresa + ".movimientoscontables where codigocuenta='" + tipoctacte + "' and rutctacte='" + cuenta + "' and mes='" & Format(fecha2, "mm") & "' and año='" & Format(fecha2, "yyyy") & "' "
'    cSql4.sql = cSql4.sql + "GROUP BY dh "
'    cSql4.Execute
'   ctadebemes = 0
'   ctahabermes = 0
'
'    If cSql4.RowsAffected > 0 Then
'
'
'        Set resultados4 = cSql4.OpenResultset
'
'         While Not resultados4.EOF
'         If resultados4(1) = "D" Then
'         ctadebemes = resultados4(0)
'         End If
'         If resultados4(1) = "H" Then
'         ctahabermes = resultados4(0)
'         End If
'         resultados4.MoveNext
'         Wend
'          resultados4.Close
'            Set resultados4 = Nothing
'
'        End If
'
'
        
        
End Function

'Public Function leerabonofactura_CU(tipo, numero, rut, cuenta, DH) As Double
'Dim csql As New rdoQuery
'Dim csql2 As New rdoQuery
'
'Dim resultados  As rdoResultset
'Dim tipo2 As String
'
'Set csql.ActiveConnection = contadb
'
'If tipo = "CT" Then
'    tipo = "FC"
'End If
'
'If tipo2 = "4" Or tipo = "1" Then tipo = "FC"
'If tipo2 = "5" Or tipo = "2" Then tipo = "ND"
'If tipo2 = "6" Or tipo = "3" Then tipo = "NC"
'
'
'csql.sql = "select ifnull(sum(monto),'0') from movimientoscontables "
'csql.sql = csql.sql & "where tipodocumento='" & tipo & "' and numerodocumento='" & numero & "' and codigocuenta='" & cuenta & "' and rutctacte='" + rut + "' and dh='" + DH + "'  "
'csql.sql = csql.sql & "group by numerodocumento "
'csql.Execute
''ariel  este pone lento
'leerabonofactura = 0
'If csql.RowsAffected > 0 Then
'    Set resultados = csql.OpenResultset
'    leerabonofactura = resultados(0)
'End If
'csql.Close
'
'Set csql2.ActiveConnection = contadb
'
'If tipo = "FC" Then tipo = "1": tipo2 = "4"
'If tipo = "ND" Then tipo = "2": tipo2 = "5"
'If tipo = "NC" Then tipo = "3": tipo2 = "6"
'
'csql2.sql = "update facturasdecompras set abono='" & leerabonofactura & "' "
'csql2.sql = csql2.sql & "where (tipo='" & tipo & "' or tipo='" + tipo2 + "') and numero='" & numero & "' and  rut='" + rut + "' "
'csql2.Execute
''Call sincronizadatos(csql2.sql, contadb, "")
'
'
'csql2.Close
'
'
'End Function
'
