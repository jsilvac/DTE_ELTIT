Attribute VB_Name = "Funciones"

Option Explicit
    
Public email_cuenta_usuario As String
Public email_cuenta_clave As String
Public email_cuenta_server As String
    
    
Public certificado_sii As String
Public clave_certificado_sii As String

    Public dte_cli_envia As String
    Public dte_respuesta_sii As String
    Public dte_respuesta_cliente As String
    
    Public dte_email_envio As String

Public CODAUTREC As String


    Public proporcional As String
    Public empresa_fae As String
    Public FOLIOini2 As String
    Public foliofin2 As String
    
    Public EXENTO As Double
    Public diesel2 As Double
    Public FOLIO_INI As String
    Public FOLIO_FIN As String
    Public codigoae As String
    Public da0 As String
    Public da1 As String
    Public da2 As String
    Public da3 As String
    Public da4 As String
    '
    Public da5 As String
    Public da6 As String
    Public da7 As String
    
    Public programafiltro As String
    Public cuentadebe As String
    Public cuentahaber As String
    Public tienecosto As String
    Public NOMBREDOCUMENTO As String
    Public rutaUpdate As String
    Public rutadestino As String
    Public nombregirador As String
    Public VENCIMIENTOREAL As String
    Public MONTOREAL As Double
    Public ctaanterior As Double
    Public ctadebe As Double
    Public ctahaber As Double
    Public ctasaldo As Double
    Public glosaeliminacionsistema As String
    Public LINEAC As Integer
    
    Public publicidad As Boolean
    Public empresarelacionada As Boolean
    Public solicitaeliminacion As String

    Public RUTA As String
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



    Public Sub AgregaFavoritos(ByRef Usuario, ByRef sistema, ByRef Aplicacion, Optional glosa As String)
    Dim csql1 As New rdoQuery
    Dim CONSULTA As String
    Set csql1.ActiveConnection = contadb
    
                csql1.sql = "insert ignore into " & clientesistema & "menu.quick_menu"
                csql1.sql = csql1.sql & "(usuario,sistema,aplicacion,glosa) values ('" & Usuario
                csql1.sql = csql1.sql & "','" & sistema & "','" & Aplicacion
                csql1.sql = csql1.sql & "','" & glosa & "')"
                csql1.Execute
'                Call SincronizaDatos(csql1.sql, conta)
                csql1.Close
                MsgBox glosa & " ha sido agregado a mis favoritos"
         
End Sub
Public Sub generarcomprobante(empresa, tipocomprobante, numerocomprobante, grilla As Grid, fecha)
  Dim k As Double
  Dim i As Double
  Dim nuevaempresa As String
  Dim rutprove As String
  Dim tipo As String
  Dim RUTEMPRE As String
  If tipocomprobante = "PA" Then tipo = "TP"
  If tipocomprobante = "CE" Then tipo = "TC"
  If tipocomprobante = "NG" Then tipo = "TN"
  
  For i = 1 To grilla.Rows - 1
    If grilla.Cell(i, 13).text <> "" Then
        rutprove = grilla.Cell(i, 13).text
        nuevaempresa = leerempresaproveedor(rutprove)
        Exit For
    End If
  Next i
  Dim NUEVOHD As String
  
  For k = 1 To grilla.Rows - 1
  If grilla.Cell(k, 9).text = "H" Then
  NUEVOHD = "D"
  Else
  NUEVOHD = "H"
  End If
    
    If grilla.Cell(k, 1).text & grilla.Cell(k, 2).text & grilla.Cell(k, 3).text = "11120001" Then
        
        Call grabarenlazado(tipo, empresaactiva & Mid(numerocomprobante, 3, 8), k, fecha, "11100010", grilla.Cell(k, 4).text, grilla.Cell(k, 5).text, grilla.Cell(k, 6).text, grilla.Cell(k, 8).text, NUEVOHD, nuevaempresa, "")
    Else
        If grilla.Cell(k, 1).text & grilla.Cell(k, 2).text & grilla.Cell(k, 3).text = "11130001" Then
             Call grabarenlazado(tipo, empresaactiva & Mid(numerocomprobante, 3, 8), k, fecha, "11500160", grilla.Cell(k, 4).text, grilla.Cell(k, 5).text, grilla.Cell(k, 6).text, grilla.Cell(k, 8).text, NUEVOHD, nuevaempresa, "")

        Else
            'NO FUNCIONO Format(Replace(rutempresa, "-", ""), "0000000000") LO DEJABA SIN 0 AL PRINCIPIO
            If InStr(LCase(rutempresa), "k") > 0 Then
                RUTEMPRE = Format(Replace(rutempresa, "-K", ""), "000000000")
                RUTEMPRE = RUTEMPRE & rut(RUTEMPRE)
            Else
                RUTEMPRE = Format(Replace(rutempresa, "-", ""), "0000000000")
            End If
            Call grabarenlazado(tipo, empresaactiva & Mid(numerocomprobante, 3, 8), k, fecha, "11200029", grilla.Cell(k, 4).text, grilla.Cell(k, 5).text, grilla.Cell(k, 6).text, grilla.Cell(k, 8).text, NUEVOHD, nuevaempresa, RUTEMPRE)
        End If
    End If
  Next k

End Sub
Public Function leerempresaproveedor(ByVal rutpro As String) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = conta
    If Mid(rutpro, 1, 1) = "0" Then
    rutpro = Mid(rutpro, 2, 9)
    rutpro = Mid(rutpro, 1, 8) & "-" & Mid(rutpro, 9, 1)
    End If
    csql.sql = "select codigoempresa from maestroempresas where rut='" & rutpro & "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerempresaproveedor = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
End Function

Public Function verificacomprobante(empr, numero, tipo, fecha) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
 
    csql.sql = "select numero from " & clientesistema & "conta" & empr & ".movimientoscontables "
    csql.sql = csql.sql & "where numero='" & numero & "' and tipo='" & tipo & "' and fecha='" & fecha & "' "
    csql.Execute
    verificacomprobante = False
    If csql.RowsAffected > 0 Then
       verificacomprobante = True
    End If
    csql.Close
    Set csql = Nothing
    
End Function
Public Sub grabarenlazado(tipo, numero, LINEA, fecha, cuenta, glosacontable, tipodocumento, numerodocumento, monto, DH, empresa, rutprove)
    Dim w As Long
    Dim tipo2 As String
    LINEA = Format(LINEA, "000")
    If tipodocumento = "FC" Then tipodocumento = "FA"
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "glosacontable"
    campos(6, 0) = "tipodocumento"
    campos(7, 0) = "numerodocumento"
    campos(8, 0) = "fechadocumento"
    campos(9, 0) = "fechavencimiento"
    campos(10, 0) = "monto"
    campos(11, 0) = "dh"
    campos(12, 0) = "creadopor"
    campos(13, 0) = "mes"
    campos(14, 0) = "año"
    campos(15, 0) = "rutctacte"
    campos(16, 0) = "centrocosto"
    campos(17, 0) = "fechacreacion"
    campos(18, 0) = "horacreacion"
    campos(19, 0) = "rutproveedor"
    campos(20, 0) = "cuenta_presupuesto"
    campos(21, 0) = "centro_gastos"
    campos(22, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = cuenta
    campos(5, 1) = leerdatos(conta, "maestroempresas", "nombre", "codigoempresa='" & empresaactiva & "'")
    campos(6, 1) = tipodocumento
    campos(7, 1) = numerodocumento
    campos(8, 1) = campos(3, 1)
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Replace(monto, ",", ".")
     
    
    campos(11, 1) = DH
    campos(12, 1) = USUARIOSISTEMA
    campos(13, 1) = Format(fecha, "mm")
    campos(14, 1) = Format(fecha, "yyyy")
    
    campos(15, 1) = Format(rutprove, "0000000000")
    campos(16, 1) = ""
    campos(17, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(18, 1) = Time$
    campos(19, 1) = Format(rutprove, "0000000000")
    campos(20, 1) = ""
    campos(21, 1) = ""
    
    campos(0, 2) = clientesistema & "conta" & empresa & ".movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
     
End Sub
Public Function verificasiexiste(rutcontable, CUENTABANCO, FECHACONTABLE, tipocontable, glosacontable) As Boolean
      Dim condicion As String
        Dim op As Integer
        campos(0, 0) = "numero"
        campos(1, 0) = ""
        
        campos(0, 2) = "movimientoscontables"
    
    condicion = "tipo = '" & tipocontable & "' and codigocuenta='" & CUENTABANCO & "' and  "
    condicion = condicion & " mes='" & Format(FECHACONTABLE, "mm") & "' and  año='" & Format(FECHACONTABLE, "yyyy") & "' "
    condicion = condicion & "and glosacontable like '" & glosacontable & "%' and linea='1' "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        verificasiexiste = True
    Else
        verificasiexiste = False
    End If
End Function
Public Function verificasiexiste2(rutcontable, CUENTABANCO, FECHACONTABLE, tipocontable, glosacontable) As Boolean
      Dim condicion As String
        Dim op As Integer
        campos(0, 0) = "numero"
        campos(1, 0) = ""
        
        campos(0, 2) = "movimientoscontables"
    
    condicion = "tipo = '" & tipocontable & "' and codigocuenta='" & CUENTABANCO & "' and  "
    condicion = condicion & " mes='" & Format(FECHACONTABLE, "mm") & "' and  año='" & Format(FECHACONTABLE, "yyyy") & "' "
    condicion = condicion & "and glosacontable like '" & glosacontable & "%' and linea='1' and rutctacte='" & rutcontable & "' "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        verificasiexiste2 = True
    Else
        verificasiexiste2 = False
    End If
End Function

Public Function leerNombreCuentaMayor(ByVal codigo As String, tipo)
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo = '" & codigo & "' "
    
    Select Case tipo
        Case 1
            condicion = condicion & " and ctacte = 1"
        Case 2
            condicion = condicion & "and crcc = 1"
        Case 3
            condicion = condicion & "and banco = 1"
    End Select
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerNombreCuentaMayor = sqlconta.response(0, 3)
    Else
        leerNombreCuentaMayor = ""
    End If
End Function
Public Function leerNombreCuentaMayorempresa(ByVal codigo As String, tipo, empre)
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = cliente_sql & "conta" & empre & ".cuentasdelmayor"
    
    condicion = "codigo = '" & codigo & "' "
    
    Select Case tipo
        Case 1
            condicion = condicion & " and ctacte = 1"
        Case 2
            condicion = condicion & "and crcc = 1"
        Case 3
            condicion = condicion & "and banco = 1"
    End Select
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerNombreCuentaMayorempresa = sqlconta.response(0, 3)
    Else
        leerNombreCuentaMayorempresa = ""
    End If
End Function

Public Function leerNombreMayor(ByVal codigo As String) As String
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo = '" & codigo & "' and año='" + Format(fechasistema, "yyyy") + "' "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerNombreMayor = sqlconta.response(0, 3)
    Else
        leerNombreMayor = ""
    End If
End Function

Public Function leerNombrerut(tipo, rut) As String
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = "cuentascorrientes"
    
    condicion = "tipo = '" & tipo & "' and rut='" + rut + "' and año='" & Format(fechasistema, "yyyy") & "'"
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerNombrerut = sqlconta.response(0, 3)
    Else
        leerNombrerut = ""
    End If
End Function


Public Function nombrectacte(rut) As String

    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + CUENTAPROVEEDOR + "' and rut=" + "'" + rut + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    nombrectacte = "*** NO CREADO *** "
    
    If sqlconta.status = 0 Then
    nombrectacte = sqlconta.response(1, 3)
    
    End If
    
End Function


Public Function comparaArchivos(File1 As String, file2 As String) As Boolean
    Dim issame As Boolean
    Dim whole As Double
    Dim part As Long
    Dim buffer1 As String
    Dim buffer2 As String
    Dim start As Long
    Dim X As Long
    Dim nf1 As Integer
    Dim nf2 As Integer

    nf1 = FreeFile
    Open File1 For Binary As #nf1
    nf2 = FreeFile
    Open file2 For Binary As #nf2
    issame = True
    If LOF(nf1) <> LOF(nf2) Then
        issame = False
        comparaArchivos = False
    Else
        whole = LOF(nf1) \ 10000
        part = LOF(nf1) Mod 10000
        buffer1 = String(10000, 0)
        buffer2 = String(10000, 0)
        start = 1
        For X = 1 To whole
            Get #nf1, start, buffer1
            Get #nf2, start, buffer2
            If buffer1 <> buffer2 Then
                issame = False
                Exit For
            End If
            start = start + 10000
        Next X
        buffer1 = String(part, 0)
        buffer2 = String(part, 0)
        Get #nf1, start, buffer1
        Get #nf2, start, buffer2
        If buffer1 <> buffer2 Then
            issame = False
        End If
        If issame = True Then
            comparaArchivos = True
        Else
            comparaArchivos = False
        End If
    End If
    Close #nf1
    Close #nf2
End Function

Sub escribeArchivoRuta(ByVal tipo As String, ByVal cadena As String, ByVal archivo As String)
    Dim NUMFIC As Integer
    NUMFIC = FreeFile
    If tipo = "SISTEMA" Then
        Open archivo For Output As #NUMFIC
        Close #NUMFIC
    End If
    NUMFIC = FreeFile
    Open archivo For Append As #NUMFIC
    Print #NUMFIC, tipo & "=" & cadena
    Close #NUMFIC
End Sub

Public Sub VisualFileCopy(ByVal SourceFileName As String, ByVal TargetFileName As String)
       Dim i As Integer
       Dim SourceFileNo As Integer
       Dim TargetFileNo As Integer
       Dim SourceFileSize As Long
       Dim CopyBuffer As String
    
       On Error GoTo FileCopyErrorRoutine
       SourceFileSize = FileLen(SourceFileName)
       CopyBuffer = Space$(25000)             'AS LARGE AS POSSIBLE UNDER 65,000
    
    '--KILL THE CURRENT TARGET FILE IF IT EXISTS
       If Len(Dir$(TargetFileName)) Then
          Kill TargetFileName
       End If
    
    '--OPEN FILES
       SourceFileNo = FreeFile
       Open SourceFileName For Binary Access Read As SourceFileNo
       TargetFileNo = FreeFile
       Open TargetFileName For Binary Access Write As TargetFileNo
    
    '--COPY SOURCE FILE TO TARGET FILE
       For i = 1 To SourceFileSize \ Len(CopyBuffer)
          Get #SourceFileNo, , CopyBuffer
    '--PROGRESS GUAGE
          Put #TargetFileNo, , CopyBuffer
          DoEvents
       Next i
    
    '--COPY ANY ODD PORTION OF THE SOURCE FILE REMAINING
       CopyBuffer = Space$(SourceFileSize - loc(TargetFileNo))
       If Len(CopyBuffer) Then
          Get #SourceFileNo, , CopyBuffer
          Put #TargetFileNo, , CopyBuffer
       End If
       Close SourceFileNo
       Close TargetFileNo
    
    Exit Sub
    
FileCopyErrorRoutine:
       MsgBox error$
       Exit Sub
    End Sub


Public Sub actualizar()
    Call escribeArchivoRuta("SISTEMA", App.path, "C:\UPDATE.TXT")
    Call escribeArchivoRuta("UPDATE", rutaUpdate & "\" & App.EXEName & ".exe", "C:\UPDATE.TXT")
    
    If ExisteArchivo(App.path & "\Update.exe") = True Then
        Call Shell(App.path & "\Update.exe", vbNormalFocus)
    End If
    
    If comparaArchivos(App.path & "\Update.exe", rutaUpdate & "\Update.exe") = False Then
        Call VisualFileCopy(rutaUpdate & "\Update.exe", App.path & "\Update.exe")
    End If
    
    Call Shell(App.path & "\Update.exe", vbNormalFocus)
End Sub

Public Sub flechas2(ByVal codigo As Integer, ByRef anterior As Object)
    If codigo = 38 Then
        If anterior.Enabled = True Then
            anterior.SetFocus
        End If
    End If
    If codigo = 40 Then
        Sendkeys "{Tab}"
    End If
End Sub


Public Function leercheque(cuenta, numero) As Boolean

    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 1) = cuenta
    campos(1, 1) = numero
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta='" + cuenta + "' and  numero='" + numero + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leercheque = True
    Else
    leercheque = False
    End If
    
End Function

Public Function leerdatoslocal(empresa, dato) As String
    Dim campos(10, 10) As String
    
    campos(0, 0) = dato
    campos(1, 0) = ""
   
    campos(0, 2) = clientesistema + "gestion.g_maestroempresas"
    condicion = "codigo=" + "'" + empresa + "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerdatoslocal = sqlconta.response(0, 3)
    Else
        leerdatoslocal = ""
    End If
    
End Function
Public Function leerNombreBodega(ByVal codigo As String) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "r_maestrobodegas_" & rubro
    condicion = "codigobodega = '" & codigo & "' AND local= '" & empresaactiva & "' AND rubro = '" & rubro & "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        leerNombreBodega = sqlconta.response(0, 3)
    Else
        leerNombreBodega = ""
    End If
End Function

Public Function leerdatos(ByRef coneccion As rdoConnection, tabla, dato, CONSULTA) As String
    Dim campos(10, 10) As String
    campos(0, 0) = dato
    campos(1, 0) = ""
    campos(0, 2) = tabla
    condicion = CONSULTA
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = coneccion
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
    leerdatos = sqlconta.response(0, 3)
    Else
    leerdatos = ""
    
    End If
    
End Function

   Function ExisteArchivo(sNombreArchivo As String) As Boolean
        Dim AttrDev%
        On Error Resume Next
        AttrDev = GetAttr(sNombreArchivo)
        If err.Number Then
            err.Clear
            ExisteArchivo = False
        Else
            ExisteArchivo = True
        End If
    End Function
Public Function saldocuentacorriente(cuenta, rut, año, empresa) As Double
    Dim saldo As Double
    Dim saldo2 As Double
    
    Dim anterior As Double
    Dim debe As Double
    Dim haber As Double
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = ""
    condicion = "tipo=" + "'" + cuenta + "' and rut='" + rut + "' and año='" + año + "'"
    campos(0, 2) = clientesistema + "conta" + empresa + ".saldosctacte"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
    saldo = sumador
    ctaanterior = sumador
    
    saldo2 = LEERSALDOSCTACTEmovi(cuenta, rut, empresa)
    
    debe = 0
    haber = 0
    
    For k = 1 To 12
    debe = debe + ctadebe2(k)
    haber = haber + ctahaber2(k)
    saldo = saldo + ctadebe2(k) - ctahaber2(k)
    
    Next k
    ctadebe = debe
    ctahaber = haber
    ctasaldo = saldo
    
    saldocuentacorriente = saldo
    
End Function

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Public Function totalotros(numero, loc) As Double

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        
        tabla = "SELECT cuenta,glosa,monto,dh "
        tabla = tabla & "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_anexopagos_" + loc + " "
        tabla = tabla & "WHERE numero= '" & numero & "' ORDER BY linea asc "
        sql.sql = tabla
        sql.Execute
        
        total = 0
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            total = 0
            While Not resultados.EOF
                If resultados(3) = "D" Then multi = 1 Else multi = -1
                total = total + (resultados(2) * multi)
                
                resultados.MoveNext
            Wend
        
        End If
    totalotros = total
    
    End Function


Public Function totalpagosotros(tipo, numero) As Double

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        
        tabla = "SELECT cuenta,glosa,monto,dh "
        tabla = tabla & "FROM " + clientesistema + "conta" + empresaactiva + ".facturasdecompras_anexospagos "
        tabla = tabla & "WHERE tipo='" + tipo + "' and numero= '" & numero & "' ORDER BY linea asc "
        sql.sql = tabla
        sql.Execute
        
        total = 0
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            total = 0
            While Not resultados.EOF
                If resultados(3) = "D" Then multi = 1 Else multi = -1
                total = total + (resultados(2) * multi)
                
                resultados.MoveNext
            Wend
        
        End If
    totalpagosotros = total
    
    End Function

Public Function leertipopago(ByVal rut As String) As String
    Dim tp As String
    
    campos(0, 0) = "modopago"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "conta.cuentascorrientes_datos_pago"
    condicion = "rut = '" & rut & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        tp = sqlconta.response(0, 3)
        If tp = "0" Then leertipopago = "CHEQUE"
        If tp = "1" Then leertipopago = "VALE VISTA"
        If tp = "3" Then leertipopago = "TRANSFERENCIA "
        
    Else
        leertipopago = "CHEQUE"
    End If
End Function
Public Function leerplazo(ByVal rut As String) As String
    Dim tp As String
    
    campos(0, 0) = "plazo"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "conta.cuentascorrientes_datos_pago"
    condicion = "rut = '" & rut & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        leerplazo = sqlconta.response(0, 3)
    Else
        leerplazo = "030"
    End If
End Function

Public Sub actualizafechacierre(fecha)
Dim csql As New rdoQuery
Set csql.ActiveConnection = conta
csql.sql = "update maestroempresas set fechacierre='" & Format(fecha, "yyyy-mm-dd") & "' "
csql.sql = csql.sql & "where codigoempresa='" & empresaactiva & "' "
csql.Execute
Call sincronizadatos(csql.sql, conta, "")

csql.Close
Set csql = Nothing

End Sub

Function Verifica_Permiso(programa As String, OPERACION As String) As Boolean
    Dim i As Integer
    Dim columna As Integer
    'agrega modifica elimina
    Dim resultados2 As rdoResultset
    
     
        If estacerrado(Format(fechasistema, "yyyy-mm-dd")) <> False And OPERACION <> "ingresa" And programa <> "Lista Estado de Resultados Comparativos" Then
        Verifica_Permiso = False
         mensaje_nopermiso = "Periodo Cerrado. Imposible Hacer Modificaciones"
        Exit Function
        Else
        mensaje_nopermiso = "Usted no tiene privilegios suficientes para realizar esta operación."
        End If
        
'    USUARIOSISTEMA = "VANTIO"
    
    Dim csql2 As New rdoQuery
        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT todas," + OPERACION + " "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta.segu_permisos "
        csql2.sql = csql2.sql + "where usuario='" + USUARIOSISTEMA + "' and programa='" + programa + "'"
        csql2.Execute
        sqlconta.glosaeliminacion = ""
        sqlconta.solicitoeliminacion = ""
        sqlconta.audit = True
        
        Verifica_Permiso = False
        If csql2.RowsAffected > 0 Then
           Set resultados2 = csql2.OpenResultset
        If resultados2(1) = 1 Or resultados2(0) = 1 Then
        If OPERACION = "elimina" Then
        frmglosaeliminacion.Show vbModal
        sqlconta.glosaeliminacion = glosaeliminacionsistema
        sqlconta.solicitoeliminacion = solicitaeliminacion

        End If
        
        Verifica_Permiso = True
        Else
        Verifica_Permiso = False
        End If
        End If
       
        
        
End Function

Public Function estacerrado(fechaactual) As Boolean
    Dim csql As New rdoQuery
    Dim resultados  As rdoResultset
    Dim año As String
    Dim MES As String
    año = Mid(fechaactual, 1, 4)
    MES = Mid(fechaactual, 6, 2)
    
    
    Set csql.ActiveConnection = conta
    csql.sql = "select estado from " + clientesistema + "conta.fechacierre "
    csql.sql = csql.sql & "where año = '" & año & "' and mes = '" & CDbl(MES) & "' and empresa='" & empresaactiva & "'"
    csql.Execute
    estacerrado = True
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        estacerrado = resultados(0)
    End If
    csql.Close
    Set resultados = Nothing
    Set csql = Nothing
    If USUARIOSISTEMA = "ERESULTADO" Then
        estacerrado = False
    End If
'    If Verifica_FORM29(fechaactual, empresaactiva) = True Then
'        estacerrado = True
'    End If
End Function

Function Verifica_FORM29(fecha, empresa) As Boolean
    Dim i As Integer
    Dim columna As Integer
    'agrega modifica elimina
    Dim resultados2 As rdoResultset
    Dim csql2  As New rdoQuery
    Dim MES As String
    Dim año As String
    MES = Format(fecha, "mm")
    año = Format(fecha, "yyyy")
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT numero  "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & empresa & ".movimientoscontables "
        csql2.sql = csql2.sql + "where tipo='IV' and mes='" & MES & "' and año='" & año & "' "
        csql2.Execute
        Verifica_FORM29 = False
        If csql2.RowsAffected > 0 Then
            Set resultados2 = csql2.OpenResultset
            Verifica_FORM29 = True
            mensaje_nopermiso = "Periodo Cerrado. Imposible Hacer Modificaciones"
        Else
            Verifica_FORM29 = False
        End If
        If USUARIOSISTEMA = "ERESULTADO" Then
            Verifica_FORM29 = False
        End If
End Function


Function leerestadocheque(cuenta, numero, monto, vencimiento) As String

   Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT cobrado,monto,vencimiento,giradoa "
        csql.sql = csql.sql + "FROM chequesdocumento where cuenta='" + cuenta + "' and numero='" + numero + "' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
                    If resultados(0) <> "0" Then
                        leerestadocheque = "1"
                    End If
                    If resultados(0) = "0" Then
                        leerestadocheque = "0"
                    End If
                    If resultados(0) <> "0" And Format(vencimiento, "yyyy-mm-dd") < Format(resultados(2), "yyyy-mm-dd") Then
                        leerestadocheque = "2"
                    End If
                    If resultados(0) <> "0" And monto <> resultados(1) Then
                        leerestadocheque = "3"
                    End If
                    
                    nombregirador = resultados(3)
                    VENCIMIENTOREAL = resultados(2)
                    MONTOREAL = resultados(1)
            Else
        leerestadocheque = "4"
                    nombregirador = "NO CONTABILIDAD"
                    VENCIMIENTOREAL = "0000-00-00"
                    MONTOREAL = "0"
        
        End If
        
End Function
Function conciliacheque(cuenta, numero, fecha) As String

   Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = contadb
        csql.sql = "update chequesdocumento set cobrado='1',fechacobro='" + Format(fecha, "yyyy-mm-dd") + "' "
        csql.sql = csql.sql + "where cuenta='" + cuenta + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, contadb, "")
        
        If csql.RowsAffected > 0 Then
        conciliacheque = "OK"
        Else
        conciliacheque = "NO"
        End If
        
End Function

Public Sub modificaordenpago(numero, loc, pago, fecha, glosadiferencia)

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        
        tabla = "update " + clientesistema + "gestion" + rubro + ".l_ordendecompra_cabeza_" + loc + " set glosadiferencia='" + glosadiferencia + "', autorizacancelacion='" + pago + "',fechaautorizacionpago='" + Format(fecha, "yyyy-mm-dd") + "' "
        tabla = tabla & "WHERE numero= '" & numero & "' "
        sql.sql = tabla
        sql.Execute
        Call sincronizadatos(sql.sql, gestionrubro, "")
        
    
    End Sub

Public Sub modificaordenpagoentre(numero, loc, pago, fecha, tipo, fechadocu)

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        If pago = "0" Then fecha = ""
        tabla = "update l_movimientos_cabeza_" + loc + " set autorizacancelacion='" + pago + "',fechaautorizacionpago='" + Format(fecha, "yyyy-mm-dd") + "' "
        tabla = tabla & "WHERE tipo='" + tipo + "' and numero= '" & numero & "' and fecha='" + Format(fechadocu, "yyyy-mm-dd") + "' "
        sql.sql = tabla
        sql.Execute
        Call sincronizadatos(sql.sql, gestionrubro, "")
        
    
    End Sub


Public Sub leerelaciones(ByRef frm As Form, ByRef grilla As Grid, rut, loc)
  Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim saldo As Double
       
        
        
        
        Set csql.ActiveConnection = contadb
        csql.sql = "select cc.tipo,cm.nombre from " + clientesistema + "conta" + loc + ".cuentascorrientes as cc," + clientesistema + "conta" + loc + ".cuentasdelmayor as cm "
        csql.sql = csql.sql + "where cm.codigo=cc.tipo and rut='" + rut + "' and cm.año='" + Format(fechasistema, "yyyy") + "' and cc.año='" + Format(fechasistema, "yyyy") + "' "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While resultados.EOF = False
        saldo = saldocuentacorriente(resultados(0), rut, Format(fechasistema, "yyyy"), loc)
        'If ctaanterior + ctadebe + ctahaber + ctasaldo <> 0 Then
        If ctaanterior + ctadebe - ctahaber <> 0 Then
        
        grilla.Rows = grilla.Rows + 1
        grilla.Cell(grilla.Rows - 1, 1).text = resultados(0)
        grilla.Cell(grilla.Rows - 1, 2).text = resultados(1)
        grilla.Cell(grilla.Rows - 1, 3).text = Format(ctaanterior, "###,###,###")
        grilla.Cell(grilla.Rows - 1, 4).text = Format(ctadebe, "###,###,###")
        grilla.Cell(grilla.Rows - 1, 5).text = Format(ctahaber, "###,###,###")
        grilla.Cell(grilla.Rows - 1, 6).text = Format(ctasaldo, "###,###,###")
        
        grilla.Cell(grilla.Rows - 1, 7).text = leerdatos(conta, "maestroempresas", "nombre", "codigoempresa='" + loc + "' ")
        
        
        End If
        resultados.MoveNext
        
        Wend
        End If
        
        csql.Close
        
    
End Sub

Public Function leebanco(ByVal codigo As String) As String

        Dim op As Integer
        Dim condicion As String
        
        campos(0, 0) = "nombre"
        campos(1, 0) = ""
        campos(0, 2) = clientesistema + "teso.maestrobancos"

        condicion = "codigobanco = '" & codigo & "' "

        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        
        leebanco = sqlconta.response(0, 3)
        
        Else
        leebanco = ""
        End If
End Function
Public Function leertiene(ByVal codigo As String, ByVal tipo As Integer) As Boolean

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo = '" & codigo & "' AND "
    
    Select Case tipo
        Case 1
            condicion = condicion & "ctacte = 1"
        Case 2
            condicion = condicion & "crcc = 1"
        Case 3
            condicion = condicion & "banco = 1"
    End Select
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leertiene = True
    Else
        leertiene = False
    End If
End Function

Public Sub SendOutlookMail(Subject As String, Recipient As _
String, Message As String)

On Error GoTo errorHandler
Dim oLapp As Object
Dim oItem As Object

Set oLapp = CreateObject("Outlook.application")
Rem - Set oLapp = CreateObject("cdo.mesagges")
Set oItem = oLapp.CreateItem(0)




With oItem
   .Subject = Subject
   .To = Recipient
    .body = Message
   .attachments.Add ("C:\comprobantedepago.html")

   .Save
  
    
End With
'
Set oLapp = Nothing
Set oItem = Nothing
'

Exit Sub

errorHandler:
Set oLapp = Nothing
Set oItem = Nothing
Exit Sub
End Sub


Public Sub leercuentascomprobante(tipo)

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "cuentadebe"
    campos(1, 0) = "cuentahaber"
    campos(2, 0) = "costo"
    campos(3, 0) = "nombredocumento"
    campos(4, 0) = ""
    campos(0, 2) = "g_maestrotipodedocumentos"
    
    condicion = "tipos = '" & tipo & "' "
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = gestion
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        cuentadebe = sqlconta.response(0, 3)
        cuentahaber = sqlconta.response(1, 3)
        tienecosto = sqlconta.response(2, 3)
        NOMBREDOCUMENTO = sqlconta.response(3, 3)
    Else
        cuentadebe = ""
        cuentahaber = ""
        tienecosto = ""
    
    End If
    
       
End Sub
Public Function leersaldomayor(codigo, fecha) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String
Dim NIVEL As String
Dim suma2 As Double


Set csql.ActiveConnection = contadb
fecha1 = Format(fecha, "yyyy") + "-01-01"
fecha2 = Format(fecha, "yyyy-mm-dd")
'        select sm.codigo,sm.debeanterior,sm.haberanterior,sm.debeanterior-sm.haberanterior+
'(select sum(if(mo.dh='D',monto,monto*-1)) from movimientoscontables as mo where fecha between '2008-01-01' and '2008-12-31' and sm.codigo=mo.codigocuenta) as saldo
'from saldosdelmayor as sm where año='2008' and codigo='11200028'
        NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        
        csql.sql = "SELECT sm.debeanterior-sm.haberanterior,"
        If NIVEL = "1" Then
        csql.sql = csql.sql + "(select sum(if(mo.dh='D',monto,monto*-1)) from movimientoscontables as mo where fecha between '" + fecha1 + "' and '" + fecha2 + "' and mid(sm.codigo,1,2)=mid(mo.codigocuenta,1,2)) as saldo "
        End If
        If NIVEL = "2" Then
        csql.sql = csql.sql + "(select sum(if(mo.dh='D',monto,monto*-1)) from movimientoscontables as mo where fecha between '" + fecha1 + "' and '" + fecha2 + "' and mid(sm.codigo,1,4)=mid(mo.codigocuenta,1,4)) as saldo "
        End If
        If NIVEL = "3" Then
        csql.sql = csql.sql + "(select sum(if(mo.dh='D',monto,monto*-1)) from movimientoscontables as mo where fecha between '" + fecha1 + "' and '" + fecha2 + "' and sm.codigo=mo.codigocuenta) as saldo "
        End If
        
        csql.sql = csql.sql + "FROM saldosdelmayor as sm "
        csql.sql = csql.sql + "WHERE año = '" & año & "' "
        csql.sql = csql.sql + "AND codigo = '" & codigo & "' "
        
        csql.sql = csql.sql + "limit 0,1 "
        csql.Execute
        leersaldomayor = 0
        
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
            If IsNull(resultados(0)) = True Then
            leersaldomayor = 0
            Else
            suma2 = 0
            If IsNull(resultados(1)) = False Then
            suma2 = resultados(1)
            End If
            
            leersaldomayor = resultados(0) + suma2
            End If
            resultados.Close
        Set resultados = Nothing
            
    End If
    
    csql.Close
    Set csql = Nothing

End Function
Public Function leersaldomayoranterior(codigo) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String

Set csql.ActiveConnection = contadb
'        select sm.codigo,sm.debeanterior,sm.haberanterior,sm.debeanterior-sm.haberanterior+
'(select sum(if(mo.dh='D',monto,monto*-1)) from movimientoscontables as mo where fecha between '2008-01-01' and '2008-12-31' and sm.codigo=mo.codigocuenta) as saldo
'from saldosdelmayor as sm where año='2008' and codigo='11200028'
        
        csql.sql = "SELECT sm.debeanterior-sm.haberanterior "
        csql.sql = csql.sql + "FROM saldosdelmayor as sm "
        csql.sql = csql.sql + "WHERE año = '" & año & "' "
        csql.sql = csql.sql + "AND codigo = '" & codigo & "' limit 0,1 "
        csql.Execute
        leersaldomayoranterior = 0
        
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
            If IsNull(resultados(0)) = True Then
            leersaldomayoranterior = 0
            Else
            leersaldomayoranterior = resultados(0)
            End If
            resultados.Close
        Set resultados = Nothing
            
    End If
    
    csql.Close
    Set csql = Nothing

End Function

 Public Sub cambiaColor(ByRef frmXP As FrameXp)
        Dim aux
        aux = frmXP.ColorBarraAbajo
        frmXP.ColorBarraAbajo = frmXP.ColorBarraArriba
        frmXP.ColorBarraArriba = aux
    End Sub
Public Function leersaldoctacte(tipo, rut, fecha) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String
Dim NIVEL As String
Dim movi As Double

Set csql.ActiveConnection = contadb
fecha1 = Format(fecha, "YYYY") + "-01-01"
fecha2 = Format(fecha, "YYYY-mm-dd")
'        select sm.codigo,sm.debeanterior,sm.haberanterior,sm.debeanterior-sm.haberanterior+
'(select sum(if(mo.dh='D',monto,monto*-1)) from movimientoscontables as mo where fecha between '2008-01-01' and '2008-12-31' and sm.codigo=mo.codigocuenta) as saldo
'from saldosdelmayor as sm where año='2008' and codigo='11200028'
        csql.sql = "SELECT sm.debeanterior-sm.haberanterior,"
        csql.sql = csql.sql + "(select sum(if(mo.dh='D',monto,monto*-1)) from movimientoscontables as mo where fecha between '" + fecha1 + "' and '" + fecha2 + "' and sm.tipo=mo.codigocuenta and sm.rut=mo.rutctacte) as saldo "
        
        csql.sql = csql.sql + "FROM saldosctacte as sm "
        csql.sql = csql.sql + "WHERE año = '" & año & "' "
   
        csql.sql = csql.sql + "AND rut = '" & rut & "' and tipo='" + tipo + "' "
        
        csql.sql = csql.sql + "limit 0,1 "
        csql.Execute
        leersaldoctacte = 0
        
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
            If IsNull(resultados(1)) = True Then
            movi = 0
            Else
            movi = resultados(1)
            End If
            leersaldoctacte = resultados(0) + movi
      
            resultados.Close
        Set resultados = Nothing
            
    End If
    
    csql.Close
    Set csql = Nothing

End Function

Public Function LEERNOMBREPROVEEDOR(rut) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "rut= '" & rut & "' and tipo='23100026' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LEERNOMBREPROVEEDOR = sqlconta.response(0, 3)
    Else
        LEERNOMBREPROVEEDOR = ""
    End If
 
    

End Function

Public Function LEERNOMBREPROVEEDORGESTION(rut) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema & "gestion" & rubro & ".r_maestroproveedores_" & rubro
    condicion = "rut= '" & rut & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LEERNOMBREPROVEEDORGESTION = sqlconta.response(0, 3)
    Else
        LEERNOMBREPROVEEDORGESTION = ""
    End If
 
    

End Function
Public Function LEERNOMBREFAMILIA(codigo) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "maestro_familias_nuevo"
    condicion = "codigo= '" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LEERNOMBREFAMILIA = sqlconta.response(0, 3)
    Else
        LEERNOMBREFAMILIA = ""
    End If
    

End Function
    
Public Function LeerNombreFamiliaSII(codigo) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "maestro_familias_tributario"
    condicion = "codigo= '" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LeerNombreFamiliaSII = sqlconta.response(0, 3)
    Else
        LeerNombreFamiliaSII = ""
    End If
    

End Function
      
      
Public Function LEERdireccionproveedor(rut) As String
    campos(0, 0) = "direccion"
    campos(1, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "rut= '" & rut & "' and tipo='23100026' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LEERdireccionproveedor = sqlconta.response(0, 3)
    Else
        LEERdireccionproveedor = ""
    End If
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


End Function
Public Function LEERciudadproveedor(rut) As String
    campos(0, 0) = "ciudad"
    campos(1, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "rut= '" & rut & "' and tipo='23100026' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LEERciudadproveedor = sqlconta.response(0, 3)
    Else
        LEERciudadproveedor = ""
    End If
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


End Function

Public Function diferenciahora(hora1, hora2) As Double
Dim hh1 As Double
Dim mm1 As Double
Dim ss1 As Double
Dim hh2 As Double
Dim mm2 As Double
Dim ss2 As Double
hh1 = CDbl(Mid(hora1, 1, 2)) * 60
mm1 = CDbl(Mid(hora1, 4, 2))
hh2 = CDbl(Mid(hora2, 1, 2)) * 60
mm2 = CDbl(Mid(hora2, 4, 2))
diferenciahora = (hh2 + mm2) - (hh1 + mm1)

End Function
Public Function leerdevoluciones(rut, loc) As Boolean

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        Dim fecha30 As String
        
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        fecha30 = DateAdd("d", -30, Format(fechasistema, "yyyy-mm-dd"))
        
        tabla = "SELECT  NUMERO,tipo "
        tabla = tabla & "FROM devoluciones_proveedores "
        tabla = tabla & "WHERE rut= '" & rut & "' and montoco='0' and local='" + loc + "' and fecha<'" + Format(fecha30, "yyyy-mm-dd") + "' "
        sql.sql = tabla
        sql.Execute
        
        leerdevoluciones = False
        If sql.RowsAffected > 0 Then
        
        
            Set resultados = sql.OpenResultset
            While resultados.EOF = False
            If guiarebajada(resultados(1), resultados(0), clientesistema + "gestion" + leerdatoslocal(loc, "rubro") + ".l_ordendecompra_anexopagos_" + loc) = False Then
            leerdevoluciones = True
            End If
            
            resultados.MoveNext
            Wend
            
            
        End If
    
    End Function
Public Function leerpublicidad(rut, loc) As Boolean

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        
        tabla = "SELECT tipo,if(tipo='1',numero,foliosii) "
        tabla = tabla & "FROM facturasdepublicidad "
        tabla = tabla & "WHERE rut= '" & rut & "' and abono='0' "
        sql.sql = tabla
        sql.Execute
       Rem  If rut = "0810941006" Then Stop
        leerpublicidad = False
        If sql.RowsAffected > 0 Then
        
        
            Set resultados = sql.OpenResultset
            While resultados.EOF = False
            If facturarebajada(resultados(0), resultados(1), clientesistema + "gestion" + leerdatoslocal(loc, "rubro") + ".l_ordendecompra_anexopagos_" + loc) = False Then
            leerpublicidad = True
            End If
            
            resultados.MoveNext
            Wend
            
            
        End If
    
    End Function
    
    
    Public Function leernombrelocal(codigo) As String

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        
        tabla = "SELECT nombre "
        tabla = tabla & "FROM " & clientesistema & "gestion" & ".g_maestroempresas "
        tabla = tabla & "WHERE codigo= '" & codigo & "' "
        sql.sql = tabla
        sql.Execute
        
        leernombrelocal = ""
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            leernombrelocal = resultados(0)
            
            
        End If
    
    End Function

Public Function guiarebajada(tipo, numero, base) As Boolean

        Dim suma As Double
        Dim rutpro As String
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        
        Dim tabla As String
Set csql.ActiveConnection = contadb
csql.sql = "select numero "
csql.sql = csql.sql & "from " + base + " "
csql.sql = csql.sql & "where tipodo='" & tipo & "' and numerodo='" + numero + "' "
csql.Execute

If numero = "0000000005" Then
'Stop
End If

guiarebajada = False

If csql.RowsAffected > 0 Then
guiarebajada = True

End If


End Function

Public Function facturarebajada(tipo, numero, base) As Boolean

        Dim suma As Double
        Dim rutpro As String
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        
        Dim tabla As String
Set csql.ActiveConnection = contadb
csql.sql = "select numero "
csql.sql = csql.sql & "from " + base + " "
csql.sql = csql.sql & "where tipodo='" & tipo & "' and numerodo='" + numero + "' "
csql.Execute
facturarebajada = False

If csql.RowsAffected > 0 Then
facturarebajada = True

End If


End Function


Public Function leerbanco(ByVal codigo As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "maestrobancos"
    
    condicion = "codigobanco = '" & codigo & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerbanco = sqlconta.response(0, 3)
    Else
        leerbanco = ""
    End If
End Function

Public Function leerdeposito(ByVal codigo As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "maestrodepositos"
    
    condicion = "codigo = '" & codigo & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerdeposito = sqlconta.response(0, 3)
    Else
        leerdeposito = ""
    End If
End Function
Public Function leertipocredito(ByVal codigo As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "creditos_bancarios." + "maestro_tipo_compromiso"
    
    condicion = "codigo = '" & codigo & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leertipocredito = sqlconta.response(0, 3)
    Else
        leertipocredito = ""
    End If
End Function
Public Function leertipoMONEDA(ByVal codigo As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "creditos_bancarios." + "maestro_tipo_monedas"
    
    condicion = "codigo = '" & codigo & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leertipoMONEDA = sqlconta.response(0, 3)
        
    Else
        leertipoMONEDA = ""
    End If
End Function


Public Function leerempresa(ByVal codigo As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "maestroempresas"
    
    condicion = "codigoempresa = '" & codigo & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerempresa = sqlconta.response(0, 3)
    Else
        leerempresa = ""
    End If
End Function

Public Function leerRutempresa(ByVal codigo As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "rut"
    campos(1, 0) = ""
    campos(0, 2) = "maestroempresas"
    
    condicion = "codigoempresa = '" & codigo & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerRutempresa = sqlconta.response(0, 3)
    Else
        leerRutempresa = ""
    End If
End Function

Public Function depositocontabilizado(numero, fecha, empresa, tipo) As Boolean

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "numero"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    
    condicion = "tipo='" + tipo + "' AND NUMERO='" + numero + "' and fecha='" + fecha + "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        depositocontabilizado = True
    Else
        depositocontabilizado = False
    End If
End Function

Public Sub eliminadepositocontabilizado(numero, fecha, empresa, tipo)

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "numero"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    
    
    condicion = "tipo='" + tipo + "' AND NUMERO='" + numero + "' and fecha='" + fecha + "'"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
End Sub

Public Sub eliminacomprobante(fecha, empresa, tipo)

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "numero"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    
    
    condicion = "tipo='" & tipo & "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "'"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
End Sub

Public Function arriendoatrasado(numero, fecha) As String

 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select * from " & clientesistema & "arriendos" & ".arriendos_mensuales as mp where numero='" + numero + "' and pagado='0' order by fecha asc limit 0,1  "
 
 csql.Execute
arriendoatrasado = "0"

 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    If Format(resultados(1), "yyyy-mm-dd") < fecha Then
    arriendoatrasado = "1"
    
    End If
    
 End If
    
  
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Function

Public Sub leerdatosarrendatario(codigopropiedad)
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb
csql.sql = "select ma.rut,ma.nombre,ma.direccion,ma.comuna,ma.fono,ma.celular from "
csql.sql = csql.sql & clientesistema & "arriendos" & ".maestro_arrendadores as ma ," & clientesistema & "arriendos" & ".maestro_propiedades mp "
csql.sql = csql.sql & "where mp.codigopropiedad='" & codigopropiedad & "' and ma.rut=mp.rutpropietario "
csql.Execute
 
If csql.RowsAffected > 0 Then
   
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
         
        DATOSARRENDATARIO(0) = resultados(0)
        DATOSARRENDATARIO(1) = resultados(1)
        DATOSARRENDATARIO(2) = resultados(2)
        DATOSARRENDATARIO(3) = resultados(3)
        DATOSARRENDATARIO(4) = resultados(4)
        DATOSARRENDATARIO(5) = resultados(5)
        
        resultados.MoveNext
    Wend

End If

End Sub
Public Function leetipoconsumo(ByVal codigo As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "consumos_basicos.maestro_tipo_consumos"
    
    condicion = "codigo = '" & codigo & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leetipoconsumo = sqlconta.response(0, 3)
    Else
        leetipoconsumo = ""
    End If
End Function

Public Function leerProveedor(ByVal rut As String) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "consumos_basicos.proveedores"
    
    condicion = "rut = '" & rut & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerProveedor = sqlconta.response(0, 3)
    Else
        leerProveedor = ""
    End If
End Function


Sub guardarmonto(fecha, monto, empresa, modifica)
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "empresa"
    campos(1, 0) = "fecha"
    campos(2, 0) = "montoplazo"
    campos(3, 0) = ""
    campos(0, 1) = empresa
    campos(1, 1) = fecha
    campos(2, 1) = monto
    campos(0, 2) = "maximopagoproveedores"
    condicion = ""
    If modifica = "1" Then condicion = "empresa= '" & empresa & "' and fecha='" & fecha & "'"
    If modifica = "1" Then op = 3 Else op = 2
    
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    modifica = "0"
End Sub
'Public Function FormularioActivo(NmbFormulario As String) As Boolean
'Dim Formulario As Form
'For Each Formulario In Forms
'If (UCase(Formulario.Name) = UCase(NmbFormulario)) Then
'FormularioActivo = True
'Exit For
'End If
'Next
'End Function
Public Function leerUF(fecha) As Double

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "monto"
    campos(1, 0) = ""
    campos(0, 2) = "maestro_uf"
    
    condicion = "fechauf = '" & Format(fecha, "yyyy-mm-dd") & "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerUF = sqlconta.response(0, 3)
    Else
        leerUF = 1
    End If
End Function

Public Function leerNOMBREcrcc(codigo) As String
    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    
    leerNOMBREcrcc = ""
    
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    csql.sql = "select nombre from " + clientesistema + "conta" + empresaactiva + ".centrosdecosto where codigo='" + codigo + "' and año='" + Format(fechasistema, "yyyy") + "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerNOMBREcrcc = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
 


End Function
Public Function leerNOMBREcrcc2(codigo, empre) As String
    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    
    leerNOMBREcrcc2 = ""
    
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    csql.sql = "select nombre from " + clientesistema + "conta" + empre + ".centrosdecosto where codigo='" + codigo + "' and año='" + Format(fechasistema, "yyyy") + "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerNOMBREcrcc2 = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
 


End Function

Public Function leerNOMBREgastos(codigo) As String
    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    
    leerNOMBREgastos = ""
    
    Set csql = New rdoQuery
    Set csql.ActiveConnection = conta
    csql.sql = "select nombre from presupuesto_centros where codigo='" + codigo + "'"
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    leerNOMBREgastos = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
 


End Function


Public Function estacontabilizado(tipo, MES, año, empresa) As String

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "numero"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    
    condicion = "tipo='" + tipo + "' AND mes='" + MES + "' and año='" + año + "' "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        estacontabilizado = "1"
    Else
        estacontabilizado = "0"
    End If
End Function

Public Function leerNOMBREcomprobantes(codigo) As String
    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    
    
    Set csql = New rdoQuery
    Set csql.ActiveConnection = conta
    csql.sql = "select nombredocumento from maestrotipodedocumentos where tipos='" + codigo + "'"
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerNOMBREcomprobantes = resultados(0)
    Else
    leerNOMBREcomprobantes = ""
    
    End If
    csql.Close
    Set csql = Nothing
    
 
End Function

Public Sub eliminacomprobantesmasivos(tipo, MES, año, empresa)

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    
    condicion = "tipo = '" & tipo & "' and mes='" + MES + "' and año='" + año + "' "
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
End Sub


Public Sub generasaldoscuentascorrientes(empresa, tipo, MES, año)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim mesa As String
Dim añoa As String
mesa = Format(MES - 1, "00")
añoa = año
'insert into tempo_saldosctacte (empresa,tipo,rut,anterior,debe,haber)
'SELECT '08',mo.codigocuenta,mo.rutctacte,
'ifnull(SUM(IF(dh='D' and mo.mes<'09' and mo.año='2009',monto,0))-SUM(IF(dh='H' and mo.mes<'09' and mo.año='2009',monto,0)),0) as saldoanterior,
'SUM(IF(dh='D' and mo.mes='09' and mo.año='2009',monto,0)) as mesd,
'SUM(IF(dh='H' and mo.mes='09' and mo.año='2009',monto,0)) as mesh
'FROM eltit_conta08.movimientoscontables as mo
'Where mo.codigocuenta='23100026'  and (mo.mes>='01' and mo.año>='2009') and (mo.mes<='09' and mo.año<='2009')
'GROUP BY codigocuenta,rutctacte
'order by mo.codigocuenta
Set csql.ActiveConnection = contadb
csql.sql = "delete from " + clientesistema + "conta.tempo_saldosctacte "
Rem csql.sql = csql.sql & "Where tipo='" + tipo + "' and empresa='" + empresa + "' "
csql.Execute
Call sincronizadatos(csql.sql, contadb, "")


Set csql.ActiveConnection = contadb
csql.sql = "insert into " + clientesistema + "conta.tempo_saldosctacte (empresa,tipo,rut,anterior,debe,haber) "
csql.sql = csql.sql & "SELECT '" + empresa + "',mo.codigocuenta,mo.rutctacte,"
csql.sql = csql.sql & "ifnull(SUM(IF(dh='D' and mo.mes<'" + MES + "' and mo.año='" + año + "',monto,0))-SUM(IF(dh='H' and mo.mes<'" + MES + "' and mo.año='" + añoa + "',monto,0)),0) as saldoanterior, "
csql.sql = csql.sql & "SUM(IF(dh='D' and mo.mes='" + MES + "' and mo.año='" + añoa + "',monto,0)) as mesd,"
csql.sql = csql.sql & "SUM(IF(dh='H' and mo.mes='" + MES + "' and mo.año='" + añoa + "',monto,0)) as mesh "
csql.sql = csql.sql & "FROM " + clientesistema + "conta" + empresa + ".movimientoscontables as mo "
csql.sql = csql.sql & "Where mo.codigocuenta='" + tipo + "'  and (mo.mes>='01' and mo.año='" + añoa + "') and (mo.mes<='" + MES + "' and mo.año='" + añoa + "') "
csql.sql = csql.sql + "GROUP BY codigocuenta,rutctacte"
csql.Execute
Call sincronizadatos(csql.sql, contadb, "")


End Sub


Public Function FormularioActivo(NmbFormulario As String) As Boolean
Dim formulario As Form
For Each formulario In Forms
If (UCase(formulario.Name) = UCase(NmbFormulario)) Then
FormularioActivo = True
Exit For
End If
Next
End Function


Public Function leerremanente(empresa, MES, año) As Double

    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    Dim mes1 As String
    Dim AÑO1 As String
    
    leerremanente = 0
    mes1 = Format(MES - 1, "00")
    AÑO1 = año
    If mes1 = "00" Then mes1 = "01": AÑO1 = año - 1
    Set csql = New rdoQuery
    Set csql.ActiveConnection = conta
    csql.sql = "select remanente from formulario29 where empresa='" + empresa + "' and mes='" + mes1 + "' and año='" + AÑO1 + "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerremanente = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
 


End Function

Public Function leerimpuestorenta(empresa, MES, año) As Double

    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    
    leerimpuestorenta = 0
    
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    csql.sql = "select sum(monto) from " + clientesistema + "remu" + empresa + ".calculoliquidaciones where codigo='IRE01' and mes='" + MES + "' and año='" + año + "' group by codigo"
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerimpuestorenta = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
 


End Function

Public Sub grabarremanante(empresa, MES, año, remanente)

    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    Dim mes1 As String
    Dim AÑO1 As String
    
    
    mes1 = Format(MES - 1, "00")
    AÑO1 = año
    If mes1 = "00" Then mes1 = "01": AÑO1 = año - 1
    Set csql = New rdoQuery
    Set csql.ActiveConnection = conta
    csql.sql = "insert into formulario29 (empresa,mes,año,remanente) "
    csql.sql = csql.sql & "values('" & empresa & "','" & MES & "','" & año & "','" & remanente & "') "
    csql.sql = csql.sql & "on duplicate key update remanente='" & remanente * -1 & "' "
    csql.Execute
    Call sincronizadatos(csql.sql, conta, "")
    
    csql.Close
    Set csql = Nothing
    
 


End Sub

Public Function LEETablaempresa(codigo, campo) As String
    campos(0, 0) = campo
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "remu.datos_empresa" 'tabla
    condicion = "empresa='" + codigo + "' "
    op = 5
    LEETablaempresa = "00"
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LEETablaempresa = sqlconta.response(0, 3)
    End If
  
End Function

Public Function leerUFmes(MES, año) As Double

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "uf"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "remu.parametroscalculo"
    
    condicion = "mes='" + MES + "' and año='" + año + "' "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerUFmes = sqlconta.response(0, 3)
    Else
        leerUFmes = 1
    End If
End Function
Public Function leerpartime(rut, MES, año, empresa) As Boolean

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "codigog"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "remu" + empresa + ".mt_semipermanente"
    
    condicion = "rut='" + rut + "' and  mes='" + MES + "' and año='" + año + "' and codigotg='0001' and (codigog='0003' or codigog='0004')"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    leerpartime = False
    
    If sqlconta.status = 0 Then
        leerpartime = True
    End If
End Function

Public Sub modifica_mayor(cuenta, sii, campo)
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = campo
    campos(1, 0) = ""
    campos(0, 1) = sii
    campos(0, 2) = clientesistema + "conta" + empresaactiva + ".cuentasdelmayor"
    condicion = "codigo='" + cuenta + "' "
    condicion = condicion & " and año='" & Format(fechasistema, "yyyy") & "' "
    op = 3
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
End Sub

Public Sub modifica_1846(cuenta, año, monto, tipo)
    Dim condicion As String
    Dim op As Integer
    If monto = "" Then Exit Sub
    campos(0, 0) = "empresa"
    campos(1, 0) = "codigo"
    campos(2, 0) = "año"
    campos(3, 0) = "monto"
    campos(4, 0) = "tipo"
    campos(5, 0) = ""
    
    campos(0, 1) = empresaactiva
    campos(1, 1) = cuenta
    campos(2, 1) = año
    campos(3, 1) = CDbl(monto)
    campos(4, 1) = tipo
    
    campos(0, 2) = clientesistema + "conta.sii_1846_datos"
    condicion = ""
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    condicion = "codigo='" + cuenta + "' and año='" + año + "'  "
    op = 3
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
End Sub

Public Function leerimpuestoFACTURA(cuenta, tipo, numero, doc, emp) As Double

    Dim csql As rdoQuery
    Dim resultados As rdoResultset
    
    leerimpuestoFACTURA = 0
    If cuenta = "1900" Then cuenta = "23200005"
    If Mid(cuenta, 1, 2) = "18" Then cuenta = "23200009"
    If tipo = "NF" And doc = "E" Then tipo = "8"
    If tipo = "NF" And doc = "" Then tipo = "4"
    
    If (tipo = "FV" Or tipo = "FA") And doc = "" Then tipo = "1"
    If (tipo = "FV" Or tipo = "FA") And doc = "E" Then tipo = "6"
    If (tipo = "FAE") Then tipo = "6"
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    
    csql.sql = "select sum(monto) from " + clientesistema + "conta" + emp + ".facturasdeventas_detalle where cuentadelmayor='" + cuenta + "' and tipo='" + tipo + "' and numero='" + numero + "' group by cuentadelmayor "
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerimpuestoFACTURA = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
 


End Function

Sub sincronizadatos(ByVal cadena2 As String, ByRef coneccion As rdoConnection, Servidor)

        Dim cadena As String
        Dim resultados2 As rdoResultset
        Dim csql2 As New rdoQuery
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        Dim columnas As String
        Dim basededatos As String
        
        Dim registros As String
        Dim i As Integer
        Dim j As Integer
        
        For i = 1 To 40
            If Mid(coneccion.Connect, i, 1) = ";" Then
                j = i
                basededatos = Mid(coneccion.Connect, 10, j - 10)
            End If
        Next i
        
            
         
        cadena2 = Replace(cadena2, "'", Chr(126))
        Set csql2.ActiveConnection = coneccion
                        cadena = "INSERT INTO " + cliente_sql + "sincroniza.sincronizador_master ("
                        cadena = cadena + "servidor,consulta,basedatos,fechacreacion,horacreacion) VALUES ("
                        cadena = cadena & "'" & Servidor & "','" + cadena2 + "','" + basededatos + "','" & Format(Date, "yyyy-mm-dd") & "','" & Time & "')"
                        'VALORES ASIGNADOS A CADA CAMPO.
                        csql2.sql = cadena
                        csql2.Execute
   
    End Sub
Public Function generacadena(response, opcion) As String
Dim cadena As String
Dim response1 As String
Dim i As Double


Select Case opcion

                Case 2:    '<<<<<   INSERTA   >>>>>>


                    cadena = "INSERT INTO " & response(0, 2) & " ("
                    response1 = ""
                    'NOMBRE DE response.
                    i = 0
                    While response(i, 0) <> ""
                        cadena = cadena & response(i, 0) + ","
                        response1 = response1 & "[" & response(i, 0) & "]"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") VALUES ("
                    
                    'VALORES ASIGNADOS A CADA CAMPO.
                    i = 0
                    While response(i, 0) <> ""
                        cadena = cadena & "'" & response(i, 1) & "',"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") "
                    cadena = cadena & "ON DUPLICATE KEY UPDATE " & response(0, 0) & " = " & response(0, 0)
                    generacadena = cadena
                    
                Case 3:    '<<<<<   ACTUALIZA   >>>>>>
                    
        
                    cadena = "UPDATE " & response(0, 2) & " SET "
                    i = 0
                    response1 = ""
                    response2 = ""
                    While response(i, 0) <> ""
                        cadena = cadena & response(i, 0) & "= '" & response(i, 1) & "',"
                        response1 = response1 & "[" & response(i, 0) & "]"
                        response2 = response2 & "[" & response(i, 1) & "]"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " WHERE " & condicion
                    generacadena = cadena
                    
                    
                Case 4:    '<<<<<   ELIMINA   >>>>>>
'                    If audit = True Then
'                        Call auditoria(Opcion, condicion)
'                    End If
                    cadena = "DELETE FROM " & response(0, 2) & " WHERE " & condicion
                    generacadena = cadena
                     
        
                Case 5:    '<<<<<   LEE   >>>>>>
                    cadena = "SELECT "
                    i = 0
                    While response(i, 0) <> ""
                        cadena = cadena & response(i, 0) & ","
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " FROM " & response(0, 2) & " WHERE " & condicion
                    generacadena = cadena
 End Select
                    
                    
 End Function


Public Function leerglosa(codigoglosa, codigo) As String
    codigo = Format(codigo, "0000")
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "remu.glosas" 'tabla
    condicion = "codigotg='" + codigoglosa + "' and codigo='" + codigo + "'  "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leerglosa = sqlconta.response(1, 3)
    Else
    leerglosa = ""
    End If
    
End Function
Public Function leercalculo(rut, MES, año, codigo) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT monto "
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where rut='" + rut + "' and mes='" + MES + "' and año='" + año + "' and codigo='" + codigo + "' "
    
    csql.Execute
    leercalculo = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leercalculo = resultados(0)
        resultados.Close
        Set resultados = Nothing
    End If
    
End Function
Public Function totalconvenios(rut) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT sum(hd.monto) "
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones as hd inner join " + clientesistema + "remu.tabladecalculo as tc on tc.codigo=hd.codigo and tc.convenio='1' "
    csql.sql = csql.sql + " WHERE rut= '" & rut & "'"
    csql.sql = csql.sql + " AND mes= '" & Format(fechasistema, "mm") & "'"
    csql.sql = csql.sql + " AND año= '" & Format(fechasistema, "yyyy") & "'"
    csql.sql = csql.sql + " ORDER BY hd.codigo"
    csql.Execute
   
    totalconvenios = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        If IsNull(resultados(0)) = False Then
        totalconvenios = resultados(0) 'cod. tabla calculo
        End If
    End If
    
End Function

Public Function leerdiferencia(rut, MES, año, codigo) As Double
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT monto "
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".liquidacionhd where rut='" & rut & "' and mes='" & MES & "' and año='" & año & "' and codtablacalculo='" & codigo & "' "
    csql.Execute
    
    leerdiferencia = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerdiferencia = resultados(0)
        resultados.Close
        Set resultados = Nothing
    End If
End Function

Public Function electronico(rut) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = conta
    csql.sql = "SELECT contable "
    csql.sql = csql.sql + " FROM " & clientesistema & "conta.proveedores_cuenta where rut='" & rut & "' "
    csql.Execute
    
    electronico = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        electronico = resultados(0)
        resultados.Close
        Set resultados = Nothing
    End If
End Function



Public Function leeranticipos(rut, MES, año, codigo) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT monto "
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".liquidacionhd where rut='" + rut + "' and mes='" + MES + "' and año='" + año + "' and codtablacalculo='" + codigo + "' "
    
    csql.Execute
    leeranticipos = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leeranticipos = resultados(0)
        resultados.Close
        Set resultados = Nothing
    End If
    
End Function
 
Public Function leercreditoplus(rut, MES, año, codigo) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT monto "
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".liquidacionhd where rut='" + rut + "' and mes='" + MES + "' and año='" + año + "' and codtablacalculo='" + codigo + "' "
    
    csql.Execute
    leercreditoplus = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leercreditoplus = resultados(0)
        resultados.Close
        Set resultados = Nothing
    End If
    
End Function
Public Function leeraquinaldos(rut, MES, año, codigo) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT monto "
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".liquidacionhd where rut='" + rut + "' and mes='" + MES + "' and año='" + año + "' and (codtablacalculo='00026' or codtablacalculo='00206') "
    
    csql.Execute
    leeraquinaldos = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leeraquinaldos = resultados(0)
        resultados.Close
        Set resultados = Nothing
    End If
    
End Function

Public Function PermisosCuentasDelMayor(Usuario, cuenta) As Boolean

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = conta
    If cuentaexiste(cuenta) = True Then
        csql.sql = "SELECT ifnull(permiso,0) as permiso "
        csql.sql = csql.sql + " from permisos_cuentas where usuario = '" & Usuario & "' and cuenta = '" & cuenta & "' limit 1"
        
        csql.Execute
        PermisosCuentasDelMayor = False
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            PermisosCuentasDelMayor = resultados(0)
            resultados.Close
            Set resultados = Nothing
        End If
    Else
        PermisosCuentasDelMayor = True
    End If
End Function

Function cuentaexiste(cuenta) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select codigo from " & clientesistema & "conta" & empresaactiva & ".cuentasdelmayor "
    csql.sql = csql.sql & "where codigo='" & cuenta & "' "
    csql.Execute
        cuentaexiste = False
    If csql.RowsAffected > 0 Then
        cuentaexiste = True
    Else
        cuentaexiste = False
    End If
End Function

Public Function leerFOLIOSIIDTE(empre, tipo, numero, fecha, caja, loc) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT numero "
    csql.sql = csql.sql + " FROM " & clientesistema & "fae" & loc & ".sv_dte" + loc + " where tipodocumento='" + tipo + "' and localdocumento='" + loc + "' and numerodocumento='" + numero + "' and fechadocumento='" & Format(fecha, "yyyy-mm-dd") & "' and cajadocumento='" + caja + "' "
    
    csql.Execute
    If numero = "0000001104" Then
'    MsgBox numero & " " & caja
    End If
    
    leerFOLIOSIIDTE = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerFOLIOSIIDTE = Format(resultados(0), "0000000000")
        resultados.Close
        Set resultados = Nothing
    End If
    
End Function
 
  
Public Sub Descargar_Forms(ElForm As String)
    On Error Resume Next
    Dim Form As Form
      
    ' bucle recursivo por los formularios del proyecto
    For Each Form In Forms
          
        ' chequea el nombre del formulario actual
          
        '.. si es distinto al del parámetro, entonces lo descarga _
         y elimina la referencia de la memoria
        If Trim(LCase(Form.Name)) <> Trim(LCase(ElForm)) Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
End Sub




'NUEVAS CONTROL DE ACTIVOS
Sub AyudaActivos_Tipos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("11s", "40s")
    cfijo = "nombre like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Tipos de Activos"

    Call cargaAyudaT(Servidor, clientesistema & "conta", Usuario, password, ".af_maestro_tipo_activos", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub AyudaActivos_Ubicaciones(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("11s", "40s")
    cfijo = "nombre like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Ubicaciones de Activos"

    Call cargaAyudaT(Servidor, clientesistema & "conta", Usuario, password, ".af_maestro_ubicaciones", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub AyudaActivos_EmpresaContable(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("11s", "40s")
    cfijo = "nombre like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Empresa Contable"

    Call cargaAyudaT(Servidor, clientesistema & "conta", Usuario, password, ".maestroempresas", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub AyudaActivos_Empresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("11s", "40s")
    cfijo = "nombre like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Empresas "

    Call cargaAyudaT(Servidor, clientesistema & "gestion", Usuario, password, ".g_maestroempresas", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub AyudaActivos_Usuarios(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("usuario", "nombre")
    largo = Array("11s", "40s")
    cfijo = "nombre like '%%'"
    cabezas = Array("Usuario", "Nombre")
    mensajeAyuda = "Ayuda de Usuarios del Sistema"

    Call cargaAyudaT(Servidor, clientesistema & "auditoria", Usuario, password, ".segu_usuarios", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Public Function LeerNombreActivos_tipo(codigo) As String
    codigo = Format(codigo, "00000")
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.af_maestro_tipo_activos"
    condicion = "codigo='" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerNombreActivos_tipo = sqlconta.response(1, 3)
    Else
    LeerNombreActivos_tipo = ""
    End If
    
End Function

Public Function LeerNombreActivos_ubicacion(codigo) As String
    codigo = Format(codigo, "0000")
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.af_maestro_ubicaciones"
    condicion = "codigo='" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerNombreActivos_ubicacion = sqlconta.response(1, 3)
    Else
    LeerNombreActivos_ubicacion = ""
    End If
    
End Function



Public Function leerrubrocomercio(comercio) As String

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        
        tabla = "SELECT rubro "
        tabla = tabla & "FROM " & clientesistema & "gestion" & ".g_maestroempresas "
        tabla = tabla & "WHERE codigo= '" & comercio & "' "
        sql.sql = tabla
        sql.Execute
        
        leerrubrocomercio = ""
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            leerrubrocomercio = resultados(0)
            
            
        End If
    
    End Function
Public Sub generacheques(empresa, cuenta, fecha)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim mesa As String
Dim añoa As String
Set csql.ActiveConnection = contadb

csql.sql = "insert ignore into " + clientesistema + "conta" + empresa + ".chequesdocumento (cuenta,numero,vencimiento,monto,emision,tipocomprobante,numerocomprobante,giradoa,fechacobro,ubicacion,cobrado,tipodocumento,fechamovimiento) "
csql.sql = csql.sql + "select mo.codigocuenta,mo.numerodocumento,mo.fechavencimiento,mo.monto,mo.fechadocumento,mo.tipo,mo.numero,mo.glosacontable,'0000-00-00','0','0','CH','0000-00-00' "
csql.sql = csql.sql + "from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo Left Join "
csql.sql = csql.sql + clientesistema + "conta" + empresa + ".chequesdocumento as ba on mo.numerodocumento=ba.numero and mo.monto=ba.monto "
csql.sql = csql.sql + " where mo.codigocuenta='" + cuenta + "' and mo.tipodocumento='CH' and mo.dh='H' AND mo.tipo<>'CB' AND mo.fecha > '" + Format(fecha, "yyyy-mm-dd") + "' "
csql.sql = csql.sql + "order by ba.numero  "

csql.Execute
Call sincronizadatos(csql.sql, contadb, "")

'
'csql.sql = "UPDATE " + clientesistema + "conta" + empresa + ".cartolasbancarias AS tb "
'csql.sql = csql.sql + " INNER JOIN cugat_teso.rc_depositos AS de ON de.numero=tb.numero "
'csql.sql = csql.sql + " Set fecha_comparacion = de.fecha_venta "
'csql.sql = csql.sql + " WHERE fecha like '%" + Mid(Format(FECHA, "yyyy-mm-dd"), 1, 7) + "%' "
'csql.Execute

End Sub


Public Sub modificaimpresa(tipo, FOLIO)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = contadb
        csql.sql = "UPDATE " + clientesistema + "fae" + CONFI_EMPRESAFAE + ".sv_dte" + CONFI_EMPRESAFAE
        csql.sql = csql.sql & " set impresa='1' WHERE tipo='" & tipo & "' and numero='" & FOLIO & "' "
        csql.Execute
        Call sincronizadatos(csql.sql, conta, "")
        csql.Close
        Set csql = Nothing
    End Sub
    
    Public Sub GENERARdocumento(tipo, caja, loc, FOLIO, fecha)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = contadb
        csql.sql = "update  " & cliente_sql & "ventas" & loc & ".sv_otros_documento_cabeza_" & loc ' & "_fc"
        csql.sql = csql.sql & " set contabilizado='E' WHERE local='" + loc + "' and tipo='" & tipo & "' and numero='" & Format(FOLIO, "0000000000") & "' "
        csql.sql = csql.sql & "and caja='" + caja + "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "'  "
        csql.Execute
        Call sincronizadatos(csql.sql, conta, "")
        csql.Close
        Set csql = Nothing
    End Sub

Public Function documentocreado(tipo, caja, loc, FOLIO, fecha) As Boolean
        Dim csql As rdoQuery
        Dim resultados As rdoResultset
        
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = contadb
        
        csql.sql = "select numero,impresa ,ifnull(fechaenviocliente,''),glosa_sii,glosa_cliente,correo_envio_cliente from " + cliente_sql + "fae" + loc + ".sv_dte" + loc
        csql.sql = csql.sql & " WHERE tipodocumento='" & tipo & "' and numerodocumento='" & Format(FOLIO, "0000000000") & "' and cajadocumento='" + caja + "' and localdocumento='" + loc + "' AND fechadocumento='" & Format(fecha, "yyyy-mm-dd") & "'"
        csql.Execute
     
        documentocreado = False
        NUMERODOCUMENTO_DTE = "0"
        dte_cli_envia = ""
        dte_respuesta_cliente = ""
        dte_respuesta_sii = ""
        dte_email_envio = ""
        
        
        If csql.RowsAffected > 0 Then
          Set resultados = csql.OpenResultset
            dte_cli_envia = resultados(2)
            dte_respuesta_cliente = resultados("glosa_cliente")
            dte_respuesta_sii = resultados("glosa_sii")
            dte_email_envio = resultados("correo_envio_cliente")
            
            documento_dte_impreso = resultados(1)
            NUMERODOCUMENTO_DTE = resultados(0)
             documentocreado = True
        End If
        
        Set csql = Nothing
    End Function
Function leer_039(empresa, MES, año) As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT ifnull(SUM(IF(fd.tipo='3'or fd.tipo='6',fd.monto*-1,fd.monto)),0) FROM facturasdeventas_detalle as fd INNER JOIN facturasdeventas AS fc "
    Rem csql.sql = "SELECT SUM(IF(fd.tipo='3',fd.monto*-1,monto)) FROM facturasdecompras_detalle AS fd INNER JOIN facturasdecompras AS fc "
    csql.sql = csql.sql + "oN fc.tipo = fd.tipo And fc.rut = fd.rut And fc.numero = fd.numero "
    
    csql.sql = csql.sql + "WHERE fc.fecha LIKE '" + año + "-" + MES + "%'AND (cuentadelmayor='23200005' or cuentadelmayor='23200009') AND (fc.tipo='1' OR fc.tipo='2' OR fc.tipo='3' ) "
    csql.Execute
    leer_039 = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
        leer_039 = resultados(0)
    End If
    
    csql.Close
    Set csql = Nothing
    
End Function
Function leer_556(empresa, MES, año) As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT ifnull(SUM(IF(fd.tipo='3' or fd.tipo='6',fd.monto*-1,monto)),0) FROM facturasdecompras_detalle AS fd INNER JOIN facturasdecompras AS fc "
    csql.sql = csql.sql + "oN fc.tipo = fd.tipo And fc.rut = fd.rut And fc.numero = fd.numero "
    csql.sql = csql.sql + "WHERE fc.mescontable ='" + MES + "' AND añocontable='" + año + "' AND (cuentadelmayor='11400005' or cuentadelmayor='11400012' or cuentadelmayor='11400009')  "

    
'    csql.sql = "SELECT SUM(IF(tipo='3' or tipo='6',monto*-1,monto)) FROM facturasdecompras_detalle WHERE fechacreacion LIKE '2014-05%'AND (cuentadelmayor='11400005' or cuentadelmayor='11400012' or cuentadelmayor='11400009')  GROUP BY cuentadelmayor "
    csql.Execute
    leer_556 = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
        leer_556 = resultados(0)
    End If
    
    csql.Close
    Set csql = Nothing
    
End Function
Public Function comprobantedigitable(tipo, programa) As Boolean
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = conta
        
       csql.sql = "select digitable from "
       csql.sql = csql.sql & "maestrotipodedocumentos "
       csql.sql = csql.sql & "where tipos='" & tipo & "' and digitable='1' "
       csql.Execute
       
       comprobantedigitable = False
        If csql.RowsAffected > 0 Then
            comprobantedigitable = True
        End If
        
        csql.sql = "SELECT * FROM " & clientesistema & "conta.segu_permisos WHERE programa LIKE '%" & programa & "%' "
        csql.sql = csql.sql & "and usuario='" & USUARIOSISTEMA & "' and autoriza='1' "
        csql.Execute
        If csql.RowsAffected > 0 Then
             comprobantedigitable = True
        End If
        
        
        csql.Close
        Set csql = Nothing
End Function


Public Function LeerEmpresaFAE(codcontable) As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = conta
        
       csql.sql = "select empresafae from "
       csql.sql = csql.sql & "maestroempresas "
       csql.sql = csql.sql & "where codigoempresa = '" & codcontable & "' "
       csql.Execute
       
       LeerEmpresaFAE = ""
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            LeerEmpresaFAE = resultados(0)
        End If
        
        
        
        csql.Close
        Set csql = Nothing
End Function



Public Function LeerFolioGuiaSII(tipo, numero, empresa) As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = contadb
        
       csql.sql = "select numero from "
       csql.sql = csql.sql & cliente_sql & "fae" & empresa & ".sv_dte" & empresa
       csql.sql = csql.sql & " where tipo='52' and numerodocumento = '" & numero & "' limit 1"
       csql.Execute
       
       LeerFolioGuiaSII = ""
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            LeerFolioGuiaSII = resultados(0)
        End If
        
        
        
        csql.Close
        Set csql = Nothing
End Function

Public Function numero_interno(tipo, FOLIO, loc) As String
        Dim csql As rdoQuery
        Dim resultados As rdoResultset
        
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = contadb
        
        csql.sql = "select numerodocumento  from " + cliente_sql + "fae" + loc + ".sv_dte" + loc
        csql.sql = csql.sql & " WHERE tipo='" & tipo & "' and numero='" & Format(FOLIO, "0000000000") & "' "
        csql.Execute
     
        numero_interno = ""
        
        If csql.RowsAffected > 0 Then
          Set resultados = csql.OpenResultset
            numero_interno = resultados(0)
            
        End If
        
        Set csql = Nothing
    End Function
Public Function LEERULTIMODTE(tipo, caja, loc) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from " + clientesistema + "ventas" + loc + ".sv_otros_documento_cabeza_" + loc + " where tipo='FV' AND caja='98' GROUP BY tipo  "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMODTE = Format(resultados(0) + 1, "0000000000")
    Else
        LEERULTIMODTE = Format(1, "0000000000")
    
    End If
    
End Function

Public Sub modifica_glosas()
        Dim csql As rdoQuery
        Dim resultados As rdoResultset
        Set csql = New rdoQuery
        Set csql.ActiveConnection = contadb
        
csql.sql = "truncate table eltit_fae00.sv_dte00_general "
'csql.sql = "delete from eltit_fae00.sv_dte00_general "
csql.Execute

csql.sql = "truncate table eltit_conta.paso_folio "
'csql.sql = "delete from eltit_conta.paso_folio "
csql.Execute

csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae00.sv_dte00 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute

csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae25.sv_dte25 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute

csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae41.sv_dte41 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute

csql.sql = "insert IGNORE into eltit_conta.paso_folio "
csql.sql = csql.sql + "select mo.tipo,mo.numero,mo.fecha,mo.linea,concat('DEVOLUCION FOLIO SII ',dte.numero,' del ',dte.fecha) as ngl, mo.glosacontable,mo.rutctacte,mo.monto,dte.numero,ifnull(dte.monto,0) as nu,concat(lpad(mid(dte.rut,1,length(dte.rut)-2),9,'0'),right(dte.rut,1)) as rut2 "
csql.sql = csql.sql + "from eltit_conta08.movimientoscontables as mo "
csql.sql = csql.sql + "left join eltit_fae00.sv_dte00_general as dte on dte.numerodocumento=mo.numerodocumento and dte.tipodocumento='G4' and concat(mid(mo.glosacontable,38,4),'-',mid(mo.glosacontable,35,2),'-',mid(mo.glosacontable,32,2))=dte.fecha /*and dte.rut=mo.rutctacte*/ "
csql.sql = csql.sql + "where mo.tipodocumento='D1' and mo.glosacontable like 'GUIA DEVOL%' and mo.fecha > '2014-01-01%' "
csql.sql = csql.sql + "Having nu <> 0 "
csql.sql = csql.sql + "order by mo.rutctacte; "
csql.Execute


csql.sql = "Update "
csql.sql = csql.sql + "eltit_conta08.movimientoscontables as mo inner join eltit_conta.paso_folio as pf on pf.tipo=mo.tipo and pf.numero=mo.numero and pf.linea=mo.linea and pf.fecha=mo.fecha and pf.monto=mo.monto "
csql.sql = csql.sql + "set mo.glosacontable=pf.ngl; "
csql.Execute

csql.sql = "truncate table eltit_fae00.sv_dte00_general "
'csql.sql = "delete from  eltit_fae00.sv_dte00_general "
csql.Execute

csql.sql = "truncate table eltit_conta.paso_folio"
'csql.sql = "delete from  eltit_conta.paso_folio"
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut,cajadocumento) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut,cajadocumento from eltit_fae00.sv_dte00 where fecha > '2014-01-01' and cajadocumento='98'; "
csql.Execute
csql.sql = "insert into eltit_conta.paso_folio "
csql.sql = csql.sql + "select mo.tipo,mo.numero,mo.fecha,mo.linea,concat('PUBLICIDAD FOLIO SII ',dte.numero,' del ',dte.fecha) as ngl, mo.glosacontable,mo.rutctacte,mo.monto,dte.numero,ifnull(dte.monto,0) as nu,concat(lpad(mid(dte.rut,1,length(dte.rut)-2),9,'0'),right(dte.rut,1)) as rut2 "
csql.sql = csql.sql + "from eltit_conta08.movimientoscontables as mo "
csql.sql = csql.sql + "left join eltit_fae00.sv_dte00_general as dte on dte.numerodocumento=mid(mo.glosacontable,20,10) and dte.tipodocumento='FV'  AND dte.cajadocumento='98' "
csql.sql = csql.sql + "where (mo.tipodocumento='PA' or mo.tipodocumento='DB') and mo.glosacontable like 'FACTURA PUBLICIDAD 0%' and mo.fecha > '2014-01-01%' "
csql.sql = csql.sql + "Having nu <> 0 "
csql.sql = csql.sql + "order by mo.rutctacte; "
csql.Execute
csql.sql = "Update "
csql.sql = csql.sql + "eltit_conta08.movimientoscontables as mo inner join eltit_conta.paso_folio as pf on pf.tipo=mo.tipo and pf.numero=mo.numero and pf.linea=mo.linea and pf.fecha=mo.fecha and pf.monto=mo.monto "
csql.sql = csql.sql + "set mo.glosacontable=pf.ngl; "
csql.Execute
csql.sql = "truncate table eltit_fae00.sv_dte00_general; "

'csql.sql = "delete from eltit_fae00.sv_dte00_general; "

csql.Execute
csql.sql = "truncate table eltit_conta.paso_folio; "
'csql.sql = "delete from eltit_conta.paso_folio; "
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae17.sv_dte17 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae18.sv_dte18 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute
csql.sql = "insert ignore into eltit_conta.paso_folio "
csql.sql = csql.sql + "select mo.tipo,mo.numero,mo.fecha,mo.linea,concat('DEVOLUCION FOLIO SII ',dte.numero,' del ',dte.fecha) as ngl, mo.glosacontable,mo.rutctacte,mo.monto,dte.numero,ifnull(dte.monto,0) as nu,concat(lpad(mid(dte.rut,1,length(dte.rut)-2),9,'0'),right(dte.rut,1)) as rut2 "
csql.sql = csql.sql + "from eltit_conta21.movimientoscontables as mo "
csql.sql = csql.sql + "left join eltit_fae00.sv_dte00_general as dte on dte.numerodocumento=mo.numerodocumento and dte.tipodocumento='G4' and concat(mid(mo.glosacontable,38,4),'-',mid(mo.glosacontable,35,2),'-',mid(mo.glosacontable,32,2))=dte.fecha /*and dte.rut=mo.rutctacte*/ "
csql.sql = csql.sql + "where mo.tipodocumento='D1' and mo.glosacontable like 'GUIA DEVOL%' and mo.fecha > '2014-01-01%' "
csql.sql = csql.sql + "Having nu <> 0 "
csql.sql = csql.sql + "order by mo.rutctacte; "
csql.Execute
csql.sql = "Update "
csql.sql = csql.sql + "eltit_conta21.movimientoscontables as mo inner join eltit_conta.paso_folio as pf on pf.tipo=mo.tipo and pf.numero=mo.numero and pf.linea=mo.linea and pf.fecha=mo.fecha and pf.monto=mo.monto "
csql.sql = csql.sql + "set mo.glosacontable=pf.ngl; "
csql.Execute

csql.sql = "truncate table eltit_fae00.sv_dte00_general; "
'csql.sql = "delete from eltit_fae00.sv_dte00_general; "
csql.Execute
csql.sql = "truncate table eltit_conta.paso_folio; "
'csql.sql = "delete from eltit_conta.paso_folio; "
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut,cajadocumento) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut,cajadocumento from eltit_fae17.sv_dte17 where fecha > '2014-01-01' and cajadocumento='98'; "
csql.Execute
csql.sql = "insert into eltit_conta.paso_folio "
csql.sql = csql.sql + "select mo.tipo,mo.numero,mo.fecha,mo.linea,concat('PUBLICIDAD FOLIO SII ',dte.numero,' del ',dte.fecha) as ngl, mo.glosacontable,mo.rutctacte,mo.monto,dte.numero,ifnull(dte.monto,0) as nu,concat(lpad(mid(dte.rut,1,length(dte.rut)-2),9,'0'),right(dte.rut,1)) as rut2 "
csql.sql = csql.sql + "from eltit_conta21.movimientoscontables as mo "
csql.sql = csql.sql + "left join eltit_fae00.sv_dte00_general as dte on dte.numerodocumento=mid(mo.glosacontable,20,10) and dte.tipodocumento='FV'  AND dte.cajadocumento='98' "
csql.sql = csql.sql + "where (mo.tipodocumento='PA' or mo.tipodocumento='DB') and mo.glosacontable like 'FACTURA PUBLICIDAD 0%' and mo.fecha > '2014-01-01%' "
csql.sql = csql.sql + "Having nu <> 0 "
csql.sql = csql.sql + "order by mo.rutctacte; "
csql.Execute
csql.sql = "Update "
csql.sql = csql.sql + "eltit_conta21.movimientoscontables as mo inner join eltit_conta.paso_folio as pf on pf.tipo=mo.tipo and pf.numero=mo.numero and pf.linea=mo.linea and pf.fecha=mo.fecha and pf.monto=mo.monto "
csql.sql = csql.sql + "set mo.glosacontable=pf.ngl; "
csql.Execute
csql.sql = "truncate table eltit_fae00.sv_dte00_general; "
'csql.sql = "delete from  eltit_fae00.sv_dte00_general  "
csql.Execute
csql.sql = "truncate table eltit_conta.paso_folio  "
'csql.sql = "delete from eltit_fae00.sv_dte00_general  "
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae42.sv_dte42 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae44.sv_dte44 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut from eltit_fae45.sv_dte45 where fecha > '2014-01-01' and tipo='52'; "
csql.Execute
csql.sql = "insert ignore into eltit_conta.paso_folio "
csql.sql = csql.sql + "select mo.tipo,mo.numero,mo.fecha,mo.linea,concat('DEVOLUCION FOLIO SII ',dte.numero,' del ',dte.fecha) as ngl, mo.glosacontable,mo.rutctacte,mo.monto,dte.numero,ifnull(dte.monto,0) as nu,concat(lpad(mid(dte.rut,1,length(dte.rut)-2),9,'0'),right(dte.rut,1)) as rut2 "
csql.sql = csql.sql + "from eltit_conta34.movimientoscontables as mo "
csql.sql = csql.sql + "left join eltit_fae00.sv_dte00_general as dte on dte.numerodocumento=mo.numerodocumento and dte.tipodocumento='G4' and concat(mid(mo.glosacontable,38,4),'-',mid(mo.glosacontable,35,2),'-',mid(mo.glosacontable,32,2))=dte.fecha /*and dte.rut=mo.rutctacte*/ "
csql.sql = csql.sql + "where mo.tipodocumento='D1' and mo.glosacontable like 'GUIA DEVOL%' and mo.fecha > '2014-01-01%' "
csql.sql = csql.sql + "Having nu <> 0 "
csql.sql = csql.sql + "order by mo.rutctacte; "
csql.Execute

csql.sql = "Update "
csql.sql = csql.sql + "eltit_conta34.movimientoscontables as mo inner join eltit_conta.paso_folio as pf on pf.tipo=mo.tipo and pf.numero=mo.numero and pf.linea=mo.linea and pf.fecha=mo.fecha and pf.monto=mo.monto "
csql.sql = csql.sql + "set mo.glosacontable=pf.ngl; "
csql.Execute
csql.sql = "truncate table eltit_fae00.sv_dte00_general; "
'csql.sql = "delete from eltit_fae00.sv_dte00_general; "
csql.Execute
csql.sql = "truncate table eltit_conta.paso_folio; "
'csql.sql = "delete from eltit_conta.paso_folio; "
csql.Execute
csql.sql = "insert into eltit_fae00.sv_dte00_general (tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut,cajadocumento) "
csql.sql = csql.sql + "select tipo,numero,fecha,tipodocumento,numerodocumento,monto,localdocumento,rut,cajadocumento from eltit_fae42.sv_dte42 where fecha > '2014-01-01' and cajadocumento='98'; "
csql.Execute
csql.sql = "insert into eltit_conta.paso_folio "
csql.sql = csql.sql + "select mo.tipo,mo.numero,mo.fecha,mo.linea,concat('PUBLICIDAD FOLIO SII ',dte.numero,' del ',dte.fecha) as ngl, mo.glosacontable,mo.rutctacte,mo.monto,dte.numero,ifnull(dte.monto,0) as nu,concat(lpad(mid(dte.rut,1,length(dte.rut)-2),9,'0'),right(dte.rut,1)) as rut2 "
csql.sql = csql.sql + "from eltit_conta34.movimientoscontables as mo "
csql.sql = csql.sql + "left join eltit_fae00.sv_dte00_general as dte on dte.numerodocumento=mid(mo.glosacontable,20,10) and dte.tipodocumento='FV'  AND dte.cajadocumento='98' "
csql.sql = csql.sql + "where (mo.tipodocumento='PA' or mo.tipodocumento='DB') and mo.glosacontable like 'FACTURA PUBLICIDAD 0%' and mo.fecha > '2014-01-01%' "
csql.sql = csql.sql + "Having nu <> 0 "
csql.sql = csql.sql + "order by mo.rutctacte; "
csql.Execute
csql.sql = "Update "
csql.sql = csql.sql + "eltit_conta34.movimientoscontables as mo inner join eltit_conta.paso_folio as pf on pf.tipo=mo.tipo and pf.numero=mo.numero and pf.linea=mo.linea and pf.fecha=mo.fecha and pf.monto=mo.monto "
csql.sql = csql.sql + "set mo.glosacontable=pf.ngl; "
csql.Execute

Rem csql.sql = "Update eltit_conta08.movimientoscontables"
Rem csql.sql = csql.sql + " SET tipodocumento='FA',numerodocumento=LPAD(MID(glosacontable,26,8),10,'0')"
Rem csql.sql = csql.sql + " WHERE glosacontable LIKE 'SALDO FACTURA PUBLICIDAD%' AND fecha>'2013-01-01' AND codigocuenta='11200028' AND (tipodocumento<>'FA' AND tipodocumento<>'FP') AND dh='H' AND MID(glosacontable,33,5)<>'' AND LPAD(MID(glosacontable,26,8),10,'0')>'0000001000';"
Rem csql.Execute


Rem csql.sql = "Update eltit_conta21.movimientoscontables"
Rem csql.sql = csql.sql + " SET tipodocumento='FA',numerodocumento=LPAD(MID(glosacontable,26,8),10,'0')"
Rem csql.sql = csql.sql + " WHERE glosacontable LIKE 'SALDO FACTURA PUBLICIDAD%' AND fecha>'2013-01-01' AND codigocuenta='11200028' AND (tipodocumento<>'FA' AND tipodocumento<>'FP') AND dh='H' AND MID(glosacontable,33,5)<>'' AND LPAD(MID(glosacontable,26,8),10,'0')>'0000001000';"
Rem csql.Execute

Rem csql.sql = "Update eltit_conta34.movimientoscontables"
Rem csql.sql = csql.sql + " SET tipodocumento='FA',numerodocumento=LPAD(MID(glosacontable,26,8),10,'0')"
Rem csql.sql = csql.sql + " WHERE glosacontable LIKE 'SALDO FACTURA PUBLICIDAD%' AND fecha>'2013-01-01' AND codigocuenta='11200028' AND (tipodocumento<>'FA' AND tipodocumento<>'FP') AND dh='H' AND MID(glosacontable,33,5)<>'' AND LPAD(MID(glosacontable,26,8),10,'0')>'0000001000';"
Rem csql.Execute

csql.sql = "UPDATE eltit_conta08.facturasdepublicidad AS fp"
csql.sql = csql.sql + " SET abono=IFNULL((SELECT SUM(IF(dh='D',monto*-1,monto)) FROM eltit_conta08.movimientoscontables AS mo WHERE mo.numerodocumento=fp.numero AND mo.codigocuenta='11200028' AND DH='H' GROUP BY mo.numerodocumento),0)"
csql.sql = csql.sql + " WHERE tipo='1' AND fecha>'2010-12-31' AND rut<>'0888888888' ;"
csql.Execute

csql.sql = "UPDATE eltit_conta08.facturasdepublicidad AS fp"
csql.sql = csql.sql + " SET abono=IFNULL((SELECT SUM(IF(dh='D',monto*-1,monto)) FROM eltit_conta08.movimientoscontables AS mo WHERE mo.numerodocumento=fp.foliosii AND mo.codigocuenta='11200028' AND DH='H' GROUP BY mo.numerodocumento),0)"
csql.sql = csql.sql + " WHERE tipo='2' AND fecha>'2010-12-31' AND rut<>'0888888888' ;"
csql.Execute

csql.sql = "UPDATE eltit_conta21.facturasdepublicidad AS fp"
csql.sql = csql.sql + " SET abono=IFNULL((SELECT SUM(IF(dh='D',monto*-1,monto)) FROM eltit_conta21.movimientoscontables AS mo WHERE mo.numerodocumento=fp.numero AND mo.codigocuenta='11200028' AND DH='H' GROUP BY mo.numerodocumento),0)"
csql.sql = csql.sql + " WHERE tipo='1' AND fecha>'2010-12-31' AND rut<>'0888888888' ;"
csql.Execute

csql.sql = "UPDATE eltit_conta21.facturasdepublicidad AS fp"
csql.sql = csql.sql + "  SET abono=IFNULL((SELECT SUM(IF(dh='D',monto*-1,monto)) FROM eltit_conta21.movimientoscontables AS mo WHERE mo.numerodocumento=fp.foliosii AND mo.codigocuenta='11200028' AND DH='H' GROUP BY mo.numerodocumento),0)"
csql.sql = csql.sql + " WHERE tipo='2' AND fecha>'2010-12-31' AND rut<>'0888888888' ;"
csql.Execute

csql.sql = "UPDATE eltit_conta34.facturasdepublicidad AS fp"
csql.sql = csql.sql + " SET abono=IFNULL((SELECT SUM(IF(dh='D',monto*-1,monto)) FROM eltit_conta34.movimientoscontables AS mo WHERE mo.numerodocumento=fp.numero AND mo.codigocuenta='11200028' AND DH='H' GROUP BY mo.numerodocumento),0)"
csql.sql = csql.sql + " WHERE tipo='1' AND fecha>'2010-12-31' AND rut<>'0888888888' ;"
csql.Execute

csql.sql = "UPDATE eltit_conta34.facturasdepublicidad AS fp"
 csql.sql = csql.sql + " SET abono=IFNULL((SELECT SUM(IF(dh='D',monto*-1,monto)) FROM eltit_conta34.movimientoscontables AS mo WHERE mo.numerodocumento=fp.foliosii AND mo.codigocuenta='11200028' AND DH='H' GROUP BY mo.numerodocumento),0)"
csql.sql = csql.sql + " WHERE tipo='2' AND fecha>'2010-12-31' AND rut<>'0888888888' ;"
csql.Execute
        
        
        Set csql = Nothing
        MsgBox "proceso concluido enter continuar"
    End Sub


Public Function proveedorelectronico(rutprove) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim fecha1 As String
    Dim fecha2 As String
    Dim NIVEL As String
    Dim suma2 As Double
    Dim LINEAS As Double
    
    
    Set csql.ActiveConnection = contadb
           
            csql.sql = "select rut "
            csql.sql = csql.sql & " from " & cliente_sql & "fae.sv_fae_proveedores  where MID(LPAD(REPLACE(rut,'-',''),10,0),1,9)='" & Mid(rutprove, 1, 9) & "' "
            csql.Execute
            proveedorelectronico = False
            If csql.RowsAffected > 0 Then
                proveedorelectronico = True
            End If
        csql.Close
        Set csql = Nothing
End Function

Public Sub modificasii(loc, tipo, FOLIO, glosa)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = contadb
        csql.sql = "UPDATE " + cliente_sql + "fae" + loc + ".sv_dte" + loc + " "
        csql.sql = csql.sql & "set glosa_sii='" + glosa + "' WHERE tipo='" + tipo + "' and numero='" + FOLIO + "' "
        csql.Execute

        csql.Close
        Set csql = Nothing
    End Sub




Public Function EsFacturadorElectronico(empresa) As Boolean
    Dim campo(10, 3) As String
    Dim op As Double
    Dim condicion As String
      campos(0, 0) = "fecharesolucion"
      campos(1, 0) = ""
      campos(0, 2) = "maestroempresas"
      condicion = "codigoempresa='" & empresa & "' and dte_obligado='S' "
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = conta
        EsFacturadorElectronico = False
      Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
            If IsDate(sqlconta.response(0, 3)) = True Then
                EsFacturadorElectronico = True
            End If
        End If

End Function



Public Function LeerFechaResolucion(empresa) As String
    Dim campo(10, 3) As String
    Dim op As Double
    Dim condicion As String
      campos(0, 0) = "fecharesolucion"
      campos(1, 0) = ""
      campos(0, 2) = "maestroempresas"
      condicion = "codigoempresa='" & empresa & "'  "
      op = 5
      
      sqlconta.response = campos
      Set sqlconta.conexion = conta
        LeerFechaResolucion = ""
      Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
            LeerFechaResolucion = sqlconta.response(0, 3)
        End If


End Function
Sub AyudaCuentaMayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    
    
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    basebus = clientesistema + "conta" + empresaactiva
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", caja, campos, cfijo, largo, 2)
    'If Val(caja.text) = 0 Then DATO1.SetFocus: GoTo no
 
  '  caja.text = pivote.text
 
    caja.Enabled = True
    caja.SetFocus
    
no:
End Sub


Public Function LeerNombreActivo(codigo) As String


    Dim condicion As String
    Dim op As Integer
    
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    
    condicion = "codigo='" & codigo & "' "
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    LeerNombreActivo = ""
    If sqlconta.status = 0 Then
        LeerNombreActivo = sqlconta.response(0, 3)
    End If

End Function


Public Function LeerActivoFijoAño(codigo) As Boolean

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        Dim fecha30 As String
        
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        tabla = "SELECT año FROM "
        tabla = tabla & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
        tabla = tabla & " WHERE codigo='" & codigo & "' and año='" & Format(fechasistema, "yyyy") & "' "
        sql.sql = tabla
        sql.Execute
        
        LeerActivoFijoAño = False
        If sql.RowsAffected > 0 Then LeerActivoFijoAño = True
 
    
    End Function
Public Function escuentacorriente(codigo) As Boolean
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo = '" & codigo & "' and ctacte='1'  "
    
    
    If sqlconta.status = 0 Then
        escuentacorriente = True
    Else
        escuentacorriente = False
        
    End If
End Function

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
        program_name & vbCrLf & _
        err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub






Public Function LeerCuentaAlternativa() As Boolean
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "cuenta"
    campos(1, 0) = "servidor"
    campos(2, 0) = "contraseña"
    campos(3, 0) = ""
    campos(4, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.maestro_correos_cuentas"
    
    condicion = "empresa = '' "
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    op = 5
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        email_cuenta_usuario = sqlconta.response(0, 3)
        email_cuenta_server = sqlconta.response(1, 3)
        email_cuenta_clave = sqlconta.response(2, 3)
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
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
       
End Sub
Public Sub leerdatos_Certificado()
     Dim csql As New rdoQuery
     Dim resultados As rdoResultset
     Set csql.ActiveConnection = temporal
     csql.sql = "select licencia,certificado from adminerp_inicio.licencias "
     csql.Execute
     If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         Usuario = resultados(0)
         password = resultados(1)
         Call Conectarconta(Servidor, "mysql", leer_certificado_digital(Usuario, "leyendo_validez_certificado_firma_sii"), leer_certificado_digital(password, "leyendo_validez_certificado_firma_sii"))
         
         Usuario = leer_certificado_digital(Usuario, "leyendo_validez_certificado_firma_sii")
         password = leer_certificado_digital(password, "leyendo_validez_certificado_firma_sii")
         
         
     Else
        MsgBox ("NO EXISTE CONFIGURACION NI LICENCIA PARA ESTE SOFTWARE")
     End If

   
    
no:
End Sub

Public Function VerificaAplicacion(ByVal archivo As String) As Boolean
    Dim hSnapShot As Long
    Dim IDAplicacion As Long
    Dim uProceso As PROCESSENTRY32
    Dim res As Long
    Dim cuenta As Double
    VerificaAplicacion = False
    hSnapShot = CreateToolhelpSnapshot(2&, 0&)
    If hSnapShot <> 0 Then
        uProceso.dwSize = Len(uProceso)
        res = ProcessFirst(hSnapShot, uProceso)
        cuenta = 0
        Do While res
            If UCase(Left$(uProceso.szExeFile, InStr(uProceso.szExeFile, Chr$(0)) - 1)) = UCase(archivo) Then
                    IDAplicacion = uProceso.th32ProcessID
                    VerificaAplicacion = True
                    Exit Function
            End If
            
            res = ProcessNext(hSnapShot, uProceso)
        Loop
        Call CloseHandle(hSnapShot)
    End If
End Function
Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub
