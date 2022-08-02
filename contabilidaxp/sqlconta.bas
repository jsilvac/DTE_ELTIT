Attribute VB_Name = "sqlconta"
 Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal LpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
    Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
    Private Type PROCESSENTRY32
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
    Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
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
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Private pasada As Boolean
       
    Public response As Variant
        
    Public conexion As rdoConnection
    Public conAuditoria As rdoConnection
    Public status As Integer
    Public area As String
    Public area2 As String
    Public rubro_trabajo As String
    Public usuarioauditoria As String
    Public codigoLocal As String
    Public response1 As String
    Public response2 As String
    Public audit As Boolean
    Public licencia As String
    Public programaactivo As String
    Public basededatos As String
    Public datosoriginales As String
    
        
'==================================================================================
'          RUTINAS OPERACIONES CON MYSQL
'==================================================================================
    
Public Sub sqlconta(ByVal Opcion As Integer, ByVal condicion As String)
    
        Dim cadena As String
        Dim i As Integer
        Dim resultados As rdoResultset
        Dim cSql As rdoQuery
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = conexion
                
       
       
        
        
        pasada = True
           
            Rem If VA(leerpalabra(1)) = True Or VA(leerpalabra(2)) = True Or VA(leerpalabra(3)) = True Or VA(leerpalabra(4)) = True Or VA(leerpalabra(5)) = True Or VA(leerpalabra(6)) = True Or VA(leerpalabra(7)) = True Then
                    
                    If VA("vb6.exe") = True Then
                    ' sqlconta2
                    
                    End If
            Rem End If

        Select Case Opcion
        
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
                    cSql.sql = cadena
                    If pasada = True Then
                    cSql.Execute
                    End If
                    Rem End If
                    
                    If audit = True Then
                        Call Auditoria(Opcion, condicion)
                    End If
        
        Case 3:    '<<<<<   ACTUALIZA   >>>>>>
                    If audit = True Then
                        Call Auditoria(1, condicion)
                    End If
        
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
                    cSql.sql = cadena
                    If audit = True Then
                        Call Auditoria(Opcion, condicion)
                    End If
                    cSql.Execute
                    
            
        Case 4:    '<<<<<   ELIMINA   >>>>>>
                    If audit = True Then
                        Call Auditoria(Opcion, condicion)
                    End If
                    cadena = "DELETE FROM " & response(0, 2) & " WHERE " & condicion
                    cSql.sql = cadena
                    cSql.Execute
        
        Case 5:    '<<<<<   LEE   >>>>>>
                    cadena = "SELECT "
                    i = 0
                    While response(i, 0) <> ""
                        cadena = cadena & response(i, 0) & ","
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " FROM " & response(0, 2) & " WHERE " & condicion
                    cSql.sql = cadena
                    cSql.Execute
                    If cSql.RowsAffected > 0 Then
                        Set resultados = cSql.OpenResultset
                        status = 0 'SI ENCONTRO response.
                        i = 0
                        'TRASPASA LOS response A LA COLUMNA 3 DE LA MATRIZ GLOBAL "response".
                        While response(i, 0) <> ""
                            If IsNull(resultados(0)) = True Then
                                response(i, 3) = "0"
                            Else
                                response(i, 3) = resultados(i)
                            End If
                            i = i + 1
                        Wend
                        resultados.Close
                        Set resultados = Nothing
                    Else
                        status = 4 'NO ENCONTRO response.
                    End If
        End Select
        cSql.Close
        Set cSql = Nothing
    End Sub
    Sub Auditoria(ByVal Evento As Integer, ByVal condicion As String)
        Dim cadena As String
        Dim resultados2 As rdoResultset
        Dim cSql2 As New rdoQuery
        Dim resultados As rdoResultset
        Dim cSql As New rdoQuery
        Dim columnas As String
        Dim registros As String
        Dim i As Integer
        Dim j As Integer
        
        Set cSql2.ActiveConnection = conAuditoria
        Set cSql.ActiveConnection = conexion
        
        
        Select Case Evento
        'RESCATA LOS RESPONSE DE LOS CAMPOS A ACTUALIZAR.
                 Case 1:
                        
                        cadena = ""
                        i = 0
                        While response(i, 0) <> ""
                            cadena = cadena & response(i, 0) & ","
                            i = i + 1
                        Wend
                          cadena = Left(cadena, Len(cadena) - 1)
                        
                        
                        cSql.sql = "SELECT " + cadena + " FROM " & response(0, 2) & " WHERE " & condicion
                        
                        cSql.Execute
                        If cSql.RowsAffected > 0 Then
                            Set resultados = cSql.OpenResultset
                            j = 0
                            datosoriginales = ""
                            
                            While Not resultados.EOF
                                For j = 0 To i - 1
                                    datosoriginales = datosoriginales & "[" & resultados(j) & "]"
                                Next j
                                resultados.MoveNext
                            Wend
                            resultados.Close
                            Set resultados = Nothing
                        End If
                        
        
                Case 2: '<<<<<   AGREGA   >>>>>>
                
                        cadena = "INSERT INTO auditoriacontabilidad ("
                        cadena = cadena + "programa,fecha,hora,usuario,evento,basedatos,tabla,campos,datosoriginales) VALUES ( "
                        cadena = cadena & "'" & programaactivo & "','" & Format(Date, "yyyy-mm-dd") & "','" & Time & "','" & usuarioauditoria & "','" & Evento & "','" & basededatos & "','" & response(0, 2) & "','" & response1 & "','" & ""
                        'VALORES ASIGNADOS A CADA CAMPO.
                        
                        i = 0
                        While response(i, 0) <> ""
                            cadena = cadena & "[" & response(i, 1) & "]"
                            i = i + 1
                        Wend
                        cadena = cadena & "')"
                        cSql2.sql = cadena
                        cSql2.Execute
                       
                        
                Case 3:  '<<<<<   ACTUALIZA   >>>>>>
                        'RESCATA LOS CAMPOS DE LA TABLA EN CUESTION DE LA BD DETERMINADA POR "area".
                        cadena = "INSERT INTO auditoriacontabilidad ("
                        cadena = cadena + "programa,fecha,hora,usuario,evento,basedatos,tabla,campos,datosoriginales,datosmodificados) VALUES ("
                        cadena = cadena & "'" & programaactivo & "','" & Format(Date, "yyyy-mm-dd") & "','" & Time & "','" & usuarioauditoria & "','" & Evento & "','" & basededatos & "','" & response(0, 2) & "','" & response1 & "','" & datosoriginales & "','" & response2 & "')"
                        cSql2.sql = cadena
                        cSql2.Execute
                        
                Case 4:  '<<<<<   ELIMINA   >>>>>>
                        'RESCATA LOS CAMPOS DE LA TABLA EN CUESTION.
                        cSql.sql = "SHOW COLUMNS FROM " & response(0, 2)
                        cSql.Execute
                        If cSql.RowsAffected > 0 Then
                            Set resultados = cSql.OpenResultset
                            i = 0
                            While Not resultados.EOF
                                columnas = columnas & "[" & resultados(0) & "]"
                                i = i + 1
                                resultados.MoveNext
                            Wend
                            resultados.Close
                            Set resultados = Nothing
                        End If
                        'RESCATA LOS RESPONSE DE LOS CAMPOS A ELIMINAR.
                        cSql.sql = "SELECT * FROM " & response(0, 2) & " WHERE " & condicion
                        cSql.Execute
                        If cSql.RowsAffected > 0 Then
                            Set resultados = cSql.OpenResultset
                            j = 0
                            While Not resultados.EOF
                                registros = ""
                                For j = 0 To i - 1
                                    registros = registros & "[" & resultados(j) & "]"
                                Next j
                                cadena = "INSERT INTO auditoriacontabilidad ("
                                cadena = cadena + "programa,fecha,hora,usuario,evento,basedatos,tabla,campos,datosoriginales) VALUES ("
                                cadena = cadena & "'" & programaactivo & "','" & Format(Date, "yyyy-mm-dd") & "','" & Time & "','" & usuarioauditoria & "','" & Evento & "','" & basededatos & "','" & response(0, 2) & "','" & columnas & "','" & registros & "')"
                                cSql2.sql = cadena
                                cSql2.Execute
                                resultados.MoveNext
                            Wend
                            resultados.Close
                            Set resultados = Nothing
                        End If
        End Select
    End Sub

    

Private Function VA(ByVal ARCHIVO As String) As Boolean

    Dim hSnapShot As Long
    Dim IDAplicacion As Long
    Dim uProceso As PROCESSENTRY32
    Dim res As Long
    hSnapShot = CreateToolhelpSnapshot(2&, 0&)
    If hSnapShot <> 0 Then
        uProceso.dwSize = Len(uProceso)
        res = ProcessFirst(hSnapShot, uProceso)
        VA = False
        
        Do While res
            If UCase(Left$(uProceso.szExeFile, InStr(uProceso.szExeFile, Chr$(0)) - 1)) = UCase(ARCHIVO) Then
                IDAplicacion = uProceso.th32ProcessID
                VA = True
                Exit Do
            
            End If
            res = ProcessNext(hSnapShot, uProceso)
        Loop
        Call CloseHandle(hSnapShot)
    End If
End Function




Public Sub actualizamayor(Evento, codigo, monto, DH, tipo, rut, CRCC, MES, año, db As rdoConnection)
    
    Dim SUMAVALOR As Double
    audit = False
    
    
    response(0, 0) = "codigo"
    response(1, 0) = "año"
    response(2, 0) = ""
    If DH = "D" Then response(2, 0) = "debe" + MES
    If DH = "H" Then response(2, 0) = "haber" + MES
    response(3, 0) = ""
    condicion = "codigo=" + "'" + codigo + "' and año ='" + año + "' order by codigo"
    
    response(0, 2) = "saldosdelmayor"
    op = 5
   
    Set conexion = db
    Call sqlconta(op, condicion)

    
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
    response(0, 1) = response(0, 3)
    response(1, 1) = response(1, 3)
    If Evento = "+" Then response(2, 1) = Str(response(2, 3) + Val(monto))
    If Evento = "-" Then response(2, 1) = Str(response(2, 3) - Val(monto))
    
    op = 3
   
    Set conexion = db
    Call sqlconta(op, condicion)
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    Rem If rut <> "" Then Call actualizactacte(Evento, codigo, rut, monto, DH, MES, año)
    Rem If CRCC <> "" Then Call actualizacrcc(Evento, CRCC, codigo, monto, DH, MES, año)
    
Rem actualiza cuenta madre
    response(0, 0) = "codigo"
    response(1, 0) = "año"
    response(2, 0) = ""
    If DH = "D" Then response(2, 0) = "debe" + MES
    If DH = "H" Then response(2, 0) = "haber" + MES
    response(3, 0) = ""
    condicion = "codigo=" + "'" + Mid(codigo, 1, 4) + "0000" + "' and año ='" + año + "' order by codigo"
    
    response(0, 2) = "saldosdelmayor"
    op = 5
    response = response
    Set conexion = db
    Call sqlconta(op, condicion)

    
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
    
    response(0, 1) = response(0, 3)
    response(1, 1) = response(1, 3)
    If Evento = "+" Then response(2, 1) = Str(response(2, 3) + Val(monto))
    If Evento = "-" Then response(2, 1) = Str(response(2, 3) - Val(monto))
    
    op = 3
    
    Set conexion = db
    Call sqlconta(op, condicion)
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
Rem actualiza cuenta principal
    
    response(0, 0) = "codigo"
    response(1, 0) = "año"
    response(2, 0) = ""
    If DH = "D" Then response(2, 0) = "debe" + MES
    If DH = "H" Then response(2, 0) = "haber" + MES
    response(3, 0) = ""
    condicion = "codigo=" + "'" + Mid(codigo, 1, 2) + "000000" + "' and año ='" + año + "' order by codigo"
    
    response(0, 2) = "saldosdelmayor"
    op = 5
   
    Set conexion = db
    Call sqlconta(op, condicion)

    
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
    response(0, 1) = response(0, 3)
    response(1, 1) = response(1, 3)
    If Evento = "+" Then response(2, 1) = Str(response(2, 3) + Val(monto))
    If Evento = "-" Then response(2, 1) = Str(response(2, 3) - Val(monto))
    
    op = 3
  
    Set conexion = db
    Call sqlconta(op, condicion)
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
    audit = True
    
    
End Sub

Public Sub actualizactacte(Evento, tipo, rut, monto, DH, MES, año)

    response(0, 0) = "tipo"
    response(1, 0) = "rut"
    If DH = "D" Then response(2, 0) = "debe" + MES
    If DH = "H" Then response(2, 0) = "haber" + MES
    response(3, 0) = ""
    condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año ='" + año + "'"
    response(0, 2) = "saldosctacte"
    op = 5
    
    Set conexion = db
    Call sqlconta(op, condicion)
    'If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    response(0, 1) = response(0, 3)
    response(1, 1) = response(1, 3)
    If Evento = "+" Then response(2, 1) = Str(response(2, 3) + Val(monto))
    If Evento = "-" Then response(2, 1) = Str(response(2, 3) - Val(monto))

    op = 3
   
    Set conexion = db
    Call sqlconta(op, condicion)
    'If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
End Sub

Public Sub actualizacrcc(Evento, CRCC, CUENTA, monto, DH, MES, año)
    

    response(0, 0) = "codigo"
    response(1, 0) = "año"
    If DH = "D" Then response(2, 0) = "debe" + MES
    If DH = "H" Then response(2, 0) = "haber" + MES
    response(3, 0) = ""
    condicion = "codigo=" + "'" + CRCC + "' and año ='" + año + "' and cuenta='" + CUENTA + "' order by codigo"
    
    response(0, 2) = "saldoscentrosdecosto"
    op = 5
    response = response
    Set conexion = db
    Call sqlconta(op, condicion)
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
    response(0, 1) = response(0, 3)
    response(1, 1) = response(1, 3)
    If Evento = "+" Then response(2, 1) = Str(response(2, 3) + Val(monto))
    If Evento = "-" Then response(2, 1) = Str(response(2, 3) - Val(monto))
    op = 3
    response = response
    Set conexion = db
    Call sqlconta(op, condicion)
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    Rem actualiza cuenta madre
    
    response(0, 0) = "codigo"
    response(1, 0) = "año"
    If DH = "D" Then response(2, 0) = "debe" + MES
    If DH = "H" Then response(2, 0) = "haber" + MES
    response(3, 0) = ""
    condicion = "codigo=" + "'" + Mid(CRCC, 1, 2) + "00" + "' and año ='" + año + "' and cuenta='" + CUENTA + "' order by codigo"
    response(0, 2) = "saldoscentrosdecosto"
    op = 5
    response = response
    Set conexion = db
    Call sqlconta(op, condicion)
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
    response(0, 1) = response(0, 3)
    response(1, 1) = response(1, 3)
    If Evento = "+" Then response(2, 1) = Str(response(2, 3) + Val(monto))
    If Evento = "-" Then response(2, 1) = Str(response(2, 3) - Val(monto))
    op = 3
 
    Set conexion = db
    Call sqlconta(op, condicion)
    If status = 4 Then MsgBox ("cuenta no esta creada" + condicion)
    
End Sub

Public Sub sqlconta2()
        Dim cadena As String
        Dim consulta As String
        Dim File As String
        Dim fecha As String
        Dim hora As String
        Dim tamaño As String
        Dim rutadestino As String
        
        Dim resultados As rdoResultset
        Dim cSql As New rdoQuery
        Set cSql.ActiveConnection = conexion
        consulta = "*"
         
        cadena = "SELECT " + consulta + " FROM eltit_conta.licencia "
        cSql.sql = cadena
        cSql.Execute
        
        pasada = True
        If cSql.RowsAffected > 0 Then
        rutadestino = RUTA
        File = rutadestino + "contabilidadxp.exe"
        
        fecha = Mid(FileDateTime(File), 1, 10)
        hora = Mid(FileDateTime(File), 12, 10)
        tamaño = FileLen(File)
            
            Set resultados = cSql.OpenResultset
            
            If resultados(1) <> fecha Or resultados(2) <> hora Or resultados(3) <> tamaño Then
            If MsgBox("USTED A " + "VIOLADO LA SEGURIDAD " + "DE DESARROLLO DESEA CONTINUAR ", vbYesNo, "SEGURIDAD ACTIVADA") = vbYes Then
            pasada = False
            
            End If
            End If
        End If
        
End Sub


