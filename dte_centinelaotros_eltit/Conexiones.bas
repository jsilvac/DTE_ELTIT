Attribute VB_Name = "Conexiones"
Option Explicit
    Public imprIMETIPO As String
    Public letras(100, 2) As String
    Public PRECIOLIBRE As Boolean
    Public diadepago As String
    Public montocredito As Double
    Public baseteso As String
    Public lista4 As String
    Public imprimepago As String
    Public tipo As String
    Public gestion As rdoConnection
    Public gestionRubro As rdoConnection
    Public gestionAuditoria As rdoConnection
    Public ventas As rdoConnection
    Public ventas2 As rdoConnection
    Public ventasRubro As rdoConnection
    Public ventasAuditoria As rdoConnection
    Public archivoxml As String
    
    Public tesoreria As rdoConnection
    Public conauditoria As rdoConnection
    Public conexionauditoria As rdoConnection
    Public mensaje_nopermiso As String
    Public mensaje_noelimina As String
    Public config As rdoConnection
    Public temporal As rdoConnection
    Public Servidor As String
    Public Servidor2 As String
    Public basedatos As String
    Public baseVentas As String
    Public baseTrigo As String
    Public baseauditoria As String
    Public usuario As String
    Public password As String
    Public empresaActiva As String
    Public localAuditoria As String
    Public rubro As String
    Public rubroAuditoria As String
    Public bodega As String
    Public idCaja As String
    Public cajera As String
    Public tipopago As String
    Public monto_total As Double
    Public rut_cliente As String
    Public sucursal_cliente As String
    Public tipo_doc As String
    Public numero_doc As String
    Public iva As Double
    Public iha As Double
    Public interesNormal As Double
    Public interesmora As Double
    Public fechasistema As String
    Public impFacturas(1 To 4) As String
    Public impBoletas(1 To 4) As String
    Public impGuias(1 To 4) As String
    Public impNCredito(1 To 4) As String
    Public impOtros(1 To 4) As String
    Public imprimeDirecto As Boolean
    Public segu As Boolean
    Public rutaUpdate As String
    Public usuarioSistema As String
    Public passwordSistema As String
    Public segurity As Boolean
    Public envia As Boolean
    Public estadoAnterior As Boolean
    Public numeroboleta As String
    Public empresanombre As String
    Public empresadireccion As String
    Public empresaciudad As String
    Public empresatelefono As String
    Public sw As Boolean
    Public autorizador As Boolean
    Public tercera_Edad As Boolean
    Public descuentogerencia As Boolean
    Public claveautorizador As String
    Public tarjetaautorizador As String
    Public claveautorizadora As String
    Public CHEQUEAPROBADO As Boolean
    Public segundosespera As String
    Public codigorespuesta As String
    Public codigoautorizacion As String
    Public cajero1 As String
    Public condicion As String
'====================================================================================
'Rutina de conexion al servidor de bases de datos
'====================================================================================
    Sub Conectar(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
        Dim cadena_conexion As String
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set gestion = New rdoConnection
        gestion.Connect = cadena_conexion
        gestion.CursorDriver = rdUseServer
        gestion.EstablishConnection
        bd = baseVentas
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set ventas = New rdoConnection
        ventas.Connect = cadena_conexion
        ventas.CursorDriver = rdUseServer
        ventas.EstablishConnection
        
        
    End Sub
    
   Sub Conectar3(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
        Dim cadena_conexion As String
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set ventas2 = New rdoConnection
        ventas2.Connect = cadena_conexion
        ventas2.CursorDriver = rdUseServer
        ventas2.EstablishConnection
        
        
    End Sub

'====================================================================================
'Rutina de conexion al servidor de bases de datos
'====================================================================================
    Sub ConectarRubro(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
        Dim cadena_conexion As String
        bd = bd & rubro
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set gestionRubro = New rdoConnection
        gestionRubro.Connect = cadena_conexion
        gestionRubro.CursorDriver = rdUseServer
        gestionRubro.EstablishConnection
        
        bd = baseVentas
'        bd = bd & rubro
        bd = bd & empresaActiva
'        bd = bd & empresaActiva
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set ventasRubro = New rdoConnection
        ventasRubro.Connect = cadena_conexion
        ventasRubro.CursorDriver = rdUseServer
        ventasRubro.EstablishConnection
        
        
        
    
'        bd = "auditoria"
'
'        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
'        Set auditoria = New rdoConnection
'        auditoria.Connect = cadena_conexion
'        auditoria.CursorDriver = rdUseServer
'        auditoria.EstablishConnection
'
    
    End Sub
    Sub Conectartemporal(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
    Dim cadena_conexion As String
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & ";PWD=" & passworddb & "; UID=" & usuariodb & ";OPTION=3"
    Set temporal = New rdoConnection
    temporal.Connect = cadena_conexion
    temporal.CursorDriver = rdUseServer
    temporal.EstablishConnection
    
End Sub
    Sub Conectarventas(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
        Dim cadena_conexion As String
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set ventasRubro = New rdoConnection
        ventasRubro.Connect = cadena_conexion
        ventasRubro.CursorDriver = rdUseServer
        ventasRubro.EstablishConnection
        
        bd = clientesistema + "auditoria"
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set conauditoria = New rdoConnection
        conauditoria.Connect = cadena_conexion
        conauditoria.CursorDriver = rdUseServer
        conauditoria.EstablishConnection

    
    End Sub
    
    Sub ConectarAuditoria(ByVal Servidor As String, ByVal rubroAuditoria As String, ByVal usuariodb As String, ByVal passworddb As String, ByVal localactivo As String)
        Dim cadena_conexion As String
        Dim bd As String
        bd = basedatos & rubroAuditoria
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set gestionAuditoria = New rdoConnection
        gestionAuditoria.Connect = cadena_conexion
        gestionAuditoria.CursorDriver = rdUseServer
        gestionAuditoria.EstablishConnection
        
        bd = baseVentas & localactivo
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set ventasAuditoria = New rdoConnection
        ventasAuditoria.Connect = cadena_conexion
        ventasAuditoria.CursorDriver = rdUseServer
        ventasAuditoria.EstablishConnection
'        sqlventas.conauditoria = ventasAuditoria
        
        
    End Sub
    Sub conecntarVentasAuditoria(ByVal Servidor As String, ByVal bd As String, ByVal usuariodb As String, ByVal passworddb As String)
        Dim cadena_conexion As String
       
        
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set ventasAuditoria = New rdoConnection
        ventasAuditoria.Connect = cadena_conexion
        ventasAuditoria.CursorDriver = rdUseServer
        ventasAuditoria.EstablishConnection
'    sqlventas.conauditoria = ventasAuditoria
        
    End Sub
'====================================================================================
'Rutina de conexion al servidor de bases de datos
'====================================================================================

'====================================================================================
'Rutina de conexion al servidor de bases de datos para un cotrol data
'====================================================================================
    Sub ConectarControlData(ByRef data As Variant, ByVal Servidor As String, ByVal bd As String, ByVal usuario As String, ByVal password As String, ByVal tabla As String)
        Dim a As String
        Dim cadena_conexion As String
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & password & "; UID=" & usuario & "; OPTION=3"
        data.ConnectionString = cadena_conexion
        data.RecordSource = tabla
        a = Right(tabla, 100)
        data.Refresh
    End Sub

    Sub ConectarConfiguracion(ByVal Servidor As String, ByVal usuariodb As String, ByVal passworddb As String)
        Dim cadena_conexion As String
        Dim bd As String
        bd = "mysql"
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set config = New rdoConnection
        config.Connect = cadena_conexion
        config.CursorDriver = rdUseServer
        config.EstablishConnection
    End Sub





Public Function leeletra(letra As String) As String
       
        Dim campos(3, 3) As String
       
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    
    campos(0, 0) = "digitos"
    campos(1, 0) = ""
    campos(0, 2) = "letras"
    condicion = "letra = '" & letra & "'"
    op = 5
    sql.response = campos
    Set sql.conexion = gestion
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
            leeletra = sql.response(0, 3)
            
        Else
            leeletra = ""
        
    End If
End Function

Public Function leeralias(codigo As String) As String
       
        Dim campos(3, 3) As String
       
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    
    campos(0, 0) = "codigobarra"
    campos(1, 0) = ""
    campos(0, 2) = "r_maestroproductos_alias_" + rubro
    condicion = "codigoalias = '" & codigo & "'"
    op = 5
    sql.response = campos
    Set sql.conexion = gestionRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
            leeralias = sql.response(0, 3)
            
        Else
            leeralias = codigo
        
    End If
End Function


Public Function leefactor(cuota As String) As String
       
        Dim campos(3, 3) As String
       
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    
    campos(0, 0) = "interesnormal"
    campos(1, 0) = ""
    campos(0, 2) = "g_maestroempresas"
    condicion = "codigo = '" & empresaActiva & "'"
    op = 5
    sql.response = campos
    Set sql.conexion = gestion
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
            leefactor = sql.response(0, 3)
            
        Else
            leefactor = ""
        
    End If
End Function
Public Function leefactorMAXIMO() As String
       
        Dim campos(3, 3) As String
       
    
    Dim op As Integer
    
    campos(0, 0) = "maximocuotas"
    campos(1, 0) = ""
    campos(0, 2) = "g_maestroempresas"
    condicion = "codigo='" + empresaActiva + "' "
    op = 5
    sqlventas.response = campos
    Set sqlventas.conexion = gestion
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
            leefactorMAXIMO = sqlventas.response(0, 3)
            
        Else
            leefactorMAXIMO = ""
        
    End If
End Function

Sub esfecha(ByRef dia As TextBox, mes As TextBox, año As TextBox, tipo As String)
Dim maximo As String

If dia.text <> "" Then
        maximo = "31"
        If mes.text = "01" Then maximo = "31"
        If mes.text = "02" Then maximo = "28"
        
        If mes.text = "02" And año.text = "1992" Then maximo = "29"
        If mes.text = "02" And año.text = "1996" Then maximo = "29"
        If mes.text = "02" And año.text = "2000" Then maximo = "29"
        If mes.text = "02" And año.text = "2004" Then maximo = "29"
        If mes.text = "02" And año.text = "2008" Then maximo = "29"
        If mes.text = "02" And año.text = "2012" Then maximo = "29"
        If mes.text = "02" And año.text = "2016" Then maximo = "29"
        If mes.text = "02" And año.text = "2020" Then maximo = "29"
        If mes.text = "02" And año.text = "2024" Then maximo = "29"
        If mes.text = "02" And año.text = "2028" Then maximo = "29"
        If mes.text = "02" And año.text = "2032" Then maximo = "29"
        If mes.text = "02" And año.text = "2036" Then maximo = "29"
        If mes.text = "02" And año.text = "2040" Then maximo = "29"
        If mes.text = "02" And año.text = "2044" Then maximo = "29"
        If mes.text = "02" And año.text = "2048" Then maximo = "29"
        If mes.text = "02" And año.text = "2052" Then maximo = "29"
        
        If mes.text = "03" Then maximo = "31"
        If mes.text = "04" Then maximo = "30"
        If mes.text = "05" Then maximo = "31"
        If mes.text = "06" Then maximo = "30"
        If mes.text = "07" Then maximo = "31"
        If mes.text = "08" Then maximo = "31"
        If mes.text = "09" Then maximo = "30"
        If mes.text = "10" Then maximo = "31"
        If mes.text = "11" Then maximo = "30"
        If mes.text = "12" Then maximo = "31"
        
        If dia.text < "01" Or dia.text > maximo Then
        dia.text = ""
        dia.SetFocus
        
        End If

End If
If mes.text <> "" Then
        maximo = "12"
        If mes.text < "01" Or mes.text > maximo Then
        
        mes.text = ""
        mes.SetFocus
        
        End If

End If
If año.text <> "" Then
        maximo = "2100"
        If año.text < "1995" Or año.text > maximo Then
        año.text = ""
        año.SetFocus
        
        End If

End If


End Sub

Public Sub grabaprincipal(programa)
    Dim cadena As String
    Dim cSql2 As New rdoQuery
                    Set cSql2.ActiveConnection = conauditoria
                    cadena = "INSERT INTO auditoriaventas ("
                    cadena = cadena + "programa,fecha,hora,usuario,evento,tabla) VALUES ( "
                    cadena = cadena & "'" & programa & "','" & Format(Date, "yyyy-mm-dd") & "','" & Time & "','" & usuarioSistema & "','0','" + NOMBREEMPRESA + "')"
                    cSql2.sql = cadena
                    
                    cSql2.Execute
End Sub
Sub Conectar_Auditoria()
'RUTINA PARA CONECTAR A LA BASE DE DATOS DE AUDITORIA
    Dim bd As String
    bd = clientesistema + "auditoria"
    
    On Error GoTo controlerror
        Call Conectar2(Servidor, bd, usuario, password)
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
