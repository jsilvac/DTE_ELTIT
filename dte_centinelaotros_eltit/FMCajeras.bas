Attribute VB_Name = "FMCajeras"
Option Explicit
    Private campos(30, 3) As String
    Public Type tipoCajera
        rut As String
        nombre As String
        direccion As String
        comuna As String
        ciudad As String
        fono As String
        celular As String
        CODIGO As String
        password As String
    End Type
    Public condiccion As String
    
'=============================================================================
'LEER CAJERA
'=============================================================================
    Public Function leerCajera(ByRef c As tipoCajera, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "rut"
        campos(1, 0) = "nombre"
        campos(2, 0) = "direccion"
        campos(3, 0) = "comuna"
        campos(4, 0) = "ciudad"
        campos(5, 0) = "fono"
        campos(6, 0) = "celular"
        campos(7, 0) = "codigoregistradora"
        campos(8, 0) = "password"
        campos(9, 0) = ""
        
        campos(0, 2) = "sv_maestrocajeras"
        
        condicion = "rut " & operador & " '" & CODIGO & "' "
        If operador = "<" Then
            condicion = condicion & "ORDER BY rut DESC"
        Else
            condicion = condicion & "ORDER BY rut ASC"
        End If
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCajera = True
            Call asigna(c, sql)
        Else
            leerCajera = False
        End If
    End Function
    
     Public Function leerNombreCajera(CODIGO) As String

        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "nombre"
        campos(1, 0) = ""
        
        campos(0, 2) = "sv_maestrocajeras"
        
        condicion = "rut= '" & CODIGO & "' "
       
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreCajera = sql.response(0, 3)
     
        End If
    End Function
    
     Public Function leerNombreCaja(CODIGO) As String

        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "descripcion"
        campos(1, 0) = ""
        
        campos(0, 2) = "sv_maestrodecajas"
        
        condicion = "local='" & empresaActiva & "' and numero= '" & CODIGO & "' "
       
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreCaja = sql.response(0, 3)
        End If
    End Function
'=============================================================================
'LEER CAJERA
'=============================================================================

'=============================================================================
'GRABAR CAJERA
'=============================================================================
    Public Sub grabarCajera(ByRef c As tipoCajera, ByVal modifica As Boolean)

        Dim op As Integer
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        Dim K As Integer
        Dim datosoriginales As String
        Call designa(c, sql)
        Set csql.ActiveConnection = ventas
        If modifica = True Then
            condicion = "WHERE rut = '" & c.rut & "'"
            csql.sql = "UPDATE " & campos(0, 2) & " SET "
            csql.sql = csql.sql & campos(0, 0) & " = '" & campos(0, 1) & "', "
            csql.sql = csql.sql & campos(1, 0) & " = '" & campos(1, 1) & "', "
            csql.sql = csql.sql & campos(2, 0) & " = '" & campos(2, 1) & "', "
            csql.sql = csql.sql & campos(3, 0) & " = '" & campos(3, 1) & "', "
            csql.sql = csql.sql & campos(4, 0) & " = '" & campos(4, 1) & "', "
            csql.sql = csql.sql & campos(5, 0) & " = '" & campos(5, 1) & "', "
            csql.sql = csql.sql & campos(6, 0) & " = '" & campos(6, 1) & "', "
            csql.sql = csql.sql & campos(7, 0) & " = '" & campos(7, 1) & "', "
            csql.sql = csql.sql & campos(8, 0) & " = '" & campos(8, 1) & "' "
            csql.sql = csql.sql & condicion
        Else
            csql.sql = "INSERT INTO " & campos(0, 2) & " ("
            csql.sql = csql.sql & campos(0, 0) & ", "
            csql.sql = csql.sql & campos(1, 0) & ", "
            csql.sql = csql.sql & campos(2, 0) & ", "
            csql.sql = csql.sql & campos(3, 0) & ", "
            csql.sql = csql.sql & campos(4, 0) & ", "
            csql.sql = csql.sql & campos(5, 0) & ", "
            csql.sql = csql.sql & campos(6, 0) & ", "
            csql.sql = csql.sql & campos(7, 0) & ", "
            csql.sql = csql.sql & campos(8, 0) & ") "
            csql.sql = csql.sql + "VALUES ('"
            csql.sql = csql.sql & campos(0, 1) & "', '"
            csql.sql = csql.sql & campos(1, 1) & "', '"
            csql.sql = csql.sql & campos(2, 1) & "', '"
            csql.sql = csql.sql & campos(3, 1) & "', '"
            csql.sql = csql.sql & campos(4, 1) & "', '"
            csql.sql = csql.sql & campos(5, 1) & "', '"
            csql.sql = csql.sql & campos(6, 1) & "', '"
            csql.sql = csql.sql & campos(7, 1) & "', '"
            csql.sql = csql.sql & campos(8, 1) & "')"
        End If
        csql.Execute
        Call consultaReplicas(csql.sql, clientesistema + "ventas")
    
    
    
    
    
    
        
        'MIDIFICA OARA CREAR CAJERAS CION EL SINCRONIZA
        
        'Set sql = New sqlventas.sqlventa
        'Call designa(c, sql)
        'condicion = ""
        'If modifica = True Then
        '    condicion = "rut = '" & c.rut & "'"
        '    op = 3
        'Else
        '    op = 2
        'End If
        'Set sql.conexion = ventas
        
        sql.audit = True: sql.programaactivo = MCajeras.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        'Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR CAJERA
'=============================================================================

'=============================================================================
'ELIMINAR CAJERA
'=============================================================================
    Public Sub eliminarCajera(ByRef c As tipoCajera)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & c.rut & "'"
        op = 4
        campos(0, 2) = "sv_maestrocajeras"
        sql.response = campos
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = MCajeras.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        Call consultaReplicas(generacadena(campos, op), clientesistema + "ventas")

    End Sub
'=============================================================================
'ELIMINAR CAJERA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef c As tipoCajera, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        c.rut = sql.response(0, 3)
        c.nombre = sql.response(1, 3)
        c.direccion = sql.response(2, 3)
        c.comuna = sql.response(3, 3)
        c.ciudad = sql.response(4, 3)
        c.fono = sql.response(5, 3)
        c.celular = sql.response(6, 3)
        c.CODIGO = sql.response(7, 3)
        c.password = sql.response(8, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef c As tipoCajera, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        campos(0, 0) = "rut"
        campos(1, 0) = "nombre"
        campos(2, 0) = "direccion"
        campos(3, 0) = "comuna"
        campos(4, 0) = "ciudad"
        campos(5, 0) = "fono"
        campos(6, 0) = "celular"
        campos(7, 0) = "codigoregistradora"
        campos(8, 0) = "password"
        campos(9, 0) = ""
        
        campos(0, 1) = c.rut
        campos(1, 1) = c.nombre
        campos(2, 1) = c.direccion
        campos(3, 1) = c.comuna
        campos(4, 1) = c.ciudad
        campos(5, 1) = c.fono
        campos(6, 1) = c.celular
        campos(7, 1) = c.CODIGO
        campos(8, 1) = c.password
        campos(9, 1) = ""
        
        campos(0, 2) = "sv_maestrocajeras"
         sql.response = campos
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



