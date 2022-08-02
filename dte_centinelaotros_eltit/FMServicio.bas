Attribute VB_Name = "FMServicio"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type tipoServicio
        rut As String
        nombre As String
        direccion As String
        comuna As String
        ciudad As String
        fono As String
        celular As String
        CODIGO As String
       
    End Type
    
'=============================================================================
'LEER CAJERA
'=============================================================================
    Public Function leerTecnico(ByRef c As tipoServicio, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "rut"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = "direccion"
        CAMPOS(3, 0) = "comuna"
        CAMPOS(4, 0) = "ciudad"
        CAMPOS(5, 0) = "fono"
        CAMPOS(6, 0) = "celular"
        CAMPOS(7, 0) = "codigo"
        CAMPOS(8, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroserviciotecnico"
        
        condicion = "rut " & operador & " '" & CODIGO & "' "
        If operador = "<" Then
            condicion = condicion & "ORDER BY rut DESC"
        Else
            condicion = condicion & "ORDER BY rut ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerTecnico = True
            Call asigna(c, sql)
        Else
            leerTecnico = False
        End If
    End Function
    
     Public Function leerNombreTecnico(CODIGO) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroserviciotecnico"
        
        condicion = "rut= '" & CODIGO & "' "
       
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreTecnico = sql.response(0, 3)
     
        End If
    End Function
    
    
'=============================================================================
'LEER CAJERA
'=============================================================================

'=============================================================================
'GRABAR CAJERA
'=============================================================================
    Public Sub grabarTecnico(ByRef c As tipoServicio, ByVal modifica As Boolean)
        
'        Dim resultados As rdoResultset
'        Dim cSql As New rdoQuery
'        Dim K As Integer
'        Dim datosoriginales As String
'
'        Call designa(c)
'        Set cSql.ActiveConnection = ventas
'        If modifica = True Then
'            condicion = "WHERE rut = '" & c.rut & "'"
'            cSql.sql = "UPDATE " & CAMPOS(0, 2) & " SET "
'            cSql.sql = cSql.sql & CAMPOS(0, 0) & " = '" & CAMPOS(0, 1) & "', "
'            cSql.sql = cSql.sql & CAMPOS(1, 0) & " = '" & CAMPOS(1, 1) & "', "
'            cSql.sql = cSql.sql & CAMPOS(2, 0) & " = '" & CAMPOS(2, 1) & "', "
'            cSql.sql = cSql.sql & CAMPOS(3, 0) & " = '" & CAMPOS(3, 1) & "', "
'            cSql.sql = cSql.sql & CAMPOS(4, 0) & " = '" & CAMPOS(4, 1) & "', "
'            cSql.sql = cSql.sql & CAMPOS(5, 0) & " = '" & CAMPOS(5, 1) & "', "
'            cSql.sql = cSql.sql & CAMPOS(6, 0) & " = '" & CAMPOS(6, 1) & "', "
'            cSql.sql = cSql.sql & CAMPOS(7, 0) & " = '" & CAMPOS(7, 1) & "' "
'
'            cSql.sql = cSql.sql & condicion
'        Else
'            cSql.sql = "INSERT INTO " & CAMPOS(0, 2) & " ("
'            cSql.sql = cSql.sql & CAMPOS(0, 0) & ", "
'            cSql.sql = cSql.sql & CAMPOS(1, 0) & ", "
'            cSql.sql = cSql.sql & CAMPOS(2, 0) & ", "
'            cSql.sql = cSql.sql & CAMPOS(3, 0) & ", "
'            cSql.sql = cSql.sql & CAMPOS(4, 0) & ", "
'            cSql.sql = cSql.sql & CAMPOS(5, 0) & ", "
'            cSql.sql = cSql.sql & CAMPOS(6, 0) & ", "
'            cSql.sql = cSql.sql & CAMPOS(7, 0) & ") "
'
'            cSql.sql = cSql.sql + "VALUES ('"
'            cSql.sql = cSql.sql & CAMPOS(0, 1) & "', '"
'            cSql.sql = cSql.sql & CAMPOS(1, 1) & "', '"
'            cSql.sql = cSql.sql & CAMPOS(2, 1) & "', '"
'            cSql.sql = cSql.sql & CAMPOS(3, 1) & "', '"
'            cSql.sql = cSql.sql & CAMPOS(4, 1) & "', '"
'            cSql.sql = cSql.sql & CAMPOS(5, 1) & "', '"
'            cSql.sql = cSql.sql & CAMPOS(6, 1) & "', '"
'            cSql.sql = cSql.sql & CAMPOS(7, 1) & "')"
'        End If
'        cSql.Execute
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(c, sql)
        condicion = ""
        If modifica = True Then
            condicion = "rut = '" & c.rut & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = MServicio.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR CAJERA
'=============================================================================

'=============================================================================
'ELIMINAR CAJERA
'=============================================================================
    Public Sub eliminarTecnico(ByRef c As tipoServicio)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & c.rut & "'"
        op = 4
        CAMPOS(0, 2) = "sv_maestroserviciotecnico"
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = MServicio.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR CAJERA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef c As tipoServicio, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        c.rut = sql.response(0, 3)
        c.nombre = sql.response(1, 3)
        c.direccion = sql.response(2, 3)
        c.comuna = sql.response(3, 3)
        c.ciudad = sql.response(4, 3)
        c.fono = sql.response(5, 3)
        c.celular = sql.response(6, 3)
        c.CODIGO = sql.response(7, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef c As tipoServicio, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "rut"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = "direccion"
        CAMPOS(3, 0) = "comuna"
        CAMPOS(4, 0) = "ciudad"
        CAMPOS(5, 0) = "fono"
        CAMPOS(6, 0) = "celular"
        CAMPOS(7, 0) = "codigo"
        CAMPOS(8, 0) = ""
        
        CAMPOS(0, 1) = c.rut
        CAMPOS(1, 1) = c.nombre
        CAMPOS(2, 1) = c.direccion
        CAMPOS(3, 1) = c.comuna
        CAMPOS(4, 1) = c.ciudad
        CAMPOS(5, 1) = c.fono
        CAMPOS(6, 1) = c.celular
        CAMPOS(7, 1) = c.CODIGO
        CAMPOS(8, 1) = ""
        CAMPOS(0, 2) = "sv_maestroserviciotecnico"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



