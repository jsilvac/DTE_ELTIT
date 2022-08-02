Attribute VB_Name = "FMVivienda"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type Vivienda
        CODIGO As String
        nombre As String
    End Type

    
'=============================================================================
'LEER BANCO
'=============================================================================
    Public Function leerVivienda(ByRef v As Vivienda, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "codigo"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrotipovivienda"
        
        condicion = "codigo " & operador & " '" & CODIGO & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY codigo DESC"
        Else
            condicion = condicion & "ORDER BY codigo ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerVivienda = True
            Call asigna(v, sql)
        Else
            leerVivienda = False
        End If
    End Function
'=============================================================================
'LEER BANCO
'=============================================================================

'=============================================================================
'GRABAR BANCO
'=============================================================================
    Public Sub grabarVivienda(ByRef v As Vivienda, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(v, sql)
        condicion = ""
        If modifica = True Then
            condicion = "codigo = '" & v.CODIGO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventas
        sql.audit = True:   sql.programaactivo = MVivienda.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR BANCO
'=============================================================================

'=============================================================================
'ELIMINAR BANCO
'=============================================================================
    Public Sub eliminarVivienda(ByRef v As Vivienda)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "codigo = '" & v.CODIGO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_maestrotipovivienda "
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True:   sql.programaactivo = MVivienda.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR BANCO
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef v As Vivienda, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        v.CODIGO = sql.response(0, 3)
        v.nombre = sql.response(1, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef v As Vivienda, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "codigo"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 1) = v.CODIGO
        CAMPOS(1, 1) = v.nombre
        CAMPOS(2, 1) = ""
        
        CAMPOS(0, 2) = "sv_maestrotipovivienda"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================


