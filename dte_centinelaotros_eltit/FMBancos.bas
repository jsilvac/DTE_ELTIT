Attribute VB_Name = "FMBancos"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type Banco
        CODIGO As String
        nombre As String
    End Type

    
'=============================================================================
'LEER BANCO
'=============================================================================
    Public Function leerBanco(ByRef b As Banco, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "codigobanco"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrobancos"
        
        condicion = "codigobanco " & operador & " '" & CODIGO & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY codigobanco DESC"
        Else
            condicion = condicion & "ORDER BY codigobanco ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerBanco = True
            Call asigna(b, sql)
        Else
            leerBanco = False
        End If
    End Function
'=============================================================================
'LEER BANCO
'=============================================================================

'=============================================================================
'GRABAR BANCO
'=============================================================================
    Public Sub grabarBanco(ByRef b As Banco, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(b, sql)
        condicion = ""
        If modifica = True Then
            condicion = "codigobanco = '" & b.CODIGO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR BANCO
'=============================================================================

'=============================================================================
'ELIMINAR BANCO
'=============================================================================
    Public Sub eliminarBanco(ByRef b As Banco)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "codigobanco = '" & b.CODIGO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_maestrobancos"
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR BANCO
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef b As Banco, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        b.CODIGO = sql.response(0, 3)
        b.nombre = sql.response(1, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef b As Banco, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "codigobanco"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 1) = b.CODIGO
        CAMPOS(1, 1) = b.nombre
        CAMPOS(2, 1) = ""
        
        CAMPOS(0, 2) = "sv_maestrobancos"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================


