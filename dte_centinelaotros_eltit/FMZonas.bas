Attribute VB_Name = "FMZonas"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type Zona
        CODIGO As String
        nombre As String
    End Type

    
'=============================================================================
'LEER ZONA
'=============================================================================
    Public Function leerZona(ByRef Z As Zona, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "codigozona"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrozonas"
        
        condicion = "codigozona " & operador & " '" & CODIGO & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY codigozona DESC"
        Else
            condicion = condicion & "ORDER BY codigozona ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerZona = True
            Call asigna(Z, sql)
        Else
            leerZona = False
        End If
    End Function
'=============================================================================
'LEER ZONA
'=============================================================================

'=============================================================================
'GRABAR ZONA
'=============================================================================
    Public Sub grabarZona(ByRef Z As Zona, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(Z, sql)
        condicion = ""
        If modifica = True Then
            condicion = "codigozona = '" & Z.CODIGO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR ZONA
'=============================================================================

'=============================================================================
'ELIMINAR ZONA
'=============================================================================
    Public Sub eliminarZona(ByRef Z As Zona)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "codigozona = '" & Z.CODIGO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_maestrozonas"
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR ZONA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef Z As Zona, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        Z.CODIGO = sql.response(0, 3)
        Z.nombre = sql.response(1, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef Z As Zona, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "codigozona"
        CAMPOS(1, 0) = "nombre"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 1) = Z.CODIGO
        CAMPOS(1, 1) = Z.nombre
        CAMPOS(2, 1) = ""
        
        CAMPOS(0, 2) = "sv_maestrozonas"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================


