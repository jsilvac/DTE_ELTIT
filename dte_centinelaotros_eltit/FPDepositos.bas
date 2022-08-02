Attribute VB_Name = "FPDepositos"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type deposito
        loc As String
        NUMERO As String
        fecha As String
        Banco As String
        MONTO As String
        TIPO As String
        rut As String
    End Type

    
'=============================================================================
'LEER PRORROGA
'=============================================================================
    Public Function leerDeposito(ByRef D As deposito, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "numero"
        CAMPOS(1, 0) = "fecha"
        CAMPOS(2, 0) = "banco"
        CAMPOS(3, 0) = "monto"
        CAMPOS(4, 0) = "tipo"
        CAMPOS(5, 0) = "rut"
        CAMPOS(6, 0) = ""
        
        CAMPOS(0, 2) = "sv_depositos"
        
        condicion = "local = '" & empresaActiva & "' AND numero " & operador & " '" & CODIGO & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY numero DESC"
        Else
            condicion = condicion & "ORDER BY numero ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerDeposito = True
            Call asigna(D, sql)
        Else
            leerDeposito = False
        End If
    End Function
'=============================================================================
'LEER PRORROGA
'=============================================================================

'=============================================================================
'GRABAR PRORROGA
'=============================================================================
    Public Sub grabarDeposito(ByRef D As deposito, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(D, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND numero = '" & D.NUMERO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR PRORROGA
'=============================================================================

'=============================================================================
'ELIMINAR PRORROGA
'=============================================================================
    Public Sub eliminarDeposito(ByRef D As deposito)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND numero = '" & D.NUMERO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_depositos"
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR PRORROGA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef D As deposito, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        D.NUMERO = sql.response(0, 3)
        D.fecha = sql.response(1, 3)
        D.Banco = sql.response(2, 3)
        D.MONTO = sql.response(3, 3)
        D.TIPO = sql.response(4, 3)
        D.rut = sql.response(5, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef D As deposito, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "fecha"
        CAMPOS(3, 0) = "banco"
        CAMPOS(4, 0) = "monto"
        CAMPOS(5, 0) = "tipo"
        CAMPOS(6, 0) = "rut"
        CAMPOS(7, 0) = ""
        
        CAMPOS(0, 1) = D.loc
        CAMPOS(1, 1) = D.NUMERO
        CAMPOS(2, 1) = D.fecha
        CAMPOS(3, 1) = D.Banco
        CAMPOS(4, 1) = D.MONTO
        CAMPOS(5, 1) = D.TIPO
        CAMPOS(6, 1) = D.rut
        CAMPOS(7, 1) = ""
        
        CAMPOS(0, 2) = "sv_depositos"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



