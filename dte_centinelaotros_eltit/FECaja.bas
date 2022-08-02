Attribute VB_Name = "FECaja"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type egreso
        FOLIO As String
        loc As String
        fecha As String
        NUMERO As String
        TIPO As String
        MONTO As String
        GLOSA As String
        Recibido As String
    End Type

    
'=============================================================================
'LEER EGRESO
'=============================================================================
    Public Function leerEgreso(ByRef e As egreso, ByVal CODIGO As String, ByVal operador As String, ByVal fecha As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "fecha"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "tipo"
        CAMPOS(3, 0) = "monto"
        CAMPOS(4, 0) = "glosa"
        CAMPOS(5, 0) = "recibido"
        CAMPOS(6, 0) = "id"
        CAMPOS(7, 0) = ""
        
        CAMPOS(0, 2) = "sv_egresoscaja_" + rubro
        
        condicion = "local = '" & empresaActiva & "' AND id " & operador & " '" & CODIGO & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY id DESC"
        Else
            condicion = condicion & "ORDER BY id ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerEgreso = True
            Call asigna(e, sql)
        Else
            leerEgreso = False
        End If
    End Function
'=============================================================================
'LEER EGRESO
'=============================================================================

'=============================================================================
'GRABAR EGRESO
'=============================================================================
    Public Sub grabarEgreso(ByRef e As egreso, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(e, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND id = '" & e.FOLIO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR EGRESO
'=============================================================================

'=============================================================================
'ELIMINAR EGRESO
'=============================================================================
    Public Sub eliminarEgreso(ByRef e As egreso)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND id = '" & e.FOLIO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_egresoscaja_" + rubro
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR EGRESO
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef e As egreso, ByRef sql As sqlventas.sqlventa)
        e.fecha = sql.response(0, 3)
        e.NUMERO = sql.response(1, 3)
        e.TIPO = sql.response(2, 3)
        e.MONTO = sql.response(3, 3)
        e.GLOSA = sql.response(4, 3)
        e.Recibido = sql.response(5, 3)
        e.FOLIO = sql.response(6, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef e As egreso, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "fecha"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "tipo"
        CAMPOS(4, 0) = "monto"
        CAMPOS(5, 0) = "glosa"
        CAMPOS(6, 0) = "recibido"
        CAMPOS(7, 0) = "id"
        CAMPOS(8, 0) = ""
        
        CAMPOS(0, 1) = e.loc
        CAMPOS(1, 1) = e.fecha
        CAMPOS(2, 1) = e.NUMERO
        CAMPOS(3, 1) = e.TIPO
        CAMPOS(4, 1) = e.MONTO
        CAMPOS(5, 1) = e.GLOSA
        CAMPOS(6, 1) = e.Recibido
        CAMPOS(7, 1) = e.FOLIO
        CAMPOS(8, 1) = ""
        
        CAMPOS(0, 2) = "sv_egresoscaja_" + rubro
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================


