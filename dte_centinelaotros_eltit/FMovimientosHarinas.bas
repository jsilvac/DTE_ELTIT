Attribute VB_Name = "FMovimientosHarinas"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type Movimientos
        codLoc As String
        tipoproducto As String
        TIPO As String
        fecha As String
        MONTO As String
    End Type

    
'=============================================================================
'LEER MOVIMIENTO
'=============================================================================
    Public Function leerMovimiento(ByRef m As Movimientos, ByVal codLoc As String, ByVal TIPO As String, ByVal fecha As String, ByVal operador As String, ByVal tipoPro As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "fecha"
        CAMPOS(3, 0) = "monto"
        CAMPOS(4, 0) = ""
        
        CAMPOS(0, 2) = "produccion"
        
        condicion = "local = '" & codLoc & "' AND tipoproducto = '" & tipoPro & "' AND tipo = '" & TIPO & "' AND fecha " & operador & " '" & fecha & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY fecha DESC"
        Else
            condicion = condicion & "ORDER BY fecha ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerMovimiento = True
            Call asigna(m, sql)
        Else
            leerMovimiento = False
        End If
    End Function
'=============================================================================
'LEER MOVIMINETO
'=============================================================================

'=============================================================================
'GRABAR MOVIMIENTO
'=============================================================================
    Public Sub grabarMovimiento(ByRef m As Movimientos, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(m, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & m.codLoc & "' AND tipoproducto = '" & m.tipoproducto & "' AND tipo = '" & m.TIPO & "' AND fecha = '" & m.fecha & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR MOVIMIENTO
'=============================================================================

'=============================================================================
'ELIMINAR MOVIMIENTO
'=============================================================================
    Public Sub eliminarMovimiento(ByRef m As Movimientos)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & m.codLoc & "' AND tipoproducto = '" & m.tipoproducto & "' AND tipo = '" & m.TIPO & "' AND fecha = '" & m.fecha & "'"
        op = 4
        CAMPOS(0, 2) = "produccion"
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR MOVIMIENTO
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef m As Movimientos, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        m.codLoc = sql.response(0, 3)
        m.TIPO = sql.response(1, 3)
        m.fecha = sql.response(2, 3)
        m.MONTO = sql.response(3, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef m As Movimientos, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "fecha"
        CAMPOS(3, 0) = "monto"
        CAMPOS(4, 0) = "tipoproducto"
        CAMPOS(5, 0) = ""
        
        CAMPOS(0, 1) = m.codLoc
        CAMPOS(1, 1) = m.TIPO
        CAMPOS(2, 1) = m.fecha
        CAMPOS(3, 1) = m.MONTO
        CAMPOS(4, 1) = m.tipoproducto
        CAMPOS(5, 1) = ""
        
        CAMPOS(0, 2) = "produccion"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



