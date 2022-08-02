Attribute VB_Name = "FCProtestados"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type protesto
        cheque As String
        rut As String
        sucursal As String
        fechaprotesto As String
        MONTO As String
        fechacheque As String
        motivo As String
        CANCELADO As String
        GLOSA As String
    End Type

    
'=============================================================================
'LEER PROTESTO
'=============================================================================
    Public Function leerProtesto(ByRef p As protesto, ByVal codigo1 As String, ByVal codigo2 As String, ByVal codigo3 As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "cheque"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "sucursal"
        CAMPOS(3, 0) = "fechaprotesto"
        CAMPOS(4, 0) = "monto"
        CAMPOS(5, 0) = "fechacheque"
        CAMPOS(6, 0) = "motivo"
        CAMPOS(7, 0) = "cancelado"
        CAMPOS(8, 0) = "glosa"
        CAMPOS(9, 0) = ""
        
        CAMPOS(0, 2) = "sv_protesto_" & empresaActiva
        
        If operador = "=" Then
            condicion = "local = '" & empresaActiva & "' AND cheque = '" & codigo1 & "' AND rut = '" & codigo2 & "' AND sucursal = '" & codigo3 & "'"
        End If
        If operador = "<" Then
            condicion = "local = '" & empresaActiva & "' AND cheque < '" & codigo1 & "' ORDER BY cheque DESC"
        End If
        If operador = ">" Then
            condicion = "local = '" & empresaActiva & "' AND cheque > '" & codigo1 & "' ORDER BY cheque ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerProtesto = True
            Call asigna(p, sql)
        Else
            leerProtesto = False
        End If
    End Function
'=============================================================================
'LEER PROTESTO
'=============================================================================

'=============================================================================
'GRABAR PROTESTO
'=============================================================================
    Public Sub grabarProtesto(ByRef p As protesto, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(p, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND cheque = '" & p.cheque & "' AND rut = '" & p.rut & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR PROTESTO
'=============================================================================

'=============================================================================
'ELIMINAR PROTESTO
'=============================================================================
    Public Sub eliminarProtesto(ByRef p As protesto)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND cheque = '" & p.cheque & "' AND rut = '" & p.rut & "' AND sucursal = '" & p.sucursal & "' "
        op = 4
        CAMPOS(0, 2) = "sv_protesto_" & empresaActiva
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR PROTESTO
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef p As protesto, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        p.cheque = sql.response(0, 3)
        p.rut = sql.response(1, 3)
        p.sucursal = sql.response(2, 3)
        p.fechaprotesto = sql.response(3, 3)
        p.MONTO = sql.response(4, 3)
        p.fechacheque = sql.response(5, 3)
        p.motivo = sql.response(6, 3)
        p.CANCELADO = sql.response(7, 3)
        p.GLOSA = sql.response(8, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef p As protesto, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "cheque"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "sucursal"
        CAMPOS(4, 0) = "fechaprotesto"
        CAMPOS(5, 0) = "monto"
        CAMPOS(6, 0) = "fechacheque"
        CAMPOS(7, 0) = "motivo"
        CAMPOS(8, 0) = "cancelado"
        CAMPOS(9, 0) = "glosa"
        CAMPOS(10, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = p.cheque
        CAMPOS(2, 1) = p.rut
        CAMPOS(3, 1) = p.sucursal
        CAMPOS(4, 1) = p.fechaprotesto
        CAMPOS(5, 1) = p.MONTO
        CAMPOS(6, 1) = p.fechacheque
        CAMPOS(7, 1) = p.motivo
        CAMPOS(8, 1) = p.CANCELADO
        CAMPOS(9, 1) = p.GLOSA
        CAMPOS(10, 1) = ""
                
        CAMPOS(0, 2) = "sv_protesto_" & empresaActiva
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



