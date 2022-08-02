Attribute VB_Name = "FdocManuales"
''''''''''''''''''''''''''''''''''''''''
'''REVISAR RELACION CON BASE DE DATOS'''
''''''''''''''''''''''''''''''''''''''''

Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type docManual
        TIPO As String
        NUMERO As String
        fEmision As String
        fVencimiento As String
        rut As String
        sucursal As String
        cajera As String
        MONTO As String
        abono As String
        obs As String
    End Type

'=============================================================================
'LEER FACTURA MANUAL
'=============================================================================
    Public Function leerDocManual(ByRef D As docManual, ByVal TIPO As String, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "tipo"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "fechaemision"
        CAMPOS(3, 0) = "vencimiento"
        CAMPOS(4, 0) = "rut"
        CAMPOS(5, 0) = "sucursal"
        CAMPOS(6, 0) = "vendedor"
        CAMPOS(7, 0) = "monto"
        CAMPOS(8, 0) = "abono"
        CAMPOS(9, 0) = "observaciones"
        CAMPOS(10, 0) = ""
        
        CAMPOS(0, 2) = "sv_documentos_cobranza_" & empresaActiva
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero " & operador & " '" & CODIGO & "'"
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
            leerDocManual = True
            Call asigna(D, sql)
        Else
            leerDocManual = False
        End If
    End Function
'=============================================================================
'LEER FACTURA MANUAL
'=============================================================================

'=============================================================================
'GRABAR FACTURA MANUAL
'=============================================================================
    Public Sub grabarDocManual(ByRef D As docManual, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(D, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND tipo = '" & D.TIPO & "' AND numero = '" & D.NUMERO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If modifica = True Then
            Call modificaDocumento(D)
        End If
    End Sub
    
    Private Sub modificaDocumento(ByRef D As docManual)
        Dim CAMPOS(10, 3) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        
        'CABEZA
        CAMPOS(0, 0) = "total"
        CAMPOS(1, 0) = "abono"
        CAMPOS(2, 0) = "vendedor"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "vencimiento"
        CAMPOS(5, 0) = ""
        
        CAMPOS(0, 1) = D.MONTO
        CAMPOS(1, 1) = D.abono
        CAMPOS(2, 1) = D.cajera
        CAMPOS(3, 1) = D.fEmision
        CAMPOS(4, 1) = D.fVencimiento
        CAMPOS(5, 1) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & D.TIPO & "' AND numero = '" & D.NUMERO & "' "
        op = 3
        
        sql.response = CAMPOS
        
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        
        'DETALLE
        CAMPOS(0, 0) = "vendedor"
        CAMPOS(1, 0) = "fecha"
        CAMPOS(2, 0) = "vencimiento"
        CAMPOS(3, 0) = ""
        
        CAMPOS(0, 1) = D.cajera
        CAMPOS(1, 1) = D.fEmision
        CAMPOS(2, 1) = D.fVencimiento
        CAMPOS(3, 1) = ""
        
        CAMPOS(0, 2) = "sv_documento_detalle_" + empresaActiva
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & D.TIPO & "' AND numero = '" & D.NUMERO & "'"
        op = 3
        
        sql.response = CAMPOS
        
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR FACTURA MANUAL
'=============================================================================

'=============================================================================
'ELIMINAR FACTURA MANUAL
'=============================================================================
    Public Sub eliminarDocManual(ByVal TIPO As String, ByVal NUMERO As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero = '" & NUMERO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_documentos_cobranza_" & empresaActiva
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR FACTURA MANUAL
'=============================================================================


'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef D As docManual, ByRef sql As sqlventas.sqlventa)
        D.TIPO = sql.response(0, 3)
        D.NUMERO = sql.response(1, 3)
        D.fEmision = sql.response(2, 3)
        D.fVencimiento = sql.response(3, 3)
        D.rut = sql.response(4, 3)
        D.sucursal = sql.response(5, 3)
        D.cajera = sql.response(6, 3)
        D.MONTO = sql.response(7, 3)
        D.abono = sql.response(8, 3)
        D.obs = sql.response(9, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef D As docManual, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "fechaemision"
        CAMPOS(4, 0) = "vencimiento"
        CAMPOS(5, 0) = "rut"
        CAMPOS(6, 0) = "sucursal"
        CAMPOS(7, 0) = "cajera"
        CAMPOS(8, 0) = "monto"
        CAMPOS(9, 0) = "abono"
        CAMPOS(10, 0) = "observaciones"
        CAMPOS(11, 0) = "vendedor"
        CAMPOS(12, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = D.TIPO
        CAMPOS(2, 1) = D.NUMERO
        CAMPOS(3, 1) = D.fEmision
        CAMPOS(4, 1) = D.fVencimiento
        CAMPOS(5, 1) = D.rut
        CAMPOS(6, 1) = D.sucursal
        CAMPOS(7, 1) = D.cajera
        CAMPOS(8, 1) = D.MONTO
        CAMPOS(9, 1) = D.abono
        CAMPOS(10, 1) = D.obs
        CAMPOS(11, 1) = D.cajera
        CAMPOS(12, 1) = ""
        
        CAMPOS(0, 2) = "sv_documentos_cobranza_" & empresaActiva
        
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================




