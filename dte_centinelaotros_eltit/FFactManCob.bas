Attribute VB_Name = "FFactManCob"
''''''''''''''''''''''''''''''''''''''''
'''REVISAR RELACION CON BASE DE DATOS'''
''''''''''''''''''''''''''''''''''''''''

Option Explicit
    Private campos(30, 3) As String
    Public Type factManuales
        tipo As String
        numero As String
        fEmision As String
        fVencimiento As String
        rut As String
        sucursal As String
        cajera As String
        monto As String
        abono As String
        obs As String
    End Type

'=============================================================================
'LEER FACTURA MANUAL
'=============================================================================
    Public Function leerFactManual(ByRef f As factManuales, ByVal codigo As String, ByVal operador As String) As Boolean
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        campos(0, 0) = "tipo"
        campos(1, 0) = "numero"
        campos(2, 0) = "fechaemision"
        campos(3, 0) = "vencimiento"
        campos(4, 0) = "rut"
        campos(5, 0) = "sucursal"
        campos(6, 0) = "cajera"
        campos(7, 0) = "monto"
        campos(8, 0) = "abono"
        campos(9, 0) = "observaciones"
        campos(10, 0) = ""
        
        campos(0, 2) = "sv_documentos_cobranza"
        
        condicion = "local = '" & empresaActiva & "' AND tipo = 'FV' AND numero " & operador & " '" & codigo & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY numero DESC"
        Else
            condicion = condicion & "ORDER BY numero ASC"
        End If
        op = 5
        sql.datos = campos
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
        If sql.estado = 0 Then
            leerFactManual = True
            Call asigna(f, sql)
        Else
            leerFactManual = False
        End If
    End Function
'=============================================================================
'LEER FACTURA MANUAL
'=============================================================================

'=============================================================================
'GRABAR FACTURA MANUAL
'=============================================================================
    Public Sub grabarFactManCob(ByRef f As factManuales, ByVal modifica As Boolean)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        Call designa(f, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND tipo = '" & f.tipo & "' AND numero = '" & f.numero & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
        If modifica = True Then
            Call modificaDocumento(f)
        End If
    End Sub
    
    Private Sub modificaDocumento(ByRef f As factManuales)
        Dim campos(10, 3) As String
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        
        'CABEZA
        campos(0, 0) = "total"
        campos(1, 0) = "abono"
        campos(2, 0) = "vendedor"
        campos(3, 0) = "fecha"
        campos(4, 0) = "vencimiento"
        campos(5, 0) = ""
        
        campos(0, 1) = f.monto
        campos(1, 1) = f.abono
        campos(2, 1) = f.cajera
        campos(3, 1) = f.fEmision
        campos(4, 1) = f.fVencimiento
        campos(5, 1) = ""
        
        campos(0, 2) = "sv_documento_cabeza"
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & f.tipo & "' AND numero = '" & f.numero & "' "
        op = 3
        
        sql.datos = campos
        
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
        
        'DETALLE
        campos(0, 0) = "vendedor"
        campos(1, 0) = "fecha"
        campos(2, 0) = "vencimiento"
        campos(3, 0) = ""
        
        campos(0, 1) = f.cajera
        campos(1, 1) = f.fEmision
        campos(2, 1) = f.fVencimiento
        campos(3, 1) = ""
        
        campos(0, 2) = "sv_documento_detalle"
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & f.tipo & "' AND numero = '" & f.numero & "'"
        op = 3
        
        sql.datos = campos
        
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
    End Sub
'=============================================================================
'GRABAR FACTURA MANUAL
'=============================================================================

'=============================================================================
'ELIMINAR FACTURA MANUAL
'=============================================================================
    Public Sub eliminarFactManCob(ByVal tipo As String, ByVal numero As String)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & tipo & "' AND numero = '" & numero & "'"
        op = 4
        campos(0, 2) = "sv_documentos_cobranza"
        sql.datos = campos
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR FACTURA MANUAL
'=============================================================================


'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef f As factManuales, ByRef sql As CSQLUtil)
        f.tipo = sql.datos(0, 3)
        f.numero = sql.datos(1, 3)
        f.fEmision = sql.datos(2, 3)
        f.fVencimiento = sql.datos(3, 3)
        f.rut = sql.datos(4, 3)
        f.sucursal = sql.datos(5, 3)
        f.cajera = sql.datos(6, 3)
        f.monto = sql.datos(7, 3)
        f.abono = sql.datos(8, 3)
        f.obs = sql.datos(9, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef f As factManuales, ByRef sql As CSQLUtil)
        Dim cad As String
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "fechaemision"
        campos(4, 0) = "vencimiento"
        campos(5, 0) = "rut"
        campos(6, 0) = "sucursal"
        campos(7, 0) = "cajera"
        campos(8, 0) = "monto"
        campos(9, 0) = "abono"
        campos(10, 0) = "observaciones"
        campos(11, 0) = ""
        
        campos(0, 1) = empresaActiva
        campos(1, 1) = f.tipo
        campos(2, 1) = f.numero
        campos(3, 1) = f.fEmision
        campos(4, 1) = f.fVencimiento
        campos(5, 1) = f.rut
        campos(6, 1) = f.sucursal
        campos(7, 1) = f.cajera
        campos(8, 1) = f.monto
        campos(9, 1) = f.abono
        campos(10, 1) = f.obs
        campos(11, 1) = ""
        
        campos(0, 2) = "sv_documentos_cobranza"
        sql.datos = campos
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



