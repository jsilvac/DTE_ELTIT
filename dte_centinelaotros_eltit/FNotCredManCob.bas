Attribute VB_Name = "FNotCredManCob"
''''''''''''''''''''''''''''''''''''''''
'''REVISAR RELACION CON BASE DE DATOS'''
''''''''''''''''''''''''''''''''''''''''

Option Explicit
    Private campos(30, 3) As String
    Public Type notCredManuales
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
'LEER NOTA CREDITO MANUAL
'=============================================================================
    Public Function leerNotCredManual(ByRef n As notCredManuales, ByVal codigo As String, ByVal operador As String) As Boolean
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        campos(0, 0) = "numero"
        campos(1, 0) = "fechaemision"
        campos(2, 0) = "vencimiento"
        campos(3, 0) = "rut"
        campos(4, 0) = "sucursal"
        campos(5, 0) = "cajera"
        campos(6, 0) = "monto * -1"
        campos(7, 0) = "abono * -1"
        campos(8, 0) = "observaciones"
        campos(9, 0) = "tipo"
        campos(10, 0) = ""
        
        campos(0, 2) = "sv_documentos_cobranza"
        
        condicion = "local = '" & empresaActiva & "' AND tipo = 'NV' AND numero " & operador & " '" & codigo & "' "
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
            leerNotCredManual = True
            Call asigna(n, sql)
        Else
            leerNotCredManual = False
        End If
    End Function
'=============================================================================
'LEER NOTA CREDITO MANUAL
'=============================================================================

'=============================================================================
'GRABAR NOTA CREDITO MANUAL
'=============================================================================
    Public Sub grabarNotCredManCob(ByRef n As notCredManuales, ByVal modifica As Boolean)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        Call designa(n, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND tipo = '" & n.tipo & "' AND numero = '" & n.numero & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
        If modifica = True Then
            Call modificaDocumento(n)
        End If
        'Call LimpiarCajas(NotCredManCob)
    End Sub
    
    Private Sub modificaDocumento(ByRef n As notCredManuales)
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
        
        campos(0, 1) = n.monto
        campos(1, 1) = n.abono
        campos(2, 1) = n.cajera
        campos(3, 1) = n.fEmision
        campos(4, 1) = n.fVencimiento
        campos(5, 1) = ""
        
        campos(0, 2) = "sv_documento_cabeza"
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & n.tipo & "' AND numero = '" & n.numero & "' "
        op = 3
        
        sql.datos = campos
        
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
        
        'DETALLE
        campos(0, 0) = "vendedor"
        campos(1, 0) = "fecha"
        campos(2, 0) = "vencimiento"
        campos(3, 0) = ""
        
        campos(0, 1) = n.cajera
        campos(1, 1) = n.fEmision
        campos(2, 1) = n.fVencimiento
        campos(3, 1) = ""
        
        campos(0, 2) = "sv_documento_detalle"
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & n.tipo & "' AND numero = '" & n.numero & "'"
        op = 3
        
        sql.datos = campos
        
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
    End Sub
'=============================================================================
'GRABAR NOTA CREDITO MANUAL
'=============================================================================

'=============================================================================
'ELIMINAR NOTA CREDITO MANUAL
'=============================================================================
    Public Sub eliminarNotCredManCob(ByRef n As notCredManuales)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & n.tipo & "' AND numero = '" & n.numero & "'"
        op = 4
        campos(0, 2) = "sv_documentos_cobranza"
        sql.datos = campos
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR NOTA CREDITO MANUAL
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef n As notCredManuales, ByRef sql As CSQLUtil)
        n.numero = sql.datos(0, 3)
        n.fEmision = sql.datos(1, 3)
        n.fVencimiento = sql.datos(2, 3)
        n.rut = sql.datos(3, 3)
        n.sucursal = sql.datos(4, 3)
        n.cajera = sql.datos(5, 3)
        n.monto = sql.datos(6, 3)
        n.abono = sql.datos(7, 3)
        n.obs = sql.datos(8, 3)
        n.tipo = sql.datos(9, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef n As notCredManuales, ByRef sql As CSQLUtil)
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
        campos(1, 1) = n.tipo
        campos(2, 1) = n.numero
        campos(3, 1) = n.fEmision
        campos(4, 1) = n.fVencimiento
        campos(5, 1) = n.rut
        campos(6, 1) = n.sucursal
        campos(7, 1) = n.cajera
        campos(8, 1) = n.monto
        campos(9, 1) = n.abono
        campos(10, 1) = n.obs
        campos(11, 1) = ""
        
        campos(0, 2) = "sv_documentos_cobranza"
        sql.datos = campos
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



