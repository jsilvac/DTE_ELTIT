Attribute VB_Name = "FCProrroga"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type prorroga
        NUMERO As String
        rut As String
        sucursal As String
        fcheque As String
        ncheque As String
        MONTO As String
        fprorroga As String
    End Type

    
'=============================================================================
'LEER PRORROGA
'=============================================================================
    Public Function leerProrroga(ByRef p As prorroga, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "numero"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "sucursal"
        CAMPOS(3, 0) = "fcheque"
        CAMPOS(4, 0) = "cheque"
        CAMPOS(5, 0) = "monto"
        CAMPOS(6, 0) = "fprorroga"
        CAMPOS(7, 0) = ""
        
        CAMPOS(0, 2) = "sv_prorroga"
        
        condicion = "local = '" & empresaActiva & "' AND numero " & operador & " '" & CODIGO & "' "
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
            leerProrroga = True
            Call asigna(p, sql)
        Else
            leerProrroga = False
        End If
    End Function
'=============================================================================
'LEER PRORROGA
'=============================================================================

'=============================================================================
'GRABAR PRORROGA
'=============================================================================
    Public Sub grabarProrroga(ByRef p As prorroga, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(p, sql)
        condicion = ""
        If modifica = True Then
            condicion = "numero = '" & p.NUMERO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        Call modificaCheque(p)
    End Sub
    
    Private Sub modificaCheque(ByRef p As prorroga)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        
        CAMPOS(0, 0) = "fechavencimiento"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 1) = p.fprorroga
        CAMPOS(1, 1) = ""
        
        CAMPOS(0, 2) = "sv_carteracheques"
        sql.response = CAMPOS
        
        condicion = "rut = '" & p.rut & "' AND numerocheque = '" & p.ncheque & "'"
        op = 3
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR PRORROGA
'=============================================================================

'=============================================================================
'ELIMINAR PRORROGA
'=============================================================================
    Public Sub eliminarProrroga(ByRef p As prorroga)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND numero = '" & p.NUMERO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_prorroga"
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
    Private Sub asigna(ByRef p As prorroga, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        p.NUMERO = sql.response(0, 3)
        p.rut = sql.response(1, 3)
        p.sucursal = sql.response(2, 3)
        p.fcheque = sql.response(3, 3)
        p.ncheque = sql.response(4, 3)
        p.MONTO = sql.response(5, 3)
        p.fprorroga = sql.response(6, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef p As prorroga, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "sucursal"
        CAMPOS(4, 0) = "fcheque"
        CAMPOS(5, 0) = "cheque"
        CAMPOS(6, 0) = "monto"
        CAMPOS(7, 0) = "fprorroga"
        CAMPOS(8, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = p.NUMERO
        CAMPOS(2, 1) = p.rut
        CAMPOS(3, 1) = p.sucursal
        CAMPOS(4, 1) = p.fcheque
        CAMPOS(5, 1) = p.ncheque
        CAMPOS(6, 1) = p.MONTO
        CAMPOS(7, 1) = p.fprorroga
        CAMPOS(8, 1) = ""
        
        CAMPOS(0, 2) = "sv_prorroga"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================


