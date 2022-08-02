Attribute VB_Name = "FCredito"
Option Explicit
    Private CAMPOS(30, 3) As String
    
    Private Type creditoCabeza
        loc As String
        TIPO As String
        NUMERO As String
        rut As String
        montocompra As String
        piecompra As String
        fecha As String
        numerocuotas As String
        montocuotas As String
        cajera As String
    End Type
    
    Private Type creditoDetalle
        loc As String
        TIPO As String
        NUMERO As String
        rut As String
        numerocuota As String
        vencimientooriginal As String
        vencimientoactual As String
        montocuota As String
        capital As String
        interesventa As String
        interesmora As String
        abono As String
        saldo As String
        folioultimopago As String
        cajerapago As String
    End Type
    
    Public Type Creditos
        cabeza As creditoCabeza
        detalle As creditoDetalle
    End Type

    
'=============================================================================
'LEER PRORROGA
'=============================================================================
'    Public Function leerDeposito(ByRef d As deposito, ByVal codigo As String, ByVal operador As String) As Boolean
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        campos(0, 0) = "numero"
'        campos(1, 0) = "fecha"
'        campos(2, 0) = "banco"
'        campos(3, 0) = "monto"
'        campos(4, 0) = "tipo"
'        campos(5, 0) = "rut"
'        campos(6, 0) = ""
'
'        campos(0, 2) = "depositos_" & empresaactiva
'
'        condicion = "numero " & operador & " '" & codigo & "'"
'        If operador = "<" Then
'            condicion = condicion & "ORDER BY numero DESC"
'        Else
'            condicion = condicion & "ORDER BY numero ASC"
'        End If
'        op = 5
'        sql.response = campos
'        Set sql.conexion = gestion
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerDeposito = True
'            Call asigna(d, sql)
'        Else
'            leerDeposito = False
'        End If
'    End Function
'=============================================================================
'LEER PRORROGA
'=============================================================================

'=============================================================================
'GRABAR CREDITO
'=============================================================================
    Public Sub grabarCredito(ByRef cr As Creditos, ByVal modifica As Boolean, ByVal DIAPAGO As String)
        Call grabarCreditoCabeza(cr.cabeza, modifica)
        Call grabarCreditoDetalle(cr.detalle, modifica, CDbl(cr.cabeza.numerocuotas), cr.cabeza.fecha, DIAPAGO)
    End Sub
    
    Private Sub grabarCreditoCabeza(ByRef cc As creditoCabeza, ByVal modifica As Boolean)
    
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designaCabeza(cc, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & cc.loc & "' AND tipo = '" & cc.TIPO & "' numero = '" & cc.NUMERO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Private Sub grabarCreditoDetalle(ByRef cd As creditoDetalle, ByVal modifica As Boolean, ByVal CUOTAS As Double, ByVal fecha As String, ByVal DIAPAGO As String)
        
        Dim op As Integer
        Dim cuota As String
        Dim i As Long
        Dim vencimiento As String
        Dim mes As String
        Dim fechaaux As String
        Set sql = New sqlventas.sqlventa
        
        condicion = ""
        If modifica = True Then
            'Call eliminarCreditoDetalle(cd)
        End If
        op = 2
                
        '''''''''''''25 dias diferencia
        'CALCULAR FECHA VENCIMIENTO
        fechaaux = DIAPAGO & "-" & Format(fecha, "mm-YYYY")
        '''''''''''''
        If fechaaux < fecha Then
            fechaaux = DateAdd("m", 1, fechaaux)
        End If
        If DateDiff("d", fecha, fechaaux) >= 25 Then
            '+1
            vencimiento = DateAdd("m", -1, fechaaux)
        Else
            '+2
            vencimiento = fechaaux
        End If
        For i = 1 To CUOTAS
            cuota = Str(i)
            cuota = Mid(cuota, 2, Len(cuota))
            cd.numerocuota = cuota
            cd.vencimientooriginal = Format(DateAdd("m", i, vencimiento), "yyyy-mm-dd")
            
            Call designaDetalle(cd, sql)
                
            Set sql.conexion = gestion
            Call sql.sqlventas(op, condicion)
        Next i
    End Sub
'=============================================================================
'GRABAR CREDITO
'=============================================================================

'=============================================================================
'ELIMINAR CREDITO
'=============================================================================
'    Public Sub eliminarCredito(ByRef d As deposito)
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        condicion = "numero = '" & d.numero & "'"
'        op = 4
'        campos(0, 2) = "depositos_" & empresaactiva
'        sql.response = campos
'        Set sql.conexion = gestion
'        Call sql.sqlventas(op, condicion)
'    End Sub
'=============================================================================
'ELIMINAR CREDITO
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
'    Private Sub asigna(ByRef d As deposito, ByRef sql as sqlventas.sqlventa)
'        Dim cad As String
'        d.numero = sql.response(0, 3)
'        d.fecha = sql.response(1, 3)
'        d.Banco = sql.response(2, 3)
'        d.monto = sql.response(3, 3)
'        d.tipo = sql.response(4, 3)
'        d.rut = sql.response(5, 3)
'    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designaCabeza(ByRef cc As creditoCabeza, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "rut"
        CAMPOS(4, 0) = "montocompra"
        CAMPOS(5, 0) = "piecompra"
        CAMPOS(6, 0) = "fecha"
        CAMPOS(7, 0) = "numerocuotas"
        CAMPOS(8, 0) = "montocuotas"
        CAMPOS(9, 0) = "cajera"
        CAMPOS(10, 0) = ""
        
        CAMPOS(0, 1) = cc.loc
        CAMPOS(1, 1) = cc.TIPO
        CAMPOS(2, 1) = cc.NUMERO
        CAMPOS(3, 1) = cc.rut
        CAMPOS(4, 1) = cc.montocompra
        CAMPOS(5, 1) = cc.piecompra
        CAMPOS(6, 1) = cc.fecha
        CAMPOS(7, 1) = cc.numerocuotas
        CAMPOS(8, 1) = cc.montocuotas
        CAMPOS(9, 1) = "" 'cc.cajera
        CAMPOS(10, 1) = ""
        
        CAMPOS(0, 2) = "cuotas_cabeza"
        sql.response = CAMPOS
    End Sub
    
    Private Sub designaDetalle(ByRef cd As creditoDetalle, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "rut"
        CAMPOS(4, 0) = "numerocuota"
        CAMPOS(5, 0) = "vencimientooriginal"
        CAMPOS(6, 0) = "vencimientoactual"
        CAMPOS(7, 0) = "montocuota"
        CAMPOS(8, 0) = "capital"
        CAMPOS(9, 0) = "interesventa"
        CAMPOS(10, 0) = "interesmora"
        CAMPOS(11, 0) = "abono"
        CAMPOS(12, 0) = "saldo"
        CAMPOS(13, 0) = "foliopago"
        CAMPOS(14, 0) = "cajerapago"
        CAMPOS(15, 0) = ""
        
        CAMPOS(0, 1) = cd.loc
        CAMPOS(1, 1) = cd.TIPO
        CAMPOS(2, 1) = cd.NUMERO
        CAMPOS(3, 1) = cd.rut
        CAMPOS(4, 1) = cd.numerocuota
        CAMPOS(5, 1) = cd.vencimientooriginal
        CAMPOS(6, 1) = cd.vencimientooriginal
        CAMPOS(7, 1) = cd.montocuota
        CAMPOS(8, 1) = "" 'cd.capital
        CAMPOS(9, 1) = "" 'cd.interesventa
        CAMPOS(10, 1) = "" 'cd.interesmora
        CAMPOS(11, 1) = "" 'cd.abono
        CAMPOS(12, 1) = "" 'cd.saldo
        CAMPOS(13, 1) = "" 'cd.folioultimopago
        CAMPOS(14, 1) = "" 'cd.cajerapago
        CAMPOS(15, 1) = ""
        
        CAMPOS(0, 2) = "cuotas_detalle"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================




