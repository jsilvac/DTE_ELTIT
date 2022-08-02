Attribute VB_Name = "FPAuditoriaVentas"
Option Explicit
    Public montoVenta As String
    Public fechaAuditIni As String
    Public fechaAuditFin As String
    Private CAMPOS(30, 3) As String
    Public titulo As String
    Private Type facturas
        cantidad As String
        Descuento As String
        nulas As String
        total As String
        folini As String
        folfin As String
    End Type
    
    Private Type boletas
        cantidad As String
        Descuento As String
        nulas As String
        total As String
        folini As String
        folfin As String
    End Type
    
    Private Type zetas
        cantidad As String
        Descuento As String
        nulas As String
        total As String
        folini As String
        folfin As String
    End Type
    
    Private Type exentas
        cantidad As String
        Descuento As String
        nulas As String
        total As String
        folini As String
        folfin As String
    End Type
    
    Private Type notascredito
        cantidad As String
        Descuento As String
        nulas As String
        total As String
        folini As String
        folfin As String
    End Type
    
    Private Type Ingresos
        efectivo As String
        Cheques As String
        tarjetas As String
        Depositos As String
        pagoClientes As String
    
    
    End Type
    
    Private Type Egresos
        egresosCaja As String
        chequesFecha As String
        CREDITO As String
        Depositos As String
    End Type
    
    Public Type auditoria
        boleta As boletas
        factura As facturas
        zeta As zetas
        exenta As exentas
        ncredito As notascredito
        egreso As Egresos
        Ingreso As Ingresos
    End Type

    
'=============================================================================
'LEER AUDITORIA
'=============================================================================
    Public Function leerauditoria(ByRef a As auditoria, fecha1 As String, fecha2 As String) As Boolean
        leerauditoria = False
        
        If leerAuditoriaBoleta(a.boleta, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaFactura(a.factura, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaZeta(a.zeta, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaExenta(a.exenta, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaNCredito(a.ncredito, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        
        
        If leerAuditoriaEgresos(a.egreso, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaNulasBoletas(a.boleta, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaNulasFacturas(a.factura, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaNulasZetas(a.zeta, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaNulasExentas(a.exenta, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        If leerAuditoriaNulasNCreditos(a.ncredito, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
        
        montoVenta = CDbl(a.boleta.total) + CDbl(a.factura.total) + CDbl(a.zeta.total) + CDbl(a.exenta.total) + CDbl(a.ncredito.total)
        
        If leerAuditoriaIngresos(a.Ingreso, fecha1, fecha2) = True Then
            leerauditoria = True
        End If
    End Function
    
    Public Function leerAuditoriaBoleta(ByRef ab As boletas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
        '
        Dim op As Integer
        Dim cajeras As String
        Dim caja As String
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(dc.numero)"
        CAMPOS(1, 0) = "FORMAT(IFNULL(SUM(dc.descuento),'0'),0)"
        CAMPOS(2, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(dc.total),'0'),0))"
        CAMPOS(3, 0) = "IFNULL(MIN(dc.numero),'')"
        CAMPOS(4, 0) = "IFNULL(MAX(dc.numero),'')"
        CAMPOS(5, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " AS dc"
        cajeras = PAuditoriaVentas.tcajera.text
        caja = PAuditoriaVentas.tcaja.text
        
        
        condicion = "dc.local = '" & localAuditoria & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.tipo = 'BV' and dc.cajera like '%" + cajeras + "%' and dc.caja like '%" + caja + "%' and caja < '90'"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        
        If sql.Status = 0 Then
            leerAuditoriaBoleta = True
            Call asignaBoleta(ab, sql)
        Else
            leerAuditoriaBoleta = False
        End If
    End Function
    
    Public Function leerAuditoriaFactura(ByRef af As facturas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
        
        Dim op As Integer
        Dim cajeras As String
        Dim caja As String
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(dc.numero)"
        CAMPOS(1, 0) = "FORMAT(IFNULL(SUM(dc.descuento),'0'),0)"
        CAMPOS(2, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(dc.total),'0'),0))"
        CAMPOS(3, 0) = "IFNULL(MIN(dc.numero),'')"
        CAMPOS(4, 0) = "IFNULL(MAX(dc.numero),'')"
        CAMPOS(5, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " AS dc"
        cajeras = PAuditoriaVentas.tcajera.text
        caja = PAuditoriaVentas.tcaja.text
        
        condicion = "dc.local = '" & localAuditoria & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.tipo = 'FV'  and dc.cajera like '%" + cajeras + "%' and dc.caja like '%" + caja + "%' and caja < '90'"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaFactura = True
            Call asignaFactura(af, sql)
        Else
            leerAuditoriaFactura = False
        End If
    End Function
    
    Public Function leerAuditoriaZeta(ByRef az As zetas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(dc.numero)"
        CAMPOS(1, 0) = "FORMAT(IFNULL(SUM(dc.descuento),'0'),0)"
        CAMPOS(2, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(dc.total),'0'),0))"
        CAMPOS(3, 0) = "IFNULL(MIN(dc.boletadesde),'')"
        CAMPOS(4, 0) = "IFNULL(MAX(dc.boletahasta),'')"
        CAMPOS(5, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " AS dc"
        
        condicion = "dc.local = '" & localAuditoria & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.tipo = 'ZE' -- AND dc.nula = 'N'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaZeta = True
            Call asignaZeta(az, sql)
        Else
            leerAuditoriaZeta = False
        End If
    End Function
    
    Public Function leerAuditoriaExenta(ByRef ae As exentas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
    
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(dc.numero)"
        CAMPOS(1, 0) = "FORMAT(IFNULL(SUM(dc.descuento),'0'),0)"
        CAMPOS(2, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(dc.total),'0'),0))"
        CAMPOS(3, 0) = "IFNULL(MIN(dc.numero),'')"
        CAMPOS(4, 0) = "IFNULL(MAX(dc.numero),'')"
        CAMPOS(5, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " AS dc"
        
        condicion = "dc.local = '" & localAuditoria & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.tipo = 'FE' -- AND dc.nula = 'N'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaExenta = True
            Call asignaExenta(ae, sql)
        Else
            leerAuditoriaExenta = False
        End If
    End Function
    
    Public Function leerAuditoriaNCredito(ByRef an As notascredito, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
       
        Dim op As Integer
        Dim cajeras As String
        Dim caja As String
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(dc.numero)"
        CAMPOS(1, 0) = "FORMAT(IFNULL(SUM(dc.descuento),'0'),0)"
        CAMPOS(2, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(dc.total),'0'),0))"
        CAMPOS(3, 0) = "IFNULL(MIN(dc.numero),'')"
        CAMPOS(4, 0) = "IFNULL(MAX(dc.numero),'')"
        CAMPOS(5, 0) = ""
        cajeras = PAuditoriaVentas.tcajera.text
        caja = PAuditoriaVentas.tcaja.text
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " AS dc"
        
        condicion = "dc.local = '" & localAuditoria & "' AND dc.cajera like '%" + cajeras + "%' and dc.caja like '%" + caja + "%' and dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND (dc.tipo = 'NB' or dc.tipo='NF') -- AND dc.nula = 'N'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaNCredito = True
            Call asignaNCredito(an, sql)
        Else
            leerAuditoriaNCredito = False
        End If
    End Function
    
    Public Function leerAuditoriaNulasBoletas(ByRef ab As boletas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
       
        Dim op As Integer
        Dim cajeras As String
        Dim caja As String
        
        Set sql = New sqlventas.sqlventa
        
        CAMPOS(0, 0) = "COUNT(numero)"
        CAMPOS(1, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(total),'0'),0))"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text
        cajeras = PAuditoriaVentas.tcajera.text
        caja = PAuditoriaVentas.tcaja.text
        
        condicion = "local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND tipo = 'BV' and nula='S' and cajera like '%" + cajeras + "%' and caja like '%" + caja + "%' and caja < '90' group by tipo"
        
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaNulasBoletas = True
            Call asignaNulasBoletas(ab, sql)
        Else
            leerAuditoriaNulasBoletas = False
            ab.nulas = "0"
        End If
    End Function
    
    Public Function leerAuditoriaNulasFacturas(ByRef af As facturas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
        
        Dim op As Integer
        Dim cajeras As String
        Dim caja As String
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(numero)"
        CAMPOS(1, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(total),'0'),0))"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text
        cajeras = PAuditoriaVentas.tcajera.text
        caja = PAuditoriaVentas.tcaja.text
        
        condicion = "local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND tipo = 'BV' and nula='S' and cajera like '%" + cajeras + "%' and caja like '%" + caja + "%' and caja  < '90' group by tipo"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaNulasFacturas = True
            Call asignaNulasFacturas(af, sql)
        Else
            leerAuditoriaNulasFacturas = False
            af.nulas = "0"
        End If
    End Function
    
    Public Function leerAuditoriaNulasZetas(ByRef az As zetas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
       
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(numero)"
        CAMPOS(1, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(total),'0'),0))"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text
        
        condicion = "local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND nula = 'S' AND tipo = 'ZE' GROUP BY tipo "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaNulasZetas = True
            Call asignaNulasZetas(az, sql)
        Else
            leerAuditoriaNulasZetas = False
            az.nulas = "0"
        End If
    End Function
    
    Public Function leerAuditoriaNulasExentas(ByRef ae As exentas, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
       
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(numero)"
        CAMPOS(1, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(total),'0'),0))"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text
        
        condicion = "local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND nula = 'S' AND tipo = 'FE' GROUP BY tipo "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaNulasExentas = True
            Call asignaNulasExentas(ae, sql)
        Else
            leerAuditoriaNulasExentas = False
            ae.nulas = "0"
        End If
    End Function
    
    Public Function leerAuditoriaNulasNCreditos(ByRef an As notascredito, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
      
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "COUNT(numero)"
        CAMPOS(1, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(total),'0'),0))"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + PAuditoriaVentas.dato1.text
        
        condicion = "local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND nula = 'S' AND tipo = 'NV' GROUP BY tipo "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaNulasNCreditos = True
            Call asignaNulasNCreditos(an, sql)
        Else
            leerAuditoriaNulasNCreditos = False
            an.nulas = "0"
        End If
        'an.cantidad = "0"
        'an.descuento = "0"
        'an.nulas = "0"
        'an.total = "0"
    End Function
    
    Public Function leerAuditoriaIngresos(ByRef ai As Ingresos, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
        leerAuditoriaIngresos = False
    
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim cajeras As String
        Dim caja As String
    
    
    For i = 1 To 14
    PAuditoriaVentas.Ingresos.Cell(i, 1).text = "0"
    Next i
        cajeras = PAuditoriaVentas.tcajera.text
        caja = PAuditoriaVentas.tcaja.text

        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasAuditoria
        
        csql.sql = "SELECT dp.tipopago,sum(dp.monto) "
        csql.sql = csql.sql & "FROM sv_documento_pagos_" + PAuditoriaVentas.dato1.text + " as dp, sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " as dc "
        csql.sql = csql.sql + "WHERE dc.caja=dp.caja and dp.tipo=dc.tipo and dp.numero=dc.numero and  dc.fecha=dp.fecha and dp.local = '" & localAuditoria & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND (dc.tipo = 'BV' OR dc.tipo='ZE' OR dc.tipo='FV') and dp.tipopago>'1' and dc.caja like '%" + caja + "%' and dc.cajera like '%" + cajeras + "%' and dc.caja < '90' GROUP BY tipoPAGO "
        
        csql.Execute
        
        If csql.RowsAffected > 0 Then
          
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
         
            PAuditoriaVentas.Ingresos.Cell(CDbl(resultado(0)), 1).text = Format(resultado(1), "###,###,###")
            
            resultado.MoveNext
            
            Wend
         End If
        Set resultado = Nothing
        csql.Close
        
        Set csql = Nothing
      
      
      Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        
        csql.sql = "SELECT sum(monto) "
        csql.sql = csql.sql & "FROM sv_cuotas_pago_cabeza "
        csql.sql = csql.sql + "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and local='" + localAuditoria + "' and cajero like  '%" + cajeras + "%' "
        
        csql.Execute
    
        If csql.RowsAffected > 0 Then
            
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
         
            PAuditoriaVentas.Ingresos.Cell(13, 1).text = Format(resultado(0), "###,###,###")
            
            resultado.MoveNext
            
            Wend
         End If
        Set resultado = Nothing
        csql.Close
      
    
    
    End Function
        
    
    
    Public Function leerAuditoriaEgresos(ByRef ae As Egresos, ByVal fecha1 As String, ByVal fecha2 As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        leerAuditoriaEgresos = False
        'EGRESOS DE CAJA
'        campos(0, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(monto),0),0))"
'        campos(1, 0) = ""
'        campos(0, 2) = "sv_egresoscaja_" + rubro
'        condicion = "local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "'"
'        op = 5
'        sql.response = campos
'        Set sql.conexion = ventasAuditoria
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerAuditoriaEgresos = True
'            ae.egresosCaja = Replace(sql.response(0, 3), ",", ".")
'        End If
        'CHEQUES A FECHA
'        campos(0, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(cch.monto),0),0))"
'        campos(1, 0) = ""
'        campos(0, 2) = "sv_carteracheques AS cch"
'        condicion = "cch.local = '" & localAuditoria & "' AND cch.fecharecepcion BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND cch.fechavencimiento > '" & fecha2 & "'"
'        op = 5
'        sql.response = campos
'        Set sql.conexion = ventasAuditoria
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerAuditoriaEgresos = True
'            ae.chequesFecha = Replace(sql.response(0, 3), ",", ".")
'        End If
       'creditos
        CAMPOS(0, 0) = "ifnull(SUM(dp.monto),0)"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = "sv_documento_pagos_" + PAuditoriaVentas.dato1.text + " AS dp INNER JOIN sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " AS dc ON dc.local = dp.local AND dc.tipo = dp.tipo AND dc.numero = dp.numero "
        condicion = "dc.local = '" & localAuditoria & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dp.tipopago > '2' and (dp.tipo='FV' OR dp.TIPO='BV' OR dp.TIPO='ZE') AND dc.nula = 'N'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerAuditoriaEgresos = True
            PAuditoriaVentas.Egresos.Cell(3, 1).text = sql.response(0, 3)
        End If
       ' Depositos
'        campos(0, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(d.monto),0),0))"
'        campos(1, 0) = ""
'        campos(0, 2) = "sv_depositos AS d"
'        condicion = "d.local = '" & localAuditoria & "' AND d.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "'"
'        op = 5
'        sql.response = campos
'        Set sql.conexion = ventasAuditoria
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerAuditoriaEgresos = True
'            ae.Depositos = Replace(sql.response(0, 3), ",", ".")
'        End If
'        campos(0, 0) = "CONCAT('$ ',FORMAT(IFNULL(SUM(monto),0),0))"
'        campos(1, 0) = ""
'        campos(0, 2) = "sv_pagos_cabeza_" & empresaActiva
'        condicion = "local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND tipopago = '3'"
'        op = 5
'        sql.response = campos
'        Set sql.conexion = ventasAuditoria
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerAuditoriaEgresos = True
'            If ae.Depositos = "" Then
'            ae.Depositos = 0
'            End If
           ' ae.Depositos = Format(CDbl(ae.Depositos) + CDbl(Replace(sql.response(0, 3), ",", ".")), "$ ###,###,##0")
            'ae.Depositos = Format(CDbl(ae.Depositos) + CDbl(sql.response(0, 3)), "$ ###,###,##0")
'        End If
    End Function
'=============================================================================
'LEER PRORROGA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaBoleta(ByRef ab As boletas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        ab.cantidad = sql.response(0, 3)
        ab.Descuento = Replace(sql.response(1, 3), ",", ".")
        ab.total = Replace(sql.response(2, 3), ",", ".")
        ab.folini = sql.response(3, 3)
        ab.folfin = sql.response(4, 3)
    End Sub
    
    Private Sub asignaFactura(ByRef af As facturas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        af.cantidad = sql.response(0, 3)
        af.Descuento = Replace(sql.response(1, 3), ",", ".")
        af.total = Replace(sql.response(2, 3), ",", ".")
        af.folini = sql.response(3, 3)
        af.folfin = sql.response(4, 3)
    End Sub
    
    Private Sub asignaZeta(ByRef az As zetas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        az.cantidad = sql.response(0, 3)
        az.Descuento = Replace(sql.response(1, 3), ",", ".")
        az.total = Replace(sql.response(2, 3), ",", ".")
        az.folini = sql.response(3, 3)
        az.folfin = sql.response(4, 3)
        'If az.folini <> "" Or az.folfin <> "" Then
        '    az.cantidad = CDbl(az.folfin) - CDbl(az.folini) + 1
        'End If
    End Sub
    
    Private Sub asignaExenta(ByRef ae As exentas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        ae.cantidad = sql.response(0, 3)
        ae.Descuento = Replace(sql.response(1, 3), ",", ".")
        ae.total = Replace(sql.response(2, 3), ",", ".")
        ae.folini = sql.response(3, 3)
        ae.folfin = sql.response(4, 3)
    End Sub
    
    Private Sub asignaNCredito(ByRef an As notascredito, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        an.cantidad = sql.response(0, 3)
        If CDbl(Replace(sql.response(1, 3), ",", ".")) < 0 Then
            an.Descuento = Format(CDbl(Replace(sql.response(1, 3), ",", ".")), "$ ###,###,##0")
        Else
            an.Descuento = Format(0 - CDbl(Replace(sql.response(1, 3), ",", ".")), "$ ###,###,##0")
        End If
        If CDbl(Replace(sql.response(2, 3), ",", ".")) < 0 Then
            an.total = Format(CDbl(Replace(sql.response(2, 3), ",", ".")), "$ ###,###,##0")
        Else
            an.total = Format(0 - CDbl(Replace(sql.response(2, 3), ",", ".")), "$ ###,###,##0")
        End If
        an.folini = sql.response(3, 3)
        an.folfin = sql.response(4, 3)
    End Sub
    
    Private Sub asignaNulasBoletas(ByRef ab As boletas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        ab.nulas = sql.response(0, 3)
        ab.cantidad = ab.cantidad - ab.nulas
        ab.total = Format(CDbl(ab.total) - CDbl(Replace(sql.response(1, 3), ",", ".")), "$ ###,###,##0")
    End Sub
    
    Private Sub asignaNulasFacturas(ByRef af As facturas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        af.nulas = sql.response(0, 3)
        af.cantidad = af.cantidad - af.nulas
        af.total = Format(CDbl(af.total) - CDbl(Replace(sql.response(1, 3), ",", ".")), "$ ###,###,##0")
    End Sub
    
    Private Sub asignaNulasZetas(ByRef az As zetas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        az.nulas = sql.response(0, 3)
        az.cantidad = az.cantidad - az.nulas
        az.total = Format(CDbl(az.total) - CDbl(Replace(sql.response(1, 3), ",", ".")), "$ ###,###,##0")
    End Sub
    
    Private Sub asignaNulasExentas(ByRef ae As exentas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        ae.nulas = sql.response(0, 3)
        ae.cantidad = ae.cantidad - ae.nulas
        ae.total = Format(CDbl(ae.total) - CDbl(Replace(sql.response(1, 3), ",", ".")), "$ ###,###,##0")
    End Sub
    
    Private Sub asignaNulasNCreditos(ByRef an As notascredito, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        an.nulas = sql.response(0, 3)
        an.cantidad = an.cantidad - an.nulas
        an.total = Format(CDbl(an.total) + CDbl(Replace(sql.response(1, 3), ",", ".")), "$ ###,###,##0")
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================




