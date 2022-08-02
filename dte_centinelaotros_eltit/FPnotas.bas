Attribute VB_Name = "FPVentas"
Option Explicit
    Public autorizado As Boolean
    Public cupo As Double
    Private campos(35, 3) As String
    Private Type ventaCabeza
        loc As String
        tipo As String
        numero As String
        fecha As String
        plazo As String
        vencimiento As String
        rut As String
        sucursal As String
        cajera As String
        notapedido As String
        notaventas As String
        ordencompra As String
        subtotal As String
        neto As String
        iva As String
        impuestoharina As String
        impuestoila As String
        impuestoespecifico As String
        exento As String
        retencionparcial As String
        retenciontotal As String
        total As String
        abono As String
        descuento As String
        contabilizado As String
        PAGADO As String
        comision As String
        fechapagocomision As String
        nula As String
        boletadesde As String
        boletahasta As String
        descuento2 As String
        
    End Type
        
    Private Type ventaDetalle
        loc As String
        tipo As String
        numero As String
        linea As String
        fecha As String
        rut As String
        sucursal As String
        codigo As String
        descripcion As String
        cantidad As String
        unidades As String
        precio As String
        descuento As String
        total As String
        vendedor As String
        pcosto As String
        bodega As String
        vencimiento As String
        numerofactura As String
        descuento2 As String
        
    End Type
        
    Public Type venta
        cabeza As ventaCabeza
        detalle As ventaDetalle
    End Type
    
'=============================================================================
'LEER VENTA
'=============================================================================
    Public Function leerVenta(ByRef v As venta, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String, Optional data As Variant, Optional Lista As Variant) As Boolean
        leerVenta = True
        If leerVentaCabeza(v.cabeza, codigo1, codigo2, operador) = False Then
            leerVenta = False
        Else
            If Not IsMissing(data) Then
                Call leerVentaDetalle(data, codigo1, v.cabeza.numero, "=", Lista)
            End If
        End If
    End Function
    
    Private Function leerVentaCabeza(ByRef vc As ventaCabeza, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String) As Boolean
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "fecha"
        campos(4, 0) = "plazo"
        campos(5, 0) = "IFNULL(vencimiento,'1900-01-01')"
        campos(6, 0) = "rut"
        campos(7, 0) = "sucursal"
        campos(8, 0) = "vendedor"
        campos(9, 0) = "notapedido"
        campos(10, 0) = "notaventa"
        campos(11, 0) = "ordencompra"
        campos(12, 0) = "subtotal"
        campos(13, 0) = "neto"
        campos(14, 0) = "iva"
        campos(15, 0) = "impuestoharina"
        campos(16, 0) = "impuestoila"
        campos(17, 0) = "impuestoespecifico"
        campos(18, 0) = "exento"
        campos(19, 0) = "retencionparcial"
        campos(20, 0) = "retenciontotal"
        campos(21, 0) = "total"
        campos(22, 0) = "abono"
        campos(23, 0) = "descuento"
        campos(24, 0) = "contabilizado"
        campos(25, 0) = "pagado"
        campos(26, 0) = "comision"
        campos(27, 0) = "IFNULL(fechapagocomision,'1900-01-01')"
        campos(28, 0) = "nula"
        campos(29, 0) = "boletadesde"
        campos(30, 0) = "boletahasta"
        campos(31, 0) = ""

        campos(0, 2) = "sv_documento_cabeza"
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' "
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
            leerVentaCabeza = True
            Call asignaCabeza(vc, sql)
        Else
            leerVentaCabeza = False
        End If
    End Function
    
    Private Sub leerVentaDetalle(ByRef data As Variant, ByVal codigo1 As String, ByVal codigo2 As String, ByVal operador As String, ByRef Lista As Variant)
        Dim tabla As String
        tabla = "SELECT CONCAT(codigo, '" & vbTab & "', descripcion, '" & vbTab & "', IF(numerofactura<>'', numerofactura, cantidad), '" & vbTab & "', unidades, '" & vbTab & "', precio, '" & vbTab & "', descuento, '" & vbTab & "', total) AS item "
        tabla = tabla & "FROM sv_documento_detalle "
        tabla = tabla & "WHERE local = '" & empresaActiva & "' AND tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' ORDER BY linea ASC"
        Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        Lista.Rows = 1
        Lista.AutoRedraw = False
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            
            While Not data.Recordset.EOF
                Lista.AddItem Replace(data.Recordset.Fields("item"), ".", ","), True
                data.Recordset.MoveNext
            Wend
        End If
        Lista.AutoRedraw = True
        Lista.Refresh
    End Sub
'=============================================================================
'LEER VENTA
'=============================================================================

'=============================================================================
'GRABAR VENTA
'=============================================================================
    Public Sub grabarVenta(ByRef v As venta, ByVal modifica As Boolean, ByRef Lista As Grid)
        Call grabarVentaCabeza(v, modifica)
        Call grabarVentaDetalle(Lista, v, modifica)
        'Call grabarVentaMovimientos(v, modifica)
    End Sub
    
    Private Sub grabarVentaCabeza(ByRef v As venta, ByVal modifica As Boolean)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        Call designaCabeza(v, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND local = '" & empresaActiva & "' AND tipo = '" & v.cabeza.tipo & "' AND numero = '" & v.cabeza.numero & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
        op = sql.estado
    End Sub
    
    Private Sub grabarVentaDetalle(ByRef Lista As Grid, ByRef v As venta, ByVal modifica As Boolean)
        Dim descripcion As String
        Dim condicion As String
        Dim op As Integer
        Dim i As Long
        Dim lin As String
        Set sql = New CSQLUtil
        
        condicion = ""
        If modifica = True Then
            Call eliminarVentaDetalle(v)
        End If
        op = 2
        
        For i = 1 To Lista.Rows - 1
            lin = Str(i)
            lin = Mid(lin, 2, Len(lin))
            If Lista.Cell(i, 1).text <> "" And Lista.Cell(i, 2).text <> "" And Lista.Cell(i, 3).text <> "" And Lista.Cell(i, 4).text <> "" And Lista.Cell(i, 5).text <> "" Then
                v.detalle.linea = String(3 - Len(lin), "0") & lin
                v.detalle.codigo = Lista.Cell(i, 1).text
                v.detalle.descripcion = Lista.Cell(i, 2).text
                
                v.detalle.unidades = Replace(Format(Lista.Cell(i, 4).text, "########0.00"), ",", ".")
                v.detalle.precio = Replace(Format(Lista.Cell(i, 5).text, "########0.00"), ",", ".")
                v.detalle.descuento = Replace(Format(Lista.Cell(i, 6).text, "########0.00"), ",", ".")
                v.detalle.descuento2 = Replace(Format(Lista.Cell(i, 8).text, "########0.00"), ",", ".")
                v.detalle.total = Replace(Format(Lista.Cell(i, 7).text, "########0.00"), ",", ".")
                v.detalle.pcosto = Lista.Cell(i, Lista.Cols - 1).text
                
                If Val(Lista.Cell(i, 1).text) = 100 Then
                    v.detalle.cantidad = "1"
                    v.detalle.numerofactura = Lista.Cell(i, 3).text
                    If v.cabeza.tipo = "NV" Then
                        Call grabarNotasFactura(v)
                        If v.detalle.codigo = "0000000000100" Then
                            Call modificarAbono(v)
                        End If
                    End If
                Else
                    v.detalle.cantidad = Replace(Format(Lista.Cell(i, 3).text, "########0.00"), ",", ".")
                    v.detalle.numerofactura = ""
                End If
                
                Call designaDetalle(v, sql, descripcion)
                
                Set sql.conexion = ventasRubro
                Call sql.SQLUTIL(op, condicion)
            End If
        Next i
    End Sub
    
    Private Sub grabarVentaMovimientos(ByRef v As venta, ByVal modifica As Boolean)
        Dim cSql As New rdoQuery
        Set cSql.ActiveConnection = gestionRubro
        If modifica = False Then
            cSql.sql = "INSERT INTO l_movimientos_detalle_" & empresaActiva & " "
            cSql.sql = cSql.sql & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
            cSql.sql = cSql.sql & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.unidades, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, ROUND(dd.unidades / dd.cantidad, 0) "
            cSql.sql = cSql.sql & "FROM " & baseVentas & rubro & ".sv_documento_detalle AS dd "
            cSql.sql = cSql.sql & "WHERE dd.local = '" & empresaActiva & "' AND dd.tipo = '" & v.detalle.tipo & "' AND dd.numero = '" & v.detalle.numero & "'"
            cSql.Execute
        Else
            cSql.sql = "DELETE FROM l_movimientos_detalle_" & empresaActiva & " "
            cSql.sql = cSql.sql & "WHERE tipo = '" & v.detalle.tipo & "' AND numero = '" & v.detalle.numero & "'"
            cSql.Execute
            modifica = False
            Call grabarVentaMovimientos(v, modifica)
        End If
        cSql.Close
        Set cSql = Nothing
    End Sub
    
    Private Sub grabarNotasFactura(ByRef v As venta)
        Dim cSql As New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "INSERT INTO sv_documento_notas "
        cSql.sql = cSql.sql & "(local, numero, localfactura, numerofactura, monto) "
        cSql.sql = cSql.sql & "VALUES('" & empresaActiva & "', '" & v.cabeza.numero & "', '" & empresaActiva & "', '" & v.detalle.numerofactura & "', '" & v.detalle.total & "') "
        cSql.sql = cSql.sql & "ON DUPLICATE KEY UPDATE local = '" & empresaActiva & "', numero = '" & v.cabeza.numero & "', localfactura = '" & empresaActiva & "', numerofactura = '" & v.detalle.numerofactura & "', monto = '" & v.detalle.total & "' "
        cSql.Execute
        cSql.Close
        Set cSql = Nothing
    End Sub
    
    Private Sub modificarAbono(ByRef v As venta)
        Dim cSql As rdoQuery
        Dim abono As Double
        Dim abonoIva As Double
        Dim abonoIha As Double
        abono = -1 * CDbl(Replace(v.detalle.total, ".", ","))
        abonoIva = abono * iva / 100
        abonoIha = abono * iha / 100
        abono = Round(abono + abonoIva + abonoIha, 0)
        
        'DOCUMENTO CABEZA(FACTURA)
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "UPDATE sv_documento_cabeza "
        cSql.sql = cSql.sql & "SET abono = abono + " & abono & " "
        cSql.sql = cSql.sql & "WHERE local = '" & v.cabeza.loc & "' AND tipo = 'FV' AND numero = '" & v.detalle.numerofactura & "' "
        cSql.Execute
        cSql.Close
        Set cSql = Nothing
        
        'DOCUMENTOS COBRANZA(FACTURA)
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "UPDATE sv_documentos_cobranza "
        cSql.sql = cSql.sql & "SET abono = abono + " & abono & " "
        cSql.sql = cSql.sql & "WHERE local = '" & v.cabeza.loc & "' AND tipo = 'FV' AND numero = '" & v.detalle.numerofactura & "' "
        cSql.Execute
        cSql.Close
        Set cSql = Nothing
        
        abono = abono * -1
        'DOCUMENTO CABEZA(NOTA CREDITO)
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "UPDATE sv_documento_cabeza "
        cSql.sql = cSql.sql & "SET abono = abono + " & abono & " "
        cSql.sql = cSql.sql & "WHERE local = '" & v.cabeza.loc & "' AND tipo = 'NV' AND numero = '" & v.cabeza.numero & "' "
        cSql.Execute
        cSql.Close
        Set cSql = Nothing
        
        'DOCUMENTO CABEZA(NOTA CREDITO)
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "UPDATE sv_documentos_cobranza "
        cSql.sql = cSql.sql & "SET abono = abono + " & abono & " "
        cSql.sql = cSql.sql & "WHERE local = '" & v.cabeza.loc & "' AND tipo = 'NV' AND numero = '" & v.cabeza.numero & "' "
        cSql.Execute
        cSql.Close
        Set cSql = Nothing
    End Sub
'=============================================================================
'GRABAR VENTA
'=============================================================================

'=============================================================================
'ELIMINAR VENTA
'=============================================================================
    Public Sub eliminarVenta(ByRef v As venta, ByRef Lista As Grid)
        Call eliminarVentaCabeza(v)
        Call eliminarVentaDetalle(v)
        Call eliminarMovimientosDetalle(v)
        Call eliminarPagos(v.cabeza.tipo, v.cabeza.numero)
        Call eliminarDocManual(v.cabeza.tipo, v.cabeza.numero)
        If v.cabeza.tipo = "NV" Then
            Call eliminarNotasFactura(v)
        End If
    End Sub
    
    Private Sub eliminarVentaCabeza(ByRef v As venta)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        condicion = "local = '" & v.cabeza.loc & "' AND tipo = '" & v.cabeza.tipo & "' AND numero = '" & v.cabeza.numero & "'"
        op = 4
        campos(0, 2) = "sv_documento_cabeza"
        sql.datos = campos
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
    End Sub
    
    Private Sub eliminarVentaDetalle(ByRef v As venta)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        condicion = "local = '" & v.cabeza.loc & "' AND tipo = '" & v.cabeza.tipo & "' AND numero = '" & v.cabeza.numero & "'"
        op = 4
        campos(0, 2) = "sv_documento_detalle"
        sql.datos = campos
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
    End Sub
    
    Private Sub eliminarMovimientosDetalle(ByRef v As venta)
'        Dim condicion As String
'        Dim op As Integer
'        Set sql = New CSQLUtil
'        condicion = "tipo = '" & v.cabeza.tipo & "' AND numero = '" & v.cabeza.numero & "'"
'        op = 4
'        campos(0, 2) = "l_movimientos_detalle_" & empresaActiva
'        sql.datos = campos
'        Set sql.conexion = gestionRubro
'        Call sql.SQLUTIL(op, condicion)
    End Sub
    
    Private Sub eliminarCobranza(ByRef v As venta)
'        Dim condicion As String
'        Dim op As Integer
'        Set sql = New CSQLUtil
'        condicion = "tipo = '" & v.cabeza.tipo & "' AND numero = '" & v.cabeza.numero & "'"
'        op = 4
'        campos(0, 2) = "l_movimientos_detalle_" & empresaActiva
'        sql.datos = campos
'        Set sql.conexion = gestionRubro
'        Call sql.SQLUTIL(op, condicion)
    End Sub
    
    Private Sub eliminarNotasFactura(ByRef v As venta)
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        condicion = "local = '" & v.cabeza.loc & "' AND numero = '" & v.cabeza.numero & "'"
        op = 4
        campos(0, 2) = "sv_documento_notas"
        sql.datos = campos
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR VENTA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaCabeza(ByRef vc As ventaCabeza, ByRef sql As CSQLUtil)
        vc.loc = sql.datos(0, 3)
        vc.tipo = sql.datos(1, 3)
        vc.numero = sql.datos(2, 3)
        vc.fecha = sql.datos(3, 3)
        vc.plazo = sql.datos(4, 3)
        vc.vencimiento = sql.datos(5, 3)
        vc.rut = sql.datos(6, 3)
        vc.sucursal = sql.datos(7, 3)
        vc.cajera = sql.datos(8, 3)
        vc.notapedido = sql.datos(9, 3)
        vc.notaventas = sql.datos(10, 3)
        vc.ordencompra = sql.datos(11, 3)
        vc.subtotal = sql.datos(12, 3)
        vc.neto = sql.datos(13, 3)
        vc.iva = sql.datos(14, 3)
        vc.impuestoharina = sql.datos(15, 3)
        vc.impuestoila = sql.datos(16, 3)
        vc.impuestoespecifico = sql.datos(17, 3)
        vc.exento = sql.datos(18, 3)
        vc.retencionparcial = sql.datos(19, 3)
        vc.retenciontotal = sql.datos(20, 3)
        vc.total = sql.datos(21, 3)
        vc.abono = sql.datos(22, 3)
        vc.descuento = sql.datos(23, 3)
        vc.contabilizado = sql.datos(24, 3)
        vc.PAGADO = sql.datos(25, 3)
        vc.comision = sql.datos(26, 3)
        vc.fechapagocomision = sql.datos(27, 3)
        vc.nula = sql.datos(28, 3)
        vc.boletadesde = sql.datos(29, 3)
        vc.boletahasta = sql.datos(30, 3)
        vc.descuento2 = sql.datos(31, 3)
        
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designaCabeza(ByRef v As venta, ByRef sql As CSQLUtil)
        Dim cad As String
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "fecha"
        campos(4, 0) = "plazo"
        campos(5, 0) = "vencimiento"
        campos(6, 0) = "rut"
        campos(7, 0) = "sucursal"
        campos(8, 0) = "cajera"
        campos(9, 0) = "notapedido"
        campos(10, 0) = "notaventa"
        campos(11, 0) = "ordencompra"
        campos(12, 0) = "subtotal"
        campos(13, 0) = "neto"
        campos(14, 0) = "iva"
        campos(15, 0) = "impuestoharina"
        campos(16, 0) = "impuestoila"
        campos(17, 0) = "impuestoespecifico"
        campos(18, 0) = "exento"
        campos(19, 0) = "retencionparcial"
        campos(20, 0) = "retenciontotal"
        campos(21, 0) = "total"
        campos(22, 0) = "abono"
        campos(23, 0) = "descuento"
        campos(24, 0) = "contabilizado"
        campos(25, 0) = "pagado"
        campos(26, 0) = "comision"
        campos(27, 0) = "fechapagocomision"
        campos(28, 0) = "nula"
        campos(29, 0) = "boletadesde"
        campos(30, 0) = "boletahasta"
        campos(31, 0) = "vendedor"
        campos(32, 0) = "descuento2"
        campos(33, 0) = ""
        
        campos(0, 1) = v.cabeza.loc
        campos(1, 1) = v.cabeza.tipo
        campos(2, 1) = v.cabeza.numero
        campos(3, 1) = v.cabeza.fecha
        campos(4, 1) = v.cabeza.plazo
        campos(5, 1) = v.cabeza.vencimiento
        campos(6, 1) = v.cabeza.rut
        campos(7, 1) = v.cabeza.sucursal
        campos(8, 1) = v.cabeza.cajera
        campos(9, 1) = v.cabeza.notapedido
        campos(10, 1) = v.cabeza.notaventas
        campos(11, 1) = v.cabeza.ordencompra
        campos(12, 1) = v.cabeza.subtotal
        campos(13, 1) = v.cabeza.neto
        campos(14, 1) = v.cabeza.iva
        campos(15, 1) = v.cabeza.impuestoharina
        campos(16, 1) = v.cabeza.impuestoila
        campos(17, 1) = v.cabeza.impuestoespecifico
        campos(18, 1) = v.cabeza.exento
        campos(19, 1) = v.cabeza.retencionparcial
        campos(20, 1) = v.cabeza.retenciontotal
        campos(21, 1) = v.cabeza.total
        campos(22, 1) = v.cabeza.abono
        campos(23, 1) = v.cabeza.descuento
        campos(24, 1) = v.cabeza.contabilizado
        campos(25, 1) = v.cabeza.PAGADO
        campos(26, 1) = v.cabeza.comision
        campos(27, 1) = v.cabeza.fechapagocomision
        campos(28, 1) = v.cabeza.nula
        campos(29, 1) = v.cabeza.boletadesde
        campos(30, 1) = v.cabeza.boletahasta
        campos(31, 1) = v.detalle.vendedor
        campos(32, 1) = v.detalle.descuento2
        
        campos(0, 2) = "sv_documento_cabeza"
        sql.datos = campos
    End Sub
    
    Private Sub designaDetalle(ByRef v As venta, ByRef sql As CSQLUtil, ByVal descripcion As String)
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "linea"
        campos(4, 0) = "fecha"
        campos(5, 0) = "rut"
        campos(6, 0) = "sucursal"
        campos(7, 0) = "codigo"
        campos(8, 0) = "descripcion"
        campos(9, 0) = "cantidad"
        campos(10, 0) = "unidades"
        campos(11, 0) = "precio"
        campos(12, 0) = "descuento"
        campos(13, 0) = "total"
        campos(14, 0) = "vendedor"
        campos(15, 0) = "pcosto"
        campos(16, 0) = "bodega"
        campos(17, 0) = "vencimiento"
        campos(18, 0) = "numerofactura"
        campos(19, 0) = "descuento2"
        campos(20, 0) = ""
    
        campos(0, 1) = v.detalle.loc
        campos(1, 1) = v.detalle.tipo
        campos(2, 1) = v.detalle.numero
        campos(3, 1) = v.detalle.linea
        campos(4, 1) = v.detalle.fecha
        campos(5, 1) = v.detalle.rut
        campos(6, 1) = v.detalle.sucursal
        campos(7, 1) = v.detalle.codigo
        campos(8, 1) = v.detalle.descripcion
        campos(9, 1) = v.detalle.cantidad
        campos(10, 1) = v.detalle.unidades
        campos(11, 1) = v.detalle.precio
        campos(12, 1) = v.detalle.descuento
        campos(13, 1) = v.detalle.total
        campos(14, 1) = v.detalle.vendedor
        campos(15, 1) = v.detalle.pcosto
        campos(16, 1) = v.detalle.bodega
        campos(17, 1) = v.detalle.vencimiento
        campos(18, 1) = v.detalle.numerofactura
        campos(19, 1) = v.detalle.descuento2
        
        
        campos(0, 2) = "sv_documento_detalle"
        sql.datos = campos
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================

    Public Sub anularDocumento(ByVal tipo As String, ByVal numero As String, Lista As Grid, vd As ventaDetalle)
        Dim condicion As String
        Dim op As Integer
        Dim i As Long
        Set sql = New CSQLUtil
        
        campos(0, 0) = "rut"
        campos(1, 0) = "subtotal"
        campos(2, 0) = "neto"
        campos(3, 0) = "iva"
        campos(4, 0) = "impuestoharina"
        campos(5, 0) = "total"
        campos(6, 0) = "abono"
        campos(7, 0) = "descuento"
        campos(8, 0) = "nula"
        campos(9, 0) = ""
        
        campos(0, 1) = "0888888888"
        campos(1, 1) = "0"
        campos(2, 1) = "0"
        campos(3, 1) = "0"
        campos(4, 1) = "0"
        campos(5, 1) = "0"
        campos(6, 1) = "0"
        campos(7, 1) = "0"
        campos(8, 1) = "S"
        campos(9, 1) = ""
        
        campos(0, 2) = "sv_documento_cabeza"
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & tipo & "' AND numero = '" & numero & "'"
        op = 3
        sql.datos = campos
        Set sql.conexion = ventasRubro
        Call sql.SQLUTIL(op, condicion)
        op = sql.estado
        
        For i = 1 To Lista.Rows - 1
            If Lista.Cell(i, 1).text <> "" Then
                Call desactualiza_stock("+", Lista.Cell(i, 1).text, "S", "N", bodega, Format(vd.fecha, "yyyy"), Lista.Cell(i, 3).text, Lista.Cell(i, 4).text, vd.fecha, vd.rut)
            End If
        Next i
    End Sub

Public Function verificarCupoCliente(ByVal rut As String, ByVal sucursal As String) As Boolean
    Call actualizarDatosCliente(rut, sucursal)
    cupo = CDbl(leerCupoClienteSucursal(rut, sucursal))
    If cupo > 0 Then
        verificarCupoCliente = True
    Else
        verificarCupoCliente = False
    End If
End Function

Public Sub enviarInformacion(ByVal rut As String, ByVal sucursal As String, ByVal tipo As String, ByVal numero As String, ByVal monto As String, ByVal glosa As String)
'    Dim cSql As rdoQuery
'    Set cSql = New rdoQuery
'    Set cSql.ActiveConnection = ventas
'    cSql.sql = "INSERT INTO sv_maestroclientes_autorizacion (local, rut, sucursal, tipodoc, numerodoc, monto, glosa, fecha, usuario) "
'    cSql.sql = cSql.sql & "VALUES('" & empresaActiva & "', '" & rut & "', '" & sucursal & "', '" & tipo & "', '" & numero & "', '" & monto & "', '" & glosa & "', '" & fechasistema & "', '" & usuarioSistema & "') "
'    cSql.sql = cSql.sql & "ON DUPLICATE KEY UPDATE monto = '" & monto & "', glosa = '" & glosa & "' "
'    cSql.Execute
'    cSql.Close
'    Set cSql = Nothing
End Sub








