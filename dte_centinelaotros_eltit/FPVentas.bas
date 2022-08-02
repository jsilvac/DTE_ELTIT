Attribute VB_Name = "FPVentas"
Option Explicit
    Public autorizado As Boolean
    Public cupo As Double
    Private campos(50, 3) As String
    Private Type ventaCabeza
        Loc As String
        Tipo As String
        Numero As String
        Fecha As String
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
        impuestoila As String
        impuestoespecifico As String
        impuestoIla13 As String
        impuestoIla15 As String
        impuestoIla27 As String
        impuestoCarne As String
        impuestoHarina As String
        exento As String
        retencionparcial As String
        retenciontotal As String
        total As String
        abono As String
        Descuento As String
        contabilizado As String
        PAGADO As String
        comision As String
        fechapagocomision As String
        nula As String
        boletadesde As String
        boletahasta As String
        descuento2 As String
        condicionesdepago As String
        transporte As String
        revisado As String
        bultos As String
        abono2 As String
        Caja As String
        foliosii As String
        vendedor As String
        
    End Type
        
    Private Type ventaDetalle
        Loc As String
        Tipo As String
        Numero As String
        linea As String
        Fecha As String
        rut As String
        sucursal As String
        CODIGO As String
        descripcion As String
        cantidad As String
        unidades As String
        precio As String
        Descuento As String
        total As String
        vendedor As String
        pcosto As String
        bodega As String
        vencimiento As String
        NUMEROFACTURA As String
        descuento2 As String
        GLOSA As String
        Caja As String
        
    End Type
        
    Public Type venta
        cabeza As ventaCabeza
        detalle As ventaDetalle
    End Type
    
'=============================================================================
'LEER VENTA
'=============================================================================
    Public Function leerVenta(ByRef v As venta, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String, Optional data As Variant, Optional lista As Variant, Optional Caja As String, Optional fechadocumento As String) As Boolean
        leerVenta = True
        If leerVentaCabeza(v.cabeza, codigo1, codigo2, operador, Caja, fechadocumento) = False Then
            
            leerVenta = False
            
        Else
            If Not IsMissing(data) Then
                Call leerVentaDetalle(data, codigo1, v.cabeza.Numero, "=", lista, v.cabeza.Caja, Format(v.cabeza.Fecha, "yyyy-mm-dd"))
            End If
        End If
    
    End Function
    
    Private Function leerVentaCabeza(ByRef vc As ventaCabeza, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String, Optional Caja As String, Optional Fecha As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "fecha"
        campos(4, 0) = "plazo"
        campos(5, 0) = "IFNULL(vencimiento,'1900-01-01')"
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
        campos(27, 0) = "IFNULL(fechapagocomision,'1900-01-01')"
        campos(28, 0) = "nula"
        campos(29, 0) = "boletadesde"
        campos(30, 0) = "boletahasta"
        campos(31, 0) = "descuento2"
        campos(32, 0) = "transporte"
        campos(33, 0) = "condicionesdepago"
        campos(34, 0) = "revisado"
        campos(35, 0) = "bultos"
        campos(36, 0) = "foliosii"
        campos(37, 0) = "caja"
        campos(38, 0) = "impuestoilarefrescos"
        campos(39, 0) = "impuestoilavinos"
        campos(40, 0) = "impuestoilalicores"
        campos(41, 0) = "impuestocarne"
        campos(42, 0) = "vendedor"
        campos(43, 0) = "horaventas"
        campos(44, 0) = ""
        campos(0, 2) = "sv_documento_cabeza_" + empresaActiva
        condicion = "local = '" & empresaActiva & "' and caja= '" & Caja & "' AND fecha='" & Fecha & "' and  tipo = '" & codigo1 & "' AND foliosii " & operador & " '" & codigo2 & "' "
       
        If operador = "<" Then
             condicion = condicion & "ORDER BY numero DESC"
        Else
             condicion = condicion & "ORDER BY numero ASC"
        End If
        
        op = 5
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            PVentas.lblfoliointerno = sql.response(2, 3)
            PVentas.lblfoliofiscal = sql.response(36, 3)
            PVentas.DatoHora = sql.response(43, 3)
            leerVentaCabeza = True
            Call asignaCabeza(vc, sql)
            
            
        
        Else
            leerVentaCabeza = False
        End If
    End Function
    
    Private Sub leerVentaDetalle(ByRef data As Variant, ByVal codigo1 As String, ByVal codigo2 As String, ByVal operador As String, ByRef lista As Variant, ByVal Caja As String, ByRef Fecha As String)
        Dim tabla As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = ventasRubro
     csql.sql = "SELECT CONCAT(codigo, '" & vbTab & "', descripcion, '" & vbTab & "', IF(numerofactura<>'', numerofactura, cantidad), '" & vbTab & "', unidades, '" & vbTab & "', precio, '" & vbTab & "', descuento, '" & vbTab & "', total) AS item "
     csql.sql = csql.sql & "FROM sv_documento_detalle_" + empresaActiva + " "
     csql.sql = csql.sql & "WHERE local = '" & empresaActiva & "' AND tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' and caja='" + Caja + "' and fecha='" + Fecha + "' ORDER BY linea ASC"
     csql.Execute

        'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        lista.Rows = 1
        lista.AutoRedraw = False
        If csql.RowsAffected > 0 Then
'            data.Recordset.MoveFirst
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                lista.AddItem Replace(resultados("item"), ".", ","), True
            resultados.MoveNext
            Wend
            Set csql = Nothing
            
            csql.Close
            Set resultados = Nothing
            
        End If
        lista.AutoRedraw = True
        lista.Refresh
    End Sub
'=============================================================================
'LEER VENTA
'=============================================================================

'=============================================================================
'GRABAR VENTA
'=============================================================================
    Public Sub grabarVenta(ByRef v As venta, ByVal modifica As Boolean, ByRef lista As Grid)
        Call grabarVentaCabeza(v, modifica)
        Call grabarVentaDetalle(lista, v, modifica)
'        Call modificafoliocaja(v.cabeza.tipo, v.cabeza.caja)
        If v.cabeza.Tipo = "FV" Or v.cabeza.Tipo = "ZE" Or v.cabeza.Tipo = "BV" Or v.cabeza.Tipo = "NV" Then
        Call grabarVentaMovimientos(v, modifica)
        End If

    End Sub
    
    Private Sub grabarVentaCabeza(ByRef v As venta, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designaCabeza(v, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND local = '" & empresaActiva & "' AND tipo = '" & v.cabeza.Tipo & "' AND numero = '" & v.cabeza.Numero & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = PVentas.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
        
        op = sql.Status
    End Sub
    
    Private Sub grabarVentaDetalle(ByRef lista As Grid, ByRef v As venta, ByVal modifica As Boolean)
        Dim descripcion As String
        
        Dim op As Integer
        Dim i As Long
        Dim lin As String
        Set sql = New sqlventas.sqlventa
        
        condicion = ""
        If modifica = True Then
            Call eliminarVentaDetalle(v)
        End If
        op = 2
        
        For i = 1 To lista.Rows - 1
            lin = Str(i)
            lin = Mid(lin, 2, Len(lin))
            If lista.Cell(i, 1).text <> "" And lista.Cell(i, 2).text <> "" And lista.Cell(i, 3).text <> "" And lista.Cell(i, 4).text <> "" And lista.Cell(i, 5).text <> "" Then
                v.detalle.linea = String(3 - Len(lin), "0") & lin
                v.detalle.CODIGO = lista.Cell(i, 1).text
                v.detalle.descripcion = lista.Cell(i, 2).text
                
                v.detalle.unidades = Replace(Format(lista.Cell(i, 3).text, "########0.00"), ",", ".")
                v.detalle.precio = Replace(Format(lista.Cell(i, 5).text, "########0.00"), ",", ".")
                v.detalle.Descuento = Replace(Format(lista.Cell(i, 6).text, "########0.00"), ",", ".")
                v.detalle.descuento2 = Replace(Format(lista.Cell(i, 8).text, "########0.00"), ",", ".")
                v.detalle.total = Replace(Format(lista.Cell(i, 7).text, "########0.00"), ",", ".")
                v.detalle.pcosto = leerCostoProducto(lista.Cell(i, 1).text)
                v.detalle.GLOSA = lista.Cell(i, 9).text
                If Val(lista.Cell(i, 1).text) = 100 Then
                    v.detalle.cantidad = "1"
                    v.detalle.NUMEROFACTURA = lista.Cell(i, 3).text
                    If v.cabeza.Tipo = "NV" Then
                        Call grabarNotasFactura(v)
                        If v.detalle.CODIGO = "0000000000100" Then
                            Call modificarAbono(v)
                        End If
                    End If
                Else
                    v.detalle.cantidad = Replace(Format(lista.Cell(i, 3).text, "########0.00"), ",", ".")
                    v.detalle.NUMEROFACTURA = ""
                Call actualiza_stock("-", v.detalle.CODIGO, "S", "N", "00", Mid(v.cabeza.Fecha, 1, 4), lista.Cell(i, 3).text, lista.Cell(i, 5).text, v.cabeza.Fecha, v.cabeza.rut, empresaActiva)
                End If
                
                Call designaDetalle(v, sql, descripcion)
                
                Set sql.conexion = ventasRubro
                sql.audit = True: sql.programaactivo = PVentas.Caption
                Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
                Call sql.sqlventas(op, condicion)
            End If
        Next i
    End Sub
    
    Private Sub grabarVentaMovimientos(ByRef v As venta, ByVal modifica As Boolean)
        Dim csql As New rdoQuery
        Set csql.ActiveConnection = gestionRubro
        If modifica = False Then
            csql.sql = "INSERT INTO l_movimientos_detalle_" & rubro & " "
            csql.sql = csql.sql & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
            If PVentas.dato1.text = "FV" Then csql.sql = csql.sql & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.cantidad, ROUND(dd.precio*1.19), ROUND(dd.total*1.19), dd.pcosto, dd.bodega, dd.bodega, 1 "
            If PVentas.dato1.text <> "FV" Then csql.sql = csql.sql & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.cantidad, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, 1 "
            
            csql.sql = csql.sql & "FROM " & baseVentas & empresaActiva & ".sv_documento_detalle_" + empresaActiva + " AS dd "
         'cSql.sql = cSql.sql & "WHERE dd.tipo='FV' OR DD.TIPO='BV' OR DD.TIPO='ZE'"
            csql.sql = csql.sql & "WHERE dd.local = '" & empresaActiva & "' "
            csql.sql = csql.sql & "AND dd.tipo = '" & v.detalle.Tipo & "' AND dd.numero = '" & v.detalle.Numero & "'"
            csql.Execute
        Else
            csql.sql = "DELETE FROM l_movimientos_detalle_" & rubro & " "
            csql.sql = csql.sql & "WHERE tipo = '" & v.detalle.Tipo & "' AND numero = '" & v.detalle.Numero & "'"
            csql.Execute
            modifica = False
            Call grabarVentaMovimientos(v, modifica)
        End If
        csql.Close
        Set csql = Nothing
    End Sub
    
    Private Sub grabarNotasFactura(ByRef v As venta)
        Dim csql As New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "INSERT INTO sv_documento_notas "
        csql.sql = csql.sql & "(local, numero, localfactura, numerofactura, monto) "
        csql.sql = csql.sql & "VALUES('" & empresaActiva & "', '" & v.cabeza.Numero & "', '" & empresaActiva & "', '" & v.detalle.NUMEROFACTURA & "', '" & v.detalle.total & "') "
        csql.sql = csql.sql & "ON DUPLICATE KEY UPDATE local = '" & empresaActiva & "', numero = '" & v.cabeza.Numero & "', localfactura = '" & empresaActiva & "', numerofactura = '" & v.detalle.NUMEROFACTURA & "', monto = '" & v.detalle.total & "' "
        csql.Execute
        csql.Close
        Set csql = Nothing
    End Sub
    
    Private Sub modificarAbono(ByRef v As venta)
        Dim csql As rdoQuery
        Dim abono As Double
        Dim abonoIva As Double
        Dim abonoIha As Double
        abono = -1 * CDbl(Replace(v.detalle.total, ".", ","))
        abonoIva = abono * iva / 100
        abonoIha = abono * iha / 100
        abono = Round(abono + abonoIva + abonoIha, 0)
        
        'DOCUMENTO CABEZA(FACTURA)
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "UPDATE sv_documento_cabeza_" + empresaActiva + " "
        csql.sql = csql.sql & "SET abono = abono + " & abono & " "
        csql.sql = csql.sql & "WHERE local = '" & v.cabeza.Loc & "' AND tipo = 'FV' AND numero = '" & v.detalle.NUMEROFACTURA & "' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventasRubro)
        csql.Close
        Set csql = Nothing
        
        'DOCUMENTOS COBRANZA(FACTURA)
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "UPDATE sv_documentos_cobranza_" & empresaActiva & " "
        csql.sql = csql.sql & "SET abono = abono + " & abono & " "
        csql.sql = csql.sql & "WHERE local = '" & v.cabeza.Loc & "' AND tipo = 'FV' AND numero = '" & v.detalle.NUMEROFACTURA & "' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventasRubro)
        csql.Close
        Set csql = Nothing
        
        abono = abono * -1
        'DOCUMENTO CABEZA(NOTA CREDITO)
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "UPDATE sv_documento_cabeza_" + empresaActiva + " "
        csql.sql = csql.sql & "SET abono = abono + " & abono & " "
        csql.sql = csql.sql & "WHERE local = '" & v.cabeza.Loc & "' AND tipo = 'NV' AND numero = '" & v.cabeza.Numero & "' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventasRubro)
        csql.Close
        Set csql = Nothing
        
        'DOCUMENTO CABEZA(NOTA CREDITO)
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "UPDATE sv_documentos_cobranza_" & empresaActiva & " "
        csql.sql = csql.sql & "SET abono = abono + " & abono & " "
        csql.sql = csql.sql & "WHERE local = '" & v.cabeza.Loc & "' AND tipo = 'NV' AND numero = '" & v.cabeza.Numero & "' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventasRubro)
        csql.Close
        Set csql = Nothing
    End Sub
'=============================================================================
'GRABAR VENTA
'=============================================================================

'=============================================================================
'ELIMINAR VENTA
'=============================================================================
    Public Sub eliminarVenta(ByRef v As venta, ByRef lista As Grid)
        Call eliminarVentaCabeza(v)
        
        Call eliminarVentaDetalle(v)
        Call eliminarMovimientosDetalle(v)
        Call eliminarPagos(v.cabeza.Tipo, v.cabeza.Numero, Format(v.cabeza.Fecha, "yyyy-mm-dd"), v.cabeza.Caja)
        Call eliminarDocManual(v.cabeza.Tipo, v.cabeza.Numero)
        If v.cabeza.Tipo = "NV" Then
            Call eliminarNotasFactura(v)
        End If
    End Sub
    
    Private Sub eliminarVentaCabeza(ByRef v As venta)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & v.cabeza.Loc & "' AND tipo = '" & v.cabeza.Tipo & "' AND numero = '" & v.cabeza.Numero & "' and caja='" & v.cabeza.Caja & "' and fecha='" & v.cabeza.Fecha & "' "
        op = 4
        campos(0, 2) = "sv_documento_cabeza_" + empresaActiva
        sql.response = campos
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = PVentas.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Private Sub eliminarVentaDetalle(ByRef v As venta)
        
        Dim op As Integer
        Call rebajarstock(PVentas.detalle, v, False)
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & v.cabeza.Loc & "' AND tipo = '" & v.cabeza.Tipo & "' AND numero = '" & v.cabeza.Numero & "' and caja='" & v.cabeza.Caja & "' and fecha='" & v.cabeza.Fecha & "' "
        op = 4
        campos(0, 2) = "sv_documento_detalle_" + empresaActiva
        sql.response = campos
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = PVentas.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
    End Sub
    Private Sub rebajarstock(ByRef lista As Grid, ByRef v As venta, ByVal modifica As Boolean)
        Dim descripcion As String
        
        Dim op As Integer
        Dim i As Long
        Dim lin As String
        
        For i = 1 To lista.Rows - 1
            lin = Str(i)
            lin = Mid(lin, 2, Len(lin))
            If lista.Cell(i, 1).text <> "" And lista.Cell(i, 2).text <> "" And lista.Cell(i, 3).text <> "" And lista.Cell(i, 4).text <> "" And lista.Cell(i, 5).text <> "" Then
                v.detalle.linea = String(3 - Len(lin), "0") & lin
                v.detalle.CODIGO = lista.Cell(i, 1).text
                Call desactualiza_stock("-", v.detalle.CODIGO, "S", "N", "00", Mid(v.cabeza.Fecha, 1, 4), lista.Cell(i, 3).text, lista.Cell(i, 5).text, v.cabeza.Fecha, v.cabeza.rut, empresaActiva)
              
            End If
        Next i
    End Sub
    
    Private Sub eliminarMovimientosDetalle(ByRef v As venta)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "tipo = '" & v.cabeza.Tipo & "' AND numero = '" & v.cabeza.Numero & "'"
        op = 4
        campos(0, 2) = "l_movimientos_detalle_" & empresaActiva
        sql.response = campos
        Set sql.conexion = gestionRubro
        sql.audit = True: sql.programaactivo = PVentas.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Private Sub eliminarCobranza(ByRef v As venta)
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        condicion = "tipo = '" & v.cabeza.tipo & "' AND numero = '" & v.cabeza.numero & "'"
'        op = 4
'        campos(0, 2) = "l_movimientos_detalle_" & empresaActiva
'        sql.response = campos
'        Set sql.conexion = gestionRubro
'        Call sql.sqlventas(op, condicion)
    End Sub
    
    Private Sub eliminarNotasFactura(ByRef v As venta)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & v.cabeza.Loc & "' AND numero = '" & v.cabeza.Numero & "'"
        op = 4
        campos(0, 2) = "sv_documento_notas"
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR VENTA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaCabeza(ByRef vc As ventaCabeza, ByRef sql As sqlventas.sqlventa)
        vc.Loc = sql.response(0, 3)
        vc.Tipo = sql.response(1, 3)
        vc.Numero = sql.response(2, 3)
        vc.Fecha = sql.response(3, 3)
        vc.plazo = sql.response(4, 3)
        vc.vencimiento = sql.response(5, 3)
        vc.rut = sql.response(6, 3)
        vc.sucursal = sql.response(7, 3)
        vc.cajera = sql.response(8, 3)
        
        vc.notapedido = sql.response(9, 3)
        vc.notaventas = sql.response(10, 3)
        vc.ordencompra = sql.response(11, 3)
        vc.subtotal = sql.response(12, 3)
        vc.neto = sql.response(13, 3)
        vc.iva = sql.response(14, 3)
        vc.impuestoHarina = sql.response(15, 3)
        vc.impuestoila = sql.response(16, 3)
        vc.impuestoespecifico = sql.response(17, 3)
        vc.exento = sql.response(18, 3)
        vc.retencionparcial = sql.response(19, 3)
        vc.retenciontotal = sql.response(20, 3)
        vc.total = sql.response(21, 3)
        vc.abono = sql.response(22, 3)
        vc.Descuento = sql.response(23, 3)
        vc.contabilizado = sql.response(24, 3)
        vc.PAGADO = sql.response(25, 3)
        vc.comision = sql.response(26, 3)
        vc.fechapagocomision = sql.response(27, 3)
        vc.nula = sql.response(28, 3)
        vc.boletadesde = sql.response(29, 3)
        vc.boletahasta = sql.response(30, 3)
        vc.descuento2 = sql.response(31, 3)
        vc.transporte = sql.response(32, 3)
        vc.condicionesdepago = sql.response(33, 3)
        vc.revisado = sql.response(34, 3)
        vc.bultos = sql.response(35, 3)
        vc.Caja = sql.response(37, 3)
        vc.foliosii = sql.response(36, 3)
        vc.impuestoIla13 = sql.response(38, 3)
        vc.impuestoIla15 = sql.response(39, 3)
        vc.impuestoIla27 = sql.response(40, 3)
        vc.impuestoCarne = sql.response(41, 3)
        vc.vendedor = sql.response(42, 3)
        
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designaCabeza(ByRef v As venta, ByRef sql As sqlventas.sqlventa)
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
        campos(33, 0) = "condicionesdepago"
        campos(34, 0) = "transporte"
        campos(35, 0) = "revisado"
        campos(36, 0) = "bultos"
        campos(37, 0) = "abono2"
        campos(38, 0) = "impuestoilarefrescos"
        campos(39, 0) = "impuestoilavinos"
        campos(40, 0) = "impuestoilalicores"
        campos(41, 0) = "impuestocarne"
        campos(42, 0) = "foliosii"
        campos(43, 0) = "caja"
        campos(44, 0) = ""
        
        campos(0, 1) = v.cabeza.Loc
        campos(1, 1) = v.cabeza.Tipo
        campos(2, 1) = v.cabeza.Numero
        campos(3, 1) = v.cabeza.Fecha
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
        campos(15, 1) = v.cabeza.impuestoHarina
        campos(16, 1) = v.cabeza.impuestoila
        campos(17, 1) = v.cabeza.impuestoespecifico
        campos(18, 1) = v.cabeza.exento
        campos(19, 1) = v.cabeza.retencionparcial
        campos(20, 1) = v.cabeza.retenciontotal
        campos(21, 1) = v.cabeza.total
        campos(22, 1) = v.cabeza.abono
        campos(23, 1) = v.cabeza.Descuento
        campos(24, 1) = v.cabeza.contabilizado
        campos(25, 1) = v.cabeza.PAGADO
        campos(26, 1) = v.cabeza.comision
        campos(27, 1) = v.cabeza.fechapagocomision
        campos(28, 1) = v.cabeza.nula
        campos(29, 1) = v.cabeza.boletadesde
        campos(30, 1) = v.cabeza.boletahasta
        campos(31, 1) = v.detalle.vendedor
        campos(32, 1) = v.detalle.descuento2
        campos(33, 1) = v.cabeza.condicionesdepago
        campos(34, 1) = v.cabeza.transporte
        campos(35, 1) = v.cabeza.revisado
        campos(36, 1) = v.cabeza.bultos
        campos(37, 1) = v.cabeza.abono2
        campos(38, 1) = v.cabeza.impuestoIla13
        campos(39, 1) = v.cabeza.impuestoIla15
        campos(40, 1) = v.cabeza.impuestoIla27
        campos(41, 1) = v.cabeza.impuestoCarne
        campos(42, 1) = v.cabeza.foliosii
        campos(43, 1) = v.cabeza.Caja
        campos(0, 2) = "sv_documento_cabeza_" + empresaActiva
        sql.response = campos
    End Sub
    
    Private Sub designaDetalle(ByRef v As venta, ByRef sql As sqlventas.sqlventa, ByVal descripcion As String)
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
        campos(20, 0) = "glosa"
        campos(21, 0) = "caja"
        campos(22, 0) = ""
    
        campos(0, 1) = v.detalle.Loc
        campos(1, 1) = v.detalle.Tipo
        campos(2, 1) = v.detalle.Numero
        campos(3, 1) = v.detalle.linea
        campos(4, 1) = v.detalle.Fecha
        campos(5, 1) = v.detalle.rut
        campos(6, 1) = v.detalle.sucursal
        campos(7, 1) = v.detalle.CODIGO
        campos(8, 1) = v.detalle.descripcion
        campos(9, 1) = v.detalle.cantidad
        campos(10, 1) = v.detalle.cantidad
        campos(11, 1) = v.detalle.precio
        campos(12, 1) = v.detalle.Descuento
        campos(13, 1) = v.detalle.total
        campos(14, 1) = v.detalle.vendedor
        campos(15, 1) = v.detalle.pcosto
        campos(16, 1) = v.detalle.bodega
        campos(17, 1) = v.detalle.vencimiento
        campos(18, 1) = v.detalle.NUMEROFACTURA
        campos(19, 1) = v.detalle.descuento2
        campos(20, 1) = v.detalle.GLOSA
        campos(21, 1) = v.detalle.Caja
        
        campos(0, 2) = "sv_documento_detalle_" + empresaActiva
        sql.response = campos
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================

    Public Sub anularDocumento(ByVal Tipo As String, ByVal Numero As String, lista As Grid, vd As ventaDetalle)
        
        Dim op As Integer
        Dim i As Long
        Set sql = New sqlventas.sqlventa
        
        campos(0, 0) = "rut"
        campos(1, 0) = "subtotal"
        campos(2, 0) = "neto"
        campos(3, 0) = "iva"
        campos(4, 0) = "impuestoharina"
        campos(5, 0) = "total"
        campos(6, 0) = "abono"
        campos(7, 0) = "descuento"
        campos(8, 0) = "nula"
        campos(9, 0) = "impuestoharina"
        campos(10, 0) = "impuestocarne"
        campos(11, 0) = "impuestoilarefrescos"
        campos(12, 0) = "impuestoilavinos"
        campos(13, 0) = "impuestoilalicores"
        campos(14, 0) = ""
        
        
        campos(0, 1) = "0888888888"
        campos(1, 1) = "0"
        campos(2, 1) = "0"
        campos(3, 1) = "0"
        campos(4, 1) = "0"
        campos(5, 1) = "0"
        campos(6, 1) = "0"
        campos(7, 1) = "0"
        campos(8, 1) = "S"
        
        campos(9, 1) = "0"
        campos(10, 1) = "0"
        campos(11, 1) = "0"
        campos(12, 1) = "0"
        campos(13, 1) = "0"
        
        campos(0, 2) = "sv_documento_cabeza_" + empresaActiva
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & Tipo & "' AND foliosii = '" & Numero & "'"
        op = 3
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        op = sql.Status
        
'        For i = 1 To Lista.Rows - 1
'            If Lista.Cell(i, 1).text <> "" Then
'                Call desactualiza_stock("+", Lista.Cell(i, 1).text, "S", "N", bodega, Format(vd.fecha, "yyyy"), Lista.Cell(i, 3).text, Lista.Cell(i, 4).text, vd.fecha, vd.rut)
'            End If
'        Next i
    End Sub

Public Function verificarCupoCliente(ByVal rut As String, ByVal sucursal As String) As Boolean
    Call actualizarDatosCliente(rut, sucursal)
    Rem cupo = CDbl(leerCupoClienteSucursal(rut, sucursal))
    If cupo > 0 Then
        verificarCupoCliente = True
    Else
        verificarCupoCliente = False
    End If
End Function

Public Sub enviarInformacion(ByVal rut As String, ByVal sucursal As String, ByVal Tipo As String, ByVal Numero As String, ByVal MONTO As String, ByVal GLOSA As String)
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

Public Sub modificafoliocaja(Tipo, Caja)
'
'        Dim op As Integer
'        Dim i As Long
'        Set sql =new sqlventas.sqlventa
'        If tipo = "VB" Then
'        campos(0, 0) = "folioboletas"
'        End If
'        If tipo = "FV" Then
'        campos(0, 0) = "foliofacturas"
'        End If
'
'        campos(0, 1) = leerfoliocaja(tipo, caja) + 1
'        campos(0, 2) = "sv_maestrodecajas"
'        condicion = "local = '" & empresaActiva & "' AND numero = '" & caja & "' "
'        op = 3
'        sql.response = campos
'        Set sql.conexion = ventas
'        Call sql.sqlventas(op, condicion)
End Sub






