Attribute VB_Name = "FPCotizacionesNotasPedido"
Option Explicit

Public Sub imprimeCotizacionNotaPedido(ByVal numerofactura As String, ByRef documento As Grid, ByRef rollo As Adodc, ByVal TIPO As String)
    Dim ss As String
    Dim i As Integer
    Dim k As Integer
    Dim cad As String
    Dim totalprod As String
    Dim descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim lineas As Integer
    Dim fecha As String
    Dim VENCIMIENTO As String
    Dim vendedor As String
    Dim notapedido As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim fono As String
    Dim O As Integer
    Dim dia As String
    Dim mes As String
    Dim ano As String
    Dim nvalor As Long
    Dim codigo As String
    Dim tiposDePago As String
    Dim resultados As rdoResultset
    Dim resultados1 As rdoResultset
    Dim resultados2 As rdoResultset
    Dim cSql As New rdoQuery
    Dim cSql1 As New rdoQuery
    Dim cSql2 As New rdoQuery
    
    
    Set cSql.ActiveConnection = ventasRubro
    cSql.sql = "SELECT d.codigo, CONCAT(d.cantidad, '" & vbTab & "', d.descripcion, '" & vbTab & vbTab & "', d.cantidad * mpf.contenido, '" & vbTab & "', d.precio, '" & vbTab & "', d.total) AS item, d.total AS totalpro, c.neto, c.iva, c.impuestoharina, c.total, DATE_FORMAT(c.fecha,'%d-%m-%Y') AS fecha, c.descuento, cl.rut, cl.nombre, cl.direccion, cl.giro, cl.ciudad, cl.comuna, cl.fono1, c.cajera, DATE_FORMAT(c.vencimiento,'%d-%m-%Y') AS vencimiento, IF(notapedido = '000000000000','',notapedido) AS notapedido "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS c INNER JOIN sv_documento_detalle_" + empresaActiva + " AS d ON c.tipo = d.tipo AND c.numero = d.numero INNER JOIN " & baseVentas & ".sv_maestroclientes AS cl ON c.rut = cl.rut INNER JOIN " & baseDatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON d.codigo = mpf.codigobarra "
    cSql.sql = cSql.sql & "WHERE c.tipo = '" & TIPO & "' AND c.numero = '" & numerofactura & "' "
    cSql.sql = cSql.sql & "ORDER BY d.linea"
    cSql.Execute
    'Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
    
    If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
        documento.Rows = 2
        documento.Cols = 8
        
        documento.DefaultFont.Name = "Arial"
        documento.DefaultFont.Size = 8
        documento.DefaultFont.Bold = False
        
        documento.Column(0).Width = 0
        documento.Column(1).Width = 45
        documento.Column(2).Width = 100
        documento.Column(3).Width = 210
        documento.Column(4).Width = 105
        documento.Column(5).Width = 70
        documento.Column(6).Width = 90
        documento.Column(7).Width = 125
        
        documento.Column(1).Alignment = cellRightCenter
        documento.Column(2).Alignment = cellCenterCenter
        documento.Column(3).Alignment = cellLeftCenter '/**/
        documento.Column(4).Alignment = cellLeftCenter '/**/
        documento.Column(5).Alignment = cellRightCenter
        documento.Column(6).Alignment = cellRightCenter
        documento.Column(7).Alignment = cellRightCenter
        
        documento.DefaultRowHeight = 13
        
        documento.PageSetup.PrintGridlines = False
        documento.AutoRedraw = False
        
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        documento.AddItem ""
        
'        rollo.Recordset.MoveFirst
        
        fecha = resultados("fecha")
        nombre = "       " & resultados("nombre")
        rut = "       " & resultados("rut")
        direccion = "       " & resultados("direccion")
        ciudad = "       " & resultados("ciudad")
        comuna = resultados("comuna")
        giro = resultados("giro")
        fono = resultados("fono1")
        VENCIMIENTO = resultados("vencimiento")
        vendedor = resultados("cajera")
        notapedido = resultados("notapedido")
        neto = resultados("neto")
        piva = resultados("iva")
        piha = resultados("impuestoharina")
        total = resultados("total")
        descuento = resultados("descuento")
        'descuento = "0"
        While Not resultados.EOF
            codigo = Right(resultados("codigo"), 4)
            'descuento = Replace(Str(CDbl(descuento) + CDbl(rollo.Recordset.Fields("totalpro")) * CDbl(rollo.Recordset.Fields("descuento")) / 100), ".", ",")
            'descuento = Mid(descuento, 2, Len(descuento))
            documento.AddItem codigo & vbTab & resultados("item"), True
            documento.Range(documento.Rows - 1, 3, documento.Rows - 1, 4).Merge
            resultados.MoveNext
        Wend
        documento.Cell(5, 4).Alignment = cellRightCenter
        documento.Cell(5, 6).text = numerofactura
        
        'TIPOS DE PAGO
        Set cSql1.ActiveConnection = ventasRubro
        cSql1.sql = "SELECT dp.tipopago, SUM(dp.monto) AS monto "
        cSql1.sql = cSql1.sql & "FROM sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero "
        cSql1.sql = cSql1.sql & "WHERE dp.tipo = '" & TIPO & "' AND dp.numero = '" & numerofactura & "' "
        cSql1.sql = cSql1.sql & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
        cSql1.Execute
'        Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
        
        If cSql1.RowsAffected > 0 Then
          Set resultados1 = cSql1.OpenResultset
            tiposDePago = ""
            While Not resultados1.EOF
                codigo = resultados1("tipopago")
                Select Case codigo
                    Case "1"    'EFECTIVO
                        tiposDePago = tiposDePago & "EFECTIVO " & Format(resultados1("monto"), "$ ###,###,##0") & " / "
                    Case "2"    'CHEQUE
                        tiposDePago = tiposDePago & "CHEQUE " & Format(resultados1("monto"), "$ ###,###,##0") & " / "
                    Case "6"    'CREDITO DIRECTO
                        tiposDePago = tiposDePago & "CREDITO DIRECTO " & Format(resultados1("monto"), "$ ###,###,##0") & " / "
                End Select
                resultados1.MoveNext
            Wend
        End If
        Set cSql1 = Nothing
        cSql1.Close
        Set resultados1 = Nothing
        
        documento.Range(4, 2, 4, 3).Merge
        documento.Range(4, 2, 4, 3).Alignment = cellCenterCenter
        documento.Cell(4, 2).text = leerNombreEmpresa(empresaActiva)
        
        documento.RowHeight(6) = 15
        documento.RowHeight(7) = 10
        documento.RowHeight(8) = 15
        documento.RowHeight(9) = 15
        documento.RowHeight(10) = 15
        documento.RowHeight(11) = 15
        
        documento.RowHeight(15) = 5
        
        'FECHA
        documento.Range(6, 1, 6, 2).Merge
        documento.Range(6, 1, 6, 2).Alignment = cellCenterCenter
        documento.Cell(6, 1).text = fecha
        
        'SEÑORES
        documento.Range(8, 2, 8, 3).Merge
        documento.Range(8, 2, 8, 3).Alignment = cellLeftCenter
        documento.Cell(8, 2).text = nombre
        'RUT
        documento.Range(9, 2, 9, 3).Merge
        documento.Range(9, 2, 9, 3).Alignment = cellLeftCenter
        documento.Cell(9, 2).text = rut
        'DIRECCION
        documento.Range(10, 2, 10, 3).Merge
        documento.Range(10, 2, 10, 3).Alignment = cellLeftCenter
        documento.Cell(10, 2).text = direccion
        'CIUDAD
        documento.Range(11, 2, 11, 3).Merge
        documento.Range(11, 2, 11, 3).Alignment = cellLeftCenter
        documento.Cell(11, 2).text = ciudad
        
        'GIRO
        documento.Range(9, 5, 9, 6).Merge
        documento.Range(9, 5, 9, 6).Alignment = cellLeftCenter
        documento.Cell(9, 5).text = giro
        'COMUNA
        documento.Range(11, 5, 11, 6).Merge
        documento.Range(11, 5, 11, 6).Alignment = cellLeftCenter
        documento.Cell(11, 5).text = comuna
        'FONO
        documento.Cell(10, 6).Alignment = cellRightCenter
        documento.Cell(10, 6).text = fono
        
        'VENCIMIENTO
        documento.Cell(9, 7).Alignment = cellCenterCenter
        documento.Cell(9, 7).text = VENCIMIENTO
        'VENDEDOR
        documento.Cell(11, 7).Alignment = cellRightCenter
        documento.Cell(11, 7).text = vendedor
        'CONDICIONES DE PAGO
        documento.Range(13, 3, 13, 6).Merge
        documento.Range(13, 3, 13, 6).Alignment = cellLeftCenter
        documento.Cell(13, 3).text = tiposDePago
        'NOTAPEDIDO
        documento.Cell(13, 7).Alignment = cellRightCenter
        documento.Cell(13, 7).text = notapedido
        
        
        For i = documento.Rows To 45
            documento.AddItem ""
        Next i
        
        'TIPOS DE PAGO
        Set cSql2.ActiveConnection = ventasRubro
        cSql2.sql = "SELECT dp.numerocheque, b.nombre, dp.monto, DATE_FORMAT(dp.vencimiento,'%d-%m-%Y') AS vencimiento "
        cSql2.sql = cSql2.sql & "FROM sv_documento_pagos_" + empresaActiva + " AS dp LEFT JOIN " & baseVentas & ".sv_maestrobancos AS b ON dp.banco = b.codigobanco "
        cSql2.sql = cSql2.sql & "WHERE dp.local = '" & empresaActiva & "' AND dp.tipopago = '2' AND dp.tipo = '" & TIPO & "' AND dp.numero = '" & numerofactura & "' "
        cSql2.sql = cSql2.sql & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
        cSql2.Execute
        'Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
        
        If cSql2.RowsAffected > 0 Then
            Set resultados2 = cSql2.OpenResultset
            i = 38
            While Not resultados2.EOF
                documento.Cell(i, 2).text = resultados2("numerocheque")
                documento.Cell(i, 3).text = resultados2("nombre")
                documento.Cell(i, 5).text = Format(resultados2("monto"), "$ ###,###,##0")
                documento.Cell(i, 6).text = resultados2("vencimiento")
                i = i + 1
                resultados2.MoveNext
            Wend
        End If
        Set cSql2 = Nothing
        cSql2.Close
        Set resultados2 = Nothing
        
        
        documento.Cell(45, 6).text = "DESCUENTO"
        documento.Cell(45, 6).Alignment = cellLeftCenter
        documento.Cell(45, 7).text = Format(descuento, "$ ###,###,##0")
        
        documento.AddItem ""
        documento.Cell(46, 7).text = Format(neto, "$ ###,###,##0")
        
        documento.AddItem ""
        documento.Cell(47, 7).text = Format(piva, "$ ###,###,##0")
        
        documento.AddItem ""
        documento.Cell(48, 7).text = Format(piha, "$ ###,###,##0")
        
        documento.AddItem ""
        documento.Cell(49, 7).text = Format(total, "$ ###,###,##0")
        
        documento.AutoRedraw = True
        documento.Refresh
        
        documento.PageSetup.LeftMargin = 0.25
        documento.PageSetup.RightMargin = 0
        documento.PageSetup.TopMargin = 2.5
        documento.PageSetup.BottomMargin = 0
        
        For i = 1 To documento.PageSetup.PaperSizes.Count
            If UCase(documento.PageSetup.PaperSizes.Item(i).PaperName) = "FACTURA" Then
                documento.PageSetup.PaperSize = documento.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        
        'Documento.DirectPrint
        documento.PrintPreview
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & numeroFactura)
    End If
End Sub
