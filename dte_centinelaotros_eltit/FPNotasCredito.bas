Attribute VB_Name = "FPNotasCredito"
Option Explicit

Public Sub imprimeNotaCredito(ByVal TIPO As String, ByVal NUMEROFACTURA As String, ByRef Documento As Grid, ByRef rollo As Adodc)
    Dim SS As String
    Dim i As Integer
    Dim k As Integer
    Dim cad As String
    Dim totalprod As String
    Dim descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim LINEAS As Integer
    Dim fecha As String
    Dim vencimiento As String
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
    Dim CODIGO As String
    Dim tiposDePago As String
    
    
    tabla = "SELECT d.codigo, CONCAT(IF(d.numerofactura<>'', d.numerofactura, d.cantidad), '" & vbTab & "', d.descripcion, '" & vbTab & vbTab & "', d.cantidad * mpf.contenido, '" & vbTab & "', d.precio, '" & vbTab & "', d.total) AS item, d.total AS totalpro, c.neto, c.iva, c.impuestoharina, c.total, DATE_FORMAT(c.fecha,'%d-%m-%Y') AS fecha, d.descuento, cl.rut, cl.nombre, cl.direccion, cl.giro, cl.ciudad, cl.comuna, cl.fono1, c.cajera, DATE_FORMAT(c.vencimiento,'%d-%m-%Y') AS vencimiento, IF(notapedido = '000000000000','',notapedido) AS notapedido "
    tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS c INNER JOIN sv_documento_detalle_" + empresaActiva + " AS d ON c.tipo = d.tipo AND c.numero = d.numero and  c.caja=d.caja and c.local=d.local and c.fecha=d.fecha INNER JOIN " & baseVentas & ".sv_maestroclientes AS cl ON c.rut = cl.rut INNER JOIN " & basedatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON d.codigo = mpf.codigobarra "
    tabla = tabla & "WHERE c.tipo = '" & TIPO & "' AND c.foliosii = '" & NUMEROFACTURA & "' and c.local='" & empresaActiva & "' and  c.caja='" & PVentas.dato30.text & "' and c.fecha= '" & PVentas.dato5.text & "-" & PVentas.dato4.text & "-" & PVentas.dato3.text & "'"
    tabla = tabla & "ORDER BY d.linea"

    Call ConectarControlData(rollo, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    
    If rollo.Recordset.RecordCount > 0 Then
        Documento.Rows = 2
        Documento.Cols = 8
        
        Documento.DefaultFont.Name = "Arial"
        Documento.DefaultFont.Size = 8
        Documento.DefaultFont.Bold = False
        
        Documento.Column(0).Width = 0
        Documento.Column(1).Width = 45
        Documento.Column(2).Width = 100
        Documento.Column(3).Width = 210
        Documento.Column(4).Width = 105
        Documento.Column(5).Width = 70
        Documento.Column(6).Width = 90
        Documento.Column(7).Width = 125
        
        Documento.Column(1).Alignment = cellRightCenter
        Documento.Column(2).Alignment = cellCenterCenter
        Documento.Column(3).Alignment = cellLeftCenter '/**/
        Documento.Column(4).Alignment = cellLeftCenter '/**/
        Documento.Column(5).Alignment = cellRightCenter
        Documento.Column(6).Alignment = cellRightCenter
        Documento.Column(7).Alignment = cellRightCenter
        
        Documento.DefaultRowHeight = 13
        
        Documento.PageSetup.PrintGridlines = False
        Documento.AutoRedraw = False
        
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        Documento.AddItem ""
        
        rollo.Recordset.MoveFirst
        
        fecha = rollo.Recordset.Fields("fecha")
        nombre = "       " & rollo.Recordset.Fields("nombre")
        rut = "       " & rollo.Recordset.Fields("rut")
        direccion = "       " & rollo.Recordset.Fields("direccion")
        ciudad = "       " & rollo.Recordset.Fields("ciudad")
        comuna = rollo.Recordset.Fields("comuna")
        giro = rollo.Recordset.Fields("giro")
        fono = rollo.Recordset.Fields("fono1")
        vencimiento = rollo.Recordset.Fields("vencimiento")
        vendedor = rollo.Recordset.Fields("cajera")
        notapedido = rollo.Recordset.Fields("notapedido")
        neto = rollo.Recordset.Fields("neto")
        piva = rollo.Recordset.Fields("iva")
        piha = rollo.Recordset.Fields("impuestoharina")
        total = rollo.Recordset.Fields("total")
        descuento = "0"
        While Not rollo.Recordset.EOF
            CODIGO = Right(rollo.Recordset.Fields("codigo"), 4)
            descuento = Replace(Str(CDbl(descuento) + CDbl(rollo.Recordset.Fields("totalpro")) * CDbl(rollo.Recordset.Fields("descuento")) / 100), ".", ",")
            descuento = Mid(descuento, 2, Len(descuento))
            Documento.AddItem CODIGO & vbTab & rollo.Recordset.Fields("item"), True
            Documento.Range(Documento.Rows - 1, 3, Documento.Rows - 1, 4).Merge
            rollo.Recordset.MoveNext
        Wend
        Documento.Cell(5, 4).Alignment = cellRightCenter
        Documento.Cell(5, 6).text = NUMEROFACTURA
        
        'TIPOS DE PAGO
        tabla = "SELECT dp.tipopago, SUM(dp.monto) AS monto "
        tabla = tabla & "FROM sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero "
        tabla = tabla & "WHERE dp.tipo = '" & TIPO & "' AND dp.numero = '" & NUMEROFACTURA & "' "
        tabla = tabla & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
    
        Call ConectarControlData(rollo, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        
        If rollo.Recordset.RecordCount > 0 Then
            rollo.Recordset.MoveFirst
            tiposDePago = ""
            While Not rollo.Recordset.EOF
                CODIGO = rollo.Recordset.Fields("tipopago")
                Select Case CODIGO
                    Case "1"    'EFECTIVO
                        tiposDePago = tiposDePago & "EFECTIVO " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
                    Case "2"    'CHEQUE
                        tiposDePago = tiposDePago & "CHEQUE " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
                    Case "6"    'CREDITO DIRECTO
                        tiposDePago = tiposDePago & "CREDITO DIRECTO " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
                End Select
                rollo.Recordset.MoveNext
            Wend
        End If
        
        Documento.Range(4, 2, 4, 3).Merge
        Documento.Range(4, 2, 4, 3).Alignment = cellCenterCenter
        Documento.Cell(4, 2).text = leerNombreEmpresa(empresaActiva)
        
        Documento.RowHeight(6) = 15
        Documento.RowHeight(7) = 10
        Documento.RowHeight(8) = 15
        Documento.RowHeight(9) = 15
        Documento.RowHeight(10) = 15
        Documento.RowHeight(11) = 15
        
        Documento.RowHeight(15) = 5
        
        'FECHA
        Documento.Range(6, 1, 6, 2).Merge
        Documento.Range(6, 1, 6, 2).Alignment = cellCenterCenter
        Documento.Cell(6, 1).text = fecha
        
        'SEÑORES
        Documento.Range(8, 2, 8, 3).Merge
        Documento.Range(8, 2, 8, 3).Alignment = cellLeftCenter
        Documento.Cell(8, 2).text = nombre
        'RUT
        Documento.Range(9, 2, 9, 3).Merge
        Documento.Range(9, 2, 9, 3).Alignment = cellLeftCenter
        Documento.Cell(9, 2).text = rut
        'DIRECCION
        Documento.Range(10, 2, 10, 3).Merge
        Documento.Range(10, 2, 10, 3).Alignment = cellLeftCenter
        Documento.Cell(10, 2).text = direccion
        'CIUDAD
        Documento.Range(11, 2, 11, 3).Merge
        Documento.Range(11, 2, 11, 3).Alignment = cellLeftCenter
        Documento.Cell(11, 2).text = ciudad
        
        'GIRO
        Documento.Range(9, 5, 9, 6).Merge
        Documento.Range(9, 5, 9, 6).Alignment = cellLeftCenter
        Documento.Cell(9, 5).text = giro
        'COMUNA
        Documento.Range(11, 5, 11, 6).Merge
        Documento.Range(11, 5, 11, 6).Alignment = cellLeftCenter
        Documento.Cell(11, 5).text = comuna
        'FONO
        Documento.Cell(10, 6).Alignment = cellRightCenter
        Documento.Cell(10, 6).text = fono
        
        'VENCIMIENTO
        Documento.Cell(9, 7).Alignment = cellCenterCenter
        Documento.Cell(9, 7).text = vencimiento
        'VENDEDOR
        Documento.Cell(11, 7).Alignment = cellRightCenter
        Documento.Cell(11, 7).text = vendedor
        'CONDICIONES DE PAGO
        Documento.Range(13, 3, 13, 6).Merge
        Documento.Range(13, 3, 13, 6).Alignment = cellLeftCenter
        Documento.Cell(13, 3).text = tiposDePago
        'NOTAPEDIDO
        Documento.Cell(13, 7).Alignment = cellRightCenter
        Documento.Cell(13, 7).text = notapedido
        
        
        For i = Documento.Rows To 45
            Documento.AddItem ""
        Next i
        
        'TIPOS DE PAGO
        tabla = "SELECT dp.numerodocumento, b.nombre, dp.monto, DATE_FORMAT(dp.vencimiento,'%d-%m-%Y') AS vencimiento "
        tabla = tabla & "FROM sv_documento_pagos_" + empresaActiva + " AS dp LEFT JOIN " & baseVentas & ".sv_maestrobancos AS b ON dp.banco = b.codigobanco "
        tabla = tabla & "WHERE dp.local = '" & empresaActiva & "' AND dp.tipopago = '2' AND dp.tipo = '" & TIPO & "' AND dp.numero = '" & NUMEROFACTURA & "' "
        tabla = tabla & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
    
        Call ConectarControlData(rollo, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        
        If rollo.Recordset.RecordCount > 0 Then
            rollo.Recordset.MoveFirst
            i = 38
            While Not rollo.Recordset.EOF
                Documento.Cell(i, 2).text = rollo.Recordset.Fields("numerocheque")
                Documento.Cell(i, 3).text = rollo.Recordset.Fields("nombre")
                Documento.Cell(i, 5).text = Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0")
                Documento.Cell(i, 6).text = rollo.Recordset.Fields("vencimiento")
                i = i + 1
                rollo.Recordset.MoveNext
            Wend
        End If
        
        Documento.Cell(45, 6).text = "DESCUENTO"
        Documento.Cell(45, 6).Alignment = cellLeftCenter
        Documento.Cell(45, 7).text = Format(descuento, "$ ###,###,##0")
        
        Documento.AddItem ""
        Documento.Cell(46, 7).text = Format(neto, "$ ###,###,##0")
        
        Documento.AddItem ""
        Documento.Cell(47, 7).text = Format(piva, "$ ###,###,##0")
        
        Documento.AddItem ""
        Documento.Cell(48, 7).text = Format(piha, "$ ###,###,##0")
        
        Documento.AddItem ""
        Documento.Cell(49, 7).text = Format(total, "$ ###,###,##0")
        
        Documento.AutoRedraw = True
        Documento.Refresh
        
        Documento.PageSetup.LeftMargin = 0.25
        Documento.PageSetup.RightMargin = 0
        Documento.PageSetup.TopMargin = 2.5
        Documento.PageSetup.BottomMargin = 0
        
        For i = 1 To Documento.PageSetup.PaperSizes.Count
            If UCase(Documento.PageSetup.PaperSizes.Item(i).PaperName) = "FACTURA" Then
                Documento.PageSetup.PaperSize = Documento.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        
        'Documento.DirectPrint
        Documento.PrintPreview
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & numeroFactura)
    End If
End Sub


Public Sub imprimenotadecredito2(ByVal NUMEROFACTURA As String, ByRef Documento As Grid, ByRef rollo As Adodc, TIPO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim cad As String
    Dim h As Integer
    Dim totalprod As String
    Dim descuento As String
    Dim neto As String
    Dim piva As String
    Dim total As String
    Dim tpago As String
    Dim LINEAS As Integer
    Dim fecha As String
    Dim O As Integer
    Dim tabla As String
    Dim totNeto As String
    Dim totIva As String
    Dim totIha As String
    Dim totIca As String
    Dim totIla As String
    Dim totaldescuento As String
    Dim porcedescuento As String
    Dim MONTO As String
    Dim tipopago As String
    Dim CAMPO1 As String * 8
    Dim campo2 As String * 8
    Dim campo3 As String * 8
    Dim campo4 As String * 8
    Dim campo5 As String * 8
    Dim campo6 As String * 8
    Dim PORCE As Double
    Dim dife As Double
    Dim exento As Double
    Dim ilarefrescos As String
    Dim ilalicores As String
    Dim ilavinos As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Documento.Rows = 1
    
 Set cSql.ActiveConnection = ventasRubro
'    cSql.sql = "SELECT dd.codigo, dd.descripcion,  dd.cantidad,  dd.precio, dd.cantidad*dd.precio, dd.total AS totalpro, dd.precio, dd.cantidad, dc.rut, dc.sucursal, dc.neto as neto, dc.iva, dc.impuestoharina AS iha, dc.impuestocarne AS ica, dc.impuestoilarefrescos as ilarefrescos, dc.impuestoilalicores as ilalicores , dc.impuestoilavinos AS ilavinos, dc.total, dc.fecha ,dc.descuento,dd.descuento as descuento2,dc.foliosii, dd.tipodocumento,dd.numerodocumento "
'    cSql.sql = cSql.sql & "from sv_documento_cabeza_" + empresaactiva + " AS dc, sv_documento_detalle_" + empresaactiva + " AS dd "
'    cSql.sql = cSql.sql & "WHERE dc.caja='" + caja + "' and dc.local = '" & empresaactiva & "' AND dc.local = dd.local AND dd.tipo = '" & TIPO & "' AND dd.numero = '" & NUMEROFACTURA & "' AND dd.tipo = dc.tipo AND dd.numero = dc.numero ORDER BY dd.linea ASC "
'    cSql.Execute
       cSql.sql = "SELECT dd.codigo, dd.descripcion,  dd.cantidad,  dd.precio, dd.cantidad*dd.precio, dd.total AS totalpro, dd.precio, dd.cantidad, dc.rut, dc.sucursal, dc.neto as neto, dc.iva, dc.impuestoharina AS iha, dc.impuestocarne AS ica, dc.impuestoilarefrescos as ilarefrescos, dc.impuestoilalicores as ilalicores , dc.impuestoilavinos AS ilavinos, dc.total, dc.fecha ,dc.descuento,dd.descuento as descuento2,dc.donacion,dc.foliosii,dc.numero,dd.tipodespacho,dd.vendedor "
    cSql.sql = cSql.sql & "from sv_documento_cabeza_" + empresaActiva + " AS dc, sv_documento_detalle_" + empresaActiva + " AS dd "
    cSql.sql = cSql.sql & "WHERE dd.caja=dc.caja and dd.local = '" & empresaActiva & "' AND dd.caja='" + PVentas.dato30.text + "' and dc.local = dd.local AND dd.tipo = '" & TIPO & "' AND dc.foliosii = '" & NUMEROFACTURA & "' AND dd.tipo = dc.tipo AND dd.numero = dc.numero and dc.fecha=dd.fecha and dd.fecha='" & PVentas.dato5.text & "-" & PVentas.dato4.text & "-" & PVentas.dato3.text & "' ORDER BY dd.linea ASC "
    cSql.Execute
    
    If cSql.RowsAffected > 0 Then
     Set resultados = cSql.OpenResultset
        exento = 0
        Documento.Rows = 1
        Documento.Cols = 6
        Documento.Rows = 40
        
'        Documento.DefaultFont.Name = "COURIER"
        Documento.DefaultFont.Size = 8
        Documento.DefaultFont.Bold = False
        Documento.Column(0).Width = 0
        Documento.Column(1).Width = 150
        Documento.Column(2).Width = 90
        Documento.Column(3).Width = 265
        Documento.Column(4).Width = 100
        Documento.Column(5).Width = 150
        Documento.Column(1).Alignment = cellRightCenter
        Documento.Column(2).Alignment = cellCenterCenter
        Documento.Column(3).Alignment = cellLeftCenter
        Documento.Column(4).Alignment = cellRightCenter
        Documento.Column(5).Alignment = cellRightCenter
  '
        Documento.DefaultRowHeight = 13
        
        Documento.PageSetup.PrintGridlines = False
        Documento.AutoRedraw = False
   
        j = 15
       
        total = resultados("total")
        fecha = resultados("fecha")
        totNeto = resultados("neto")
        totIva = resultados("iva")
        totIha = resultados("iha")
        totIca = resultados("ica")
        ilarefrescos = resultados("ilarefrescos")
        ilalicores = resultados("ilalicores")
        ilavinos = resultados("ilavinos")
        totaldescuento = resultados("descuento")
        porcedescuento = resultados("descuento2")
        descuento = "0"
                      
        Rem NUMERO
        Documento.Cell(3, 5).text = "F/O: " & resultados("numero")
'        Documento.Cell(4, 5).text = "F/O: " & Resultados("tpdocumento") & " " & Resultados("numerodocumento")
        Documento.Range(2, 2, 2, 3).Merge
        Documento.Range(2, 2, 2, 3).Alignment = cellCenterCenter
        Documento.Cell(2, 2).text = leerNombreEmpresa(empresaActiva)
        Documento.Cell(2, 5).text = "F/F: " & resultados("foliosii")

        
        
        'SEÑORES
        Documento.Range(6, 2, 6, 3).Merge
        Documento.Range(6, 2, 6, 3).Alignment = cellLeftCenter
        Documento.Cell(6, 2).text = PVentas.lblRazon.Caption
        'FECHA
        Documento.Cell(6, 5).text = fecha
          
        'DIRECCION
        Documento.Range(7, 2, 7, 3).Merge
        Documento.Range(7, 2, 7, 3).Alignment = cellLeftCenter
        Documento.Cell(7, 2).text = PVentas.LBLDIRECCION.Caption
        
        
      'RUT
'        Documento.Range(8, 4, 8, 5).Merge
'        Documento.Range(8, 4, 8, 5).Alignment = cellLeftCenter
'        Documento.Cell(9, 5).Alignment = cellCenterCenter
        Documento.Cell(7, 5).text = "     " + Format(PVentas.dato6.text, "###,###,##0") & "-" & PVentas.lblDV.Caption
 
 
        'GIRO
        Documento.Range(8, 2, 8, 3).Merge
        Documento.Range(8, 2, 8, 3).Alignment = cellLeftCenter
        Documento.Cell(8, 2).text = leerGiroCliente(PVentas.dato6.text & PVentas.lblDV.Caption, PVentas.dato7.text)

        
        'CIUDAD
'        Documento.Range(10, 4, 10, 5).Merge
'        Documento.Range(11, 5, 11, 5).Alignment = cellCenterCenter
        Documento.Cell(8, 5).text = "     " + PVentas.LBLCIUDAD.Caption
    
        'DESCUENTO
'        Documento.Range(10, 4, 10, 5).Merge
'        Documento.Range(10, 4, 10, 5).Alignment = cellLeftCenter
'        Documento.Cell(10, 4).text = "D:" + Descuento
'
        
        
        LINEAS = 15
        While Not resultados.EOF
            descuento = Str(CDbl(descuento) + Int(resultados("cantidad") * resultados("precio")) - Int(resultados("totalpro")))
            descuento = Mid(descuento, 2, Len(descuento))
            
            LINEAS = LINEAS + 1
            
            
            Documento.Cell(LINEAS, 1).text = resultados(0)
            dife = CDbl(resultados(2) - resultados(2))
            If dife <> 0 Then
             Documento.Cell(LINEAS, 2).text = Format(resultados(2), "###,##0,00")
            Else
             Documento.Cell(LINEAS, 2).text = Format(resultados(2), "###,###")
            End If
            Documento.Cell(LINEAS, 3).text = resultados(1) & " " & resultados(22) & "-" & resultados(23)
            Documento.Cell(LINEAS, 4).text = Format(resultados(3), " $ ###,###,###")
            Documento.Cell(LINEAS, 5).text = Format(resultados(4), " $ ###,###,###")
            
            resultados.MoveNext
        Wend
        MONTO = WORDNUM(Format(total, "########0"), "PESO", "PESOS", 0)
        
        Rem monto = numToLet(Format(total, "########0"), "PESO", "PESOS", 0)
        If totaldescuento <> 0 Then
        PORCE = totaldescuento / (CDbl(total) + CDbl(totaldescuento)) * 100
        'Documento.Cell(28, 4).text = "(-)%" + Format(PORCE, "##")
        Documento.Cell(28, 5).text = Format(totaldescuento * -1, " $ ###,###,##0")
        
        
        End If
        Set cSql = Nothing
        cSql.Close
        Set resultados = Nothing
        
        Documento.Cell(29, 4).Alignment = cellLeftCenter
        Documento.Cell(29, 4).text = "  NETO"
        Documento.Cell(29, 5).text = Format(totNeto, " $ ####,##0")
        
        Documento.Cell(30, 4).Alignment = cellLeftCenter
        Documento.Cell(30, 4).text = "  IVA"
        Documento.Cell(30, 5).text = Format(totIva, " $ ###,##0")
        
        Documento.Cell(31, 4).Alignment = cellLeftCenter
        Documento.Cell(31, 4).text = "  EXENTO"
        Documento.Cell(31, 5).text = Format(exento, " $ ###,###,##0")
        
        Documento.Cell(32, 4).Alignment = cellLeftCenter
        Documento.Cell(32, 4).text = "  TOTAL"
        Documento.Cell(32, 5).text = Format(total, " $ ###,###,##0")

        Documento.Range(30, 1, 30, 3).Merge
        Documento.Range(30, 1, 30, 3).Alignment = cellLeftCenter
        Documento.Cell(30, 1).text = "                        " + MONTO
        Documento.Range(29, 1, 31, 3).WrapText = True

        Documento.Range(31, 2, 31, 3).Merge
        Documento.Range(31, 2, 31, 3).Alignment = cellLeftCenter
        Documento.Range(31, 2, 31, 3).FontBold = True
        Documento.Cell(31, 2).text = "   ILA 13     " + "   ILA 15     " + "   ILA 27     " + "   HARINA     " + "   CARNE      "
        
        
        CAMPO1 = String(7 - Len(ilarefrescos), 32) & Format(ilarefrescos, "####,##0")
        campo2 = String(7 - Len(ilavinos), 32) & Format(ilavinos, "####,##0")
        campo3 = String(7 - Len(ilalicores), 32) & Format(ilalicores, "####,##0")
        campo4 = String(7 - Len(totIha), 32) & Format(totIha, "####,##0")
        campo5 = String(7 - Len(totIca), 32) & Format(totIca, "####,##0")
         cad = "  " & CAMPO1 & "          " & campo2 & "          " & campo3 & "          " & campo4 & "          " & campo5
        

         Documento.Range(32, 2, 32, 3).Merge
         Documento.Range(32, 2, 32, 3).Alignment = cellLeftCenter
         Documento.Range(32, 2, 32, 3).FontBold = True
        
        Documento.Cell(32, 2).text = cad
        
        Documento.AutoRedraw = True
        Documento.Refresh
        

        
        Documento.PageSetup.LeftMargin = 0.25
        Documento.PageSetup.RightMargin = 0
        Documento.PageSetup.TopMargin = 2.5
        Documento.PageSetup.BottomMargin = 0
        
        For i = 1 To Documento.PageSetup.PaperSizes.Count
            If UCase(Documento.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
                Documento.PageSetup.PaperSize = Documento.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        
        Documento.PageSetup.PrintGridlines = False

        Documento.PrintPreview

        
'        Documento.PrintPreview
        
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & NUMEROFACTURA)
    End If

End Sub
