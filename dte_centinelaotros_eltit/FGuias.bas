Attribute VB_Name = "FGuias"
Option Explicit
    Private campos(30, 5) As String
    Private c As Cliente

Public Sub imprimeGuia(ByVal TIPO As String, ByVal NUMERO As String, ByRef Impresion As Grid, ByRef data As Adodc)
    Select Case TIPO
        Case "GM", "FM"
            Call imprimeGuiaMolienda(TIPO, NUMERO, Impresion, data)
        Case "GD", "ZE"
            Call imprimeGuiaDespacho(TIPO, NUMERO, Impresion, data)
    End Select
End Sub

Private Sub imprimeGuiaDespacho(ByVal TIPO As String, ByVal NUMERO As String, ByRef Impresion As Grid, ByRef data As Adodc)
    Dim i As Integer
    Dim descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim fecha As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim codigo As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set cSql.ActiveConnection = ventasRubro
    cSql.sql = "SELECT d.codigo, CONCAT(IF(d.tipo = 'NV','',d.cantidad), '" & vbTab & "', d.descripcion, '" & vbTab & vbTab & "', d.unidades, '" & vbTab & "', '$ ', FORMAT(d.precio,0), '" & vbTab & "', '$ ', FORMAT(d.total,0)) AS item, d.total AS totalpro, c.neto, c.iva, c.impuestoharina, c.total, DATE_FORMAT(c.fecha,'%d-%m-%Y') AS fecha, c.descuento, c.rut, c.sucursal "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS c INNER JOIN sv_documento_detalle_" + empresaActiva + " AS d ON c.local = d.local AND c.tipo = d.tipo AND c.numero = d.numero INNER JOIN " & baseDatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON d.codigo = mpf.codigobarra "
    cSql.sql = cSql.sql & "WHERE c.local = '" & empresaActiva & "' AND c.tipo = '" & TIPO & "' AND c.numero = '" & NUMERO & "' "
    cSql.sql = cSql.sql & "ORDER BY d.linea"
    cSql.Execute
'    Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
    
    If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
        Impresion.Rows = 2
        Impresion.Cols = 8
        
        Impresion.DefaultFont.Name = "Arial"
        Impresion.DefaultFont.Size = 8
        Impresion.DefaultFont.Bold = False
        
        Impresion.Column(0).Width = 0
        Impresion.Column(1).Width = 45
        Impresion.Column(2).Width = 100
        Impresion.Column(3).Width = 210
        Impresion.Column(4).Width = 105
        Impresion.Column(5).Width = 70
        Impresion.Column(6).Width = 90
        Impresion.Column(7).Width = 125
        
        Impresion.Column(1).Alignment = cellRightCenter
        Impresion.Column(2).Alignment = cellCenterCenter
        Impresion.Column(3).Alignment = cellLeftCenter '/**/
        Impresion.Column(4).Alignment = cellLeftCenter '/**/
        Impresion.Column(5).Alignment = cellRightCenter
        Impresion.Column(6).Alignment = cellRightCenter
        Impresion.Column(7).Alignment = cellRightCenter
        
        Impresion.DefaultRowHeight = 13
        
        Impresion.PageSetup.PrintGridlines = False
        Impresion.AutoRedraw = False
        
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        
        data.Recordset.MoveFirst
        
        Call LEERCLIENTE(c, resultados("rut"), resultados("sucursal"), "=")
        
        fecha = resultados("fecha")
        fecha = Format(fecha, "dd") & "                            " & Format(fecha, "mmmm") & "                                             " & Right(Format(fecha, "yyyy"), 1) & "        "
        nombre = "       " & c.nombre
        rut = "       " & Left(c.rut, Len(c.rut) - 1) & "-" & Right(c.rut, 1)
        direccion = "       " & c.direccion
        ciudad = "       " & c.ciudad
        comuna = "       " & c.comuna
        giro = "       " & c.giro
        neto = resultados("neto")
        piva = resultados("iva")
        piha = resultados("impuestoharina")
        total = resultados("total")
        descuento = resultados("descuento")
        While Not resultados.EOF
            codigo = Right(resultados("codigo"), 4)
            Impresion.AddItem codigo & vbTab & Replace(resultados("item"), ",", "."), True
            Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, 4).Merge
            resultados.MoveNext
        Wend
        Set cSql = Nothing
        cSql.Close
        Set resultados = Nothing
        
        Impresion.Cell(5, 4).Alignment = cellRightCenter
        Impresion.Cell(5, 6).text = NUMERO
        
        Impresion.RowHeight(6) = 15
        Impresion.RowHeight(7) = 10
        Impresion.RowHeight(8) = 15
        Impresion.RowHeight(9) = 15
        Impresion.RowHeight(10) = 15
        Impresion.RowHeight(11) = 15
        
        Impresion.RowHeight(15) = 5
        
        'empresa
        Impresion.Range(2, 2, 2, 3).Merge
        Impresion.Range(2, 2, 2, 3).Alignment = cellLeftGeneral
        Impresion.Cell(2, 2).text = leerNombreEmpresa(empresaActiva)
        'FECHA
        Impresion.Range(6, 1, 6, 2).Merge
        Impresion.Range(6, 1, 6, 2).Alignment = cellCenterCenter
        Impresion.Cell(6, 1).text = fecha
        Impresion.Range(6, 2, 7, 3).Merge
        Impresion.Range(6, 2, 7, 3).Alignment = cellRightBottom
        
        'SEÑORES
        Impresion.Range(8, 2, 9, 3).Merge
        Impresion.Range(8, 2, 8, 3).Alignment = cellLeftCenter
        Impresion.Cell(8, 2).text = nombre
        
        'RUT
        Impresion.Range(8, 7, 9, 7).Merge
        Impresion.Range(8, 7, 9, 7).Alignment = cellLeftCenter
        Impresion.Cell(8, 7).text = rut
        
        'DIRECCION
        Impresion.Range(10, 2, 11, 3).Merge
        Impresion.Range(10, 2, 10, 3).Alignment = cellLeftCenter
        Impresion.Cell(10, 2).text = direccion
        
        'CIUDAD
        Impresion.Range(10, 7, 11, 7).Merge
        Impresion.Range(11, 7, 11, 7).Alignment = cellLeftCenter
        Impresion.Cell(10, 7).text = ciudad
        
        'GIRO
        Impresion.Range(12, 2, 13, 3).Merge
        Impresion.Range(12, 2, 13, 3).Alignment = cellLeftTop
        Impresion.Cell(12, 2).text = giro
        
        'COMUNA
        Impresion.Range(12, 7, 13, 7).Merge
        Impresion.Range(12, 7, 13, 7).Alignment = cellLeftTop
        Impresion.Cell(12, 7).text = comuna
        
        For i = Impresion.Rows To 35
            Impresion.AddItem ""
        Next i
        
        Impresion.AddItem ""
        Impresion.Cell(35, 6).text = "TOTAL"
        Impresion.Cell(35, 7).text = Format(total, "$ ###,###,##0")
        
        Impresion.AutoRedraw = True
        Impresion.Refresh
        
        Impresion.PageSetup.LeftMargin = 0.5
        Impresion.PageSetup.RightMargin = 0
        Impresion.PageSetup.TopMargin = 2.5
        Impresion.PageSetup.BottomMargin = 0
        
        
        Call verificaImpresora(3, Impresion)
        
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & numeroFactura)
    End If
End Sub

Private Sub imprimeGuiaMolienda(ByVal TIPO As String, ByVal NUMERO As String, ByRef Impresion As Grid, ByRef data As Adodc)
    Dim i As Integer
    Dim cadena As String
    Dim fecha As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim codigo As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set cSql.ActiveConnection = ventasRubro
    cSql.sql = "SELECT gm.numero, DATE_FORMAT(gm.fecha,'%d-%m-%Y') AS fecha, gm.rut, gm.sucursal, gm.harina, gm.afrecho, gm.harinilla, gm.tipodocumento, gm.numerodocumento, gm.valor "
    cSql.sql = cSql.sql & "FROM sv_guiasmolienda AS gm "
    cSql.sql = cSql.sql & "WHERE gm.local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero = '" & NUMERO & "'"
    cSql.Execute
'    Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
    
    If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
        Impresion.Rows = 2
        Impresion.Cols = 8
        
        Impresion.DefaultFont.Name = "Arial"
        Impresion.DefaultFont.Size = 8
        Impresion.DefaultFont.Bold = False
        
        Impresion.Column(0).Width = 0
        Impresion.Column(1).Width = 45
        Impresion.Column(2).Width = 100
        Impresion.Column(3).Width = 210
        Impresion.Column(4).Width = 105
        Impresion.Column(5).Width = 70
        Impresion.Column(6).Width = 90
        Impresion.Column(7).Width = 125
        
        Impresion.Column(1).Alignment = cellRightCenter
        Impresion.Column(2).Alignment = cellCenterCenter
        Impresion.Column(3).Alignment = cellLeftCenter '/**/
        Impresion.Column(4).Alignment = cellRightCenter '/**/
        Impresion.Column(5).Alignment = cellRightCenter
        Impresion.Column(6).Alignment = cellRightCenter
        Impresion.Column(7).Alignment = cellRightCenter
        
        Impresion.DefaultRowHeight = 13
        
        Impresion.PageSetup.PrintGridlines = False
        Impresion.AutoRedraw = False
        
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        Impresion.AddItem ""
        
        data.Recordset.MoveFirst
        
        Call LEERCLIENTE(c, resultados("rut"), resultados("sucursal"), "=")
        
        fecha = resultados("fecha")
        fecha = Format(fecha, "dd") & "                            " & Format(fecha, "mmmm") & "                                             " & Right(Format(fecha, "yyyy"), 1) & "        "
        nombre = "       " & c.nombre
        rut = "       " & Left(c.rut, Len(c.rut) - 1) & "-" & Right(c.rut, 1)
        direccion = "       " & c.direccion
        ciudad = "       " & c.ciudad
        comuna = "       " & c.comuna
        giro = "       " & c.giro
        
        '*********************************
        cadena = vbTab & vbTab
        cadena = cadena & "HARINA FLOR FINA" & vbTab
        cadena = cadena & Format(Replace(resultados("harina"), ".", ","), "###,###,##0.00")
        Impresion.AddItem cadena, True
        
        cadena = vbTab & vbTab
        cadena = cadena & "AFRECHO" & vbTab
        cadena = cadena & Format(Replace(resultados("afrecho"), ".", ","), "###,###,##0.00")
        Impresion.AddItem cadena, True
        
        cadena = vbTab & vbTab
        cadena = cadena & "HARINILLA" & vbTab
        cadena = cadena & Format(Replace(resultados("harinilla"), ".", ","), "###,###,##0.00")
        Impresion.AddItem cadena, True
        
        Impresion.AddItem "", True
        Impresion.AddItem "", True
        
        Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 5).Borders(cellEdgeBottom) = cellThick
        
        Impresion.AddItem "", True
        
        cadena = vbTab
        cadena = cadena & "SERVICIO DE MOLIENDA"
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 5).Merge
        
        cadena = vbTab
        If resultados("tipodocumento") = "BV" Then
            cadena = cadena & "SEGUN BOLETA NUMERO "
        Else
            cadena = cadena & "SEGUN FACTURA NUMERO "
        End If
        cadena = cadena & resultados("numerodocumento")
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 5).Merge
        
        Impresion.AddItem "", True
        
        cadena = vbTab
        If resultados("tipodocumento") = "BV" Then
            cadena = cadena & "VALOR BOLETA          "
        Else
            cadena = cadena & "VALOR FACTURA          "
        End If
        cadena = cadena & Format(resultados("valor"), "$ ###,###,##0")
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 5).Merge
        
        Impresion.AddItem "", True
        
        Impresion.Range(Impresion.Rows - 6, 2, Impresion.Rows - 1, 5).FontBold = True
        Impresion.Range(Impresion.Rows - 6, 2, Impresion.Rows - 1, 5).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 5).Borders(cellEdgeBottom) = cellThick
        '*********************************
        
        Impresion.Cell(5, 4).Alignment = cellRightCenter
        Impresion.Cell(5, 6).text = NUMERO
        
        Impresion.RowHeight(6) = 15
        Impresion.RowHeight(7) = 10
        Impresion.RowHeight(8) = 15
        Impresion.RowHeight(9) = 15
        Impresion.RowHeight(10) = 15
        Impresion.RowHeight(11) = 15
        
        Impresion.RowHeight(15) = 5
        
        'FECHA
        Impresion.Range(6, 1, 6, 2).Merge
        Impresion.Range(6, 1, 6, 2).Alignment = cellCenterCenter
        Impresion.Cell(6, 1).text = fecha
        Impresion.Range(6, 2, 7, 3).Merge
        Impresion.Range(6, 2, 7, 3).Alignment = cellRightBottom
        
        'SEÑORES
        Impresion.Range(8, 2, 9, 3).Merge
        Impresion.Range(8, 2, 8, 3).Alignment = cellLeftCenter
        Impresion.Cell(8, 2).text = nombre
        
        'RUT
        Impresion.Range(8, 7, 9, 7).Merge
        Impresion.Range(8, 7, 9, 7).Alignment = cellLeftCenter
        Impresion.Cell(8, 7).text = rut
        
        'DIRECCION
        Impresion.Range(10, 2, 11, 3).Merge
        Impresion.Range(10, 2, 10, 3).Alignment = cellLeftCenter
        Impresion.Cell(10, 2).text = direccion
        
        'CIUDAD
        Impresion.Range(10, 7, 11, 7).Merge
        Impresion.Range(11, 7, 11, 7).Alignment = cellLeftCenter
        Impresion.Cell(10, 7).text = ciudad
        
        'GIRO
        Impresion.Range(12, 2, 13, 3).Merge
        Impresion.Range(12, 2, 13, 3).Alignment = cellLeftTop
        Impresion.Cell(12, 2).text = giro
        
        'COMUNA
        Impresion.Range(12, 7, 13, 7).Merge
        Impresion.Range(12, 7, 13, 7).Alignment = cellLeftTop
        Impresion.Cell(12, 7).text = comuna
        
        Impresion.AutoRedraw = True
        Impresion.Refresh
        
        Impresion.PageSetup.LeftMargin = 0.5
        Impresion.PageSetup.RightMargin = 0
        Impresion.PageSetup.TopMargin = 2.5
        Impresion.PageSetup.BottomMargin = 0
        
        Call verificaImpresora(3, Impresion)
        
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & numeroFactura)
    End If
End Sub

