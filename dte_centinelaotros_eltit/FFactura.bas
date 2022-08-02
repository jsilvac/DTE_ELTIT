Attribute VB_Name = "FFactura"
Option Explicit
    Private CAMPOS(30, 5) As String
    
    
    Public Sub imprimeFactura(ByVal numerofactura As String, ByRef documento As Grid, ByRef rollo As Adodc)
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim cad As String
    Dim h As Integer
    Dim totalprod As String
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim total As String
    Dim tpago As String
    Dim lineas As Integer
    Dim fecha As String
    Dim o As Integer
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
    Dim CAMPO1 As String * 11
    Dim campo2 As String * 11
    Dim campo3 As String * 11
    Dim campo4 As String * 11
    Dim campo5 As String * 11
    Dim campo6 As String * 11
    Dim porce As Double
    Dim dife As Double
    Dim exento As Double
    Dim ilarefrescos As String
    Dim ilalicores As String
    Dim ilavinos As String
    Dim dona As String
    Dim foliofiscal As String
    Dim rut As String
    Dim horaventafactura As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
   
    
    documento.Rows = 1
    
   Set csql.ActiveConnection = ventasRubro
    csql.sql = "SELECT dd.codigo, dd.descripcion,  dd.cantidad,  dd.precio, dd.cantidad*dd.precio, dd.total AS totalpro, dd.precio, dd.cantidad, dc.rut, dc.sucursal, dc.neto as neto, dc.iva, dc.impuestoharina AS iha, dc.impuestocarne AS ica, dc.impuestoilarefrescos as ilarefrescos, dc.impuestoilalicores as ilalicores , dc.impuestoilavinos AS ilavinos, dc.total, dc.fecha ,dc.descuento,dd.descuento as descuento2,dc.donacion,dc.foliosii,dc.numero,dd.tipodespacho,dd.vendedor,dc.horaventas "
    csql.sql = csql.sql & "from sv_documento_cabeza_" + empresaActiva + " AS dc, sv_documento_detalle_" + empresaActiva + " AS dd "
    csql.sql = csql.sql & "WHERE dd.caja=dc.caja and dc.fecha = dd.fecha and dd.local = '" & empresaActiva & "' AND dd.caja='" + PVentas.dato30.text + "' and dc.local = dd.local AND dd.tipo = 'FV' AND dc.foliosii = '" & numerofactura & "' AND dd.tipo = dc.tipo AND dd.numero = dc.numero and dd.fecha='" & PVentas.dato5.text & "-" & PVentas.dato4.text & "-" & PVentas.dato3.text & "' ORDER BY dd.linea ASC "
    csql.Execute
         
    If csql.RowsAffected > 0 Then

     Set resultados = csql.OpenResultset
        exento = 0
        documento.Rows = 1
        documento.Cols = 6
        documento.Rows = 63
        
        documento.DefaultFont.Name = "arial"
      
        horaventafactura = resultados("horaventas")
        documento.DefaultFont.Size = 8
        documento.DefaultFont.Bold = False
        documento.Column(0).Width = 0
        documento.Column(1).Width = 150
        documento.Column(2).Width = 90
        documento.Column(3).Width = 265
        documento.Column(4).Width = 100
        documento.Column(5).Width = 150
        documento.Column(1).Alignment = cellRightCenter
        documento.Column(2).Alignment = cellCenterCenter
        documento.Column(3).Alignment = cellLeftCenter
        documento.Column(4).Alignment = cellRightCenter
        documento.Column(5).Alignment = cellRightCenter
  '
        documento.DefaultRowHeight = 15
        
        documento.PageSetup.PrintGridlines = False
        documento.AutoRedraw = False
'    Grid2.PageSetup.PrintGridlines = False
'    Grid2.AutoRedraw = False


        
        j = 15
'        vendedor = Resultados("vendedor")
        total = resultados("total")
        fecha = resultados("fecha")
'        c.cc.rut = resultados("rut")
        rut = resultados("rut")
'        c.cc.sucursal = resultados("sucursal")
        totNeto = resultados("neto")
        totIva = resultados("iva")
        totIha = resultados("iha")
        totIca = resultados("ica")
'        totIla = resultados("ila")
        ilarefrescos = resultados("ilarefrescos")
        ilalicores = resultados("ilalicores")
        ilavinos = resultados("ilavinos")
        totaldescuento = resultados("descuento")
        porcedescuento = resultados("descuento2")
        exento = CDbl(ilarefrescos) + CDbl(ilalicores) + CDbl(ilavinos) + CDbl(totIca) + CDbl(totIha)
        
'        Call leerClienteCliente(c, "=")
        Descuento = "0"
        dona = resultados("donacion")
        
        
       
        Rem NUMERO
'        Documento.Cell(3, 5).Alignment = cellRightCenter
        documento.Cell(4, 1).text = "CAJA:" + PVentas.dato30.text + " F/O:" + resultados("numero")
                
        Rem Documento.Cell(4, 5).text = "F/O: " & NUMEROFACTURA
        documento.Range(3, 2, 3, 3).Merge
        documento.Range(3, 2, 3, 3).Alignment = cellCenterCenter
        documento.Cell(3, 2).text = leerNombreEmpresa(empresaActiva)
         
        documento.Cell(3, 1).text = "F/F: " & resultados("foliosii")
        foliofiscal = resultados("foliosii")
        'SEÑORES
        documento.Range(7, 2, 7, 3).Merge
        documento.Range(7, 2, 7, 3).Alignment = cellLeftCenter
        
        documento.Cell(7, 2).text = PVentas.lblRazon.Caption
           'FECHA
'        Documento.Range(7, 1, 7, 3).Merge
'        Documento.Range(7, 1, 7, 3).Alignment = cellLeftCenter
        documento.Cell(7, 5).text = fecha
          
        'DIRECCION
        documento.Range(9, 2, 9, 3).Merge
        documento.Range(9, 2, 9, 3).Alignment = cellLeftCenter
        documento.Cell(9, 2).text = PVentas.LBLDIRECCION.Caption
        
        
      'RUT
'        Documento.Range(8, 4, 8, 5).Merge
'        Documento.Range(8, 4, 8, 5).Alignment = cellLeftCenter
'        Documento.Cell(9, 5).Alignment = cellCenterCenter
        documento.Cell(9, 5).text = "     " + Format(PVentas.dato6.text, "###,###,##0") & "-" & PVentas.lbldv.Caption
 
        'GIRO
        documento.Range(11, 2, 11, 3).Merge
        documento.Range(11, 2, 11, 3).Alignment = cellLeftCenter
        documento.Cell(11, 2).text = leerGiroCliente(PVentas.dato6.text & PVentas.lbldv.Caption, PVentas.dato7.text)
        
        Rem tipo pago
'        tipopago = leerpago("FV", NUMEROFACTURA)
'

        'CIUDAD
'        Documento.Range(10, 4, 10, 5).Merge
'        Documento.Range(11, 5, 11, 5).Alignment = cellCenterCenter
        documento.Cell(11, 5).text = "     " + PVentas.LBLCIUDAD.Caption
        
        'DESCUENTO
'        Documento.Range(10, 4, 10, 5).Merge
'        Documento.Range(10, 4, 10, 5).Alignment = cellLeftCenter
'        Documento.Cell(10, 4).text = "D:" + Descuento
'
        
        
        lineas = 16
        While Not resultados.EOF
            Descuento = Str(CDbl(Descuento) + Int(resultados("cantidad") * resultados("precio")) - Int(resultados("totalpro")))
            Descuento = Mid(Descuento, 2, Len(Descuento))
            
            lineas = lineas + 1
            
            
            documento.Cell(lineas, 1).text = resultados(0)
            dife = CDbl(Int(resultados(2)) - resultados(2))
            If dife <> 0 Then
             documento.Cell(lineas, 2).text = Format(resultados(2), "###,##0.000")
            Else
             documento.Cell(lineas, 2).text = Format(resultados(2), "###,###")
            End If
            If resultados("tipodespacho") <> "" Then
            documento.Cell(lineas, 3).text = resultados(1) & " DESP. " & resultados("tipodespacho") & "-" & leerNombreTipoDespacho(resultados("tipodespacho"))
            Else
            documento.Cell(lineas, 3).text = resultados(1)
            End If
            
            documento.Cell(lineas, 4).text = Format(resultados(3), " $ ###,###,###")
            documento.Cell(lineas, 5).text = Format(resultados(4), " $ ###,###,###")
            
            resultados.MoveNext
        Wend
        MONTO = WORDNUM(Format(total, "########0"), "PESO", "PESOS", 0)
        
        Rem monto = numToLet(Format(total, "########0"), "PESO", "PESOS", 0)
        If totaldescuento <> 0 Then
        porce = totaldescuento / (CDbl(total) + CDbl(totaldescuento)) * 100
        'Documento.Cell(28, 4).text = "(-)%" + Format(PORCE, "##")
        documento.Cell(47, 4).Alignment = cellLeftCenter
        documento.Cell(47, 4).text = "  DESC"
        documento.Cell(47, 5).text = Format(totaldescuento * -1, " $ ###,###,##0")
        
        
        End If
        Set csql = Nothing
        csql.Close
        Set resultados = Nothing
          
'        Documento.Range(30, 1, 30, Documento.Cols - 1).Merge
'        Documento.Range(30, 1, 30, Documento.Cols - 1).Alignment = cellLeftCenter
'        Documento.Cell(30, 1).text = "    " + monto
'        Documento.Range(31, 1, 31, 2).Merge
        
'        If CDbl(porcedescuento) <> 0 Then
'        Documento.Cell(31, 1).text = "Dcto " + porcedescuento + "% N/Inc *OFE*"
'        End If
        

        documento.Cell(48, 4).Alignment = cellLeftCenter
        documento.Cell(48, 4).text = "  NETO"
        documento.Cell(48, 5).text = Format(totNeto, " $ ####,##0")
        
        documento.Cell(49, 4).Alignment = cellLeftCenter
        documento.Cell(49, 4).text = "  IVA"
        documento.Cell(49, 5).text = Format(totIva, " $ ###,##0")
        
        documento.Cell(50, 4).Alignment = cellLeftCenter
        'ariel cambia palabara exento por "otros impuetos"
        documento.Cell(50, 4).text = "  Otros impuestos"
        documento.Cell(50, 5).text = Format(exento, " $ ###,###,##0")
        
        documento.Cell(51, 4).Alignment = cellLeftCenter
        documento.Cell(51, 4).text = "  TOTAL"
        documento.Cell(51, 5).text = Format(total, " $ ###,###,##0")
        
        If CDbl(dona) > 0 Then
        documento.Cell(52, 4).Alignment = cellLeftCenter
        documento.Cell(52, 4).text = entidaddonacion
        
        documento.Cell(52, 5).text = Format(dona, " $ ###,###,##0")
        End If
        
        documento.Range(46, 1, 48, 3).Merge
        documento.Range(46, 1, 48, 3).Alignment = cellLeftCenter
        documento.Cell(46, 1).text = "                    " + MONTO
        documento.Range(46, 1, 48, 3).WrapText = True

   If ilarefrescos + ilavinos + ilalicores + totIha + totIca <> 0 Then
'
        documento.Range(55, 2, 55, 5).Merge
        documento.Range(55, 2, 55, 5).Alignment = cellLeftCenter
        documento.Range(55, 2, 55, 5).FontBold = True
        documento.Cell(55, 2).text = "      ILA 13     " + "      ILA 15     " + "      ILA 27     " + "      HARINA     " + "      CARNE      "
      
      

    'PALABRAIMPUESTO = CAMPO1(1) + CAMPO1(2) + CAMPO1(3) + CAMPO1(4) + CAMPO1(5)
    'Grid2.Cell(59, 2).text = PALABRAIMPUESTO
 

      
        
        CAMPO1 = String(10 - Len(ilarefrescos), 32) & Format(ilarefrescos, "####,##0")
        campo2 = String(10 - Len(ilavinos), 32) & Format(ilavinos, "####,##0")
        
        campo3 = String(10 - Len(ilalicores), 32) & Format(ilalicores, "####,##0")
       
        campo4 = String(10 - Len(totIha), 32) & Format(totIha, "####,##0")
        
        campo5 = String(10 - Len(totIca), 32) & Format(totIca, "####,##0")
       
'        campo6 = String(7 - Len(total), 32) & Format(total, "####,##0")

'       Documento   = "   ILA 13     " + "   ILA 15     " + "   ILA 27     " + "   HARINA     " + "   CARNE      "
               cad = "  " & CAMPO1 & "          " & campo2 & "          " & campo3 & "          " & campo4 & "          " & campo5
     
'        Documento.Range(59, 2, 59, Documento.Cols - 1).Merge
'        Documento.RowHeight(59) = 20
         documento.Range(56, 2, 56, 5).Merge
         documento.Range(56, 2, 56, 5).Alignment = cellLeftCenter
         documento.Range(56, 2, 56, 5).FontBold = True
         documento.Cell(56, 2).text = cad
        End If
        Call leercredito("FV", PVentas.dato30.text, foliofiscal, rut)
        
        If cantidadcuotas <> "" Then
            documento.Range(56, 1, 56, 4).Merge
            documento.Range(57, 1, 57, 4).Merge
            documento.Range(58, 1, 58, 4).Merge
            documento.Range(59, 1, 59, 4).Merge
            documento.Range(60, 1, 60, 4).Merge
            documento.Range(61, 1, 61, 4).Merge
            documento.Range(62, 1, 62, 4).Merge

        documento.Cell(56, 1).Border(cellEdgeTop) = cellThick
        documento.Cell(56, 1).Alignment = cellLeftCenter
        documento.Cell(57, 1).Alignment = cellLeftCenter
        documento.Cell(58, 1).Alignment = cellLeftCenter
        documento.Cell(59, 1).Alignment = cellLeftCenter
        documento.Cell(60, 1).Alignment = cellLeftCenter
        documento.Cell(61, 1).Alignment = cellLeftCenter
        documento.Cell(62, 1).Alignment = cellLeftCenter
        
        documento.Cell(56, 1).text = "FIRMA :"
        documento.Cell(57, 1).text = "Yo " + leerNombreCliente(rutcredito) + ""
        documento.Cell(58, 1).text = "CI." + Mid(rutcredito, 1, 9) + "-" + Mid(rutcredito, 10, 1) + " autorizo segun contrato PALGUIN LTDA " + "Cargar a mi cuenta " + cantidadcuotas + " cuotas de " + Format(montocuotas, "$ ###,###,###")
        documento.Cell(59, 1).text = "Primer Vencimiento " + Format(primervencimiento, "dd-mm-yyyy")
        documento.Cell(60, 1).text = "TOTAL CREDITO :" & Format(montocredito2, "$ ###,###,###") & " PIE :" & Format(CDbl(total) - CDbl(montocredito2), "$ ###,###,##0")
        End If
                
        
 
        documento.Range(2, 1, 2, 5).Merge
        documento.Range(2, 2, 2, 5).Alignment = cellCenterCenter
        If horaventafactura <> "" Then
            documento.Cell(2, 1).text = "cajero(a);" + PVentas.lblcajera.Caption + " HORA :" & horaventafactura
        Else
            documento.Cell(2, 1).text = "cajero(a);" + PVentas.lblcajera.Caption + " HORA :" & Time
        End If
               
         
        documento.AutoRedraw = True
        documento.Refresh
        
        documento.PageSetup.LeftMargin = 0.25
        documento.PageSetup.RightMargin = 0
        documento.PageSetup.TopMargin = 3
        documento.PageSetup.BottomMargin = 0
        
        For i = 1 To documento.PageSetup.PaperSizes.Count
            If UCase(documento.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
                documento.PageSetup.PaperSize = documento.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        
 
        documento.PageSetup.PrintGridlines = False
         If imprimeDirecto = True Then
        documento.DirectPrint
        Else
        documento.PrintPreview
        End If
'        cantidadCUOTAS = ""
'        montocuotas = ""
'        rutcredito = ""
'        primervencimiento = ""
'
    End If
End Sub
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

'Public Sub imprimeFactura(ByVal TIPO As String, ByVal NUMEROFACTURA As String, ByRef Documento As Grid, ByRef rollo As Adodc)
'    Dim i As Integer
'    Dim descuento As String
'    Dim neto As String
'    Dim piva As String
'    Dim piha As String
'    Dim total As String
'    Dim fecha As String
'    Dim vencimiento As String
'    Dim vendedor As String
'    Dim notapedido As String
'    Dim nombre As String
'    Dim rut As String
'    Dim direccion As String
'    Dim ciudad As String
'    Dim comuna As String
'    Dim giro As String
'    Dim fono As String
'    Dim CODIGO As String
'    Dim tiposDePago As String
'    Dim tabla As String
'    Dim c As Cliente
'    Dim LINEAS As Double
'    Dim condiciones As String
'    Dim transporte As String
'    Dim K As Integer
'    Dim cSql As New rdoQuery
'    Dim resultados As rdoResultset
'    Dim cSql1 As New rdoQuery
'    Dim resultados1 As rdoResultset
'    Dim cSql2 As New rdoQuery
'    Dim resultados2 As rdoResultset
'
'
'
'    Set cSql.ActiveConnection = ventasRubro
'    cSql.sql = "SELECT d.cantidad,d.codigo, d.descripcion, d.precio, d.descuento as descuento2,d.total AS totalpro, c.neto, c.iva, c.impuestoharina, c.total, DATE_FORMAT(c.fecha,'%d-%m-%Y') AS fecha, c.descuento, c.rut, c.sucursal, c.cajera, DATE_FORMAT(c.vencimiento,'%d-%m-%Y') AS vencimiento, IF(notapedido = '0000000000','',notapedido) AS notapedido "
'    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS c INNER JOIN sv_documento_detalle_" + empresaActiva + " AS d ON c.local = d.local AND c.tipo = d.tipo AND c.numero = d.numero INNER JOIN " & basedatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON d.codigo = mpf.codigobarra "
'    cSql.sql = cSql.sql & "WHERE c.local = '" & empresaActiva & "' AND c.tipo = '" & TIPO & "' AND c.foliosii = '" & NUMEROFACTURA & "' and caja='" & caja & "' "
'    cSql.sql = cSql.sql & "ORDER BY d.linea"
'    cSql.Execute
''    Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
'
'    If cSql.RowsAffected > 0 Then
'    Set resultados = cSql.OpenResultset
'
'        Documento.Rows = 1
'
'        Documento.Rows = 70
'        Documento.Cols = 7
'
'
'        Documento.DefaultFont.Name = "ARIAL"
'        Documento.DefaultFont.Size = 8
'        Documento.DefaultFont.Bold = False
'
'        Documento.Column(0).Width = 0
'        Documento.Column(1).Width = 100
'        Documento.Column(2).Width = 100
'        Documento.Column(3).Width = 300
'        Documento.Column(4).Width = 90
'        Documento.Column(5).Width = 50
'        Documento.Column(6).Width = 100
'
'
'        Documento.Column(1).Alignment = cellRightCenter
'        Documento.Column(2).Alignment = cellRightCenter
'        Documento.Column(3).Alignment = cellLeftCenter
'        Documento.Column(4).Alignment = cellRightCenter
'        Documento.Column(5).Alignment = cellRightCenter
'        Documento.Column(6).Alignment = cellRightCenter
'
'        Documento.DefaultRowHeight = 13
'        Rem imprime lineas
'
'        Documento.PageSetup.PrintGridlines = False
'
'        Documento.AutoRedraw = False
'
'
'        Call LEERCLIENTE(c, resultados("rut"), resultados("sucursal"), "=")
'
'        fecha = resultados("fecha")
'        nombre = "       " & c.nombre
'        For K = 1 To 2
'        If Mid(c.rut, K, 1) = "0" Then Mid(c.rut, K, 1) = " "
'        Next K
'
'        rut = "       " & Left(c.rut, Len(c.rut) - 1) & "-" & Right(c.rut, 1)
'        direccion = "       " & c.direccion
'        ciudad = "       " & c.ciudad
'        comuna = c.comuna
'        giro = c.giro
'        fono = c.fono1
'        vencimiento = resultados("vencimiento")
'        vendedor = resultados("cajera")
'        notapedido = resultados("notapedido")
'        neto = resultados("neto")
'        piva = resultados("iva")
'        piha = resultados("impuestoharina")
'        total = resultados("total")
'        descuento = resultados("descuento")
'        'descuento = "0"
'        condiciones = PVentas.dato8.text
'        transporte = PVentas.dato9.text
'
'        Documento.RowHeight(6) = 17
'        Documento.RowHeight(7) = 17
'        Documento.RowHeight(8) = 17
'        Documento.RowHeight(9) = 17
'        Documento.RowHeight(10) = 17
''        documento.RowHeight(11) = 15
''
''        documento.RowHeight(15) = 15
'         'NUMERO
'        Documento.Range(1, 5, 1, 6).Merge
'
'        Documento.Cell(1, 5).Alignment = cellLeftGeneral
'        Documento.Cell(1, 5).Font.Size = 10
'        Documento.Cell(1, 5).Font.Bold = True
'
'
'
'        Documento.Cell(1, 5).text = Mid(NUMEROFACTURA, 4, 7)
'        'EMPRESA
'
'        Documento.Range(2, 2, 2, 3).Merge
'        Documento.Range(2, 2, 2, 3).Alignment = cellLeftGeneral
'        Documento.Cell(2, 2).text = leerNombreEmpresa(empresaActiva)
'
'
'        'FECHA
'        Documento.Range(4, 2, 4, 3).Merge
'        Documento.Range(4, 2, 4, 3).Alignment = cellLeftCenter
'
'        Documento.Cell(4, 2).text = Format(fecha, "dd") + "                      " + MonthName(Format(fecha, "mm")) + "                                              " + Format(fecha, "yyyy")
'
'        'SEÑORES
'        Documento.Range(6, 2, 6, 3).Merge
'        Documento.Range(6, 2, 6, 3).Alignment = cellLeftCenter
'        Documento.Cell(6, 2).text = nombre
'
'        'RUT
'        Documento.Range(7, 2, 7, 3).Merge
'        Documento.Range(7, 2, 7, 3).Alignment = cellLeftCenter
'        Documento.Cell(7, 2).text = direccion
'
'        'DIRECCION
'        Documento.Range(8, 2, 8, 3).Merge
'        Documento.Range(8, 2, 8, 3).Alignment = cellLeftCenter
'        Documento.Cell(8, 2).text = giro
'
'        'CONDICIONES DE PAGO
'        Documento.Range(9, 2, 9, 3).Merge
'        Documento.Range(9, 2, 9, 3).Alignment = cellLeftCenter
'        Documento.Cell(9, 2).text = "                " + condiciones
'
'        'VENDEDOR
'        Documento.Range(10, 2, 10, 3).Merge
'        Documento.Range(10, 2, 10, 3).Alignment = cellLeftCenter
'        Documento.Cell(10, 2).text = "      " + vendedor + " " + PVentas.lblvendedor.Caption
'
'
'
'        'rut
'        Documento.Range(6, 5, 6, 6).Merge
'        Documento.Range(6, 5, 6, 6).Alignment = cellLeftCenter
'        Documento.Cell(6, 5).text = rut
'        'COMUNA
'        Documento.Range(7, 5, 7, 6).Merge
'        Documento.Range(7, 5, 7, 6).Alignment = cellLeftCenter
'        Documento.Cell(7, 5).text = ciudad
'        'FONO
'        Documento.Range(8, 5, 8, 6).Merge
'        Documento.Range(8, 5, 8, 6).Alignment = cellLeftCenter
'        Documento.Cell(8, 5).text = fono
'        'notapdedido
'        Documento.Range(10, 5, 10, 6).Merge
'        Documento.Range(10, 5, 10, 6).Alignment = cellLeftCenter
'        Documento.Cell(10, 5).text = transporte
'        ' LINEAS DE FACTURA
'
'         LINEAS = 14
'
'
'        While Not resultados.EOF
'            LINEAS = LINEAS + 1
'
'            Documento.Cell(LINEAS, 1).text = resultados("cantidad")
'            Documento.Cell(LINEAS, 2).text = resultados("codigo")
'            Documento.Cell(LINEAS, 3).text = "                  " + resultados("descripcion")
'            Documento.Cell(LINEAS, 4).text = Format(resultados("precio"), "###,###,###")
'            Documento.Cell(LINEAS, 5).text = Format(resultados("descuento2"), "% ##.00")
'            Documento.Cell(LINEAS, 6).text = Format(resultados("totalpro"), "###,###,###")
'
'            Documento.Range(Documento.Rows - 1, 3, Documento.Rows - 1, 4).Merge
'            resultados.MoveNext
'        Wend
'         Set cSql = Nothing
'         cSql.Close
'         Set resultados = Nothing
'
'
'        'TIPOS DE PAGO
'        Set cSql1.ActiveConnection = ventasRubro
'        cSql1.sql = "SELECT dp.tipopago, SUM(dp.monto) AS monto "
'        cSql1.sql = cSql1.sql & "FROM sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero "
'        cSql1.sql = cSql1.sql & "WHERE dp.tipo = '" & TIPO & "' AND dp.numero = '" & NUMEROFACTURA & "' "
'        cSql1.sql = cSql1.sql & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
'        cSql1.Execute
'        'Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
'
'        If cSql1.RowsAffected > 0 Then
'            Set resultados1 = cSql1.OpenResultset
'            tiposDePago = ""
'            While Not resultados1.EOF
'                CODIGO = resultados1("tipopago")
'                Select Case CODIGO
'                    Case "1"    'EFECTIVO
'                        If CDbl(resultados1("monto")) <> 0 Then
'                            tiposDePago = tiposDePago & "EFECTIVO " & Format(resultados1("monto"), "$ ###,###,##0") & " / "
'                        End If
'                    Case "2", "3"   'CHEQUE
'                        tiposDePago = tiposDePago & "CHEQUE " & Format(resultados1("monto"), "$ ###,###,##0") & " / "
'                    Case "4"    'CREDITO DIRECTO
'                        tiposDePago = tiposDePago & "CREDITO DIRECTO " & Format(resultados1("monto"), "$ ###,###,##0") & " / "
'                End Select
'                resultados1.MoveNext
'            Wend
'        End If
'        Set cSql1 = Nothing
'        cSql1.Close
'        Set resultados1 = Nothing
'
'
'        For i = Documento.Rows To 70
'        Documento.AddItem ""
'        Next i
'
'        'CHEQUES
'        Set cSql2.ActiveConnection = ventasRubro
'        cSql2.sql = "SELECT dp.numerodocumento, IFNULL(b.nombre,'') AS nombre, dp.monto, DATE_FORMAT(dp.vencimiento,'%d-%m-%Y') AS vencimiento "
'        cSql2.sql = cSql2.sql & "FROM sv_documento_pagos_" + empresaActiva + " AS dp LEFT JOIN " & baseVentas & ".sv_maestrobancos AS b ON dp.banco = b.codigobanco "
'        cSql2.sql = cSql2.sql & "WHERE dp.local = '00' AND dp.tipopago = '2' AND dp.tipo = '" & TIPO & "' AND dp.numero = '" & NUMEROFACTURA & "' "
'        cSql2.sql = cSql2.sql & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
'        cSql2.Execute
'        'Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
'        Documento.Range(45, 3, 45, 6).Merge
'        Documento.Range(45, 3, 45, 6).Alignment = cellCenterCenter
'        Documento.Cell(45, 3).text = "  ** REVISADO POR:" + PVentas.dato20.text + " BULTOS :" + PVentas.dato21.text + " **"
'        Documento.Cell(46, 2).text = tiposDePago
'        Documento.Range(46, 2, 46, 6).Merge
'        Documento.Range(46, 2, 46, 6).Alignment = cellLeftCenter
'
'        If cSql2.RowsAffected > 0 Then
'           Set resultados2 = cSql2.OpenResultset
'            i = 47
'            While Not resultados2.EOF
'                Documento.Cell(i, 2).text = resultados2("numerocheque")
'                Documento.Cell(i, 3).text = resultados2("nombre")
'                Documento.Cell(i, 5).text = Format(resultados2("monto"), "$ ###,###,##0")
'                Documento.Cell(i, 6).text = resultados2("vencimiento")
'                i = i + 1
'                resultados2.MoveNext
'            Wend
'        End If
'        Set cSql2 = Nothing
'        cSql2.Close
'        Set resultados2 = Nothing
'
'        Documento.RowHeight(50) = 17
'        Documento.RowHeight(51) = 17
'        Documento.RowHeight(52) = 17
'        Documento.RowHeight(53) = 17
'        Documento.RowHeight(54) = 17
'
'        Documento.Cell(50, 2).text = numToLet(total, "PESO", "PESOS", "CENTAVO", "CENTAVOS", 0)
'        Documento.Range(50, 2, 50, 5).Merge
'        Documento.Range(50, 2, 50, 5).Alignment = cellLeftCenter
'        If descuento <> 0 Then
'        Documento.Cell(48, 4).text = "DESCUENTO"
'        Documento.Cell(48, 4).Alignment = cellCenterCenter
'        Documento.Cell(48, 6).text = Format(descuento * -1, "$ ###,###,##0")
'        End If
'        Documento.AddItem ""
'        Documento.Cell(50, 6).text = Format(neto, "$ ###,###,##0")
'
'        Documento.AddItem ""
'        Documento.Cell(52, 4).text = Str(iva) & "%"
'        Documento.Cell(52, 4).Alignment = cellCenterCenter
'        Documento.Cell(52, 6).text = Format(piva, "$ ###,###,##0")
'
'        Documento.AddItem ""
'        Documento.Cell(54, 6).text = Format(total, "$ ###,###,##0")
'
'        Documento.AutoRedraw = True
'        Documento.Refresh
'
'        Documento.PageSetup.LeftMargin = 0.5
'        Documento.PageSetup.RightMargin = 0
'        Documento.PageSetup.TopMargin = 2.5
'        Documento.PageSetup.BottomMargin = 0
'
'
'
'        For i = 1 To Documento.PageSetup.PaperSizes.Count
'            If UCase(Documento.PageSetup.PaperSizes.Item(i).PaperName) = "FACTURA" Then
'                Documento.PageSetup.PaperSize = Documento.PageSetup.PaperSizes.Item(i).Kind
'                Exit For
'            End If
'        Next i
'
'        If TIPO = "NV" Then
'
'            Call verificaImpresora(4, Documento)
'
'        Else
'
'            Call verificaImpresora(1, Documento)
'
'        End If
'    Else
'        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & numeroFactura)
'    End If
'End Sub

'numToLet(Me.txtUserName.text, "PESO", "PESOS", "CENTAVO", "CENTAVOS", 0)
Function numToLet(ByVal NUMERO As Variant, Optional TipoCambioSingular As String = "", Optional TipoCambioPlural As String, Optional subTipoCambioSingular As String, Optional subTipoCambioPlural As String, Optional xInternal As Long = 0) As String
     Dim snum As String, vNum() As String, x As Long, Y As Long, Z As Long, sTmp As String
     Dim D1 As String, D2 As String, D3 As String, D4 As String, DFinal As String
     Dim tNum As String, B1 As Boolean, B2 As Boolean, B3 As Boolean
     Dim wNum() As String, xNums As String, xWords As String, Nombres() As String
     
     '***********************************************************************************************
     '* Esta función convierte números en palabras, sin importar el contexto donde se encuentren    *
     '* La presición (por limitancia del lenguaje) es de 28B, Ej: 9999999999999999999999999999 max. *
     '***********************************************************************************************
     
     'Convierte el valor en un string
     snum = Trim(CStr(NUMERO))
            
     'Procesa cada número que exista en la variable por separado
     If xInternal = 0 Then
        'Separa los números limpios de las palabras y los procesa por separado (no incluye números con letras)
        wNum = Split(snum, " ")
        For x = 0 To UBound(wNum)
            'Concatena los strings o números según corresponda
            If IsNumeric(wNum(x)) Then
               'Separa los enteros de los decimales para procesarlos por separado
               If Int(Val(wNum(x))) < wNum(x) Then
                  D1 = Int(Val(wNum(x)))
                  D2 = Mid(CStr(wNum(x)), Len(D1) + 2)
                  DFinal = DFinal & IIf(D1 < 0, "menos ", "") & numToLet(D1, TipoCambioSingular, TipoCambioPlural, 1) & " con "
                  DFinal = DFinal & numToLet(D2, subTipoCambioSingular, subTipoCambioPlural, , , 1) & " "
               Else
                  DFinal = DFinal & IIf(wNum(x) < 0, "menos ", "") & numToLet(wNum(x), TipoCambioSingular, TipoCambioPlural, subTipoCambioSingular, subTipoCambioPlural, 1) & " "
               End If
            Else
               DFinal = DFinal & wNum(x) & " "
            End If
        Next
     Else
        
        'ELimina el signo
        If Not IsNumeric(Left(snum, 1)) Then
           snum = Mid(snum, 2)
        End If
     
        'Elimina cualquier formato posible (incluye valores científicos)
        snum = Format(snum, "0")
        
        'Completa con ceros a la izquierda hasta obtener una longitud múltiplo de 3
        Do While Len(snum) Mod 3 <> 0
           snum = "0" & snum
        Loop
     
        'Dimenciona un arreglo con espacio para cada una de las centenas
        ReDim vNum(Len(snum) / 3 - 1)
        
        'Carga el arreglo con las centenas que corresponda
        For x = 0 To UBound(vNum, 1)
            vNum(x) = Mid(snum, (x + 1) * 3 - 2, 3)
        Next
         
        'Si el arreglo contiene una sola centena, la convierte en palabras
        If UBound(vNum, 1) = 0 Then
            'Asigna los dígitos de la centena y recuerda si son mayores que cero
            D3 = Left(snum, 1): B3 = Val(D3) > 0
            D2 = Mid(snum, 2, 1): B2 = Val(D2) > 0
            D1 = Right(snum, 1): B1 = Val(D1) > 0
            
            'Procesa las unidades
            Select Case D1
                   Case "1": DFinal = "un"
                   Case "2": DFinal = "dos"
                   Case "3": DFinal = "tres"
                   Case "4": DFinal = "cuatro"
                   Case "5": DFinal = "cinco"
                   Case "6": DFinal = "seis"
                   Case "7": DFinal = "siete"
                   Case "8": DFinal = "ocho"
                   Case "9": DFinal = "nueve"
            End Select
            
            'Procesa las decenas
            Select Case D2
                   Case "1"
                        'Maneja lógica del retrasado mental que puso nombres ilógicos a algunos números.
                        Select Case D1
                               Case "0": DFinal = "diez"
                               Case "1": DFinal = "once"
                               Case "2": DFinal = "doce"
                               Case "3": DFinal = "trece"
                               Case "4": DFinal = "catorce"
                               Case "5": DFinal = "quince"
                               Case "6": DFinal = "dieciséis"
                               Case Else
                                    DFinal = "dieci" & DFinal
                        End Select
                   Case "2"
                        If B1 Then
                           If D1 = "2" Then DFinal = "dós"
                           If D1 = "3" Then DFinal = "trés"
                           DFinal = "veinti" & DFinal
                        Else
                           DFinal = "veinte"
                        End If
                   Case "3": If B1 Then DFinal = "treinta y " & DFinal Else DFinal = "treinta"
                   Case "4": If B1 Then DFinal = "cuarenta y " & DFinal Else DFinal = "cuarenta"
                   Case "5": If B1 Then DFinal = "cincuenta y " & DFinal Else DFinal = "cincuenta"
                   Case "6": If B1 Then DFinal = "sesenta y " & DFinal Else DFinal = "sesenta"
                   Case "7": If B1 Then DFinal = "setenta y " & DFinal Else DFinal = "setenta"
                   Case "8": If B1 Then DFinal = "ochenta y " & DFinal Else DFinal = "ochenta"
                   Case "9": If B1 Then DFinal = "noventa y " & DFinal Else DFinal = "noventa"
            End Select
            
            'Procesa las centenas
            Select Case D3
                   Case "1": If B1 Or B2 Then DFinal = "ciento " & DFinal Else DFinal = "cien"
                   Case "2": If B1 Or B2 Then DFinal = "doscientos " & DFinal Else DFinal = "doscientos"
                   Case "3": If B1 Or B2 Then DFinal = "trescientos " & DFinal Else DFinal = "trescientos"
                   Case "4": If B1 Or B2 Then DFinal = "cuatrocientos " & DFinal Else DFinal = "cuatrocientos"
                   Case "5": If B1 Or B2 Then DFinal = "quinientos " & DFinal Else DFinal = "quinientos"
                   Case "6": If B1 Or B2 Then DFinal = "seiscientos " & DFinal Else DFinal = "seiscientos"
                   Case "7": If B1 Or B2 Then DFinal = "setecientos " & DFinal Else DFinal = "setecientos"
                   Case "8": If B1 Or B2 Then DFinal = "ochocientos " & DFinal Else DFinal = "ochocientos"
                   Case "9": If B1 Or B2 Then DFinal = "novecientos " & DFinal Else DFinal = "novecientos"
            End Select
            
            'Si es la ejecución principal efectua algunos arreglines
            If xInternal = 1 Then
               'Validación del cero
               If DFinal = "" Then DFinal = "cero"
               'Validación de terminados en "un"
               If Right(DFinal, 2) = "un" And TipoCambioSingular = "" Then DFinal = DFinal & "o"
            End If
            
        Else 'Si es más de una centena, las separa y procesa independientemente
            Y = -1
            Z = 1
            For x = UBound(vNum) To 0 Step -1
                Y = Y + 1
                
                'Convierte la centena en palabras
                tNum = numToLet(vNum(x), xInternal:=2)
                
                'Arregla la terminación "uno" cuando corresponde
                If Y = 0 And Right(tNum, 2) = "un" And TipoCambioSingular & TipoCambioPlural = "" Then tNum = tNum + "o"
                
                'Genera un valor temporal para poder modificar
                sTmp = tNum
                
                'Asigna los nombres genéricos principales
                Nombres = Split(" mil , millón , millones , billón , billones , trillón , trillones , cuatrillón , cuatrillones , quintillón , quintillones , sextillón , sextillones , septillón , septillones , octillón , octillones, nonillón , nonillones , decillón , decillones , undecillón , undecillones , duodecillón , duodecillones , tredecillón , tredecillones , cuatordecillón , cuatordecillones , quindecillón , quindecillones , sexdecillón , sexdecillones , septendecillón , septendecillones , octodecillón , octodecillones , novendecillón , novendecillones , vigintillón , vigintillones ", ",")
                
                'Controla que el índice de nombres no salga de los límites
                If Y > UBound(Nombres) Then
                   numToLet = "?"
                   Exit Function
                End If
                
                'Asigna los nombres correspondientes
                If Y Mod 2 > 0 Then
                   D1 = Nombres(0)
                   D2 = Nombres(Y - 1)
                ElseIf Y > 0 Then
                   D1 = Nombres(Y - 1)
                   D2 = Nombres(Y)
                Else
                   D1 = "": D2 = ""
                End If
                
                'Actualiza el nombre del número
                Select Case Y Mod 2
                       Case 0: If sTmp = "un" Then sTmp = sTmp & D1 Else sTmp = sTmp & IIf(tNum = "", "", D2)
                       Case Else
                            If sTmp = "un" Then sTmp = ""
                            sTmp = sTmp & IIf(tNum = "", "", D1)
                            If x = 0 And Y > 1 Then
                               If InStr(1, DFinal, D2, vbTextCompare) = 0 Then sTmp = sTmp & Mid(D2, 2)
                            End If
                End Select
                DFinal = sTmp & DFinal
            Next
        End If
     End If
     
     'Aplica el tipo de moneda cuando corresponda
     If xInternal = 1 Then DFinal = DFinal & " " & IIf(Format(snum, "#0") = "1", TipoCambioSingular, TipoCambioPlural)

     'Asigna el número en palabras
      numToLet = UCase(Trim(DFinal))
End Function

Public Sub imprimeNOTAPEDIDO(ByVal TIPO As String, ByVal numerofactura As String, ByRef documento As Grid, ByRef rollo As Adodc)
    Dim i As Integer
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
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
    Dim CODIGO As String
    Dim tiposDePago As String
    Dim tabla As String
    Dim c As Cliente
    Dim lineas As Double
    Dim cDP As String
    Dim transporte As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
     Dim objReportTitle As FlexCell.ReportTitle
    
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "SELECT d.cantidad,d.codigo, d.descripcion, d.precio, d.descuento as descuento2,d.total AS totalpro, c.neto, c.iva, c.impuestoharina, c.total, DATE_FORMAT(c.fecha,'%d-%m-%Y') AS fecha, c.descuento, c.rut, c.sucursal, c.cajera, DATE_FORMAT(c.vencimiento,'%d-%m-%Y') AS vencimiento, c.transporte, c.condicionesdepago "
    csql.sql = csql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS c INNER JOIN sv_documento_detalle_" + empresaActiva + " AS d ON c.local = d.local AND c.tipo = d.tipo AND c.numero = d.numero INNER JOIN " & basedatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON d.codigo = mpf.codigobarra "
    csql.sql = csql.sql & "WHERE c.local = '" & empresaActiva & "' AND c.tipo = '" & TIPO & "' AND c.numero = '" & numerofactura & "' "
    csql.sql = csql.sql & "ORDER BY d.linea"
    csql.Execute
'    Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
    
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
        documento.Rows = 1
        documento.Rows = 70
        documento.Cols = 11
                
        
        documento.DefaultFont.Name = "Arial"
        documento.DefaultFont.Size = 8
        documento.DefaultFont.Bold = False
        
        documento.Column(0).Width = 0
        documento.Column(1).Width = 90
        documento.Column(2).Width = 90
        documento.Column(3).Width = 200
        documento.Column(4).Width = 60
        documento.Column(5).Width = 40
        documento.Column(6).Width = 70
        documento.Column(7).Width = 70
        documento.Column(8).Width = 40
        documento.Column(9).Width = 40
        documento.Column(10).Width = 40
           

        
        documento.Column(1).Alignment = cellRightCenter
        documento.Column(2).Alignment = cellRightCenter
        documento.Column(3).Alignment = cellLeftCenter
        documento.Column(4).Alignment = cellRightCenter
        documento.Column(5).Alignment = cellRightCenter
        documento.Column(6).Alignment = cellRightCenter
        documento.Column(7).Alignment = cellLeftCenter
        documento.Column(8).Alignment = cellRightCenter
        documento.Column(9).Alignment = cellRightCenter
        documento.Column(10).Alignment = cellRightCenter
        
        documento.DefaultRowHeight = 13
        Rem imprime lineas
        
        documento.PageSetup.PrintGridlines = False
        
        documento.AutoRedraw = False
        
        rollo.Recordset.MoveFirst
        
        Call LEERCLIENTE(c, resultados("rut"), resultados("sucursal"), "=")
        
        fecha = resultados("fecha")
        nombre = c.nombre
        rut = Left(c.rut, Len(c.rut) - 1) & "-" & Right(c.rut, 1)
        direccion = c.direccion
        ciudad = c.ciudad
        comuna = c.comuna
        giro = c.giro
        fono = c.fono1
        vencimiento = resultados("vencimiento")
        vendedor = resultados("cajera")
        transporte = PVentas.dato9.text
        
        cDP = PVentas.dato8.text
        neto = resultados("neto")
        piva = resultados("iva")
        piha = resultados("impuestoharina")
        total = resultados("total")
        Descuento = resultados("descuento")
        'descuento = "0"
      
       
         'NUMERO
        
        documento.Range(1, 1, 1, 3).Merge
        documento.Range(1, 1, 1, 3).Alignment = cellLeftCenter
        documento.Cell(1, 1).text = leerNombreEmpresa(empresaActiva)
        documento.Range(2, 1, 2, 3).Merge
        documento.Range(2, 1, 2, 3).Alignment = cellLeftCenter
        documento.Cell(2, 1).text = leerDireccionEmpresa(empresaActiva)
        documento.Range(3, 1, 3, 3).Merge
        documento.Range(3, 1, 3, 3).Alignment = cellLeftCenter
        documento.Cell(3, 1).text = ""
        
        documento.Range(5, 1, 5, 6).Merge
        documento.Range(5, 1, 5, 6).Alignment = cellCenterCenter
        documento.Range(5, 1, 5, 6).FontSize = 10
        documento.Range(5, 1, 5, 6).FontBold = True
        
        
        
        documento.Cell(5, 1).text = "NOTA DE PEDIDO :" + numerofactura
        'EMPRESA
        
        
        'FECHA
        For i = 6 To 15
        documento.Range(i, 2, i, 3).Merge
        documento.Range(i, 2, i, 2).Alignment = cellLeftCenter
        
        
        documento.Range(i, 2, i, 2).FontBold = True
        documento.Range(i, 1, i, 1).Alignment = cellLeftCenter
        
        
        Next i
        
        
        documento.Cell(6, 1).text = "FECHA   "
        documento.Cell(6, 2).text = fecha
        
        documento.Cell(7, 1).text = "RUT "
        documento.Cell(7, 2).text = rut
        
                
        documento.Cell(8, 1).text = "NOMBRE "
        documento.Cell(8, 2).text = nombre
        
        documento.Cell(9, 1).text = "DIRECCION "
        documento.Cell(9, 2).text = direccion
        
        documento.Cell(10, 1).text = "CIUDAD   "
        documento.Cell(10, 2).text = ciudad
        
        documento.Cell(11, 1).text = "FONO     "
        documento.Cell(11, 2).text = fono
        
        documento.Cell(12, 1).text = "GIRO     "
        documento.Cell(12, 2).text = giro
        
        
        documento.Cell(13, 1).text = "VENDEDOR  "
        documento.Cell(13, 2).text = vendedor + " " + PVentas.lblVendedor.Caption
         
         documento.Cell(14, 1).text = "CONDICIONES"
        documento.Cell(14, 2).text = cDP
         documento.Cell(15, 1).text = "TRANSPORTE  "
        documento.Cell(15, 2).text = transporte


        lineas = 17
            
            
            documento.Range(lineas, 1, lineas, 6).Alignment = cellCenterGeneral
            For i = 1 To 10
            
            documento.Cell(lineas, i).Border(cellEdgeBottom) = cellThick
            documento.Cell(lineas, i).Border(cellEdgeTop) = cellThick
            documento.Cell(lineas, i).Border(cellEdgeLeft) = cellThick
            documento.Cell(lineas, i).Border(cellEdgeRight) = cellThick
            Next i
            
          
            documento.Cell(lineas, 1).text = "CANT."
            documento.Cell(lineas, 2).text = "CODIGO"
            documento.Cell(lineas, 3).text = "DESCRIPCION"
            documento.Cell(lineas, 4).text = "PRECIO"
            documento.Cell(lineas, 5).text = "Dcto"
            documento.Cell(lineas, 6).text = "TOTAL"
            documento.Cell(lineas, 7).text = "UBICACION"
            documento.Cell(lineas, 8).text = "SK B00"
            documento.Cell(lineas, 9).text = "SK B01"
            documento.Cell(lineas, 10).text = "SK OTRAS"
            
        
         lineas = 17
        
        
        While Not rollo.Recordset.EOF
            lineas = lineas + 1
            
            documento.Cell(lineas, 1).text = rollo.Recordset.Fields("cantidad")
            documento.Cell(lineas, 2).text = rollo.Recordset.Fields("codigo")
            documento.Cell(lineas, 3).text = Mid(resultados("descripcion"), 1, 50)
            documento.Cell(lineas, 4).text = Format(resultados("precio"), "###,###,###")
            documento.Cell(lineas, 5).text = Format(resultados("descuento2"), "% ##.00")
            documento.Cell(lineas, 6).text = Format(resultados("totalpro"), "###,###,###")
            documento.Cell(lineas, 7).text = leerubicacion(resultados("codigo"))
            documento.Cell(lineas, 8).text = leerstock(resultados("codigo"), "00")
            documento.Cell(lineas, 9).text = leerstock(resultados("codigo"), "01")
            documento.Cell(lineas, 10).text = leerstock(resultados("codigo"), "02") + leerstock(resultados("codigo"), "03") + leerstock(resultados("codigo"), "04") + leerstock(resultados("codigo"), "05") + leerstock(resultados("codigo"), "06") + leerstock(resultados("codigo"), "07")
  
            
            documento.Range(documento.Rows - 1, 3, documento.Rows - 1, 4).Merge
            resultados.MoveNext
        Wend
         Set csql = Nothing
         csql.Close
         Set resultados = Nothing
         
        
'        'TIPOS DE PAGO
'        tabla = "SELECT dp.tipopago, SUM(dp.monto) AS monto "
'        tabla = tabla & "FROM sv_documento_pagos AS dp INNER JOIN sv_documento_cabeza AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero "
'        tabla = tabla & "WHERE dp.tipo = '" & tipo & "' AND dp.numero = '" & numerofactura & "' "
'        tabla = tabla & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
'
'        Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
'
'        If rollo.Recordset.RecordCount > 0 Then
'            rollo.Recordset.MoveFirst
'            tiposDePago = ""
'            While Not rollo.Recordset.EOF
'                codigo = rollo.Recordset.Fields("tipopago")
'                Select Case codigo
'                    Case "1"    'EFECTIVO
'                        If CDbl(rollo.Recordset.Fields("monto")) <> 0 Then
'                            tiposDePago = tiposDePago & "EFECTIVO " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
'                        End If
'                    Case "2", "3"   'CHEQUE
'                        tiposDePago = tiposDePago & "CHEQUE " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
'                    Case "4"    'CREDITO DIRECTO
'                        tiposDePago = tiposDePago & "CREDITO DIRECTO " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
'                End Select
'                rollo.Recordset.MoveNext
'            Wend
'        End If
'
'
'        For i = documento.Rows To 70
'        documento.AddItem ""
'        Next i
'
'        'CHEQUES
'        tabla = "SELECT dp.numerocheque, IFNULL(b.nombre,'') AS nombre, dp.monto, DATE_FORMAT(dp.vencimiento,'%d-%m-%Y') AS vencimiento "
'        tabla = tabla & "FROM sv_documento_pagos AS dp LEFT JOIN " & baseVentas & ".sv_maestrobancos AS b ON dp.banco = b.codigobanco "
'        tabla = tabla & "WHERE dp.local = '00' AND dp.tipopago = '2' AND dp.tipo = '" & tipo & "' AND dp.numero = '" & numerofactura & "' "
'        tabla = tabla & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
'
'        Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
'
'        documento.Cell(46, 2).text = tiposDePago
'        documento.Range(46, 2, 46, 6).Merge
'        documento.Range(46, 2, 46, 6).Alignment = cellLeftCenter
'        If rollo.Recordset.RecordCount > 0 Then
'            rollo.Recordset.MoveFirst
'            i = 47
'            While Not rollo.Recordset.EOF
'                documento.Cell(i, 2).text = rollo.Recordset.Fields("numerocheque")
'                documento.Cell(i, 3).text = rollo.Recordset.Fields("nombre")
'                documento.Cell(i, 5).text = Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0")
'                documento.Cell(i, 6).text = rollo.Recordset.Fields("vencimiento")
'                i = i + 1
'                rollo.Recordset.MoveNext
'            Wend
'        End If
        
'        documento.Cell(50, 2).text = numToLet(total, "PESO", "PESOS", "CENTAVO", "CENTAVOS", 0)
'        documento.Range(50, 2, 50, 4).Merge
'        documento.Range(50, 2, 50, 4).Alignment = cellLeftCenter
            
            documento.Range(61, 4, 67, 6).Borders(cellEdgeBottom) = cellThick
            documento.Range(61, 4, 67, 6).Borders(cellEdgeTop) = cellThick
            documento.Range(61, 4, 67, 6).Borders(cellEdgeLeft) = cellThick
            documento.Range(61, 4, 67, 6).Borders(cellEdgeRight) = cellThick
            documento.Range(61, 1, 61, 10).Borders(cellEdgeTop) = cellThick
            
            
        If Descuento <> 0 Then
        documento.Cell(60, 4).text = "DESCUENTO"
        documento.Cell(60, 4).Alignment = cellCenterCenter
        documento.Cell(60, 6).text = Format(Descuento * -1, "$ ###,###,##0")
        End If
        
        documento.AddItem ""
        documento.Cell(61, 4).text = "NETO "
        documento.Cell(61, 4).Alignment = cellLeftCenter
        documento.Cell(61, 6).text = Format(neto, "$ ###,###,##0")
        
        documento.AddItem ""
        documento.Cell(62, 4).text = "I.V.A " + Str(iva) & "%"
        documento.Cell(62, 4).Alignment = cellLeftCenter
        documento.Cell(62, 6).text = Format(piva, "$ ###,###,##0")
        
        documento.AddItem ""
        documento.Cell(63, 4).text = "TOTAL "
        documento.Cell(63, 4).Alignment = cellLeftCenter
        
        documento.Cell(63, 6).text = Format(total, "$ ###,###,##0")
        
        documento.AutoRedraw = True
        documento.Refresh
        
        documento.PageSetup.LeftMargin = 0.5
        documento.PageSetup.RightMargin = 0
        documento.PageSetup.TopMargin = 2.5
        documento.PageSetup.BottomMargin = 0
'        documento.PageSetup.PrintGridlines = True
          
        
        
'        documento.Images.Add App.Path & "\SKORPIOS.BMP", "Logo"
'        Set objReportTitle = New FlexCell.ReportTitle
'        objReportTitle.ImageKey = "Logo"
'        documento.ReportTitles.Add objReportTitle

        
        
        For i = 1 To documento.PageSetup.PaperSizes.Count
            If UCase(documento.PageSetup.PaperSizes.Item(i).PaperName) = "FACTURA" Then
                documento.PageSetup.PaperSize = documento.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        
        If TIPO = "NV" Then
            
            Call verificaImpresora(4, documento)
            
        Else
        
            Call verificaImpresora(1, documento)
            
        End If
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & numeroFactura)
    End If
End Sub

Public Function leerubicacion(CODIGO) As String
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim numCaja As String
    Dim NUMERO As Double
    
    
    
    CAMPOS(0, 0) = "ubicacion"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "r_maestroproductos_stock_00"
    condicion = "local = '00' AND codigo='" + CODIGO + "' and bodega='01' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sql.response = CAMPOS
    Set sql.conexion = gestionRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
           
           leerubicacion = sql.response(0, 3)
    Else
           leerubicacion = "S/D"
    
    End If
End Function
Private Function leerstock(CODIGO, bodega)
    Dim a As Integer
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim saldo As Double

        Set csql.ActiveConnection = gestionRubro
        csql.sql = "SELECT stockactual "
        csql.sql = csql.sql + "FROM r_maestroproductos_stock_" & rubro & " "
        csql.sql = csql.sql + "WHERE año='" + Format(fechasistema, "yyyy") + "' AND codigo='" + CODIGO + "' AND bodega='" + bodega + "' "
        csql.Execute
       
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
           leerstock = resultados(0)
            resultados.Close
            Set resultados = Nothing
            Else
            leerstock = 0
        End If
       
End Function

Public Sub imprimeFactura2(ByVal numerofactura As String, ByRef documento As Grid, ByRef rollo As Adodc)
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim cad As String
    Dim h As Integer
    Dim totalprod As String
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim total As String
    Dim tpago As String
    Dim lineas As Integer
    Dim fecha As String
    Dim o As Integer
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
    Dim CAMPO1 As String * 11
    Dim campo2 As String * 11
    Dim campo3 As String * 11
    Dim campo4 As String * 11
    Dim campo5 As String * 11
    Dim campo6 As String * 11
    Dim porce As Double
    Dim dife As Double
    Dim exento As Double
    Dim ilarefrescos As String
    Dim ilalicores As String
    Dim ilavinos As String
    Dim dona As String
    Dim foliofiscal As String
    Dim rut As String
    
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
   
    
    documento.Rows = 1
    
   Set csql.ActiveConnection = ventasRubro
    csql.sql = "SELECT dd.codigo, dd.descripcion,  dd.cantidad,  dd.precio, dd.cantidad*dd.precio, dd.total AS totalpro, dd.precio, dd.cantidad, dc.rut, dc.sucursal, dc.neto as neto, dc.iva, dc.impuestoharina AS iha, dc.impuestocarne AS ica, dc.impuestoilarefrescos as ilarefrescos, dc.impuestoilalicores as ilalicores , dc.impuestoilavinos AS ilavinos, dc.total, dc.fecha ,dc.descuento,dd.descuento as descuento2,dc.donacion,dc.foliosii,dc.numero,dd.tipodespacho,dd.vendedor "
    csql.sql = csql.sql & "from sv_documento_cabeza_" + empresaActiva + " AS dc, sv_documento_detalle_" + empresaActiva + " AS dd "
    csql.sql = csql.sql & "WHERE dd.caja=dc.caja and dc.fecha=dd.fecha  and dc.local = dd.local  and dd.numero = dc.numero and dc.local = '" & empresaActiva & "' AND dc.caja='" & PVentas.dato30.text & "'  AND dd.tipo = 'FV' AND dc.numero = '" & numerofactura & "' AND dd.tipo = dc.tipo  and dd.fecha='" & Format(fechasistema, "yyyy-mm-dd") & "' ORDER BY dd.linea ASC "
    csql.Execute
         
    If csql.RowsAffected > 0 Then

     Set resultados = csql.OpenResultset
        exento = 0
        documento.Rows = 1
        documento.Cols = 6
        documento.Rows = 63
        
        documento.DefaultFont.Name = "arial"
      
        
        documento.DefaultFont.Size = 8
        documento.DefaultFont.Bold = False
        documento.Column(0).Width = 0
        documento.Column(1).Width = 150
        documento.Column(2).Width = 90
        documento.Column(3).Width = 265
        documento.Column(4).Width = 100
        documento.Column(5).Width = 150
        documento.Column(1).Alignment = cellRightCenter
        documento.Column(2).Alignment = cellCenterCenter
        documento.Column(3).Alignment = cellLeftCenter
        documento.Column(4).Alignment = cellRightCenter
        documento.Column(5).Alignment = cellRightCenter
  '
        documento.DefaultRowHeight = 15
        
        documento.PageSetup.PrintGridlines = False
        documento.AutoRedraw = False
'    Grid2.PageSetup.PrintGridlines = False
'    Grid2.AutoRedraw = False


        
        j = 15
'        vendedor = Resultados("vendedor")
        total = resultados("total")
        fecha = resultados("fecha")
     
        rut = resultados("rut")

        totNeto = resultados("neto")
        totIva = resultados("iva")
        totIha = resultados("iha")
        totIca = resultados("ica")
'        totIla = resultados("ila")
        ilarefrescos = resultados("ilarefrescos")
        ilalicores = resultados("ilalicores")
        ilavinos = resultados("ilavinos")
        totaldescuento = resultados("descuento")
        porcedescuento = resultados("descuento2")
        exento = CDbl(ilarefrescos) + CDbl(ilalicores) + CDbl(ilavinos) + CDbl(totIca) + CDbl(totIha)
        
    
        Descuento = "0"
        dona = resultados("donacion")
        
        
       
        Rem NUMERO
'        Documento.Cell(3, 5).Alignment = cellRightCenter
        documento.Cell(4, 1).text = "CAJA:" + PVentas.dato30.text + " F/O:" + resultados("numero")
                
        Rem Documento.Cell(4, 5).text = "F/O: " & NUMEROFACTURA
        documento.Range(3, 2, 3, 3).Merge
        documento.Range(3, 2, 3, 3).Alignment = cellCenterCenter
        documento.Cell(3, 2).text = leerNombreEmpresa(empresaActiva)
         
        documento.Cell(3, 1).text = "F/F: " & resultados("foliosii")
        foliofiscal = resultados("foliosii")
        'SEÑORES
        documento.Range(7, 2, 7, 3).Merge
        documento.Range(7, 2, 7, 3).Alignment = cellLeftCenter
        
        documento.Cell(7, 2).text = PVentas.lblRazon.Caption
           'FECHA
'        Documento.Range(7, 1, 7, 3).Merge
'        Documento.Range(7, 1, 7, 3).Alignment = cellLeftCenter
        documento.Cell(7, 5).text = fecha
          
        'DIRECCION
        documento.Range(9, 2, 9, 3).Merge
        documento.Range(9, 2, 9, 3).Alignment = cellLeftCenter
        documento.Cell(9, 2).text = PVentas.LBLDIRECCION.Caption
        
        
      'RUT
'        Documento.Range(8, 4, 8, 5).Merge
'        Documento.Range(8, 4, 8, 5).Alignment = cellLeftCenter
'        Documento.Cell(9, 5).Alignment = cellCenterCenter
        documento.Cell(9, 5).text = "     " + Format(PVentas.dato6.text, "###,###,##0") & "-" & PVentas.lbldv.Caption
 
        'GIRO
        documento.Range(11, 2, 11, 3).Merge
        documento.Range(11, 2, 11, 3).Alignment = cellLeftCenter
        documento.Cell(11, 2).text = leerGiroCliente(PVentas.dato6.text & PVentas.lbldv.Caption, PVentas.dato7.text)

        
        Rem tipo pago
'        tipopago = leerpago("FV", NUMEROFACTURA)
'

        'CIUDAD
'        Documento.Range(10, 4, 10, 5).Merge
'        Documento.Range(11, 5, 11, 5).Alignment = cellCenterCenter
        documento.Cell(11, 5).text = "     " + PVentas.LBLCIUDAD.Caption
    
        
        'DESCUENTO
'        Documento.Range(10, 4, 10, 5).Merge
'        Documento.Range(10, 4, 10, 5).Alignment = cellLeftCenter
'        Documento.Cell(10, 4).text = "D:" + Descuento
'
        
           
        
        lineas = 16
        While Not resultados.EOF
            Descuento = Str(CDbl(Descuento) + Int(resultados("cantidad") * resultados("precio")) - Int(resultados("totalpro")))
            Descuento = Mid(Descuento, 2, Len(Descuento))
            
            lineas = lineas + 1
            
            
            documento.Cell(lineas, 1).text = resultados(0)
            dife = CDbl(Int(resultados(2)) - resultados(2))
            If dife <> 0 Then
             documento.Cell(lineas, 2).text = Format(resultados(2), "###,##0.000")
            Else
             documento.Cell(lineas, 2).text = Format(resultados(2), "###,###")
            End If
            If resultados("tipodespacho") <> "" Then
            documento.Cell(lineas, 3).text = resultados(1) & " DESP. " & resultados("tipodespacho") & "-" & leerNombreTipoDespacho(resultados("tipodespacho"))
            Else
            documento.Cell(lineas, 3).text = resultados(1)
            End If
            
            documento.Cell(lineas, 4).text = Format(resultados(3), " $ ###,###,###")
            documento.Cell(lineas, 5).text = Format(resultados(4), " $ ###,###,###")
            
            resultados.MoveNext
        Wend
        MONTO = WORDNUM(Format(total, "########0"), "PESO", "PESOS", 0)
        
        Rem monto = numToLet(Format(total, "########0"), "PESO", "PESOS", 0)
        If totaldescuento <> 0 Then
        porce = totaldescuento / (CDbl(total) + CDbl(totaldescuento)) * 100
        'Documento.Cell(28, 4).text = "(-)%" + Format(PORCE, "##")
        documento.Cell(47, 4).Alignment = cellLeftCenter
        documento.Cell(47, 4).text = "  DESC"
        documento.Cell(47, 5).text = Format(totaldescuento * -1, " $ ###,###,##0")
        
        
        End If
        Set csql = Nothing
        csql.Close
        Set resultados = Nothing
          
'        Documento.Range(30, 1, 30, Documento.Cols - 1).Merge
'        Documento.Range(30, 1, 30, Documento.Cols - 1).Alignment = cellLeftCenter
'        Documento.Cell(30, 1).text = "    " + monto
'        Documento.Range(31, 1, 31, 2).Merge
        
'        If CDbl(porcedescuento) <> 0 Then
'        Documento.Cell(31, 1).text = "Dcto " + porcedescuento + "% N/Inc *OFE*"
'        End If
        

        documento.Cell(48, 4).Alignment = cellLeftCenter
        documento.Cell(48, 4).text = "  NETO"
        documento.Cell(48, 5).text = Format(totNeto, " $ ####,##0")
        
        documento.Cell(49, 4).Alignment = cellLeftCenter
        documento.Cell(49, 4).text = "  IVA"
        documento.Cell(49, 5).text = Format(totIva, " $ ###,##0")
        
        documento.Cell(50, 4).Alignment = cellLeftCenter
        'ariel cambia palabara exento por "otros impuetos"
        documento.Cell(50, 4).text = "  Otros impuestos"
        documento.Cell(50, 5).text = Format(exento, " $ ###,###,##0")
        
        documento.Cell(51, 4).Alignment = cellLeftCenter
        documento.Cell(51, 4).text = "  TOTAL"
        documento.Cell(51, 5).text = Format(total, " $ ###,###,##0")
        
        If CDbl(dona) > 0 Then
        documento.Cell(52, 4).Alignment = cellLeftCenter
        documento.Cell(52, 4).text = entidaddonacion
        
        documento.Cell(52, 5).text = Format(dona, " $ ###,###,##0")
        End If
        
        documento.Range(46, 1, 48, 3).Merge
        documento.Range(46, 1, 48, 3).Alignment = cellLeftCenter
        documento.Cell(46, 1).text = "                    " + MONTO
        documento.Range(46, 1, 48, 3).WrapText = True

   If ilarefrescos + ilavinos + ilalicores + totIha + totIca <> 0 Then
'
        documento.Range(55, 2, 55, 5).Merge
        documento.Range(55, 2, 55, 5).Alignment = cellLeftCenter
        documento.Range(55, 2, 55, 5).FontBold = True
        documento.Cell(55, 2).text = "      ILA 13     " + "      ILA 15     " + "      ILA 27     " + "      HARINA     " + "      CARNE      "
      
      

    'PALABRAIMPUESTO = CAMPO1(1) + CAMPO1(2) + CAMPO1(3) + CAMPO1(4) + CAMPO1(5)
    'Grid2.Cell(59, 2).text = PALABRAIMPUESTO
 

      
        
        CAMPO1 = String(10 - Len(ilarefrescos), 32) & Format(ilarefrescos, "####,##0")
        campo2 = String(10 - Len(ilavinos), 32) & Format(ilavinos, "####,##0")
        
        campo3 = String(10 - Len(ilalicores), 32) & Format(ilalicores, "####,##0")
       
        campo4 = String(10 - Len(totIha), 32) & Format(totIha, "####,##0")
        
        campo5 = String(10 - Len(totIca), 32) & Format(totIca, "####,##0")
       
'        campo6 = String(7 - Len(total), 32) & Format(total, "####,##0")

'       Documento   = "   ILA 13     " + "   ILA 15     " + "   ILA 27     " + "   HARINA     " + "   CARNE      "
               cad = "  " & CAMPO1 & "          " & campo2 & "          " & campo3 & "          " & campo4 & "          " & campo5
     
'        Documento.Range(59, 2, 59, Documento.Cols - 1).Merge
'        Documento.RowHeight(59) = 20
         documento.Range(56, 2, 56, 5).Merge
         documento.Range(56, 2, 56, 5).Alignment = cellLeftCenter
         documento.Range(56, 2, 56, 5).FontBold = True
         documento.Cell(56, 2).text = cad
        End If
'        Call leercredito("FV", PVentas.dato30.text, foliofiscal, rut)
'        If cantidadCUOTAS <> "" Then
'
'        Documento.Range(56, 1, 56, 4).Merge
'        Documento.Range(57, 1, 57, 4).Merge
'        Documento.Range(58, 1, 58, 4).Merge
'        Documento.Range(59, 1, 59, 4).Merge
'        Documento.Range(60, 1, 60, 4).Merge
'        Documento.Range(61, 1, 61, 4).Merge
'        Documento.Range(62, 1, 62, 4).Merge
'
'        Documento.Cell(56, 1).Border(cellEdgeTop) = cellThick
'        Documento.Cell(56, 1).Alignment = cellLeftCenter
'        Documento.Cell(57, 1).Alignment = cellLeftCenter
'        Documento.Cell(58, 1).Alignment = cellLeftCenter
'        Documento.Cell(59, 1).Alignment = cellLeftCenter
'        Documento.Cell(60, 1).Alignment = cellLeftCenter
'        Documento.Cell(61, 1).Alignment = cellLeftCenter
'        Documento.Cell(62, 1).Alignment = cellLeftCenter
'
'
'        Documento.Cell(56, 1).text = "FIRMA :"
'        Documento.Cell(57, 1).text = "Yo " + leerNombreCliente(rutcredito) + ""
'        Documento.Cell(58, 1).text = "CI." + Mid(rutcredito, 1, 9) + "-" + Mid(rutcredito, 10, 1) + " autorizo segun contrato PALGUIN LTDA " + "Cargar a mi cuenta " + cantidadCUOTAS + " cuotas de " + Format(montocuotas, "$ ###,###,###")
'        Documento.Cell(59, 1).text = "Primer Vencimiento " + Format(primervencimiento, "dd-mm-yyyy")
'        Documento.Cell(60, 1).text = "TOTAL CREDITO :" + Format(montocredito, "$ ###,###,###") + " PIE :" + Format(CDbl(montototalventa) - CDbl(montocredito), "$ ###,###,###")
'
'
'        End If
                
        
 
        documento.Range(2, 1, 2, 5).Merge
        documento.Range(2, 2, 2, 5).Alignment = cellCenterCenter
        documento.Cell(2, 1).text = "cajero(a);" & PVentas.lblcajera.Caption & " HORA :" & Time
'        If vendedor <> "" Then
'        Documento.Range(60, 1, 60, 5).Merge
'        Documento.Range(60, 2, 60, 5).Alignment = cellCenterCenter
'        Documento.Cell(60, 1).text = "Vendedor(a);" + vendedor & " " & leerNombreVendedor(vendedor) + " "
'        End If
               
         
        documento.AutoRedraw = True
        documento.Refresh
        

        
        documento.PageSetup.LeftMargin = 0.25
        documento.PageSetup.RightMargin = 0
        documento.PageSetup.TopMargin = 3
        documento.PageSetup.BottomMargin = 0
        
        For i = 1 To documento.PageSetup.PaperSizes.Count
            If UCase(documento.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
                documento.PageSetup.PaperSize = documento.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        
        
        
'        Documento.PageSetup.PaperWidth = 21.6
'        Documento.PageSetup.PaperHeight = 16.5
'        'Documento.PageSetup.PrintGridlines = True
        documento.PageSetup.PrintGridlines = False
 
        documento.PrintPreview
'
'    Grid2.PageSetup.PrintGridlines = False
'    'Grid2.DirectPrint
'    Grid2.PrintPreview
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & NUMEROFACTURA)
    End If
End Sub

