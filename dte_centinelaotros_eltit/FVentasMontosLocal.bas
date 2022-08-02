Attribute VB_Name = "FVentasMontosLocal"
Option Explicit

Public Sub generaInformeVML(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("RESUMEN DE VENTAS POR MONTO POR LOCAL - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), codLoc, impresion)
    If detalle = True Then
        altoFila = impresion.DefaultRowHeight
        ultimaFila = detalleVentas(data, impresion, TIPO, codLoc, fecha1, fecha2)
        For i = 1 To ultimaFila
            impresion.RowHeight(i) = 0
        Next i
        Call impresion.HPageBreaks.Add(impresion.Rows - 1)
    End If
    Call resumenVentas(data, impresion, TIPO, codLoc, fecha1, fecha2)
    Call resumenDepto(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function detalleVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim cantidad As Double
    Dim kilos As Double
    Dim PRECIO As Double
    Dim total As Double
    
    rubAux = leerRubro(codLoc)
    tabla = "SELECT CONCAT(' ', dd.tipo, '    ', dd.numero, '    ', dd.linea, '   ', dd.fecha, '   ', dd.codigo, '" & vbTab & vbTab & "', dd.descripcion, '" & vbTab & vbTab & vbTab & vbTab & "') AS item1, CONCAT(FORMAT(IF(dd.tipo = 'NV', 1, dd.cantidad),0), '" & vbTab & "', FORMAT(dd.unidades,0), '" & vbTab & "') AS item2, CONCAT('$ ', FORMAT(dd.precio,0), '" & vbTab & "', '$ ', FORMAT(dd.total,0)) AS item3, IF(dd.tipo = 'NV', 1, dd.cantidad) AS cantidad, dd.unidades, dd.precio, dd.total, dd.tipo, dd.descuento "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' "
    tabla = tabla & "WHERE dc.local = '" & codLoc & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "ORDER BY dc.fecha, dc.tipo, dc.numero, dd.linea ASC"
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        'TITULO
        impresion.AddItem "DETALLE DE VENTAS", True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.AddItem "", True
        'CABEZA
        impresion.AddItem "TIPO    NUMERO    LINEA       FECHA            CODIGO" & vbTab & vbTab & "DESCRIPCION" & vbTab & vbTab & vbTab & vbTab & "CANTIDAD" & vbTab & "         " & vbTab & "PRECIO" & vbTab & "TOTAL", True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 6).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellLeftCenter
        impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        cantidad = 0
        kilos = 0
        PRECIO = 0
        total = 0
        
        data.Recordset.MoveFirst
        While Not data.Recordset.EOF
            cantidad = cantidad + CDbl(data.Recordset.Fields("cantidad"))
            kilos = kilos + CDbl(data.Recordset.Fields("unidades"))
            PRECIO = PRECIO + CDbl(data.Recordset.Fields("precio"))
            If data.Recordset.Fields("tipo") = "FV" Or data.Recordset.Fields("tipo") = "NV" Then
                'If CDbl(data.Recordset.Fields("descuento")) <> 0 Then Stop
                total = total + CDbl(data.Recordset.Fields("total")) - CDbl(data.Recordset.Fields("total")) * CDbl(data.Recordset.Fields("descuento")) / 100 + CDbl(data.Recordset.Fields("total")) * (iva / 100) + CDbl(data.Recordset.Fields("total")) * (iha / 100)
            Else
                'If CDbl(data.Recordset.Fields("descuento")) <> 0 Then Stop
                total = total + CDbl(data.Recordset.Fields("total")) - CDbl(data.Recordset.Fields("total")) * CDbl(data.Recordset.Fields("descuento")) / 100
            End If
            impresion.AddItem data.Recordset.Fields("item1") & Replace(data.Recordset.Fields("item2"), ",", ".") & Replace(data.Recordset.Fields("item3"), ",", "."), True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
            impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 6).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellLeftCenter
            impresion.Range(impresion.Rows - 1, 7, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
            data.Recordset.MoveNext
        Wend
        impresion.AddItem "", True
        impresion.AddItem "TOTAL DOCUMENTOS" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & Format(cantidad, "###,###,##0") & vbTab & Format(kilos, "###,###,##0") & vbTab & Format(PRECIO, "$ ###,###,##0") & vbTab & Format(total, "$ ###,###,##0"), True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 7, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.AddItem "", True
    End If
    detalleVentas = impresion.Rows - 1
End Function

Private Sub resumenVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim rubAux As String
    Dim cadena As String
    Dim i As Integer
    Dim CODIGO As String
    Dim tipoDoc As String
    Dim impuesto As String
    Dim suma(7) As Double
    Dim sumKilos As Double
    Dim sumCantidad As Double
    Dim sumSubTotal As Double
    Dim sumDescuento As Double
    Dim sumNeto As Double
    Dim sumIva As Double
    Dim sumIha As Double
    Dim sumTotal As Double
    
    rubAux = leerRubro(codLoc)
     
    tabla = "SELECT dd.tipo, dd.codigo, mi.nombre AS impuesto, dd.descripcion, SUM(dd.unidades) AS kilos, IF(dd.tipo = 'NV', 1, SUM(dd.cantidad)) AS cantidad, dd.precio, dd.descuento "
    tabla = tabla & "FROM sv_documento_detalle_" + codLoc + " As dd INNER JOIN sv_documento_cabeza_" + codLoc + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & basedatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & basedatos & ".g_maestroimpuestos As mi ON mpf.codigoimpuesto = mi.codigo "
    tabla = tabla & "WHERE dd.local = '" & codLoc & "' AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " and dd.fecha=dc.fecha and dd.caja=dc.caja "
    tabla = tabla & "GROUP BY codigo, tipo, nombre, precio, descuento "
    tabla = tabla & "ORDER BY codigo, tipo, nombre ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & codLoc, usuario, password, tabla)
    For i = 0 To 7
        suma(i) = 0
    Next i
    
    If data.Recordset.RecordCount > 0 Then
        cadena = "RESUMEN DE VENTAS "
        If InStr(1, TIPO, "FV", vbBinaryCompare) <> 0 Then
            If InStr(1, TIPO, "BV", vbBinaryCompare) <> 0 Then
                cadena = cadena & "TODAS"
            Else
                cadena = cadena & "FACTURAS"
            End If
        Else
            cadena = cadena & "BOLETAS"
        End If
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        
        impresion.AddItem "", True
        
        cadena = "CODIGO" & vbTab
        cadena = cadena & "DESCRIPCION" & vbTab
        cadena = cadena & "P.P." & vbTab
        cadena = cadena & "         " & vbTab
        cadena = cadena & "CANTIDAD" & vbTab
        cadena = cadena & "SUBTOTAL" & vbTab
        cadena = cadena & "DESCUENTO" & vbTab
        cadena = cadena & "NETO" & vbTab
        cadena = cadena & "IVA" & vbTab
        cadena = cadena & "IHA" & vbTab
        cadena = cadena & "TOTAL"
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        data.Recordset.MoveFirst
        CODIGO = data.Recordset.Fields("codigo")
        
        sumKilos = 0
        sumCantidad = 0
        sumSubTotal = 0
        sumDescuento = 0
        sumNeto = 0
        sumIva = 0
        sumIha = 0
        sumTotal = 0
        
        While Not data.Recordset.EOF
            If CODIGO = data.Recordset.Fields("codigo") Then
                tipoDoc = data.Recordset.Fields("tipo")
                impuesto = data.Recordset.Fields("impuesto")
                Rem sumKilos = sumKilos + data.Recordset.Fields("         ")
                sumCantidad = sumCantidad + data.Recordset.Fields("cantidad")
                
                sumSubTotal = sumSubTotal + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0)
                               
                sumDescuento = sumDescuento + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)
                Select Case tipoDoc
                    Case "FV", "NV"
                        sumNeto = sumNeto + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)
                        Select Case impuesto
                            Case "IVA"
                                sumIva = sumIva + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iva / 100 + 0.1, 0)
                                sumIha = sumIha
                            Case "IHA"
                                sumIva = sumIva + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iva / 100 + 0.1, 0)
                                sumIha = sumIha + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iha / 100 + 0.1, 0)
                            Case "EXENTO"
                                sumIva = sumIva
                                sumIha = sumIha
                        End Select
                    Case "BV", "ZE"
                        sumNeto = sumNeto + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) / (1 + iva / 100) + 0.1, 0)
                        Select Case impuesto
                            Case "IVA", "IHA"
                                sumIva = sumIva + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) / (1 + iva / 100) * iva / 100 + 0.1, 0)
                                
                                sumIha = sumIha
                            Case "EXENTO"
                                sumIva = sumIva
                                sumIha = sumIha
                        End Select
                    Case "FE"
                        sumNeto = sumNeto + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)
                        sumIva = 0
                        sumIha = 0
                End Select
            Else
                sumKilos = 1
                sumTotal = sumNeto + sumIva + sumIha
                sumSubTotal = sumNeto + sumDescuento
                cadena = CODIGO & vbTab
                cadena = cadena & leerNombreProducto(CODIGO) & vbTab
                cadena = cadena & Format(sumNeto / sumKilos, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
                cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
                cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
                impresion.AddItem cadena, True
                CODIGO = data.Recordset.Fields("codigo")
                
                suma(0) = suma(0) + sumKilos
                suma(1) = suma(1) + sumCantidad
                suma(2) = suma(2) + sumSubTotal
                suma(3) = suma(3) + sumDescuento
                suma(4) = suma(4) + sumNeto
                suma(5) = suma(5) + sumIva
                suma(6) = suma(6) + sumIha
                suma(7) = suma(7) + sumTotal
                
                sumKilos = 0
                sumCantidad = 0
                sumSubTotal = 0
                sumDescuento = 0
                sumNeto = 0
                sumIva = 0
                sumIha = 0
                sumTotal = 0
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        sumKilos = 1
        data.Recordset.MovePrevious
        sumTotal = sumNeto + sumIva + sumIha
        cadena = CODIGO & vbTab
        cadena = cadena & leerNombreProducto(CODIGO) & vbTab
        cadena = cadena & Format(sumNeto / sumKilos, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
        cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
        cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
        impresion.AddItem cadena, True
        
        suma(0) = suma(0) + sumKilos
        suma(1) = suma(1) + sumCantidad
        suma(2) = suma(2) + sumSubTotal
        suma(3) = suma(3) + sumDescuento
        suma(4) = suma(4) + sumNeto
        suma(5) = suma(5) + sumIva
        suma(6) = suma(6) + sumIha
        suma(7) = suma(7) + sumTotal
        
        cadena = "TOTAL GENERAL" & vbTab & vbTab & vbTab
        cadena = cadena & Format(suma(0), "###,###,##0") & vbTab
        cadena = cadena & Format(suma(1), "###,###,##0") & vbTab
        For i = 2 To 7
            cadena = cadena & Format(suma(i), "$ ###,###,##0") & vbTab
        Next i
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    End If
End Sub

Private Sub resumenDepto(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim rubAux As String
    Dim cadena As String
    Dim i As Integer
    Dim codigoLinea As String
    Dim codigoDepto As String
    Dim codigoSeccion As String
    Dim linea As String
    Dim depto As String
    Dim tipoDoc As String
    Dim impuesto As String
    Dim suma(7) As Double
    Dim sumaDepto(7) As Double
    Dim sumKilos As Double
    Dim sumCantidad As Double
    Dim sumSubTotal As Double
    Dim sumDescuento As Double
    Dim sumNeto As Double
    Dim sumIva As Double
    Dim sumIha As Double
    Dim sumTotal As Double
    
    rubAux = leerRubro(codLoc)
    tabla = ""
    tabla = "SELECT dd.tipo, dd.codigo, mi.nombre AS impuesto, dd.descripcion, IF(dd.tipo = 'NV', IF(dd.codigo = '0000000000100' OR dd.codigo = '0000000000101', 0, -1 * SUM(dd.unidades)), SUM(dd.unidades)) AS kilos, IF(dd.tipo = 'NV', /*IF(dd.codigo = '0000000000100' OR dd.codigo = '0000000000101', 0,*/ -1 * SUM(dd.cantidad)/*)*/, SUM(dd.cantidad)) AS cantidad, dd.precio, dd.descuento, mpf.codigolinea, mpf.codigodepto, mpf.codigoseccion  "
    tabla = tabla & "FROM sv_documento_detalle_" + codLoc + " As dd INNER JOIN sv_documento_cabeza_" + codLoc + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & basedatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & basedatos & ".g_maestroimpuestos As mi ON mpf.codigoimpuesto = mi.codigo "
    tabla = tabla & "WHERE dd.local = '" & codLoc & "' AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " and dd.fecha=dc.fecha and dd.caja=dc.caja "
    tabla = tabla & "GROUP BY codigo, tipo, nombre, precio, descuento, codigolinea, codigodepto, codigoseccion "
    tabla = tabla & "ORDER BY codigoseccion, codigodepto, codigolinea ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & codLoc, usuario, password, tabla)
    
    For i = 0 To 7
        suma(i) = 0
        sumaDepto(i) = 0
    Next i
    If data.Recordset.RecordCount > 0 Then
        impresion.AddItem "", True
        impresion.AddItem "", True
        cadena = "RESUMEN DE VENTAS TOTALIZADO POR DEPARTAMENTOS"
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        
        impresion.AddItem "", True
        
        cadena = "CODIGO" & vbTab
        cadena = cadena & "DESCRIPCION" & vbTab
        cadena = cadena & "P.P." & vbTab
        cadena = cadena & "         " & vbTab
        cadena = cadena & "CANTIDAD" & vbTab
        cadena = cadena & "SUBTOTAL" & vbTab
        cadena = cadena & "DESCUENTO" & vbTab
        cadena = cadena & "NETO" & vbTab
        cadena = cadena & "IVA" & vbTab
        cadena = cadena & "IHA" & vbTab
        cadena = cadena & "TOTAL"
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        data.Recordset.MoveFirst
        codigoLinea = data.Recordset.Fields("codigolinea")
        codigoDepto = data.Recordset.Fields("codigodepto")
        codigoSeccion = data.Recordset.Fields("codigoseccion")
        
        sumKilos = 0
        sumCantidad = 0
        sumSubTotal = 0
        sumDescuento = 0
        sumNeto = 0
        sumIva = 0
        sumIha = 0
        sumTotal = 0
        
        While Not data.Recordset.EOF
            If codigoDepto = data.Recordset.Fields("codigodepto") Then
                If codigoLinea = data.Recordset.Fields("codigolinea") Then
                    tipoDoc = data.Recordset.Fields("tipo")
                    impuesto = data.Recordset.Fields("impuesto")
                    'sumKilos = sumKilos + data.Recordset.Fields("         ")
                    sumCantidad = sumCantidad + data.Recordset.Fields("cantidad")
                    sumSubTotal = sumSubTotal + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0)
                    sumDescuento = sumDescuento + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)
                    Select Case tipoDoc
                        Case "FV"
                            sumNeto = sumNeto + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)
                            Select Case impuesto
                                Case "IVA"
                                    sumIva = sumIva + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iva / 100 + 0.1, 0)
                                    sumIha = sumIha
                                Case "IHA"
                                    sumIva = sumIva + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iva / 100 + 0.1, 0)
                                    sumIha = sumIha + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iha / 100 + 0.1, 0)
                                Case "EXENTO"
                                    sumIva = sumIva
                                    sumIha = sumIha
                            End Select
                        Case "NV"
                            sumNeto = sumNeto + Round(-1 * CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)
                            Select Case impuesto
                                Case "IVA"
                                    sumIva = sumIva + Round((Round(-1 * CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iva / 100 + 0.1, 0)
                                    sumIha = sumIha
                                Case "IHA"
                                    sumIva = sumIva + Round((Round(-1 * CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iva / 100 + 0.1, 0)
                                    sumIha = sumIha + Round((Round(-1 * CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) * iha / 100 + 0.1, 0)
                                Case "EXENTO"
                                    sumIva = sumIva
                                    sumIha = sumIha
                            End Select
                        Case "BV", "ZE"
                            sumNeto = sumNeto + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) / (1 + iva / 100) + 0.1, 0)
                            Select Case impuesto
                                Case "IVA", "IHA"
                                    sumIva = sumIva + Round((Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)) / (1 + iva / 100) * iva / 100 + 0.1, 0)
                                    sumIha = sumIha
                                Case "EXENTO"
                                    sumIva = sumIva
                                    sumIha = sumIha
                            End Select
                        Case "FE"
                            sumNeto = sumNeto + Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")), 0) - Round(CDbl(data.Recordset.Fields("cantidad")) * CDbl(data.Recordset.Fields("precio")) * CDbl(data.Recordset.Fields("descuento")) / 100 + 0.1, 0)
                            sumIva = 0
                            sumIha = 0
                    End Select
                Else
                    sumKilos = 1
                    sumTotal = sumNeto + sumIva + sumIha
                    cadena = codigoLinea & vbTab
                    cadena = cadena & leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux) & vbTab
                    cadena = cadena & Format(sumNeto / sumKilos, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
                    cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
                    cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
                    impresion.AddItem cadena, True
                    codigoLinea = data.Recordset.Fields("codigolinea")
                    
                    suma(0) = suma(0) + sumKilos
                    suma(1) = suma(1) + sumCantidad
                    suma(2) = suma(2) + sumSubTotal
                    suma(3) = suma(3) + sumDescuento
                    suma(4) = suma(4) + sumNeto
                    suma(5) = suma(5) + sumIva
                    suma(6) = suma(6) + sumIha
                    suma(7) = suma(7) + sumTotal
                    
                    sumaDepto(0) = sumaDepto(0) + sumKilos
                    sumaDepto(1) = sumaDepto(1) + sumCantidad
                    sumaDepto(2) = sumaDepto(2) + sumSubTotal
                    sumaDepto(3) = sumaDepto(3) + sumDescuento
                    sumaDepto(4) = sumaDepto(4) + sumNeto
                    sumaDepto(5) = sumaDepto(5) + sumIva
                    sumaDepto(6) = sumaDepto(6) + sumIha
                    sumaDepto(7) = sumaDepto(7) + sumTotal
                    
                    sumKilos = 0
                    sumCantidad = 0
                    sumSubTotal = 0
                    sumDescuento = 0
                    sumNeto = 0
                    sumIva = 0
                    sumIha = 0
                    sumTotal = 0
                    data.Recordset.MovePrevious
                End If
            Else
                sumKilos = 1
                sumTotal = sumNeto + sumIva + sumIha
                cadena = codigoLinea & vbTab
                cadena = cadena & leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux) & vbTab
                cadena = cadena & Format(sumNeto / sumKilos, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
                cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
                cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
                impresion.AddItem cadena, True
                codigoLinea = data.Recordset.Fields("codigolinea")
                
                suma(0) = suma(0) + sumKilos
                suma(1) = suma(1) + sumCantidad
                suma(2) = suma(2) + sumSubTotal
                suma(3) = suma(3) + sumDescuento
                suma(4) = suma(4) + sumNeto
                suma(5) = suma(5) + sumIva
                suma(6) = suma(6) + sumIha
                suma(7) = suma(7) + sumTotal
                
                sumaDepto(0) = sumaDepto(0) + sumKilos
                sumaDepto(1) = sumaDepto(1) + sumCantidad
                sumaDepto(2) = sumaDepto(2) + sumSubTotal
                sumaDepto(3) = sumaDepto(3) + sumDescuento
                sumaDepto(4) = sumaDepto(4) + sumNeto
                sumaDepto(5) = sumaDepto(5) + sumIva
                sumaDepto(6) = sumaDepto(6) + sumIha
                sumaDepto(7) = sumaDepto(7) + sumTotal
                
                sumKilos = 0
                sumCantidad = 0
                sumSubTotal = 0
                sumDescuento = 0
                sumNeto = 0
                sumIva = 0
                sumIha = 0
                sumTotal = 0
            
                cadena = codigoDepto & vbTab
                cadena = "TOTAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubAux) & vbTab & vbTab & vbTab
                cadena = cadena & Format(sumaDepto(0), "###,###,##0") & vbTab
                cadena = cadena & Format(sumaDepto(1), "###,###,##0") & vbTab
                For i = 2 To 7
                    cadena = cadena & Format(sumaDepto(i), "$ ###,###,##0") & vbTab
                Next i
                impresion.AddItem cadena, True
                impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellLeftCenter
                impresion.AddItem "", True
                
                For i = 0 To 7
                    sumaDepto(i) = 0
                Next i
                
                codigoSeccion = data.Recordset.Fields("codigoseccion")
                codigoDepto = data.Recordset.Fields("codigodepto")
                codigoLinea = data.Recordset.Fields("codigolinea")
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        sumKilos = 1
        data.Recordset.MovePrevious
        sumTotal = sumNeto + sumIva + sumIha
        cadena = codigoLinea & vbTab
        cadena = cadena & leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux) & vbTab
        cadena = cadena & Format(sumNeto / sumKilos, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
        cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
        cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
        impresion.AddItem cadena, True
        
        suma(0) = suma(0) + sumKilos
        suma(1) = suma(1) + sumCantidad
        suma(2) = suma(2) + sumSubTotal
        suma(3) = suma(3) + sumDescuento
        suma(4) = suma(4) + sumNeto
        suma(5) = suma(5) + sumIva
        suma(6) = suma(6) + sumIha
        suma(7) = suma(7) + sumTotal
        
        sumaDepto(0) = sumaDepto(0) + sumKilos
        sumaDepto(1) = sumaDepto(1) + sumCantidad
        sumaDepto(2) = sumaDepto(2) + sumSubTotal
        sumaDepto(3) = sumaDepto(3) + sumDescuento
        sumaDepto(4) = sumaDepto(4) + sumNeto
        sumaDepto(5) = sumaDepto(5) + sumIva
        sumaDepto(6) = sumaDepto(6) + sumIha
        sumaDepto(7) = sumaDepto(7) + sumTotal
        
        cadena = codigoDepto & vbTab
        cadena = "TOTAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubAux) & vbTab & vbTab & vbTab
        cadena = cadena & Format(sumaDepto(0), "###,###,##0") & vbTab
        cadena = cadena & Format(sumaDepto(1), "###,###,##0") & vbTab
        For i = 2 To 7
            cadena = cadena & Format(sumaDepto(i), "$ ###,###,##0") & vbTab
        Next i
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellLeftCenter
        impresion.AddItem "", True
        
        cadena = "TOTAL GENERAL" & vbTab & vbTab & vbTab
        cadena = cadena & Format(suma(0), "###,###,##0") & vbTab
        cadena = cadena & Format(suma(1), "###,###,##0") & vbTab
        For i = 2 To 7
            cadena = cadena & Format(suma(i), "$ ###,###,##0") & vbTab
        Next i
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    End If
End Sub
