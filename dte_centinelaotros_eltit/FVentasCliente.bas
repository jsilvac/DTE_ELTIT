Attribute VB_Name = "FVentasCliente"
Option Explicit

Public Sub generaInformeVC(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String, ByVal rut As String)
    Dim i As Long
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("RESUMEN DE VENTAS POR CLIENTE - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), codLoc, Impresion)
    Call resumenVentas(data, Impresion, TIPO, codLoc, fecha1, fecha2, rut)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub resumenVentas(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String, ByVal rut As String)
    Dim tabla As String
    Dim rubAux As String
    Dim cadena As String
    Dim i As Integer
    Dim codigo As String
    Dim sucursal As String
    Dim tipoDoc As String
    Dim impuesto As String
    Dim suma(2, 7) As Double
    Dim sumKilos As Double
    Dim sumCantidad As Double
    Dim sumSubTotal As Double
    Dim sumDescuento As Double
    Dim sumNeto As Double
    Dim sumIva As Double
    Dim sumIha As Double
    Dim sumTotal As Double
    
    rubAux = leerRubro(codLoc)
    
    tabla = "SELECT dd.tipo, dd.codigo, mi.nombre AS impuesto, dd.descripcion, SUM(dd.unidades) AS kilos, IF(dd.tipo = 'NV', 1, SUM(dd.cantidad)) AS cantidad, dd.precio, dd.descuento, dc.sucursal "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & baseDatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & baseDatos & ".g_maestroimpuestos As mi ON mpf.codigoimpuesto = mi.codigo "
    tabla = tabla & "WHERE dd.local = '" & codLoc & "' AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " AND dd.rut = '" & rut & "' "
    tabla = tabla & "GROUP BY codigo, tipo, nombre, precio, descuento, sucursal "
    tabla = tabla & "ORDER BY codigo, tipo, nombre, sucursal ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    For i = 0 To 7
        suma(0, i) = 0
        suma(1, i) = 0
    Next i
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        sucursal = data.Recordset.Fields("sucursal")
        
        cadena = vbTab & "CLIENTE  :  " & leerNombreClienteSucursal(rut, sucursal)
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 4).Merge
        cadena = vbTab & "SUCURSAL :  " & leerDireccionCliente(rut, sucursal) & vbTab & vbTab & vbTab & leerCiudadCliente(rut, sucursal)
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 4).Merge
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 7).Merge
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 7).Alignment = cellLeftCenter
        Impresion.AddItem "", True
        cadena = "CODIGO" & vbTab
        cadena = cadena & "DESCRIPCION" & vbTab
        cadena = cadena & "         " & vbTab
        cadena = cadena & "CANTIDAD" & vbTab
        cadena = cadena & "SUBTOTAL" & vbTab
        cadena = cadena & "DESCUENTO" & vbTab
        cadena = cadena & "NETO" & vbTab
        cadena = cadena & "IVA" & vbTab
        cadena = cadena & "IHA" & vbTab
        cadena = cadena & "TOTAL"
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        codigo = data.Recordset.Fields("codigo")
        
        sumKilos = 0
        sumCantidad = 0
        sumSubTotal = 0
        sumDescuento = 0
        sumNeto = 0
        sumIva = 0
        sumIha = 0
        sumTotal = 0
        
        While Not data.Recordset.EOF
            If sucursal = data.Recordset.Fields("sucursal") Then
                If codigo = data.Recordset.Fields("codigo") Then
                    tipoDoc = data.Recordset.Fields("tipo")
                    impuesto = data.Recordset.Fields("impuesto")
                    sumKilos = sumKilos + data.Recordset.Fields("         ")
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
                    sumTotal = sumNeto + sumIva + sumIha
                    cadena = codigo & vbTab
                    cadena = cadena & leerNombreProducto(codigo) & vbTab
                    cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
                    cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
                    cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
                    cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
                    Impresion.AddItem cadena, True
                    codigo = data.Recordset.Fields("codigo")
                    
                    suma(0, 0) = suma(0, 0) + sumKilos
                    suma(0, 1) = suma(0, 1) + sumCantidad
                    suma(0, 2) = suma(0, 2) + sumSubTotal
                    suma(0, 3) = suma(0, 3) + sumDescuento
                    suma(0, 4) = suma(0, 4) + sumNeto
                    suma(0, 5) = suma(0, 5) + sumIva
                    suma(0, 6) = suma(0, 6) + sumIha
                    suma(0, 7) = suma(0, 7) + sumTotal
                    
                    suma(1, 0) = suma(1, 0) + sumKilos
                    suma(1, 1) = suma(1, 1) + sumCantidad
                    suma(1, 2) = suma(1, 2) + sumSubTotal
                    suma(1, 3) = suma(1, 3) + sumDescuento
                    suma(1, 4) = suma(1, 4) + sumNeto
                    suma(1, 5) = suma(1, 5) + sumIva
                    suma(1, 6) = suma(1, 6) + sumIha
                    suma(1, 7) = suma(1, 7) + sumTotal
                    
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
                sumTotal = sumNeto + sumIva + sumIha
                cadena = codigo & vbTab
                cadena = cadena & leerNombreProducto(codigo) & vbTab
                cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
                cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
                cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
                cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
                Impresion.AddItem cadena, True
                
                codigo = data.Recordset.Fields("codigo")
                
                suma(0, 0) = suma(0, 0) + sumKilos
                suma(0, 1) = suma(0, 1) + sumCantidad
                suma(0, 2) = suma(0, 2) + sumSubTotal
                suma(0, 3) = suma(0, 3) + sumDescuento
                suma(0, 4) = suma(0, 4) + sumNeto
                suma(0, 5) = suma(0, 5) + sumIva
                suma(0, 6) = suma(0, 6) + sumIha
                suma(0, 7) = suma(0, 7) + sumTotal
                
                suma(1, 0) = suma(1, 0) + sumKilos
                suma(1, 1) = suma(1, 1) + sumCantidad
                suma(1, 2) = suma(1, 2) + sumSubTotal
                suma(1, 3) = suma(1, 3) + sumDescuento
                suma(1, 4) = suma(1, 4) + sumNeto
                suma(1, 5) = suma(1, 5) + sumIva
                suma(1, 6) = suma(1, 6) + sumIha
                suma(1, 7) = suma(1, 7) + sumTotal
                
                cadena = "TOTAL SUCURSAL" & vbTab & vbTab
                cadena = cadena & Format(suma(1, 0), "###,###,##0") & vbTab
                cadena = cadena & Format(suma(1, 1), "###,###,##0") & vbTab
                For i = 2 To 7
                    cadena = cadena & Format(suma(1, i), "$ ###,###,##0") & vbTab
                Next i
                Impresion.AddItem cadena, True
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                Impresion.AddItem "", True
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellDot
                
                For i = 0 To 7
                    suma(1, i) = 0
                Next i
                sumKilos = 0
                sumCantidad = 0
                sumSubTotal = 0
                sumDescuento = 0
                sumNeto = 0
                sumIva = 0
                sumIha = 0
                sumTotal = 0
                
                Impresion.AddItem "", True
                codigo = data.Recordset.Fields("codigo")
                sucursal = data.Recordset.Fields("sucursal")
                
                cadena = vbTab & "CLIENTE  :  " & leerNombreClienteSucursal(rut, sucursal)
                Impresion.AddItem cadena, True
                Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 4).Merge
                cadena = vbTab & "SUCURSAL :  " & leerDireccionCliente(rut, sucursal) & vbTab & vbTab & vbTab & leerCiudadCliente(rut, sucursal)
                Impresion.AddItem cadena, True
                Impresion.Range(Impresion.Rows - 1, 2, Impresion.Rows - 1, 4).Merge
                Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 7).Merge
                Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 7).Alignment = cellLeftCenter
                Impresion.AddItem "", True
                cadena = "CODIGO" & vbTab
                cadena = cadena & "DESCRIPCION" & vbTab
                cadena = cadena & "         " & vbTab
                cadena = cadena & "CANTIDAD" & vbTab
                cadena = cadena & "SUBTOTAL" & vbTab
                cadena = cadena & "DESCUENTO" & vbTab
                cadena = cadena & "NETO" & vbTab
                cadena = cadena & "IVA" & vbTab
                cadena = cadena & "IHA" & vbTab
                cadena = cadena & "TOTAL"
                Impresion.AddItem cadena, True
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
                
                Impresion.AddItem "", True
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        data.Recordset.MovePrevious
        sumTotal = sumNeto + sumIva + sumIha
        cadena = codigo & vbTab
        cadena = cadena & leerNombreProducto(codigo) & vbTab
        cadena = cadena & Format(sumKilos, "###,###,##0") & vbTab
        cadena = cadena & Format(sumCantidad, "###,###,##0") & vbTab
        cadena = cadena & Format(sumSubTotal, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumDescuento, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumNeto, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumIva, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumIha, "$ ###,###,##0") & vbTab
        cadena = cadena & Format(sumTotal, "$ ###,###,##0") & vbTab
        Impresion.AddItem cadena, True
        
        suma(0, 0) = suma(0, 0) + sumKilos
        suma(0, 1) = suma(0, 1) + sumCantidad
        suma(0, 2) = suma(0, 2) + sumSubTotal
        suma(0, 3) = suma(0, 3) + sumDescuento
        suma(0, 4) = suma(0, 4) + sumNeto
        suma(0, 5) = suma(0, 5) + sumIva
        suma(0, 6) = suma(0, 6) + sumIha
        suma(0, 7) = suma(0, 7) + sumTotal
        
        suma(1, 0) = suma(1, 0) + sumKilos
        suma(1, 1) = suma(1, 1) + sumCantidad
        suma(1, 2) = suma(1, 2) + sumSubTotal
        suma(1, 3) = suma(1, 3) + sumDescuento
        suma(1, 4) = suma(1, 4) + sumNeto
        suma(1, 5) = suma(1, 5) + sumIva
        suma(1, 6) = suma(1, 6) + sumIha
        suma(1, 7) = suma(1, 7) + sumTotal
        
        cadena = "TOTAL SUCURSAL" & vbTab & vbTab
        cadena = cadena & Format(suma(1, 0), "###,###,##0") & vbTab
        cadena = cadena & Format(suma(1, 1), "###,###,##0") & vbTab
        For i = 2 To 7
            cadena = cadena & Format(suma(1, i), "$ ###,###,##0") & vbTab
        Next i
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        
        cadena = "TOTAL GENERAL" & vbTab & vbTab
        cadena = cadena & Format(suma(0, 0), "###,###,##0") & vbTab
        cadena = cadena & Format(suma(0, 1), "###,###,##0") & vbTab
        For i = 2 To 7
            cadena = cadena & Format(suma(0, i), "$ ###,###,##0") & vbTab
        Next i
        'impresion.AddItem cadena, True
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    End If
End Sub


