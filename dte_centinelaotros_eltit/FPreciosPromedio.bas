Attribute VB_Name = "FPreciosPromedio"
Option Explicit
    
Public Sub generaInformePP(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("LISTADO DE PRECIOS PROMEDIO - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, Impresion)
    If detalle = True Then
        altoFila = Impresion.DefaultRowHeight
        ultimaFila = PreciosPromedioDetalle(data, Impresion, TIPO, fecha1, fecha2)
        For i = 1 To ultimaFila
            Impresion.RowHeight(i) = 0
        Next i
        Call Impresion.HPageBreaks.Add(Impresion.Rows - 1)
    End If
    Call PreciosPromedioLinea(data, Impresion, TIPO, fecha1, fecha2)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Function PreciosPromedioDetalle(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim codigo As String
    Dim i As Integer
    Dim j As Integer
    Dim cont As Integer
    Dim promKilos As Double
    Dim promMonto As Double
    Dim cadena As String
    Dim cad As String
    Dim numloc As Integer
    
    tabla = "SELECT dd.local, dd.tipo, dd.codigo, dd.descripcion, @kilos := @kilos + SUM(IF(dd.tipo = 'NV', -1 * dd.unidades, dd.unidades)) AS kilos, @monto := @monto + SUM(IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(dd.total / (1 + " & iva & " / 100),0), dd.total)) AS monto, mpf.codigolinea, mpf.codigodepto, @monto:=0, @kilos:=0 "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & baseDatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON dd.codigo = mpf.codigobarra "
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " AND (mpf.codigodepto = '00001' OR mpf.codigodepto = '00002' OR mpf.codigodepto = '00101' OR mpf.codigodepto = '00102') "
    tabla = tabla & "GROUP BY local, codigo "
    tabla = tabla & "ORDER BY codigo, local ASC "
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    data.Recordset.Requery
    If data.Recordset.RecordCount > 0 Then
        For i = 0 To cantLocales
            cabezaLocales(1, i) = "0"
            cabezaLocales(2, i) = "0"
            cabezaLocales(3, i) = "0"
        Next i
        
        data.Recordset.MoveFirst
        codigo = data.Recordset.Fields("codigo")
        cabezaLocales(1, 1) = data.Recordset.Fields("codigo")
        cabezaLocales(1, 2) = data.Recordset.Fields("descripcion")
        While Not data.Recordset.EOF
            If codigo = data.Recordset.Fields("codigo") Then
                numloc = CDbl(data.Recordset.Fields("local")) + 3
                cabezaLocales(1, numloc) = Format(data.Recordset.Fields("         "), "###,###,##0")
                cabezaLocales(2, numloc) = Format(data.Recordset.Fields("monto"), "$ ###,###,##0")
                cabezaLocales(3, numloc) = Format(Round(CDbl(data.Recordset.Fields("monto")) / CDbl(data.Recordset.Fields("         ")), 2), "$ ###,###,##0")
            Else
                promKilos = 0
                promMonto = 0
                For j = 1 To 3
                    cadena = ""
                    cont = 0
                    Select Case j
                        Case 1
                            cad = "         "
                        Case 2
                            cad = "MONTO"
                        Case 3
                            cad = "PROMEDIO"
                    End Select
                    For i = 1 To cantLocales - 1
                        If i = 3 Then
                            cadena = cadena & cad & vbTab
                        End If
                        cadena = cadena & cabezaLocales(j, i) & vbTab
                        If i >= 3 And Val(Format(cabezaLocales(j, i), "########0")) <> 0 Then
                            If j = 1 Then
                                promKilos = promKilos + CDbl(cabezaLocales(j, i))
                            End If
                            If j = 2 Then
                                promMonto = promMonto + CDbl(cabezaLocales(j, i))
                            End If
                            cont = cont + 1
                        End If
                    Next i
                    Select Case j
                        Case 1
                            promKilos = promKilos / cont
                            cadena = cadena & Format(promKilos, "###,###,##0.00")
                        Case 2
                            promMonto = promMonto / cont
                            cadena = cadena & Format(promMonto, "$ ###,###,##0.00")
                        Case 3
                            cadena = cadena & Format(promMonto / promKilos, "$ ###,###,##0.00")
                            promKilos = 0
                            promMonto = 0
                    End Select
                    Impresion.AddItem cadena, True
                Next j
                Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 1).Merge
                Impresion.Range(Impresion.Rows - 3, 2, Impresion.Rows - 1, 2).Merge
                Impresion.Range(Impresion.Rows - 3, 2, Impresion.Rows - 1, 2).WrapText = True
                Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellDot
                
                For i = 0 To cantLocales
                    cabezaLocales(1, i) = "0"
                    cabezaLocales(2, i) = "0"
                    cabezaLocales(3, i) = "0"
                Next i
                
                codigo = data.Recordset.Fields("codigo")
                cabezaLocales(1, 1) = data.Recordset.Fields("codigo")
                cabezaLocales(1, 2) = data.Recordset.Fields("descripcion")
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        promKilos = 0
        promMonto = 0
        For j = 1 To 3
            cadena = ""
            cont = 0
            Select Case j
                Case 1
                    cad = "         "
                Case 2
                    cad = "MONTO"
                Case 3
                    cad = "PROMEDIO"
            End Select
            For i = 1 To cantLocales - 1
                If i = 3 Then
                    cadena = cadena & cad & vbTab
                End If
                cadena = cadena & cabezaLocales(j, i) & vbTab
                If i >= 3 And Val(Format(cabezaLocales(j, i), "########0")) <> 0 Then
                    If j = 1 Then
                        promKilos = promKilos + CDbl(cabezaLocales(j, i))
                    End If
                    If j = 2 Then
                        promMonto = promMonto + CDbl(cabezaLocales(j, i))
                    End If
                    cont = cont + 1
                End If
            Next i
            Select Case j
                Case 1
                    promKilos = promKilos / cont
                    cadena = cadena & Format(promKilos, "###,###,##0.00")
                Case 2
                    promMonto = promMonto / cont
                    cadena = cadena & Format(promMonto, "$ ###,###,##0.00")
                Case 3
                    cadena = cadena & Format(promMonto / promKilos, "$ ###,###,##0.00")
                    promKilos = 0
                    promMonto = 0
            End Select
            Impresion.AddItem cadena, True
        Next j
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 1).Merge
        Impresion.Range(Impresion.Rows - 3, 2, Impresion.Rows - 1, 2).Merge
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellDot
    End If
    Impresion.AddItem "", True
    PreciosPromedioDetalle = Impresion.Rows - 1
End Function

Private Sub PreciosPromedioLinea(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim codigo As String
    Dim i As Integer
    Dim j As Integer
    Dim cont As Integer
    Dim promKilos As Double
    Dim promMonto As Double
    Dim cadena As String
    Dim cad As String
    Dim numloc As Integer
    Dim codigo2 As String
    Dim sumadores(10) As Double
    Dim sumadores2(10) As Double
    tabla = "SELECT dd.local, dd.tipo, ml.codigolinea AS codigo, ml.nombre AS descripcion, @kilos := @kilos + SUM(IF(dd.tipo = 'NV', -1 * dd.unidades, dd.unidades)) AS kilos, @monto := @monto + SUM(IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(dd.total / (1 + 19 / 100),0), dd.total)) AS monto, mpf.codigolinea, mpf.codigodepto, @monto:=0, @kilos:=0 "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & baseDatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & baseDatos & rubro & ".r_maestrolineas_" & rubro & " AS ml ON mpf.codigoseccion = ml.codigoseccion AND mpf.codigodepto = ml.codigodepto AND mpf.codigolinea = ml.codigolinea "
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "AND (mpf.codigodepto = '00001' OR mpf.codigodepto = '00002' OR mpf.codigodepto = '00101' OR mpf.codigodepto = '00102') "
    tabla = tabla & "GROUP BY local, codigodepto, codigo "
    tabla = tabla & "ORDER BY codigodepto, codigo, local ASC "
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    data.Recordset.Requery
    If data.Recordset.RecordCount > 0 Then
        For i = 0 To cantLocales
            cabezaLocales(1, i) = "0"
            cabezaLocales(2, i) = "0"
            cabezaLocales(3, i) = "0"
        Next i
        
        data.Recordset.MoveFirst
        codigo = data.Recordset.Fields("codigodepto") + data.Recordset.Fields("codigo")
        codigo2 = data.Recordset.Fields("codigodepto")
        cabezaLocales(1, 1) = data.Recordset.Fields("codigo")
        cabezaLocales(1, 2) = data.Recordset.Fields("descripcion")
        While Not data.Recordset.EOF
            
            If codigo = data.Recordset.Fields("codigodepto") + data.Recordset.Fields("codigo") Then
                numloc = CDbl(data.Recordset.Fields("local")) + 3
                cabezaLocales(1, numloc) = Format(data.Recordset.Fields("         "), "###,###,##0")
                cabezaLocales(2, numloc) = Format(data.Recordset.Fields("monto"), "$ ###,###,##0")
                cabezaLocales(3, numloc) = Format(Round(CDbl(data.Recordset.Fields("monto")) / CDbl(data.Recordset.Fields("         ")), 2), "$ ###,##0.00")
            Else
                promKilos = 0
                promMonto = 0
                For j = 1 To 3
                    cadena = ""
                    cont = 0
                    Select Case j
                        Case 1
                            cad = "         "
                        Case 2
                            cad = "MONTO"
                        Case 3
                            cad = "PROMEDIO"
                    End Select
                    For i = 1 To cantLocales - 1
                        If i = 3 Then
                            cadena = cadena & cad & vbTab
                        End If
                        cadena = cadena & cabezaLocales(j, i) & vbTab
                        If i >= 3 And Val(Format(cabezaLocales(j, i), "########0")) <> 0 Then
                            If j = 1 Then
                                promKilos = promKilos + CDbl(cabezaLocales(j, i))
                            End If
                            If j = 2 Then
                                promMonto = promMonto + CDbl(cabezaLocales(j, i))
                            End If
                            cont = cont + 1
                        End If
                    Next i
                    Select Case j
                        Case 1
                            promKilos = promKilos '/ cont
                            cadena = cadena & Format(promKilos, "###,###,##0.00")
                        Case 2
                            promMonto = promMonto '/ cont
                            cadena = cadena & Format(promMonto, "$ ###,###,##0.00")
                        Case 3
                            cadena = cadena & Format(promMonto / promKilos, "$ ###,###,##0.00")
                            
                            promKilos = 0
                            promMonto = 0
                    End Select
                    Impresion.AddItem cadena, True
                Next j
                Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 1).Merge
                Impresion.Range(Impresion.Rows - 3, 2, Impresion.Rows - 1, 2).Merge
                Impresion.Range(Impresion.Rows - 3, 2, Impresion.Rows - 1, 2).WrapText = True
                Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellDot
                
                For i = 0 To cantLocales
                    cabezaLocales(1, i) = "0"
                    cabezaLocales(2, i) = "0"
                    cabezaLocales(3, i) = "0"
                Next i
                
                If codigo2 <> data.Recordset.Fields("codigodepto") Then
                    'imprimetotal
                End If
                
                
                codigo = data.Recordset.Fields("codigodepto") + data.Recordset.Fields("codigo")
                codigo2 = data.Recordset.Fields("codigodepto")
                
                cabezaLocales(1, 1) = data.Recordset.Fields("codigo")
                cabezaLocales(1, 2) = data.Recordset.Fields("descripcion")
                data.Recordset.MovePrevious
            
            
            End If
            data.Recordset.MoveNext
        Wend
        promKilos = 0
        promMonto = 0
        For j = 1 To 3
            cadena = ""
            cont = 0
            Select Case j
                Case 1
                    cad = "         "
                Case 2
                    cad = "MONTO"
                Case 3
                    cad = "PROMEDIO"
            End Select
            For i = 1 To cantLocales - 1
                If i = 3 Then
                    cadena = cadena & cad & vbTab
                End If
                cadena = cadena & cabezaLocales(j, i) & vbTab
                If i >= 3 And Val(Format(cabezaLocales(j, i), "########0")) <> 0 Then
                    If j = 1 Then
                        promKilos = promKilos + CDbl(cabezaLocales(j, i))
                    End If
                    If j = 2 Then
                        promMonto = promMonto + CDbl(cabezaLocales(j, i))
                    End If
                    cont = cont + 1
                End If
            Next i
            Select Case j
                Case 1
                    promKilos = promKilos '/ cont
                    cadena = cadena & Format(promKilos, "###,###,##0.00")
                Case 2
                    promMonto = promMonto '/ cont
                    cadena = cadena & Format(promMonto, "$ ###,###,##0.00")
                Case 3
                    cadena = cadena & Format(promMonto / promKilos, "$ ###,###,##0.00")
                    promKilos = 0
                    promMonto = 0
            End Select
            Impresion.AddItem cadena, True
        Next j
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 1).Merge
        Impresion.Range(Impresion.Rows - 3, 2, Impresion.Rows - 1, 2).Merge
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellDot
    End If
    Impresion.AddItem "", True
    
    Dim total(4, 10) As Double
    Dim totalgeneral(2, 3) As Double
    
    For i = 0 To 10
        total(0, i) = 0
        total(1, i) = 0
        total(2, i) = 0
        total(3, i) = 0
        total(4, i) = 0
    Next i
    
    tabla = "SELECT dc.local, md.nombre AS descripcion, @kilos := @kilos + SUM(IF(dd.tipo = 'NV',IF(dd.codigo = '0000000000100' OR dd.codigo = '0000000000101' , 0, -1 * dd.unidades), dd.unidades)) AS kilos, @monto := @monto + SUM(IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(dd.total / (1 + 19 / 100),0), ROUND(dd.total,0))) AS monto, mpf.codigolinea, IF(mpf.codigodepto = '00001' OR mpf.codigodepto = '00101', '1', '2') AS codigodepto, @monto:=0, @kilos:=0 "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & baseDatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & baseDatos & rubro & ".r_maestrodepartamentos_" & rubro & " AS md ON mpf.codigoseccion = md.codigoseccion AND mpf.codigodepto = md.codigodepto "
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "AND (mpf.codigodepto = '00001' or mpf.codigodepto='00002' or mpf.codigodepto='00101' or mpf.codigodepto='00102') "
    tabla = tabla & "GROUP BY local, codigodepto "
    tabla = tabla & "ORDER BY local, codigodepto ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    data.Recordset.Requery
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        While Not data.Recordset.EOF
            numloc = Val(data.Recordset.Fields("local"))
            If data.Recordset.Fields("codigodepto") = "1" Then
                total(1, numloc) = total(1, numloc) + CDbl(data.Recordset.Fields("         "))
                total(2, numloc) = total(2, numloc) + CDbl(data.Recordset.Fields("monto"))
                totalgeneral(1, 1) = totalgeneral(1, 1) + CDbl(data.Recordset.Fields("         "))
                totalgeneral(1, 2) = totalgeneral(1, 2) + CDbl(data.Recordset.Fields("monto"))
            End If
            If data.Recordset.Fields("codigodepto") = "2" Then
                total(3, numloc) = total(3, numloc) + CDbl(data.Recordset.Fields("         "))
                total(4, numloc) = total(4, numloc) + CDbl(data.Recordset.Fields("monto"))
                totalgeneral(2, 1) = totalgeneral(2, 1) + CDbl(data.Recordset.Fields("         "))
                totalgeneral(2, 2) = totalgeneral(2, 2) + CDbl(data.Recordset.Fields("monto"))
            End If
            data.Recordset.MoveNext
        Wend
        
        cadena = "TOTAL GENERAL" & vbTab & vbTab & "         "
        For i = 0 To 10
            cadena = cadena & vbTab & Format(total(1, i), "###,###,##0")
        Next i
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        
        cadena = vbTab & vbTab & "MONTO"
        For i = 0 To 10
            cadena = cadena & vbTab & Format(total(2, i), "$ ###,###,##0")
        Next i
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 2, 1, Impresion.Rows - 1, 2).Merge
        Impresion.Range(Impresion.Rows - 2, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter

        Impresion.AddItem "", True
        Impresion.AddItem "", True
        
        totalgeneral(1, 3) = totalgeneral(1, 2) / totalgeneral(1, 1)
        totalgeneral(2, 3) = totalgeneral(2, 2) / totalgeneral(2, 1)
        
        Impresion.AddItem "PROMEDIO GENERAL HARINAS " & vbTab & vbTab & vbTab & "         " & vbTab & Format(totalgeneral(1, 1), "###,###,##0"), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 4).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 6).Merge
        Impresion.AddItem "PROMEDIO GENERAL HARINAS " & vbTab & vbTab & vbTab & "MONTO" & vbTab & Format(totalgeneral(1, 2), "$ ###,###,##0"), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 4).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 6).Merge
        Impresion.AddItem "PROMEDIO GENERAL HARINAS " & vbTab & vbTab & vbTab & "PROM." & vbTab & Format(totalgeneral(1, 3), "$ ###,###,##0"), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 4).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 6).Merge
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 3).Merge
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 3).Alignment = cellLeftCenter
        
        Impresion.AddItem "", True
        
        Impresion.AddItem "PROMEDIO GENERAL SUBPRODUCTOS " & vbTab & vbTab & vbTab & "         " & vbTab & Format(totalgeneral(2, 1), "###,###,##0"), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 4).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 6).Merge
        Impresion.AddItem "PROMEDIO GENERAL SUBPRODUCTOS " & vbTab & vbTab & vbTab & "MONTO" & vbTab & Format(totalgeneral(2, 2), "$ ###,###,##0"), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 4).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 6).Merge
        Impresion.AddItem "PROMEDIO GENERAL SUBPRODUCTOS " & vbTab & vbTab & vbTab & "PROM." & vbTab & Format(totalgeneral(2, 3), "$ ###,###,##0"), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 4).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 6).Merge
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 3).Merge
        Impresion.Range(Impresion.Rows - 3, 1, Impresion.Rows - 1, 3).Alignment = cellLeftCenter
    End If
End Sub


