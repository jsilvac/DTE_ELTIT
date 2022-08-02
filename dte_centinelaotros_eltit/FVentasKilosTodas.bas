Attribute VB_Name = "FVentasKilosTodas"
Option Explicit
    
Public Sub generaInformeVK(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("ESTADISTICA DE VENTAS POR KILOS - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, Impresion)
    Call resumenDepto(data, Impresion, TIPO, fecha1, fecha2)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub resumenDepto(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim codigoLinea As String
    Dim codigoDepto As String
    Dim codigoSeccion As String
    Dim linea As String
    Dim tipoDoc As String
    Dim i As Integer
    Dim cadena As String
    Dim numloc As Integer
    Dim mostrar As Boolean
    Dim suma As Double
    
    tabla = "SELECT dd.local, IF(dd.tipo = 'FV' OR dd.tipo = 'FE', '1', IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', '2', IF(dd.tipo = 'NV', '3', '0'))) AS tipo, ml.nombre AS descripcion, IF(dd.tipo = 'NV', 0 - SUM(dd.unidades), SUM(dd.unidades)) AS kilos, mpf.codigolinea, mpf.codigodepto, mpf.codigoseccion "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & baseDatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & baseDatos & rubro & ".r_maestrolineas_" & rubro & " AS ml ON /*mpf.codigoseccion = ml.codigoseccion AND*/ mpf.codigodepto = ml.codigodepto AND mpf.codigolinea = ml.codigolinea "
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " AND (mpf.codigodepto = '00001' OR mpf.codigodepto = '00002' OR mpf.codigodepto = '00101' OR mpf.codigodepto = '00102') AND mpf.codigobarra <> '0000000000100' AND mpf.codigobarra <> '0000000000101' AND mpf.codigobarra <> '0000000000901' "
    tabla = tabla & "GROUP BY local, codigodepto, codigolinea, tipo "
    tabla = tabla & "ORDER BY codigodepto, codigolinea, tipo, local ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    
    If data.Recordset.RecordCount > 0 Then
        For i = 0 To cantLocales
            cabezaLocales(1, i) = "0"
            cabezaLocales(2, i) = "0"
            cabezaLocales(3, i) = "0"
            cabezaLocales(4, i) = "0"
            cabezaLocales(5, i) = "0"
            cabezaLocales(6, i) = "0"
            cabezaLocales(7, i) = "0"
            cabezaLocales(8, i) = "0"
        Next i
        
        data.Recordset.MoveFirst
        tipoDoc = data.Recordset.Fields("tipo")
        codigoLinea = data.Recordset.Fields("codigolinea")
        codigoDepto = data.Recordset.Fields("codigodepto")
        codigoSeccion = data.Recordset.Fields("codigoseccion")
        linea = data.Recordset.Fields("descripcion")
        cabezaLocales(0, 1) = codigoLinea
        cabezaLocales(0, 2) = data.Recordset.Fields("descripcion")
        suma = 0
        While Not data.Recordset.EOF
            If codigoDepto = data.Recordset.Fields("codigodepto") Then
                If codigoLinea = data.Recordset.Fields("codigolinea") Then
                    If tipoDoc = data.Recordset.Fields("tipo") Then
                        numloc = CDbl(data.Recordset.Fields("local")) + 3
                        cabezaLocales(1, numloc) = Format(data.Recordset.Fields("         "), "###,###,##0") 'para la linea
                        cabezaLocales(2, numloc) = Format(CDbl(cabezaLocales(2, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0") 'para el total por linea
                        Select Case tipoDoc
                            Case "1"
                                cabezaLocales(3, numloc) = Format(CDbl(cabezaLocales(3, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0") 'para el total por facturas
                                cabezaLocales(6, numloc) = Format(CDbl(cabezaLocales(6, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0") 'para el total general facturas
                            Case "2"
                                cabezaLocales(4, numloc) = Format(CDbl(cabezaLocales(4, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0") 'para el total por boletas
                                cabezaLocales(7, numloc) = Format(CDbl(cabezaLocales(7, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0") 'para el total general boletas
                            Case "3"
                                cabezaLocales(5, numloc) = Format(CDbl(cabezaLocales(5, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0") 'para el total por notas de credito
                                cabezaLocales(8, numloc) = Format(CDbl(cabezaLocales(8, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0") 'para el total general notas de credito
                        End Select
                    Else
                        cadena = codigoLinea & vbTab
                        Select Case tipoDoc
                            Case "1"
                                cadena = cadena & "FAC "
                            Case "2"
                                cadena = cadena & "BOL "
                            Case "3"
                                cadena = cadena & "NCR "
                        End Select
                        cadena = cadena & linea & vbTab
                        suma = 0
                        For i = 3 To cantLocales
                            suma = suma + CDbl(cabezaLocales(1, i))
                            cadena = cadena & cabezaLocales(1, i) & vbTab
                        Next i
                        cadena = cadena & Format(suma, "###,###,##0")
                        Impresion.AddItem cadena, True
                        For i = 0 To cantLocales
                            cabezaLocales(1, i) = "0"
                        Next i
                        tipoDoc = data.Recordset.Fields("tipo")
                        data.Recordset.MovePrevious
                    End If
                Else
                    cadena = codigoLinea & vbTab
                    Select Case tipoDoc
                        Case "1"
                            cadena = cadena & "FAC "
                        Case "2"
                            cadena = cadena & "BOL "
                        Case "3"
                            cadena = cadena & "NCR "
                    End Select
                    cadena = cadena & linea & vbTab
                    suma = 0
                    For i = 3 To cantLocales
                        suma = suma + CDbl(cabezaLocales(1, i))
                        cadena = cadena & cabezaLocales(1, i) & vbTab
                    Next i
                    cadena = cadena & Format(suma, "###,###,##0")
                    Impresion.AddItem cadena, True
                    
                    cadena = vbTab & vbTab
                    suma = 0
                    For i = 3 To cantLocales
                        suma = suma + CDbl(cabezaLocales(2, i))
                        cadena = cadena & cabezaLocales(2, i) & vbTab
                    Next i
                    cadena = cadena & Format(suma, "###,###,##0")
                    Impresion.AddItem cadena, True
                    Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                    Impresion.AddItem "", True
                    
                    For i = 0 To cantLocales
                        cabezaLocales(1, i) = "0"
                        cabezaLocales(2, i) = "0"
                    Next i
                    tipoDoc = data.Recordset.Fields("tipo")
                    codigoLinea = data.Recordset.Fields("codigolinea")
                    linea = data.Recordset.Fields("descripcion")
                    data.Recordset.MovePrevious
                End If
            Else
                cadena = codigoLinea & vbTab
                Select Case tipoDoc
                    Case "1"
                        cadena = cadena & "FAC "
                    Case "2"
                        cadena = cadena & "BOL "
                    Case "3"
                        cadena = cadena & "NCR "
                End Select
                cadena = cadena & linea & vbTab
                suma = 0
                For i = 3 To cantLocales
                    suma = suma + CDbl(cabezaLocales(1, i))
                    cadena = cadena & cabezaLocales(1, i) & vbTab
                Next i
                cadena = cadena & Format(suma, "###,###,##0")
                Impresion.AddItem cadena, True
                
                cadena = vbTab & vbTab
                suma = 0
                For i = 3 To cantLocales
                    suma = suma + CDbl(cabezaLocales(2, i))
                    cadena = cadena & cabezaLocales(2, i) & vbTab
                Next i
                cadena = cadena & Format(suma, "###,###,##0")
                Impresion.AddItem cadena, True
                Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                Impresion.AddItem "", True
                
                For i = 0 To cantLocales
                    cabezaLocales(1, i) = "0"
                Next i
                
                'TOTAL FACTURAS
                mostrar = False
                cadena = "TOTAL FAC " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
                suma = 0
                For i = 3 To cantLocales
                    If Val(cabezaLocales(3, i)) <> 0 Then
                        mostrar = True
                    End If
                    suma = suma + CDbl(cabezaLocales(3, i))
                    cadena = cadena & cabezaLocales(3, i) & vbTab
                    cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(3, i)), "###,###,##0")
                Next i
                If mostrar = True Then
                    cadena = cadena & Format(suma, "####,###,#0")
                    Impresion.AddItem cadena, True
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                End If
                'TOTAL BOLETAS
                mostrar = False
                cadena = "TOTAL BOL " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
                suma = 0
                For i = 3 To cantLocales
                    If Val(cabezaLocales(4, i)) <> 0 Then
                        mostrar = True
                    End If
                    suma = suma + CDbl(cabezaLocales(4, i))
                    cadena = cadena & cabezaLocales(4, i) & vbTab
                    cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(4, i)), "###,###,##0")
                Next i
                If mostrar = True Then
                    cadena = cadena & Format(suma, "###,###,##0")
                    Impresion.AddItem cadena, True
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                End If
                'TOTAL NOTAS CREDITO
                mostrar = False
                cadena = "TOTAL NCR " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
                suma = 0
                For i = 3 To cantLocales
                    If Val(cabezaLocales(5, i)) <> 0 Then
                        mostrar = True
                    End If
                    suma = suma + CDbl(cabezaLocales(5, i))
                    cadena = cadena & cabezaLocales(5, i) & vbTab
                    cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(5, i)), "###,###,##0")
                Next i
                If mostrar = True Then
                    cadena = cadena & Format(suma, "###,###,##0")
                    Impresion.AddItem cadena, True
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                End If
                'TOTAL PARCIAL POR DEPTO
                mostrar = False
                cadena = "TOTAL PARCIAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
                suma = 0
                For i = 3 To cantLocales
                    If Val(cabezaLocales(1, i)) <> 0 Then
                        mostrar = True
                    End If
                    suma = suma + CDbl(cabezaLocales(1, i))
                    cadena = cadena & cabezaLocales(1, i) & vbTab
                Next i
                If mostrar = True Then
                    cadena = cadena & Format(suma, "###,###,##0")
                    Impresion.AddItem cadena, True
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                    Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                    Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellDot
                End If
                
                Impresion.AddItem "", True
                
                If codigoDepto = "00002" Then
                    Impresion.RowHeight(Impresion.Rows - 1) = 0
                    Call Impresion.HPageBreaks.Add(Impresion.Rows - 1)
                End If
                
                For i = 0 To cantLocales
                    cabezaLocales(1, i) = "0"
                    cabezaLocales(2, i) = "0"
                    cabezaLocales(3, i) = "0"
                    cabezaLocales(4, i) = "0"
                    cabezaLocales(5, i) = "0"
                Next i
                tipoDoc = data.Recordset.Fields("tipo")
                codigoLinea = data.Recordset.Fields("codigolinea")
                linea = data.Recordset.Fields("descripcion")
                codigoDepto = data.Recordset.Fields("codigodepto")
                
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        cadena = codigoLinea & vbTab
        Select Case tipoDoc
            Case "1"
                cadena = cadena & "FAC "
            Case "2"
                cadena = cadena & "BOL "
            Case "3"
                cadena = cadena & "NCR "
        End Select
        cadena = cadena & linea & vbTab
        suma = 0
        For i = 3 To cantLocales
            suma = suma + CDbl(cabezaLocales(1, i))
            cadena = cadena & cabezaLocales(1, i) & vbTab
        Next i
        cadena = cadena & Format(suma, "###,###,##0")
        Impresion.AddItem cadena, True
        
        cadena = vbTab & vbTab
        suma = 0
        For i = 3 To cantLocales
            suma = suma + CDbl(cabezaLocales(2, i))
            cadena = cadena & cabezaLocales(2, i) & vbTab
        Next i
        cadena = cadena & Format(suma, "###,###,##0")
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        Impresion.AddItem "", True
        
        For i = 0 To cantLocales
            cabezaLocales(1, i) = "0"
        Next i
        
        'TOTAL FACTURAS
        mostrar = False
        cadena = "TOTAL FAC " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
        suma = 0
        For i = 3 To cantLocales
            If Val(cabezaLocales(3, i)) <> 0 Then
                mostrar = True
            End If
            suma = suma + CDbl(cabezaLocales(3, i))
            cadena = cadena & cabezaLocales(3, i) & vbTab
            cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(3, i)), "###,###,##0")
        Next i
        If mostrar = True Then
            cadena = cadena & Format(suma, "###,###,##0")
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        End If
        'TOTAL BOLETAS
        mostrar = False
        cadena = "TOTAL BOL " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
        suma = 0
        For i = 3 To cantLocales
            If Val(cabezaLocales(4, i)) <> 0 Then
                mostrar = True
            End If
            suma = suma + CDbl(cabezaLocales(4, i))
            cadena = cadena & cabezaLocales(4, i) & vbTab
            cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(4, i)), "###,###,##0")
        Next i
        If mostrar = True Then
            cadena = cadena & Format(suma, "###,###,##0")
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        End If
        'TOTAL NOTAS CREDITO
        mostrar = False
        cadena = "TOTAL NCR " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
        suma = 0
        For i = 3 To cantLocales
            If Val(cabezaLocales(5, i)) <> 0 Then
                mostrar = True
            End If
            suma = suma + CDbl(cabezaLocales(5, i))
            cadena = cadena & cabezaLocales(5, i) & vbTab
            cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(5, i)), "###,###,##0")
        Next i
        If mostrar = True Then
            cadena = cadena & Format(suma, "###,###,##0")
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        End If
        'TOTAL PARCIAL POR DEPTO
        mostrar = False
        cadena = "TOTAL PARCIAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubro) & vbTab & vbTab
        suma = 0
        For i = 3 To cantLocales
            If Val(cabezaLocales(1, i)) <> 0 Then
                mostrar = True
            End If
            suma = suma + CDbl(cabezaLocales(1, i))
            cadena = cadena & cabezaLocales(1, i) & vbTab
        Next i
        If mostrar = True Then
            cadena = cadena & Format(suma, "###,###,##0")
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellDot
        End If
        'impresion.AddItem "", True
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        ''TOTALES GENERALES
        'For i = 1 To cantLocales
        '    cabezaLocales(1, i) = "0"
        'Next i
        'TOTAL FACTURAS
        'mostrar = False
        'cadena = "TOTAL GENERAL FACTURAS " & vbTab & vbTab
        'suma = 0
        'For i = 3 To cantLocales
        '    If Val(cabezaLocales(6, i)) <> 0 Then
        '        mostrar = True
        '    End If
        '    suma = suma + CDbl(cabezaLocales(6, i))
        '    cadena = cadena & cabezaLocales(6, i) & vbTab
        '    cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(6, i)), "###,###,##0")
        'Next i
        'If mostrar = True Then
        '    cadena = cadena & Format(suma, "###,###,##0")
        '    impresion.AddItem cadena, True
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellLeftCenter
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        'End If
        ''TOTAL BOLETAS
        'mostrar = False
        'cadena = "TOTAL GENERAL BOLETAS " & vbTab & vbTab
        'suma = 0
        'For i = 3 To cantLocales
        '    If Val(cabezaLocales(7, i)) <> 0 Then
        '        mostrar = True
        '    End If
        '    suma = suma + CDbl(cabezaLocales(7, i))
        '    cadena = cadena & cabezaLocales(7, i) & vbTab
        '    cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(7, i)), "###,###,##0")
        'Next i
        'If mostrar = True Then
        '    cadena = cadena & Format(suma, "###,###,##0")
        '    impresion.AddItem cadena, True
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellLeftCenter
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        'End If
        ''TOTAL NOTAS CREDITO
        'mostrar = False
        'cadena = "TOTAL GENERAL NOTAS DE CREDITO " & vbTab & vbTab
        'suma = 0
        'For i = 3 To cantLocales
        '    If Val(cabezaLocales(8, i)) <> 0 Then
        '        mostrar = True
        '    End If
        '    suma = suma + CDbl(cabezaLocales(8, i))
        '    cadena = cadena & cabezaLocales(8, i) & vbTab
        '    cabezaLocales(1, i) = Format(CDbl(cabezaLocales(1, i)) + CDbl(cabezaLocales(8, i)), "###,###,##0")
        'Next i
        'If mostrar = True Then
        '    cadena = cadena & Format(suma, "###,###,##0")
        '    impresion.AddItem cadena, True
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellLeftCenter
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        'End If
        'TOTAL GENERAL
        'mostrar = False
        'cadena = "TOTAL GENERAL " & vbTab & vbTab
        'suma = 0
        'For i = 3 To cantLocales
        '    If Val(cabezaLocales(1, i)) <> 0 Then
        '        mostrar = True
        '    End If
        '    suma = suma + CDbl(cabezaLocales(1, i))
        '    cadena = cadena & cabezaLocales(1, i) & vbTab
        'Next i
        'If mostrar = True Then
        '    cadena = cadena & Format(suma, "###,###,##0")
        '    impresion.AddItem cadena, True
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellLeftCenter
        '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        '    impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        'End If
    End If
End Sub


