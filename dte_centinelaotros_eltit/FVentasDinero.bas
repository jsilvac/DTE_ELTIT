Attribute VB_Name = "FVentasDinero"
Option Explicit
    Private deptos(1 To 5, 0 To 9) As String
    Private totales(10) As Double
    Private totales2(10) As Double
    Private listaempresa As String
    
Public Sub generaInformeVD(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    
    deptos(1, 1) = "3"
    deptos(1, 2) = "4"
    deptos(1, 3) = "7"
    deptos(1, 4) = "5"
    deptos(1, 5) = "6"
    deptos(1, 6) = "3"
    deptos(1, 7) = "0"
    deptos(1, 8) = "0"
    deptos(1, 9) = "8"
    
    deptos(2, 1) = "01"
    deptos(2, 2) = "02"
    deptos(2, 3) = "03"
    deptos(2, 4) = "04"
    deptos(2, 5) = "05"
    deptos(2, 6) = "06"
    deptos(2, 7) = "07"
    deptos(2, 8) = "08"
    deptos(2, 9) = "09"
    
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("RESUMEN DE VENTAS POR DINERO - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    Call resumenVentas(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function resumenVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    Dim csql As rdoQuery
    Dim resultado As rdoResultset
    Dim orden As String
    Dim i As Integer

    rubAux = rubro
    tabla = "SELECT IF(dc.tipo = 'FV', 1, IF(dc.tipo = 'NV', 2, IF(dc.tipo = 'BV' OR dc.tipo = 'ZE', 4, IF(dc.tipo = 'FE', 3, 0)))) AS orden, dc.local, dc.tipo, SUM(dc.neto) AS neto, SUM(dc.iva) AS iva, SUM(dc.impuestoharina) AS iha, SUM(dc.total) AS total, COUNT(numero) AS cantidad "
    tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc "
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.tipo <> 'GD' AND dc.tipo <> 'GM' AND dc.tipo <> 'FM' "
    tabla = tabla & "GROUP BY local, orden ORDER BY orden, local "
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    
    If data.Recordset.RecordCount > 0 Then
        cadena = "SUCURSAL" & vbTab
        cadena = cadena & "DOC. EMITID." & vbTab
        cadena = cadena & "NETO HARINAS" & vbTab
        cadena = cadena & "NETO SUBPRODUC." & vbTab
        cadena = cadena & "NETO ENVASES" & vbTab
        cadena = cadena & "NETO TRIGO" & vbTab
        cadena = cadena & "NETO MAQUILA" & vbTab
        cadena = cadena & "NETO OTROS" & vbTab
        cadena = cadena & "TOTAL NETOS" & vbTab
        cadena = cadena & "TOTAL IVA" & vbTab
        cadena = cadena & "TOTAL RETENCION" & vbTab
        cadena = cadena & "TOTAL GENERAL" & vbTab
        impresion.AddItem cadena, True
        impresion.RowHeight(impresion.Rows - 1) = impresion.DefaultRowHeight * 1.75
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).WrapText = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin

        data.Recordset.MoveFirst
        'listaempresa = data.Recordset.Fields(0)
        orden = data.Recordset.Fields(0)
        
        While Not data.Recordset.EOF
            If orden <> data.Recordset.Fields(0) Then
                Call imprimetotaltipo(impresion, orden)
                orden = data.Recordset.Fields(0)
            End If
            
            listaempresa = data.Recordset.Fields(1)
            Select Case orden
                Case "1"
                    TIPO = "dc.tipo = 'FV'"
                Case "2"
                    TIPO = "dc.tipo = 'NV'"
                Case "4"
                    TIPO = "(dc.tipo = 'BV' OR dc.tipo = 'ZE')"
                Case "3"
                    TIPO = "dc.tipo = 'FE'"
            End Select
            Set csql = New rdoQuery
            Set csql.ActiveConnection = ventasRubro
            
            Rem If numeroDoc = "0001622156" Then Stop
            
            'cSql.sql = "SELECT md.codigodepto, mpf.codigoimpuesto, ROUND(SUM(dd.total - dd.total * dd.descuento / 100),0) AS monto, md.nombre "
            'cSql.sql = "SELECT md.codigodepto, mpf.codigoimpuesto, IF(dd.tipo = 'BV' OR dd.tipo = 'ZE',ROUND(SUM(dd.total / (1 + " & Replace(iva / 100, ",", ".") & ")),0), dd.total) AS monto, md.nombre "
            csql.sql = "SELECT md.codigodepto, mpf.codigoimpuesto, ROUND(SUM(dd.total),0) AS monto, md.nombre, dc.tipo "
            csql.sql = csql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_detalle_" + empresaActiva + " AS dd ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & basedatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & basedatos & rubro & ".r_maestrodepartamentos_" & rubro & " As md ON mpf.codigoseccion = md.codigoseccion AND mpf.codigodepto = md.codigodepto "
            csql.sql = csql.sql & "WHERE dc.local = '" & listaempresa & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
            csql.sql = csql.sql & "GROUP BY dc.numero, codigodepto "
            csql.sql = csql.sql & "ORDER BY dc.numero, codigodepto ASC "
            
            csql.Execute
            
            If csql.RowsAffected > 0 Then
                harinas = 0
                subproductos = 0
                envases = 0
                trigo = 0
                maquila = 0
                otros = 0
                Set resultado = csql.OpenResultset
                While Not resultado.EOF
                    tipoDoc = resultado("tipo")
                    Select Case resultado(0)
                        Case "00001", "00101"   'HARINAS
                            If tipoDoc = "BV" Or tipoDoc = "ZE" Then
                                harinas = harinas + Round(resultado(2) / (1 + iva / 100), 0)
                            Else
                                harinas = harinas + resultado(2)
                            End If
                        Case "00002", "00102"   'SUBPRODUCTOS
                            If tipoDoc = "BV" Or tipoDoc = "ZE" Then
                                subproductos = subproductos + Round(resultado(2) / (1 + iva / 100), 0)
                            Else
                                subproductos = subproductos + resultado(2)
                            End If
                        Case "00004", "00104"   'ENVASES
                            If tipoDoc = "BV" Or tipoDoc = "ZE" Then
                                envases = envases + Round(resultado(2) / (1 + iva / 100), 0)
                            Else
                                envases = envases + resultado(2)
                            End If
                        Case "00005", "00105"   'TRIGO
                            If tipoDoc = "BV" Or tipoDoc = "ZE" Then
                                trigo = trigo + Round(resultado(2) / (1 + iva / 100), 0)
                            Else
                                trigo = trigo + resultado(2)
                            End If
                        Case "00003", "00103"   'MAQUILA
                            If tipoDoc = "BV" Or tipoDoc = "ZE" Then
                                maquila = maquila + Round(resultado(2) / (1 + iva / 100), 0)
                            Else
                                maquila = maquila + resultado(2)
                            End If
                        Case Else   'OTROS
                            If tipoDoc = "BV" Or tipoDoc = "ZE" Then
                                otros = otros + Round(resultado(2) / (1 + iva / 100), 0)
                            Else
                                otros = otros + resultado(2)
                            End If
                    End Select
                    resultado.MoveNext
                Wend
            End If

            csql.Close
            
            cadena = data.Recordset.Fields(1) & " " & leerNombreEmpresa(data.Recordset.Fields(1)) & vbTab
            cadena = cadena & data.Recordset.Fields(7) & vbTab
            cadena = cadena & harinas & vbTab
            cadena = cadena & subproductos & vbTab
            cadena = cadena & envases & vbTab
            cadena = cadena & trigo & vbTab
            cadena = cadena & maquila & vbTab
            cadena = cadena & otros & vbTab
            cadena = cadena & data.Recordset.Fields(3) & vbTab
            cadena = cadena & data.Recordset.Fields(4) & vbTab
            cadena = cadena & data.Recordset.Fields(5) & vbTab
            cadena = cadena & data.Recordset.Fields(6) & vbTab
            
            totales(0) = totales(0) + CDbl(data.Recordset.Fields(7))
            totales(1) = totales(1) + harinas
            totales(2) = totales(2) + subproductos
            totales(3) = totales(3) + envases
            totales(4) = totales(4) + trigo
            totales(5) = totales(5) + maquila
            totales(6) = totales(6) + otros
            totales(7) = totales(7) + CDbl(data.Recordset.Fields(3))
            totales(8) = totales(8) + CDbl(data.Recordset.Fields(4))
            totales(9) = totales(9) + CDbl(data.Recordset.Fields(5))
            totales(10) = totales(10) + CDbl(data.Recordset.Fields(6))
            
            totales2(1) = totales2(1) + harinas
            totales2(2) = totales2(2) + subproductos
            totales2(3) = totales2(3) + envases
            totales2(4) = totales2(4) + trigo
            totales2(5) = totales2(5) + maquila
            totales2(6) = totales2(6) + otros
            totales2(7) = totales2(7) + CDbl(data.Recordset.Fields(3))
            totales2(8) = totales2(8) + CDbl(data.Recordset.Fields(4))
            totales2(9) = totales2(9) + CDbl(data.Recordset.Fields(5))
            totales2(10) = totales2(10) + CDbl(data.Recordset.Fields(6))
            impresion.AddItem cadena, True
            
            harinas = 0
            subproductos = 0
            envases = 0
            trigo = 0
            maquila = 0
            otros = 0
            
            data.Recordset.MoveNext
        Wend
    
    End If
    Call imprimetotaltipo(impresion, orden)
    
    impresion.AddItem "", True
    cadena = "TOTAL GENERAL " & vbTab & vbTab
    cadena = cadena & totales2(1) & vbTab
    cadena = cadena & totales2(2) & vbTab
    cadena = cadena & totales2(3) & vbTab
    cadena = cadena & totales2(4) & vbTab
    cadena = cadena & totales2(5) & vbTab
    cadena = cadena & totales2(6) & vbTab
    cadena = cadena & totales2(7) & vbTab
    cadena = cadena & totales2(8) & vbTab
    cadena = cadena & totales2(9) & vbTab
    cadena = cadena & totales2(10) & vbTab
    impresion.AddItem cadena, True
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.AddItem "", True
End Function

Private Sub imprimetotaltipo(ByRef impresion As Grid, ByVal TIPO As String)
    Dim cadena As String
    Dim i As Integer
    Select Case TIPO
        Case "1"
            cadena = "TOTAL FACTURAS " & vbTab
        Case "2"
            cadena = "TOTAL NOTAS CREDITO " & vbTab
        Case "3"
            cadena = "TOTAL FACTURAS EXENTAS " & vbTab
        Case "4"
            cadena = "TOTAL BOLETAS " & vbTab
    End Select
    cadena = cadena & totales(0) & vbTab
    cadena = cadena & totales(1) & vbTab
    cadena = cadena & totales(2) & vbTab
    cadena = cadena & totales(3) & vbTab
    cadena = cadena & totales(4) & vbTab
    cadena = cadena & totales(5) & vbTab
    cadena = cadena & totales(6) & vbTab
    cadena = cadena & totales(7) & vbTab
    cadena = cadena & totales(8) & vbTab
    cadena = cadena & totales(9) & vbTab
    cadena = cadena & totales(10) & vbTab
    impresion.AddItem cadena, True
    For i = 0 To 10
        totales(i) = 0
    Next i
    'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.AddItem "", True
End Sub




Private Function resumenVentas1(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim CODIGO As String
    Dim codigoDepto As String
    Dim i As Integer
    Dim cadena As String
    Dim numloc As String
    Dim orden As String
    Dim depto As Integer
    Dim pos As Integer
    
    rubAux = rubro
    tabla = "SELECT IF(dc.tipo = 'FV', 1, IF(dc.tipo = 'NV', 2, IF(dc.tipo = 'BV' OR dc.tipo = 'ZE', 3, IF(dc.tipo = 'FE', 4, 0)))) AS orden, dd.local, dd.tipo, ROUND(SUM(dd.total),0) AS monto, md.codigodepto, md.nombre "
    tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_detalle_" + empresaActiva + " AS dd ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & basedatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra INNER JOIN " & basedatos & rubAux & ".r_maestrodepartamentos_" & rubAux & " As md ON mpf.codigoseccion = md.codigoseccion AND mpf.codigodepto = md.codigodepto "
    tabla = tabla & "WHERE dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "GROUP BY local, orden, codigodepto "
    tabla = tabla & "ORDER BY orden, local, codigodepto ASC "
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        'TITULO
        impresion.AddItem "RESUMEN DE VENTAS POR DINERO", True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.AddItem "", True
        
        cadena = "SUCURSAL" & vbTab
        cadena = cadena & "DOC. EMITID." & vbTab
        cadena = cadena & "NETO HARINAS" & vbTab
        cadena = cadena & "NETO SUBPRODUC." & vbTab
        cadena = cadena & "NETO ENVASES" & vbTab
        cadena = cadena & "NETO TRIGO" & vbTab
        cadena = cadena & "NETO MAQUILA" & vbTab
        cadena = cadena & "NETO OTROS" & vbTab
        cadena = cadena & "TOTAL NETOS" & vbTab
        cadena = cadena & "TOTAL IVA" & vbTab
        cadena = cadena & "TOTAL RETENCION" & vbTab
        cadena = cadena & "TOTAL GENERAL" & vbTab
        impresion.AddItem cadena, True
        impresion.RowHeight(impresion.Rows - 1) = impresion.DefaultRowHeight * 1.75
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).WrapText = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        For i = 0 To 9
            deptos(3, i) = "0"
            deptos(4, i) = "0"
            deptos(5, i) = "0"
        Next i
        
        data.Recordset.MoveFirst
        orden = data.Recordset.Fields("orden")
        numloc = data.Recordset.Fields("local")
        While Not data.Recordset.EOF
            codigoDepto = Right(data.Recordset.Fields("codigodepto"), 2)
            If orden = data.Recordset.Fields("orden") Then
                If numloc = data.Recordset.Fields("local") Then
                    For i = 1 To 9
                        If codigoDepto = deptos(2, i) Then
                            pos = CDbl(deptos(1, i))
                            deptos(3, pos) = CDbl(deptos(3, pos)) + CDbl(data.Recordset.Fields("monto"))
                            deptos(4, pos) = CDbl(deptos(4, pos)) + CDbl(data.Recordset.Fields("monto"))
                            deptos(5, pos) = CDbl(deptos(5, pos)) + CDbl(data.Recordset.Fields("monto"))
                            Exit For
                        End If
                    Next i
                Else
                    deptos(3, 1) = numloc & " " & leerNombreEmpresa(numloc)
                    deptos(3, 2) = leerDocumentosLocal(numloc, rubAux, orden, fecha1, fecha2)
                    cadena = ""
                    For i = 1 To 9
                        If i > 2 Then
                            Select Case orden
                                Case "1"
                                    cadena = cadena & deptos(3, i) & vbTab
                                Case "2"
                                    cadena = cadena & 0 - CDbl(deptos(3, i)) & vbTab
                                Case "3"
                                    cadena = cadena & Format(CDbl(deptos(3, i)) / (1 + iva / 100) + 0.1, "########0") & vbTab
                                Case "4"
                                    cadena = cadena & deptos(3, i) & vbTab
                            End Select
                        Else
                            cadena = cadena & deptos(3, i) & vbTab
                        End If
                    Next i
                    impresion.AddItem cadena, True
                    impresion.Cell(impresion.Rows - 1, 0).text = orden
                    For i = 1 To 9
                        deptos(3, i) = "0"
                    Next i
                    numloc = data.Recordset.Fields("local")
                    data.Recordset.MovePrevious
                End If
            Else
                deptos(3, 1) = numloc & " " & leerNombreEmpresa(numloc)
                deptos(3, 2) = leerDocumentosLocal(numloc, rubAux, orden, fecha1, fecha2)
                cadena = ""
                For i = 1 To 9
                    cadena = cadena & deptos(3, i) & vbTab
                Next i
                impresion.AddItem cadena, True
                impresion.Cell(impresion.Rows - 1, 0).text = orden
                
                Select Case orden
                    Case "1"
                        cadena = "FACTURAS"
                    Case "2"
                        cadena = "N.CREDITO"
                    Case "3"
                        cadena = "BOLETAS"
                    Case "4"
                        cadena = "EXENTAS"
                End Select
                
                deptos(4, 1) = "TOTAL GENERAL " & cadena
                deptos(4, 2) = ""
                cadena = ""
                For i = 1 To 9
                    cadena = cadena & deptos(4, i) & vbTab
                Next i
                impresion.AddItem cadena, True
                impresion.Cell(impresion.Rows - 1, 0).text = orden
                impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                impresion.AddItem "", True
                For i = 0 To 9
                    deptos(3, i) = "0"
                    deptos(4, i) = "0"
                Next i
                numloc = data.Recordset.Fields("local")
                orden = data.Recordset.Fields("orden")
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        
        deptos(3, 1) = numloc & " " & leerNombreEmpresa(numloc)
        deptos(3, 2) = leerDocumentosLocal(numloc, rubAux, orden, fecha1, fecha2)
        cadena = ""
        For i = 1 To 9
            cadena = cadena & deptos(3, i) & vbTab
        Next i
        impresion.AddItem cadena, True
        impresion.Cell(impresion.Rows - 1, 0).text = orden
        'impresion.AddItem "", True
        
        Select Case orden
            Case "1"
                cadena = "FACTURAS"
            Case "2"
                cadena = "N.CREDITO"
            Case "3"
                cadena = "BOLETAS"
            Case "4"
                cadena = "EXENTAS"
        End Select
        
        deptos(4, 1) = "TOTAL GENERAL " & cadena
        deptos(4, 2) = ""
        cadena = ""
        For i = 1 To 9
            cadena = cadena & deptos(4, i) & vbTab
        Next i
        impresion.AddItem cadena, True
        impresion.Cell(impresion.Rows - 1, 0).text = orden
        impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.AddItem "", True
        
        deptos(5, 1) = "TOTAL GENERAL"
        deptos(5, 2) = ""
        cadena = ""
        For i = 1 To 9
            cadena = cadena & deptos(5, i) & vbTab
        Next i
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
    End If
    Call sumaGrilla(impresion)
End Function

Private Sub sumaGrilla(ByRef impresion As Grid)
    Dim i As Long
    Dim j As Long
    Dim totalNeto As Double
    Dim totalIVA As Double
    Dim totalIHA As Double
    Dim total As Double
    Dim escribir As Boolean
    
    For i = 1 To impresion.Rows - 1
        totalNeto = 0
        totalIVA = 0
        totalIHA = 0
        total = 0
        For j = 1 To 8
            If IsNumeric(impresion.Cell(i, j).text) = True Then
                totalNeto = totalNeto + CDbl(impresion.Cell(i, j).text)
                escribir = True
            Else
                escribir = False
            End If
        Next j
        Select Case impresion.Cell(i, 0).text
            Case "1", "2"
                totalIVA = Round(totalNeto * iva / 100 + 0.1, 0)
                totalIHA = Round(totalNeto * iha / 100 + 0.1, 0)
            Case "3"
                totalIVA = Round(totalNeto - totalNeto / (1 + iva / 100) + 0.1, 0)
                totalIHA = 0
            Case "4"
                totalIVA = 0
                totalIHA = 0
        End Select
        total = totalNeto + totalIVA + totalIHA
        If escribir = True Then
            impresion.Cell(i, 9).text = totalNeto
            impresion.Cell(i, 10).text = totalIVA
            impresion.Cell(i, 11).text = totalIHA
            impresion.Cell(i, 12).text = total
        End If
    Next i
    totalIVA = 0
    totalNeto = 0
    total = 0
    For i = 1 To impresion.Rows - 1
        If IsNumeric(impresion.Cell(i, 10).text) = True And impresion.Cell(i, 0).text <> "" Then
            If impresion.Cell(i, 10).Font.Bold = True Then
                totalIVA = totalIVA + CDbl(impresion.Cell(i, 10).text)
                totalIHA = totalIHA + CDbl(impresion.Cell(i, 11).text)
                total = total + CDbl(impresion.Cell(i, 12).text)
            End If
        End If
    Next i
    If i > 1 Then
        impresion.Cell(i - 1, 10).text = totalIVA
        impresion.Cell(i - 1, 11).text = totalIHA
        impresion.Cell(i - 1, 12).text = total
    End If
End Sub

Private Function leerDocumentosLocal(ByVal numloc As String, ByVal rubAux As String, ByVal orden As String, ByVal fecha1 As String, ByVal fecha2 As String) As String
    
    Dim CAMPOS(2, 3) As String
    Dim op As Integer
    Dim TIPO As String
    Set sql = New sqlventas.sqlventa
    CAMPOS(0, 0) = "IFNULL(COUNT(*),0)"
    CAMPOS(1, 0) = ""
    
    CAMPOS(0, 2) = baseVentas & rubAux & ".sv_documento_cabeza_" + empresaActiva + " AS dd"
    
    Select Case orden
        Case "1"
            TIPO = "tipo = 'FV'"
        Case "2"
            TIPO = "tipo = 'NV'"
        Case "3"
            TIPO = "(tipo = 'BV' OR tipo = 'ZE')"
        Case "4"
            TIPO = "tipo = 'FE'"
    End Select
    condicion = "dd.local = '" & numloc & "' AND dd.nula = 'N' AND " & TIPO & " AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
    op = 5
    sql.response = CAMPOS
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        leerDocumentosLocal = sql.response(0, 3)
    Else
        leerDocumentosLocal = ""
    End If
End Function





