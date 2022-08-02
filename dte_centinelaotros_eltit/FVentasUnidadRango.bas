Attribute VB_Name = "FVentasUnidadRango"
Option Explicit
    
Public Sub generaInformeVUR(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("RESUMEN DE VENTAS POR UNIDAD POR RANGO - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, Impresion)
    Call resumenVentas(data, Impresion, TIPO, codLoc, fecha1, fecha2)
    Call resumenDepto(data, Impresion, TIPO, codLoc, fecha1, fecha2)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Function resumenVentas(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim codigo As String
    Dim i As Integer
    Dim cadena As String
    Dim numloc As Integer
    
    rubAux = rubro
    tabla = "SELECT dd.local, dd.codigo, dd.descripcion, SUM(dd.unidades) AS kilos "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' "
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "GROUP BY local, codigo "
    tabla = tabla & "ORDER BY codigo, local ASC "
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        'TITULO
        Impresion.AddItem "RESUMEN DE VENTAS POR KILOS", True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.AddItem "", True
        
        cadena = "CODIGO" & vbTab
        cadena = cadena & "DESCRIPCION" & vbTab
        
        For i = 3 To cantLocales
            cadena = cadena & cabezaLocales(0, i) & vbTab
        Next i
        
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        For i = 0 To cantLocales
            cabezaLocales(1, i) = "0"
            cabezaLocales(2, i) = "0"
        Next i
        
        data.Recordset.MoveFirst
        codigo = data.Recordset.Fields("codigo")
        cabezaLocales(1, 1) = data.Recordset.Fields("codigo")
        cabezaLocales(1, 2) = data.Recordset.Fields("descripcion")
        While Not data.Recordset.EOF
            If codigo = data.Recordset.Fields("codigo") Then
                numloc = CDbl(data.Recordset.Fields("local")) + 3
                cabezaLocales(1, numloc) = Format(data.Recordset.Fields("         "), "###,###,##0")
                cabezaLocales(2, numloc) = Format(CDbl(cabezaLocales(2, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0")
            Else
                cadena = ""
                For i = 1 To cantLocales
                    cadena = cadena & cabezaLocales(1, i) & vbTab
                Next i
                Impresion.AddItem cadena, True
                
                For i = 0 To cantLocales
                    cabezaLocales(1, i) = "0"
                Next i
                
                codigo = data.Recordset.Fields("codigo")
                cabezaLocales(1, 1) = data.Recordset.Fields("codigo")
                cabezaLocales(1, 2) = data.Recordset.Fields("descripcion")
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        cadena = ""
        For i = 1 To cantLocales
            cadena = cadena & cabezaLocales(1, i) & vbTab
        Next i
        Impresion.AddItem cadena, True
        
        cadena = "TOTAL GENERAL" & vbTab & vbTab
        For i = 3 To cantLocales
            cadena = cadena & cabezaLocales(2, i) & vbTab
        Next i
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellCenterCenter
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
    End If
    Impresion.AddItem "", True
End Function

Private Sub resumenDepto(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim rubAux As String
    Dim codigo As String
    Dim i As Integer
    Dim cadena As String
    Dim numloc As Integer
    Dim codigoSeccion As String
    Dim codigoDepto As String
    Dim codigoLinea As String
    
    rubAux = rubro
    
    tabla = "SELECT dd.local, dd.codigo, SUM(dd.unidades) AS kilos, mpf.codigolinea, mpf.codigodepto, mpf.codigoseccion  "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " As dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' INNER JOIN " & baseDatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra "
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "GROUP BY local, codigolinea, codigodepto, codigoseccion "
    tabla = tabla & "ORDER BY codigoseccion, codigodepto, codigolinea, local ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        'TITULO
        Impresion.AddItem "RESUMEN DE VENTAS POR KILOS TOTALIZADO POR DEPARTAMENTOS", True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.AddItem "", True
        
        cadena = "CODIGO" & vbTab
        cadena = cadena & "DESCRIPCION" & vbTab
        
        For i = 3 To cantLocales
            cadena = cadena & cabezaLocales(0, i) & vbTab
        Next i
        
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        For i = 0 To cantLocales
            cabezaLocales(1, i) = "0"
            cabezaLocales(2, i) = "0"
            cabezaLocales(3, i) = "0"
        Next i
        
        data.Recordset.MoveFirst
        codigo = data.Recordset.Fields("codigo")
        codigoSeccion = data.Recordset.Fields("codigoseccion")
        codigoDepto = data.Recordset.Fields("codigodepto")
        codigoLinea = data.Recordset.Fields("codigolinea")
        cabezaLocales(1, 1) = codigoLinea
        cabezaLocales(1, 2) = leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux)
        While Not data.Recordset.EOF
            If codigoDepto = data.Recordset.Fields("codigodepto") Then
                If codigoLinea = data.Recordset.Fields("codigolinea") Then
                    numloc = CDbl(data.Recordset.Fields("local")) + 3
                    cabezaLocales(1, numloc) = Format(data.Recordset.Fields("         "), "###,###,##0")
                    cabezaLocales(2, numloc) = Format(CDbl(cabezaLocales(2, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0")
                    cabezaLocales(3, numloc) = Format(CDbl(cabezaLocales(3, numloc)) + CDbl(data.Recordset.Fields("         ")), "###,###,##0")
                Else
                    cadena = ""
                    For i = 1 To cantLocales
                        cadena = cadena & cabezaLocales(1, i) & vbTab
                    Next i
                    Impresion.AddItem cadena, True
                    
                    For i = 0 To cantLocales
                        cabezaLocales(1, i) = "0"
                    Next i
                    
                    codigoLinea = data.Recordset.Fields("codigolinea")
                    cabezaLocales(1, 1) = codigoLinea
                    cabezaLocales(1, 2) = leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux)
                    data.Recordset.MovePrevious
                End If
            Else
                cadena = ""
                For i = 1 To cantLocales
                    cadena = cadena & cabezaLocales(1, i) & vbTab
                Next i
                Impresion.AddItem cadena, True
                
                cabezaLocales(2, 1) = "TOTAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubAux)
                cabezaLocales(2, 2) = ""
                cadena = ""
                For i = 1 To cantLocales
                    cadena = cadena & cabezaLocales(2, i) & vbTab
                Next i
                Impresion.AddItem cadena, True
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                
                
                For i = 0 To cantLocales
                    cabezaLocales(1, i) = "0"
                    cabezaLocales(2, i) = "0"
                Next i
                
                Impresion.AddItem "", True
                
                codigoSeccion = data.Recordset.Fields("codigoseccion")
                codigoDepto = data.Recordset.Fields("codigodepto")
                codigoLinea = data.Recordset.Fields("codigolinea")
                cabezaLocales(1, 1) = codigoLinea
                cabezaLocales(1, 2) = leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux)
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        cadena = ""
        For i = 1 To cantLocales
            cadena = cadena & cabezaLocales(1, i) & vbTab
        Next i
        Impresion.AddItem cadena, True
        
        cabezaLocales(2, 1) = "TOTAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubAux)
        cabezaLocales(2, 2) = ""
        cadena = ""
        For i = 1 To cantLocales
            cadena = cadena & cabezaLocales(2, i) & vbTab
        Next i
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        Impresion.AddItem "", True
        cadena = "TOTAL GENERAL" & vbTab & vbTab
        For i = 3 To cantLocales
            cadena = cadena & cabezaLocales(3, i) & vbTab
        Next i
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 2).Alignment = cellCenterCenter
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
    End If
End Sub

