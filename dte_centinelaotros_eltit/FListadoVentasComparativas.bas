Attribute VB_Name = "FListadoVentasComparativas"
Option Explicit

Public Sub generaInformeLVC(ByRef data1 As Adodc, ByRef data2 As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String)
    Dim i As Long
    impresion.Rows = 2
    impresion.AutoRedraw = False
    
    Call cargaCabeza("LISTADO COMPARATIVO DE VENTAS POR DEPARTAMENTO POR KILOS", codLoc, impresion)
    Call resumenVentasComparativas(data1, data2, impresion, TIPO, codLoc)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Sub resumenVentasComparativas(ByRef data1 As Adodc, ByRef data2 As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String)
    Dim tabla As String
    Dim cadena As String
    Dim rubAux As String
    Dim codigoSeccion As String
    Dim codigoDepto As String
    Dim codigoLinea As String
    Dim prom As Double
    Dim periodo As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim totalkilosMes As Double
    Dim totalkilosAño As Double
    Dim totalNetoMes As Double
    Dim totalNetoAño As Double
    
    fecha1 = Format(DateSerial(Year(fechasistema) - 1, 1, 1), "yyyy-mm-dd")
    fecha2 = Format(DateSerial(Year(fechasistema), 13, 0), "yyyy-mm-dd")
    rubAux = leerRubro(codLoc)
    
    tabla = "SELECT mpf.codigoseccion, mpf.codigodepto, mpf.codigolinea, SUM(dd.unidades) AS kilos, IF(mpf.codigodepto = '00002', ROUND(SUM(dd.total),0), IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(SUM(dd.total / " & Replace((1 + iva / 100), ",", ".") & "),0), ROUND(SUM(dd.total),0))) AS neto, 'TODO' AS periodo "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " AS dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " As dc ON dd.local = dc.local AND dd.tipo = dc.tipo AND dd.numero = dc.numero AND dc.nula = 'N' INNER JOIN " & basedatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra "
    tabla = tabla & "WHERE dd.local = '" & codLoc & "' AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "GROUP BY codigoseccion, codigodepto, codigolinea "
    tabla = tabla & "UNION "
    
    fecha1 = Format(DateSerial(Year(fechasistema), Month(fechasistema), 1), "yyyy-mm-dd")
    fecha2 = Format(DateSerial(Year(fechasistema), Month(fechasistema) + 1, 0), "yyyy-mm-dd")
    
    tabla = tabla & "SELECT mpf.codigoseccion, mpf.codigodepto, mpf.codigolinea, SUM(dd.unidades) AS kilos, IF(mpf.codigodepto = '00002', ROUND(SUM(dd.total),0), IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(SUM(dd.total / " & Replace((1 + iva / 100), ",", ".") & "),0), ROUND(SUM(dd.total),0))) AS neto, 'MES' AS periodo "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " AS dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " As dc ON dd.local = dc.local AND dd.tipo = dc.tipo AND dd.numero = dc.numero AND dc.nula = 'N' INNER JOIN " & basedatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra "
    tabla = tabla & "WHERE dd.local = '" & codLoc & "' AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " "
    tabla = tabla & "GROUP BY codigoseccion, codigodepto, codigolinea "
    tabla = tabla & "ORDER BY codigoseccion, codigodepto, codigolinea, periodo ASC "
    Call ConectarControlData(data1, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data1.Recordset.RecordCount > 0 Then
        Call bordesCabeza1(impresion)
        
        totalkilosMes = 0
        totalkilosAño = 0
        totalNetoMes = 0
        totalNetoAño = 0
        data1.Recordset.MoveFirst
        codigoSeccion = data1.Recordset.Fields("codigoseccion")
        codigoDepto = data1.Recordset.Fields("codigodepto")
        While Not data1.Recordset.EOF
            codigoLinea = data1.Recordset.Fields("codigolinea")
            periodo = data1.Recordset.Fields("periodo")
            If codigoSeccion = data1.Recordset.Fields("codigoseccion") And codigoDepto = data1.Recordset.Fields("codigodepto") Then
                If periodo = "MES" Then
                    prom = Round(CDbl(data1.Recordset.Fields("neto")) / CDbl(data1.Recordset.Fields("         ")) + 0.001, 2)
                    cadena = leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux) & vbTab
                    cadena = cadena & data1.Recordset.Fields("         ") & vbTab
                    cadena = cadena & data1.Recordset.Fields("neto") & vbTab
                    cadena = cadena & prom & vbTab
                    totalkilosMes = totalkilosMes + CDbl(data1.Recordset.Fields("         "))
                    totalNetoMes = totalNetoMes + CDbl(data1.Recordset.Fields("neto"))
                Else
                    If cadena = "" Then
                        cadena = leerNombreLinea(codigoSeccion, codigoDepto, codigoLinea, rubAux) & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab
                    End If
                    prom = Round(CDbl(data1.Recordset.Fields("neto")) / CDbl(data1.Recordset.Fields("         ")) + 0.001, 2)
                    cadena = cadena & data1.Recordset.Fields("         ") & vbTab
                    cadena = cadena & data1.Recordset.Fields("neto") & vbTab
                    cadena = cadena & prom
                    totalkilosAño = totalkilosAño + CDbl(data1.Recordset.Fields("         "))
                    totalNetoAño = totalNetoAño + CDbl(data1.Recordset.Fields("neto"))
                    impresion.AddItem cadena, True
                    cadena = ""
                End If
            Else
                If totalkilosMes = 0 Then
                    prom = 0
                Else
                    prom = Round(totalNetoMes / totalkilosMes + 0.001, 2)
                End If
                cadena = "TOTAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubAux) & vbTab
                cadena = cadena & totalkilosMes & vbTab
                cadena = cadena & totalNetoMes & vbTab
                cadena = cadena & prom & vbTab
                If totalkilosAño = 0 Then
                    prom = 0
                Else
                    prom = Round(totalNetoAño / totalkilosAño + 0.001, 2)
                End If
                cadena = cadena & totalkilosAño & vbTab
                cadena = cadena & totalNetoAño & vbTab
                cadena = cadena & prom
                impresion.AddItem cadena, True
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                
                cadena = ""
                impresion.AddItem "", True
                
                Call comparaAños(data2, impresion, rubAux, codigoSeccion, codigoDepto, codLoc, TIPO)
                Call bordesCabeza1(impresion)
                
                totalkilosMes = 0
                totalkilosAño = 0
                totalNetoMes = 0
                totalNetoAño = 0
                codigoSeccion = data1.Recordset.Fields("codigoseccion")
                codigoDepto = data1.Recordset.Fields("codigodepto")
                data1.Recordset.MovePrevious
            End If
            data1.Recordset.MoveNext
        Wend
        If totalkilosMes = 0 Then
            prom = 0
        Else
            prom = Round(totalNetoMes / totalkilosMes + 0.001, 2)
        End If
        cadena = "TOTAL " & leerNombreDepto(codigoSeccion, codigoDepto, rubAux) & vbTab
        cadena = cadena & totalkilosMes & vbTab
        cadena = cadena & totalNetoMes & vbTab
        cadena = cadena & prom & vbTab
        If totalkilosAño = 0 Then
            prom = 0
        Else
            prom = Round(totalNetoAño / totalkilosAño + 0.001, 2)
        End If
        cadena = cadena & totalkilosAño & vbTab
        cadena = cadena & totalNetoAño & vbTab
        cadena = cadena & prom
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        
        cadena = ""
        impresion.AddItem "", True
        
        Call comparaAños(data2, impresion, rubAux, codigoSeccion, codigoDepto, codLoc, TIPO)
    End If
End Sub

Private Sub comparaAños(ByRef data As Adodc, ByRef impresion As Grid, ByVal rubAux As String, ByVal codigoSeccion As String, ByVal codigoDepto As String, ByVal codLoc As String, ByVal TIPO As String)
    Dim tabla As String
    Dim cadena1 As String
    Dim cadena2 As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim i As Integer
    Dim j As Integer
    Dim prom As Double
    Dim año As String
    Dim mes As Integer
    Dim entro As Boolean
    Dim sumaMesesKilosSegundo As Double
    Dim sumaMesesKilosPrimero As Double
    Dim sumaMesesNetosSegundo As Double
    Dim sumaMesesNetosPrimero As Double
    
    fecha1 = Format(DateSerial(Year(fechasistema) - 1, 1, 1), "yyyy-mm-dd")
    fecha2 = Format(DateSerial(Year(fechasistema), 13, 0), "yyyy-mm-dd")
    
    primerAño = Year(fecha1)
    segundoAño = Year(fecha2)
    
    tabla = "SELECT DATE_FORMAT(dd.fecha, '%Y') AS año, DATE_FORMAT(dd.fecha, '%m') AS mes, SUM(dd.unidades) AS kilos, IF(mpf.codigodepto = '00002', ROUND(SUM(dd.total),0), IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(SUM(dd.total / " & Replace((1 + iva / 100), ",", ".") & "),0), ROUND(SUM(dd.total),0))) AS neto "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " AS dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " As dc ON dd.local = dc.local AND dd.tipo = dc.tipo AND dd.numero = dc.numero AND dc.nula = 'N' INNER JOIN " & basedatos & rubAux & ".r_maestroproductos_fijo_" & rubAux & " AS mpf ON dd.codigo = mpf.codigobarra "
    tabla = tabla & "WHERE dd.local = '" & codLoc & "' AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & "  AND mpf.codigoseccion = '" & codigoSeccion & "' AND mpf.codigodepto = '" & codigoDepto & "' "
    tabla = tabla & "GROUP BY año, mes "
    tabla = tabla & "ORDER BY mes ASC, año DESC "
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        Call bordesCabeza2(impresion)
        
        data.Recordset.MoveFirst
        i = 1
        cadena1 = ""
        cadena2 = ""
        sumaMesesKilosPrimero = 0
        sumaMesesKilosSegundo = 0
        sumaMesesNetosPrimero = 0
        sumaMesesNetosSegundo = 0
        While Not data.Recordset.EOF
            mes = Val(data.Recordset.Fields("mes"))
            If i = mes Then
                If segundoAño = data.Recordset.Fields("año") Then
                    cadena1 = meses(i) & vbTab
                    prom = Round(CDbl(data.Recordset.Fields("neto")) / CDbl(data.Recordset.Fields("         ")) + 0.001, 2)
                    cadena1 = cadena1 & data.Recordset.Fields("         ") & vbTab
                    cadena1 = cadena1 & data.Recordset.Fields("neto") & vbTab
                    cadena1 = cadena1 & prom & vbTab
                    sumaMesesKilosSegundo = sumaMesesKilosSegundo + CDbl(data.Recordset.Fields("         "))
                    sumaMesesNetosSegundo = sumaMesesNetosSegundo + CDbl(data.Recordset.Fields("neto"))
                    entro = True
                    data.Recordset.MoveNext
                Else
                    cadena1 = meses(i) & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab
                End If
            Else
                cadena1 = meses(i) & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab
                If entro = True Then
                    data.Recordset.MovePrevious
                End If
            End If
            If data.Recordset.EOF = False Then
                mes = Val(data.Recordset.Fields("mes"))
                If i = mes Then
                    If primerAño = data.Recordset.Fields("año") Then
                        prom = Round(CDbl(data.Recordset.Fields("neto")) / CDbl(data.Recordset.Fields("         ")) + 0.001, 2)
                        cadena2 = cadena2 & data.Recordset.Fields("         ") & vbTab
                        cadena2 = cadena2 & data.Recordset.Fields("neto") & vbTab
                        cadena2 = cadena2 & prom & vbTab
                        sumaMesesKilosPrimero = sumaMesesKilosSegundo + CDbl(data.Recordset.Fields("         "))
                        sumaMesesNetosPrimero = sumaMesesNetosSegundo + CDbl(data.Recordset.Fields("neto"))
                        entro = True
                    Else
                        cadena2 = "0" & vbTab & "0" & vbTab & "0" & vbTab
                    End If
                Else
                    If entro = True Then
                        data.Recordset.MovePrevious
                    End If
                End If
            Else
                entro = False
            End If
                cadena2 = cadena2 & "0" & vbTab & "0" & vbTab & "0" & vbTab
                impresion.AddItem cadena1 & cadena2, True
                If entro = True Then
                    data.Recordset.MoveNext
                End If
                i = i + 1
                cadena1 = ""
                cadena2 = ""
                entro = False
            'End If
        Wend
        For j = i To 12
            impresion.AddItem meses(j) & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
        Next j
        
        If sumaMesesKilosSegundo = 0 Then
            prom = 0
        Else
            prom = Round(sumaMesesNetosSegundo / sumaMesesKilosSegundo + 0.001, 2)
        End If
        cadena1 = "TOTAL AÑO" & vbTab & sumaMesesKilosSegundo & vbTab & sumaMesesNetosSegundo & vbTab & prom & vbTab
        If sumaMesesKilosPrimero = 0 Then
            prom = 0
        Else
            prom = Round(sumaMesesNetosPrimero / sumaMesesKilosPrimero + 0.001, 2)
        End If
        cadena2 = sumaMesesKilosPrimero & vbTab & sumaMesesNetosPrimero & vbTab & prom
        impresion.AddItem cadena1 & cadena2, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        
        
        impresion.AddItem "", True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellDot
    End If
End Sub


Private Sub bordesCabeza1(ByRef impresion As Grid)
    Dim cadena As String
    'CABEZA
    cadena = "DESCRIPCION" & vbTab
    cadena = cadena & "INFORMACION DEL MES" & vbTab & vbTab & vbTab
    cadena = cadena & "INFORMACION ACUMULADA" & vbTab & vbTab & vbTab
    impresion.AddItem cadena, True
    
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, 4).Merge
    impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, impresion.Cols - 1).Merge
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
    
    cadena = vbTab
    cadena = cadena & "         " & vbTab
    cadena = cadena & "NETO" & vbTab
    cadena = cadena & "PRECIO PROMEDIO"
    cadena = cadena & cadena
    impresion.AddItem cadena, True
    
    impresion.Range(impresion.Rows - 2, 1, impresion.Rows - 1, 1).Merge
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 2, impresion.Cols - 1).Alignment = cellCenterCenter
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
End Sub

Private Sub bordesCabeza2(ByRef impresion As Grid)
    Dim cadena As String
    'CABEZA
    cadena = "MESES" & vbTab
    cadena = cadena & "AÑO " & segundoAño & vbTab & vbTab & vbTab
    cadena = cadena & "AÑO " & primerAño & vbTab & vbTab & vbTab
    impresion.AddItem cadena, True
    
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, 4).Merge
    impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, impresion.Cols - 1).Merge
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
    
    cadena = vbTab
    cadena = cadena & "         " & vbTab
    cadena = cadena & "NETO" & vbTab
    cadena = cadena & "PRECIO PROMEDIO"
    cadena = cadena & cadena
    impresion.AddItem cadena, True
    
    impresion.Range(impresion.Rows - 2, 1, impresion.Rows - 1, 1).Merge
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 2, impresion.Cols - 1).Alignment = cellCenterCenter
    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
End Sub




