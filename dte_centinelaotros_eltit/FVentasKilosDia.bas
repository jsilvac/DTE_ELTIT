Attribute VB_Name = "FVentasKilosDia"
Option Explicit
    
Public Sub generaInformeVKD(ByRef data As Adodc, ByRef Impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Impresion.Rows = 2
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("VENTAS POR KILOS POR DIA", empresaActiva, Impresion)
    Call ventaKilosDia(data, Impresion, fecha1, fecha2)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub ventaKilosDia(ByRef data As Adodc, ByRef Impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim i As Integer
    Dim j As Integer
    Dim fecha As String
    Dim cadena As String
    Dim mostrar As Boolean
    Dim dia As String
    Dim semana As String
    Dim suma(1 To 2, 1 To 6) As Double
    Dim entro As Boolean
    Dim ultimo As Integer
    
    tabla = "SELECT DATE_FORMAT(dd.fecha, '%d') AS dia, @uni1:=@uni1 + SUM(IF(mpf.codigodepto = '00001',(unidades),'0')) AS harinasmol, @uni2:=@uni2 + SUM(IF(mpf.codigodepto = '00002',(unidades),'0')) AS submol, @uni3:=@uni3 + SUM(IF(mpf.codigodepto = '00005',(unidades),'0')) AS trigomol, @uni4:=@uni4 + SUM(IF(mpf.codigodepto = '00101',(unidades),'0')) AS harinasall, @uni5:=@uni5 + SUM(IF(mpf.codigodepto = '00102',(unidades),'0')) AS suball, @uni6:=@uni6 + SUM(IF(mpf.codigodepto = '00105',(unidades),'0')) AS trigoall, @uni1:=0, @uni2:=0, @uni3:=0, @uni4:=0, @uni5:=0, @uni6:=0 "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " AS dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dd.local = dc.local AND dd.tipo = dc.tipo AND dd.numero = dc.numero AND dc.nula = 'N' INNER JOIN " & baseDatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON dd.codigo = mpf.codigobarra "
    tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND (mpf.codigodepto = '00001' OR mpf.codigodepto = '00101' OR mpf.codigodepto = '00002' OR mpf.codigodepto = '00102' OR mpf.codigodepto = '00005' OR mpf.codigodepto = '00105') AND (dc.tipo = 'FV' OR dc.tipo = 'BV' OR dc.tipo = 'ZE') AND mpf.codigobarra <> '0000000000100' AND mpf.codigobarra <> '0000000000101' AND mpf.codigobarra <> '0000000000901'"
    tabla = tabla & "GROUP BY dia "
    tabla = tabla & "ORDER BY dia ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    data.Recordset.Requery
    If data.Recordset.RecordCount > 0 Then
        For j = 1 To 6
            suma(1, j) = 0
            suma(2, j) = 0
        Next j
        
        data.Recordset.MoveFirst
        entro = False
        i = 1
        fecha = fecha1
        semana = calculaSemana(fecha)
        'cadena = "SEMANA N° " & semana & " DEL AÑO"
        'impresion.AddItem cadena, True
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        While Not data.Recordset.EOF
            dia = data.Recordset.Fields("dia")
            fecha = Format(fecha1, "yyyy-mm-") & Format(i, "00")
            'If semana = calculaSemana(fecha) Then
                If i = Val(dia) Then
                    cadena = Format(fecha, "ddd") & " " & dia & vbTab
                    cadena = cadena & data.Recordset.Fields("harinasmol") & vbTab
                    cadena = cadena & data.Recordset.Fields("submol") & vbTab
                    cadena = cadena & data.Recordset.Fields("trigomol") & vbTab
                    cadena = cadena & data.Recordset.Fields("harinasall") & vbTab
                    cadena = cadena & data.Recordset.Fields("suball") & vbTab
                    cadena = cadena & data.Recordset.Fields("trigoall")
                    suma(1, 1) = suma(1, 1) + CDbl(data.Recordset.Fields("harinasmol"))
                    suma(1, 2) = suma(1, 2) + CDbl(data.Recordset.Fields("submol"))
                    suma(1, 3) = suma(1, 3) + CDbl(data.Recordset.Fields("trigomol"))
                    suma(1, 4) = suma(1, 4) + CDbl(data.Recordset.Fields("harinasall"))
                    suma(1, 5) = suma(1, 5) + CDbl(data.Recordset.Fields("suball"))
                    suma(1, 6) = suma(1, 6) + CDbl(data.Recordset.Fields("trigoall"))
                    suma(2, 1) = suma(2, 1) + CDbl(data.Recordset.Fields("harinasmol"))
                    suma(2, 2) = suma(2, 2) + CDbl(data.Recordset.Fields("submol"))
                    suma(2, 3) = suma(2, 3) + CDbl(data.Recordset.Fields("trigomol"))
                    suma(2, 4) = suma(2, 4) + CDbl(data.Recordset.Fields("harinasall"))
                    suma(2, 5) = suma(2, 5) + CDbl(data.Recordset.Fields("suball"))
                    suma(2, 6) = suma(2, 6) + CDbl(data.Recordset.Fields("trigoall"))
                    entro = True
                Else
                    cadena = Format(fecha, "ddd") & " " & Format(i, "00") & vbTab
                    cadena = cadena & "0" & vbTab
                    cadena = cadena & "0" & vbTab
                    cadena = cadena & "0" & vbTab
                    cadena = cadena & "0" & vbTab
                    cadena = cadena & "0" & vbTab
                    cadena = cadena & "0"
                    If entro = True Then
                        data.Recordset.MovePrevious
                    End If
                End If
                Impresion.AddItem cadena, True
                Impresion.Cell(Impresion.Rows - 1, Impresion.Cols - 1).Border(cellEdgeBottom) = cellThin
                i = i + 1
            'Else
                'cadena = "TOT.S." & vbTab
                'cadena = cadena & suma(1, 1) & vbTab
                'cadena = cadena & suma(1, 2) & vbTab
                'cadena = cadena & suma(1, 3) & vbTab
                'cadena = cadena & suma(1, 4) & vbTab
                'cadena = cadena & suma(1, 5) & vbTab
                'cadena = cadena & suma(1, 6)
                'impresion.AddItem cadena, True
                'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                'impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                
                'impresion.AddItem "", True
                'semana = calculaSemana(fecha)
                'cadena = "SEMANA N° " & semana & " DEL AÑO"
                'impresion.AddItem cadena, True
                'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
                'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
                'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                'For j = 1 To 6
                '    suma(1, j) = 0
                'Next j
                'If entro = True Then
                '    data.Recordset.MovePrevious
                'End If
            'End If
            If entro = True Then
                data.Recordset.MoveNext
            End If
        Wend
        ultimo = Val(DateSerial(Year(fecha1), Month(fecha1) + 1, 0))
        For j = i To ultimo
            fecha = Format(fecha1, "yyyy-mm-") & Format(j, "00")
            If semana = calculaSemana(fecha) Then
                cadena = Format(fecha, "ddd") & " " & Format(j, "00") & vbTab
                cadena = cadena & "0" & vbTab
                cadena = cadena & "0" & vbTab
                cadena = cadena & "0" & vbTab
                cadena = cadena & "0" & vbTab
                cadena = cadena & "0" & vbTab
                cadena = cadena & "0"
                Impresion.AddItem cadena, True
            Else
                'cadena = "TOT.S." & vbTab
                'cadena = cadena & suma(1, 1) & vbTab
                'cadena = cadena & suma(1, 2) & vbTab
                'cadena = cadena & suma(1, 3) & vbTab
                'cadena = cadena & suma(1, 4) & vbTab
                'cadena = cadena & suma(1, 5) & vbTab
                'cadena = cadena & suma(1, 6)
                'impresion.AddItem cadena, True
                'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                'impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                'impresion.AddItem "", True
                
                semana = calculaSemana(fecha)
                'If j < ultimo Then
                '    cadena = "SEMANA N° " & semana & " DEL AÑO"
                '    impresion.AddItem cadena, True
                '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
                '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
                '    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                'End If
                For i = 1 To 6
                    suma(1, i) = 0
                Next i
                j = j - 1
            End If
        Next j
        'cadena = "TOT.S." & vbTab
        'cadena = cadena & suma(1, 1) & vbTab
        'cadena = cadena & suma(1, 2) & vbTab
        'cadena = cadena & suma(1, 3) & vbTab
        'cadena = cadena & suma(1, 4) & vbTab
        'cadena = cadena & suma(1, 5) & vbTab
        'cadena = cadena & suma(1, 6)
        'impresion.AddItem cadena, True
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        'impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        'impresion.AddItem "", True
        
        cadena = "TOT.M." & vbTab
        cadena = cadena & suma(2, 1) & vbTab
        cadena = cadena & suma(2, 2) & vbTab
        cadena = cadena & suma(2, 3) & vbTab
        cadena = cadena & suma(2, 4) & vbTab
        cadena = cadena & suma(2, 5) & vbTab
        cadena = cadena & suma(2, 6)
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    End If
    For i = 2 To Impresion.Rows - 1
        Impresion.RowHeight(i) = 26
    Next i
End Sub



