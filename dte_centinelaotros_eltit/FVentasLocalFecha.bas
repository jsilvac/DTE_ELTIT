Attribute VB_Name = "FVentasLocalFecha"
Option Explicit

Public Sub generaInformeVLF(ByRef data As Adodc, ByRef impresion As Grid, ByVal codLoc As String, ByVal iniciotxt As String, ByVal finaltxt As String)
    Dim i As Long
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("RESUMEN DE VENTAS (FACTURAS) POR LOCAL - DESDE " & iniciotxt & " HASTA " & finaltxt, codLoc, impresion)
    Call detalleVentas(data, impresion, codLoc, iniciotxt, finaltxt)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function detalleVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal codLoc As String, ByVal iniciotxt As String, ByVal finaltxt As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim totales(1 To 3, 0 To 7) As Double
    Dim i As Integer
    Dim cSql As rdoQuery
    Dim resultado As rdoResultset
    rubAux = leerRubro(codLoc)
    tabla = "SELECT CONCAT(dc.tipo, ' ', dc.numero, '" & vbTab & "', DATE_FORMAT(dc.fecha, '%d-%m-%Y'), '" & vbTab & "', dc.cajera, '" & vbTab & "', IF(dc.nula = 'N', CONCAT(LEFT(mc.rut,9), '-', RIGHT(mc.rut,1), ' ', mc.nombre), 'DOCUMENTO NULO'), '" & vbTab & "', CASE dp.tipopago WHEN '1' THEN 'EFE' WHEN '2' THEN 'CHE' WHEN '3' THEN 'TCB' WHEN '4' THEN 'TDB' WHEN '5' THEN 'CRD' WHEN '6' THEN 'CRT' ELSE 'OTR' END, '" & vbTab & "', IF(dc.nula='N', dc.subtotal, '0'), '" & vbTab & "', IF(dc.nula='N', dc.descuento, '0'), '" & vbTab & "', IF(dc.nula='N', dc.total , '0')) AS item, dp.tipopago, IF(dc.nula='N', dc.subtotal, 0) AS monto, IF(dc.nula='N', dc.descuento, 0) AS descuento, IF(dc.nula='N', dc.total , 0) AS total "
    tabla = tabla & "FROM sv_documento_cabeza_" & codLoc & " as dc INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dc.rut = mc.rut AND mc.sucursal = '0' INNER JOIN sv_documento_pagos_" & codLoc & " AS dp ON dc.local = dp.local AND dc.tipo = dp.tipo AND dc.numero = dp.numero "
    tabla = tabla & "WHERE dc.local = '" & codLoc & "' AND dc.numero>='" & iniciotxt & "' AND dc.numero<='" & finaltxt & "' AND dc.tipo = 'FV' AND dp.monto > 0 "
    tabla = tabla & "ORDER BY  dc.tipo, dc.numero ASC "
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        For i = 0 To 7
            totales(1, i) = 0
            totales(2, i) = 0
            totales(3, i) = 0
        Next i
        While Not data.Recordset.EOF
            impresion.AddItem data.Recordset.Fields("item"), False
            totales(1, 0) = totales(1, 0) + CDbl(data.Recordset.Fields("monto"))
            totales(2, 0) = totales(2, 0) + CDbl(data.Recordset.Fields("descuento"))
            totales(3, 0) = totales(3, 0) + CDbl(data.Recordset.Fields("total"))
            Select Case data.Recordset.Fields("tipopago")
                Case "1"
                    totales(3, 1) = totales(3, 1) + CDbl(data.Recordset.Fields("total"))
                Case "2"
                    totales(3, 2) = totales(3, 2) + CDbl(data.Recordset.Fields("total"))
                Case "3"
                    totales(3, 3) = totales(3, 3) + CDbl(data.Recordset.Fields("total"))
                Case "4"
                    totales(3, 4) = totales(3, 4) + CDbl(data.Recordset.Fields("total"))
                Case "5"
                    totales(3, 5) = totales(3, 5) + CDbl(data.Recordset.Fields("total"))
                Case "6"
                    totales(3, 6) = totales(3, 6) + CDbl(data.Recordset.Fields("total"))
                Case Else
                    totales(3, 7) = totales(3, 7) + CDbl(data.Recordset.Fields("total"))
            End Select
            data.Recordset.MoveNext
        Wend
        impresion.AddItem vbTab & vbTab & vbTab & "TOTALES" & vbTab & vbTab & totales(1, 0) & vbTab & totales(2, 0) & vbTab & totales(3, 0), False
        impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Cell(impresion.Rows - 1, 4).Alignment = cellRightCenter
        impresion.AddItem "", False
        impresion.AddItem "DETALLE FORMAS DE PAGO", False
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).FontBold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Alignment = cellCenterCenter
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas
        cSql.sql = "SELECT codigo, IF(codigo>6, 'OTROS', nombre) AS nombre "
        cSql.sql = cSql.sql & "FROM sv_tiposdepagoclientes "
        cSql.sql = cSql.sql & "ORDER BY codigo ASC"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultado = cSql.OpenResultset
            While Not resultado.EOF
                Select Case resultado("codigo")
                    Case "1"
                        impresion.AddItem resultado("nombre") & vbTab & vbTab & vbTab & Format(totales(3, 1), "$ ###,###,##0"), False
                    Case "2"
                        impresion.AddItem resultado("nombre") & vbTab & vbTab & vbTab & Format(totales(3, 2), "$ ###,###,##0"), False
                    Case "3"
                        impresion.AddItem resultado("nombre") & vbTab & vbTab & vbTab & Format(totales(3, 3), "$ ###,###,##0"), False
                    Case "4"
                        impresion.AddItem resultado("nombre") & vbTab & vbTab & vbTab & Format(totales(3, 4), "$ ###,###,##0"), False
                    Case "5"
                        impresion.AddItem resultado("nombre") & vbTab & vbTab & vbTab & Format(totales(3, 5), "$ ###,###,##0"), False
                    Case "6"
                        impresion.AddItem resultado("nombre") & vbTab & vbTab & vbTab & Format(totales(3, 6), "$ ###,###,##0"), False
                    Case Else
                        impresion.AddItem resultado("nombre") & vbTab & vbTab & vbTab & Format(totales(3, 7), "$ ###,###,##0"), False
                End Select
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Alignment = cellLeftCenter
                impresion.Cell(impresion.Rows - 1, 4).Alignment = cellRightCenter
                resultado.MoveNext
            Wend
            impresion.AddItem "TOTAL" & vbTab & vbTab & vbTab & Format(totales(3, 0), "$ ###,###,##0"), False
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).FontBold = True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Borders(cellEdgeTop) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Alignment = cellLeftCenter
            impresion.Cell(impresion.Rows - 1, 4).Alignment = cellRightCenter
            resultado.Close
            Set resultado = Nothing
        End If
        cSql.Close
        Set cSql = Nothing
    End If
End Function

