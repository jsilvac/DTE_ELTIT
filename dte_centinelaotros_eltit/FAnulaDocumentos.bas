Attribute VB_Name = "FAnulaDocumentos"
Option Explicit

Public Sub listadoDocumentos(ByRef data As Adodc, ByRef impresion As Grid, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("DOCUMENTOS ANULADOS POR LOCAL - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), codLoc, impresion)
    Call detalleVentas(data, impresion, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function detalleVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim totales(1 To 3, 0 To 7) As Double
    Dim i As Integer
    Dim tipo As String
    
    rubAux = leerRubro(codLoc)
    tabla = "SELECT CONCAT(dc.tipo, ' ', dc.numero, '" & vbTab & "', DATE_FORMAT(dc.fecha, '%d-%m-%Y'), '" & vbTab & "', dc.cajera, '" & vbTab & "', LEFT(mc.rut,9), '-', RIGHT(mc.rut,1), ' ', mc.nombre, '" & vbTab & "', CASE dp.tipopago WHEN '1' THEN 'EFE' WHEN '2' THEN 'CHE' WHEN '3' THEN 'TCB' WHEN '4' THEN 'TDB' WHEN '5' THEN 'CRD' WHEN '6' THEN 'CRT' ELSE 'OTR' END, '" & vbTab & "', dc.total, '" & vbTab & "', dc.descuento, '" & vbTab & "', dc.total - dc.descuento) AS item, dp.tipopago, dc.total AS monto, dc.descuento, dc.total - dc.descuento AS total, dc.tipo "
    tabla = tabla & "FROM sv_documento_cabeza_" & codLoc & " as dc INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dc.rut = mc.rut AND mc.sucursal = '0' INNER JOIN sv_documento_pagos_" & codLoc & " AS dp ON dc.local = dp.local AND dc.tipo = dp.tipo AND dc.numero = dp.numero "
    tabla = tabla & "WHERE dc.local = '" & codLoc & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.nula = 'N' "
    tabla = tabla & "ORDER BY dc.tipo, dc.fecha, dc.tipo, dc.numero ASC "
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        For i = 0 To 7
            totales(1, i) = 0
            totales(2, i) = 0
            totales(3, i) = 0
        Next i
        tipo = data.Recordset.Fields("tipo")
        While Not data.Recordset.EOF
            If tipo <> data.Recordset.Fields("tipo") Then
                impresion.AddItem "", False
                impresion.Cell(impresion.Rows - 1, impresion.Cols - 1).CellType = cellTextBox
                tipo = data.Recordset.Fields("tipo")
            Else
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
            End If
        Wend
        impresion.AddItem vbTab & vbTab & vbTab & "TOTALES" & vbTab & vbTab & totales(1, 0) & vbTab & totales(2, 0) & vbTab & totales(3, 0), False
        impresion.Cell(impresion.Rows - 1, impresion.Cols - 1).CellType = cellTextBox
        impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Cell(impresion.Rows - 1, 4).Alignment = cellRightCenter
    End If
End Function


