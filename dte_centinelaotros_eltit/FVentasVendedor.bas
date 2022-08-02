Attribute VB_Name = "FVentasVendedor"
Option Explicit

Public Sub generaInformeVV(ByRef data As Adodc, ByRef Impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("LIBRO DE MOLIENDA - MES " & Format(fecha1, "mmmm"), empresaActiva, Impresion)
    Call ventasVendedor(data, Impresion, fecha1, fecha2)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub ventasVendedor(ByRef data As Adodc, ByRef Impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim cadena As String
    Dim total As Double
    Dim i As Integer
    Dim vendedor As String
    
    tabla = "SELECT CONCAT(dc.tipo, ' ', dc.numero) AS doc, dc.fecha, dc.rut, dc.sucursal, dc.local, dc.neto, dd.vendedor, dc.tipo "
    tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_detalle_" + empresaActiva + " AS dd ON dc.local = dd.local AND dc.tipo = dd.tipo AND dc.numero = dd.numero AND dc.nula = 'N' "
    tabla = tabla & "WHERE dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dd.vendedor <> '' AND (dc.tipo = 'FV' OR dc.tipo = 'FE' OR dc.tipo = 'BV' OR dc.tipo = 'NV') "
    tabla = tabla & "ORDER BY vendedor, tipo, local, fecha ASC "
    
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        total = 0
        vendedor = data.Recordset.Fields("vendedor")
        
        Impresion.AddItem vendedor & "  " & leerNombreVendedor(vendedor), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        While Not data.Recordset.EOF
            If vendedor = data.Recordset.Fields("vendedor") Then
                cadena = data.Recordset.Fields("doc") & vbTab
                cadena = cadena & data.Recordset.Fields("fecha") & vbTab
                cadena = cadena & data.Recordset.Fields("rut") & vbTab
                cadena = cadena & leerNombreClienteSucursal(data.Recordset.Fields("rut"), data.Recordset.Fields("sucursal")) & vbTab
                cadena = cadena & data.Recordset.Fields("local") & vbTab
                cadena = cadena & Replace(data.Recordset.Fields("neto"), ".", ",")
                
                total = Round(total + CDbl(data.Recordset.Fields("neto")), 2)
                
                Impresion.AddItem cadena, True
            Else
                cadena = vbTab & vbTab & vbTab & "TOTAL VENDEDOR " & vbTab & vbTab
                cadena = cadena & total
                Impresion.AddItem cadena, True
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                Impresion.Range(Impresion.Rows - 1, Impresion.Cols - 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 5).Merge
                Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 5).Alignment = cellCenterCenter
                Impresion.AddItem "", True
                total = 0
                vendedor = data.Recordset.Fields("vendedor")
                Impresion.AddItem vendedor & "  " & leerNombreVendedor(vendedor), True
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        cadena = vbTab & vbTab & vbTab & "TOTAL VENDEDOR " & vbTab & vbTab
        cadena = cadena & total
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, Impresion.Cols - 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 5).Merge
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 5).Alignment = cellCenterCenter
        
        Impresion.AddItem "", True
    End If
End Sub





