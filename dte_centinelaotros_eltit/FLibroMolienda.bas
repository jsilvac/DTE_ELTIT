Attribute VB_Name = "FLibroMolienda"
Option Explicit

Public Sub generaInformeLM(ByRef data As Adodc, ByRef impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("LIBRO DE MOLIENDA - MES " & Format(fecha1, "mmmm"), empresaActiva, impresion)
    Call libroM(data, impresion, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Sub libroM(ByRef data As Adodc, ByRef impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim cadena As String
    Dim total(1 To 3, 1 To 4) As Double
    Dim i As Integer
    Dim dia As String
    
    tabla = "SELECT DATE_FORMAT(fecha, '%d') AS dia, rut, sucursal, trigo, numero, CONCAT(tipodocumento, ' ', numerodocumento) AS doc, harina, ROUND(afrecho + harinilla + impurezas, 2) AS subproductos, valor, tipodocumento "
    tabla = tabla & "FROM sv_guiasmolienda "
    tabla = tabla & "WHERE local = '" & empresaActiva & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
    tabla = tabla & "ORDER BY dia, numero ASC"
    
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        For i = 1 To 4
            total(1, i) = 0
            total(2, i) = 0
            total(3, i) = 0
        Next i
        dia = data.Recordset.Fields("dia")
        While Not data.Recordset.EOF
            If dia = data.Recordset.Fields("dia") Then
                cadena = data.Recordset.Fields("dia") & vbTab
                cadena = cadena & leerNombreClienteSucursal(data.Recordset.Fields("rut"), data.Recordset.Fields("sucursal")) & vbTab
                cadena = cadena & data.Recordset.Fields("rut") & vbTab
                cadena = cadena & Replace(data.Recordset.Fields("trigo"), ".", ",") & vbTab
                cadena = cadena & data.Recordset.Fields("numero") & vbTab
                cadena = cadena & data.Recordset.Fields("doc") & vbTab
                cadena = cadena & data.Recordset.Fields("harina") & vbTab
                cadena = cadena & data.Recordset.Fields("subproductos") & vbTab
                cadena = cadena & data.Recordset.Fields("valor")
                
                total(1, 1) = Round(total(1, 1) + CDbl(data.Recordset.Fields("trigo")), 2)
                total(1, 2) = Round(total(1, 2) + CDbl(data.Recordset.Fields("harina")), 2)
                total(1, 3) = Round(total(1, 3) + CDbl(data.Recordset.Fields("subproductos")), 2)
                total(1, 4) = Round(total(1, 4) + CDbl(data.Recordset.Fields("valor")), 2)
                
                If data.Recordset.Fields("tipodocumento") = "FV" Then
                    total(2, 1) = Round(total(2, 1) + CDbl(data.Recordset.Fields("trigo")), 2)
                    total(2, 2) = Round(total(2, 2) + CDbl(data.Recordset.Fields("harina")), 2)
                    total(2, 3) = Round(total(2, 3) + CDbl(data.Recordset.Fields("subproductos")), 2)
                    total(2, 4) = Round(total(2, 4) + CDbl(data.Recordset.Fields("valor")), 2)
                Else
                    total(3, 1) = Round(total(3, 1) + CDbl(data.Recordset.Fields("trigo")), 2)
                    total(3, 2) = Round(total(3, 2) + CDbl(data.Recordset.Fields("harina")), 2)
                    total(3, 3) = Round(total(3, 3) + CDbl(data.Recordset.Fields("subproductos")), 2)
                    total(3, 4) = Round(total(3, 4) + CDbl(data.Recordset.Fields("valor")), 2)
                End If
                
                impresion.AddItem cadena, True
            Else
                cadena = "TOTAL DIA " & dia & vbTab & vbTab & vbTab
                cadena = cadena & total(1, 1) & vbTab & vbTab & vbTab
                cadena = cadena & total(1, 2) & vbTab
                cadena = cadena & total(1, 3) & vbTab
                cadena = cadena & total(1, 4)
                impresion.AddItem cadena, True
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Alignment = cellCenterCenter
                impresion.AddItem "", True
                For i = 1 To 4
                    total(1, i) = 0
                Next i
                dia = data.Recordset.Fields("dia")
                data.Recordset.MovePrevious
            End If
            data.Recordset.MoveNext
        Wend
        cadena = "TOTAL DIA " & dia & vbTab & vbTab & vbTab
        cadena = cadena & total(1, 1) & vbTab & vbTab & vbTab
        cadena = cadena & total(1, 2) & vbTab
        cadena = cadena & total(1, 3) & vbTab
        cadena = cadena & total(1, 4)
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Alignment = cellCenterCenter
        
        impresion.AddItem "", True
        
        cadena = "TOTAL FACTURAS " & vbTab & vbTab & vbTab
        cadena = cadena & total(2, 1) & vbTab & vbTab & vbTab
        cadena = cadena & total(2, 2) & vbTab
        cadena = cadena & total(2, 3) & vbTab
        cadena = cadena & total(2, 4)
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Alignment = cellCenterCenter
        
        cadena = "TOTAL BOLETAS " & vbTab & vbTab & vbTab
        cadena = cadena & total(3, 1) & vbTab & vbTab & vbTab
        cadena = cadena & total(3, 2) & vbTab
        cadena = cadena & total(3, 3) & vbTab
        cadena = cadena & total(3, 4)
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Alignment = cellCenterCenter
        
        cadena = "TOTAL GENERAL " & vbTab & vbTab & vbTab
        cadena = cadena & total(2, 1) + total(3, 1) & vbTab & vbTab & vbTab
        cadena = cadena & total(2, 2) + total(3, 2) & vbTab
        cadena = cadena & total(2, 3) + total(3, 3) & vbTab
        cadena = cadena & total(2, 4) + total(3, 4)
        impresion.AddItem cadena, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Alignment = cellCenterCenter
    End If
End Sub




