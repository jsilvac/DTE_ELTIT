Attribute VB_Name = "FComisionesVendedor"
Option Explicit
    Public nombreVendedor As String

Public Sub generaInformeCV(ByRef data As Adodc, ByRef data2 As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal vendedor As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("COMISIONES VENDEDOR " & vendedor & " " & nombreVendedor & " - DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, Impresion)
    Call comisionesVendedor(data, data2, Impresion, TIPO, vendedor, fecha1, fecha2)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub comisionesVendedor(ByRef data As Adodc, ByRef data2 As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal vendedor As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim cadena As String
    Dim rut As String
    Dim sucursal As String
    Dim com As Double
    Dim comision As Double
    Dim dias As String
    Dim documento As String
    Dim i As Integer
    Dim loc As Integer
    Dim codLoc As String
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Set cSql.ActiveConnection = ventasRubro
    
    cantLocales = leerCantidadLocales
    ReDim suma(cantLocales + 1, 2) As Double
    
    cSql.sql = "SELECT dc.local, CONCAT(dc.tipo, ' ', dc.numero) AS doc, DATE_FORMAT(dc.vencimiento, '%d-%m-%Y') AS vencimiento, dc.rut, dc.sucursal, DATE_FORMAT(pd.fecha, '%d-%m-%Y') AS pago, pd.numero AS numero, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS dias, dc.total, mv.comision AS com, (mv.comision * dc.neto / 100) AS comision, dc.neto "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_pagos_detalle_" & empresaActiva & " AS pd ON dc.local = pd.local AND dc.tipo = pd.tipo AND dc.numero = pd.documento AND dc.nula = 'N' INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON dc.vendedor = mv.codigo "
    cSql.sql = cSql.sql & "WHERE dc.vendedor = '" & vendedor & "' AND pd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & TIPO & " AND dc.total = dc.abono AND dc.fechapagocomision = '0000-00-00' "
    cSql.sql = cSql.sql & "GROUP BY dc.numero "
    
    cSql.sql = cSql.sql & "UNION "
    
    cSql.sql = cSql.sql & "SELECT dc.local, CONCAT(dc.tipo, ' ', dc.numero) AS doc, DATE_FORMAT(dc.vencimiento, '%d-%m-%Y') AS vencimiento, dc.rut, dc.sucursal, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS pago, dc.numero AS numero, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS dias, dc.total, mv.comision AS com, (mv.comision * dc.neto / 100) AS comision, dc.neto "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_pagos_" + empresaActiva + " AS dp ON dc.local = dp.local AND dc.tipo = dp.tipo AND dc.numero = dp.numero AND dc.nula = 'N' INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON dc.vendedor = mv.codigo "
    cSql.sql = cSql.sql & "WHERE dc.vendedor = '" & vendedor & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND (dp.tipopago = '1' OR dp.tipopago = '2' AND dp.tipopago = '3') AND dp.monto > 0 AND " & TIPO & " AND dc.total = dc.abono AND dc.fechapagocomision = '0000-00-00' "
    cSql.sql = cSql.sql & "GROUP BY dc.numero "
    
    cSql.sql = cSql.sql & "UNION "
    
    cSql.sql = cSql.sql & "SELECT dc.local, CONCAT(dc.tipo, ' ', dc.numero) AS doc, DATE_FORMAT(dc.vencimiento, '%d-%m-%Y') AS vencimiento, dc.rut, dc.sucursal, DATE_FORMAT(dp.fecha, '%d-%m-%Y') AS pago, dc.numero AS numero, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS dias, -1 * dc.total AS total, mv.comision AS com, -1 * (mv.comision * dc.neto / 100) AS comision, dc.neto "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_pagos_" + empresaActiva + " AS dp ON dc.local = dp.local AND dc.tipo = dp.tipo AND dc.numero = dp.numero AND dc.nula = 'N' INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON dc.vendedor = mv.codigo INNER JOIN sv_protesto_" & empresaActiva & " AS pr ON dp.local = pr.local AND dp.numerocheque = pr.cheque "
    cSql.sql = cSql.sql & "WHERE dc.vendedor = '" & vendedor & "' AND pr.fechacheque BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND pr.fechacheque = pr.cancelado AND " & TIPO & " AND dc.total = dc.abono AND dc.fechapagocomision = '0000-00-00' "
    cSql.sql = cSql.sql & "GROUP BY dc.numero "
    
    cSql.sql = cSql.sql & "UNION "
    
    cSql.sql = cSql.sql & "SELECT dc.local, CONCAT(dc.tipo, ' ', dc.numero) AS doc, DATE_FORMAT(dc.vencimiento, '%d-%m-%Y') AS vencimiento, dc.rut, dc.sucursal, DATE_FORMAT(dp.fecha, '%d-%m-%Y') AS pago, dc.numero AS numero, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS dias, dc.total, mv.comision AS com, (mv.comision * dc.neto / 100) AS comision, dc.neto "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_pagos_" + empresaActiva + " AS dp ON dc.local = dp.local AND dc.tipo = dp.tipo AND dc.numero = dp.numero AND dc.nula = 'N' INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON dc.vendedor = mv.codigo INNER JOIN sv_protesto_" & empresaActiva & " AS pr ON dp.local = pr.local AND dp.numerocheque = pr.cheque "
    cSql.sql = cSql.sql & "WHERE dc.vendedor = '" & vendedor & "' AND pr.cancelado BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND pr.fechacheque < pr.cancelado AND " & TIPO & " AND dc.total = dc.abono AND dc.fechapagocomision = '0000-00-00' "
    cSql.sql = cSql.sql & "GROUP BY dc.numero "
    
    cSql.sql = cSql.sql & "UNION "
    
    cSql.sql = cSql.sql & "SELECT dc.local, CONCAT(dc.tipo, ' ', dc.numero) AS doc, DATE_FORMAT(dc.vencimiento, '%d-%m-%Y') AS vencimiento, dc.rut, dc.sucursal, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS pago, dc.numero AS numero, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS dias, dc.total, mv.comision AS com, (mv.comision * dc.neto / 100) AS comision, dc.neto "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_notas AS dn ON dc.local = dn.local AND dc.tipo = 'FV' AND dc.numero = dn.numerofactura AND dc.nula = 'N' INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON dc.vendedor = mv.codigo "
    cSql.sql = cSql.sql & "WHERE dc.vendedor = '" & vendedor & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.neto + dn.monto = dc.abono / " & Replace(1 + (iva + iha) / 100, ",", ".") & " AND dc.fechapagocomision = '0000-00-00' "
    cSql.sql = cSql.sql & "GROUP BY dc.numero "
    
    cSql.sql = cSql.sql & "UNION "
    
    cSql.sql = cSql.sql & "SELECT dc.local, CONCAT(dc.tipo, ' ', dc.numero) AS doc, DATE_FORMAT(dc.vencimiento, '%d-%m-%Y') AS vencimiento, dc.rut, dc.sucursal, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS pago, dc.numero AS numero, DATE_FORMAT(dc.fecha, '%d-%m-%Y') AS dias, dc.total, mv.comision AS com, (mv.comision * dc.neto / 100) AS comision, dc.neto "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON dc.vendedor = mv.codigo "
    cSql.sql = cSql.sql & "WHERE dc.vendedor = '" & vendedor & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.fechapagocomision = '0000-00-00' AND dc.tipo = 'NV' AND dc.nula = 'N'  AND dc.total = dc.abono "
    cSql.sql = cSql.sql & "GROUP BY dc.numero "
    
    cSql.sql = cSql.sql & "UNION "
    
    cSql.sql = cSql.sql & "SELECT p.local, CONCAT(cch.tipodocumento, ' ', cch.numero) AS doc, DATE_FORMAT(p.fechacheque, '%d-%m-%Y') AS vencimiento, p.rut, p.sucursal, DATE_FORMAT(cch.fecharecepcion, '%d-%m-%Y') AS pago, p.cheque AS numero, DATE_FORMAT(p.fechaprotesto, '%d-%m-%Y') AS dias, p.monto, mv.comision AS com, (mv.comision * -1 * p.monto / 100) AS comision, -1 * p.monto AS neto "
    cSql.sql = cSql.sql & "FROM sv_protesto_" & empresaActiva & " AS p INNER JOIN sv_carteracheques AS cch ON p.local = cch.local AND p.cheque = cch.numerocheque INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON p.vendedor = mv.codigo "
    cSql.sql = cSql.sql & "WHERE p.vendedor = '" & vendedor & "' AND p.fechaprotesto BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND  p.cancelado = p.fechacheque "
    cSql.sql = cSql.sql & "GROUP BY p.cheque "
    
    cSql.sql = cSql.sql & "UNION "
    
    cSql.sql = cSql.sql & "SELECT p.local, CONCAT(cch.tipodocumento, ' ', cch.numero) AS doc, DATE_FORMAT(p.fechacheque, '%d-%m-%Y') AS vencimiento, p.rut, p.sucursal, DATE_FORMAT(cch.fecharecepcion, '%d-%m-%Y') AS pago, p.cheque AS numero, DATE_FORMAT(p.fechaprotesto, '%d-%m-%Y') AS dias, p.monto, mv.comision AS com, (mv.comision * p.monto / 100) AS comision, p.monto AS neto "
    cSql.sql = cSql.sql & "FROM sv_protesto_" & empresaActiva & " AS p INNER JOIN sv_carteracheques AS cch ON p.local = cch.local AND p.cheque = cch.numerocheque INNER JOIN " & baseVentas & ".sv_maestrovendedores AS mv ON p.vendedor = mv.codigo "
    cSql.sql = cSql.sql & "WHERE p.vendedor = '" & vendedor & "' AND p.cancelado BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND  p.cancelado <> p.fechacheque "
    cSql.sql = cSql.sql & "GROUP BY p.cheque "
    
    cSql.sql = cSql.sql & "ORDER BY numero ASC "
    cSql.Execute
    
   ' Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
    If cSql.RowsAffected > 0 Then
'        data.Recordset.MoveFirst
          Set resultados = cSql.OpenResultset
        For i = 0 To cantLocales + 1
            suma(i, 0) = 0
            suma(i, 1) = 0
            suma(i, 2) = 0
        Next i
        
        com = 0
        comision = 0
        While Not resultados.EOF
            rut = resultados("rut")
            sucursal = resultados("sucursal")
            loc = Val(resultados("local"))
            'DIAS PAGO COMPARAR FECHA EMISION FACTURA CON VENCIMIENTO CHEQUE
            dias = resultados("dias")
            documento = resultados("doc")
            dias = leerDias(documento, dias, data2)
            com = leerComisionCliente(rut, sucursal)
            If com = 0 Then
                com = CDbl(resultados("com"))
                comision = CDbl(resultados("comision"))
            Else
                comision = CDbl(resultados("neto")) * com / 100
            End If
            cadena = documento & vbTab
            cadena = cadena & resultados("vencimiento") & vbTab
            cadena = cadena & rut & " " & leerNombreClienteSucursal(rut, sucursal) & vbTab
            cadena = cadena & resultados("pago") & vbTab
            cadena = cadena & resultados("numero") & vbTab
            cadena = cadena & dias & vbTab
            cadena = cadena & resultados("total") & vbTab
            cadena = cadena & Format(com, "#0.0") & vbTab
            cadena = cadena & Format(comision, "###,###,##0")
            
            suma(cantLocales + 1, 0) = suma(cantLocales + 1, 0) + CDbl(resultados("total"))
            suma(cantLocales + 1, 1) = CDbl(resultados("com"))
            suma(cantLocales + 1, 2) = suma(cantLocales + 1, 2) + CDbl(resultados("comision"))
            
            suma(loc, 0) = suma(loc, 0) + CDbl(resultados("total"))
            suma(loc, 1) = CDbl(resultados("com"))
            suma(loc, 2) = suma(loc, 2) + CDbl(resultados("comision"))
            
            Impresion.AddItem cadena & vbTab & "1", True
            Impresion.Cell(Impresion.Rows - 1, 0).text = resultados("local")
            If CDbl(resultados("comision")) < 0 Then
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            End If
            resultados.MoveNext
        Wend
        Set cSql = Nothing
        cSql.Close
        Set resultados = Nothing
        
        cadena = vbTab & vbTab & vbTab
        cadena = cadena & "TOTAL VENDEDOR" & vbTab & vbTab & vbTab
        cadena = cadena & suma(cantLocales + 1, 0) & vbTab
        cadena = cadena & suma(cantLocales + 1, 1) & vbTab
        cadena = cadena & Format(suma(cantLocales + 1, 2), "###,###,##0")
        Impresion.AddItem cadena, True
        Impresion.Cell(Impresion.Rows - 1, Impresion.Cols - 1).CellType = cellTextBox
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 6).Merge
        Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 6).Alignment = cellLeftCenter
        Impresion.Range(Impresion.Rows - 1, 7, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.AddItem "", True
        Impresion.Cell(Impresion.Rows - 1, Impresion.Cols - 1).CellType = cellTextBox
        For i = 0 To cantLocales
            codLoc = Format(i, "00")
            If suma(i, 0) <> 0 Then
                cadena = vbTab & vbTab & vbTab
                cadena = cadena & "TOTAL " & leerNombreEmpresa(codLoc) & vbTab & vbTab & vbTab
                cadena = cadena & suma(i, 0) & vbTab
                cadena = cadena & suma(i, 1) & vbTab
                cadena = cadena & Format(suma(i, 2), "###,###,##0")
                Impresion.AddItem cadena, True
                Impresion.Cell(Impresion.Rows - 1, Impresion.Cols - 1).CellType = cellTextBox
                Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 6).Merge
                Impresion.Range(Impresion.Rows - 1, 4, Impresion.Rows - 1, 6).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            End If
        Next i
    End If
End Sub

Private Function leerDias(ByVal documento As String, ByVal fecha As String, ByRef data As Adodc) As String
    Dim tabla As String
    Dim TIPO As String
    Dim NUMERO As String
    Dim dias As String
    Dim numeropago As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Set cSql.ActiveConnection = ventasRubro
    
    TIPO = Left(documento, 2)
    NUMERO = Right(documento, 10)
    
    cSql.sql = "SELECT CASE dp.tipopago WHEN '1' THEN DATE_FORMAT(dp.fecha, '%d-%m-%Y') WHEN '2' THEN DATE_FORMAT(dp.vencimiento, '%d-%m-%Y') WHEN '3' THEN DATE_FORMAT(dp.vencimiento, '%d-%m-%Y') WHEN '4' THEN CASE pc.tipopago WHEN '1' THEN DATE_FORMAT(pc.fecha, '%d-%m-%Y') WHEN '2' THEN DATE_FORMAT(cch.fechavencimiento, '%d-%m-%Y') WHEN '3' THEN DATE_FORMAT(pc.fecha, '%d-%m-%Y') END END AS fecha "
    cSql.sql = cSql.sql & "FROM sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_pagos_detalle_" & empresaActiva & " AS pd ON dp.local = pd.local AND dp.tipo = pd.tipo AND dp.numero = pd.documento INNER JOIN sv_pagos_cabeza_" & empresaActiva & " AS pc ON pd.local = pc.local AND pd.numero = pc.numero LEFT JOIN sv_carteracheques As cch ON pc.local = cch.local AND pc.numero = cch.numero AND 'PA' = cch.tipodocumento "
    cSql.sql = cSql.sql & "WHERE dp.local = '" & empresaActiva & "' AND dp.tipo = '" & TIPO & "' AND dp.numero = '" & NUMERO & "' "
    cSql.sql = cSql.sql & "ORDER BY fecha ASC "
    cSql.Execute
    'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
    
    If cSql.RowsAffected > 0 Then
'        data.Recordset.MoveFirst
          Set resultados = cSql.OpenResultset
        dias = resultados("fecha")
        dias = DateDiff("d", fecha, dias)
    Else
        dias = "0"
    End If
    leerDias = dias
End Function
