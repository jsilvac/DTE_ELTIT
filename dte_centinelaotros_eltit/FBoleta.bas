Attribute VB_Name = "FBoleta"
Option Explicit
    Private CAMPOS(30, 5) As String

Public Sub imprimeBoletaMatPun(ByVal NUMERO As String, ByRef impresion As Grid, ByRef data As Adodc)
    Dim i As Integer
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim fecha As String
    Dim vencimiento As String
    Dim vendedor As String
    Dim notapedido As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim fono As String
    Dim CODIGO As String
    Dim tiposDePago As String
    Dim cad As String
    Dim tabla As String
    Dim c As Cliente
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventasRubro
    

    csql.sql = "SELECT d.codigo, d.cantidad, d.descripcion, d.precio, d.cantidad * d.precio AS totalpro, c.total, DATE_FORMAT(c.fecha,'%d-%m-%Y') AS fecha, c.descuento, c.rut, c.sucursal "
    csql.sql = csql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS c INNER JOIN sv_documento_detalle_" + empresaActiva + " AS d ON c.local = d.local AND c.tipo = d.tipo AND c.numero = d.numero "
    csql.sql = csql.sql & "WHERE c.local = '" & empresaActiva & "' AND c.tipo = 'BV' AND c.numero = '" & NUMERO & "' "
    csql.sql = csql.sql & "ORDER BY d.linea ASC"
    csql.Execute
 '   Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
    
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
        impresion.Rows = 2
        impresion.Cols = 6
        
        impresion.DefaultFont.Name = "Arial"
        impresion.DefaultFont.Size = 8
        impresion.DefaultFont.Bold = False
        
        impresion.Column(0).Width = 0
        impresion.Column(1).Width = 45
        impresion.Column(2).Width = 100
        impresion.Column(3).Width = 210
        impresion.Column(4).Width = 85
        impresion.Column(5).Width = 120
        
        impresion.Column(1).Alignment = cellRightCenter
        impresion.Column(2).Alignment = cellCenterCenter
        impresion.Column(3).Alignment = cellLeftCenter '/**/
        impresion.Column(4).Alignment = cellRightCenter '/**/
        impresion.Column(5).Alignment = cellRightCenter
        
        impresion.DefaultRowHeight = 13
        
        impresion.PageSetup.PrintGridlines = False
        impresion.AutoRedraw = False
        
        For i = 1 To 23
            impresion.AddItem "", True
        Next i
        
        'data.Recordset.MoveFirst
        
        Call LEERCLIENTE(c, resultados("rut"), resultados("sucursal"), "=")
        
        cad = resultados("fecha")
        fecha = "   " & Format(cad, "dddd") & " "
        fecha = fecha & Format(cad, "dd") & " de "
        fecha = fecha & Format(cad, "mmmm") & " de "
        fecha = fecha & Format(cad, "yyyy")
        nombre = "       " & c.nombre
        rut = "       " & Left(c.rut, Len(c.rut) - 1) & "-" & Right(c.rut, 1)
        direccion = "       " & c.direccion
        ciudad = "       " & c.ciudad
        comuna = c.comuna
        giro = c.giro
        fono = c.fono1
        total = Format(resultados("total"), "$ ###,###,##0") & "               "
        Descuento = Format(CDbl(resultados("descuento")) * -1, "$ ###,###,##0") & "               "
        i = 8
        While Not resultados.EOF
            impresion.Cell(i, 1).text = Right(resultados("codigo"), 4)
            impresion.Cell(i, 2).text = Format(resultados("cantidad"), "###,###,##0.00")
            impresion.Cell(i, 3).text = resultados("descripcion")
            impresion.Cell(i, 4).text = Format(resultados("precio"), "$ ###,###,##0.00")
            impresion.Cell(i, 5).text = Format(resultados("totalpro"), "$ ###,###,##0.00")
            i = i + 1
            resultados.MoveNext
        Wend
        Set csql = Nothing
        csql.Close
        Set resultados = Nothing
        
        impresion.Cell(2, 5).Alignment = cellRightCenter
        impresion.Cell(2, 5).text = NUMERO
        
        'FECHA
        impresion.Range(2, 2, 2, 4).Merge
        impresion.Range(2, 2, 2, 4).Alignment = cellLeftCenter
        impresion.Cell(2, 2).text = fecha
        
        'SEÑORES
        impresion.Range(3, 3, 3, 5).Merge
        impresion.Range(3, 3, 3, 5).Alignment = cellLeftCenter
        impresion.Cell(3, 3).text = nombre
        'RUT
        impresion.Range(4, 2, 4, 3).Merge
        impresion.Range(4, 2, 4, 3).Alignment = cellLeftCenter
        impresion.Cell(4, 2).text = rut
        
        impresion.Cell(22, 5).text = Descuento
        
        impresion.Cell(23, 5).text = total
        
        impresion.AutoRedraw = True
        impresion.Refresh
        
        impresion.PageSetup.LeftMargin = 0.5
        impresion.PageSetup.RightMargin = 0
        impresion.PageSetup.TopMargin = 2
        impresion.PageSetup.BottomMargin = 0
        
        Call verificaImpresora(2, impresion)
        
    Else
        'Call mensaje.mostrarMensaje("ERROR", "NO EXISTE NINGUNA BOLETA CON ESE NUMERO", "NUMERO DE BOLETA = " & numeroFactura)
    End If
End Sub

Public Function leerUltimoFolio(ByVal TIPO As String) As String
    '
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    
    CAMPOS(0, 0) = "IFNULL(MAX(numero) + 1,'0000000001')"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva
    condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "'"
    op = 5
    sql.response = CAMPOS
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        If sql.response(0, 3) <> "" And sql.response(0, 3) <> "0" Then
            leerUltimoFolio = sql.response(0, 3)
        Else
            leerUltimoFolio = "0000000001"
        End If
    End If
End Function

Public Sub MODIFICAFOLIOSII(ByVal TIPO As String, ByVal NUMERO As String, ByVal caja As String, ByVal nuevofolio As String)
    '
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
 nuevofolio = Format(nuevofolio, "0000000000")
    CAMPOS(0, 0) = "foliosii"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 1) = nuevofolio
   
    
    
    CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva
    condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' and numero='" + NUMERO + "' and caja='" & caja & "' "
    op = 3
    sql.response = CAMPOS
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    
End Sub


