Attribute VB_Name = "FFactura"
Option Explicit
    Private campos(30, 5) As String

'Public Function leerUltimaFactura() As String
'    Dim condicion As String
'    Dim op As Integer
'    Dim sql As New CSqlUtil
'
'    campos(0, 0) = "foliofactura"
'    campos(1, 0) = ""
'    campos(0, 2) = "maestrodecajas"
'    condicion = "codigo = '" & puntoVenta.caja & "' AND local = '" & empresa & "'"
'    op = 5
'    sql.datos = campos
'    Set sql.Conexion = db
'    Call sql.SqlUtil(op, condicion)
'    If sql.estado = 0 Then
'        If sql.datos(0, 3) <> "" Then
'            leerUltimaFactura = sql.datos(0, 3)
'        Else
'            leerUltimaFactura = "0000000001"
'        End If
'    End If
'End Function
'Public Function leerUltimoFolio(ByVal tipo As String) As String
'    Dim condicion As String
'    Dim op As Integer
'    Dim sql As New CSQLUtil
'
'    campos(0, 0) = "MAX(numero) + 1"
'    campos(1, 0) = ""
'    campos(0, 2) = "l_movimientos_cabeza_" & empresaactiva
'    condicion = "tipo = '" & tipo & "'"
'    op = 5
'    sql.datos = campos
'    Set sql.conexion = gestionRubro
'    Call sql.SQLUTIL(op, condicion)
'    If sql.estado = 0 Then
'        If sql.datos(0, 3) <> "" And sql.datos(0, 3) <> "0" Then
'            leerUltimoFolio = sql.datos(0, 3)
'        Else
'            leerUltimoFolio = "0000000001"
'        End If
'    End If
'End Function

'Public Sub modificaFolioFactura(ByVal folio As String)
'    Dim condicion As String
'    Dim op As Integer
'    Dim sql As New CSQLUtil
'
'    campos(0, 0) = "foliofactura"
'    campos(1, 0) = ""
'
'    campos(0, 1) = folio
'    campos(1, 1) = ""
'
'    campos(0, 2) = "sv_maestrodecajas"
'    condicion = "codigo = '" & puntoVenta.caja & "' AND local = '" & empresa & "'"
'    op = 3
'    sql.datos = campos
'    Set sql.conexion = ventas
'    Call sql.SQLUTIL(op, condicion)
'End Sub

Public Sub imprimeFactura(ByVal numeroFactura As String, ByRef Documento As Grid, ByRef rollo As Adodc)
    Dim ss As String
    Dim i As Integer
    Dim k As Integer
    Dim cad As String
    Dim totalprod As String
    Dim descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim lineas As Integer
    Dim fecha As String
    Dim vencimiento As String
    Dim Vendedor As String
    Dim notapedido As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim fono As String
    Dim o As Integer
    Dim dia As String
    Dim mes As String
    Dim ano As String
    Dim nvalor As Long
    Dim codigo As String
    Dim tiposDePago As String
    
    Documento.Rows = 2
    Documento.Cols = 7
    
    Documento.DefaultFont.Name = "Arial"
    Documento.DefaultFont.Size = 8
    Documento.DefaultFont.Bold = False
    
    Documento.Column(0).Width = 0
    Documento.Column(1).Width = 110
    Documento.Column(2).Width = 100
    Documento.Column(3).Width = 210
    Documento.Column(4).Width = 105
    Documento.Column(5).Width = 100
    Documento.Column(6).Width = 130
    'Documento.Column(7).Width = 125
    
    Documento.Column(1).Alignment = cellRightCenter
    Documento.Column(2).Alignment = cellCenterCenter
    Documento.Column(3).Alignment = cellLeftCenter '/**/
    Documento.Column(4).Alignment = cellLeftCenter '/**/
    Documento.Column(5).Alignment = cellRightCenter
    Documento.Column(6).Alignment = cellRightCenter
    'Documento.Column(7).Alignment = cellRightCenter
    
    Documento.DefaultRowHeight = 13
    
    Documento.PageSetup.PrintGridlines = False
    Documento.AutoRedraw = False
    
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.AddItem ""
        
    'CABEZA
    tabla = "SELECT dc.fecha, dc.numero, mc.nombre, dc.rut, mc.direccion, mc.ciudad, mc.giro, mc.comuna, dc.neto, dc.iva, dc.total, (dc.descuento * dc.total / (100 - dc.descuento)) AS descuento "
    tabla = tabla & "FROM sv_documento_cabeza AS dc INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dc.rut = mc.rut "
    tabla = tabla & "WHERE dc.local = '" & empresaactiva & "' AND dc.tipo = 'FV' AND dc.numero = '" & numeroFactura & "'"

    Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
    
    If rollo.Recordset.RecordCount > 0 Then
        rollo.Recordset.MoveFirst
        
        fecha = "                                   "
        fecha = fecha & Format(rollo.Recordset.Fields("fecha"), "dd")
        fecha = fecha & "                         "
        fecha = fecha & Format(rollo.Recordset.Fields("fecha"), "mmmm")
        fecha = fecha & "                                             "
        fecha = fecha & Format(rollo.Recordset.Fields("fecha"), "yyyy")
        nombre = "       " & rollo.Recordset.Fields("nombre")
        rut = rollo.Recordset.Fields("rut")
        direccion = "       " & rollo.Recordset.Fields("direccion")
        ciudad = rollo.Recordset.Fields("ciudad")
        comuna = rollo.Recordset.Fields("comuna")
        giro = "       " & rollo.Recordset.Fields("giro")
        neto = rollo.Recordset.Fields("neto")
        piva = rollo.Recordset.Fields("iva")
        total = rollo.Recordset.Fields("total")
        descuento = rollo.Recordset.Fields("descuento")
        
    End If
    
    'DETALLE
    tabla = "SELECT CONCAT(d.cantidad, '" & vbTab & "', ' ', d.codigo, '" & vbTab & "', d.descripcion, '" & vbTab & vbTab & "', d.precio, '" & vbTab & "', d.total) AS item, d.descuento * d.total / 100 AS descuento, d.total "
    tabla = tabla & "FROM sv_documento_detalle  AS d "
    tabla = tabla & "WHERE d.local = '" & empresaactiva & "' AND d.tipo = 'FV' AND d.numero = '" & numeroFactura & "'"

    Call ConectarControlData(rollo, servidor, baseVentas & rubro, usuario, password, tabla)
        
    If rollo.Recordset.RecordCount > 0 Then
        rollo.Recordset.MoveFirst
        
        While Not rollo.Recordset.EOF
            Documento.AddItem rollo.Recordset.Fields("item"), True
            Documento.Range(Documento.Rows - 1, 3, Documento.Rows - 1, 4).Merge
            rollo.Recordset.MoveNext
        Wend
        Documento.Cell(5, 4).Alignment = cellRightCenter
        Documento.Cell(5, 6).text = numeroFactura
    End If
    
    tiposDePago = leerTiposDePago(numeroFactura, "FV", rollo)
        
    Documento.Range(4, 2, 4, 3).Merge
    Documento.Range(4, 2, 4, 3).Alignment = cellCenterCenter
    'Documento.Cell(4, 2).text = leerNombreEmpresa(empresaactiva)
    
    'Documento.RowHeight(6) = 15
    'Documento.RowHeight(7) = 10
    'Documento.RowHeight(8) = 15
    'Documento.RowHeight(9) = 15
    'Documento.RowHeight(10) = 15
    'Documento.RowHeight(11) = 15
    
    Documento.RowHeight(7) = 20
    
    'FECHA
    Documento.Range(8, 1, 8, 3).Merge
    Documento.Range(8, 1, 8, 3).Alignment = cellLeftCenter
    Documento.Cell(8, 1).text = fecha
    
    'SEÑORES
    Documento.Range(11, 2, 11, 3).Merge
    Documento.Range(11, 2, 11, 3).Alignment = cellLeftCenter
    Documento.Cell(11, 2).text = nombre
    'RUT
    'Documento.Range(11, 2, 11, 3).Merge
    Documento.Cell(11, 6).Alignment = cellLeftCenter
    Documento.Cell(11, 6).text = rut
    'DIRECCION
    Documento.Range(13, 2, 13, 3).Merge
    Documento.Range(13, 2, 13, 3).Alignment = cellLeftCenter
    Documento.Cell(13, 2).text = direccion
    'CIUDAD
    'Documento.Range(13, 2, 13, 3).Merge
    Documento.Cell(13, 6).Alignment = cellLeftCenter
    Documento.Cell(13, 6).text = ciudad
    'GIRO
    Documento.Range(15, 2, 15, 3).Merge
    Documento.Range(15, 2, 15, 3).Alignment = cellLeftCenter
    Documento.Cell(15, 2).text = giro
    'COMUNA
    'Documento.Range(15, 5, 15, 6).Merge
    Documento.Cell(15, 6).Alignment = cellLeftCenter
    Documento.Cell(15, 6).text = comuna
    'CONDICIONES DE PAGO
    'Documento.Range(13, 3, 13, 6).Merge
    'Documento.Range(13, 3, 13, 6).Alignment = cellLeftCenter
    'Documento.Cell(13, 3).text = tiposDePago
    
    
    For i = Documento.Rows To 63
        Documento.AddItem ""
    Next i
    
    If descuento <> 0 Then
        Documento.Cell(63, 5).text = "DESCUENTO"
        Documento.Cell(63, 5).Alignment = cellLeftCenter
        Documento.Cell(63, 6).text = Format(descuento, "$ ###,###,##0")
    End If
    
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.Cell(65, 6).text = Format(neto, "$ ###,###,##0")
    
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.Cell(67, 6).text = Format(piva, "$ ###,###,##0")
    
    Documento.AddItem ""
    Documento.AddItem ""
    Documento.Cell(69, 6).text = Format(total, "$ ###,###,##0")
    
    Documento.AutoRedraw = True
    Documento.Refresh
    
    Documento.PageSetup.LeftMargin = 0.25
    Documento.PageSetup.RightMargin = 0
    Documento.PageSetup.TopMargin = 2.5
    Documento.PageSetup.BottomMargin = 0
    
    For i = 1 To Documento.PageSetup.PaperSizes.Count
        If UCase(Documento.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
            Documento.PageSetup.PaperSize = Documento.PageSetup.PaperSizes.Item(i).Kind
            Exit For
        End If
    Next i
    
    'Documento.DirectPrint
    Documento.PrintPreview
End Sub

'Public Function verificarDocumento(ByVal mostrar As Boolean, ByVal tipo As String, ByVal numero As String) As Boolean
'    Dim condicion As String
'    Dim op As Integer
'    Dim sql As New CSQLUtil
'
'    campos(0, 0) = "numero"
'    campos(1, 0) = ""
'    campos(0, 2) = "datodocumento"
'    condicion = "local = '" & empresa & "' AND tipo = '" & tipo & "' AND numero = '" & numero & "'"
'    op = 5
'    sql.datos = campos
'    Set sql.conexion = db
'    Call sql.SQLUTIL(op, condicion)
'    If sql.estado = 0 Then
'        verificarDocumento = True
'        If mostrar = True Then
'            Call mensaje.mostrarMensaje("ERROR", "EL NUMERO DE BOLETA INGRESADO YA EXISTE.", "BOLETA NUMERO: " & numero)
'        End If
'    Else
'        verificarDocumento = False
'    End If
'End Function

'Public Sub modificarDocumento(ByVal tipo As String, ByVal numeroAntiguo As String, ByVal numeroNuevo As String)
'    Dim condicion As String
'    Dim op As Integer
'    Dim sql As New CSQLUtil
'
'    campos(0, 0) = "numero"
'    campos(1, 0) = ""
'    campos(0, 1) = numeroNuevo
'    campos(1, 1) = ""
'    campos(0, 2) = "datodocumento"
'    condicion = "local = '" & empresa & "' AND numero = '" & numeroAntiguo & "' AND tipo = '" & tipo & "'"
'    op = 3
'    sql.datos = campos
'    Set sql.conexion = db
'    Call sql.SQLUTIL(op, condicion)
'
'    campos(0, 0) = "numero"
'    campos(1, 0) = ""
'    campos(0, 1) = numeroNuevo
'    campos(1, 1) = ""
'    campos(0, 2) = "detalledocumentos"
'    condicion = "local = '" & empresa & "' AND numero = '" & numeroAntiguo & "' AND tipo = '" & tipo & "'"
'    op = 3
'    sql.datos = campos
'    Set sql.conexion = db
'    Call sql.SQLUTIL(op, condicion)
'
'End Sub


Public Function Numero_Texto(nvalor As Long) As String
    
    Dim Mon_Esc, QueES As String
    Dim k As String
    ReDim UNI(15) As String
    ReDim Dec(9) As String
    Dim Z, Num, var As Variant
    Dim c, D, u, v, i As Integer
    Dim textnum As Long
    If Len(nvalor) = 0 Then                        'Si no se ingresa Valor se Devuelve Vacío
        textnum = "": Exit Function
    End If
    If nvalor = 0 Or nvalor > 1E+17 Then
       Mon_Esc = IIf(nvalor = 0, "CERO", "*")
    End If
    ' ------------ UNIDADES ----------------------------------
    UNI(1) = "UN"
    UNI(2) = "DOS"
    UNI(3) = "TRES"
    UNI(4) = "CUATRO"
    UNI(5) = "CINCO"
    UNI(6) = "SEIS"
    UNI(7) = "SIETE"
    UNI(8) = "OCHO"
    UNI(9) = "NUEVE"
    UNI(10) = "DIEZ"
    UNI(11) = "ONCE"
    UNI(12) = "DOCE"
    UNI(13) = "TRECE"
    UNI(14) = "CATORCE"
    UNI(15) = "QUINCE"
    ' ------------ DECENAS ----------------------------------
    Dec(3) = "TREINTA"
    Dec(4) = "CUARENTA"
    Dec(5) = "CINCUENTA"
    Dec(6) = "SESENTA"
    Dec(7) = "SETENTA"
    Dec(8) = "OCHENTA"
    Dec(9) = "NOVENTA"
    
    Num = String$(19 - Len(Str(Trim(nvalor))), Space(1))
    Num = Num + Trim(Str(nvalor))
    i = 1
    Z = ""
    
    Do While True
       k = Mid(Num, 18 - (i * 3 - 1), 3)
    
       If k = Space(3) Then
          Exit Do
       End If
    
       c = Val(Mid(k, 1, 1))
       D = Val(Mid(k, 2, 1))
       u = Val(Mid(k, 3, 1))
       v = Val(Mid(k, 2, 2))
    
       If i > 1 Then
          If (i = 2 Or i = 4) And Val(k) > 0 Then
             Z = " MIL " + Z
          End If
          If i = 3 And Val(Mid(Num, 7, 6)) > 0 Then
             If Val(k) = 1 Then
                Z = " MILLON " + Z
             Else
                Z = " MILLONES " + Z
             End If
          End If
          If i = 5 And Val(k) > 0 Then
             If Val(k) = 1 Then
                Z = " BILLON " + Z
             Else
                Z = " BILLONES " + Z
             End If
          End If
       End If
    
       If v > 0 Then
          Select Case v
                 Case 0 To 15
                      Z = UNI(v) + Z
                 Case 0 To 19
                      Z = " DIECI" + UNI(u) + Z
                 Case 20
                      Z = " VEINTE " + Z
                 Case 0 To 29
                      Z = " VEINTI" + UNI(u) + Z
                 Case Else
                      If u = 0 Then
                         Z = Dec(D) + Z
                      Else
                         Z = Dec(D) + " Y " + UNI(u) + Z
                      End If
          End Select
       End If
    
       If c > 0 Then
          If c = 1 Then
             If v = 0 Then
                Z = " CIEN " + Z
             Else
                Z = " CIENTO " + Z
             End If
          End If
          If c = 2 Or c = 3 Or c = 4 Or c = 6 Or c = 8 Then
             Z = UNI(c) + "CIENTOS " + Z
          End If
          If c = 5 Then
             Z = " QUINIENTOS " + Z
          End If
          If c = 7 Then
             Z = " SETECIENTOS " + Z
          End If
          If c = 9 Then
             Z = " NOVECIENTOS " + Z
          End If
       End If
    
       i = i + 1
    Loop
    
    Mon_Esc = Trim(Z)
    ' CAMBIA "UNO MIL ..." POR "MIL..."
    If Mid(Mon_Esc, 1, 7) = "UN MIL " Then
        Mon_Esc = "MIL " + Trim(Mid(Mon_Esc, 7, Len(Mon_Esc)))
    End If
    Numero_Texto = Mon_Esc + " PESOS" + QueES
End Function






















