VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form vistaDTE 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VISTA ESTANDAR XML CLIENTE"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FlexCell.Grid Grid2 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1720
      BackColorBkg    =   14737632
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IMPRIMIR"
      Height          =   375
      Left            =   11640
      TabIndex        =   1
      Top             =   8280
      Width           =   1455
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12303
      BackColor       =   16761024
      Caption         =   "VISTA DOCUMENTO ELECTRONICO ESTANDAR"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin FlexCell.Grid Grid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   12091
         BackColorBkg    =   16777215
         BackColorScrollBar=   16777215
         BackColorSel    =   16777215
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   16777215
         Rows            =   30
         SelectionMode   =   3
      End
   End
End
Attribute VB_Name = "vistaDTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Grid1.PrintPreview
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)
cargagrila

End Sub




Public Sub BUSCADTE(tipo, rut, numero, empresadte)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim archivoxml As String
    Dim LINEA As Double
    Dim bases As String
            Set csql.ActiveConnection = contadb
        csql.sql = "SELECT numero,fecha,rut,nombre,fecharecepcion,monto,xml,glosadte "
        csql.sql = csql.sql + "FROM " & clientesistema & "fae" & empresadte & ".sv_dte" & empresadte & "_recibidos "
        csql.sql = csql.sql + "WHERE rut like '%" & rut & "%' "
        csql.sql = csql.sql & " and tipo='" & tipo & "' "
        csql.sql = csql.sql & " and numero='" & numero & "' "
 
        csql.Execute
        LINEA = 0
        Grid2.Rows = 1
        Grid2.Refresh
   Grid2.AutoRedraw = False
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            
            While Not resultados.EOF
            Grid2.Rows = Grid2.Rows + 1
            LINEA = LINEA + 1
            Grid2.Cell(LINEA, 1).text = resultados(0)
            Grid2.Cell(LINEA, 2).text = Format(resultados(1), "dd-mm-yyyy")
            Grid2.Cell(LINEA, 3).text = resultados(2)
            Grid2.Cell(LINEA, 4).text = resultados(3)
            Grid2.Cell(LINEA, 5).text = Format(resultados(4), "dd-mm-yyyy")
             Grid2.Cell(LINEA, 6).text = resultados(5)
             If InStr(resultados(7), "Rechazado") > 0 Then
                MsgBox "PROVEEDOR HA ENVIADO DOCUMENTO CON FORMATO NO VALIDO" & vbNewLine & " ERROR" & vbNewLine & resultados(7), vbCritical, "CONTACTAR CON SOPORTE"
                 
             Else
                archivoxml = resultados(6)
            End If
                Call leercontenidodte(archivoxml)
      
 

             '  Grid2.Cell(LINEA, 7).text = leercontenidodte(archivoxml)
            resultados.MoveNext
            Wend
            resultados.Close
        Set resultados = Nothing

        End If
        
         Grid2.AutoRedraw = True
   Grid2.Refresh
   Grid2.Enabled = True
    
        End Sub
 


Sub cargagrila()
   Rem DATOS DE LA COLUMNA
   Dim FORMATOGRILLA(20, 20) As String
    FORMATOGRILLA(1, 1) = "NUMERO"
    FORMATOGRILLA(1, 2) = "FECHA"
    FORMATOGRILLA(1, 3) = "RUT"
    FORMATOGRILLA(1, 4) = "NOMBRE"
    FORMATOGRILLA(1, 5) = "RECEPCION"
    FORMATOGRILLA(1, 6) = "MONTO"
    FORMATOGRILLA(1, 7) = "ORDEN"
    FORMATOGRILLA(1, 8) = "SELECCIONAR"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "9"
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "30"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "0"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = "###,###,###"
    FORMATOGRILLA(4, 7) = ""
    FORMATOGRILLA(4, 8) = "S"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "false"
    FORMATOGRILLA(5, 8) = "FALSE"
    
    Grid2.Cols = 8
    Grid2.Rows = 1

    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
    Grid2.BackColor1 = RGB(231, 235, 247)
    Grid2.BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).Width = CDbl(FORMATOGRILLA(2, k)) * Grid2.DefaultFont.Size
        Grid2.Column(k).MaxLength = CDbl(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid2.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
       
    Next k
'    Grid2.Column(8).CellType = cellCheckBox
    Grid2.Column(0).Width = 0
    Grid2.Range(0, 0, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
    Grid2.Enabled = True
End Sub


Sub leercontenidodte(archivo)
 Dim archivo2 As String
 Dim fae_td As String
 Dim fae_emisor As String
 Dim fae_numero As String
 Dim LINEA(100) As Double
 Dim lin As Double
 Dim detalle(100, 10) As String
 Dim impuesto(500, 50) As String
 fae_td = LeeXML(archivo, "siid:TipoDTE>")
 fae_emisor = LeeXML(archivo, "siid:RUTEmisor>")
 fae_numero = LeeXML(archivo, "siid:Folio>")
 fae_fecha = LeeXML(archivo, "siid:FchEmis>")
 fae_VENCIMIENTO = LeeXML(archivo, "siid:FchVenc>")
 Dim impue As String
 fae_neto = LeeXML(archivo, "siid:MntNeto>")
 fae_IVA = LeeXML(archivo, "siid:IVA>")
 FAE_EXENTO = LeeXML(archivo, "siid:MntExe>")
 Fae_total = LeeXML(archivo, "siid:MntTotal>")
 
  
 fae_emisor_razon = LeeXML(archivo, "siid:RznSoc>")
 fae_emisor_giro = LeeXML(archivo, "siid:GiroEmis>")
 fae_emisor_direccion = LeeXML(archivo, "siid:DirOrigen>")
 fae_emisor_COMUNA = LeeXML(archivo, "siid:CmnaOrigen>")
 fae_emisor_CIUDAD = LeeXML(archivo, "siid:CiudadOrigen>")
 fae_emisor_sucursal = LeeXML(archivo, "siid:Sucursal>")
 
 fae_RECEPTOR_RUT = LeeXML(archivo, "siid:RUTRecep>")
 fae_receptor_razon = LeeXML(archivo, "siid:RznSocRecep>")
 fae_receptor_giro = LeeXML(archivo, "siid:GiroRecep>")
 fae_receptor_direccion = LeeXML(archivo, "siid:DirRecep>")
 fae_RECEPTOR_COMUNA = LeeXML(archivo, "siid:CmnaRecep>")
 fae_RECEPTOR_CIUDAD = LeeXML(archivo, "siid:CiudadRecep>")
 Rem MsgBox LeeXML(ARCHIVO, "siid:TasaImp>")
 For lin = 1 To 100
 If archivo2 <> "" Then
 archivo = Replace(archivo, archivo2, "")
 archivo2 = ""
 End If
    archivo2 = LeeXML(archivo, "siid:NroLinDet>" & lin)
  
    If archivo2 = "" Then
    LINEAS = lin - 1
        Exit For
    Else
          detalle(lin, 0) = LeeXML(archivo2, "siid:VlrCodigo>")
          detalle(lin, 1) = LeeXML(archivo2, "siid:NmbItem>")
          If Len(detalle(lin, 1)) < 3 Then
          detalle(lin, 1) = LeeXML(archivo2, "siid:DscItem>")
          End If
          detalle(lin, 2) = LeeXML(archivo2, "siid:QtyItem>")
          detalle(lin, 3) = LeeXML(archivo2, "siid:UnmdItem>")
          detalle(lin, 4) = LeeXML(archivo2, "siid:PrcItem>")
          detalle(lin, 5) = LeeXML(archivo2, "siid:MontoItem>")
          detalle(lin, 6) = LeeXML(archivo2, "siid:CodImpAdic>")
          If detalle(lin, 6) <> "" Then
            detalle(lin, 7) = leerdatos(contadb, clientesistema & "gestion.g_maestroimpuestos", "porcentaje", "codigofae='" & detalle(lin, 6) & "' ")
            If detalle(lin, 6) = 28 Then
                detalle(lin, 7) = LeeXML(archivo, "siid:TasaImp>")
                detalle(lin, 8) = LeeXML(archivo, "siid:MontoImp>")
                If impuesto(lin, 1) = "" Then impuesto(lin, 1) = 0
                impuesto(lin, 1) = CDbl(impuesto(lin, 1)) + CDbl(detalle(lin, 8))
                impuesto(lin, 2) = "DIESEL"
    
            Else
                impue = detalle(lin, 6)
                detalle(lin, 8) = (detalle(lin, 5) / 100) * detalle(lin, 7)
                If impuesto(impue, 1) = "" Then impuesto(impue, 1) = 0
                impuesto(impue, 1) = CDbl(impuesto(impue, 1)) + CDbl(detalle(lin, 8))
                impuesto(impue, 2) = leerdatos(contadb, clientesistema & "gestion.g_maestroimpuestos", "nombre", "codigofae='" & detalle(lin, 6) & "' ")
           End If
           End If
End If
    

 Next lin
 Grid1.RowHeight(0) = 0
 Grid1.Column(0).Width = 0
 Grid1.ExtendLastCol = True
 
 Grid1.Cols = 1
  Grid1.Cols = 11
  Grid1.Rows = 12
 
 
Grid1.Range(1, 7, 1, 10).Merge
Grid1.Cell(1, 7).text = "RUT : " & fae_emisor
 Grid1.Range(2, 7, 2, 10).Merge
 Select Case fae_td
 Case 33
    Grid1.Cell(2, 7).text = "FACTURA"
 Case 34
    Grid1.Cell(2, 7).text = "FACTURA EXENTA"
Case 61
    Grid1.Cell(2, 7).text = "NOTA DE CREDITO"
Case 52
    Grid1.Cell(2, 7).text = "GUIA DE DESPACHO"
Case 56
    Grid1.Cell(2, 7).text = "NOTA DE DEBITO"
 End Select
 
Grid1.Range(3, 7, 3, 10).Merge
Grid1.Cell(3, 7).text = " Nº " & fae_numero
Grid1.Range(1, 2, 1, 6).Merge
    Grid1.Cell(1, 1).text = "RAZON SOCIAL": Grid1.Cell(1, 2).text = fae_emisor_razon
Grid1.Range(2, 2, 2, 6).Merge
    Grid1.Cell(2, 1).text = "GIRO": Grid1.Cell(2, 2).text = fae_emisor_giro
Grid1.Range(3, 2, 3, 6).Merge
    Grid1.Cell(3, 1).text = "DIRECCION": Grid1.Cell(3, 2).text = fae_emisor_direccion
Grid1.Range(4, 2, 4, 6).Merge
    Grid1.Cell(4, 1).text = "COMUNA": Grid1.Cell(4, 2).text = fae_emisor_COMUNA
Grid1.Range(5, 2, 5, 6).Merge
    Grid1.Cell(5, 1).text = "CIUDAD": Grid1.Cell(5, 2).text = fae_emisor_CIUDAD
 
 Grid1.Range(1, 7, 3, 9).FontBold = True
 Grid1.Range(1, 7, 3, 9).Alignment = cellCenterCenter
 Grid1.Range(1, 7, 3, 9).Borders(cellEdgeRight) = cellThick
 Grid1.Range(1, 7, 3, 9).Borders(cellEdgeLeft) = cellThick
 Grid1.Range(1, 7, 1, 9).Borders(cellEdgeTop) = cellThick
 Grid1.Range(3, 7, 3, 9).Borders(cellEdgeBottom) = cellThick
 
 Grid1.Range(7, 1, 10, Grid1.Cols - 1).BackColor = &HE0E0E0
 Grid1.RowHeight(6) = 10
 
 
 Grid1.Range(7, 1, 10, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
 Grid1.Range(7, 1, 10, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
 Grid1.Range(7, 1, 10, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
 Grid1.Range(7, 1, 10, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
  
    Grid1.Cell(7, 1).text = "FECHA": Grid1.Cell(7, 2).text = fae_fecha
    Grid1.Cell(7, 3).text = "VENCIMIENTO":     Grid1.Cell(7, 4).text = fae_VENCIMIENTO
    Grid1.Cell(7, 7).text = "RUT":      Grid1.Cell(7, 8).text = fae_RECEPTOR_RUT
     Grid1.Range(7, 1, 10, 1).FontBold = True
     Grid1.Range(7, 7, 10, 7).FontBold = True
 Grid1.Range(8, 2, 8, 6).Merge
    Grid1.Cell(8, 1).text = "SEÑORES":     Grid1.Cell(8, 2).text = fae_receptor_razon
    Grid1.Cell(8, 7).text = "COMUNA":     Grid1.Cell(8, 8).text = fae_RECEPTOR_COMUNA
 Grid1.Range(9, 2, 9, 6).Merge
    Grid1.Cell(9, 1).text = "DIRECCION":     Grid1.Cell(9, 2).text = fae_receptor_direccion
      Grid1.Cell(9, 7).text = "CIUDAD":       Grid1.Cell(9, 8).text = fae_RECEPTOR_CIUDAD
 Grid1.Range(10, 2, 10, 6).Merge
    Grid1.Cell(10, 1).text = "GIRO":     Grid1.Cell(10, 2).text = fae_RECEPTOR_COMUNA

  Grid1.RowHeight(11) = 10
 
 'detalle
 Grid1.AddItem "", True
 
 Grid1.Cell(12, 1).text = "CODIGO"
 Grid1.Cell(12, 3).text = "DESCRIPCION"
 Grid1.Cell(12, 7).text = "U/M"
 Grid1.Cell(12, 8).text = "CANTIDAD"
 Grid1.Cell(12, 9).text = "PRECIO"
 Grid1.Cell(12, 10).text = "TOTAL"
  
 Grid1.Column(7).Width = 70
 Grid1.Column(6).Width = 50
 Grid1.Column(5).Width = 50
   Grid1.Range(12, 1, 12, 2).Merge
  Grid1.Range(12, 3, 12, 6).Merge
  Grid1.Range(12, 1, 12, Grid1.Cols - 1).FontBold = True
  Grid1.Range(12, 1, 12, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
  Grid1.Range(12, 1, 12, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
  Grid1.Range(12, 1, 12, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
  
      Grid1.AddItem "", True
 
  
 For k = 1 To LINEAS
  Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 2).Merge
  Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 6).Merge
  Grid1.Range(Grid1.Rows - 1, 7, Grid1.Rows - 1, 10).Alignment = cellRightCenter
  'Grid1.Range(Grid1.rows - 1, 1, Grid1.rows - 1, Grid1.cols - 1).Borders(cellInsideHorizontal) = cellThin
  Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
     Grid1.Cell(Grid1.Rows - 1, 1).text = detalle(k, 0)
     Grid1.Cell(Grid1.Rows - 1, 3).text = detalle(k, 1)
     Grid1.Cell(Grid1.Rows - 1, 7).text = detalle(k, 3)
     Grid1.Cell(Grid1.Rows - 1, 8).text = Format(Replace(detalle(k, 2), ".", ","), "###,###,##0.000")
     Grid1.Cell(Grid1.Rows - 1, 9).text = Format(Replace(detalle(k, 4), ".", ","), "###,###,##0")
     Grid1.Cell(Grid1.Rows - 1, 10).text = Format(Replace(detalle(k, 5), ".", ","), "###,###,##0")
     Grid1.AddItem "", True
 Next k
 Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
 
 Grid1.AddItem "", True
 Grid1.Cell(Grid1.Rows - 1, 10).text = Format(Replace(fae_neto, ".", ","), "###,###,##0")
 Grid1.Cell(Grid1.Rows - 1, 9).text = "NETO"
 Grid1.Range(Grid1.Rows - 1, 7, Grid1.Rows - 1, 10).Alignment = cellRightCenter
 
 
    Grid1.AddItem "", True
    Grid1.Cell(Grid1.Rows - 1, 10).text = Format(Replace(FAE_EXENTO, ".", ","), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 9).text = "EXENTO"
 Grid1.Range(Grid1.Rows - 1, 7, Grid1.Rows - 1, 10).Alignment = cellRightCenter
 
 Grid1.AddItem "", True
 Grid1.Cell(Grid1.Rows - 1, 10).text = Format(Replace(fae_IVA, ".", ","), "###,###,##0")
 Grid1.Cell(Grid1.Rows - 1, 9).text = "IVA"
 Grid1.Range(Grid1.Rows - 1, 7, Grid1.Rows - 1, 10).Alignment = cellRightCenter
 
 For n = 1 To 280
    If impuesto(n, 1) <> "" Then
        Grid1.AddItem "", True
        Grid1.Cell(Grid1.Rows - 1, 7).text = "IMPUESTO " & impuesto(n, 2)
        Grid1.Range(Grid1.Rows - 1, 7, Grid1.Rows - 1, 9).Merge
        Grid1.Cell(Grid1.Rows - 1, 10).text = Format(Replace(impuesto(n, 1), ".", ","), "###,###,##0")
    Grid1.Range(Grid1.Rows - 1, 7, Grid1.Rows - 1, 10).Alignment = cellRightCenter
        
    End If
 Next n
 
 
 Grid1.AddItem "", True
 Grid1.Cell(Grid1.Rows - 1, 10).text = Format(Replace(Fae_total, ".", ","), "###,###,##0")
 Grid1.Cell(Grid1.Rows - 1, 9).text = "TOTAL"
  Grid1.Range(Grid1.Rows - 1, 7, Grid1.Rows - 1, 10).Alignment = cellRightCenter

 
 Grid1.SelectionMode = cellSelectionByRow
 Grid1.AutoRedraw = True
 Grid1.Refresh
 Grid1.PageSetup.LeftMargin = 0.5
 Grid1.PageSetup.RightMargin = 0.5
 Grid1.PageSetup.TopMargin = 1
 
Rem Grid1.PrintPreview
End Sub


Public Function LeeXML(archivo, campo) As String
 Dim cadena As String
 Dim inicio As Double
 Dim strfinal As String
 Dim FINAL As Double
 Dim pasar As Double
 Dim desde As String
 Dim hasta As String
1:
 If InStr(campo, "siid:NroLinDet>") > 0 Then
    strfinal = "</siid:Detalle>"
    pasar = Val(Right(campo, 1)) - 1
 Else
    strfinal = "</" & campo
 End If
 
 If pasar > 0 Then
    FINAL = ((InStr(archivo, strfinal)))
    inicio = InStr(archivo, campo)
    archivo = Mid(archivo, FINAL + Len(strfinal), Len(archivo))
    largo = Len(campo)
     inicio = InStr(archivo, campo)
    FINAL = ((InStr(archivo, strfinal)))
    If inicio > 0 Then inicio = inicio + largo
    
    largo = FINAL - (inicio)
 
 Else
    largo = Len(campo)
    inicio = InStr(archivo, campo)
    If inicio > 0 Then inicio = inicio + largo
    FINAL = ((InStr(archivo, strfinal)))
    largo = FINAL - (inicio)
 End If
 

 If inicio > 0 And largo > 0 Then LeeXML = Mid(archivo, inicio, largo)
 

End Function
