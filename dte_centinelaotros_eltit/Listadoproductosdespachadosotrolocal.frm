VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Listadodespachootrolocal 
   Caption         =   "Mercaderia Por Despachar"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   18441
      BackColor       =   16761024
      Caption         =   "DETALLE"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   675
         Left            =   5040
         TabIndex        =   16
         Top             =   9720
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "LOCAL DESPACHO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox combolocaldespacho 
            Height          =   315
            Left            =   45
            TabIndex        =   17
            Top             =   270
            Width           =   4485
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   675
         Left            =   120
         TabIndex        =   14
         Top             =   9720
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "LOCAL ORIGEN"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox combolocal 
            Height          =   315
            Left            =   45
            TabIndex        =   15
            Top             =   270
            Width           =   4485
         End
      End
      Begin VB.TextBox HASTA1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   11955
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   9960
         Width           =   375
      End
      Begin VB.TextBox HASTA2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12315
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "fecha"
         Top             =   9960
         Width           =   375
      End
      Begin VB.TextBox HASTA3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12675
         MaxLength       =   4
         TabIndex        =   9
         Tag             =   "fecha"
         Top             =   9960
         Width           =   615
      End
      Begin VB.TextBox DESDE1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "fecha"
         Top             =   9960
         Width           =   375
      End
      Begin VB.TextBox DESDE2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10920
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "fecha"
         Top             =   9960
         Width           =   375
      End
      Begin VB.TextBox DESDE3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   11280
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "fecha"
         Top             =   9960
         Width           =   615
      End
      Begin VB.CommandButton cmdgenerar 
         BackColor       =   &H00FF8080&
         Caption         =   "GENERAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13680
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   9840
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   9840
         Width           =   1335
      End
      Begin VB.CommandButton cmdlimpiar 
         BackColor       =   &H00FF8080&
         Caption         =   "LIMPIAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   9840
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   9360
         Width           =   17775
         _ExtentX        =   31353
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   -1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin FlexCell.Grid Grid1 
         Height          =   9015
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   17895
         _ExtentX        =   31565
         _ExtentY        =   15901
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10560
         TabIndex        =   13
         Top             =   9720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HASTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   11955
         TabIndex        =   12
         Top             =   9720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Listadodespachootrolocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fechadesde As String
Dim fechahasta As String

Sub leerMercaderiaDespachada(desde, hasta, localorigen, localdespacho)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim filtro As String
    Dim FILTRO2 As String
    Dim totaluni1 As Double
    Dim totaluni2 As Double
    Dim total1 As Double
    Dim total2 As Double
    
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select numero,rut,ifnull(DATE_FORMAT(fecha,'%d-%m-%Y'),'0'),codigo,descripcion,cantidad,tipodocumento,numerodocumento "
    csql.sql = csql.sql & "from " & baseVentas & localorigen & ".sv_guia_despacho_entrega_" & localorigen & " where localdocumento='" & localdespacho & "' "
    csql.sql = csql.sql & "and fecha between '" & desde & "'and '" & hasta & "' order by rut"
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0
        Grid1.Rows = 1
        Set resultados = csql.OpenResultset
        filtro = resultados(1)
        FILTRO2 = filtro
        Grid1.AutoRedraw = False
        While Not resultados.EOF
            
            If filtro <> FILTRO2 Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 2).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontSize = 8
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
                Grid1.Cell(Grid1.Rows - 1, 1).text = filtro
                Grid1.Cell(Grid1.Rows - 1, 8).text = totaluni1
                Grid1.Cell(Grid1.Rows - 1, 9).text = total1
                totaluni1 = 0
                total1 = 0
                filtro = resultados(1)
                barra.Max = barra.Max + 1
                barra.Value = barra.Value + 1
                barra.Refresh
            End If
 
        
        
            barra.Value = barra.Value + 1
            barra.Refresh
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 1).text = "GD"
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
            Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(1)
            Grid1.Cell(Grid1.Rows - 1, 5).text = leerNombreCliente(resultados(1))
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(3)
            Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(4)
            Grid1.Cell(Grid1.Rows - 1, 8).text = resultados(5)
            Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(6)
            Grid1.Cell(Grid1.Rows - 1, 10).text = resultados(7)
  
            
            totaluni1 = totaluni1 + CDbl(resultados(5))
            totaluni2 = totaluni2 + CDbl(resultados(5))
            total1 = total1 + 1
            total2 = total2 + 1
            
            resultados.MoveNext
            If Not resultados.EOF Then
                 FILTRO2 = resultados(1)
            End If
        Wend
    End If
    csql.Close
   
    Set resultados = Nothing
    Set csql = Nothing
    
    Grid1.AutoRedraw = True
    Grid1.Refresh
      If Grid1.Rows > 1 Then
        ' SUMA FINAL
            Grid1.Rows = Grid1.Rows + 1
            barra.Max = barra.Max + 1
            Grid1.Column(1).Locked = False
            Grid1.Column(2).Locked = False
            Grid1.Column(3).Locked = False
            Grid1.Column(4).Locked = False
            Grid1.Cell(Grid1.Rows - 1, 1).text = filtro
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 2).Merge
            
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontSize = 8
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
 
            Grid1.Cell(Grid1.Rows - 1, 8).text = totaluni1
            totaluni1 = 0
            total1 = 0
            barra.Max = barra.Max + 1
            barra.Value = barra.Value + 1
            barra.Refresh
      End If
    
End Sub

Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer, ByVal impresion As Grid)
        Dim formatogrilla(20, 20) As String
        Dim i As Integer
        'tipo,numero,fecha,rut,codigo,descripcion,cantidad,despachado,vendedor
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TP"
        formatogrilla(1, 2) = "NUMERO"
        formatogrilla(1, 3) = "FECHA"
        formatogrilla(1, 4) = "RUT"
        formatogrilla(1, 5) = "NOMBRE"
        formatogrilla(1, 6) = "CODIGO"
        formatogrilla(1, 7) = "DESCRIPCION"
        formatogrilla(1, 8) = "CANT"
        formatogrilla(1, 9) = "TIPO DOC"
        formatogrilla(1, 10) = "NUM DOC"
 
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "3"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "10"
        formatogrilla(2, 5) = "20"
        formatogrilla(2, 6) = "10"
        formatogrilla(2, 7) = "20"
        formatogrilla(2, 8) = "5"
        formatogrilla(2, 9) = "4"
        formatogrilla(2, 10) = "10"
 
         
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "D"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "S"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = ""
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
        formatogrilla(4, 10) = ""
       
 
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
      
 
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        formatogrilla(6, 10) = ""
     
 
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
        formatogrilla(7, 10) = ""
    
        
 
 
        Rem ANCHO
        formatogrilla(8, 1) = "3"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "7"
        formatogrilla(8, 5) = "30"
        formatogrilla(8, 6) = "9"
        formatogrilla(8, 7) = "30"
        formatogrilla(8, 8) = "5"
        formatogrilla(8, 9) = "4"
        formatogrilla(8, 10) = "7"
    
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)

        impresion.Column(0).Width = 0
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        impresion.DefaultFont.Size = 7
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.SelectionMode = cellSelectionByRow
        impresion.AllowUserSort = True
    End Sub

Private Sub cmdgenerar_Click()
Call leerMercaderiaDespachada(fechadesde, fechahasta, Mid(combolocal.text, 1, 2), Mid(combolocaldespacho.text, 1, 2))
End Sub
Private Sub cmdImprimir_Click()
    If Grid1.Rows > 1 Then
        Call Titulos("LISTADO PRODUCTOS POR DESPACHAR", Grid1)
        Grid1.AutoRedraw = False
        Grid1.Range(1, 1, 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
        Grid1.PageSetup.HeaderMargin = 0.5
        Grid1.PageSetup.TopMargin = 2
        Grid1.PageSetup.LeftMargin = 1
        Grid1.PageSetup.RightMargin = 1
        Grid1.PageSetup.BottomMargin = 2
        Grid1.PageSetup.FooterMargin = 1
        Grid1.PageSetup.BlackAndWhite = True
        Grid1.PageSetup.Orientation = cellLandscape
        Grid1.PageSetup.PrintFixedRow = True
        Grid1.AutoRedraw = True
        Grid1.PrintPreview
    End If
End Sub
Sub Titulos(titulo1 As String, impresion As Grid)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    impresion.FixedRowColStyle = Fixed3D
    impresion.CellBorderColorFixed = vbButtonShadow
    impresion.ShowResizeTips = False
    impresion.ReportTitles.Clear
    impresion.PageSetup.CenterHorizontally = True
    impresion.PageSetup.Orientation = cellPortrait
    impresion.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    impresion.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    impresion.PageSetup.HeaderAlignment = cellLeft
    impresion.PageSetup.HeaderFont.Name = "Verdana"
    impresion.PageSetup.HeaderFont.Size = 8
    impresion.PageSetup.HeaderFont.Italic = True
    
    'TITULOS DEL REPORTE
  
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    impresion.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & " usuario:" + usuarioSistema
    impresion.PageSetup.FooterAlignment = cellRight
    impresion.PageSetup.FooterFont.Name = "Verdana"
    impresion.PageSetup.FooterFont.Size = 7

    
End Sub

Private Sub cmdlimpiar_Click()
Grid1.Rows = 1
End Sub

 

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    DESDE1.text = ceros(DESDE1)
    If DESDE1.text = "00" Then DESDE1.text = Format(fechasistema, "dd")
    DESDE2.SetFocus
End If
End Sub
Private Sub DESDE2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    DESDE2.text = ceros(DESDE2)
    If DESDE2.text = "00" Then DESDE2.text = Format(fechasistema, "mm")
    DESDE3.SetFocus
End If
End Sub
Private Sub DESDE3_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    DESDE3.text = ceros(DESDE3)
    If DESDE3.text = "0000" Then DESDE3.text = Format(fechasistema, "yyyy")
    If IsDate(DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text) = True Then
        fechadesde = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
        HASTA1.SetFocus
    Else
        MsgBox "FECHA INVALIDA", vbCritical, "ATENCION"
        DESDE1.text = ""
        DESDE3.text = ""
        DESDE2.text = ""
        DESDE1.SetFocus
    End If
End If
End Sub
Private Sub HASTA1_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    HASTA1.text = ceros(HASTA1)
    If HASTA1.text = "00" Then HASTA1.text = Format(fechasistema, "dd")
    HASTA2.SetFocus
End If
End Sub
Private Sub hasta2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    HASTA2.text = ceros(HASTA2)
    If HASTA2.text = "00" Then HASTA2.text = Format(fechasistema, "mm")
    HASTA3.SetFocus
End If
End Sub
Private Sub hasta3_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    HASTA3.text = ceros(HASTA3)
    If HASTA3.text = "0000" Then HASTA3.text = Format(fechasistema, "yyyy")
    If IsDate(HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text) = True Then
        fechahasta = HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text
        cmdgenerar.SetFocus
    Else
        MsgBox "FECHA INVALIDA", vbCritical, "ATENCION"
        HASTA1.text = ""
        HASTA3.text = ""
        HASTA2.text = ""
        HASTA1.SetFocus
    End If
End If
End Sub
Private Sub Form_Load()
    Dim palabras As String
    Dim k As Double
    
    Call CargaGrillaInforme(1, 11, Grid1)
    Call LEErlocales
    For k = 1 To combolocal.ListCount
        palabras = combolocal.List(k)
        If empresaActiva = Mid(palabras, 1, 2) Then
            combolocal.text = combolocal.List(k)
            Exit For
        End If
    Next k
End Sub
 
Private Sub opt2_Click()
    Call CargaGrillaInforme(1, 12, Grid1)
    DESDE1.SetFocus
    
End Sub
Private Sub opt3_Click()
    Call CargaGrillaInforme(1, 12, Grid1)
    DESDE1.SetFocus
End Sub
Function leerultimomovimiento(FOLIO) As String
    Dim tabla As String
        tabla = "select glosa from sv_movimientos_garantias_" & empresaActiva & " "
        tabla = tabla & "where folio='" & FOLIO & "' order by folio desc limit 0,1 "
        Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
                leerultimomovimiento = ""
                If data.Recordset.RecordCount > 0 Then
                    data.Recordset.MoveFirst
                    While Not data.Recordset.EOF
                         leerultimomovimiento = data.Recordset.Fields("glosa")
                         data.Recordset.MoveNext
                    Wend
                End If
End Function
Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                combolocal.AddItem (resultados(0) + " " + resultados(1))
                combolocaldespacho.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        combolocaldespacho.text = combolocal.List(0)
        End If
        
End Sub
