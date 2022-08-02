VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form cambiosdeprecio 
   BackColor       =   &H00FF8080&
   Caption         =   "PANTALLA HISTORICO CAMBIOS DE PRECIO"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5985
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5175
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RETORNO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2610
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5175
      Width           =   2850
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4965
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8758
      BackColor       =   16761024
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid1 
         Height          =   4470
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   7885
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "cambiosdeprecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

'GRID1.PrintPreview

Call cabezaInforme("", GRID1, "LISTADO DE CAMBIOS DE PRECIOS", 1)
        GRID1.PageSetup.HeaderMargin = 1
        GRID1.PageSetup.TopMargin = 1
        GRID1.PageSetup.LeftMargin = 1.5
        GRID1.PageSetup.RightMargin = 1
        GRID1.PageSetup.PrintFixedRow = True
        GRID1.PageSetup.BlackAndWhite = True
        GRID1.PageSetup.Orientation = cellLandscape
        GRID1.Range(0, 0, 0, GRID1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Call verificaImpresora(5, GRID1)
End Sub

Private Sub Form_Load()
CARGAGRILLA
Lee_CAMBIOS
End Sub
Sub CARGAGRILLA()
    GRID1.Cols = 9
    GRID1.Column(0).Width = 0
    GRID1.Column(1).Width = 60
    GRID1.Column(2).Width = 60
    GRID1.Column(3).Width = 90
    GRID1.Column(4).Width = 200
    GRID1.Column(5).Width = 80
    GRID1.Column(6).Width = 80
    GRID1.Column(7).Width = 80
    GRID1.Column(7).Width = 100
    
    GRID1.Column(0).Locked = True
    GRID1.Column(1).Locked = True
    GRID1.Column(2).Locked = True
    GRID1.Column(3).Locked = True
    GRID1.Column(4).Locked = True
    GRID1.Column(5).Locked = True
    GRID1.Column(6).Locked = True
    GRID1.Column(7).Locked = True
    GRID1.Cell(0, 1).text = "FECHA"
    GRID1.Cell(0, 2).text = "HORA"
    GRID1.Cell(0, 3).text = "CODIGO"
    GRID1.Cell(0, 4).text = "DESCRIPCION"
    GRID1.Cell(0, 5).text = "TIPOPRECIO"
    GRID1.Cell(0, 6).text = "PRECIO"
    GRID1.Cell(0, 7).text = "ANTERIOR"
    GRID1.Cell(0, 8).text = "USUARIO"
    
    GRID1.Range(0, 1, 0, 8).Alignment = cellCenterGeneral
    GRID1.Range(0, 1, 0, 8).FontSize = 7
    GRID1.Range(0, 1, 0, 8).FontBold = True
    GRID1.Range(0, 1, 0, 8).Borders(cellEdgeBottom) = cellThick
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
GRID1.Rows = 1
End Sub

Sub Lee_CAMBIOS()
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Dim codigo_producto As String
    Dim linea As Double
    
        Set cSql.ActiveConnection = gestionRubro
        cSql.sql = "SELECT r.fecha,r.horamodificacion,r.codigo,m.descripcion,r.codigoprecio,r.precionuevo,r.precioanterior,r.usuario  FROM r_maestroproductos_cambiosdeprecio_" + rubro + " as r "
        cSql.sql = cSql.sql + "INNER JOIN r_maestroproductos_fijo_" & rubro & " AS m ON r.codigo=m.codigobarra "
        cSql.sql = cSql.sql + "WHERE local = '00' and codigo='" + CambioPrecio.dato1.text + "' order by fecha,horamodificacion"
        cSql.Execute
        GRID1.Rows = cSql.RowsAffected + 1
        GRID1.AutoRedraw = False
    linea = 0
        If cSql.RowsAffected > 0 Then
            
            Set resultados = cSql.OpenResultset
           
            While Not resultados.EOF
           linea = linea + 1
                GRID1.Cell(linea, 1).text = resultados(0)
                GRID1.Cell(linea, 2).text = resultados(1)
                GRID1.Cell(linea, 3).text = resultados(2)
                GRID1.Cell(linea, 4).text = resultados(3)
                GRID1.Cell(linea, 5).text = resultados(4)
                GRID1.Cell(linea, 6).text = Format(resultados(5), "$ ##,###,###")
                GRID1.Cell(linea, 7).text = Format(resultados(6), "$ ##,###,###")
                GRID1.Cell(linea, 8).text = resultados(7)
                
               
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        Else
            
        End If
        GRID1.AutoRedraw = True
        GRID1.Refresh
End Sub

