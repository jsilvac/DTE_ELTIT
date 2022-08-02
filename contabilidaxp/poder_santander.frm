VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form poder_santander 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLANILLA DE TRASPASO DE FONDOS SANTANDER"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   2778
      BackColor       =   16761024
      Caption         =   "DATOS DEL PODER"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar En Fecha"
         Height          =   375
         Left            =   6840
         TabIndex        =   30
         Top             =   960
         Width           =   1455
      End
      Begin Contabilidadxp.BotonMyERP CmdAsientos 
         Height          =   615
         Left            =   11040
         TabIndex        =   29
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         Caption         =   "GENERA ASIENTOS CONTABLES"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox pivote 
         Height          =   375
         Left            =   13200
         TabIndex        =   27
         Top             =   5760
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dato3 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   840
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   503
         _Version        =   393216
         Format          =   160563201
         CurrentDate     =   41673
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "TRASPASO DE FONDOS DEL DIA"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "COMITE DE CREDITOS OFICINA"
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "SRES. CENTRO DE PROCESO REGIONAL"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " REF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9551
      BackColor       =   16761024
      Caption         =   "DETALLE"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid2 
         Height          =   2655
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4683
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   2175
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   3836
         BackColor       =   16761024
         Caption         =   "CARGOS/ABONOS"
         CaptionEstilo3D =   2
         BackColor       =   16761024
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox monto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   23
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox rut2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   9
            TabIndex        =   19
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox rut1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   9
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin XPFrame.FrameXp FrameXp3 
            Height          =   1215
            Left            =   9480
            TabIndex        =   14
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   2143
            BackColor       =   16761024
            Caption         =   "OPCIONES"
            CaptionEstilo3D =   2
            BackColor       =   16761024
            ForeColor       =   8438015
            BordeColor      =   -2147483639
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Contabilidadxp.BotonMyERP CmdImprime 
               Height          =   855
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1508
               Caption         =   "Imprimir"
               PicturePosition =   0
               Picture         =   "poder_santander.frx":0000
               PictureHover    =   "poder_santander.frx":0D2C
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16761024
            End
            Begin Contabilidadxp.BotonMyERP CmdElimina 
               Height          =   855
               Left            =   1200
               TabIndex        =   16
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1508
               Caption         =   "Eliminar"
               PicturePosition =   0
               Picture         =   "poder_santander.frx":1AED
               PictureHover    =   "poder_santander.frx":280A
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16761024
            End
         End
         Begin VB.Label cta2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5520
            TabIndex        =   26
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label cta1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5520
            TabIndex        =   25
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   240
            TabIndex        =   24
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label LblNombreAbono 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3360
            TabIndex        =   22
            Top             =   1200
            Width           =   4815
         End
         Begin VB.Label dv2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3000
            TabIndex        =   21
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " RUT ABONO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   240
            TabIndex        =   20
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblNombreCargo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3360
            TabIndex        =   18
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label dv1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3000
            TabIndex        =   17
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " RUT CARGO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   1095
         Left            =   4800
         TabIndex        =   3
         Top             =   3720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1931
         Cols            =   5
         DefaultFontSize =   8.25
         Enabled         =   0   'False
         Rows            =   30
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Picture         =   "poder_santander.frx":35A3
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3540
   End
End
Attribute VB_Name = "poder_santander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub impresion()
Dim rutalogo As String
If Grid2.Rows - 1 = 2 Then Exit Sub

'Image1.Picture = LoadPicture("")
'Image1.Picture = LoadPicture(App.path & "\stander.jpg")
If ExisteArchivo(rutalogo) = True Then Kill rutalogo
rutalogo = App.path & "\logotemporal.jpg"
    SavePicture Image1.Picture, rutalogo
  
    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.Rows = 1
    Grid1.Cols = 1
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellPortrait
    Grid1.PageSetup.LeftMargin = 1.5
    Grid1.PageSetup.RightMargin = 1.5
    Grid1.PageSetup.TopMargin = 1
      
    Grid1.PageSetup.PrintTitleRows = 0
    
    'Logo
    Grid1.Images.Add rutalogo, "Logo"
    'Set objReportTitle = New FlexCell.ReportTitle
    'objReportTitle.ImageKey = "Logo"
    'objReportTitle.Align = CellLeft
    'Grid1.ReportTitles.Add objReportTitle

    Grid1.Rows = 12
    Grid1.Cols = 7
    Grid1.Range(1, 2, 1, Grid1.Cols - 1).Merge
    Grid1.Range(2, 2, 2, Grid1.Cols - 1).Merge
    Grid1.Range(3, 2, 3, Grid1.Cols - 1).Merge
    Grid1.Range(4, 2, 4, Grid1.Cols - 1).Merge
    
    Grid1.Range(5, 1, 9, Grid1.Cols - 1).Merge
    Grid1.Range(5, 1, 9, Grid1.Cols - 1).WrapText = True
    Grid1.Range(10, 1, 10, 3).Merge
    Grid1.Range(10, 4, 10, 5).Merge
    Grid1.Range(10, 6, 11, 6).Merge
    Grid1.Range(10, 6, 11, 6).WrapText = True
    Grid1.Column(2).Width = 180
    Grid1.Column(4).Width = 180
    
    Grid1.Cell(1, 1).text = "A"
    Grid1.Cell(2, 1).text = "De"
    Grid1.Cell(3, 1).text = "Fecha"
    Grid1.Cell(4, 1).text = "Ref."
    
    Grid1.Cell(1, 2).text = dato1.text
    Grid1.Cell(2, 2).text = dato2.text
    Grid1.Cell(3, 2).text = dato3
    Grid1.Cell(4, 2).text = dato4.text
    
    Grid1.Cell(5, 1).text = "Por el presente instrumento, faculto al Banco Santander Chile para cargar mi(s) cuenta(s) corriente(s) "
    Grid1.Cell(5, 1).text = Grid1.Cell(5, 1).text & "indicada(s) con el objeto de traspasar fondos o cubrir saldos deudor(es) de la(s) siguientes cuenta(s) "
    Grid1.Cell(5, 1).text = Grid1.Cell(5, 1).text & "corriente(s) de terceros bajo mi exclusiva responsabilidad, sin ulterior responsabilidad para el Banco."
    
    Grid1.Cell(10, 1).text = "CUENTA DE CARGO"
    Grid1.Cell(10, 4).text = "CUENTA DE ABONO"
    Grid1.Cell(11, 1).text = "RUT"
    Grid1.Cell(11, 2).text = "NOMBRE"
    Grid1.Cell(11, 3).text = "CTA.CTE."
    
    Grid1.Cell(11, 4).text = "NOMBRE"
    Grid1.Cell(11, 5).text = "CTA.CTE"
    Grid1.Cell(10, 6).text = "MONTO"
    Grid1.Column(6).FormatString = "###,###,###"
    Grid1.Column(6).Alignment = cellRightCenter
    Grid1.Range(10, 1, 11, Grid1.Cols - 1).FontBold = True
    Grid1.Range(10, 1, 11, Grid1.Cols - 1).Alignment = cellCenterCenter
    
    For n = 3 To Grid2.Rows - 1
        Grid1.AddItem "", True
        Grid1.Cell(Grid1.Rows - 1, 1).text = Grid2.Cell(n, 1).text
        Grid1.Cell(Grid1.Rows - 1, 2).text = Grid2.Cell(n, 2).text
        Grid1.Cell(Grid1.Rows - 1, 3).text = Grid2.Cell(n, 3).text
        Grid1.Cell(Grid1.Rows - 1, 4).text = Grid2.Cell(n, 4).text
        Grid1.Cell(Grid1.Rows - 1, 5).text = Grid2.Cell(n, 5).text
        Grid1.Cell(Grid1.Rows - 1, 6).text = Grid2.Cell(n, 6).text
    Next n
    
    Call Grid1.InsertRow(1, 3)
    Grid1.Range(1, 1, 3, 3).Merge
    
    Grid1.Cell(1, 1).SetImage "Logo"
    Grid1.Range(4, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
    Grid1.Range(4, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    Grid1.Range(4, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
    Grid1.Range(4, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
    Grid1.Range(4, 1, 4, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
    
    
    Grid1.Rows = Grid1.Rows + 10
    Grid1.Range(Grid1.Rows - 2, Grid1.Cols - 2, Grid1.Rows - 2, Grid1.Cols - 1).Merge
    Grid1.Range(Grid1.Rows - 2, Grid1.Cols - 2, Grid1.Rows - 2, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid1.Range(Grid1.Rows - 2, Grid1.Cols - 2, Grid1.Rows - 2, Grid1.Cols - 1).Alignment = cellCenterCenter
    Grid1.Cell(Grid1.Rows - 2, Grid1.Cols - 2).text = "FIRMA REPRESENTANTE LEGAL"
    
    Grid1.PageSetup.PrintGridlines = False
    Grid1.PageSetup.BlackAndWhite = True
    Grid1.PrintPreview
  
End Sub




Sub GrabarCabeza(a, de, fecha, ref)
    Dim condicion As String
    Dim op As Integer
    Call EliminaCabeza(fecha)
    campos(0, 0) = "a"
    campos(1, 0) = "de"
    campos(2, 0) = "fecha"
    campos(3, 0) = "ref"
    campos(4, 0) = "usuariocreacion"
    campos(5, 0) = "fechacreacion"
    campos(6, 0) = ""
    
    campos(0, 1) = a
    campos(1, 1) = de
    campos(2, 1) = Format(fecha, "yyyy-mm-dd")
    campos(3, 1) = ref
    campos(4, 1) = USUARIOSISTEMA
    campos(5, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    
    campos(0, 2) = clientesistema & "conta.poder_santander_cabeza"
    condicion = ""
    op = 2
    
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
 
End Sub


Sub GrabarDetalle(fecha, LINEA, rut, NOMBRE, cuenta, monto, c)
    Dim condicion As String
    Dim op As Integer
   
    campos(0, 0) = "fecha"
    campos(1, 0) = "rut"
    campos(2, 0) = "nombre"
    campos(3, 0) = "cuenta"
    campos(4, 0) = "monto"
    campos(5, 0) = "dh"
    campos(6, 0) = "linea"
    campos(7, 0) = ""
    
        
    campos(0, 1) = Format(fecha, "yyyy-mm-dd")
    campos(1, 1) = rut
    campos(2, 1) = NOMBRE
    campos(3, 1) = cuenta
    campos(4, 1) = monto
    campos(5, 1) = c
    campos(6, 1) = LINEA
    
    campos(0, 2) = clientesistema & "conta.poder_santander_detalle"
    condicion = "" ' fecha='" & Format(fecha, "yyyy-mm-dd") & "' "
    op = 2
    
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
 
End Sub

Sub EliminaCabeza(fecha)
'    Dim condicion As String
    Dim op As Integer
    
    campos(0, 2) = clientesistema & "conta.poder_santander_cabeza"
    condicion = "fecha = '" & Format(fecha, "yyyy-mm-dd") & "'"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
End Sub
Sub EliminaDetalle(fecha, Optional LINEA As String)
  '  Dim condicion As String
    Dim op As Integer
    
    campos(0, 2) = clientesistema & "conta.poder_santander_detalle"
    condicion = "fecha = '" & Format(fecha, "yyyy-mm-dd") & "'"
    If LINEA <> "" Then
        condicion = condicion & " and linea ='" & LINEA & "' "
    End If
    
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    
End Sub

Private Sub BotonMyERP2_Click()

End Sub

Private Sub BotonMyERP1_Click()

End Sub



Private Sub CmdElimina_Click()
If Verifica_Permiso(Me.Caption, "elimina") = True Then
    If MsgBox("SE VA A ELIMINAR ESTE REGISTRO" & vbNewLine & " DESEA CONTINUAR?", vbYesNo, "ELIMINANDO  - A T E E N C I O N") = vbYes Then
        Call EliminaCabeza(dato3)
        Call EliminaDetalle(dato3)
        Grid2.Rows = 3
    End If
Else
    MsgBox "NO TIENE PRIVILEGIOS PARA ELIMINAR DATOS EN ESTE MODULO", vbCritical, "ATENCION"

End If
End Sub
 
Private Sub cmdImprime_Click()
Call impresion
End Sub

Private Sub Command1_Click()
 Call LeerPoder(dato3)

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub
 
Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
     
End Sub

Private Sub Form_Load()
Call CARGAGRILLA
dato3 = fechasistema
Call CENTRAR(Me)
End Sub

Private Sub Grid2_DblClick()
Dim row As Integer
row = Grid2.ActiveCell.row
row = Grid2.Cell(row, 0).text
Call LeerLinea(dato3, row)
monto.SetFocus
End Sub

Private Sub monto_GotFocus()
Call cargatexto(monto)
End Sub

Private Sub monto_KeyPress(KeyAscii As Integer)
snum = 0: KeyAscii = esNumero(KeyAscii)
Dim lin As Integer
If KeyAscii = 13 Then
If dv1.Caption = Empty Then rut1.SetFocus: Exit Sub
If dv2.Caption = Empty Then rut2.SetFocus: Exit Sub

If lblNombreCargo.Caption = Empty Then rut1.SetFocus: Exit Sub
If LblNombreAbono.Caption = Empty Then rut2.SetFocus: Exit Sub
If monto = "" Then monto.SetFocus: Exit Sub
If rut1 = rut2 Then rut2.SetFocus: Exit Sub
If monto.Tag <> "" Then
    lin = monto.Tag
    Call EliminaDetalle(dato3, monto.Tag)
    If monto = 0 Then GoTo fin
Else
    lin = ObtenerLinea(dato3)
End If
    Call GrabarCabeza(dato1.text, dato2.text, dato3, dato4.text)
    Call GrabarDetalle(dato3, lin, rut1 & dv1, lblNombreCargo, cta1, monto, "C")
    Call GrabarDetalle(dato3, lin, rut2 & dv2, LblNombreAbono, cta2, monto, "A")
fin:
    Call LeerPoder(dato3)
    monto.Tag = ""
    rut1.text = Empty
    rut2.text = Empty
    monto.text = 0
    rut1.SetFocus
End If
End Sub

Private Sub rut1_Change()
lblNombreCargo.Caption = Empty
dv1.Caption = Empty
cta1.Caption = Empty
End Sub

Private Sub rut1_GotFocus()
Call cargatexto(rut1)
End Sub

Private Sub rut1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    Call ayudaempresa2(rut1)
    Call rut1_KeyPress(13)
End If
End Sub

Private Sub rut1_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(rut1)
        dv1.Caption = rut(rut1)
        lblNombreCargo.Caption = LeerNombreEmpresa2(rut1)
        cta1.Caption = LeerCuentaBancaria(rut1)
        rut2.SetFocus
    End If
End Sub

Private Sub rut2_Change()
LblNombreAbono.Caption = Empty
dv2.Caption = Empty
cta2.Caption = Empty
End Sub

Private Sub rut2_GotFocus()
Call cargatexto(rut2)
End Sub

Private Sub rut2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    Call ayudaempresa2(rut2)
    Call rut2_KeyPress(13)
End If
End Sub

Private Sub rut2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
    Call ceros(rut2)
        dv2.Caption = rut(rut2)
        LblNombreAbono.Caption = LeerNombreEmpresa2(rut2)
        cta2.Caption = LeerCuentaBancaria(rut2)
        monto.SetFocus
    End If
End Sub

Public Function LeerNombreEmpresa2(rutempresa) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.maestroempresas"
    condicion = "rut='" & Val(rutempresa) & "-" & rut(Format(rutempresa, "000000000")) & "' "
    op = 5
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerNombreEmpresa2 = sqlconta.response(0, 3)
    End If
  
End Function


Public Function LeerCodigoEmpresa2(rutempresa) As String
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.maestroempresas"
    condicion = "rut='" & Val(rutempresa) & "-" & rut(Format(rutempresa, "000000000")) & "' "
    op = 5
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerCodigoEmpresa2 = sqlconta.response(0, 3)
    End If
  
End Function


Public Function LeerCuentaBancaria(rutempresa) As String
    campos(0, 0) = "cuentabancaria"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.maestroempresas"
    condicion = "rut='" & Val(rutempresa) & "-" & rut(Format(rutempresa, "000000000")) & "' "
    op = 5
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerCuentaBancaria = sqlconta.response(0, 3)
    End If
  
End Function

Sub ayudaempresa2(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("10s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", pivote, campos, cfijo, largo, 2)
    If pivote <> "" Then caja = Mid(pivote, 1, Len(Replace(pivote, "-", "")) - 1)
    
    caja.Enabled = True
    
    caja.SetFocus
    
End Sub


Sub CARGAGRILLA()
Dim rutalogo As String

    Dim i As Integer
    
    Grid2.Cols = 1
    Grid2.Rows = 3
    Grid2.Cols = 9
    Grid2.RowHeight(0) = 0
    Grid2.Column(0).Width = 0
    Grid2.Range(1, 1, 1, 3).Merge
    Grid2.Range(1, 4, 1, 6).Merge
    Grid2.Range(1, 7, 2, 7).Merge
    Grid2.Cell(1, 1).text = "CUENTA DE CARGO"
    Grid2.Cell(1, 4).text = "CUENTA DE ABONO"
    Grid2.Cell(2, 1).text = "RUT"
    Grid2.Cell(2, 2).text = "NOMBRE"
    Grid2.Cell(2, 3).text = "CTA.CTE."
    
    Grid2.Cell(2, 4).text = "NOMBRE"
    Grid2.Cell(2, 5).text = "CTA.CTE"
    Grid2.Cell(2, 6).text = "MONTO"
    Grid2.Cell(1, 7).text = "CONTAB."
    Grid2.Cell(1, 8).text = "rut"
    Grid2.Column(6).Alignment = cellRightCenter
    
    Grid2.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    Grid2.Column(3).Locked = True
    Grid2.Column(4).Locked = True
    Grid2.Column(5).Locked = True
    Grid2.Column(6).Locked = True
    
    Grid2.Column(1).Width = 80
    Grid2.Column(2).Width = 200
    Grid2.Column(3).Width = 90
    
    Grid2.Column(4).Width = 200
    Grid2.Column(5).Width = 90
    Grid2.Column(6).Width = 80
    Grid2.Column(7).Width = 60
    Grid2.Column(8).Width = 0
    Grid2.Column(6).FormatString = "###,###,##0"
    
    Grid2.Range(1, 1, 2, Grid2.Cols - 1).FontBold = True
    Grid2.Range(1, 1, 2, Grid2.Cols - 1).Alignment = cellCenterCenter
    Grid2.ExtendLastCol = False
    Grid2.PageSetup.PrintGridlines = True
    Grid2.PageSetup.BlackAndWhite = True

End Sub



Function LeerPoder(fecha) As Boolean
    campos(0, 0) = "usuariocreacion"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.poder_santander_cabeza"
    condicion = "fecha = '" & Format(fecha, "yyyy-mm-dd") & "' "
    op = 5
    Grid2.Rows = 3
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    LeerPoder = False
    If sqlconta.status = 0 Then
        LeerPoder = True 'sqlconta.response(0, 3)
    Else
    GoTo no
    End If
    
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Set csql = New rdoQuery
        Set csql.ActiveConnection = conta
        
        csql.sql = "select * from " & clientesistema & "conta.poder_santander_detalle"
        csql.sql = csql.sql & " where fecha = '" & Format(fecha, "yyyy-mm-dd") & "'"
        csql.sql = csql.sql & " order by linea,dh "
        csql.Execute
        
    If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While resultados.EOF = False
                If resultados("dh") = "A" Then
                    Grid2.AddItem "", True
                    Grid2.Cell(Grid2.Rows - 1, 0).text = resultados("linea")
                    Grid2.Cell(Grid2.Rows - 1, 4).text = LeerNombreEmpresa2(Mid(resultados("rut"), 1, 9))
                    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados("cuenta")
                    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados("monto")
                    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados("rut")
                Else
                    Grid2.Cell(Grid2.Rows - 1, 1).text = Format(Mid(resultados("rut"), 1, 9), "###,###,###") & "-" & Right(resultados("rut"), 1)
                    Grid2.Cell(Grid2.Rows - 1, 2).text = LeerNombreEmpresa2(Mid(resultados("rut"), 1, 9))
                    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados("cuenta")
                End If
                Grid2.Cell(Grid2.Rows - 1, 7).CellType = cellCheckBox
            If resultados("dh") <> "A" Then
            RutLoc = Mid(Replace(Grid2.Cell(Grid2.Rows - 1, 1).text, ".", ""), 1, 8)
            EMPRESAORIGEN = LeerCodigoEmpresa2(RutLoc)
                
                Grid2.Cell(Grid2.Rows - 1, 7).text = contabilizado(resultados("monto"), "TF", Grid2.Cell(Grid2.Rows - 1, 4).text, "D", "001", fecha, EMPRESAORIGEN)
                Grid2.Cell(Grid2.Rows - 1, 7).Locked = True
                End If
                resultados.MoveNext
            Wend
        End If
        Set resultados = Nothing
        csql.Close
        rut1.SetFocus
no:
End Function



Function ObtenerLinea(fecha) As String
    campos(0, 0) = "ifnull(max(linea),0)+1"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.poder_santander_detalle"
    condicion = "fecha='" & Format(fecha, "yyyy-mm-dd") & "' "
    op = 5
    ObtenerLinea = 1
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    ObtenerLinea = sqlconta.response(0, 3)
    End If

End Function



Sub LeerLinea(fecha, LINEA)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Set csql = New rdoQuery
        Set csql.ActiveConnection = conta
        
        csql.sql = "select * from " & clientesistema & "conta.poder_santander_detalle"
        csql.sql = csql.sql & " where fecha = '" & Format(fecha, "yyyy-mm-dd") & "' and linea ='" & LINEA & "' "
        csql.sql = csql.sql & " order by linea,dh "
        csql.Execute
        
    If csql.RowsAffected > 0 Then
            monto.Tag = LINEA
            Set resultados = csql.OpenResultset
            While resultados.EOF = False
                If resultados("dh") = "A" Then
                    rut2 = Mid(resultados("rut"), 1, 9)
                    dv2.Caption = Right(resultados("rut"), 1)
                    LblNombreAbono = resultados("nombre")
                    monto = resultados("monto")
                    cta2 = resultados("cuenta")
                Else
                    rut1 = Mid(resultados("rut"), 1, 9)
                    dv1.Caption = Right(resultados("rut"), 1)
                    lblNombreCargo = resultados("nombre")
                    cta1 = resultados("cuenta")
                End If
                resultados.MoveNext
            Wend
        End If
        Set resultados = Nothing
        csql.Close
End Sub
 
 Private Sub CmdAsientos_Click()
Call generacomprobantes
Call LeerPoder(dato3)

End Sub
 
Sub generacomprobantes()
          Dim k As Double
       Dim tipo As String
        Dim lin As String
        Dim Num As String
    Dim CuentaD As String
    Dim CuentaH As String
Dim codigocontable As String
Dim rutcontable As String
     Dim glosa As String
     Dim FOLIO As String
        Dim DH As String
     Dim fecha As String
        Dim td As String
     Dim monto As Double
    Dim CONTAORIGEN As String
    Dim CONTADESTINO As String
    
    For k = 3 To Grid2.Rows - 1
        If Grid2.Cell(k, 7).BooleanValue = False Then
            
          Rem ORIGEN
            RutLoc = Grid2.Cell(k, 8).text
            RutLoc = Left(RutLoc, Len(RutLoc) - 1)
            codigocontable = LeerCodigoEmpresa2(RutLoc)
            CONTADESTINO = LeerCodigoEmpresa2(RutLoc)
            RutLoc = Mid(Replace(Grid2.Cell(k, 1).text, ".", ""), 1, 8)
            CONTAORIGEN = LeerCodigoEmpresa2(RutLoc)
            
                    td = "TF"
                
             FOLIO = ultimoFolio(td, CONTAORIGEN)
             monto = Grid2.Cell(k, 6).text
             LINEA = "0001"
            If CONTADESTINO = "01" Then
            cuenta = "23100001"
            Else
            cuenta = "1160" & Format(CONTADESTINO, "0000")
            End If
             
             glosa = "TRANF. DE FONDOS A " & Grid2.Cell(k, 4).text
                DH = "D"
             fecha = Format(dato3, "yyyy-mm-dd")
                
Call grabarcomprobante_lineas(td, FOLIO, LINEA, fecha, cuenta, "", "", "", glosa, td, FOLIO, fecha, fecha, monto, DH, Format(fecha, "mm"), Format(fecha, "yyyy"), "", CONTAORIGEN)
             LINEA = "0002"
             cuenta = "11500100"
             glosa = "TRANF. DE FONDOS A " & Grid2.Cell(k, 4).text
             DH = "H"
          Call grabarcomprobante_lineas(td, FOLIO, LINEA, fecha, cuenta, "", "", "", glosa, td, FOLIO, fecha, fecha, monto, DH, Format(fecha, "mm"), Format(fecha, "yyyy"), "", CONTAORIGEN)

        Rem DESTINO

            td = "RF"
            FOLIO = ultimoFolio(td, CONTADESTINO)
            monto = Grid2.Cell(k, 6).text
            LINEA = "0001"
            cuenta = "11500100"
            glosa = "TRANF. DE FONDOS " & Grid2.Cell(k, 2).text
            DH = "D"
            fecha = Format(dato3, "yyyy-mm-dd")

       Call grabarcomprobante_lineas(td, FOLIO, LINEA, fecha, cuenta, "", "", "", glosa, td, FOLIO, fecha, fecha, monto, DH, Format(fecha, "mm"), Format(fecha, "yyyy"), "", CONTADESTINO)
            LINEA = "0002"
            cuenta = "1160" & Format(CONTAORIGEN, "0000")
            
            If CONTAORIGEN = "01" Then
            cuenta = "23100001"
            End If
            
            If CONTAORIGEN = "01" And CONTADESTINO = "28" Then
            cuenta = "23100002"
            End If
            
            
            glosa = "TRANF. DE FONDOS " & Grid2.Cell(k, 2).text
            DH = "H"
Call grabarcomprobante_lineas(td, FOLIO, LINEA, fecha, cuenta, "", "", "", glosa, td, FOLIO, fecha, fecha, monto, DH, Format(fecha, "mm"), Format(fecha, "yyyy"), "", CONTADESTINO)
'
'
'
'
         
    
        End If
    Next k
End Sub
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, _
                            glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, _
                            monto, DH, MES, año, rutproveedor, codigocontable)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(19, 1) = Time
    campos(20, 1) = rutproveedor

    campos(0, 2) = clientesistema & "conta" & codigocontable & ".movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub


Sub MarcarLinea(fecha, LINEA, rut)
    campos(0, 0) = "contabilizado"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 1) = 1
    campos(1, 1) = ""
    
    campos(0, 2) = clientesistema & "conta.poder_santander_detalle"
    condicion = "fecha='" & fecha & "' and linea ='" & LINEA & "' "
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
End Sub



Public Function ExisteCta(codigo, empresa) As Boolean
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.cuentasdelmayor"
    condicion = "año='" & Format(fechasistema, "yyyy-mm-dd") & "' and codigo='" & codigo & "' "
    op = 5
    ExisteCta = False
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    ExisteCta = True ' sqlconta.response(0, 3)
    End If
  
End Function


Function ultimoFolio(tipo, empresa) As String
    campos(0, 0) = "LPAD(IFNULL(MAX(numero),0)+1,10,0) as num"
    campos(1, 0) = ""
    campos(2, 0) = ""
    campos(0, 2) = clientesistema & "conta" & empresa & ".movimientoscontables"
    condicion = "tipo='" & tipo & "' and año='" + Format(fechasistema, "yyyy") + "' and mes='" + Format(fechasistema, "mm") + "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        ultimoFolio = sqlconta.response(0, 3)
    End If
    End Function
Function contabilizado(monto, tipo, glosa, DH, LINEA, fecha, empresa) As String

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Set csql = New rdoQuery
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT * FROM eltit_conta" + empresa + ".movimientoscontables "
        csql.sql = csql.sql + " where monto='" & monto & "' and tipo='" + tipo + "' and glosacontable like '%" + Mid(glosa, 1, 40) + "%' and dh='" + DH + "' and linea ='" + LINEA + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "'"

        csql.Execute
        contabilizado = "0"
        
    If csql.RowsAffected > 0 Then
        contabilizado = "1"
        
        End If
        Set resultados = Nothing
        csql.Close
End Function

