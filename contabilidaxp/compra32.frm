VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form compra032 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Libros de Compras Aceptados en el SII"
   ClientHeight    =   10410
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   19110
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   694
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1274
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   16080
      TabIndex        =   9
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   19080
      TabIndex        =   0
      Top             =   10410
      Width           =   19110
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   9690
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   17092
      BackColor       =   16761024
      Caption         =   "LISTADO"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command9 
         Caption         =   "MARCAR NOTAS DE CREDITOS"
         Height          =   255
         Left            =   11760
         TabIndex        =   30
         Top             =   9240
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   "CONTABILIZAR NO RECIBIDAS"
         Height          =   255
         Left            =   8640
         TabIndex        =   29
         Top             =   9240
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Solo Inconsistencias"
         Height          =   255
         Left            =   12480
         TabIndex        =   27
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pendientes de Aceptacion SII"
         Height          =   255
         Left            =   12480
         TabIndex        =   26
         Top             =   480
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Aceptados en el SII"
         Height          =   255
         Left            =   12480
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Carga archivos SII"
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin XPFrame.FrameXp CARGATXT 
         Height          =   4560
         Left            =   3360
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   8043
         BackColor       =   16761024
         Caption         =   "BUSCAR "
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
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FF8080&
            Caption         =   "PROCESAR LIBROS"
            Height          =   465
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   3480
            Width           =   2625
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FF8080&
            Caption         =   "RETORNO"
            Height          =   465
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3465
            Width           =   2625
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FF8080&
            Caption         =   "PROCESAR"
            Height          =   465
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3465
            Width           =   2625
         End
         Begin VB.DirListBox Dir1 
            Height          =   2565
            Left            =   180
            TabIndex        =   18
            Top             =   765
            Width           =   3855
         End
         Begin VB.TextBox ARCHIVO 
            Height          =   285
            Left            =   4230
            TabIndex        =   17
            Top             =   3060
            Width           =   4275
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   180
            TabIndex        =   16
            Top             =   315
            Width           =   3855
         End
         Begin VB.FileListBox File1 
            Height          =   2235
            Left            =   4230
            TabIndex        =   15
            Top             =   315
            Width           =   4275
         End
         Begin MSComctlLib.ProgressBar barra2 
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   4080
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ARCHIVO SELECCIONADO"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4230
            TabIndex        =   23
            Top             =   2790
            Width           =   4290
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   8880
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command3 
         Caption         =   "EXCEL"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   9240
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GENERA INFORME"
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "IMPRIMIR"
         Height          =   255
         Left            =   4800
         TabIndex        =   3
         Top             =   9240
         Width           =   2535
      End
      Begin FlexCell.Grid Grid2 
         Height          =   7560
         Left            =   0
         TabIndex        =   2
         Top             =   1320
         Width           =   19080
         _ExtentX        =   33655
         _ExtentY        =   13335
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   735
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1296
         BackColor       =   16744576
         Caption         =   "MES"
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
         Begin VB.ComboBox COMBOMES 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   3615
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   735
         Left            =   9480
         TabIndex        =   6
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         BackColor       =   16744576
         Caption         =   " AÑO"
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
         Begin VB.ComboBox COMBOAÑO 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   5055
      End
   End
End
Attribute VB_Name = "compra032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public saldoglobal As Double
Private moneda As String
Private rutpropi As String

Private MODIFI As Integer

 

Private Sub Check1_Click()
Command2_Click


End Sub

Private Sub COMBOAÑO_Change()
'leerCONSUMOS
End Sub

Private Sub COMBOMES_Change()
'leerCONSUMOS
End Sub

Private Sub Command1_Click()
Titulos
Grid2.PrintPreview

End Sub
Sub Titulos()
    

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid2.FixedRowColStyle = Fixed3D
    Grid2.CellBorderColorFixed = vbButtonShadow
    Grid2.ShowResizeTips = False
    Grid2.PageSetup.Orientation = cellPortrait
    Grid2.DefaultFont.Size = 7.5
    Grid2.PageSetup.PrintFixedRow = True
    Grid2.ReportTitles.Clear
    Grid2.PageSetup.CenterHorizontally = True
    Grid2.PageSetup.PrintTitleRows = 0
    Grid2.PageSetup.BlackAndWhite = False
    Grid2.PageSetup.Orientation = cellLandscape
    'Logo
  
    'ENCABEZADO DE PAGINA
    Grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid2.PageSetup.HeaderAlignment = CellLeft
    Grid2.PageSetup.HeaderFont.Name = "Verdana"
    Grid2.PageSetup.HeaderFont.Size = 8
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE CONSUMOS Y SUS ESTADOS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
        
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Underline = True
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    Grid2.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D " & vbCrLf & "Usuario:" & USUARIOSISTEMA
    Grid2.PageSetup.FooterAlignment = cellRight
    Grid2.PageSetup.FooterFont.Name = "Verdana"
    Grid2.PageSetup.FooterFont.Size = 7
    Grid2.PageSetup.LeftMargin = 0.5
    Grid2.PageSetup.RightMargin = 0.5
    
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThick
    
    
    
    
    
End Sub

Private Sub Command2_Click()
MES = COMBOMES.ListIndex + 1
    If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
 If Option1.Value = True Then
    
Label1.Caption = "FECHA ACTUALIZACION ACEPTADOS " + fecha_aceptados(MES, año)
DIFERENCIA = DateDiff("d", fecha_aceptados(MES, año), Date)
If DIFERENCIA > 1 Then
Rem MsgBox "EL ARCHIVO DE COMPARACION TIENE " + DIFERENCIA + " DE DIFERENCIA FAVOR ACTUALIZAR "
End If
    
    
    LEERcompras_aceptados
 End If
 If Option2.Value = True Then
Label1.Caption = "FECHA ACTUALIZACION PENDIENTES " + fecha_pendientes
DIFERENCIA = DateDiff("d", fecha_pendientes, Date)
If DIFERENCIA > 1 Then
MsgBox "EL ARCHIVO DE COMPARACION TIENE " + DIFERENCIA + " DE DIFERENCIA FAVOR ACTUALIZAR "
End If
    
    
    LEERcompras_PENDIENTES
 End If
 
 
 End Sub

Sub CARGAGRILLA_PENDIENTES()
    Dim formatogrilla2(50, 50)
    formatogrilla2(1, 1) = "NRO"
    formatogrilla2(1, 2) = "TIPO DOC"
    formatogrilla2(1, 3) = "NOMBRE"
    formatogrilla2(1, 4) = "RUT PROVEEDOR"
    formatogrilla2(1, 5) = "FOLIO"
    formatogrilla2(1, 6) = "FECHA DOC."
    formatogrilla2(1, 7) = "FECHA RECEPCION"
    formatogrilla2(1, 8) = "MONTO EXENTO"
    formatogrilla2(1, 9) = "MONTO NETO"
    formatogrilla2(1, 10) = "MONTO IVA RECUPERABLE"
    formatogrilla2(1, 11) = "MONTO IVA NO RECUPERABLE"
    formatogrilla2(1, 12) = "CODIGO IVA NO REC"
    formatogrilla2(1, 13) = "MONTO TOTAL"
    formatogrilla2(1, 14) = "MONTO NETO ACTIVO FIJO"
    formatogrilla2(1, 15) = "IVA ACTIVO FIJO"
    formatogrilla2(1, 16) = "IVA USO COMUN"
    formatogrilla2(1, 17) = "IMPTO. SIN DERECHO A CREDITO"
    formatogrilla2(1, 18) = "IVA NO RETENIDO"
    formatogrilla2(1, 19) = "NCE O NDE SOBRE FACT. DE COMPRA"
    formatogrilla2(1, 20) = "CODIGO OTRO IMPUESTO"
    formatogrilla2(1, 21) = "VALOR OTRO IMPUESTO"
    formatogrilla2(1, 22) = "TASA OTRO IMPUESTO"
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "5"
    formatogrilla2(2, 3) = "20"
    formatogrilla2(2, 4) = "8"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "8"
    formatogrilla2(2, 7) = "8"
    formatogrilla2(2, 8) = "10"
    formatogrilla2(2, 9) = "10"
    formatogrilla2(2, 10) = "10"
    formatogrilla2(2, 11) = "10"
    formatogrilla2(2, 12) = "10"
    formatogrilla2(2, 13) = "10"
    formatogrilla2(2, 14) = "10"
    formatogrilla2(2, 15) = "10"
    formatogrilla2(2, 16) = "10"
    formatogrilla2(2, 17) = "10"
    formatogrilla2(2, 18) = "10"
    formatogrilla2(2, 19) = "10"
    formatogrilla2(2, 20) = "10"
    formatogrilla2(2, 21) = "10"
    formatogrilla2(2, 22) = "10"
    
   
   
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "D"
    formatogrilla2(3, 7) = "D"
    formatogrilla2(3, 8) = "D"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    formatogrilla2(3, 12) = "N"
    formatogrilla2(3, 13) = "N"
    formatogrilla2(3, 14) = "N"
    formatogrilla2(3, 15) = "N"
    formatogrilla2(3, 16) = "N"
    formatogrilla2(3, 17) = "N"
    formatogrilla2(3, 18) = "N"
    formatogrilla2(3, 19) = "N"
    formatogrilla2(3, 20) = "N"
    formatogrilla2(3, 21) = "N"
    formatogrilla2(3, 22) = "N"
    
   
    
    Rem FORMATO GRILLA
 
    formatogrilla2(4, 8) = "###,###,###,##0"
    formatogrilla2(4, 9) = "###,###,###,##0"
    formatogrilla2(4, 10) = "###,###,###,##0"
    formatogrilla2(4, 11) = "###,###,###,##0"
    formatogrilla2(4, 13) = "###,###,###,##0"
    formatogrilla2(4, 14) = "###,###,###,##0"
    formatogrilla2(4, 15) = "###,###,###,##0"
    formatogrilla2(4, 16) = "###,###,###,##0"
    formatogrilla2(4, 17) = "###,###,###,##0"
    formatogrilla2(4, 18) = "###,###,###,##0"
    formatogrilla2(4, 21) = "###,###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "TRUE"
    formatogrilla2(5, 11) = "TRUE"
    formatogrilla2(5, 12) = "TRUE"
    formatogrilla2(5, 13) = "TRUE"
    formatogrilla2(5, 14) = "TRUE"
    formatogrilla2(5, 15) = "TRUE"
    formatogrilla2(5, 16) = "TRUE"
    formatogrilla2(5, 17) = "TRUE"
    formatogrilla2(5, 18) = "TRUE"
    formatogrilla2(5, 19) = "TRUE"
    formatogrilla2(5, 20) = "TRUE"
    formatogrilla2(5, 21) = "TRUE"
    formatogrilla2(5, 22) = "TRUE"
    Rem VALOR MAXIMO
    
    Grid2.Cols = 23
    Grid2.Rows = 1
    Grid2.AllowUserResizing = True
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = False
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
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftCenter
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
  
    
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).WrapText = True
     Grid2.RowHeight(0) = 40
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
    End Sub


Private Sub Command3_Click()
    If Grid2.Rows > 0 Then
        Call Grid2.ExportToExcel("", True, True)
    End If
End Sub

Private Sub Command6_Click()
    MsgBox "RECUERDE TENER EL ARCHIVO CON EL NOMBRE ORIGINAL Y .CSV EN LA CARPETA u:\aceptados_sii "
    CARGATXT.Visible = True
    Call Command7_Click
    CARGATXT.Visible = False
    MsgBox "PROCESO TERMINADO ", vbInformation, "ATENCION"
End Sub

Private Sub Command7_Click()
    Dim disco As String
    Dim s As Double
    
    disco = "U:"
    Dir1.path = disco
    File1.path = disco + "\aceptados_sii\"
    File1.Pattern = "*.csv"
    
    File1.Refresh
    Dim o As Double
    Dim origen As String
    Dim destino As String
    
    
    For o = 0 To File1.ListCount - 1
    
            ARCHIVO.text = File1.List(o)
                    If UCase(Right(ARCHIVO.text, 3)) = "CSV" And InStr(ARCHIVO.text, "RCV_COMPRA_REGISTRO") > 0 Then
                              Rem CARGATXT.Visible = True
    
                               TRASPASADATOS2
                    
    
                    Rem CARGATXT.Visible = False
                             origen = "u:\aceptados_sii\" + ARCHIVO.text
                             destino = "u:\aceptados_sii_usados\" + ARCHIVO.text
                             Call FileCopy(origen, destino)
    
                             Kill origen
                    End If
    
                     If UCase(Right(ARCHIVO.text, 3)) = "CSV" And InStr(ARCHIVO.text, "RCV_COMPRA_PENDIENTE") > 0 Then
                              Rem CARGATXT.Visible = True
    
                               TRASPASADATOS3
                    
    
                    Rem CARGATXT.Visible = False
                             origen = "u:\aceptados_sii\" + ARCHIVO.text
                             destino = "u:\aceptados_sii_usados\" + ARCHIVO.text
                             Call FileCopy(origen, destino)
    
                             Kill origen
                    End If
    
    
    Next o
    
    
            


End Sub


Private Sub command8_Click()
For Q = 1 To Grid2.Rows - 1
If Grid2.Cell(Q, 1).text = 1 And Check1.Value = 1 Then
    lc_rut = Grid2.Cell(Q, 4).text
         CUENTA2 = Right(lc_rut, 1)
    lc_rut = Format(Mid(lc_rut, 1, Len(lc_rut) - 2), "000000000")
    lc_rut = lc_rut + CUENTA2
    If lee_factura_de_compra(Grid2.Cell(Q, 2).text, Format(Grid2.Cell(Q, 5).text, "0000000000"), lc_rut) = False Then
        
        lc_tipodte = Grid2.Cell(Q, 2).text
        lc_folio = Grid2.Cell(Q, 5).text
        lc_fchemis = Grid2.Cell(Q, 6).text
        
        lc_neto = Grid2.Cell(Q, 10).text
        lc_iva = Grid2.Cell(Q, 11).text
        lc_exento = Grid2.Cell(Q, 9).text
        lc_total = Grid2.Cell(Q, 14).text
        
        lc_tipo(1) = Grid2.Cell(Q, 24).text
        lc_monto(1) = Grid2.Cell(Q, 25).text
        
         Call grabafactura(lc_tipodte, lc_folio, lc_fchemis, lc_fchemis, lc_rut, lc_neto, lc_iva, lc_exento, 0, lc_total)
         Call crearcuentacorriente(lc_rut, Grid2.Cell(Q, 3).text, "", "", "", "", "")
     End If
End If
Next Q
Command2_Click

End Sub

Private Sub Command9_Click()
If Check1.Value = 1 Then
For Q = 1 To Grid2.Rows - 1
If Grid2.Cell(Q, 2).text = "61" Then
    Grid2.Cell(Q, 1).text = "1"
End If
       
Next Q
End If


End Sub

Private Sub Form_Load()
'Call CENTRAR(Me)
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
 
    Rem Call RECUPERAFECHA
    For k = 1 To 12
    COMBOMES.AddItem MonthName(k)
    Next k
    COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
    For k = 2000 To Val(Format(fechasistema, "yyyy"))
    COMBOAÑO.AddItem k
    Next k
    COMBOAÑO.ListIndex = k - 2001
    
    'Call CARGAPERMISO(Me.Name)
    
    
     

End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub


Sub CARGAGRILLA_ACEPTADOS()
    Dim formatogrilla2(50, 50)
    formatogrilla2(1, 1) = "NRO"
    formatogrilla2(1, 2) = "TIPO DOC"
    formatogrilla2(1, 3) = "NOMBRE"
    formatogrilla2(1, 4) = "RUT PROVEEDOR"
    formatogrilla2(1, 5) = "FOLIO"
    formatogrilla2(1, 6) = "FECHA DOC"
    formatogrilla2(1, 7) = "FECHA RECEPCION"
    formatogrilla2(1, 8) = "FECHA ACUSE"
    formatogrilla2(1, 9) = "MONTO EXENTO"
    formatogrilla2(1, 10) = "MONTO NETO"
    formatogrilla2(1, 11) = "MONTO IVA RECUPERABLE"
    formatogrilla2(1, 12) = "MONTO IVA NO RECUPERABLE"
    formatogrilla2(1, 13) = "CODIGO IVA NO REC"
    formatogrilla2(1, 14) = "MONTO TOTAL"
    formatogrilla2(1, 15) = "MONTO NETO ACTIVO FIJO"
    formatogrilla2(1, 16) = "IVA ACTIVA FIJO"
    formatogrilla2(1, 17) = "IVA USO COMUN"
    formatogrilla2(1, 18) = "IMPTO SIN DERECHO A CREDITO"
    formatogrilla2(1, 19) = "IVA NO RETENIDO"
    formatogrilla2(1, 20) = "TABACOS PUROS"
    formatogrilla2(1, 21) = "TABACOS CIGARRILLOS"
    formatogrilla2(1, 22) = "TABACOS ELABORADOS"
    formatogrilla2(1, 23) = "NCE O NDE SOBRE FAC DE COMPRA"
    formatogrilla2(1, 24) = "CODIGO OTRO IMPUESTO"
    formatogrilla2(1, 25) = "VALOR OTRO IMPUESTO"
    formatogrilla2(1, 26) = "TASA OTRO IMPUESTO"

    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "5"
    formatogrilla2(2, 3) = "20"
    formatogrilla2(2, 4) = "8"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "8"
    formatogrilla2(2, 7) = "8"
    formatogrilla2(2, 8) = "8"
    formatogrilla2(2, 9) = "10"
    formatogrilla2(2, 10) = "10"
    formatogrilla2(2, 11) = "10"
    formatogrilla2(2, 12) = "10"
    formatogrilla2(2, 13) = "10"
    formatogrilla2(2, 14) = "10"
    formatogrilla2(2, 15) = "10"
    formatogrilla2(2, 16) = "10"
    formatogrilla2(2, 17) = "10"
    formatogrilla2(2, 18) = "10"
    formatogrilla2(2, 19) = "10"
    formatogrilla2(2, 20) = "10"
    formatogrilla2(2, 21) = "10"
    formatogrilla2(2, 22) = "10"
    formatogrilla2(2, 23) = "10"
    formatogrilla2(2, 24) = "10"
    formatogrilla2(2, 25) = "10"
    formatogrilla2(2, 26) = "10"
    
  
   
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "D"
    formatogrilla2(3, 7) = "D"
    formatogrilla2(3, 8) = "D"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    formatogrilla2(3, 12) = "N"
    formatogrilla2(3, 13) = "N"
    formatogrilla2(3, 14) = "N"
    formatogrilla2(3, 15) = "N"
    formatogrilla2(3, 16) = "N"
    formatogrilla2(3, 17) = "N"
    formatogrilla2(3, 18) = "N"
    formatogrilla2(3, 19) = "N"
    formatogrilla2(3, 20) = "N"
    formatogrilla2(3, 21) = "N"
    formatogrilla2(3, 22) = "N"
    formatogrilla2(3, 23) = "N"
    formatogrilla2(3, 24) = "N"
    formatogrilla2(3, 25) = "N"
    formatogrilla2(3, 26) = "N"
   
    
   
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 9) = " ###,###,###,##0"
    formatogrilla2(4, 10) = " ###,###,###,##0"
    formatogrilla2(4, 11) = " ###,###,###,##0"
    formatogrilla2(4, 13) = " ###,###,###,##0"
    formatogrilla2(4, 14) = " ###,###,###,##0"
    formatogrilla2(4, 15) = " ###,###,###,##0"
    formatogrilla2(4, 16) = " ###,###,###,##0"
    formatogrilla2(4, 17) = " ###,###,###,##0"
    formatogrilla2(4, 18) = " ###,###,###,##0"
    formatogrilla2(4, 25) = " ###,###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "TRUE"
    formatogrilla2(5, 11) = "TRUE"
    formatogrilla2(5, 12) = "TRUE"
    formatogrilla2(5, 13) = "TRUE"
    formatogrilla2(5, 14) = "TRUE"
    formatogrilla2(5, 15) = "TRUE"
    formatogrilla2(5, 16) = "TRUE"
    formatogrilla2(5, 17) = "TRUE"
    formatogrilla2(5, 18) = "TRUE"
    formatogrilla2(5, 19) = "TRUE"
    formatogrilla2(5, 20) = "TRUE"
    formatogrilla2(5, 21) = "TRUE"
    formatogrilla2(5, 22) = "TRUE"
    formatogrilla2(5, 23) = "TRUE"
    formatogrilla2(5, 24) = "TRUE"
    formatogrilla2(5, 25) = "TRUE"
    formatogrilla2(5, 26) = "TRUE"
    
   
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 27
    Grid2.Rows = 1
    Grid2.AllowUserResizing = True
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = False
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
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).WrapText = True
     Grid2.RowHeight(0) = 40
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
Grid2.Column(1).CellType = cellCheckBox
Grid2.Column(1).Locked = False


    
    
    End Sub


 Public Sub LEERcompras_aceptados()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 Dim MES As String
 Dim año As String
 
 
 CARGAGRILLA_ACEPTADOS
 On Error GoTo salida:
 año = COMBOAÑO.text
    MES = COMBOMES.ListIndex + 1
    If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
    
    
 Set csql.ActiveConnection = db
csql.sql = "SELECT * FROM sii_lc_" + MES + "_" + año
 
 
 csql.Execute
nova:
 
 Grid2.Rows = 1
 Grid2.AutoRedraw = False
 
 If csql.RowsAffected > 0 Then
    barra.Max = csql.RowsAffected + 1
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    If Check1.Value = 1 And esta_en_libro_compras(resultados(1), resultados(4), resultados(3), resultados("montototal"), empresaactiva, MES, año) = True Then
     GoTo paso:
    
    End If
    
    
    Grid2.Rows = Grid2.Rows + 1
    barra.Value = Grid2.Rows
    
    Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
    Grid2.Cell(Grid2.Rows - 1, 3).text = LEERNOMBREPROVEEDOR(resultados(3))
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    If IsNull(resultados(6)) = False Then
        Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Else
         Grid2.Cell(Grid2.Rows - 1, 7).text = "0000-00-00"
    End If
    
    If IsNull(resultados(7)) = False Then
        Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Else
        Grid2.Cell(Grid2.Rows - 1, 8).text = "0000-00-00"
    End If
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
    Grid2.Cell(Grid2.Rows - 1, 10).text = resultados(9)
    Grid2.Cell(Grid2.Rows - 1, 11).text = resultados(10)
    Grid2.Cell(Grid2.Rows - 1, 12).text = resultados(11)
    Grid2.Cell(Grid2.Rows - 1, 13).text = resultados(12)
    Grid2.Cell(Grid2.Rows - 1, 14).text = resultados(13)
    Grid2.Cell(Grid2.Rows - 1, 15).text = resultados(14)
    Grid2.Cell(Grid2.Rows - 1, 16).text = resultados(15)
    Grid2.Cell(Grid2.Rows - 1, 17).text = resultados(16)
    Grid2.Cell(Grid2.Rows - 1, 18).text = resultados(17)
    Grid2.Cell(Grid2.Rows - 1, 19).text = resultados(18)
    Grid2.Cell(Grid2.Rows - 1, 20).text = resultados(19)
    Grid2.Cell(Grid2.Rows - 1, 21).text = resultados(20)
    Grid2.Cell(Grid2.Rows - 1, 22).text = resultados(21)
    Grid2.Cell(Grid2.Rows - 1, 23).text = resultados(22)
    Grid2.Cell(Grid2.Rows - 1, 24).text = resultados(23)
    Grid2.Cell(Grid2.Rows - 1, 25).text = resultados(24)
    Grid2.Cell(Grid2.Rows - 1, 26).text = resultados(25)
    If esta_en_libro_compras(resultados(1), resultados(4), resultados(3), resultados("montototal"), empresaactiva, MES, año) = True Then
    Grid2.Cell(Grid2.Rows - 1, 3).BackColor = vbGreen
    
    End If
    
    
    
    
   
paso:
    resultados.MoveNext



    Wend


  End If
 
  
  Grid2.AutoRedraw = True
  Grid2.Refresh
  
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
Exit Sub
salida:
MsgBox "mes no esta procesado "
 End Sub
 
Public Sub LEERcompras_PENDIENTES()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 Dim MES As String
 Dim año As String
 
 
 CARGAGRILLA_PENDIENTES
 On Error GoTo salida:
 año = COMBOAÑO.text
 MES = COMBOMES.ListIndex + 1
 If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
    
    
 Set csql.ActiveConnection = db
csql.sql = "SELECT * FROM sii_lp_99 "
 
 
 csql.Execute
nova:
 
 Grid2.Rows = 1
 Grid2.AutoRedraw = False
 
 If csql.RowsAffected > 0 Then
    barra.Max = csql.RowsAffected + 1
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    If Check1.Value = 1 And esta_en_libro_compras(resultados(1), resultados(4), resultados(3), resultados("montototal"), empresaactiva, MES, año) = True Then
     GoTo paso:
    
    End If
    
    
    Grid2.Rows = Grid2.Rows + 1
    barra.Value = Grid2.Rows
    
    Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
    Grid2.Cell(Grid2.Rows - 1, 3).text = LEERNOMBREPROVEEDOR(resultados(3))
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    If IsNull(resultados(6)) = False Then
        Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Else
         Grid2.Cell(Grid2.Rows - 1, 7).text = "0000-00-00"
    End If
    
    If IsNull(resultados(7)) = False Then
        Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Else
        Grid2.Cell(Grid2.Rows - 1, 8).text = "0000-00-00"
    End If
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
    Grid2.Cell(Grid2.Rows - 1, 10).text = resultados(9)
    Grid2.Cell(Grid2.Rows - 1, 11).text = resultados(10)
    Grid2.Cell(Grid2.Rows - 1, 12).text = resultados(11)
    Grid2.Cell(Grid2.Rows - 1, 13).text = resultados(12)
    Grid2.Cell(Grid2.Rows - 1, 14).text = resultados(13)
    Grid2.Cell(Grid2.Rows - 1, 15).text = resultados(14)
    Grid2.Cell(Grid2.Rows - 1, 16).text = resultados(15)
    Grid2.Cell(Grid2.Rows - 1, 17).text = resultados(16)
    Grid2.Cell(Grid2.Rows - 1, 18).text = resultados(17)
    Grid2.Cell(Grid2.Rows - 1, 19).text = resultados(18)
    Grid2.Cell(Grid2.Rows - 1, 20).text = resultados(19)
    Grid2.Cell(Grid2.Rows - 1, 21).text = resultados(20)
    
    If esta_en_libro_compras(resultados(1), resultados(4), resultados(3), resultados("montototal"), empresaactiva, MES, año) = True Then
    Grid2.Cell(Grid2.Rows - 1, 3).BackColor = vbGreen
    
    End If

    
    
    
paso:

    resultados.MoveNext



    Wend


  End If
 
  
  Grid2.AutoRedraw = True
  Grid2.Refresh
  
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
Exit Sub
salida:
MsgBox "mes no esta procesado "
 End Sub
 

Public Sub LEERVENTAS()
' Dim csql As New rdoQuery
' Dim resultados As rdoResultset
'
'
' CARGAGRILLA2
' Set csql.ActiveConnection = db
' empresa_fae = leerdatos(conta, "maestroempresas", "empresafae", "codigoempresa='" + empresaactiva + "' ")
' csql.SQL = "SELECT fc.tipo,fc.rut,cc.nombre,fc.numero,fc.fecha,fc.iva,fc.total,ifnull(f.iva,0) as iva,ifnull(f.total,0) as total,f.tipo FROM "
' csql.SQL = csql.SQL & "facturasdeventas AS fc LEFT JOIN cuentascorrientes AS cc ON  fc.rut=cc.rut LEFT JOIN "
' csql.SQL = csql.SQL & cliente_sql & "fae" & empresa_fae & ".sv_dte_libros_sii_ventas AS f "
' csql.SQL = csql.SQL & "ON f.rut=fc.rut AND f.numero=fc.numero AND f.fecha=fc.fecha "
' csql.SQL = csql.SQL & "WHERE fc.tipo<>'' AND cc.año='" & COMBOAÑO.text & "' AND cc.tipo='11200027' AND "
' csql.SQL = csql.SQL & "fc.fecha LIKE '" & COMBOAÑO.text & "-" & Format(COMBOMES.ListIndex + 1, "00") & "%' "
' If CHK1.Value = 1 Then
''    csql.sql = csql.sql & " having fc.total=total or fc.iva<>iva "
'    csql.SQL = csql.SQL & " and f.total=0 "
' End If
' csql.SQL = csql.SQL & " ORDER BY fc.fecha "
'
' csql.Execute
'
'
' Grid2.Rows = 1
' Grid2.AutoRedraw = False
'
' If csql.RowsAffected > 0 Then
'    barra.Max = csql.RowsAffected + 1
'    Set resultados = csql.OpenResultset
'    While resultados.EOF = False
'    Grid2.Rows = Grid2.Rows + 1
'    barra.Value = Grid2.Rows
'
'                If resultados(0) = "1" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FA"
'                If resultados(0) = "2" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "ND"
'                If resultados(0) = "3" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
'                If resultados(0) = "4" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
'                If resultados(0) = "5" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEX"
'                If resultados(0) = "6" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FAE"
'                If resultados(0) = "7" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NDE"
'                If resultados(0) = "8" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NCE"
'                If resultados(0) = "9" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FE"
'                If resultados(0) = "0" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEE"
'                If resultados(0) = "L" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "LFE"
'
'
'    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
'    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
'    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
'    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
'    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
'    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
'    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
'    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
'
'   If resultados(8) = 0 Then
'     Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = vbRed
'   End If
'
'    resultados.MoveNext
'
'
'
'    Wend
'
'
'  End If
'  Call buscanoencontradosventas(COMBOAÑO.text, Format(COMBOMES.ListIndex + 1, "00"), empresa_fae)
'
'  Grid2.AutoRedraw = True
'  Grid2.Refresh
'
' csql.Close
' Set csql = Nothing
' Set resultados = Nothing

 End Sub
 
 Sub buscanoencontradoscompras(año, MES, loc)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = db
    
     
    csql.sql = " SELECT tipo,rut,numero,fecha,iva,total FROM " & cliente_sql & "conta" & loc & ".sv_dte_aceptados_sii_compras "
    csql.sql = csql.sql & "WHERE  mescontable='" & MES & "' AND añocontable='" & año & "' "
    csql.sql = csql.sql & "AND numero NOT IN (SELECT numero FROM  facturasdecompras AS fc  "
    csql.sql = csql.sql & "WHERE fc.añocontable='" & año & "' AND fc.mescontable='" & MES & "')"
    csql.Execute
'    Grid2.Rows = Grid2.Rows + 1
    If csql.RowsAffected > 0 Then
        Grid2.AutoRedraw = False
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Grid2.Rows = Grid2.Rows + 1
            
                If resultados(0) = "30" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FA"
                If resultados(0) = "55" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "ND"
                If resultados(0) = "60" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
                If resultados(0) = "33" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FAE"
                If resultados(0) = "56" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NDE"
                If resultados(0) = "61" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NCE"
                If resultados(0) = "46" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FC"
                If resultados(0) = "914" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "IM"
                If resultados(0) = "32" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FE"
                If resultados(0) = "34" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEE"
                If resultados(0) = "43" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "LFE"
                
                
         
            Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
            Grid2.Cell(Grid2.Rows - 1, 3).text = LEERNOMBREPROVEEDOR(resultados(1))
            Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(2)
            Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(3)
            Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(4)
            Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(5)
            Grid2.Cell(Grid2.Rows - 1, 6).text = "0"
            Grid2.Cell(Grid2.Rows - 1, 7).text = "0"
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = vbYellow
            resultados.MoveNext
        Wend
        Grid2.AutoRedraw = True
        Grid2.Refresh
        
    End If
    
    
 End Sub
 Public Function LEERNOMBREPROVEEDOR(rut) As String
    CAMPOS(0, 0) = "razonsocial"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = clientesistema + "fae.sv_fae_proveedores"
    condicion = "rut= '" & rut & "'  "
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LEERNOMBREPROVEEDOR = sqlconta.response(0, 3)
    Else
        LEERNOMBREPROVEEDOR = ""
    End If
 
    

End Function
 Sub buscanoencontradosventas(año, MES, loc)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = db
    
     
    csql.sql = " SELECT tipo,rut,numero,fecha,iva,total FROM " & cliente_sql & "fae" & loc & ".sv_dte_libros_sii_ventas "
    csql.sql = csql.sql & "WHERE  mescontable='" & MES & "' AND añocontable='" & año & "' "
    csql.sql = csql.sql & "AND numero NOT IN (SELECT numero FROM  facturasdeventas AS fc  "
    csql.sql = csql.sql & "WHERE fc.fecha like '" & año & "-" & MES & "%') "
    csql.Execute
'    Grid2.Rows = Grid2.Rows + 1
    If csql.RowsAffected > 0 Then
        Grid2.AutoRedraw = False
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Grid2.Rows = Grid2.Rows + 1
            
                If resultados(0) = "30" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FA"
                If resultados(0) = "55" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "ND"
                If resultados(0) = "60" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
                If resultados(0) = "33" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FAE"
                If resultados(0) = "56" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NDE"
                If resultados(0) = "61" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NCE"
                If resultados(0) = "46" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FC"
                If resultados(0) = "914" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "IM"
                If resultados(0) = "32" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FE"
                If resultados(0) = "34" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEE"
                If resultados(0) = "43" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "LFE"
                
                
         
            Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
            Grid2.Cell(Grid2.Rows - 1, 3).text = LEERNOMBREPROVEEDOR(resultados(1))
            Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(2)
            Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(3)
            Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(4)
            Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(5)
            Grid2.Cell(Grid2.Rows - 1, 6).text = "0"
            Grid2.Cell(Grid2.Rows - 1, 7).text = "0"
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = vbYellow
            resultados.MoveNext
        Wend
        Grid2.AutoRedraw = True
        Grid2.Refresh
        
    End If
    
    
 End Sub
 

Private Function leemonedas(codigo) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = conta

csql.sql = "select nombremoneda from " & clientesistema & "arriendos" & ".maestro_monedas where codigomoneda='" & codigo & "'"
csql.Execute
leemonedas = ""
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leemonedas = resultados(0)
End If
Set resultados = Nothing
csql.Close
Set csql = Nothing

End Function

Public Function LEERULTIMOFOLIOcontrato() As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = db

            csql.sql = "select IFNULL(max(numero),0) from " + clientesistema + "arriendos.contratos_arriendo"
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIOcontrato = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function


 
Sub ayudaTIPO(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigo", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Tipos de Consumos Basicos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "consumos_basicos", Usuario, password, "maestro_tipo_consumos", caja, CAMPOS, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

 

    
 
Sub ayudaempresa(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigoempresa", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Empresas"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestroempresas", caja, CAMPOS, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub




'Private Sub dato1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 27 Then Unload Me
'    snum = 0: KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 Then
'
'    Call ceros(dato1)
'    If leetipoconsumo(dato1.text) <> "" Then
'    LBLTIPO.Caption = leetipoconsumo(dato1.text)
'    leerCONSUMOS
'
'    Else
'    dato1.SetFocus
'    End If
'    End If
'
'
'
'
'End Sub
 

 

Sub TRASPASADATOS3()
Dim lin As Double
Dim palabras() As String
Dim palabras2() As String
Dim varipaso As String
Dim empresacontable As String
Dim datos As Variant
Dim periodo As String
Dim tipolibro As String


barra2.Max = 100
barra2.Refresh
barra2.Value = 0
Close 20

palabras() = Split(ARCHIVO.text, "_")
periodo = Mid(Replace(palabras(4), ".csv", ""), 1, 6)
empresacontable = leercodigocontable2(palabras(3))
tipolibro = palabras(1)

Call carga_archivo2(ARCHIVO, Mid(periodo, 5, 2), Mid(periodo, 1, 4), empresacontable)

'End If

End Sub

Sub TRASPASADATOS2()
Dim lin As Double
Dim palabras() As String
Dim palabras2() As String
Dim varipaso As String
Dim empresacontable As String
Dim datos As Variant
Dim periodo As String
Dim tipolibro As String


barra2.Max = 100
barra2.Refresh
barra2.Value = 0
Close 20

palabras() = Split(ARCHIVO.text, "_")
periodo = Mid(Replace(palabras(4), ".csv", ""), 1, 6)
empresacontable = leercodigocontable2(palabras(3))
tipolibro = palabras(1)

Call carga_archivo(ARCHIVO, Mid(periodo, 5, 2), Mid(periodo, 1, 4), empresacontable)

'End If

End Sub



Sub GRABACARTOLALIBROS(nro, tipodoc, tipocompra, rutproveedor, FOLIO, fechadocto, fecharecepcion, fechaacuse, montoexento, montoneto, montoivarecuperable, montoivanorecuperable, codigoivanorec, montototal _
, montonetoactivofijo, ivaactivofijo, ivausocomun, imptosinderechoacredito, ivanoretenido, tabacospuros, tabacoscigarrillos, tabacoselaborados, nceondesobrefacdecompra _
, codigootroimpuesto, valorotroimpuesto, tasaotroimpuesto, loc, periodo)
    Dim tipo_dte2 As String
    Dim rut3 As String
    Dim dato As Variant
    If InStr(1, rutproveedor, "-") = 0 Then GoTo no:
    dato = Split(rutproveedor, "-")
 

    
    CAMPOS(0, 0) = "nro"
    CAMPOS(1, 0) = "tipodoc"
    CAMPOS(2, 0) = "tipocompra"
    CAMPOS(3, 0) = "rutproveedor"
    CAMPOS(4, 0) = "folio"
    CAMPOS(5, 0) = "fechadocto"
    CAMPOS(6, 0) = "fecharecepcion"
    CAMPOS(7, 0) = "fechaacuse"
    CAMPOS(8, 0) = "montoexento"
    CAMPOS(9, 0) = "montoneto"
    CAMPOS(10, 0) = "montoivarecuperable"
    CAMPOS(11, 0) = "montoivanorecuperable"
    CAMPOS(12, 0) = "codigoivanorec"
    CAMPOS(13, 0) = "montototal"
    CAMPOS(14, 0) = "montonetoactivofijo"
    CAMPOS(15, 0) = "ivaactivofijo"
    CAMPOS(16, 0) = "ivausocomun"
    CAMPOS(17, 0) = "imptosinderechoacredito"
    CAMPOS(18, 0) = "ivanoretenido"
    CAMPOS(19, 0) = "tabacospuros"
    CAMPOS(20, 0) = "tabacoscigarrillos"
    CAMPOS(21, 0) = "tabacoselaborados"
    CAMPOS(22, 0) = "nceondesobrefacdecompra"
    CAMPOS(23, 0) = "codigootroimpuesto"
    CAMPOS(24, 0) = "valorotroimpuesto"
    CAMPOS(25, 0) = "tasaotroimpuesto"
    CAMPOS(26, 0) = "mescontable"
    CAMPOS(27, 0) = "añocontable"
    CAMPOS(28, 0) = ""
 
    CAMPOS(0, 1) = nro
    CAMPOS(1, 1) = tipodoc
    CAMPOS(2, 1) = tipocompra
    CAMPOS(3, 1) = Format(dato(0), "000000000") & dato(1)
    CAMPOS(4, 1) = Format(FOLIO, "0000000000")
    CAMPOS(5, 1) = Format(fechadocto, "yyyy-mm-dd")
    CAMPOS(6, 1) = Format(fecharecepcion, "yyyy-mm-dd")
    CAMPOS(7, 1) = Format(fechaacuse, "yyyy-mm-dd")
    CAMPOS(8, 1) = montoexento
    CAMPOS(9, 1) = montoneto
    CAMPOS(10, 1) = montoivarecuperable
    CAMPOS(11, 1) = montoivanorecuperable
    CAMPOS(12, 1) = codigoivanorec
    CAMPOS(13, 1) = montototal
    CAMPOS(14, 1) = montonetoactivofijo
    CAMPOS(15, 1) = ivaactivofijo
    CAMPOS(16, 1) = ivausocomun
    CAMPOS(17, 1) = imptosinderechoacredito
    CAMPOS(18, 1) = ivanoretenido
    CAMPOS(19, 1) = tabacospuros
    CAMPOS(20, 1) = tabacoscigarrillos
    CAMPOS(21, 1) = tabacoselaborados
    CAMPOS(22, 1) = nceondesobrefacdecompra
    CAMPOS(23, 1) = codigootroimpuesto
    CAMPOS(24, 1) = valorotroimpuesto
    CAMPOS(25, 1) = tasaotroimpuesto
    CAMPOS(26, 1) = Mid(periodo, 5, 2)
    CAMPOS(27, 1) = Mid(periodo, 1, 4)
 
    CAMPOS(0, 2) = clientesistema + "conta" + loc + ".sv_dte_aceptados_sii_compras"
    
           
    op = 2
    
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
no:
End Sub
Public Function leercodigocontable2(rut) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = conta

        csql.sql = "SELECT codigocontable  "
        csql.sql = csql.sql & "FROM " & clientesistema & "gestion.g_maestroempresas "
        csql.sql = csql.sql & "WHERE rut like '%" & Replace(rut, "-", "") & "%' "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
        leercodigocontable2 = resultado(0)
        Else
        leercodigocontable2 = ""
        End If
End Function


Public Function carga_archivo(ARCHIVO, mesLC, añoLC, empresa) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = conta
        
        Rem borra la tabla
        On Error GoTo paso1:
        csql.sql = "DROP TABLE " + clientesistema + "conta" + empresa + ".sii_lc_" + mesLC + "_" + añoLC + " "
        csql.Execute
paso1:
        Rem genera la tabla
        
        csql.sql = "CREATE TABLE " + clientesistema + "conta" + empresa + ".sii_lc_" + mesLC + "_" + añoLC + " ( PRIMARY KEY(`nro`,`tipodoc`,`rutproveedor`,`folio`,`fechadocto`,`mescontable`,`añocontable`) )ENGINE=MYISAM COLLATE = latin1_swedish_ci COMMENT = '' SELECT `nro`, `tipodoc`, `tipocompra`, `rutproveedor`, `folio`, `fechadocto`, `fecharecepcion`, `fechaacuse`, `montoexento`, `montoneto`, `montoivarecuperable`, `montoivanorecuperable`, `codigoivanorec`, `montototal`, `montonetoactivofijo`, `ivaactivofijo`, `ivausocomun`, `imptosinderechoacredito`, `ivanoretenido`, `tabacospuros`, `tabacoscigarrillos`, `tabacoselaborados`, `nceondesobrefacdecompra`, `codigootroimpuesto`, `valorotroimpuesto`, `tasaotroimpuesto`, `mescontable`, `añocontable` FROM " + clientesistema + "conta.sii_lc_00_0000;"
        csql.Execute

        Rem carga el archivo
        
        csql.sql = "LOAD DATA LOCAL INFILE 'u://aceptados_sii/" + ARCHIVO + "' "
        csql.sql = csql.sql + "INTO TABLE " + clientesistema + "conta" + empresa + ".sii_lc_" + mesLC + "_" + añoLC + " FIELDS TERMINATED BY ';';"

        csql.Execute

        Rem borra el titulo


        csql.sql = "DELETE FROM " + clientesistema + "conta" + empresa + ".sii_lc_" + mesLC + "_" + añoLC + " WHERE tipodoc='Tip';"

        csql.Execute

End Function

Public Function carga_archivo2(ARCHIVO, mesLC, añoLC, empresa) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = conta
        
        Rem borra la tabla
        On Error GoTo paso1:
        csql.sql = "DROP TABLE " + clientesistema + "conta" + empresa + ".sii_lp_99 "
        csql.Execute
paso1:
        Rem genera la tabla
        
        csql.sql = "CREATE TABLE " + clientesistema + "conta" + empresa + ".sii_lp_99 ( PRIMARY KEY(`nro`,`tipodoc`,`rutproveedor`,`folio`) )ENGINE=MYISAM COLLATE = latin1_swedish_ci COMMENT = '' SELECT `nro`, `tipodoc`, `tipocompra`, `rutproveedor`, `folio`, `fechadocto`, `fecharecepcion`, `montoexento`, `montoneto`, `montoivarecuperable`, `montoivanorecuperable`, `codigoivanorec`, `montototal`, `montonetoactivofijo`, `ivaactivofijo`, `ivausocomun`, `imptosinderechoacredito`, `ivanoretenido`, `nceondesobrefactdecompra`, `codigootroimpuesto`, `valorotroimpuesto`, `tasaotroimpuesto`, `fecha` FROM `eltit_conta`.`sii_lp_00_0000` ;"
        
        csql.Execute

        Rem carga el archivo
        
        csql.sql = "LOAD DATA LOCAL INFILE 'u://aceptados_sii/" + ARCHIVO + "' "
        csql.sql = csql.sql + "INTO TABLE " + clientesistema + "conta" + empresa + ".sii_lp_99 FIELDS TERMINATED BY ';';"

        csql.Execute

        Rem borra el titulo


        csql.sql = "DELETE FROM " + clientesistema + "conta" + empresa + ".sii_lp_99 WHERE tipodoc='Tipo';"

        csql.Execute

End Function


Private Sub Grid2_DblClick()
Rem Shell "u:\Python27\python.exe u:\RegistroReclamo\get-pip.pyget-pip.py -O 3 -r " + Grid2.Cell(Grid2.ActiveCell.row, 1).text

End Sub

Private Sub Option1_Click()
Command2_Click

End Sub

Private Sub Option2_Click()
Command2_Click

End Sub
