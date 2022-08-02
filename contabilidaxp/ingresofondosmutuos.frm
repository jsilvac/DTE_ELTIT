VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form inver04 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Ingreso Inversiones"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   1230
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15270
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   615
      Left            =   12000
      TabIndex        =   29
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdimprimir 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Imprimir"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7440
      Width           =   1575
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7260
      Left            =   180
      TabIndex        =   6
      Top             =   45
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   12806
      BackColor       =   16744576
      Caption         =   "Numero Inversion"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "RETIRO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3105
         TabIndex        =   27
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "INVERSION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   26
         Top             =   1800
         Value           =   -1  'True
         Width           =   2535
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   4920
         Left            =   180
         TabIndex        =   18
         Top             =   2250
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   8678
         BackColor       =   16761024
         Caption         =   "Cartola de Fondos Mutuos"
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
            Height          =   4560
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   14685
            _ExtentX        =   25903
            _ExtentY        =   8043
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   1305
         Width           =   375
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "fecha"
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "fecha"
         Top             =   315
         Width           =   555
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   990
         Width           =   1455
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1410
         Left            =   7680
         TabIndex        =   12
         Top             =   600
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   2487
         BackColor       =   8454016
         Caption         =   "VALORES EN PESOS"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox TAZA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2655
            MaxLength       =   20
            TabIndex        =   24
            Tag             =   "fecha"
            Top             =   675
            Width           =   1455
         End
         Begin VB.TextBox dato5 
            BackColor       =   &H00E1FFFD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "fecha"
            Top             =   315
            Width           =   375
         End
         Begin VB.TextBox dato6 
            BackColor       =   &H00E1FFFD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3015
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "fecha"
            Top             =   315
            Width           =   375
         End
         Begin VB.TextBox dato7 
            BackColor       =   &H00E1FFFD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3465
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "fecha"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox monto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2655
            MaxLength       =   20
            TabIndex        =   13
            Tag             =   "fecha"
            Top             =   990
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TAZA INVERSION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   45
            TabIndex        =   25
            Top             =   675
            Width           =   2445
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FECHA INVERSION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   45
            TabIndex        =   23
            Top             =   315
            Width           =   2445
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   45
            TabIndex        =   14
            Top             =   1035
            Width           =   2445
         End
      End
      Begin VB.Label LBLEMPRESA 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   2835
         TabIndex        =   17
         Top             =   1305
         Width           =   4785
      End
      Begin VB.Label LBLTIPO 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   2835
         TabIndex        =   16
         Top             =   630
         Width           =   4785
      End
      Begin VB.Label LBLBANCO 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   2835
         TabIndex        =   15
         Top             =   315
         Width           =   4785
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMPRESA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1305
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   945
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   630
         Width           =   1815
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1800
      Left            =   180
      TabIndex        =   3
      Top             =   7290
      Width           =   7065
      _cx             =   12462
      _cy             =   3175
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "inver04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private tipocuenta As String
    Private cc As Integer
    Private FORMATOGRILLA(100, 20)
    Private formatogrilla2(100, 20)
    Private cdi As Integer
    Private CANDO As Integer
    Private existe As String
    Private AFECTO As Double
    Private EXENTO As Double
    Private MODIFI As String
        
    Private AUXILIAR(1000, 3) As String
    
    Private respu As String
    Private tipoctacte As String
    Private nlineas As Double
    Private DOCU(6) As String
    Private grilladetalle(1000, 13) As String
    Private SALDOPE As Double
    Private NETO As Double
     Private CUENTAMAYOR(999) As String
     
     Private TIENECTACTE(999) As String
     Private TIENECRCC(999) As String
     Private TIENEBANCO(999) As String
     Private TIENEILA(999) As String
     Private TIENEICA(999) As String
     Private TIENEIHA(999) As String
     Private TIENEACTIVO(999) As String
     Private MES As String
     Private año As String
     
     
    






Private Sub cmdimprimir_Click()
Dim titulo As String
If Grid1.Rows > 1 Then

    Call Titulos
    Grid1.PrintPreview
    
End If
End Sub

Sub Titulos()
    

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.PageSetup.Orientation = cellPortrait
    Grid1.DefaultFont.Size = 7.5
    Grid1.Column(1).Width = 140
    Grid1.Column(2).Width = 150
    Grid1.Column(3).Width = 60
    Grid1.Column(4).Width = 60
    Grid1.Column(5).Width = 60
    Grid1.Column(6).Width = 70
    Grid1.Column(7).Width = 55
    Grid1.Column(8).Width = 70
    Grid1.Column(9).Width = 55
   
    Grid1.PageSetup.PrintFixedRow = True
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.PrintTitleRows = 0
    Grid1.PageSetup.BlackAndWhite = False
 
    'Logo
  
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "INVERSIONES FONDOS MUTUOS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
        
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Underline = True
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D " & vbCrLf & "Usuario:" & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    Grid1.PageSetup.LeftMargin = 0.5
    Grid1.PageSetup.RightMargin = 0.5
    
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
     
End Sub
Private Sub dato5_GotFocus()
leer
leerfondos

End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO5): Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    ' If KeyAscii = 42 And SUMADEBE = SUMAHABER Then grabarcomprobante:retorno: dato3.Enabled = True: dato3.SetFocus
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
no:
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call Pregunta(dato7, TAZA)
    
    End If
    
End Sub
Private Sub taza_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call Pregunta(TAZA, monto)
    
    End If
    
End Sub

Private Sub Form_Load()
CENTRAR Me
iva = 19
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    sc = 0
    opciones.Visible = False
CARGAGRILLA
leerfondos

End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudabancos(dato1)
    If KeyCode = 38 Then Unload Me: GoTo no:
    
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub
Sub ayudabancos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigobanco", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Bancos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestrobancos", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudadepositos(dato2)
    
    Call flechas(dato1, dato3, KeyCode)
End Sub
Sub ayudadepositos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Depositos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestrodepositos", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub


Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub
 
 Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato4)
    
    Call flechas(dato3, DATO5, KeyCode)


End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Empresas"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestroempresas", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub


Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(DATO5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, monto, KeyCode)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        
    Call ceros(dato1)
    If leerbanco(dato1.text) <> "" Then
    lblBanco.Caption = leerbanco(dato1.text)
    dato2.Enabled = True
    
    dato2.SetFocus
    Else
    dato1.SetFocus
    End If
    End If
    
    
    
    
End Sub
Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato2)
    If leerdeposito(dato2.text) <> "" Then
    LBLTIPO.Caption = leerdeposito(dato2.text)
    dato3.Enabled = True
    
    dato3.SetFocus
    Else
    dato2.SetFocus
    End If
    
    End If
    
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato4)
    If leerempresa(dato4.text) <> "" Then
    LBLEMPRESA.Caption = leerempresa(dato4.text)
    DATO5.Enabled = True
    
    DATO5.SetFocus
    Else
    dato4.SetFocus
    End If
    
    
    End If
    
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus: caja.SelStart = 0
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus: caja.SelStart = 0
End Sub
Sub grabar(banco, tipo, numero, empresa, EVENTO, monto, fecha, TAZA)
   
    campos(0, 0) = "banco"
    campos(1, 0) = "tipo"
    campos(2, 0) = "numero"
    campos(3, 0) = "empresa"
    campos(4, 0) = "evento"
    campos(5, 0) = "monto"
    campos(6, 0) = "fecha"
    campos(7, 0) = "taza"
    campos(8, 0) = ""
    
    
    campos(0, 1) = banco
    campos(1, 1) = tipo
    campos(2, 1) = numero
    campos(3, 1) = empresa
    campos(4, 1) = EVENTO
    campos(5, 1) = monto
    campos(6, 1) = fecha
    campos(7, 1) = Replace(TAZA, ",", ".")
    
    If MODIFI = "1" Then
    condicion = "banco='" + banco + "' and tipo='" + tipo + "' and numero='" + numero + "' and empresa='" + empresa + "' "
    op = 3
    
    Else
    op = 2
    
    condicion = ""
    End If
    campos(0, 2) = "inver_fondosmutuos"
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
MODIFI = "0"
leer

End Sub


Sub leer()

    campos(0, 0) = "banco"
    campos(1, 0) = "tipo"
    campos(2, 0) = "numero"
    campos(3, 0) = "empresa"
    campos(4, 0) = "evento"
    campos(5, 0) = "monto"
    campos(6, 0) = "fecha"
    campos(7, 0) = "taza"
    campos(8, 0) = ""
    
    
    
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero='" + dato3.text + "' and empresa='" + dato4.text + "' "
    campos(0, 2) = "inver_fondosmutuos"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
        
    If sqlconta.status = 0 Then
    MODIFI = "1"
    carga
    opciones.Visible = True
    disponible (True)
    opciones.SetFocus
    
    
    
    End If

End Sub
Sub carga()
    disponible (True)
    
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato3.text = sqlconta.response(2, 3)
    dato4.text = sqlconta.response(3, 3)
    DATO5.text = Mid(sqlconta.response(6, 3), 1, 2)
    dato6.text = Mid(sqlconta.response(6, 3), 4, 2)
    dato7.text = Mid(sqlconta.response(6, 3), 7, 4)
    TAZA.text = Format(sqlconta.response(7, 3), "%##,##0.000")
    monto.text = Format(sqlconta.response(5, 3), "###,###,###")
    
    Call leerbanco(dato1.text)
    Call leerdeposito(dato2.text)
    Call leerempresa(dato4.text)
    
End Sub



Sub ELIMINAR()
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero='" + dato3.text + "' and empresa='" + dato4.text + "' "
    campos(0, 2) = "inver_fondosmutuos"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
no:
End Sub




Private Sub MANUAL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
    If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
    If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
    If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
    If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub



Private Sub monto_KeyPress(KeyAscii As Integer)
Dim EVENTO As String

KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And monto.text <> "" Then
If Option1.Value = True Then
EVENTO = "1"
Else
EVENTO = "2"
End If

Call grabar(dato1.text, dato2.text, dato3.text, dato4.text, EVENTO, monto.text, dato7.text + "-" + dato6.text + "-" + DATO5.text, TAZA.text)
retorno
End If

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then retorno

If command = "modifica" Then
    MODIFI = "1"
    dato1.Enabled = True
    dato1.SetFocus
End If
If command = "elimina" Then
    
    ELIMINAR
    retorno
End If


End Sub


Sub retorno()


opciones.Visible = False
limpia
disponible (False)
MODIFI = ""
dato1.Enabled = True
dato1.SetFocus
leerfondos

End Sub


Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    monto.text = ""
    TAZA.text = ""
    lblBanco.Caption = ""
    LBLEMPRESA.Caption = ""
    LBLTIPO.Caption = ""
    
    
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus: caja.SelStart = 0
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub


Sub ayudacrcc(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "no"
    pivote2.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote2, campos, cfijo, largo, 2)
    
    pivote2.text = ""
End Sub

Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    monto.Enabled = condicion
    
End Sub



Private Sub opciones_GotFocus()
MANUAL.SetFocus

End Sub


Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 10)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "BANCO"
    FORMATOGRILLA(1, 2) = "EMPRESA"
    FORMATOGRILLA(1, 3) = "TIPO"
    FORMATOGRILLA(1, 4) = "NUMERO"
    FORMATOGRILLA(1, 5) = "FECHA"
    FORMATOGRILLA(1, 6) = "INVERSION"
    FORMATOGRILLA(1, 7) = "RETIRO"
    FORMATOGRILLA(1, 8) = "SALDO"
    FORMATOGRILLA(1, 9) = "TAZA"
    
     
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "20"
    FORMATOGRILLA(2, 2) = "20"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "D"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,###,##0"
    FORMATOGRILLA(4, 7) = "###,###,###,##0"
    FORMATOGRILLA(4, 8) = "###,###,###,##0"
    FORMATOGRILLA(4, 9) = "###,###,##0.000"
    
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    
    Grid1.Cols = 10
    Grid1.Rows = 2
    
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub
Private Sub leerfondos()

Dim resultados As rdoResultset
    
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim saldo As Double
    
    LINEA = 0
 saldo = 0
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT banco,empresa,tipo,numero,fecha,monto,evento,taza  "
        csql.sql = csql.sql + "FROM inver_fondosmutuos order by banco,empresa,fecha,evento"
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = csql.RowsAffected + 1
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(resultados(2), "yyyy") + "/" + Format(resultados(2), "mm") + "/" + Format(resultados(2), "dd")
        
         While Not resultados.EOF
          LINEA = LINEA + 1
             Grid1.Cell(LINEA, 1).text = leerbanco(resultados(0))
             Grid1.Cell(LINEA, 2).text = leerempresa(resultados(1))
             Grid1.Cell(LINEA, 3).text = leerdeposito(resultados(2))
             Grid1.Cell(LINEA, 4).text = resultados(3)
             Grid1.Cell(LINEA, 5).text = resultados(4)
             If resultados(6) = "1" Then
             Grid1.Cell(LINEA, 6).text = resultados(5)
             Grid1.Cell(LINEA, 7).text = "0"
             Else
             Grid1.Cell(LINEA, 7).text = resultados(5)
             Grid1.Cell(LINEA, 6).text = "0"
             End If
             
             Grid1.Cell(LINEA, 8).text = saldo + CDbl(Grid1.Cell(LINEA, 6).text) - CDbl(Grid1.Cell(LINEA, 7).text)
             Grid1.Cell(LINEA, 9).text = resultados(7)
             If LINEA > 0 Then
             saldo = CDbl(Grid1.Cell(LINEA, 8).text)
             End If
             
             
             
             resultados.MoveNext
          If resultados.EOF = False Then
       
          End If
   
                   Wend
End If
 
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
