VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form activos_listado 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Activos"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15000
   Tag             =   "maestro_arrendatarios"
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   8055
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14208
      BackColor       =   16761024
      Caption         =   "LISTADO DE ACTIVOS"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin FlexCell.Grid grid1 
         Height          =   7755
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13679
         BackColor1      =   12648384
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   375
      Left            =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   2540
      BackColor       =   16761024
      Caption         =   "Listado de Acivos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BordeColor      =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "codigopropiedad"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "codigopropiedad"
         Top             =   720
         Width           =   855
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   975
         Left            =   6600
         TabIndex        =   7
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1720
         BackColor       =   16761024
         Caption         =   "FILTRAR PORESTADO  :"
         CaptionEstilo3D =   2
         BackColor       =   16761024
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "SOLO ACTIVOS DADOS DE BAJA"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.TextBox dato0 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "codigopropiedad"
         Top             =   360
         Width           =   1335
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   960
         Left            =   -5760
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   1693
         BackColor       =   12648384
         Caption         =   "TIPOS CLIENTES"
         CaptionEstilo3D =   1
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox Combotipos 
            Height          =   315
            Left            =   45
            TabIndex        =   2
            Text            =   "Combo1"
            Top             =   315
            Width           =   4875
         End
      End
      Begin XPFrame.FrameXp FrmOpciones 
         Height          =   1140
         Left            =   12000
         TabIndex        =   16
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2011
         BackColor       =   16761024
         Caption         =   "OPCIONES"
         CaptionEstilo3D =   2
         BackColor       =   16761024
         ForeColor       =   8438015
         BordeColor      =   192
         ColorBarraArriba=   255
         ColorBarraAbajo =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         ColorTextShadow =   192
         Begin Contabilidadxp.BotonMyERP opcion 
            Height          =   855
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
            Caption         =   "Generar"
            PicturePosition =   0
            Picture         =   "activos_listado.frx":0000
            PictureHover    =   "activos_listado.frx":0CB6
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16761024
         End
         Begin Contabilidadxp.BotonMyERP opcion 
            Height          =   855
            Index           =   4
            Left            =   960
            TabIndex        =   18
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
            Caption         =   "Retorno"
            PicturePosition =   0
            Picture         =   "activos_listado.frx":1A17
            PictureHover    =   "activos_listado.frx":2740
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16761024
         End
         Begin Contabilidadxp.BotonMyERP opcion 
            Height          =   855
            Index           =   2
            Left            =   1800
            TabIndex        =   19
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
            Caption         =   "Imprimir"
            PicturePosition =   0
            Picture         =   "activos_listado.frx":34E4
            PictureHover    =   "activos_listado.frx":4210
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3360
         TabIndex        =   15
         Top             =   1080
         Width           =   3090
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ubicacion"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1095
         Width           =   1770
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3360
         TabIndex        =   12
         Top             =   720
         Width           =   3090
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo Activo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label lblnombreactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   3090
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo Activo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   735
         Width           =   1770
      End
   End
End
Attribute VB_Name = "activos_listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 12)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "DESCRIPCION"
    FORMATOGRILLA(1, 3) = "TIPO"
    FORMATOGRILLA(1, 4) = ""
    FORMATOGRILLA(1, 5) = ""
    FORMATOGRILLA(1, 6) = ""
    FORMATOGRILLA(1, 7) = ""
    FORMATOGRILLA(1, 8) = ""
    FORMATOGRILLA(1, 9) = ""
    FORMATOGRILLA(1, 10) = ""
    FORMATOGRILLA(1, 11) = ""
    
     
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "35"
    FORMATOGRILLA(2, 3) = "20"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "5"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "15"
    FORMATOGRILLA(2, 9) = "15"
    FORMATOGRILLA(2, 10) = "15"
    FORMATOGRILLA(2, 11) = "15"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "D"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
   
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = ""
    
    
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
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    
    Grid1.Cols = 4
    Grid1.Rows = 1
    
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

Private Sub Form_Load()
Call CARGAGRILLA
Call CENTRAR(Me)
GenerarListado
End Sub

Sub GenerarListado()
Dim csql As New rdoQuery
Dim resultados  As rdoResultset
'Dim FECHACONSULTA As String
'FECHACONSULTA = Format(DateAdd("D", dias, fecharecepcion), "yyyy-mm-dd")



Set csql.ActiveConnection = conta

csql.sql = "select * from af_maestro_activos  "

csql.Execute

 
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid1.AutoRedraw = True
    Grid1.AddItem "", True
    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados("codigobarras")
    Grid1.Cell(Grid1.Rows - 1, 2).text = resultados("descripcion")
    Grid1.Cell(Grid1.Rows - 1, 3).text = resultados("tipo") & " " & LeerNombreActivos_tipo(resultados("tipo"))
    
    
    resultados.MoveNext
    
    Wend
    
End If
Grid1.AutoRedraw = False
Grid1.Refresh


End Sub
