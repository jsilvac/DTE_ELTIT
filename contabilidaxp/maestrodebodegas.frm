VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form maestro07 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Bodegas"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   427
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc mb 
      Height          =   330
      Left            =   6960
      Top             =   6000
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3975
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   7215
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   5
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   4
         Top             =   2400
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   3
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   2
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bodega / SV"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3975
         Left            =   0
         Top             =   0
         Width           =   7215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion de Bodegas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Otros"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   600
      TabIndex        =   14
      Top             =   4800
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\remuxp\barra_opciones.swf"
      Src             =   "c:\remuxp\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   3975
      Left            =   600
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "maestro07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dato_GotFocus(Index As Integer)
'   CAMBIA EL COLOR DE FONDO DEL LOS CUADRO DE
'   TEXTO CADA VEZ QUE OBTENGAN O PIERDAN EL FOCO.
'----------------------------------------------------
    
        dato(Index).BackColor = &HFFFF&
        dato(Index).SelStart = 0
        dato(Index).SelLength = dato(Index).MaxLength
End Sub

Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 672
    Me.ScaleHeight = 577
    'CARGA LA BARRA DE TITULO
    Rem swfBarra.Width = Me.ScaleWidth
    Rem swfBarra.Height = Me.ScaleHeight
    Rem Call swfBarra.LoadMovie(0, Interfaces.path + "\Data\Barra_Titulo.swf")
    'CARGA EL BOTON NUEVO
    Rem Call swfNuevo.LoadMovie(0, Interfaces.path + "\Data\btn_nuevo.swf")
    'OBTENER POSICION DEL FORMULARIO
    posx2 = Me.ScaleWidth
    posy2 = Me.ScaleHeight
    posx1 = (Interfaces.equiAncho(Screen.Width) - posx2) \ 2
    posy1 = (Interfaces.equiAlto(Screen.Height) - posy2) \ 2
    'CARGADO DE LA IMAGEN DEGRADADA
    apis.Degradado Me, Principal, posx1, posx2, posy1, posy2, 150
       
    'FLAG = 0 SE GRABA/MODIFICA  FLAG = 1 YA SE GUARDO EN BD
    
    Call Conecta_Maestro_Bodegas
    Call Conectar_BD

End Sub

