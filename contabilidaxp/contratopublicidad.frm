VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form contratopublicidad 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contrato Publicidad"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   4215
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   7455
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         Format          =   "dd/mm/aaaa"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00D8E1FC&
         Height          =   315
         ItemData        =   "contratopublicidad.frx":0000
         Left            =   1680
         List            =   "contratopublicidad.frx":000D
         TabIndex        =   12
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00D8E1FC&
         Height          =   315
         ItemData        =   "contratopublicidad.frx":003A
         Left            =   1680
         List            =   "contratopublicidad.frx":004D
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion de Contrato Publicidad"
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
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Facturar"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   4215
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "En Base a"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde "
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   4215
      Left            =   720
      Top             =   600
      Width           =   7455
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   720
      TabIndex        =   10
      Top             =   5280
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
End
Attribute VB_Name = "contratopublicidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 786
    Me.ScaleHeight = 593
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


End Sub
