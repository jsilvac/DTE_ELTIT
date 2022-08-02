VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form maestroparametrosfinancieros 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros Financieros Local"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10110
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   655
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   7695
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   29
         Top             =   6360
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   27
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   22
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   20
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox dato5 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos de Cobranzas Mas de"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos de Cobranzas Mas de"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos de Cobranzas Mas de"
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Parametros Morosidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Descento Autorizado No.3"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Descento Autorizado No.2"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Parametros Ventas Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Descento Autorizado No.1"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento Gerencia"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   7695
         Left            =   0
         Top             =   0
         Width           =   8895
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "3 Cuotas Precio Contado"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximo de Cuotas"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Interes Mensual en Ventas"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion Parametros Financieros"
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
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Interes Mensual Morosos"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerancia Cupo Autorizado"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximo de Dias sin Costo"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   4920
         Width           =   1935
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   720
      TabIndex        =   30
      Top             =   8280
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
      Height          =   7695
      Left            =   720
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "maestroparametrosfinancieros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
dato1.SetFocus

End Sub

Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 674
    Me.ScaleHeight = 655
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
