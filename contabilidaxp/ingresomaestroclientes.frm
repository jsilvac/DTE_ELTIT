VERSION 5.00
Begin VB.Form ventas01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10860
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   682
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   724
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2655
      Left            =   600
      TabIndex        =   31
      Top             =   6480
      Width           =   9615
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1440
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   58
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label40 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   57
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   56
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   55
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   54
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   53
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   52
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   51
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   50
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   49
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Boletas"
         Height          =   255
         Left            =   3600
         TabIndex        =   47
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Prorrogas"
         Height          =   255
         Left            =   3600
         TabIndex        =   46
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   45
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   44
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   43
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Boletas"
         Height          =   255
         Left            =   3600
         TabIndex        =   39
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label21 
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
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Cupo"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Protestados"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   600
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   9615
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Prorrogas"
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Usado"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5775
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   4800
         Width           =   5055
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   4440
         Width           =   5055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox dato5 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3960
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato6 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox dato9 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox dato8 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox dato7 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox dato10 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox dato11 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox dato14 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox dato13 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   3720
         Width           =   5055
      End
      Begin VB.TextBox dato12 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   3360
         Width           =   5055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Dscuento"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Credito"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   2880
         X2              =   3000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label1 
         Caption         =   "Ingreso Maestro de Clientes"
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
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RUT"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   5775
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Comuna"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefonos"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Giro"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   1335
      End
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   2655
      Left            =   720
      Top             =   6600
      Width           =   9615
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5775
      Left            =   1800
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "Ventas01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 724
    Me.ScaleHeight = 682
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

