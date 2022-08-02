VERSION 5.00
Begin VB.Form ventas10 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantalla de Depositos"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2655
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Depositos"
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
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   2655
      Left            =   1080
      Top             =   840
      Width           =   6495
   End
End
Attribute VB_Name = "ventas10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 595
    Me.ScaleHeight = 297
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
