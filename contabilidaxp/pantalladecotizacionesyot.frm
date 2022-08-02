VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ventas09 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   695
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5415
      Left            =   720
      TabIndex        =   34
      Top             =   3120
      Width           =   10575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2535
         Left            =   240
         TabIndex        =   46
         Top             =   840
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4471
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFF2F7&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   10215
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00D8E1FC&
            Height          =   285
            Left            =   7200
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00D8E1FC&
            Height          =   285
            Left            =   6120
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00D8E1FC&
            Height          =   285
            Left            =   0
            TabIndex        =   41
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8640
            TabIndex        =   45
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label dato8 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PUERTO USB DE 2.0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1200
            TabIndex        =   44
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   8640
            TabIndex        =   40
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "UNIDADES"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6120
            TabIndex        =   39
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CODIGO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DESCRIPCION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            TabIndex        =   37
            Top             =   0
            Width           =   4935
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PRECIO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7200
            TabIndex        =   36
            Top             =   0
            Width           =   1455
         End
      End
      Begin MSAdodcLib.Adodc vmp 
         Height          =   330
         Left            =   240
         Top             =   3240
         Visible         =   0   'False
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
         Connect         =   "DSN=maestroproductos1"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "maestroproductos1"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT codigoproducto,descripcion,pventadetalle,stockcritico FROM maestroproductos ORDER BY descripcion"
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   5415
         Left            =   120
         Top             =   0
         Width           =   10575
      End
   End
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   13
         Left            =   4200
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   12
         Left            =   9480
         TabIndex        =   15
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   11
         Left            =   9120
         TabIndex        =   14
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   10
         Left            =   8760
         TabIndex        =   13
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   7
         Left            =   6600
         TabIndex        =   12
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   5040
         MaxLength       =   9
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   0
         Left            =   6600
         TabIndex        =   9
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   7
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   3
         Left            =   6600
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   4
         Left            =   6960
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   5
         Left            =   7320
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtventas 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   14
         Left            =   1320
         TabIndex        =   1
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "N/P"
         Height          =   255
         Left            =   3720
         TabIndex        =   33
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
         Height          =   255
         Left            =   7080
         TabIndex        =   31
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
         Height          =   255
         Left            =   7680
         TabIndex        =   30
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Condiciones de Pago"
         Height          =   495
         Left            =   5520
         TabIndex        =   29
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   4320
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "RUT"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   5520
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2535
         Left            =   0
         Top             =   0
         Width           =   11655
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Comuna"
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2640
         X2              =   2760
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5415
      Left            =   840
      Top             =   3240
      Width           =   10575
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   2535
      Left            =   360
      Top             =   360
      Width           =   11655
   End
End
Attribute VB_Name = "ventas09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 820
    Me.ScaleHeight = 697
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

