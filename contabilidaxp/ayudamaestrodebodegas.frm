VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ayudamaestrodebodegas 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda Maestro de Bodegas"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   360
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.TextBox buscaproductos 
         BackColor       =   &H00D8E1FC&
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   5160
         Width           =   4095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "ayudamaestrodebodegas.frx":0000
         Height          =   4335
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   15380136
         ForeColor       =   128
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "codigobodega"
            Caption         =   "  CODIGO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "bodega"
            Caption         =   "                           NOMBRE BODEGA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   5595,024
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   5895
         Left            =   0
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Maestro de Bodega"
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
         TabIndex        =   3
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar por Nombre de Bodega"
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
         TabIndex        =   2
         Top             =   5280
         Width           =   2655
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5895
      Left            =   360
      Top             =   480
      Width           =   7575
   End
End
Attribute VB_Name = "ayudamaestrodebodegas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buscaproductos_Change()
With ayudamaestrodebodegas
         .Adodc1.ConnectionString = "driver={MySQL ODBC 3.51 Driver};" & _
         "server=localhost;uid=root;pwd=;database=conta01"
        .Adodc1.RecordSource = "SELECT codigobodega,bodega FROM maestrobodegas WHERE bodega >= '" & buscaproductos.text & "' ORDER BY bodega"
        .Adodc1.Refresh
    End With
    
End Sub

Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 1024
    Me.ScaleHeight = 768
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


Call listatrabajadores

End Sub

Sub listatrabajadores()

'******************************************************************************
'Inicio Codigo para conexion a base de datos hacia Data Grid
'******************************************************************************


    With ayudamaestrodebodegas
        .Adodc1.ConnectionString = "driver={MySQL ODBC 3.51 Driver};" & _
        "server=localhost;uid=root;pwd=;database=conta01"
        .Adodc1.RecordSource = "SELECT codigobodega,bodega FROM maestrobodegas WHERE bodega >= '" & buscaproductos.text & "' ORDER BY bodega"
        .Adodc1.Refresh
    End With
    
'******************************************************************************
'Fin Codigo para conexion a base de datos hacia Data Grid
'******************************************************************************
   
End Sub


