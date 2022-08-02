VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form DespachoHarina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DESPACHO HARINA"
   ClientHeight    =   7200
   ClientLeft      =   3555
   ClientTop       =   3375
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9255
   Begin VB.CommandButton CmdListado 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IMPRIMIR LISTADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4800
      Width           =   1575
   End
   Begin XPFrame.FrameXp FrameDespachos 
      Height          =   1455
      Left            =   960
      TabIndex        =   26
      Top             =   5520
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2566
      BackColor       =   16761024
      Caption         =   "DATOS DEL DESPACHO"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "REIMPRIMIR VALE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox DespachoAutoriza 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   31
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox DespachoHora 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox DespachoFecha 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label DG 
         Caption         =   "Label5"
         Height          =   375
         Left            =   8160
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AUTORIZADO POR"
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
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HORA DESPACHO"
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
         Left            =   3840
         TabIndex        =   30
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA DESPACHO"
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
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
   End
   Begin XPFrame.FrameXp FrameDespacho 
      Height          =   2295
      Left            =   1560
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4048
      BackColor       =   16761024
      Caption         =   "AUTORIZAR DESPACHO"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox AutorizaRutVerificador 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton CmdAutoriza 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ACEPTA&R"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox AutorizaRut 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   20
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label AutorizaNombre 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp frmdatos 
      Height          =   7215
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12726
      BackColor       =   16744576
      Caption         =   "DATOS"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "VISTA PREVIA DE IMPRESION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4800
         Width           =   1575
      End
      Begin FlexCell.Grid Grid1 
         Height          =   495
         Left            =   1080
         TabIndex        =   34
         Top             =   8280
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CommandButton CmdDespachar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&DESPACHAR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox despachonumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox dato2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox dato9 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   5655
      End
      Begin VB.CommandButton CmdRetorno 
         BackColor       =   &H00FFC0C0&
         Caption         =   "RETORNO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4800
         Width           =   1815
      End
      Begin FlexCell.Grid productos 
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5530
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.TextBox dato8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox dato7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox dato6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox dato3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3480
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbldv 
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
         Left            =   2880
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CLIENTE"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA"
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
         Left            =   3480
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CAJA"
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
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nº DOCUMENTO"
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
         Left            =   5160
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "DespachoHarina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private c As vendedor
Dim fecha As String

 Private Sub CargaGrillaProductos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Dim formatogrilla(10, 10) As String
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CANTIDAD"
        formatogrilla(1, 2) = "CODIGO"
        formatogrilla(1, 3) = "DESCRIPCION"
        formatogrilla(1, 4) = "PRECIO"
        formatogrilla(1, 5) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "2"
        formatogrilla(2, 2) = "13"
        formatogrilla(2, 3) = "80"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "###0.000"
        formatogrilla(4, 2) = "0000000000000"
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = "$ ###,###,##0"
        formatogrilla(4, 5) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
 
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
 
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
    
        
        Rem ANCHO
        formatogrilla(8, 1) = "7"
        formatogrilla(8, 2) = "11"
        formatogrilla(8, 3) = "14"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "9"
            
        productos.Cols = col
        productos.Rows = row
        productos.AllowUserResizing = False
        productos.DisplayFocusRect = False
        productos.ExtendLastCol = True
        productos.BoldFixedCell = False
        productos.DrawMode = cellOwnerDraw
        productos.Appearance = Flat
        productos.ScrollBarStyle = Flat
        productos.FixedRowColStyle = Flat
        productos.BackColorFixed = RGB(90, 158, 214)
        productos.BackColorFixedSel = RGB(110, 180, 230)
        productos.BackColorBkg = RGB(90, 158, 214)
        productos.BackColorScrollBar = RGB(231, 235, 247)
        productos.BackColor1 = RGB(231, 235, 247)
        productos.BackColor2 = RGB(239, 243, 255)
        productos.GridColor = RGB(148, 190, 231)
     
        productos.Column(0).Width = 0
        For i = 1 To col - 1
            productos.Cell(0, i).text = formatogrilla(1, i)
            'productos.Column(i).Width = Val(formatoGrilla(8, i)) * (productos.Cell(0, i).Font.Size + 1.25): K = formatoGrilla.FontSize(148, 190, 231, 111)
            productos.Column(i).MaxLength = Val(formatogrilla(2, i))
            productos.Column(i).FormatString = formatogrilla(4, i)
            productos.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                productos.Column(i).Alignment = cellRightCenter
            Else
                productos.Column(i).Alignment = cellLeftCenter
            End If
            
        Next i
        productos.Column(3).Width = "200"
        productos.Cell(0, 1).Alignment = cellCenterCenter
        productos.Cell(0, 2).Alignment = cellCenterCenter
        productos.Cell(0, 3).Alignment = cellCenterCenter
        productos.Cell(0, 4).Alignment = cellCenterCenter
        productos.Cell(0, 5).Alignment = cellCenterCenter

        'Productos.Enabled = True
    End Sub
 



Private Sub AutorizaNombre_Change()
If AutorizaRut <> Empty And AutorizaNombre <> Empty Then
CmdAutoriza.Enabled = True
CmdAutoriza.SetFocus
End If
End Sub

Private Sub AutorizaRut_Change()
AutorizaRutVerificador = Empty
AutorizaNombre = Empty
End Sub

Private Sub AutorizaRut_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
        AutorizaRut = Empty
        FrameDespacho.Visible = False
        CmdDespachar.SetFocus
    End If
KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And AutorizaRut.text <> Empty Then
        
        AutorizaRut = ceros(AutorizaRut)
        AutorizaRutVerificador = rut(AutorizaRut)
        If leerVendedor(c, AutorizaRut.text & AutorizaRutVerificador.text, "=") = True Then
        'AutorizaNombre = leerNombreCliente(AutorizaRut & AutorizaRutVerificador)
        AutorizaNombre = c.nombre
        Else
        MsgBox "EL VENDEDOR NO EXISTE", vbCritical, "ATENCION"
        AutorizaRut.SelStart = 0
        AutorizaRut.SelLength = Len(AutorizaRut)
        AutorizaRut.SetFocus
        
     End If
     Else
     AutorizaNombre = Empty
     CmdAutoriza.Enabled = False
    End If
    
End Sub


Public Sub pausa(tiempo, estado, GLOSA)
Sleep (tiempo)
End Sub
Sub agregarclientedespacho(NUMERO, tipoDoc, numdoc, caja, fecha, rutcli)
    Dim csql As New rdoQuery
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "update sv_valesharinas set cliente='" & rutcli & "' where "
    csql.sql = csql.sql & "numero='" & NUMERO & "' and tipodocumento='" & tipoDoc & "' and numerodocumento='" & numdoc & "' and caja='" & caja & "' and fecha= '" & fecha & "' "
    csql.Execute
    csql.Close
    
End Sub
Private Sub dato1_KeyPress(KeyAscii As Integer)
 '   KeyAscii = esNumero(KeyAscii)
  '  If KeyAscii = 13 And dato1.text <> "" Then
   '     dato1.text = ceros(dato1)
    '    dato2.SetFocus
   ' End If
End Sub

Private Sub CmdAutoriza_Click()
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro

If MsgBox("Se va a autorizar el despacho de Harina" & vbCr & _
"Nº :" & despachonumero & "    Desea continuar??", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

    csql.sql = "update " & clientesistema & "ventas" & empresaActiva & ".sv_valesharinas "
    csql.sql = csql.sql + " set fechaentrega = '" & (Format(Date, "yyyy") & "-" & Format(Date, "mm") & "-" & Format(Date, "dd")) & " ' , horaentrega = '" & Time & "' , entregadopor = '" & AutorizaRut & "' "
    csql.sql = csql.sql & "where numero = '" & Mid(despachonumero, 3, 13) & "'"
    csql.Execute
    Call sincronizadatos(csql.sql, ventasRubro)
    ImprimeTicketAutoriza
    GoTo fin

Else
    AutorizaRut = Empty
    FrameDespacho.Visible = False
    DespachoAutoriza = Empty
    despachonumero.SetFocus
    Exit Sub
End If
    GoTo fin
Exit Sub
fin:
    despachonumero = Empty
    AutorizaRut = Empty
    CmdDespachar.Enabled = False
    FrameDespacho.Visible = False
    CmdAutoriza.Enabled = False
    CmdRetorno.SetFocus
    
End Sub

Private Sub CmdAutoriza_KeyPress(KeyAscii As Integer)
AutorizaRut = Empty
FrameDespacho.Visible = False
despachonumero.SetFocus
End Sub

Private Sub CmdDespachar_Click()
FrameDespacho.Visible = True
AutorizaRut.SetFocus
End Sub

Private Sub CmdDespachar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then despachonumero.SetFocus
End Sub

Private Sub CmdListado_Click()
  imprimirticket
End Sub

Private Sub CmdRetorno_Click()
If despachonumero = Empty Then
    Unload Me
Else
    despachonumero = Empty
    despachonumero.SetFocus
  
End If
End Sub
Private Sub Command1_Click()
'imprimirticket
ImprimeTicketAutoriza
End Sub

Private Sub Command2_Click()
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = ventasRubro

With Grid1
    .Cols = 7
    .Rows = 6
    .RowHeight(0) = 0
    .PageSetup.TopMargin = 3
    .PageSetup.LeftMargin = 1
        
    With .Range(1, 1, 1, 6)       'titulo del documento
      .Merge
      .Borders(cellEdgeLeft) = cellThick
      .Borders(cellEdgeRight) = cellThick
      .Borders(cellEdgeTop) = cellThick
      .Borders(cellEdgeBottom) = cellThick
      .FontBold = True
      .Alignment = cellCenterCenter
    End With
    With .Range(2, 1, 2, 6)       'titulo del documento
      .Merge
      .Alignment = cellCenterCenter
    End With
    With .Range(3, 1, 3, 6)       'titulo del documento
      .Merge
      .Alignment = cellCenterCenter
    End With
   
    With .Range(5, 1, 5, 6)       'titulo del documento
     .Borders(cellEdgeLeft) = cellThick
      .Borders(cellEdgeRight) = cellThick
      .Borders(cellEdgeTop) = cellThick
      .Borders(cellEdgeBottom) = cellThick
      .Borders(cellInsideVertical) = cellThick
      .Alignment = cellCenterCenter
    End With
        .Column(1).Width = "50"
        .Column(2).Width = "70"
        .Column(3).Width = "20"
        .Column(4).Width = "35"
        
    .Cell(1, 1).text = "LISTADO DE TICKETS HARINAS"
    .Cell(2, 1).text = Replace(leerNombreEmpresa(empresaActiva), " Ltda.", "") & " - Rut:" & Format(Mid(leerRutEmpresa(empresaActiva), 1, 10), "###,###,###") & "-" & Mid(leerRutEmpresa(empresaActiva), 10, 1)
    .Cell(3, 1).text = leerDireccionEmpresa(empresaActiva) & " - " & fechasistema
    
    .Cell(5, 1).text = "FOLIO"
    .Cell(5, 2).text = "FECHA"
    .Cell(5, 3).text = "TP"
    .Cell(5, 4).text = "CAJA"
    .Cell(5, 5).text = "CLIENTE"
    .Cell(5, 6).text = "ESTADO"
    
 
 
    csql.sql = "select numero,fecha,tipodocumento,caja,cliente,entregadopor from "
    csql.sql = csql.sql & clientesistema & "ventas" & empresaActiva & ".sv_valesharinas "
    csql.sql = csql.sql & "where fecha = '" & Format(Date, "yyyy-mm-dd") & "'"
    csql.Execute
    Dim row, col As Long
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        
            .Rows = .Rows + 1
            .Cell(.Rows - 1, 1).text = Mid(resultados(0), 5, 10)
            .Cell(.Rows - 1, 2).text = resultados(1)
            .Cell(.Rows - 1, 3).text = resultados(2)
            .Cell(.Rows - 1, 4).text = resultados(3)
            .Cell(.Rows - 1, 5).text = resultados(4)
          '  MsgBox resultados(5)
            If resultados(5) = Empty Then
                .Cell(.Rows - 1, 6).text = "PENDIENTE"
            Else
                .Cell(.Rows - 1, 6).text = "DESPACHADO"
            End If
        resultados.MoveNext
    Wend
   ' .Range(1, 6, .Rows - 1, 5).FontSize = 5
'    CmdDespachar.Enabled = True
'    CmdDespachar.SetFocus
End If
                
    .PrintPreview
    
   ' .PrintDialog
End With
'End If

End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 13 And dato2.text <> "" Then
    dato3.SetFocus
 End If
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        dato6.text = ceros(dato6)
        If dato6.text = "00" Then dato6.text = Format(fechasistema, "mm")
        dato7.SetFocus
    End If
End Sub
Private Sub dato8_Change()
lbldv.Caption = rut(dato8)
dato9.text = leerNombreCliente(dato8.text & lbldv.Caption)
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And dato8 <> Empty Then
    dato8.text = ceros(dato8)
    dato9.text = leerNombreCliente(dato8.text & lbldv.Caption)
End If
End Sub

Private Sub DespachoAutoriza_Change()
If DespachoFecha <> Empty Then
    If DespachoHora <> Empty Then
        If DespachoAutoriza <> Empty Then
            
            MsgBox "DESPACHO NUMERO : " & despachonumero & " YA AUTORIZADO", vbExclamation, "ATENCION"
            FrameDespachos.Visible = True
            FrameDespacho.Visible = False
            CmdDespachar.Enabled = False
            despachonumero.SelStart = 0
            despachonumero.SelLength = Len(despachonumero)
            despachonumero.SetFocus
        End If
    End If
Else
AutorizaNombre = Empty
End If
    
'    DG.Caption = rut(DespachoAutoriza)
    
End Sub

Private Sub DespachoAutoriza_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And DespachoAutoriza <> Empty Then
End Sub

Private Sub despachonumero_Change()
With Me
    .DespachoAutoriza = Empty
    .DespachoFecha = Empty
    .DespachoHora = Empty
    
    .dato2 = Empty
    .dato3 = Empty
    .dato4 = Empty
    .dato5 = Empty
    .dato6 = Empty
    .dato7 = Empty
    .dato8 = Empty
    .dato9 = Empty
.productos.Rows = 1
.FrameDespachos.Visible = False
End With

End Sub

Private Sub despachonumero_GotFocus()
despachonumero.SelStart = 0
despachonumero.SelLength = Len(despachonumero)
End Sub

Private Sub despachonumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then despachonumero = Empty
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And despachonumero.text <> Empty Then
        despachonumero = ceros(despachonumero)
        Call LeerDespachoHarina(Mid(despachonumero, 3, 12), empresaActiva)
        Call leerproductosharina(dato2, dato3, dato4.text)
        If dato8.text = Empty Then
            dato8.Locked = False
            MsgBox "INGRESE EL RUT DEL CLIENTE"
            dato8.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Call CargaGrillaProductos(1, 6)

End Sub

Sub leerproductosharina(TIPO, NUMERO, caja)
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro

csql.sql = "select dd.cantidad,dd.codigo,dd.descripcion,dd.precio,dd.total from sv_documento_detalle_" & empresaActiva & " as dd, "
csql.sql = csql.sql & "sv_documento_cabeza_" & empresaActiva & " as dc "
csql.sql = csql.sql & "where dc.numero=dd.numero and dc.local=dd.local and dc.caja=dd.caja and dc.fecha=dd.fecha and "
csql.sql = csql.sql & " dc.tipo=dd.tipo and dd.tipo='" & TIPO & "' and dc.foliosii='" & NUMERO & "' and dd.caja='" & caja & "'"
'cSql.sql = cSql.sql & "and dd.fecha='" & fecha & "' "
csql.Execute
productos.Rows = 1
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        If productomarcado(resultados(1)) = True Then
            productos.Rows = productos.Rows + 1
            productos.Cell(productos.Rows - 1, 1).text = resultados(0)
            productos.Cell(productos.Rows - 1, 2).text = resultados(1)
            productos.Cell(productos.Rows - 1, 3).text = resultados(2)
    'productos.Column(2).Width = Len(resultados(2))
            productos.Cell(productos.Rows - 1, 4).text = resultados(3)
            productos.Cell(productos.Rows - 1, 5).text = resultados(4)
        End If
        resultados.MoveNext
    Wend
    
'    CmdDespachar.Enabled = True
'    CmdDespachar.SetFocus
End If
End Sub
Sub limpia()
    'dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    lbldv.Caption = ""
    productos.Rows = 1
    fecha = ""
    'dato1.SetFocus
End Sub

    Function existecliente(rut) As Boolean
        Dim csql As New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "select rut from sv_maestroclientes where rut='" & rut & "' "
        csql.Execute
        existecliente = False
        If csql.RowsAffected > 0 Then
            existecliente = True
        End If
        csql.Close
        Set csql = Nothing
    
    End Function

Public Sub leerdeafuera()
    'fecha = dato7.text & "-" & dato6.text & "-" & dato5.text
    'Call leerproductosharina(dato2.text, dato3.text, dato4.text, fecha)
End Sub

Public Sub LeerDespachoHarina(NUMERO, loc)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select * from " & clientesistema & "ventas" & loc & ".sv_valesharinas "
    csql.sql = csql.sql & "where numero = '" & NUMERO & "'"
    csql.Execute
    'leerultimodespachoharina = "0000000001"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        With DespachoHarina
            .dato2 = resultados(4)
            .dato3 = resultados(5)
            .dato4 = resultados(6)
            .dato8 = Mid(resultados(2), 1, 10)
            .dato5 = Mid(resultados(3), 1, 3)
            .dato6 = Mid(resultados(3), 4, 5)
            .dato7 = Mid(resultados(3), 7, 9)
           If resultados(11) <> Empty Then
                .DespachoFecha = resultados(11)
           Else
                .DespachoFecha = ""
           End If
           If resultados(12) <> Empty Then
                .DespachoHora = resultados(12)
           Else
             .DespachoHora = ""
           End If
           If resultados(13) <> Empty Then
            .DespachoAutoriza = resultados(13)
           Else
            .DespachoAutoriza = ""
           '.FrameDespacho.Visible = True
            .CmdDespachar.Enabled = True
            .CmdDespachar.SetFocus
           End If
        End With
        
    Else
    MsgBox "DOCUMENTO NO EXISTE O VENTA SIN PRODUCTOS MARCADOS " & vbCrLf & _
    "                     VERIFIQUE LOS DATOS", vbCritical, "ATENCION"
    despachonumero.SelStart = 0
    despachonumero.SelLength = Len(despachonumero)
    End If
    csql.Close
    
End Sub
Sub imprimirticket()
    Dim numfic As Integer
    Dim K As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    numfic = 20
    Close numfic
    
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #numfic
    
    Print #numfic, Chr$(27); Chr$(64)
    Print #numfic, "DETALLE DE DESPACHO DE HARINAS"
    Print #numfic, leerNombreEmpresa(empresaActiva)
    Print #numfic, "RUT: " & Mid(leerRutEmpresa(empresaActiva), 1, 10) & "-" & Mid(leerRutEmpresa(empresaActiva), 10, 1)
    Print #numfic, fechasistema
    Print #numfic, " "
    Set csql.ActiveConnection = ventas
    'cSql.sql = "select numero,tipodocumento,numerodocumento,fecha,caja,cliente,codigo,cantidad,total,entregadopor,fechaentrega from "
    csql.sql = "select * from "
    csql.sql = csql.sql & clientesistema & "ventas" & empresaActiva & ".sv_valesharinas "
    csql.sql = csql.sql & "where fecha = '" & Format(Date, "yyyy-mm-dd") & "' and fechaentrega <> '0000-00-00'  order by horaentrega desc " 'between '2009-09-01' and '2009-09-30'" '
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
    Print #numfic, "________________________________________"
    Print #numfic, "TIPO:"; resultados(4); " Numero:"; Format(resultados(5), "########0"); " Fecha:"; resultados(3)
    Print #numfic, "CAJA:"; resultados(6); resultados(8); DescripcionProducto(resultados(7)); " Total "; Format(resultados(10), "$###,###,##0")
    Print #numfic, "CLIENTE :"; resultados(2); " "; leerNombreCliente(resultados(2))
    Print #numfic, "Despachado el: "; resultados(11); " "; "a las :" & resultados(12)
    
    'If resultados(9) = Empty Then
    '      Print #numfic, "PENDIENTE"
    'Else
     Print #numfic, "DESPACHADO POR: "; leerNombreVendedor(resultados(13) & rut(resultados(13)))
        
    'End If
    resultados.MoveNext
    Wend
    End If
    Print #numfic, "_________________________________ "
    Print #numfic, "DESPACHOS AUTORIZADOS DEL DIA:"; Despachosharinas(Format(Date, "yyyy-mm-dd"), Format(Date, "yyyy-mm-dd"), 1)
    Print #numfic, "DESPACHOS PENDIENTES DEL DIA :"; Despachosharinas(Format(Date, "yyyy-mm-dd"), "0000-00-00", 1)
    Print #numfic, "DESPACHOS AUTORIZADOS OTRA FECHA :"; Despachosharinas(Format(Date, "yyyy-mm-dd"), "0000-00-00", 3)
    Print #numfic, "TOTAL DESPACHOS AUTORIZADOS  :"; Despachosharinas(Date, Format(fechasistema, "yyyy-mm-dd"), 2)
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, "_________________________________ "
    Print #numfic, "FIRMA GUARDIA RESPONSABLE "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, Chr(27); "i"
    '''''''''''''''''''''
    'PRE-VENTA
    '''''''''''''''''''''
    Close #numfic
    'If Impresora = 0 Then Shell "notepad impresion.txt"
 
 End Sub
Sub ImprimeTicketAutoriza()
    Dim numfic As Integer
    Dim K As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    numfic = 20
    Close numfic
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #numfic
    Print #numfic, Chr$(27); Chr$(64)
    Print #numfic, "COMPROBANTE DE DESPACHO DE HARINAS"
    Print #numfic, leerNombreEmpresa(empresaActiva)
    Print #numfic, "RUT: " & Mid(leerRutEmpresa(empresaActiva), 1, 10) & "-" & Mid(leerRutEmpresa(empresaActiva), 10, 1)
    Print #numfic, leerDireccionEmpresa(empresaActiva) ' fechasistema
    Print #numfic, fechasistema & " " & Time
    Set csql.ActiveConnection = ventas
    
    csql.sql = "select * from "
    csql.sql = csql.sql & clientesistema & "ventas" & empresaActiva & ".sv_valesharinas "
    csql.sql = csql.sql & "where numero = '" & Mid(despachonumero, 3, 13) & "' and  fecha = '" & dato7 & "-" & dato6 & "-" & dato5 & "'"
    'if productos.Rows = 2 then csql.sql = csql.sql & "and linea = '1' limit 1"
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
    Print #numfic, "___________________________________"
    Print #numfic, " Numero :"; Mid(resultados(0), 6, 10); " TIPO :"; resultados(4); " FOLIO :"; resultados(5)
    Print #numfic, " FECHA :"; resultados(3); " CAJA :"; resultados(6); " TOTAL "; Format(resultados(10), "$###,###,###")
    Print #numfic, "Cant. :"; resultados(8); DescripcionProducto(resultados(7))
    Print #numfic, " CLIENTE :"; resultados(2); " "; leerNombreCliente(resultados(2))
    Print #numfic, " DESPACHADO POR: "; leerNombreVendedor(resultados(13) & rut(resultados(13)))
    resultados.MoveNext
    Wend
    End If
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, "_________________________________ "
    Print #numfic, "    FIRMA GUARDIA RESPONSABLE     "
    Print #numfic, " "
    Print #numfic, "**Este Comprobante debe ir adjunto**"
    Print #numfic, "        **al ticket original**      "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, " "
    Print #numfic, Chr(27); "i"
    '''''''''''''''''''''
    'PRE-VENTA
    '''''''''''''''''''''
    Close #numfic
    'If Impresora = 0 Then Shell "notepad impresion.txt"
 
 End Sub

Function DescripcionProducto(cod) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = gestionRubro
csql.sql = "select descripcion from r_maestroproductos_fijo_" & empresaActiva
csql.sql = csql.sql & " where codigobarra = '" & cod & "' limit 1"
csql.Execute
    
If csql.RowsAffected > 0 Then
   Set resultados = csql.OpenResultset
    DescripcionProducto = resultados(0)
End If
End Function
