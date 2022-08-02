VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ListadoComprasPorClientes 
   BackColor       =   &H00FF8080&
   Caption         =   "Listado Compras Por Clientes"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   14340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox SUCU 
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   6315
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   11139
      BackColor       =   12582912
      Caption         =   "CUOTAS PENDIENTES"
      CaptionEstilo3D =   1
      BackColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdimprimir 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5520
         Width           =   2085
      End
      Begin VB.CommandButton RETORNAR 
         BackColor       =   &H00FF8080&
         Caption         =   "RETORNO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5520
         Width           =   2085
      End
      Begin FlexCell.Grid Grid1 
         Height          =   5055
         Left            =   120
         TabIndex        =   3
         Top             =   225
         Width           =   14145
         _ExtentX        =   24950
         _ExtentY        =   8916
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
      Begin VB.Label lblayuda 
         BackStyle       =   0  'Transparent
         Caption         =   $"ComprasPorClientes.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6720
         TabIndex        =   36
         Top             =   5400
         Width           =   4815
      End
      Begin VB.Label lbl11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   11
         Left            =   9765
         TabIndex        =   11
         Top             =   4365
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   9765
         TabIndex        =   10
         Top             =   4725
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   7425
         TabIndex        =   9
         Top             =   4725
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Interes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   2
         Left            =   7425
         TabIndex        =   8
         Top             =   4365
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4860
         TabIndex        =   7
         Top             =   4725
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Cuotas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   4815
         TabIndex        =   6
         Top             =   4365
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   12
         Left            =   12015
         TabIndex        =   5
         Top             =   4365
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   12015
         TabIndex        =   4
         Top             =   4725
         Visible         =   0   'False
         Width           =   2130
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1815
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3201
      BackColor       =   16761024
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
      Begin VB.CommandButton cmdconsultar 
         BackColor       =   &H00FF8080&
         Caption         =   "Generar Informe"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox rut2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   0
         Top             =   360
         Width           =   1635
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   900
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   1588
         BackColor       =   16761024
         Caption         =   "Fecha Consultar y/o numero documento"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.TextBox NUMERO 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   33
            Tag             =   "fecha"
            Top             =   570
            Width           =   1575
         End
         Begin VB.TextBox HASTA1 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1470
            MaxLength       =   2
            TabIndex        =   29
            Tag             =   "fecha"
            Top             =   570
            Width           =   375
         End
         Begin VB.TextBox HASTA2 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   28
            Tag             =   "fecha"
            Top             =   570
            Width           =   375
         End
         Begin VB.TextBox HASTA3 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2190
            MaxLength       =   4
            TabIndex        =   27
            Tag             =   "fecha"
            Top             =   570
            Width           =   615
         End
         Begin VB.TextBox DESDE1 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   75
            MaxLength       =   2
            TabIndex        =   26
            Tag             =   "fecha"
            Top             =   570
            Width           =   375
         End
         Begin VB.TextBox DESDE2 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   435
            MaxLength       =   2
            TabIndex        =   25
            Tag             =   "fecha"
            Top             =   570
            Width           =   375
         End
         Begin VB.TextBox DESDE3 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   795
            MaxLength       =   4
            TabIndex        =   24
            Tag             =   "fecha"
            Top             =   570
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NUMERO"
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
            TabIndex        =   34
            Top             =   330
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DESDE"
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
            Left            =   75
            TabIndex        =   31
            Top             =   330
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HASTA"
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
            Left            =   1470
            TabIndex        =   30
            Top             =   330
            Width           =   1335
         End
      End
      Begin VB.Label lbldv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3360
         TabIndex        =   22
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Rut Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   30
         TabIndex        =   21
         Top             =   390
         Width           =   1680
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3720
         TabIndex        =   20
         Top             =   360
         Width           =   8580
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Autorizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   6960
         TabIndex        =   19
         Top             =   945
         Width           =   2325
      End
      Begin VB.Label lblCupo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   6960
         TabIndex        =   18
         Top             =   1305
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Utilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   4
         Left            =   9240
         TabIndex        =   17
         Top             =   945
         Width           =   2310
      End
      Begin VB.Label lblUtilizado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   9240
         TabIndex        =   16
         Top             =   1305
         Width           =   2310
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   11520
         TabIndex        =   15
         Top             =   945
         Width           =   2460
      End
      Begin VB.Label lblDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   11520
         TabIndex        =   14
         Top             =   1305
         Width           =   2460
      End
   End
End
Attribute VB_Name = "ListadoComprasPorClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim fechadesde As String
 Dim fechahasta As String
 Dim tazainteresmora As Double
 Dim totalusado As Double
 Dim moratotal As Double
 Dim fechacom As String
 

Private Sub cmdconsultar_Click()
    Call LEErcuotas(rut2.text + lblDV.Caption, fechadesde, fechahasta)
        Grid1.Rows = Grid1.Rows + 4
        Grid1.Column(2).Locked = False
        Grid1.Column(3).Locked = False
        Grid1.Column(4).Locked = False
        Grid1.Column(5).Locked = False
        Grid1.Column(6).Locked = False
        Grid1.Column(7).Locked = False
        Grid1.Column(8).Locked = False
        Grid1.Column(9).Locked = False
        Grid1.Column(10).Locked = False
        
        Grid1.Range(Grid1.Rows - 1, 2, Grid1.Rows - 1, 10).Merge
        Grid1.Cell(Grid1.Rows - 1, 2).text = "____________________"
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 2, Grid1.Rows - 1, 10).Merge
        Grid1.Cell(Grid1.Rows - 1, 2).text = " PROMOTORA PALGUIN"
End Sub

Private Sub cmdImprimir_Click()
If Grid1.Rows > 1 Then
    Call imprimir
End If
End Sub
   Private Sub imprimir()
        Dim i As Long
        
        Call Titulos("LISTADO DE COMPRAS CLIENTES")
        
        Grid1.AutoRedraw = False
        Grid1.Range(1, 1, 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
        Grid1.PageSetup.HeaderMargin = 0.5
    
        Grid1.PageSetup.TopMargin = 1
        Grid1.PageSetup.LeftMargin = 0.5
        Grid1.PageSetup.RightMargin = 0.5
        Grid1.PageSetup.BottomMargin = 2
        Grid1.PageSetup.FooterMargin = 1
        Grid1.PageSetup.BlackAndWhite = True
        Grid1.PageSetup.Orientation = cellLandscape
        Grid1.PageSetup.PrintFixedRow = True


        
        
        Call verificaImpresora(5, Grid1)
        
        Grid1.AutoRedraw = True
    End Sub
    Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellLandscape
    Grid1.PageSetup.PrintTitleRows = 1
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid1.PageSetup.HeaderAlignment = cellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    Grid1.PageSetup.HeaderFont.Italic = True
    
    'TITULOS DEL REPORTE
  
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CLIENTE : " & rut2.text & "-" & lblDV.Caption & "  " & lblNombre.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & " usuario:" + usuarioSistema
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7

    
End Sub

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        desde1.text = ceros(desde1)
        If desde1.text = "00" Then desde1.text = Format(fechasistema, "dd")
        desde2.SetFocus
    End If
End Sub
Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        desde2.text = ceros(desde2)
        If desde2.text = "00" Then desde2.text = Format(fechasistema, "mm")
        desde3.SetFocus
    End If
End Sub
Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        desde3.text = ceros(desde3)
        If desde3.text = "0000" Then desde3.text = Format(fechasistema, "yyyy")
        fechadesde = desde3.text & "-" & desde2.text & "-" & desde1.text
        If IsDate(fechadesde) = True Then
            hasta1.SetFocus
        Else
            MsgBox "FECHA INVALIDA", vbCritical, "ATENCION"
            desde1.text = ""
            desde2.text = ""
            desde3.text = ""
            fechadesde = ""
            desde1.SetFocus
        End If
    End If
End Sub
Private Sub HASTA1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        hasta1.text = ceros(hasta1)
        If hasta1.text = "00" Then hasta1.text = Format(fechasistema, "dd")
        hasta2.SetFocus
    End If
End Sub
Private Sub HASTA2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        hasta2.text = ceros(hasta2)
        If hasta2.text = "00" Then hasta2.text = Format(fechasistema, "mm")
        hasta3.SetFocus
    End If
End Sub
Private Sub HASTA3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        hasta3.text = ceros(hasta3)
        If hasta3.text = "0000" Then hasta3.text = Format(fechasistema, "yyyy")
         fechahasta = hasta3.text & "-" & hasta2.text & "-" & hasta1.text
        If IsDate(fechahasta) = True Then
           cmdconsultar.SetFocus
        Else
            MsgBox "FECHA INVALIDA", vbCritical, "ATENCION"
            hasta1.text = ""
            hasta2.text = ""
            hasta3.text = ""
            fechahasta = ""
            hasta1.SetFocus
        End If
       
    End If
End Sub
Private Sub Form_Load()
Call CARGAGRILLA
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
 
 



Private Sub NUMERO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
NUMERO.text = ceros(NUMERO)
End If
End Sub

Private Sub RETORNAR_Click()
Grid1.Rows = 1
rut2.text = ""
lblDV.Caption = ""
lblNombre.Caption = ""
lblCupo.Caption = "$ 0"
lblUtilizado.Caption = "$ 0"
lblDisponible.Caption = "$ 0"
total1.Caption = "$ 0"
total2.Caption = "$ 0"
total3.Caption = "$ 0"
total4.Caption = "$ 0"
desde1.text = ""
desde2.text = ""
desde3.text = ""
hasta1.text = ""
hasta2.text = ""
hasta3.text = ""
NUMERO.text = ""

rut2.SetFocus
End Sub

Private Sub rut2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Call ayudaCliente(rut2, SUCU, lblDV)
  End If
End Sub
Private Sub rut2_KeyPress(KeyAscii As Integer)
          KeyAscii = esNumero(KeyAscii)
           If KeyAscii = 13 And rut2.text <> "" And Val(rut2.text) <> 0 Then
             rut2.text = ceros(rut2)
             lblDV.Caption = rut(rut2.text)
             If LEERCLIENTE(rut2.text + lblDV.Caption) = True Then
                desde1.SetFocus
             Else
                MsgBox ("CLIENTE NO CORRESPONDE A CLIENTE A CREDITO O NO TIENE CUPO ASIGNADO")
                rut2.SetFocus
             End If
        End If
End Sub

Public Function LEERCLIENTE(rut) As Boolean

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas
        
        cSql.sql = "SELECT mc.diapago,mc.cupodirecto,mc.cupoutilizadodirecto,mc.nombre,mc.direccion "
        cSql.sql = cSql.sql & "FROM sv_maestroclientes as mc "
        cSql.sql = cSql.sql & "WHERE mc.rut='" + rut + "' "
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultado = cSql.OpenResultset
            While Not resultado.EOF
        
        lblCupo.Caption = Format(resultado(1), "###,###,##0")
        
        lblNombre.Caption = resultado(3)

        
        If resultado(1) > 0 Then
                LEERCLIENTE = True
        Else
                LEERCLIENTE = True
        
        End If
        
            resultado.MoveNext
            Wend
        Else
        LEERCLIENTE = False
        
        End If
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
    End Function
'
'    Sub LEErcuotas(rut)
'
'        Dim cSql As rdoQuery
'        Dim resultado As rdoResultset
'        Dim i As Integer
'        Dim diasmora As Double
'        Dim saldo As Double
'        Dim interes As Double
'        Dim t1 As Double
'        Dim t2 As Double
'        Dim total As Double
'        Dim porcecondo1 As Double
'        Dim porcecondo2 As Double
'        Dim cuota As Double
'        Dim interescuota As Double
'        Dim capital As Double
'        Dim cuotabase As Double
'
'        Set cSql = New rdoQuery
'        Set cSql.ActiveConnection = ventas
'
'        cSql.sql = "SELECT *  "
'        cSql.sql = cSql.sql & "FROM sv_cuotas_detalle "
'        cSql.sql = cSql.sql & "WHERE rut='" + rut + "' " ' and ( (montocuota-abono)>0 or ((interesmora+montocuota)-abono)>0)
'        cSql.sql = cSql.sql & "order by vencimientoactual "
'        cSql.Execute
'        totalusado = 0
'        moratotal = 0
'        If cSql.RowsAffected > 0 Then
'
'            Set resultado = cSql.OpenResultset
'
'        Grid1.Rows = 1
'        Grid1.AutoRedraw = False
'
'        totalusado = 0
'        moratotal = 0
'        While Not resultado.EOF
'        Grid1.Rows = Grid1.Rows + 1
'        Grid1.Cell(Grid1.Rows - 1, 0).text = resultado(0)
'        Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(1) & " " & resultado(2)
'        Grid1.Cell(Grid1.Rows - 1, 2).text = Format(resultado(4), "00") & " / " & Format(resultado(12), "00")
'        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(6), "dd/mm/yyyy")
'        cuotabase = resultado(7)
'        cuota = resultado(7)
'
'        interescuota = resultado(7) - resultado("capitalcuota")
'
'
'        capital = resultado("capitalcuota")
'
'        cuota = cuotabase
'
'
'
'        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(cuota, "###,###,###")
'        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(8), "###,###,###")
'
'        saldo = (cuota + resultado("interesmora")) - resultado(8)
'
'
'        tazainteresmora = leerInteresMora("00")
'        porcecondo2 = 1 - (CDbl(0) / 100)
'        tazainteresmora = tazainteresmora * porcecondo2
'        If resultado(1) <> "CA" Then
'        diasmora = DateDiff("d", resultado(6), fechasistema)
'        Else
'        diasmora = 0
'        End If
'        If diasmora <= diasgracia Then
'        diasmora = 0
'        Else
'
'        End If
'
'        interes = Round(saldo * ((tazainteresmora * diasmora) / 100), 0)
'
'        total = saldo + interes
'        If saldo = 0 Then
'        Grid1.Cell(Grid1.Rows - 1, 6).text = "0"
'        Else
'         Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
'        End If
'        Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
'        Grid1.Cell(Grid1.Rows - 1, 8).text = interes
'        If total = 0 Then
'        Grid1.Cell(Grid1.Rows - 1, 9).text = "0"
'        Else
'        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(total, "###,###,###")
'        End If
'        Grid1.Cell(Grid1.Rows - 1, 10).text = "0"
'        Grid1.Cell(Grid1.Rows - 1, 11).text = resultado(13)
'        Grid1.Cell(Grid1.Rows - 1, 12).text = resultado("capitalcuota")
'
'
'
'        totalusado = totalusado + total
'        If interes <> 0 Then moratotal = moratotal + total
'        If Format(resultado(6), "yyyy-mm") <= Format(fechasistema, "yyyy-mm") Then
'        Grid1.Cell(Grid1.Rows - 1, 10).text = Grid1.Cell(Grid1.Rows - 1, 9).text
'        t1 = t1 + saldo
'        t2 = t2 + interes
'
'        End If
'
'            resultado.MoveNext
'            Wend
'        Else
'
'        End If
'        Set resultado = Nothing
'        cSql.Close
'        Set cSql = Nothing
'        Grid1.AutoRedraw = True
'        Grid1.Refresh
'        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
'        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
'        total4.Caption = Format(totalusado, "###,###,##0")
'        total1.Caption = Format(t1, "###,###,##0")
'        total2.Caption = Format(t2, "###,###,##0")
'        total3.Caption = Format(t1 + t2, "###,###,##0")
'
'    SUMAPAGOS
'    End Sub
    
    
    Sub LEErcuotas(rut, desde, hasta)

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim total As Double
        Dim porcecondo1 As Double
        Dim porcecondo2 As Double
        Dim cuota As Double
        Dim interescuota As Double
        Dim capital As Double
        Dim cuotabase As Double
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT *  "
        cSql.sql = cSql.sql & "FROM sv_cuotas_detalle "
        If NUMERO.text <> "" And NUMERO.text <> "0000000000" Then
        cSql.sql = cSql.sql & "WHERE rut='" + rut + "' and numero='" + NUMERO.text + "' AND fechacompra between '" & desde & "' and '" & hasta & "' "
        Else
        cSql.sql = cSql.sql & "WHERE rut='" + rut + "' and  fechacompra between '" & desde & "' and '" & hasta & "'"
        End If
        cSql.sql = cSql.sql & "order by fechacompra,local,tipo,numero,numerocuota asc "
        cSql.Execute
        totalusado = 0
        moratotal = 0
        If cSql.RowsAffected > 0 Then
            Set resultado = cSql.OpenResultset
            Grid1.Rows = 2
            Grid1.AutoRedraw = False
            totalusado = 0
            moratotal = 0
            While Not resultado.EOF
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 0).text = resultado(0)
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(1) & " " & resultado(2)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(4) & " / " & resultado(12)
                Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(6), "dd/mm/yyyy")
                cuotabase = resultado(7)
                cuota = resultado(7)
                interescuota = resultado(7) - resultado("capitalcuota")
                capital = resultado("capitalcuota")
                cuota = cuotabase
                Grid1.Cell(Grid1.Rows - 1, 4).text = Format(cuota, "###,###,###")
                Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(8), "###,###,###")
                saldo = (cuota) - resultado(8)
                Grid1.Cell(Grid1.Rows - 1, 10).text = Format(saldo, "###,###,###")
                tazainteresmora = leerInteresMora("00")
                diasmora = DateDiff("d", resultado(6), fechasistema)
                If diasmora > 0 Then
                    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HFF&
                End If
                 If diasmora <= diasgracia Then diasmora = 0
                 interes = Round(saldo * ((tazainteresmora / 100 / 30) * diasmora), 0)
                 total = saldo + interes
                If saldo = 0 Then
                    Grid1.Cell(Grid1.Rows - 1, 6).text = "0"
                Else
                    Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
                End If
                Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
                Grid1.Cell(Grid1.Rows - 1, 8).text = interes
                If total = 0 Then
                    Grid1.Cell(Grid1.Rows - 1, 9).text = "0"
                Else
                    Grid1.Cell(Grid1.Rows - 1, 9).text = Format(saldo, "###,###,###")
                End If
                If Not IsNull(resultado("fechacompra")) = True Then
                    fechacom = Format(resultado("fechacompra"), "dd-mm-yyyy")
                End If
        
                Grid1.Cell(Grid1.Rows - 1, 11).text = fechacom + "  " + resultado(13)
                Grid1.Cell(Grid1.Rows - 1, 12).text = resultado("capitalcuota")
        
                totalusado = totalusado + total
                If interes <> 0 Then moratotal = moratotal + total
                If Format(resultado(6), "yyyy-mm") <= Format(fechasistema, "yyyy-mm") Then
                    Grid1.Cell(Grid1.Rows - 1, 10).text = Grid1.Cell(Grid1.Rows - 1, 9).text
                    t1 = t1 + saldo
                    t2 = t2 + interes
                End If
            
                resultado.MoveNext
            Wend
         Else
            MsgBox "NO SE HAN ENCONTRADO RESULTADOS", vbInformation, "ATENCION"
         End If
       
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
        
        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
        total4.Caption = Format(totalusado, "###,###,##0")
        total1.Caption = Format(t1, "###,###,##0")
        total2.Caption = Format(t2, "###,###,##0")
        total3.Caption = Format(t1 + t2, "###,###,##0")
    
  
'    Call LEErcuotasACUMULADAS(rut)
    End Sub
    
    
    
 
Sub CARGAGRILLA()
    Grid1.Cols = 13
    
    Grid1.Column(0).Width = 20
    Grid1.Column(1).Width = 90
    Grid1.Column(2).Width = 50
    Grid1.Column(3).Width = 80
    Grid1.Column(4).Width = 50
    Grid1.Column(5).Width = 50
    Grid1.Column(6).Width = 0
    Grid1.Column(7).Width = 0
    Grid1.Column(8).Width = 0
    Grid1.Column(9).Width = 0
    Grid1.Column(10).Width = 50
    Grid1.Column(11).Width = 600
    Grid1.Column(12).Width = 0
   
    
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    Grid1.Column(8).Locked = True
    Grid1.Column(9).Locked = True
    Grid1.Column(10).Locked = True
    Grid1.Column(11).Locked = True
    Grid1.Column(3).CellType = cellCalendar
    
    Grid1.Cell(0, 0).text = "LO"
    Grid1.Cell(0, 1).text = "DOCUMENTO"
    Grid1.Cell(0, 2).text = "N.CUOTA"
    Grid1.Cell(0, 3).text = "VENCIMIENTO"
    Grid1.Cell(0, 4).text = "CUOTA"
    Grid1.Cell(0, 5).text = "ABONO"
    Grid1.Cell(0, 6).text = "SALDO"
    Grid1.Cell(0, 7).text = "DIAS MORA"
    Grid1.Cell(0, 8).text = "INTERES"
    Grid1.Cell(0, 9).text = "TOTAL"
    Grid1.Cell(0, 10).text = "SALDO"
    
    Grid1.Cell(0, 11).text = "DETALLE COMPRAS "
    Grid1.Cell(0, 12).text = "CAPITAL"
    
    Grid1.Column(4).Alignment = cellRightTop
    Grid1.Column(5).Alignment = cellRightTop
    Grid1.Column(6).Alignment = cellRightTop
    Grid1.Column(7).Alignment = cellRightTop
    Grid1.Column(8).Alignment = cellRightTop
    Grid1.Column(9).Alignment = cellRightTop
    Grid1.Column(10).Alignment = cellRightTop
    
    Grid1.Column(11).Alignment = cellLeftCenter
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid1.Rows = 1
   
End Sub

 
