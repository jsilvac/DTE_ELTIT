VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado7 
   Caption         =   "LISTA ESTADOS DE COBRANZA"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   15105
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   13573
      BackColor       =   16761024
      Caption         =   ""
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
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   6960
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin FlexCell.Grid GRID1 
         Height          =   6690
         Left            =   100
         TabIndex        =   2
         Top             =   240
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   11800
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   2910
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   5133
      BackColor       =   16761024
      Caption         =   ""
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
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "RETORNO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2400
         Width           =   2760
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2400
         Width           =   2760
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1230
         Left            =   12960
         TabIndex        =   13
         Top             =   270
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   2170
         BackColor       =   16744576
         Caption         =   "Detalle del Listado"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Detallado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   135
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FF8080&
            Caption         =   "Acumulado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   135
            TabIndex        =   14
            Top             =   765
            Width           =   1680
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000C000&
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox sucu 
         Height          =   285
         Left            =   15240
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "0"
         Top             =   2040
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         TabIndex        =   3
         Top             =   3480
         Width           =   1230
      End
      Begin XPFrame.FrameXp frmmes 
         Height          =   660
         Left            =   10920
         TabIndex        =   5
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "MES"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOMES 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   45
            TabIndex        =   6
            Top             =   225
            Width           =   3615
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   2100
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   3704
         BackColor       =   16744576
         Caption         =   "CLIENTES X RUT"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox rut1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   90
            MaxLength       =   9
            TabIndex        =   9
            Tag             =   "proveedor"
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label lblg 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " GIRO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5520
            TabIndex        =   30
            Top             =   600
            Width           =   570
         End
         Begin VB.Label lbld 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DIRECION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   945
         End
         Begin VB.Label lblf 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FONO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7920
            TabIndex        =   28
            Top             =   270
            Width           =   585
         End
         Begin VB.Label lblgiro 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   6150
            TabIndex        =   26
            Top             =   600
            Width           =   4530
         End
         Begin VB.Label lblsaldo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   525
            Left            =   6600
            TabIndex        =   22
            Top             =   1320
            Width           =   4065
         End
         Begin VB.Label lblutilizado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   525
            Left            =   3240
            TabIndex        =   20
            Top             =   1320
            Width           =   3330
         End
         Begin VB.Label lblcupo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000007&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   525
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   3105
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DISPONIBLE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6600
            TabIndex        =   25
            Top             =   1080
            Width           =   4065
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CUPO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   3105
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "UTILIZADO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   23
            Top             =   1080
            Width           =   3330
         End
         Begin VB.Label lbldireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Top             =   600
            Width           =   4410
         End
         Begin VB.Label lblnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   2025
            TabIndex        =   11
            Top             =   270
            Width           =   5865
         End
         Begin VB.Label lblDV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1530
            TabIndex        =   10
            Top             =   270
            Width           =   375
         End
         Begin VB.Label lblfono 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   8520
            TabIndex        =   27
            Top             =   270
            Width           =   2145
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   1230
         Left            =   11025
         TabIndex        =   16
         Top             =   270
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   2170
         BackColor       =   16744576
         Caption         =   "Tipo Listado"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FF8080&
            Caption         =   "Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   135
            TabIndex        =   18
            Top             =   765
            Width           =   1680
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FF8080&
            Caption         =   "Todo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   135
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "tmplistado7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CARGARDESDEAFUERAtmp()
    rut1_KeyPress (13)
 End Sub
Private Sub COMBOMES_Click()
LEErCREDITOS
End Sub


Private Sub Command1_Click()

Call Titulos("LISTADO DE CUOTAS POR VENCER ")
GRID1.PageSetup.Orientation = cellPortrait

GRID1.PageSetup.HeaderMargin = 0.5
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.PrintTitleRows = 0
GRID1.PageSetup.TopMargin = 2
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0.5
GRID1.PageSetup.BottomMargin = 3
GRID1.PageSetup.FooterMargin = 2
GRID1.PageSetup.BlackAndWhite = True


GRID1.PrintPreview
End Sub

Private Sub Command2_Click()
Call CargaGrillaGRID1(1, 11)
LEErCREDITOS

End Sub



Private Sub Command3_Click()
GRID1.Rows = 1
LEErCREDITOS
End Sub

Private Sub Command4_Click()
GRID1.Rows = 1
rut1.text = ""
lblDV.Caption = ""
lblnombre.Caption = ""
lblfono.Caption = ""
lbldireccion.Caption = ""
lblgiro.Caption = ""
lblcupo.Caption = ""
lblutilizado.Caption = ""
lblsaldo.Caption = ""
rut1.SetFocus
End Sub

Private Sub Form_Activate()
rut1.SetFocus
End Sub

Private Sub Form_Load()
Call CargaGrillaGRID1(1, 11)

       For K = 1 To 12
    COMBOMES.AddItem MonthName(K)
    Next K
    COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
    mes = Format(COMBOMES.ListIndex + 1, "00")


End Sub

 Private Sub CargaGrillaGRID1(ByVal Row As Integer, ByVal Col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "LO"
        formatogrilla(1, 2) = "F.COMPRA"
        formatogrilla(1, 3) = "TD"
        formatogrilla(1, 4) = "NUMERO"
        formatogrilla(1, 5) = "COMPRA"
        formatogrilla(1, 6) = "VENCIMIENTO"
        formatogrilla(1, 7) = "CUOTA/DE"
        formatogrilla(1, 8) = "M.CUOTA"
        formatogrilla(1, 9) = "INT.MORA"
        formatogrilla(1, 10) = "TOTAL"
        
        Rem ANCHO
        formatogrilla(8, 1) = "3"
        formatogrilla(8, 2) = "8"
        formatogrilla(8, 3) = "4"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "25"
        formatogrilla(8, 6) = "10"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
        formatogrilla(8, 10) = "8"
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "D"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "D"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""

        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"

        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        
            
        GRID1.Cols = Col
        GRID1.Rows = Row
        GRID1.AllowUserResizing = True
        GRID1.DisplayFocusRect = False
        GRID1.ExtendLastCol = True
        GRID1.BoldFixedCell = False
        GRID1.DrawMode = cellOwnerDraw
        GRID1.Appearance = Flat
        GRID1.ScrollBarStyle = Flat
        GRID1.FixedRowColStyle = Flat
        GRID1.BackColorFixed = RGB(90, 158, 214)
        GRID1.BackColorFixedSel = RGB(110, 180, 230)
        GRID1.BackColorBkg = RGB(90, 158, 214)
        GRID1.BackColorScrollBar = RGB(231, 235, 247)
        GRID1.BackColor1 = RGB(231, 235, 247)
        GRID1.BackColor2 = RGB(239, 243, 255)
        GRID1.GridColor = RGB(148, 190, 231)
        
        GRID1.Column(0).Width = 0
        For i = 1 To Col - 1
            GRID1.Cell(0, i).text = formatogrilla(1, i)
            GRID1.Column(i).Width = Val(formatogrilla(8, i)) * (GRID1.Cell(0, i).Font.Size)
            GRID1.Column(i).MaxLength = Val(formatogrilla(2, i))
            GRID1.Column(i).FormatString = formatogrilla(4, i)
            GRID1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                GRID1.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                GRID1.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                GRID1.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        GRID1.Range(0, 0, 0, GRID1.Cols - 1).Alignment = cellCenterCenter
        GRID1.Enabled = True
    End Sub
'**
Sub LEErCREDITOS()

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim CREDITO As Double
        Dim usado As Double
        Dim disponible As Double
        Dim mora As Double
        Dim total1 As Double
        Dim total2 As Double
        Dim total3 As Double
        Dim total4 As Double
        Dim total5 As Double
        Dim ACUMULADO As Double
        Dim FECHAMORA As String
        Dim fechaseparacion As String
        Dim PASADAS As Double
        
        Dim fecha1 As String
        Dim fecha2 As String
        mes = Format(COMBOMES.ListIndex + 1, "00")
        fecha1 = "2000-01-01"
        fecha2 = Format(fechasistema, "yyyy") + "-" + mes + "-" + "31"
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas
        pasada = 0
        cSql.sql = "SELECT cd.local,cd.fechacompra,cd.tipo,cd.numero,cd.glosacompra,cd.vencimientoactual,cd.numerocuota,cd.cantidadcuotas,cd.montocuota-cd.abono "
        cSql.sql = cSql.sql & "FROM sv_cuotas_detalle as cd "
        cSql.sql = cSql.sql & "WHERE cd.rut='" + rut1.text + lblDV.Caption + "' and montocuota>abono  "
        
        If Option4.Value = True Then
           cSql.sql = cSql.sql & "and mid(vencimientoactual,1,7)='" + Mid(fecha2, 1, 7) + "' "
        End If
        cSql.sql = cSql.sql & "order by cd.vencimientoactual "
        cSql.Execute
        
    If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
'        If Option1.Value = True Then separador = resultado(4)
'        If Option2.Value = True Then separador = resultado(6)
'
            BARRA.Max = cSql.RowsAffected + 1
            fechaseparacion = Format(resultado(5), "yyyy-mm")
            GRID1.Rows = 1
            GRID1.AutoRedraw = False
        
            total1 = 0
            total2 = 0
            total3 = 0
            total4 = 0
            total5 = 0
        
        
        While Not resultado.EOF
        
        'BARRA.Value = BARRA.Value + 1
        
            tazainteresmora = leerInteresMora("00")
            diasmora = DateDiff("d", resultado(5), fechasistema)
        
             If diasmora <= diasgracia Then diasmora = 0
                mora = Round(resultado(8) * ((tazainteresmora / 100 / 30) * diasmora), 0)
                ACUMULADO = ACUMULADO + (resultado(7) + mora)
                If fechaseparacion <> Format(resultado(5), "yyyy-mm") Then
                    GRID1.Rows = GRID1.Rows + 1
                    GRID1.Cell(GRID1.Rows - 1, 5).text = "TOTAL MES "
                    GRID1.Cell(GRID1.Rows - 1, 6).text = Format(fechaseparacion, "mm-yyyy")
                    GRID1.Cell(GRID1.Rows - 1, 8).text = Format(total1, "###,###,###")
                    GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total2, "###,###,###")
                    GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total3, "###,###,###")
                    total1 = 0
                    total2 = 0
                    total3 = 0
                    GRID1.Range(GRID1.Rows - 1, 5, GRID1.Rows - 1, 10).Borders(cellEdgeTop) = cellThin
                    fechaseparacion = Format(resultado(5), "yyyy-mm")
                    
                    If Option2.Value = False Then
                        GRID1.Rows = GRID1.Rows + 2
                    Else
                        GRID1.Rows = GRID1.Rows + 1
                    End If
        
                 End If
        
        
                 If Option1.Value = True Or Format(resultado(5), "yyyy-mm") <= Format(fechasistema, "yyyy-mm") Then
        
                    GRID1.Rows = GRID1.Rows + 1
                    GRID1.Cell(GRID1.Rows - 1, 1).text = resultado(0)
                    If IsNull(resultado(1)) = False Then
                        GRID1.Cell(GRID1.Rows - 1, 2).text = Format(resultado(1), "dd-mm-yyyy")
                    End If
        
                    GRID1.Cell(GRID1.Rows - 1, 3).text = resultado(2)
                    GRID1.Cell(GRID1.Rows - 1, 4).text = resultado(3)
                    GRID1.Cell(GRID1.Rows - 1, 5).text = resultado(4)
                    GRID1.Cell(GRID1.Rows - 1, 6).text = Format(resultado(5), "dd-mm-yyyy")
                    GRID1.Cell(GRID1.Rows - 1, 7).text = resultado(6) & "/" & resultado(7)
                    GRID1.Cell(GRID1.Rows - 1, 8).text = Format(resultado(8), "###,###,###")
                    GRID1.Cell(GRID1.Rows - 1, 9).text = Format(mora, "###,###,###")
                    GRID1.Cell(GRID1.Rows - 1, 10).text = Format(resultado(8) + mora, "###,###,###")
                  End If
                  total1 = total1 + resultado(8)
                  total2 = total2 + mora
                  total3 = total3 + (resultado(8) + mora)
                  total11 = total11 + resultado(8)
                  total12 = total12 + mora
                  total13 = total13 + (resultado(8) + mora)
                
                  resultado.MoveNext
        Wend
    Else
       
    End If
         
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Cell(GRID1.Rows - 1, 5).text = "TOTAL MES"
        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(fechaseparacion, "mm-yyyy")
        GRID1.Cell(GRID1.Rows - 1, 8).text = Format(total1, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total2, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total3, "###,###,###")
        total1 = 0
        total2 = 0
        total3 = 0
        GRID1.Range(GRID1.Rows - 1, 5, GRID1.Rows - 1, 10).Borders(cellEdgeTop) = cellThin
        
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 5, GRID1.Rows - 1, 10).Borders(cellEdgeTop) = cellThick
        
        
        GRID1.Cell(GRID1.Rows - 1, 5).text = "TOTALES GENERALES"
        GRID1.Cell(GRID1.Rows - 1, 8).text = Format(total11, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total12, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total13, "###,###,###")
        
        
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        GRID1.AutoRedraw = True
        GRID1.Refresh
    End Sub

Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    GRID1.FixedRowColStyle = Fixed3D
    GRID1.CellBorderColorFixed = vbButtonShadow
    GRID1.ShowResizeTips = False
    GRID1.ReportTitles.Clear
    GRID1.PageSetup.CenterHorizontally = True
    GRID1.PageSetup.Orientation = cellPortrait
    
    
      
    GRID1.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    GRID1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    GRID1.PageSetup.HeaderAlignment = cellLeft
    GRID1.PageSetup.HeaderFont.Name = "Verdana"
    GRID1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
   
'    If Option1.Value = True Then tipoListado = "CLIENTES MAAT"
'    If Option2.Value = True Then tipoListado = "CLIENTES SKORPIOS"
'    If Option3.Value = True Then tipoListado = "CLIENTES TODOS"
'
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "ESTADO DE CUENTA "
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CLIENTE     :" + rut1.text + "-" + lblDV.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "NOMBRE     :" + lblnombre.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "DIRECCION     : " + lbldireccion.Caption
    
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CUPO     : " & Format(leerCupoCliente(rut1.text + lblDV.Caption), "$ ###,###,###")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CUPO UTILIZADO  : " & Format(LEErcreditoutilizado(rut_cliente, "9999-99-99"), "$ ###,###,###")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " "
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
  
    
    
    
    'PIE DE PAGINA
    GRID1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & usuarioSistema
    GRID1.PageSetup.FooterAlignment = cellRight
    GRID1.PageSetup.FooterFont.Name = "Verdana"
    GRID1.PageSetup.FooterFont.Size = 7
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeTop) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeBottom) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeLeft) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeRight) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideVertical) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).FontBold = True
           
End Sub




Private Sub Option1_Click()
Call Command3_Click
End Sub

Private Sub Option2_Click()
Call Command3_Click
End Sub

Private Sub Option3_Click()
Call Command3_Click
End Sub

Private Sub Option4_Click()
Call Command3_Click
End Sub

Private Sub rut1_GotFocus()
        Call VerificarCajas(Me, rut1)
        Call selecciona(rut1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
End Sub

Private Sub rut1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
            Call ayudaCliente(rut1, sucu, lblDV)
        Else
            Call Flechas(KeyCode, rut1)
        End If
End Sub

Private Sub rut1_KeyPress(KeyAscii As Integer)

 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And rut1.text <> "" And Val(rut1.text) <> 0 Then
            rut1.text = ceros(rut1)
            lblDV.Caption = rut(rut1.text)
            rut_cliente = rut1.text + lblDV.Caption
            lblnombre.Caption = leerNombreCliente(rut_cliente)
            lbldireccion.Caption = leerDireccionCliente(rut_cliente, "0")
            lblfono.Caption = leerFonoCliente(rut_cliente, "0")
            lblgiro.Caption = leerGiroCliente(rut_cliente, "0")
            
            lblcupo = leerCupoCliente(rut1.text & lblDV.Caption)
            lblutilizado.Caption = LEErcreditoutilizado(rut1.text & lblDV.Caption, "9999-99-99")
            lblsaldo.Caption = Format(CDbl(lblcupo.Caption) - CDbl(lblutilizado.Caption), "###,###,##0")
            lblcupo.Caption = Format(lblcupo.Caption, "###,###,##0")
            lblutilizado.Caption = Format(lblutilizado.Caption, "###,###,##0")
           
            ' COMBOMES.SetFocus
            
           Command3_Click
           
        End If
End Sub

