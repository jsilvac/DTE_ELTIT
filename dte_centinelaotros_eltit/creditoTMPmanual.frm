VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form creditoTMPMANUAL 
   BackColor       =   &H00AE1118&
   BorderStyle     =   0  'None
   Caption         =   "Crédito"
   ClientHeight    =   9075
   ClientLeft      =   210
   ClientTop       =   1710
   ClientWidth     =   14580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8310
      Left            =   45
      TabIndex        =   1
      Top             =   675
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   14658
      BackColor       =   16744576
      Caption         =   "Venta a Crédito"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   330
         Left            =   13770
         TabIndex        =   20
         Top             =   45
         Width           =   465
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1815
         Left            =   90
         TabIndex        =   16
         Top             =   2790
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   3201
         BackColor       =   49344
         Caption         =   "Calcula Cuotas"
         CaptionEstilo3D =   1
         BackColor       =   49344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton ACEPTAR 
            BackColor       =   &H0000FF00&
            Caption         =   "ACEPTAR"
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
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1350
            Width           =   2085
         End
         Begin VB.CommandButton CALCULAR 
            BackColor       =   &H0000FF00&
            Caption         =   "CALCULAR"
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
            Left            =   2295
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1350
            Width           =   2085
         End
         Begin VB.TextBox CUOTAS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2520
            TabIndex        =   0
            Text            =   "3"
            Top             =   810
            Width           =   2130
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CUOTAS"
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
            Height          =   330
            Left            =   2475
            TabIndex        =   17
            Top             =   495
            Width           =   2220
         End
         Begin VB.Label VENCIMIENTO 
            Alignment       =   2  'Center
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
            Left            =   7110
            TabIndex        =   25
            Top             =   810
            Width           =   2025
         End
         Begin VB.Label VALORCUOTA 
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
            Left            =   4860
            TabIndex        =   24
            Top             =   810
            Width           =   2055
         End
         Begin VB.Label MONTO 
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
            Left            =   225
            TabIndex        =   23
            Top             =   810
            Width           =   2055
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MONTO CREDITO"
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
            Height          =   330
            Left            =   180
            TabIndex        =   21
            Top             =   495
            Width           =   2130
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PRIMER VENCIMIENTO"
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
            Height          =   330
            Left            =   7065
            TabIndex        =   19
            Top             =   495
            Width           =   2130
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MONTO CUOTA"
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
            Height          =   330
            Left            =   4815
            TabIndex        =   18
            Top             =   495
            Width           =   2130
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   3480
         Left            =   90
         TabIndex        =   14
         Top             =   4680
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   6138
         BackColor       =   49344
         CaptionEstilo3D =   1
         BackColor       =   49344
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
         Begin FlexCell.Grid Grid1 
            Height          =   3120
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   13830
            _ExtentX        =   24395
            _ExtentY        =   5503
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4200
         Left            =   9720
         TabIndex        =   28
         Top             =   450
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   7408
         BackColor       =   49344
         Caption         =   "SIMULADOR DE CUOTAS"
         CaptionEstilo3D =   1
         BackColor       =   49344
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
         Begin FlexCell.Grid Grid2 
            Height          =   3840
            Left            =   90
            TabIndex        =   29
            Top             =   225
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   6773
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.Label lblrut 
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
         Left            =   2025
         TabIndex        =   22
         Top             =   495
         Width           =   1785
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
         Left            =   120
         TabIndex        =   13
         Top             =   480
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
         Left            =   3870
         TabIndex        =   12
         Top             =   495
         Width           =   5655
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
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
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblDireccion 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   960
         Width           =   7575
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Día de Pago"
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label lblDiaPago 
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
         Left            =   2025
         TabIndex        =   8
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Autorizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblCupo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Utilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   4
         Left            =   3240
         TabIndex        =   5
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label lblUtilizado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3240
         TabIndex        =   4
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   6480
         TabIndex        =   3
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   6480
         TabIndex        =   2
         Top             =   2280
         Width           =   3135
      End
   End
End
Attribute VB_Name = "creditoTMPMANUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CALCULAR_Click()
CALCULACUOTA


End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub CUOTAS_GotFocus()
Call cargatexto(CUOTAS)

End Sub

Private Sub CUOTAS_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If CUOTAS.text <> "" Then
If KeyAscii = 13 And CUOTAS.text <> "" And CUOTAS.text <> "0" And CDbl(CUOTAS.text) <= leefactorMAXIMO Then
CALCULAR_Click
End If


End If

End Sub

Private Sub Form_Load()
CARGAGRILLA
CARGAGRILLA2

End Sub

Sub CARGAGRILLA()
    Grid1.Cols = 9
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 3
    Grid1.Column(2).Width = 80
    Grid1.Column(3).Width = 80
    Grid1.Column(4).Width = 80
    Grid1.Column(5).Width = 80
    Grid1.Column(6).Width = 80
    Grid1.Column(7).Width = 80
    Grid1.Column(8).Width = 80
    
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    Grid1.Column(8).Locked = True
    
    Grid1.Cell(0, 1).text = "TD"
    Grid1.Cell(0, 2).text = "NUMERO"
    Grid1.Cell(0, 3).text = "FECHA"
    Grid1.Cell(0, 4).text = "CUOTA"
    Grid1.Cell(0, 5).text = "VENCIMIENTO"
    Grid1.Cell(0, 6).text = "MOROSIDAD"
    Grid1.Cell(0, 7).text = "INTERES"
    Grid1.Cell(0, 8).text = "TOTAL"
    
    Grid1.Range(0, 1, 0, 8).Alignment = cellCenterGeneral
    Grid1.Range(0, 1, 0, 8).FontSize = 8
    Grid1.Range(0, 1, 0, 8).FontBold = True
    Grid1.Range(0, 1, 0, 8).Borders(cellEdgeBottom) = cellThick
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid1.Rows = 1
    
   
End Sub
Sub CARGAGRILLA2()
    Grid2.Cols = 4
    Grid2.Column(0).Width = 0
    Grid2.Column(1).Width = 40
    Grid2.Column(2).Width = 60
    Grid2.Column(3).Width = 60
    
    Grid2.Column(0).Locked = True
    Grid2.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    Grid2.Column(3).Locked = True
    
    Grid2.Cell(0, 1).text = "CUOTAS"
    Grid2.Cell(0, 2).text = "VALOR"
    Grid2.Cell(0, 3).text = "TOTAL"
    
    Grid2.Column(1).Alignment = cellRightTop
    Grid2.Column(2).Alignment = cellRightTop
     Grid2.Column(3).Alignment = cellRightTop
    Grid2.Rows = leefactorMAXIMO + 1
    
    

End Sub



Sub CALCULACUOTA()
Dim FACTOR As Double

Dim CUOTA As Double
FACTOR = leefactor(CUOTAS.text)
CUOTA = Int((CDbl(Replace(MONTO.Caption, ",", "")) * FACTOR / CDbl(CUOTAS.text)) + 0.5)

VALORCUOTA.Caption = Format(CUOTA, "###,###,###")

CALCULATODASCUOTAS

End Sub
Sub CALCULATODASCUOTAS()
Dim K As Integer
Dim CUOTA As Double
Dim fin As Double

fin = leefactorMAXIMO

For K = 1 To fin

FACTOR = leefactor(Str(K))
CUOTA = Int((CDbl(Replace(MONTO.Caption, ",", "")) * FACTOR / K) + 0.5)
Grid2.Cell(K, 1).text = K

Grid2.Cell(K, 2).text = Format(CUOTA, "###,###,##0")
Grid2.Cell(K, 3).text = Format(CUOTA * K, "###,###,##0")

Next K

End Sub
Sub CALCULAPRIMERVENCIMIENTO()


End Sub

