VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form prove0016 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelacion Ordenes de Compra"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14895
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   610
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   993
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   32
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      BackColor       =   8454016
      Caption         =   " Mis Datos"
      BackColor       =   8454016
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton botonmisaccesos 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp frmcheque 
      Height          =   2535
      Left            =   3960
      TabIndex        =   16
      Top             =   4320
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   4471
      BackColor       =   16761024
      Caption         =   "PANTALLA DATOS DEL CHEQUE"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080FFFF&
         Caption         =   "RETORNO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox dato4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1530
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "INICIAR PROCESO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox dato3 
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
         Left            =   3420
         MaxLength       =   4
         TabIndex        =   19
         Top             =   405
         Width           =   735
      End
      Begin VB.TextBox dato2 
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
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   18
         Top             =   405
         Width           =   375
      End
      Begin VB.TextBox dato1 
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   17
         Tag             =   "codigo"
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO INICIAL"
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
         Height          =   240
         Left            =   1800
         TabIndex        =   24
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label lblBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   810
         Width           =   5145
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BANCO"
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
         Left            =   990
         TabIndex        =   20
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9090
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   16034
      BackColor       =   8454016
      Caption         =   "PANTALLA PAGO DE ORDENES DE COMPRA"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      ForeColor       =   65535
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   16384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   252
         Left            =   480
         TabIndex        =   36
         Top             =   8160
         Visible         =   0   'False
         Width           =   972
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   8640
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox ORDEN 
         BackColor       =   &H00FFC0C0&
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
         Left            =   13050
         MaxLength       =   10
         TabIndex        =   26
         Top             =   8235
         Width           =   1500
      End
      Begin VB.CommandButton BUSCAR 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Busca Orden"
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
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8235
         Width           =   1320
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Generar Pagos Electronicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8190
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8190
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8190
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1852
         BackColor       =   8454016
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   8438015
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton Option3 
            BackColor       =   &H0080FF80&
            Caption         =   "x Orden"
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
            Left            =   13440
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080FF80&
            Caption         =   "Devo.Automatica"
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
            Left            =   12720
            TabIndex        =   29
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H0080FF80&
            Caption         =   "Mensual"
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
            Left            =   12240
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0080FF80&
            Caption         =   "Diario"
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
            Left            =   11400
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "LISTAR"
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
            Left            =   11350
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   675
            Width           =   1335
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
            BackColor       =   8454016
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ForeColor       =   65535
            ColorBarraArriba=   16384
            ColorBarraAbajo =   8454016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOMES 
               Height          =   315
               Left            =   45
               TabIndex        =   10
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   11
            Top             =   270
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1191
            BackColor       =   8454016
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ForeColor       =   65535
            ColorBarraArriba=   16384
            ColorBarraAbajo =   8454016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOAÑO 
               Height          =   315
               Left            =   90
               TabIndex        =   12
               Top             =   270
               Width           =   2865
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   675
            Left            =   6705
            TabIndex        =   13
            Top             =   270
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   8454016
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ForeColor       =   65535
            ColorBarraArriba=   16384
            ColorBarraAbajo =   8454016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox ComboLOCAL 
               Height          =   315
               Left            =   90
               TabIndex        =   14
               Top             =   270
               Width           =   4395
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   135
         TabIndex        =   3
         Top             =   1485
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11774
         BackColor       =   8454016
         Caption         =   "LISTADO ORDENES A CANCELAR"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   8438015
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid1 
            Height          =   6360
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   11218
            BackColorFixed  =   8454016
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   32768
            Rows            =   30
            DateFormat      =   2
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Generar Pagos Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8190
         Width           =   2535
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "prove0016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lineassiguientes As Double
Private contabilizacionespendientes As Boolean
Private devolucionespendientes As Boolean

Private localfiltro As String
Private pagada As Boolean
Private NCHEQUE As Double
Private rutreal As String
Private DIFERENCIA As Double
Private contabilizada As String
Private lineafinal As Double
Private numerocontable As String
Private FECHACONTABLE As String
Private lineacontable As Double
Private rutcontable As String
Private tipocontable As String
Private TOTALCheque As Double
Private fechacheque As String
Private NOMBREGIRADO As String
Private montoharina As Double
Private montocarne As Double





Private Sub Command1_Click()
imprimir
End Sub



Private Sub Command2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1

leer


End Sub



Private Sub Command3_Click()
Dim k As Double
Dim rutprove As String
Dim tipo As String
Dim ordenes As Double
lineacontable = 15

año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(Format(fechasistema, "yyyy-mm-dd")) <> True Then


For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 14).text = "1" And Mid(Grid1.Cell(k, 1).text, 1, 2) = "TR" Then
        If rutprove <> Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1) Or lineacontable + leerlineas(Grid1.Cell(k, 2).text) > 10 Then
            If TOTALCheque <> 0 Then
            Call grabarcheque(TOTALCheque)
            End If
        fechacheque = Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd")
        NOMBREGIRADO = Grid1.Cell(k, 4).text
        FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
        tipocontable = "DB"
        numerocontable = LEERFOLIOCE("DB")
        lineacontable = 0
        rutprove = Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1)
        rutcontable = rutprove
        TOTALCheque = 0
        End If
        
        Call GRABARCOMPROBANTE(Grid1.Cell(k, 2).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd"), Grid1.Cell(k, 1).text, Grid1.Cell(k, 16).text)
        
    End If
Next k
        If TOTALCheque <> 0 Then
        Call grabarcheque(TOTALCheque)
        TOTALCheque = 0
        End If
Call modifica_glosas
leer
Else
MsgBox "MES YA CERRADO "

End If


End Sub

Private Sub Command4_Click()
'preguntar granate ariel

año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(Format(fechasistema, "yyyy-mm-dd")) <> True Then
frmcheque.Visible = True
dato1.SetFocus
Else
MsgBox "MES YA CERRADO"
End If

End Sub

Private Sub Command5_Click()


Dim k As Double
Dim rutprove As String
Dim tipo As String

If LBLBANCO.Caption <> "" Then

NCHEQUE = CDbl(dato4.text) - 1



For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 14).text = "1" And Mid(Grid1.Cell(k, 1).text, 1, 2) = "CH" Then
        If rutprove <> Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1) Then
        If TOTALCheque <> 0 Then
        Call grabarcheque(TOTALCheque)
        End If
        fechacheque = Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd")
        NOMBREGIRADO = Grid1.Cell(k, 4).text
        FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
        tipocontable = "PA"
        numerocontable = LEERFOLIOCE("PA")
        lineacontable = 0
        rutprove = Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1)
        rutcontable = rutprove
        TOTALCheque = 0
        End If
        
        Call GRABARCOMPROBANTE(Grid1.Cell(k, 2).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd"), Grid1.Cell(k, 1).text, Grid1.Cell(k, 16).text)
        
    End If
Next k
        If TOTALCheque <> 0 Then
        Call grabarcheque(TOTALCheque)
        TOTALCheque = 0
        End If
    
leer

End If
Call modifica_glosas
frmcheque.Visible = False
dato4.text = ""
End Sub

Private Sub Command6_Click()
frmcheque.Visible = False
dato4.text = ""

End Sub

Private Sub Command7_Click()
Call modifica_glosas

End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato4)
If leercheque(dato1.text + dato2.text + dato3.text, dato4.text) = True Then
MsgBox ("EL NUMERO DE CHEQUE YA ESTA EMITIDO")
dato4.text = ""
dato4.SetFocus
Else

Command5.SetFocus
End If

End If

End Sub

Private Sub Form_Load()

CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA
Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion00", Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
LEErlocales
frmcheque.Visible = False


End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub
Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub

Private Sub Label16_Click()
End Sub

Sub limpia()
    
    
End Sub

Sub imprimir()
Dim titulo As String
titulo = "ORDENES PENDIENTES DE PAGO " + COMBOMES.text + " " + COMBOAÑO.text
Call CABEZAS2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FormatoGrilla(10, 20)
    Grid1.DefaultFont.Size = 8
      
    FormatoGrilla(1, 1) = "TP"
    FormatoGrilla(1, 2) = "NUMERO"
    FormatoGrilla(1, 3) = "RUT"
    FormatoGrilla(1, 4) = "PROVEEDOR"
    FormatoGrilla(1, 5) = "FECHA"
    FormatoGrilla(1, 6) = "COMPRA"
    FormatoGrilla(1, 7) = "FACTURAS"
    FormatoGrilla(1, 8) = "OTROS"
    FormatoGrilla(1, 9) = "PAGAR"
    FormatoGrilla(1, 10) = "DIFERENCIA"
    FormatoGrilla(1, 11) = "PUB"
    FormatoGrilla(1, 12) = "DEV"
    FormatoGrilla(1, 13) = "OK"
    FormatoGrilla(1, 14) = "PAGA"
    FormatoGrilla(1, 15) = "FECHA PAGO"
    FormatoGrilla(1, 16) = "GLOSA DIFERENCIA"
    FormatoGrilla(1, 17) = "GLOSA DIGITADA"
    FormatoGrilla(1, 18) = "PROVEEDOR ORIGINAL"
    FormatoGrilla(1, 19) = "ANT"
  
    
    Rem LARGO DE LOS DATOS
    FormatoGrilla(2, 1) = "6"
    FormatoGrilla(2, 2) = "9"
    FormatoGrilla(2, 3) = "9"
    FormatoGrilla(2, 4) = "20"
    FormatoGrilla(2, 5) = "8"
    FormatoGrilla(2, 6) = "8"
    FormatoGrilla(2, 7) = "8"
    FormatoGrilla(2, 8) = "8"
    FormatoGrilla(2, 9) = "8"
    FormatoGrilla(2, 10) = "8"
    FormatoGrilla(2, 11) = "3"
    FormatoGrilla(2, 12) = "3"
    FormatoGrilla(2, 13) = "3"
    FormatoGrilla(2, 14) = "5"
    FormatoGrilla(2, 15) = "10"
    FormatoGrilla(2, 16) = "30"
    FormatoGrilla(2, 17) = "30"
    FormatoGrilla(2, 18) = "30"
    FormatoGrilla(2, 19) = "3"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FormatoGrilla(3, 1) = "S"
    FormatoGrilla(3, 2) = "S"
    FormatoGrilla(3, 3) = "S"
    FormatoGrilla(3, 4) = "S"
    FormatoGrilla(3, 5) = "S"
    FormatoGrilla(3, 6) = "N"
    FormatoGrilla(3, 7) = "N"
    FormatoGrilla(3, 8) = "N"
    FormatoGrilla(3, 9) = "N"
    FormatoGrilla(3, 10) = "N"
    FormatoGrilla(3, 11) = "S"
    FormatoGrilla(3, 12) = "S"
    FormatoGrilla(3, 13) = "S"
    FormatoGrilla(3, 14) = "S"
    FormatoGrilla(3, 15) = "S"
    FormatoGrilla(3, 16) = "S"
    FormatoGrilla(3, 17) = "S"
    FormatoGrilla(3, 18) = "S"
    FormatoGrilla(3, 19) = "S"
    
    Rem FORMATO GRILLA
    FormatoGrilla(4, 6) = "##,###,##0"
    FormatoGrilla(4, 7) = "##,###,##0"
    FormatoGrilla(4, 8) = "##,###,##0"
    FormatoGrilla(4, 9) = "##,###,##0"
    FormatoGrilla(4, 10) = "##,###,##0"
    
    Rem LOCCKED
    For k = 1 To 19
    FormatoGrilla(5, k) = "TRUE"
    
    Next k
    
  
    FormatoGrilla(5, 14) = "FALSE"
    FormatoGrilla(5, 15) = "FALSE"
    FormatoGrilla(5, 16) = "FALSE"
    
    Grid1.Cols = 20
    Grid1.Rows = 2
    
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = FormatoGrilla(1, k)
        Grid1.Column(k).Width = Val(FormatoGrilla(2, k)) * Grid1.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(FormatoGrilla(2, k))
        Grid1.Column(k).FormatString = FormatoGrilla(4, k)
        Grid1.Column(k).Locked = FormatoGrilla(5, k)
        If FormatoGrilla(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FormatoGrilla(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
    
    Grid1.Column(11).CellType = cellCheckBox
    Grid1.Column(12).CellType = cellCheckBox
    Grid1.Column(13).CellType = cellCheckBox
    Grid1.Column(14).CellType = cellCheckBox
    Grid1.Column(15).CellType = cellCalendar
    Grid1.Column(17).CellType = cellTextBox
    Grid1.Column(18).CellType = cellTextBox
    Grid1.Column(19).CellType = cellCheckBox
    Grid1.Column(1).Locked = False
    
    Grid1.Column(1).CellType = cellComboBox
    
    
    
    With Grid1.ComboBox(1)
        '.Locked = True
        .AutoComplete = True
        .Font.Name = "Courier New"
        .AddItem "CHEQUES"
        .AddItem "TRANSFERENCIA"
        
    End With
    
    
End Sub



Private Sub monto_Click()
End Sub

Public Sub leer()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim montofacturas As Double
    Dim OTROS As Double
    Dim apagar As Double
    Dim saldoctacte As String
    Dim sipu As String
    Dim DEVOLU As Boolean
    Dim glosa As String
    Dim plazo As String
    Dim fila As Double
    Dim tcompra As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
        Set csql.ActiveConnection = gestionrubro
'
'        csql.sql = "SELECT DISTINCT 'OC',oc.numero,oc.proveedor,mp.nombre,oc.fecha,oc.montorecepcionado,"
        'csql.sql = "oc.autorizacancelacion,oc.fechaautorizacionpago,mc.glosa,mc.tipoglosa,oc.glosadiferencia,oc.fecharecepcion "
'        csql.sql = csql.sql + "FROM " & clientesistema & "gestion" & rubro & ".l_ordendecompra_cabeza_" + localfiltro + " as oc "
        ' csql.sql inner join " + clientesistema + "gestion" + rubro + ".r_maestroproveedores_" + rubro + " as mp "
        'inner join " + clientesistema + "gestion" + rubro + ".l_movimientos_cabeza_" + localfiltro + " as mc on
'        (oc.proveedor=mp.rut and oc.numero=mc.numero and mc.tipo='OC' ) "
'        If Option1.Value = False Then
'            csql.sql = csql.sql + "where oc.fecharecepcion>='" + fecha1 + "' AND oc.fecharecepcion<='" + fecha2 + "' "
'        Else
'            csql.sql = csql.sql + "where oc.fecharecepcion='" & Format(fechasistema, "yyyy-mm-dd") + "' "
'        End If
'
''        csql.sql = csql.sql + "order by mp.nombre,oc.fecharecepcion "
'        csql.sql = csql.sql + "order by oc.rutcontable,oc.fecharecepcion "
'        csql.sql = csql.sql + " "
        
        
        
      csql.sql = "SELECT DISTINCT 'OC',oc.numero,oc.rutcontable,ifnull(cc.nombre,''),oc.fecha,oc.montorecepcionado,oc.autorizacancelacion, "
      csql.sql = csql.sql & "oc.fechaautorizacionpago,mc.glosa,mc.tipoglosa,oc.glosadiferencia,oc.fecharecepcion,mp.nombre "
      csql.sql = csql.sql & "FROM " & clientesistema & "gestion" & rubro & ".l_ordendecompra_cabeza_" & localfiltro & " as oc "
      csql.sql = csql.sql & "inner join " + clientesistema + "gestion" + rubro + ".r_maestroproveedores_" + rubro + " as mp "
      csql.sql = csql.sql & "inner join " & clientesistema & "gestion" & rubro & ".l_movimientos_cabeza_" & localfiltro & " as mc "
      csql.sql = csql.sql & "on (oc.proveedor=mp.rut and oc.numero=mc.numero and mc.tipo='OC' ) "
      csql.sql = csql.sql & "left join " & basedatos & ".cuentascorrientes as cc on oc.rutcontable=cc.rut and cc.tipo='" & CUENTAPROVEEDOR & "' "
      csql.sql = csql.sql & "and cc.año='" & Format(fechasistema, "yyyy") & "' "
        If Option2.Value = True Then
            csql.sql = csql.sql + "where oc.fecharecepcion>='" + fecha1 + "' AND oc.fecharecepcion<='" + fecha2 + "' "
        End If
        If Option1.Value = True Then
            csql.sql = csql.sql + "where oc.fecharecepcion='" & Format(fechasistema, "yyyy-mm-dd") + "' "
        End If
        If Option3.Value = True Then
            csql.sql = csql.sql + "where oc.numero='" & ORDEN.text + "' "
        End If
        
        
        csql.sql = csql.sql + "order by cc.nombre,oc.fecharecepcion "
        csql.sql = csql.sql + " "
        
        
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        barra.Value = 0
        
       contabilizacionespendientes = False
       devolucionespendientes = False
       
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
           barra.Value = barra.Value + 1
            montofacturas = LEERcompras(resultados(1))
            If pagada = False And montofacturas > 0 Then
             Grid1.Rows = Grid1.Rows + 1
             LINEA = LINEA + 1
             Rem If resultados(1) = "0000228059" Then Stop
             Grid1.Cell(LINEA, 1).text = leertipopago(rutreal)
             Rem Grid1.Cell(linea, 1).text = "CHEQUES"
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = Mid(rutreal, 1, 9) + "-" + Mid(rutreal, 10, 1)
'             Grid1.Cell(linea, 4).text = leerdatos(db, "cuentascorrientes", "nombre", "rut='" + rutreal + "' and tipo='" + CUENTAPROVEEDOR + "' and año='" & Format(fechasistema, "yyyy") & "'")
             Grid1.Cell(LINEA, 4).text = resultados(3)
             Grid1.Cell(LINEA, 18).text = resultados("nombre")
             If IsNull(resultados(3)) = False Then
             Grid1.Cell(LINEA, 5).text = resultados(4)
             End If
             tcompra = resultados(5) + montoharina + montocarne + LEERenlace(resultados(1))
             Grid1.Cell(LINEA, 6).text = resultados(5) + montoharina + montocarne + LEERenlace(resultados(1))
             Grid1.Cell(LINEA, 7).text = montofacturas
             OTROS = totalotros(resultados(1), localfiltro)
             Grid1.Cell(LINEA, 8).text = OTROS
             If montofacturas > resultados(5) Then
             apagar = resultados(5)
             DIFERENCIA = montofacturas - (tcompra + montoharina + montocarne)
                     If DIFERENCIA <= 1000 Then
                        DIFERENCIA = 0
                        apagar = montofacturas + OTROS
                    Else
                        apagar = resultados(5) + OTROS
             End If
             Else
             apagar = montofacturas + OTROS
             DIFERENCIA = 0
             End If
             Grid1.Cell(LINEA, 9).text = apagar
             Grid1.Cell(LINEA, 10).text = DIFERENCIA
             
             
             If leerpublicidad(rutreal, localfiltro) = True Then
             
             Grid1.Cell(LINEA, 11).text = "1"
             Else
             Grid1.Cell(LINEA, 11).text = "0"
             
             End If
             
             If leerdevoluciones(rutreal, localfiltro) = True Then
             Grid1.Cell(LINEA, 12).text = "1"
             If Check1.Value = "1" Then
                Call grabardevoluciones(LINEA)
                OTROS = totalotros(resultados(1), localfiltro)
                Grid1.Cell(LINEA, 8).text = OTROS
             If leerdevoluciones(rutreal, localfiltro) = False Then
             Grid1.Cell(LINEA, 12).text = "0"
             End If
             
             End If
             
             
             Else
             Grid1.Cell(LINEA, 12).text = "0"
             End If
             
             Grid1.Cell(LINEA, 13).text = contabilizada
             Grid1.Cell(LINEA, 14).text = resultados(6)
            
             
             If IsNull(resultados(7)) = True Then
                 plazo = leerplazo(rutreal)
         
             Grid1.Cell(LINEA, 15).text = Format(leevencimiento(resultados(1), plazo, Format(resultados(11), "yyyy-mm-dd")), "dd-mm-yyyy")
             fila = LINEA
             Call modificaordenpago(Grid1.Cell(fila, 2).text, localfiltro, Grid1.Cell(fila, 14).text, Format(Grid1.Cell(fila, 15).text, "yyyy-mm-dd"), Grid1.Cell(fila, 16).text)

             End If
               If saldocuentacorriente("23100027", rutreal, Format(fechasistema, "yyyy"), empresaactiva) <> 0 Then
                      Grid1.Cell(LINEA, 19).text = "1"
                 End If
                 
             If IsNull(resultados(7)) = False Then
             Grid1.Cell(LINEA, 15).text = resultados(7)
             End If
             glosa = ""
             If DIFERENCIA <> 0 Then
             glosa = "DIFERENCIA POR ACLARAR"
             If resultados("tipoglosa") = "1" Then glosa = "DIFERENCIA PRECIOS"
             If resultados("tipoglosa") = "2" Then glosa = "DIFERENCIA CANTIDADES"
             End If
             If resultados("glosadiferencia") <> "" Then
             glosa = resultados("glosadiferencia")
             End If
             Grid1.Cell(LINEA, 16).text = glosa
             End If
                    
'             If IsNull(resultados(8)) = False Then
'                     Grid1.Cell(linea, 17).text = resultados(8)
'             End If
            If Grid1.Cell(LINEA, 12).text = "1" Then
            
            End If
            
            If Grid1.Cell(LINEA, 12).text = "0" And Grid1.Cell(LINEA, 13).text = "1" Then
            Rem GRID1.Cell(linea, 14).text = "1"
            End If
            
              
 
            
            resultados.MoveNext
                        
            Wend
      End If
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      
      
End Sub
Sub limpiar()


End Sub

Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ComboLOCAL.text
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub


Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub



Private Sub Grid1_Click()
Dim fila As Double
fila = Grid1.ActiveCell.row
If Grid1.ActiveCell.col = 14 Then
If Grid1.Cell(Grid1.ActiveCell.row, 13).text = "0" Then
Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = "0"
End If
End If
If Grid1.ActiveCell.col = 14 And Grid1.Cell(fila, 12).text = "1" Then
Call MsgBox("debe rebajar devoluciones antes cancelar")
Grid1.Cell(fila, 14).text = "0"
End If
If Grid1.ActiveCell.col = 14 And Grid1.Cell(fila, 19).text = "1" Then
    Call MsgBox("PROVEEDOR TIENE UN ANTICIPO PENDIENTE POR REBAJAR RECTIFIQUE ANTES DE CANCELAR ")
    Grid1.Cell(fila, 14).text = "0"
    Exit Sub
End If
If Grid1.ActiveCell.col = 14 And Grid1.Cell(fila, 15).text = "" Then
Call MsgBox("debe colocar fecha de pago antes de autorizar")
Grid1.Cell(fila, 14).text = "0"
End If
If Grid1.ActiveCell.col = 14 Or Grid1.ActiveCell.col = 15 Or Grid1.ActiveCell.col = 16 Then
If Grid1.Cell(fila, 14).text = "0" Then
Grid1.Cell(fila, 15).text = ""
End If
Call modificaordenpago(Grid1.Cell(fila, 2).text, localfiltro, Grid1.Cell(fila, 14).text, Format(Grid1.Cell(fila, 15).text, "yyyy-mm-dd"), Grid1.Cell(fila, 16).text)
End If
If Grid1.ActiveCell.col = 12 Then
prove0005.dato3.text = Mid(Grid1.Cell(fila, 3).text, 1, 9)
prove0005.DV.Caption = Mid(Grid1.Cell(fila, 3).text, 11, 1)
Call prove0005.LEERGUIAS
prove0005.Show vbModal

End If
If Grid1.ActiveCell.col = 11 Then
publi0003.dato3.text = Mid(Grid1.Cell(fila, 3).text, 1, 9)
publi0003.DV.Caption = Mid(Grid1.Cell(fila, 3).text, 11, 1)
Call publi0003.LEERGUIAS
publi0003.Show
End If
End Sub

Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col = 2 Then
localorden = localfiltro
compra02.dato1.text = Grid1.Cell(Grid1.ActiveCell.row, 2).text
compra02.Show vbModal
End If
If Grid1.ActiveCell.col = 8 Then
localorden = localfiltro
ordenotros.montoorden.Caption = Format(CDbl(Grid1.Cell(Grid1.ActiveCell.row, 9).text), "###,###,###,###")
ordenotros.numero = Grid1.Cell(Grid1.ActiveCell.row, 2).text
ordenotros.LBLPROVEEDOR.Caption = Grid1.Cell(Grid1.ActiveCell.row, 3).text + " " + Grid1.Cell(Grid1.ActiveCell.row, 4).text
ordenotros.Show vbModal
End If
'If Grid1.ActiveCell.col = 10 Then
'localorden = localfiltro
'
'GLOSARECEPCION.Show vbModal
'
'End If
'
'

End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub



Public Function leefactura(tipo, numero, rut) As Boolean
    Dim ABONO As Double
    Dim tipo2 As String
    Dim fech As String
    tipo2 = tipo
    CAMPOS(0, 0) = "tipo"
    CAMPOS(1, 0) = "numero"
    CAMPOS(2, 0) = "total"
    CAMPOS(3, 0) = "abono"
    CAMPOS(4, 0) = "fecha"
    CAMPOS(5, 0) = ""
    
    If tipo = "FA" Then tipo = "1"
    If tipo = "ND" Then tipo = "2"
    If tipo = "NC" Then tipo = "3"
    If tipo = "FAE" Then tipo = "4"
    If tipo = "NDE" Then tipo = "5"
    If tipo = "NCE" Then tipo = "6"
    If tipo = "FC" Then tipo = "7"
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    CAMPOS(0, 2) = "facturasdecompras"
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    contabilizada = False
    pagada = False
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    
    leefactura = True
    If tipo = "1" Or tipo = "4" Then tipo2 = "FC"
    ABONO = leerabonofactura(tipo2, numero, rut, CUENTAPROVEEDOR, "D", sqlconta.response(4, 3))
    
    If ABONO <> 0 Then
    pagada = True
    End If
    Else
    pagada = True
    End If
    
        

End Function

Public Function LEERFOLIOCE(tipo) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = db
            csql.sql = "select max(numero) from movimientoscontables where mes = '" & Format(Format(fechasistema, "mm"), "00") & "' AND año = '" & Format(fechasistema, "yyyy") & "' and tipo='" + tipo + "' "
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
        If IsNull(resultados(0)) = False Then
        LEERFOLIOCE = Format(resultados(0) + 1, "0000000000")
        Else
        LEERFOLIOCE = Format(1, "0000000000")
        End If
        
    End If
    
End Function
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, NumeroDocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim CAMPOS(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    CAMPOS(0, 0) = "tipo"
    CAMPOS(1, 0) = "numero"
    CAMPOS(2, 0) = "linea"
    CAMPOS(3, 0) = "fecha"
    CAMPOS(4, 0) = "codigocuenta"
    CAMPOS(5, 0) = "tipoctacte"
    CAMPOS(6, 0) = "rutctacte"
    CAMPOS(7, 0) = "centrocosto"
    CAMPOS(8, 0) = "glosacontable"
    CAMPOS(9, 0) = "tipodocumento"
    CAMPOS(10, 0) = "numerodocumento"
    CAMPOS(11, 0) = "fechadocumento"
    CAMPOS(12, 0) = "fechavencimiento"
    CAMPOS(13, 0) = "monto"
    CAMPOS(14, 0) = "dh"
    CAMPOS(15, 0) = "creadopor"
    CAMPOS(16, 0) = "mes"
    CAMPOS(17, 0) = "año"
    CAMPOS(18, 0) = "fechacreacion"
    CAMPOS(19, 0) = "horacreacion"
    CAMPOS(20, 0) = "rutproveedor"
    CAMPOS(21, 0) = ""
    
    CAMPOS(0, 1) = tipo
    CAMPOS(1, 1) = numero
    CAMPOS(2, 1) = LINEA
    CAMPOS(3, 1) = Format(fecha, "yyyy-mm-dd")
    CAMPOS(4, 1) = codigocuenta
    CAMPOS(5, 1) = tipoctacte
    CAMPOS(6, 1) = rutctacte
    CAMPOS(7, 1) = centrocosto
    CAMPOS(8, 1) = glosacontable
    CAMPOS(9, 1) = tipodocumento
    CAMPOS(10, 1) = NumeroDocumento
    CAMPOS(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    CAMPOS(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    CAMPOS(13, 1) = monto

    CAMPOS(14, 1) = DH
    CAMPOS(15, 1) = creadopor
    CAMPOS(16, 1) = MES
    CAMPOS(17, 1) = año
    
    CAMPOS(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    CAMPOS(19, 1) = horacreacion
    CAMPOS(20, 1) = rutproveedor

    CAMPOS(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub

Public Function LEERcompras(ORDEN) As Double
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim total As Double
    Dim multi As Double
    Dim pasada As String
    
    
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT tipo,total,rut,numero,ordendecompra "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_detalle_facturas_" + localfiltro + " WHERE ordendecompra='" + ORDEN + "' and (tipo='FA' or tipo='FAE' or tipo='OE' or tipo='FC' or tipo='NCE') "
        csql.sql = csql.sql + "ORDER BY ordendecompra "
        csql.Execute
        total = 0
        contabilizada = "0"
        montocarne = 0
        montoharina = 0
        rutreal = ""
        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
            rutreal = resultados(2)
            While Not resultados.EOF
               
               If resultados(0) = "NCE" Or resultados(0) = "NC" Then multi = -1 Else multi = 1
               total = total + (resultados(1) * multi)
               
               montocarne = montocarne + LEERMONTOIMPUESTO(resultados(0), resultados(3), ORDEN, "11400012")
               montoharina = montoharina + LEERMONTOIMPUESTO(resultados(0), resultados(3), ORDEN, "11400005")
             Rem  If resultados(2) = "0802034008" Then Stop
               If leefactura(resultados(0), resultados(3), resultados(2)) = False Then
               
               pasada = "1"
               Else
               contabilizada = "1"
               End If
               
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        LEERcompras = total
        If pasada = "1" Then
        pasada = ""
        contabilizada = "0"
        End If
        
End Function
Sub GRABARCOMPROBANTE(ORDEN, montocheque, DIFERENCIA, fechacheque, TIPOPAGO, glosadiferencia)
    Dim DH As String
    Dim numero As String
    Dim LINEA As Double
    Dim fecha As Date
    Dim rut As String
    Dim tipodocumento As String
    Dim NumeroDocumento As String
    Dim fechadocumento As String
    Dim fechavencimiento As String
    Dim MES As String
    Dim año As String
    Dim monto As Double
    Dim CUENTABANCO As String
   
    Dim tipo2 As String
    Dim TIPO3 As String
    Dim DOCUMENTOPAGO As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery

        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT tipo,total,rut,numero,fecha "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_detalle_facturas_" + localfiltro + " WHERE ordendecompra='" + ORDEN + "' and total<>0 "
        csql.sql = csql.sql + "ORDER BY ordendecompra "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            DH = "D"
            If resultados(0) = "NC" Or resultados(0) = "NCE" Then
            DH = "H"
            End If
            fecha = Format(fechasistema, "yyyy-mm-dd")
            rut = resultados(2)
            tipodocumento = resultados(0)
            TIPO3 = resultados(0)
            If tipodocumento = "FA" Then tipo2 = "1": TIPO3 = "FC"
            If tipodocumento = "ND" Then tipo2 = "2": TIPO3 = "ND"
            If tipodocumento = "NC" Then tipo2 = "3": TIPO3 = "NC"
            If tipodocumento = "FAE" Then tipo2 = "4": TIPO3 = "FC"
            If tipodocumento = "NDE" Then tipo2 = "5": TIPO3 = "ND"
            If tipodocumento = "NCE" Then tipo2 = "6": TIPO3 = "NC"
            If Mid(TIPOPAGO, 1, 2) = "CH" Then
            DOCUMENTOPAGO = "PA"
            Else
            DOCUMENTOPAGO = "DB"
            End If
            
            tipodocumento = TIPO3
            NumeroDocumento = resultados(3)
            fechadocumento = resultados(4)
            fechavencimiento = fechadocumento
            MES = Format(fechasistema, "mm")
            año = Format(fechasistema, "yyyy")
            monto = resultados(1)
            
            If DH = "D" Then
            TOTALCheque = TOTALCheque + monto
            Else
            TOTALCheque = TOTALCheque - monto
            End If
            
             lineacontable = lineacontable + 1
            Call grabarcomprobante_lineas(DOCUMENTOPAGO, numerocontable, lineacontable, fecha, CUENTAPROVEEDOR, " ", rut, " ", "CANCELA DOCUMENTO", tipodocumento, NumeroDocumento, fechadocumento, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rut)
            
            Call abonofactura(tipo2, NumeroDocumento, rut, monto)
           
            resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
   Rem graba linea  diferencia
        
        If DIFERENCIA <> 0 Then
           
        lineacontable = lineacontable + 1
            
            monto = DIFERENCIA
            DH = "H"
            If DH = "D" Then
            TOTALCheque = TOTALCheque + monto
            Else
            TOTALCheque = TOTALCheque - monto
            End If
            
            Call grabarcomprobante_lineas(DOCUMENTOPAGO, numerocontable, lineacontable, fecha, cuentadiferencia, " ", rut, " ", glosadiferencia, "OC", ORDEN, fechadocumento, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rut)
        End If
        lineafinal = lineacontable
        Call leerotros(DOCUMENTOPAGO, numerocontable, fecha, localfiltro, rut, ORDEN, lineacontable, MES, año)
        
        lineacontable = lineafinal
End Sub
Sub grabarcheque(montocheque As Double)
Dim tipodocumento As String
Dim NumeroDocumento As String
Dim CUENTABANCO As String
Dim fechavencimiento As String
Dim monto As Double
Dim DH As String



    Rem graba cheque
        
        NCHEQUE = NCHEQUE + 1
        lineacontable = lineacontable + 1
        
        If tipocontable = "PA" Then
            tipodocumento = "CH"
            NumeroDocumento = Format(NCHEQUE, "0000000000")
            CUENTABANCO = dato1.text + dato2.text + dato3.text
            fechavencimiento = fechacheque
            monto = montocheque
            
            
            Else
            
            tipodocumento = "DB"
            NumeroDocumento = Format(numerocontable, "0000000000")
            CUENTABANCO = "11130001"
            fechavencimiento = fechacheque
            monto = montocheque
            
        End If
        
        DH = "H"
        Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, FECHACONTABLE, CUENTABANCO, " ", "", " ", NOMBREGIRADO, tipodocumento, NumeroDocumento, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), Format(Date, "yyyy-mm-dd"), Time, rutcontable)
        If tipocontable = "PA" Then
        fecha = Format(fechasistema, "yyyy-mm-dd")
        Call grabacheque(CUENTABANCO, NumeroDocumento, fecha, monto, fechavencimiento, "PA", numerocontable, NOMBREGIRADO, "0")
        End If
End Sub

Public Sub leerotros(tipo, numero, fecha, loc, rut, ORDEN, LINEA, MES, año)
        Dim tipocontable As String
        Dim numerocontable As String
        
        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        
        tabla = "SELECT cuenta,glosa,monto,dh,tipodo,numerodo,rut "
        tabla = tabla & "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_anexopagos_" + loc + " "
        tabla = tabla & "WHERE numero= '" & ORDEN & "' ORDER BY linea asc "
        sql.sql = tabla
        sql.Execute
        
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            While Not resultados.EOF
                LINEA = LINEA + 1
                If leerNombreCuentaMayor(resultados(0), 1) = "" Then
                rut = ""
                Else
                rut = resultados(6)
                End If
                
                If resultados(3) = "D" Then
                TOTALCheque = TOTALCheque + resultados(2)
                Else
                TOTALCheque = TOTALCheque - resultados(2)
                End If
                tipocontable = tipo
                numerocontable = numero
                If resultados(4) = "DM" Or resultados(4) = "D1" Then
                tipocontable = resultados(4)
                numerocontable = resultados(5)
                Call abonoGUIADEVOLUCION(tipocontable, numerocontable, tipo, numero, fecha, resultados(2))
                
                End If
                If resultados(4) = "1" Then
                tipocontable = "FP"
                numerocontable = resultados(5)
                
                Call abonopublicidad(resultados(4), resultados(5), resultados(2))
                
                End If
                
                If resultados(4) = "2" Then
                tipocontable = "FA"
                numerocontable = resultados(5)
                
                Call abonopublicidad(resultados(4), resultados(5), resultados(2))
                
                End If
                
                Call grabarcomprobante_lineas(tipo, numero, LINEA, fecha, resultados(0), " ", rut, " ", resultados(1), tipocontable, numerocontable, fecha, fecha, resultados(2), resultados(3), USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rut)
                
                resultados.MoveNext
            Wend
        
        End If
    lineafinal = LINEA
    
    
    End Sub



Sub grabacheque(cuenta, numero, emision, monto, vencimiento, tipocomprobante, numerocomprobante, giradoa, ubicacion)
    CAMPOS(0, 0) = "cuenta"
    CAMPOS(1, 0) = "numero"
    CAMPOS(2, 0) = "emision"
    CAMPOS(3, 0) = "monto"
    CAMPOS(4, 0) = "vencimiento"
    CAMPOS(5, 0) = "tipocomprobante"
    CAMPOS(6, 0) = "numerocomprobante"
    CAMPOS(7, 0) = "giradoa"
    CAMPOS(8, 0) = "ubicacion"
    CAMPOS(9, 0) = ""
    
    CAMPOS(0, 1) = cuenta
    CAMPOS(1, 1) = numero
    CAMPOS(2, 1) = emision
    CAMPOS(3, 1) = monto
    CAMPOS(4, 1) = vencimiento
    CAMPOS(5, 1) = tipocomprobante
    CAMPOS(6, 1) = numerocomprobante
    CAMPOS(7, 1) = giradoa
    CAMPOS(8, 1) = "0"
    CAMPOS(0, 2) = "chequesdocumento"
       
    op = 2
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
End Sub


Sub abonofactura(tipo, numero, rut, monto)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = db
    csql.sql = "update facturasdecompras set abono = abono + " & monto & " "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and rut='" + rut + "' and numero='" + numero + "'"
    csql.Execute
'    Call sincronizadatos(csql.sql, db, "")
    
    csql.Close
    Set csql = Nothing
End Sub
Sub abonoGUIADEVOLUCION(tipo, numero, TIPOCO, NUMEROCO, fechaco, montoco)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = db
    csql.sql = "update devoluciones_proveedores set tipoco='" & TIPOCO & "',numeroco='" & NUMEROCO & "',fechaco='" & Format(fechaco, "yyyy-mm-dd") & "',montoco='" & montoco & "' "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and numero='" + numero + "'"
    csql.Execute
    Call sincronizadatos(csql.sql, db, "")
    
    csql.Close
    Set csql = Nothing
End Sub
Sub abonopublicidad(tipo, numero, montoco)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = db
    
    csql.sql = "update facturasdepublicidad set abono=abono+'" & montoco & "' "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and numero='" + numero + "'"
    csql.Execute
    Call sincronizadatos(csql.sql, db, "")
    
    csql.Close
    Set csql = Nothing
End Sub


    Private Sub dato1_GotFocus()
        Call cargatexto(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
'****************************************************************************
'GOTFOCUS
'****************************************************************************

'****************************************************************************
'KEYDOWN
'****************************************************************************
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 38 Then Unload Me: GoTo no:
        If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
        Call flechas(dato1, dato2, KeyCode)
no:
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato1, dato3, KeyCode)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato2, dato4, KeyCode)
    End Sub
'*********************************************
'KEYDOWN
'****************************************************************************

'****************************************************************************
'KEYPRESS
'****************************************************************************
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
           
          dato2.SetFocus
          
           
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
           dato3.SetFocus
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            LBLBANCO.Caption = leerNombreCuentaMayor(dato1.text & dato2.text & dato3.text, 3)
            If LBLBANCO.Caption <> "" Then
                
            dato4.SetFocus
            End If
        
        End If
    End Sub

Sub ayudamayor(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "' AND banco='1'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    basebus = clientesistema + "conta" + empresaactiva
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", PIVOTE, CAMPOS, cfijo, largo, 2)
    If Val(PIVOTE.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(PIVOTE.text, 1, 2)
    dato2.text = Mid(PIVOTE.text, 3, 2)
    dato3.text = Mid(PIVOTE.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
no:
End Sub



Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
If (col = 14 Or col = 15 Or col = 16) And row <> NewRow Then
Call modificaordenpago(Grid1.Cell(row, 2).text, localfiltro, Grid1.Cell(row, 14).text, Grid1.Cell(row, 15).text, Grid1.Cell(row, 16).text)

End If

End Sub

'Sub ayudamayor(ByRef caja As TextBox)
'    Dim campos As Variant
'    Dim cfijo As Variant
'    Dim largo As Variant
'    campos = Array("codigo", "nombre")
'    largo = Array("8s", "40s")
'    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
'    cabezas = Array("cuenta", "nombre")
'    mensajeAyuda = "Ayuda tipo de Cuentas mayor"
'
'    Call cargaAyudaT(servidor, basebus, usuario, password, "cuentasdelmayor", dato1, campos, cfijo, largo, 2)
'    caja.Enabled = True
'    caja.SetFocus
'End Sub
Private Sub ORDEN_GotFocus()
Call cargatexto(ORDEN)
End Sub

Private Sub ORDEN_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
If Option3.Value = False Then

Call ceros(ORDEN)
Call BUSCAR_Click
Else
Call ceros(ORDEN)
Call Command2_Click

End If

End If

End Sub

Private Sub BUSCAR_Click()
 Dim i As Integer
 
  For i = 1 To Grid1.Rows - 1
            If Mid(Grid1.Cell(i, 2).text, 1, 10) = ORDEN.text Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
            End If
        Next i
End Sub

Private Function LEERMONTOIMPUESTO(tipo, numero, ORDEN, cuenta) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
 
        Set csql.ActiveConnection = gestionrubro

            csql.sql = "select monto from " + clientesistema + "gestion" + rubro + ".l_ordendecompra_impuestos_" + localfiltro + " where cuenta = '" & cuenta & "' and tipo='" + tipo + "' and numero='" + numero + "' and numeroorden='" + ORDEN + "' "
            
            csql.Execute
    LEERMONTOIMPUESTO = 0
    If csql.RowsAffected > 0 Then
    
    Set resultados = csql.OpenResultset
    LEERMONTOIMPUESTO = resultados(0)
    
    End If
    
End Function

Public Function leerlineas(ORDEN) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery

        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT tipo,total,rut,numero,fecha "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + rubro + ". l_ordendecompra_detalle_facturas_" + localfiltro + " WHERE ordendecompra='" + ORDEN + "' "
        csql.sql = csql.sql + "ORDER BY ordendecompra "
        csql.Execute
        leerlineas = 0
        If csql.RowsAffected > 0 Then
        leerlineas = csql.RowsAffected + leerlineasotros(ORDEN) + 1
        End If
        
        
End Function

Public Function leerlineasotros(ORDEN) As Double


        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        
        tabla = "SELECT cuenta,glosa,monto,dh "
        tabla = tabla & "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_anexopagos_" + localfiltro + " "
        tabla = tabla & "WHERE numero= '" & ORDEN & "' ORDER BY linea asc "
        sql.sql = tabla
        sql.Execute
        leerlineasotros = 0
        If sql.RowsAffected > 0 Then
        leerlineasotros = sql.RowsAffected
        
                End If
    
    
    End Function

Sub grabardevoluciones(LINEA)
        Dim suma As Double
        Dim rutpro As String
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        Dim LINEAS As Double
        Dim glosa_guia_sii As String
        Dim tabla As String
Set csql.ActiveConnection = db
rutpro = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
csql.sql = "select dp.fecha,dp.numero,dp.monto,dp.tipo,dp.numero "
csql.sql = csql.sql & " from devoluciones_proveedores as dp "
csql.sql = csql.sql & " left join cuentascorrientes as cc on (dp.rut=cc.rut "
csql.sql = csql.sql & " and cc.tipo='" + CUENTAPROVEEDOR + "'"
csql.sql = csql.sql & " AND cc.año='" + Format(fechasistema, "yyyy") + "') "
csql.sql = csql.sql & " where dp.rut='" & rutpro & "' and dp.montoco='0' "
csql.sql = csql.sql & " and local='" + localfiltro + "' "
csql.sql = csql.sql & " order by cc.nombre "
csql.Execute

If csql.RowsAffected > 0 Then
  LINEAS = 0
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        If guiarebajada(resultados(3), resultados(4), clientesistema + "gestion" + leerdatoslocal(localfiltro, "rubro") + ".l_ordendecompra_anexopagos_" + localfiltro) = False Then
    
        LINEAS = LINEAS + 1
        If resultados("tipo") = "D1" Then
            glosa_guia_sii = "GD FOLIO SII " & LeerFolioGuiaSII("D1", resultados(4), leerdatoslocal(localfiltro, "rubro")) & " FECHA " & Format(resultados(0), "dd-mm-yyyy")
        Else
            glosa_guia_sii = "GUIA DEVOLUCION " & resultados(1) & " DEL " & Format(resultados(0), "dd-mm-yyyy")
        End If
        
        Call grabarEspecialesguias(Grid1.Cell(LINEA, 2).text, LINEAS, "11200044", glosa_guia_sii, resultados(2), "H", resultados(3), resultados(4))
       End If
        resultados.MoveNext
    Wend
End If

csql.Close
Set csql = Nothing
Set resultados = Nothing

End Sub

Public Sub grabarEspecialesguias(numero, LINEA, cuenta, glosa, monto, DH, TIPODO, numerodo)
        Dim condicion As String
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        CAMPOS(0, 0) = "numero"
        CAMPOS(1, 0) = "linea"
        CAMPOS(2, 0) = "cuenta"
        CAMPOS(3, 0) = "glosa"
        CAMPOS(4, 0) = "monto"
        CAMPOS(5, 0) = "dh"
        CAMPOS(6, 0) = "tipodo"
        CAMPOS(7, 0) = "numerodo"
        
        CAMPOS(8, 0) = ""
        
        CAMPOS(0, 1) = numero
        CAMPOS(1, 1) = LINEA
        CAMPOS(2, 1) = cuenta
        CAMPOS(3, 1) = glosa
        CAMPOS(4, 1) = monto
        CAMPOS(5, 1) = DH
        CAMPOS(6, 1) = TIPODO
        CAMPOS(7, 1) = numerodo
        
        CAMPOS(0, 2) = clientesistema & "gestion" & rubro & ".l_ordendecompra_anexopagos_" & localfiltro
                
        
        condicion = ""
        op = 2
        sqlconta.response = CAMPOS
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
    End Sub

Public Sub grabarEspecialespublicidad(numero, LINEA, cuenta, glosa, monto, DH, TIPODO, numerodo)
        Dim condicion As String
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        CAMPOS(0, 0) = "numero"
        CAMPOS(1, 0) = "linea"
        CAMPOS(2, 0) = "cuenta"
        CAMPOS(3, 0) = "glosa"
        CAMPOS(4, 0) = "monto"
        CAMPOS(5, 0) = "dh"
        CAMPOS(6, 0) = "tipodo"
        CAMPOS(7, 0) = "numerodo"
        
        CAMPOS(8, 0) = ""
        
        CAMPOS(0, 1) = numero
        CAMPOS(1, 1) = LINEA
        CAMPOS(2, 1) = cuenta
        CAMPOS(3, 1) = glosa
        CAMPOS(4, 1) = monto
        CAMPOS(5, 1) = DH
        CAMPOS(6, 1) = TIPODO
        CAMPOS(7, 1) = numerodo
        
        CAMPOS(0, 2) = clientesistema & "gestion" & rubro & ".l_ordendecompra_anexopagos_" & localfiltro
                
        
        condicion = ""
        op = 2
        sqlconta.response = CAMPOS
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
    End Sub

Public Function LEERenlace(ORDEN) As Double
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim total As Double
    Dim multi As Double
    Dim pasada As String
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT ode.ordenconfactura,ode.ordenenlazada,oc.montorecepcionado "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_enlace_factura_" + localfiltro + " as ode left join " + clientesistema + "gestion" + rubro + ".l_ordendecompra_cabeza_" + localfiltro + " as oc on (oc.numero=ode.ordenenlazada) "
        csql.sql = csql.sql + "where ordenconfactura='" + ORDEN + "' "
        
        csql.Execute
       
        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
               
               total = total + (resultados(2))
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
       LEERenlace = total
End Function

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
