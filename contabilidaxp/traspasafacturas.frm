VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form proceso03 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Facturas"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15885
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   697
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1059
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   27
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      BackColor       =   -2147483632
      Caption         =   " Mis Datos"
      BackColor       =   -2147483632
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
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
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
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   10350
      Left            =   120
      TabIndex        =   2
      Top             =   45
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   18256
      BackColor       =   12632256
      Caption         =   "TRASPASO DE FACTURAS DE VENTA"
      CaptionEstilo3D =   1
      BackColor       =   12632256
      ColorBarraArriba=   4210752
      ColorBarraAbajo =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command3 
         Caption         =   "Busca Folio"
         Height          =   375
         Left            =   10440
         TabIndex        =   38
         Top             =   8520
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "TRASPASA CONTABILIDAD"
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8520
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8520
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   15480
         _ExtentX        =   27305
         _ExtentY        =   1852
         BackColor       =   12632256
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   12632256
         ColorBarraArriba=   4210752
         ColorBarraAbajo =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtcaja 
            Height          =   285
            Left            =   12000
            MaxLength       =   2
            TabIndex        =   34
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Command2 
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
            Left            =   11970
            TabIndex        =   7
            Top             =   675
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
            BackColor       =   8421504
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   8421504
            ForeColor       =   65535
            ColorBarraArriba=   12632256
            ColorBarraAbajo =   4210752
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
            BackColor       =   8421504
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   8421504
            ForeColor       =   65535
            ColorBarraArriba=   12632256
            ColorBarraAbajo =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            BackColor       =   8421504
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   8421504
            ForeColor       =   65535
            ColorBarraArriba=   12632256
            ColorBarraAbajo =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox ComboLOCAL 
               Height          =   315
               Left            =   45
               TabIndex        =   14
               Top             =   270
               Width           =   4395
            End
         End
         Begin VB.Label lblcaja 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   12480
            TabIndex        =   35
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "CAJA"
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
            Height          =   255
            Left            =   11520
            TabIndex        =   33
            Top             =   400
            Width           =   495
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   135
         TabIndex        =   3
         Top             =   1485
         Width           =   15525
         _ExtentX        =   27384
         _ExtentY        =   11774
         BackColor       =   14737632
         Caption         =   "LISTADO DE FACTURAS DE VENTA EMITIDAS"
         CaptionEstilo3D =   1
         BackColor       =   14737632
         ColorBarraArriba=   8421504
         ColorBarraAbajo =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin XPFrame.FrameXp buscafolio 
            Height          =   1575
            Left            =   5040
            TabIndex        =   36
            Top             =   2280
            Visible         =   0   'False
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2778
            BackColor       =   14737632
            Caption         =   "Busca Folio"
            CaptionEstilo3D =   1
            BackColor       =   14737632
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
               BackColor       =   &H00E0E0E0&
               Caption         =   "NC"
               Height          =   195
               Index           =   1
               Left            =   600
               TabIndex        =   41
               Top             =   840
               Width           =   975
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Factura"
               Height          =   195
               Index           =   0
               Left            =   600
               TabIndex        =   40
               Top             =   480
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.CommandButton Command5 
               Caption         =   "C E R R A R"
               Height          =   255
               Left            =   600
               TabIndex        =   39
               Top             =   1200
               Width           =   3495
            End
            Begin VB.TextBox dato1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   1680
               MaxLength       =   10
               TabIndex        =   37
               Top             =   480
               Width           =   2295
            End
         End
         Begin FlexCell.Grid Grid1 
            Height          =   6360
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   15435
            _ExtentX        =   27226
            _ExtentY        =   11218
            BackColorFixed  =   4210752
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   16711680
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1170
         Left            =   7695
         TabIndex        =   15
         Top             =   9000
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   2064
         BackColor       =   14737632
         Caption         =   "Rangos de Fecha"
         CaptionEstilo3D =   1
         BackColor       =   14737632
         ColorBarraArriba=   8421504
         ColorBarraAbajo =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin CoolButtons.cool_Button command8 
            Height          =   375
            Left            =   6000
            TabIndex        =   16
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            SkinId          =   "13"
            Caption         =   "Cambia Fecha"
         End
         Begin VB.Label hastafecha 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   20
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label desdefecha 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   18
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   1935
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   930
         Left            =   120
         TabIndex        =   22
         Top             =   9000
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   1640
         BackColor       =   14737632
         Caption         =   "TIPOS"
         CaptionEstilo3D =   1
         BackColor       =   14737632
         ColorBarraArriba=   8421504
         ColorBarraAbajo =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.OptionButton optElectronicas 
            Caption         =   "ELECTRONICAS"
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
            Left            =   3120
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton optNormales 
            Caption         =   "NORMALES"
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
            Left            =   1440
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "TODAS"
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
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   8160
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Esta como Electronica"
         Height          =   375
         Left            =   5520
         TabIndex        =   32
         Top             =   9360
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  FOLIO SALTADO"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   31
         Top             =   8880
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "No esta en Dte"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   5520
         TabIndex        =   30
         Top             =   9840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Factura Electronica"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   8520
         Width           =   1455
      End
   End
End
Attribute VB_Name = "proceso03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String


Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
Call Conectarventas(Servidor, clientesistema + "ventas" + localfiltro, Usuario, password)
leer


End Sub



Private Sub Command3_Click()
buscafolio.Visible = True
dato1.SetFocus



End Sub


Private Sub Command4_Click()
    Dim k As Double
        If Verifica_FORM29(COMBOAÑO.text & "-" & COMBOMES.ListIndex + 1 & "-01", empresaactiva) = False Then
            For k = 1 To Grid1.Rows - 1
                If Grid1.Cell(k, 16).text = "0" Then
                    Call grabafactura(k)
                End If
            Next k
        Else
            MsgBox mensaje_nopermiso, vbCritical, "ATENCION"
        End If
    leer
End Sub

Private Sub Command5_Click()
buscafolio.Visible = False
dato1.text = ""
End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        'Call ceros(dato1)
        If Option1(0).Value = True Then MsgBox buscafactura("33", dato1.text)
        If Option1(1).Value = True Then MsgBox buscafactura("61", dato1.text)
        
        dato1.text = ""
    End If
End Sub

Private Sub Form_Load()
CENTRAR Me


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA

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
desdefecha.Caption = "01-" + Format(fechasistema, "mm-yyyy")
hastafecha.Caption = fechasistema


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
titulo = "LISTADO DE FACTURAS EMITIDAS " + COMBOMES.text + " " + COMBOAÑO.text
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




 
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "RUT"
    FORMATOGRILLA(1, 4) = "CLIENTE"
    FORMATOGRILLA(1, 5) = "FECHA"
    FORMATOGRILLA(1, 6) = "NETO"
    FORMATOGRILLA(1, 7) = "IVA"
    FORMATOGRILLA(1, 8) = "I.R.AZU"
    FORMATOGRILLA(1, 9) = "I.VINO "
    FORMATOGRILLA(1, 10) = "I.LICOR"
    FORMATOGRILLA(1, 11) = "I.HARINA"
    FORMATOGRILLA(1, 12) = "I.CARNE"
    FORMATOGRILLA(1, 13) = "I.NO AZU"
    FORMATOGRILLA(1, 14) = "I.CERVEZ"
    
    FORMATOGRILLA(1, 15) = "TOTAL  "
    FORMATOGRILLA(1, 16) = "CONTA"
    FORMATOGRILLA(1, 17) = "CAJERA"
    FORMATOGRILLA(1, 18) = "TIPO"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "8"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "23"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "6"
    FORMATOGRILLA(2, 8) = "6"
    FORMATOGRILLA(2, 9) = "6"
    FORMATOGRILLA(2, 10) = "6"
    FORMATOGRILLA(2, 11) = "6"
    FORMATOGRILLA(2, 12) = "6"
    FORMATOGRILLA(2, 13) = "6"
    FORMATOGRILLA(2, 14) = "6"
    
    FORMATOGRILLA(2, 15) = "8"
    FORMATOGRILLA(2, 16) = "5"
    FORMATOGRILLA(2, 17) = "2"
    FORMATOGRILLA(2, 18) = "3"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "N"
    FORMATOGRILLA(3, 16) = "S"
   
   
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    FORMATOGRILLA(4, 8) = "##,###,##0"
    FORMATOGRILLA(4, 9) = "##,###,##0"
    FORMATOGRILLA(4, 10) = "##,###,##0"
    FORMATOGRILLA(4, 11) = "##,###,##0"
    FORMATOGRILLA(4, 12) = "##,###,##0"
    FORMATOGRILLA(4, 13) = "##,###,##0"
    FORMATOGRILLA(4, 14) = "##,###,##0"
    FORMATOGRILLA(4, 15) = "##,###,##0"
    
    Rem LOCCKED
    For k = 1 To 18
        FORMATOGRILLA(5, k) = "TRUE"
    Next k
        
    
    Grid1.Cols = 19
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
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
    Grid1.Column(16).CellType = cellCheckBox
    
    
End Sub



Private Sub monto_Click()
End Sub

Private Sub leer()

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
    Dim tipodoc As String
    Dim docdte As String
    Dim foliocon As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
    
        Set csql.ActiveConnection = ventaslocal
        If localfiltro <> "52" Then
            csql.sql = "SELECT dc.tipo,dc.foliosii,dc.rut,mc.nombre,dc.fecha,dc.neto,dc.iva, "
            csql.sql = csql.sql & "dc.impuestoilarefrescos,dc.impuestoilavinos,dc.impuestoilalicores, "
            csql.sql = csql.sql & "dc.impuestoharina,dc.impuestocarne,dc.impuestoespecifico,dc.retencionparcial,dc.total,dc.caja,dc.contabilizado,dc.numero "
            csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " as dc," + clientesistema + "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql + "where dc.rut=mc.rut and mc.sucursal='0' and "
            csql.sql = csql.sql & "(dc.tipo='FV' OR dc.tipo='NB' or dc.tipo='NF') and fecha "
            csql.sql = csql.sql & "between '" + fecha1 + "' and '" + fecha2 + "' "
            If txtcaja.text <> "" And txtcaja.text <> "00" Then
                csql.sql = csql.sql & " and dc.caja='" & txtcaja.text & "' "
            End If
        Else
            csql.sql = "SELECT dc.tipo,dc.foliosii,dc.rut,ifnull(mc.nombre,'RUT NO CREADO'),dc.fecha,if(dc.tipo<>'FE',if(dc.dolar<>'0',(dc.neto*dc.dolar),dc.neto),dc.total*dc.dolar),if(dc.tipo<>'FE',if(dc.dolar<>'0',(dc.iva*dolar),dc.iva),0), "
            csql.sql = csql.sql & "dc.impuestoilarefrescos,dc.impuestoilavinos,dc.impuestoilalicores, "
            csql.sql = csql.sql & "dc.impuestoharina,dc.impuestocarne,'0','0',if(dc.dolar<>'0',(dc.total*dolar),dc.total),dc.caja,dc.contabilizado,dc.numero "
            csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " as dc left join " + clientesistema + "ventas.sv_maestroclientes as mc on "
            csql.sql = csql.sql + " dc.rut=mc.rut and mc.sucursal='0' where "
            csql.sql = csql.sql & "(dc.tipo='FV' OR dc.tipo='NB' or dc.tipo='NF' or dc.tipo='FE') and fecha "
            csql.sql = csql.sql & "between '" + fecha1 + "' and '" + fecha2 + "' "
             If txtcaja.text <> "" And txtcaja.text <> "00" Then
                csql.sql = csql.sql & " and dc.caja='" & txtcaja.text & "' "
            End If
        End If
        
        csql.sql = csql.sql + "order by dc.foliosii "
        csql.Execute
        total = 0
        total2 = 0
        
        Grid1.Rows = 1

        If csql.RowsAffected > 0 Then
            Grid1.AutoRedraw = False
            Grid1.Rows = csql.RowsAffected + 1
            barra.Max = csql.RowsAffected + 1
            barra.Value = 0
            Set resultados = csql.OpenResultset
            fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
                     
             LINEA = LINEA + 1
             tipodoc = resultados(0)
             barra.Value = barra.Value + 1
            
            If localfiltro <> "52" Then
             Grid1.Cell(LINEA, 0).text = resultados(16)
             Grid1.Cell(LINEA, 1).text = tipodoc
             Grid1.Cell(LINEA, 2).text = resultados("foliosii")
             Else
             Grid1.Cell(LINEA, 0).text = resultados(16)
             Grid1.Cell(LINEA, 1).text = tipodoc
             Grid1.Cell(LINEA, 2).text = resultados("foliosii")
             
            End If
              
             If resultados("contabilizado") <> "Z" Then
                If empresaactiva <> "15" Then
                    Grid1.Range(LINEA, 1, LINEA, Grid1.Cols - 1).BackColor = &HC0FFC0
                Else
                    If tipodoc <> "FE" Then
                        Grid1.Range(LINEA, 1, LINEA, Grid1.Cols - 1).BackColor = &HC0FFC0
                    End If
                End If
                
                Select Case tipodoc
                    Case "FV"
                        Grid1.Cell(LINEA, 1).text = "FAE"
                    Case "ND"
                        Grid1.Cell(LINEA, 1).text = "NDE"
                    Case "NB", "NF"
                        Grid1.Cell(LINEA, 1).text = "NCE"
                                 
                End Select
             End If
             
             If resultados(0) = "FE" Then
                Grid1.Cell(LINEA, 3).text = "000000001-9"
                Grid1.Cell(LINEA, 4).text = "EXTRANJERIA"
             Else
                Grid1.Cell(LINEA, 3).text = Mid(resultados(2), 1, 9) + "-" + Mid(resultados(2), 10, 1)
                Grid1.Cell(LINEA, 4).text = resultados(3)
             End If
             
             
             Grid1.Cell(LINEA, 5).text = resultados(4)
             Grid1.Cell(LINEA, 6).text = resultados(5)
             Grid1.Cell(LINEA, 7).text = resultados(6)
             Grid1.Cell(LINEA, 8).text = resultados(7)
             Grid1.Cell(LINEA, 9).text = resultados(8)
             Grid1.Cell(LINEA, 10).text = resultados(9)
             Grid1.Cell(LINEA, 11).text = resultados(10)
             Grid1.Cell(LINEA, 12).text = resultados(11)
             Grid1.Cell(LINEA, 13).text = resultados(12)
             Grid1.Cell(LINEA, 14).text = resultados(13)
             Grid1.Cell(LINEA, 15).text = resultados(14)
             Grid1.Cell(LINEA, 16).text = leefactura(LINEA)
             Grid1.Cell(LINEA, 17).text = resultados(15)
             Grid1.Cell(LINEA, 18).text = Mid(resultados(0), 2, 1)
             If resultados("contabilizado") <> "Z" Then
                If Grid1.Cell(LINEA, 1).text = "FAE" Then docdte = "33"
                If Grid1.Cell(LINEA, 1).text = "NCE" Then docdte = "61"
                If Grid1.Cell(LINEA, 1).text = "FE" Then docdte = "34"
                If empresaactiva = "15" And Grid1.Cell(LINEA, 1).text = "FE" Then docdte = "110"
                
                If numeroresolucion <> "0" Then
                Grid1.Cell(LINEA, 2).text = leerFOLIOSIIDTE(empresaactiva, tipodoc, resultados("numero"), resultados(4), resultados("caja"), localfiltro)
                End If
                If Grid1.Cell(LINEA, 2).text = "" Then
                Grid1.Cell(LINEA, 2).BackColor = vbBlue
                Grid1.Cell(LINEA, 2).text = resultados("foliosii")
                End If
            End If
            If foliocon <> CDbl(Grid1.Cell(LINEA, 2).text) Then
            foliocon = CDbl(Grid1.Cell(LINEA, 2).text)
             Grid1.Range(LINEA, 1, LINEA, Grid1.Cols - 1).BackColor = vbRed
            
            
            End If
            Rem If Val(Grid1.Cell(linea, 2).text) = 93 Then Stop
            foliocon = foliocon + 1
            
             
             
             Call leectacte(resultados(2))
            resultados.MoveNext
       
            Wend
            Grid1.AutoRedraw = True
            Grid1.Refresh
End If
      
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

Sub eliminafactura(tipo, numero)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_pagos_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, gestionrubro, "")
        

        
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'
'End If
'leer
End Sub

Sub grabafactura(LINEA)
    Dim netos As Double
    Dim DH As String
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim cuenta As String
    Dim DH2 As String
    Dim tipodoc As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "total"
    campos(9, 0) = "fechadigitacion"
    campos(10, 0) = "crcc"
    campos(11, 0) = "caja"
    campos(12, 0) = "tnc"
    campos(13, 0) = ""
    
    
    If Grid1.Cell(LINEA, 1).text = "FV" Then TIPOCON = "1": DH = "D": DH2 = "H": tipodoc = "FA"
    If Grid1.Cell(LINEA, 1).text = "NB" Then TIPOCON = "3": DH = "H": DH2 = "D": tipodoc = "NF"
    If Grid1.Cell(LINEA, 1).text = "NF" Then TIPOCON = "4": DH = "H": DH2 = "D": tipodoc = "NB"
    If Grid1.Cell(LINEA, 1).text = "FE" Then TIPOCON = "5": DH = "D": DH2 = "H": tipodoc = "FE"
    
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "6": DH = "D": DH2 = "H": tipodoc = "EF"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "7": DH = "D": DH2 = "H": tipodoc = "ED"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "8": DH = "H": DH2 = "D": tipodoc = "EC"
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(3, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
  
        campos(4, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
 
    campos(5, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 7).text, ",", ".")
    exentos = CDbl(Grid1.Cell(LINEA, 8).text) + CDbl(Grid1.Cell(LINEA, 9).text) + CDbl(Grid1.Cell(LINEA, 10).text) + CDbl(Grid1.Cell(LINEA, 11).text) + CDbl(Grid1.Cell(LINEA, 12).text) + CDbl(Grid1.Cell(LINEA, 13).text) + CDbl(Grid1.Cell(LINEA, 14).text)
    campos(7, 1) = Str(exentos)
    campos(8, 1) = Replace(Grid1.Cell(LINEA, 15).text, ",", ".")
    campos(9, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(10, 1) = leerdatoslocal(localfiltro, "codigocrcc")
    campos(11, 1) = Grid1.Cell(LINEA, 17).text
    If Grid1.Cell(LINEA, 1).text = "NCE" Then
        campos(12, 1) = Grid1.Cell(LINEA, 18).text
    End If
    
    condicion = ""
    campos(0, 2) = "facturasdeventas"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
    
    If Grid1.Cell(LINEA, 17).text = "99" Then
    cuenta = leerdatos(conta, "maestroempresas", "cuentacreditoer", "codigoempresa='" + empresaactiva + "' ")
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "001", campos(2, 1), cuenta, "", campos(4, 1), "", "CARGA DOCUMENTO VENTAS " + tipodoc, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(8, 1), DH, USUARIOSISTEMA, Format(campos(2, 1), "MM"), Format(campos(2, 1), "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "002", campos(2, 1), ivadebito, "", "", "", "CARGA IVA VENTAS " + tipodoc, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(6, 1), DH2, USUARIOSISTEMA, Format(campos(2, 1), "MM"), Format(campos(2, 1), "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    
    End If
    
    Call grabardetallefactura(LINEA, campos(2, 1), campos(11, 1))


End Sub

Sub grabardetallefactura(LINEA, fecha2, caja)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim ilas As Double
    Dim CRCC As String
    Dim DH As String
    Dim DH2 As String
    Dim tipodoc As String
    Dim fecha As Date
    fecha = Format(fecha2, "yyyy-mm-dd")
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = ""
    
'    If Grid1.Cell(LINEA, 1).text = "FV" Then TIPOCON = "1": DH = "D": DH2 = "H": tipodoc = "FA"
'    If Grid1.Cell(LINEA, 1).text = "NB" Then TIPOCON = "3": DH = "H": DH2 = "D": tipodoc = "NF"
'    If Grid1.Cell(LINEA, 1).text = "NF" Then TIPOCON = "4": DH = "H": DH2 = "D": tipodoc = "NB"
'    If Grid1.Cell(LINEA, 1).text = "FE" Then TIPOCON = "5": DH = "D": DH2 = "H": tipodoc = "FE"
'
'    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "6": DH = "D": DH2 = "H": tipodoc = "FA"
'    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "7": DH = "D": DH2 = "H": tipodoc = "ND"
'    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "8": DH = "H": DH2 = "D": tipodoc = "NC"
    
    
    If Grid1.Cell(LINEA, 1).text = "FV" Then TIPOCON = "1": DH = "D": DH2 = "H": tipodoc = "FA"
    If Grid1.Cell(LINEA, 1).text = "NB" Then TIPOCON = "3": DH = "H": DH2 = "D": tipodoc = "NF"
    If Grid1.Cell(LINEA, 1).text = "NF" Then TIPOCON = "4": DH = "H": DH2 = "D": tipodoc = "NB"
    If Grid1.Cell(LINEA, 1).text = "FE" Then TIPOCON = "5": DH = "D": DH2 = "H": tipodoc = "FE"
    
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "6": DH = "D": DH2 = "H": tipodoc = "EF"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "7": DH = "D": DH2 = "H": tipodoc = "ED"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "8": DH = "H": DH2 = "D": tipodoc = "EC"
    
    
    
    Rem  CALCULA netos
    
    lin = 3
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentaingresoventa")
    
    If caja = "99" Then
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentaingresootrasventas")
    End If
    
    
    campos(5, 1) = "INGRESOS VENTAS " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(7, 1) = DH2
    campos(8, 1) = leerdatoslocal(localfiltro, "codigocrcc")
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If Grid1.Cell(LINEA, 17).text = "99" Then
 
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", campos(8, 1), campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    End If
Rem CALCULA ILAS refrescos

    ilas = CDbl(Grid1.Cell(LINEA, 8).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 8) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
        If Format(fechasistema, "yyyy-mm-dd") < "2017-01-01" Then
            campos(4, 1) = leerdatoslocal(localfiltro, "cuentailarefrescos")
        Else
             campos(4, 1) = "23300010"
        End If
    
    campos(5, 1) = "ILA REFRESCOS " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = ""
    
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If Grid1.Cell(LINEA, 15).text = "99" Then
     Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    End If
    End If

Rem CALCULA sin azucar

    ilas = CDbl(Grid1.Cell(LINEA, 13).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 8) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    
         If Format(fechasistema, "yyyy-mm-dd") < "2017-01-01" Then
            campos(4, 1) = "11400017"
        Else
             campos(4, 1) = "23300017"
        End If
        
        
    
    
    campos(5, 1) = "ILA SIN AZUCAR  " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = ""
    
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If Grid1.Cell(LINEA, 17).text = "99" Then
     Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    End If
    End If

Rem CALCULA CERVEZAS

    ilas = CDbl(Grid1.Cell(LINEA, 14).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 8) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
   
    
     If Format(fechasistema, "yyyy-mm-dd") < "2017-01-01" Then
            campos(4, 1) = "11400014"
        Else
             campos(4, 1) = "23300014"
        End If
        
        
    campos(5, 1) = "ILA CERVEZAS " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = ""
    
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If Grid1.Cell(LINEA, 17).text = "99" Then
     Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    End If
    End If
    
    
    
Rem CALCULA ILAS VINOS

    ilas = CDbl(Grid1.Cell(LINEA, 9).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 8) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    
     If Format(fechasistema, "yyyy-mm-dd") < "2017-01-01" Then
            campos(4, 1) = leerdatoslocal(localfiltro, "cuentailavinos")
        Else
             campos(4, 1) = "23300011"
        End If
        
        
    
    campos(5, 1) = "ILA VINOS " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
     If Grid1.Cell(LINEA, 17).text = "99" Then
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
    End If
    
Rem CALCULA ILAS LICORES

    ilas = CDbl(Grid1.Cell(LINEA, 10).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 8) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
        If Format(fechasistema, "yyyy-mm-dd") < "2017-01-01" Then
           campos(4, 1) = leerdatoslocal(localfiltro, "cuentailalicores")
        Else
             campos(4, 1) = "23300013"
        End If
        
    
    campos(5, 1) = "ILA LICORES " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
     If Grid1.Cell(LINEA, 17).text = "99" Then
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
    End If
    
    
    
Rem CALCULA HARINA
    ilas = CDbl(Grid1.Cell(LINEA, 11).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "harinaventa")
    campos(5, 1) = "IMPUESTO HARINA " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
     If Grid1.Cell(LINEA, 17).text = "99" Then
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    End If
    End If

Rem CALCULA CARNE
    ilas = CDbl(Grid1.Cell(LINEA, 12).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "carneventas")
    campos(5, 1) = "IMPUESTO CARNE " + Mid(ComboLOCAL.text, 4, 30)
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If Grid1.Cell(LINEA, 17).text = "99" Then
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    End If
    End If
    
   
    
    
End Sub

Public Function leefactura(LINEA) As String

    Dim TIPOCON As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    If Grid1.Cell(LINEA, 1).text = "FV" Then TIPOCON = "1"
    If Grid1.Cell(LINEA, 1).text = "NB" Then TIPOCON = "3"
    If Grid1.Cell(LINEA, 1).text = "NF" Then TIPOCON = "4"
    If Grid1.Cell(LINEA, 1).text = "FE" Then TIPOCON = "5"
    
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "6"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "7"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "8"
    
    condicion = "tipo='" + TIPOCON + "' and numero='" + Grid1.Cell(LINEA, 2).text + "' and fecha='" + Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd") + "' "
    campos(0, 2) = "facturasdeventas"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leefactura = "1"
    Else
    leefactura = "0"
    End If
    
    

End Function

Sub leectacte(rut)
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + cuentacliente + "' and rut=" + "'" + rut + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
    Call crearcuentacorriente(rut)
    End If
    
End Sub
Sub crearcuentacorriente(rut)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion

            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            


End Sub
'cSql.SQL = "INSERT INTO l_movimientos_detalle_" & empresaactiva & " "
'            cSql.SQL = cSql.SQL & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
'            cSql.SQL = cSql.SQL & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.unidades, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, ROUND(dd.unidades / dd.cantidad, 0) "
'            cSql.SQL = cSql.SQL & "FROM " & baseVentas & rubro & ".sv_documento_detalle_" + empresaactiva + " as dd "
'            cSql.SQL = cSql.SQL & "WHERE dd.local = '" & empresaactiva & "' AND dd.tipo = '" & v.detalle.tipo & "' AND dd.numero = '" & v.detalle.numero & "'"
'            cSql.Execute
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub txtcaja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call ayudaCaja(txtcaja, Mid(ComboLOCAL.text, 1, 2))
    End If
End Sub
Sub ayudaCaja(ByRef caja As TextBox, ByVal empr As String)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("numero", "descripcion", "local")
    largo = Array("7n", "50s", "5n")
    cfijo = "local = '" & empr & "'"
    basebus = clientesistema + "ventas"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "sv_maestrodecajas", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus

End Sub
 
Private Sub txtcaja_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(txtcaja)
        lblcaja.Caption = leernombrecaja(Mid(ComboLOCAL.text, 1, 2), txtcaja.text)
        Command2.SetFocus
    End If
End Sub
Function leernombrecaja(loc, codigo) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select descripcion "
    csql.sql = csql.sql & "from " & clientesistema & "ventas.sv_maestrodecajas "
    csql.sql = csql.sql & "where local='" & loc & "'  and numero='" & codigo & "' "
    csql.Execute
    leernombrecaja = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
       leernombrecaja = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
End Function

    

Private Function buscafactura(tipo, numero) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT fecha "
        csql.sql = csql.sql + "FROM eltit_fae" & Mid(ComboLOCAL.text, 1, 2) & ".sv_dte" & Mid(ComboLOCAL.text, 1, 2) & " WHERE numero=" & numero & " and tipo='" & tipo & "'"
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            buscafactura = resultados(0)
        Else
         buscafactura = ""
        End If
       Set resultados = Nothing
        
End Function

