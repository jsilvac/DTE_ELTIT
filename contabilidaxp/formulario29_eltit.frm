VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "clbutn.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form auxiliar99 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESUMEN FORMULARIO 29"
   ClientHeight    =   5190
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   5295
   Begin XPFrame.FrameXp fechas 
      Height          =   1935
      Left            =   1800
      TabIndex        =   20
      Top             =   6720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      BackColor       =   14737632
      Caption         =   "Rangos de Fecha"
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
      Alignment       =   1
      Begin CoolButtons.cool_Button command8 
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   1320
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
   End
   Begin XPFrame.FrameXp OPCIONES 
      Height          =   4965
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8758
      BackColor       =   16761024
      Caption         =   "Resumen de Formulario 29"
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
      Begin CoolButtons.cool_Button COMMAND2 
         Height          =   495
         Left            =   1320
         TabIndex        =   12
         Top             =   3600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Genera Informe"
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1095
         Left            =   7440
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         BackColor       =   16761024
         Caption         =   "Datos"
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
         Begin VB.OptionButton datos2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Fecha Digitacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton datos1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   2055
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   4080
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1095
         Left            =   8040
         TabIndex        =   3
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         BackColor       =   16761024
         Caption         =   "Resumen"
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
         Begin VB.OptionButton RESUMEN2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Resumido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton RESUMEN1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detallado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   3255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5741
         BackColor       =   16761024
         Caption         =   "Configuracion"
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
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   855
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "EMPRESA"
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
            Begin VB.TextBox DATO1 
               Height          =   285
               Left            =   240
               TabIndex        =   17
               Text            =   "01"
               Top             =   360
               Width           =   375
            End
            Begin VB.Label empresanombre 
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   840
               TabIndex        =   16
               Top             =   360
               Width           =   3255
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   855
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
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
               Height          =   315
               Left            =   240
               TabIndex        =   10
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   855
            Left            =   120
            TabIndex        =   9
            Top             =   2280
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "A?O"
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
            Begin VB.ComboBox COMBOA?O 
               Height          =   315
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   3855
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   960
         Left            =   6720
         TabIndex        =   13
         Top             =   3360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1693
         BackColor       =   16761024
         Caption         =   "Detalle Imputaciones"
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
         Begin VB.OptionButton DETALLE1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton DETALLE2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   495
            TabIndex        =   14
            Top             =   540
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   990
         Left            =   6600
         TabIndex        =   26
         Top             =   5280
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   1746
         BackColor       =   16761024
         Caption         =   "TIPO DE IMPRESION"
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
         Begin VB.OptionButton original 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Original"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   29
            Top             =   315
            Width           =   1575
         End
         Begin VB.OptionButton timbrado 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Timbrado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   28
            Top             =   630
            Width           =   1695
         End
         Begin VB.TextBox FOLIO 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2115
            MaxLength       =   8
            TabIndex        =   27
            Top             =   315
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp FrameXp10 
         Height          =   2235
         Left            =   360
         TabIndex        =   30
         Top             =   7200
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   3942
         BackColor       =   16761024
         Caption         =   "FILTROS DE IMPRESION"
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
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Importaciones"
            Height          =   375
            Left            =   45
            TabIndex        =   36
            Top             =   1820
            Width           =   2685
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Facturas de Compras"
            Height          =   375
            Left            =   45
            TabIndex        =   35
            Top             =   1530
            Width           =   2685
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Notas de Credito"
            Height          =   375
            Left            =   45
            TabIndex        =   34
            Top             =   1215
            Width           =   2910
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Todas"
            Height          =   375
            Left            =   45
            TabIndex        =   33
            Top             =   225
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Facturas"
            Height          =   375
            Left            =   45
            TabIndex        =   32
            Top             =   540
            Width           =   2055
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Notas de Debito"
            Height          =   375
            Left            =   45
            TabIndex        =   31
            Top             =   855
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   840
         TabIndex        =   37
         Top             =   4320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         BackColor       =   16744576
         Caption         =   " Mis Datos"
         BackColor       =   16744576
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
         Begin VB.CommandButton botonmisfavoritos 
            Caption         =   "Mis Favoritos"
            Height          =   255
            Left            =   2040
            TabIndex        =   39
            Top             =   280
            Width           =   1335
         End
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   280
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "auxiliar99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormatoGrilla(20, 20)
Private lin As Double
Private tipoprove As String
Private plan(2000, 3) As Variant
Private canplan As Double
Private total(10) As Double
Private totalh(10) As Double
Private remanente As Double
Private VENTASEXENTAS As Double


Private detalle(30, 10) As Double
Private TIPOS(20) As String
Private TIPOS2(20) As String

Private mes As String
Private a?o As String
Private totaldocumentos As Double








Private Sub Command2_Click()
Dim TIMBRA As String
Dim i As Integer

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(servidor, clientesistema + "conta" + DATO1.text, Usuario, password)

a?o = COMBOA?O.text
mes = COMBOMES.ListIndex + 1
If Val(mes) < 10 Then mes = "0" + Mid(Str(mes), 2, 1) Else mes = Mid(Str(mes), 2, 2)

CARGAmayor
leermayor
Call CARGAGRILLA(infogrilla)
For k = 1 To 2000
plan(k, 3) = 0
Next k
For k = 1 To 30
For i = 1 To 10
detalle(k, i) = 0
Next i
Next k
Call Consulta_Informe_ventas(infogrilla)
Call Consulta_boletas(infogrilla)
Call Consulta_honorarios(infogrilla)

Call Consulta_Informe(infogrilla)


infogrilla.Visible = True
infogrilla.Caption = "LISTA RESUMEN DE I.V.A": grillainformes.Tag = "auxiliar05" & TIMBRA & FOLIO.text

infogrilla.Show


End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)


End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(DATO1)
    
End Sub

Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + DATO1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then DATO1.SetFocus: GoTo no:
    COMBOMES.SetFocus
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub

Private Sub datos1_Click()
If datos2.Value = True Then fechas.Visible = True
If datos2.Value = False Then fechas.Visible = False

End Sub

Private Sub datos2_Click()
If datos2.Value = True Then fechas.Visible = True
If datos2.Value = False Then fechas.Visible = False

End Sub

Private Sub Form_Load()

CENTRAR Me

Dim i As Integer
Dim k As Integer

TIPOS2(1) = "FACTURAS "
TIPOS2(2) = "NOTAS DE DEBITO"
TIPOS2(3) = "NOTAS DE CREDITO FACTURAS"
TIPOS2(4) = "FACTURAS EXPORTACION"
TIPOS2(5) = "FACTURAS EXENTAS"
TIPOS2(6) = "TOTAL VENTAS FACTURAS   "
TIPOS2(7) = "BOLETAS   "
TIPOS2(8) = "NOTAS DE CREDITO BOLETAS"
TIPOS2(9) = "TOTAL VENTAS BOLETAS    "
TIPOS2(10) = "TOTAL VENTAS GENERALES   "

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO"
TIPOS(4) = "FACTURAS ELECTRONICAS"
TIPOS(5) = "NOTAS DE DEBITO ELECTRONICAS"
TIPOS(6) = "NOTAS DE CREDITO ELECTRONICAS"
TIPOS(7) = "FACTURAS ACTIVO FIJO"
TIPOS(8) = "FACTURAS COMPRAS PROPIAS"
TIPOS(9) = "IMPORTACIONES"
TIPOS(10) = "TOTAL COMPRAS "




Option1.Value = True

    
Call Conectar_BD
Call Conectarconta(servidor, clientesistema + "conta", Usuario, password)
For i = 1 To 10
For k = 1 To 30
detalle(k, i) = 0
Next k

Next i
OPCIONES.Visible = True

original.Value = True

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOA?O.AddItem k
Next k
COMBOA?O.ListIndex = k - 2001
DATO1.text = empresaactiva
empresanombre.Caption = nombreempresa
datos1.Value = True
RESUMEN1.Value = True
DETALLE1.Value = True
desdefecha.Caption = fechasistema
hastafecha.Caption = fechasistema

fechas.Visible = False

End Sub


    
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    

    Dim PASO As String
        totaldocumentos = 0
        Set csql.ActiveConnection = temporal
        csql.sql = "SELECT folio,fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,iva,exento,impuestoespecifico,retencion,total,fc.electronica,fc.activo "
        csql.sql = csql.sql + "FROM facturasdecompras as fc,cuentascorrientes as cc "
        If Option1.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo<>'' and "
        If Option2.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='4') and "
        If Option3.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='2' or fc.tipo='5') and "
        If Option4.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='3' or fc.tipo='6') and "
        If Option5.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='7' and "
        If Option6.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='8' and "
        If datos2.Value = False Then csql.sql = csql.sql + "fc.rut=cc.rut and cc.a?o='" + COMBOA?O.text + "' and cc.tipo='" + tipoprove + "' and a?ocontable='" + a?o + "' and mescontable='" + mes + "' order by fecha "
        If datos2.Value = True Then csql.sql = csql.sql + "fc.rut=cc.rut and cc.tipo='" + tipoprove + "' and fc.fechadigitacion>='" + Format(desdefecha.Caption, "yyyy-mm-dd") + "' and fc.fechadigitacion<='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "' order by fecha "
        
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
          total(7) = 0
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected
        barra.Value = 0
        Set resultados = csql.OpenResultset
        lin = 0
         While Not resultados.EOF
             If resultados(1) = "3" Or resultados(1) = "6" Then multi = -1 Else multi = 1
             total(1) = total(1) + resultados(6) * multi
             total(2) = total(2) + resultados(7) * multi
             total(3) = total(3) + resultados(8) * multi
             total(4) = total(4) + resultados(9) * multi
             total(5) = total(5) + resultados(10) * multi
             total(6) = total(6) + resultados(11) * multi
                          
                            
                          If resultados(1) = "1" And resultados(13) <> "S" Then
                          detalle(11, 1) = detalle(11, 1) + 1
                          detalle(11, 2) = detalle(11, 2) + resultados(6)
                          detalle(11, 3) = detalle(11, 3) + resultados(7)
                          detalle(11, 4) = detalle(11, 4) + resultados(8)
                          detalle(11, 5) = detalle(11, 5) + resultados(9)
                          detalle(11, 6) = detalle(11, 6) + resultados(10)
                          detalle(11, 7) = detalle(11, 7) + resultados(11)
                          End If
                          
                          If resultados(1) = "2" Then
                          detalle(12, 1) = detalle(12, 1) + 1
                          detalle(12, 2) = detalle(12, 2) + resultados(6)
                          detalle(12, 3) = detalle(12, 3) + resultados(7)
                          detalle(12, 4) = detalle(12, 4) + resultados(8)
                          detalle(12, 5) = detalle(12, 5) + resultados(9)
                          detalle(12, 6) = detalle(12, 6) + resultados(10)
                          detalle(12, 7) = detalle(12, 7) + resultados(11)
                          End If
                          
                          If resultados(1) = "3" Then
                          detalle(13, 1) = detalle(13, 1) + 1
                          detalle(13, 2) = detalle(13, 2) + resultados(6)
                          detalle(13, 3) = detalle(13, 3) + resultados(7)
                          detalle(13, 4) = detalle(13, 4) + resultados(8)
                          detalle(13, 5) = detalle(13, 5) + resultados(9)
                          detalle(13, 6) = detalle(13, 6) + resultados(10)
                          detalle(13, 7) = detalle(13, 7) + resultados(11)
                          End If
                          
                          If resultados(1) = "4" And resultados(13) <> "S" Then
                          detalle(14, 1) = detalle(14, 1) + 1
                          detalle(14, 2) = detalle(14, 2) + resultados(6)
                          detalle(14, 3) = detalle(14, 3) + resultados(7)
                          detalle(14, 4) = detalle(14, 4) + resultados(8)
                          detalle(14, 5) = detalle(14, 5) + resultados(9)
                          detalle(14, 6) = detalle(14, 6) + resultados(10)
                          detalle(14, 7) = detalle(14, 7) + resultados(11)
                          
                          End If
                          
                          If resultados(1) = "5" Then
                          detalle(15, 1) = detalle(15, 1) + 1
                          detalle(15, 2) = detalle(15, 2) + resultados(6)
                          detalle(15, 3) = detalle(15, 3) + resultados(7)
                          detalle(15, 4) = detalle(15, 4) + resultados(8)
                          detalle(15, 5) = detalle(15, 5) + resultados(9)
                          detalle(15, 6) = detalle(15, 6) + resultados(10)
                          detalle(15, 7) = detalle(15, 7) + resultados(11)
                          
                          
                          End If
                          
                          If resultados(1) = "6" Then
                          detalle(16, 1) = detalle(16, 1) + 1
                          detalle(16, 2) = detalle(16, 2) + resultados(6)
                          detalle(16, 3) = detalle(16, 3) + resultados(7)
                          detalle(16, 4) = detalle(16, 4) + resultados(8)
                          detalle(16, 5) = detalle(16, 5) + resultados(9)
                          detalle(16, 6) = detalle(16, 6) + resultados(10)
                          detalle(16, 7) = detalle(16, 7) + resultados(11)
                          
                          End If
                          
                          If resultados(13) = "S" And resultados(1) <> "3" And resultados(1) <> "6" Then
                          detalle(17, 1) = detalle(17, 1) + 1
                          detalle(17, 2) = detalle(17, 2) + resultados(6)
                          detalle(17, 3) = detalle(17, 3) + resultados(7)
                          detalle(17, 4) = detalle(17, 4) + resultados(8)
                          detalle(17, 5) = detalle(17, 5) + resultados(9)
                          detalle(17, 6) = detalle(17, 6) + resultados(10)
                          detalle(17, 7) = detalle(17, 7) + resultados(11)
                          
                          End If
                          
                          If resultados(1) = "7" Then
                          detalle(18, 1) = detalle(18, 1) + 1
                          detalle(18, 2) = detalle(18, 2) + resultados(6)
                          detalle(18, 3) = detalle(18, 3) + resultados(7)
                          detalle(18, 4) = detalle(18, 4) + resultados(8)
                          detalle(18, 5) = detalle(18, 5) + resultados(9)
                          detalle(18, 6) = detalle(18, 6) + resultados(10)
                          detalle(18, 7) = detalle(18, 7) + resultados(11)
                          
                          
                          End If
                          
                          If resultados(1) = "8" Then
                          detalle(19, 1) = detalle(19, 1) + 1
                          detalle(19, 2) = detalle(19, 2) + resultados(6)
                          detalle(19, 3) = detalle(19, 3) + resultados(7)
                          detalle(19, 4) = detalle(19, 4) + resultados(8)
                          detalle(19, 5) = detalle(19, 5) + resultados(9)
                          detalle(19, 6) = detalle(19, 6) + resultados(10)
                          detalle(19, 7) = detalle(19, 7) + resultados(11)
                          
                          End If
                          
                          
             
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            detalle(20, 1) = 0
            detalle(20, 2) = total(1)
            detalle(20, 3) = total(2)
            detalle(20, 4) = total(3)
            detalle(20, 5) = total(4)
            detalle(20, 6) = total(5)
            detalle(20, 7) = total(6)
            
     
Call totallibro(infogrilla)
barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub

Sub totallibro(infogrilla As grillainformes)
    Dim totales(20) As Double
    Dim totales2(20) As Double
    Dim i As Integer
    Dim IVAPAGAR As Double
    Dim ppm As Double
    Dim UNICO As Double
    Dim FORMU As Double
    Dim remanentemesiguiente As Double
    
    
    Dim TOTALge As Double
    infogrilla.Grid1.DefaultFont.Size = 10
    lin = 0
    TOTALge = 0
    
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 30
'    infogrilla.Grid1.Range(lin, 5, lin + 25, 12).Borders(cellEdgeTop) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 25, 12).Borders(cellEdgeLeft) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 25, 12).Borders(cellEdgeRight) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 25, 12).Borders(cellEdgeBottom) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 25, 12).Borders(cellInsideHorizontal) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 25, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 5).text = "Cant."
    infogrilla.Grid1.Cell(lin, 6).text = "Documentos"
    infogrilla.Grid1.Cell(lin, 7).text = "Neto"
    infogrilla.Grid1.Cell(lin, 8).text = "i.v.a"
    infogrilla.Grid1.Cell(lin, 9).text = "exento"
    infogrilla.Grid1.Cell(lin, 10).text = "diesel"
    infogrilla.Grid1.Cell(lin, 11).text = "retencion"
    infogrilla.Grid1.Cell(lin, 12).text = "total"
    
    
    For k = 1 To 10
Rem     If detalle(k, 2) <> 0 Then
    lin = lin + 1
            If k = 6 Or k = 9 Or k = 10 Then
            infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
            End If
    infogrilla.Grid1.Cell(lin, 6).text = TIPOS2(k)
    For i = 2 To 7
        If k = 3 Or k = 4 Or k = 8 Then
        detalle(k, i) = detalle(k, i) * -1
        End If
    
    Next i
    
        If k <> 6 And k <> 9 And k <> 10 Then
        infogrilla.Grid1.Cell(lin, 5).text = Format(detalle(k, 1), "###,###,##0")
            Else
        infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
        infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
        infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
        infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
        infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
        
        End If
    
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 3), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 4), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 10).text = Format(detalle(k, 6), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 11).text = Format(detalle(k, 7), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 12).text = Format(detalle(k, 8), "###,###,##0")
    totales(1) = totales(1) + detalle(k, 1)
    totales(2) = totales(2) + detalle(k, 2)
    totales(3) = totales(3) + detalle(k, 3)
    totales(4) = totales(4) + detalle(k, 4)
    totales(5) = totales(5) + detalle(k, 5)
    totales(6) = totales(6) + detalle(k, 6)
    totales(7) = totales(7) + detalle(k, 7)
Rem     End If
    
    Next k
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 6).text = "TOTAL VENTAS EXENTAS  "
    
    infogrilla.Grid1.Cell(lin, 12).text = Format(VENTASEXENTAS, "###,###,##0")
    
    
    
    lin = lin + 2
    For k = 11 To 20
Rem     If detalle(k, 2) <> 0 Then
    lin = lin + 1
    If k = 20 Then
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    End If
    
    
    infogrilla.Grid1.Cell(lin, 6).text = TIPOS(k - 10)
    infogrilla.Grid1.Cell(lin, 5).text = Format(detalle(k, 1), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 3), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 4), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 10).text = Format(detalle(k, 5), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 11).text = Format(detalle(k, 6), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 12).text = Format(detalle(k, 7), "###,###,##0")
    
Rem    End If
    Next k
    Rem MUESTRA REMANENTE
    lin = lin + 1
    infogrilla.Grid1.Cell(lin, 6).text = "TOTAL REMANENTE ANTERIOR "
    remanente = leerremanente(DATO1.text, mes, a?o)
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(remanente, "###,###,##0")
    
    Rem MUESTRA IVA A PAGAR
    IVAPAGAR = detalle(10, 3) - detalle(20, 3) - remanente
    lin = lin + 1
     infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    If IVAPAGAR < 0 Then
        infogrilla.Grid1.Cell(lin, 6).text = "REMANENTE MES SIGUIENTE  "
        remanentemesiguiente = IVAPAGAR
    Else
    infogrilla.Grid1.Cell(lin, 6).text = "TOTAL I.V.A A PAGAR      "
    End If
    infogrilla.Grid1.Cell(lin, 8).text = Format(IVAPAGAR, "###,###,##0")
    If IVAPAGAR < 0 Then IVAPAGAR = 0
    ppm = detalle(9, 2) * (leerdatos(conta, "maestroempresas", "ppm", "codigoempresa='" + DATO1.text + "'") / 100)
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 6).text = "PPM A PAGAR  "
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(ppm, "###,###,##0")
    
    
    
    For k = 22 To 22
    lin = lin + 1
    
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 6).text = "BOLETAS DE HONORARIOS "
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 1), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 3), "###,###,##0")
    
    Next k
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    UNICO = leerimpuestorenta(DATO1.text, mes, a?o)
    
    infogrilla.Grid1.Cell(lin, 6).text = "IMPUESTO UNICO TRABAJADORES "
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(UNICO, "###,###,##0")
    
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 6).text = "TOTAL FORMULARIO 29 "
    FORMU = IVAPAGAR + UNICO + detalle(21, 2) + ppm
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(FORMU, "###,###,##0")
    
    
    Call grabarremanante(DATO1.text, mes, a?o, remanentemesiguiente)
    
               
    End Sub
    





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
    
    infogrilla.Grid1.DefaultFont.Name = "ARIAL"
    
    
    
    FormatoGrilla(1, 1) = ""
    FormatoGrilla(1, 2) = ""
    FormatoGrilla(1, 3) = ""
    FormatoGrilla(1, 4) = ""
    FormatoGrilla(1, 5) = ""
    FormatoGrilla(1, 6) = ""
    FormatoGrilla(1, 7) = ""
    FormatoGrilla(1, 8) = ""
    FormatoGrilla(1, 9) = ""
    FormatoGrilla(1, 10) = ""
    FormatoGrilla(1, 11) = ""
    
    FormatoGrilla(1, 12) = ""
    FormatoGrilla(1, 13) = ""
    FormatoGrilla(1, 14) = ""
    Rem LARGO DE LOS DATOS
    
    FormatoGrilla(2, 1) = "0"
    FormatoGrilla(2, 2) = "0"
    FormatoGrilla(2, 3) = "0"
    FormatoGrilla(2, 4) = "0"
    FormatoGrilla(2, 5) = "8"
    FormatoGrilla(2, 6) = "30"
    FormatoGrilla(2, 7) = "12"
    FormatoGrilla(2, 8) = "12"
    FormatoGrilla(2, 9) = "12"
    FormatoGrilla(2, 10) = "12"
    FormatoGrilla(2, 11) = "12"
    FormatoGrilla(2, 12) = "12"
    FormatoGrilla(2, 13) = "30"
    FormatoGrilla(2, 14) = "12"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FormatoGrilla(3, 1) = "N"
    FormatoGrilla(3, 2) = "S"
    FormatoGrilla(3, 3) = "S"
    FormatoGrilla(3, 4) = "S"
    FormatoGrilla(3, 5) = "N"
    FormatoGrilla(3, 6) = "S"
    FormatoGrilla(3, 7) = "N"
    FormatoGrilla(3, 8) = "N"
    FormatoGrilla(3, 9) = "N"
    FormatoGrilla(3, 10) = "N"
    FormatoGrilla(3, 11) = "N"
    FormatoGrilla(3, 12) = "N"
    FormatoGrilla(3, 13) = "S"
    FormatoGrilla(3, 14) = "N"
    Rem FORMATO GRILLA
    FormatoGrilla(4, 5) = "###,###"
    FormatoGrilla(4, 7) = "###,###,###"
    FormatoGrilla(4, 8) = "###,###,###"
    FormatoGrilla(4, 9) = "###,###,###"
    FormatoGrilla(4, 10) = "###,###,###"
    FormatoGrilla(4, 11) = "###,###,###"
    FormatoGrilla(4, 12) = "###,###,###"
    FormatoGrilla(4, 14) = "###,###,###"
    
    Rem LOCCKED
    For k = 1 To 14
    FormatoGrilla(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 13
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FormatoGrilla(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FormatoGrilla(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FormatoGrilla(2, k))
        infogrilla.Grid1.Column(k).FormatString = FormatoGrilla(4, k)
        infogrilla.Grid1.Column(k).Locked = FormatoGrilla(5, k)
        If FormatoGrilla(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FormatoGrilla(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub leermayor()
    tipoprove = CUENTAPROVEEDOR
    

    
End Sub

'Sub Consultadetalle(MES As String, a?o As String)
'Dim multi As Integer
'
'Dim resultados2 As rdoResultset
'    Dim cSql2 As New rdoQuery
'        Set cSql2.ActiveConnection = db
'        cSql2.SQL = "SELECT cuentadelmayor,dfc.tipo,sum(dfc.monto)"
'        cSql2.SQL = cSql2.SQL + "FROM facturasdecompras as fc,detallefacturasdecompra as dfc "
'        cSql2.SQL = cSql2.SQL + "where a?ocontable='" + a?o + "' and mescontable='" + MES + "'"
'        cSql2.SQL = cSql2.SQL + " and fc.tipo=dfc.tipo and fc.numero=dfc.numero and fc.rut=dfc.rut"
'        cSql2.SQL = cSql2.SQL + " group by cuentadelmayor,dfc.tipo "
'
'        cSql2.Execute
'
'
'        If cSql2.RowsAffected > 0 Then
'        Set resultados2 = cSql2.OpenResultset
'
'         While Not resultados2.EOF
'         For K = 1 To canplan
'         If resultados2(1) = "3" Then multi = -1 Else multi = 1
'         If resultados2(0) = plan(K, 1) Then plan(K, 3) = plan(K, 3) + (resultados2(2) * multi): infogrilla.Grid1.Cell(lin, 11).text = plan(K, 2): K = canplan + 1
'         Next K
'          resultados2.MoveNext
'
'
'         Wend
'
'          resultados2.Close
'
'        End If
'
'End Sub
Sub CARGAmayor()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = temporal
        csql.sql = "SELECT codigo,nombre,tipo "
        csql.sql = csql.sql + "FROM cuentasdelmayor where a?o='" + COMBOA?O.text + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        linea = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             linea = linea + 1
             plan(linea, 1) = resultados(0)
             plan(linea, 2) = resultados(1)
             plan(linea, 3) = 0

            resultados.MoveNext
            Wend
        End If
canplan = linea
   

End Sub

Sub Consultadetalle(tipo, numero, rut, infogrilla As grillainformes)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = temporal
        csql2.sql = "SELECT cuentadelmayor,monto "
        csql2.sql = csql2.sql + "FROM facturasdecompras_detalle "
        csql2.sql = csql2.sql + "where tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' order by linea "
        csql2.Execute
        
        If csql2.RowsAffected > 0 Then
        barra.Max = barra.Max + csql2.RowsAffected - 1
        
        Set resultados2 = csql2.OpenResultset
        linpaso = 0
        While Not resultados2.EOF
          
          For k = 1 To canplan
          If tipo = 3 Or tipo = 6 Then multi = -1 Else multi = 1
          If resultados2(0) = plan(k, 1) Then plan(k, 3) = plan(k, 3) + (resultados2(1) * multi)
          If resultados2(0) = plan(k, 1) And DETALLE1.Value = True Then
            If linpaso = 1 And csql2.RowsAffected > 1 Then
            lin = lin + 1: infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            End If
          
            infogrilla.Grid1.Cell(lin, 13).text = plan(k, 2): infogrilla.Grid1.Cell(lin, 14).text = resultados2(1): k = canplan + 1: linpaso = 1
          
          End If
          
            
          Next k
          resultados2.MoveNext
                

         Wend

          resultados2.Close

        End If

End Sub

Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(servidor, basebus, Usuario, password, "maestroempresas", DATO1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub

Private Sub FrameXp9_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

End Sub

Private Sub RESUMEN2_Click()
DETALLE2.Value = True

End Sub
Sub Consulta_Informe_ventas(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim FOLIO As Double
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    
    If datos1.Value = True Then
    fecha1 = a?o + "-" + mes + "-" + "01"
    fecha2 = a?o + "-" + mes + "-" + "31"
    Else
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
    
    End If
    
        Set csql.ActiveConnection = temporal
        
        csql.sql = "SELECT '',fc.tipo,numero,fecha,fc.rut,'',neto,iva,exento,'0','0',total "
        csql.sql = csql.sql + "FROM facturasdeventas as fc "
        csql.sql = csql.sql + "WHERE fc.tipo<>'' and "
        csql.sql = csql.sql + " fecha >= '" + fecha1 + "' and fecha <= '" + fecha2 + "' "
        csql.sql = csql.sql + " order by fecha,tipo,numero "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        Set resultados = csql.OpenResultset
        lin = 0
        FOLIO = CDbl(resultados(1))
     
        While Not resultados.EOF
             If resultados(1) > "2" And resultados(1) < "5" Then multi = -1 Else multi = 1
            If resultados(1) <> "3" Then
             total(1) = total(1) + resultados(6) * multi
             total(2) = total(2) + resultados(7) * multi
             total(3) = total(3) + resultados(8) * multi
             total(4) = total(4) + resultados(9) * multi
             total(5) = total(5) + resultados(10) * multi
             total(6) = total(6) + resultados(11) * multi
            End If
            
             If resultados(1) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9): detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 8) = detalle(1, 8) + resultados(11)
             If resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 8) = detalle(2, 8) + resultados(11)
             If resultados(1) = "4" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 8) = detalle(3, 8) + resultados(11)
             If resultados(1) = "5" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + resultados(8):: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 8) = detalle(4, 8) + resultados(11)
             If resultados(1) = "9" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(6): detalle(5, 3) = detalle(5, 3) + resultados(7):: detalle(5, 4) = detalle(5, 4) + resultados(8):: detalle(5, 5) = detalle(5, 5) + resultados(9): detalle(5, 6) = detalle(5, 6) + resultados(10): detalle(5, 8) = detalle(5, 8) + resultados(11)
             
             If resultados(1) = "3" Then detalle(8, 1) = detalle(8, 1) + 1: detalle(8, 2) = detalle(8, 2) + resultados(6): detalle(8, 3) = detalle(8, 3) + resultados(7):: detalle(8, 4) = detalle(8, 4) + resultados(8):: detalle(8, 5) = detalle(8, 5) + resultados(9): detalle(8, 6) = detalle(8, 6) + resultados(10): detalle(8, 8) = detalle(8, 8) + resultados(11)
             If resultados(1) = "6" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9): detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 8) = detalle(1, 8) + resultados(11)
             If resultados(1) = "7" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 8) = detalle(2, 8) + resultados(11)
             If resultados(1) = "8" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 8) = detalle(3, 8) + resultados(11)
             
             resultados.MoveNext
             
           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
     
                        detalle(6, 1) = 0
                        detalle(6, 2) = total(1)
                        detalle(6, 3) = total(2)
                        detalle(6, 4) = total(3)
                        detalle(6, 5) = total(4)
                        detalle(6, 6) = total(5)
                        detalle(6, 8) = total(6)
     

barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub

Sub Consulta_boletas(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim NETO As Double
    Dim iva As Double
    

    Dim PASO As String
        Set csql.ActiveConnection = temporal
        csql.sql = "SELECT fecha,caja,numerozeta,boletainicial,boletafinal,(boletafinal-boletainicial+1) as diferencia,round(monto/1.19),monto-round(monto/1.19),exento,total "
        csql.sql = csql.sql + "FROM boletasdeventa "
        csql.sql = csql.sql + "where mid(fecha,1,7) = '" + a?o + "-" + mes + "'  "
        csql.sql = csql.sql + "order by fecha "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
        VENTASEXENTAS = 0
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        
        Set resultados = csql.OpenResultset
        lin = 0
         While Not resultados.EOF
            If resultados(8) = 0 Then
    
             total(1) = total(1) + resultados(5)
             total(2) = total(2) + resultados(6)
             total(3) = total(3) + resultados(7)
             total(4) = total(4) + resultados(8)
             total(5) = total(5) + resultados(9)
             Else
             NETO = Round(resultados(8) / 1.19, 0)
             iva = resultados(8) - NETO
             total(1) = total(1) - 0
             total(2) = total(2) - NETO
             total(3) = total(3) - iva
             total(4) = total(4) - 0
             total(5) = total(5) - resultados(8)
             VENTASEXENTAS = VENTASEXENTAS + resultados(8)
             End If
             
        
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            detalle(7, 1) = total(1)
            detalle(7, 2) = total(2)
            detalle(7, 3) = total(3)
            detalle(7, 4) = total(4)
            detalle(7, 7) = total(5)
            
            
            detalle(9, 1) = total(1) + detalle(8, 1) * -1
            detalle(9, 2) = total(2) + detalle(8, 2) * -1
            detalle(9, 3) = total(3) + detalle(8, 3) * -1
            detalle(9, 4) = total(4) + detalle(8, 4) * -1
            detalle(9, 5) = detalle(8, 5) * -1
            detalle(9, 6) = detalle(8, 6) * -1
            detalle(9, 7) = total(5) + detalle(8, 7) * -1
                                                
            detalle(10, 1) = detalle(6, 1) + detalle(9, 1)
            detalle(10, 2) = detalle(6, 2) + detalle(9, 2)
            detalle(10, 3) = detalle(6, 3) + detalle(9, 3)
            detalle(10, 4) = detalle(6, 4) + detalle(9, 4)
            detalle(10, 5) = detalle(6, 5) + detalle(9, 5)
            detalle(10, 6) = detalle(6, 6) + detalle(9, 6)
            detalle(10, 7) = detalle(6, 7) + detalle(9, 7)
            
           
     

End Sub

Sub Consulta_honorarios(infogrilla As grillainformes)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim PASO As String
        Set csql.ActiveConnection = temporal
        csql.sql = "SELECT fc.tipo,numero,fecha,fc.rut,cc.nombre,monto,retencion,liquido "
        csql.sql = csql.sql + "FROM boletasdehonorarios as fc,cuentascorrientes as cc "
        csql.sql = csql.sql + "where fc.rut=cc.rut and cc.a?o='" + COMBOA?O.text + "' and cc.tipo='" + cuentahonorarios + "' and a?ocontable='" + a?o + "' and mescontable='" + mes + "' order by tipo,fecha "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        totalh(1) = 0
        totalh(2) = 0
        totalh(3) = 0
        totalh(4) = 0
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        Set resultados = csql.OpenResultset
        lin = 0
        While Not resultados.EOF
             totalh(1) = totalh(1) + resultados(5)
             totalh(2) = totalh(2) + resultados(6)
             totalh(3) = totalh(3) + resultados(7)
             resultados.MoveNext
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            detalle(22, 1) = 0
            detalle(22, 2) = totalh(2)
            detalle(22, 3) = 0
            
         
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

