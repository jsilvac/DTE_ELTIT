VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Begin VB.Form electro06 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Compras"
   ClientHeight    =   8700
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   8325
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
   ScaleHeight     =   8700
   ScaleWidth      =   8325
   Begin XPFrame.FrameXp fechas 
      Height          =   1695
      Left            =   1800
      TabIndex        =   20
      Top             =   6960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2990
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
         Left            =   1800
         TabIndex        =   21
         Top             =   1200
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
      Height          =   6765
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11933
      BackColor       =   16761024
      Caption         =   "Lista Librod e Compras"
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Solo Las Contabilizadas Automaticas"
         Height          =   255
         Left            =   3600
         TabIndex        =   37
         Top             =   5880
         Width           =   3855
      End
      Begin CoolButtons.cool_Button COMMAND2 
         Height          =   495
         Left            =   4500
         TabIndex        =   12
         Top             =   3720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Genera Informe"
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
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
            Height          =   255
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
         Left            =   135
         TabIndex        =   1
         Top             =   6360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   1560
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
         Left            =   3240
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
            Caption         =   "A�O"
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
            Begin VB.ComboBox COMBOA�O 
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
         Left            =   240
         TabIndex        =   13
         Top             =   2760
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
         Left            =   3825
         TabIndex        =   26
         Top             =   4275
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
         Height          =   2475
         Left            =   240
         TabIndex        =   30
         Top             =   3840
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   4366
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
   End
End
Attribute VB_Name = "electro06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private tipoprove As String
Private plan(2000, 3) As Variant
Private canplan As Double
Private total(10) As Double
Private detalle(10, 10) As Double
Private TIPOS(9) As String
Private mes As String
Private a�o As String
Private totaldocumentos As Double








Private Sub COMMAND2_Click()
Dim TIMBRA As String

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(servidor, clientesistema + "conta" + DATO1.text, usuario, password)

a�o = COMBOA�O.text
mes = COMBOMES.ListIndex + 1
If Val(mes) < 10 Then mes = "0" + Mid(Str(mes), 2, 1) Else mes = Mid(Str(mes), 2, 2)

CARGAmayor
leermayor
Call CARGAGRILLA(infogrilla)
For k = 1 To 2000
plan(k, 3) = 0
Next k
For k = 1 To 10
detalle(k, 1) = 0
detalle(k, 2) = 0
detalle(k, 3) = 0
detalle(k, 4) = 0
detalle(k, 5) = 0
detalle(k, 6) = 0
detalle(k, 7) = 0
detalle(k, 8) = 0
detalle(k, 9) = 0
detalle(k, 10) = 0

Next k

Call Consulta_Informe(infogrilla)


infogrilla.Visible = True
infogrilla.Caption = "LIBRO DE COMPRAS"


grillainformes.Tag = "auxiliar05" & TIMBRA & FOLIO.text

infogrilla.Show


End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)


End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(DATO1)
    
End Sub

Sub leer()
    CAMPOS(0, 0) = "codigoempresa"
    CAMPOS(1, 0) = "nombre"
    CAMPOS(2, 0) = ""
    CAMPOS(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + DATO1.text + "'"
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.Status = 4 Then DATO1.SetFocus: GoTo no:
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

Centrar Me

Dim i As Integer
Dim k As Integer

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO"
TIPOS(4) = "FACTURAS ELECTRONICAS"
TIPOS(5) = "NOTAS DE DEBITO ELECTRONICAS"
TIPOS(6) = "NOTAS DE CREDITO ELECTRONICAS"
TIPOS(7) = "FACTURAS ACTIVO FIJO"
TIPOS(8) = "FACTURAS COMPRAS PROPIAS"
TIPOS(9) = "IMPORTACIONES"

Option1.Value = True

    
Call Conectar_BD
Call Conectarconta(servidor, clientesistema + "conta", usuario, password)
For i = 1 To 10
For k = 1 To 10
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
COMBOA�O.AddItem k
Next k
COMBOA�O.ListIndex = k - 2001
DATO1.text = empresaActiva
empresanombre.Caption = NOMBREEMPRESA
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
        If datos2.Value = False Then csql.sql = csql.sql + "fc.rut=cc.rut and cc.a�o='" + COMBOA�O.text + "' and cc.tipo='" + tipoprove + "' and a�ocontable='" + a�o + "' and mescontable='" + mes + "' order by fecha "
        If datos2.Value = True Then
        csql.sql = csql.sql + "fc.rut=cc.rut and cc.tipo='" + tipoprove + "' and cc.a�o='" + COMBOA�O.text + "'  and fc.fechadigitacion>='" + Format(desdefecha.Caption, "yyyy-mm-dd") + "' and fc.fechadigitacion<='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "' "
        If Check1.Value = "1" Then
        csql.sql = csql.sql + " and comentario='CENTRALIZACION AUTOMATICA' "
        End If
        
        csql.sql = csql.sql + " order by fecha "
        End If
        
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
'    If datos2.Value = True And resultados(3) < desdefecha.Caption Then GoTo PASO:
'    If datos2.Value = True And resultados(3) > hastafecha.Caption Then GoTo PASO:
'
         If RESUMEN1.Value = True Then
             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
             For k = 0 To 11
             infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
             
             Next k
             multi = 1
                totaldocumentos = totaldocumentos + 1
                If resultados(1) = "1" Then infogrilla.Grid1.Cell(lin, 2).text = "FA"
                If resultados(1) = "2" Then infogrilla.Grid1.Cell(lin, 2).text = "ND"
                If resultados(1) = "3" Then infogrilla.Grid1.Cell(lin, 2).text = "NC": multi = -1
                If resultados(1) = "4" Then infogrilla.Grid1.Cell(lin, 2).text = "FAE"
                If resultados(1) = "5" Then infogrilla.Grid1.Cell(lin, 2).text = "NDE"
                If resultados(1) = "6" Then infogrilla.Grid1.Cell(lin, 2).text = "NCE": multi = -1
                If resultados(1) = "7" Then infogrilla.Grid1.Cell(lin, 2).text = "FC"
                If resultados(1) = "8" Then infogrilla.Grid1.Cell(lin, 2).text = "IM"
                
                infogrilla.Grid1.Cell(lin, 7).text = resultados(6) * multi
                infogrilla.Grid1.Cell(lin, 8).text = resultados(7) * multi
                infogrilla.Grid1.Cell(lin, 9).text = resultados(8) * multi
                infogrilla.Grid1.Cell(lin, 10).text = resultados(9) * multi
                infogrilla.Grid1.Cell(lin, 11).text = resultados(10) * multi
                infogrilla.Grid1.Cell(lin, 12).text = resultados(11) * multi
                
                infogrilla.Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 9) + "-" + Mid(resultados(4), 10, 1)
                Rem If resultados(12) = "S" Then infogrilla.Grid1.Cell(lin, 2).text = infogrilla.Grid1.Cell(lin, 2).text + "E"
         
         End If
             If resultados(1) = "3" Or resultados(1) = "6" Then multi = -1 Else multi = 1
             total(1) = total(1) + resultados(6) * multi
             total(2) = total(2) + resultados(7) * multi
             total(3) = total(3) + resultados(8) * multi
             total(4) = total(4) + resultados(9) * multi
             total(5) = total(5) + resultados(10) * multi
             total(6) = total(6) + resultados(11) * multi
                          
                          Rem If resultados(1) = "7" And resultados(13) <> "S" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9):: detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 7) = detalle(1, 7) + resultados(11)
                          If resultados(1) = "1" And resultados(13) <> "S" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9):: detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 7) = detalle(1, 7) + resultados(11)
                          If resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 7) = detalle(2, 7) + resultados(11)
                          If resultados(1) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 7) = detalle(3, 7) + resultados(11)
                          If resultados(1) = "4" And resultados(13) <> "S" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + resultados(8):: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 7) = detalle(4, 7) + resultados(11)
                          If resultados(1) = "5" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(6): detalle(5, 3) = detalle(5, 3) + resultados(7):: detalle(5, 4) = detalle(5, 4) + resultados(8):: detalle(5, 5) = detalle(5, 5) + resultados(9): detalle(5, 6) = detalle(5, 6) + resultados(10): detalle(5, 7) = detalle(5, 7) + resultados(11)
                          If resultados(1) = "6" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(6): detalle(6, 3) = detalle(6, 3) + resultados(7):: detalle(6, 4) = detalle(6, 4) + resultados(8):: detalle(6, 5) = detalle(6, 5) + resultados(9): detalle(6, 6) = detalle(6, 6) + resultados(10): detalle(6, 7) = detalle(6, 7) + resultados(11)
                          If resultados(13) = "S" And resultados(1) <> "3" And resultados(1) <> "6" Then detalle(7, 1) = detalle(7, 1) + 1: detalle(7, 2) = detalle(7, 2) + resultados(6): detalle(7, 3) = detalle(7, 3) + resultados(7):: detalle(7, 4) = detalle(7, 4) + resultados(8):: detalle(7, 5) = detalle(7, 5) + resultados(9): detalle(7, 6) = detalle(7, 6) + resultados(10): detalle(7, 7) = detalle(7, 7) + resultados(11)
                          If resultados(1) = "7" Then detalle(8, 1) = detalle(8, 1) + 1: detalle(8, 2) = detalle(8, 2) + resultados(6): detalle(8, 3) = detalle(8, 3) + resultados(7):: detalle(8, 4) = detalle(8, 4) + resultados(8):: detalle(8, 5) = detalle(8, 5) + resultados(9): detalle(8, 6) = detalle(8, 6) + resultados(10): detalle(8, 7) = detalle(8, 7) + resultados(11)
                          If resultados(1) = "8" Then detalle(9, 1) = detalle(9, 1) + 1: detalle(9, 2) = detalle(9, 2) + resultados(6): detalle(9, 3) = detalle(9, 3) + resultados(7):: detalle(9, 4) = detalle(9, 4) + resultados(8):: detalle(9, 5) = detalle(9, 5) + resultados(9): detalle(9, 6) = detalle(9, 6) + resultados(10): detalle(9, 7) = detalle(9, 7) + resultados(11)
                          
             
             
'                          If resultados(12) <> "S" And resultados(13) <> "S" And resultados(1) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9):: detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 7) = detalle(1, 7) + resultados(11)
'                          If resultados(12) <> "S" And resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 7) = detalle(2, 7) + resultados(11)
'                          If resultados(12) <> "S" And resultados(1) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 7) = detalle(3, 7) + resultados(11)
'                          If resultados(12) = "S" And resultados(1) = "1" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + resultados(8):: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 7) = detalle(4, 7) + resultados(11)
'                          If resultados(12) = "S" And resultados(1) = "2" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(6): detalle(5, 3) = detalle(5, 3) + resultados(7):: detalle(5, 4) = detalle(5, 4) + resultados(8):: detalle(5, 5) = detalle(5, 5) + resultados(9): detalle(5, 6) = detalle(5, 6) + resultados(10): detalle(5, 7) = detalle(5, 7) + resultados(11)
'                          If resultados(12) = "S" And resultados(1) = "3" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(6): detalle(6, 3) = detalle(6, 3) + resultados(7):: detalle(6, 4) = detalle(6, 4) + resultados(8):: detalle(6, 5) = detalle(6, 5) + resultados(9): detalle(6, 6) = detalle(6, 6) + resultados(10): detalle(6, 7) = detalle(6, 7) + resultados(11)
'                          If resultados(13) = "S" And resultados(1) = "1" Then detalle(7, 1) = detalle(7, 1) + 1: detalle(7, 2) = detalle(7, 2) + resultados(6): detalle(7, 3) = detalle(7, 3) + resultados(7):: detalle(7, 4) = detalle(7, 4) + resultados(8):: detalle(7, 5) = detalle(7, 5) + resultados(9): detalle(7, 6) = detalle(7, 6) + resultados(10): detalle(7, 7) = detalle(7, 7) + resultados(11)
'
              Call Consultadetalle(resultados(1), resultados(2), resultados(4), infogrilla)
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
     
Call totallibro(infogrilla)
barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub

Sub totallibro(infogrilla As grillainformes)
    
    Dim TOTALge As Double
      lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 7, lin, 12).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 6).text = "TOTAL DOCUMENTOS  " & Format(totaldocumentos, "###,###,###")
        infogrilla.Grid1.Cell(lin, 7).text = total(1)
        infogrilla.Grid1.Cell(lin, 8).text = total(2)
        infogrilla.Grid1.Cell(lin, 9).text = total(3)
        infogrilla.Grid1.Cell(lin, 10).text = total(4)
        infogrilla.Grid1.Cell(lin, 11).text = total(5)
        infogrilla.Grid1.Cell(lin, 12).text = total(6)
    
    TOTALge = 0
    lin = lin + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 11
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 5).text = "Cant."
    infogrilla.Grid1.Cell(lin, 6).text = "Documentos"
    infogrilla.Grid1.Cell(lin, 7).text = "Neto"
    infogrilla.Grid1.Cell(lin, 8).text = "i.v.a"
    infogrilla.Grid1.Cell(lin, 9).text = "exento"
    infogrilla.Grid1.Cell(lin, 10).text = "diesel"
    infogrilla.Grid1.Cell(lin, 11).text = "retencion"
    infogrilla.Grid1.Cell(lin, 12).text = "total"
    
    
    
    For k = 1 To 9
    lin = lin + 1
    
    infogrilla.Grid1.Cell(lin, 6).text = TIPOS(k)
    infogrilla.Grid1.Cell(lin, 5).text = Format(detalle(k, 1), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 3), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 4), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 10).text = Format(detalle(k, 5), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 11).text = Format(detalle(k, 6), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 12).text = Format(detalle(k, 7), "###,###,##0")
    
    Next k
    
    
    
    
    
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    lin = lin + 2
    For k = 1 To canplan
    If plan(k, 3) <> 0 Then
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 5).text = plan(k, 1)
        infogrilla.Grid1.Cell(lin, 6).text = plan(k, 2)
        infogrilla.Grid1.Cell(lin, 7).text = plan(k, 3)
        TOTALge = TOTALge + plan(k, 3)
        End If
    Next k
        lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Range(lin, 6, lin, 7).Borders(cellEdgeTop) = cellThin
        
        
        
        
        
        infogrilla.Grid1.Cell(lin, 6).text = "TOTAL DETALLE"
         infogrilla.Grid1.Cell(lin, 7).text = TOTALge
               
    End Sub
    





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "FOLIO"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "RUT"
    FORMATOGRILLA(1, 6) = "PROVEEDOR"
    FORMATOGRILLA(1, 7) = "NETO"
    FORMATOGRILLA(1, 8) = "IVA"
    FORMATOGRILLA(1, 9) = "EXENTO"
    FORMATOGRILLA(1, 10) = "IMPTO DIESEL"
    FORMATOGRILLA(1, 11) = "RETENCION"
    
    FORMATOGRILLA(1, 12) = "TOTAL"
    FORMATOGRILLA(1, 13) = " CUENTA "
    FORMATOGRILLA(1, 14) = " MONTO "
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "4"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "30"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "9"
    FORMATOGRILLA(2, 12) = "9"
    FORMATOGRILLA(2, 13) = "30"
    FORMATOGRILLA(2, 14) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    FORMATOGRILLA(4, 10) = "###,###,###"
    FORMATOGRILLA(4, 11) = "###,###,###"
    FORMATOGRILLA(4, 12) = "###,###,###"
    FORMATOGRILLA(4, 14) = "###,###,###"
    
    Rem LOCCKED
    For k = 1 To 14
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 15
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
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub leermayor()
    tipoprove = CUENTAPROVEEDOR
    

    
End Sub

'Sub Consultadetalle(MES As String, a�o As String)
'Dim multi As Integer
'
'Dim resultados2 As rdoResultset
'    Dim cSql2 As New rdoQuery
'        Set cSql2.ActiveConnection = db
'        cSql2.SQL = "SELECT cuentadelmayor,dfc.tipo,sum(dfc.monto)"
'        cSql2.SQL = cSql2.SQL + "FROM facturasdecompras as fc,detallefacturasdecompra as dfc "
'        cSql2.SQL = cSql2.SQL + "where a�ocontable='" + a�o + "' and mescontable='" + MES + "'"
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
        csql.sql = csql.sql + "FROM cuentasdelmayor where a�o='" + COMBOA�O.text + "' "
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

Sub Consultadetalle(TIPO, NUMERO, rut, infogrilla As grillainformes)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set cSql2.ActiveConnection = temporal
        cSql2.sql = "SELECT cuentadelmayor,monto "
        cSql2.sql = cSql2.sql + "FROM facturasdecompras_detalle "
        cSql2.sql = cSql2.sql + "where tipo='" + TIPO + "' and numero='" + NUMERO + "' and rut='" + rut + "' order by linea "
        cSql2.Execute
        
        If cSql2.RowsAffected > 0 Then
        barra.Max = barra.Max + cSql2.RowsAffected - 1
        
        Set resultados2 = cSql2.OpenResultset
        linpaso = 0
        While Not resultados2.EOF
          
          For k = 1 To canplan
          If TIPO = 3 Or TIPO = 6 Then multi = -1 Else multi = 1
          If resultados2(0) = plan(k, 1) Then plan(k, 3) = plan(k, 3) + (resultados2(1) * multi)
          If resultados2(0) = plan(k, 1) And DETALLE1.Value = True Then
            If linpaso = 1 And cSql2.RowsAffected > 1 Then
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
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(servidor, basebus, usuario, password, "maestroempresas", DATO1, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub

Private Sub FrameXp9_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

End Sub

Private Sub RESUMEN2_Click()
DETALLE2.Value = True

End Sub
