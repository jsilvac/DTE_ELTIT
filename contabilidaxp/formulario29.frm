VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form auxiliar99 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RESUMEN FORMULARIO 29"
   ClientHeight    =   5010
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
   ScaleHeight     =   5010
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
      Height          =   4845
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8546
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
         Left            =   1440
         TabIndex        =   12
         Top             =   3960
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
         Top             =   4440
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
               Locked          =   -1  'True
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
            Caption         =   "AÑO"
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
            Begin VB.ComboBox COMBOAÑO 
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
   End
End
Attribute VB_Name = "auxiliar99"
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
Private totalH(10) As Double
Private remanente As Double
Private VENTASEXENTAS As Double


Private detalle(40, 10) As Double
Private TIPOS(20) As String
Private TIPOS2(20) As String

Private MES As String
Private año As String
Private totaldocumentos As Double








Private Sub COMMAND2_Click()
Dim TIMBRA As String
Dim i As Integer

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)

año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

CARGAmayor
leermayor
Call CARGAGRILLA(infogrilla)
For k = 1 To 2000
    plan(k, 3) = 0
Next k
For k = 1 To 40
    For i = 1 To 10
        detalle(k, i) = 0
    Next i
Next k
Call Consulta_Informe_ventas(infogrilla)
Call Consulta_boletas(infogrilla)
Call Consulta_boletas_tbk(infogrilla)
Call Consulta_boletas_exentas(infogrilla)
'Call Consulta_boletas_exe(infogrilla)
Call Consulta_honorarios(infogrilla)

Call Consulta_Informe(infogrilla)


infogrilla.Visible = True
infogrilla.cmdcomprobante.Visible = True
infogrilla.Caption = "LISTA RESUMEN DE I.V.A": grillainformes.Tag = "auxiliar05" & TIMBRA & FOLIO.text

infogrilla.Show


End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)


End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
    
End Sub

Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
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
TIPOS2(4) = "NOTAS DE CREDITO BOLETAS"
TIPOS2(5) = "FACTURAS EXPORTACION"
TIPOS2(6) = "FACTURAS EXENTAS"
TIPOS2(7) = "TOTAL VENTAS FACTURAS   "
TIPOS2(8) = "BOLETAS AFECTAS  "
TIPOS2(9) = "BOLETAS TRANSBANK "
TIPOS2(10) = "BOLETAS EXENTAS  "
TIPOS2(11) = "TOTAL VENTAS BOLETAS    "
TIPOS2(12) = "TOTAL VENTAS GENERALES   "

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO"
TIPOS(4) = "FACTURAS ACTIVO FIJO"
TIPOS(5) = "FACTURAS COMPRAS PROPIAS"
TIPOS(6) = "COMPRAS SUPERMERCADOS "
TIPOS(7) = "IMPORTACIONES"
TIPOS(8) = "TOTAL COMPRAS "




Option1.Value = True

    
Call Conectar_BD
Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
For i = 1 To 10
For k = 1 To 30
detalle(k, i) = 0
Next k

Next i
opciones.Visible = True

original.Value = True

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
dato1.text = empresaactiva
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
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT folio,fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,IF(ivanorecuperable=1,0,iva) AS iva,exento,impuestoespecifico,retencion,total,fc.electronica,fc.activo,fc.comprasuper "
        csql.sql = csql.sql + "FROM facturasdecompras as fc,cuentascorrientes as cc "
        If Option1.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo<>'' and "
        If Option2.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='4') and "
        If Option3.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='2' or fc.tipo='5') and "
        If Option4.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='3' or fc.tipo='6') and "
        If Option5.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='7' and "
        If Option6.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='8' and "
        If datos2.Value = False Then csql.sql = csql.sql + "fc.rut=cc.rut and cc.año='" + COMBOAÑO.text + "' and cc.tipo='" + tipoprove + "' and añocontable='" + año + "' and mescontable='" + MES + "' order by fecha "
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
                          
                            
                          If (resultados(1) = "1" Or resultados(1) = "4") And resultados(13) <> "S" And resultados(14) <> "1" Then
                          detalle(13, 1) = detalle(13, 1) + 1
                          detalle(13, 2) = detalle(13, 2) + resultados(6)
                          detalle(13, 3) = detalle(13, 3) + resultados(7)
                          detalle(13, 4) = detalle(13, 4) + resultados(8)
                          detalle(13, 5) = detalle(13, 5) + resultados(9)
                          detalle(13, 6) = detalle(13, 6) + resultados(10)
                          detalle(13, 7) = detalle(13, 7) + resultados(11)
                          End If
                          
                          If resultados(1) = "2" Or resultados(1) = "5" Then
                          detalle(14, 1) = detalle(14, 1) + 1
                          detalle(14, 2) = detalle(14, 2) + resultados(6)
                          detalle(14, 3) = detalle(14, 3) + resultados(7)
                          detalle(14, 4) = detalle(14, 4) + resultados(8)
                          detalle(14, 5) = detalle(14, 5) + resultados(9)
                          detalle(14, 6) = detalle(14, 6) + resultados(10)
                          detalle(14, 7) = detalle(14, 7) + resultados(11)
                          End If
                          
                          If resultados(1) = "3" Or resultados(1) = "6" Then
                          detalle(15, 1) = detalle(15, 1) + 1
                          detalle(15, 2) = detalle(15, 2) + resultados(6)
                          detalle(15, 3) = detalle(15, 3) + resultados(7)
                          detalle(15, 4) = detalle(15, 4) + resultados(8)
                          detalle(15, 5) = detalle(15, 5) + resultados(9)
                          detalle(15, 6) = detalle(15, 6) + resultados(10)
                          detalle(15, 7) = detalle(15, 7) + resultados(11)
                          End If
                          
                          If resultados(13) = "S" Then
                          detalle(16, 1) = detalle(16, 1) + 1
                          detalle(16, 2) = detalle(16, 2) + resultados(6)
                          detalle(16, 3) = detalle(16, 3) + resultados(7)
                          detalle(16, 4) = detalle(16, 4) + resultados(8)
                          detalle(16, 5) = detalle(16, 5) + resultados(9)
                          detalle(16, 6) = detalle(16, 6) + resultados(10)
                          detalle(16, 7) = detalle(16, 7) + resultados(11)
                          
                          End If
                          
                          If resultados(1) = "7" Then
                          detalle(17, 1) = detalle(17, 1) + 1
                          detalle(17, 2) = detalle(17, 2) + resultados(6)
                          detalle(17, 3) = detalle(17, 3) + resultados(7)
                          detalle(17, 4) = detalle(17, 4) + resultados(8)
                          detalle(17, 5) = detalle(17, 5) + resultados(9)
                          detalle(17, 6) = detalle(17, 6) + resultados(10)
                          detalle(17, 7) = detalle(17, 7) + resultados(11)
                          
                          
                          End If
                          
                          
                          If resultados(14) = "1" Then
                          detalle(18, 1) = detalle(18, 1) + 1
                          detalle(18, 2) = detalle(18, 2) + resultados(6)
                          detalle(18, 3) = detalle(18, 3) + resultados(7)
                          detalle(18, 4) = detalle(18, 4) + resultados(8)
                          detalle(18, 5) = detalle(18, 5) + resultados(9)
                          detalle(18, 6) = detalle(18, 6) + resultados(10)
                          detalle(18, 7) = detalle(18, 7) + resultados(11)
                          
                          End If
                          
                          
                          If resultados(1) = "4" Then
                          detalle(40, 1) = detalle(40, 1) + 1
                          detalle(40, 2) = detalle(40, 2) + resultados(6)
                          detalle(40, 3) = detalle(40, 3) + resultados(7)
                          detalle(40, 4) = detalle(40, 4) + resultados(8)
                          detalle(40, 5) = detalle(40, 5) + resultados(9)
                          detalle(40, 6) = detalle(40, 6) + resultados(10)
                          detalle(40, 7) = detalle(40, 7) + resultados(11)
                          
                          
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
    Dim IVARETENIDOTERCEROS As Double
    Dim IVAANTICIPADO As Double
    Dim IVAFACTURACOMPRA As Double
    Dim totalhonorario As Double
    
    
    Dim TOTALge As Double
    infogrilla.Grid1.DefaultFont.Size = 10
    lin = 0
    TOTALge = 0
    
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 40
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
    
    
    For k = 1 To 12
Rem     If detalle(k, 2) <> 0 Then
    lin = lin + 1
            If k = 7 Or k = 11 Or k = 12 Then
            infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
            End If
    infogrilla.Grid1.Cell(lin, 6).text = TIPOS2(k)
    For i = 2 To 9
        If k = 3 Or k = 4 Then
        detalle(k, i) = detalle(k, i) * -1
        End If
    
    Next i
    
        If k <> 7 And k <> 11 And k <> 12 Then
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
    For k = 13 To 20
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
    
    
    infogrilla.Grid1.Cell(lin, 6).text = TIPOS(k - 12)
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
    remanente = leerremanente(dato1.text, MES, año)
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(remanente, "###,###,##0")
    
    Rem MUESTRA IVA A PAGAR
    IVAPAGAR = detalle(12, 3) - detalle(20, 3) - remanente
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
    ppm = (detalle(10, 2) + detalle(10, 4)) * (leerdatos(conta, "maestroempresas", "ppm", "codigoempresa='" + dato1.text + "'") / 100)
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 6).text = "PPM A PAGAR  " & Format((leerdatos(conta, "maestroempresas", "ppm", "codigoempresa='" + dato1.text + "'")), "%###.00")
    
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
    
    infogrilla.Grid1.Cell(lin, 6).text = "BOLETAS DE HONORARIOS (151)"
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 1), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 3), "###,###,##0")
    totalhonorario = detalle(k, 2)
    Next k
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    UNICO = leerimpuestorenta(dato1.text, MES, año)
    
    infogrilla.Grid1.Cell(lin, 6).text = "IMPUESTO UNICO (048)"
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(UNICO, "###,###,##0")
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    IVARETENIDOTERCEROS = leer_039(dato1.text, MES, año)
    
    infogrilla.Grid1.Cell(lin, 6).text = "IVA TOTAL RETENIDO (039)"
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(IVARETENIDOTERCEROS, "###,###,##0")
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    IVAANTICIPADO = leer_556(dato1.text, MES, año)
    
    infogrilla.Grid1.Cell(lin, 6).text = "IVA ANTICIPADO (556)"
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(IVAANTICIPADO, "###,###,##0")
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
'    IVAFACTURACOMPRA = infogrilla.Grid1.Cell(1151, 8).text
IVAFACTURACOMPRA = infogrilla.Grid1.Cell(11, 8).text
    
    infogrilla.Grid1.Cell(lin, 6).text = "IVA PARCIAL RET (554)"
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(IVAFACTURACOMPRA, "###,###,##0")
    
    
    
    
    lin = lin + 1
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 6).text = "TOTAL FORMULARIO 29 "
    FORMU = IVAPAGAR + UNICO + totalhonorario + ppm + IVARETENIDOTERCEROS - IVAANTICIPADO + IVAFACTURACOMPRA
    
    infogrilla.Grid1.Cell(lin, 8).text = Format(FORMU, "###,###,##0")
    
    
    For k = 40 To 40
    Rem     If detalle(k, 2) <> 0 Then
    lin = lin + 2
    infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).FontBold = True
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin, 12).Borders(cellInsideVertical) = cellThin
    
    
    infogrilla.Grid1.Cell(lin, 6).text = "DOCUMENTOS ELETRONICOS "
    infogrilla.Grid1.Cell(lin, 5).text = Format(detalle(k, 1), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 3), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 4), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 10).text = Format(detalle(k, 5), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 11).text = Format(detalle(k, 6), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 12).text = Format(detalle(k, 7), "###,###,##0")
    
Rem    End If
    Next k
    
    
    Call grabarremanante(dato1.text, MES, año, remanentemesiguiente)
    
'    Call cargarformulario29
    
               
    End Sub
    
Sub cargarformulario29()
    Load form29
    form29.Show
    form29.COMMAND2_Click
     
End Sub

 
'Sub leerposiciones(codigo)
'    Dim csql As New rdoQuery
'    Dim resultados As rdoResultset
'    Set csql.ActiveConnection = contadb
'        csql.sql = "select posicionx,posiciony,hoja from "
'        csql.sql = csql.sql & " maestro_codigof29 "
'        csql.sql = csql.sql & " where codigosii='" & codigo & "' "
'        csql.Execute
'        posx = 0
'        posy = 0
'        hoja = 0
'
'    If csql.RowsAffected > 0 Then
'        Set resultados = csql.OpenResultset
'        posx = resultados(0)
'        posy = resultados(1)
'        hoja = resultados(2)
'    End If
'    csql.Close
'    Set csql = Nothing
'
'
'End Sub


Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
    
    infogrilla.Grid1.DefaultFont.Name = "ARIAL"
    
    
    
    FORMATOGRILLA(1, 1) = ""
    FORMATOGRILLA(1, 2) = ""
    FORMATOGRILLA(1, 3) = ""
    FORMATOGRILLA(1, 4) = ""
    FORMATOGRILLA(1, 5) = ""
    FORMATOGRILLA(1, 6) = ""
    FORMATOGRILLA(1, 7) = ""
    FORMATOGRILLA(1, 8) = ""
    FORMATOGRILLA(1, 9) = ""
    FORMATOGRILLA(1, 10) = ""
    FORMATOGRILLA(1, 11) = ""
    
    FORMATOGRILLA(1, 12) = ""
    FORMATOGRILLA(1, 13) = ""
    FORMATOGRILLA(1, 14) = ""
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "0"
    FORMATOGRILLA(2, 2) = "0"
    FORMATOGRILLA(2, 3) = "0"
    FORMATOGRILLA(2, 4) = "0"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "30"
    FORMATOGRILLA(2, 7) = "12"
    FORMATOGRILLA(2, 8) = "12"
    FORMATOGRILLA(2, 9) = "12"
    FORMATOGRILLA(2, 10) = "12"
    FORMATOGRILLA(2, 11) = "12"
    FORMATOGRILLA(2, 12) = "12"
    FORMATOGRILLA(2, 13) = "30"
    FORMATOGRILLA(2, 14) = "12"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "N"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "N"
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
    FORMATOGRILLA(4, 5) = "###,###"
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

'Sub Consultadetalle(MES As String, año As String)
'Dim multi As Integer
'
'Dim resultados2 As rdoResultset
'    Dim cSql2 As New rdoQuery
'        Set cSql2.ActiveConnection = contadb
'        cSql2.SQL = "SELECT cuentadelmayor,dfc.tipo,sum(dfc.monto)"
'        cSql2.SQL = cSql2.SQL + "FROM facturasdecompras as fc,detallefacturasdecompra as dfc "
'        cSql2.SQL = cSql2.SQL + "where añocontable='" + año + "' and mescontable='" + MES + "'"
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
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo "
        csql.sql = csql.sql + "FROM cuentasdelmayor where año='" + COMBOAÑO.text + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             plan(LINEA, 1) = resultados(0)
             plan(LINEA, 2) = resultados(1)
             plan(LINEA, 3) = 0

            resultados.MoveNext
            Wend
        End If
canplan = LINEA
   

End Sub

Sub Consultadetalle(tipo, numero, rut, infogrilla As grillainformes)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
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
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub

Private Sub FrameXp9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

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
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    Else
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
    
    End If
    
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT '',fc.tipo,numero,fecha,fc.rut,'',neto,iva,exento,'0','0',total,tnc "
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
           multi = 1
           If resultados(1) = "3" Or resultados(1) = "4" Or resultados(1) = "8" Then multi = -1
            
             total(1) = total(1) + resultados(6) * multi
             total(2) = total(2) + resultados(7) * multi
             total(3) = total(3) + resultados(8) * multi
             total(4) = total(4) + resultados(9) * multi
             total(5) = total(5) + resultados(10) * multi
             total(6) = total(6) + resultados(11) * multi
            
            
             If resultados(1) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9): detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 8) = detalle(1, 8) + resultados(11)
             If resultados(1) = "6" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9): detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 8) = detalle(1, 8) + resultados(11)
             If resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 8) = detalle(2, 8) + resultados(11)
             If resultados(1) = "7" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 8) = detalle(2, 8) + resultados(11)
             If resultados(1) = "4" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 8) = detalle(3, 8) + resultados(11)
             If resultados(1) = "8" And resultados("tnc") = "F" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 8) = detalle(3, 8) + resultados(11)
             If resultados(1) = "3" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + resultados(8):: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 8) = detalle(4, 8) + resultados(11)
             If resultados(1) = "8" And resultados("tnc") = "B" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + resultados(8):: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 8) = detalle(4, 8) + resultados(11)
             If resultados(1) = "5" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(6): detalle(5, 3) = detalle(5, 3) + resultados(7):: detalle(5, 4) = detalle(5, 4) + resultados(8):: detalle(5, 5) = detalle(5, 5) + resultados(9): detalle(5, 6) = detalle(5, 6) + resultados(10): detalle(5, 8) = detalle(5, 8) + resultados(11)
             If resultados(1) = "9" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(6): detalle(6, 3) = detalle(6, 3) + resultados(7):: detalle(6, 4) = detalle(6, 4) + resultados(8):: detalle(6, 5) = detalle(6, 5) + resultados(9): detalle(6, 6) = detalle(6, 6) + resultados(10): detalle(6, 8) = detalle(6, 8) + resultados(11)
             
             resultados.MoveNext
             
           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
     
                        detalle(7, 1) = 0
                        detalle(7, 2) = total(1)
                        detalle(7, 3) = total(2)
                        detalle(7, 4) = total(3)
                        detalle(7, 5) = total(4)
                        detalle(7, 6) = total(5)
                        detalle(7, 8) = total(6)
     

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
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,caja,numerozeta,boletainicial,boletafinal,(boletafinal-boletainicial+1) as diferencia,"
        csql.sql = csql.sql & "round(monto/1.19),monto-round(monto/1.19),exento,total,cigarro "
        csql.sql = csql.sql + "FROM boletasdeventa "
        csql.sql = csql.sql + "where mid(fecha,1,7) = '" + año + "-" + MES + "' and estbk='0'   "
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
            If resultados("cigarro") = 0 And resultados("exento") > 0 Then GoTo otro:
            If resultados("cigarro") = 0 Then
    
                total(1) = total(1) + resultados(5)
                total(2) = total(2) + resultados(6)
                total(3) = total(3) + resultados(7)
                total(4) = total(4) + resultados(8)
                total(6) = total(6) + resultados(9)
                VENTASEXENTAS = VENTASEXENTAS + resultados(8)
             Else
                NETO = Round(resultados(8) / 1.19, 0)
                iva = resultados(8) - NETO
                total(1) = total(1) - 0
                total(2) = total(2) - NETO
                total(3) = total(3) - iva
                total(4) = total(4) + resultados(8)
                total(5) = total(5) - resultados(8)
                total(6) = total(6) + resultados(9)
              End If
             
otro:
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            detalle(8, 1) = total(1)
            detalle(8, 2) = total(2)
            detalle(8, 3) = total(3)
            detalle(8, 4) = total(4)
            detalle(8, 5) = 0
            detalle(8, 6) = 0
            detalle(8, 7) = 0
            detalle(8, 8) = total(6)
            

End Sub


Sub Consulta_boletas_exentas(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim NETO As Double
    Dim iva As Double
    

    Dim PASO As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,caja,numerozeta,boletainicial,boletafinal,(boletafinal-boletainicial+1) as diferencia,"
        csql.sql = csql.sql & "round(monto/1.19),monto-round(monto/1.19),exento,total,cigarro "
        csql.sql = csql.sql + "FROM boletasdeventa "
        csql.sql = csql.sql + "where mid(fecha,1,7) = '" + año + "-" + MES + "' and estbk='0' and cigarro='0' and exento>0 "
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
          
    
                total(1) = total(1) + resultados(5)
                total(2) = total(2) + resultados(6)
                total(3) = total(3) + resultados(7)
                total(4) = total(4) + resultados(8)
                total(6) = total(6) + resultados(9)
                VENTASEXENTAS = VENTASEXENTAS + resultados(8)
             
             
        
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            detalle(10, 1) = total(1)
            detalle(10, 2) = total(2)
            detalle(10, 3) = total(3)
            detalle(10, 4) = total(4)
            detalle(10, 5) = 0
            detalle(10, 6) = 0
            detalle(10, 7) = 0
            detalle(10, 8) = total(6)
            
            
             
            detalle(11, 1) = detalle(8, 1) + detalle(9, 1) + detalle(10, 1)
            detalle(11, 2) = detalle(8, 2) + detalle(9, 2) + detalle(10, 2)
            detalle(11, 3) = detalle(8, 3) + detalle(9, 3) + detalle(10, 3)
            detalle(11, 4) = detalle(8, 4) + detalle(9, 4) + detalle(10, 4)
            detalle(11, 5) = detalle(8, 5) + detalle(9, 5) + detalle(10, 5)
            detalle(11, 6) = detalle(8, 6) + detalle(9, 6) + detalle(10, 6)
            detalle(11, 7) = detalle(8, 7) + detalle(9, 7) + detalle(10, 7)
            detalle(11, 8) = detalle(8, 8) + detalle(9, 8) + detalle(10, 8)
            
            detalle(12, 1) = detalle(7, 1) + detalle(11, 1)
            detalle(12, 2) = detalle(7, 2) + detalle(11, 2)
            detalle(12, 3) = detalle(7, 3) + detalle(11, 3)
            detalle(12, 4) = detalle(7, 4) + detalle(11, 4)
            detalle(12, 5) = detalle(7, 5) + detalle(11, 5)
            detalle(12, 6) = detalle(7, 6) + detalle(11, 6)
            detalle(12, 7) = detalle(7, 7) + detalle(11, 7)
            detalle(12, 8) = detalle(7, 8) + detalle(11, 8)
            

End Sub
Sub Consulta_boletas_tbk(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim NETO As Double
    Dim iva As Double
    

    Dim PASO As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,caja,numerozeta,boletainicial,boletafinal,(boletafinal-boletainicial+1) as diferencia,round(monto/1.19),monto-round(monto/1.19),exento,total "
        csql.sql = csql.sql + "FROM boletasdeventa "
        csql.sql = csql.sql + "where mid(fecha,1,7) = '" + año + "-" + MES + "' and estbk='1' "
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
            'If resultados(8) = 0 Then
    
             total(1) = total(1) + resultados(5)
             total(2) = total(2) + resultados(6)
             total(3) = total(3) + resultados(7)
             total(4) = total(4) + resultados(8)
             total(6) = total(6) + resultados(9)
'             Else
'             NETO = Round(resultados(8) / 1.19, 0)
'             iva = resultados(8) - NETO
'             total(1) = total(1) - 0
'             total(2) = total(2) - NETO
'             total(3) = total(3) - iva
'             total(4) = total(4) - 0
'             total(5) = total(5) - resultados(8)
             VENTASEXENTAS = VENTASEXENTAS + resultados(8)
Rem              End If
             
        
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            detalle(9, 1) = total(1)
            detalle(9, 2) = total(2)
            detalle(9, 3) = total(3)
            detalle(9, 4) = total(4)
            detalle(9, 5) = 0
            detalle(9, 6) = 0
            detalle(9, 7) = 0
            detalle(9, 8) = total(6)
            
            
           
     

End Sub

Sub Consulta_boletas_exe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim NETO As Double
    Dim iva As Double
    

    Dim PASO As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,caja,numerozeta,boletainicial,boletafinal,(boletafinal-boletainicial+1) as diferencia,round(monto/1.19),monto-round(monto/1.19),exento,total "
        csql.sql = csql.sql + "FROM boletasdeventa "
        csql.sql = csql.sql + "where mid(fecha,1,7) = '" + año + "-" + MES + "' and estbk='2' "
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
            'If resultados(8) = 0 Then
    
             total(1) = total(1) + resultados(5)
             total(2) = total(2) + resultados(6)
             total(3) = total(3) + resultados(7)
             total(4) = total(4) + resultados(8)
             total(6) = total(6) + resultados(9)
'             Else
'             NETO = Round(resultados(8) / 1.19, 0)
'             iva = resultados(8) - NETO
'             total(1) = total(1) - 0
'             total(2) = total(2) - NETO
'             total(3) = total(3) - iva
'             total(4) = total(4) - 0
'             total(5) = total(5) - resultados(8)
             VENTASEXENTAS = VENTASEXENTAS + resultados(8)
Rem              End If
             
        
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
        
            detalle(10, 1) = total(1)
            detalle(10, 2) = total(2)
            detalle(10, 3) = total(3)
            detalle(10, 4) = total(4)
            detalle(10, 5) = 0
            detalle(10, 6) = 0
            detalle(10, 7) = 0
            detalle(10, 8) = total(6)
            
            
            detalle(11, 1) = detalle(8, 1) + detalle(9, 1) + detalle(10, 1)
            detalle(11, 2) = detalle(8, 2) + detalle(9, 2) + detalle(10, 2)
            detalle(11, 3) = detalle(8, 3) + detalle(9, 3) + detalle(10, 3)
            detalle(11, 4) = detalle(8, 4) + detalle(9, 4) + detalle(10, 4)
            detalle(11, 5) = detalle(8, 5) + detalle(9, 5) + detalle(10, 5)
            detalle(11, 6) = detalle(8, 6) + detalle(9, 6) + detalle(10, 6)
            detalle(11, 7) = detalle(8, 7) + detalle(9, 7) + detalle(10, 7)
            detalle(11, 8) = detalle(8, 8) + detalle(9, 8) + detalle(10, 8)
            
            detalle(12, 1) = detalle(7, 1) + detalle(11, 1)
            detalle(12, 2) = detalle(7, 2) + detalle(11, 2)
            detalle(12, 3) = detalle(7, 3) + detalle(11, 3)
            detalle(12, 4) = detalle(7, 4) + detalle(11, 4)
            detalle(12, 5) = detalle(7, 5) + detalle(11, 5)
            detalle(12, 6) = detalle(7, 6) + detalle(11, 6)
            detalle(12, 7) = detalle(7, 7) + detalle(11, 7)
            detalle(12, 8) = detalle(7, 8) + detalle(11, 8)
            
           
     

End Sub

Sub Consulta_honorarios(infogrilla As grillainformes)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim PASO As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fc.tipo,numero,fecha,fc.rut,cc.nombre,monto,retencion,liquido "
        csql.sql = csql.sql + "FROM boletasdehonorarios as fc,cuentascorrientes as cc "
        csql.sql = csql.sql + "where fc.rut=cc.rut and cc.año='" + COMBOAÑO.text + "' and cc.tipo='" + cuentahonorarios + "' and añocontable='" + año + "' and mescontable='" + MES + "' order by tipo,fecha "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        totalH(1) = 0
        totalH(2) = 0
        totalH(3) = 0
        totalH(4) = 0
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        Set resultados = csql.OpenResultset
        lin = 0
        While Not resultados.EOF
             totalH(1) = totalH(1) + resultados(5)
             totalH(2) = totalH(2) + resultados(6)
             totalH(3) = totalH(3) + resultados(7)
             resultados.MoveNext
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            detalle(22, 1) = 0
            detalle(22, 2) = totalH(2)
            detalle(22, 3) = 0
            
         
End Sub

