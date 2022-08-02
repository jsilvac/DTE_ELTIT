VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form infocarne 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Ventas"
   ClientHeight    =   8460
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   8325
   Begin XPFrame.FrameXp fechas 
      Height          =   1935
      Left            =   1755
      TabIndex        =   20
      Top             =   6480
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
      Height          =   6285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11086
      BackColor       =   16761024
      Caption         =   "Lista Libro de Ventas"
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
         Left            =   360
         TabIndex        =   12
         Top             =   3960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Genera Informe"
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1095
         Left            =   240
         TabIndex        =   2
         Top             =   360
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
            Caption         =   "Rango Fecha"
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton datos1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Mensual"
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
         Top             =   5850
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
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton RESUMEN1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detallado"
            Height          =   375
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4215
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7435
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
               Left            =   240
               TabIndex        =   17
               Text            =   "01"
               Top             =   360
               Width           =   375
            End
            Begin VB.Label empresanombre 
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
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp9 
            Height          =   855
            Left            =   120
            TabIndex        =   30
            Top             =   3240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "Centros de Costo"
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
            Begin VB.ComboBox Combocrcc 
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
               Left            =   240
               TabIndex        =   31
               Top             =   360
               Width           =   3855
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
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
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton DETALLE2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "No"
            Height          =   375
            Left            =   480
            TabIndex        =   14
            Top             =   600
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   855
         Left            =   1170
         TabIndex        =   26
         Top             =   4725
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1508
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
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton timbrado 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Timbrado"
            Height          =   255
            Left            =   2160
            TabIndex        =   28
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox FOLIO 
            Height          =   285
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   27
            Top             =   360
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "infocarne"
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
Private TIPOS(6) As String
Private MES As String
Private año As String
Private centro As String








Private Sub COMMAND2_Click()
Dim TIMBRA As String

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(servidor, clientesistema + "conta" + DATO1.text, USUARIO, password)
centro = Mid(Combocrcc.text, 1, 2) + Mid(Combocrcc.text, 4, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

CARGAmayor
leermayor
Call CARGAGRILLA(infogrilla)
For K = 1 To 10
detalle(K, 1) = 0
detalle(K, 2) = 0
detalle(K, 3) = 0
detalle(K, 4) = 0
detalle(K, 5) = 0
detalle(K, 6) = 0
detalle(K, 7) = 0
detalle(K, 8) = 0
detalle(K, 9) = 0
detalle(K, 10) = 0
Next K
Call Consulta_Informe(infogrilla)
infogrilla.Visible = True
infogrilla.Caption = "LIBRO DE VENTAS": grillainformes.Tag = "auxiliar44" & TIMBRA & FOLIO.text
infogrilla.CABEZA.Caption = "LISTADO DE VENTAS " + Combocrcc.text
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
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then DATO1.SetFocus: GoTo no:
    COMBOMES.SetFocus
    empresanombre.Caption = SQLUTIL.datos(1, 3)
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
Dim K As Integer

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO FACTURAS"
TIPOS(4) = "NOTAS DE CREDITO BOLETAS"
    
    Call Conectar_BD
    Call Conectarconta(servidor, clientesistema + "conta", USUARIO, password)
For i = 1 To 10
For K = 1 To 10
detalle(K, i) = 0
Next K

Next i
OPCIONES.Visible = True

original.Value = True

For K = 1 To 12
COMBOMES.AddItem MonthName(K)
Next K
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For K = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem K
Next K
COMBOAÑO.ListIndex = K - 2001
DATO1.text = empresaactiva
empresanombre.Caption = nombreempresa
datos1.Value = True
RESUMEN1.Value = True
DETALLE1.Value = True
    desdefecha.Caption = fechasistema
    hastafecha.Caption = fechasistema

fechas.Visible = False
CARGAcrcc

End Sub


    
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    

    Dim PASO As String
        Set cSql.ActiveConnection = temporal
        cSql.SQL = "SELECT fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,iva,exento,total "
        cSql.SQL = cSql.SQL + "FROM facturasdeventas as fc,cuentascorrientes as cc "
        If centro = "9999" Then cSql.SQL = cSql.SQL + "where fc.rut=cc.rut and cc.tipo='" + cuentacliente + "' and fecha >= '" + año + "/" + MES + "/" + "01" + "' and fecha <= '" + año + "/" + MES + "/" + "31' order by tipo,numero "
        If centro <> "9999" Then cSql.SQL = cSql.SQL + "where fc.crcc='" + centro + "' and fc.rut=cc.rut and cc.tipo='" + cuentacliente + "' and fecha >= '" + año + "/" + MES + "/" + "01" + "' and fecha <= '" + año + "/" + MES + "/" + "31' order by tipo,numero "
        
        cSql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        If cSql.RowsAffected > 0 Then
        barra.Max = cSql.RowsAffected + 1
        
        Set resultados = cSql.OpenResultset
        lin = 0
         While Not resultados.EOF
    If datos2.Value = True And resultados(3) < desdefecha.Caption Then GoTo PASO:
    If datos2.Value = True And resultados(3) > hastafecha.Caption Then GoTo PASO:
    
         If RESUMEN1.Value = True Then
             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             For K = 0 To 8
             infogrilla.Grid1.Cell(lin, K + 1).text = resultados(K)
             Next K
             multi = 1
                If resultados(0) = "1" Then infogrilla.Grid1.Cell(lin, 1).text = "FA"
                If resultados(0) = "2" Then infogrilla.Grid1.Cell(lin, 1).text = "ND"
                If resultados(0) = "3" Then infogrilla.Grid1.Cell(lin, 1).text = "NC": multi = -1
                If resultados(0) = "4" Then infogrilla.Grid1.Cell(lin, 1).text = "NB": multi = -1
             
                infogrilla.Grid1.Cell(lin, 6).text = resultados(5) * multi
                infogrilla.Grid1.Cell(lin, 7).text = resultados(6) * multi
                infogrilla.Grid1.Cell(lin, 8).text = resultados(7) * multi
                infogrilla.Grid1.Cell(lin, 9).text = resultados(8) * multi
                infogrilla.Grid1.Cell(lin, 4).text = Mid(resultados(3), 1, 9) + "-" + Mid(resultados(3), 10, 1)

         
         End If
             If resultados(0) = "3" Then multi = -1 Else multi = 1
             total(1) = total(1) + resultados(5) * multi
             total(2) = total(2) + resultados(6) * multi
             total(3) = total(3) + resultados(7) * multi
             total(4) = total(4) + resultados(8) * multi
             If resultados(0) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(5): detalle(1, 3) = detalle(1, 3) + resultados(6):: detalle(1, 4) = detalle(1, 4) + resultados(7):: detalle(1, 5) = detalle(1, 5) + resultados(8)
             If resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(5): detalle(2, 3) = detalle(2, 3) + resultados(6):: detalle(2, 4) = detalle(2, 4) + resultados(7):: detalle(2, 5) = detalle(2, 5) + resultados(8)
             If resultados(2) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(5): detalle(3, 3) = detalle(3, 3) + resultados(6):: detalle(3, 4) = detalle(3, 4) + resultados(7):: detalle(3, 5) = detalle(3, 5) + resultados(8)
             If resultados(3) = "4" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(5): detalle(4, 3) = detalle(4, 3) + resultados(6):: detalle(4, 4) = detalle(4, 4) + resultados(7):: detalle(4, 5) = detalle(4, 5) + resultados(8)
             
              Call Consultadetalle(resultados(0), resultados(1), resultados(2), infogrilla)
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
        infogrilla.Grid1.Range(lin, 6, lin, 9).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "TOTALES"
        infogrilla.Grid1.Cell(lin, 6).text = total(1)
        infogrilla.Grid1.Cell(lin, 7).text = total(2)
        infogrilla.Grid1.Cell(lin, 8).text = total(3)
        infogrilla.Grid1.Cell(lin, 9).text = total(4)
      
    
    TOTALge = 0
    lin = lin + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 9
    infogrilla.Grid1.Range(lin, 4, lin + 4, 9).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 4, 9).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 4, 9).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 4, 9).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 4, 9).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 4, 9).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 4).text = "Cant."
    infogrilla.Grid1.Cell(lin, 5).text = "Documentos"
    infogrilla.Grid1.Cell(lin, 6).text = "Neto"
    infogrilla.Grid1.Cell(lin, 7).text = "i.v.a"
    infogrilla.Grid1.Cell(lin, 8).text = "exento"
    infogrilla.Grid1.Cell(lin, 9).text = "total"
    
    For K = 1 To 4
    lin = lin + 1
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
    infogrilla.Grid1.Cell(lin, 5).text = TIPOS(K)
    infogrilla.Grid1.Cell(lin, 4).text = Format(detalle(K, 1), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 6).text = Format(detalle(K, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(K, 3), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(K, 4), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(K, 5), "###,###,##0")
    
    Next K
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    lin = lin + 2
    For K = 1 To canplan
    If plan(K, 3) <> 0 Then
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 4).text = plan(K, 1)
        infogrilla.Grid1.Cell(lin, 5).text = plan(K, 2)
        infogrilla.Grid1.Cell(lin, 6).text = plan(K, 3)
        TOTALge = TOTALge + plan(K, 3)
        End If
    Next K
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 5).text = "TOTAL DETALLE"
        infogrilla.Grid1.Range(lin, 6, lin, 6).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 6).text = TOTALge
               
    End Sub
    





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "RUT"
    FORMATOGRILLA(1, 5) = "PROVEEDOR"
    FORMATOGRILLA(1, 6) = "NETO"
    FORMATOGRILLA(1, 7) = "IVA"
    FORMATOGRILLA(1, 8) = "EXENTO"
    FORMATOGRILLA(1, 9) = "TOTAL"
    FORMATOGRILLA(1, 10) = "GLOSA"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "30"
    FORMATOGRILLA(2, 6) = "9"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "0"
    
    
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
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,###"
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    
    Rem LOCCKED
    For K = 1 To 10
    FORMATOGRILLA(5, K) = "TRUE"
    Next K
    
    infogrilla.Grid1.Cols = 11
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
    
    For K = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, K).text = FORMATOGRILLA(1, K)
        infogrilla.Grid1.Column(K).Width = Val(FORMATOGRILLA(2, K)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(K).MaxLength = Val(FORMATOGRILLA(2, K))
        infogrilla.Grid1.Column(K).FormatString = FORMATOGRILLA(4, K)
        infogrilla.Grid1.Column(K).Locked = FORMATOGRILLA(5, K)
        If FORMATOGRILLA(3, K) = "N" Then infogrilla.Grid1.Column(K).Alignment = cellRightCenter
        If FORMATOGRILLA(3, K) = "D" Then infogrilla.Grid1.Column(K).CellType = cellCalendar
        
    Next K
End Sub

Sub leermayor()
    tipoprove = cuentaproveedor
    

    
End Sub

'Sub Consultadetalle(MES As String, año As String)
'Dim multi As Integer
'
'Dim resultados2 As rdoResultset
'    Dim cSql2 As New rdoQuery
'        Set cSql2.ActiveConnection = db
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
    Dim cSql As New rdoQuery
    Dim LINEAS As Integer
    
   
        Set cSql.ActiveConnection = temporal
        cSql.SQL = "SELECT codigo,nombre,tipo "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        linea = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
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

Sub Consultadetalle(tipo, numero, fecha As Date, infogrilla As grillainformes)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
       
        
        Set cSql2.ActiveConnection = temporal
        cSql2.SQL = "SELECT cuentadelmayor,monto "
        cSql2.SQL = cSql2.SQL + "FROM facturasdeventas_detalle "
        cSql2.SQL = cSql2.SQL + "where tipo='" + tipo + "' and numero='" + numero + "'  order by linea "
        cSql2.Execute

        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset

        While Not resultados2.EOF
          For K = 1 To canplan
          If tipo = 3 Then multi = -1 Else multi = 1
          If resultados2(0) = plan(K, 1) Then plan(K, 3) = plan(K, 3) + (resultados2(1) * multi)
          If resultados2(0) = plan(K, 1) And DETALLE1.Value = True Then infogrilla.Grid1.Cell(lin, 10).text = plan(K, 2): K = canplan + 1

          Next K
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
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestroempresas", DATO1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub

Sub CARGAcrcc()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim LINEAS As Integer
    
   
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre "
        cSql.SQL = cSql.SQL + "FROM centrosdecosto "
        cSql.SQL = cSql.SQL + "order by codigo"
        cSql.Execute
        linea = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
             While Not resultados.EOF
             linea = linea + 1
             Combocrcc.AddItem (Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + " " + resultados(1))
             
            resultados.MoveNext
            Wend
        End If
        Combocrcc.AddItem ("99.99" + " " + "TODOS")
            
        Combocrcc.text = Combocrcc.List(linea)
        
   

End Sub

