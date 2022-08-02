VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form infoge01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESTADOS DE RESULTADO"
   ClientHeight    =   8460
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   8325
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   4920
      TabIndex        =   34
      Top             =   7680
      Width           =   3255
      _ExtentX        =   5741
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
      Alignment       =   1
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp fechas 
      Height          =   1935
      Left            =   1320
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
      Height          =   6780
      Left            =   765
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11959
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
      Begin CoolButtons.cool_Button COMMAND2 
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   5535
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Genera Informe"
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   735
         Left            =   1035
         TabIndex        =   2
         Top             =   4725
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   1296
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
         Begin VB.OptionButton datos3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cuentas Globales"
            Height          =   330
            Left            =   2745
            TabIndex        =   32
            Top             =   270
            Width           =   1875
         End
         Begin VB.OptionButton datos2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cuentas Madre"
            Height          =   330
            Left            =   1170
            TabIndex        =   19
            Top             =   270
            Width           =   1470
         End
         Begin VB.OptionButton datos1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Todas"
            Height          =   375
            Left            =   90
            TabIndex        =   18
            Top             =   225
            Width           =   1515
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   6000
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
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
         Height          =   4440
         Left            =   1170
         TabIndex        =   6
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7832
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
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Todas Las Empresas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2040
            TabIndex        =   33
            Top             =   1200
            Width           =   2415
         End
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   855
            Left            =   120
            TabIndex        =   7
            Top             =   240
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
               Left            =   960
               TabIndex        =   16
               Top             =   240
               Width           =   3255
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   855
            Left            =   120
            TabIndex        =   8
            Top             =   1560
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
            Top             =   2520
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
         Begin XPFrame.FrameXp FrameXp9 
            Height          =   855
            Left            =   135
            TabIndex        =   30
            Top             =   3480
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
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
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
         Left            =   315
         TabIndex        =   26
         Top             =   4635
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
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
Attribute VB_Name = "infoge01"
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
Private TIPOS(7) As String
Private MES As String
Private año As String
  Private totales(13) As Double







Private Sub Check1_Click()
Combocrcc.text = "99.99 TODOS"
End Sub

Private Sub COMMAND2_Click()
If Check1.Value = "1" Then
Combocrcc.text = "99.99 TODAS "
End If

destino = "ESTADORESULTADO"

borraacumulado
If Check1.Value = False Then
Call generaacumulado(empresaactiva)
Else
Call listatodaslasempresas


End If

Dim TIMBRA As String

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)

año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)



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
If Mid(Combocrcc.text, 1, 5) = "99.99" Then
infogrilla.CABEZA.Caption = "ESTADO DE RESULTADO"
Call Consulta_Informe(infogrilla)
Else
Call Consulta_Informecrcc(infogrilla, Mid(Combocrcc.text, 1, 2) + Mid(Combocrcc.text, 4, 2))
End If



infogrilla.Visible = True
infogrilla.Caption = "ESTADO DE RESULTADOS ": grillainformes.Tag = "infoge01" & TIMBRA & FOLIO.text

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

Private Sub Form_Load()

CENTRAR Me

Dim i As Integer
Dim k As Integer

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO"
TIPOS(4) = "FACTURAS ELECTRONICAS"
TIPOS(5) = "NOTAS DE DEBITO ELECTRONICAS"
TIPOS(6) = "NOTAS DE CREDITO ELECTRONICAS"
TIPOS(7) = "FACTURAS ACTIVO FIJO"
    
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
For i = 1 To 10
For k = 1 To 10
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
CARGAcrcc
If MsgBox("SE RECOMIENDA ACTUALIZAR LOS DATOS ANTES DE PROCESAR DESEA ACTUALIZAR ", vbYesNo) = vbYes Then
proceso02.Show vbModal

End If




End Sub


    
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
  
    
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim tip As String
    Dim tota As Double
    Dim MES As Double
    
    Dim PASO As String
    tip = "1"
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigo,nombre,sum(enero),sum(febrero),sum(marzo),sum(abril),sum(mayo),sum(junio),sum(julio),sum(agosto),sum(septiembre),sum(octubre),sum(noviembre),sum(diciembre) from estadoresultado where año='" + COMBOAÑO.text + "' group by año,codigo "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        For k = 1 To 13
        totales(k) = 0
        Next k
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected
        MES = COMBOMES.ListIndex + 1
        
        Set resultados = csql.OpenResultset
        lin = 0
         While Not resultados.EOF
             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             For k = 1 To 2
             infogrilla.Grid1.Cell(lin, k).text = resultados(k - 1)
             Next k
             tota = 0
             
             For k = 3 To 14
             If k <= MES + 2 Then
             infogrilla.Grid1.Cell(lin, k).text = resultados(k - 1) * -1
             tota = tota + resultados(k - 1) * -1
             Else
             infogrilla.Grid1.Cell(lin, k).text = "0"
             
             
             End If
             
             Next k
             infogrilla.Grid1.Cell(lin, 15).text = tota
             
             If Mid(resultados(0), 5, 4) = "0000" Then
             infogrilla.Grid1.Range(lin, 0, lin, 15).FontSize = 7
             infogrilla.Grid1.Range(lin, 0, lin, 15).FontBold = True
             End If
             If Mid(resultados(0), 5, 4) <> "0000" Then
                    For k = 1 To 13
                      totales(k) = totales(k) + Val(infogrilla.Grid1.Cell(lin, k + 2).text)
                     Next k
             End If
             
             
             
             multi = 1
         

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

Sub Consulta_Informecrcc(infogrilla As grillainformes, CRCC As String)
Dim resultados As rdoResultset
  
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim tip As String
    Dim tota As Double
    
    Dim PASO As String
    tip = "1"
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT cm.codigo,cm.nombre"
        For k = 1 To MES
        csql.sql = csql.sql + ",sm.debe" + Format(k, "00") + "-sm.haber" + Format(k, "00")
        Next k
        For k = MES + 1 To 12
        csql.sql = csql.sql + ",'0'"
        Next k
        csql.sql = csql.sql + " FROM cuentasdelmayor as cm,saldoscentrosdecosto as sm "
        csql.sql = csql.sql + "where cm.codigo=sm.cuenta and sm.año=cm.año and sm.año='" + Format(fechasistema, "yyyy") + "' and mid(cm.codigo,1,1)>'2' and sm.codigo='" + CRCC + "' "
        If datos2.Value = True Then
        csql.sql = csql.sql + "and mid(sm.codigo,5,4)='0000' and mid(sm.codigo,3,2)<>'00' "
        End If
        If datos3.Value = True Then
        csql.sql = csql.sql + "and mid(sm.codigo,3,6)='000000' "
        End If
        
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        For k = 1 To 13
        totales(k) = 0
        Next k
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected
        
        Set resultados = csql.OpenResultset
        lin = 0
         While Not resultados.EOF
             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             For k = 1 To 2
             infogrilla.Grid1.Cell(lin, k).text = resultados(k - 1)
             Next k
             tota = 0
             For k = 3 To 14
             infogrilla.Grid1.Cell(lin, k).text = resultados(k - 1) * -1
             tota = tota + resultados(k - 1) * -1
             Next k
             infogrilla.Grid1.Cell(lin, 15).text = tota
             
             If Mid(resultados(0), 5, 4) = "0000" Then
             infogrilla.Grid1.Range(lin, 0, lin, 15).FontSize = 7
             infogrilla.Grid1.Range(lin, 0, lin, 15).FontBold = True
             End If
             If Mid(resultados(0), 5, 4) <> "0000" Then
                    For k = 3 To 15
                     totales(k - 2) = totales(k - 2) + CDbl(infogrilla.Grid1.Cell(lin, k).text)
                     Next k
             End If
             
             
             
             multi = 1
         

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
        infogrilla.Grid1.Range(lin, 0, lin, 15).FontSize = 7
        infogrilla.Grid1.Range(lin, 0, lin, 15).FontBold = True
        
        infogrilla.Grid1.Range(lin, 0, lin, 15).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 2).text = "TOTALES"
        For k = 1 To 13
        infogrilla.Grid1.Cell(lin, k + 2).text = totales(k)
        Next k
        
End Sub
               
   





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7
    
    
    FORMATOGRILLA(1, 1) = "CUENTA"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    
    FORMATOGRILLA(1, 3) = "ENERO"
    FORMATOGRILLA(1, 4) = "FEBRERO"
    FORMATOGRILLA(1, 5) = "MARZO"
    FORMATOGRILLA(1, 6) = "ABRIL"
    FORMATOGRILLA(1, 7) = "MAYO"
    FORMATOGRILLA(1, 8) = "JUNIO"
    FORMATOGRILLA(1, 9) = "JULIO"
    FORMATOGRILLA(1, 10) = "AGOSTO"
    FORMATOGRILLA(1, 11) = "SEPTIEMBRE"
    FORMATOGRILLA(1, 12) = "OCTUBRE"
    FORMATOGRILLA(1, 13) = "NOVIEMBRE"
    FORMATOGRILLA(1, 14) = "DICIEMBRE"
    FORMATOGRILLA(1, 15) = "TOTAL"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "30"
    For k = 3 To 15
    FORMATOGRILLA(2, k) = "14"
    Next k
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    For k = 3 To 15
    FORMATOGRILLA(3, k) = "N"
    Next k
    For k = 3 To 15
    FORMATOGRILLA(4, k) = "##,###,###,##0"
    Next k
    
    
    Rem LOCCKED
    For k = 1 To 15
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 16
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
    infogrilla.Grid1.FrozenCols = 2
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

Private Sub Option1_Click()

End Sub

Sub CARGAcrcc()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM centrosdecosto where año='" + COMBOAÑO.text + "' "
        csql.sql = csql.sql + "order by codigo"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             Combocrcc.AddItem (Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + " " + resultados(1))
             
            resultados.MoveNext
            Wend
        End If
        Combocrcc.AddItem ("99.99" + " " + "TODOS")
            
        Combocrcc.text = Combocrcc.List(LINEA)
        
   

End Sub
Sub generaacumulado(empresa)
Dim resultados As rdoResultset
Dim csql As New rdoQuery
Dim MES As String

       MES = Format(Format(COMBOMES.ListIndex) + 1, "00")
       
        
        Set csql.ActiveConnection = contadb
        csql.sql = "insert into " + clientesistema + "conta.estadoresultado (empresa,año,codigo,nombre,enero,febrero,marzo,abril,mayo,junio,julio,agosto,septiembre,octubre,noviembre,diciembre) "
        csql.sql = csql.sql + "SELECT '" + empresa + "','" + COMBOAÑO.text + "',cm.codigo,cm.nombre "
        For k = 1 To 12
        csql.sql = csql.sql + ",sm.debe" + Format(k, "00") + "-sm.haber" + Format(k, "00")
        Next k
        csql.sql = csql.sql + " FROM " + clientesistema + "conta" + empresa + ".cuentasdelmayor as cm," + clientesistema + "conta" + empresa + ".saldosdelmayor as sm "
        csql.sql = csql.sql + "where cm.codigo=sm.codigo and sm.año='" + COMBOAÑO.text + "' and mid(sm.codigo,1,1)>'2' and cm.año='" + COMBOAÑO.text + "' "
        If datos2.Value = True Then
        csql.sql = csql.sql + "and mid(sm.codigo,5,4)='0000' and mid(sm.codigo,3,2)<>'00' "
        End If
        If datos3.Value = True Then
        csql.sql = csql.sql + "and mid(sm.codigo,3,6)='000000' "
        End If
        
        csql.Execute
        
End Sub
Sub borraacumulado()
Dim resultados As rdoResultset
Dim csql As New rdoQuery
Dim MES As String

        
        
        Set csql.ActiveConnection = contadb
        csql.sql = "delete from " + clientesistema + "conta.estadoresultado where año='" + COMBOAÑO.text + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, contadb, "")
        
        
End Sub

Sub listatodaslasempresas()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa "
        csql.sql = csql.sql + "FROM maestroempresas where estadoresultado='1' "
        csql.sql = csql.sql + "order by codigoempresa "
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
            Call generaacumulado(resultados(0))
            resultados.MoveNext
            Wend
        End If
        
   

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
