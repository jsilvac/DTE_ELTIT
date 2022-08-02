VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form informa01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista  Archivos Maestros"
   ClientHeight    =   5955
   ClientLeft      =   -150
   ClientTop       =   510
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   2040
      TabIndex        =   32
      Top             =   5280
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
         TabIndex        =   34
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp opciones 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9128
      BackColor       =   16744576
      Caption         =   "Lista Archivos Maestros"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
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
      Begin XPFrame.FrameXp opcion1 
         Height          =   2775
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
         BackColor       =   16761024
         Caption         =   "Cuentas del Mayor"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   1455
            Left            =   3480
            TabIndex        =   15
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   2566
            BackColor       =   16761024
            Caption         =   "SALDOS"
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
            Alignment       =   1
            Begin VB.OptionButton saldocm1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Sin Saldos"
               Height          =   495
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton saldocm2 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Con Saldos"
               Height          =   495
               Left            =   120
               TabIndex        =   16
               Top             =   840
               Width           =   1935
            End
         End
         Begin XPFrame.FrameXp FrameXp3 
            Height          =   1455
            Left            =   480
            TabIndex        =   13
            Top             =   360
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
            BackColor       =   16761024
            Caption         =   "OPCIONES"
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
            Alignment       =   1
            Begin VB.OptionButton cmtoda 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Plan de Cuentas"
               Height          =   495
               Left            =   360
               TabIndex        =   14
               Top             =   480
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin CoolButtons.cool_Button COMMAND5 
            Height          =   495
            Left            =   2160
            TabIndex        =   30
            Top             =   2160
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "GENERA INFORME"
         End
      End
      Begin XPFrame.FrameXp opcion3 
         Height          =   2775
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
         BackColor       =   16761024
         Caption         =   "Centros de Costo"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
         Begin XPFrame.FrameXp FrameXp2 
            Height          =   1335
            Left            =   3480
            TabIndex        =   10
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2355
            BackColor       =   16761024
            Caption         =   "SALDOS"
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
            Alignment       =   1
            Begin VB.OptionButton SALDOCC2 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Con Saldos"
               Height          =   495
               Left            =   360
               TabIndex        =   12
               Top             =   720
               Width           =   1935
            End
            Begin VB.OptionButton SALDOCC1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Sin Saldos"
               Height          =   495
               Left            =   360
               TabIndex        =   11
               Top             =   240
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin XPFrame.FrameXp FrameXp1 
            Height          =   1335
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2355
            BackColor       =   16761024
            Caption         =   "OPCIONES"
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
            Alignment       =   1
            Begin VB.OptionButton cctoda 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Todas los Centros"
               Height          =   495
               Left            =   240
               TabIndex        =   9
               Top             =   360
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin CoolButtons.cool_Button COMMAND6 
            Height          =   495
            Left            =   2400
            TabIndex        =   29
            Top             =   2280
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "GENERA INFORME"
         End
      End
      Begin XPFrame.FrameXp opcion2 
         Height          =   3735
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6588
         BackColor       =   16761024
         Caption         =   "Cuentas Corrientes"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "EXTERNOS"
            Height          =   255
            Left            =   3120
            TabIndex        =   37
            Top             =   3240
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "RELACIONADOS"
            Height          =   255
            Left            =   1320
            TabIndex        =   36
            Top             =   3240
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "TODOS"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   3240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin CoolButtons.cool_Button command2 
            Height          =   495
            Left            =   4560
            TabIndex        =   28
            Top             =   3120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "GENERA INFORME"
         End
         Begin XPFrame.FrameXp cuentacorriente 
            Height          =   975
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   1720
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
            Begin VB.TextBox dato1 
               BackColor       =   &H00E1FFFD&
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
               Left            =   1200
               MaxLength       =   8
               TabIndex        =   25
               Tag             =   "tipo"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Codigo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2400
               TabIndex        =   26
               Top             =   360
               Width           =   3135
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   1575
            Left            =   3480
            TabIndex        =   21
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2778
            BackColor       =   16761024
            Caption         =   "SALDOS"
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
            Alignment       =   1
            Begin VB.OptionButton SALDOCT2 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Con Saldos"
               Height          =   495
               Left            =   240
               TabIndex        =   23
               Top             =   840
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton SALDOCT1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Sin Saldos"
               Height          =   495
               Left            =   240
               TabIndex        =   22
               Top             =   240
               Width           =   1935
            End
         End
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   1575
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2778
            BackColor       =   16761024
            Caption         =   "OPCIONES"
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
            Alignment       =   1
            Begin VB.OptionButton cttodatipo 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Todas las Cuentas de Un tipo"
               Height          =   495
               Left            =   120
               TabIndex        =   20
               Top             =   960
               Value           =   -1  'True
               Width           =   2415
            End
            Begin VB.OptionButton cttoda 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Todos los Tipos"
               Height          =   495
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   1935
            End
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   195
         Left            =   270
         TabIndex        =   31
         Top             =   4680
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00F29980&
         Caption         =   "Centros de Costo"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00F29980&
         Caption         =   "Cuentas Corrientes"
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
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00F29980&
         Caption         =   "Cuentas Del  Mayor"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   5280
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "informa01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private saldo As Double
Private dedonde As Integer
Private tipoctacte As String
Private tipocuenta(4) As String
Private TIPOAUXILIAR(3) As String
Private SALDOANTERIOR As Double
Private totaldebe As Double
Private totalhaber As Double






Private Sub busca_Click()

End Sub




Private Sub COMMAND2_Click()
If cttodatipo.Value = True And Val(dato1.text) = 0 Then dato1.SetFocus: GoTo no:

dedonde = 2


If cttoda.Value = True Then Call ACEPTA(2)

If cttodatipo.Value = True Then Call ACEPTA(2)
no:
End Sub

Private Sub Command5_Click()
dedonde = 1

If cmtoda.Value = True Then Call ACEPTA(1)

End Sub
Sub ACEPTA(opcion)
Dim infogrilla As grillainformes
Set infogrilla = New grillainformes
If opcion = 1 Then infogrilla.Caption = "LISTA PLAN DE CUENTAS"
If opcion = 2 Then infogrilla.Caption = "LISTA CUENTAS CORRIENTES"
If opcion = 3 Then infogrilla.Caption = "LISTA CENTROS DE COSTO"

If opcion = 1 Then Call leecuentas(infogrilla)
If opcion = 2 Then Call leercuentascorrientes(infogrilla)
If opcion = 3 Then Call leecrcc(infogrilla)

If opcion = 1 Then grillainformes.Tag = "informa01_1"
If opcion = 2 Then grillainformes.Tag = "informa01_2"
If opcion = 3 Then grillainformes.Tag = "informa01_3"


infogrilla.Grid1.Visible = True

infogrilla.Show

End Sub

Private Sub Command6_Click()
dedonde = 3

If cctoda.Value = True Then Call ACEPTA(3)


End Sub

Private Sub Command7_Click()
opcion1.Visible = True
opcion2.Visible = False
opcion3.Visible = False


End Sub

Private Sub command8_Click()
opcion1.Visible = False
opcion2.Visible = True
opcion3.Visible = False

End Sub
Private Sub Command9_Click()
opcion1.Visible = False
opcion2.Visible = False
opcion3.Visible = True

End Sub

Private Sub cool_Button1_Click()

End Sub

Private Sub cttoda_Click()
If cttodatipo.Value = False Then cuentacorriente.Visible = False

End Sub

Private Sub cttodatipo_Click()
If cttodatipo.Value = True Then cuentacorriente.Visible = True: dato1.SetFocus

If cttodatipo.Value = False Then cuentacorriente.Visible = False

End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudatipocuenta(dato1)
End Sub
Sub ayudatipocuenta(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "CTACTE <> '0'"
    cabezas = Array("cuenta", "nombre")
    mensajeAyuda = "Ayuda tipo de Cuentas Corrientes"
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub LEETIPOCTACTE()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)

   If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
   Label1.Caption = sqlconta.response(1, 3)

no:
End Sub



Private Sub dato1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command2.SetFocus

End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    
lin = 0



tipocuenta(1) = "ACTIVO"
tipocuenta(2) = "PASIVO"
tipocuenta(3) = "RESULTADO"
TIPOAUXILIAR(0) = ""
TIPOAUXILIAR(1) = "CENTRO COSTO"
TIPOAUXILIAR(2) = "BANCO"

cuentacorriente.Visible = True
opcion1.Visible = False
opcion2.Visible = True
opcion3.Visible = False


End Sub



    





Sub formatoplan(infogrilla As grillainformes)

Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7
    
    
    FORMATOGRILLA(1, 1) = "CUENTA"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "TIPO    "
    FORMATOGRILLA(1, 4) = "CTA.CTE "
    FORMATOGRILLA(1, 5) = "CRCC"
    FORMATOGRILLA(1, 6) = "BANCO   "
    FORMATOGRILLA(1, 7) = "I.LICOR "
    FORMATOGRILLA(1, 8) = "I.CARNE "
    FORMATOGRILLA(1, 9) = "I.HARINA"
    FORMATOGRILLA(1, 10) = "ACT.FIJO"
    
    If saldocm1.Value = False Then FORMATOGRILLA(1, 11) = "SALDO DEBE"
    If saldocm1.Value = False Then FORMATOGRILLA(1, 12) = "SALDO HABER"
    If saldocm1.Value = True Then FORMATOGRILLA(1, 11) = ""
    If saldocm1.Value = True Then FORMATOGRILLA(1, 12) = ""
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "30"
    FORMATOGRILLA(2, 3) = "20"
    FORMATOGRILLA(2, 4) = "7"
    FORMATOGRILLA(2, 5) = "7"
    FORMATOGRILLA(2, 6) = "7"
    FORMATOGRILLA(2, 7) = "7"
    FORMATOGRILLA(2, 8) = "7"
    FORMATOGRILLA(2, 9) = "7"
    FORMATOGRILLA(2, 10) = "7"

    If saldocm1.Value = False Then FORMATOGRILLA(2, 11) = "15"
    If saldocm1.Value = False Then FORMATOGRILLA(2, 12) = "15"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    
    If saldocm1.Value = False Then FORMATOGRILLA(3, 11) = "N"
    If saldocm1.Value = False Then FORMATOGRILLA(3, 12) = "N"
    
    
    Rem FORMATO GRILLA
    
    If saldocm1.Value = False Then FORMATOGRILLA(4, 11) = "###,###,###,##0"
    If saldocm1.Value = False Then FORMATOGRILLA(4, 12) = "###,###,###,##0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    
    If saldocm1.Value = False Then FORMATOGRILLA(5, 11) = "TRUE"
    If saldocm1.Value = False Then FORMATOGRILLA(5, 12) = "TRUE"
    
    If saldocm1.Value = True Then infogrilla.Grid1.Cols = 11
    If saldocm1.Value = False Then infogrilla.Grid1.Cols = 13
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
     infogrilla.AutoRedraw = False
     
    infogrilla.Grid1.DisplayFocusRect = False
    infogrilla.Grid1.ExtendLastCol = True
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
For k = 4 To 10
infogrilla.Grid1.Column(k).CellType = cellCheckBox
Next k
infogrilla.AutoRedraw = True
infogrilla.Refresh
infogrilla.Grid1.ExtendLastCol = False



End Sub
Sub formatocrcc(infogrilla As grillainformes)

Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
    
    
    FORMATOGRILLA(1, 1) = "CUENTA"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    If SALDOCC1.Value = False Then FORMATOGRILLA(1, 3) = "SALDO DEBE"
    If SALDOCC1.Value = False Then FORMATOGRILLA(1, 4) = "SALDO HABER"
    If SALDOCC1.Value = True Then FORMATOGRILLA(1, 3) = ""
    If SALDOCC1.Value = True Then FORMATOGRILLA(1, 4) = ""
    FORMATOGRILLA(1, 5) = ""
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "50"
   
    If SALDOCC1.Value = False Then FORMATOGRILLA(2, 3) = "15"
    If SALDOCC1.Value = False Then FORMATOGRILLA(2, 4) = "15"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    If SALDOCC1.Value = False Then FORMATOGRILLA(3, 3) = "N"
    If SALDOCC1.Value = False Then FORMATOGRILLA(3, 4) = "N"
    
    
    Rem FORMATO GRILLA
    
    If SALDOCC1.Value = False Then FORMATOGRILLA(4, 3) = "###,###,###,##0"
    If SALDOCC1.Value = False Then FORMATOGRILLA(4, 4) = "###,###,###,##0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    If SALDOCC1.Value = False Then FORMATOGRILLA(5, 3) = "TRUE"
    If SALDOCC1.Value = False Then FORMATOGRILLA(5, 4) = "TRUE"
    
    If SALDOCC1.Value = True Then infogrilla.Grid1.Cols = 3
    If SALDOCC1.Value = False Then infogrilla.Grid1.Cols = 6
    FORMATOGRILLA(5, 5) = "TRUE"
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    infogrilla.Grid1.ExtendLastCol = True
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
Sub formatoctacte(infogrilla As grillainformes)

Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
    
    
    FORMATOGRILLA(1, 1) = "RUT "
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "ANTERIOR"
    FORMATOGRILLA(1, 4) = "DEBE"
    FORMATOGRILLA(1, 5) = "HABER"
    FORMATOGRILLA(1, 6) = "SALDO"
    
    
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "12"
    FORMATOGRILLA(2, 2) = "30"
   
    FORMATOGRILLA(2, 3) = "11"
    FORMATOGRILLA(2, 4) = "11"
    FORMATOGRILLA(2, 5) = "11"
    FORMATOGRILLA(2, 6) = "11"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    
    
    Rem FORMATO GRILLA
    
    FORMATOGRILLA(4, 3) = "###,###,###,##0"
    FORMATOGRILLA(4, 4) = "###,###,###,##0"
    FORMATOGRILLA(4, 5) = "###,###,###,##0"
    FORMATOGRILLA(4, 6) = "###,###,###,##0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    If SALDOCT1.Value = True Then infogrilla.Grid1.Cols = 3
    If SALDOCT1.Value = False Then infogrilla.Grid1.Cols = 7
    FORMATOGRILLA(5, 6) = "TRUE"
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    infogrilla.Grid1.ExtendLastCol = False
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


Sub leecuentas(ByRef infogrilla As grillainformes)
Call formatoplan(infogrilla)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre,tipo,ctacte,crcc,banco,ila,ica,iha,activo "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
       
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        
        LINEAS = 0
        infogrilla.AutoRedraw = False
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        
        LINEAS = LINEAS + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        If Mid(resultados2(0), 5, 4) = "0000" Then
        With infogrilla.Grid1.Range(LINEAS, 0, LINEAS, infogrilla.Grid1.Cols - 1)
        .FontBold = True
        .FontUnderline = True
        End With
        End If
      
        infogrilla.Grid1.Cell(LINEAS, 1).text = Mid(resultados2(0), 1, 2) + "." + Mid(resultados2(0), 3, 2) + "." + Mid(resultados2(0), 5, 4)
        infogrilla.Grid1.Cell(LINEAS, 2).text = resultados2(1)
        infogrilla.Grid1.Cell(LINEAS, 3).text = tipocuenta(resultados2(2))
        infogrilla.Grid1.Cell(LINEAS, 4).text = resultados2(3)
        infogrilla.Grid1.Cell(LINEAS, 5).text = resultados2(4)
        infogrilla.Grid1.Cell(LINEAS, 6).text = resultados2(5)
        infogrilla.Grid1.Cell(LINEAS, 7).text = resultados2(6)
        infogrilla.Grid1.Cell(LINEAS, 8).text = resultados2(7)
        infogrilla.Grid1.Cell(LINEAS, 9).text = resultados2(8)
        infogrilla.Grid1.Cell(LINEAS, 10).text = resultados2(9)
       
        If saldocm1.Value = False And Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERSALDOS(resultados2(0))
        If saldocm1.Value = False And saldo >= 0 Then infogrilla.Grid1.Cell(LINEAS, 11).text = saldo
        If saldocm1.Value = False And saldo < 0 Then infogrilla.Grid1.Cell(LINEAS, 12).text = saldo * -1
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
      infogrilla.AutoRedraw = True
      infogrilla.Refresh
        

End Sub
Sub LEERSALDOS(cuenta)
Dim SUMD As Double
Dim SUMH As Double
Dim anted As Double
Dim anteh As Double
Dim DIFE As Double

    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = "debe01"
    campos(5, 0) = "debe02"
    campos(6, 0) = "debe03"
    campos(7, 0) = "debe04"
    campos(8, 0) = "debe05"
    campos(9, 0) = "debe06"
    campos(10, 0) = "debe07"
    campos(11, 0) = "debe08"
    campos(12, 0) = "debe09"
    campos(13, 0) = "debe10"
    campos(14, 0) = "debe11"
    campos(15, 0) = "debe12"
    campos(16, 0) = "haber01"
    campos(17, 0) = "haber02"
    campos(18, 0) = "haber03"
    campos(19, 0) = "haber04"
    campos(20, 0) = "haber05"
    campos(21, 0) = "haber06"
    campos(22, 0) = "haber07"
    campos(23, 0) = "haber08"
    campos(24, 0) = "haber09"
    campos(25, 0) = "HABER10"
    campos(26, 0) = "HABER11"
    campos(27, 0) = "HABER12"
    campos(28, 0) = ""
    
    condicion = "codigo=" + "'" + cuenta + "' and año ='" + Format(fechasistema, "yyyy") + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
    
    anted = sqlconta.response(2, 3)
    anteh = sqlconta.response(3, 3)
    SUMD = 0: SUMH = 0
For k = 1 To CDbl(Format(fechasistema, "mm"))
SUMD = SUMD + sqlconta.response(k + 3, 3)
SUMH = SUMH + sqlconta.response(k + 15, 3)
Next
saldo = anted - anteh + SUMD - SUMH
End Sub

Sub DATOSSALDOS(cuenta)
Call LEERSALDOS(cuenta)
End Sub
Sub leecrcc(infogrilla As grillainformes)
Call formatocrcc(infogrilla)
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    infogrilla.Grid1.AutoRedraw = False
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM centrosdecosto "
       
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
       
        If csql2.RowsAffected > 0 Then
     
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        infogrilla.Grid1.Rows = LINEAS + 1
        If Mid(resultados2(0), 3, 2) = "00" Then
        With infogrilla.Grid1.Range(LINEAS, 0, LINEAS, infogrilla.Grid1.Cols - 1)
        .FontBold = True
        .FontUnderline = True
        End With
        End If
        
        
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(LINEAS, 1).text = Mid(resultados2(0), 1, 2) + "." + Mid(resultados2(0), 3, 2)
        infogrilla.Grid1.Cell(LINEAS, 2).text = resultados2(1)
     
        If SALDOCC1.Value = False Then LEERSALDOS (resultados2(0))
        If SALDOCC1.Value = False And saldo >= 0 Then infogrilla.Grid1.Cell(LINEAS, 3).text = saldo
        If SALDOCC1.Value = False And saldo < 0 Then infogrilla.Grid1.Cell(LINEAS, 4).text = saldo * -1
       
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    infogrilla.Grid1.AutoRedraw = False
    infogrilla.Grid1.Refresh
    
    

End Sub

Sub leercuentascorrientes(infogrilla As grillainformes)
 Call formatoctacte(infogrilla)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim tipodecuenta As String
    Dim tiposaldo As String
    Dim suma1 As Double
    Dim suma2 As Double
    Dim suma3 As Double
    Dim suma4 As Double
    
    Dim mesa As String
    Dim añoa As String
    
    Dim saldo2 As Double
    Dim saldoan As Double
    mesa = Format(fechasistema, "mm")
    añoa = Format(fechasistema, "yyyy")
    Call generasaldoscuentascorrientes(empresaactiva, dato1.text, mesa, añoa)
    infogrilla.Grid1.AutoRedraw = False
        Set csql2.ActiveConnection = contadb
        If SALDOCT2.Value = True Then
        
        csql2.sql = "SELECT sc.tipo,sc.rut,sc.debeanterior-sc.haberanterior,ifnull(sa.anterior,0) as ante,ifnull(sa.debe,0) as debe,ifnull(sa.haber,0) as haber "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empresaactiva + ".saldosctacte as sc left join " + clientesistema + "conta.tempo_saldosctacte as sa "
        csql2.sql = csql2.sql + "on sa.rut=sc.rut and sa.tipo=sc.tipo "
        csql2.sql = csql2.sql + "where sc.año='" + añoa + "' and (sc.debeanterior-sc.haberanterior<>'0' or sa.anterior<>'0' or sa.debe<>'0' or sa.haber<>'0') "
        
        If Option2.Value = True Then
            csql2.sql = csql2.sql & " AND sc.rut IN (SELECT LPAD(REPLACE(rut,'-',''),10,0) AS rut FROM " & clientesistema & "conta.maestroempresas) "
        End If
        If Option3.Value = True Then
            csql2.sql = csql2.sql & " AND sc.rut NOT IN (SELECT LPAD(REPLACE(rut,'-',''),10,0) AS rut FROM " & clientesistema & "conta.maestroempresas ) "
        End If
        
        If cttodatipo.Value = True Then
            csql2.sql = csql2.sql + " and  sc.tipo='" + dato1.text + "' "
        End If
            csql2.sql = csql2.sql + " order by sc.tipo,sc.rut "
        Else
        csql2.sql = "SELECT sc.tipo,sc.rut,sc.nombre "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes as sc "
        If cttodatipo.Value = True Then
            csql2.sql = csql2.sql + " where  sc.tipo='" + dato1.text + "' "
        End If
        csql2.sql = csql2.sql + " order by sc.tipo,sc.nombre "
       
       
       End If
       
        csql2.Execute
        LINEAS = 0
        
        suma1 = 0
        suma2 = 0
        suma3 = 0
        suma4 = 0
       
       barra.Value = 0
        If csql2.RowsAffected > 0 Then
        barra.Max = csql2.RowsAffected + 1
        
        Set resultados2 = csql2.OpenResultset
        
        
        While Not resultados2.EOF
        If tiposaldo <> resultados2(0) Then
        barra.Max = barra.Max + 1
        barra.Value = barra.Value + 1
        LINEAS = LINEAS + 1
        
        infogrilla.Grid1.Rows = LINEAS + 1
        infogrilla.Grid1.Range(LINEAS, 1, LINEAS, 2).FontBold = True
        infogrilla.Grid1.Cell(LINEAS, 1).text = resultados2(0)
        infogrilla.Grid1.Cell(LINEAS, 2).text = leerNombreMayor(resultados2(0))
        tiposaldo = resultados2(0)
        End If
               
      Rem   If resultados2(3) + resultados2(4) + resultados2(5) + resultados2(6) <> 0 Then
        barra.Value = infogrilla.Grid1.Rows - 1
        LINEAS = LINEAS + 1
        infogrilla.Grid1.Rows = LINEAS + 1
        infogrilla.Grid1.Cell(LINEAS, 1).text = Mid(resultados2(1), 1, 9) + "-" + Mid(resultados2(1), 10, 1)
        infogrilla.Grid1.Cell(LINEAS, 2).text = leerNombrerut(resultados2(0), resultados2(1))
        
        
       If SALDOCT2.Value = True Then
              
        infogrilla.Grid1.Cell(LINEAS, 3).text = resultados2(2) + resultados2(3)
        infogrilla.Grid1.Cell(LINEAS, 4).text = resultados2(4)
        infogrilla.Grid1.Cell(LINEAS, 5).text = resultados2(5)
        infogrilla.Grid1.Cell(LINEAS, 6).text = resultados2(2) + resultados2(3) + resultados2(4) - resultados2(5)
        suma1 = suma1 + infogrilla.Grid1.Cell(LINEAS, 3).text
        suma2 = suma2 + infogrilla.Grid1.Cell(LINEAS, 4).text
        suma3 = suma3 + infogrilla.Grid1.Cell(LINEAS, 5).text
        suma4 = suma4 + infogrilla.Grid1.Cell(LINEAS, 6).text
        
        End If
        Rem End If
    
        
        resultados2.MoveNext
        Wend
          resultados2.Close
            Set resultados2 = Nothing
'       If cSql2.RowsAffected > 0 And SALDOCT1.Value = False Then LINEAS = LINEAS + 2: Call totaltipo(PASO, saldo2, LINEAS)

        End If
       If SALDOCT2.Value = True Then
        LINEAS = LINEAS + 2
        infogrilla.Grid1.Rows = LINEAS + 1
        infogrilla.Grid1.Range(LINEAS, 1, LINEAS, 6).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Range(LINEAS, 1, LINEAS, 6).FontBold = True
              
        infogrilla.Grid1.Cell(LINEAS, 2).text = "TOTALES "
        infogrilla.Grid1.Cell(LINEAS, 3).text = suma1
        infogrilla.Grid1.Cell(LINEAS, 4).text = suma2
        infogrilla.Grid1.Cell(LINEAS, 5).text = suma3
        infogrilla.Grid1.Cell(LINEAS, 6).text = suma4
        End If
     infogrilla.Grid1.AutoRedraw = True
     infogrilla.Grid1.Refresh
     
        

End Sub
Sub titulotipo(infogrilla As grillainformes, tipo, NOMBRE, LINEA)
infogrilla.Grid1.Rows = LINEA + 2
With infogrilla.Grid1.Range(LINEA + 1, 0, LINEA + 1, infogrilla.Grid1.Cols - 1)
        .FontBold = True
        .FontUnderline = True
        End With
infogrilla.Grid1.Cell(LINEA, 1).text = ""
infogrilla.Grid1.Cell(LINEA, 2).text = ""

infogrilla.Grid1.Cell(LINEA + 1, 1).text = tipo
infogrilla.Grid1.Cell(LINEA + 1, 2).text = NOMBRE
dedonde = 1
Call LEERSALDOS(tipo)
dedonde = 2
If saldo >= 0 And SALDOCT1.Value = False Then infogrilla.Grid1.Cell(LINEA + 1, 3).text = saldo
If saldo < 0 And SALDOCT1.Value = False Then infogrilla.Grid1.Cell(LINEA + 1, 4).text = saldo * -1

End Sub


Sub SALDOSctacte(tipo, rut)

Call LEERSALDOSCTACTE(tipo, rut)
sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
For k = 1 To CDbl(Format(fechasistema, "mm"))
sumador = sumador + Val(sqlconta.response(k + 4, 3)) - Val(sqlconta.response(k + 16, 3))
Next k
saldo = sumador
sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
For k = 1 To CDbl(Format(fechasistema, "mm") - 1)
sumador = sumador + Val(sqlconta.response(k + 4, 3)) - Val(sqlconta.response(k + 16, 3))
Next k
SALDOANTERIOR = sumador
totaldebe = 0
totalhaber = 0


For k = CDbl(Format(fechasistema, "mm")) To CDbl(Format(fechasistema, "mm"))
totaldebe = Val(sqlconta.response(k + 4, 3))
totalhaber = Val(sqlconta.response(k + 16, 3))
Next k

End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Sub LEERSALDOSFECHA(cuenta, fecha1, fecha2)
Dim resultados3 As rdoResultset
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
        
        Set cSql3.ActiveConnection = contadb
        cSql3.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
        If dedonde = 1 Then cSql3.sql = cSql3.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and fecha>'" + fecha1 + "' and fecha<'" + fecha2 + "'"
        If dedonde = 2 Then cSql3.sql = cSql3.sql + "FROM movimientoscontables where tipoctacte='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha>'" + fecha1 + "' and fecha<'" + fecha2 + "'"
        If dedonde = 3 Then cSql3.sql = cSql3.sql + "FROM movimientoscontables where centrocosto='" + cuenta + "' and fecha>'" + fecha1 + "' and fecha<'" + fecha2 + "'"
        
        cSql3.sql = cSql3.sql + "order by codigocuenta,fecha"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
             
             resultados3.MoveNext
           
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If
 
End Sub

Private Sub dato1_LostFocus()
LEETIPOCTACTE

End Sub

Sub LEERSALDOSCTACTE(tipoctacte, cuenta)
   Dim resultados3 As rdoResultset
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = "debe01"
    campos(6, 0) = "debe02"
    campos(7, 0) = "debe03"
    campos(8, 0) = "debe04"
    campos(9, 0) = "debe05"
    campos(10, 0) = "debe06"
    campos(11, 0) = "debe07"
    campos(12, 0) = "debe08"
    campos(13, 0) = "debe09"
    campos(14, 0) = "debe10"
    campos(15, 0) = "debe11"
    campos(16, 0) = "debe12"
    campos(17, 0) = "haber01"
    campos(18, 0) = "haber02"
    campos(19, 0) = "haber03"
    campos(20, 0) = "haber04"
    campos(21, 0) = "haber05"
    campos(22, 0) = "haber06"
    campos(23, 0) = "haber07"
    campos(24, 0) = "haber08"
    campos(25, 0) = "haber09"
    campos(26, 0) = "haber10"
    campos(27, 0) = "haber11"
    campos(28, 0) = "haber12"
    campos(29, 0) = ""
    condicion = "tipo=" + "'" + tipoctacte + "' and rut='" + cuenta + "' and año='" + Format(fechasistema, "yyyy") + "'"
    campos(0, 2) = "saldosctacte"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
    saldo = sumador
    Rem acumula fecha
        fecha1 = Format(fechasistema, "yyyy") + "-01-01"

    
        
        Set cSql3.ActiveConnection = contadb
        cSql3.sql = "SELECT SUM(monto),dh "
        cSql3.sql = cSql3.sql + "FROM movimientoscontables where codigocuenta='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha<'" + fecha1 + "' and fecha>='" + Format(fechasistema, "yyyy") + "-01-01" + "' "
        cSql3.sql = cSql3.sql + "GROUP BY DH"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
         If resultados3(1) = "D" Then saldo = saldo + resultados3(0)
         If resultados3(1) = "H" Then saldo = saldo - resultados3(0)
         resultados3.MoveNext
           
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If
End Sub

Private Sub botonmisaccesos_Click()
    programafiltro = Me.Caption
    misaccesos.Show
End Sub

Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

