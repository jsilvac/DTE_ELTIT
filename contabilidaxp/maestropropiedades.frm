VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form arriendo01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Propiedades"
   ClientHeight    =   7035
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   8370
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   5280
      TabIndex        =   31
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5010
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8837
      BackColor       =   16744576
      Caption         =   "DATOS  "
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
      Alignment       =   1
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "codigopropiedad"
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "direccion"
         Top             =   1035
         Width           =   6015
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "nombrepropiedad"
         Top             =   675
         Width           =   6015
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Comuna"
         Top             =   1395
         Width           =   3255
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "ciudad"
         Top             =   1755
         Width           =   3255
      End
      Begin VB.TextBox dato7 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   1
         TabIndex        =   7
         Tag             =   "tipoarriendo"
         Top             =   2475
         Width           =   375
      End
      Begin VB.TextBox dato8 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   1
         TabIndex        =   6
         Tag             =   "gastoscomunes"
         Top             =   2835
         Width           =   375
      End
      Begin VB.TextBox dato6 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "monto"
         Top             =   2115
         Width           =   1095
      End
      Begin VB.TextBox dato10 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "monedaarriendo"
         Top             =   3555
         Width           =   375
      End
      Begin VB.TextBox dato9 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1725
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "rutpropietario"
         Top             =   3195
         Width           =   1095
      End
      Begin VB.Label lblarrendatario 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   3195
         Width           =   4530
      End
      Begin VB.Label lblmoneda 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         Top             =   3555
         Width           =   4815
      End
      Begin VB.Label lbl 
         BackColor       =   &H00FF8080&
         Caption         =   " S - Si  -  N - No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   2835
         Width           =   4695
      End
      Begin VB.Label lbltipo 
         BackColor       =   &H00FF8080&
         Caption         =   " M - Mensual  -  A - Anual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   2475
         Width           =   4335
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   26
         Top             =   315
         Width           =   1530
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   25
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   24
         Top             =   1035
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comuna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   23
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ciudad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   22
         Top             =   1755
         Width           =   1530
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   21
         Top             =   2115
         Width           =   1530
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gastos Comunes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   100
         TabIndex        =   20
         Top             =   2835
         Width           =   1530
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut Propietario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   19
         Top             =   3195
         Width           =   1530
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Arriendo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   18
         Top             =   2475
         Width           =   1530
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda Arriendo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   100
         TabIndex        =   17
         Top             =   3555
         Width           =   1530
      End
      Begin VB.Label dv 
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         TabIndex        =   16
         Top             =   3195
         Width           =   255
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   8340
      TabIndex        =   14
      Top             =   7035
      Width           =   8370
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3735
      Left            =   8400
      TabIndex        =   11
      Top             =   240
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3495
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   12582912
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   16107953
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   16776436
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3735
         Left            =   0
         Top             =   0
         Width           =   4695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   900
      TabIndex        =   10
      Top             =   5580
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   3735
      Left            =   8520
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "arriendo01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private MODIFI As Integer

 
 

Private Sub dato1_GotFocus()
Call cargatexto(dato1)
End Sub

Private Sub dato2_GotFocus()
If MODIFI = 0 Then Call leer
Call cargatexto(dato2)
End Sub
Private Sub dato4_GotFocus()
Call cargatexto(dato4)
End Sub
Private Sub dato5_GotFocus()
Call cargatexto(DATO5)
End Sub
Private Sub dato6_GotFocus()
Call cargatexto(dato6)
End Sub
Private Sub dato7_GotFocus()
Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()
Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
Call cargatexto(dato9)
End Sub
Private Sub dato10_GotFocus()
 
 Call cargatexto(dato10)
End Sub
 
Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then Call ayudapropiedades(dato1)
    Call flechas(dato1, dato2, KeyCode)
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
       Call flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
      Call flechas(dato2, dato4, KeyCode)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, DATO5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(DATO5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudaarrendatarios(dato9)
        Call flechas(dato8, dato10, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudamonedas(dato10)
        Call flechas(dato9, dato10, KeyCode)
End Sub


 

 Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3

Rem Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)

 

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato1.text <> "" Then
'        Call ceros(dato1)
        Call Pregunta(dato1, dato2)
    End If
End Sub
Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato2.text) <> "" Then
        Call Pregunta(dato2, dato3)
    End If
End Sub
Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato3.text) <> "" Then Call Pregunta(dato3, dato4)
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato4.text) <> "" Then Call Pregunta(dato4, DATO5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(DATO5.text) <> "" Then Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato6.text <> "" Then Call Pregunta(dato6, dato7)
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato7.text) <> "" And (dato7.text = "M" Or dato7.text = "A") Then Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And dato8.text <> "" And (dato8.text = "S" Or dato8.text = "N") Then Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato9.text <> "" Then
        Call ceros(dato9)
        DV.Caption = rut(dato9)
          If leearrendatario(dato9.text & DV.Caption) = True Then
            Call Pregunta(dato9, dato10)
        Else
        MsgBox " PROPIETARIO NO EXISTE O RUT MAL INGRESADO", vbCritical, "ATENCION"
        End If
    End If
    

End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
   
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato10.text <> "" Then
        Call ceros(dato10)
        If leemonedas(dato10.text) = True Then
  
    If Verifica_Permiso(Me.Caption, "agrega") = True Then
      grabar
    End If
            retorno
        End If
    End If
End Sub
 

Sub leer()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = DATO5.Tag
    campos(5, 0) = dato6.Tag
    campos(6, 0) = dato7.Tag
    campos(7, 0) = dato8.Tag
    campos(8, 0) = dato9.Tag
    campos(9, 0) = dato10.Tag
    campos(10, 0) = ""
    campos(0, 2) = clientesistema & "arriendos" & ".maestro_propiedades"
    condicion = "codigopropiedad = '" & dato1.text & "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato2.SetFocus: GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
        
no:
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = DATO5.Tag
    campos(5, 0) = dato6.Tag
    campos(6, 0) = dato7.Tag
    campos(7, 0) = dato8.Tag
    campos(8, 0) = dato9.Tag
    campos(9, 0) = dato10.Tag
    campos(10, 0) = ""
    campos(0, 2) = clientesistema & "arriendos" & ".maestro_propiedades"
    condicion = " codigopropiedad > '" & dato1.text & "' order by codigopropiedad asc "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    
no:
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = DATO5.Tag
    campos(5, 0) = dato6.Tag
    campos(6, 0) = dato7.Tag
    campos(7, 0) = dato8.Tag
    campos(8, 0) = dato9.Tag
    campos(9, 0) = dato10.Tag
    campos(10, 0) = ""
    campos(0, 2) = clientesistema & "arriendos" & ".maestro_propiedades"
    condicion = " codigopropiedad < '" & dato1.text & "' order by codigopropiedad desc "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    
no:
   
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato3.text = sqlconta.response(2, 3)
    dato4.text = sqlconta.response(3, 3)
    DATO5.text = sqlconta.response(4, 3)
    dato6.text = sqlconta.response(5, 3)
    dato7.text = sqlconta.response(6, 3)
    If sqlconta.response(7, 3) = 0 Then dato8.text = "N" Else dato8.text = "S"
    dato9.text = Mid(sqlconta.response(8, 3), 1, 9)
    DV.Caption = rut(dato9.text)
    Call leearrendatario(sqlconta.response(8, 3))
    dato10.text = sqlconta.response(9, 3)
    Call leemonedas(dato10.text)
  
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    DATO5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
    dato9.Locked = condicion
    dato10.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    dato10.Enabled = condicion
   
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudapropiedades(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigopropiedad", "direccion")
    largo = Array("11s", "40s")
    cfijo = "codigopropiedad like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Propiedades"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_propiedades", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub ayudamonedas(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigomoneda", "nombremoneda")
    largo = Array("11s", "40s")
    cfijo = "codigomoneda like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Monedas"
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_monedas", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = DATO5.Tag
    campos(5, 0) = dato6.Tag
    campos(6, 0) = dato7.Tag
    campos(7, 0) = dato8.Tag
    campos(8, 0) = dato9.Tag
    campos(9, 0) = dato10.Tag
    campos(10, 0) = ""
   
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = dato3.text
    campos(3, 1) = dato4.text
    campos(4, 1) = DATO5.text
    campos(5, 1) = dato6.text
    campos(6, 1) = dato7.text
    If dato8.text = "S" Then
        campos(7, 1) = "1"
    Else
        campos(7, 1) = "0"
    End If
    
    campos(8, 1) = dato9.text & DV.Caption
    campos(9, 1) = dato10.text
    campos(0, 2) = clientesistema & "arriendos" & ".maestro_propiedades"
    If MODIFI = 1 Then condicion = "codigopropiedad ='" & dato1.text & "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
Sub ELIMINAR()
   campos(0, 2) = clientesistema & "arriendos" & ".maestro_propiedades"
    condicion = "codigopropiedad ='" & dato1.text & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub
  

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
 
If command = "modifica" Then
    If Verifica_Permiso(Me.Caption, "modifica") = True Then
        modifica
    End If
End If

If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        ELIMINA
    End If
End If

If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
 



End Sub
Sub ELIMINA()
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus
 
End Sub

Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
dato2.SetFocus
MODIFI = 1

End Sub
Sub retorno()

disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
no:

    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    DV.Caption = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    LBLMONEDA.Caption = ""
    lblarrendatario.Caption = ""
End Sub
 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub

Private Function leemonedas(codigo) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select nombremoneda from " & clientesistema & "arriendos" & ".maestro_monedas where codigomoneda='" & codigo & "'"
csql.Execute
leemonedas = False
If csql.RowsAffected > 0 Then
leemonedas = True
Set resultados = csql.OpenResultset
LBLMONEDA.Caption = resultados(0)
Else
LBLMONEDA.Caption = ""

End If
Set resultados = Nothing
csql.Close
Set csql = Nothing

End Function

Private Function leearrendatario(rutarrendatario) As Boolean
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select nombre from " & clientesistema & "arriendos" & ".maestro_arrendadores "
 csql.sql = csql.sql & "where rut='" & rutarrendatario & "' "
 csql.Execute
 leearrendatario = False
 
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leearrendatario = True
    lblarrendatario.Caption = resultados(0)
 Else
    lblarrendatario.Caption = ""
 End If
 
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Function
 

Sub ayudaarrendatarios(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "rut like '%%'"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendatarios"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_arrendatarios", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
