VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form digitaanalisis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digita Cuenta de Analisis"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmrut 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5106
      BackColor       =   16744576
      Caption         =   ""
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
      Begin VB.TextBox dato0 
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
         Left            =   1395
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "rut"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   0
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox DATO20 
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
         Left            =   1395
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "rut"
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton RUTOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FCE2E2&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "centro"
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
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblcentro 
         BackColor       =   &H00FCE2E2&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label lblglosa 
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lblcuenta 
         BackColor       =   &H00FF0000&
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label nombrectacte 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   6855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "codigo"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
   End
End
Attribute VB_Name = "digitaanalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ayudatipoconsumo(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigo", "nombre")
    largo = Array("11s", "40s")
    cfijo = "codigo like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Centros de Gastos"
       
    Call cargaAyudaT(servidor, clientesistema & "conta", usuario, password, ".presupuesto_centros", caja, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Private Sub dato0_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then Call ayudatipoconsumo(dato0)
    
End Sub

Private Sub dato0_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
 Call ceros(dato0)

If leerNOMBREgastos(dato0.text) <> "" Then
   
lblcentro.Caption = leerNOMBREgastos(dato0.text)
DATO20.SetFocus
Else
MsgBox "centro de gastos no creado "

dato0.SetFocus

End If



End If

End Sub

Private Sub dato20_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(DATO20)
   
    If leernombreanalisis(lblcuenta.Caption, DATO20.text) <> "" Then
    nombrectacte.Caption = leernombreanalisis(lblcuenta.Caption, DATO20.text)
    RUTOK.SetFocus
    
    End If
    
    
    
    End If
End Sub
Private Sub dato20_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    
    If KeyCode = vbKeyF2 Then Call ayudaanalisis(DATO20)
    End Sub


Private Sub Label1_Click()

End Sub

Sub ayudaanalisis(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigo", "nombre")
    largo = Array("6n", "40s")
    cfijo = "cuenta='" & lblcuenta.Caption & "' "
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Analsis " + lblglosa.Caption
    Call cargaAyudaT(servidor, clientesistema + "conta", usuario, password, "presupuesto_detalle", caja, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub




Private Sub Form_Unload(Cancel As Integer)
If nombrectacte.Caption = "" Then
Cancel = 1

End If
End Sub

Private Sub RUTOK_Click()
If nombrectacte.Caption <> "" Then
DIGITA_ANALISIS_CODIGO = DATO20.text
DIGITA_ANALISIS_NOMBRE = nombrectacte.Caption
DIGITA_CENTROS_CODIGO = dato0.text
DIGITA_CENTROS_NOMBRE = lblcentro.Caption

Unload Me


Else
DATO20.SetFocus
End If
End Sub
