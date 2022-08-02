VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ingreso05 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Ingreso de  Boletas o Zetas"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   1110
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   1560
      TabIndex        =   26
      Top             =   5160
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   10398
      BackColor       =   16744576
      Caption         =   "RESUMEN DE BOLETAS"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox chkcigarros 
         BackColor       =   &H00FF8080&
         Caption         =   "CIGARROS O RECARGAS"
         Height          =   495
         Left            =   480
         TabIndex        =   30
         Top             =   4200
         Width           =   2415
      End
      Begin VB.CheckBox CH_TBK 
         BackColor       =   &H00FF8080&
         Caption         =   "BOLETAS TRANSBANK"
         Height          =   495
         Left            =   3120
         TabIndex        =   29
         Top             =   4200
         Width           =   2415
      End
      Begin VB.TextBox DATO8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   22
         Text            =   "0"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox DATO10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox DATO4 
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox DATO9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   8
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox DATO7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "0"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox DATO6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00E1FFFD&
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00E1FFFD&
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
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "fecha"
         Top             =   360
         Width           =   375
      End
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
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "fecha"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label total 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   2040
         TabIndex        =   25
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL VENTAS"
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
         TabIndex        =   24
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MONTO EXENTO"
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
         TabIndex        =   23
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO ZETA"
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
         TabIndex        =   21
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO DE CAJA"
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
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CENTRO COSTO"
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
         TabIndex        =   19
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MONTO AFECTO"
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
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BOLETA FINAL"
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
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BOLETA INICIAL"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   13
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox LINEAS 
      Height          =   285
      Left            =   5760
      MaxLength       =   3
      TabIndex        =   12
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   11
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   5895
      _cx             =   10398
      _cy             =   2778
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
End
Attribute VB_Name = "ingreso05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private tipocuenta As String
    Private cc As Integer
    Private FORMATOGRILLA(100, 20)
    Private formatogrilla2(100, 20)
    Private cdi As Integer
    Private CANDO As Integer
    Private existe As String
    Private AFECTO As Double
    Private EXENTO As Double
    Private MODIFI As String
        
    Private AUXILIAR(1000, 3) As String
    
    Private respu As String
    Private tipoctacte As String
    Private nlineas As Double
    Private DOCU(6) As String
    Private grilladetalle(1000, 13) As String
    Private SALDOPE As Double
    Private NETO As Double
     Private CUENTAMAYOR(999) As String
     
     Private TIENECTACTE(999) As String
     Private TIENECRCC(999) As String
     Private TIENEBANCO(999) As String
     Private TIENEILA(999) As String
     Private TIENEICA(999) As String
     Private TIENEIHA(999) As String
     Private TIENEACTIVO(999) As String
     Private MES As String
     Private año As String
     
     
    
Private Sub COMMAND2_Click()

End Sub











 

Private Sub dato1_Change()
If Val(dato1.text) > 31 Then dato1.text = ""
End Sub


Private Sub dato1_LostFocus()
If dato1.text = "00" Then dato2.Enabled = True: dato3.Enabled = True: dato1.text = Mid(fechasistema, 1, 2): dato2.text = Mid(fechasistema, 4, 2): dato3.text = Mid(fechasistema, 7, 4): dato4.Enabled = True: dato4.SetFocus

End Sub

Private Sub dato2_Change()
If Val(dato4.text) > 12 Or Val(dato4.text) < 1 Then dato4.text = ""
End Sub

Private Sub DATO3_LostFocus()
If DATO5.text < "1900" Or DATO5.text > Format(fechasistema, "YYYY") Then DATO5.text = ""

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
Private Sub dato4_GotFocus()
Call cargatexto(dato4)
End Sub
Private Sub dato5_GotFocus()
leerboleta
Call cargatexto(DATO5)
End Sub
Private Sub dato6_GotFocus()
Call cargatexto(dato6)
End Sub
Private Sub dato7_GotFocus()
Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()
AFECTO = CDbl(Replace(dato7.text, ",", ""))
dato7.text = Format(AFECTO, "#,###,###,##0")
total.Caption = Format(AFECTO + EXENTO, "#,###,###,##0")


Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
EXENTO = CDbl(Replace(dato8.text, ",", ""))
dato8.text = Format(EXENTO, "#,###,###,##0")
total.Caption = Format(AFECTO + EXENTO, "#,###,###,##0")

Call cargatexto(dato9)
End Sub



Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call ayudacrcc(dato9)
    End If
End Sub
Sub ayudacrcc(caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    caja.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", caja, campos, cfijo, largo, 2)
 
End Sub
Private Sub Form_Load()
CENTRAR Me
iva = 19
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    sc = 0
    opciones.Visible = False

End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    Call flechas(dato1, dato2, KeyCode)
no:
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


Private Sub dato1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub
Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)

End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato4): Call Pregunta(dato4, DATO5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO5): Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    ' If KeyAscii = 42 And SUMADEBE = SUMAHABER Then grabarcomprobante:retorno: dato3.Enabled = True: dato3.SetFocus
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
no:
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato9)
        If existecrcc(dato9) = True Then
        Call Pregunta(dato9, dato10)
        Else
            MsgBox "CENTRO DE COSTO NO EXISTE POR FAVOR VERIFICAR", vbCritical, "ATENCION"
        End If
    End If
End Sub
Function existecrcc(dato) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select nombre from centrosdecosto "
    csql.sql = csql.sql & "where codigo='" & dato & "' and año='" & Format(fechasistema, "yyyy") & "' #"
    csql.Execute
    existecrcc = False
    If csql.RowsAffected > 0 Then
        existecrcc = True
    End If
    csql.Close
    Set csql = Nothing
    
End Function
Private Sub dato10_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato9): grabar: retorno
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus: caja.SelStart = 0
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus: caja.SelStart = 0
End Sub
Sub grabar()
    Dim netos As Double
    Dim DH As String
    Call ELIMINAR
    campos(0, 0) = "fecha"
    campos(1, 0) = "caja"
    campos(2, 0) = "boletainicial"
    campos(3, 0) = "boletafinal"
    campos(4, 0) = "monto"
    campos(5, 0) = "exento"
    campos(6, 0) = "centrocosto"
    campos(7, 0) = "numerozeta"
    campos(8, 0) = "total"
    campos(9, 0) = "estbk"
    campos(10, 0) = "cigarro"
    campos(11, 0) = ""
    
    
    
    campos(0, 1) = dato3.text + "-" + dato2.text + "-" + dato1.text
    campos(1, 1) = dato4.text
    campos(2, 1) = DATO5.text
    campos(3, 1) = dato6.text
    campos(4, 1) = Replace(dato7.text, ".", "")
    campos(5, 1) = Replace(dato8.text, ".", "")
    campos(6, 1) = dato9.text
    campos(7, 1) = dato10.text
    campos(8, 1) = Replace(total.Caption, ".", "")
    campos(9, 1) = CH_TBK.Value
    campos(10, 1) = chkcigarros.Value
    
    
    condicion = ""
    campos(0, 2) = "boletasdeventa"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    


End Sub


Sub leerboleta()
    campos(0, 0) = "fecha"
    campos(1, 0) = "caja"
    campos(2, 0) = "boletainicial"
    campos(3, 0) = "boletafinal"
    campos(4, 0) = "monto"
    campos(5, 0) = "exento"
    campos(6, 0) = "centrocosto"
    campos(7, 0) = "numerozeta"
    campos(8, 0) = "total"
    campos(9, 0) = "estbk"
    campos(10, 0) = "cigarro"
    campos(11, 0) = ""
    
    campos(0, 2) = "boletasdeventa"
    condicion = "fecha=" + "'" + dato3.text + "-" + dato2.text + "-" + dato1.text + "'" + " and caja=" + "'" + dato4.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then MODIFI = 1: carga: opciones.Visible = True: disponible (True):  opciones.SetFocus


End Sub
Sub carga()
    disponible (True)
    
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 4, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 7, 4)
    dato4.text = sqlconta.response(1, 3)
    DATO5.text = sqlconta.response(2, 3)
    dato6.text = sqlconta.response(3, 3)
    dato7.text = Format(sqlconta.response(4, 3), "##,###,###,##0")
    dato8.text = Format(sqlconta.response(5, 3), "##,###,###,##0")
    total.Caption = Format(sqlconta.response(8, 3), "##,###,###,##0")
    dato9.text = sqlconta.response(6, 3)
    dato10.text = sqlconta.response(7, 3)
    CH_TBK.Value = sqlconta.response(9, 3)
    chkcigarros.Value = sqlconta.response(10, 3)
End Sub



Sub ELIMINAR()
    
    campos(0, 2) = "boletasdeventa"
    condicion = "fecha=" + "'" + dato3.text + "-" + dato2.text + "-" + dato1.text + "'" + " and caja=" + "'" + dato4.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
no:
End Sub




Private Sub MANUAL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
    If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
    If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
    If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
    If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub


Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" And MODIFI = 0 Then retorno
If command = "retorno" And MODIFI = 1 Then grabar: retorno
If command = "modifica" Then
dato2.Enabled = True: dato2.SetFocus
End If


If command = "elimina" Then
If Verifica_Permiso(Me.Caption, "elimina") Then
ELIMINAR
retorno
Else
  MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
End If
End If


End Sub


Sub retorno()


opciones.Visible = False
limpia
disponible (False)

dato1.Enabled = True
dato1.SetFocus

End Sub


Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = "0"
    dato8.text = "0"
    dato9.text = ""
    dato10.text = "0"
    total.Caption = "0"
    CH_TBK.Value = 0
    chkcigarros.Value = 0
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus: caja.SelStart = 0
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

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
    
    total.Enabled = condicion
   
    
End Sub



Private Sub opciones_GotFocus()
MANUAL.SetFocus

End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
