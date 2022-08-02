VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form cambioLocal 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Local Activo"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1395
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   2461
      BackColor       =   16744576
      Caption         =   "Configurar Local Activo"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Local Actual"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nuevo Local"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lblActual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   420
         Width           =   4515
      End
      Begin VB.Label lblNuevo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   900
         Width           =   3675
      End
   End
   Begin XPFrame.FrameXp frmAceptar 
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "A   C   E   P   T   A   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "cambioLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private contador As Integer

Private Sub Cambiar()
    If dato1.text <> "" And lblNuevo.Caption <> "" Then
        empresaActiva = dato1.text
        
        rubro = leerRubro(empresaActiva)
        Call ConectarRubro(servidor, baseDatos, usuario, password)
        Call Conectarventas(servidor, baseVentas & empresaActiva, usuario, password)
        If empresaActiva = "00" Then
        bodega = "00"
        Else
        bodega = "01"
        End If
        Unload Me
    End If
End Sub

Private Sub dato1_GotFocus()
    Call selecciona(dato1)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call ayudaLocalesRubro(dato1)
    Else
        Call Flechas(KeyCode, dato1)
    End If
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If contador = 0 Then
            dato1.text = ceros(dato1)
            lblNuevo.Caption = leerNombreEmpresa(dato1.text)
            If lblNuevo.Caption <> "" Then
                SendKeys "{Tab}"
                contador = 1
            End If
        Else
            Call Cambiar
        End If
    End If
End Sub

Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 38 Then
        If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    contador = 0
    lblActual = leerNombreEmpresa(empresaActiva)
End Sub

    Private Sub frmAceptar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmAceptar)
        frmAceptar.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmAceptar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmAceptar)
        frmAceptar.CaptionEstilo3D = Inserted
        Call Cambiar
    End Sub
