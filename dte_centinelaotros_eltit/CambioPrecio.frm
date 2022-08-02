VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form CambioPrecio 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantalla Consulta precios"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6855
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6255
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Historico de Cambios Precios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6165
      Width           =   2085
   End
   Begin VB.CommandButton retorno 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6255
      Width           =   1815
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5835
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   10292
      BackColor       =   12648384
      Caption         =   "Información del Producto"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPFrame.FrameXp MODIPRE 
         Height          =   2085
         Left            =   45
         TabIndex        =   16
         Top             =   3690
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   3678
         BackColor       =   49344
         Caption         =   "MODIFICA PRECIOS"
         CaptionEstilo3D =   1
         BackColor       =   49344
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
         Begin VB.TextBox dato2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4230
            MaxLength       =   9
            TabIndex        =   1
            Top             =   360
            Width           =   1995
         End
         Begin VB.TextBox dato3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4230
            MaxLength       =   9
            TabIndex        =   2
            Top             =   900
            Width           =   1995
         End
         Begin VB.TextBox dato4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4230
            MaxLength       =   9
            TabIndex        =   3
            Top             =   1485
            Width           =   1995
         End
         Begin VB.Label lbl4 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NUEVO PRECIO PUBLICO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   90
            TabIndex        =   19
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label lbl5 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NUEVO PRECIO MAYORISTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   90
            TabIndex        =   18
            Top             =   900
            Width           =   3975
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PRECIO CARACOL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   90
            TabIndex        =   17
            Top             =   1485
            Width           =   3975
         End
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1980
         MaxLength       =   13
         TabIndex        =   0
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRECIO CARACOL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   180
         TabIndex        =   13
         Top             =   3150
         Width           =   3975
      End
      Begin VB.Label precio3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   4320
         TabIndex        =   12
         Top             =   3150
         Width           =   1995
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRECIO MAYORISTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   180
         TabIndex        =   11
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label precio2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   4320
         TabIndex        =   10
         Top             =   2640
         Width           =   1995
      End
      Begin VB.Label precio1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   4320
         TabIndex        =   9
         Top             =   2100
         Width           =   1995
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRECIO PUBLICO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   180
         TabIndex        =   8
         Top             =   2100
         Width           =   3975
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1020
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Width           =   6180
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CODIGO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   180
         TabIndex        =   6
         Top             =   420
         Width           =   1635
      End
   End
End
Attribute VB_Name = "CambioPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Private segurity As Boolean

Private Sub Command1_Click()
cambiosdeprecio.Show vbModal

End Sub

Private Sub Command2_Click()
MODIPRE.Visible = True
dato2.SetFocus


End Sub

Private Sub Form_Activate()
    If segurity = True Then
        Seguridad.Show vbModal
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            Unload Me
        Case 38
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
    End Select
End Sub

Private Sub Form_Load()
    Call Centrar(Me)
    titCaption = Me.Caption
    'segurity = Not Verificar(usuarioSistema, passwordSistema)
    MODIPRE.Visible = False
    
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Productos"
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub
    
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaProductotxt(dato1)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
      Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        
        If KeyAscii > 65 And dato1.text = "" Then
         dato1.text = leeletra(Chr(KeyAscii))
         End If
         KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato1.text <> "" Then
            dato1.text = ceros(dato1)
            lblNombre.Caption = leerNombreProducto(dato1.text)
            precio1.Caption = leerPrecioProducto2(dato1.text, "01")
            precio2.Caption = leerPrecioProducto2(dato1.text, "02")
            precio3.Caption = leerPrecioProducto2(dato1.text, "03")
            dato2.text = leerPrecioProducto2(dato1.text, "01")
            dato3.text = leerPrecioProducto2(dato1.text, "02")
            dato4.text = leerPrecioProducto2(dato1.text, "03")
            
            If lblNombre.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
         
    
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumeroDecimal(dato3, KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumeroDecimal(dato3, KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumeroDecimal(dato4, KeyAscii)
        If KeyAscii = 13 Then
            Call modificarPrecio(dato1.text, dato2.text, "01", precio1.Caption)
            Call modificarPrecio(dato1.text, dato3.text, "02", precio2.Caption)
            Call modificarPrecio(dato1.text, dato4.text, "03", precio3.Caption)
            precio1.Caption = ""
            precio2.Caption = ""
            precio3.Caption = ""
            
            lblNombre.Caption = ""
            Call LimpiarCajas(Me)
            MODIPRE.Visible = False
            
            dato1.SetFocus
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
Private Sub Label2_Click()

End Sub

Private Sub retorno_Click()
            precio1.Caption = ""
            precio2.Caption = ""
            precio3.Caption = ""
            
            lblNombre.Caption = ""
            MODIPRE.Visible = False
            Call LimpiarCajas(Me)
            dato1.SetFocus

End Sub

Private Sub Text1_Change()

End Sub
