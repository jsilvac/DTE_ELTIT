VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form CProtestados 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheques Protestados"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6300
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   3255
      Left            =   120
      TabIndex        =   16
      Top             =   135
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5741
      BackColor       =   16744576
      Caption         =   "Datos Protesto"
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
      Begin XPFrame.FrameXp frmComentario 
         Height          =   675
         Left            =   2880
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   1.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Poner fecha distinta a fecha cheque para cancelar."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   3075
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox dato3 
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
         Left            =   4560
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   720
         Width           =   315
      End
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
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox dato2 
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
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox dato4 
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
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox dato5 
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
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox dato6 
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   555
      End
      Begin VB.TextBox dato7 
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
         Left            =   4560
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox dato8 
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
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox dato9 
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
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox dato10 
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   1800
         Width           =   555
      End
      Begin VB.TextBox dato11 
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
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox dato15 
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
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "proveedor"
         Top             =   2880
         Width           =   4575
      End
      Begin VB.TextBox dato14 
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "proveedor"
         Top             =   2520
         Width           =   555
      End
      Begin VB.TextBox dato13 
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
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "proveedor"
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox dato12 
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
         TabIndex        =   11
         Tag             =   "proveedor"
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lbl10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sucursal"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " N' Cheque"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Protesto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Cliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblNombre 
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
         Left            =   120
         TabIndex        =   22
         Top             =   1005
         Width           =   5895
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Cheque"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblDv 
         Alignment       =   1  'Right Justify
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
         Left            =   2700
         TabIndex        =   20
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Motivo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cancelado"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lbl9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Glosa"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   1215
      End
   End
   Begin VB.PictureBox manual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   0
      Width           =   555
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1275
      Left            =   360
      TabIndex        =   15
      Top             =   3480
      Width           =   5955
      _cx             =   10504
      _cy             =   2249
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "CProtestados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private p As protesto
    Private modifica As Boolean

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    
    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
    End Sub
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
    End Sub
    
    Private Sub dato8_GotFocus()
        Call VerificarCajas(Me, dato8)
        Call selecciona(dato8)
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
    End Sub
    
    Private Sub dato10_GotFocus()
        Call VerificarCajas(Me, dato10)
        Call selecciona(dato10)
    End Sub
    
    Private Sub dato11_GotFocus()
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
    End Sub
    
    Private Sub dato12_GotFocus()
        Call VerificarCajas(Me, dato12)
        Call selecciona(dato12)
        If frmComentario.Visible = False Then
            frmComentario.Visible = True
        End If
    End Sub
    
    Private Sub dato13_GotFocus()
        Call VerificarCajas(Me, dato13)
        Call selecciona(dato13)
        If frmComentario.Visible = False Then
            frmComentario.Visible = True
        End If
    End Sub
    
    Private Sub dato14_GotFocus()
        Call VerificarCajas(Me, dato14)
        Call selecciona(dato14)
        If frmComentario.Visible = False Then
            frmComentario.Visible = True
        End If
    End Sub
    
    Private Sub dato15_GotFocus()
        Call VerificarCajas(Me, dato15)
        Call selecciona(dato15)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCliente(dato2, dato3, lblDV)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato5)
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato7)
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato8)
    End Sub
    
    Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato9)
    End Sub
    
    Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato10)
    End Sub
    
    Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato11)
    End Sub
    
    Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato12)
    End Sub
    
    Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato13)
    End Sub
    
    Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato14)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato2.text <> "" Then
            dato2.text = ceros(dato2)
            lblDV.Caption = rut(dato2.text)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If leerProtesto(p, dato1.text, dato2.text & lblDV.Caption, dato3.text, "=") = True Then
                Call structtoctrl
                opciones.SetFocus
            Else
                Call HabilitarCajas(Me, modifica)
                lblNombre.Caption = leerNombreClienteSucursal(dato2.text & lblDV.Caption, dato3.text)
                If lblNombre.Caption <> "" Then
                    SendKeys "{Tab}"
                End If
            End If
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato8.text = ceros(dato8)
            If dato8.text = "00" Then
                dato8.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato9.text = ceros(dato9)
            If dato9.text = "00" Then
                dato9.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato10.text = ceros(dato10)
            If dato10.text = "0000" Then
                dato10.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato12_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato12.text = ceros(dato12)
            If dato12.text = "00" Then
                If modifica = True Then
                    dato12.text = Format(fechasistema, "dd")
                End If
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato13_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato13.text = ceros(dato13)
            If dato13.text = "00" Then
                If modifica = True Then
                    dato13.text = Format(fechasistema, "mm")
                End If
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato14_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato14.text = ceros(dato14)
            If dato14.text = "0000" Then
                If modifica = True Then
                    dato14.text = Format(fechasistema, "yyyy")
                End If
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato15_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            Call ctrltostruct
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
'    Private Sub dato4_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato4.text) = dato4.MaxLength Then
'            Call dato4_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato5_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato5.text) = dato5.MaxLength Then
'            Call dato5_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato6_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato6.text) = dato6.MaxLength Then
'            Call dato6_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato8_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato8.text) = dato8.MaxLength Then
'            Call dato8_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato9_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato9.text) = dato9.MaxLength Then
'            Call dato9_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato10_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato10.text) = dato10.MaxLength Then
'            Call dato10_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato12_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato12.text) = dato12.MaxLength Then
'            Call dato12_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato13_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato13.text) = dato13.MaxLength Then
'            Call dato13_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato14_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato14.text) = dato14.MaxLength Then
'            Call dato14_KeyPress(13)
'        End If
'    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato2_LostFocus()
        Call limpiaBarra(2)
    End Sub
Private Sub dato4_LostFocus()
Call esfecha(dato4, dato5, dato6, "dd")
End Sub
Private Sub dato5_LostFocus()
Call esfecha(dato4, dato5, dato6, "mm")
End Sub
Private Sub dato6_LostFocus()
Call esfecha(dato4, dato5, dato6, "yyyy")
End Sub
Private Sub dato8_LostFocus()
Call esfecha(dato8, dato9, dato10, "dd")
End Sub
Private Sub dato9_LostFocus()
Call esfecha(dato8, dato9, dato10, "mm")
End Sub
Private Sub dato10_LostFocus()
Call esfecha(dato8, dato9, dato10, "yyyy")
End Sub

    Private Sub dato12_LostFocus()
        If frmComentario.Visible = True Then
            frmComentario.Visible = False
        End If
        Call esfecha(dato12, dato13, dato14, "dd")

    End Sub
    
    Private Sub dato13_LostFocus()
        If frmComentario.Visible = True Then
            frmComentario.Visible = False
        End If
        Call esfecha(dato12, dato13, dato14, "mm")

    End Sub
    
    Private Sub dato14_LostFocus()
        If frmComentario.Visible = True Then
            frmComentario.Visible = False
        End If
        Call esfecha(dato12, dato13, dato14, "yyyy")
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

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
        modifica = False
        Call Centrar(Me)
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        p.cheque = dato1.text
        p.rut = dato2.text & lblDV.Caption
        p.sucursal = dato3.text
        p.fechaprotesto = dato6.text & dato5.text & dato4.text
        p.MONTO = dato7.text
        p.fechacheque = dato10.text & dato9.text & dato8.text
        p.motivo = dato11.text
        If modifica = False Then
            p.CANCELADO = dato10.text & dato9.text & dato8.text
        Else
            p.CANCELADO = dato14.text & dato13.text & dato12.text
        End If
        p.glosa = dato15.text
        Call grabarProtesto(p, modifica)
        Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        Dim cad As String
        dato1.text = p.cheque
        cad = p.rut
        dato2.text = Left(cad, Len(cad) - 1)
        lblDV.Caption = Right(cad, 1)
        dato3.text = p.sucursal
        dato4.text = Format(p.fechaprotesto, "dd")
        dato5.text = Format(p.fechaprotesto, "mm")
        dato6.text = Format(p.fechaprotesto, "yyyy")
        dato7.text = p.MONTO
        dato8.text = Format(p.fechacheque, "dd")
        dato9.text = Format(p.fechacheque, "mm")
        dato10.text = Format(p.fechacheque, "yyyy")
        dato11.text = p.motivo
        If p.fechacheque = p.CANCELADO Then
            dato12.text = "00"
            dato13.text = "00"
            dato14.text = "0000"
        Else
            dato12.text = Format(p.CANCELADO, "dd")
            dato13.text = Format(p.CANCELADO, "mm")
            dato14.text = Format(p.CANCELADO, "yyyy")
        End If
        dato15.text = p.glosa
        lblNombre.Caption = leerNombreClienteSucursal(dato2.text & lblDV.Caption, dato3.text)
        Call DeshabilitarCajas(Me)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

Private Sub manual_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, Asc("r"), Asc("R")
            Call retorno
        Case Asc("a"), Asc("A"), 37
            Call anterior
        Case Asc("s"), Asc("S"), 39
            Call siguiente
        Case Asc("m"), Asc("M")
            Call modificar
        Case Asc("e"), Asc("E"), 46
            Call eliminar
    End Select
End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                Call modificar
            Case "elimina"
                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                Call eliminar
                End If
            Case "imprime"
            Case "movimientos"
            Case "historico"
            Case "retorno"
                Call retorno
            Case "anterior"
                Call anterior
            Case "siguiente"
                Call siguiente
        End Select
    End Sub
    
    Private Sub modificar()
        modifica = True
        Call HabilitarCajas(Me, modifica)
        dato1.Enabled = False
        dato2.Enabled = False
        dato3.Enabled = False
        dato4.SetFocus
    End Sub
    
    Private Sub eliminar()
        Call eliminarProtesto(p)
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(CProtestados)
        Call LimpiarLabels(CProtestados)
        modifica = False
        Call DeshabilitarCajas(Me)
        dato1.SetFocus
    End Sub
        
    Private Sub anterior()
        If leerProtesto(p, dato1.text, dato2.text & lblDV.Caption, dato3.text, "<") = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerProtesto(p, dato1.text, dato2.text & lblDV.Caption, dato3.text, ">") = True Then
            structtoctrl
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

Private Sub opciones_GotFocus()
    manual.SetFocus
End Sub
