VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form consumo02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Proveedores"
   ClientHeight    =   6705
   ClientLeft      =   4575
   ClientTop       =   3945
   ClientWidth     =   7065
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   3960
      TabIndex        =   22
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4845
      Left            =   135
      TabIndex        =   12
      Top             =   135
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   8546
      BackColor       =   16773879
      Caption         =   "Maestro de Proveedores"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato1 
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
         Left            =   1590
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "rut"
         Top             =   420
         Width           =   1095
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
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "nombre"
         Top             =   780
         Width           =   4950
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
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "comuna"
         Top             =   1500
         Width           =   2415
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
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "direccion"
         Top             =   1140
         Width           =   4950
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
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   5
         Tag             =   "ciudad"
         Top             =   1860
         Width           =   2415
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
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "fono1"
         Top             =   2220
         Width           =   1575
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
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "fono2"
         Top             =   2580
         Width           =   1575
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
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "fax"
         Top             =   2940
         Width           =   1575
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
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "contacto"
         Top             =   3300
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
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
         Height          =   285
         Left            =   270
         TabIndex        =   21
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
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
         Height          =   285
         Left            =   270
         TabIndex        =   20
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Direccion"
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
         Height          =   285
         Left            =   270
         TabIndex        =   19
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comuna"
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
         Height          =   285
         Left            =   270
         TabIndex        =   18
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
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
         Height          =   285
         Left            =   270
         TabIndex        =   17
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono1"
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
         Height          =   285
         Left            =   270
         TabIndex        =   16
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono2"
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
         Height          =   285
         Left            =   270
         TabIndex        =   15
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fax"
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
         Height          =   285
         Left            =   270
         TabIndex        =   14
         Top             =   2940
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Mail"
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
         Height          =   285
         Left            =   270
         TabIndex        =   13
         Top             =   3300
         Width           =   1215
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2790
         TabIndex        =   1
         Top             =   420
         Width           =   285
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin MSAdodcLib.Adodc PROVEEDOR_DATA 
      Height          =   375
      Left            =   2880
      Top             =   6795
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   225
      TabIndex        =   10
      Top             =   5355
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
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   4815
      Left            =   360
      Top             =   270
      Width           =   6600
   End
End
Attribute VB_Name = "consumo02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CUENTAPROVEEDOR As String
Private MODIFI As Integer

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


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaProveedor(dato2)
    Call flechas(dato1, dato2, KeyCode)
    If KeyCode = 27 Then Unload Me
    If KeyCode = 38 Then Unload Me
End Sub

Private Sub dato1_LostFocus()
    DV.Caption = rut(dato1)
    If sl = 0 Then Call leer
    sl = 0
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
    Call flechas(dato8, dato9, KeyCode)
End Sub


Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000
    
    
    sc = 0
    opciones.Visible = False
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato3, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato4, DATO5)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(DATO5, dato6)
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato6, dato7)
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato7, dato8)
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato8, dato9)
End Sub

Private Sub dato9_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    sc = 1
    grabar
    End If
    
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'NOMBRE
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'COMUNA
    campos(4, 0) = DATO5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = ""
    
    campos(0, 2) = clientesistema + "consumos_basicos.proveedores"
    condicion = "rut=" + "'" + dato1.text + DV.Caption + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        Call carga
        opciones.Visible = True
        Call disponible(True)
        Call habilita(True)
        opciones.SetFocus
    Else
        If Verifica_Permiso(Me.Caption, "agrega") = True Then
            dato2.Enabled = True
            dato2.SetFocus
        Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
            dato1.SelStart = 0
            dato1.SelLength = Len(dato1.text)
            dato1.SetFocus
        End If
    End If
    
End Sub

Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'NOMBRE
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'COMUNA
    campos(4, 0) = DATO5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = ""
    
    campos(0, 2) = clientesistema + "consumos_basicos.proveedores"
    condicion = "rut>" + "'" + dato1.text + DV.Caption + "' ORDER BY rut"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub

Sub leeranterior()
    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'NOMBRE
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'COMUNA
    campos(4, 0) = DATO5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = ""
    
    campos(0, 2) = clientesistema + "consumos_basicos.proveedores"
    condicion = "rut<" + "'" + dato1.text + DV.Caption + "' ORDER BY rut DESC"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub

Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 9)
    DV.Caption = Mid(sqlconta.response(0, 3), 10, 1)
    dato2.text = sqlconta.response(1, 3)
    dato3.text = sqlconta.response(2, 3)
    dato4.text = sqlconta.response(3, 3)
    DATO5.text = sqlconta.response(4, 3)
    dato6.text = sqlconta.response(5, 3)
    dato7.text = sqlconta.response(6, 3)
    dato8.text = sqlconta.response(7, 3)
    dato9.text = sqlconta.response(8, 3)
    

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
    
End Sub
'Sub Conecta_Maestro_Secciones()
'    'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
'    With maestro02
'        .mpro.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=basesdedatos"
'    End With
'End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub ayudaProveedor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("13n", "30s")
    cfijo = "no"
    mensajeAyuda = "Ayuda Proveedores"
    cabezas = Array("rut", "nombre")

    Call cargaAyudaT(Servidor, basedatos, Usuario, password, clientesistema + "consumos_basicos.proveedores", dato1, campos, cfijo, largo, 2)
    If dato1.text = "" Then dato1.SetFocus: GoTo no:
    caja.Enabled = True
    caja.SetFocus
no:
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()

    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'NOMBRE
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'COMUNA
    campos(4, 0) = DATO5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = ""
    
    campos(0, 1) = dato1.text & DV.Caption 'RUT
    campos(1, 1) = dato2.text 'NOMBRE
    campos(2, 1) = dato3.text 'DIRECCION
    campos(3, 1) = dato4.text 'COMUNA
    campos(4, 1) = DATO5.text '5
    campos(5, 1) = dato6.text '6
    campos(6, 1) = dato7.text '7
    campos(7, 1) = dato8.text '8
    campos(8, 1) = dato9.text '9
    
    campos(0, 2) = clientesistema + "consumos_basicos.proveedores"
    If MODIFI = 1 Then condicion = "rut = '" & dato1.text & DV.Caption & "'"
    If MODIFI = 1 Then op = 3 Else op = 2: condicion = ""
    MODIFI = 0
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    status = sqlconta.status
    If op = 3 Then Call leer
    If op = 2 Then Call retorno

End Sub

Sub ELIMINAR()
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        campos(0, 2) = clientesistema + "consumos_basicos.proveedores"
        condicion = "rut=" + "'" + dato1.text + DV.Caption + "'"
        op = 4
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
    Else
        MsgBox "no permiso para eliminar ", vbCritical + vbOKOnly, "Permiso Denegado"
        dato1.SelStart = 0
        dato1.SelLength = Len(dato1.text)
        dato1.SetFocus
    End If
End Sub

Private Sub Label9_Click()

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
    If command = "retorno" Then retorno
    If command = "modifica" Then
        If Verifica_Permiso(Me.Caption, "modifica") = True Then
            Call disponible(True)
            Call habilita(False)
            dato1.Enabled = False
            dato2.SetFocus
            MODIFI = 1
        Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
        End If
    End If
    If command = "elimina" Then
            If MsgBox("REALMENTE DESEA ELIMINAR", vbYesNo) = vbYes Then
                If Verifica_Permiso(Me.Caption, "elimina") = True Then
                    Call disponible(True)
                    Call habilita(False)
                    Call ELIMINAR
                    Call retorno
                    dato1.SetFocus
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
            End If
        
    End If
    If command = "siguiente" Then leersiguiente
    If command = "anterior" Then leeranterior
End Sub

Sub retorno()
    Call disponible(True)
    Call habilita(False)
    Call limpia
    opciones.Visible = False
    dato1.SetFocus
End Sub

Sub limpia()
    DV.Caption = ""
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    
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
