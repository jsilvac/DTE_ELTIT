VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ventas01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Clientes"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   559
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   Begin MSAdodcLib.Adodc mpro 
      Height          =   375
      Left            =   7080
      Top             =   8520
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      Begin VB.TextBox dato15 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   34
         Tag             =   "cupo"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox dato14 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   32
         Tag             =   "descuento"
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox dato13 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   30
         Tag             =   "credito"
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   28
         Tag             =   "nombre"
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox dato12 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   26
         Tag             =   "plazo"
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox dato11 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   25
         Tag             =   "contacto"
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox dato10 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   24
         Tag             =   "giro"
         Top             =   3120
         Width           =   4695
      End
      Begin VB.TextBox dato9 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "fax"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox dato8 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "fono2"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox dato7 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "fono1"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox dato6 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   20
         Tag             =   "ciudad"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "direccion"
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   12
         Tag             =   "comuna"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3960
         MaxLength       =   1
         TabIndex        =   10
         Tag             =   "sucursal"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "rutcliente"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Cupo"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Credito"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label label 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   27
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Giro"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fono2"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fono1"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label label 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   8
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label label 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label label 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comuna"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   5535
         Left            =   0
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rut"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   6600
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "\\servidor\e\gestion comercial\barra_opciones.swf"
      Src             =   "\\servidor\e\gestion comercial\barra_opciones.swf"
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
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5535
      Left            =   480
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "ventas01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaclientes(dato2)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato1_LostFocus()
    If sl = 0 Then leer
sl = 0
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato8, dato10, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato9, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato10, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato13, dato15, KeyCode)
End Sub

Private Sub Form_Activate()

dato1.SetFocus
End Sub

Private Sub Form_Load()
    Call Conectar_BD
    sc = 0
    opciones.Visible = False
End Sub

Private Sub DATO1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato3, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato4, dato5)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato5, dato6)
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato6, dato7)
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato7, dato8)
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato8, dato9)
End Sub

Private Sub dato9_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato9, dato10)
End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato10, dato11)
End Sub

Private Sub dato11_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato11, dato12)
End Sub

Private Sub dato12_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato12, dato13)
End Sub

Private Sub dato13_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato13, dato14)
End Sub

Private Sub dato14_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato14, dato15)
End Sub

Private Sub dato15_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then GRABAR: leer:
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'SUCURSAL
    campos(2, 0) = dato3.Tag 'NOMBRE
    campos(3, 0) = dato4.Tag 'DIRECCION
    campos(4, 0) = dato5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = dato10.Tag '10
    campos(10, 0) = dato11.Tag '11
    campos(11, 0) = dato12.Tag '12
    campos(12, 0) = dato13.Tag '13
    campos(13, 0) = dato14.Tag '14
    campos(14, 0) = dato15.Tag '15
    
    
    campos(0, 2) = "maestroclientes"
    condicion = "rutcliente=" + "'" + dato1.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'NOMBRE
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'COMUNA
    campos(4, 0) = dato5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = dato10.Tag '10
    campos(10, 0) = dato11.Tag '11
    campos(11, 0) = dato12.Tag '12
    campos(12, 0) = dato13.Tag '13
    campos(13, 0) = dato14.Tag '14
    campos(14, 0) = dato15.Tag '15

    campos(0, 2) = "maestroclientes"
    condicion = "rutcliente>" + "'" + dato1.text + "' order by rutcliente"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'NOMBRE
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'COMUNA
    campos(4, 0) = dato5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = dato10.Tag '10
    campos(10, 0) = dato11.Tag '11
    campos(11, 0) = dato12.Tag '12
    campos(12, 0) = dato13.Tag '13
    campos(13, 0) = dato14.Tag '14
    campos(14, 0) = dato15.Tag '15

    campos(0, 2) = "maestroclientes"
    condicion = "rutcliente<" + "'" + dato1.text + "' order by rutcliente"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
 
End Sub

Sub carga()
    habilita (True)
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = SQLUTIL.datos(1, 3)
    dato3.text = SQLUTIL.datos(2, 3)
    dato4.text = SQLUTIL.datos(3, 3)
    dato5.text = SQLUTIL.datos(4, 3)
    dato6.text = SQLUTIL.datos(5, 3)
    dato7.text = SQLUTIL.datos(6, 3)
    dato8.text = SQLUTIL.datos(7, 3)
    dato9.text = SQLUTIL.datos(8, 3)
    dato10.text = SQLUTIL.datos(9, 3)
    dato11.text = SQLUTIL.datos(10, 3)
    dato12.text = SQLUTIL.datos(11, 3)
    dato13.text = SQLUTIL.datos(12, 3)
    dato14.text = SQLUTIL.datos(13, 3)
    dato15.text = SQLUTIL.datos(14, 3)
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    dato5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
    dato9.Locked = condicion
    dato10.Locked = condicion
    dato11.Locked = condicion
    dato12.Locked = condicion
    dato13.Locked = condicion
    dato14.Locked = condicion
    dato15.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    dato5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    dato10.Enabled = condicion
    dato11.Enabled = condicion
    dato12.Enabled = condicion
    dato13.Enabled = condicion
    dato14.Enabled = condicion
    dato15.Enabled = condicion
        
End Sub
'Sub Conecta_Maestro_Secciones()
'    'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
'    With maestro02
'        .mpro.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=conta01"
'    End With
'End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub ayudaclientes(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("rutcliente", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestroclientes", dato1, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub GRABAR()

    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'SUCURSAL
    campos(2, 0) = dato3.Tag 'NOMBRE
    campos(3, 0) = dato4.Tag 'DIRECCION
    campos(4, 0) = dato5.Tag '5
    campos(5, 0) = dato6.Tag '6
    campos(6, 0) = dato7.Tag '7
    campos(7, 0) = dato8.Tag '8
    campos(8, 0) = dato9.Tag '9
    campos(9, 0) = dato10.Tag '10
    campos(10, 0) = dato11.Tag '11
    campos(11, 0) = dato12.Tag '12
    campos(12, 0) = dato13.Tag '13
    campos(13, 0) = dato14.Tag '14
    campos(14, 0) = dato15.Tag '15
        
    campos(0, 1) = dato1.text 'RUT
    campos(1, 1) = dato2.text 'SUCURSAL
    campos(2, 1) = dato3.text 'NOMBRE
    campos(3, 1) = dato4.text 'DIRECCION
    campos(4, 1) = dato5.text '5
    campos(5, 1) = dato6.text '6
    campos(6, 1) = dato7.text '7
    campos(7, 1) = dato8.text '8
    campos(8, 1) = dato9.text '9
    campos(9, 1) = dato10.text '10
    campos(10, 1) = dato11.text '11
    campos(11, 1) = dato12.text '12
    campos(12, 1) = dato13.text '13
    campos(13, 1) = dato14.text '14
    campos(14, 1) = dato15.text '15
    
    campos(0, 2) = "maestroclientes"
    If modifi = 1 Then condicion = "rutcliente=" + "'" + dato1.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    modifi = 0
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.ESTADO




End Sub
Sub ELIMINAR()
    
    campos(0, 2) = "maestroclientes"
    condicion = "rutcliente=" + "'" + dato1.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
End Sub


Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then disponible (True): habilita (False): limpia: opciones.Visible = False: dato1.SetFocus
If command = "modifica" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.SetFocus: modifi = 1
If command = "elimina" Then disponible (True): habilita (False): ELIMINAR: limpia: opciones.Visible = False: dato1.SetFocus
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
End Sub

Sub limpia()


    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
    dato15.text = ""
    
End Sub
