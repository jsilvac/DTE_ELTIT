VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form TRASPASA 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Cuentas del Mayor"
   ClientHeight    =   6675
   ClientLeft      =   2235
   ClientTop       =   1425
   ClientWidth     =   8805
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc mcm 
      Height          =   375
      Left            =   2400
      Top             =   6840
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
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "PROCESA TRASPASO"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   1710
         Width           =   3375
      End
      Begin VB.FileListBox File1 
         Height          =   1065
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox ARCHIVO 
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   1320
         Width           =   3375
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   270
         TabIndex        =   4
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   7215
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         Height          =   6135
         Left            =   0
         Top             =   0
         Width           =   7935
      End
      Begin VB.Label NOMBRETIPO2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label NOMBRETIPO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   1560
         Width           =   1695
      End
   End
End
Attribute VB_Name = "TRASPASA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Right(ARCHIVO.text, 3) <> "txt" Then GoTo no:
TRASPASADATOS
no:


End Sub

Private Sub Dir1_Change()
Dir1.path = Drive1.Drive
File1.path = Dir1.path
File1.Pattern = "*.TXT"

End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
File1.path = Dir1.path
File1.Pattern = "*.TXT"




End Sub

Private Sub File1_DblClick()
K = File1.ListIndex

ARCHIVO.text = File1.List(K)



End Sub

Private Sub Form_Load()
    Call Conectar_BD
End Sub

Sub TRASPASADATOS()
añotraspaso = Mid(ARCHIVO.text, 3, 4)

Open File1.path + "\" + ARCHIVO.text For Input As #20

20 If EOF(20) Then
   Close 20
   Unload Me
   GoTo no:
   End If
   
Line Input #20, VARIPASO
Label1.Caption = VARIPASO
Label1.Refresh

GRABAMCMDATOS2

'For K = 1 To Len(VARIPASO)
'
'    If Mid(VARIPASO, K, 1) = "'" Then Mid(VARIPASO, K, 1) = " "
'    Next K
'
'If Mid(VARIPASO, 1, 8) = "MCMDATOS" Then GRABAMCMDATOS
'If Mid(VARIPASO, 1, 9) = "MCMSALDOS" Then GRABAMCMSALDOS
'If Mid(VARIPASO, 1, 8) = "MCTDATOS" Then GRABAMCTDATOS
'If Mid(VARIPASO, 1, 9) = "MCTSALDOS" Then GRABAMCTSALDOS
'If Mid(VARIPASO, 1, 8) = "MCCDATOS" Then GRABAMCCDATOS
'Rem If Mid(VARIPASO, 1, 9) = "MCCSALDOS" Then GRABAMCCSALDOS
'If Mid(VARIPASO, 1, 9) = "MACODATOS" Then GRABAMACODATOS
'If Mid(VARIPASO, 1, 8) = "MACOCTAS" Then GRABAMACOCUENTAS
'If Mid(VARIPASO, 1, 9) = "MOVIDATOS" Then GRABAMOVIMIENTOS
''If Mid(VARIPASO, 1, 9) = "MAVEDATOS" Then GRABAMAVEDATOS
''If Mid(VARIPASO, 1, 8) = "MAVECTAS" Then GRABAMAVECUENTAS
''If Mid(VARIPASO, 1, 9) = "MAFCDATOS" Then GRABAMAFCDATOS
''If Mid(VARIPASO, 1, 8) = "MAFCCTAS" Then GRABAMAFCCUENTAS
'If Mid(VARIPASO, 1, 9) = "MAHODATOS" Then GRABAMAHODATOS
'If Mid(VARIPASO, 1, 8) = "HONOCTAS" Then GRABAMAHOCUENTAS
'If Mid(VARIPASO, 1, 11) = "CHEQUEDATOS" Then GRABACHEQUES
''If Mid(VARIPASO, 1, 9) = "ZETADATOS" Then GRABAZETAS
'

GoTo 20:
no:

End Sub


Sub GRABAMCMDATOS()
    
    
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    grilla(num + 1) = añotraspaso
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "ctacte"
    campos(4, 0) = "crcc"
    campos(5, 0) = "banco"
    campos(6, 0) = "ila"
    campos(7, 0) = "ica"
    campos(8, 0) = "iha"
    campos(9, 0) = "activo"
    campos(10, 0) = "año"
    campos(11, 0) = ""
    
    For K = 0 To 11
    campos(K, 1) = grilla(K + 2)
    Next K
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + grilla(2) + "' and año='" + año + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then GoTo no:
       

    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
no:
End Sub

Sub GRABAMCMDATOS2()
    Dim tipocue As String
    
    
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    grilla(1) = Mid(VARIPASO, 1, 2) + Mid(VARIPASO, 4, 2) + Mid(VARIPASO, 7, 4)
    grilla(2) = Mid(VARIPASO, 11, 50)
    grilla(2) = Replace(grilla(2), Chr(9), "")
    If Mid(VARIPASO, 1, 1) >= "0" And Mid(VARIPASO, 1, 1) <= "9" Then
    grilla(3) = Mid(grilla(1), 1, 1)
    grilla(4) = "2008"
    año = "2008"
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "año"
    campos(4, 0) = ""
    
    
    campos(0, 1) = grilla(1)
    campos(1, 1) = grilla(2)
    campos(2, 1) = grilla(3)
    campos(3, 1) = grilla(4)
    
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + grilla(1) + "' and año='" + año + "'"

    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then GoTo no:


    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
no:
    End If
End Sub


Sub GRABAMCMSALDOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
      
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
    campos(0, 1) = grilla(2)
    campos(1, 1) = Mid(ARCHIVO.text, 3, 4)
    For K = 2 To 27
    campos(K, 1) = grilla(K + 1)
  
    Next K
    
    campos(0, 2) = "saldosdelmayor"
    condicion = "codigo=" + "'" + grilla(2) + "' and año ='" + grilla(3) + "'"
'    op = 5
 '   SQLUTIL.datos = campos
 '   Set SQLUTIL.conexion = db
  '  Call SQLUTIL.SQLUTIL(op, condicion)
   ' If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
no:
End Sub
Sub GRABAMCTDATOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = "'" Then Mid(VARIPASO, K, 1) = " "
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    grilla(num + 1) = añotraspaso
      
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "nombre"
    campos(3, 0) = "direccion"
    campos(4, 0) = "comuna"
    campos(5, 0) = "ciudad"
    campos(6, 0) = "giro"
    campos(7, 0) = "fono"
    campos(8, 0) = "fax"
    campos(9, 0) = "celular"
    campos(10, 0) = "email"
    campos(11, 0) = "contacto"
    campos(12, 0) = "dest_cheque"
    campos(13, 0) = "año"
    campos(14, 0) = ""
    
    grilla(3) = "0" + grilla(3)
    For K = 0 To 13
    campos(K, 1) = grilla(K + 2)
  
    Next K
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + grilla(2) + "' and rut ='" + grilla(3) + "'"
    'op = 5
    'SQLUTIL.datos = campos
    'Set SQLUTIL.conexion = db
    'Call SQLUTIL.SQLUTIL(op, condicion)
    'If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
no:
End Sub

Sub GRABAMCTSALDOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
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
    campos(26, 0) = "HABER10"
    campos(27, 0) = "HABER11"
    campos(28, 0) = "HABER12"
    campos(29, 0) = ""

    campos(0, 1) = grilla(2)
    grilla(3) = "0" + grilla(3)
    campos(1, 1) = grilla(3)
    campos(2, 1) = Mid(ARCHIVO.text, 3, 4)
    For K = 3 To 28
    campos(K, 1) = grilla(K + 1)
    Next K
    
    campos(0, 2) = "saldosctacte"
    condicion = "tipo=" + "'" + grilla(2) + "' and rut ='" + grilla(3) + "'and año ='" + grilla(4) + "'"
    'op = 5
    'SQLUTIL.datos = campos
    'Set SQLUTIL.conexion = db
    'Call SQLUTIL.SQLUTIL(op, condicion)
    'If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
no:
End Sub

Sub GRABAMCCDATOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "año"
    campos(3, 0) = ""
    grilla(4) = año
    For K = 0 To 2
    campos(K, 1) = grilla(K + 2)
    Next K
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + grilla(2) + "'"
    
    'op = 5
    'SQLUTIL.datos = campos
    'Set SQLUTIL.conexion = db
    'Call SQLUTIL.SQLUTIL(op, condicion)
    'If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
    Call grabarcuentas(grilla(2))
no:
End Sub

'Sub GRABAMCCSALDOS()
'    Dim grilla(50) As String
'    For K = 1 To 50: grilla(K) = "": Next K
'    num = 0: pasa = 0: ini = 1
'    For K = 1 To Len(VARIPASO)
'    pasa = pasa + 1
'    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, pasa - 1): ini = ini + pasa: pasa = 0
'    Next K
'
'    campos(0, 0) = "codigo"
'    campos(1, 0) = "año"
'    campos(2, 0) = "debeanterior"
'    campos(3, 0) = "haberanterior"
'    campos(4, 0) = "debe01"
'    campos(5, 0) = "debe02"
'    campos(6, 0) = "debe03"
'    campos(7, 0) = "debe04"
'    campos(8, 0) = "debe05"
'    campos(9, 0) = "debe06"
'    campos(10, 0) = "debe07"
'    campos(11, 0) = "debe08"
'    campos(12, 0) = "debe09"
'    campos(13, 0) = "debe10"
'    campos(14, 0) = "debe11"
'    campos(15, 0) = "debe12"
'    campos(16, 0) = "haber01"
'    campos(17, 0) = "haber02"
'    campos(18, 0) = "haber03"
'    campos(19, 0) = "haber04"
'    campos(20, 0) = "haber05"
'    campos(21, 0) = "haber06"
'    campos(22, 0) = "haber07"
'    campos(23, 0) = "haber08"
'    campos(24, 0) = "haber09"
'    campos(25, 0) = "HABER10"
'    campos(26, 0) = "HABER11"
'    campos(27, 0) = "HABER12"
'    campos(28, 0) = ""
'    campos(0, 1) = grilla(2)
'    campos(1, 1) = Mid(ARCHIVO.text, 3, 4)
'    For K = 2 To 27
'    campos(K, 1) = grilla(K + 1)
'
'    Next K
'
'    campos(0, 2) = "saldoscentrosdecosto"
'    condicion = "codigo=" + "'" + grilla(2) + "' and año ='" + grilla(3) + "'"
'    'op = 5
'    'SQLUTIL.datos = campos
'    'Set SQLUTIL.conexion = db
'    'Call SQLUTIL.SQLUTIL(op, condicion)
'    'If SQLUTIL.estado = 0 Then GoTo no:
'    op = 2
'    SQLUTIL.datos = campos
'    Set SQLUTIL.conexion = db
'    Call SQLUTIL.SQLUTIL(op, condicion)
'
'no:
'End Sub
Sub GRABAMOVIMIENTOS()
    Dim RUTPROVEEDOR As String
    
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    
    If Mid(VARIPASO, K, 1) = "'" Then Mid(VARIPASO, K, 1) = " "
    Next K
    For K = 1 To Len(VARIPASO)
    
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
      
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "glosacontable"
    campos(6, 0) = "tipoctacte"
    campos(7, 0) = "rutctacte"
    campos(8, 0) = "tipodocumento"
    campos(9, 0) = "fechavencimiento"
    campos(10, 0) = "monto"
    campos(11, 0) = "dh"
    campos(12, 0) = "numerodocumento"
    campos(13, 0) = "centrocosto"
    campos(14, 0) = "creadopor"
    campos(15, 0) = "fechadocumento"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "rutproveedor"
    
    campos(19, 0) = ""
    If grilla(2) = "0" Then grilla(2) = "CI"
    If grilla(2) = "1" Then grilla(2) = "CI"
    If grilla(2) = "2" Then grilla(2) = "CE"
    If grilla(2) = "3" Then grilla(2) = "CT"
    If grilla(2) = "4" Then grilla(2) = "CA"
    If grilla(2) = "5" Then grilla(2) = "FT"
    If grilla(2) = "6" Then grilla(2) = "BH"
    If grilla(2) = "7" Then grilla(2) = "FC"
    If grilla(2) = "8" Then grilla(2) = "DC"
    If grilla(2) = "9" Then grilla(2) = "NC"
    grilla(3) = "00" + grilla(3)
    grilla(14) = "00" + grilla(14)
    grilla(4) = "0" + grilla(4)
    grilla(17) = grilla(5)
    grilla(18) = Mid(grilla(5), 5, 2)
    grilla(19) = Mid(grilla(5), 1, 4)

    If grilla(2) = "FC" And grilla(4) = "001" Then
    
    campos(18, 1) = grilla(9)
    
    End If
    
    For K = 0 To 17
    campos(K, 1) = grilla(K + 2)

    Next K
    
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo=" + "'" + grilla(2) + "' and numero ='" + grilla(3) + "' and linea='" + grilla(4) + "'"
 '  op = 5
  ' SQLUTIL.datos = campos
  ' Set SQLUTIL.conexion = db
  ' Call SQLUTIL.SQLUTIL(op, condicion)
  ' If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
no:
End Sub

Sub GRABAMACODATOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "rut"
    campos(4, 0) = "neto"
    campos(5, 0) = "iva"
    campos(6, 0) = "exento"
    campos(7, 0) = "total"
    campos(8, 0) = "mescontable"
    campos(9, 0) = "añocontable"
    campos(10, 0) = "fechavencimiento"
    campos(11, 0) = "retencion"
    campos(12, 0) = "abono"
    campos(13, 0) = "folio"
    campos(14, 0) = "comentario"
    campos(15, 0) = "electronica"
    campos(16, 0) = "fechadigitacion"
    campos(17, 0) = ""
'    If grilla(2) = 4 Then grilla(2) = "1": grilla(12) = "E"
'    If grilla(2) = 5 Then grilla(2) = "2": grilla(12) = "E"
'    If grilla(2) = 6 Then grilla(2) = "3": grilla(12) = "E"
    
    campos(0, 1) = grilla(2)
    campos(1, 1) = "00" + grilla(3)
    campos(2, 1) = grilla(4)
    campos(3, 1) = "0" + grilla(5)
    campos(4, 1) = grilla(6)
    campos(5, 1) = grilla(7)
    campos(6, 1) = grilla(8)
    campos(7, 1) = grilla(9)
    campos(8, 1) = grilla(10)
    campos(9, 1) = Mid(ARCHIVO.text, 3, 4)
    campos(10, 1) = grilla(15)
    campos(11, 1) = "0"
    campos(12, 1) = grilla(17)
    campos(13, 1) = grilla(14)
    campos(14, 1) = ""
    campos(15, 1) = grilla(12)
    campos(16, 1) = grilla(16)
    If grilla(16) > "20061231" Then
    
    campos(0, 2) = "facturasdecompras"
  ' condicion = "tipo=" + "'" + grilla(2) + "'and numero='" + grilla(3) + "' and rut='" + "0" + grilla(5) + "'"

   'op = 5
   'SQLUTIL.datos = campos
   'Set SQLUTIL.conexion = db
   'Call SQLUTIL.SQLUTIL(op, condicion)
   'If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
End If

no:
End Sub

Sub GRABAMACOCUENTAS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
      
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "rut"
    campos(3, 0) = "linea"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "centrodecosto"
    campos(8, 0) = "dh"
    campos(9, 0) = "rutctacte"
    
    campos(10, 0) = ""
    If grilla(2) = 4 Then grilla(2) = "1"
    If grilla(2) = 5 Then grilla(2) = "2"
    If grilla(2) = 6 Then grilla(2) = "3"
    
    campos(0, 1) = grilla(2)
    campos(1, 1) = "00" + grilla(3)
    campos(2, 1) = "0" + grilla(4)
    campos(3, 1) = grilla(5)
    campos(4, 1) = grilla(6)
    campos(5, 1) = grilla(7)
    campos(6, 1) = grilla(8)
    campos(7, 1) = grilla(9)
    campos(8, 1) = grilla(10)
    campos(9, 1) = ""
    campos(10, 1) = ""
    
    
    
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = "tipo=" + "'" + grilla(2) + "' and numero ='" + grilla(3) + "'and linea ='" + grilla(5) + "' and rut='" + "0" + grilla(4) + "'"
  ' Op = 5
  ' SQLUTIL.datos = campos
  ' Set SQLUTIL.conexion = db
  ' Call SQLUTIL.SQLUTIL(op, condicion)
  ' If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
   
no:
End Sub
Sub GRABAMAVEDATOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "rut"
    campos(4, 0) = "fechavencimiento"
    campos(5, 0) = "añocontable"
    campos(6, 0) = "mescontable"
    campos(7, 0) = "neto"
    campos(8, 0) = "iva"
    campos(9, 0) = "exento"
    campos(10, 0) = "impuestoespecifico"
    campos(11, 0) = "total"
    campos(12, 0) = ""
    For K = 0 To 11
    campos(K, 1) = grilla(K + 2)
    Next K
    campos(0, 2) = "facturasdeventas"
    condicion = "tipo=" + "'" + grilla(2) + "'and numero='" + grilla(3) + "'"
'   op = 5
'   SQLUTIL.datos = campos
'   Set SQLUTIL.conexion = db
'   Call SQLUTIL.SQLUTIL(op, condicion)
'   If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
no:
End Sub

Sub GRABAMAVECUENTAS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(2, 0) = "rut"
    campos(2, 0) = "linea"
    campos(3, 0) = "cuentadelmayor"
    campos(4, 0) = "glosa"
    campos(5, 0) = "monto"
    campos(6, 0) = "dh"
    campos(7, 0) = "centrodecostos"
    campos(8, 0) = "tipoctacte"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = ""
    campos(0, 1) = grilla(2)
    campos(1, 1) = Mid(ARCHIVO.text, 3, 4)
    For K = 2 To 9
    campos(K, 1) = grilla(K + 1)
    Next K
    campos(0, 2) = "detallefacturadeventa"
    condicion = "tipo=" + "'" + grilla(2) + "' and numero ='" + grilla(3) + "'and linea ='" + grilla(4) + "'"
'   op = 5
'   SQLUTIL.datos = campos
'   Set SQLUTIL.conexion = db
'   Call SQLUTIL.SQLUTIL(op, condicion)
'   If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
no:
End Sub
Sub GRABAMAFCDATOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "rut"
    campos(4, 0) = "fechavencimiento"
    campos(5, 0) = "añocontable"
    campos(6, 0) = "mescontable"
    campos(7, 0) = "neto"
    campos(8, 0) = "iva"
    campos(9, 0) = "exento"
    campos(10, 0) = "retenido"
    campos(11, 0) = "total"
    campos(12, 0) = ""
    For K = 0 To 11
    campos(K, 1) = grilla(K + 2)
    Next K
    campos(0, 2) = "facturasdecompraspropias"
    condicion = "tipo=" + "'" + grilla(2) + "'and numero='" + grilla(3) + "'"
'   op = 5
'   SQLUTIL.datos = campos
'   Set SQLUTIL.conexion = db
'   Call SQLUTIL.SQLUTIL(op, condicion)
'   If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
no:
End Sub

Sub GRABAMAFCCUENTAS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "cuentasdelmayor"
    campos(4, 0) = "glosa"
    campos(5, 0) = "monto"
    campos(6, 0) = "dh"
    campos(7, 0) = "centrodecosto"
    campos(8, 0) = "tipoctacte"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = ""
    campos(0, 1) = grilla(2)
    campos(1, 1) = Mid(ARCHIVO.text, 3, 4)
    For K = 2 To 9
    campos(K, 1) = grilla(K + 1)
    Next K
    campos(0, 2) = "detallefacturaspropias"
    condicion = "tipo=" + "'" + grilla(2) + "' and numero ='" + grilla(3) + "'and linea ='" + grilla(4) + "'"
'   op = 5
'   SQLUTIL.datos = campos
'   Set SQLUTIL.conexion = db
'   Call SQLUTIL.SQLUTIL(op, condicion)
'   If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
no:
End Sub
Sub GRABAMAHODATOS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "rut"
    campos(4, 0) = "mescontable"
    campos(5, 0) = "añocontable"
    campos(6, 0) = "monto"
    campos(7, 0) = "retencion"
    campos(8, 0) = "liquido"
    campos(9, 0) = "fechavencimiento"
    campos(10, 0) = ""
    
    campos(0, 1) = grilla(2)
    campos(1, 1) = "00" + grilla(3)
    campos(2, 1) = grilla(4)
    campos(3, 1) = "0" + grilla(5)
    campos(4, 1) = Mid(grilla(4), 5, 2)
    campos(5, 1) = Mid(grilla(4), 1, 4)
    campos(6, 1) = grilla(6)
    campos(7, 1) = grilla(7)
    campos(8, 1) = grilla(9)
    campos(9, 1) = grilla(4)
 
    
    

    campos(0, 2) = "boletasdehonorarios"
    condicion = "tipo=" + "'" + grilla(2) + "' and numero ='" + grilla(3) + "' and rut='" + "0" + grilla(4) + "'"
       
'   op = 5
'   SQLUTIL.datos = campos
'   Set SQLUTIL.conexion = db
'   Call SQLUTIL.SQLUTIL(op, condicion)
'   If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
no:
End Sub

Sub GRABAMAHOCUENTAS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
      
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "rut"
    campos(3, 0) = "linea"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "centrodecosto"
    campos(8, 0) = "dh"
    campos(9, 0) = "rutctacte"
    
    campos(10, 0) = ""
 
    
    campos(0, 1) = grilla(2)
    campos(1, 1) = "00" + grilla(3)
    campos(2, 1) = "0" + grilla(4)
    campos(3, 1) = grilla(5)
    campos(4, 1) = grilla(6)
    campos(5, 1) = grilla(7)
    campos(6, 1) = grilla(8)
    campos(7, 1) = grilla(9)
    campos(8, 1) = grilla(10)
    campos(9, 1) = ""
    campos(10, 1) = ""
    
    
    
    campos(0, 2) = "boletasdehonorarios_detalle"
    condicion = "tipo=" + "'" + grilla(2) + "' and numero ='" + grilla(3) + "'and linea ='" + grilla(5) + "' and rut='" + "0" + grilla(4) + "'"
  ' Op = 5
  ' SQLUTIL.datos = campos
  ' Set SQLUTIL.conexion = db
  ' Call SQLUTIL.SQLUTIL(op, condicion)
  ' If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
   
no:
End Sub


Sub GRABACHEQUES()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
      
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "numerocomprobante"
    campos(6, 0) = "fechacobro"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "ubicacion"
    campos(9, 0) = "cobrado"
    campos(10, 0) = "tipocomprobante"
    campos(11, 0) = ""
    If Val(grilla(8)) <> 0 Then grilla(11) = "1"
    grilla(3) = "00" + grilla(3)
    grilla(7) = "0000" + grilla(7)
    grilla(8) = Mid(grilla(8), 5, 4) + "-" + Mid(grilla(8), 3, 2) + "-" + Mid(grilla(8), 1, 2)
    
    For K = 2 To 11
    campos(K - 2, 1) = grilla(K)
    Next K
    campos(10, 1) = "CE"
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta=" + "'" + grilla(2) + "' and numero ='" + grilla(3) + "'"
'   op = 5
'   SQLUTIL.datos = campos
'   Set SQLUTIL.conexion = db
'   Call SQLUTIL.SQLUTIL(op, condicion)
'   If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
no:
End Sub

Sub GRABAZETAS()
    Dim grilla(50) As String
    For K = 1 To 50: grilla(K) = "": Next K
    num = 0: PASA = 0: ini = 1
    For K = 1 To Len(VARIPASO)
    PASA = PASA + 1
    If Mid(VARIPASO, K, 1) = Chr(123) Then num = num + 1: grilla(num) = Mid(VARIPASO, ini, PASA - 1): ini = ini + PASA: PASA = 0
    Next K
      
    campos(0, 0) = "fecha"
    campos(1, 0) = "caja"
    campos(2, 0) = "numero"
    campos(3, 0) = "zetas"
    campos(4, 0) = "desde"
    campos(5, 0) = "haber"
    campos(6, 0) = "monto"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "exento"
    campos(9, 0) = ""
    campos(0, 1) = grilla(2)
    campos(1, 1) = Mid(ARCHIVO.text, 3, 4)
    For K = 2 To 8
    campos(K, 1) = grilla(K + 1)
    Next K
    campos(0, 2) = "zetas"
'   condicion = "fecha=" + "'" + grilla(2) + "' and caja ='" + grilla(3) + "'"
'   op = 5
'   SQLUTIL.datos = campos
'   Set SQLUTIL.conexion = db
'   Call SQLUTIL.SQLUTIL(op, condicion)
'   If SQLUTIL.estado = 0 Then GoTo no:
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
no:
End Sub



Sub grabarcuentas(CRCC As String)

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = db
        cSql2.SQL = "SELECT codigo,nombre "
        cSql2.SQL = cSql2.SQL + "FROM cuentasdelmayor where crcc='1' and año='" + año + "' "
       
'        cSql2.SQL = cSql2.SQL + "order by codigo"
        cSql2.Execute
        LINEAS = 0
        
        If cSql2.RowsAffected > 0 Then
         
        
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        Call grabar2(CRCC, resultados2(0), Mid(fechasistema, 7, 4))
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    
    
    

End Sub

Sub grabar2(CRCC, cuenta, año)
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "cuenta"
    campos(3, 0) = ""
    campos(0, 1) = CRCC
    campos(1, 1) = año
    campos(2, 1) = cuenta
    campos(0, 2) = "saldoscentrosdecosto"
    op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    

End Sub

