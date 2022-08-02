VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form ingreso101 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Pantalla de  Ventas"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   1440
   ClientWidth     =   15165
   FillColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   15165
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   8880
      TabIndex        =   24
      Top             =   7800
      Width           =   6015
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEBE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   120
         Width           =   1575
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
         Index           =   5
         Left            =   2640
         TabIndex        =   32
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
         Index           =   6
         Left            =   2040
         TabIndex        =   31
         Top             =   960
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
         Index           =   7
         Left            =   2040
         TabIndex        =   30
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HABER"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label debe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFDF2&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label saldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFDF2&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label haber 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFDF2&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   6135
      Left            =   0
      TabIndex        =   22
      Top             =   1560
      Width           =   15135
      Begin FlexCell.Grid Grid1 
         Height          =   5775
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   14865
         _ExtentX        =   26220
         _ExtentY        =   10186
         Appearance      =   0
         BackColor1      =   16776690
         BackColor2      =   16117210
         BackColorBkg    =   16107953
         BackColorFixed  =   16107953
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         Cols            =   5
         DefaultFontSize =   8.25
         DefaultFontBold =   -1  'True
         DisablePrintButton=   -1  'True
         DisplayRowIndex =   -1  'True
         GridColor       =   16761024
         Rows            =   30
         DateFormat      =   2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   15135
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre de Centro de Costo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10320
         TabIndex        =   21
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre Cuenta del Mayor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre Cuenta Corriente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5200
         TabIndex        =   19
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label nombrecrcc 
         BackColor       =   &H00FFFDF2&
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
         Left            =   10320
         TabIndex        =   18
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label nombrectacte 
         BackColor       =   &H00FFFDF2&
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
         Left            =   5200
         TabIndex        =   17
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label nombremayor 
         BackColor       =   &H00FFFDF2&
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
         TabIndex        =   16
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.TextBox pivote4 
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox pivote3 
      Height          =   285
      Left            =   0
      MaxLength       =   12
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox pivote2 
      Height          =   375
      Left            =   0
      MaxLength       =   8
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox lineas 
      Height          =   285
      Left            =   0
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame cabeza 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C9B1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   14895
      Begin VB.TextBox dato0 
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
         Left            =   720
         MaxLength       =   2
         TabIndex        =   1
         Top             =   120
         Width           =   420
      End
      Begin VB.TextBox DATO2 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox DATO3 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9840
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   120
         Width           =   375
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox DATO1 
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
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   2
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label tipocompro 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   14895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA :"
         Height          =   285
         Left            =   8640
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FOLIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox PIVOTE 
      Height          =   375
      Left            =   6480
      MaxLength       =   13
      TabIndex        =   0
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   8400
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
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
   End
End
Attribute VB_Name = "ingreso101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private modifi As Integer
    Private formatogrilla(10, 20) As String
    Private sg As String
    Private tipoctacte As String
    Private SUMADEBE As Double
    Private SUMAHABER As Double
    Private CANLI As Integer
    Private grilladetalle(1000, 13) As String
    Private o As Integer
    Private ef As String
    Private cc As String
    Private respu As String
    
    
    
    
    
    
    
    


Private Sub Command3_Click()
    cabeza.Enabled = True
    Grid1.Enabled = True
    
End Sub

Private Sub dato0_GotFocus()
Call cargatexto(dato0)

End Sub

Private Sub dato0_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudatipos(dato0)
End Sub

Private Sub DATO1_GotFocus()
ULTIMO
Call cargatexto(DATO1)
Call leetipos(dato0)

End Sub

Private Sub dato1_LostFocus()

leecomprobante
End Sub

Private Sub dato2_GotFocus()
    
    Call cargatexto(DATO2)
End Sub

Private Sub DATO2_LostFocus()
    
    If DATO2.text = "00" Then DATO3.Enabled = True: DATO4.Enabled = True: DATO2.text = Mid(Date, 1, 2): DATO3.text = Mid(Date, 4, 2): DATO4.text = Mid(Date, 7, 4): Grid1.Enabled = True: Grid1.SetFocus

End Sub

Private Sub dato3_GotFocus()
    Call cargatexto(DATO3)
End Sub
Private Sub dato4_GotFocus()
    
    Call cargatexto(DATO4)
End Sub

Private Sub dato4_LostFocus()
    Call ESFECHA(DATO2.text, DATO3.text, DATO4.text)
    If ef = "N" Then DATO2.SetFocus
    If ef = "S" Then Grid1.Enabled = True

End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato0, DATO2, KeyCode)
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO1, DATO3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO2, DATO4, KeyCode)
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO3, DATO4, KeyCode)
End Sub
Private Sub DATO0_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato0, DATO1)
End Sub

Private Sub DATO1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO1): Call Pregunta(DATO1, DATO2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO2): Call Pregunta(DATO2, DATO3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO3): Call Pregunta(DATO3, DATO4)
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Grid1.SetFocus
    
End Sub



Private Sub Form_Load()
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
    Dim SS As String
    Dim margen As Double
    Dim suma As Double
    Call CARGAGRILLA(2, 13)
    opciones.Visible = False
    Grid1.Enabled = False
    
     
End Sub

Sub CARGAGRILLA(row, col)
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "C1"
    formatogrilla(1, 2) = "C2"
    formatogrilla(1, 3) = "C3"
    formatogrilla(1, 4) = "TC"
    formatogrilla(1, 5) = "RUT"
    formatogrilla(1, 6) = "CRCC"
    formatogrilla(1, 7) = "GLOSA"
    formatogrilla(1, 8) = "TP"
    formatogrilla(1, 9) = "NUMERO"
    formatogrilla(1, 10) = "F.VENCI."
    formatogrilla(1, 11) = "MONTO"
    formatogrilla(1, 12) = "D/H"
    formatogrilla(1, 13) = ""
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "2"
    formatogrilla(2, 2) = "2"
    formatogrilla(2, 3) = "4"
    formatogrilla(2, 4) = "3"
    formatogrilla(2, 5) = "10"
    formatogrilla(2, 6) = "4"
    formatogrilla(2, 7) = "60"
    formatogrilla(2, 8) = "2"
    formatogrilla(2, 9) = "10"
    formatogrilla(2, 10) = "10"
    formatogrilla(2, 11) = "11"
    formatogrilla(2, 12) = "1"
    formatogrilla(2, 13) = "1"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "C"
    formatogrilla(3, 2) = "C"
    formatogrilla(3, 3) = "C"
    formatogrilla(3, 4) = "S"
    formatogrilla(3, 5) = "C"
    formatogrilla(3, 6) = "C"
    formatogrilla(3, 7) = "S"
    formatogrilla(3, 8) = "S"
    formatogrilla(3, 9) = "C"
    formatogrilla(3, 10) = "D"
    formatogrilla(3, 11) = "N"
    formatogrilla(3, 12) = "S"
    formatogrilla(3, 13) = "S"
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = ""
    formatogrilla(4, 4) = ""
    formatogrilla(4, 5) = ""
    formatogrilla(4, 6) = ""
    formatogrilla(4, 7) = ""
    formatogrilla(4, 8) = ""
    formatogrilla(4, 9) = ""
    formatogrilla(4, 10) = "dd-mm-yyyy"
    formatogrilla(4, 11) = " ###,###,##0"
    formatogrilla(4, 12) = ""
    formatogrilla(4, 13) = ""
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    formatogrilla(5, 3) = "FALSE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "FALSE"
    formatogrilla(5, 6) = "FALSE"
    formatogrilla(5, 7) = "FALSE"
    formatogrilla(5, 8) = "FALSE"
    formatogrilla(5, 9) = "FALSE"
    formatogrilla(5, 10) = "FALSE"
    formatogrilla(5, 11) = "FALSE"
    formatogrilla(5, 12) = "FALSE"
    formatogrilla(5, 13) = "TRUE"
    Rem VALOR MINIMO
    formatogrilla(6, 1) = ""
    formatogrilla(6, 2) = ""
    formatogrilla(6, 3) = ""
    formatogrilla(6, 4) = ""
    formatogrilla(6, 5) = ""
    formatogrilla(6, 6) = ""
    formatogrilla(6, 7) = ""
    formatogrilla(6, 8) = ""
    formatogrilla(6, 9) = ""
    formatogrilla(6, 10) = ""
    formatogrilla(6, 11) = ""
    formatogrilla(6, 12) = ""
    formatogrilla(6, 13) = ""
    Rem VALOR MAXIMO
    formatogrilla(7, 1) = ""
    formatogrilla(7, 2) = ""
    formatogrilla(7, 3) = ""
    formatogrilla(7, 4) = ""
    formatogrilla(7, 5) = ""
    formatogrilla(7, 6) = ""
    formatogrilla(7, 7) = ""
    formatogrilla(7, 8) = ""
    formatogrilla(7, 9) = ""
    formatogrilla(7, 10) = ""
    formatogrilla(7, 11) = "999999999"
    formatogrilla(7, 12) = ""
    formatogrilla(7, 13) = ""
    Grid1.Cols = col
    Grid1.Rows = row
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    'Grid1.Appearance = Flat
    'Grid1.ScrollBarStyle = Flat
    'Grid1.FixedRowColStyle = Flat
    'Grid1.BackColorFixed = RGB(90, 158, 214)
    'Grid1.BackColorFixedSel = RGB(110, 180, 230)
    'Grid1.BackColorBkg = RGB(90, 158, 214)
    'Grid1.BackColorScrollBar = RGB(231, 235, 247)
    'Grid1.BackColor1 = RGB(231, 235, 247)
    'Grid1.BackColor2 = RGB(239, 243, 255)
    'Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 4 * 8.8
    For K = 1 To col - 1
        Grid1.Cell(0, K).text = formatogrilla(1, K)
        If K < 5 Then Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 10.5
        If K > 4 Then Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 8.8
        If K = 8 Then Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 10.5
        If K = 7 Then Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 6.8
        If K = 12 Then Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 14
        If K = 13 Then Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 1
        Grid1.Column(K).MaxLength = Val(formatogrilla(2, K))
        Grid1.Column(K).FormatString = formatogrilla(4, K)
        Grid1.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then Grid1.Column(K).Alignment = cellRightCenter
        If formatogrilla(3, K) = "D" Then Grid1.Column(K).CellType = cellCalendar
        
        'Grid1.Column(7).CellType = cellComboBox
    Next K
        
    'Grid1.Column(1).CellType = cellTextBox
    'Grid1.Column(2).CellType = cellComboBox
    'Grid1.Column(2).MaxLength = 4 '// Important for ComboBox
    'Grid1.Column(3).CellType = cellCheckBox
    'Grid1.Column(4).CellType = cellCalendar
    'Grid1.Column(5).CellType = cellButton
    'Grid1.Column(6).CellType = cellHyperLink
    'Grid1.Column(4).FormatString = "mm-dd-yyyy"
    ' Grid1.Range(1, 6, Grid1.Rows - 1, 6).ForeColor = RGB(0, 0, 128)
Rem     Grid1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 35 And Grid1.ActiveCell.col = 1 Then graba
    Rem If KeyCode = 38 And Grid1.ActiveCell.row = Grid1.Rows - 1 Then SG = "S" Else SG = "N"
    If Grid1.ActiveCell.col = "1" And KeyCode = vbKeyF2 Then Call ayudamayor(Grid1.ActiveCell.row, Grid1.ActiveCell.col)
    If Grid1.ActiveCell.col = "5" And KeyCode = vbKeyF2 Then Call ayudactacte(Grid1.ActiveCell.row, Grid1.ActiveCell.col)
    If Grid1.ActiveCell.col = "6" And KeyCode = vbKeyF2 Then Call ayudacrcc(Grid1.ActiveCell.row, Grid1.ActiveCell.col)
   
    End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Grid1.ActiveCell.col = 12 And Chr(KeyAscii) <> "D" And Chr(KeyAscii) <> "H" Then KeyAscii = 0
    If Grid1.ActiveCell.col = 1 And Chr(KeyAscii) = "*" Then graba
    Rem If formatogrilla(3, Grid1.ActiveCell.col) = "S" Then Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = UCase(Grid1.ActiveCell.text)
    If formatogrilla(3, Grid1.ActiveCell.col) = "N" Then snum = 1: KeyAscii = esNumero(KeyAscii)
    If formatogrilla(3, Grid1.ActiveCell.col) = "C" Then snum = 1: KeyAscii = esNumero(KeyAscii)
End Sub

Sub graba()
    If modifi = 1 Then elimina: modifi = 0
    
    For K = 1 To Grid1.Rows - 1
        If Grid1.Cell(K, 1).text <> "" Then Call GRABAR
    Next K
    final
End Sub

Sub final()
    Grid1.AutoRedraw = False
    cabeza.Enabled = True
    Grid1.Enabled = True
    Grid1.Rows = 2
    DATO1.SetFocus
    DATO2.text = dia
    DATO3.text = mes
    DATO4.text = año
    
    For K = 1 To 12
    Grid1.Cell(1, K).text = ""
    Next K
    opciones.Visible = False
Grid1.AutoRedraw = True
Grid1.Refresh

End Sub

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, newrow As Long, newcol As Long, Cancel As Boolean)
    Dim TEXTO As String
    Dim dv As String
    If row = 0 And col = 0 Then newrow = 1: newcol = 1: GoTo no:
    If col = 1 And newcol > 11 Then newcol = 11
    If col = 12 And newcol > col Then newcol = 1
    TEXTO = Grid1.Cell(row, col).text
    For K = 1 To 12
    If col = K And row > 1 And Grid1.Cell(row, col).text = "" Then Grid1.Cell(row, col).text = Grid1.Cell(row - 1, col).text
    Next K
    If col = 12 And row > 1 And Grid1.Cell(row - 1, col).text = "D" Then Grid1.Cell(row, col).text = "H"
    If col = 12 And row > 1 And Grid1.Cell(row - 1, col).text = "H" Then Grid1.Cell(row, col).text = "D"
    If col = 1 And row = Grid1.Rows - 1 And newrow < row Then GoTo paso2:
    If newcol = 6 And newcol < col And Mid(Grid1.Cell(row, 4).text, 1, 1) <> "S" Then newcol = 5
    If newcol = 5 And newcol < col And Mid(Grid1.Cell(row, 4).text, 2, 2) = "00" Then newcol = 3
    If col = 1 Then PIVOTE.MaxLength = 2: PIVOTE.text = Grid1.Cell(row, col).text: Call ceros(PIVOTE): Grid1.Cell(row, col).text = PIVOTE.text
    If col = 2 Then PIVOTE.MaxLength = 2: PIVOTE.text = Grid1.Cell(row, col).text: Call ceros(PIVOTE): Grid1.Cell(row, col).text = PIVOTE.text
    If col = 3 Then PIVOTE.MaxLength = 4: PIVOTE.text = Grid1.Cell(row, col).text: Call ceros(PIVOTE): Grid1.Cell(row, col).text = PIVOTE.text
    If col = 3 And newcol = 4 Then Call leermayor(row, col)
10: If newcol = 4 And col < newcol And Mid(Grid1.Cell(row, 4).text, 2, 2) = "00" Then newcol = 6: GoTo 20:
    If col = 5 Then PIVOTE.MaxLength = 9: PIVOTE.text = Grid1.Cell(row, col).text: Call ceros(PIVOTE): Grid1.Cell(row, col).text = PIVOTE.text
    If col = 5 Then: dv = rut(Grid1.Cell(row, col).text): Grid1.Cell(row, col).text = PIVOTE.text + dv
    If col = 5 And newcol = 6 Then Call leerctacte(row, col)
20: If newcol = 6 And col < newcol And Mid(Grid1.Cell(row, 4).text, 1, 1) <> "S" Then newcol = 7: GoTo 30:
    If col = 6 Then PIVOTE.MaxLength = 4: PIVOTE.text = Grid1.Cell(row, col).text: Call ceros(PIVOTE): Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = PIVOTE.text
    If newcol = 7 And col = 6 Then Call leercrcc(row, col)
30: If col = 9 Then PIVOTE.MaxLength = 10: PIVOTE.text = Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text: Call ceros(PIVOTE): Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = PIVOTE.text
    If col = 10 And Grid1.Cell(row, col).text = "" Then Grid1.Cell(row, col).text = Mid(Date$, 4, 2) + "-" + Mid(Date$, 1, 2) + "-" + Mid(Date$, 7, 4)
paso2:
    If col > 11 Then SUMAR
    If row = Grid1.Rows - 1 And col = Grid1.Cols - 1 And newcol = 1 Then Grid1.Rows = Grid1.Rows + 1: newrow = Grid1.Rows - 1
    For K = 1 To newcol - 1
    
    If Grid1.Cell(row, K).text = "" And K <> 4 And K <> 5 And K <> 6 Then newcol = K: Exit For
    Next K
no:
End Sub





Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub leeultima()
End Sub

Sub GRABAR()
    Dim tipo2 As String
    lineas.text = K
    Call ceros(lineas)
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    
    campos(16, 0) = ""
    campos(0, 1) = dato0.text
    campos(1, 1) = DATO1.text
    campos(2, 1) = lineas.text
    campos(3, 1) = DATO4.text + DATO3.text + DATO2.text
    campos(4, 1) = Grid1.Cell(K, 1).text + Grid1.Cell(K, 2).text + Grid1.Cell(K, 3).text
    campos(5, 1) = Grid1.Cell(K, 4).text
    campos(6, 1) = Grid1.Cell(K, 5).text
    campos(7, 1) = Grid1.Cell(K, 6).text
    campos(8, 1) = Grid1.Cell(K, 7).text
    campos(9, 1) = Grid1.Cell(K, 8).text
    campos(10, 1) = Grid1.Cell(K, 9).text
    campos(11, 1) = campos(3, 1)
    campos(12, 1) = Grid1.Cell(K, 10).text
    campos(13, 1) = Replace(Grid1.Cell(K, 11).text, ",", ".")
    campos(14, 1) = Grid1.Cell(K, 12).text
    campos(15, 1) = USUARIO
    campos(0, 2) = "movimientoscontables"
    condicion = ""
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    tipo2 = Mid(Grid1.Cell(K, 4).text, 2, 2)
    If tipo2 <> "00" Then Call actualizactacte(K, tipo2, Grid1.Cell(K, 5).text)
    If Mid(Grid1.Cell(K, 4).text, 1, 1) = "S" Then Call actualizacrcc(K, Grid1.Cell(K, 6).text)
    If Mid(Grid1.Cell(K, 4).text, 1, 1) = "S" Then Call actualizacrcc(K, Mid(Grid1.Cell(K, 6).text, 1, 2) + "00")
    Call actualizamayor(K, Grid1.Cell(K, 1).text + Grid1.Cell(K, 2).text + Grid1.Cell(K, 3).text)
    Call actualizamayor(K, Grid1.Cell(K, 1).text + Grid1.Cell(K, 2).text + "0000")
    Call actualizamayor(K, Grid1.Cell(K, 1).text + "000000")
End Sub



Sub leecomprobante()
    Dim lin As Integer
    
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String

    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT tipo,numero,linea,fecha,codigocuenta,tipoctacte,rutctacte,centrocosto,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh "
        cSql.SQL = cSql.SQL + "FROM movimientoscontables "
        cSql.SQL = cSql.SQL + "WHERE tipo='" + dato0.text + "' and numero='" & DATO1.text & "' order by linea"
        cSql.Execute

        CANLI = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
             CANLI = CANLI + 1
                DATO2.text = Mid(resultados(3), 1, 2)
                DATO3.text = Mid(resultados(3), 4, 2)
                DATO4.text = Mid(resultados(3), 7, 4)
                
                grilladetalle(CANLI, 1) = Mid(resultados(4), 1, 2)
                grilladetalle(CANLI, 2) = Mid(resultados(4), 3, 2)
                grilladetalle(CANLI, 3) = Mid(resultados(4), 5, 4)
                grilladetalle(CANLI, 4) = resultados(5)
                grilladetalle(CANLI, 5) = resultados(6)
                grilladetalle(CANLI, 6) = resultados(7)
                grilladetalle(CANLI, 7) = resultados(8)
                grilladetalle(CANLI, 8) = resultados(9)
                grilladetalle(CANLI, 9) = resultados(10)
                grilladetalle(CANLI, 10) = resultados(11)
                grilladetalle(CANLI, 11) = resultados(12)
                grilladetalle(CANLI, 12) = resultados(13)

                resultados.MoveNext
            Wend
            cargadorcomprobante
            resultados.Close
            Set resultados = Nothing
        End If
    End With
   If cSql.RowsAffected > 0 Then opciones.Visible = True: Grid1.Enabled = False: cabeza.Enabled = False: opciones.SetFocus
no:
End Sub
Sub cargadorcomprobante()
    Grid1.AutoRedraw = False
    
    
    Grid1.Rows = CANLI + 1
    
    For K = 1 To CANLI
    Grid1.Cell(K, 1).text = grilladetalle(K, 1)
    Grid1.Cell(K, 2).text = grilladetalle(K, 2)
    Grid1.Cell(K, 3).text = grilladetalle(K, 3)
    Grid1.Cell(K, 4).text = grilladetalle(K, 4)
    Grid1.Cell(K, 5).text = grilladetalle(K, 5)
    Grid1.Cell(K, 6).text = grilladetalle(K, 6)
    Grid1.Cell(K, 7).text = grilladetalle(K, 7)
    Grid1.Cell(K, 8).text = grilladetalle(K, 8)
    Grid1.Cell(K, 9).text = grilladetalle(K, 9)
    Grid1.Cell(K, 10).text = grilladetalle(K, 10)
    Grid1.Cell(K, 11).text = grilladetalle(K, 11)
    Grid1.Cell(K, 12).text = grilladetalle(K, 12)

    SUMAR

    Next K
    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    
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
    If command = "retorno" Then final
    If command = "modifica" Then modifica
    If command = "elimina" Then elimina: final
    If command = "siguiente" Then SIGUIENTE
    If command = "anterior" Then anterior
    If command = "imprime" Then IMPRIMIR
End Sub

Sub modifica()
    modifi = 1
    Grid1.Enabled = True
    Grid1.SetFocus
End Sub

Sub IMPRIMIR()
    Grid1.PrintDialog
    Grid1.PrintPreview
End Sub

Sub elimina()
    Dim tipo2 As String
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + dato0.text + "' and numero='" + DATO1.text + "' order by numero desc"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    For K = 1 To CANLI
    tipo2 = Mid(grilladetalle(K, 4), 2, 2)
    If tipo2 <> "00" Then Call desactualizactacte(K, tipo2, grilladetalle(K, 5))
    If Mid(grilladetalle(K, 4), 1, 1) = "S" Then Call desactualizacrcc(K, grilladetalle(K, 6))
    If Mid(grilladetalle(K, 4), 1, 1) = "S" Then Call desactualizacrcc(K, Mid(grilladetalle(K, 6), 1, 2) + "00")
    Call desactualizamayor(K, grilladetalle(K, 1) + grilladetalle(K, 2) + grilladetalle(K, 3))
    Call desactualizamayor(K, grilladetalle(K, 1) + grilladetalle(K, 2) + "0000")
    Call desactualizamayor(K, grilladetalle(K, 1) + "000000")
    Next K

End Sub

Sub ESFECHA(ByRef DIAS As Integer, ByRef mes As Integer, ByRef ANO As Integer)
    
    If DIAS < 1 Or DIAS > 31 Then ef = "N": GoTo no
    If mes < 1 Or mes > 12 Then ef = "N": GoTo no
    If ANO < 1999 Then ef = "N": GoTo no:
    ef = "S"

no:

End Sub




Sub leermayor(row As Long, col As Long)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "ctacte"
    campos(3, 0) = "centrocosto"
    campos(4, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + Grid1.Cell(row, 1).text + Grid1.Cell(row, 2).text + Grid1.Cell(row, 3).text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
 
    If SQLUTIL.ESTADO = 4 Then
    RESPUESTA = "N"
    Grid1.Cell(row, col).text = ""
    Grid1.Cell(row, col + 1).text = ""
    Grid1.Cell(row, col + 2).text = ""
    Grid1.Cell(row, col + 5).text = ""
    Else
    RESPUESTA = "S"
    nombremayor.Caption = SQLUTIL.datos(1, 3)
    tipoctacte = SQLUTIL.datos(2, 3)
    If Grid1.Cell(row, 5).text = "" Then Grid1.Cell(row, 5).text = String(10, 32)
    If Grid1.Cell(row, 6).text = "" Then Grid1.Cell(row, 6).text = String(4, 32)
    If Grid1.Cell(row, 4).text = "" Then Grid1.Cell(row, 4).text = "N00"
    If tipoctacte <> "00" Then Grid1.Cell(row, 4).text = "C" + tipoctacte
    cc = SQLUTIL.datos(3, 3)
    If cc = 3 Then Grid1.Cell(row, 4).text = "B" + tipoctacte
    If cc = 2 Then Grid1.Cell(row, 4).text = "S" + tipoctacte
no:
       
    
    End If

End Sub


Sub SIGUIENTE()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + dato0.text + "' and numero>'" + DATO1.text + "' order by numero asc"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then dato0.text = SQLUTIL.datos(0, 3): DATO1.text = SQLUTIL.datos(1, 3): leecomprobante
End Sub

Sub anterior()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + dato0.text + "' and numero<'" + DATO1.text + "' order by numero desc"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then dato0.text = SQLUTIL.datos(0, 3): DATO1.text = SQLUTIL.datos(1, 3): leecomprobante
    
End Sub

Private Sub opciones_GotFocus()
    MANUAL.SetFocus
End Sub
Sub ayudamayor(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    cfijo = "no"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentasdelmayor", pivote2, campos, cfijo, largo, 2)
    Grid1.Cell(row, col).text = Mid(pivote2.text, 1, 2)
    Grid1.Cell(row, col + 1).text = Mid(pivote2.text, 3, 2)
    Grid1.Cell(row, col + 2).text = Mid(pivote2.text, 5, 4)
    Rem Call leermayor(row, col)
    respu = ""
    If pivote2.text <> "" Then Call leermayor(row, col): respu = "S"
    pivote2.text = ""
    
End Sub
Sub ayudacrcc(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "no"
    pivote2.MaxLength = 4
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "centrosdecosto", pivote2, campos, cfijo, largo, 2)
    Grid1.Cell(row, col).text = pivote2.text
    pivote2.text = ""
End Sub
Sub ayudactacte(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    cabezas = Array("rut", "nombre")
    largo = Array("12n", "40s")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    cfijo = "tipo='" & tipoctacte & "'"
    pivote2.MaxLength = 10
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentascorrientes", pivote2, campos, cfijo, largo, 2)
    Grid1.Cell(row, col).text = pivote2.text
    pivote2.text = ""
End Sub

Sub leerctacte(row As Long, col As Long)
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + Mid(Grid1.Cell(row, 4).text, 2, 2) + "' and rut=" + "'" + Grid1.Cell(row, col).text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Grid1.Cell(row, col).text = "": GoTo no:
    nombrectacte.Caption = SQLUTIL.datos(1, 3)
no:
End Sub

Sub leercrcc(row As Long, col As Long)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + Grid1.Cell(row, 6).text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Grid1.Cell(row, 6).text = "": GoTo no:
    nombrecrcc.Caption = SQLUTIL.datos(1, 3)
no:
End Sub
Sub SUMAR()
SUMADEBE = 0
SUMAHABER = 0

For o = 1 To Grid1.Rows - 1
If Grid1.Cell(o, 12).text = "D" Then SUMADEBE = SUMADEBE + Grid1.Cell(o, 11).text
If Grid1.Cell(o, 12).text = "H" Then SUMAHABER = SUMAHABER + Grid1.Cell(o, 11).text
Next o
debe.Caption = Format(SUMADEBE, "###,###,###,##0")
haber.Caption = Format(SUMAHABER, "###,###,###,##0")
saldo.Caption = Format(SUMADEBE - SUMAHABER, "###,###,###,##0")
End Sub

Sub ayudatipos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("tipos", "nombredocumento")
    cabezas = Array("TIPOS", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Tipos de Documentos"
    cfijo = "no"
    
    Call cargaAyudaT(servidor, "conta", USUARIO, password, "maestrotipodedocumentos", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    Call leetipos(caja)
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

Sub leetipos(caja As TextBox)
    
    campos(0, 0) = "tipos"
    campos(1, 0) = "nombredocumento"
    campos(2, 0) = ""
    
    campos(0, 2) = "maestrotipodedocumentos"
    condicion = "tipos=" + "'" + caja.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.ESTADO = 4 Then caja.text = "": caja.SetFocus:  GoTo no:
    tipocompro.Caption = SQLUTIL.datos(1, 3)
    
    

no:

End Sub

Sub actualizamayor(row, cuentamayor As String)
    Dim tipo2 As String

    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + DATO3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + DATO3.text
    campos(3, 0) = ""
    
    condicion = "codigo=" + "'" + cuentamayor + "' and año ='" + DATO4.text + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.ESTADO = 4 Then Stop
   
    VARIPASO = Replace(Grid1.Cell(row, 11).text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto + Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    
    
End Sub

Sub desactualizamayor(row, cuentamayor As String)
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + DATO3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + DATO3.text
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + cuentamayor + "' and año ='" + DATO4.text + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    If SQLUTIL.ESTADO = 4 Then Stop
   
    VARIPASO = Replace(grilladetalle(row, 11), ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    
    campos(2, 1) = Str(varimonto - Val(VARIPASO))
    
    op = 3
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    

    
End Sub
Sub actualizactacte(row, tipo As String, rut As String)
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + DATO3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + DATO3.text
    campos(3, 0) = ""
    condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año ='" + DATO4.text + "' order by tipo"
    campos(0, 2) = "saldosctacte"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.ESTADO = 4 Then Stop
   
    VARIPASO = Replace(Grid1.Cell(row, 11).text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto + Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    

    
End Sub

Sub desactualizactacte(row, tipo As String, rut As String)
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + DATO3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + DATO3.text
    campos(3, 0) = ""
    condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año ='" + DATO4.text + "' order by tipo"
    campos(0, 2) = "saldosctacte"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.ESTADO = 4 Then Stop
   
    VARIPASO = Replace(Grid1.Cell(row, 11).text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto - Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    

    
End Sub



Sub actualizacrcc(row, crcc As String)
    

    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + DATO3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + DATO3.text
    campos(3, 0) = ""
    
    condicion = "codigo=" + "'" + crcc + "' and año ='" + DATO4.text + "' order by codigo"
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.ESTADO = 4 Then Stop
   
    VARIPASO = Replace(Grid1.Cell(row, 11).text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto + Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    
    
End Sub

Sub desactualizacrcc(row, crcc As String)
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + DATO3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + DATO3.text
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + crcc + "' and año ='" + DATO4.text + "' order by codigo"
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    VARIPASO = Replace(grilladetalle(row, 11), ".", "")
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto - Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    
End Sub

Sub ULTIMO()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + dato0.text + "' order by numero desc"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    DATO1.text = SQLUTIL.datos(1, 3) + 1

End Sub

