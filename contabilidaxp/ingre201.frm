VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form ingreso010 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes Contables"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14235
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   14235
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   4080
      MaxLength       =   8
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   7560
      TabIndex        =   12
      Top             =   8760
      Width           =   6015
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
         TabIndex        =   21
         Top             =   360
         Width           =   1695
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
         TabIndex        =   20
         Top             =   360
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
         TabIndex        =   19
         Top             =   360
         Width           =   1575
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
         TabIndex        =   18
         Top             =   120
         Width           =   1695
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
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   6015
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
         TabIndex        =   16
         Top             =   1560
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
         TabIndex        =   15
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
         Index           =   5
         Left            =   2640
         TabIndex        =   14
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEBE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame comprocabeza 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13815
      Begin VB.TextBox dato5 
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
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "fecha"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox dato4 
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
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   240
         Width           =   375
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
         Left            =   8160
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   240
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
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "tipo"
         Top             =   240
         Width           =   375
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
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "numero"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA EMISION"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6600
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   13815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
         Height          =   255
         Left            =   9960
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIPO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO   :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label nombrecomprobante 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame comprocuerpo 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   13815
      Begin FlexCell.Grid Grid1 
         Height          =   6975
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   12303
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Height          =   7215
         Left            =   0
         Top             =   0
         Width           =   13815
      End
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   735
      Left            =   7680
      Top             =   8880
      Width           =   6015
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   1455
      Left            =   240
      Top             =   1200
      Width           =   13815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   5535
      Left            =   240
      Top             =   2880
      Width           =   13815
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   240
      TabIndex        =   24
      Top             =   8520
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   615
      Left            =   240
      Top             =   360
      Width           =   13815
   End
End
Attribute VB_Name = "ingreso010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub dato1_Change()
If Val(DATO1.text) < 1 Or Val(DATO1.text) > CANDO Then DATO1.Enabled = True: DATO1.text = "": DATO1.SetFocus
End Sub

Private Sub DATO1_GotFocus()

limpia
limpia2


Call cargatexto(DATO1)

End Sub

Private Sub dato19_Change()
If dato19.text <> "D" And dato19.text <> "H" Then dato19.text = ""
End Sub

Private Sub dato19_GotFocus()
If dato19.text = "H" And modifi = 0 Then dato19.text = "D": GoTo PASO:
If dato19.text = "D" And modifi = 0 Then dato19.text = "H": GoTo PASO:
PASO:
Call cargatexto(dato19)
End Sub

Private Sub dato2_GotFocus()

Call cargatexto(DATO2)
If Val(DATO1.text) < 1 Or Val(DATO1.text) > CANDO Then DATO1.text = "": DATO1.SetFocus: GoTo no:
nombrecomprobante.Caption = DOCU(Val(DATO1.text))
no:
End Sub

Private Sub dato3_GotFocus()

If modifi = 0 Then LEERMOVIMIENTOS

If VARIPASO <> "0" Then opciones.Visible = True: comprodatos.Enabled = False: comprocabeza.Enabled = False: comprocuerpo.Enabled = False: opciones.SetFocus: limpia2: CREANDO = "": GoTo no:
CREANDO = "S": grilladocumentoS.Enabled = True

If Val(DATO2.text) = 0 Then DATO2.text = "": DATO2.Enabled = True: DATO2.SetFocus
Call cargatexto(DATO3)
no:
End Sub
Private Sub dato4_GotFocus()
If DATO3.text = "00" Then DATO4.Enabled = True: dato5.Enabled = True: dato6.Enabled = True: DATO3.text = Mid(fechasistema, 1, 2): DATO4.text = Mid(fechasistema, 4, 2): dato5.text = Mid(fechasistema, 7, 4): dato6.SetFocus
Call cargatexto(DATO4)
End Sub

Private Sub dato5_GotFocus()
Call cargatexto(dato5)
End Sub




Private Sub dato6_GotFocus()

Call cargatexto(dato6)
Call ESFECHA(Val(DATO3.text), Val(DATO4.text), Val(dato5.text))

If VARIPASO = "N" Then DATO3.text = "": DATO4.text = "": dato5.text = "": DATO3.SetFocus

End Sub


Private Sub dato7_GotFocus()
Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()
Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
Call cargatexto(dato9)
PIVOTE.text = dato6.text + dato7.text + dato8.text
leemayor
If tipocuenta = "00" Then dato9.text = "": dato10.Enabled = True: dato10.text = "": dato11.Enabled = True: dato11.SetFocus
End Sub

Private Sub dato10_GotFocus()
Call cargatexto(dato10)
   If tipocuenta = "00" Then dato10.Enabled = True: dato11.Enabled = True: dato11.SetFocus

End Sub

Private Sub dato11_GotFocus()
Call cargatexto(dato11)
If tipocuenta <> "00" Then pivote2.text = dato9.text + dato10.text: leectacte: GoTo no:
If tipocentro <> "2" Then dato11.text = "": dato12.Enabled = True: dato12.SetFocus
no:
If tipocuenta = "00" Then leemayor
End Sub

Private Sub dato12_GotFocus()
Call cargatexto(dato12)
If tipocentro = "2" Then leecrcc
If tipocentro <> "2" Then dato11.text = ""
Rem If tipocuenta <> "00" Then leectacte
If tipocentro <> "2" And tipocuenta = "00" Then leemayor
End Sub
Private Sub dato13_GotFocus()
Call cargatexto(dato13)
End Sub

Private Sub dato14_GotFocus()
Call cargatexto(dato14)
leetipos
End Sub
Private Sub dato15_GotFocus()
Call cargatexto(dato15)
End Sub
Private Sub dato16_GotFocus()
If dato15.text = "00" Then dato18.Enabled = True: dato16.Enabled = True: dato17.Enabled = True: dato15.text = Mid(fechasistema, 1, 2): dato16.text = Mid(fechasistema, 4, 2): dato17.text = Mid(fechasistema, 7, 4): dato18.SetFocus
Call cargatexto(dato16)
End Sub
Private Sub dato17_GotFocus()
Call cargatexto(dato17)
End Sub

Private Sub dato18_GotFocus()
Call cargatexto(dato18)
Call ESFECHA(Val(dato15.text), Val(dato16.text), Val(dato17.text))
If VARIPASO = "N" Then dato15.text = "": dato16.text = "": dato17.text = "": dato15.SetFocus

End Sub


Sub ESFECHA(ByRef DIAS As Integer, ByRef mes As Integer, ByRef ANO As Integer)
If DIAS < 1 Or DIAS > 31 Then VARIPASO = "N": GoTo no
If mes < 1 Or mes > 12 Then VARIPASO = "N": GoTo no
If ANO < 2005 Then VARIPASO = "N": GoTo no:
VARIPASO = "S"
no:

End Sub


Private Sub eliminacomprobante_Click()

opcionelimina.Visible = False
eliminatodo.Visible = True
Evento.Visible = True
textoevento.Caption = "PROCESO DE ELIMINACION DEL COMPROBANTE COMPLETO"


End Sub

Private Sub eliminalinea_Click()
comprocuerpo.Enabled = True
grilladocumentoS.Enabled = True
Evento.Visible = True
textoevento.Caption = "PROCESO DE ELIMINACION SELECCIONE UNA LINEA Y DOBLE CLICK"
opcionelimina.Visible = False
End Sub

Option Explicit
    Private modifi As Integer
    Private datogrilla2(100, 5) As String

