VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form DATOSPAGO 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATOS PROVEEDOR"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4800
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   8467
      BackColor       =   16761024
      Caption         =   "Antecedentes Bancarios"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPFrame.FrameXp FRMTIPO 
         Height          =   1455
         Left            =   2880
         TabIndex        =   24
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2566
         BackColor       =   16744576
         Caption         =   "TIPOS DE PAGO"
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
         Alignment       =   1
         Begin VB.Label Label13 
            BackColor       =   &H00FF8080&
            Caption         =   "3 = TRANSFERENCIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FF8080&
            Caption         =   "1 = CHEQUES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FF8080&
            Caption         =   "0 = VALE VISTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox dato11 
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
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   22
         Tag             =   "Comuna"
         Top             =   3720
         Width           =   555
      End
      Begin VB.TextBox dato10 
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
         Left            =   1710
         MaxLength       =   200
         TabIndex        =   7
         Tag             =   "fono"
         Top             =   3285
         Width           =   5955
      End
      Begin VB.TextBox dato9 
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
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "fono"
         Top             =   2880
         Width           =   6000
      End
      Begin VB.TextBox dato8 
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
         Left            =   1710
         MaxLength       =   12
         TabIndex        =   5
         Tag             =   "fono"
         Top             =   2475
         Width           =   1815
      End
      Begin VB.TextBox dato7 
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
         Left            =   1710
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "giro"
         Top             =   2115
         Width           =   2505
      End
      Begin VB.TextBox dato6 
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
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "ciudad"
         Top             =   1755
         Width           =   555
      End
      Begin VB.TextBox dato5 
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
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "Comuna"
         Top             =   1395
         Width           =   555
      End
      Begin VB.TextBox dato3 
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
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "nombre"
         Top             =   675
         Width           =   6015
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   9
         Tag             =   "rut"
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox dato4 
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
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   1
         Tag             =   "direccion"
         Top             =   1035
         Width           =   300
      End
      Begin VB.Label lbltipo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plazo Pago"
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
         TabIndex        =   23
         Top             =   3720
         Width           =   1530
      End
      Begin VB.Label lblbanco 
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
         Height          =   330
         Left            =   2400
         TabIndex        =   21
         Top             =   1395
         Width           =   5550
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email "
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
         Left            =   90
         TabIndex        =   19
         Top             =   3285
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre Retira"
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
         Left            =   90
         TabIndex        =   18
         Top             =   2880
         Width           =   1530
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2925
         TabIndex        =   17
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut Retira"
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
         Left            =   90
         TabIndex        =   16
         Top             =   2475
         Width           =   1530
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Corriente"
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
         Left            =   90
         TabIndex        =   15
         Top             =   2115
         Width           =   1530
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sucursal"
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
         Left            =   90
         TabIndex        =   14
         Top             =   1755
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         Left            =   90
         TabIndex        =   13
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modo Pago"
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
         Left            =   90
         TabIndex        =   12
         Top             =   1035
         Width           =   1530
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   90
         TabIndex        =   11
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut"
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
         Left            =   90
         TabIndex        =   10
         Top             =   315
         Width           =   1530
      End
   End
   Begin CoolButtons.cool_Button graba 
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   4920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "&Grabar"
   End
End
Attribute VB_Name = "DATOSPAGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MODI As Boolean

Private Sub dato4_LostFocus()
frmTipo.Visible = False
End Sub

Private Sub graba_Click()
If dato11.text <> "" Then
grabar
Unload Me
Else
MsgBox ("DEBE INGRESAR TODOS LOS CAMPOS")
dato3.SetFocus

End If

End Sub

Private Sub dato3_GotFocus()
Call cargatexto(dato4)
End Sub


Private Sub dato4_GotFocus()
frmTipo.Visible = True
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
Private Sub dato10_GotFocus()
Call cargatexto(dato10)
End Sub
Private Sub dato3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dato4.SetFocus
End If

End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
If dato4.text = "1" Or dato4.text = "3" Or dato4.text = "0" Then
If dato4.text = "1" Then LBLTIPO.Caption = "VALE VISTA"
If dato4.text = "0" Then LBLTIPO.Caption = "CHEQUES"
If dato4.text = "3" Then LBLTIPO.Caption = "TRANSFERENCIA"

DATO5.SetFocus
Else
dato4.SetFocus
End If

End If

End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 Then
Call ceros(DATO5)

If leebanco(DATO5.text) <> "" Then

lblBanco.Caption = leebanco(DATO5.text)
dato6.SetFocus
Else
DATO5.SetFocus

End If
End If

End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dato7.SetFocus
End If

End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dato8.SetFocus
End If

End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dato9.SetFocus
End If

End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dato10.SetFocus
End If

End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
dato11.SetFocus
End If

End Sub

Private Sub dato11_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then

Call ceros(dato11)
graba.SetFocus

End If

End Sub


Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call flechas(dato3, dato4, KeyCode)
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
        Call flechas(dato8, dato10, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato9, dato3, KeyCode)
End Sub

Private Sub Form_Load()
dato2.text = maestro02.dato2.text
DV.Caption = maestro02.DV.Caption
dato3.text = maestro02.dato4.text
frmTipo.Visible = False

leer
End Sub
Sub leer()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = "modopago"
    campos(3, 0) = "banco"
    campos(4, 0) = "sucursal"
    campos(5, 0) = "cuentacorriente"
    campos(6, 0) = "rutretira"
    campos(7, 0) = "nombreretira"
    campos(8, 0) = "email"
    campos(9, 0) = "plazo"
    campos(10, 0) = ""
    campos(0, 2) = "cuentascorrientes_datos_pago"
    condicion = "rut=" + "'" + dato2.text + DV.Caption + "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    MODI = True
    
    carga
    Else
    MODI = False
    
   End If
no:
End Sub
Sub grabar()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = "modopago"
    campos(3, 0) = "banco"
    campos(4, 0) = "sucursal"
    campos(5, 0) = "cuentacorriente"
    campos(6, 0) = "rutretira"
    campos(7, 0) = "nombreretira"
    campos(8, 0) = "email"
    campos(9, 0) = "plazo"
    campos(10, 0) = ""
    campos(0, 1) = dato2.text + DV.Caption
    campos(1, 1) = dato3.text
    campos(2, 1) = dato4.text
    campos(3, 1) = DATO5.text
    campos(4, 1) = dato6.text
    campos(5, 1) = dato7.text
    campos(6, 1) = dato8.text
    campos(7, 1) = dato9.text
    campos(8, 1) = dato10.text
    campos(9, 1) = dato11.text
    
    
    
    
    campos(0, 2) = "cuentascorrientes_datos_pago"
    condicion = "rut=" + "'" + dato2.text + DV.Caption + "' "
    If MODI = True Then
    op = 3
    Else
    condicion = ""
    op = 2
    End If
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    

End Sub

Sub carga()
    dato2.text = Mid(sqlconta.response(0, 3), 1, 9)
    DV.Caption = Mid(sqlconta.response(0, 3), 10, 1)
    dato3.text = sqlconta.response(1, 3)
    dato4.text = sqlconta.response(2, 3)
    DATO5.text = sqlconta.response(3, 3)
    dato6.text = sqlconta.response(4, 3)
    dato7.text = sqlconta.response(5, 3)
    dato8.text = sqlconta.response(6, 3)
    dato9.text = sqlconta.response(7, 3)
    dato10.text = sqlconta.response(8, 3)
    dato11.text = sqlconta.response(9, 3)
    lblBanco.Caption = leebanco(DATO5.text)
If dato4.text = "1" Then LBLTIPO.Caption = "VALE VISTA"
If dato4.text = "0" Then LBLTIPO.Caption = "CHEQUES"
If dato4.text = "3" Then LBLTIPO.Caption = "TRANSFERENCIA"
    
End Sub

