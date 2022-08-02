VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form muestrach 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribuye Cheques"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   7200
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5910
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   10425
      BackColor       =   16744576
      Caption         =   "DATOS DEL CHEQUE"
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
      Begin VB.TextBox DATO5 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   28
         Tag             =   "nombre"
         Top             =   2790
         Width           =   270
      End
      Begin VB.CommandButton ver 
         Caption         =   "Command1"
         Height          =   285
         Left            =   8100
         TabIndex        =   27
         Top             =   450
         Visible         =   0   'False
         Width           =   780
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   555
         Left            =   90
         TabIndex        =   25
         Top             =   8505
         Visible         =   0   'False
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   979
         BackColor       =   49344
         Caption         =   "Cheques pendientes de Cobro"
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
         Begin FlexCell.Grid Grid1 
            Height          =   5505
            Left            =   0
            TabIndex        =   26
            Top             =   225
            Width           =   9600
            _ExtentX        =   16933
            _ExtentY        =   9710
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "codigo"
         Top             =   285
         Width           =   375
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   2
         TabIndex        =   2
         Top             =   285
         Width           =   375
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   3
         Top             =   285
         Width           =   735
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "nombre"
         Top             =   630
         Width           =   1215
      End
      Begin XPFrame.FrameXp TIPOS 
         Height          =   2880
         Left            =   6210
         TabIndex        =   30
         Top             =   2835
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5080
         BackColor       =   16761024
         Caption         =   "DISTRIBUCION DE CHEQUES"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRILLATIPO 
            Height          =   2520
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   4445
            _Version        =   393216
            BackColor       =   16107953
            ForeColor       =   16711680
            Rows            =   3
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   16777152
            BackColorBkg    =   16761024
            GridColor       =   16744576
            GridColorFixed  =   14282751
            GridColorUnpopulated=   14282751
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Label lblfecha 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   34
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado del Cheque"
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
         Left            =   180
         TabIndex        =   33
         Top             =   3195
         Width           =   5640
      End
      Begin VB.Label lbldis 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   180
         TabIndex        =   32
         Top             =   3465
         Width           =   5670
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distribucion"
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
         Left            =   90
         TabIndex        =   29
         Top             =   2790
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Banco"
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
         Left            =   90
         TabIndex        =   24
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Numero"
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
         Left            =   90
         TabIndex        =   23
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
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
         Left            =   90
         TabIndex        =   22
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vencimiento"
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
         Left            =   90
         TabIndex        =   21
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Girado a"
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
         Left            =   90
         TabIndex        =   20
         Top             =   1710
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comprobante"
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
         Left            =   90
         TabIndex        =   19
         Top             =   2070
         Width           =   1455
      End
      Begin VB.Label monto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1620
         TabIndex        =   18
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label vencimiento 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1620
         TabIndex        =   17
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label giradoa 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1620
         TabIndex        =   16
         Top             =   1710
         Width           =   5400
      End
      Begin VB.Label comprobante 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label numero 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1530
         TabIndex        =   14
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Numero"
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
         Left            =   1530
         TabIndex        =   13
         Top             =   2070
         Width           =   1230
      End
      Begin VB.Label fechaemision 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2745
         TabIndex        =   12
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Emision"
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
         Left            =   2745
         TabIndex        =   11
         Top             =   2070
         Width           =   1455
      End
      Begin VB.Label fechacobro 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4140
         TabIndex        =   10
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Cobro"
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
         Left            =   4185
         TabIndex        =   9
         Top             =   2070
         Width           =   1410
      End
      Begin VB.Label estado 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5580
         TabIndex        =   8
         Top             =   2340
         Width           =   1470
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cobrado"
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
         Left            =   5580
         TabIndex        =   7
         Top             =   2070
         Width           =   1455
      End
   End
End
Attribute VB_Name = "muestrach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub Command1_Click()
imprimir
End Sub


Private Sub Cobrar_Click()
    campos(0, 0) = "cobrado"
    campos(1, 0) = "fechacobro"
    campos(2, 0) = ""
    campos(0, 1) = "1"
    campos(1, 1) = año + "-" + MES + "-" + dia
    
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta=" + "'" + dato1.text + dato2.text + dato3.text + "' and numero='" + dato4.text + "' order by numero"
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   Rem  VER_Click
    dato4.text = ""
    dato4.SetFocus
    limpiar
End Sub

Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub
Private Sub dato3_GotFocus()
Call cargatexto(dato3)
End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudamayor(dato4)
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub dato4_GotFocus()
Rem VER_Click

End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyF2 Then Call ayudacheques(dato5)
    Call flechas(dato3, dato4, KeyCode)

End Sub

Private Sub dato5_GotFocus()
Call cargatexto(DATO5)

End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
dato4.text = ""
DATO5.text = "0"
dato4.SetFocus
lbldis.Caption = ""
fechacobro.Caption = ""
estado.Caption = ""
End If
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And DATO5.text <> "" Then
modifica
leercheque

End If



End Sub

Private Sub dato5_LostFocus()
If DATO5.text > "9" Then
DATO5.SetFocus

End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    monto.Caption = ""
    vencimiento.Caption = ""
    giradoa.Caption = ""
    comprobante.Caption = ""
    numero.Caption = ""
    fechaemision.Caption = ""
End If

End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA
GRILLATIPOS
dato1.text = da1
dato2.text = da2
dato3.text = da3
dato4.text = da4
Call dato4_KeyPress(13)


    

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
    If KeyAscii = 13 Then
    Call ceros(dato3)
    Call Pregunta(dato3, dato4)
   Rem  VER_Click
    
    
End If

End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato4)

    leercheque
    
    
    End If
    
End Sub

Sub leercheque()
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "fechacobro"
    campos(9, 0) = "ubicacion"
    campos(10, 0) = "fechamovimiento"
    campos(11, 0) = "cobrado"
    campos(12, 0) = ""
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta=" + "'" + dato1.text + dato2.text + dato3.text + "' and numero='" + dato4.text + "' order by numero"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    dato4.text = sqlconta.response(1, 3)
    monto.Caption = Format(sqlconta.response(3, 3), "###,###,###,###")
    vencimiento.Caption = Mid(sqlconta.response(4, 3), 1, 2) + "-" + Mid(sqlconta.response(4, 3), 4, 2) + "-" + Mid(sqlconta.response(4, 3), 7, 4)
    giradoa.Caption = sqlconta.response(7, 3)
    comprobante.Caption = sqlconta.response(5, 3)
    numero.Caption = sqlconta.response(6, 3)
    fechaemision.Caption = Mid(sqlconta.response(2, 3), 1, 2) + "-" + Mid(sqlconta.response(2, 3), 4, 2) + "-" + Mid(sqlconta.response(2, 3), 7, 4)
    lblfecha.Caption = Format(sqlconta.response(10, 3), "dd-mm-yyyy")
    
    estado.Caption = "SIN COBRAR"
    If sqlconta.response(8, 3) <> "01/01/1800" Then
    fechacobro.Caption = Mid(sqlconta.response(8, 3), 1, 2) + "-" + Mid(sqlconta.response(8, 3), 4, 2) + "-" + Mid(sqlconta.response(8, 3), 7, 4)
    estado.Caption = "COBRADO"
    End If
    
    'ariel agrega nueva opciones de cheque anulado
    
    If sqlconta.response(11, 3) = "1" Then estado.Caption = "COBRADO AUTOMATICO"
    If sqlconta.response(11, 3) = "2" Then estado.Caption = "COBRADO MANUAL"
    If sqlconta.response(11, 3) = "3" Then estado.Caption = "CADUCADO"
    If sqlconta.response(11, 3) = "4" Then estado.Caption = "O.N.P"
    If sqlconta.response(11, 3) = "5" Then estado.Caption = "ANULADO"
    
    
    DATO5.text = sqlconta.response(9, 3)
   
    lbldis.Caption = GRILLATIPO.TextMatrix(CDbl(DATO5.text), 1)
    
    Else
    MsgBox ("NUMERO DE CHEQUE NO EXISTE")
   Unload Me
   
    
    
    
    End If
   
    
    
End Sub


Sub leer()
    Rem lee cuenta madre
  
lee2:    Rem lee cuenta madre
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    
    
no:
   
End Sub
Sub leersiguiente()
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "fechacobro"
    campos(9, 0) = "ubicacion"
    campos(10, 0) = ""

    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta=" + "'" + dato1.text + dato2.text + dato3.text + "' and numero>'" + dato4.text + "' order by numero"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    dato4.text = sqlconta.response(1, 3)
    
    monto.Caption = Format(sqlconta.response(3, 3), "###,###,###,###")
    vencimiento.Caption = Mid(sqlconta.response(4, 3), 1, 2) + "-" + Mid(sqlconta.response(4, 3), 4, 2) + "-" + Mid(sqlconta.response(4, 3), 7, 4)
    giradoa.Caption = sqlconta.response(7, 3)
    comprobante.Caption = sqlconta.response(5, 3)
    numero.Caption = sqlconta.response(6, 3)
    fechaemision.Caption = Mid(sqlconta.response(2, 3), 1, 2) + "-" + Mid(sqlconta.response(2, 3), 4, 2) + "-" + Mid(sqlconta.response(2, 3), 7, 4)
    If sqlconta.response(8, 3) <> "01/01/1800" Then fechacobro.Caption = Mid(sqlconta.response(8, 3), 1, 2) + "-" + Mid(sqlconta.response(8, 3), 4, 2) + "-" + Mid(sqlconta.response(8, 3), 7, 4)

    estado.Caption = sqlconta.response(9, 3)
    

    disponible (True)
    habilita (True)
    
no:
   
    
End Sub
Sub leeranterior()
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "fechacobro"
    campos(9, 0) = "ubicacion"
    campos(10, 0) = "fechamovimiento"
    campos(11, 0) = ""
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta=" + "'" + dato1.text + dato2.text + dato3.text + "' and numero<'" + dato4.text + "' order by numero desc"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    dato4.text = sqlconta.response(1, 3)
    
    monto.Caption = Format(sqlconta.response(3, 3), "###,###,###,###")
    vencimiento.Caption = Mid(sqlconta.response(4, 3), 1, 2) + "-" + Mid(sqlconta.response(4, 3), 4, 2) + "-" + Mid(sqlconta.response(4, 3), 7, 4)
    giradoa.Caption = sqlconta.response(7, 3)
    comprobante.Caption = sqlconta.response(5, 3)
    numero.Caption = sqlconta.response(6, 3)
    fechaemision.Caption = Mid(sqlconta.response(2, 3), 1, 2) + "-" + Mid(sqlconta.response(2, 3), 4, 2) + "-" + Mid(sqlconta.response(2, 3), 7, 4)
    lblfecha.Caption = Format(sqlconta.response(10, 3), "dd-mm-yyyy")
    If sqlconta.response(8, 3) <> "01/01/1800" Then fechacobro.Caption = Mid(sqlconta.response(8, 3), 1, 2) + "-" + Mid(sqlconta.response(8, 3), 4, 2) + "-" + Mid(sqlconta.response(8, 3), 7, 4)
    estado.Caption = sqlconta.response(9, 3)
    disponible (True)
    habilita (True)
    
no:
   
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    dato4.text = sqlconta.response(1, 3)
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    

End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "banco='1'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    dato2.text = Mid(pivote.text, 3, 2)
    dato3.text = Mid(pivote.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
    
no:
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub ELIMINAR()
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub


Private Sub lblhistorico_Click(Index As Integer)

End Sub






Private Sub liberar_Click()
    campos(0, 0) = "cobrado"
    campos(1, 0) = "fechacobro"
    campos(2, 0) = ""
    campos(0, 1) = "0"
    campos(1, 1) = "0000/00/00"
    
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta=" + "'" + dato1.text + dato2.text + dato3.text + "' and numero='" + dato4.text + "' order by numero"
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    VER_Click
   
    dato4.text = ""
    dato4.SetFocus
    limpiar
    

End Sub

Sub retorno()
disponible (True)
habilita (False)
limpia

dato1.Enabled = True
dato1.SetFocus
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    monto.Caption = ""
    vencimiento.Caption = ""
    giradoa.Caption = ""
    comprobante.Caption = ""
    numero.Caption = ""
    fechaemision.Caption = ""
    fechacobro.Caption = ""
    estado.Caption = ""

End Sub

Sub imprimir()
    
   
End Sub
Sub grilla()
    
End Sub
Sub cabeza()
    

End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        csql.sql = csql.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4): colu(1) = 15: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                dato(3) = resultados(2) + " " + DOCU$(Val(resultados(2)))
                dato(4) = resultados(3)
                dato(5) = resultados(4)
                dato(6) = resultados(5) + " " + DOCU2$(Val(resultados(5)))
                colu(3) = 15: tipodato(3) = "s"
                colu(4) = 3: tipodato(4) = "s"
                colu(5) = 20: tipodato(5) = "s"
                colu(6) = 20: tipodato(6) = "s"
                 cancolu = 6
                grilla
                resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    

End Sub


Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 10)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "NUMERO"
    FORMATOGRILLA(1, 2) = "GIRADO A"
    FORMATOGRILLA(1, 3) = "MONTO"
    FORMATOGRILLA(1, 4) = "VENCIMIENTO"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "40"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "D"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = "###,###,###,##0"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    
    Grid1.Cols = 5
    Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
   Grid1.BackColorFixed = RGB(90, 158, 214)
   Grid1.BackColorFixedSel = RGB(110, 180, 230)
   Grid1.BackColorBkg = RGB(90, 158, 214)
   Grid1.BackColorScrollBar = RGB(231, 235, 247)
   Grid1.BackColor1 = RGB(231, 235, 247)
   Grid1.BackColor2 = RGB(239, 243, 255)
   Grid1.GridColor = RGB(148, 190, 231)
   Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub



Private Sub VER_Click()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim total2 As Double
    LINEA = 0
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT emision,tipocomprobante,numerocomprobante,numero,giradoa,vencimiento,monto,fechacobro,cobrado "
        csql.sql = csql.sql + "FROM chequesdocumento where cuenta='" + dato1.text + dato2.text + dato3.text + "'"
        csql.sql = csql.sql + "order by cuenta,vencimiento"
        csql.Execute
        total = 0
        total2 = 0
        
        
        
        Grid1.Rows = csql.RowsAffected + 1
        
        
        If csql.RowsAffected > 0 Then
        
        Grid1.Rows = csql.RowsAffected + 1
        
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
          
         
         
         If resultados(8) = "0" Then
          LINEA = LINEA + 1
             Grid1.Cell(LINEA, 1).text = resultados(3)
             Grid1.Cell(LINEA, 2).text = resultados(4)
             Grid1.Cell(LINEA, 3).text = resultados(6)
             Grid1.Cell(LINEA, 4).text = resultados(5)
             total = total + resultados(6)
          Else
          End If
          resultados.MoveNext

         Wend
End If
        Rem TOTALES.Caption = Format(total, "###,###,###,##0")
End Sub
Sub limpiar()
monto.Caption = ""
vencimiento.Caption = ""
giradoa.Caption = ""
comprobante.Caption = ""
numero.Caption = ""
fechaemision.Caption = ""
fechacobro.Caption = ""
estado.Caption = ""


End Sub
Sub GRILLATIPOS()
GRILLATIPO.Cols = 2
GRILLATIPO.Rows = 10
GRILLATIPO.ColWidth(0) = 200 * 2
GRILLATIPO.ColWidth(1) = 200 * 10

GRILLATIPO.TextMatrix(0, 0) = "0"
GRILLATIPO.TextMatrix(1, 0) = "1"
GRILLATIPO.TextMatrix(2, 0) = "2"
GRILLATIPO.TextMatrix(3, 0) = "3"
GRILLATIPO.TextMatrix(4, 0) = "4"
GRILLATIPO.TextMatrix(5, 0) = "5"
GRILLATIPO.TextMatrix(6, 0) = "6"
GRILLATIPO.TextMatrix(7, 0) = "7"
GRILLATIPO.TextMatrix(8, 0) = "8"
GRILLATIPO.TextMatrix(9, 0) = "9"

GRILLATIPO.TextMatrix(0, 1) = "EMITIDO SIN FIRMA"
GRILLATIPO.TextMatrix(1, 1) = "ENVIADO POR CORREO"
GRILLATIPO.TextMatrix(2, 1) = "EN PODER SECRETARIA"
GRILLATIPO.TextMatrix(3, 1) = "RETENIDO POR PUBLICIDAD"
GRILLATIPO.TextMatrix(4, 1) = "RETENIDO POR GERENCIA"
GRILLATIPO.TextMatrix(5, 1) = "ENTREGADO A PROVEEDOR"
GRILLATIPO.TextMatrix(6, 1) = "GIRADO A FECHA   "
GRILLATIPO.TextMatrix(7, 1) = "ENVIADO BUSES JAC"
GRILLATIPO.TextMatrix(8, 1) = "ENVIADO BUSES TURBUS"
GRILLATIPO.TextMatrix(9, 1) = "DEPOSITO EN CTA CTE "

CANDO = 9



End Sub

Sub modifica()
    
    campos(0, 0) = "ubicacion"
    campos(1, 0) = ""
    campos(0, 1) = DATO5.text
    
    campos(0, 2) = "chequesdocumento"
    
    condicion = "cuenta=" + "'" + dato1.text + dato2.text + dato3.text + "' and numero='" + dato4.text + "' "
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
End Sub

