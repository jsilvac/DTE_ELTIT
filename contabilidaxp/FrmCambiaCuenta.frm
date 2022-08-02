VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form FrmCambiaCuenta 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "cambua cuenta comprobante"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9340
      BackColor       =   16761024
      Caption         =   "MODIFICA CUENTA COMPROBANTE CONTABLE"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   4194304
      ColorBarraArriba=   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   31
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   28
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3720
         Width           =   7335
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   2040
         MaxLength       =   9
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
      Begin Contabilidadxp.BotonMyERP CmdAceptar 
         Height          =   495
         Left            =   1320
         TabIndex        =   11
         Top             =   4680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "ACEPTAR"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Contabilidadxp.BotonMyERP CmdCancelar 
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   4680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "CANCELAR"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO DOCUMENTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " N. DOCUMENTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LINEA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   3720
         TabIndex        =   22
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label LINEA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5520
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " GLOSA CONTABLE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   7335
      End
      Begin VB.Label dv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label fecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label numero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label td 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblnombrectacte 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label lblnombrecrcc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label lblnombrecta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3600
         TabIndex        =   13
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTACTE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CENTRO COSTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CUENTA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmCambiaCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" & Format(fechasistema, "yyyy") & "' "
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", caja, campos, cfijo, largo, 2)
    If Val(caja) = 0 Then dato1.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus
no:
End Sub

Private Sub CmdAceptar_Click()
If DATO5.text = "CH" And dato6.text <> dato8.text Then
   If verificachequeemitido(dato1.text, dato6.text) = True Then
      MsgBox "NUMERO CHEQUE YA ESTA EMITIDO POR FAVOR REVISAR", vbCritical, "ATENCION"
        dato6.text = dato8.text
        dato6.SetFocus
        Exit Sub
   End If
End If

    If leerNombreMayor(dato1.text) <> "" Then
    Call MODIFICALINEA(td.Caption, numero.Caption, fecha.Caption, LINEA.Caption, dato1.text, dato3.text + dv.Caption, dato2.text, dato4.text, DATO5.text, dato6.text, dato8.text)
    End If
    
    Unload Me

End Sub
Function verificachequeemitido(cuenta, numero) As Boolean
    Dim csql As New rdoQuery
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select numero from chequesdocumento "
    csql.sql = csql.sql & "where numero='" & numero & "' and cuenta='" & cuenta & "' "
    csql.Execute
    verificachequeemitido = False
    If csql.RowsAffected > 0 Then
        verificachequeemitido = True
    End If
    csql.Close
    
End Function
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub dato1_Change()
lblnombrecta.Caption = Empty
End Sub

Private Sub dato1_GotFocus()
Call cargatexto(dato1)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato1)
        lblnombrecta.Caption = leerNombreMayor2(dato1)
    End If
End Sub

Private Sub dato2_Change()
lblnombrecrcc.Caption = Empty
End Sub

Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
        Call ceros(dato2)
     lblnombrecrcc.Caption = leerNOMBREcrcc(dato2)
    End If
End Sub

Private Sub dato3_Change()
dv.Caption = Empty
lblnombrectacte.Caption = Empty
End Sub

Private Sub dato3_GotFocus()
Call cargatexto(dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
    Call ceros(dato3)
    dv.Caption = rut(dato3)
    lblnombrectacte.Caption = nombrectacte(dato3.text & dv.Caption)
    End If
End Sub
 
Private Sub dato4_GotFocus()
Call cargatexto(dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato6)
        If dato6.text <> dato8.text Then
        If verificachequeemitido(dato1.text, dato6.text) = True Then
            MsgBox "NUMERO CHEQUE YA ESTA EMITIDO POR FAVOR REVISAR", vbCritical, "ATENCION"
            dato6.text = dato8.text
            dato6.SetFocus
        End If
        End If
    End If
    
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)
End Sub

Sub MODIFICALINEA(tipo, numero, fecha, LINEA, cuenta, rut, CRCC, glosa, tipodoc, NUMERODOC, numeroantiguo)
    campos(0, 0) = "codigocuenta"
    campos(1, 0) = "rutctacte"
    campos(2, 0) = "centrocosto"
    campos(3, 0) = "glosacontable"
    campos(4, 0) = "tipodocumento"
    campos(5, 0) = "numerodocumento"
    campos(6, 0) = ""
    
    campos(0, 1) = cuenta
    campos(1, 1) = rut
    campos(2, 1) = CRCC
    campos(3, 1) = glosa
    campos(4, 1) = tipodoc
    campos(5, 1) = NUMERODOC
    
    
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and linea='" + LINEA + "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' "
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If tipodoc = "CH" And Mid(cuenta, 1, 4) = "1112" And dato6.text <> dato8.text Then
        Call eliminacheque(cuenta, numeroantiguo, tipo, numero)
        Call grabacheque(dato9.text, dato10.text, dato11.text)
    End If
    
End Sub
Sub eliminacheque(cuenta, numero, tipocomprobante, numerocomprobante)
 
    campos(0, 0) = "cobrado"
    campos(1, 0) = "fechacobro"
    campos(2, 0) = "fechamovimiento"
    campos(3, 0) = ""
    
    
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta = '" & cuenta & "' AND numero = '" & numero & "' AND tipocomprobante = '" & tipocomprobante & "' AND numerocomprobante = '" & numerocomprobante & "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        dato9.text = sqlconta.response(0, 3)
        dato10.text = sqlconta.response(1, 3)
        dato11.text = sqlconta.response(2, 3)
        
    End If
    
    
    campos(0, 0) = ""
    
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta = '" & cuenta & "' AND numero = '" & numero & "' AND tipocomprobante = '" & tipocomprobante & "' AND numerocomprobante = '" & numerocomprobante & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
 
End Sub
Sub grabacheque(cobrado, fechacobro, fechamovimiento)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "ubicacion"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "fechamovimiento"
    campos(11, 0) = "cobrado"
    campos(12, 0) = "fechacobro"
    campos(13, 0) = ""
    
    
    
    campos(0, 1) = cartolamantencion.Grid1.Cell(cartolamantencion.Grid1.ActiveCell.row, 5).text
    campos(1, 1) = dato6.text
    campos(2, 1) = Format(fecha.Caption, "yyyy-mm-dd")
    campos(3, 1) = cartolamantencion.Grid1.Cell(cartolamantencion.Grid1.ActiveCell.row, 12).text 'leerdatos(contadb, "movimientoscontables", "monto", "tipo='" & td.Caption & "' and numero='" & numero.Caption & "' and fecha='" & Format(fecha.Caption, "yyyy-mm-dd") & "' and linea='" & LINEA.Caption & "' ")
    campos(4, 1) = Format(cartolamantencion.Grid1.Cell(cartolamantencion.Grid1.ActiveCell.row, 10).text, "yyyy-mm-dd")
    campos(5, 1) = td.Caption
    campos(6, 1) = numero.Caption
    campos(7, 1) = cartolamantencion.Grid1.Cell(cartolamantencion.Grid1.ActiveCell.row, 6).text
    campos(8, 1) = "0"
    campos(9, 1) = "CH"
    campos(10, 1) = Format(fechamovimiento, "yyyy-mm-dd")
    campos(11, 1) = cobrado
    campos(12, 1) = Format(fechacobro, "yyyy-mm-dd")
    
    campos(0, 2) = "chequesdocumento"
       
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
End Sub
Public Function leerNombreMayor2(codigo) As String
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = "crcc"
    campos(2, 0) = "ctacte"
    campos(3, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo = '" & codigo & "' "
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerNombreMayor2 = sqlconta.response(0, 3)
        If sqlconta.response(1, 3) = 1 Then 'CRCC
            Label1(3).Visible = True
                dato2.Visible = True
        lblnombrecrcc.Visible = True
                   dato2.text = Empty
        Else
            Label1(3).Visible = False
                dato2.text = Empty
                dato2.Visible = False
      lblnombrecrcc.Visible = False
        
        End If
        
        If sqlconta.response(2, 3) = 1 Then 'CTACTE
            Label1(5).Visible = True
                dato3.Visible = True
      lblnombrectacte.Visible = True
                   dv.Visible = True
                   dato3.text = Empty
        Else
            Label1(5).Visible = False
                dato3.Visible = False
                   dato3.text = Empty
                   dv.Visible = False
        lblnombrectacte.Visible = False
        End If
        
    Else
        leerNombreMayor2 = ""
    End If
End Function

