VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro11 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Cuentas del Mayor"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11490
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   712
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   8655
      Left            =   15
      TabIndex        =   4
      Top             =   45
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   15266
      BackColor       =   16744576
      Caption         =   "Datos"
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   12632256
      ColorBarraAbajo =   12582912
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
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6360
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   11055
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "nombre"
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox dato1 
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin CoolButtons.cool_Button graba 
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   8160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Grabar"
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Cuenta"
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   8040
      TabIndex        =   10
      Top             =   9960
      Width           =   3255
      _ExtentX        =   5741
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
         Left            =   1800
         TabIndex        =   12
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   280
         Width           =   1455
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1455
      Left            =   480
      TabIndex        =   3
      Top             =   9000
      Width           =   7455
      _cx             =   13150
      _cy             =   2566
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
End
Attribute VB_Name = "maestro11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub Command1_Click()
maestro21.Show

End Sub


Private Sub COMMAND2_Click()
MAESTRO20.Show

End Sub

Private Sub dato1_GotFocus()
leecrcc

Call cargatexto(dato1)
End Sub


Private Sub dato4_LostFocus()
Rem If modifi = 0 Then graba.Visible = True

End Sub




Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)

no:
End Sub








Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption
End Sub

Private Sub Form_Load()

'dibu1.FileName = App.path & "\archivo.gif"
'dibu2.FileName = App.path & "\archivo.gif"


    
    Call Conectar_BD

    sc = 0
    opciones.Visible = False


Rem Call RECUPERAFECHA
    
leecrcc

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    leer
    End If
    
End Sub



Sub leer()
    Rem lee cuenta madre
         
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = clientesistema + "conta.sii_1846"
      condicion = "codigo=" + "'" + dato1.text + "' "
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = conta
      Call sqlconta.sqlconta(op, condicion)
      If sqlconta.status = 0 Then
      
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    
    DATOSSALDOS
   
    
    opciones.SetFocus
    graba.Visible = False
Else
dato1.SetFocus

End If

   
End Sub

Sub carga()
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    graba.Visible = False
    
    
    
    

fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    
   
  
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    
 

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
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    basebus = clientesistema + "conta" + empresaactiva
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    caja.Enabled = True
    caja.SetFocus
    
no:
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub grabar()
       
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
    
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = ""
    campos(0, 2) = "cuentasdelmayor"
    If MODIFI = 1 Then condicion = "codigo=" + "'" + dato1.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If MODIFI = 0 Then grabar2
    MODIFI = 0
no:
retorno
End Sub
Sub grabar2()
     sqlconta.audit = False
     
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = Format(fechasistema, "yyyy")
    campos(0, 2) = "saldosdelmayor"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call GRABACRCC
    

End Sub
Sub GRABACRCC()
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM centrosdecosto "
       
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
       
        If csql2.RowsAffected > 0 Then
     
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
         Call GRABAR3(resultados2(0), dato1.text)
         
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    

End Sub
Sub GRABAR3(CRCC, cuenta)
     sqlconta.audit = False
     
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "cuenta"
    campos(3, 0) = ""
    campos(0, 1) = CRCC
    campos(1, 1) = Mid(fechasistema, 7, 4)
    campos(2, 1) = cuenta

    campos(0, 2) = "saldoscentrosdecosto"
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    sqlconta.audit = True
    

End Sub


Sub ELIMINAR()
    
    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub






Private Sub graba_Click()
grabar
End Sub

Private Sub List1_Click()
If Val(Mid(List1.text, 4, 2)) <> 0 Then

dato1.Enabled = True
dato2.Enabled = True

dato1.text = Mid(List1.text, 3, 10)
dato2.SetFocus
End If



End Sub

Private Sub List1_DblClick()
If Mid(List1.text, 17, 2) <> "  " Then
dato1.text = Mid(List1.text, 1, 5)
dato2.Enabled = True

dato2.SetFocus
End If


End Sub

Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub


Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then retorno
'If command = "modifica" And PERMISOPROGRAMA(3) = "N" Then Call NOPERMISO(3)
If command = "modifica" Then

 If Verifica_Permiso(Me.Caption, "modifica") = True Then
    modifica
 Else
     MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
 End If
    
End If
'If command = "elimina" And PERMISOPROGRAMA(4) = "N" Then Call NOPERMISO(4)

If command = "elimina" Then
 If Verifica_Permiso(Me.Caption, "elimina") = True Then
    ELIMINA
 Else
     MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
 End If
End If
If command = "imprime" Then imprimir
If command = "movimientos" Then movimientos

End Sub
Sub ELIMINA()

If saldoglobal = 0 Then
If Verifica_Permiso(Me.Caption, "elimina") = True Then
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus

Else
 MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
End If
Else
MsgBox ("CUENTA CON MOVIMIENTOS IMPOSIBLE ELIMINAR ")
End If
End Sub
Sub modifica()
If saldoglobal = 0 Then

disponible (True)
habilita (False)
dato1.Enabled = False
dato2.Enabled = False

MODIFI = 1
graba.Visible = True
Else
MsgBox ("CUENTA CON MOVIMIENTOS IMPOSIBLE MODIFICAR ")
End If

End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
graba.Visible = False

dato1.Enabled = True
dato1.SetFocus
MODIFI = 0



End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    
    


graba.Visible = False

End Sub

Sub imprimir()
    
End Sub
Sub grilla()

End Sub
Sub CABEZA()
    
End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo,ctacte,crcc "
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

Sub DATOSSALDOS()

End Sub



Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub



Private Sub opciones_GotFocus()
graba.Visible = False

MANUAL.SetFocus

End Sub


Sub leeplandecuenta()
leecrcc
End Sub
Sub leecrcc()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta.sii_1846 "
       
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
         List1.Clear
         
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        If Mid(resultados2(0), 6, 2) = "00" Then
        List1.AddItem "": List1.AddItem resultados2(0) + " " + resultados2(1): List1.AddItem ""
        Else
        List1.AddItem "   " + resultados2(0) + " " + resultados2(1)
        
        
        End If
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    
    
    

End Sub

Sub movimientos()
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub
