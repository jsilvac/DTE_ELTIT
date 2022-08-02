VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form seguri01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Cuentas del Mayor"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6000
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
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7935
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "nombre"
         Top             =   1320
         Width           =   6015
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   4
         Top             =   840
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "codigo"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1095
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "\\eltitxp\contabilidad 2005\barra_opciones.swf"
      Src             =   "\\eltitxp\contabilidad 2005\barra_opciones.swf"
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
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   3615
      Left            =   360
      Top             =   480
      Width           =   7935
   End
End
Attribute VB_Name = "seguri01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub Command1_Click()
IMPRIMIR
End Sub


Private Sub DATO1_GotFocus()
grillasaldos
Call cargatexto(dato1)
End Sub
Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub
Private Sub dato3_GotFocus()
If modifi = 0 Then leer
Call cargatexto(dato3)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo NO:
    If KeyCode = vbKeyF2 Then Call ayudacentrocosto(dato3)
    Call flechas(dato1, dato2, KeyCode)
NO:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato3, KeyCode)
End Sub
Private Sub DESDE1_GotFocus()
Call cargatexto(DESDE1)
End Sub
Private Sub DESDE2_GotFocus()
Call cargatexto(DESDE2)
End Sub
Private Sub DESDE3_GotFocus()
Call cargatexto(DESDE3)
End Sub
Private Sub HABER1_GotFocus()
Call cargatexto(HABER1)
End Sub
Private Sub HABER2_GotFocus()
Call cargatexto(HABER2)
End Sub
Private Sub HABER3_GotFocus()
Call cargatexto(HABER3)
End Sub


Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then Call ceros(DESDE1): Call Pregunta(DESDE1, DESDE2)
End Sub
Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then Call ceros(DESDE2): Call Pregunta(DESDE1, DESDE3)
End Sub
Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DESDE3): Call Pregunta(DESDE2, HASTA1)
End Sub
Private Sub HASTA1_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(HASTA1): Call Pregunta(DESDE3, HASTA2)
End Sub
Private Sub HASTA2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(HASTA2): Call Pregunta(HASTA1, HASTA3)
End Sub
Private Sub HASTA3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(HASTA2): Call Pregunta(HASTA2, DESDE1)
End Sub


Private Sub Form_Load()
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3

Call RECUPERAFECHA


End Sub

Private Sub DATO1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then GRABAR: retorno
End Sub


Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then GoTo NO:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
NO:
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo>" + "'" + dato1.text + dato2.text + "' order by codigo"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then GoTo NO:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
NO:
   
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo<" + "'" + dato1.text + dato2.text + "' order by codigo"

    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
   If SQLUTIL.ESTADO = 4 Then GoTo NO:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
    
NO:
   
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = Mid(SQLUTIL.datos(0, 3), 1, 2)
    dato2.text = Mid(SQLUTIL.datos(0, 3), 3, 2)
    dato3.text = SQLUTIL.datos(1, 3)
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudacentrocosto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "no"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "centrosdecosto", PIVOTE, campos, cfijo, largo, 2)
    If Val(PIVOTE.text) = 0 Then dato1.SetFocus: GoTo NO
    dato2.Enabled = True
    dato1.text = Mid(PIVOTE.text, 1, 2)
    dato2.text = Mid(PIVOTE.text, 3, 2)
    
    caja.Enabled = True
    caja.SetFocus
NO:
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub GRABAR()
    
    
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = ""
    campos(0, 1) = dato1.text + dato2.text
    campos(1, 1) = dato2.text
    campos(0, 2) = "centrosdecosto"
    If modifi = 1 Then condicion = "codigo=" + "'" + dato1.text + dato2.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If modifi = 0 Then GRABAR2
    modifi = 0
End Sub
Sub GRABAR2()
     
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
    campos(25, 0) = "haber10"
    campos(26, 0) = "haber11"
    campos(27, 0) = "haber12"
    
    campos(28, 0) = ""
    campos(0, 1) = dato1.text + dato2.text
    campos(1, 1) = año

    For K = 2 To 27
    campos(K, 1) = "0"
    Next K
    campos(0, 2) = "saldoscentrosdecosto"
    op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    

End Sub

Sub ELIMINAR()
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub

Private Sub LISTATIPOS_DblClick()
dato5.text = Mid(LISTATIPOS.text, 2, 1)
NOMBRETIPO.Caption = Mid(LISTATIPOS.text, 4, 20)
dato6.SetFocus

End Sub


Private Sub LISTATIPOS2_Click()
dato8.text = Mid(LISTATIPOS2.text, 2, 1)
NOMBRETIPO2.Caption = Mid(LISTATIPOS2.text, 4, 20)


End Sub

Private Sub NOIMPRIME_Click()
FECHAS.Visible = False

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then retorno
If command = "modifica" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.Enabled = False:  dato3.SetFocus: modifi = 1
If command = "elimina" Then disponible (True): habilita (False): ELIMINAR: limpia: opciones.Visible = False: dato1.SetFocus
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "imprime" Then IMPRIMIR
If command = "movimientos" Then CARTOLA
End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
End Sub

Sub IMPRIMIR()
    informes.info.Clear
    largopagina = 65
    tituloinforme = "plan de cuentas"
    titu(1) = "CODIGO"
    titu(2) = "NOMBRE DE CUENTA"
    titu(3) = "TIPO"
    titu(4) = "CTACTE"
    titu(5) = "NOMBRE CTACTE"
    titu(6) = "AUXILIAR"
    lineas = 70
    Consulta_Informe
    informes.Show
    
End Sub
Sub grilla()
    palabra = ""
      
    For K = 1 To cancolu
    If tipodato(K) = "s" Or tipodato(K) = "S" Then dato(K) = dato(K) & String(colu(K) - Len(dato(K)), 32)
    If tipodato(K) = "n" Or tipodato(K) = "N" Then dato(K) = String(colu(K) - Len(dato(K)), 32) & dato(K)
    palabra = palabra & dato(K)
    Next K
    If lineas > largopagina Then Call cabeza
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    If Mid(dato(1), 7, 4) <> "0000" Then informes.info.AddItem ("    " + palabra)
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (Mid(palabra, 1, 40))
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    lineas = lineas + 1

End Sub
Sub cabeza()
    informes.info.AddItem ("")
    informes.info.AddItem ("")
    pagina = pagina + 1
    


    informes.info.AddItem ("NOMBRE EMPRESA          " + tituloinforme + "                                   PAGINA " + Str$(pagina))
    informes.info.AddItem ("DIRECCION EMPRESA                                                                              " + Mid(Date$, 4, 2) + "-" + Mid(Date$, 1, 2) + "-" + Mid(Date$, 7, 4))
    informes.info.AddItem ("RUT EMPRESA                                                                                    " + Time$)
    informes.info.AddItem ("                                " + tituloinforme)
    informes.info.AddItem String(132, "_")
    TITULOS = ""
    For K = 1 To cancolu
    titu(K) = titu(K) & String(colu(K) - Len(titu(K)), 32)
    TITULOS = TITULOS & titu(K)
    Next K
    informes.info.AddItem (TITULOS)
    informes.info.AddItem String(132, "_")

lineas = 8

End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4): colu(1) = 15: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                dato(3) = resultados(2) + " " + DOCU$(Val(resultados(2)))
                dato(4) = resultados(3)
                dato(5) = resultados(4)
                dato(6) = resultados(5) + " " + DOCU2$(Val(resultados(5)))
                colu(3) = 10: tipodato(3) = "s"
                colu(4) = 10: tipodato(4) = "s"
                colu(5) = 20: tipodato(5) = "s"
                colu(6) = 10: tipodato(6) = "s"
                 cancolu = 6
                grilla
                resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    End With

End Sub

Sub DATOSSALDOS()

LEERSALDOS
SUMADOR = Val(SQLUTIL.datos(2, 3)) - Val(SQLUTIL.datos(3, 3))
SALDOS.TextMatrix(1, 1) = Format(SQLUTIL.datos(2, 3), "###,###,##0")
SALDOS.TextMatrix(1, 2) = Format(SQLUTIL.datos(3, 3), "###,###,##0")
SALDOS.TextMatrix(1, 3) = Format(SUMADOR, "###,###,##0")
For K = 4 To 15
SALDOS.TextMatrix(K - 2, 1) = Format(SQLUTIL.datos(K, 3), "###,###,##0")
SALDOS.TextMatrix(K - 2, 2) = Format(SQLUTIL.datos(K + 12, 3), "###,###,##0")
SUMADOR = SUMADOR + Val(SQLUTIL.datos(K, 3)) - Val(SQLUTIL.datos(K + 12, 3))
SALDOS.TextMatrix(K - 2, 3) = Format(SUMADOR, "###,###,##0")
Next K

End Sub
Sub grillasaldos()
SALDOS.Cols = 4
SALDOS.Rows = 14
SALDOS.ColWidth(0) = 120 * 12
SALDOS.ColWidth(1) = 120 * 8
SALDOS.ColWidth(2) = 120 * 8
SALDOS.ColWidth(3) = 120 * 8
SALDOS.TextMatrix(0, 0) = "MESES   "
SALDOS.TextMatrix(0, 1) = "DEBE    "
SALDOS.TextMatrix(0, 2) = "HABER   "
SALDOS.TextMatrix(0, 3) = "SALDO   "
SALDOS.TextMatrix(1, 0) = "AÑO ANTERIOR"
SALDOS.TextMatrix(2, 0) = "ENERO"
SALDOS.TextMatrix(3, 0) = "FEBRERO"
SALDOS.TextMatrix(4, 0) = "MARZO"
SALDOS.TextMatrix(5, 0) = "ABRIL"
SALDOS.TextMatrix(6, 0) = "MAYO"
SALDOS.TextMatrix(7, 0) = "JUNIO"
SALDOS.TextMatrix(8, 0) = "JULIO"
SALDOS.TextMatrix(9, 0) = "AGOSTO"
SALDOS.TextMatrix(10, 0) = "SEPTIEMBRE"
SALDOS.TextMatrix(11, 0) = "OCTUBRE"
SALDOS.TextMatrix(12, 0) = "NOVIEMBRE "
SALDOS.TextMatrix(13, 0) = "DICIEMBRE "
For K = 1 To 13
SALDOS.TextMatrix(K, 1) = "0"
SALDOS.TextMatrix(K, 2) = "0"
SALDOS.TextMatrix(K, 3) = "0"
Next K
End Sub

Sub LEERSALDOS()
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
    campos(25, 0) = "haber10"
    campos(26, 0) = "haber11"
    campos(27, 0) = "haber12"
    campos(28, 0) = ""
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "'"
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    grillasaldos
End Sub

Sub CARTOLA()
FECHAS.Visible = True
DESDE1.text = "01"
DESDE2.text = mes
DESDE3.text = año
HASTA1.text = "31"
HASTA2.text = mes
HASTA3.text = año

DESDE1.SetFocus
End Sub

Sub movimientos()



cartolas.Caption = "CARTOLA CUENTA DEL MAYOR"
cartolas.titulocartola = dato1.text + "." + dato2.text + "." + dato3.text + "   " + dato4.text

cartolas.grilla.Cols = 15
cartolas.grilla.Rows = 2
cartolas.grilla.ColWidth(0) = 120 * 8
cartolas.grilla.ColWidth(1) = 120 * 3
cartolas.grilla.ColWidth(2) = 120 * 10
cartolas.grilla.ColWidth(3) = 120 * 4
cartolas.grilla.ColWidth(4) = 120 * 8
cartolas.grilla.ColWidth(5) = 120 * 25
cartolas.grilla.ColWidth(6) = 120 * 3
cartolas.grilla.ColWidth(7) = 120 * 10
cartolas.grilla.ColWidth(8) = 120 * 8
cartolas.grilla.ColWidth(9) = 120 * 10
cartolas.grilla.ColWidth(10) = 120 * 10
cartolas.grilla.ColWidth(11) = 120 * 10
cartolas.Show


Rem TITULOS
cartolas.grilla.TextMatrix(0, 0) = "FECHA"
cartolas.grilla.TextMatrix(0, 1) = "TIPO"
cartolas.grilla.TextMatrix(0, 2) = "NUMERO"
cartolas.grilla.TextMatrix(0, 3) = "LINEA"
cartolas.grilla.TextMatrix(0, 4) = "CUENTA"
cartolas.grilla.TextMatrix(0, 5) = "GLOSA"
cartolas.grilla.TextMatrix(0, 6) = "TD"
cartolas.grilla.TextMatrix(0, 7) = "NUMERO"
cartolas.grilla.TextMatrix(0, 8) = "VENCIMIENTO"
cartolas.grilla.TextMatrix(0, 9) = "DEBE"
cartolas.grilla.TextMatrix(0, 10) = "HABER"
cartolas.grilla.TextMatrix(0, 11) = "SALDO"
LEERMOVIMIENTOS
GoTo NO:
SUMADOR = Val(SQLUTIL.datos(2, 3)) - Val(SQLUTIL.datos(3, 3))
SALDOS.TextMatrix(1, 1) = Format(SQLUTIL.datos(2, 3), "###,###,##0")
SALDOS.TextMatrix(1, 2) = Format(SQLUTIL.datos(3, 3), "###,###,##0")
SALDOS.TextMatrix(1, 3) = Format(SUMADOR, "###,###,##0")
For K = 4 To 15
SALDOS.TextMatrix(K - 2, 1) = Format(SQLUTIL.datos(K, 3), "###,###,##0")
SALDOS.TextMatrix(K - 2, 2) = Format(SQLUTIL.datos(K + 12, 3), "###,###,##0")
SUMADOR = SUMADOR + Val(SQLUTIL.datos(K, 3)) - Val(SQLUTIL.datos(K + 12, 3))
SALDOS.TextMatrix(K - 2, 3) = Format(SUMADOR, "###,###,##0")
Next K
NO:
End Sub

Sub LEERMOVIMIENTOS()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh "
        cSql.SQL = cSql.SQL + "FROM movimientoscontables"
        PIVOTE.text = dato1.text + dato2.text + dato3.text
        cSql.SQL = cSql.SQL + " where codigocuenta = " + "'" + PIVOTE.text + "' AND FECHA>=" + "'" + DESDE3.text + DESDE2.text + DESDE1.text + "' and fecha<=" + "'" + HASTA3.text + HASTA2.text + HASTA1.text + "'  ORDER BY FECHA "
    
        cSql.Execute
        linea = 0: SUMADOR = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                linea = linea + 1
                cartolas.grilla.Rows = linea + 2
                For K = 0 To 8
                cartolas.grilla.TextMatrix(linea, K) = resultados(K)
                Next K
                If resultados(10) = "D" Then cartolas.grilla.TextMatrix(linea, 9) = Format(resultados(K), "###,###,##0"): SUMADOR = SUMADOR + Val(resultados(K))
                If resultados(10) = "H" Then cartolas.grilla.TextMatrix(linea, 10) = Format(resultados(K), "###,###,##0"): SUMADOR = SUMADOR - Val(resultados(K))
                cartolas.grilla.TextMatrix(linea, 11) = Format(SUMADOR, "###,###,##0"):
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing

        End If
    End With

End Sub

Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub

Private Sub SIIMPRIME_Click()
FECHAS.Visible = False

movimientos

End Sub

