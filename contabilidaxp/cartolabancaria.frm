VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form banco02 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6750
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5175
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
      BackColor       =   16744576
      Caption         =   "CARTOLAS DE BANCO"
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   " Consolidado"
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   3720
         Width           =   1575
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1215
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2143
         BackColor       =   49344
         Caption         =   "CODIGO DE BANCO"
         CaptionEstilo3D =   1
         BackColor       =   49344
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
         Alignment       =   1
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
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "codigo"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox dato2 
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
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   1
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox dato3 
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
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   2
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   60
            TabIndex        =   14
            Top             =   720
            Width           =   4695
         End
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1815
         Left            =   1080
         TabIndex        =   6
         Top             =   1800
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3201
         BackColor       =   14737632
         Caption         =   "Rangos de Fecha"
         CaptionEstilo3D =   1
         BackColor       =   14737632
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
         Begin CoolButtons.cool_Button command8 
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            SkinId          =   "13"
            Caption         =   "Cambia Fecha"
         End
         Begin VB.Label hastafecha 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   11
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label desdefecha 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
      End
      Begin CoolButtons.cool_Button GENERA 
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   3720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "GENERA INFORME"
      End
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   3480
         TabIndex        =   15
         Top             =   4440
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
            TabIndex        =   17
            Top             =   280
            Width           =   1335
         End
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   280
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "banco02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private FORMATOGRILLA(20, 20)
    Private lin As Double
    Private saldo As Double
    Private dedonde As Integer
    Private tipoctacte As String

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub command8_Click()
    Call retornofecha(desdefecha, hastafecha)
End Sub

'****************************************************************************
'GOTFOCUS
'****************************************************************************
    Private Sub dato1_GotFocus()
        Call cargatexto(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
'****************************************************************************
'GOTFOCUS
'****************************************************************************

'****************************************************************************
'KEYDOWN
'****************************************************************************
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 38 Then Unload Me: GoTo no:
        If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
        Call flechas(dato1, dato2, KeyCode)
no:
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato1, dato3, KeyCode)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato2, dato3, KeyCode)
    End Sub
'****************************************************************************
'KEYDOWN
'****************************************************************************

'****************************************************************************
'KEYPRESS
'****************************************************************************
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            lblBanco.Caption = leerNombreCuentaMayor(dato1.text & dato2.text & dato3.text, 3)
            If lblBanco.Caption <> "" Then
                SendKeys "{Tab}"
            dato1.SetFocus
            End If
        End If
    End Sub
'****************************************************************************
'KEYPRESS
'****************************************************************************

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

Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "' AND banco='1'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    basebus = clientesistema + "conta" + empresaactiva
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

Sub LEERMOVIMIENTOS(infogrilla As grillainformes, cuenta, NOMBRE)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
    fecha2 = Mid(hastafecha.Caption, 7, 4) + "-" + Mid(hastafecha.Caption, 4, 2) + "-" + Mid(hastafecha.Caption, 1, 2)
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
    If dedonde = 1 Then csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "'"
    If dedonde = 2 Then csql.sql = csql.sql + "FROM movimientoscontables where tipoctacte='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "'"
    If dedonde = 3 Then csql.sql = csql.sql + "FROM movimientoscontables where centrocosto='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "'"
    csql.sql = csql.sql + "order by codigocuenta,fecha,tipo,numero,linea"
    csql.Execute
        
    If dedonde <> 2 Then Call DATOSSALDOS(cuenta)
       
    For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = False
    Next k
    Rem If saldo <> 0 Then
    lin = lin + 1
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
                
    infogrilla.Grid1.Range(lin, 1, lin, 6).Merge
        
    infogrilla.Grid1.Cell(lin, 1).CellType = cellTextBox
        
    infogrilla.Grid1.Cell(lin, 10).CellType = cellTextBox
        
    infogrilla.Grid1.Cell(lin, 1).text = Mid(cuenta, 1, 2) + "." + Mid(cuenta, 3, 2) + "." + Mid(cuenta, 5, 4) + " " + NOMBRE
        
    If dedonde = 2 Then infogrilla.Grid1.Cell(lin, 7).text = tipoctacte
    infogrilla.Grid1.Cell(lin, 10).text = "SALDO-->"
    infogrilla.Grid1.Cell(lin, 13).text = saldo
    infogrilla.Grid1.Range(lin, 0, lin, infogrilla.Grid1.Cols - 1).FontBold = False
    infogrilla.Grid1.Range(lin, 0, lin, infogrilla.Grid1.Cols - 1).FontUnderline = True
    Rem End If
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            lin = lin + 1
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            For k = 0 To 9
                infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
            Next k
            If resultados(11) = "D" Then infogrilla.Grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10): saldo = saldo + resultados(10)
            If resultados(11) = "H" Then infogrilla.Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10): saldo = saldo - resultados(10)
            infogrilla.Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
            infogrilla.Grid1.Cell(lin, 13).text = saldo
            resultados.MoveNext
        Wend
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        'Call totalcomprobante(infogrilla, lin)
        resultados.Close
        Set resultados = Nothing
    End If
    For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = True
    Next k
End Sub

Sub LEERCHEQUES(infogrilla As grillainformes, cuenta, NOMBRE)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    LINEA = lin + 1
    
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT cuenta,numero,emision,monto,vencimiento,tipocomprobante,numerocomprobante,giradoa,fechacobro,ubicacion "
    csql.sql = csql.sql + "FROM chequesdocumento where cuenta='" + cuenta + "' and (cobrado='0' or fechacobro>'" + Format(hastafecha.Caption, "yyyy-mm-dd") + "') and emision<='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "'  "
    
    csql.sql = csql.sql + "order by cuenta,vencimiento"
    csql.Execute
    infogrilla.AutoRedraw = False
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
    infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 2, 6).text = "CHEQUES PENDIENTES DE COBRO"
    infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 2, 6).Font.Bold = True
    infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 2, 6).Alignment = cellCenterCenter
    
    
    
    
    
    
    infogrilla.Grid1.Range(infogrilla.Grid1.Rows - 1, 1, infogrilla.Grid1.Rows - 1, infogrilla.Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
    infogrilla.Grid1.Range(infogrilla.Grid1.Rows - 2, 1, infogrilla.Grid1.Rows - 2, infogrilla.Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
    
    
    
    For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = False
    Next k
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            LINEA = LINEA + 1
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            infogrilla.Grid1.Cell(LINEA, 1).text = resultados(2)
            infogrilla.Grid1.Cell(LINEA, 2).text = resultados(5)
            infogrilla.Grid1.Cell(LINEA, 3).text = resultados(6)
            infogrilla.Grid1.Cell(LINEA, 5).text = resultados(0)
            infogrilla.Grid1.Cell(LINEA, 5).text = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4)
            
            infogrilla.Grid1.Cell(LINEA, 6).text = resultados(7)
            infogrilla.Grid1.Cell(LINEA, 7).text = "CH"
            infogrilla.Grid1.Cell(LINEA, 8).text = resultados(1)
            infogrilla.Grid1.Cell(LINEA, 9).text = resultados(2)
            infogrilla.Grid1.Cell(LINEA, 10).text = resultados(4)
            infogrilla.Grid1.Cell(LINEA, 12).text = resultados(3)
            saldo = saldo + resultados(3)
            infogrilla.Grid1.Cell(LINEA, 13).text = saldo
            resultados.MoveNext
        Wend
    
    End If
infogrilla.AutoRedraw = True

End Sub

Sub totalcomprobante(infogrilla As grillainformes, row)
    infogrilla.Grid1.Range(row, 11, row, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(row, 1, row, 12).FontSize = 8
    
    
    infogrilla.Grid1.Range(row, 1, row, 12).FontBold = False
    infogrilla.Grid1.Range(row, 1, row, 12).FontUnderline = True
    infogrilla.Grid1.Cell(row, 10).CellType = cellTextBox
    infogrilla.Grid1.Cell(row, 10).text = "TOTAL "
    infogrilla.Grid1.Cell(row, 11).text = anted
    infogrilla.Grid1.Cell(row, 12).text = anteh
    lin = lin + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
        
    anted = 0: anteh = 0
End Sub
    
Sub CARGAGRILLA(infogrilla As grillainformes)
    Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
    
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "LINEA"
    FORMATOGRILLA(1, 5) = "CUENTA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "TP"
    FORMATOGRILLA(1, 8) = "NUMERO"
    FORMATOGRILLA(1, 9) = "EMISION"
    FORMATOGRILLA(1, 10) = "VENCIMIENTO"
    FORMATOGRILLA(1, 11) = "DEBE"
    FORMATOGRILLA(1, 12) = "HABER"
    FORMATOGRILLA(1, 13) = "SALDO"
    FORMATOGRILLA(1, 14) = "NOMBRE CUENTA"
    FORMATOGRILLA(1, 15) = "CUENTA CORRIENTE"
    FORMATOGRILLA(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "9"
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "28"
    FORMATOGRILLA(2, 7) = "3"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "11"
    FORMATOGRILLA(2, 11) = "11"
    FORMATOGRILLA(2, 12) = "11"
    FORMATOGRILLA(2, 13) = "11"
    FORMATOGRILLA(2, 14) = "30"
    FORMATOGRILLA(2, 15) = "30"

    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "D"
    FORMATOGRILLA(3, 10) = "D"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 11) = "###,###,###,##0"
    FORMATOGRILLA(4, 12) = "###,###,###,##0"
    FORMATOGRILLA(4, 13) = "###,###,###,##0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    
    infogrilla.Grid1.Cols = 15
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub LEERSALDOS(cuenta)
    Dim resultados3 As rdoResultset
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    
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
    
    If dedonde = 1 Then condicion = "codigo=" + "'" + cuenta + "' and año='" + Mid(desdefecha.Caption, 7, 4) + "' order by codigo"
    If dedonde = 3 Then condicion = "codigo=" + "'" + cuenta + "' and año='" + Mid(desdefecha.Caption, 7, 4) + "' order by codigo"
    
    If dedonde = 1 Then campos(0, 2) = "saldosdelmayor"
    If dedonde = 3 Then campos(0, 2) = "saldoscentrosdecosto"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    'If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(2, 3)) - Val(sqlconta.response(3, 3))
    saldo = sumador
    Rem acumula fecha
    fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
    Set cSql3.ActiveConnection = contadb
    cSql3.sql = "SELECT sum(monto),dh "
    cSql3.sql = cSql3.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and fecha<'" + fecha1 + "' and fecha>='" + Format(fechasistema, "yyyy") + "-01-01" + "' "
        
    cSql3.sql = cSql3.sql + "group by dh"
    cSql3.Execute
        
    If cSql3.RowsAffected > 0 Then
        Set resultados3 = cSql3.OpenResultset
        While Not resultados3.EOF
            If resultados3(1) = "D" Then saldo = saldo + resultados3(0)
            If resultados3(1) = "H" Then saldo = saldo - resultados3(0)
            resultados3.MoveNext
        Wend
        resultados3.Close
        Set resultados3 = Nothing
    End If
End Sub

Sub DATOSSALDOS(cuenta)
    Call LEERSALDOS(cuenta)
End Sub

Private Sub Form_Load()
    Call CENTRAR(Me)
    fechas.Visible = True
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    desdefecha.Caption = fechasistema
    hastafecha.Caption = fechasistema
    lin = 0
End Sub

Private Sub GENERA_Click()
    lin = 0
    Dim infogrilla As grillainformes
    Set infogrilla = New grillainformes
    Call CARGAGRILLA(infogrilla)
    infogrilla.Caption = "CARTOLA BANCARIA"
    dedonde = 1
    Call LEERMOVIMIENTOS(infogrilla, dato1.text + dato2.text + dato3.text, lblBanco)
    Call LEERCHEQUES(infogrilla, dato1.text + dato2.text + dato3.text, lblBanco)
    grillainformes.Tag = "banco02"
    infogrilla.Grid1.Visible = True
    
    infogrilla.Show
End Sub
