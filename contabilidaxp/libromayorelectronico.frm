VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "clbutn.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form electronico01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LIBRO MAYOR ANALITICO"
   ClientHeight    =   5805
   ClientLeft      =   240
   ClientTop       =   1290
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin CoolButtons.cool_Button GENERA 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "GENERA INFORME"
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Formato Envio SII"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   4200
      Width           =   2895
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   855
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      BackColor       =   16761024
      Caption         =   "TIPO DE IMPRESION"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.TextBox FOLIO 
         Height          =   285
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton timbrado 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprime Timbrado"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton original 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprime Original"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5953
      BackColor       =   16761024
      Caption         =   "Configuracion"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   735
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1296
         BackColor       =   16744576
         Caption         =   "EMPRESA"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox DATO1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Text            =   "01"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label empresanombre 
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
            Height          =   375
            Left            =   840
            TabIndex        =   4
            Top             =   360
            Width           =   3255
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   855
         Left            =   1080
         TabIndex        =   5
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "MES"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOMES 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   3855
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   855
         Left            =   1080
         TabIndex        =   7
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "AÑO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOAÑO 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   3855
         End
      End
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   3600
      TabIndex        =   15
      Top             =   5040
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "electronico01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormatoGrilla(10, 20) As String
Private sumas(10) As Double
Private suma(10) As Double
Private sumas2(10) As Double
Private sumas3(10) As Double
Private montos(5) As Double
Private lin As Double





Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
    
End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub


Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    COMBOMES.SetFocus
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub

Private Sub Form_Load()
CENTRAR Me

 Call Conectar_BD
 Call Conectarconta(servidor, clientesistema + "conta", Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
dato1.text = empresaactiva
empresanombre.Caption = nombreempresa
original.Value = True



End Sub


Sub ACEPTA(opcion)
Dim TIMBRA As String

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes
If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"
If opcion = 1 Then infogrilla.Caption = "LIBRO MAYOR ANALITICO": grillainformes.Tag = "auxiliar03" & TIMBRA & folio.text
infogrilla.CABEZA.Caption = "LIBRO MAYOR ANALITICO de " & COMBOMES.text & " del " & año + " de la empresa " + empresanombre.Caption
infogrilla.titulofinal.Caption = "LIBRO MAYOR ANALITICO DE " & COMBOMES.text & " DEL " & año
lin = 0
If Check1.Value = 0 Then
Call CARGAGRILLA(infogrilla)
Else
Call CARGAGRILLA2(infogrilla)

End If

Call leecuentas(infogrilla)

infogrilla.Visible = True

infogrilla.Show

End Sub


Private Sub GENERA_Click()
Call Conectartemporal(servidor, clientesistema + "conta" + dato1.text, Usuario, password)
año = COMBOAÑO.text
mes = COMBOMES.ListIndex + 1
If Val(mes) < 10 Then mes = "0" + Mid(Str(mes), 2, 1) Else mes = Mid(Str(mes), 2, 2)
Call ACEPTA(1)
Unload Me

End Sub
Sub LEERMOVIMIENTOS(cuenta, NOMBRE, infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    
        Set csql.ActiveConnection = temporal
        
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,rutctacte "
        csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and mes='" + mes + "' and año='" + año + "' "
        csql.sql = csql.sql + "order by fecha"
        csql.Execute
       
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        Call DATOSSALDOS(cuenta)
        For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = False
        Next k
        
        infogrilla.Grid1.Range(lin, 1, lin, 12).FontBold = True
        infogrilla.Grid1.Range(lin, 1, lin, 12).FontUnderline = True
        
        
        
        
        infogrilla.Grid1.Range(lin, 1, lin, 6).Merge
        
        infogrilla.Grid1.Cell(lin, 1).CellType = cellTextBox
        
        infogrilla.Grid1.Cell(lin, 10).CellType = cellTextBox
        
        infogrilla.Grid1.Cell(lin, 1).text = NOMBRE
        infogrilla.Grid1.Cell(lin, 10).text = "SALDO-->"
        
        infogrilla.Grid1.Cell(lin, 13).text = saldo

        
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
             infogrilla.Grid1.Cell(lin, 13).text = saldo
            If Check1.Value = 1 Then
            infogrilla.Grid1.Cell(lin, 14).text = resultados(12)
        End If
             resultados.MoveNext
           
         Wend
          lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
         
         'Call totalcomprobante(lin, infogrilla)
          resultados.Close
            Set resultados = Nothing

        End If

End Sub

Sub totalcomprobante(Row, infogrilla As grillainformes)
    'infogrilla.Grid1.Range(Row, 1, Row, 12).FontBold = True
    infogrilla.Grid1.Range(Row, 1, Row, 12).FontUnderline = True
        
    
    infogrilla.Grid1.Range(Row, 11, Row, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Cell(Row, 10).CellType = cellTextBox
    infogrilla.Grid1.Cell(Row, 10).text = "TOTAL "
    infogrilla.Grid1.Cell(Row, 11).text = anted
    infogrilla.Grid1.Cell(Row, 12).text = anteh
    lin = lin + 2
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
        
        anted = 0: anteh = 0: saldo = 0
    End Sub
    





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7
    
    
    FormatoGrilla(1, 1) = "FECHA"
    FormatoGrilla(1, 2) = "TP"
    FormatoGrilla(1, 3) = "NUMERO"
    FormatoGrilla(1, 4) = "LINEA"
    FormatoGrilla(1, 5) = "CUENTA"
    FormatoGrilla(1, 6) = "GLOSA"
    FormatoGrilla(1, 7) = "TP"
    FormatoGrilla(1, 8) = "NUMERO"
    FormatoGrilla(1, 9) = "EMISION"
    FormatoGrilla(1, 10) = "VENCIMIENTO"
    FormatoGrilla(1, 11) = "DEBE"
    FormatoGrilla(1, 12) = "HABER"
    FormatoGrilla(1, 13) = "SALDO"
    FormatoGrilla(1, 14) = "NOMBRE CUENTA"
    FormatoGrilla(1, 15) = "CUENTA CORRIENTE"
    FormatoGrilla(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FormatoGrilla(2, 1) = "10"
    FormatoGrilla(2, 3) = "10"
    FormatoGrilla(2, 4) = "4"
    FormatoGrilla(2, 5) = "0"
    FormatoGrilla(2, 6) = "30"
    FormatoGrilla(2, 7) = "3"
    FormatoGrilla(2, 8) = "10"
    FormatoGrilla(2, 9) = "0"
    FormatoGrilla(2, 10) = "10"
    FormatoGrilla(2, 11) = "11"
    FormatoGrilla(2, 12) = "11"
    FormatoGrilla(2, 13) = "12"
    FormatoGrilla(2, 14) = "30"
    FormatoGrilla(2, 15) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FormatoGrilla(3, 1) = "D"
    FormatoGrilla(3, 2) = "S"
    FormatoGrilla(3, 3) = "S"
    FormatoGrilla(3, 4) = "S"
    FormatoGrilla(3, 5) = "S"
    FormatoGrilla(3, 6) = "S"
    FormatoGrilla(3, 7) = "S"
    FormatoGrilla(3, 8) = "S"
    FormatoGrilla(3, 9) = "D"
    FormatoGrilla(3, 10) = "D"
    FormatoGrilla(3, 11) = "N"
    FormatoGrilla(3, 12) = "N"
    FormatoGrilla(3, 13) = "N"
    FormatoGrilla(3, 14) = "S"
    FormatoGrilla(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    FormatoGrilla(4, 11) = "###,###,###,##0"
    FormatoGrilla(4, 12) = "###,###,###,##0"
    FormatoGrilla(4, 13) = "###,###,###,##0"
    Rem LOCCKED
    FormatoGrilla(5, 1) = "TRUE"
    FormatoGrilla(5, 2) = "TRUE"
    FormatoGrilla(5, 3) = "TRUE"
    FormatoGrilla(5, 4) = "TRUE"
    FormatoGrilla(5, 5) = "TRUE"
    FormatoGrilla(5, 6) = "TRUE"
    FormatoGrilla(5, 7) = "TRUE"
    FormatoGrilla(5, 8) = "TRUE"
    FormatoGrilla(5, 9) = "TRUE"
    FormatoGrilla(5, 10) = "TRUE"
    FormatoGrilla(5, 11) = "TRUE"
    FormatoGrilla(5, 12) = "TRUE"
    FormatoGrilla(5, 13) = "TRUE"
    FormatoGrilla(5, 14) = "TRUE"
    FormatoGrilla(5, 15) = "TRUE"
    
    infogrilla.Grid1.Cols = 14
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
        
        infogrilla.Grid1.Cell(0, k).text = FormatoGrilla(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FormatoGrilla(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FormatoGrilla(2, k))
        infogrilla.Grid1.Column(k).FormatString = FormatoGrilla(4, k)
        infogrilla.Grid1.Column(k).Locked = FormatoGrilla(5, k)
        If FormatoGrilla(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FormatoGrilla(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub CARGAGRILLA2(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7
    
    
    FormatoGrilla(1, 1) = "FECHA"
    FormatoGrilla(1, 2) = "TP"
    FormatoGrilla(1, 3) = "NUMERO"
    FormatoGrilla(1, 4) = "LINEA"
    FormatoGrilla(1, 5) = "CUENTA"
    FormatoGrilla(1, 6) = "GLOSA"
    FormatoGrilla(1, 7) = "TP"
    FormatoGrilla(1, 8) = "NUMERO"
    FormatoGrilla(1, 9) = "EMISION"
    FormatoGrilla(1, 10) = "VENCIMIENTO"
    FormatoGrilla(1, 11) = "DEBE"
    FormatoGrilla(1, 12) = "HABER"
    FormatoGrilla(1, 13) = "SALDO"
    FormatoGrilla(1, 14) = "RUT"
    FormatoGrilla(1, 15) = "CUENTA CORRIENTE"
    FormatoGrilla(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FormatoGrilla(2, 1) = "10"
    FormatoGrilla(2, 3) = "10"
    FormatoGrilla(2, 4) = "4"
    FormatoGrilla(2, 5) = "10"
    FormatoGrilla(2, 6) = "30"
    FormatoGrilla(2, 7) = "3"
    FormatoGrilla(2, 8) = "10"
    FormatoGrilla(2, 9) = "0"
    FormatoGrilla(2, 10) = "10"
    FormatoGrilla(2, 11) = "11"
    FormatoGrilla(2, 12) = "11"
    FormatoGrilla(2, 13) = "12"
    FormatoGrilla(2, 14) = "30"
    FormatoGrilla(2, 15) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FormatoGrilla(3, 1) = "D"
    FormatoGrilla(3, 2) = "S"
    FormatoGrilla(3, 3) = "S"
    FormatoGrilla(3, 4) = "S"
    FormatoGrilla(3, 5) = "S"
    FormatoGrilla(3, 6) = "S"
    FormatoGrilla(3, 7) = "S"
    FormatoGrilla(3, 8) = "S"
    FormatoGrilla(3, 9) = "D"
    FormatoGrilla(3, 10) = "D"
    FormatoGrilla(3, 11) = "N"
    FormatoGrilla(3, 12) = "N"
    FormatoGrilla(3, 13) = "N"
    FormatoGrilla(3, 14) = "S"
    FormatoGrilla(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    FormatoGrilla(4, 11) = "###,###,###,##0"
    FormatoGrilla(4, 12) = "###,###,###,##0"
    FormatoGrilla(4, 13) = "###,###,###,##0"
    Rem LOCCKED
    FormatoGrilla(5, 1) = "TRUE"
    FormatoGrilla(5, 2) = "TRUE"
    FormatoGrilla(5, 3) = "TRUE"
    FormatoGrilla(5, 4) = "TRUE"
    FormatoGrilla(5, 5) = "TRUE"
    FormatoGrilla(5, 6) = "TRUE"
    FormatoGrilla(5, 7) = "TRUE"
    FormatoGrilla(5, 8) = "TRUE"
    FormatoGrilla(5, 9) = "TRUE"
    FormatoGrilla(5, 10) = "TRUE"
    FormatoGrilla(5, 11) = "TRUE"
    FormatoGrilla(5, 12) = "TRUE"
    FormatoGrilla(5, 13) = "TRUE"
    FormatoGrilla(5, 14) = "TRUE"
    FormatoGrilla(5, 15) = "TRUE"
    
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
        
        infogrilla.Grid1.Cell(0, k).text = FormatoGrilla(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FormatoGrilla(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FormatoGrilla(2, k))
        infogrilla.Grid1.Column(k).FormatString = FormatoGrilla(4, k)
        infogrilla.Grid1.Column(k).Locked = FormatoGrilla(5, k)
        If FormatoGrilla(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FormatoGrilla(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub leecuentas(infogrilla As grillainformes)
BARRA.Visible = True
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
infogrilla.Grid1.AutoRedraw = False

        Set csql2.ActiveConnection = temporal
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor  "
        csql2.sql = csql2.sql + "WHERE año='" + COMBOAÑO.text + "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        BARRA.Max = csql2.RowsAffected + 4
        BARRA.Value = 0
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        If Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERMOVIMIENTOS(resultados2(0), resultados2(1), infogrilla)
        BARRA.Value = BARRA.Value + 1
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        infogrilla.Grid1.Column(8).Locked = True
        infogrilla.Grid1.Column(9).Locked = True
        infogrilla.Grid1.Column(10).Locked = True
  BARRA.Visible = False
  
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh


End Sub

Sub LEERSALDOS(cuenta)
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
    
    condicion = "codigo=" + "'" + cuenta + "' and año='" + año + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = temporal
    Call sqlconta.sqlconta(op, condicion)
    
 '   If sqlconta.status = 4 Then Stop

End Sub
Sub DATOSSALDOS(cuenta)

Call LEERSALDOS(cuenta)
sumador = Val(sqlconta.response(2, 3)) - Val(sqlconta.response(3, 3))
For k = 1 To Val(mes) - 1
sumador = sumador + Val(sqlconta.response(k + 3, 3)) - Val(sqlconta.response(k + 15, 3))
Next k
saldo = sumador
End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

