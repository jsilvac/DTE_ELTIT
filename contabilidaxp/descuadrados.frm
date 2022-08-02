VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form control01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5145
   ClientLeft      =   240
   ClientTop       =   1290
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   5280
      TabIndex        =   12
      Top             =   4440
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
      Alignment       =   1
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   280
         Width           =   1335
      End
   End
   Begin CoolButtons.cool_Button GENERA 
      Height          =   495
      Left            =   3015
      TabIndex        =   0
      Top             =   4320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "GENERA INFORME"
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   5100
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8996
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
         Left            =   315
         TabIndex        =   5
         Top             =   3915
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   3195
         Left            =   1305
         TabIndex        =   2
         Top             =   450
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5636
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
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FF8080&
            Caption         =   "Anual"
            Height          =   330
            Left            =   2700
            TabIndex        =   11
            Top             =   2790
            Width           =   1140
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Mensual"
            Height          =   330
            Left            =   585
            TabIndex        =   10
            Top             =   2790
            Width           =   1140
         End
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
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   855
            Left            =   0
            TabIndex        =   6
            Top             =   855
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
               Height          =   315
               Left            =   240
               TabIndex        =   7
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   855
            Left            =   0
            TabIndex        =   8
            Top             =   1815
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
               Height          =   315
               Left            =   240
               TabIndex        =   9
               Top             =   360
               Width           =   3855
            End
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
            Left            =   765
            TabIndex        =   4
            Top             =   360
            Width           =   3255
         End
      End
   End
End
Attribute VB_Name = "control01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(10, 20) As String
Private sumas(10) As Double
Private suma(10) As Double
Private sumas2(10) As Double
Private sumas3(10) As Double
Private montos(5) As Double
Private lin As Double
Private año As String
Private MES As String


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
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
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
   
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub

Private Sub Form_Load()
CENTRAR Me

 Call Conectar_BD
 Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
dato1.text = empresaactiva
Option1.Value = True

dato1.text = empresaactiva
empresanombre.Caption = nombreempresa
End Sub


Sub ACEPTA(opcion)
Dim TIMBRA As String

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

If opcion = 1 Then infogrilla.Caption = "BUSCA DESCUADRADOS ": grillainformes.Tag = "control01"
Rem infogrilla.CABEZA.Caption = "BUSCA POR MONTO  de " & COMBOMES.text & " del " & año + " de la empresa " + empresanombre.Caption

lin = 0
Call CARGAGRILLA(infogrilla)

Call Consulta_InformeS(infogrilla)

infogrilla.Visible = True

infogrilla.Show

End Sub


Private Sub GENERA_Click()
Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)
año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

Call ACEPTA(1)
Unload Me

End Sub


    
Sub Consulta_InformeS(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim busca As String
    Dim LLAVE As Double
    Dim suma As Double
    
    
  
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,numero,fecha,"
        csql.sql = csql.sql + "SUM(CASE WHEN (dh = 'D') THEN monto ELSE 0 END),SUM(CASE WHEN (dh = 'H') THEN monto ELSE 0 END) "
        
        csql.sql = csql.sql + "FROM movimientoscontables "
        If Option1.Value = True Then
        csql.sql = csql.sql + "where mes='" + MES + "' and año='" + año + "' "
        Else
        csql.sql = csql.sql + "where año='" + año + "' "
        
        End If
        
        csql.sql = csql.sql + "GROUP by tipo,numero,fecha "
        csql.sql = csql.sql + "ORDER BY tipo,numero,fecha "
        
        csql.Execute

        
        
        
        
        infogrilla.Grid1.AutoRedraw = False
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
        
        lin = 0: LLAVE = 0: suma = 0
         While Not resultados.EOF
            If resultados(4) - resultados(3) <> 0 Then
            lin = lin + 1
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            infogrilla.Grid1.Cell(lin, 1).text = resultados(0)
            infogrilla.Grid1.Cell(lin, 2).text = resultados(1)
            infogrilla.Grid1.Cell(lin, 3).text = resultados(2)
            infogrilla.Grid1.Cell(lin, 4).text = resultados(3)
            infogrilla.Grid1.Cell(lin, 5).text = resultados(4)
            infogrilla.Grid1.Cell(lin, 6).text = resultados(3) - resultados(4)
            suma = suma + resultados(3) - resultados(4)
            End If
            
             resultados.MoveNext
            
           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
            lin = lin + 1
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            infogrilla.Grid1.Cell(lin, 6).text = suma
            
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh

End Sub

Sub totalcomprobante(row, infogrilla As grillainformes)
    
    
    
    
    
    
    With infogrilla.Grid1.Range(row, 11, row, 12)
    
    .Borders(cellEdgeTop) = cellThin
    
    
    
     End With
   With infogrilla.Grid1.Range(row, 1, row, 12)
   .FontBold = True
    .FontUnderline = True
    End With
    
    
    
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
    infogrilla.Grid1.DefaultFont.Size = 7.5
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "DEBE"
    FORMATOGRILLA(1, 5) = "HABER"
    FORMATOGRILLA(1, 6) = "DIFERENCIA"
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "5"
    FORMATOGRILLA(2, 2) = "20"
    FORMATOGRILLA(2, 3) = "20"
    FORMATOGRILLA(2, 4) = "20"
    FORMATOGRILLA(2, 5) = "20"
    FORMATOGRILLA(2, 6) = "20"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "D"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 4) = "###,###,###,###"
    FORMATOGRILLA(4, 5) = "###,###,###,###"
    FORMATOGRILLA(4, 6) = "###,###,###,###"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    
    infogrilla.Grid1.Cols = 7
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

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
