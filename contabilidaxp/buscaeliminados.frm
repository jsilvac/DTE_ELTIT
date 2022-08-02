VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form control04 
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
      TabIndex        =   10
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
         Height          =   2610
         Left            =   1305
         TabIndex        =   2
         Top             =   450
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4604
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
            Caption         =   "Cuentas Corrientes"
            Height          =   330
            Left            =   2475
            TabIndex        =   9
            Top             =   2115
            Width           =   1635
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Cuentas Del Mayor"
            Height          =   330
            Left            =   405
            TabIndex        =   8
            Top             =   2115
            Value           =   -1  'True
            Width           =   1680
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
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   855
            Left            =   0
            TabIndex        =   6
            Top             =   900
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
               TabIndex        =   7
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
Attribute VB_Name = "control04"
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

If Option1.Value = True Then infogrilla.Caption = "BUSCA CUENTAS MAYOR ELIMINADAS ": grillainformes.Tag = "control04"
If Option1.Value = False Then infogrilla.Caption = "BUSCA CUENTAS CORRIENTES ELIMINADAS ": grillainformes.Tag = "control04"

Rem infogrilla.CABEZA.Caption = "BUSCA POR MONTO  de " & COMBOMES.text & " del " & año + " de la empresa " + empresanombre.Caption

lin = 0
Call CARGAGRILLA(infogrilla)

Call Consulta_InformeS(infogrilla)

infogrilla.Visible = True

infogrilla.Show

End Sub


Private Sub GENERA_Click()
Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)

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
    Dim AÑOPROCESO As String
    AÑOPROCESO = Mid(COMBOAÑO.text, 1, 4)
    
      Set csql.ActiveConnection = contadb
      
      If Option1.Value = True Then
      csql.sql = "select mo.tipo,mo.numero,mo.fecha,mo.rutctacte,mo.codigocuenta,mo.glosacontable,cm.nombre "
      csql.sql = csql.sql + "from movimientoscontables as mo left join  cuentasdelmayor as cm on mo.codigocuenta=cm.codigo and cm.año='" & AÑOPROCESO & "' "
      csql.sql = csql.sql + "where mo.fecha like '" & AÑOPROCESO & "%'  and isnull(cm.nombre) "
      csql.sql = csql.sql + "group by codigocuenta "
      csql.sql = csql.sql + "order by cm.nombre "
      csql.Execute
      Else
      csql.sql = "select mo.tipo,mo.numero,mo.fecha,mo.rutctacte,mo.codigocuenta,mo.glosacontable,cc.nombre "
      csql.sql = csql.sql + "from movimientoscontables as mo "
      csql.sql = csql.sql + "left join  cuentascorrientes as cc on mo.codigocuenta=cc.tipo and cc.año='" & AÑOPROCESO & "' and cc.rut=mo.rutctacte "
      csql.sql = csql.sql + "inner join cuentasdelmayor as cm on mo.codigocuenta=cm.codigo and cm.año='" & AÑOPROCESO & "' "
      csql.sql = csql.sql + "where mo.fecha like '" + AÑOPROCESO + "%' and rutctacte<>'' and isnull(cc.nombre) and cm.ctacte='1' "
      csql.sql = csql.sql + "group by codigocuenta,rutctacte "
      csql.sql = csql.sql + "order by cc.nombre "
      csql.Execute
      End If
  
        
        
        
        infogrilla.Grid1.AutoRedraw = False
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
        
        lin = 0: LLAVE = 0: suma = 0
         While Not resultados.EOF
            
            lin = lin + 1
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            infogrilla.Grid1.Cell(lin, 1).text = resultados(0)
            infogrilla.Grid1.Cell(lin, 2).text = resultados(1)
            infogrilla.Grid1.Cell(lin, 3).text = resultados(2)
            infogrilla.Grid1.Cell(lin, 4).text = resultados(3)
            infogrilla.Grid1.Cell(lin, 5).text = resultados(4)
            infogrilla.Grid1.Cell(lin, 6).text = resultados(5)
            
            
            
             resultados.MoveNext
            
           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
           
            
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
    FORMATOGRILLA(1, 4) = "RUT"
    FORMATOGRILLA(1, 5) = "CUENTA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "5"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "40"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "D"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "S"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    
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
