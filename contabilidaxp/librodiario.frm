VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "clbutn.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form auxiliar04 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LIBRO DIARIO"
   ClientHeight    =   6345
   ClientLeft      =   240
   ClientTop       =   1290
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Version SII"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Formato Para SII XML Emisor Dte o 3328"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   3480
      Width           =   3975
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   855
      Left            =   1320
      TabIndex        =   10
      Top             =   4560
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
   Begin CoolButtons.cool_Button GENERA 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   5520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "GENERA INFORME"
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   4215
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7435
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
         Top             =   3720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   855
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
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
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
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
         Height          =   1095
         Left            =   1440
         TabIndex        =   7
         Top             =   2280
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1931
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
      Left            =   4680
      TabIndex        =   14
      Top             =   5640
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
         TabIndex        =   16
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "auxiliar04"
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
Private ANTED2 As Double
Private ANTEH2 As Double







Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(DATO1)
    
End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", DATO1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub


Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + DATO1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then DATO1.SetFocus: GoTo no:
    COMBOMES.SetFocus
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
DATO1.text = empresaactiva
empresanombre.Caption = nombreempresa
original.Value = True



End Sub


Sub ACEPTA(opcion)
Dim TIMBRA As String

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes
If Check2.Value = 1 Then
xmllibrodiario = True

End If

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"
If opcion = 1 Then infogrilla.Caption = "LIBRO DIARIO": grillainformes.Tag = "auxiliar04" & TIMBRA & FOLIO.text
infogrilla.CABEZA.Caption = "LIBRO DIARIO de " & COMBOMES.text & " del " & año + " de la empresa " + empresanombre.Caption
infogrilla.titulofinal.Caption = "LIBRO DIARIO DE " & COMBOMES.text & " DEL " & año
lin = 0
Call CARGAGRILLA(infogrilla)

Call Consulta_InformeS(infogrilla)

infogrilla.Visible = True

infogrilla.Show

End Sub


Private Sub GENERA_Click()
Call Conectartemporal(Servidor, clientesistema + "conta" + DATO1.text, Usuario, password)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1

If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

Call ACEPTA(1)
Unload Me

End Sub


    
Sub Consulta_InformeS(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fechacom As String
    Dim lineascom As Double
    
    Dim totales(31, 3) As Variant
    
        Set csql.ActiveConnection = temporal
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,tiposii "
        If Check1.Value = 1 Then
        csql.sql = csql.sql + "FROM movimientoscontables_resumen ," + clientesistema + "conta.maestrotipodedocumentos as td WHERE mes='" + MES + "' and año='" + año + "' and tipos=tipo "
        Else
        csql.sql = csql.sql + "FROM movimientoscontables ," + clientesistema + "conta.maestrotipodedocumentos as td WHERE mes='" + MES + "' and año='" + año + "' and tipos=tipo "
        
        End If
        
        csql.sql = csql.sql + "order by tipo,numero,linea" 'fecha,tiposii,numero,linea "
        csql.Execute

        
        
        
        
        infogrilla.Grid1.AutoRedraw = False
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0: lineascom = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        lin = 0: PASO = resultados(1) + resultados(2)
        fechacom = resultados(0)
         
         While Not resultados.EOF
          lin = lin + 1
             lineascom = lineascom + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             If resultados(1) + resultados(2) <> PASO Then
             Call totalcomprobante(lin, infogrilla)
             totales(Format(fechacom, "dd"), 1) = totales(Format(fechacom, "dd"), 1) + 1
             totales(Format(fechacom, "dd"), 2) = totales(Format(fechacom, "dd"), 2) + lineascom
             lineascom = 0
             fechacom = resultados(0)
             
             
             
             End If
             
             For k = 0 To 9
             infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             infogrilla.Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             barra.Value = barra.Value + 1
             If resultados(11) = "D" Then
             totales(Format(fechacom, "dd"), 3) = totales(Format(fechacom, "dd"), 3) + resultados(10)
             infogrilla.Grid1.Cell(lin, 11).text = resultados(10)
             anted = anted + resultados(10)
             ANTED2 = ANTED2 + resultados(10)
             Else
             infogrilla.Grid1.Cell(lin, 12).text = resultados(10)
             anteh = anteh + resultados(10)
             ANTEH2 = ANTEH2 + resultados(10)
             End If
             PASO = resultados(1) + resultados(2)
             resultados.MoveNext

           
         Wend
          infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
          lin = lin + 1
          Call totalcomprobante(lin, infogrilla)
          totales(Format(fechacom, "dd"), 1) = totales(Format(fechacom, "dd"), 1) + 1
           totales(Format(fechacom, "dd"), 2) = totales(Format(fechacom, "dd"), 2) + lineascom
             
             lineascom = 0
            
             
          infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
          lin = lin + 1
          Call totalcomprobante2(lin, infogrilla)
          
          resultados.Close
            Set resultados = Nothing

        End If
        If xmllibrodiario = True Then
        For k = 1 To 31
          infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             lin = infogrilla.Grid1.Rows - 1
            infogrilla.Grid1.Cell(lin, 7).text = k
            infogrilla.Grid1.Cell(lin, 8).text = totales(k, 1)
            infogrilla.Grid1.Cell(lin, 11).text = totales(k, 2)
            infogrilla.Grid1.Cell(lin, 12).text = totales(k, 3)
          
          Next k
          
        
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
Sub totalcomprobante2(row, infogrilla As grillainformes)
    
    
    With infogrilla.Grid1.Range(row, 11, row, 12)
    
    .Borders(cellEdgeTop) = cellThin
    
    
    
     End With
   With infogrilla.Grid1.Range(row, 1, row, 12)
   .FontBold = True
    .FontUnderline = True
    End With
    
    
    
    infogrilla.Grid1.Cell(row, 10).CellType = cellTextBox
    
    
    infogrilla.Grid1.Cell(row, 10).text = "TOTAL GENERAL"
    infogrilla.Grid1.Cell(row, 11).text = ANTED2
    infogrilla.Grid1.Cell(row, 12).text = ANTEH2
    lin = lin + 2
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
        
        ANTED2 = 0: ANTEH2 = 0
    End Sub
    
 


Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7.5
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "NL"
    FORMATOGRILLA(1, 5) = "CUENTA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "TP"
    FORMATOGRILLA(1, 8) = "NUMERO"
    FORMATOGRILLA(1, 9) = "EMISION"
    FORMATOGRILLA(1, 10) = "VENCIMIENTO"
    FORMATOGRILLA(1, 11) = "DEBE"
    FORMATOGRILLA(1, 12) = "HABER"
    FORMATOGRILLA(1, 13) = "NOMBRE CUENTA"
    FORMATOGRILLA(1, 14) = "CUENTA CORRIENTE"
    FORMATOGRILLA(1, 15) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "3"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "25"
    FORMATOGRILLA(2, 7) = "3"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "0"
    FORMATOGRILLA(2, 10) = "0"
    FORMATOGRILLA(2, 11) = "12"
    FORMATOGRILLA(2, 12) = "12"
    FORMATOGRILLA(2, 13) = "30"
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
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 11) = "###,###,###,###"
    FORMATOGRILLA(4, 12) = "###,###,###,###"
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
    
    infogrilla.Grid1.Cols = 13
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

