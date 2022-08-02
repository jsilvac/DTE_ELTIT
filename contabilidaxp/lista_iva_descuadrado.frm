VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form iva_descuadrado 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEMA AUTOMATICO BANCO SANTANDER"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12435
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   659
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   829
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmprincipal 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   3625
      BackColor       =   16744576
      Caption         =   "DATOS DEL CHEQUE"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483639
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprimir"
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
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   1545
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Excel"
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
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1545
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Genera Informe"
         Height          =   375
         Left            =   10680
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox PIVOTE 
         Height          =   285
         Left            =   18360
         TabIndex        =   1
         Top             =   8160
         Visible         =   0   'False
         Width           =   735
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   855
         Left            =   0
         TabIndex        =   7
         Top             =   0
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
            TabIndex        =   8
            Top             =   360
            Width           =   3855
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   855
         Left            =   0
         TabIndex        =   9
         Top             =   960
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
            TabIndex        =   10
            Top             =   360
            Width           =   3855
         End
      End
   End
   Begin XPFrame.FrameXp frminforme 
      Height          =   7785
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   13732
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483639
      ColorBarraArriba=   16744576
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
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   7560
         Width           =   12360
         _ExtentX        =   21802
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7335
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   12938
         AllowUserSort   =   -1  'True
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
   End
End
Attribute VB_Name = "iva_descuadrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecha1 As String
Dim fecha2 As String
Dim CtaMayorCrcc As Boolean
Dim CtaMayorCtaCte As Boolean
Sub CARGAGRILLA()
Dim FORMATOGRILLA(50, 50) As String
'    infogrilla.Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "FOLIO"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "RUT"
    FORMATOGRILLA(1, 6) = "PROVEEDOR"
    FORMATOGRILLA(1, 7) = "NETO"
    FORMATOGRILLA(1, 8) = "IVA"
    FORMATOGRILLA(1, 9) = "IVA REAL"
    FORMATOGRILLA(1, 10) = "DIFERENCIA"
    FORMATOGRILLA(1, 11) = ""
    
    FORMATOGRILLA(1, 12) = "TOTAL"
    FORMATOGRILLA(1, 13) = " CUENTA "
    FORMATOGRILLA(1, 14) = " MONTO "
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "0"
    FORMATOGRILLA(2, 2) = "4"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "35"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "9"
    FORMATOGRILLA(2, 12) = "9"
    FORMATOGRILLA(2, 13) = "0"
    FORMATOGRILLA(2, 14) = "0"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    FORMATOGRILLA(4, 10) = "###,###,###"
    FORMATOGRILLA(4, 11) = "###,###,###"
    FORMATOGRILLA(4, 12) = "###,###,###"
    FORMATOGRILLA(4, 14) = "###,###,###"
    
    Rem LOCCKED
    For k = 1 To 14
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    Grid1.Cols = 11
    Grid1.Rows = 2
    
     'grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
   'grid1.BackColorFixed = RGB(90, 158, 214)
   ' grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' grid1.BackColorBkg = RGB(90, 158, 214)
   ' grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' grid1.BackColor1 = RGB(231, 235, 247)
   ' grid1.BackColor2 = RGB(239, 243, 255)
   ' grid1.GridColor = RGB(148, 190, 231)
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


Sub GenerarInforme(año, MES)
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Grid1.AutoRedraw = False
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT folio,fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,iva,exento"
        csql2.sql = csql2.sql & " ,impuestoespecifico,retencion,total,fc.electronica"
        csql2.sql = csql2.sql & " ,fc.activo,fc.comentario FROM facturasdecompras as fc "
        csql2.sql = csql2.sql & " ,cuentascorrientes as cc "
        csql2.sql = csql2.sql & " WHERE fc.tipo<>'' and fc.rut=cc.rut and "
        csql2.sql = csql2.sql & " cc.año=añocontable and cc.tipo='" & CUENTAPROVEEDOR & "'"
        csql2.sql = csql2.sql & " and añocontable ='" & Format(año, "0000") & "' and mescontable ='" & Format(MES, "00") & "'"
        csql2.sql = csql2.sql & " HAVING ABS(iva-ROUND((neto/100)*19))>=10  order by folio,fecha"
        csql2.Execute
        
        Grid1.Rows = 1
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        barra.Max = csql2.RowsAffected + 1
        barra.Value = 0
        
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        barra.Value = barra.Value + 1
            Grid1.AddItem "", True
            
                If resultados2(1) = "1" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "FA"
                If resultados2(1) = "2" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "ND"
                If resultados2(1) = "3" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "NC": multi = -1
                If resultados2(1) = "4" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "FAE"
                If resultados2(1) = "5" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "NDE"
                If resultados2(1) = "6" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "NCE": multi = -1
                If resultados2(1) = "7" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "FC"
                If resultados2(1) = "8" Then Grid1.Cell(Grid1.Rows - 1, 2).text = "IM"
                
            
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
            'Grid1.Cell(Grid1.Rows - 1, 2).text = resultados2(1)
            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados2(2)
            Grid1.Cell(Grid1.Rows - 1, 4).text = resultados2(3)
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados2(4)
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados2(5)
            Grid1.Cell(Grid1.Rows - 1, 7).text = resultados2(6)
            Grid1.Cell(Grid1.Rows - 1, 8).text = resultados2(7)
            Grid1.Cell(Grid1.Rows - 1, 9).text = (resultados2("neto") / 100) * 19
            Grid1.Cell(Grid1.Rows - 1, 10).text = Grid1.Cell(Grid1.Rows - 1, 9).text - Grid1.Cell(Grid1.Rows - 1, 8).text
       
        resultados2.MoveNext
        Wend
          
        resultados2.Close
        Set resultados2 = Nothing
        End If
        
Grid1.AutoRedraw = True
Grid1.Refresh

End Sub

Sub LEERMOVIMIENTOS(infogrilla As Grid, cuenta, NOMBRE)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    dedonde = 1
    barra.Value = 0
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
        fecha2 = Mid(hastafecha.Caption, 7, 4) + "-" + Mid(hastafecha.Caption, 4, 2) + "-" + Mid(hastafecha.Caption, 1, 2)
    
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento"
        csql.sql = csql.sql & ",fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte,monto "
        csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' "
        csql.sql = csql.sql & " and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        csql.sql = csql.sql + "order by codigocuenta,fecha,tipo,numero,linea"
        csql.Execute
        
        Call LEERSALDOS(cuenta)
        
 
        
        
        If saldo <> 0 Or csql.RowsAffected <> 0 Then
        lin = lin + 1
        Grid1.Rows = Grid1.Rows + 1
        
        For k = 1 To 6
        Grid1.Column(k).Locked = False
        Next k
                
        Grid1.Range(lin, 1, lin, 6).Merge
      
        Grid1.Cell(lin, 1).CellType = cellTextBox
        
        Grid1.Cell(lin, 10).CellType = cellTextBox
        
        If dedonde = 1 Then
        Grid1.Cell(lin, 1).text = cuenta & " " + NOMBRE
        End If
        
        'If dedonde = 2 Then Grid1.Cell(lin, 6).text = nombrectacte
        Grid1.Cell(lin, 10).text = "SALDO-->"
        
        Grid1.Cell(lin, 13).text = saldo
        Grid1.Range(lin, 0, lin, Grid1.Cols - 1).FontBold = True
        Grid1.Range(lin, 0, lin, Grid1.Cols - 1).FontUnderline = True
        
        
        End If
        
        If csql.RowsAffected > 0 Then
        
        
        Set resultados = csql.OpenResultset
        barra.Max = csql.RowsAffected + 10
         While Not resultados.EOF
         barra.Value = barra.Value + 1
'        If dedonde = 1 And Check2.Value = 1 Then
'        If resultados(15) > 2 Then GoTo dale:
'        End If
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
             If IsNull(resultados("rutctacte")) = False Then Grid1.Cell(lin, 0).text = resultados("rutctacte")
             For k = 0 To 9
             If IsNull(resultados(k)) = False Then Grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             If resultados(11) = "D" Then Grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10): saldo = saldo + resultados(10)
             If resultados(11) = "H" Then Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10): saldo = saldo - resultados(10)
'             Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             Grid1.Cell(lin, 13).text = saldo
             If resultados("rutctacte") <> "" Then
                Grid1.Cell(lin, 14).text = resultados("rutctacte")
                Grid1.Cell(lin, 15).text = leerNombrerut(dato1 & dato2 & dato3, resultados("rutctacte"))
             End If
             If resultados("centrocosto") <> "" Then Grid1.Cell(lin, 16).text = resultados("centrocosto") & " " & leerNOMBREcrcc(resultados("centrocosto"))
             
dale:             resultados.MoveNext
          
         Wend
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
         
         Call totalcomprobante(Grid1, lin)
          resultados.Close
            Set resultados = Nothing

        End If
 For k = 1 To 6
        Grid1.Column(k).Locked = True
        
        Next k
        barra.Value = 0
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub COMBOAÑO_KeyPress(KeyAscii As Integer)
keyasii = 0
End Sub

Private Sub COMBOMES_KeyPress(KeyAscii As Integer)
keyasii = 0
End Sub

Private Sub Command1_Click()
Grid1.PageSetup.Orientation = cellLandscape
Grid1.PageSetup.PrintFixedRow = True
Call Grid1.PrintPreview
End Sub

Private Sub Command10_Click()
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
Call GenerarInforme(año, MES)
End Sub
 
 
 

Private Sub Command5_Click()
Call Grid1.ExportToExcel("", True, True)
End Sub

Private Sub Form_Load()

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001


Call CARGAGRILLA
desdefecha = "01-01-" & Format(fechasistema, "yyyy")
hastafecha = Format(fechasistema, "dd-mm-yyyy")
End Sub

'
'
'Private Sub dato1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 27 Then Unload Me
'    snum = 0: KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
'End Sub
'
'Private Sub dato2_KeyPress(KeyAscii As Integer)
'    snum = 0: KeyAscii = esNumero(KeyAscii)
'
'    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
'End Sub
'
'Private Sub dato3_KeyPress(KeyAscii As Integer)
'    snum = 0: KeyAscii = esNumero(KeyAscii)
'
'    If KeyAscii = 13 Then
'
'    Call ceros(dato3)
'    lblnombrecuenta.Caption = leerNombreMayor(dato1.text + dato2.text + dato3.text)
'    Call Pregunta(dato3, dato3)
'    CARGAGRILLA
'
'    Call leercuenta(dato1.text + dato2.text + dato3.text, "", "")
'
'End If
'
'End Sub

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
   
    


Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    
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


Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" & Format(fechasistema, "yyyy") & "' "
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
    
    condicion = "codigo=" + "'" + cuenta + "'and año='" + Mid(desdefecha.Caption, 7, 4) + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(2, 3)) - Val(sqlconta.response(3, 3))
  
    saldo = sumador
Rem acumula fecha
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)

    
        
        Set cSql3.ActiveConnection = contadb
        cSql3.sql = "SELECT SUM(monto),dh "
         cSql3.sql = cSql3.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and fecha<'" + fecha1 + "' and fecha>='" + Format(fechasistema, "yyyy") + "-01-01" + "' "
     
        
        cSql3.sql = cSql3.sql + "GROUP by DH"
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

Sub totalcomprobante(infogrilla As Grid, row)
    Grid1.Range(row, 11, row, 12).Borders(cellEdgeTop) = cellThin
    Grid1.Range(row, 1, row, 12).FontBold = True
    Grid1.Range(row, 1, row, 12).FontUnderline = True
    
    
    Grid1.Cell(row, 10).CellType = cellTextBox
    Grid1.Cell(row, 10).text = "TOTAL "
    Grid1.Cell(row, 11).text = anted
    Grid1.Cell(row, 12).text = anteh
    lin = lin + 2
    Grid1.Rows = Grid1.Rows + 2
        
    anted = 0: anteh = 0: saldo = 0
    End Sub

Private Sub Grid1_DblClick()
Dim dia As String

If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "FAE" Or Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NCE" Or Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NCD" Then
If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "FAE" Then electro88.tipo.text = "33"
If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NCE" Then electro88.tipo.text = "61"
If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NDE" Then electro88.tipo.text = "56"
End If


electro88.FOLIO.text = Grid1.Cell(Grid1.ActiveCell.row, 3).text

electro88.Show vbModal

End Sub


Function LeerTipoCtaMayor(codigo) As String
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "ctacte"
    campos(3, 0) = "crcc"
    campos(4, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo=" + "'" + codigo + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    CtaMayorCtaCte = False
    CtaMayorCrcc = False
    If sqlconta.status = 4 Then
    
    Else
    If sqlconta.response(2, 3) = 1 Then CtaMayorCtaCte = True
    If sqlconta.response(3, 3) = 1 Then CtaMayorCrcc = True
    LeerTipoCtaMayor = sqlconta.response(1, 3)
    End If
no:

End Function
