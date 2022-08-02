VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form detalleimpuestos 
   Caption         =   "DETALLE IMPUESTOS"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp impuestos 
      Height          =   4290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   7567
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid2 
         Height          =   3210
         Left            =   45
         TabIndex        =   1
         Top             =   270
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   5662
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         ExtendLastCol   =   -1  'True
         Rows            =   30
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "<ESC> PASAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2295
         TabIndex        =   2
         Top             =   3735
         Width           =   2490
      End
   End
End
Attribute VB_Name = "detalleimpuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If


End Sub

Private Sub Form_Load()
FormatoImpuestos
CargaImpuestos

End Sub
Sub Grabar_Impuestos()
    Dim tipodoc As String
    tipodoc = Mid(Rcompra02.Grid1.Cell(Rcompra02.Grid1.ActiveCell.row, 1).text, 1, 3)
    
    
    Call Elimina_Impuestos(tipodoc, Rcompra02.Grid1.Cell(Rcompra02.Grid1.ActiveCell.row, 2).text, Rcompra02.DATO5.text & Rcompra02.DV.Caption)

    For k = 1 To Grid2.Rows - 1
        If CDbl(Grid2.Cell(k, 3).text) > 0 Then
            campos(0, 0) = "tipo"
            campos(1, 0) = "numero"
            campos(2, 0) = "rut"
            campos(3, 0) = "cuenta"
            campos(4, 0) = "monto"
            campos(5, 0) = "numeroorden"
            campos(6, 0) = ""
            
            campos(0, 1) = tipodoc
            campos(1, 1) = Rcompra02.Grid1.Cell(Rcompra02.Grid1.ActiveCell.row, 2).text
            campos(2, 1) = Rcompra02.dato7.text & Rcompra02.dv2.Caption
            campos(3, 1) = Grid2.Cell(k, 1).text
            campos(4, 1) = Grid2.Cell(k, 3).text
            campos(5, 1) = Rcompra02.dato1.text
            

            campos(0, 2) = "l_ordendecompra_impuestos_" & localorden
            
            
            condicion = ""
            op = 2
            sqlconta.response = campos
            Set sqlconta.conexion = gestionrubro
            Call sqlconta.sqlconta(op, condicion)
        End If
    Next k
    For k = 1 To Grid2.Rows - 1
        Grid2.Cell(k, 3).text = ""
    Next k

End Sub

Sub Elimina_Impuestos(tipo_documento As String, numero_documento As String, rut As String)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset

    Set csql.ActiveConnection = gestionrubro
    csql.sql = "DELETE FROM l_ordendecompra_impuestos_" & localorden & " "
    csql.sql = csql.sql + "WHERE tipo = '" & tipo_documento & "' "
    csql.sql = csql.sql + "AND numero = '" & numero_documento & "' "
    csql.sql = csql.sql + "AND numeroorden = '" & Rcompra02.dato1.text & "'"
    
    csql.Execute
    Call sincronizadatos(csql.sql, gestionrubro, "")
    
End Sub


Sub SumaImpuestos()
    Dim valor As Double
    
    valor = 0
    For k = 1 To Grid2.Rows - 1
        If Grid2.Cell(k, 3).text = "" Then Grid2.Cell(k, 3).text = "0"
        valor = valor + CDbl(Grid2.Cell(k, 3).text)
    Next k
    Rcompra02.Grid1.Cell(Rcompra02.Grid1.ActiveCell.row, 8).text = valor
    
    
End Sub


Sub FormatoImpuestos()
    Grid2.Cols = 4
    Grid2.Rows = 1
    
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
    Grid2.BackColor1 = RGB(231, 235, 247)
    Grid2.BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
    Grid2.Cell(0, 1).text = "CODIGO"
    Grid2.Cell(0, 2).text = "IMPUESTO"
    Grid2.Cell(0, 3).text = "MONTO"
    Grid2.Column(0).Width = 0
    Grid2.Column(1).Width = "80"
    Grid2.Column(2).Width = "150"
    Grid2.Column(3).Width = "100"
    Grid2.Column(3).Alignment = cellRightGeneral
    
    Grid2.Column(3).FormatString = " ###,###,##0"
    Grid2.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    Grid2.Column(3).Locked = False
    Grid2.Column(3).Mask = cellNumeric
    
    
End Sub

Sub CargaImpuestos()

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim tipodoc As String
    Dim total As Double

    Set csql2.ActiveConnection = contadb
    csql2.sql = "SELECT codigo,nombre "
    csql2.sql = csql2.sql + "FROM cuentasdelmayor "
    csql2.sql = csql2.sql + "WHERE (ila <>'0' OR iha <>'0' OR ica <>'0') AND año='" + Format(fechasistema, "yyyy") + "' "
    csql2.sql = csql2.sql + "ORDER BY codigo"
    csql2.Execute
    tipodoc = Mid(Rcompra02.Grid1.Cell(Rcompra02.Grid1.ActiveCell.row, 1).text, 1, 3)
    
    
    
    Grid2.Rows = 1
    If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
            total = LeerImpuestos(tipodoc, Rcompra02.Grid1.Cell(Rcompra02.Grid1.ActiveCell.row, 2).text, Rcompra02.DATO5.text & Rcompra02.DV.Caption, resultados2(0).Value)
            Grid2.AddItem resultados2(0) & vbTab & resultados2(1) & vbTab & total, True
            resultados2.MoveNext
        Wend
        resultados2.Close
        Set resultados2 = Nothing
        
    End If
Grid2.AddItem ""

Grid2.Cell(1, 3).SetFocus


End Sub

Function LeerImpuestos(tipo_documento As String, numero_documento As String, rut As String, codigo_impuesto As String) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery

    Set csql.ActiveConnection = gestionrubro
    csql.sql = "SELECT monto "
    csql.sql = csql.sql + "FROM l_ordendecompra_impuestos_" & localorden & " "
    csql.sql = csql.sql + "WHERE tipo = '" & tipo_documento & "' "
    csql.sql = csql.sql + "AND numero = '" & numero_documento & "' "
    csql.sql = csql.sql + "AND numeroorden = '" & Rcompra02.dato1.text & "' "
    csql.sql = csql.sql + "AND cuenta = '" & codigo_impuesto & "'"
    csql.Execute

    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            LeerImpuestos = resultados(0).Value
            resultados.MoveNext
        Wend
        resultados.Close
        Set resultados = Nothing
    Else
        LeerImpuestos = 0
    End If
End Function


Private Sub Grid2_KeyPress(KeyAscii As Integer)

    
    If Grid2.ActiveCell.row = Grid2.Rows - 1 Then
    If KeyAscii = 13 Then
    
    Call SumaImpuestos
    Call Grabar_Impuestos
    Unload Me
    End If
    End If

End Sub

Private Sub Grid2_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
   
    If NewCol <> 3 Then NewCol = 3

End Sub

