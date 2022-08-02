VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form listadoCodigosSinCuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codigos Sin Cuenta"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8916
      BackColor       =   16744576
      Caption         =   "DETALLE CODIGOS SIN CUENTA"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid1 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8281
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "listadoCodigosSinCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call CARGAGRILLA
End Sub


Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 10)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "MONTO"
    FORMATOGRILLA(1, 3) = "GLOSA"
    FORMATOGRILLA(1, 4) = ""
    FORMATOGRILLA(1, 5) = "EMISION"
    FORMATOGRILLA(1, 6) = "VENCIMIENTO"
    FORMATOGRILLA(1, 7) = "MONTO"
    FORMATOGRILLA(1, 8) = "TAZA"
    FORMATOGRILLA(1, 9) = "INTERES"
    FORMATOGRILLA(1, 10) = "TOTAL"
    
     
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "30"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"
    FORMATOGRILLA(2, 10) = "10"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "D"
    FORMATOGRILLA(3, 6) = "D"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
   
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "###,###,###,##0"
    FORMATOGRILLA(4, 8) = "###,###,##0.000"
    FORMATOGRILLA(4, 9) = "###,###,###,##0"
    FORMATOGRILLA(4, 10) = "###,###,###,##0"
  
    
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
    
    Grid1.Cols = 4
    Grid1.Rows = 1
    
 Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
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


Public Sub buscaCuentas(empresa, año, MES)
 Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        Dim totalD As Double
        Dim totalH As Double
        Dim tabla As String
        Set sql.ActiveConnection = contadb

        tabla = " SELECT li.codigohd,SUM(li.monto) AS monto,li.glosa AS glosa ,ac.contable,ac.codigo,ac.dh,li.columnalibro"
        tabla = tabla & " FROM " & clientesistema & "remu" & empresa & ".calculoliquidaciones li"
        tabla = tabla & " LEFT JOIN eltit_remu.asiento_contable AS ac ON ac.codigo=li.codigohd"
        tabla = tabla & " WHERE mes='" & MES & "' AND año='" & año & "' "
        tabla = tabla & " AND MID(li.codigohd,1,1)<>'I'"
        tabla = tabla & " AND MID(li.codigohd,1,2)<>'B0'"
        tabla = tabla & " AND MID(li.codigohd,1,1)<>'A'"
        tabla = tabla & " AND MID(li.codigohd,1,2)<>'CA'"
        tabla = tabla & " AND MID(li.codigohd,1,1)<>' '"
        tabla = tabla & " AND  li.codigohd <>'00001'"
        tabla = tabla & " AND MID(li.codigohd,1,2)<>'CF'"
        tabla = tabla & " AND MID(li.codigohd,1,5)<>'CONIS'"
        tabla = tabla & " AND MID(li.codigohd,1,2)<>'CT'"
        tabla = tabla & " AND MID(li.codigohd,1,1)<>'D'"
        tabla = tabla & " AND MID(li.codigohd,1,1)<>'F'"
        tabla = tabla & " AND MID(li.codigohd,1,5)<>'HDO01'"
        tabla = tabla & " AND MID(li.codigohd,1,5)<>'HE001'"
         tabla = tabla & " AND MID(li.codigohd,1,1)<>'P'"
         tabla = tabla & " AND MID(li.codigohd,1,2)<>'SB'"
         tabla = tabla & " AND MID(li.codigohd,1,2)<>'SE'"
         tabla = tabla & " AND MID(li.codigohd,1,2)<>'ST'"
         tabla = tabla & " AND MID(li.codigohd,1,2)<>'SI'"
          tabla = tabla & " AND MID(li.codigohd,1,2)<>'TD'"
         tabla = tabla & " AND MID(li.codigohd,1,3)<>'THG'"
         tabla = tabla & " AND MID(li.codigohd,1,3)<>'THI'"
         tabla = tabla & " AND MID(li.codigohd,1,1)<>'W'"
         tabla = tabla & " AND MID(li.codigohd,1,5)<>'XDO01'"
         tabla = tabla & " AND MID(li.codigohd,1,5)<>'H2O01'"
         tabla = tabla & " AND MID(li.codigohd,1,5)<>'X2O01' "
         tabla = tabla & " AND MID(li.codigohd,1,5)<>'MUT00' "
        
        tabla = tabla & " GROUP BY li.codigohd"
        
        tabla = tabla & " Having IsNull(codigo) = True"
        tabla = tabla & " ORDER BY li.codigohd"
         sql.sql = tabla
        sql.Execute
       Grid1.Rows = 1
         
        If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
            While resultados.EOF = False
                Grid1.Rows = Grid1.Rows + 1
                    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
                    Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
                    Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
                    total = total + resultados("monto")
            resultados.MoveNext
            Wend
            
        End If
                    Grid1.Rows = Grid1.Rows + 1
                   
                    Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL"
                    Grid1.Cell(Grid1.Rows - 1, 2).text = total
 
                    
        Grid1.AutoRedraw = True
        Grid1.Refresh
End Sub
