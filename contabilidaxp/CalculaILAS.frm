VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form informeIlas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo de Impuesto ILA"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   1931
      BackColor       =   16761024
      Caption         =   "Lista Calculo de Impuesto ILA"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command3 
         Caption         =   "Exportar Excel"
         Height          =   255
         Left            =   10920
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   10800
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Informe"
         Height          =   255
         Left            =   10800
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   12091
      BackColor       =   16761024
      Caption         =   "Informe"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   6360
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6015
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   10610
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp5 
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1296
      BackColor       =   16761024
      Caption         =   "Maestro de Empresas"
      BackColor       =   16761024
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox comboempresas 
         BackColor       =   &H00FFC0C0&
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "informeIlas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 12)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "LICORES"
    FORMATOGRILLA(1, 2) = "CREDITO DEL MES"
    FORMATOGRILLA(1, 3) = "PROP.COMPRAS"
    FORMATOGRILLA(1, 4) = "%"
    FORMATOGRILLA(1, 5) = "PROP. VENTAS"
    FORMATOGRILLA(1, 6) = ""
    FORMATOGRILLA(1, 7) = ""
    FORMATOGRILLA(1, 8) = ""
    FORMATOGRILLA(1, 9) = ""
    FORMATOGRILLA(1, 10) = ""
    FORMATOGRILLA(1, 11) = ""
    
     
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "30"
    FORMATOGRILLA(2, 2) = "15"
    FORMATOGRILLA(2, 3) = "15"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "15"
    FORMATOGRILLA(2, 6) = "5"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "15"
    FORMATOGRILLA(2, 9) = "15"
    FORMATOGRILLA(2, 10) = "15"
    FORMATOGRILLA(2, 11) = "15"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
   
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(5, 2) = "###,###,###,##0"
    FORMATOGRILLA(5, 3) = "###,###,###,##0"
    FORMATOGRILLA(5, 4) = "% ###,###,###,##0.00"
    FORMATOGRILLA(5, 5) = "###,###,###,##0"
    
    
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
     Grid1.Cols = 1
    Grid1.Cols = 6
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
        If FORMATOGRILLA(3, k) = "N" Then
            Grid1.Column(k).Alignment = cellRightCenter
            Grid1.Column(k).Mask = cellNumeric
        End If
        
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
End Sub

Private Sub Command1_Click()
  Call GenerarInforme
End Sub

Private Sub COMMAND2_Click()
Call Titulos(Grid1)
Grid1.PrintPreview
End Sub

Private Sub Form_Load()
Call CARGAGRILLA
CARGAempresas
End Sub

 
Sub CARGAempresas()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre from "
        csql.sql = csql.sql & clientesistema & "gestion.g_maestroempresas "
        csql.sql = csql.sql & " where codigocontable='" & empresaactiva & "' "
        csql.sql = csql.sql & " order by codigo"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             comboempresas.AddItem (resultados(0) & " " & resultados(1))
             
            resultados.MoveNext
            Wend
        End If
        comboempresas.AddItem ("99 " & "TODOS")
            
        comboempresas.text = "99 " & "TODOS"

End Sub

Private Sub GenerarInforme()
Dim n As Double
Dim resultados As rdoResultset
Dim tabla As String
Dim basedatos As String
Dim rubro_loc As String
Dim periodo As String
Dim csql As New rdoQuery
CARGAGRILLA
'GoTo PASO
If MsgBox("TIMPO ESTIMADO 2 MINUTOS" & vbNewLine & "DESEA CONTINUAR?", vbYesNo, "ATENCION") = vbNo Then
    Exit Sub
End If

Me.ProgressBar1.Max = 10
Me.ProgressBar1.Value = 0
  '  Call Conectar_BD
    periodo = Format(fechasistema, "yyyy-mm")
    Set csql.ActiveConnection = contadb
    
    tabla = clientesistema & "conta.resumen_ilas "
    
    csql.sql = "TRUNCATE TABLE " & tabla
    csql.Execute
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    Select Case empresaactiva
        Case "08"
            ProgressBar1.Value = ProgressBar1.Value + 1
            empresa = "00"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            '-----
            ProgressBar1.Value = ProgressBar1.Value + 1
            empresa = "25"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            '-----
            ProgressBar1.Value = ProgressBar1.Value + 1
            empresa = "41"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            '------
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            csql.sql = "INSERT INTO " & tabla
            csql.sql = csql.sql & " SELECT mpf.codigobarra,mpf.codigoimpuesto,"
        
            csql.sql = csql.sql & "(IFNULL((SELECT SUM(total) FROM eltit_ventas00.sv_ila_detalle_00 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas00.sv_ila_detalle_00 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0)+"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas25.sv_ila_detalle_25 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas25.sv_ila_detalle_25 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0)+"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas41.sv_ila_detalle_41 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas41.sv_ila_detalle_41 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0)) AS FAC,"
            
            csql.sql = csql.sql & " (IFNULL((SELECT SUM(total) FROM eltit_ventas00.sv_ila_detalle_00 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas00.sv_ila_detalle_00 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0)+"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas25.sv_ila_detalle_25 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas25.sv_ila_detalle_25 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0) +"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas41.sv_ila_detalle_41 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas41.sv_ila_detalle_41 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0) ) AS BOL"

            csql.sql = csql.sql & " FROM " & clientesistema & "gestion" & rubro_loc
            csql.sql = csql.sql & ".r_maestroproductos_fijo_" & rubro_loc & " AS mpf"
            
            csql.sql = csql.sql & " Where (mpf.codigoimpuesto='00001' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00002' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00003' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00006' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00007')"
            csql.sql = csql.sql & " GROUP BY codigobarra "
            csql.sql = csql.sql & " ORDER BY codigobarra DESC "
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
        Case "21"
            empresa = "17"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            '------
            empresa = "18"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            '------
            
            csql.sql = "INSERT INTO " & tabla
            csql.sql = csql.sql & " SELECT mpf.codigobarra,mpf.codigoimpuesto,"
        
            csql.sql = csql.sql & " (IFNULL((SELECT SUM(total) FROM eltit_ventas17.sv_ila_detalle_17 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas17.sv_ila_detalle_17 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0)+"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas18.sv_ila_detalle_18 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas18.sv_ila_detalle_18 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0) ) AS FAC,"
            csql.sql = csql.sql & " (IFNULL((SELECT SUM(total) FROM eltit_ventas17.sv_ila_detalle_17 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas17.sv_ila_detalle_17 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0)+"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas18.sv_ila_detalle_18 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas18.sv_ila_detalle_18 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0) ) AS BOL"
            
            csql.sql = csql.sql & " FROM " & clientesistema & "gestion" & rubro_loc
            csql.sql = csql.sql & ".r_maestroproductos_fijo_" & rubro_loc & " AS mpf"
            
            csql.sql = csql.sql & " Where (mpf.codigoimpuesto='00001' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00002' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00003' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00006' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00007')"
            csql.sql = csql.sql & " GROUP BY codigobarra "
            csql.sql = csql.sql & " ORDER BY codigobarra DESC "
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
        Case "34"
            empresa = "42"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            '------
            empresa = "44"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            '------
            ProgressBar1.Value = ProgressBar1.Value + 1
            empresa = "45"
            basedatos = clientesistema & "ventas" & empresa
            rubro_loc = leerrubrocomercio(empresa)
            
            csql.sql = "TRUNCATE TABLE " & basedatos & ".sv_ila_detalle_" & empresa
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            csql.sql = " INSERT INTO " & basedatos & ".sv_ila_detalle_" & empresa
            csql.sql = csql.sql & " SELECT codigo,tipo,SUM(total) FROM " & basedatos
            csql.sql = csql.sql & ".sv_documento_detalle_" & empresa
            csql.sql = csql.sql & " WHERE MID(fecha,1,7) = '" & Format(fechasistema, "yyyy-mm") & "' "
            csql.sql = csql.sql & " GROUP BY tipo,codigo "
            csql.Execute
            '------
            ProgressBar1.Value = ProgressBar1.Value + 1
            csql.sql = "INSERT INTO " & tabla
            csql.sql = csql.sql & " SELECT mpf.codigobarra,mpf.codigoimpuesto,"
        
        
            csql.sql = csql.sql & "(IFNULL((SELECT SUM(total) FROM eltit_ventas42.sv_ila_detalle_42 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas42.sv_ila_detalle_42 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0)+"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas44.sv_ila_detalle_44 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas44.sv_ila_detalle_44 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0) +"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas45.sv_ila_detalle_45 AS lmd WHERE tipo='FV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas45.sv_ila_detalle_45 AS lmd WHERE tipo='NF' AND mpf.codigobarra=lmd.codigobarra),0) ) AS FAC,"
            
            csql.sql = csql.sql & " (IFNULL((SELECT SUM(total) FROM eltit_ventas42.sv_ila_detalle_42 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas42.sv_ila_detalle_42 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0)+"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas44.sv_ila_detalle_44 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas44.sv_ila_detalle_44 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0)+"
            
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas45.sv_ila_detalle_45 AS lmd WHERE tipo='BV' AND mpf.codigobarra=lmd.codigobarra),0)-"
            csql.sql = csql.sql & " IFNULL((SELECT SUM(total) FROM eltit_ventas45.sv_ila_detalle_45 AS lmd WHERE tipo='NB' AND mpf.codigobarra=lmd.codigobarra),0) ) AS BOL"
            
            csql.sql = csql.sql & " FROM " & clientesistema & "gestion" & rubro_loc
            csql.sql = csql.sql & ".r_maestroproductos_fijo_" & rubro_loc & " AS mpf"
            csql.sql = csql.sql & " Where (mpf.codigoimpuesto='00001' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00002' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00003' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00006' "
            csql.sql = csql.sql & " OR mpf.codigoimpuesto='00007')"
            csql.sql = csql.sql & " GROUP BY codigobarra "
            csql.sql = csql.sql & " ORDER BY codigobarra DESC "
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Value + 1
         Case Else
            Exit Sub
    End Select
PASO:
Set csql.ActiveConnection = contadb
            csql.sql = " SELECT ila.impuesto,mi.nombre,mi.contable,SUM(ila.factura),SUM(ila.boleta) ,"
            csql.sql = csql.sql & " (SELECT SUM(IF(fd.tipo<>'3' AND fd.tipo<>'6',fd.monto,fd.monto*-1)) "
            csql.sql = csql.sql & " FROM " & clientesistema & "conta" & empresaactiva & ".facturasdecompras AS fc"
            csql.sql = csql.sql & " INNER JOIN eltit_conta" & empresaactiva & ".facturasdecompras_detalle AS fd "
            csql.sql = csql.sql & " ON fc.tipo=fd.tipo AND fc.numero=fd.numero AND fc.rut=fd.rut "
            csql.sql = csql.sql & " AND fc.añocontable='" & Format(fechasistema, "YYYY") & "' AND fc.mescontable='" & Format(fechasistema, "MM") & "' "
            csql.sql = csql.sql & " Where fd.cuentadelmayor = mi.contable"
            csql.sql = csql.sql & " GROUP BY fd.cuentadelmayor) AS creditos"
            csql.sql = csql.sql & " , (SELECT SUM(IF(fd.tipo='1' OR fd.tipo='6',fd.monto, fd.monto*-1))  "
            csql.sql = csql.sql & " FROM " & clientesistema & "conta" & empresaactiva & ".facturasdeventas_detalle AS fd"
            csql.sql = csql.sql & " WHERE  fd.fechacreacion LIKE '" & Format(fechasistema, "yyyy-mm") & "%' "
            csql.sql = csql.sql & " AND  fd.cuentadelmayor = mi.contable2   AND (fd.tipo='1' OR fd.tipo='6' OR fd.tipo='4' OR fd.tipo='8')  GROUP BY fd.cuentadelmayor) AS debitos"
                       
'             csql.sql = csql.sql & ", (SELECT SUM(IF(fd.tipo=1 or fd.tipo=6,fd.monto, fd.monto*-1)) "
'            csql.sql = csql.sql & " FROM " & clientesistema & "conta" & empresaactiva & ".facturasdeventas AS fc"
'            csql.sql = csql.sql & " INNER JOIN eltit_conta" & empresaactiva & ".facturasdeventas_detalle AS fd "
'            csql.sql = csql.sql & " ON fc.tipo=fd.tipo AND fc.numero=fd.numero AND fc.rut=fd.rut "
'            csql.sql = csql.sql & " AND fc.fecha like '" & Format(fechasistema, "YYYY-mm") & "' and (fc.tipo='1' or fc.tipo='6' or fc.tipo='4' or fc.tipo='8')  "
'            csql.sql = csql.sql & " Where fd.cuentadelmayor = mi.contable"
'            csql.sql = csql.sql & " GROUP BY fd.cuentadelmayor) AS debitos "
            
            csql.sql = csql.sql & " FROM " & clientesistema & "conta.resumen_ilas AS ila"
            csql.sql = csql.sql & " INNER JOIN " & clientesistema & "gestion.g_maestroimpuestos AS mi"
            csql.sql = csql.sql & " ON mi.codigo=ila.impuesto"
            csql.sql = csql.sql & " GROUP BY ila.impuesto;"
                            
            csql.Execute
            ProgressBar1.Value = ProgressBar1.Max
            Grid1.Rows = 1
            Grid1.AutoRedraw = False
            ProgressBar1.Value = 0
            Dim factura As Double
            Dim boleta As Double
            Dim TOTALVENTA As Double
            Dim porce As Double
            Dim TOTAL1 As Double
            Dim total2 As Double
            Dim total3 As Double
            If csql.RowsAffected > 0 Then
            ProgressBar1.Max = csql.RowsAffected
            Set resultados = csql.OpenResultset
            While resultados.EOF = False
            ProgressBar1.Value = ProgressBar1.Value + 1
                Grid1.Rows = Grid1.Rows + 1
                    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
                    Grid1.Cell(Grid1.Rows - 1, 2).text = Format(resultados(5), "###,###,###,##0")
                    
                    factura = resultados(3)
                    boleta = resultados(4)
                    TOTALVENTA = factura + boleta
                    porce = factura / TOTALVENTA
           
                    Grid1.Cell(Grid1.Rows - 1, 3).text = Format(Round(resultados(5) * porce), "###,###,###,##0")
                    Grid1.Cell(Grid1.Rows - 1, 4).text = Round(porce * 100, 2)
                    Grid1.Cell(Grid1.Rows - 1, 5).text = Format(Round(resultados(6)), "###,###,###,##0")
                    If Grid1.Cell(Grid1.Rows - 1, 3).text = "" Then Grid1.Cell(Grid1.Rows - 1, 3).text = 0
                    If Grid1.Cell(Grid1.Rows - 1, 5).text = "" Then Grid1.Cell(Grid1.Rows - 1, 5).text = 0
                    TOTAL1 = TOTAL1 + Grid1.Cell(Grid1.Rows - 1, 3).text
                    total2 = total2 + Grid1.Cell(Grid1.Rows - 1, 5).text
                    
                resultados.MoveNext
            Wend
             
            End If
            
            Grid1.AddItem "", True
            Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTALES"
            Grid1.Cell(Grid1.Rows - 1, 3).text = Format(Round(TOTAL1), "###,###,###,##0")
            Grid1.Cell(Grid1.Rows - 1, 5).text = Format(Round(total2), "###,###,###,##0")
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
            Grid1.AutoRedraw = True
            Grid1.Refresh
            
End Sub




Sub Titulos(lista As Grid)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    lista.FixedRowColStyle = Fixed3D
    lista.CellBorderColorFixed = vbButtonShadow
    lista.ShowResizeTips = False
    lista.ReportTitles.Clear
    lista.PageSetup.CenterHorizontally = True
    lista.PageSetup.Orientation = cellPortrait
    
      
    lista.PageSetup.PrintTitleRows = 0
    
    'Logo
'    lista.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    lista.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    lista.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    lista.PageSetup.HeaderAlignment = CellLeft
    lista.PageSetup.HeaderFont.Name = "Verdana"
    lista.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CALCULO DE IMPUESTO ILA  |  " & "EMITIDO  :  " & Format(fechasistema, "dd-MM-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
    
      
    
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = tipoListado
'    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
'    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold = True
'    objReportTitle.Align = cellCenter
'    objReportTitle.PrintOnAllPages = True
'    lista.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    lista.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    lista.PageSetup.FooterAlignment = cellRight
    lista.PageSetup.FooterFont.Name = "Verdana"
    lista.PageSetup.FooterFont.Size = 7
    lista.PageSetup.LeftMargin = 0.5
    lista.PageSetup.RightMargin = 0.5
    lista.PageSetup.PrintFixedRow = True
End Sub
