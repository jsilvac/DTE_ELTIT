VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form proceso07 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " LISTADO DE LIQUIDACIONES DEL MES"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   6960
      TabIndex        =   17
      Top             =   8280
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
         TabIndex        =   19
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp frmcheque 
      Height          =   2535
      Left            =   5400
      TabIndex        =   7
      Top             =   4560
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   4471
      BackColor       =   16761024
      Caption         =   "PANTALLA DATOS DEL CHEQUE"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   4680
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "codigo"
         Top             =   405
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
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   11
         Top             =   405
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
         Left            =   3420
         MaxLength       =   4
         TabIndex        =   10
         Top             =   405
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "INICIAR PROCESO"
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
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox dato4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1530
         Width           =   1815
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BANCO"
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
         Left            =   990
         TabIndex        =   15
         Top             =   390
         Width           =   1455
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
         Left            =   90
         TabIndex        =   14
         Top             =   810
         Width           =   5145
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO INICIAL"
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
         Height          =   240
         Left            =   1800
         TabIndex        =   13
         Top             =   1260
         Width           =   1815
      End
   End
   Begin XPFrame.FrameXp FrameXP1 
      Height          =   8055
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   14208
      BackColor       =   16773879
      Caption         =   " LISTA DE LIQUIDACIONES"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7320
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "GENERAR INFORME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7335
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "GENERAR COMPROBANTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7320
         Width           =   3255
      End
      Begin FlexCell.Grid GRILLALIQUIDACION 
         Height          =   600
         Left            =   585
         TabIndex        =   3
         Top             =   7335
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   1058
         BackColor1      =   14737632
         BackColor2      =   14737632
         BackColorActiveCellSel=   16777088
         BackColorBkg    =   16761024
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         SelectionBorderColor=   16744576
         DefaultFontName =   "Arial"
         DefaultFontSize =   9.75
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   4
         DateFormat      =   2
      End
      Begin FlexCell.Grid GridLiquida 
         Height          =   6690
         Left            =   240
         TabIndex        =   1
         Top             =   405
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   11800
         BackColor1      =   16761024
         BackColor2      =   16761024
         BackColorActiveCellSel=   16777088
         BackColorBkg    =   16761024
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         SelectionBorderColor=   16744576
         Cols            =   3
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   4
         DateFormat      =   2
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "proceso07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ruttemporal As String
Dim cabezas1 As Variant
Dim fila As Integer
Dim columna As Integer
Dim i As Integer
Dim NCHEQUE As Double
Private TOTALCheque As Double
Private fechacheque As String
Private NOMBREGIRADO As String
Private FECHACONTABLE As String
Private numerocontable As String
Private tipocontable As String
Private lineacontable As Double
Private rutcontable As String
Private empresaconsulta As String

Private lineafinal As Double

 

Private Sub Command1_Click()
    IniciaGrid23
    empresaconsulta = empresaactiva
    Call Me.IniciaGrid1

End Sub

Private Sub COMMAND2_Click()
año = Format(fechasistema, "yyyy")
MES = Format(fechasistema, "mm")

If estacerrado(Format(fechasistema, "yyyy-mm-dd")) <> True Then
    If MsgBox("DESDEA PAGO MANUAL O ELECTRONICO,SI PARA MANUAL , NO PARA ELECTRONICO", vbYesNo, "ATENCION") = vbYes Then
        frmcheque.Visible = True
        dato1.SetFocus
    Else
        Call pagoelectronico
    End If
Else
    MsgBox "MES YA CERRADO"
End If
'GENERATXT
End Sub

Private Sub Command3_Click()

Call imprimir(GridLiquida)

End Sub
Sub imprimir(grilla As Grid)
Dim objReportTitle As FlexCell.ReportTitle
    
    
    grilla.FixedRowColStyle = Fixed3D
    grilla.CellBorderColorFixed = vbButtonShadow
    grilla.ShowResizeTips = False
    grilla.PageSetup.Orientation = cellPortrait
    
    
    
    
    
'    grilla.DefaultFont.Size = 7
    
    
    grilla.PageSetup.PrintFixedRow = True
    grilla.ReportTitles.Clear
    grilla.PageSetup.CenterHorizontally = False
    grilla.PageSetup.PrintTitleRows = 1
    grilla.PageSetup.BlackAndWhite = True
    
    
    'ENCABEZADO DE PAGINA
    
    grilla.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa
    grilla.PageSetup.HeaderAlignment = CellLeft
    grilla.PageSetup.HeaderFont.Name = "Verdana"
    grilla.PageSetup.HeaderFont.Size = 12
    grilla.PageSetup.HeaderFont.Bold = True
    grilla.PageSetup.HeaderMargin = 0.5
    
    
    
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE LIQUIDACIONES EMITIDAS PARA PAGO"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    grilla.ReportTitles.Add objReportTitle
    
   
    'PIE DE PAGINA
    grilla.PageSetup.LeftMargin = 0.5
    grilla.PageSetup.RightMargin = 0.1
    grilla.PageSetup.TopMargin = 3
    grilla.PageSetup.BottomMargin = 0.5
    
    
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeLeft) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeTop) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeBottom) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeRight) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellInsideHorizontal) = cellNone

    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellInsideVertical) = cellNone
    
    grilla.PrintPreview
    
    
    
    
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato4)
If leercheque(dato1.text + dato2.text + dato3.text, dato4.text) = True Then
MsgBox ("EL NUMERO DE CHEQUE YA ESTA EMITIDO")
dato4.text = ""
dato4.SetFocus
Else

Command5.SetFocus
End If

End If

End Sub

Private Sub Command5_Click()
    
Dim k As Double
Dim rutprove As String
Dim tipo As String
Dim CUENTABANCO As String
Dim fechavencimiento  As String
Dim monto As Double
Dim DH As String


If lblBanco.Caption <> "" Then
If empresaconsulta <> empresaactiva Then MsgBox "EL LISTADO ES DE OTRA EMPRESA " & empresaconsulta, vbCritical, "ATENCION": GoTo no:

NCHEQUE = CDbl(dato4.text) - 1

  tipocontable = "CE"
        numerocontable = LEERFOLIOCE("CE")
        lineacontable = 0
        TOTALCheque = 0
For k = 1 To GridLiquida.Rows - 1
   
'        If rutprove <> Mid(GridLiquida.Cell(k, 3).text, 1, 9) + Mid(GridLiquida.Cell(k, 3).text, 11, 1) Then
'        If TOTALCheque <> 0 Then
'        Call grabarcheque(TOTALCheque)
'        End If
       
        
'        End If
 If GridLiquida.Cell(k, 9).text = "1" Then
         fechacheque = Format(fechasistema, "yyyy-mm-dd")
        NOMBREGIRADO = GridLiquida.Cell(k, 2).text
        FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
      
        rutprove = Mid(GridLiquida.Cell(k, 1).text, 1, 9) + Mid(GridLiquida.Cell(k, 1).text, 10, 1)
        rutcontable = rutprove
        
        CUENTABANCO = "23100028"
        fechavencimiento = Format(fechasistema, "yyyy-mm-dd")
        monto = CDbl(GridLiquida.Cell(k, 8).text)
        DH = "D"
       
    If verificasiexiste2(rutcontable, CUENTABANCO, FECHACONTABLE, tipocontable, "CANC. SUELDO ") = False Then
        lineacontable = lineacontable + 1
        Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, _
        FECHACONTABLE, CUENTABANCO, " ", rutcontable, " ", "CANC. SUELDO " & NOMBREGIRADO, _
        tipocontable, numerocontable, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, _
        Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), _
        Format(Date, "yyyy-mm-dd"), Time, rutcontable)
        TOTALCheque = TOTALCheque + monto
    Else
        MsgBox "YA EXISTE COMPROBANTE ", vbCritical, "ATENCION"
        frmcheque.Visible = False
        dato4.text = ""
        Exit Sub
    End If
        
'        Call GRABARCOMPROBANTE(GridLiquida.Cell(k, 2).text, GridLiquida.Cell(k, 9).text, GridLiquida.Cell(k, 10).text, Format(GridLiquida.Cell(k, 15).text, "yyyy-mm-dd"), GridLiquida.Cell(k, 1).text, GridLiquida.Cell(k, 16).text)
        
    End If
Next k
    
        If TOTALCheque <> 0 Then
        Call grabarcheque(TOTALCheque)
        TOTALCheque = 0
        End If
        
'leer

End If
no:
frmcheque.Visible = False
dato4.text = ""
End Sub
 

Private Sub Form_Activate()
'sqlremu.audit = True
'sqlremu.programaactivo = Me.Caption
 If Verifica_Permiso(Me.Caption, "autoriza") = False Then
        Command2.Enabled = False
  End If
End Sub

'***********************************************************************
'***********************************************************************
Private Sub Form_Load()
'    Call configuracion.Conectar_BD 'db
'    Call configuracion.ConectarRemu(servidor, clientesistema + "remu", usuario, password) 'remu
    MODIFI = 0
    '----------------------------------------------
  cabezas1 = Array("RUT", "NOMBRE", "TIPO", "BANCO", "CUENTA", "LIQUIDO", "CONVENIOS", "A PAGAR", "REVISA", "")
    Call CargaGrilla1(1, 10, GridLiquida, cabezas1)
  frmcheque.Visible = False
 
  Call CENTRAR(Me)
End Sub

Private Sub GridLiquida_DblClick()
    If GridLiquida.Cell(GridLiquida.ActiveCell.row, GridLiquida.Cols - 1).text = "0" Then
        GridLiquida.Cell(GridLiquida.ActiveCell.row, GridLiquida.Cols - 1).text = "1"
    Else
        GridLiquida.Cell(GridLiquida.ActiveCell.row, GridLiquida.Cols - 1).text = "0"
    End If
      If GridLiquida.Cell(GridLiquida.ActiveCell.row, GridLiquida.Cols - 2).text = "" Then
    If GridLiquida.Cell(GridLiquida.ActiveCell.row, GridLiquida.Cols - 1).text = "1" Then
        MsgBox "NO PUEDE SELECCIONAR SI NO TIENE TOTAL A PAGAR", vbCritical, "ATENCION"
        GridLiquida.Cell(GridLiquida.ActiveCell.row, GridLiquida.Cols - 1).text = "0"
    End If
  End If
  If GridLiquida.ActiveCell.col = 9 Then
        If Verifica_Permiso(Me.Caption, "autoriza") = True Then
            Call grabarrevisado(Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), GridLiquida.Cell(GridLiquida.ActiveCell.row, 1).text, empresaactiva, GridLiquida.Cell(GridLiquida.ActiveCell.row, 9).text, Format(fechasistema, "yyyy-mm-dd"))
'            Call IniciaGrid1
        End If
  End If
End Sub

'************************************************************************
'************************************************************************
Private Sub MANUAL_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27 'esc
            Unload Me
    End Select
End Sub

'************************************************************************
'************************************************************************
Private Sub GridLiquida_Click()


  
End Sub

Private Sub GridLiquida_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    fila = GridLiquida.ActiveCell.row
    columna = GridLiquida.ActiveCell.col
    Select Case KeyCode
        Case 27 'esc
            MANUAL.SetFocus
        Case 77 'M:modificar
        Case 46 'suprimir/eliminar
    End Select
End Sub

'************************************************************************
'************************************************************************
Sub IniciaGrid1()
    Call CargaGridLiquida(GridLiquida)
End Sub

Sub CargaGridLiquida(grilla As Grid)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Dim LIQUIDO As Double
    Dim CONVENIO As Double
    Dim saldo As Double
    Dim TOTALB1 As Double
    Dim TOTALB2 As Double
    Dim TOTALB3 As Double
    Dim TOTALN1 As Double
    Dim TOTALN2 As Double
    Dim TOTALN3 As Double
    
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT mt.rut,mt.nombre,cb.tipo,cb.banco,cb.cuenta "
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".mt_fijo as mt LEFT join " + clientesistema + "remu.cuentasbancarias as cb on mt.rut=cb.rut "
    csql.sql = csql.sql + "where mt.mes='" + Format(fechasistema, "mm") + "' and mt.año='" + Format(fechasistema, "yyyy") + "' and (mid(mt.fecharet,1,7)>='" + Format(fechasistema, "yyyy-mm") + "' or fecharet='0000-00-00') AND cb.tipo <> '0'"
                        'arielito
    
    csql.sql = csql.sql + " ORDER BY nombre "
    csql.Execute
    TOTALB1 = 0
    TOTALB2 = 0
    TOTALB3 = 0
    TOTALN1 = 0
    TOTALN2 = 0
    TOTALN3 = 0
    grilla.AutoRedraw = False
    
    
    grilla.Rows = 1
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
        While Not resultados.EOF
             LIQUIDO = leercalculo(resultados(0), Format(fechasistema, "mm"), Format(fechasistema, "YYYY"), "LI001")
            CONVENIO = totalconvenios(resultados(0))
            saldo = LIQUIDO - CONVENIO
            If saldo > 0 Then
             grilla.Rows = grilla.Rows + 1
            LINEA = grilla.Rows - 1
            grilla.Cell(LINEA, 1).text = resultados(0)
            grilla.Cell(LINEA, 2).text = resultados(1)
            If IsNull(resultados(2)) = False Then
            grilla.Cell(LINEA, 3).text = resultados(2)
            grilla.Cell(LINEA, 4).text = resultados(3)
            grilla.Cell(LINEA, 5).text = resultados(4)
            End If
'            Call calculaliquidaciones.calculaliquidaciones(resultados(0), Format(fechasistema, "mm"), Format(fechasistema, "YYYY"), empresaactiva)
           
            grilla.Cell(LINEA, 6).text = Format(LIQUIDO, "###,###,###")
            grilla.Cell(LINEA, 7).text = Format(CONVENIO, "###,###,###")
            grilla.Cell(LINEA, 8).text = Format(saldo, "###,###,###")
            grilla.Cell(LINEA, 9).text = leerestado(empresaactiva, Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), resultados(0))
            If IsNull(resultados(2)) = False Then
            TOTALB1 = TOTALB1 + LIQUIDO
            TOTALB2 = TOTALB2 + CONVENIO
            TOTALB3 = TOTALB3 + saldo
            Else
            TOTALN1 = TOTALN1 + LIQUIDO
            TOTALN2 = TOTALN2 + CONVENIO
            TOTALN3 = TOTALN3 + saldo
            
            
            End If
            End If
            resultados.MoveNext
            LINEA = LINEA + 1
        Wend
        resultados.Close
        Set resultados = Nothing
    End If
    
    grilla.Rows = grilla.Rows + 1
    grilla.Range(grilla.Rows - 1, 1, grilla.Rows - 1, grilla.Cols - 1).Borders(cellEdgeTop) = cellThin
    grilla.Range(grilla.Rows - 1, 1, grilla.Rows - 1, grilla.Cols - 1).FontBold = True
    
    
    
    
    grilla.Cell(grilla.Rows - 1, 2).text = "TOTAL SUELDOS A PAGAR X BANCO "
    grilla.Cell(grilla.Rows - 1, 6).text = Format(TOTALB1, "#,###,###,###")
    grilla.Cell(grilla.Rows - 1, 7).text = Format(TOTALB2, "#,###,###,###")
    grilla.Cell(grilla.Rows - 1, 8).text = Format(TOTALB3, "#,###,###,###")
    grilla.Rows = grilla.Rows + 1
    grilla.Range(grilla.Rows - 1, 1, grilla.Rows - 1, grilla.Cols - 1).FontBold = True
    
    grilla.Cell(grilla.Rows - 1, 2).text = "TOTAL SUELDOS A PAGAR DIRECTOS "
    grilla.Cell(grilla.Rows - 1, 6).text = Format(TOTALN1, "#,###,###,###")
    grilla.Cell(grilla.Rows - 1, 7).text = Format(TOTALN2, "#,###,###,###")
    grilla.Cell(grilla.Rows - 1, 8).text = Format(TOTALN3, "#,###,###,###")
    
    grilla.AutoRedraw = True
    grilla.Refresh
End Sub

Sub CargaGrilla1(numRow, numCol, grilla As Grid, camposgrid As Variant)
    Dim i As Integer
    Dim FORMATOGRILLA(50, 50) As String
    
    i = 0
    While (camposgrid(i) <> "")
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, i + 1) = camposgrid(i)
        i = i + 1
    Wend
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "6"
    'formatogrilla(2, 2) = "25"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE
    FORMATOGRILLA(3, 1) = "N"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    
    'formatogrilla(3, 2) = "S"
    Rem FORMATO GRILLA
    For i = 1 To 6
        FORMATOGRILLA(4, i) = ""
    Next i
    Rem LOCCKED
    For i = 1 To 8
        FORMATOGRILLA(5, i) = "TRUE"
    Next i
    Rem ancho
    FORMATOGRILLA(6, 1) = "9"
    FORMATOGRILLA(6, 2) = "30"
    FORMATOGRILLA(6, 3) = "5"
    FORMATOGRILLA(6, 4) = "5"
    FORMATOGRILLA(6, 5) = "15"
    FORMATOGRILLA(6, 6) = "8"
    FORMATOGRILLA(6, 7) = "8"
    FORMATOGRILLA(6, 8) = "8"
    FORMATOGRILLA(6, 9) = "8"
    
    With grilla
        .Cols = numCol
        .Rows = numRow
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .ExtendLastCol = False
        .BoldFixedCell = False
        .DrawMode = cellOwnerDraw
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        For k = 1 To numCol - 1
            .Cell(0, k).text = FORMATOGRILLA(1, k)
            .Column(k).Width = Val(FORMATOGRILLA(6, k)) * .Cell(0, k).Font.Size + 1.25
         
            .Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
            .Column(k).FormatString = FORMATOGRILLA(4, k)
            .Column(k).Locked = True 'formatogrilla(5, K)
            If FORMATOGRILLA(3, k) = "N" Then
                .Column(k).Alignment = cellRightCenter
                
                .Column(k).Mask = cellNumeric
            End If
            If FORMATOGRILLA(3, k) = "S" Then
                .Column(k).Alignment = cellLeftCenter
                .Column(k).Mask = cellUpper
            End If
        Next k
        .Range(0, 0, 0, .Cols - 1).Alignment = cellCenterCenter
    .Column(0).Width = 0
    .Column(9).CellType = cellCheckBox
    If Verifica_Permiso(Me.Caption, "autoriza") = True Then
        .Column(9).Locked = False
    Else
        .Column(9).Locked = True
    End If
    
    
    
    
    End With '//grilla

End Sub

Sub IniciaGrid23()
    Dim CABEZAS2 As Variant
    
    CABEZAS2 = Array("HABERES", "BASE", "MONTO", "DESCUENTOS", "BASE", "MONTO", "")
    Call CargaGrilla23(1, 7, GRILLALIQUIDACION, CABEZAS2)
    'cabezas2 = Array("DESCUENTOS", "BASE CAL.", "MONTO", "")
    'Call CargaGrilla23(1, 4, GridDescuentos, cabezas2)
End Sub

Sub CargaGrilla23(numRow, numCol, grilla As Grid, camposgrid As Variant)
    Dim i As Integer
    Dim FORMATOGRILLA(50, 50) As String
    
    i = 0
    While (camposgrid(i) <> "")
        FORMATOGRILLA(1, i + 1) = camposgrid(i) 'encabezados
        i = i + 1
    Wend
    Rem LARGO DE LOS DATOS
    For i = 1 To 6
        FORMATOGRILLA(2, i) = "10"
    Next i
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    Rem FORMATO GRILLA
    For i = 1 To 9
        FORMATOGRILLA(4, i) = ""
    Next i
    Rem LOCCKED
    For i = 1 To 9
        FORMATOGRILLA(5, i) = "FALSE"
    Next i
    Rem ancho
    FORMATOGRILLA(6, 1) = "30"
    FORMATOGRILLA(6, 2) = "7"
    FORMATOGRILLA(6, 3) = "10"
    FORMATOGRILLA(6, 4) = "30"
    FORMATOGRILLA(6, 5) = "7"
    FORMATOGRILLA(6, 6) = "10"
    grilla.Column(0).Width = 0
    
    With grilla
        .Cols = numCol
        .Rows = numRow
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .ExtendLastCol = False
        .BoldFixedCell = False
        .DrawMode = cellOwnerDraw
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .Column(0).Width = 0
        For k = 1 To numCol - 1
            .Cell(0, k).text = FORMATOGRILLA(1, k)
            .Column(k).Width = Val(FORMATOGRILLA(6, k)) * .Cell(0, k).Font.Size + 1.25
            .Column(0).Width = 0
            .Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
            .Column(k).FormatString = FORMATOGRILLA(4, k)
            .Column(k).Locked = True 'formatogrilla(5, K)
            If FORMATOGRILLA(3, k) = "N" Then
                .Column(k).Alignment = cellRightCenter
                .Column(k).Mask = cellNumeric
            End If
            If FORMATOGRILLA(3, k) = "S" Then
                .Column(k).Alignment = cellLeftCenter
                .Column(k).Mask = cellUpper
            End If
            If FORMATOGRILLA(3, k) = "D" Then
                .Column(k).CellType = cellCalendar
                .Column(k).Mask = cellNumeric
            End If
            '.Column(7).CellType = cellComboBox
        Next k
        '.Range(0, 1, 0, 3).Merge
        '.Cell(0, 1).text = "CUENTA"
        .Range(0, 0, 0, .Cols - 1).Alignment = cellCenterCenter
    End With '//grilla
End Sub
'Sub liquidacion()
'    Dim W As Integer
'    Dim contadord As Integer
'    Dim contadorh As Integer
'    GRILLALIQUIDACION.Rows = 1
'
'    GRILLALIQUIDACION.Rows = 18
'    contadord = 0
'    contadorh = 0
'
'    For W = 1 To LINEAC
'        Select Case CALCULOS(W, 6)
'        Case "H"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadord = contadord + 1
'                If CALCULOS(W, 1) = "TOTAL HABERES GENERALES " Then
'                contadord = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadord, 1).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'                GRILLALIQUIDACION.Cell(contadord, 2).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadord, 3).text = Format(CALCULOS(W, 3), "###,###,###")
'                If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadord, 1, contadord, 3).FontBold = True
'                End If
'
'
'            End If
'        Case "D"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadorh = contadorh + 1
'                If CALCULOS(W, 1) = "TOTAL DESCUENTOS GENERALES " Then
'                contadorh = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadorh, 4).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'
'                GRILLALIQUIDACION.Cell(contadorh, 5).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadorh, 6).text = Format(CALCULOS(W, 3), "###,###,###")
'
'            If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadorh, 4, contadorh, 6).FontBold = True
'                End If
'
'
'            End If
'        End Select
'    Next W
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeBottom) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeLeft) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeTop) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeRight) = cellThin
'GRILLALIQUIDACION.Range(13, 3, 13, 3).Borders(cellEdgeRight) = cellThin
'For W = 1 To 6
'GRILLALIQUIDACION.Column(W).Locked = False
'Next W
'
'GRILLALIQUIDACION.Range(15, 1, 15, 6).Merge
'GRILLALIQUIDACION.Cell(15, 1).text = "SON :" + WORDNUM(GRILLALIQUIDACION.Cell(14, 6).text, "PESO", "PESOS", 0)
'For W = 1 To 6
'GRILLALIQUIDACION.Column(W).Locked = True
'Next W
'GRILLALIQUIDACION.Cell(17, 4).text = "RECIBI CONFORME ____________________ "
'
'End Sub
''
''Sub liquidacion()
'    Dim W As Integer
'    Dim contadord As Integer
'    Dim contadorh As Integer
'    GRILLALIQUIDACION.Rows = 1
'
'    GRILLALIQUIDACION.Rows = 15
'    contadord = 0
'    contadorh = 0
'
'    For W = 1 To LINEAC
'        Select Case CALCULOS(W, 6)
'        Case "H"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadord = contadord + 1
'                If CALCULOS(W, 1) = "TOTAL HABERES GENERALES " Then
'                contadord = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadord, 1).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'                GRILLALIQUIDACION.Cell(contadord, 2).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadord, 3).text = Format(CALCULOS(W, 3), "###,###,###")
'                If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadord, 1, contadord, 3).FontBold = True
'                End If
'
'
'            End If
'        Case "D"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadorh = contadorh + 1
'                If CALCULOS(W, 1) = "TOTAL DESCUENTOS GENERALES " Then
'                contadorh = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadorh, 4).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'
'                GRILLALIQUIDACION.Cell(contadorh, 5).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadorh, 6).text = Format(CALCULOS(W, 3), "###,###,###")
'
'            If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadorh, 4, contadorh, 6).FontBold = True
'                End If
'
'
'            End If
'        End Select
'    Next W
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeBottom) = cellThin
'
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeLeft) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeTop) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeRight) = cellThin
'GRILLALIQUIDACION.Range(13, 3, 13, 3).Borders(cellEdgeRight) = cellThin
'GRILLALIQUIDACION.Cell(14, 1).text = WORDNUM(Format(dato6.text, "########0"), "PESO", "PESOS", 0)
'
'
'
'End Sub

Sub GENERATXT()
Dim k As Double
Dim s As Double
Dim L As Double

Dim MATRIX(11, 3) As String
Dim contador As Double
Dim signo As String
Dim VARIABLE As String
Dim BLANCO As String * 9
Dim rut As String * 11
Dim NOMBRE As String * 30
Dim direccion As String * 30
Dim comuna As String * 15
Dim ciudad As String * 15

'MODALIDAD DE PAGO       ( 1)  ( 1) ok
'BANCO           ( 2)  ( 3) ok
'NRO CTA(5)(18) ok
'SI MOD PAGO=1 THEN '999'(23)  ( 3) ok
'monto PAGO(26)(8) ok
'RUT                     (34)  (11) ok
'APELLIDOS Nombres(45)(30)
'FEC NAC(75)(10)
'DIRECCION               (85)  (30)

Dim MODALIDAD As String * 1
Dim SUCURSAL As String * 3
Dim cuenta As String * 18
Dim banco As String * 3
Dim RUTRETIRA As String * 12
Dim NOMBRERETIRA As String * 30
Dim numerodo As String * 8
Dim MONTODO As String * 11
Dim SIGNODO As String * 1
Dim TOTALGENERAL As String * 8
Dim sumatotal As Double
Dim modopago As String * 3
Dim fechanacimiento As String * 10


Close 20
On Error GoTo error:
Open "c:\sueldos\sueldos_" + empresaactiva + "_" + Format(fechasistema, "dd-mm-yyyy") + ".TXT" For Output As #20

contador = 0
For k = 1 To GridLiquida.Rows - 3
If GridLiquida.Cell(k, 8).text <> "" Then
    If GridLiquida.Cell(k, 9).text = "1" Then

        rut = "0" + GridLiquida.Cell(k, 1).text
        NOMBRE = GridLiquida.Cell(k, 2).text

        TOTALGENERAL = Format(CDbl(GridLiquida.Cell(k, 8).text), "00000000")
        banco = GridLiquida.Cell(k, 4).text
        cuenta = GridLiquida.Cell(k, 5).text
        direccion = String(40, 32)
        comuna = String(40, 32)
        ciudad = String(40, 32)
        MODALIDAD = GridLiquida.Cell(k, 3).text

        direccion = leerdirecciontrabajador(GridLiquida.Cell(k, 1).text)
        fechanacimiento = Replace(leerfechanacimiento(GridLiquida.Cell(k, 1).text), "-", "/")

        If MODALIDAD = "1" Then
            modopago = "999"
        Else
            modopago = "   "
        End If

        VARIABLE = MODALIDAD + banco + cuenta + modopago + TOTALGENERAL + rut + NOMBRE + fechanacimiento + direccion + ciudad + ""
        If MODALIDAD <> "0" Then
            Print #20, VARIABLE
        End If
    End If
End If
Next k
Close 20

Shell "notepad c:\sueldos\sueldos_" + empresaactiva + "_" + Format(fechasistema, "dd-mm-yyyy") + ".TXT"
Exit Sub
error:
MsgBox "la carpera c:\sueldos\ no esta creada en este equipo"
End Sub

Function leerdirecciontrabajador(rut) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select direccion,ciudad from mt_fijo where rut ='" & rut & "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerdirecciontrabajador = resultados(0) & " " & resultados(1)
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
End Function
Function leerfechanacimiento(rut) As String
        Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select ifnull(fechanac,'') from mt_fijo where rut ='" & rut & "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerfechanacimiento = Format(resultados(0), "dd/mm/yyyy")
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
End Function
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
        Call flechas(dato2, dato4, KeyCode)
    End Sub
'*********************************************
'KEYDOWN
'****************************************************************************

'****************************************************************************
'KEYPRESS
'****************************************************************************
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
           
          dato2.SetFocus
          
           
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
           dato3.SetFocus
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            lblBanco.Caption = leerNombreCuentaMayor(dato1.text & dato2.text & dato3.text, 3)
            If lblBanco.Caption <> "" Then
                
            dato4.SetFocus
            End If
        
        End If
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
Sub grabarcheque(montocheque As Double)
Dim tipodocumento As String
Dim numerodocumento As String
Dim CUENTABANCO As String
Dim fechavencimiento As String
Dim monto As Double
Dim DH As String



    Rem graba cheque
        
        NCHEQUE = NCHEQUE + 1
        lineacontable = lineacontable + 1
        
        If tipocontable = "CE" Then
            tipodocumento = "CH"
            numerodocumento = Format(NCHEQUE, "0000000000")
            CUENTABANCO = dato1.text + dato2.text + dato3.text
            fechavencimiento = fechacheque
            monto = montocheque
            
            
            Else
            
            tipodocumento = "DB"
            numerodocumento = Format(numerocontable, "0000000000")
            CUENTABANCO = "11130001"
            fechavencimiento = fechacheque
            monto = montocheque
            
        End If
        
        DH = "H"
        NOMBREGIRADO = nombreempresa
        Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, FECHACONTABLE, CUENTABANCO, " ", "", " ", NOMBREGIRADO, tipodocumento, numerodocumento, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), Format(Date, "yyyy-mm-dd"), Time, rutcontable)
        If tipocontable = "CE" Then
        fecha = Format(fechasistema, "yyyy-mm-dd")
        Call grabacheque(CUENTABANCO, numerodocumento, fecha, monto, fechavencimiento, tipocontable, numerocontable, NOMBREGIRADO, "0")
        End If
End Sub
Public Function LEERFOLIOCE(tipo) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = contadb
            csql.sql = "select max(numero) from movimientoscontables where mes = '" & Format(Format(fechasistema, "mm"), "00") & "' AND año = '" & Format(fechasistema, "yyyy") & "' and tipo='" + tipo + "' "
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
        If IsNull(resultados(0)) = False Then
        LEERFOLIOCE = Format(resultados(0) + 1, "0000000000")
        Else
        LEERFOLIOCE = Format(1, "0000000000")
        End If
        
    End If
    
End Function
Sub GRABARCOMPROBANTE(ORDEN, montocheque, DIFERENCIA, fechacheque, TIPOPAGO, glosadiferencia)
    Dim DH As String
    Dim numero As String
    Dim LINEA As Double
    Dim fecha As Date
    Dim rut As String
    Dim tipodocumento As String
    Dim numerodocumento As String
    Dim fechadocumento As String
    Dim fechavencimiento As String
    Dim MES As String
    Dim año As String
    Dim monto As Double
    Dim CUENTABANCO As String
   
    Dim tipo2 As String
    Dim TIPO3 As String
    Dim DOCUMENTOPAGO As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery

        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT tipo,total,rut,numero,fecha "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_detalle_facturas_" + empresaactiva + " WHERE ordendecompra='" + ORDEN + "' "
        csql.sql = csql.sql + "ORDER BY ordendecompra "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            DH = "D"
            If resultados(0) = "NC" Or resultados(0) = "NCE" Then
            DH = "H"
            End If
            fecha = Format(fechasistema, "yyyy-mm-dd")
            rut = resultados(2)
            tipodocumento = resultados(0)
            TIPO3 = resultados(0)
            If tipodocumento = "FA" Then tipo2 = "1": TIPO3 = "FC"
            If tipodocumento = "ND" Then tipo2 = "2": TIPO3 = "ND"
            If tipodocumento = "NC" Then tipo2 = "3": TIPO3 = "NC"
            If tipodocumento = "FAE" Then tipo2 = "4": TIPO3 = "FC"
            If tipodocumento = "NDE" Then tipo2 = "5": TIPO3 = "ND"
            If tipodocumento = "NCE" Then tipo2 = "6": TIPO3 = "NC"
            If Mid(TIPOPAGO, 1, 2) = "CH" Then
            DOCUMENTOPAGO = "PA"
            Else
            DOCUMENTOPAGO = "DB"
            End If
            
            tipodocumento = TIPO3
            numerodocumento = resultados(3)
            fechadocumento = resultados(4)
            fechavencimiento = fechadocumento
            MES = Format(fechasistema, "mm")
            año = Format(fechasistema, "yyyy")
            monto = resultados(1)
            
            If DH = "D" Then
            TOTALCheque = TOTALCheque + monto
            Else
            TOTALCheque = TOTALCheque - monto
            End If
            
             lineacontable = lineacontable + 1
            Call grabarcomprobante_lineas(DOCUMENTOPAGO, numerocontable, lineacontable, fecha, CUENTAPROVEEDOR, " ", rut, " ", "CANCELA DOCUMENTO", tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rut)
            
'            Call abonofactura(tipo2, numerodocumento, rut, monto)
           
            resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
   Rem graba linea  diferencia
        
        If DIFERENCIA <> 0 Then
           
        lineacontable = lineacontable + 1
            
            monto = DIFERENCIA
            DH = "H"
            If DH = "D" Then
            TOTALCheque = TOTALCheque + monto
            Else
            TOTALCheque = TOTALCheque - monto
            End If
            
            Call grabarcomprobante_lineas(DOCUMENTOPAGO, numerocontable, lineacontable, fecha, cuentadiferencia, " ", rut, " ", glosadiferencia, "OC", ORDEN, fechadocumento, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rut)
        End If
        lineafinal = lineacontable
'        Call leerotros(DOCUMENTOPAGO, numerocontable, FECHA, "00", rut, ORDEN, lineacontable, mes, año)
        
        lineacontable = lineafinal
End Sub
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If rutctacte <> "" Then
        Call existerut(año, codigocuenta, rutctacte, empresaactiva)
    End If
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub
Sub grabacheque(cuenta, numero, emision, monto, vencimiento, tipocomprobante, numerocomprobante, giradoa, ubicacion)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "ubicacion"
    campos(9, 0) = ""
    
    campos(0, 1) = cuenta
    campos(1, 1) = numero
    campos(2, 1) = emision
    campos(3, 1) = monto
    campos(4, 1) = vencimiento
    campos(5, 1) = tipocomprobante
    campos(6, 1) = numerocomprobante
    campos(7, 1) = giradoa
    campos(8, 1) = "0"
    campos(0, 2) = "chequesdocumento"
       
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
End Sub

Sub grabarrevisado(MES, año, ruttrabajador, empresa, estado, fecha)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "delete from " & clientesistema & "remu.mt_fijo_revisado "
    csql.sql = csql.sql & "where empresa='" & empresa & "' and año='" & año & "' "
    csql.sql = csql.sql & "and mes='" & MES & "' and ruttrabajador='" & ruttrabajador & "' "
    csql.Execute
'    Call sincronizadatos(csql.sql, contadb, "")
    
    
    csql.sql = "insert into " & clientesistema & "remu.mt_fijo_revisado (empresa,año,mes,ruttrabajador,estado,fecha)"
    csql.sql = csql.sql & "values('" & empresa & "','" & año & "','" & MES & "','" & ruttrabajador & "','" & estado & "','" & fecha & "') "
    csql.Execute
    
'    Call sincronizadatos(csql.sql, contadb, "")
    csql.Close
    Set csql = Nothing
    
    
End Sub
Function leerestado(empresaconsulta, MES, año, ruttrabajador) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select estado "
    csql.sql = csql.sql & "from " & clientesistema & "remu.mt_fijo_revisado "
    csql.sql = csql.sql & "where empresa='" & empresaconsulta & "' and año='" & año & "' "
    csql.sql = csql.sql & "and mes='" & MES & "' and ruttrabajador='" & ruttrabajador & "' "
    csql.Execute
        leerestado = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerestado = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
    
End Function


Public Sub existerut(año, tipo, rut, empresa)
 
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = nombretraba(rut, empresa)
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' and año='" + año + "'  "
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    Else
    
    Call grabar(año, tipo, rut, nombretraba(rut, empresa), empresa)
    
    End If

    
    End Sub
Public Function nombretraba(rut, empresa) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' "
    campos(0, 2) = clientesistema + "remu" + empresa + ".mt_fijo "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    nombretraba = sqlconta.response(0, 3)
    Else
    nombretraba = ""
    End If
    End Function

Sub grabar(año, tipo, rut, NOMBRE, empresa)
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = NOMBRE
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
     Call grabar2(año, tipo, rut, empresa)
    
    End Sub
Sub grabar2(año, tipo, rut, empresa)
      
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = ""
    
    campos(0, 1) = año
    campos(1, 1) = tipo
    campos(2, 1) = rut
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".saldosctacte"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
Sub pagoelectronico()
    Dim k As Double
Dim rutprove As String
Dim tipo As String
Dim CUENTABANCO As String
Dim fechavencimiento  As String
Dim monto As Double
Dim DH As String
 
    If empresaconsulta <> empresaactiva Then MsgBox "EL LISTADO ES DE OTRA EMPRESA " & empresaconsulta, vbCritical, "ATENCION": GoTo no:
       
        tipocontable = "CE"
        numerocontable = LEERFOLIOCE("CE")
         NCHEQUE = numerocontable
        lineacontable = 0
        TOTALCheque = 0
        For k = 1 To GridLiquida.Rows - 1
            If GridLiquida.Cell(k, 9).text = "1" Then
                fechacheque = Format(fechasistema, "yyyy-mm-dd")
                NOMBREGIRADO = GridLiquida.Cell(k, 2).text
                FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
                rutprove = Mid(GridLiquida.Cell(k, 1).text, 1, 9) + Mid(GridLiquida.Cell(k, 1).text, 10, 1)
                rutcontable = rutprove
                CUENTABANCO = "23100028"
                fechavencimiento = Format(fechasistema, "yyyy-mm-dd")
                monto = CDbl(GridLiquida.Cell(k, 8).text)
                DH = "D"
                If verificasiexiste2(rutcontable, CUENTABANCO, FECHACONTABLE, tipocontable, "CANC. SUELDO ") = False Then
                    lineacontable = lineacontable + 1
                    Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, _
                    FECHACONTABLE, CUENTABANCO, " ", rutcontable, " ", "CANC. SUELDO " & NOMBREGIRADO, _
                    tipocontable, numerocontable, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, _
                    Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), _
                    Format(Date, "yyyy-mm-dd"), Time, rutcontable)
                    TOTALCheque = TOTALCheque + monto
                Else
                    MsgBox "YA EXISTE COMPROBANTE ", vbCritical, "ATENCION"
                    frmcheque.Visible = False
                    dato4.text = ""
                    Exit Sub
                End If
            End If
        Next k
    
        If TOTALCheque <> 0 Then
        Call grabarcheque2(TOTALCheque)
        TOTALCheque = 0
        End If
        
'leer
MsgBox "COMPROBANTE CE N°" & numerocontable & " GENERADO EXITASAMENTE ", vbInformation, "ATENCION"
  
no:
frmcheque.Visible = False
dato4.text = ""
End Sub
Sub grabarcheque2(montocheque As Double)
Dim tipodocumento As String
Dim numerodocumento As String
Dim CUENTABANCO As String
Dim fechavencimiento As String
Dim monto As Double
Dim DH As String



    Rem graba cheque
        
        NCHEQUE = NCHEQUE + 1
        lineacontable = lineacontable + 1
        
        
            tipodocumento = "DB"
            numerodocumento = Format(numerocontable, "0000000000")
            CUENTABANCO = "11500160"
            fechavencimiento = fechacheque
            monto = montocheque
            
       
        DH = "H"
        NOMBREGIRADO = nombreempresa
        Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, FECHACONTABLE, CUENTABANCO, " ", "", " ", NOMBREGIRADO, tipodocumento, numerodocumento, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), Format(Date, "yyyy-mm-dd"), Time, rutcontable)
 
End Sub
