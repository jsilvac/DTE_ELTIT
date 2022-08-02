VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form muestraorden 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUESTRA ORDENES DE COMPRA"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15015
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   619
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1001
   Begin XPFrame.FrameXp bodega 
      Height          =   2400
      Left            =   3555
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4233
      BackColor       =   16761024
      Caption         =   "RUT EMPRESA PROVEEDORA"
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
      Begin VB.CommandButton continuar 
         BackColor       =   &H00808080&
         Caption         =   "Continuar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4050
         TabIndex        =   17
         Top             =   1710
         Width           =   1335
      End
      Begin VB.TextBox dato7 
         BackColor       =   &H00C0FFFF&
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
         Left            =   90
         MaxLength       =   9
         TabIndex        =   16
         Tag             =   "rut"
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label nombrecontable 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   1035
         Width           =   5505
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT CONTABLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label dv2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1575
         TabIndex        =   18
         Top             =   630
         Width           =   285
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9105
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   16060
      BackColor       =   12648447
      Caption         =   "PANTALLA ORDENES DE COMPRA"
      CaptionEstilo3D =   1
      BackColor       =   12648447
      ForeColor       =   65535
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   16384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   690
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1217
         BackColor       =   12648447
         Caption         =   "DATOS DEL DOCUMENTO"
         CaptionEstilo3D =   1
         BackColor       =   12648447
         ForeColor       =   8438015
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox loc 
            Height          =   285
            Left            =   1170
            MaxLength       =   2
            TabIndex        =   10
            Top             =   315
            Width           =   330
         End
         Begin VB.TextBox numero 
            Height          =   285
            Left            =   4185
            MaxLength       =   10
            TabIndex        =   9
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label LBLPROVEEDOR 
            BackStyle       =   0  'Transparent
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
            Height          =   330
            Left            =   8595
            TabIndex        =   12
            Top             =   270
            Width           =   5910
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "LOCAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   135
            TabIndex        =   11
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "NUMERO DE ORDEN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            TabIndex        =   8
            Top             =   315
            Width           =   2400
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   3930
         Left            =   90
         TabIndex        =   3
         Top             =   1080
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   6932
         BackColor       =   8454016
         Caption         =   "DETALLE ORDEN DE COMPRA"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   1563884
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid3 
            Height          =   3570
            Left            =   0
            TabIndex        =   4
            Top             =   270
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   6297
            BackColorFixed  =   12648447
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   32768
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   3930
         Left            =   135
         TabIndex        =   6
         Top             =   5040
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   6932
         BackColor       =   8454016
         Caption         =   "DETALLE DOCUMENTOS RECIBIDOS"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   1563884
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton Command1 
            Caption         =   "&Eliminar Detalle"
            Height          =   255
            Left            =   6795
            TabIndex        =   14
            Top             =   3375
            Width           =   2505
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Grabar Documentos"
            Height          =   255
            Left            =   4275
            TabIndex        =   13
            Top             =   3375
            Width           =   2280
         End
         Begin FlexCell.Grid Grid1 
            Height          =   2715
            Left            =   0
            TabIndex        =   7
            Top             =   270
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   4789
            BackColorFixed  =   12648447
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   32768
            Rows            =   30
         End
      End
   End
   Begin MSAdodcLib.Adodc ordenes 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "muestraorden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String


Private Sub Command1_Click()
    Dim i As Integer
    
    Call eliminaPago
    Grid1.Rows = 2
    For i = 1 To Grid1.Cols - 1
        Grid1.Column(i).Locked = False
    Next i
    Grid1.AddItem ""
    Grid1.Cell(1, 1).SetFocus

End Sub



Private Sub Command2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
mes = COMBOMES.ListIndex + 1

leer


End Sub



Private Sub Command3_Click()
    Dim i As Integer
    
    If Grid1.Cell(1, 10).text <> "" Or Mid(Grid1.Cell(1, 1).text, 1, 2) = "OE" Then
        Call grabarPago
    End If

End Sub


Private Sub Command4_Click()
Dim k As Double
For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 14).text = "1" Then
        Call grabafactura(k, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text)
        
    End If
Next k
leer
End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
    sc = 0
planillaoc
planillafacturasdecompra

Call Conectarventas(servidor, clientesistema + "ventas00", usuario, password)
Call Conectargestion(servidor, clientesistema + "gestion", usuario, password)
Call Conectargestionrubro(servidor, clientesistema + "gestion00", usuario, password)
numero.text = numerodeorden
loc.text = localorden
leerecepcion

LeerPagos
bodega.Visible = False


End Sub








Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub




Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub


Private Sub lblhistorico_Click(Index As Integer)

End Sub




Private Sub Label16_Click()
End Sub

Sub limpia()
    
    
End Sub

Sub IMPRIMIR()
Dim titulo As String
titulo = "LISTADO DE FACTURAS EMITIDAS " + COMBOMES.text + " " + COMBOAÑO.text
Call cabezas2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub


Private Sub monto_Click()
End Sub

Private Sub leer()

Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + mes + "-" + "01"
    fecha2 = año + "-" + mes + "-" + "31"
    
        Set cSql.ActiveConnection = gestionrubro
        cSql.sql = "SELECT 'OC',numero,proveedor,fecha,montocomprado,montorecepcionado "
        cSql.sql = cSql.sql + "FROM l_ordendecompra_cabeza_" + localfiltro + " "
        cSql.sql = cSql.sql + "where fecharecepcion>='" + fecha1 + "' AND fecharecepcion<='" + fecha2 + "' order by proveedor "
        cSql.sql = cSql.sql + ""
        cSql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
             If leefactura(resultados(0), resultados(1), resultados(2)) = "0" Then
            Grid1.Rows = Grid1.Rows + 1
             
             LINEA = LINEA + 1
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = Mid(resultados(2), 1, 9) + "-" + Mid(resultados(2), 10, 1)
             Grid1.Cell(LINEA, 4).text = ""
             If IsNull(resultados(3)) = False Then
             Grid1.Cell(LINEA, 5).text = resultados(3)
             End If
             Grid1.Cell(LINEA, 6).text = resultados(4)
             Grid1.Cell(LINEA, 7).text = resultados(5)
             Grid1.Cell(LINEA, 8).text = leercompras(resultados(1))
             Grid1.Cell(LINEA, 9).text = "0"
             Grid1.Cell(LINEA, 10).text = "0"
            End If
            resultados.MoveNext
       
            Wend
End If
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      
      
End Sub
Sub limpiar()


End Sub

Sub cabezas2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ComboLOCAL.text
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub


Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = gestion
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        cSql.sql = cSql.sql + "ORDER BY codigo "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub
Sub eliminafactura(tipo, numero)
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = ventaslocal
        cSql.sql = "delete "
        cSql.sql = cSql.sql + "FROM sv_documento_cabeza_" + localfiltro + " "
        cSql.sql = cSql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        cSql.Execute
        cSql.sql = "delete "
        cSql.sql = cSql.sql + "FROM sv_documento_detalle_" + localfiltro + " "
        cSql.sql = cSql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        cSql.Execute
        cSql.sql = "delete "
        cSql.sql = cSql.sql + "FROM sv_documento_pagos_" + localfiltro + " "
        cSql.sql = cSql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        cSql.Execute
        
        Set cSql.ActiveConnection = gestionrubro
        cSql.sql = "delete "
        cSql.sql = cSql.sql + "FROM l_movimientos_detalle_" + localfiltro + " "
        cSql.sql = cSql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        cSql.Execute

        
End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub

Sub grabafactura(LINEA, tipo, orden)
    Dim netos As Double
    Dim DH As String
    Dim DH2 As String
    
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim ELECTRONICA As String
    Dim TIPODOC As String
    Dim fecha As Date
    
    campos(0, 0) = "numero"
    campos(1, 0) = "proveedor"
    campos(2, 0) = "nombre"
    campos(3, 0) = "fecharecepcion"
    campos(4, 0) = "montorecepcionado"
    campos(5, 0) = "montodocumentos"
    campos(18, 0) = ""
 
    If Grid1.Cell(LINEA, 1).text = "FA" Then TIPOCON = "1": ELECTRONICA = "N": TIPODOC = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "2": ELECTRONICA = "N": TIPODOC = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "3": ELECTRONICA = "N": TIPODOC = "NC": DH = "D": DH2 = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": ELECTRONICA = "S": TIPODOC = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": ELECTRONICA = "S": TIPODOC = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": ELECTRONICA = "S": TIPODOC = "NC": DH = "D": DH2 = "H"
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(3, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(4, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(5, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 7).text, ",", ".")
    exentos = CDbl(Grid1.Cell(LINEA, 8).text) + CDbl(Grid1.Cell(LINEA, 9).text) + CDbl(Grid1.Cell(LINEA, 10).text) + CDbl(Grid1.Cell(LINEA, 11).text) + CDbl(Grid1.Cell(LINEA, 12).text)
    campos(7, 1) = Str(exentos)
    campos(8, 1) = "0"
    campos(9, 1) = Replace(Grid1.Cell(LINEA, 13).text, ",", ".")
    campos(10, 1) = año
    campos(11, 1) = Format(mes, "00")
    campos(12, 1) = "CENTRALIZACION AUTOMATICA"
        
    campos(13, 1) = ELECTRONICA
    campos(14, 1) = "N"
    campos(15, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(16, 1) = LEERULTIMOFOLIO
    campos(17, 1) = "0"
    
    condicion = ""
    campos(0, 2) = "facturasdecompras"
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db

    Call SQLUTIL.SQLUTIL(op, condicion)
    k = SQLUTIL.estado
    fecha = Format(campos(3, 1), "yyyy-mm-dd")
    Call grabarcomprobante_lineas(TIPODOC, campos(1, 1), "001", fecha, cuentaproveedor, "", campos(4, 1), "", "CENTRALIZA DOCUMENTO DE COMPRAS " + Grid1.Cell(LINEA, 1).text, TIPODOC, campos(1, 1), campos(2, 1), campos(3, 1), campos(9, 1), DH, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    Call grabarcomprobante_lineas(TIPODOC, campos(1, 1), "002", fecha, ivacredito, "", "", "", "CENTRALIZACION I.V.A", TIPODOC, campos(1, 1), campos(2, 1), campos(3, 1), campos(6, 1), DH2, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    Call grabardetallefactura(LINEA, tipo, orden, fecha, campos(11, 1), campos(10, 1))


End Sub

Sub grabardetallefactura(LINEA, tipo, orden, fecha, mes, año)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim ilas As Double
    Dim CRCC As String
    Dim CUENTA As String
    Dim DH As String
    Dim NOMBRE As String
    Dim TIPODOC As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = ""
    If localfiltro = "00" Then CRCC = "0101"
    If localfiltro = "41" Then CRCC = "0104"
    If Grid1.Cell(LINEA, 1).text = "FA" Then TIPOCON = "1": TIPODOC = "FC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "2": TIPODOC = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "3": TIPODOC = "NC": DH = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": TIPODOC = "FC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": TIPODOC = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": TIPODOC = "NC": DH = "H"
    
    If tipo = "DI" Then CUENTA = "11350008": NOMBRE = "DIARIOS"
    If tipo = "ME" Then CUENTA = "11350001": NOMBRE = "MERCADERIAS"
    If tipo = "CI" Then CUENTA = "11350007": NOMBRE = "CIGARRILLOS"
    If tipo = "FR" Then CUENTA = "11350002": NOMBRE = "FRUTAS"
    If tipo = "CA" Then CUENTA = "11350003": NOMBRE = "CARNICERIA"
    If tipo = "FI" Then CUENTA = "11350004": NOMBRE = "FIAMBRERIA"
    If tipo = "PA" Then CUENTA = "11350007": NOMBRE = "PANADERIA"
    If tipo = "EM" Then CUENTA = "11350006": NOMBRE = "MATERIAL EMPAQUE"
    

Rem CALCULA NETOS

    lin = 3
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = CUENTA
    campos(5, 1) = "O/C " + orden + " " + NOMBRE
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(7, 1) = DH
    campos(8, 1) = CRCC
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    Call grabarcomprobante_lineas(TIPODOC, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", campos(3, 1), "", campos(5, 1), TIPODOC, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, mes, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    
    
Rem CALCULA ILAS
    ilas = CDbl(Grid1.Cell(LINEA, 8).text) + CDbl(Grid1.Cell(LINEA, 9).text) + CDbl(Grid1.Cell(LINEA, 10).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = "11400006"
    campos(5, 1) = "O/C " + orden + " IMPUESTO ILAS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    Call grabarcomprobante_lineas(TIPODOC, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), TIPODOC, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, mes, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
    
Rem CALCULA HARINA
    ilas = CDbl(Grid1.Cell(LINEA, 11).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = "11400005"
    campos(5, 1) = "O/C " + orden + " IMPUESTO HARINAS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    Call grabarcomprobante_lineas(TIPODOC, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), TIPODOC, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, mes, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA carne
    ilas = CDbl(Grid1.Cell(LINEA, 12).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = "11400012"
    campos(5, 1) = "O/C " + orden + " IMPUESTO CARNE"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    Call grabarcomprobante_lineas(TIPODOC, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), TIPODOC, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, mes, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
    
   
    
    
End Sub

Public Function leefactura(tipo, numero, rut) As String

    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    If tipo = "FA" Then tipo = "1"
    If tipo = "ND" Then tipo = "2"
    If tipo = "NC" Then tipo = "3"
    If tipo = "FAE" Then tipo = "4"
    If tipo = "NDE" Then tipo = "5"
    If tipo = "NCE" Then tipo = "6"
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    campos(0, 2) = "facturasdecompras"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then
    leefactura = "1"
    
    Else
    leefactura = "0"
    
    End If
    
    

End Function

Public Function nombrectacte(rut) As String

    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + cuentaproveedor + "' and rut=" + "'" + rut + "' and año='" + año + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    nombrectacte = "*** NO CREADO *** "
    
    If SQLUTIL.estado = 0 Then
    nombrectacte = SQLUTIL.datos(1, 3)
    
    End If
    
End Function
Sub crearcuentacorriente(rut)
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = gestion

            cSql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            cSql.sql = cSql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            cSql.sql = cSql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            cSql.sql = cSql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            cSql.sql = cSql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            cSql.Execute
            
            cSql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            cSql.sql = cSql.sql & "(año,tipo,rut) "
            cSql.sql = cSql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
            cSql.sql = cSql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            cSql.sql = cSql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            cSql.Execute


End Sub
'cSql.SQL = "INSERT INTO l_movimientos_detalle_" & empresaactiva & " "
'            cSql.SQL = cSql.SQL & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
'            cSql.SQL = cSql.SQL & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.unidades, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, ROUND(dd.unidades / dd.cantidad, 0) "
'            cSql.SQL = cSql.SQL & "FROM " & baseVentas & rubro & ".sv_documento_detalle_" + empresaactiva + " as dd "
'            cSql.SQL = cSql.SQL & "WHERE dd.local = '" & empresaactiva & "' AND dd.tipo = '" & v.detalle.tipo & "' AND dd.numero = '" & v.detalle.numero & "'"
'            cSql.Execute

Public Function LEERULTIMOFOLIO() As String

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = db

            cSql.sql = "select max(folio) from facturasdecompras where mescontable = '" & Format(mes, "00") & "' AND añocontable = '" & año & "' "
            
            cSql.Execute
    If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
    
    
        LEERULTIMOFOLIO = Format(resultados(0) + 1, "000000")
    End If
    
End Function
Public Function LEERMONTOIMPUESTO(tipo, numero, rut, CUENTA) As Double

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = gestionrubro

            cSql.sql = "select monto from l_ordendecompra_impuestos_" + localfiltro + " where cuenta = '" & CUENTA & "' and tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
            
            cSql.Execute
    LEERMONTOIMPUESTO = 0
    If cSql.RowsAffected > 0 Then
    
    Set resultados = cSql.OpenResultset
    LEERMONTOIMPUESTO = resultados(0)
    
    End If
    
End Function
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, mes, año, fechacreacion, horacreacion, rutproveedor)
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
    campos(16, 1) = mes
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    
    Call SQLUTIL.SQLUTIL(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub

Public Function leercompras(orden) As Double
Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim total As Double
    Dim MULTI As Double
    
        Set cSql.ActiveConnection = gestionrubro
        cSql.sql = "SELECT tipo,total "
        cSql.sql = cSql.sql + "FROM l_ordendecompra_detalle_facturas_" + localfiltro + " WHERE ordendecompra='" + orden + "' "
        cSql.sql = cSql.sql + "ORDER BY ordendecompra "
        cSql.Execute
        total = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
               If resultados(0) = "NCE" Or resultados(0) = "NC" Then MULTI = -1 Else MULTI = 1
               total = total + (resultados(1) * MULTI)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        leercompras = total
End Function

Private Sub Grid2_AfterReorderColumn(ByVal OriginalPosition As Long, ByVal NewPosition As Long)

End Sub
Sub planillaoc()
    Rem DATOS DE LA COLUMNA
    Dim formatogrilla(10, 10)
    Grid3.DefaultFont.Size = 8
    Grid3.DefaultFont.Bold = True
    
    formatogrilla(1, 1) = "CODIGO"
    formatogrilla(1, 2) = "DESCRIPCION"
    formatogrilla(1, 3) = "CAJAS"
    formatogrilla(1, 4) = "UxC"
    formatogrilla(1, 5) = "UNIDADES"
    formatogrilla(1, 6) = "P.UNI."
    formatogrilla(1, 7) = "TOTAL C/IVA"
    formatogrilla(1, 8) = "TOTAL NETO "
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "13"
    formatogrilla(2, 2) = "37"
    formatogrilla(2, 3) = "9"
    formatogrilla(2, 4) = "6"
    formatogrilla(2, 5) = "9"
    formatogrilla(2, 6) = "9"
    formatogrilla(2, 7) = "13"
    formatogrilla(2, 8) = "13"


    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "N"
    formatogrilla(3, 4) = "N"
    formatogrilla(3, 5) = "N"
    formatogrilla(3, 6) = "N"
    formatogrilla(3, 7) = "N"
    formatogrilla(3, 8) = "N"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = "#,###,##0.0"
    formatogrilla(4, 4) = "#,###,##0.0"
    formatogrilla(4, 5) = "#,###,##0.0"
    formatogrilla(4, 6) = "#,###,##0.0"
    formatogrilla(4, 7) = "#,###,##0.0"
    formatogrilla(4, 8) = "#,###,##0.0"
    
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "FALSE"
    formatogrilla(5, 4) = "FALSE"
    formatogrilla(5, 5) = "TRUE"
       
     
     
  
    formatogrilla(5, 6) = "TRUE"
    
    formatogrilla(5, 7) = "TRUE"
    formatogrilla(5, 8) = "TRUE"

Grid3.Cols = 9
Grid3.Rows = 1

    
    
    
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    Grid3.BackColorFixed = RGB(90, 158, 214)
    Grid3.BackColorFixedSel = RGB(110, 180, 214)
    Grid3.BackColorBkg = RGB(90, 158, 214)
    Grid3.BackColorScrollBar = RGB(231, 235, 247)
    Grid3.BackColor1 = RGB(231, 235, 247)
    Grid3.BackColor2 = RGB(239, 243, 255)
    Grid3.GridColor = RGB(148, 190, 231)
    For k = 1 To Grid3.Cols - 1
        Grid3.Cell(0, k).text = formatogrilla(1, k)
        Grid3.Column(k).Width = Val(formatogrilla(2, k)) * Grid3.DefaultFont.Size
        Grid3.Column(k).MaxLength = Val(formatogrilla(2, k))
        Grid3.Column(k).FormatString = formatogrilla(4, k)
        Grid3.Column(k).Locked = formatogrilla(5, k)
        If formatogrilla(3, k) = "N" Then Grid3.Column(k).Alignment = cellRightCenter
       
    Next k
Grid3.Column(0).Width = 30
Grid3.Cell(0, 0).text = "lin"
Grid3.Range(0, 0, 0, Grid3.Cols - 1).Alignment = cellCenterCenter
    

Grid3.Enabled = False
End Sub

Sub planillafacturasdecompra()
    Dim formatogrilla(10, 20)
    Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7
    Grid1.DefaultFont.Bold = True
    
    formatogrilla(1, 1) = "TIPO"
    formatogrilla(1, 2) = "NUMERO"
    formatogrilla(1, 3) = "FECHA"
    formatogrilla(1, 4) = "VENCIMIENTO"
    formatogrilla(1, 5) = "NETO"
    formatogrilla(1, 6) = "IVA"
    formatogrilla(1, 7) = "EXENTO"
    formatogrilla(1, 8) = "IMPUESTOS"
    formatogrilla(1, 9) = "TOTAL"
    formatogrilla(1, 10) = "TIPO"
    formatogrilla(1, 11) = "BONIFICACION"
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "20"
    formatogrilla(2, 2) = "12"
    formatogrilla(2, 3) = "12"
    formatogrilla(2, 4) = "12"
    formatogrilla(2, 5) = "12"
    formatogrilla(2, 6) = "12"
    formatogrilla(2, 7) = "12"
    formatogrilla(2, 9) = "12"
    formatogrilla(2, 9) = "12"
    formatogrilla(2, 10) = "12"
    formatogrilla(2, 11) = "12"


    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "C"
    formatogrilla(3, 2) = "N"
    formatogrilla(3, 3) = "D"
    formatogrilla(3, 4) = "D"
    formatogrilla(3, 5) = "N"
    formatogrilla(3, 6) = "N"
    formatogrilla(3, 7) = "N"
    formatogrilla(3, 9) = "N"
    formatogrilla(3, 9) = "N"
    formatogrilla(3, 10) = "C"
    formatogrilla(3, 11) = "CH"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = ""
    formatogrilla(4, 4) = ""
    formatogrilla(4, 6) = "#,###,##0.0"
    formatogrilla(4, 7) = "#,###,##0.0"
    formatogrilla(4, 9) = "#,###,##0.0"
    formatogrilla(4, 9) = "#,###,##0.0"
    
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    formatogrilla(5, 3) = "FALSE"
    formatogrilla(5, 4) = "FALSE"
    formatogrilla(5, 5) = "FALSE"
    formatogrilla(5, 6) = "FALSE"
    formatogrilla(5, 7) = "FALSE"
    formatogrilla(5, 8) = "TRUE"
    formatogrilla(5, 9) = "TRUE"
    formatogrilla(5, 10) = "FALSE"
    formatogrilla(5, 11) = "FALSE"

    Grid1.Cols = 12
    Grid1.Rows = 2
    
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.BackColorFixed = RGB(90, 158, 214)
    Grid1.BackColorFixedSel = RGB(110, 180, 214)
    Grid1.BackColorBkg = RGB(90, 158, 214)
    Grid1.BackColorScrollBar = RGB(231, 235, 247)
    Grid1.BackColor1 = RGB(231, 235, 247)
    Grid1.BackColor2 = RGB(239, 243, 255)
    Grid1.GridColor = RGB(148, 190, 231)
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla(1, k)
        Grid1.Column(k).Width = Val(formatogrilla(2, k)) * Grid3.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(formatogrilla(2, k))
        Grid1.Column(k).FormatString = formatogrilla(4, k)
        Grid1.Column(k).Locked = formatogrilla(5, k)
        If formatogrilla(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        If formatogrilla(3, k) = "C" Then Grid1.Column(k).CellType = cellComboBox
        If formatogrilla(3, k) = "CH" Then Grid1.Column(k).CellType = cellCheckBox
    Next k
    Grid1.Column(0).Width = 0
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
    
    With Grid1.ComboBox(1)
        '.Locked = False
        .AutoComplete = True
        .Font.Name = "Courier New"
        .AddItem "FA FACTURA" '1
        .AddItem "ND NOTA DEBITO" '2
        .AddItem "NC NOTA CREDITO" '3
        .AddItem "FAE FACTURA ELECTRONICA" '1
        .AddItem "NDE NOTA DEBITO ELECTRONICA" '2
        .AddItem "NCE NOTA CREDITO ELECTRONICA" '3
        .AddItem "OE ORDEN DE ENLACE" '4
        .AddItem "GD DESPACHO" '4
    
    
    End With
    With Grid1.ComboBox(10)
        '.Locked = True
        .AutoComplete = True
        .Font.Name = "Courier New"
        .AddItem "MERCADERIAS"
        .AddItem "CIGARRILLOS"
        .AddItem "FRUTAS Y VERDURAS"
        .AddItem "CARNICERIA"
        .AddItem "FIAMBRERIA"
        .AddItem "PANADERIA"
        .AddItem "EMPAQUE"
        .AddItem "DIARIOS"
        
    End With

Grid1.Enabled = False
End Sub

Private Function LeerPagos()
   Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim total As Double
    Dim MULTI As Double
    
        Set cSql.ActiveConnection = gestionrubro
        cSql.sql = "SELECT CONCAT(tipo, '" & vbTab & "', numero, '" & vbTab & "', fecha, '" & vbTab & "', vencimiento, '" & vbTab & "', neto, '" & vbTab & "', iva, '" & vbTab & "', exento, '" & vbTab & "', impuestos, '" & vbTab & "', total, '" & vbTab & "', categoria, '" & vbTab & "', bonificacion) AS item "
        cSql.sql = cSql.sql + " FROM l_ordendecompra_detalle_facturas_" & loc.text & " "
        cSql.sql = cSql.sql + "WHERE ordendecompra = '" & numero.text & "' order by linea "
        cSql.Execute
        
        Grid1.AutoRedraw = False
        Grid1.Rows = 1
      
        If cSql.RowsAffected > 0 Then
        
        Set resultados = cSql.OpenResultset
        While resultados.EOF = False
            
            Grid1.AddItem resultados(0), True
        
        resultados.MoveNext
        Wend
        End If
Grid1.AutoRedraw = True
Grid1.Refresh


    
End Function

Sub leerecepcion()
    Dim lin As Integer
    Dim suma As Double
    Dim sql As String
   Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
   
        Set cSql.ActiveConnection = gestionrubro
        
        Rem Call leerProveedor
        
        cSql.sql = "SELECT codigo,r_maestroproductos_fijo_" & RUBRO & ".descripcion,linea,cantidad,uxc,unidades,precio,total,bodega,fecha "
        cSql.sql = cSql.sql + "FROM r_maestroproductos_fijo_" & RUBRO & ",l_movimientos_detalle_" + loc.text + " "
        cSql.sql = cSql.sql + "WHERE codigobarra=codigo and tipo='OC' AND numero='" + numero.text + "' order by linea "
    
       cSql.Execute
       
        
        If cSql.RowsAffected > 0 Then
        
            Grid3.Rows = cSql.RowsAffected + 1
            
            suma = 0: lin = 0
'            If movi.Recordset.EOF = False Then
'            dato6.text = movi.Recordset.Fields(8)
'            dato2.text = Format(movi.Recordset.Fields(9), "dd")
'            dato3.text = Format(movi.Recordset.Fields(9), "mm")
'            dato4.text = Format(movi.Recordset.Fields(9), "yyyy")
'
'            BODEGARECEPCION.Caption = leerNombreBodega(dato6.text)
'            BODE.text = dato6.text
'            End If
'
            Set resultados = cSql.OpenResultset
            
            
            While Not resultados.EOF
                lin = lin + 1
               Grid3.Cell(lin, 0).text = lin
                Grid3.Cell(lin, 1).text = resultados(0)
                Grid3.Cell(lin, 2).text = resultados(1)
                Grid3.Cell(lin, 3).text = resultados(3)
                Grid3.Cell(lin, 4).text = resultados(4)
                Grid3.Cell(lin, 5).text = resultados(5)
                Grid3.Cell(lin, 6).text = resultados(6)
                Grid3.Cell(lin, 7).text = resultados(7)
                Grid3.Cell(lin, 8).text = CDbl(resultados(7) / 1.19)
                
                Grid3.Column(3).Locked = False
                suma = suma + resultados(7)
                resultados.MoveNext
                
            Wend
        If lin <> 0 Then
'            EXISTE = "S"
            Grid3.Enabled = False
            'OPCIONES.Visible = True
            'OPCIONES.SetFocus
            'calcularecepcion
           ' Command4.Visible = True
            
        End If
        End If
       ' If lin = 0 Then EXISTE = "N"
End Sub

Sub grabarPago()
    Dim i As Integer
    Dim cmps(5, 3) As String
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "vencimiento"
    campos(4, 0) = "neto"
    campos(5, 0) = "iva"
    campos(6, 0) = "exento"
    campos(7, 0) = "impuestos"
    campos(8, 0) = "total"
    campos(9, 0) = "categoria"
    campos(10, 0) = "bonificacion"
    campos(11, 0) = "ordendecompra"
    campos(12, 0) = "linea"
    campos(13, 0) = "rut"
    campos(14, 0) = ""
    cmps(0, 0) = "ordenconfactura"
    cmps(1, 0) = "ordenenlazada"
    cmps(2, 0) = ""
    For i = 1 To Grid1.Rows - 1
        If Grid1.Cell(i, 1).text = "OE ORDEN DE ENLACE" Then
            cmps(0, 1) = Grid1.Cell(i, 2).text
            cmps(1, 1) = DATO1.text
            cmps(2, 1) = ""
            cmps(0, 2) = "l_ordendecompra_enlace_factura_" & empresaactiva
            condicion = ""
            op = 2
            SQLUTIL.datos = cmps
            Set SQLUTIL.conexion = gestionrubro
            Call SQLUTIL.SQLUTIL(op, condicion)
        End If
            If Grid1.Cell(i, 10).text <> "" Or Mid(Grid1.Cell(i, 1).text, 1, 2) = "OE" Then
                campos(0, 1) = Grid1.Cell(i, 1).text
                campos(1, 1) = Grid1.Cell(i, 2).text
                campos(2, 1) = Format(Grid1.Cell(i, 3).text, "yyyy-mm-dd")
                campos(3, 1) = Format(Grid1.Cell(i, 4).text, "yyyy-mm-dd")
                campos(4, 1) = Grid1.Cell(i, 5).text
                campos(5, 1) = Grid1.Cell(i, 6).text
                campos(6, 1) = Grid1.Cell(i, 7).text
                campos(7, 1) = Grid1.Cell(i, 8).text
                campos(8, 1) = Replace(Grid1.Cell(i, 9).text, ".", "")
                campos(9, 1) = Grid1.Cell(i, 10).text
                campos(12, 1) = Str(i)
                campos(13, 1) = dato7.text + dv2.Caption
                campos(10, 1) = Grid1.Cell(i, 11).text
                campos(11, 1) = DATO1.text
                campos(0, 2) = "l_ordendecompra_detalle_facturas_" & empresaactiva
                condicion = ""
                op = 2
                SQLUTIL.datos = campos
                Set SQLUTIL.conexion = gestionrubro
                Call SQLUTIL.SQLUTIL(op, condicion)
            End If
      
    Next i
End Sub

Private Sub eliminaPago()
    Dim orden As String
    Call ConectarControlData(ordenes, servidor, "eltit_gestion" & RUBRO, usuario, password, "SELECT DISTINCT oef.ordenenlazada FROM l_ordendecompra_enlace_factura_" & empresaactiva & " AS oef, l_ordendecompra_detalle_facturas_" & empresaactiva & " AS opm WHERE oef.ordenconfactura = '" & numerodeorden & "' AND oef.ordenconfactura = opm.ordendecompra ORDER BY oef.ordenenlazada ASC")
    campos(0, 2) = "l_ordendecompra_enlace_factura_" & empresaactiva
    If ordenes.Recordset.RecordCount > 0 Then
        ordenes.Recordset.MoveFirst
        While Not ordenes.Recordset.EOF
            orden = ordenes.Recordset.Fields("ordenenlazada")
            condicion = "ordenconfactura = '" & numerodeorden & "' AND ordenenlazada = '" & orden & "'"
            op = 4
            SQLUTIL.datos = campos
            Set SQLUTIL.conexion = gestionrubro
            Call SQLUTIL.SQLUTIL(op, condicion)
            ordenes.Recordset.MoveNext
        Wend
    Else
        condicion = "ordenenlazada = '" & numerodeorden & "'"
        op = 4
        SQLUTIL.datos = campos
        Set SQLUTIL.conexion = gestionrubro
        Call SQLUTIL.SQLUTIL(op, condicion)
    End If
    campos(0, 2) = "l_ordendecompra_detalle_facturas_" & empresaactiva
    condicion = "ordendecompra = '" & numerodeorden & "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = gestionrubro
    Call SQLUTIL.SQLUTIL(op, condicion)
End Sub

