VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form AuditoriaDetalles 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmDetalle 
      Height          =   7755
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13679
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin FlexCell.Grid impresion 
         Height          =   7275
         Left            =   60
         TabIndex        =   1
         Top             =   420
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12832
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   255
         Left            =   11220
         TabIndex        =   2
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   16761024
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
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   1380
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   60
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   8220
      TabIndex        =   3
      Top             =   7860
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "I   M   P   R   I   M   I   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "AuditoriaDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public fechaini As String
    Public fechafin As String
    Public codLocal As String
    Public codRubro As String
    Public informe As String
    Public tipoInforme As Integer

Private Sub Form_Activate()
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    frmDetalle.Caption = titulo
    If informe = "INGRESOS" Then
        Select Case tipoInforme
            Case 1
                Call efectivo
            Case 2
                Call Cheques
            Case 3
'                Call credito
            Case 4
            Case 6
                Call pagos
        End Select
    End If
    If informe = "EGRESOS" Then
        Select Case tipoInforme
            Case 1
            Case 2
            Case 3
            Case 4
        End Select
    End If
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub Form_Load()
    Call CargaGrillaImpresion(1, 40)
End Sub

'****************************************************************************
'Formato de la Grilla Impresion
'****************************************************************************
    Private Sub CargaGrillaImpresion(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Impresion.Cols = col
        Impresion.Rows = row
        Impresion.AllowUserResizing = False
        Impresion.DisplayFocusRect = False
        Impresion.ExtendLastCol = False
        Impresion.BoldFixedCell = False
        Impresion.DrawMode = cellOwnerDraw
        Impresion.Appearance = Flat
        Impresion.ScrollBarStyle = Flat
        Impresion.FixedRowColStyle = Flat
        Impresion.BackColorFixed = RGB(90, 158, 214)
        Impresion.BackColorFixedSel = RGB(110, 180, 230)
        Impresion.BackColorBkg = RGB(90, 158, 214)
        Impresion.BackColorScrollBar = RGB(231, 235, 247)
        Impresion.BackColor1 = RGB(231, 235, 247)
        Impresion.BackColor2 = RGB(239, 243, 255)
        Impresion.GridColor = RGB(148, 190, 231)
        
        Impresion.Column(0).Width = 0
        Impresion.RowHeight(0) = 0
        
        For i = 1 To col - 1
            Impresion.Column(i).Width = 2.25 * (Impresion.Cell(0, i).Font.Size)
        Next i
    End Sub
'****************************************************************************
'Formato de la Grilla Impresion
'****************************************************************************

    Private Sub efectivo()
        Dim tabulador As String
        Dim i As Integer
        Dim cadena As String
        Impresion.Rows = 1
        Impresion.AutoRedraw = False
        tabulador = ""
        For i = 1 To 5
            tabulador = tabulador & vbTab
        Next i
        Impresion.AddItem "", True
        cadena = tabulador
        cadena = cadena & "Efectivo" & tabulador
        cadena = cadena & "=" & vbTab
        cadena = cadena & "Total Ingresos" & tabulador
        cadena = cadena & "-" & vbTab
        cadena = cadena & "Cheques" & tabulador
        cadena = cadena & "-" & vbTab
        cadena = cadena & "Creditos" & tabulador
        cadena = cadena & "-" & vbTab
        cadena = cadena & "Depositos"
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 10).Merge
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).Merge
        Impresion.Range(Impresion.Rows - 1, 18, Impresion.Rows - 1, 22).Merge
        Impresion.Range(Impresion.Rows - 1, 24, Impresion.Rows - 1, 28).Merge
        Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 34).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        
        Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 10).FontBold = True
        
        Impresion.AddItem "", True
        cadena = tabulador & tabulador & vbTab
        cadena = cadena & "Total Ingresos" & tabulador
        cadena = cadena & "=" & vbTab
        cadena = cadena & "Total Documentos" & tabulador
        cadena = cadena & "+" & vbTab
        cadena = cadena & "Pagos de Clientes"
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).Merge
        Impresion.Range(Impresion.Rows - 1, 18, Impresion.Rows - 1, 22).Merge
        Impresion.Range(Impresion.Rows - 1, 24, Impresion.Rows - 1, 28).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).FontBold = True
        
        Impresion.AddItem "", True
        cadena = tabulador & tabulador & vbTab
        cadena = cadena & "Total Ingresos" & tabulador
        cadena = cadena & "=" & vbTab
        cadena = cadena & PAuditoriaVentas.Documentos.Cell(6, 4).text & tabulador
        cadena = cadena & "+" & vbTab
        cadena = cadena & PAuditoriaVentas.Ingresos.Cell(6, 1).text
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).Merge
        Impresion.Range(Impresion.Rows - 1, 18, Impresion.Rows - 1, 22).Merge
        Impresion.Range(Impresion.Rows - 1, 24, Impresion.Rows - 1, 28).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).FontBold = True
        
        Impresion.AddItem "", True
        cadena = tabulador & tabulador & vbTab
        cadena = cadena & "Total Ingresos" & tabulador
        cadena = cadena & "=" & vbTab
        cadena = cadena & PAuditoriaVentas.Ingresos.Cell(5, 1).text
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).Merge
        Impresion.Range(Impresion.Rows - 1, 18, Impresion.Rows - 1, 22).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).FontBold = True
        
        
        Impresion.AddItem "", True
        Impresion.AddItem "", True
        cadena = tabulador
        cadena = cadena & "Efectivo" & tabulador
        cadena = cadena & "=" & vbTab
        cadena = cadena & PAuditoriaVentas.Ingresos.Cell(5, 1).text & tabulador
        cadena = cadena & "-" & vbTab
        cadena = cadena & PAuditoriaVentas.Ingresos.Cell(2, 1).text & tabulador
        cadena = cadena & "-" & vbTab
        cadena = cadena & PAuditoriaVentas.Ingresos.Cell(3, 1).text & tabulador
        cadena = cadena & "-" & vbTab
        cadena = cadena & PAuditoriaVentas.Ingresos.Cell(4, 1).text
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 10).Merge
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).Merge
        Impresion.Range(Impresion.Rows - 1, 18, Impresion.Rows - 1, 22).Merge
        Impresion.Range(Impresion.Rows - 1, 24, Impresion.Rows - 1, 28).Merge
        Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 34).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        
        Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 10).FontBold = True
        
        Impresion.AddItem "", True
        cadena = tabulador
        cadena = cadena & "Efectivo" & tabulador
        cadena = cadena & "=" & vbTab
        cadena = cadena & PAuditoriaVentas.Ingresos.Cell(1, 1).text
        Impresion.AddItem cadena, True
        Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 10).Merge
        Impresion.Range(Impresion.Rows - 1, 12, Impresion.Rows - 1, 16).Merge
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
        
        Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 16).FontBold = True
        
        Impresion.AutoRedraw = True
        Impresion.Refresh
    End Sub

    Private Sub pagos()
        Dim tabla As String
        Dim NUMERO As String
        Dim rut As String
        Dim dias As Double
        Dim fecha As String
        Dim i As Integer
        Dim TIPO As String
        Dim Cliente As String
        Dim cadena As String
        Dim tabulador As String
        Dim resultados As rdoResultset
        Dim cSql As New rdoQuery
        Dim resultados1  As rdoResultset
        Dim cSql1 As New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        Set cSql.ActiveConnection = ventasRubro
        
        tabulador = ""
        For i = 1 To 4
            tabulador = tabulador & vbTab
        Next i
        'DATOS
        cSql.sql = "SELECT CONCAT(pc.numero, '" & tabulador & "', DATE_FORMAT(pc.fecha,'%d-%m-%Y'), '" & tabulador & "') AS item1, pc.rut, CONCAT('" & tabulador & tabulador & tabulador & vbTab & "', CASE pc.tipopago WHEN '1' THEN '1 EFECTIVO' WHEN '2' THEN '2 CHEQUE' WHEN '3' THEN '3 DEPOSITO' ELSE '' END, '" & tabulador & vbTab & vbTab & "', CONCAT('$ ', FORMAT(pc.monto,0))) AS item2, pc.numero, pc.rut, DATE_FORMAT(pc.fecha,'%d-%m-%Y') AS fecha, pc.tipopago "
        cSql.sql = cSql.sql & "FROM sv_pagos_cabeza_" & empresaActiva & " AS pc "
        cSql.sql = cSql.sql & "WHERE pc.local = '" & codLocal & "' AND pc.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' ORDER BY pc.numero ASC"
'        Call ConectarControlData(data, servidor, baseVentas & codRubro, usuario, password, tabla)
        cSql.Execute
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
            'LISTADO DE CLIENTES CON PAGOS
            'TITULO
            Impresion.AddItem "LISTADO DE PAGO DE CLIENTES", True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            'CABEZA
            'PRIMERA LINEA
            cadena = "NUMERO" & tabulador
            cadena = cadena & "FECHA" & tabulador
            cadena = cadena & "RUT" & tabulador
            cadena = cadena & "NOMBRE" & tabulador & tabulador & tabulador & vbTab
            cadena = cadena & "FORMA PAGO" & tabulador & vbTab & vbTab
            cadena = cadena & "MONTO PAGO" & tabulador
            cadena = cadena & "VENCIMIENTO"
            Impresion.AddItem cadena, True
            'UNION
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 4).Merge
            Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 8).Merge
            Impresion.Range(Impresion.Rows - 1, 9, Impresion.Rows - 1, 12).Merge
            Impresion.Range(Impresion.Rows - 1, 13, Impresion.Rows - 1, 25).Merge
            Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 31).Merge
            Impresion.Range(Impresion.Rows - 1, 32, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            'SEGUNDA LINEA
            cadena = tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "DOCUMENTO" & tabulador & vbTab
            cadena = cadena & "FECHA" & tabulador
            cadena = cadena & "MONTO" & tabulador
            cadena = cadena & "PLAZO"
            Impresion.AddItem cadena, True
            'UNION
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 18).Merge
            Impresion.Range(Impresion.Rows - 1, 19, Impresion.Rows - 1, 23).Merge
            Impresion.Range(Impresion.Rows - 1, 24, Impresion.Rows - 1, 27).Merge
            Impresion.Range(Impresion.Rows - 1, 28, Impresion.Rows - 1, 31).Merge
            Impresion.Range(Impresion.Rows - 1, 32, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            Impresion.Range(Impresion.Rows - 1, 19, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 18).Borders(cellEdgeTop) = cellThin
            Impresion.Range(Impresion.Rows - 1, 18, Impresion.Rows - 1, 18).Borders(cellEdgeRight) = cellThin
            
'            data.Recordset.MoveFirst
            While Not resultados.EOF
                NUMERO = resultados("numero")
                rut = resultados("rut")
                fecha = resultados("fecha")
                TIPO = resultados("tipopago")
                Cliente = rut & tabulador & leerNombreCliente(rut)
                If TIPO = "2" Then
                    cadena = resultados("item1") & Cliente & resultados("item2") & tabulador & leerVencimientoPago(NUMERO)
                Else
                    cadena = resultados("item1") & Cliente & resultados("item2") & tabulador & fecha
                End If
                Impresion.AddItem Replace(cadena, ",", "."), True
                'UNION
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 4).Merge
                Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 8).Merge
                Impresion.Range(Impresion.Rows - 1, 9, Impresion.Rows - 1, 12).Merge
                Impresion.Range(Impresion.Rows - 1, 13, Impresion.Rows - 1, 25).Merge
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 31).Merge
                Impresion.Range(Impresion.Rows - 1, 32, Impresion.Rows - 1, 35).Merge
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
                'ALINEACION
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 4).Alignment = cellCenterCenter
                Impresion.Range(Impresion.Rows - 1, 5, Impresion.Rows - 1, 8).Alignment = cellCenterCenter
                Impresion.Range(Impresion.Rows - 1, 9, Impresion.Rows - 1, 12).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 13, Impresion.Rows - 1, 25).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 31).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 32, Impresion.Rows - 1, 35).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
                
                'LISTADO DE DOCUMENTOS PAGADOS
                
                'tabla = "SELECT CONCAT(pd.tipo, ' ', pd.documento, '" & tabulador & vbTab & "', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & tabulador & "', CONCAT('$ ', FORMAT(pd.monto,0)), '" & tabulador & "') AS item, DATE_FORMAT(dc.vencimiento,'%d-%m-%Y') AS fecha "
                'tabla = tabla & "FROM sv_pagos_detalle AS pd INNER JOIN sv_documento_cabeza AS dc ON pd.tipo = dc.tipo AND pd.documento = dc.numero "
                'tabla = tabla & "WHERE pd.local = '" & codLocal & "' AND pd.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND pd.numero = '" & numero & "' "
                'tabla = tabla & "UNION "
                cSql1.sql = "SELECT CONCAT(pd.tipo, ' ', pd.documento, '" & tabulador & vbTab & "', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & tabulador & "', CONCAT('$ ', FORMAT(pd.monto,0)), '" & tabulador & "') AS item, DATE_FORMAT(dc.vencimiento,'%d-%m-%Y') AS fecha "
                cSql1.sql = cSql1.sql & "FROM sv_pagos_detalle_" & empresaActiva & " AS pd INNER JOIN sv_documentos_cobranza_" & empresaActiva & " AS dc ON pd.tipo = dc.tipo AND pd.documento = dc.numero "
                cSql1.sql = cSql1.sql & "WHERE pd.local = '" & codLocal & "' AND pd.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND pd.numero = '" & NUMERO & "' "
                cSql1.sql = cSql1.sql & "ORDER BY fecha ASC"
'                Call ConectarControlData(data2, servidor, baseVentas & codRubro, usuario, password, tabla)
                cSql1.Execute
                If cSql1.RowsAffected > 0 Then
                   Set resultados1 = cSql1.OpenResultset
'                    data2.Recordset.MoveFirst
                    While Not resultados1.EOF
                        cadena = tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
                        cadena = cadena & Replace(resultados1("item"), ",", ".")
                        dias = DateDiff("d", resultados1("fecha"), fecha)
                        cadena = cadena & "     DIAS PAGO " & dias
                        Impresion.AddItem cadena, True
                        'UNION
                        Impresion.Range(Impresion.Rows - 1, 19, Impresion.Rows - 1, 23).Merge
                        Impresion.Range(Impresion.Rows - 1, 24, Impresion.Rows - 1, 27).Merge
                        Impresion.Range(Impresion.Rows - 1, 28, Impresion.Rows - 1, 31).Merge
                        Impresion.Range(Impresion.Rows - 1, 32, Impresion.Rows - 1, Impresion.Cols - 1).Merge
                        'ALINEACION
                        Impresion.Range(Impresion.Rows - 1, 19, Impresion.Rows - 1, 23).Alignment = cellCenterCenter
                        Impresion.Range(Impresion.Rows - 1, 24, Impresion.Rows - 1, 27).Alignment = cellCenterCenter
                        Impresion.Range(Impresion.Rows - 1, 28, Impresion.Rows - 1, 31).Alignment = cellRightCenter
                        Impresion.Range(Impresion.Rows - 1, 32, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellLeftCenter
                        resultados1.MoveNext
                    Wend
                End If
                cSql1.Close
                Set cSql1 = Nothing
                Set resultados = Nothing
                
                Impresion.Range(Impresion.Rows - 1, 19, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Impresion.AddItem "", True
                resultados.MoveNext
            Wend
            cSql.Close
            Set cSql = Nothing
            Set resultados = Nothing
            
        End If
    End Sub

    Private Sub Cheques()
        Dim tabla As String
        Dim cadena As String
        Dim fila As Long
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim fecha As Double
        Dim dia As Double
        Dim rut As String
        Dim tabulador As String
        Dim resultados As rdoResultset
        Dim resultados1 As rdoResultset
        Dim cSql1 As New rdoQuery
        Dim cSql As New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
      
        
        'CHEQUES POR VENTAS
        tabulador = ""
        For i = 1 To 3
            tabulador = tabulador & vbTab
        Next i
'        cSql.sql = "SELECT CONCAT(CONCAT(dp.tipo,' ', dp.numero), '" & tabulador & vbTab & vbTab & "', CONCAT(dp.tipopago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', dp.banco, '" & tabulador & "', dp.plaza, '" & tabulador & "',dp.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(dp.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(dp.vencimiento,'%d-%m-%Y')) AS item3, IFNULL(dp.vencimiento,'') AS vencimiento, dp.monto, dc.rut "
'        cSql.sql = cSql.sql & "FROM sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero AND dc.nula = 'N'"
'        cSql.sql = cSql.sql & "WHERE dp.local = '" & codLocal & "' AND dp.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND dp.tipopago = '2' "
'        cSql.sql = cSql.sql & "ORDER BY dp.tipo, dp.numero ASC"
''        Call ConectarControlData(data, servidor, baseVentas & codRubro, usuario, password, tabla)
'        cSql.Execute
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
            'TITULO
            Impresion.AddItem "CHEQUES RECIBIDOR POR VENTAS", True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            fecha = 0
            dia = 0
            'CABEZA
            cadena = "DOCUMENTO" & tabulador & vbTab & vbTab
            cadena = cadena & "DETALLE CUENTA" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
            cadena = cadena & "BANCO" & tabulador
            cadena = cadena & "PLAZA" & tabulador
            cadena = cadena & "NUMERO" & tabulador & vbTab
            cadena = cadena & "MONTO" & tabulador & tabulador
            cadena = cadena & "VENCIMIENTO"
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            'UNION DE CELDAS
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
            Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
            Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
            Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            
'            data.Recordset.MoveFirst
            While Not resultados.EOF
                rut = resultados("rut")
                If resultados("vencimiento") > fechasistema Then
                    fecha = fecha + CDbl(resultados("monto"))
                Else
                    dia = dia + CDbl(resultados("monto"))
                End If
                Impresion.AddItem resultados("item1") & "   " & leerNombreCliente(rut) & resultados("item2") & Replace(resultados("item3"), ",", "."), True
                'UNION DE CELDAS
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
                'ALINEACION
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Alignment = cellCenterCenter
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
                resultados.MoveNext
            Wend
        End If
        cSql.Close
        Set cSql = Nothing
        Set resultados = Nothing
        
        Impresion.AddItem "", True
        
        'CHEQUES POR PAGOS DE CLIENTES
        tabulador = ""
        For i = 1 To 3
            tabulador = tabulador & vbTab
        Next i
        Set cSql1.ActiveConnection = ventasRubro
'        cSql1.sql = "SELECT CONCAT(CONCAT('PA', ' ', pc.numero), '" & tabulador & vbTab & vbTab & "', CONCAT(pc.tipopago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', c.banco, '" & tabulador & "', c.plaza, '" & tabulador & "',c.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(c.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(c.fechavencimiento,'%d-%m-%Y')) AS item3, IFNULL(c.fechavencimiento,'') AS fechavencimiento, c.monto, pc.rut , pc.numero "
'        cSql1.sql = cSql1.sql & "FROM sv_pagos_cabeza_" & empresaActiva & " AS pc INNER JOIN sv_carteracheques AS c ON pc.local = c.local AND pc.numero = c.numero AND c.tipodocumento = 'PA' "
'        cSql1.sql = cSql1.sql & "WHERE pc.local = '" & codLocal & "' AND pc.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND pc.tipopago = '2' "
'        cSql1.sql = cSql1.sql & "ORDER BY numero ASC"
'        cSql1.Execute
'
        'tabla = "SELECT CONCAT(CONCAT(pd.tipo,' ', pd.documento), '" & tabulador & vbTab & vbTab & "', CONCAT(pd.formapago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', c.banco, '" & tabulador & "', c.plaza, '" & tabulador & "',c.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(c.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(c.fechavencimiento,'%d-%m-%Y')) AS item3, IFNULL(c.fechavencimiento,'') AS fechavencimiento, c.monto, pd.rut "
        'tabla = tabla & "FROM sv_pagos_detalle AS pd INNER JOIN sv_carteracheques AS c ON pd.local = c.local AND pd.numero = c.numero /*INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON pd.rut = mc.rut*/ LEFT JOIN " & baseVentas & ".sv_maestrobancos AS mb ON c.banco = mb.codigobanco INNER JOIN sv_pagos_cabeza AS pc ON pd.local = pc.local AND pd.numero = pc.numero "
        'tabla = tabla & "WHERE pd.local = '" & codLocal & "' AND pd.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND pc.tipopago = '2' ORDER BY pd.numero ASC"
'        Call ConectarControlData(data, servidor, baseVentas & codRubro, usuario, password, tabla)
        
        If cSql1.RowsAffected > 0 Then
        Set resultados1 = cSql1.OpenResultset
            'TITULO
            Impresion.AddItem "CHEQUES RECIBIDOR POR PAGOS", True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            'fecha = 0
            'dia = 0
            'CABEZA
            cadena = "DOCUMENTO" & tabulador & vbTab & vbTab
            cadena = cadena & "DETALLE CUENTA" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
            cadena = cadena & "BANCO" & tabulador
            cadena = cadena & "PLAZA" & tabulador
            cadena = cadena & "NUMERO" & tabulador & vbTab
            cadena = cadena & "MONTO" & tabulador & tabulador
            cadena = cadena & "VENCIMIENTO"
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            'UNION DE CELDAS
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
            Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
            Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
            Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            
'            data.Recordset.MoveFirst
            While Not resultados1.EOF
                rut = resultados1("rut")
                If resultados1("fechavencimiento") > fechasistema Then
                    fecha = fecha + CDbl(resultados1("monto"))
                Else
                    dia = dia + CDbl(resultados1("monto"))
                End If
                Impresion.AddItem resultados1("item1") & "   " & leerNombreCliente(rut) & resultados1("item2") & Replace(resultados1("item3"), ",", "."), True
                'UNION DE CELDAS
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
                'ALINEACION
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Alignment = cellCenterCenter
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
                resultados1.MoveNext
            Wend
            Set cSql1 = Nothing
            cSql1.Close
            Set resultados1 = Nothing
            
            Impresion.AddItem "", True
            
            tabulador = ""
            For i = 1 To 5
                tabulador = tabulador & vbTab
            Next i
            Impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES A FECHA" & tabulador & tabulador & Format(fecha, "$ ###,###,##0"), True
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            Impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES AL DIA" & tabulador & tabulador & Format(dia, "$ ###,###,##0"), True
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            Impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES RECIBIDOS" & tabulador & tabulador & Format(dia + fecha, "$ ###,###,##0"), True
            Impresion.Range(Impresion.Rows - 2, 20, Impresion.Rows - 2, 35).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        End If
    End Sub
    
    Private Sub Creditos()
        Dim tabla As String
        Dim cadena As String
        Dim fila As Long
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim fecha As Double
        Dim dia As Double
        Dim rut As String
        Dim tabulador As String
        Dim resultados As rdoResultset
        Dim cSql As New rdoQuery
        Dim resultados1 As rdoResultset
        Dim cSql1 As New rdoQuery
        
        Set cSql.ActiveConnection = ventasRubro
        
        
        'LISTADO DE CREDITOS
        tabulador = ""
        For i = 1 To 6
            tabulador = tabulador & vbTab
        Next i
        cSql.sql = "SELECT CONCAT(CONCAT(dp.tipo,' ', dp.numero), '" & tabulador & vbTab & vbTab & "', CONCAT(dp.tipopago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', dp.banco, '" & tabulador & "', dp.plaza, '" & tabulador & "',dp.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(dp.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(dp.vencimiento,'%d-%m-%Y')) AS item3, IFNULL(dp.vencimiento,'') AS vencimiento, dp.monto, dc.rut "
        cSql.sql = cSql.sql & "FROM sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero AND dc.nula = 'N'"
        cSql.sql = cSql.sql & "WHERE dp.local = '" & codLocal & "' AND dp.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND dp.tipopago = '2' ORDER BY dp.tipo, dp.numero ASC"
        cSql.Execute
       
       If cSql.RowsAffected > 0 Then
       Set resultados = cSql.OpenResultset
            'TITULO
            Impresion.AddItem "CHEQUES RECIBIDOR POR VENTAS", True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            fecha = 0
            dia = 0
            'CABEZA
            cadena = "DOCUMENTO" & tabulador & vbTab & vbTab
            cadena = cadena & "DETALLE CUENTA" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
            cadena = cadena & "BANCO" & tabulador
            cadena = cadena & "PLAZA" & tabulador
            cadena = cadena & "NUMERO" & tabulador & vbTab
            cadena = cadena & "MONTO" & tabulador & tabulador
            cadena = cadena & "VENCIMIENTO"
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            'UNION DE CELDAS
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
            Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
            Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
            Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            
'            data.Recordset.MoveFirst
            While Not resultados.EOF
                rut = resultados("rut")
                If resultados("vencimiento") > fechasistema Then
                    fecha = fecha + CDbl(resultados("monto"))
                Else
                    dia = dia + CDbl(resultados("monto"))
                End If
                Impresion.AddItem resultados("item1") & "   " & leerNombreCliente(rut) & resultados("item2") & Replace(resultados("item3"), ",", "."), True
                'UNION DE CELDAS
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
                'ALINEACION
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Alignment = cellCenterCenter
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
                resultados.MoveNext
            Wend
        End If
        Set cSql = Nothing
        cSql.Close
        Set resultados = Nothing
        
        Impresion.AddItem "", True
        
        'CHEQUES POR PAGOS DE CLIENTES
        tabulador = ""
        For i = 1 To 3
            tabulador = tabulador & vbTab
        Next i
        
        cSql1.sql = "SELECT CONCAT(CONCAT('PA', ' ', pc.numero), '" & tabulador & vbTab & vbTab & "', CONCAT(pc.tipopago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', c.banco, '" & tabulador & "', c.plaza, '" & tabulador & "',c.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(c.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(c.fechavencimiento,'%d-%m-%Y')) AS item3, IFNULL(c.fechavencimiento,'') AS fechavencimiento, c.monto, pc.rut , pc.numero "
        cSql1.sql = cSql1.sql & "FROM sv_pagos_cabeza_" & empresaActiva & " AS pc INNER JOIN sv_carteracheques AS c ON pc.local = c.local AND pc.numero = c.numero AND c.tipodocumento = 'PA' "
        cSql1.sql = cSql1.sql & "ORDER BY numero ASC"
        cSql1.Execute
        
        'tabla = "SELECT CONCAT(CONCAT(pd.tipo,' ', pd.documento), '" & tabulador & vbTab & vbTab & "', CONCAT(pd.formapago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', c.banco, '" & tabulador & "', c.plaza, '" & tabulador & "',c.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(c.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(c.fechavencimiento,'%d-%m-%Y')) AS item3, IFNULL(c.fechavencimiento,'') AS fechavencimiento, c.monto, pd.rut "
        'tabla = tabla & "FROM sv_pagos_detalle AS pd INNER JOIN sv_carteracheques AS c ON pd.local = c.local AND pd.numero = c.numero /*INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON pd.rut = mc.rut*/ LEFT JOIN " & baseVentas & ".sv_maestrobancos AS mb ON c.banco = mb.codigobanco INNER JOIN sv_pagos_cabeza AS pc ON pd.local = pc.local AND pd.numero = pc.numero "
        'tabla = tabla & "WHERE pd.local = '" & codLocal & "' AND pd.fecha BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND pc.tipopago = '2' ORDER BY pd.numero ASC"
 '       Call ConectarControlData(data, servidor, baseVentas & codRubro, usuario, password, tabla)
        If cSql1.RowsAffected > 0 Then
            'TITULO
            Set resultados1 = cSql1.OpenResultset
            Impresion.AddItem "CHEQUES RECIBIDOR POR PAGOS", True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            'fecha = 0
            'dia = 0
            'CABEZA
            cadena = "DOCUMENTO" & tabulador & vbTab & vbTab
            cadena = cadena & "DETALLE CUENTA" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
            cadena = cadena & "BANCO" & tabulador
            cadena = cadena & "PLAZA" & tabulador
            cadena = cadena & "NUMERO" & tabulador & vbTab
            cadena = cadena & "MONTO" & tabulador & tabulador
            cadena = cadena & "VENCIMIENTO"
            Impresion.AddItem cadena, True
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            'UNION DE CELDAS
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
            Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
            Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
            Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            
'            data.Recordset.MoveFirst
            While Not resultados1.EOF
                rut = resultados1("rut")
                If resultados1("fechavencimiento") > fechasistema Then
                    fecha = fecha + CDbl(resultados1("monto"))
                Else
                    dia = dia + CDbl(resultados1("monto"))
                End If
                Impresion.AddItem resultados1("item1") & "   " & leerNombreCliente(rut) & resultados1("item2") & Replace(resultados1("item3"), ",", "."), True
                'UNION DE CELDAS
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Merge
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Merge
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Merge
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Merge
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Merge
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Merge
                'ALINEACION
                Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, 5).Alignment = cellCenterCenter
                Impresion.Range(Impresion.Rows - 1, 6, Impresion.Rows - 1, 19).Alignment = cellLeftCenter
                Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 22).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 23, Impresion.Rows - 1, 25).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 26, Impresion.Rows - 1, 29).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
                Impresion.Range(Impresion.Rows - 1, 36, Impresion.Rows - 1, Impresion.Cols - 1).Alignment = cellCenterCenter
            resultados1.MoveNext
            Wend
            Set cSql1 = Nothing
            cSql1.Close
            Set resultados = Nothing
            Impresion.AddItem "", True
            
            tabulador = ""
            For i = 1 To 5
                tabulador = tabulador & vbTab
            Next i
            Impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES A FECHA" & tabulador & tabulador & Format(fecha, "$ ###,###,##0"), True
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            Impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES AL DIA" & tabulador & tabulador & Format(dia, "$ ###,###,##0"), True
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
            
            Impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES RECIBIDOS" & tabulador & tabulador & Format(dia + fecha, "$ ###,###,##0"), True
            Impresion.Range(Impresion.Rows - 2, 20, Impresion.Rows - 2, 35).Borders(cellEdgeBottom) = cellThin
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Merge
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Merge
            Impresion.Range(Impresion.Rows - 1, 20, Impresion.Rows - 1, 29).Alignment = cellLeftCenter
            Impresion.Range(Impresion.Rows - 1, 30, Impresion.Rows - 1, 35).Alignment = cellRightCenter
            Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        End If
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmCerrar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        Impresion.PageSetup.HeaderMargin = 2
    
        Impresion.PageSetup.TopMargin = 2
        Impresion.PageSetup.LeftMargin = 0.5
        Impresion.PageSetup.RightMargin = 0.5
        Impresion.PageSetup.BottomMargin = 2
        
        Impresion.PageSetup.FooterMargin = 2
        Impresion.PageSetup.BlackAndWhite = True
        
        Call verificaImpresora(5, Impresion)
        
    End Sub


