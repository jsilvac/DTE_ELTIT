VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form detallePagos 
   BackColor       =   &H00AE1118&
   BorderStyle     =   0  'None
   Caption         =   "Detalle de Pagos"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   9128
      BackColor       =   16744576
      Caption         =   "Detalle Pago"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox pivote 
         Height          =   315
         Left            =   4800
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdEsc 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ESC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4020
         Width           =   735
      End
      Begin FlexCell.Grid pagos 
         Height          =   3480
         Left            =   0
         TabIndex        =   0
         Top             =   360
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   6138
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
      Begin VB.CommandButton cmdF9 
         BackColor       =   &H00FFC0C0&
         Caption         =   " F9 o *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4620
         Visible         =   0   'False
         Width           =   735
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   3900
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2143
         BackColor       =   16744576
         Caption         =   "CANCELADO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
         ColorTextShadow =   0
         Begin VB.Label lblCancelado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   33
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   765
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   4335
         End
      End
      Begin FlexCell.Grid Impresion 
         Height          =   735
         Left            =   60
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   1980
         Top             =   3840
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label lblVolver 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VOLVER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   6120
         TabIndex        =   10
         Top             =   4020
         Width           =   3135
      End
      Begin VB.Label lblTerminar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FINALIZAR VENTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   6120
         TabIndex        =   2
         Top             =   4620
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin XPFrame.FrameXp FrameXp5 
      Height          =   1575
      Left            =   135
      TabIndex        =   3
      Top             =   315
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   2778
      BackColor       =   16744576
      Caption         =   "TOTAL A PAGAR"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   0
      Begin VB.Label lblTotalPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   44.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   1005
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   10110
      End
   End
End
Attribute VB_Name = "detallePagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private p As pagos
    Private formatogrilla(20, 20) As String
   ' Private tipopago As String
    Private modifica As Boolean
    Public lectura As Boolean
    Private resto As Double
    Private grabo As Boolean

Private Sub cmdEsc_Click()
    PVentas.imprimio = False
    If modifica = False Then
        Unload Me
    Else
        If grabo = True Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdF9_Click()
    If resto <= 0 Then
        Call ctrltostruct
        If modifica = False Then
           
                Select Case PVentas.dato1.text
                    Case "BV"
                        numeroboleta = PVentas.dato2.text
                        impresionboleta.Show vbModal
                        
                        
                    Case "FV", "FE"
'                        Call imprimeFactura(PVentas.dato1.text, PVentas.dato27.text, impresion, PVentas.data)
                    Case "GD", "GM"
                        Call imprimeGuia(PVentas.dato1.text, PVentas.dato27.text, impresion, PVentas.data)
                End Select
            
        End If
        modifica = False
        grabo = True
        Unload Me
        PVentas.dato1.Enabled = True
        PVentas.retorno
        PVentas.dato1.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If lectura = False Then
        pagos.Cell(1, 1).SetFocus
        cmdF9.Caption = "F9 o *"
        lblTerminar.Caption = "IMPRIMIR VENTA"
    Else
        cmdF9.Caption = "F8"
        lblTerminar.Caption = "MODIFICAR PAGO"
        cmdF9.Visible = True
        lblTerminar.Visible = True
        modifica = True
        lectura = False
        grabo = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF9, 106
            Call cmdF9_Click
            PVentas.imprimio = True
        Case vbKeyF8
            Call eliminarPagos(PVentas.dato1.text, PVentas.dato27.text, PVentas.dato5.text + "-" + PVentas.dato4.text + "-" + PVentas.dato3.text, PVentas.dato30.text)
            Call eliminarDocManual(PVentas.dato1.text, PVentas.dato27.text)
'            Call eliminarCheque(PVentas.dato1.text, PVentas.dato2.text)
            pagos.Rows = 1
            pagos.SelectionMode = cellSelectionFree
            pagos.AddItem "", True
            pagos.Cell(1, 1).SetFocus
            cmdF9.Caption = "F9 o *"
            lblTerminar.Caption = "IMPRIMIR VENTA"
            cmdF9.Visible = False
            lblTerminar.Visible = False
            grabo = False
        Case 27
            Call cmdEsc_Click
    End Select
End Sub

'****************************************************************************
'Formato de la Grilla Pagos
'****************************************************************************
    Private Sub CargaGrillaPagos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TIPO PAGO"
        formatogrilla(1, 2) = "MONTO"
        formatogrilla(1, 3) = "NUMERO"
        formatogrilla(1, 4) = "BANCO"
        formatogrilla(1, 5) = "CUENTA"
        formatogrilla(1, 6) = "VENCIMIENTO"
        formatogrilla(1, 7) = ""
        formatogrilla(1, 8) = ""
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "25"
        formatogrilla(2, 2) = "9"
        formatogrilla(2, 3) = "7"
        formatogrilla(2, 4) = "3"
        formatogrilla(2, 5) = "11"
        formatogrilla(2, 6) = "2"
        formatogrilla(2, 7) = "2"
        formatogrilla(2, 8) = "4"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = "$ ###,###,##0"
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = "00"
        formatogrilla(4, 7) = "00"
        formatogrilla(4, 8) = "0000"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "20"
        formatogrilla(8, 2) = "10"
        formatogrilla(8, 3) = "12"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "3"
        formatogrilla(8, 7) = "3"
        formatogrilla(8, 8) = "4"
            
        pagos.Cols = col
        pagos.Rows = row
        pagos.AllowUserResizing = False
        pagos.DisplayFocusRect = False
        pagos.ExtendLastCol = True
        pagos.BoldFixedCell = False
        pagos.DrawMode = cellOwnerDraw
        pagos.Appearance = Flat
        pagos.ScrollBarStyle = Flat
        pagos.FixedRowColStyle = Flat
        pagos.BackColorFixed = RGB(90, 158, 214)
        pagos.BackColorFixedSel = RGB(110, 180, 230)
        pagos.BackColorBkg = RGB(90, 158, 214)
        pagos.BackColorScrollBar = RGB(231, 235, 247)
        pagos.BackColor1 = RGB(231, 235, 247)
        pagos.BackColor2 = RGB(239, 243, 255)
        pagos.GridColor = RGB(148, 190, 231)
        
        pagos.Column(0).Width = 0
        For i = 1 To col - 1
            pagos.Cell(0, i).text = formatogrilla(1, i)
            pagos.Column(i).Width = Val(formatogrilla(8, i)) * (pagos.Cell(0, i).Font.Size + 1.25)
            pagos.Column(i).MaxLength = Val(formatogrilla(2, i))
            pagos.Column(i).FormatString = formatogrilla(4, i)
            pagos.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                pagos.Column(i).Alignment = cellRightCenter
            Else
                pagos.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        
        pagos.Range(0, 6, 0, pagos.Cols - 1).Merge
        pagos.Range(0, 0, 0, pagos.Cols - 1).Alignment = cellCenterCenter
        pagos.Column(1).CellType = cellComboBox
        'pagos.Column(6).CellType = cellCalendar
         leerTipoPago
       
        pagos.ComboBox(1).AutoComplete = True
        
    End Sub

        Sub leerTipoPago()
       
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql & "FROM sv_tiposdepagoclientes "
        csql.sql = csql.sql & "order by codigo asc "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

        Set resultado = csql.OpenResultset
      
        While Not resultado.EOF
        
          pagos.ComboBox(1).AddItem CDbl(resultado(0)) & " - " & resultado(1)
       
            resultado.MoveNext
            Wend
        Else
      

        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Sub
   Function leerTipoPagoIndividual(nombre) As String
       
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql & "FROM sv_tiposdepagoclientes "
        csql.sql = csql.sql & "where nombre='" & nombre & "' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

        Set resultado = csql.OpenResultset
      
        While Not resultado.EOF
        
        leerTipoPagoIndividual = resultado(0) & " - " & resultado(1)
       
            resultado.MoveNext
            Wend
        Else
      

        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function

Private Sub Form_Load()
    modifica = False
    lectura = False
    grabo = False
    Call CargaGrillaPagos(2, 9)
    lblTotalPagar.Caption = Format(PVentas.dato26.text, "$ ###,###,##0")
    pagos.Cell(1, 2).text = saldo(1, 1)
End Sub


Private Sub pagos_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 13 Then
          
        Select Case pagos.ActiveCell.col
            Case 1
                 
                If autorizado = False Then
                    Select Case Left(pagos.ActiveCell.text, 1)
                        Case "2", "3"
                            If verificarChequesCliente(rut_cliente, sucursal_cliente) = True Then
                                Call enviarInformacion(rut_cliente, sucursal_cliente, PVentas.dato1.text, PVentas.dato2.text, Format(PVentas.lblTotal.Caption, "########0"), "CHEQUES PENDIENTES")
                                'Call mensaje.mostrarMensaje("Información", "El cliente " & rut_cliente & " posee cheques protestados o prorrogados.", "Solicite autorizaión")
                            End If
                        Case "9"
                            If verificarFacturasCliente(rut_cliente, sucursal_cliente) = True Then
                                Call enviarInformacion(rut_cliente, sucursal_cliente, PVentas.dato1.text, PVentas.dato2.text, Format(PVentas.lblTotal.Caption, "########0"), "DOCUMENTOS PENDIENTES")
                                'Call mensaje.mostrarMensaje("Información", "El cliente " & rut_cliente & " posee documentos pendientes de pago.", "Solicite autorizaión")
                            End If
                        Case "8"
                        Load creditoTMP
                        
                            creditoTMP.rut2.text = PVentas.dato6.text
                            creditoTMP.NUMERO.text = PVentas.dato2.text
                            creditoTMP.TIPO.text = PVentas.dato1.text
                            
                            
                            creditoTMP.lbldv.Caption = PVentas.lbldv.Caption
                            
                            
                            creditoTMP.MONTO = Format(pagos.Cell(pagos.ActiveCell.row, 2).text, "###,###,##0")
                            creditoTMP.MONTODOCUMENTO.Caption = creditoTMP.MONTO.text
                            
                            creditoTMP.NUMERO.Locked = True
                            
                            
                            creditoTMP.Show vbModal
                            
                    End Select
                End If
            Case 3
                If pagos.ActiveCell.text <> "" Then
                    pivote.MaxLength = 7
                    pivote.text = pagos.ActiveCell.text
                    pivote.text = ceros(pivote)
                    pagos.ActiveCell.text = pivote.text
                End If
            Case 4
                If pagos.ActiveCell.text <> "" Then
                    pivote.MaxLength = 3
                    pivote.text = pagos.ActiveCell.text
                    pivote.text = ceros(pivote)
                    pagos.ActiveCell.text = pivote.text
                End If
            Case 6
                If pagos.ActiveCell.text = "" Or pagos.ActiveCell.text = "00" Then
                    pagos.ActiveCell.text = Format(fechasistema, "dd")
                    'vacio = False
                    'Call pagos_LeaveCell(pagos.ActiveCell.row, 6, pagos.ActiveCell.row, 7, False)
                End If
            Case 7
                If pagos.ActiveCell.text = "" Or pagos.ActiveCell.text = "00" Then
                    pagos.ActiveCell.text = Format(fechasistema, "mm")
                    'vacio = False
                    'Call pagos_LeaveCell(pagos.ActiveCell.row, 7, pagos.ActiveCell.row, 8, False)
                End If
            Case 8
                If pagos.ActiveCell.text = "" Or pagos.ActiveCell.text = "0000" Then
                    pagos.ActiveCell.text = Format(fechasistema, "yyyy")
                    'vacio = False
                    'Call pagos_LeaveCell(pagos.ActiveCell.row, 8, pagos.ActiveCell.row, 1, False)
                End If
        End Select
    End If
End Sub

Private Sub pagos_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)



End Sub

Private Sub Pagos_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    Dim i As Integer
   
    ''''''''''''''''''''''''''''''''''
    'MANEJO DE LOS MOVIMIENTOS
    ''''''''''''''''''''''''''''''''''
    If lectura = False Then
        If row = NewRow Then
            If pagos.Cell(row, col).text = "" Then
                If NewCol > col Then
                    NewCol = col
                Else
                    If NewCol = 1 And pagos.Cols - 1 = col Then
                        NewCol = col
                    End If
                End If
            Else
                If NewCol - 1 > col Then
                    NewCol = col
                End If
            End If
            If col = pagos.Cols - 1 Then
                If NewCol = 1 Then
                    If pagos.Rows - 1 = NewRow Then
                        pagos.Rows = pagos.Rows + 1
                        NewCol = 1
                        NewRow = pagos.Rows - 1
                    End If
                End If
            End If
        Else
            For i = 1 To pagos.Cols - 1
                If pagos.Cell(NewRow, i).text = "" Then
                    NewCol = i
                    Exit For
                End If
            Next i
        End If
        ''''''''''''''''''''''''''''''''''
        'MANEJO DE LOS MOVIMIENTOS
        ''''''''''''''''''''''''''''''''''
            
        ''''''''''''''''''''''''''''''''''
        'VALIDACIONES PARA ESTA GRILLA
        ''''''''''''''''''''''''''''''''''
        If row = NewRow Then
            If pagos.Cell(row, 1).text <> "" Then
                tipopago = pagos.Cell(row, 1).text
                tipopago = Left(tipopago, 1)
            End If
            Select Case NewCol
                Case 1
                    
                Case 2
                    
                Case 3
                    Select Case tipopago
                        Case "1"
                            If pagos.ActiveCell.row <= pagos.Rows - 1 Then
                                pagos.Range(row, NewCol, row, pagos.Cols - 1).ClearText
                                If pagos.Rows - 1 = row Then
                                    pagos.Rows = pagos.Rows + 1
                                End If
                                pagos.Range(row, NewCol, row, pagos.Cols - 1).Locked = True
                                NewCol = 1
                                NewRow = pagos.Rows - 1
                            End If
                            pagos.Cell(pagos.Rows - 1, 2).text = saldo(row, NewRow)
                        Case "2", "3"
                            pagos.Range(row, NewCol, row, pagos.Cols - 1).Locked = False
                        'Case "5"
                        '    If pagos.ActiveCell.row <= pagos.Rows - 1 Then
                        '        pagos.Range(row, NewCol, row, pagos.Cols - 1).ClearText
                        '        pagos.Rows = pagos.Rows + 1
                        '        pagos.Range(row, NewCol, row, pagos.Cols - 1).Locked = True
                        '        Load credito
                        '        credito.txtRut.text = PVentas.dato6.text
                        '        credito.lblDV.Caption = PVentas.lblDV.Caption
                        '        'Credito.lblMonto.Caption = Pagos.ActiveCell.text
                        '        credito.Show vbModal
                        '        NewCol = 1
                        '        NewRow = pagos.Rows - 1
                        '    Else
                        '        If pagos.ActiveCell.row <= pagos.Rows - 1 Then
                        '            pagos.Range(row, NewCol, row, pagos.Cols - 1).ClearText
                        '            pagos.Range(row, NewCol, row, pagos.Cols - 1).Locked = True
                        '            Load credito
                        '            credito.txtRut.text = PVentas.dato6.text
                        '            credito.lblDV.Caption = PVentas.lblDV.Caption
                        '            'Credito.lblMonto.Caption = Pagos.ActiveCell.text
                        '            credito.Show vbModal
                        '            NewCol = 1
                        '            NewRow = pagos.Rows - 1
                        '        End If
                        '    End If
                        '    pagos.Cell(pagos.Rows - 1, 2).text = saldo(row, NewRow)
                        
                        Case 4, 5, 6, 7, 9, 8, 10
                            If pagos.ActiveCell.row <= pagos.Rows - 1 Then
                                pagos.Range(row, NewCol, row, pagos.Cols - 1).ClearText
                                If pagos.Rows - 1 = row Then
                                    pagos.Rows = pagos.Rows + 1
                                End If
                                pagos.Range(row, NewCol, row, pagos.Cols - 1).Locked = True
                                NewCol = 1
                                NewRow = pagos.Rows - 1
                            End If
                            pagos.Cell(pagos.Rows - 1, 2).text = saldo(row, NewRow)
                        Case 11
                        
                            If pagos.ActiveCell.row <= pagos.Rows - 1 Then
                                pagos.Range(row, NewCol, row, pagos.Cols - 1).ClearText
                                pagos.Range(row, col, row, pagos.Cols - 1).Locked = True
                            End If
                    End Select
                Case 4
                Case 5
                Case 6
            End Select
        Else
            pagos.Cell(pagos.Rows - 1, 2).text = saldo(row, NewRow)
            If pagos.Cell(NewRow, 1).text <> "" Then
                tipopago = pagos.Cell(NewRow, 1).text
                tipopago = Left(tipopago, 1)
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''
    'VALIDACIONES PARA ESTA GRILLA
    ''''''''''''''''''''''''''''''''''
End Sub

Private Function saldo(ByVal row As Long, ByVal NewRow As Long) As Double
    Dim i As Long
    saldo = 0
    For i = 1 To pagos.Rows - 2
        If pagos.Cell(i, 2).text <> "" Then
            saldo = saldo + CDbl(pagos.Cell(i, 2).text)
        End If
    Next i
    lblCancelado.Caption = Format(saldo, "$ ###,###,##0")
    saldo = CDbl(Right(lblTotalPagar.Caption, Len(lblTotalPagar.Caption) - 1)) - saldo
    resto = saldo
    If saldo <= 0 And pagos.Cell(1, 1).text <> "" And NewRow > row Then
        pagos.Cell(NewRow, 1).text = leerTipoPagoIndividual("VUELTO")
        cmdF9.Visible = True
        lblTerminar.Visible = True
        cmdF9.SetFocus
    End If
End Function

    Private Sub ctrltostruct()
        p.tipodocumento = PVentas.dato1.text
        p.rut = PVentas.dato6.text & PVentas.lbldv.Caption
        p.sucursal = PVentas.dato7.text
        p.numeroDocumento = ceros(PVentas.dato27)
        p.foliosii = PVentas.dato2.text
        
        p.fecha = PVentas.dato5.text & "-" & PVentas.dato4.text & "-" & PVentas.dato3.text
        p.vencDoc = PVentas.dato5.text & "-" & PVentas.dato4.text & "-" & PVentas.dato3.text
        Call eliminarPagos(p.tipodocumento, p.numeroDocumento, Format(p.fecha, "yyyy-mm-dd"), PVentas.dato30.text)
        Call grabarPagos(pagos, p, modifica, PVentas.dato7.text)
        'Call retorno
    End Sub

    Private Function verificarChequesCliente(ByVal rut As String, ByVal sucursal As String) As Boolean
        
        Dim CAMPOS(3, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "protestos"
        CAMPOS(1, 0) = "prorrogas"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            If sql.response(0, 3) <> "0" Or sql.response(1, 3) <> "0" Then
                verificarChequesCliente = True
                Call bloquearCliente(rut, sucursal)
            Else
                verificarChequesCliente = False
            End If
        Else
            verificarChequesCliente = False
        End If
    End Function

    Public Function verificarFacturasCliente(ByVal rut As String, ByVal sucursal As String) As Boolean
        
        Dim CAMPOS(3, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "IFNULL(COUNT(*),0)"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_documentos_cobranza_" & empresaActiva
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "' AND (tipo = 'BV' OR tipo = 'FV') AND monto > abono AND vencimiento <= '" & fechasistema & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            If sql.response(0, 3) <> "0" Then
                verificarFacturasCliente = True
                Call bloquearCliente(rut, sucursal)
            Else
                verificarFacturasCliente = False
            End If
        Else
            verificarFacturasCliente = False
        End If
    End Function



