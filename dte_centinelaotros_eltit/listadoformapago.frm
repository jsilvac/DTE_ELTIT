VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form listadoformapago 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Ventas"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   11475
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   120
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   7260
      Left            =   60
      TabIndex        =   9
      Top             =   1500
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   12806
      BackColor       =   16744576
      Caption         =   "Informe"
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
         Height          =   6780
         Left            =   45
         TabIndex        =   10
         Top             =   360
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   11959
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1335
      Left            =   45
      TabIndex        =   6
      Top             =   90
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   2355
      BackColor       =   16744576
      Caption         =   "Ingreso de Información"
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
      Begin VB.ComboBox formasdepago 
         Height          =   315
         Left            =   675
         TabIndex        =   13
         Text            =   "2 - CHEQUE PROPIO"
         Top             =   810
         Width           =   3210
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Generar Informe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8505
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   495
         Width           =   2040
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   6540
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   6060
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   7020
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
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
         Height          =   315
         Left            =   660
         TabIndex        =   8
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
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
         Height          =   315
         Left            =   4620
         TabIndex        =   7
         Top             =   420
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   4095
      TabIndex        =   11
      Top             =   8865
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
Attribute VB_Name = "listadoformapago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private TIPO As String
    Private detalle As Boolean
    Private fecha1 As String
    Private fecha2 As String

Private Sub Command1_Click()
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
           Call CargaGrillaInforme(1, 4)
            Call generaInforme
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
    End Sub

    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub

    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub

    Private Sub dato5_GotFocus()
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub

    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato5)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            SendKeys "{Tab}"
            Call CargaGrillaInforme(1, 4)
            Call generaInforme
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
       
    '========================================================
    'LostFocus
    '========================================================
  
Private Sub dato1_LostFocus()
Call limpiaBarra(2)
Call esfecha(dato1, dato2, dato3, "dd")
End Sub
Private Sub dato2_LostFocus()
Call esfecha(dato1, dato2, dato3, "mm")
End Sub
Private Sub dato3_LostFocus()
Call esfecha(dato1, dato2, dato3, "yyyy")
End Sub

Private Sub dato4_LostFocus()
Call esfecha(dato4, dato5, dato6, "dd")
End Sub
Private Sub dato5_LostFocus()
Call esfecha(dato4, dato5, dato6, "mm")
End Sub
Private Sub dato6_LostFocus()
Call esfecha(dato4, dato5, dato6, "yyyy")
End Sub

    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 27
                Unload Me
            Case 38
                If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                    Unload Me
                End If
        End Select
    End Sub
    
    Private Sub Form_Load()
        Call Centrar(Me)
        Call CargaGrillaInforme(1, 4)
        Call leerTipos
        detalle = False
        dato1.text = "01"
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        dato4.text = Format(fechasistema, "dd")
        dato5.text = Format(fechasistema, "mm")
        dato6.text = Format(fechasistema, "yyyy")
    
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "TIPO DOCUMENTO"
        formatoGrilla(1, 2) = "NUMERO DOCUMENTO"
        formatoGrilla(1, 3) = "TOTAL VENTA"
    
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "4"
        formatoGrilla(2, 2) = "10"
        formatoGrilla(2, 3) = "10"
       
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "C"
        formatoGrilla(3, 3) = "N"
       
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
       
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
       
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
       
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
       
        Rem ANCHO
        formatoGrilla(8, 1) = "15"
        formatoGrilla(8, 2) = "20"
        formatoGrilla(8, 3) = "10"
        
        Impresion.Cols = col
        Impresion.Rows = row
        Impresion.Range(0, 0, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        Impresion.Range(0, 0, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        Impresion.Range(0, 0, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        Impresion.AllowUserResizing = False
        Impresion.DisplayFocusRect = False
        Impresion.ExtendLastCol = True
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
        Impresion.RowHeight(0) = Impresion.DefaultRowHeight * 1.75
        Impresion.Range(0, 1, 0, Impresion.Cols - 1).WrapText = True
        
        For i = 1 To Impresion.Cols - 1
            Impresion.Cell(0, i).text = formatoGrilla(1, i)
            Impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (Impresion.Cell(0, i).Font.Size + 1.25)
            Impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            Impresion.Column(i).FormatString = formatoGrilla(4, i)
            Impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                Impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                Impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                Impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Impresion.Range(0, 1, 0, Impresion.Cols - 1).Alignment = cellCenterCenter
        Impresion.Range(0, 1, 0, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
       
        
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************



Private Sub formasdepago_DblClick()
Call Command1_Click
End Sub

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        empresa
        Call imprimir
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        
        Call Titulos("INFORME VENTAS X TIPO DE PAGO")


        
        Impresion.AutoRedraw = False
        Impresion.Range(1, 1, 1, 3).Borders(cellEdgeTop) = cellThick
        Impresion.PageSetup.HeaderMargin = 0.5
    
        Impresion.PageSetup.TopMargin = 1
        Impresion.PageSetup.LeftMargin = 0.5
        Impresion.PageSetup.RightMargin = 0
        Impresion.PageSetup.BottomMargin = 2
     
        
        Impresion.PageSetup.FooterMargin = 1
        Impresion.PageSetup.BlackAndWhite = True
        Impresion.PageSetup.Orientation = cellPortrait
        Impresion.PageSetup.PrintFixedRow = True
        Impresion.PageSetup.BlackAndWhite = True
        
        Call verificaImpresora(5, Impresion)
        
        Impresion.AutoRedraw = True
    End Sub
    
  
    
    
    Sub leerTipos()
       
        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql & "FROM sv_tiposdepagoclientes "
        cSql.sql = cSql.sql & "order by codigo asc "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

        Set resultado = cSql.OpenResultset
      
        While Not resultado.EOF
         If resultado(1) <> "VUELTO" Then
         If resultado(1) <> "EFECTIVO" Then
         formasdepago.AddItem CDbl(resultado(0)) & " - " & resultado(1)
         End If
         End If
       
            resultado.MoveNext
            Wend
        Else
      

        End If
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
    End Sub
Sub generaInforme()
 Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim linea As Double
        Dim total As Double
           
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventasRubro

        cSql.sql = "SELECT tipo,numero,monto "
        cSql.sql = cSql.sql & "FROM sv_documento_pagos_" + empresaActiva + " "
        cSql.sql = cSql.sql & "where tipopago='" + Mid(formasdepago.text, 1, 2) + "' and fecha between '" + fecha1 + "' and '" + fecha2 + "' "
        cSql.Execute
        total = 0
        If cSql.RowsAffected > 0 Then

        Set resultado = cSql.OpenResultset
       Impresion.AutoRedraw = False
        
        Impresion.Rows = cSql.RowsAffected + 2
        linea = 1
        While Not resultado.EOF
        linea = linea + 1
            total = total + resultado(2)
            Impresion.Cell(linea, 1).text = resultado(0)
            Impresion.Cell(linea, 2).text = resultado(1)
            Impresion.Cell(linea, 3).text = Format(resultado(2), "$ ###,###,###")
           
            resultado.MoveNext
            Wend
         

        End If
        Set resultado = Nothing
        cSql.Close
         
        Set cSql = Nothing
        Impresion.AutoRedraw = True
        Impresion.Refresh
        If total > 0 Then
        linea = linea + 1
        Impresion.Rows = Impresion.Rows + 1
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).Borders(cellEdgeLeft) = cellThick
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).Borders(cellEdgeRight) = cellThick
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThick
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).FontSize = 8
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).FontBold = True
        
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).BackColor = RGB(90, 158, 214)
        Impresion.Range(linea, 1, linea, Impresion.Cols - 1).ForeColor = 0
        Impresion.Cell(linea, 1).text = "TOTAL VENTAS "
        Impresion.Cell(linea, Impresion.Cols - 1).text = Format(total, "$ ###,###,###")
        
       
       Impresion.AutoRedraw = True
       Impresion.Refresh
       End If
        
End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Impresion.FixedRowColStyle = Fixed3D
    Impresion.CellBorderColorFixed = vbButtonShadow
    Impresion.ShowResizeTips = False
    Impresion.ReportTitles.Clear
    Impresion.PageSetup.CenterHorizontally = True
    Impresion.PageSetup.Orientation = cellPortrait
    
      
    Impresion.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Impresion.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Impresion.PageSetup.HeaderAlignment = cellLeft
    Impresion.PageSetup.HeaderFont.Name = "Verdana"
    Impresion.PageSetup.HeaderFont.Size = 8
    Impresion.PageSetup.HeaderFont.Italic = True
    
    'TITULOS DEL REPORTE
  
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Impresion.ReportTitles.Add objReportTitle
  
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FORMA DE PAGO: " + Replace(Mid(formasdepago.text, "4", "20"), "-", "") & "  |  " & "Periodo Desde " & Format(fecha1, "dd-mm-yyyy") & " Hasta " & Format(fecha2, "dd-mm-yyyy") & " "
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Impresion.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Impresion.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" + " usuario:" + usuarioSistema
    Impresion.PageSetup.FooterAlignment = cellRight
    Impresion.PageSetup.FooterFont.Name = "Verdana"
    Impresion.PageSetup.FooterFont.Size = 7

    
End Sub
