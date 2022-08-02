VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ComprasCliente 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historico de Compras por Clientes por Kilos"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15195
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1755
      Left            =   3300
      TabIndex        =   4
      Top             =   60
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3096
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
      Begin XPFrame.FrameXp frmIndividual 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1720
         BackColor       =   16744576
         Caption         =   "Informe por Cliente"
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
            Left            =   1560
            MaxLength       =   9
            TabIndex        =   2
            Tag             =   "proveedor"
            Top             =   420
            Width           =   1095
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
            Left            =   3060
            MaxLength       =   1
            TabIndex        =   3
            Tag             =   "proveedor"
            Top             =   420
            Width           =   375
         End
         Begin VB.Label lblDV 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H0080FFFF&
            Height          =   315
            Left            =   2700
            TabIndex        =   8
            Top             =   420
            Width           =   315
         End
         Begin VB.Label lblNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H0080FFFF&
            Height          =   315
            Left            =   3480
            TabIndex        =   7
            Top             =   420
            Width           =   4695
         End
         Begin VB.Label lbl2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cliente"
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
            Left            =   120
            TabIndex        =   6
            Top             =   420
            Width           =   1335
         End
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4560
         TabIndex        =   1
         Top             =   420
         Width           =   2595
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Individual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   420
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   120
      Top             =   8040
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
      Height          =   6015
      Left            =   60
      TabIndex        =   9
      Top             =   1920
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   10610
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
         Height          =   5535
         Left            =   60
         TabIndex        =   10
         Top             =   420
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   9763
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   11760
      TabIndex        =   11
      Top             =   8040
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
Attribute VB_Name = "ComprasCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private tipo As String
    Private posX As Single
    Private posY As Single
        
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Clientes"
    End Sub

    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case vbKeyF2
                Call ayudaCliente(dato1, dato2, lblDV)
            Case Else
                Call Flechas(KeyCode, dato1)
        End Select
    End Sub

    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato1.text <> "" Then
            dato1.text = ceros(dato1)
            lblDV.Caption = rut(dato1.text)
            SendKeys "{Tab}"
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            lblNombre.Caption = leerNombreClienteSucursal(dato1.text & lblDV.Caption, dato2.text)
            If lblNombre.Caption <> "" Then
                Call generaInformeHCK(data, impresion, tipo, dato1.text & lblDV.Caption, dato2.text)
            End If
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
        Call CargaGrillaInforme(1, 15)
        tipo = "(dd.tipo = 'FV' OR dd.tipo = 'FE' OR dd.tipo = 'NV' OR dd.tipo = 'BV' OR dd.tipo = 'ZE' OR dd.tipo = 'FE')"
    End Sub

    Private Sub opt1_Click()
        If opt1.Value = True Then
            Unload ComprasClienteGrafico
            impresion.Rows = 1
            frmIndividual.Enabled = True
            dato1.SetFocus
        End If
    End Sub

    Private Sub opt2_Click()
        If opt2.Value = True Then
            frmIndividual.Enabled = False
            dato1.text = ""
            dato2.text = ""
            lblDV.Caption = ""
            lblNombre.Caption = ""
            Unload ComprasClienteGrafico
            Call generaInformeHCK(data, impresion, tipo, dato1.text & lblDV.Caption, dato2.text)
        End If
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        Dim ancho As Integer
        Call leerMeses
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CLIENTE"
        formatoGrilla(1, 2) = "AÑO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "50"
        formatoGrilla(2, 2) = "4"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "20"
        formatoGrilla(8, 2) = "4"
        
        If cantMeses > 2 Then
            ancho = Val(72 / cantMeses)
        Else
            ancho = Val(104 / (cantMeses + 3))
        End If
        For i = 1 To cantMeses
            If cantMeses > 10 Then
                formatoGrilla(1, i + 2) = Left(meses(i), 8)
            Else
                formatoGrilla(1, i + 2) = meses(i)
            End If
            
            formatoGrilla(2, i + 2) = "10"
            
            formatoGrilla(3, i + 2) = "N"
            
            formatoGrilla(4, i + 2) = "###,###,##0"
            
            formatoGrilla(5, i + 2) = "FALSE"
            
            formatoGrilla(6, i + 2) = ""
            
            formatoGrilla(7, i + 2) = ""
            
            formatoGrilla(8, i + 2) = ancho
        Next i
        
        formatoGrilla(1, i + 2) = "TOTAL"
            
        formatoGrilla(2, i + 2) = "10"
        
        formatoGrilla(3, i + 2) = "N"
        
        formatoGrilla(4, i + 2) = "###,###,##0"
        
        formatoGrilla(5, i + 2) = "FALSE"
        
        formatoGrilla(6, i + 2) = ""
        
        formatoGrilla(7, i + 2) = ""
        
        If cantMeses > 2 Then
            formatoGrilla(8, i + 2) = "8"
        Else
            formatoGrilla(8, 1) = ancho
            formatoGrilla(8, 2) = ancho
            formatoGrilla(8, i + 2) = ancho
        End If
        
        impresion.Cols = i + 3
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            Else
                impresion.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        impresion.AutoRedraw = False
        
        impresion.PageSetup.HeaderMargin = 2
    
        impresion.PageSetup.TopMargin = 2
        impresion.PageSetup.LeftMargin = 2
        impresion.PageSetup.RightMargin = 1.5
        impresion.PageSetup.BottomMargin = 2
        
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellLandscape
        impresion.PageSetup.PrintFixedRow = True
        
        Call verificaImpresora(5, impresion)
        
        impresion.AutoRedraw = True
    End Sub
    
    Private Sub leerMeses()
        Dim i As Integer
        Dim fecha As String
        For i = 1 To 12
            fecha = Format(fechasistema, "yyyy") & "-" & i & "-01"
            meses(i) = UCase(Format(fecha, "mmmm"))
            If Val(Format(fechasistema, "mm")) = i Then
                cantMeses = i
            End If
        Next i
    End Sub

    Private Sub impresion_Click()
        Dim i As Integer
        Dim ancho As Single
        Dim alto As Single
        Static fila As Long
        
        Call GetCursorPos(mouse)
        posX = mouse.X
        posY = mouse.Y
        
        If fila = impresion.MouseRow Then
            Unload ComprasClienteGrafico
            fila1 = 0
            fila2 = 0
            fila = 0
        Else
            fila = impresion.MouseRow
            If fila = -1 Then
                fila = 0
            Else
                If opt1.Value = True Then
                    If (fila Mod 2) = 0 Then
                        fila1 = fila - 1
                        fila2 = fila
                    Else
                        fila1 = fila
                        fila2 = fila + 1
                    End If
                Else
                    If (fila Mod 3) <> 0 Then
                        If (fila / 3) - (fila \ 3) < 0.5 Then
                            fila1 = fila
                            fila2 = fila + 1
                        Else
                            fila1 = fila - 1
                            fila2 = fila
                        End If
                    Else
                        fila1 = 0
                        fila2 = 0
                    End If
                End If
                If fila1 <> 0 Then
                    nombreCliente = impresion.Cell(fila2, 1).text
                    Unload ComprasClienteGrafico
                    Load ComprasClienteGrafico
                    With ComprasClienteGrafico
                        ancho = .Width
                        alto = .Height
                        .Left = posX * Screen.TwipsPerPixelX
                        .Top = posY * Screen.TwipsPerPixelY
                        If ancho + .Left > Screen.Width Then
                            .Left = .Left - ancho
                        End If
                        If alto + .Top > Screen.Height - 500 Then
                            .Top = .Top - alto
                        End If
                        .Show
                    End With
                End If
            End If
        End If
    End Sub

    








