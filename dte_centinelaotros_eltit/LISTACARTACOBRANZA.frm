VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado8 
   Caption         =   "LISTA CARTAS DE COBRANZA"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1320
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   2328
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1005
         Left            =   9405
         TabIndex        =   18
         Top             =   270
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1773
         BackColor       =   16761024
         Caption         =   "Fechas"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   1
            Tag             =   "proveedor"
            Top             =   270
            Width           =   435
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
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   2
            Tag             =   "proveedor"
            Top             =   270
            Width           =   435
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
            Left            =   2430
            MaxLength       =   4
            TabIndex        =   3
            Tag             =   "proveedor"
            Top             =   270
            Width           =   795
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
            Left            =   2430
            MaxLength       =   4
            TabIndex        =   6
            Tag             =   "proveedor"
            Top             =   645
            Width           =   795
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
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   4
            Tag             =   "proveedor"
            Top             =   645
            Width           =   435
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
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   5
            Tag             =   "proveedor"
            Top             =   645
            Width           =   435
         End
         Begin VB.Label lbl2 
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
            Left            =   90
            TabIndex        =   20
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lbl1 
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
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF8080&
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13545
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox SUCU 
         Height          =   285
         Left            =   7695
         MaxLength       =   1
         TabIndex        =   13
         Text            =   "0"
         Top             =   1260
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13005
         TabIndex        =   11
         Top             =   3015
         Width           =   1230
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   660
         Left            =   90
         TabIndex        =   14
         Top             =   315
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "CLIENTES X RUT"
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
         Begin VB.TextBox rut1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   90
            MaxLength       =   9
            TabIndex        =   0
            Tag             =   "proveedor"
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label lblDV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1530
            TabIndex        =   16
            Top             =   270
            Width           =   375
         End
         Begin VB.Label lblnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   2025
            TabIndex        =   15
            Top             =   270
            Width           =   5085
         End
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8340
      Left            =   45
      TabIndex        =   7
      Top             =   1350
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   14711
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   7335
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5895
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7830
         Width           =   2760
      End
      Begin FlexCell.Grid GRID1 
         Height          =   6990
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   12330
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "tmplistado8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub COMBOMES_Click()
LEErCREDITOS
End Sub

Private Sub Command1_Click()

Call Titulos("CARTA COBRANZA")

Grid1.PageSetup.HeaderMargin = 0
Grid1.PageSetup.PrintFixedRow = True

Grid1.PageSetup.TopMargin = 0.5
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0.5
Grid1.PageSetup.BottomMargin = 3
Grid1.PageSetup.FooterMargin = 2
Grid1.PageSetup.BlackAndWhite = True
Grid1.PrintPreview
End Sub

Private Sub Command2_Click()
Call CargaGrillaGRID1(1, 11)
LEErCREDITOS

End Sub



Private Sub Command3_Click()
LEErCREDITOS
End Sub

 Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, DATO1)
        Call selecciona(DATO1)
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

    Private Sub DATO5_GotFocus()
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
    
    Private Sub DATO6_GotFocus()
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
        Call Flechas(KeyCode, DATO1)
    End Sub

    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, DATO1)
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
            DATO1.text = ceros(DATO1)
            If DATO1.text = "00" Then
                DATO1.text = Format(fechasistema, "dd")
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
            fecha1 = dato3.text & "-" & dato2.text & "-" & DATO1.text
            fecha2 = ""
            
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
    
    Private Sub DATO5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub DATO6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & DATO1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            SendKeys "{Tab}"
            
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================

    '========================================================
    'KeyUp
    '========================================================
    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato1.text) = dato1.MaxLength Then
        '    Call dato1_KeyPress(13)
        'End If
    End Sub
    
    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato2.text) = dato2.MaxLength Then
        '    Call dato2_KeyPress(13)
        'End If
    End Sub
    
    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato3.text) = dato3.MaxLength Then
        '    Call dato3_KeyPress(13)
        'End If
    End Sub
    Private Sub dato1_LostFocus()
    Call esfecha(DATO1, dato2, dato3, "dd")
    End Sub
    Private Sub dato2_LostFocus()
    Call esfecha(DATO1, dato2, dato3, "mm")
    End Sub
    Private Sub dato3_LostFocus()
    Call esfecha(DATO1, dato2, dato3, "yyyy")
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
   

Private Sub Form_Activate()
rut1.SetFocus
End Sub

Private Sub Form_Load()
Call CargaGrillaGRID1(1, 11)


End Sub

 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "LO"
        formatogrilla(1, 2) = "F.COMPRA"
        formatogrilla(1, 3) = "TD"
        formatogrilla(1, 4) = "NUMERO"
        formatogrilla(1, 5) = "COMPRA"
        formatogrilla(1, 6) = "VENCIMIENTO"
        formatogrilla(1, 7) = "CUOTA/DE"
        formatogrilla(1, 8) = "M.CUOTA"
        formatogrilla(1, 9) = "INT.MORA"
        formatogrilla(1, 10) = "TOTAL"
        
        Rem ANCHO
        formatogrilla(8, 1) = "2"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "4"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "25"
        formatogrilla(8, 6) = "10"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
        formatogrilla(8, 10) = "8"
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "D"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "D"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""

        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"

        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        
            
        Grid1.Cols = col
        Grid1.Rows = row
        Grid1.AllowUserResizing = False
        Grid1.DisplayFocusRect = False
        Grid1.ExtendLastCol = True
        Grid1.BoldFixedCell = False
        Grid1.DrawMode = cellOwnerDraw
        Grid1.Appearance = Flat
        Grid1.ScrollBarStyle = Flat
        Grid1.FixedRowColStyle = Flat
        Grid1.BackColorFixed = RGB(90, 158, 214)
        Grid1.BackColorFixedSel = RGB(110, 180, 230)
        Grid1.BackColorBkg = RGB(90, 158, 214)
        Grid1.BackColorScrollBar = RGB(231, 235, 247)
        Grid1.BackColor1 = RGB(231, 235, 247)
        Grid1.BackColor2 = RGB(239, 243, 255)
        Grid1.GridColor = RGB(148, 190, 231)
        
        Grid1.Column(0).Width = 0
        For i = 1 To col - 1
            Grid1.Cell(0, i).text = formatogrilla(1, i)
            Grid1.Column(i).Width = Val(formatogrilla(8, i)) * (Grid1.Cell(0, i).Font.Size)
            Grid1.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid1.Column(i).FormatString = formatogrilla(4, i)
            Grid1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid1.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                Grid1.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                Grid1.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
        Grid1.Enabled = True
    End Sub
'**
Sub LEErCREDITOS()

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim CREDITO As Double
        Dim usado As Double
        Dim disponible As Double
        Dim mora As Double
        Dim total1 As Double
        Dim total2 As Double
        Dim total3 As Double
        Dim total4 As Double
        Dim total5 As Double
        Dim ACUMULADO As Double
        Dim FECHAMORA As String
        Dim MESMORA As String
        Dim AÑOMORA As String
        Dim linea As Double
        
        
        Dim fecha1 As String
        Dim fecha2 As String
        
        fecha1 = dato3.text + "-" + dato2.text + "-" + DATO1.text
        fecha2 = dato6.text + "-" + dato5.text + "-" + dato4.text
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT cd.local,cd.fechacompra,cd.tipo,cd.numero,cd.glosacompra,cd.vencimientoactual,cd.numerocuota,cd.cantidadcuotas,cd.montocuota-cd.abono "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle as cd "
        csql.sql = csql.sql & "WHERE cd.rut='" + rut1.text + lbldv.Caption + "' and  cd.vencimientoactual between '" + "1900-01-01" + "' and '" + fecha2 + "' and montocuota>abono  "
        
        
        csql.sql = csql.sql & "order by cd.vencimientoactual "
        
        
        
        csql.Execute
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
'        If Option1.Value = True Then separador = resultado(4)
'        If Option2.Value = True Then separador = resultado(6)
        
        
        barra.Max = csql.RowsAffected + 1
        
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        total1 = 0
        total2 = 0
        total3 = 0
        total4 = 0
        total5 = 0
        
        While Not resultado.EOF
        
        tazainteresmora = leerInteresMora("00")
        diasmora = DateDiff("d", resultado(5), fechasistema)
        If diasmora <= diasgracia Then diasmora = 0
        mora = Round(resultado(8) * ((tazainteresmora / 30 * diasmora) / 100), 0)
        
        
        ACUMULADO = ACUMULADO + (resultado(8) + mora)
        barra.Value = Grid1.Rows
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(0)
        If IsNull(resultado(1)) = False Then
        Grid1.Cell(Grid1.Rows - 1, 2).text = Format(resultado(1), "dd-mm-yyyy")
        End If
        
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultado(2)
        Grid1.Cell(Grid1.Rows - 1, 4).text = resultado(3)
        Grid1.Cell(Grid1.Rows - 1, 5).text = resultado(4)
        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(resultado(5), "dd-mm-yyyy")
        Grid1.Cell(Grid1.Rows - 1, 7).text = resultado(6) & "/" & resultado(7)
        Grid1.Cell(Grid1.Rows - 1, 8).text = Format(resultado(8), "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(mora, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 10).text = Format(resultado(8) + mora, "###,###,###")
     
        total1 = total1 + resultado(8)
        total2 = total2 + mora
        total3 = total3 + (resultado(8) + mora)
        
        total11 = total11 + resultado(8)
        total12 = total12 + mora
        total13 = total13 + (resultado(8) + mora)
        
        

        
    
            resultado.MoveNext
            Wend
        Else
       
        End If
         
        
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 6, Grid1.Rows - 1, 10).Borders(cellEdgeTop) = cellThick
          
        Grid1.Cell(Grid1.Rows - 1, 6).text = "TOTALES GENERALES"
        Grid1.Cell(Grid1.Rows - 1, 8).text = Format(total11, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(total12, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 10).text = Format(total13, "###,###,###")
       
    Grid1.Column(1).Locked = False
    Grid1.Column(2).Locked = False
    Grid1.Column(3).Locked = False
    Grid1.Column(4).Locked = False
    Grid1.Column(5).Locked = False
    Grid1.Column(6).Locked = False
    Grid1.Column(7).Locked = False
    Grid1.Column(8).Locked = False
    Grid1.Column(9).Locked = False
    Grid1.Column(10).Locked = False
     For K = 1 To 10
        linea = linea + 1
     Next K
        
    Grid1.Rows = Grid1.Rows + linea
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 8
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid1.Cell(Grid1.Rows - 1, 1).text = " "
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 9
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid1.Cell(Grid1.Rows - 1, 1).text = "        Se le solicita tenga a bien regularizar cuanto antes esta situación.  "
   
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 9
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid1.Cell(Grid1.Rows - 1, 1).text = " "
   
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 9
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid1.Cell(Grid1.Rows - 1, 1).text = "         Saluda a usted cordialmente,  "
   
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 8
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid1.Cell(Grid1.Rows - 1, 1).text = " "
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 8
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellCenterCenter
    Grid1.Cell(Grid1.Rows - 1, 1).text = "Departamento de Cobranzas  "
    
 
    
     Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 50
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid1.Cell(Grid1.Rows - 1, 1).text = " "
    
    Grid1.Rows = Grid1.Rows + 2
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).FontSize = 9
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).Alignment = cellLeftCenter
    Grid1.Cell(Grid1.Rows - 1, 1).text = "Nota: en caso de haber recibido esta cobranza con posterioridad a la regularización de esta deuda; "
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).FontSize = 9
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 9).Alignment = cellLeftCenter
    Grid1.Cell(Grid1.Rows - 1, 1).text = "por favor, sírvase dejar sin efecto esta comunicación. Estas cuotas se encuentran sin intereses."
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Merge
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontSize = 9
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).FontBold = True
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 8).Alignment = cellCenterCenter
    Grid1.Cell(Grid1.Rows - 1, 1).text = ""
        
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
    End Sub


Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    Dim K As Integer
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellPortrait

    
      
    Grid1.PageSetup.PrintTitleRows = 1
    
    'Logo
    
'    Grid1.Images.Add App.Path & "\logo.jpg", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = cellLeft
'    objReportTitle.Height = 60
'    Grid1.ReportTitles.Add objReportTitle
    
    
''    'ENCABEZADO DE PAGINA
'    GRID1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & leerRutEmpresa(empresaActiva)
'    GRID1.PageSetup.HeaderAlignment = cellLeft
'    GRID1.PageSetup.HeaderFont.Name = "Verdana"
'    GRID1.PageSetup.HeaderFont.Size = 8

  'ENCABEZADO DE PAGINA
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = direccionempresa & vbCrLf & comunaempresa
'    GRID1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & leerRutEmpresa(empresaActiva)
    objReportTitle.Font.Name = "Verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Align = cellLeft
    Grid1.ReportTitles.Add objReportTitle
'    GRID1.PageSetup.HeaderAlignment = cellLeft
'    GRID1.PageSetup.HeaderFont.Name = "Verdana"
'    GRID1.PageSetup.HeaderFont.Size = 8
''
    'TITULOS DEL REPORTE
    
   
'    If Option1.Value = True Then tipoListado = "CLIENTES MAAT"
'    If Option2.Value = True Then tipoListado = "CLIENTES SKORPIOS"
'    If Option3.Value = True Then tipoListado = "CLIENTES TODOS"
'
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = comunaempresa & "," & Format(fechasistema, ("dd")) & " de " & MonthName(Month(Now)) & " " & Format(fechasistema, ("yyyy")) & "   "
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold
    objReportTitle.Align = cellRight
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "SEÑOR(A)." & vbCrLf & lblNombre.Caption & vbCrLf & leerDireccionCliente(rut1.text & lbldv.Caption, "0") & vbCrLf & Replace(leerCiudadCliente(rut1.text & lbldv.Caption, "0"), "     ", "")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "PRESENTE"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Italic = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "            Comunicamos a ustedes que al día de hoy en nuestros registros aun mantiene " & vbCrLf & "  una deuda, que detalla a continuación: " & vbCrLf

    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
   
    'PIE DE PAGINA
'    Grid1.PageSetup.Footer = "info@importadoraskorpios.cl   Manuel Montt 533,  Coronel Centro  Fono (41) 277 3347 Fax (41) 271 1942"
    Grid1.PageSetup.FooterAlignment = cellCenter
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
     Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
     Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
     Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
     Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
     Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
     Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
     Grid1.Range(0, 1, 0, Grid1.Cols - 1).FontBold = True
       
End Sub



Private Sub rut1_GotFocus()
   
        Call selecciona(rut1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
End Sub

Private Sub rut1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
            Call ayudaCliente(rut1, SUCU, lbldv)
        Else
            Call Flechas(KeyCode, rut1)
        End If
End Sub

Private Sub rut1_KeyPress(KeyAscii As Integer)

 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And rut1.text <> "" And Val(rut1.text) <> 0 Then
            rut1.text = ceros(rut1)
            lbldv.Caption = rut(rut1.text)
            rut_cliente = rut1.text + lbldv.Caption
            lblNombre.Caption = leerNombreCliente(rut_cliente)
            LEErCREDITOS
            
           
        End If
End Sub

