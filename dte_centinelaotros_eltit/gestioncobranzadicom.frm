VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado12 
   Caption         =   "PANTALLA GESTION COBRANZA"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   0
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1095
      Left            =   90
      TabIndex        =   8
      Top             =   45
      Width           =   15885
      _ExtentX        =   28019
      _ExtentY        =   1931
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "GENERA INFORME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11385
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   540
         Width           =   2085
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   960
         Left            =   5985
         TabIndex        =   12
         Top             =   45
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   1693
         BackColor       =   12648384
         Caption         =   "TIPOS CLIENTES"
         CaptionEstilo3D =   1
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox Combotipos 
            Height          =   315
            Left            =   45
            TabIndex        =   13
            Text            =   "Combo1"
            Top             =   315
            Width           =   4875
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1005
         Left            =   405
         TabIndex        =   14
         Top             =   0
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
            TabIndex        =   4
            Tag             =   "proveedor"
            Top             =   645
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
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "proveedor"
            Top             =   645
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
            Left            =   2430
            MaxLength       =   4
            TabIndex        =   5
            Tag             =   "proveedor"
            Top             =   645
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
            Left            =   2430
            MaxLength       =   4
            TabIndex        =   2
            Tag             =   "proveedor"
            Top             =   270
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
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   1
            Tag             =   "proveedor"
            Top             =   270
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
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "proveedor"
            Top             =   270
            Width           =   435
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
            TabIndex        =   16
            Top             =   270
            Width           =   1335
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
            TabIndex        =   15
            Top             =   630
            Width           =   1335
         End
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8385
      Left            =   45
      TabIndex        =   7
      Top             =   1215
      Width           =   15900
      _ExtentX        =   28046
      _ExtentY        =   14790
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
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FF8080&
         Caption         =   "Desmarcar Todas"
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
         Left            =   12960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7680
         Width           =   2040
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "Marcar Todas"
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
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   7680
         Width           =   1440
      End
      Begin VB.CheckBox cobrarcartas 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cobrar Dicom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9480
         TabIndex        =   19
         Top             =   7800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin FlexCell.Grid Grid2 
         Height          =   195
         Left            =   315
         TabIndex        =   18
         Top             =   7740
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   344
         BackColorBkg    =   -2147483644
         BackColorFixed  =   16777215
         BackColorFixedSel=   -2147483641
         BackColorScrollBar=   -2147483641
         BackColorSel    =   16777215
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   9.75
         Rows            =   30
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF8080&
         Caption         =   "Generar Dicom"
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7680
         Width           =   2760
      End
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   240
         Left            =   45
         TabIndex        =   11
         Top             =   7380
         Width           =   15840
         _ExtentX        =   27940
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "Imprimir"
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
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7680
         Width           =   2760
      End
      Begin FlexCell.Grid GRID1 
         Height          =   7080
         Left            =   45
         TabIndex        =   9
         Top             =   270
         Width           =   15810
         _ExtentX        =   27887
         _ExtentY        =   12488
         AllowUserSort   =   -1  'True
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "tmplistado12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FOLIOINTERNO As Double



Private Sub Command1_Click()

Call Titulos("LISTADO DE CLIENTES CON DEUDA")
Grid1.PageSetup.Orientation = cellLandscape


Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0.5
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.FooterMargin = 2
Grid1.PageSetup.BlackAndWhite = True

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick



Grid1.PrintPreview
End Sub

Private Sub Command2_Click()
LEErclientes


End Sub

Private Sub Command3_Click()
Dim rut As String
Dim total As String
Dim nro As String
Dim i As Integer
Dim dire As String

Call CargaGrillacobranza(1, 13)
FOLIOINTERNO = 0
For K = 1 To Grid1.Rows - 1
    
    If Grid1.Cell(K, 9).text = "1" Then
        FOLIOINTERNO = FOLIOINTERNO + 1
        pivote.text = Grid1.Cell(K, 1).text
        pivote.MaxLength = 10
        rut = ceros(pivote)
        pivote.text = CDbl(Grid1.Cell(K, 6).text) & "00"
        pivote.MaxLength = 13
        total = ceros(pivote)
        nro = ""
        dire = Grid1.Cell(K, 3).text
        For s = 1 To Len(dire)
        If Mid(dire, s, 1) >= "0" And Mid(dire, s, 1) <= "9" Then
        nro = nro + Mid(dire, s, 1)
        End If
        Next s
        If nro = "" Then
           nro = "S/N"
        End If
        
        Call LEErCREDITOS("659490", leernombreempresacontable("28"), rut, Grid1.Cell(K, 2).text, Grid1.Cell(K, 3).text, nro, Grid1.Cell(K, 4).text, Format(Grid1.Cell(K, 8).text, "yyyymmdd"), "TC", "MOROSIDAD", "$", total)
        Call grabareventocarta(Grid1.Cell(K, 1).text, Format(fechasistema, "yyyy-mm-dd"), Time, Format(Grid1.Cell(K, 8).text, "yyyy-mm-dd"), "02", Grid1.Cell(K, 6).text, "GASTO COBRANZA 3")
 
    End If

Next K
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid2.Refresh
Call Grid2.ExportToExcel("", True)

End Sub
Sub grabareventocarta(rut, fecha, HORA, fechamorosidad, evento, MONTO, GLOSA)

Dim CAMPOS(10, 10) As String
Dim op As Integer

 

    CAMPOS(0, 0) = "rut"
    CAMPOS(1, 0) = "fecha"
    CAMPOS(2, 0) = "hora"
    CAMPOS(3, 0) = "fechamorosidad"
    CAMPOS(4, 0) = "evento"
    CAMPOS(5, 0) = "monto"
    CAMPOS(6, 0) = "glosa"
    CAMPOS(7, 0) = ""
    CAMPOS(0, 1) = rut
    CAMPOS(1, 1) = fecha
    CAMPOS(2, 1) = HORA
    CAMPOS(3, 1) = fechamorosidad
    CAMPOS(4, 1) = evento
    CAMPOS(5, 1) = CDbl(MONTO)
    CAMPOS(6, 1) = GLOSA
    CAMPOS(0, 2) = "sv_cobranza_gestion"
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    op = 2
    condicion = ""
    Call sqlventas.sqlventas(op, condicion)
    Call grabarcuotacobranza(rut, fecha, evento, GLOSA)
 
    
End Sub


Private Sub Command4_Click()
Dim K As Double
For K = 1 To Grid1.Rows - 2
    If Format(Grid1.Cell(K, 8).text, "yyyy-mm-dd") > Format(Grid1.Cell(K, 10).text, "yyyy-mm-dd") Then
    Grid1.Cell(K, 9).text = "1"
    End If
Next K

End Sub

Private Sub Command5_Click()
Dim K As Double
For K = 1 To Grid1.Rows - 2
    Grid1.Cell(K, 9).text = "0"
Next K
End Sub

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
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
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

Private Sub Form_Load()
Call CargaGrillaGRID1(1, 11)

LEErTIPOSCLIENTES

dato1.text = "01"
dato2.text = "01"
dato3.text = "1995"
dato4.text = Format(fechasistema, "dd")
dato5.text = Format(fechasistema, "mm")
dato6.text = Format(fechasistema, "yyyy")

End Sub

 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "RUT"
        formatogrilla(1, 2) = "CLIENTE"
        formatogrilla(1, 3) = "DIRECCION"
        formatogrilla(1, 4) = "COMUNA"
        formatogrilla(1, 5) = "FONO"
        formatogrilla(1, 6) = "MORA"
        formatogrilla(1, 7) = "DIASMORA"
        formatogrilla(1, 8) = "FECHA MORA"
        formatogrilla(1, 9) = "DICOM"
        formatogrilla(1, 10) = "ULTIMO DICOM"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "D"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = "0000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = "$ ###,###,##0"
        formatogrilla(4, 7) = ""

        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"

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
        
        Rem ANCHO
        formatogrilla(8, 1) = "8"
        formatogrilla(8, 2) = "25"
        formatogrilla(8, 3) = "25"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "5"
        formatogrilla(8, 10) = "8"
            
        Grid1.Cols = col
        Grid1.Rows = row
        Grid1.AllowUserResizing = True
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
            Grid1.Column(i).Width = Val(formatogrilla(8, i)) * (Grid1.Cell(0, i).Font.Size + 1.25)
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
    Grid1.Column(9).CellType = cellCheckBox
    
    End Sub
'**
Sub LEErclientes()

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
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.fono1,sum(case when cd.vencimientoactual<'" + Format(fechasistema, "yyyy-mm-dd") + "' then cd.montocuota-cd.abono else '0' end)  "
        csql.sql = csql.sql + "FROM sv_maestroclientes as mc inner join sv_cuotas_detalle as cd on (cd.rut=mc.rut) "
        If Mid(Combotipos.text, 1, 2) <> "99" Then
            csql.sql = csql.sql + "and mc.tipocliente='" + Mid(Combotipos.text, 1, 2) + "' "
        End If
        csql.sql = csql.sql + "group by cd.rut order by mc.nombre "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            barra.Max = csql.RowsAffected + 1
            barra.Value = 0
            Grid1.Rows = 1
            Grid1.AutoRedraw = False
        
            total1 = 0
            total2 = 0
            total3 = 0
            total4 = 0
        
            While Not resultado.EOF
                LEErdiasmora (resultado(0))
                If Format(FECHAMORA, "yyyy-mm-dd") >= dato3.text + "-" + dato2.text + "-" + dato1.text And Format(FECHAMORA, "yyyy-mm-dd") <= dato6.text + "-" + dato5.text + "-" + dato4.text Then
                    mora = resultado(5)
                    If mora <> 0 Then
                        barra.Value = barra.Value + 1
                        barra.Refresh
                        Grid1.Rows = Grid1.Rows + 1
                        Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(0)
                        Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(1)
                        Grid1.Cell(Grid1.Rows - 1, 3).text = resultado(2)
                        Grid1.Cell(Grid1.Rows - 1, 4).text = resultado(3)
                        Grid1.Cell(Grid1.Rows - 1, 5).text = resultado(4)
                        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(mora, "###,###,###")
                        total1 = total1 + CREDITO
                        total2 = total2 + usado
                        total3 = total3 + disponible
                        total4 = total4 + mora
                        Grid1.Cell(Grid1.Rows - 1, 10).text = LeerUltimaCartaCobranza(resultado(0), "02")
                        If mora <> 0 Then
                            Grid1.Cell(Grid1.Rows - 1, 7).text = LEErdiasmora(resultado(0))
                            Grid1.Cell(Grid1.Rows - 1, 8).text = FECHAMORA
                            If Format(FECHAMORA, "yyyy-mm-dd") > Format(Grid1.Cell(Grid1.Rows - 1, 10).text, "yyyy-mm-dd") Then
                                Grid1.Cell(Grid1.Rows - 1, 9).text = "1"
                            End If
                        End If

                    End If
                End If
                resultado.MoveNext
            Wend
        Else
       
        End If
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 6).Borders(cellEdgeTop) = cellThick
        Grid1.Cell(Grid1.Rows - 1, 2).text = "TOTALES GENERALES"
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(total1, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(total2, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(total3, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(total4, "###,###,###")
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
 
    End Sub

Private Sub Grid1_Click()
If Grid1.Rows > 1 Then
    If Format(Grid1.Cell(Grid1.ActiveCell.row, 8).text, "yyyy-mm-dd") > Format(Grid1.Cell(Grid1.ActiveCell.row, 10).text, "yyyy-mm-dd") Then
        Grid1.Cell(Grid1.ActiveCell.row, 9).text = 1
    Else
        Grid1.Cell(Grid1.ActiveCell.row, 9).text = 0
        MsgBox "YA A INGRESADO A DICOM EN  ESTA FECHA", vbInformation, "ATECION"
    End If
End If
End Sub

Private Sub Grid1_DblClick()
'creditoPAGOSTMP.rut2.text = Mid(GRID1.Cell(GRID1.ActiveCell.row, 1).text, 1, 9)
'creditoPAGOSTMP.lbldv.Caption = Mid(GRID1.Cell(GRID1.ActiveCell.row, 1).text, 10, 1)
'creditoPAGOSTMP.Show
'
End Sub

Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    
      
    Grid1.PageSetup.PrintTitleRows = 1
    
 
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid1.PageSetup.HeaderAlignment = cellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE COBRANZA MOROSOS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "AL DIA  :  " & Format(fechasistema, "dd-mm-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub


Sub LEErTIPOSCLIENTES()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    
        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM sv_tiposdeclientes "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        linea = 1
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                linea = linea + 1
                Combotipos.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
      Combotipos.AddItem ("99 TODOS")
      
      
                
        Combotipos.text = Combotipos.List(linea - 1)
        End If
        
End Sub

 Private Sub CargaGrillacobranza(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "EMPRESA"
        formatogrilla(1, 3) = "RUT"
        formatogrilla(1, 4) = "NOMBRE"
        formatogrilla(1, 5) = "DIRECCION"
        formatogrilla(1, 6) = "NRO"
        formatogrilla(1, 7) = "CIUDAD"
        formatogrilla(1, 8) = "VENCIMT"
        formatogrilla(1, 9) = "DOC"
        formatogrilla(1, 10) = "TIPO DEUDA"
        formatogrilla(1, 11) = "M"
        formatogrilla(1, 12) = "D E U D A"
        
        Rem ANCHO
        formatogrilla(8, 1) = "6"
        formatogrilla(8, 2) = "40"
        formatogrilla(8, 3) = "8"
        formatogrilla(8, 4) = "40"
        formatogrilla(8, 5) = "40"
        formatogrilla(8, 6) = "5"
        formatogrilla(8, 7) = "20"
        formatogrilla(8, 8) = "7"
        formatogrilla(8, 9) = "4"
        formatogrilla(8, 10) = "9"
        formatogrilla(8, 11) = "2"
        formatogrilla(8, 12) = "10"
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 8) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "S"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = ""
        formatogrilla(3, 12) = "N"
        
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
        
            
        Grid2.Cols = col
        Grid2.Rows = row
        Grid2.AllowUserResizing = False
        Grid2.DisplayFocusRect = False
        Grid2.ExtendLastCol = True
        Grid2.BoldFixedCell = False
        Grid2.DrawMode = cellOwnerDraw
        Grid2.Appearance = Flat
        Grid2.ScrollBarStyle = Flat
        Grid2.FixedRowColStyle = Flat
''        Grid2.BackColorFixed = RGB(90, 158, 214)
''        Grid2.BackColorFixedSel = RGB(110, 180, 230)
'        Grid2.BackColorBkg = RGB(90, 158, 214)
'        Grid2.BackColorScrollBar = RGB(231, 235, 247)
'        Grid2.BackColor1 = RGB(231, 235, 247)
'        Grid2.BackColor2 = RGB(239, 243, 255)
'        Grid2.GridColor = RGB(148, 190, 231)
''        Grid2.DefaultFont.Bold = True
        
        Grid2.Column(0).Width = 0
        For i = 1 To col - 1
            Grid2.Cell(0, i).text = formatogrilla(1, i)
            Grid2.Column(i).Width = Val(formatogrilla(8, i)) * (Grid2.Cell(0, i).Font.Size)
            Grid2.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid2.Column(i).FormatString = formatogrilla(4, i)
            Grid2.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid2.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                Grid2.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                Grid2.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Grid2.Range(0, 0, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
        Grid2.Enabled = True
    End Sub

Sub LEErCREDITOS(CODIGO, EMPRESA, rut, nombre, direccion, nro, ciudad, VENCIMT, DOC, TIPODEUDA, m, deuda)
        
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = CODIGO
        Grid2.Cell(Grid2.Rows - 1, 2).text = EMPRESA
        Grid2.Cell(Grid2.Rows - 1, 3).text = rut
        Grid2.Cell(Grid2.Rows - 1, 4).text = nombre
        Grid2.Cell(Grid2.Rows - 1, 5).text = direccion
        Grid2.Cell(Grid2.Rows - 1, 6).text = nro
        Grid2.Cell(Grid2.Rows - 1, 7).text = ciudad
        Grid2.Cell(Grid2.Rows - 1, 8).text = VENCIMT
        Grid2.Cell(Grid2.Rows - 1, 9).text = DOC
        Grid2.Cell(Grid2.Rows - 1, 10).text = TIPODEUDA
        Grid2.Cell(Grid2.Rows - 1, 11).text = m
        Grid2.Cell(Grid2.Rows - 1, 12).text = deuda
  

 
    End Sub

Sub Titulos2(titulo1, nombre, direccion, comuna, fono, FOLIO)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    Dim K As Integer
    
    Grid2.FixedRowColStyle = Fixed3D
    Grid2.CellBorderColorFixed = vbButtonShadow
    Grid2.ShowResizeTips = False
    Grid2.ReportTitles.Clear
    Grid2.PageSetup.CenterHorizontally = True
    Grid2.PageSetup.Orientation = cellPortrait

    
      
    Grid2.PageSetup.PrintTitleRows = 0
    
    'Logo
    
'    grid2.Images.Add App.Path & "\logo.jpg", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = cellLeft
'    objReportTitle.Height = 60
'    grid2.ReportTitles.Add objReportTitle
    
    
''    'ENCABEZADO DE PAGINA
'    grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & leerRutEmpresa(empresaActiva)
'    grid2.PageSetup.HeaderAlignment = cellLeft
'    grid2.PageSetup.HeaderFont.Name = "Verdana"
'    grid2.PageSetup.HeaderFont.Size = 8

  'ENCABEZADO DE PAGINA
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "EMPRESAS CREDITOS ELTIT " & vbCrLf & "PROMOTORA PALGUIN FRESIA 289 Sdo Piso " & vbCrLf & "FONO :441349 ANEXO 350"
'    grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & leerRutEmpresa(empresaActiva)
    objReportTitle.Font.Name = "Verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Align = cellLeft
    Grid2.ReportTitles.Add objReportTitle
'    grid2.PageSetup.HeaderAlignment = cellLeft
'    grid2.PageSetup.HeaderFont.Name = "Verdana"
'    grid2.PageSetup.HeaderFont.Size = 8
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
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "SEÑOR(A)." & vbCrLf & nombre & vbCrLf & direccion & vbCrLf & comuna & vbCrLf & comuna
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "PRESENTE"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Italic = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = False
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "            Estima Cliente , por medio de la presente informamos a ud. que en nuestros registros de cobranza " & vbCrLf & "aun se encuentran la siguientes cuotas impagas "
    

    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
   
    'PIE DE PAGINA
  Rem   Grid2.PageSetup.Footer = " ESTIMADOS SEÑORES " & vbCrLf & "PRUEBA"
    
    
    Grid2.PageSetup.Footer = "FOLIO INTERNO :" + Format(FOLIO, "00000")
    
    
    Grid2.PageSetup.FooterAlignment = cellCenter
    Grid2.PageSetup.FooterFont.Name = "Verdana"
    Grid2.PageSetup.FooterFont.Size = 7
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).FontBold = True
       
End Sub

