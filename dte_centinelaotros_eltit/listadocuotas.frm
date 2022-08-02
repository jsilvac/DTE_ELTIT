VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado6 
   Caption         =   "LISTADO CUOTAS PENDIENTES"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1005
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   1773
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
         Height          =   960
         Left            =   5310
         TabIndex        =   14
         Top             =   0
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1693
         BackColor       =   12648384
         Caption         =   "ORDENADOS POR"
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
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "VENCIMIENTO"
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
            Left            =   135
            TabIndex        =   16
            Top             =   630
            Width           =   2760
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "CLIENTE"
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
            Left            =   135
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
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
         Left            =   3645
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   585
         Width           =   1545
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
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   9
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
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   8
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
         Left            =   2715
         MaxLength       =   4
         TabIndex        =   7
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
         Left            =   2715
         MaxLength       =   4
         TabIndex        =   6
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
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   5
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
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   270
         Width           =   435
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   960
         Left            =   7515
         TabIndex        =   17
         Top             =   0
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1693
         BackColor       =   12648384
         Caption         =   "LISTADO"
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
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ACUMULADO"
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
            Left            =   135
            TabIndex        =   19
            Top             =   270
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "DETALLADO"
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
            Left            =   135
            TabIndex        =   18
            Top             =   630
            Width           =   2760
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   960
         Left            =   9765
         TabIndex        =   20
         Top             =   0
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
            TabIndex        =   21
            Text            =   "Combo1"
            Top             =   315
            Width           =   4875
         End
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
         Left            =   315
         TabIndex        =   11
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
         Left            =   315
         TabIndex        =   10
         Top             =   630
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8475
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   14949
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
         Left            =   135
         TabIndex        =   13
         Top             =   7425
         Width           =   14595
         _ExtentX        =   25744
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
         TabIndex        =   3
         Top             =   7965
         Width           =   2760
      End
      Begin FlexCell.Grid GRID1 
         Height          =   6945
         Left            =   45
         TabIndex        =   2
         Top             =   315
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   12250
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "tmplistado6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Call Titulos("LISTADO DE CUOTAS POR VENCER ")
GRID1.PageSetup.Orientation = cellLandscape
GRID1.PageSetup.HeaderMargin = 0.5
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.TopMargin = 2
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0.5
GRID1.PageSetup.BottomMargin = 3
GRID1.PageSetup.FooterMargin = 2
GRID1.PageSetup.BlackAndWhite = True


GRID1.PrintPreview
End Sub

Private Sub Command2_Click()
Call CargaGrillaGRID1(1, 13)
LEErCREDITOS

End Sub



Private Sub Form_Activate()
dato1.SetFocus
End Sub

Private Sub Form_Load()
Call CargaGrillaGRID1(1, 13)
dato1.text = Format(fechasistema, "dd")
dato2.text = Format(fechasistema, "mm")
dato3.text = Format(fechasistema, "yyyy")
dato4.text = Format(fechasistema, "dd")
dato5.text = Format(fechasistema, "mm")
dato6.text = Format(fechasistema, "yyyy")
LEErTIPOSCLIENTES





End Sub

 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "LO"
        formatogrilla(1, 2) = "F.COMPRA"
        formatogrilla(1, 3) = "TD"
        formatogrilla(1, 4) = "NUMERO"
        formatogrilla(1, 5) = "RUT"
        formatogrilla(1, 6) = "CLIENTE"
        formatogrilla(1, 7) = "VENCIMIENTO"
        formatogrilla(1, 8) = "N.CUOTA"
        formatogrilla(1, 9) = "M.CUOTA"
        formatogrilla(1, 10) = "INT.MORA"
        formatogrilla(1, 11) = "TOTAL"
        formatogrilla(1, 12) = "ACUMULADO"
        
        
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
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "S"
        formatogrilla(3, 7) = "D"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = "N"
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
        
        Rem ANCHO
        formatogrilla(8, 1) = "2"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "4"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "25"
        formatogrilla(8, 7) = "10"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
        formatogrilla(8, 10) = "8"
        formatogrilla(8, 11) = "8"
        formatogrilla(8, 12) = "8"
            
        GRID1.Cols = col
        GRID1.Rows = row
        GRID1.AllowUserResizing = False
        GRID1.DisplayFocusRect = False
        GRID1.ExtendLastCol = True
        GRID1.BoldFixedCell = False
        GRID1.DrawMode = cellOwnerDraw
        GRID1.Appearance = Flat
        GRID1.ScrollBarStyle = Flat
        GRID1.FixedRowColStyle = Flat
        GRID1.BackColorFixed = RGB(90, 158, 214)
        GRID1.BackColorFixedSel = RGB(110, 180, 230)
        GRID1.BackColorBkg = RGB(90, 158, 214)
        GRID1.BackColorScrollBar = RGB(231, 235, 247)
        GRID1.BackColor1 = RGB(231, 235, 247)
        GRID1.BackColor2 = RGB(239, 243, 255)
        GRID1.GridColor = RGB(148, 190, 231)
        
        GRID1.Column(0).Width = 0
        For i = 1 To col - 1
            GRID1.Cell(0, i).text = formatogrilla(1, i)
            GRID1.Column(i).Width = Val(formatogrilla(8, i)) * (GRID1.Cell(0, i).Font.Size + 1.25)
            GRID1.Column(i).MaxLength = Val(formatogrilla(2, i))
            GRID1.Column(i).FormatString = formatogrilla(4, i)
            GRID1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                GRID1.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                GRID1.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                GRID1.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        GRID1.Range(0, 0, 0, GRID1.Cols - 1).Alignment = cellCenterCenter
        GRID1.Enabled = True
    End Sub
'**
Sub LEErCREDITOS()

        Dim cSql As rdoQuery
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
        
        
        Dim fecha1 As String
        Dim fecha2 As String
        fecha1 = dato3.text + "-" + dato2.text + "-" + dato1.text
        fecha2 = dato6.text + "-" + dato5.text + "-" + dato4.text
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT cd.local,cd.fechacompra,cd.tipo,cd.numero,cd.rut,mc.nombre,cd.vencimientoactual,cd.numerocuota,cd.montocuota-cd.abono "
        cSql.sql = cSql.sql & "FROM sv_maestroclientes as mc,sv_cuotas_detalle as cd "
        cSql.sql = cSql.sql & "WHERE cd.rut=mc.rut and cd.vencimientoactual between '" + fecha1 + "' and '" + fecha2 + "' and montocuota>abono  "
        If Mid(Combotipos.text, 1, 2) <> "99" Then
        cSql.sql = cSql.sql + "and mc.tipocliente='" + Mid(Combotipos.text, 1, 2) + "' "
        End If
        
        
        
        If Option2.Value = True Then
        cSql.sql = cSql.sql & "order by cd.vencimientoactual,mc.nombre "
        End If
        If Option1.Value = True Then
        cSql.sql = cSql.sql & "order by mc.nombre,cd.vencimientoactual "
        End If
        
        
        
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
        If Option1.Value = True Then separador = resultado(4)
        If Option2.Value = True Then separador = resultado(6)
        
        
        BARRA.Max = cSql.RowsAffected + 1
        
        GRID1.Rows = 1
        GRID1.AutoRedraw = False
        
        total1 = 0
        total2 = 0
        total3 = 0
        total4 = 0
        total5 = 0
        
        While Not resultado.EOF
        If Option2.Value = True And separador <> resultado(6) Then
        GRID1.Rows = GRID1.Rows + 1
        BARRA.Max = BARRA.Max + 1
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 11).Borders(cellEdgeTop) = cellThick
        GRID1.Cell(GRID1.Rows - 1, 6).text = "TOTAL VENCIMIENTO"
        GRID1.Cell(GRID1.Rows - 1, 7).text = Format(separador, "dd-mm-yyyy")
        
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total1, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total2, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total3, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 12).text = Format(ACUMULADO, "###,###,###")
        
        total1 = 0
        total2 = 0
        total3 = 0
        separador = resultado(6)
        
        End If
        
        If Option1.Value = True And separador <> resultado(4) Then
        BARRA.Max = BARRA.Max + 1
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 12).Borders(cellEdgeTop) = cellThick
        GRID1.Cell(GRID1.Rows - 1, 5).text = "TOTAL "
        GRID1.Cell(GRID1.Rows - 1, 6).text = leerNombreCliente(separador)
        
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total1, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total2, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total3, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 12).text = Format(ACUMULADO, "###,###,###")
        
        total1 = 0
        total2 = 0
        total3 = 0
        separador = resultado(4)
        
        End If
        
        
        tazainteresmora = leerInteresMora("00")
        diasmora = DateDiff("d", resultado(6), fechasistema)
        If diasmora <= diasgracia Then diasmora = 0
        mora = Round(resultado(8) * ((tazainteresmora / 30 * diasmora) / 100), 0)
        ACUMULADO = ACUMULADO + (resultado(8) + mora)
        If Option3.Value = True Then
        BARRA.Value = GRID1.Rows
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Cell(GRID1.Rows - 1, 1).text = resultado(0)
        If IsNull(resultado(1)) = False Then
        GRID1.Cell(GRID1.Rows - 1, 2).text = Format(resultado(1), "dd-mm-yyyy")
        End If
        
        GRID1.Cell(GRID1.Rows - 1, 3).text = resultado(2)
        GRID1.Cell(GRID1.Rows - 1, 4).text = resultado(3)
        GRID1.Cell(GRID1.Rows - 1, 5).text = Mid(resultado(4), 1, 9) + "-" + Mid(resultado(4), 10, 1)
        GRID1.Cell(GRID1.Rows - 1, 6).text = resultado(5)
        GRID1.Cell(GRID1.Rows - 1, 7).text = Format(resultado(6), "dd-mm-yyyy")
        GRID1.Cell(GRID1.Rows - 1, 8).text = Format(resultado(7), "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(resultado(8), "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(mora, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(resultado(8) + mora, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 12).text = Format(ACUMULADO, "###,###,###")
        
        End If
        total1 = total1 + resultado(8)
        total2 = total2 + mora
        total3 = total3 + (resultado(8) + mora)
        
        total11 = total11 + resultado(8)
        total12 = total12 + mora
        total13 = total13 + (resultado(8) + mora)
        
        

        
'
'        saldo = resultado(7) - resultado(8)
'        tazainteresmora = leerInteresMora("00")
'        diasmora = DateDiff("d", resultado(6), fechasistema)
'
'
'        If diasmora <= diasgracia Then diasmora = 0
'
'
'        interes = Round(saldo * ((tazainteresmora  * diasmora) / 100), 0)
'
'        t1 = t1 + saldo
'        t2 = t2 + interes
'
'        total = saldo + interes
'        If saldo = 0 Then
'
'        GRID1.Cell(GRID1.Rows - 1, 6).text = "0"
'        Else
'         GRID1.Cell(GRID1.Rows - 1, 6).text = Format(saldo, "###,###,###")
'        End If
'
'        GRID1.Cell(GRID1.Rows - 1, 7).text = diasmora
'        GRID1.Cell(GRID1.Rows - 1, 8).text = interes
'        If total = 0 Then
'        GRID1.Cell(GRID1.Rows - 1, 9).text = "0"
'        Else
'
'        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total, "###,###,###")
'        End If
'        GRID1.Cell(GRID1.Rows - 1, 10).text = "0"
'
'        GRID1.Cell(GRID1.Rows - 1, 11).text = resultado(13)
'
'        totalusado = totalusado + total
'        If interes <> 0 Then moratotal = moratotal + total
'
            
            resultado.MoveNext
            Wend
        Else
       
        End If
         If Option2.Value = True Then
        GRID1.Rows = GRID1.Rows + 1
        BARRA.Max = BARRA.Max + 1
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 11).Borders(cellEdgeTop) = cellThick
        GRID1.Cell(GRID1.Rows - 1, 6).text = "TOTAL VENCIMIENTO"
        GRID1.Cell(GRID1.Rows - 1, 7).text = Format(separador, "dd-mm-yyyy")
        
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total1, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total2, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total3, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 12).text = Format(ACUMULADO, "###,###,###")
        
        total1 = 0
        total2 = 0
        total3 = 0
        
        End If
        
        If Option1.Value = True Then
        BARRA.Max = BARRA.Max + 1
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 11).Borders(cellEdgeTop) = cellThick
        GRID1.Cell(GRID1.Rows - 1, 5).text = "TOTAL "
        GRID1.Cell(GRID1.Rows - 1, 6).text = leerNombreCliente(separador)
        
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total1, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total2, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total3, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 12).text = Format(ACUMULADO, "###,###,###")
        
        total1 = 0
        total2 = 0
        total3 = 0
      
        
        End If
        
        
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 11).Borders(cellEdgeTop) = cellThick
        
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = "TOTALES GENERALES"
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total11, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total12, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total13, "###,###,###")
        
        
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        GRID1.AutoRedraw = True
        GRID1.Refresh
    End Sub


Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    GRID1.FixedRowColStyle = Fixed3D
    GRID1.CellBorderColorFixed = vbButtonShadow
    GRID1.ShowResizeTips = False
    GRID1.ReportTitles.Clear
    GRID1.PageSetup.CenterHorizontally = True
    GRID1.PageSetup.Orientation = cellLandscape
    
      
    GRID1.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    GRID1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    GRID1.PageSetup.HeaderAlignment = cellLeft
    GRID1.PageSetup.HeaderFont.Name = "Verdana"
    GRID1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "ENTRE EL DIA  :  " & dato1.text + "-" + dato2.text + "-" + dato3.text & " y " & dato4.text + "-" + dato5.text + "-" + dato6.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
'    If Option1.Value = True Then tipoListado = "CLIENTES MAAT"
'    If Option2.Value = True Then tipoListado = "CLIENTES SKORPIOS"
'    If Option3.Value = True Then tipoListado = "CLIENTES TODOS"
'
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = tipoListado
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    GRID1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & usuarioSistema
    GRID1.PageSetup.FooterAlignment = cellRight
    GRID1.PageSetup.FooterFont.Name = "Verdana"
    GRID1.PageSetup.FooterFont.Size = 7
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeTop) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeBottom) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeLeft) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeRight) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideVertical) = cellThin
    GRID1.Range(0, 1, 0, GRID1.Cols - 1).FontBold = True
    
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
           dato2.SetFocus
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
           dato3.SetFocus
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
            
            dato4.SetFocus
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
        dato5.SetFocus
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
           dato6.SetFocus
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
           Command2.SetFocus
            
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
   
Private Sub Option1_Click()
LEErCREDITOS
End Sub

Private Sub Option2_Click()
LEErCREDITOS
End Sub

Private Sub Option3_Click()
LEErCREDITOS
End Sub

Private Sub Option4_Click()
LEErCREDITOS
End Sub
Sub LEErTIPOSCLIENTES()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim linea As Double
    
        Set cSql.ActiveConnection = ventas
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM sv_tiposdeclientes "
        cSql.sql = cSql.sql + "ORDER BY codigo "
        cSql.Execute
        linea = 1
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
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

