VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado3 
   Caption         =   "LISTADO CREDITOS OTORGADOS"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1005
      Left            =   90
      TabIndex        =   7
      Top             =   120
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
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pucon"
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
         Left            =   9045
         TabIndex        =   17
         Top             =   585
         Width           =   1140
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ambos"
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
         Left            =   10305
         TabIndex        =   16
         Top             =   585
         Width           =   1320
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Villarrica"
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
         Left            =   7470
         TabIndex        =   15
         Top             =   585
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Skorpios"
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
         Left            =   9450
         TabIndex        =   14
         Top             =   990
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Maat"
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
         Left            =   7155
         TabIndex        =   13
         Top             =   1035
         Visible         =   0   'False
         Width           =   2175
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
         Left            =   4545
         TabIndex        =   12
         Top             =   540
         Width           =   2175
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
         Left            =   2205
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
         Left            =   1755
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
         Left            =   2655
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
         Left            =   2655
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
         Left            =   2205
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
         Left            =   1755
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
      Height          =   8025
      Left            =   45
      TabIndex        =   6
      Top             =   1215
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   14155
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
         Left            =   5805
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7470
         Width           =   2760
      End
      Begin FlexCell.Grid GRID1 
         Height          =   7080
         Left            =   45
         TabIndex        =   8
         Top             =   315
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   12488
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "tmplistado3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Call Titulos("LISTADO DE CREDITOS OTORGADOS")
GRID1.PageSetup.Orientation = cellLandscape
GRID1.PageSetup.HeaderMargin = 0.5
GRID1.PageSetup.PrintFixedRow = True



GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0.5
GRID1.PageSetup.BottomMargin = 3
GRID1.PageSetup.FooterMargin = 2
GRID1.PageSetup.BlackAndWhite = True

GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeTop) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeBottom) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeLeft) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeRight) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideVertical) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideHorizontal) = cellThin




GRID1.PrintPreview
End Sub

Private Sub Command2_Click()
Call CargaGrillaGRID1(1, 12)
LEErCREDITOS

End Sub



Private Sub Form_Load()
Call CargaGrillaGRID1(1, 13)
dato1.text = Format(fechasistema, "DD")
dato2.text = Format(fechasistema, "MM")
dato3.text = Format(fechasistema, "YYYY")
dato4.text = Format(fechasistema, "DD")
dato5.text = Format(fechasistema, "MM")
dato6.text = Format(fechasistema, "YYYY")
Option3.Value = True
If empresaActiva = "42" Or empresaActiva = "43" Then
Option3.Value = True
Else
Option4.Value = True

End If


LEErCREDITOS

End Sub

 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatoGrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "LOCAL"
        formatoGrilla(1, 2) = "FECHA"
        formatoGrilla(1, 3) = "TIPO"
        formatoGrilla(1, 4) = "NUMERO"
        formatoGrilla(1, 5) = "RUT"
        formatoGrilla(1, 6) = "CLIENTE"
        formatoGrilla(1, 7) = "CUOTAS"
        formatoGrilla(1, 8) = "VALOR"
        formatoGrilla(1, 9) = "TOTAL"
        formatoGrilla(1, 10) = "CAPITAL"
        formatoGrilla(1, 11) = "INTERESES"
        formatoGrilla(1, 12) = "CAJERA"
        
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "10"
        formatoGrilla(2, 2) = ""
        formatoGrilla(2, 3) = ""
        formatoGrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "D"
        formatoGrilla(3, 3) = "S"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "S"
        formatoGrilla(3, 7) = "N"
        formatoGrilla(3, 8) = "N"
        formatoGrilla(3, 9) = "N"
        formatoGrilla(3, 10) = "N"
        formatoGrilla(3, 11) = "N"
        formatoGrilla(3, 12) = "S"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""

        Rem LOCCKED
        formatoGrilla(5, 1) = "TRUE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "TRUE"
        formatoGrilla(5, 4) = "TRUE"
        formatoGrilla(5, 5) = "TRUE"
        formatoGrilla(5, 6) = "TRUE"
        formatoGrilla(5, 7) = "TRUE"
        formatoGrilla(5, 8) = "TRUE"
        formatoGrilla(5, 9) = "TRUE"
        formatoGrilla(5, 10) = "TRUE"
        formatoGrilla(5, 11) = "TRUE"
        formatoGrilla(5, 12) = "TRUE"

        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "4"
        formatoGrilla(8, 2) = "7"
        formatoGrilla(8, 3) = "4"
        formatoGrilla(8, 4) = "8"
        formatoGrilla(8, 5) = "8"
        formatoGrilla(8, 6) = "25"
        formatoGrilla(8, 7) = "6"
        formatoGrilla(8, 8) = "8"
        formatoGrilla(8, 9) = "8"
        formatoGrilla(8, 10) = "8"
        formatoGrilla(8, 11) = "8"
        formatoGrilla(8, 12) = "8"
            
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
            GRID1.Cell(0, i).text = formatoGrilla(1, i)
            GRID1.Column(i).Width = Val(formatoGrilla(8, i)) * (GRID1.Cell(0, i).Font.Size + 1.25)
            GRID1.Column(i).MaxLength = Val(formatoGrilla(2, i))
            GRID1.Column(i).FormatString = formatoGrilla(4, i)
            GRID1.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                GRID1.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                GRID1.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
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
        Dim total21 As Double
        Dim total22 As Double
        Dim total23 As Double
        Dim total24 As Double
        Dim total25 As Double
        
        Dim fecha1 As String
        Dim fecha2 As String
        Dim loc As String
        
        fecha1 = dato3.text + "-" + dato2.text + "-" + dato1.text
        fecha2 = dato6.text + "-" + dato5.text + "-" + dato4.text
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT cd.local,cd.fechacompra,cd.tipo,cd.numero,cd.rut,mc.nombre,cd.cantidadcuotas,cd.montocuota,cd.cantidadcuotas*cd.montocuota,cd.montocredito,cajera "
        cSql.sql = cSql.sql & "FROM sv_maestroclientes as mc,sv_cuotas_detalle as cd "
        cSql.sql = cSql.sql & "WHERE cd.rut=mc.rut and cd.fechacompra between '" + fecha1 + "' and '" + fecha2 + "' and cd.tipo<>'CA' and cd.tipo<>'GC' "
'        If Option1.Value = True Then
'        cSql.sql = cSql.sql & "and mc.credito='M' "
'        End If
'        If Option2.Value = True Then
'        cSql.sql = cSql.sql & "and mc.credito='T' "
'        End If
'
        If Option3.Value = True Then
        cSql.sql = cSql.sql & "and (cd.local='42' or cd.local='43')"
        End If
        
        If Option4.Value = True Then
        cSql.sql = cSql.sql & "and (cd.local<>'42' and cd.local<>'43')"
        End If
        
        
        cSql.sql = cSql.sql & "group by rut,tipo,numero order by cd.local,cd.rut"
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
            
        GRID1.Rows = 1
       GRID1.AutoRedraw = False
        
        total1 = 0
        total2 = 0
        total3 = 0
        total4 = 0
        total5 = 0
        total21 = 0
        total22 = 0
        total23 = 0
        total24 = 0
        total25 = 0
        loc = resultado(0)
        
        While Not resultado.EOF
        GRID1.Cols = 13
        GRID1.Rows = GRID1.Rows + 1
        
        If loc <> resultado(0) Then
        
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 11).Borders(cellEdgeTop) = cellThick
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = "TOTAL LOCAL"
        
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total3, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total4, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total5, "###,###,###")
        GRID1.Rows = GRID1.Rows + 2
        
        loc = resultado(0)
        total3 = 0
        total4 = 0
        total5 = 0
        End If
        
        
        
                
        GRID1.Cell(GRID1.Rows - 1, 1).text = resultado(0)
        GRID1.Cell(GRID1.Rows - 1, 2).text = resultado(1)
        GRID1.Cell(GRID1.Rows - 1, 3).text = Format(resultado(2), "dd-mm-yyyy")
        GRID1.Cell(GRID1.Rows - 1, 4).text = resultado(3)
        GRID1.Cell(GRID1.Rows - 1, 5).text = Mid(resultado(4), 1, 9) + "-" + Mid(resultado(4), 10, 1)
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(resultado(5), "###,###")
        GRID1.Cell(GRID1.Rows - 1, 7).text = Format(resultado(6), "###,###")
        GRID1.Cell(GRID1.Rows - 1, 8).text = Format(resultado(7), "###,###")
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(resultado(8), "###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(resultado(9), "###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(resultado(8) - resultado(9), "###,###")
        GRID1.Cell(GRID1.Rows - 1, 12).text = leerNombreCajera(resultado(10) + rut(resultado(10)))
        
        
        
        total1 = total1 + resultado(6)
        total2 = total2 + resultado(7)
        total3 = total3 + resultado(8)
        total4 = total4 + resultado(9)
        total5 = total5 + (resultado(8) - resultado(9))
        
        total21 = total21 + resultado(6)
        total22 = total22 + resultado(7)
        total23 = total23 + resultado(8)
        total24 = total24 + resultado(9)
        total25 = total25 + (resultado(8) - resultado(9))
        
        

        
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
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 11).Borders(cellEdgeTop) = cellThick
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = "TOTAL LOCAL"
        
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total3, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total4, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total5, "###,###,###")
        GRID1.Rows = GRID1.Rows + 2
        
       
       
        total3 = 0
        total4 = 0
        total5 = 0
        
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 6, GRID1.Rows - 1, 11).Borders(cellEdgeTop) = cellThick
        
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = "TOTALES GENERALES"
        
       ' GRID1.Cell(GRID1.Rows - 1, 7).text = Format(total1, "###,###,###")
       ' GRID1.Cell(GRID1.Rows - 1, 8).text = Format(total2, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 9).text = Format(total23, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 10).text = Format(total24, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 11).text = Format(total25, "###,###,###")
        
        
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        GRID1.AutoRedraw = True
        GRID1.Refresh
'        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
'        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
'        totaldeuda.Caption = Format(totalusado, "###,###,##0")
'        moroso.Caption = Format(moratotal, "###,###,##0")
'
'
'        totalcuotas.Caption = Format(t1, "###,###,##0")
'        totalmora.Caption = Format(t2, "###,###,##0")
'
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
    
      
    GRID1.PageSetup.PrintTitleRows = 0
    
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
    
    tipoListado = "CLIENTES TODOS"
    
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

Private Sub Option5_Click()
LEErCREDITOS


End Sub
