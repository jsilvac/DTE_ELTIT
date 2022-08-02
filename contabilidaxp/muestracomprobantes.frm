VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form muestracomprobantes 
   BorderStyle     =   0  'None
   Caption         =   "Muestra Comprobantes"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFavoritos 
      BackColor       =   &H0080FF80&
      Caption         =   "Agregar a Favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   1935
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8745
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   15425
      BackColor       =   16744576
      Caption         =   "MUESTRA COMPROBANTES"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   65535
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
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6855
         Left            =   45
         TabIndex        =   22
         Top             =   810
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   12091
         BackColor       =   16744576
         Caption         =   "DATOS COMPROBANTE"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin FlexCell.Grid Grid1 
            Height          =   6585
            Left            =   0
            TabIndex        =   23
            Top             =   225
            Width           =   14985
            _ExtentX        =   26432
            _ExtentY        =   11615
            BackColor1      =   15859171
            BackColorFixed  =   12648384
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14625
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   45
         Width           =   375
      End
      Begin VB.Frame cabeza 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   45
         TabIndex        =   11
         Top             =   315
         Width           =   14895
         Begin VB.TextBox dato0 
            BackColor       =   &H00E1FFFD&
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
            Left            =   720
            MaxLength       =   2
            TabIndex        =   16
            Top             =   120
            Width           =   420
         End
         Begin VB.TextBox DATO2 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   9480
            MaxLength       =   2
            TabIndex        =   15
            Tag             =   "fecha"
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox DATO3 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   9840
            MaxLength       =   2
            TabIndex        =   14
            Tag             =   "fecha"
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox DATO4 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   10200
            MaxLength       =   4
            TabIndex        =   13
            Tag             =   "fecha"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox DATO1 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
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
            Left            =   6600
            MaxLength       =   10
            TabIndex        =   12
            Top             =   120
            Width           =   1425
         End
         Begin VB.Label tipocompro 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   20
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label Label5 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TIPO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label6 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FECHA :"
            Height          =   285
            Left            =   8640
            TabIndex        =   18
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label3 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FOLIO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5400
            TabIndex        =   17
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Datos personales"
         Height          =   735
         Left            =   8820
         TabIndex        =   1
         Top             =   7830
         Width           =   6015
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DEBE"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label label 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   9
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label label 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   6
            Left            =   2040
            TabIndex        =   8
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label label 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   7
            Left            =   2040
            TabIndex        =   7
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HABER"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   6
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SALDO"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4200
            TabIndex        =   5
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label debe 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label saldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   4200
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label haber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2280
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "muestracomprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private canli As Integer
    Private grilladetalle(5000, 13) As String
    Private FORMATOGRILLA(100, 100) As String
Sub leecomprobante()
    Dim lin As Integer
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut2 As String

    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,numero,linea,fecha,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh,rutctacte,centrocosto "
        csql.sql = csql.sql + "FROM movimientoscontables "
            
        csql.sql = csql.sql + "WHERE tipo='" + DATO0.text + "' and numero='" & dato1.text & "'and año='" + da4 + "' and mes='" + da3 + "' order by linea"
        csql.Execute

        canli = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
             Rem fechacon = resultados(3)
             canli = canli + 1
                rut2 = resultados(2)
                
                dato2.text = Mid(resultados(3), 1, 2)
                dato3.text = Mid(resultados(3), 4, 2)
                dato4.text = Mid(resultados(3), 7, 4)
                
                
                grilladetalle(canli, 1) = Mid(resultados(4), 1, 2)
                grilladetalle(canli, 2) = Mid(resultados(4), 3, 2)
                grilladetalle(canli, 3) = Mid(resultados(4), 5, 4)
                grilladetalle(canli, 4) = resultados(5)
                grilladetalle(canli, 5) = resultados(6)
                grilladetalle(canli, 6) = resultados(7)
                grilladetalle(canli, 7) = resultados(8)
                grilladetalle(canli, 8) = resultados(9)
                grilladetalle(canli, 9) = resultados(10)
                grilladetalle(canli, 10) = resultados(11)
                grilladetalle(canli, 11) = resultados(12)
                If resultados(6) = "FA" And resultados(7) <> "0000000000" And resultados(10) = "D" Then
               Rem  Call leecentrofactura(resultados(6), resultados(7), resultados(11))
                
                End If
                

                resultados.MoveNext
            Wend
            cargadorcomprobante
            resultados.Close
            Set resultados = Nothing
        End If
    
Rem leerglosa

no:
End Sub
Sub cargadorcomprobante()
    Dim LINEA As Long
    Grid1.AutoRedraw = False
    
    
    Grid1.Rows = canli + 1
    
    For k = 1 To canli
   ' CUENTAMAYOR(K) = grilladetalle(K, 1)
    Grid1.Cell(k, 0).text = k
    Grid1.Cell(k, 1).text = grilladetalle(k, 1)
    Grid1.Cell(k, 2).text = grilladetalle(k, 2)
    Grid1.Cell(k, 3).text = grilladetalle(k, 3)
    Grid1.Cell(k, 4).text = grilladetalle(k, 4)
    Grid1.Cell(k, 5).text = grilladetalle(k, 5)
    Grid1.Cell(k, 6).text = grilladetalle(k, 6)
    Grid1.Cell(k, 7).text = grilladetalle(k, 7)
    Grid1.Cell(k, 8).text = grilladetalle(k, 8)
    Grid1.Cell(k, 9).text = grilladetalle(k, 9)
    Grid1.Cell(k, 10).text = leerNombreMayor(grilladetalle(k, 1) + grilladetalle(k, 2) + grilladetalle(k, 3))
    
    Grid1.Cell(k, 11).text = ""
    Grid1.Cell(k, 12).text = ""
    
    Grid1.Cell(k, 13).text = grilladetalle(k, 10)
    Grid1.Cell(k, 14).text = grilladetalle(k, 11)
    
    
    LINEA = k
    
    'Call leermayor(linea, 9999)
    'If Val(Mid(grilladetalle(linea, 10), 1, 9)) <> 0 Then Call leerctacte(linea, 9999)
    'If Val(grilladetalle(linea, 11)) <> 0 Then Call leercrcc(linea, 9999)

    SUMAR

    Next k
    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    
End Sub

 
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_Load()

    Call CARGAGRILLA(2, 16)
DATO0.text = da0
dato1.text = da1
dato2.text = da2
dato3.text = da3
dato4.text = da4


leecomprobante


End Sub
Sub CARGAGRILLA(row, col)
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "C1"
    FORMATOGRILLA(1, 2) = "C2"
    FORMATOGRILLA(1, 3) = "C3"
    FORMATOGRILLA(1, 4) = "GLOSA"
    FORMATOGRILLA(1, 5) = "TP"
    FORMATOGRILLA(1, 6) = "NUMERO"
    FORMATOGRILLA(1, 7) = "F.VENCI."
    FORMATOGRILLA(1, 8) = "MONTO"
    FORMATOGRILLA(1, 9) = "D/H"
    FORMATOGRILLA(1, 10) = "MAYOR"
    FORMATOGRILLA(1, 11) = "CTACTE"
    FORMATOGRILLA(1, 12) = "CRCC"
    FORMATOGRILLA(1, 13) = "RUT"
    FORMATOGRILLA(1, 14) = "CRCC"
    FORMATOGRILLA(1, 15) = "ANALISIS"
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "2"
    FORMATOGRILLA(2, 2) = "2"
    FORMATOGRILLA(2, 3) = "4"
    FORMATOGRILLA(2, 4) = "60"
    FORMATOGRILLA(2, 5) = "2"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "12"
    FORMATOGRILLA(2, 9) = "3"
    FORMATOGRILLA(2, 10) = "15"
    FORMATOGRILLA(2, 11) = "15"
    FORMATOGRILLA(2, 12) = "15"
    FORMATOGRILLA(2, 13) = "10"
    FORMATOGRILLA(2, 14) = "4"
    FORMATOGRILLA(2, 15) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "C"
    FORMATOGRILLA(3, 2) = "C"
    FORMATOGRILLA(3, 3) = "C"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "C"
    FORMATOGRILLA(3, 7) = "D"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    FORMATOGRILLA(4, 7) = ""
    FORMATOGRILLA(4, 8) = "$ ###,###,##0"
    FORMATOGRILLA(4, 9) = ""
    FORMATOGRILLA(4, 10) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    
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
    Grid1.Column(0).Width = 4 * 8.8
    Grid1.Column(1).Width = 2 * 10
    Grid1.Column(2).Width = 2 * 10
    Grid1.Column(3).Width = 4 * 10
    Grid1.Column(4).Width = 40 * 9
    Grid1.Column(5).Width = 3 * 9
    Grid1.Column(6).Width = 8 * 9
    Grid1.Column(7).Width = 8 * 9
    Grid1.Column(8).Width = 12 * 9
    Grid1.Column(9).Width = 3 * 9
    Grid1.Column(13).Width = 100
    Grid1.Column(14).Width = 40
    Grid1.Column(15).Width = 100
    For k = 1 To col - 1
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then
            Grid1.Column(k).Alignment = cellRightCenter
            Grid1.Column(k).Mask = cellNumeric
        End If
        If FORMATOGRILLA(3, k) = "S" Then
            Grid1.Column(k).Alignment = cellLeftCenter
            'Grid1.Column(K).Mask = cellUpper
        End If
        If FORMATOGRILLA(3, k) = "D" Then
            Grid1.Column(k).CellType = cellCalendar
            Grid1.Column(k).Mask = cellNumeric
        End If
        
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    
    Grid1.Column(0).Locked = False
    Grid1.Column(1).Locked = False
    Grid1.Column(2).Locked = False
    Grid1.Column(3).Locked = False
    
    
    
    Grid1.Range(0, 1, 0, 3).Merge
    Grid1.Cell(0, 1).text = "CUENTA"
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    
    
End Sub

Sub SUMAR()
Dim o As Integer

sumadebe = 0
sumahaber = 0

For o = 1 To Grid1.Rows - 1
If Grid1.Cell(o, 12).text = "D" Then sumadebe = sumadebe + Grid1.Cell(o, 11).text
If Grid1.Cell(o, 12).text = "H" Then sumahaber = sumahaber + Grid1.Cell(o, 11).text
Next o
debe.Caption = Format(sumadebe, "###,###,###,##0")
haber.Caption = Format(sumahaber, "###,###,###,##0")
saldo.Caption = Format(sumadebe - sumahaber, "###,###,###,##0")
End Sub

Private Sub Grid1_DblClick()

If Grid1.Cell(Grid1.ActiveCell.row, 5).text = "CH" And Grid1.Cell(Grid1.ActiveCell.row, 9).text = "H" Then
    da1 = Grid1.Cell(Grid1.ActiveCell.row, 1).text
    da2 = Grid1.Cell(Grid1.ActiveCell.row, 2).text
    da3 = Grid1.Cell(Grid1.ActiveCell.row, 3).text
    da4 = Grid1.Cell(Grid1.ActiveCell.row, 6).text
    muestrach.Show vbModal
End If
If (Grid1.Cell(Grid1.ActiveCell.row, 5).text = "NC" Or Grid1.Cell(Grid1.ActiveCell.row, 5).text = "FC") And Grid1.Cell(Grid1.ActiveCell.row, 13).text <> "" Then
    If Grid1.Cell(Grid1.ActiveCell.row, 5).text = "NC" Then
        If Mid(Grid1.Cell(Grid1.ActiveCell.row, 4).text, 1, 31) = "CENTRALIZA DOCUMENTO DE COMPRAS" Then
        da5 = "6"
        Else
        da5 = "3"
        End If
    
    Else
    If Mid(Grid1.Cell(Grid1.ActiveCell.row, 4).text, 1, 31) = "CENTRALIZA DOCUMENTO DE COMPRAS" Then
        da5 = "4"
        Else
        da5 = "1"
        End If
    
    End If
    
    da6 = Grid1.Cell(Grid1.ActiveCell.row, 6).text
    da7 = Grid1.Cell(Grid1.ActiveCell.row, 13).text
'    If MsgBox("DESEA SOLO VER LA FACTURA ", vbYesNo) = vbYes Then
'
    Load ingreso22
    ingreso22.Show vbModal
'    Else
'    ingreso04.dato1.text = da5
'    ingreso04.dato2.text = da2
'    ingreso04.dato9.text = da6
'    ingreso04.Show vbModal
'
'
'
'End If

End If

End Sub
