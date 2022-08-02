VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cartolamantencion 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEMA AUTOMATICO BANCO SANTANDER"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14610
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   659
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   974
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmprincipal 
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   3625
      BackColor       =   16744576
      Caption         =   "DATOS DEL CHEQUE"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483639
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox ccdato1 
         Appearance      =   0  'Flat
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
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "codigo"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox ccdato2 
         Appearance      =   0  'Flat
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
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   1665
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Genera Informe"
         Height          =   375
         Left            =   8520
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox dato1 
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
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "codigo"
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   1
         Top             =   270
         Width           =   375
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1665
         Left            =   10320
         TabIndex        =   6
         Top             =   240
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   2937
         BackColor       =   14737632
         Caption         =   "Rangos de Fecha"
         CaptionEstilo3D =   2
         BackColor       =   14737632
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         ColorBarraArriba=   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin CoolButtons.cool_Button command8 
            Height          =   375
            Left            =   1440
            TabIndex        =   7
            Top             =   1260
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            SkinId          =   "13"
            Caption         =   "Cambia Fecha"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde Fecha"
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
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta Fecha"
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
            Height          =   375
            Left            =   2160
            TabIndex        =   10
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label desdefecha 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label hastafecha 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2160
            TabIndex        =   8
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   2
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox PIVOTE 
         Height          =   285
         Left            =   18360
         TabIndex        =   5
         Top             =   8160
         Visible         =   0   'False
         Width           =   735
      End
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   7200
         TabIndex        =   18
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         BackColor       =   16744576
         Caption         =   " Mis Datos"
         BackColor       =   16744576
         BordeColor      =   4194304
         ColorBarraArriba=   4194304
         ColorBarraAbajo =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton botonmisfavoritos 
            Caption         =   "Mis Favoritos"
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   280
            Width           =   1335
         End
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   280
            Width           =   1455
         End
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CRCC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label ccnombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3360
         TabIndex        =   23
         Top             =   720
         Width           =   4590
      End
      Begin VB.Label lblnombrecuenta 
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   4545
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp frminforme 
      Height          =   7785
      Left            =   0
      TabIndex        =   14
      Top             =   2040
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   13732
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483639
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ProgressBar barra 
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   7560
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7335
         Left            =   0
         TabIndex        =   16
         Top             =   240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   12938
         AllowUserSort   =   -1  'True
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
   End
End
Attribute VB_Name = "cartolamantencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecha1 As String
Dim fecha2 As String
Dim CtaMayorCrcc As Boolean
Dim CtaMayorCtaCte As Boolean
Sub CARGAGRILLA()
Dim FORMATOGRILLA(50, 50) As String
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "LIN"
    FORMATOGRILLA(1, 5) = "CUENTA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "TP"
    FORMATOGRILLA(1, 8) = "NUMERO"
    FORMATOGRILLA(1, 9) = "EMISION"
    FORMATOGRILLA(1, 10) = "VENCE"
    FORMATOGRILLA(1, 11) = "DEBE"
    FORMATOGRILLA(1, 12) = "HABER"
    FORMATOGRILLA(1, 13) = "SALDO"
    FORMATOGRILLA(1, 14) = "RUT CTACTE"
    FORMATOGRILLA(1, 15) = "NOMBRE CTACTE"
    FORMATOGRILLA(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "3"
    
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "28"
    FORMATOGRILLA(2, 7) = "3"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "0"
    FORMATOGRILLA(2, 10) = "8"
    FORMATOGRILLA(2, 11) = "11"
    FORMATOGRILLA(2, 12) = "11"
    FORMATOGRILLA(2, 13) = "11"
    FORMATOGRILLA(2, 14) = "10"
    FORMATOGRILLA(2, 15) = "20"
    FORMATOGRILLA(2, 16) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "D"
    FORMATOGRILLA(3, 10) = "D"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    FORMATOGRILLA(3, 16) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 11) = "###,###,###,##0"
    FORMATOGRILLA(4, 12) = "###,###,###,##0"
    FORMATOGRILLA(4, 13) = "###,###,###,##0"
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
    FORMATOGRILLA(5, 16) = "TRUE"
    FORMATOGRILLA(5, 17) = "TRUE"
    
    Grid1.Cols = 17
    Grid1.Rows = 2
    
    Grid1.DisplayFocusRect = False
    Grid1.BoldFixedCell = False
    Grid1.ExtendLastCol = True
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
    
End Sub


Sub leercuenta(cuenta, CRCC, fecha1, fecha2)
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Grid1.AutoRedraw = False
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
        csql2.sql = csql2.sql + "and codigo='" & cuenta & "' "
        csql2.sql = csql2.sql + "and año ='" & Format(fechasistema, "yyyy") & "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        
        Grid1.Rows = 1
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        barra.Max = csql2.RowsAffected + 1
        barra.Value = 0
        
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        barra.Value = barra.Value + 1
        
        
        If Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERMOVIMIENTOS(Grid1, resultados2(0), resultados2(1), ccdato1 & ccdato2)
       
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        Grid1.Column(8).Locked = True
        Grid1.Column(9).Locked = True
        Grid1.Column(10).Locked = True
        
Grid1.AutoRedraw = True
Grid1.Refresh

End Sub

Sub LEERMOVIMIENTOS(infogrilla As Grid, cuenta, NOMBRE, Optional CRCC As String)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    dedonde = 1
    barra.Value = 0
    
    If ccnombre.Caption = "" Then CRCC = ""
    
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
        fecha2 = Mid(hastafecha.Caption, 7, 4) + "-" + Mid(hastafecha.Caption, 4, 2) + "-" + Mid(hastafecha.Caption, 1, 2)
    
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento"
        csql.sql = csql.sql & ",fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte,monto "
        csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' "
        csql.sql = csql.sql & " and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        If CRCC <> "" Then
        
            csql.sql = csql.sql & " and centrocosto='" & CRCC & "' "
        End If
        csql.sql = csql.sql + "order by codigocuenta,fecha,tipo,numero,linea"
        csql.Execute
        
        Call LEERSALDOS(cuenta)
        
 
        
        
        If saldo <> 0 Or csql.RowsAffected <> 0 Then
        lin = lin + 1
        Grid1.Rows = Grid1.Rows + 1
        
        For k = 1 To 6
        Grid1.Column(k).Locked = False
        Next k
                
        Grid1.Range(lin, 1, lin, 6).Merge
      
        Grid1.Cell(lin, 1).CellType = cellTextBox
        
        Grid1.Cell(lin, 10).CellType = cellTextBox
        
        If dedonde = 1 Then
        Grid1.Cell(lin, 1).text = cuenta & " " + NOMBRE
        End If
        
        'If dedonde = 2 Then Grid1.Cell(lin, 6).text = nombrectacte
        Grid1.Cell(lin, 10).text = "SALDO-->"
        
        Grid1.Cell(lin, 13).text = saldo
        Grid1.Range(lin, 0, lin, Grid1.Cols - 1).FontBold = True
        Grid1.Range(lin, 0, lin, Grid1.Cols - 1).FontUnderline = True
        
        
        End If
        
        If csql.RowsAffected > 0 Then
        
        
        Set resultados = csql.OpenResultset
        barra.Max = csql.RowsAffected + 10
         While Not resultados.EOF
         barra.Value = barra.Value + 1
'        If dedonde = 1 And Check2.Value = 1 Then
'        If resultados(15) > 2 Then GoTo dale:
'        End If
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
             If IsNull(resultados("rutctacte")) = False Then Grid1.Cell(lin, 0).text = resultados("rutctacte")
             For k = 0 To 9
             If IsNull(resultados(k)) = False Then Grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             If resultados(11) = "D" Then Grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10): saldo = saldo + resultados(10)
             If resultados(11) = "H" Then Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10): saldo = saldo - resultados(10)
'             Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             Grid1.Cell(lin, 13).text = saldo
             If resultados("rutctacte") <> "" Then
                Grid1.Cell(lin, 14).text = resultados("rutctacte")
                Grid1.Cell(lin, 15).text = leerNombrerut(dato1 & dato2 & dato3, resultados("rutctacte"))
             End If
             If resultados("centrocosto") <> "" Then Grid1.Cell(lin, 16).text = resultados("centrocosto") & " " & leerNOMBREcrcc(resultados("centrocosto"))
             
dale:             resultados.MoveNext
          
         Wend
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
         
         Call totalcomprobante(Grid1, lin)
          resultados.Close
            Set resultados = Nothing

        End If
 For k = 1 To 6
        Grid1.Column(k).Locked = True
        
        Next k
        barra.Value = 0
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub ccdato1_Change()
ccnombre.Caption = Empty
End Sub

Private Sub ccdato1_GotFocus()
Call cargatexto(ccdato1)
End Sub

Private Sub ccdato1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudacrcc(ccdato1, ccdato2)
End Sub

Private Sub ccdato1_KeyPress(KeyAscii As Integer)
snum = 0: KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    Call ceros(ccdato1)
    ccdato2.SetFocus
End If
End Sub

Private Sub ccdato2_Change()
ccnombre.Caption = Empty
End Sub

Private Sub ccdato2_GotFocus()
Call cargatexto(ccdato2)
End Sub

Private Sub ccdato2_KeyPress(KeyAscii As Integer)
snum = 0: KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    Call ceros(ccdato2)
    ccnombre.Caption = leerNOMBREcrcc(ccdato1 & ccdato2)
End If
End Sub

Private Sub Command10_Click()
Call leercuenta(dato1 & dato2 & dato3, ccdato1 & ccdato2, desdefecha, hastafecha)
End Sub

Private Sub Command5_Click()
Grid1.PageSetup.Orientation = cellLandscape
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0.5
Grid1.PageSetup.Zoom = 85
Grid1.PrintPreview (90)
End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
Call Command10_Click
End Sub

Private Sub dato1_GotFocus()
Call cargatexto(dato1)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
End Sub

Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub

Private Sub dato3_GotFocus()
Call cargatexto(dato3)
End Sub

Private Sub Form_Load()
Me.Width = Screen.Width - 300
frmprincipal.Width = Me.Width - 100
frminforme.Width = frmprincipal.Width - 100
Grid1.Width = Me.Width - 200
barra.Width = Grid1.Width
fechas.Left = frmprincipal.Width - fechas.Width
FrameQuickMenu.Left = fechas.Left + fechas.Width + 10
Call CARGAGRILLA
desdefecha = "01-01-" & Format(fechasistema, "yyyy")
hastafecha = Format(fechasistema, "dd-mm-yyyy")
End Sub



Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)

    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
   
    Call ceros(dato3)
    LblnombreCuenta.Caption = leerNombreMayor(dato1.text + dato2.text + dato3.text)
    Call Pregunta(dato3, dato3)
    CARGAGRILLA
    
    Call leercuenta(dato1.text + dato2.text + dato3.text, "", "", "")
   
End If

End Sub

Sub leer()
    Rem lee cuenta madre
  
lee2:    Rem lee cuenta madre
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    
    
no:
   
End Sub
   
    


Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    
    
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    
    

End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" & Format(fechasistema, "yyyy") & "' "
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    dato2.text = Mid(pivote.text, 3, 2)
    dato3.text = Mid(pivote.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
    
no:
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub




Sub LEERSALDOS(cuenta)
Dim resultados3 As rdoResultset
    
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = "debe01"
    campos(5, 0) = "debe02"
    campos(6, 0) = "debe03"
    campos(7, 0) = "debe04"
    campos(8, 0) = "debe05"
    campos(9, 0) = "debe06"
    campos(10, 0) = "debe07"
    campos(11, 0) = "debe08"
    campos(12, 0) = "debe09"
    campos(13, 0) = "debe10"
    campos(14, 0) = "debe11"
    campos(15, 0) = "debe12"
    campos(16, 0) = "haber01"
    campos(17, 0) = "haber02"
    campos(18, 0) = "haber03"
    campos(19, 0) = "haber04"
    campos(20, 0) = "haber05"
    campos(21, 0) = "haber06"
    campos(22, 0) = "haber07"
    campos(23, 0) = "haber08"
    campos(24, 0) = "haber09"
    campos(25, 0) = "HABER10"
    campos(26, 0) = "HABER11"
    campos(27, 0) = "HABER12"
    campos(28, 0) = ""
    
    condicion = "codigo=" + "'" + cuenta + "'and año='" + Mid(desdefecha.Caption, 7, 4) + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(2, 3)) - Val(sqlconta.response(3, 3))
  
    saldo = sumador
Rem acumula fecha
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)

    
        
        Set cSql3.ActiveConnection = contadb
        cSql3.sql = "SELECT SUM(monto),dh "
         cSql3.sql = cSql3.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and fecha<'" + fecha1 + "' and fecha>='" + Format(fechasistema, "yyyy") + "-01-01" + "' "
     
        
        cSql3.sql = cSql3.sql + "GROUP by DH"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
         If resultados3(1) = "D" Then saldo = saldo + resultados3(0)
         If resultados3(1) = "H" Then saldo = saldo - resultados3(0)
         
             
             resultados3.MoveNext
           
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If

End Sub

Sub totalcomprobante(infogrilla As Grid, row)
    Grid1.Range(row, 11, row, 12).Borders(cellEdgeTop) = cellThin
    Grid1.Range(row, 1, row, 12).FontBold = True
    Grid1.Range(row, 1, row, 12).FontUnderline = True
    
    
    Grid1.Cell(row, 10).CellType = cellTextBox
    Grid1.Cell(row, 10).text = "TOTAL "
    Grid1.Cell(row, 11).text = anted
    Grid1.Cell(row, 12).text = anteh
    lin = lin + 2
    Grid1.Rows = Grid1.Rows + 2
        
    anted = 0: anteh = 0: saldo = 0
    End Sub

Private Sub Grid1_DblClick()
Dim row As Integer
row = Grid1.ActiveCell.row
If row > 0 Then
    If Grid1.Cell(row, 1).text <> "" And Grid1.Cell(row, 3).text <> "" Then
        Load FrmCambiaCuenta
        With FrmCambiaCuenta
              .td = Grid1.Cell(row, 2).text
              .LINEA = Format(Grid1.Cell(row, 4).text, "000")

              .numero = Grid1.Cell(row, 3).text
              .fecha = Grid1.Cell(row, 1).text
              .dato1 = Grid1.Cell(row, 5).text
              .dato4 = Grid1.Cell(row, 6).text
              .DATO5.text = Grid1.Cell(row, 7).text
              .dato6.text = Grid1.Cell(row, 8).text
              
              .dato7.text = Grid1.Cell(row, 7).text
              .dato8.text = Grid1.Cell(row, 8).text
              
              
              .DV.Caption = Right(Grid1.Cell(row, 14).text, 1)
              .lblnombrecta.Caption = LeerTipoCtaMayor(.dato1)
              
              If CtaMayorCrcc = True Then
              .dato2.Visible = True
              .Label1(3).Visible = True
              .lblnombrecrcc.Visible = True
                .dato2 = Mid(Grid1.Cell(row, 16).text, 1, 4)
                .lblnombrecrcc.Caption = leerNOMBREcrcc(.dato2)
              End If
              
              If CtaMayorCtaCte = True Then
              .dato3.Visible = True
              .lblnombrecta.Visible = True
              .dato3 = Mid(Grid1.Cell(row, 14).text, 1, 9)
              .lblnombrectacte.Caption = LeerTipoCtaMayor(.dato1)
              .Label1(5).Visible = True
              End If
              If .DATO5.text = "CH" And dato1.text & dato2.text = "1112" Then
                 .Label1(8).Visible = True
                 .Label1(9).Visible = True
                 .DATO5.Visible = True
                 .dato6.Visible = True
              End If
             .Show 1
        Call Command10_Click
        End With
        
    End If

End If


End Sub


Function LeerTipoCtaMayor(codigo) As String
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "ctacte"
    campos(3, 0) = "crcc"
    campos(4, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo=" + "'" + codigo + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    CtaMayorCtaCte = False
    CtaMayorCrcc = False
    If sqlconta.status = 4 Then
    
    Else
    If sqlconta.response(2, 3) = 1 Then CtaMayorCtaCte = True
    If sqlconta.response(3, 3) = 1 Then CtaMayorCrcc = True
    LeerTipoCtaMayor = sqlconta.response(1, 3)
    End If
no:

End Function


Sub ayudacrcc(primero As TextBox, segundo As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "año='" + año + "'"
    pivote.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote, campos, cfijo, largo, 2)
    primero.text = Mid(pivote.text, 1, 2)
    segundo.text = Mid(pivote.text, 3, 2)
    pivote.text = ""
End Sub

