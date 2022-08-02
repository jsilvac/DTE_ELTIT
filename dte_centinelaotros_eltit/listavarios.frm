VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form informe05 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Varios Digitados"
   ClientHeight    =   8775
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14655
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1320
      Left            =   45
      TabIndex        =   4
      Top             =   90
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   2328
      BackColor       =   16761024
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
         Caption         =   "Genera Informe"
         Height          =   285
         Left            =   12600
         TabIndex        =   15
         Top             =   765
         Width           =   1860
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5175
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   405
         Width           =   615
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4815
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "fecha"
         Top             =   405
         Width           =   375
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4455
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "fecha"
         Top             =   405
         Width           =   375
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2100
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "fecha"
         Top             =   405
         Width           =   615
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1740
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "fecha"
         Top             =   405
         Width           =   375
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "fecha"
         Top             =   405
         Width           =   375
      End
      Begin VB.ComboBox cmbRazon 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   855
         Width           =   5790
      End
      Begin MSAdodcLib.Adodc razon 
         Height          =   330
         Left            =   0
         Top             =   0
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
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   1230
         Left            =   10170
         TabIndex        =   16
         Top             =   45
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   2170
         BackColor       =   12648384
         Caption         =   "LOCALES"
         CaptionEstilo3D =   1
         BackColor       =   12648384
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
         Begin VB.OptionButton local1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Pucon"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   19
            Top             =   315
            Value           =   -1  'True
            Width           =   1860
         End
         Begin VB.OptionButton local2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Villarrica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   630
            Width           =   1860
         End
         Begin VB.OptionButton local3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ambos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   17
            Top             =   945
            Width           =   1860
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA TERMINO"
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
         Height          =   255
         Left            =   2790
         TabIndex        =   14
         Top             =   450
         Width           =   1545
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIO"
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
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RAZON SOCIAL"
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
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   855
         Width           =   1695
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6810
      Left            =   45
      TabIndex        =   2
      Top             =   1440
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   12012
      BackColor       =   16761024
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
      Begin FlexCell.Grid grid1 
         Height          =   6735
         Left            =   0
         TabIndex        =   3
         Top             =   225
         Width           =   14490
         _ExtentX        =   25559
         _ExtentY        =   11880
         Cols            =   8
         DefaultFontSize =   8.25
         Rows            =   1
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin FlexCell.Grid Impresion 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
      DateFormat      =   2
   End
   Begin VB.CommandButton cmdImprime 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Imprimir"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   3855
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   0
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
End
Attribute VB_Name = "informe05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(20, 20)
    Private fecha1 As String
    Private fecha2 As String
    Private titulo As String
    
Private Sub cmdImprime_Click()
Dim TITU As String
Dim TITU2 As String
Dim titu3 As String

If local1.Value = True Then titu3 = "PUCON"
If local2.Value = True Then titu3 = "VILLARRICA"
If local3.Value = True Then titu3 = "AMBOS "

TITU2 = "PERIODO " + dato1.text + "-" + dato2.text + "-" + dato3.text + " AL " + dato4.text + "-" + dato5.text + "-" + dato6.text
Call cabezas2("LISTADO VARIOS INGRESADOS " + TITU + titu3, TITU2)


grid1.PrintPreview

End Sub
Sub cabezas2(titulo, titulo2)
Dim objReportTitle As FlexCell.ReportTitle
grid1.ReportTitles.Clear
grid1.PageSetup.Orientation = cellLandscape




DATOSEMPRESA(1) = "EMPRESAS ELTIT "

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo2
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = cellLeft
        grid1.ReportTitles.Add objReportTitle
    Next k
    
    grid1.Range(0, 0, 0, grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
    grid1.Range(0, 0, 0, grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
    grid1.Range(0, 0, 0, grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
    grid1.Range(0, 0, 0, grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
    grid1.Range(0, 0, 0, grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    grid1.Range(0, 0, 0, grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
    
    
    
    
    
    
With grid1.PageSetup
        
        .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        .LeftMargin = 0.5
        .RightMargin = 0.5
        .BlackAndWhite = True
        .PrintFixedRow = True
        
        
        
        
        
        
        
End With

End Sub

Private Sub Command1_Click()
LEERCARTERA

End Sub

'****************************************************************************
'Manejo de los Controles
'****************************************************************************
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    Private Sub DATO1_GotFocus()
        Call cargatexto(dato1)
    End Sub

    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call cargatexto(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        Call cargatexto(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call cargatexto(dato6)
    End Sub
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    
    '****************************************************************************
    'KEYDOWN
    '****************************************************************************
    Private Sub DATO1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato1, KeyCode, dato1)
    End Sub

    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato2, KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato3, KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato4, KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato5, KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato6, KeyCode, dato5)
    End Sub
    '****************************************************************************
    'KEYDOWN
    '****************************************************************************
    
    '****************************************************************************
    'KEYPRESS
    '****************************************************************************
    Private Sub DATO1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            dato4.text = dato1.text
            dato5.text = dato2.text
            dato6.text = dato3.text
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato4)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato5)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato6)
            SendKeys "{Tab}"
        End If
        fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
    End Sub
    '****************************************************************************
    'KEYPRESS
    '****************************************************************************

    Private Sub Form_Load()
        Call Centrar(Me)
        dato1.text = Format(fechasistema, "dd")
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        dato4.text = Format(fechasistema, "dd")
        dato5.text = Format(fechasistema, "mm")
        dato6.text = Format(fechasistema, "yyyy")
        'cmbTipo.ListIndex = 0
        cargaRazon
        
        
    End Sub
Sub LEERCARTERA()
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim TOTAL As Double
    Dim CHEQUES As Double
    
    Call CargaGrilla(1, 11)
    fecha1 = dato3.text + "-" + dato2.text + "-" + dato1.text
    fecha2 = dato6.text + "-" + dato5.text + "-" + dato4.text
    TOTAL = 0
    CHEQUES = 0
    grid1.AutoRedraw = False
      
      
        Set cSql2.ActiveConnection = teso
        cSql2.sql = "SELECT fecha,cajera,local,cuenta,glosa,tipo,numero,monto,dh,rut,crcc "
        cSql2.sql = cSql2.sql + "FROM rc_contabilidad where fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        
        If Mid(cmbRazon.text, 1, 5) <> "TODOS" Then
        cSql2.sql = cSql2.sql + "and local='" + Mid(cmbRazon.text, 1, 2) + "' "
        End If
        
        If local1.Value = True Then
        cSql2.sql = cSql2.sql + "and local<>'43' and local<>'42' and local<>'77' "
        End If
        If local2.Value = True Then
        cSql2.sql = cSql2.sql + "and (local='43' or local='42' or local='77') "
        End If
        
        
        
        cSql2.sql = cSql2.sql + "order by fecha,local,cuenta "
        cSql2.Execute
        grid1.Rows = 1
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        grid1.Rows = grid1.Rows + 1
        grid1.Cell(grid1.Rows - 1, 1).text = Format(resultados2(0), "dd-mm-yyyy")
        
        grid1.Cell(grid1.Rows - 1, 2).text = resultados2(1)
        grid1.Cell(grid1.Rows - 1, 3).text = leerLocal(resultados2(2))
        grid1.Cell(grid1.Rows - 1, 4).text = resultados2(3)
        grid1.Cell(grid1.Rows - 1, 5).text = leecuentacontable(resultados2(3))
        
        grid1.Cell(grid1.Rows - 1, 6).text = resultados2(4)
        grid1.Cell(grid1.Rows - 1, 7).text = resultados2(5)
        grid1.Cell(grid1.Rows - 1, 8).text = resultados2(6)
        
        grid1.Cell(grid1.Rows - 1, 9).text = Format(resultados2(7), "###,###,###")
        grid1.Cell(grid1.Rows - 1, 10).text = resultados2(8)
        
        
        
        TOTAL = TOTAL + resultados2(7)
        CHEQUES = CHEQUES + 1
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing
        grid1.Rows = grid1.Rows + 1
        
        End If
      
        grid1.Cell(grid1.Rows - 1, 9).Font.Size = 10
        grid1.Cell(grid1.Rows - 1, 9).Font.Bold = True
        grid1.Cell(grid1.Rows - 1, 9).text = Format(TOTAL, "###,###,###")
          
        grid1.Cell(grid1.Rows - 1, 6).CellType = cellTextBox
        grid1.Cell(grid1.Rows - 1, 6).Font.Size = 10
        grid1.Cell(grid1.Rows - 1, 6).Font.Bold = True
        grid1.Cell(grid1.Rows - 1, 6).text = "MONTO"
        
        
      
        
        grid1.Range(grid1.Rows - 1, 4, grid1.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
        grid1.Range(grid1.Rows - 1, 4, grid1.Rows - 1, 9).Borders(cellEdgeBottom) = cellThin
        grid1.Range(grid1.Rows - 1, 4, grid1.Rows - 1, 9).Borders(cellEdgeLeft) = cellThin
        grid1.Range(grid1.Rows - 1, 4, grid1.Rows - 1, 9).Borders(cellEdgeRight) = cellThin
        grid1.AutoRedraw = True
        
        
        grid1.Refresh
        
        
        
    End Sub
    
'****************************************************************************
'Manejo de los Controles
'****************************************************************************


Sub CargaGrilla(ByVal row As Long, ByVal col As Long)
    Dim i As Long
    Rem DATOS DE LA COLUMNA
    grid1.DefaultFont.Size = 8
    formatogrilla(1, 1) = "INGRESO"
    formatogrilla(1, 2) = "CAJERA"
    formatogrilla(1, 3) = "LOCAL"
    formatogrilla(1, 4) = "CUENTA"
    formatogrilla(1, 5) = "NOMBRE"
    formatogrilla(1, 6) = "GLOSA"
    formatogrilla(1, 7) = "TIPO"
    formatogrilla(1, 8) = "NUMERO"
    formatogrilla(1, 9) = "MONTO"
    formatogrilla(1, 10) = "D/H"
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "10"
    formatogrilla(2, 3) = "10"
    formatogrilla(2, 4) = "10"
    formatogrilla(2, 5) = "10"
    formatogrilla(2, 6) = "10"
    formatogrilla(2, 7) = "10"
    formatogrilla(2, 8) = "10"
    formatogrilla(2, 9) = "10"
    formatogrilla(2, 10) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "D"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "S"
    formatogrilla(3, 4) = "S"
    formatogrilla(3, 5) = "S"
    formatogrilla(3, 6) = "S"
    formatogrilla(3, 7) = "S"
    formatogrilla(3, 8) = "S"
    formatogrilla(3, 9) = "N"
    formatogrilla(3, 10) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = ""
    formatogrilla(4, 4) = ""
    formatogrilla(4, 5) = ""
    formatogrilla(4, 6) = ""
    formatogrilla(4, 7) = ""
    formatogrilla(4, 8) = ""
    formatogrilla(4, 9) = "###,###,###"
    formatogrilla(4, 10) = ""
    
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    formatogrilla(5, 3) = "FALSE"
    formatogrilla(5, 4) = "FALSE"
    formatogrilla(5, 5) = "FALSE"
    formatogrilla(5, 6) = "FALSE"
    formatogrilla(5, 7) = "FALSE"
    formatogrilla(5, 8) = "FALSE"
    formatogrilla(5, 9) = "FALSE"
    formatogrilla(5, 10) = "FALSE"
    
    Rem ANCHO
    formatogrilla(6, 1) = "8"
    formatogrilla(6, 2) = "8"
    formatogrilla(6, 3) = "20"
    formatogrilla(6, 4) = "10"
    formatogrilla(6, 5) = "20"
    formatogrilla(6, 6) = "20"
    formatogrilla(6, 7) = "4"
    formatogrilla(6, 8) = "8"
    formatogrilla(6, 9) = "10"
    formatogrilla(6, 10) = "3"
    
    grid1.Cols = col
    grid1.Rows = row
    
    
    grid1.AllowUserResizing = False
    grid1.DisplayFocusRect = False
    grid1.ExtendLastCol = True
    grid1.BoldFixedCell = False
    grid1.DrawMode = cellOwnerDraw
    grid1.Appearance = Flat
    grid1.ScrollBarStyle = Flat
    grid1.FixedRowColStyle = Flat

    grid1.BackColorFixed = RGB(90, 158, 214)
    grid1.BackColorFixedSel = RGB(110, 180, 230)
    grid1.BackColorBkg = RGB(90, 158, 214)
    grid1.BackColorScrollBar = RGB(231, 235, 247)
    grid1.BackColor1 = RGB(231, 235, 247)
    grid1.BackColor2 = RGB(239, 243, 255)
    grid1.GridColor = RGB(148, 190, 231)
    grid1.Column(0).Width = 0
    
    For i = 1 To grid1.Cols - 1
        grid1.Cell(0, i).text = formatogrilla(1, i)
        grid1.Column(i).Width = Val(formatogrilla(6, i)) * grid1.DefaultFont.Size
        grid1.Column(i).MaxLength = Val(formatogrilla(2, i))
        grid1.Column(i).FormatString = formatogrilla(4, i)
        grid1.Column(i).Locked = formatogrilla(5, i)
        If formatogrilla(3, i) = "N" Then grid1.Column(i).Alignment = cellRightCenter
        If formatogrilla(3, i) = "S" Then grid1.Column(i).Alignment = cellLeftCenter
        If formatogrilla(3, i) = "D" Then grid1.Column(i).CellType = cellCalendar
    Next i
    grid1.Range(0, 0, 0, grid1.Cols - 1).Alignment = cellCenterCenter
    
    
End Sub



Private Sub cargaRazon()
    Dim tabla As String
    tabla = "SELECT CONCAT(codigoempresa, ' ', nombre) AS item FROM maestroempresas ORDER BY codigoempresa"
    Call ConectarControlData(razon, servidor, clientesistema + "conta", usuario, password, tabla)
    If razon.Recordset.RecordCount > 0 Then
        razon.Recordset.MoveFirst
        While Not razon.Recordset.EOF
            cmbRazon.AddItem razon.Recordset.Fields("item")
            razon.Recordset.MoveNext
        Wend
        cmbRazon.AddItem "TODOS"
        cmbRazon.text = "TODOS"
    End If
End Sub



Private Function leerEmpresa(ByVal CODIGO As String) As String
    Dim condicion As String
    Dim op As Integer
    'Set sql = New CSQLUtil
    campos(0, 0) = "nombre"
    campos(1, 0) = "rut"
    campos(2, 0) = "direccion"
    campos(3, 0) = "ciudad"
    campos(4, 0) = ""
    
    campos(0, 2) = "maestroempresas"

    condicion = "codigoempresa = '" & CODIGO & "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then
        leerEmpresa = SQLUTIL.datos(0, 3) & vbCrLf & SQLUTIL.datos(1, 3) & vbCrLf & SQLUTIL.datos(2, 3) & vbCrLf & SQLUTIL.datos(3, 3)
    Else
        leerEmpresa = ""
    End If
End Function


Private Sub Informe_AfterReorderColumn(ByVal OriginalPosition As Long, ByVal NewPosition As Long)

End Sub

Private Sub local1_Click()
Command1_Click

End Sub

Private Sub local2_Click()
Command1_Click

End Sub

Private Sub local3_Click()
Command1_Click

End Sub

Private Sub Option1_Click()
Command1_Click

End Sub

Private Sub Option2_Click()
Command1_Click

End Sub

Private Sub Option3_Click()
Command1_Click

End Sub
