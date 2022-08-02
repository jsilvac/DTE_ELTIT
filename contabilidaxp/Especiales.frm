VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form flujocaja2 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7830
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmDatos 
      Height          =   7755
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   13679
      BackColor       =   16761024
      Caption         =   "Ingreso de Datos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ColorBarraArriba=   12582912
      ColorBarraAbajo =   4194304
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
      Begin VB.TextBox dato5 
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
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "fecha"
         Top             =   6960
         Width           =   615
      End
      Begin VB.TextBox dato4 
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
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "fecha"
         Top             =   6960
         Width           =   375
      End
      Begin VB.TextBox dato3 
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
         Left            =   6960
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   6960
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "IMPRIMIR"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "ELIMINAR TODOS LOS PRECIOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9495
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7155
         Width           =   2040
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   1
         Top             =   6960
         Width           =   4935
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         MaxLength       =   15
         TabIndex        =   2
         Top             =   6960
         Width           =   1365
      End
      Begin XPFrame.FrameXp frmLista 
         Height          =   6075
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   10716
         BackColor       =   16761024
         Caption         =   "Lista de Datos"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ColorBarraAbajo =   16711680
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
         Begin FlexCell.Grid lista 
            Height          =   5355
            Left            =   90
            TabIndex        =   3
            Top             =   420
            Width           =   8370
            _ExtentX        =   14764
            _ExtentY        =   9446
            Cols            =   5
            DefaultFontSize =   9.75
            Rows            =   1
            SelectionMode   =   1
         End
         Begin MSAdodcLib.Adodc data 
            Height          =   330
            Left            =   60
            Top             =   4560
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
         Begin VB.Label LBLTOTAL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   5160
            TabIndex        =   10
            Top             =   5760
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   8280
         TabIndex        =   5
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   32896
         ColorBarraAbajo =   12648447
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
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEMANA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
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
         Height          =   315
         Left            =   5280
         TabIndex        =   7
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GLOSA"
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
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   6600
         Width           =   4905
      End
   End
End
Attribute VB_Name = "flujocaja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private FORMATOGRILLA(10, 10) As String
    Private modifica As Boolean
    Private preciocosto As Double
        
Private Sub COMMAND2_Click()
imprimir

End Sub
Sub imprimir()
Dim titulo As String
titulo = frmLista.Caption
Call CABEZAS2(titulo, "N", "000000000")
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeBottom) = cellThick
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeLeft) = cellThick
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeTop) = cellThick
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeRight) = cellThick
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellInsideHorizontal) = cellThick
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellInsideVertical) = cellThick
lista.DefaultFont.Size = 8
lista.PageSetup.Orientation = cellPortrait
lista.PageSetup.PrintFixedRow = True
lista.PageSetup.BottomMargin = 2
lista.PageSetup.TopMargin = 1
lista.PageSetup.LeftMargin = 1
lista.PageSetup.RightMargin = 0
lista.PageSetup.BlackAndWhite = True
lista.PageSetup.PrintGridlines = False
lista.PrintPreview 100

   
End Sub
Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
lista.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
    
    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 1
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = flujocaja.LBLEMPRESA.Caption
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        lista.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        lista.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        lista.ReportTitles.Add objReportTitle
        
    End If
    
With lista.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call cargatexto(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato1, dato2, KeyCode)
    End Sub
    
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
        If KeyAscii = 13 And dato1.text <> "" Then
        dato2.SetFocus
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        Dim Precio As String
        
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato2.text <> "" Then
            
            If modifica = False Then
                Call grabarEspeciales
                Call leerEspeciales(tipoflujo, fechaflujo)
            Else
                dato3.SetFocus
                
            End If
        
        End If
    End Sub
    
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaLista(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, 1) = "GLOSA"
        FORMATOGRILLA(1, 2) = "MONTO"
        
        Rem LARGO DE LOS DATOS
        FORMATOGRILLA(2, 1) = "40"
        FORMATOGRILLA(2, 2) = "15"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        FORMATOGRILLA(3, 1) = "S"
        FORMATOGRILLA(3, 2) = "N"
        
        Rem FORMATO GRILLA
        FORMATOGRILLA(4, 1) = ""
        FORMATOGRILLA(4, 2) = "###,###,###,##0"
        
        Rem LOCCKED
        FORMATOGRILLA(5, 1) = "TRUE"
        FORMATOGRILLA(5, 2) = "TRUE"
        
        Rem VALOR MINIMO
        FORMATOGRILLA(6, 1) = ""
        FORMATOGRILLA(6, 2) = ""
        
        Rem VALOR MAXIMO
        FORMATOGRILLA(7, 1) = ""
        FORMATOGRILLA(7, 2) = ""
        
        Rem ANCHO
        FORMATOGRILLA(8, 1) = "20"
        FORMATOGRILLA(8, 2) = "15"
            
        lista.Cols = col
        lista.Rows = row
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = True
        lista.BoldFixedCell = False
        lista.DrawMode = cellOwnerDraw
        lista.Appearance = Flat
        lista.ScrollBarStyle = Flat
        lista.FixedRowColStyle = Flat
        lista.BackColorFixed = RGB(90, 214, 158)
        lista.BackColorFixedSel = RGB(110, 230, 180)
        lista.BackColorBkg = RGB(90, 214, 158)
        lista.BackColorScrollBar = RGB(231, 247, 235)
        lista.BackColor1 = RGB(231, 247, 235)
        lista.BackColor2 = RGB(239, 255, 243)
        lista.GridColor = RGB(148, 231, 190)
        
        lista.Column(0).Width = 0
        For i = 1 To col - 1
            lista.Cell(0, i).text = FORMATOGRILLA(1, i)
            lista.Column(i).Width = Val(FORMATOGRILLA(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(FORMATOGRILLA(2, i))
            lista.Column(i).FormatString = FORMATOGRILLA(4, i)
            lista.Column(i).Locked = FORMATOGRILLA(5, i)
            If FORMATOGRILLA(3, i) = "N" Then
                lista.Column(i).Alignment = cellRightCenter
            Else
                lista.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        lista.Range(0, 1, 0, lista.Cols - 1).Alignment = cellCenterCenter
        lista.Enabled = True
    
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================
    Private Sub leerEspeciales(tipoespecial, fecha)
        Dim tabla As String
        Dim DIFE As Double
        Dim total As Double
        
        tabla = "SELECT glosa,monto,id "
        tabla = tabla & "FROM " + clientesistema + "conta.flujo_caja "
        tabla = tabla & "WHERE tipo='" + tipoespecial + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' and empresa='" + empresaflujo + "' "
        Call ConectarControlData(data, Servidor, basedatos, Usuario, password, tabla)
        lista.Rows = 1
        lista.AutoRedraw = False
        total = 0
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                lista.Rows = lista.Rows + 1
                lista.Cell(lista.Rows - 1, 0).text = data.Recordset.Fields(2)
                lista.Cell(lista.Rows - 1, 1).text = data.Recordset.Fields(0)
                lista.Cell(lista.Rows - 1, 2).text = data.Recordset.Fields(1)
               total = total + data.Recordset.Fields(1)
               
               
                data.Recordset.MoveNext
            Wend
        lista.AutoRedraw = True
        lista.Refresh
        End If
    lbltotal.Caption = Format(total, "###,###,###")
    dato1.text = ""
    dato2.text = ""
    
    End Sub
'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub grabarEspeciales()
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        
        campos(0, 0) = "tipo"
        campos(1, 0) = "fecha"
        campos(2, 0) = "glosa"
        campos(3, 0) = "monto"
        campos(4, 0) = "empresa"
        campos(5, 0) = ""
        campos(0, 1) = tipoflujo
        campos(1, 1) = Format(fechaflujo, "yyyy-mm-dd")
        campos(2, 1) = dato1.text
        campos(3, 1) = dato2.text
        campos(4, 1) = empresaflujo
        
        campos(0, 2) = "flujo_caja"
        
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    
 
    End Sub
'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub modificaEspeciales(id)
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "tipo"
        campos(1, 0) = "fecha"
        campos(2, 0) = "glosa"
        campos(3, 0) = "monto"
        campos(4, 0) = "empresa"
        campos(5, 0) = ""
        campos(0, 1) = tipoflujo
        campos(1, 1) = DATO5.text + "-" + dato4.text + "-" + dato3.text
        campos(2, 1) = dato1.text
        campos(3, 1) = dato2.text
        campos(4, 1) = empresaflujo
        campos(0, 2) = "flujo_caja"
        condicion = "tipo = '" & tipoflujo & "' AND fecha = '" & Format(fechaflujo, "yyyy-mm-dd") & "' and id='" & id & "' "
        op = 3
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
        modifica = False
    End Sub
'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================
    Private Sub eliminarEspeciales(id)
    
        Dim condicion As String
        Dim campos(1, 3) As String
        Dim op As Integer
        condicion = "tipo = '" & tipoflujo & "' AND fecha = '" & Format(fechaflujo, "yyyy-mm-dd") & "' and id='" + id + "' and empresa='" + empresaflujo + "' "
        op = 4
        campos(0, 2) = "flujo_caja"
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    End Sub

Private Sub dato3_GotFocus()
Call cargatexto(dato3)

End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If dato3.text >= "01" And dato3.text <= "31" Then
dato4.SetFocus

End If

End If

End Sub

Private Sub dato4_GotFocus()
Call cargatexto(dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If dato4.text >= "01" Or dato4.text <= "12" Then
DATO5.SetFocus

End If

End If

End Sub

Private Sub dato5_GotFocus()
Call cargatexto(DATO5)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And DATO5.text > "2000" Then
     Call modificaEspeciales(lista.Cell(lista.ActiveCell.row, 0).text)
     
     Call leerEspeciales(tipoflujo, fechaflujo)
     End If
     
     
End Sub

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        Call CargaGrillaLista(1, 3)
        Call leerEspeciales(tipoflujo, fechaflujo)
        frmLista.Caption = glosaflujo + " " + Format(fechaflujo, "dd-mm-yyyy")
        DATO5.text = Format(fechaflujo, "yyyy")
        dato4.text = Format(fechaflujo, "mm")
        dato3.text = Format(fechaflujo, "dd")
        
        modifica = False
    End Sub
    
    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = RAISED
    End Sub

    Private Sub lista_DblClick()
        dato1.text = lista.Cell(lista.ActiveCell.row, 1).text
        dato2.text = lista.Cell(lista.ActiveCell.row, 2).text
        
        modifica = True
        dato1.SetFocus
    End Sub

    Private Sub lista_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        Select Case KeyCode
            Case 46
                If lista.ActiveCell.row > 0 Then
                    Call eliminarEspeciales(lista.Cell(lista.ActiveCell.row, 0).text)
                    lista.RemoveItem (lista.ActiveCell.row)
                End If
        End Select
    End Sub

    Private Function revisaCodigo() As String
        Dim i As Long
        revisaCodigo = ""
        For i = 1 To lista.Rows - 1
            If lista.Cell(i, 1).text = dato1.text And lista.Cell(i, 3).text = dato2.text Then
                revisaCodigo = lista.Cell(i, 4).text
                Exit For
            End If
        Next i
    End Function

    
    












