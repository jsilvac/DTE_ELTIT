VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form AdminCajas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   8760
   ClientLeft      =   645
   ClientTop       =   1095
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp detalle 
      Height          =   6495
      Left            =   2160
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11456
      BackColor       =   8421504
      Caption         =   "DETALLE CAJA"
      CaptionEstilo3D =   1
      BackColor       =   8421504
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HabilitarArrastre=   -1  'True
      Begin FlexCell.Grid GridDetalle 
         Height          =   6015
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   10610
         BackColor1      =   -2147483629
         BackColor2      =   14737632
         BackColorActiveCellSel=   16761024
         Cols            =   2
         DefaultFontSize =   8.25
         Rows            =   2
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   9720
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   8400
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin XPFrame.FrameXp FrameComandos 
      Height          =   7455
      Left            =   13680
      TabIndex        =   6
      Top             =   8640
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   13150
      BackColor       =   16744576
      Caption         =   "COMANDOS"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   0
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4260
         BackColor       =   16744576
         Caption         =   "EJECUAR COMANDO"
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
         Begin VB.CommandButton CmdComando 
            BackColor       =   &H00FF8080&
            Caption         =   "EJECUTAR COMANDO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2040
            Width           =   3255
         End
         Begin VB.OptionButton comandos 
            BackColor       =   &H00FF8080&
            Caption         =   "HACER PING"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   12
            Top             =   1680
            Width           =   3015
         End
         Begin VB.OptionButton comandos 
            BackColor       =   &H00FF8080&
            Caption         =   "CERRAR SESION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   11
            Top             =   1320
            Width           =   3015
         End
         Begin VB.OptionButton comandos 
            BackColor       =   &H00FF8080&
            Caption         =   "REINICIAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   10
            Top             =   960
            Width           =   3015
         End
         Begin VB.OptionButton comandos 
            BackColor       =   &H00FF8080&
            Caption         =   "APAGAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   9
            Top             =   600
            Width           =   3015
         End
         Begin VB.OptionButton comandos 
            BackColor       =   &H00FF8080&
            Caption         =   "ENCENDER"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Width           =   3015
         End
      End
   End
   Begin XPFrame.FrameXp Cajas 
      Height          =   7455
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13150
      BackColor       =   16744576
      Caption         =   "CAJAS"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin FlexCell.Grid GridPrincipal 
         Height          =   7095
         Left            =   -120
         TabIndex        =   5
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   12515
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   2
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   1720
      BackColor       =   16744576
      Caption         =   "SERVIDOR LOCAL"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CheckBox VentasLocales 
         BackColor       =   &H00FF8080&
         Caption         =   "BUSCAR VENTAS LocalHost"
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
         Left            =   10560
         TabIndex        =   20
         ToolTipText     =   "La Busqueda se demorará mas"
         Top             =   480
         Width           =   3135
      End
      Begin VB.OptionButton OpTodosLocales 
         BackColor       =   &H00FF8080&
         Caption         =   "TODOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton OpLocalActivo 
         BackColor       =   &H00FF8080&
         Caption         =   "LOCAL ACTIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton CmdConecar 
         BackColor       =   &H00FF8080&
         Caption         =   "CONECTAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TxtServidor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label EstadoConexion 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado de la conexion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SERVIDOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin XPFrame.FrameXp FrameXp6 
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      BackColor       =   16761024
      Caption         =   "Seleccione el Local a Procesar"
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
      Alignment       =   1
      Begin VB.ComboBox combolocal 
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   4485
      End
   End
End
Attribute VB_Name = "AdminCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LocalHostCaja As New rdoConnection
Dim ServidorCaja As New rdoConnection
Private Sub CmdConecar_Click()

GridPrincipal.Rows = 1
GridPrincipal.Refresh
If VentasLocales.Value = 1 Then
    Call CARGAGRILLA(GridPrincipal, 1, 15)
Else
        Call CARGAGRILLA(GridPrincipal, 1, 14)
End If
If OpLocalActivo = True Then
TxtServidor.text = ServidorVentas(empresaActiva)
    Call ConectarServerCaja(TxtServidor, "root", "123")
    Call LeerConfiguracion(GridPrincipal, empresaActiva)
End If

If OpTodosLocales = True Then
For Q = 0 To (combolocal.ListCount - 2)
    TxtServidor.text = ServidorVentas(Mid(combolocal.List(Q), 1, 2))
    Call ConectarServerCaja(TxtServidor, "root", "123")
    Call LeerConfiguracion(GridPrincipal, Mid(combolocal.List(Q), 1, 2))
Next
End If
End Sub
Private Sub Form_Load()
LEErlocales
TxtServidor.text = ServidorVentas(empresaActiva)
Me.Caption = "ADMINISTRACION DE CAJAS " & leerNombreEmpresa(empresaActiva)
Call CARGAGRILLA(GridPrincipal, 1, 13)
End Sub
Private Function ServidorVentas(loc As String) As String
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
Set csql.ActiveConnection = gestion
    csql.sql = "select servidorventas "
    csql.sql = csql.sql & " from g_maestroempresas where codigo = '" & loc & "'"
csql.Execute
        
If csql.RowsAffected > 0 Then
   Set resultados = csql.OpenResultset
    ServidorVentas = resultados(0)
Else
    ServidorVentas = 0
End If
End Function
Private Sub ConectarCajas(loc As String)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
On Error Resume Next
Set csql.ActiveConnection = ventas
    csql.sql = "select numero,nombreequipo,ipprimaria,ipsecundaria from"
    csql.sql = csql.sql & " sv_maestrodecajas where local = '" & loc & "'"
    csql.sql = csql.sql & " order by numero asc"
csql.Execute
'----
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    progreso.Visible = True
    progreso.Max = csql.RowsAffected
    Dim r As Long
    While Not resultados.EOF
With gridcajas
   .Rows = .Rows + 1
   r = .Rows - 1
    progreso.Value = r

    .Cell(r, 1).text = resultados(0)
    .Cell(r, 2).text = resultados(1)
    .Cell(r, 3).text = resultados(2)
    .Cell(r, 4).text = resultados(3)
    .Cell(r, 5).text = resultados(4)
End With
resultados.MoveNext
Wend
resultados.Close
progreso.Visible = False
End If
End Sub
Private Sub CARGAGRILLA(gridcajas As Grid, ByVal row As Integer, ByVal col As Integer)
    Dim formatogrilla(10, 20) As String
    Dim i As Integer
        Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "LOCAL"
    formatogrilla(1, 2) = "CAJA"
    formatogrilla(1, 3) = "IP CAJA"
    formatogrilla(1, 4) = "SERVIDOR"
    formatogrilla(1, 5) = "USUARIO SQL"
    formatogrilla(1, 6) = "Imp Boleta."
    formatogrilla(1, 7) = "Impresora Factura"
    formatogrilla(1, 8) = "Impresora Nota C."
    formatogrilla(1, 9) = "Impresora Pagos"
    formatogrilla(1, 10) = "Sincroniza Local"
    formatogrilla(1, 11) = "Ruta Actualizacion"
    formatogrilla(1, 12) = "VERSION PROGRAMA"
    formatogrilla(1, 13) = "ESTADO"
    formatogrilla(1, 14) = "VENTAS MODO LOCAL"
    
    Rem ANCHO DE LAS CELDAS
        formatogrilla(8, 1) = "5"
        formatogrilla(8, 2) = "4"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "6"
        formatogrilla(8, 6) = "6"
        formatogrilla(8, 7) = "6"
        formatogrilla(8, 8) = "6"
        formatogrilla(8, 9) = "6"
        formatogrilla(8, 10) = "6"
        formatogrilla(8, 11) = "14"
        formatogrilla(8, 12) = "10"
        formatogrilla(8, 13) = "8"
        formatogrilla(8, 14) = "6"
    
    Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "15"
        formatogrilla(2, 4) = "15"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "20"
        formatogrilla(2, 7) = "9"
      
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "C"
        formatogrilla(3, 2) = "C"
        formatogrilla(3, 3) = "C"
        formatogrilla(3, 4) = "C"
        formatogrilla(3, 5) = "C"
        formatogrilla(3, 6) = "C"
        formatogrilla(3, 7) = "C"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
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
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"
        formatogrilla(5, 13) = "TRUE"
        formatogrilla(5, 14) = "TRUE"
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        Rem ANCHO

        gridcajas.Cols = col
        gridcajas.Rows = row
        gridcajas.Range(0, 0, gridcajas.Rows - 1, gridcajas.Cols - 1).Borders(cellEdgeBottom) = cellNone
        gridcajas.Range(0, 0, gridcajas.Rows - 1, gridcajas.Cols - 1).Borders(cellEdgeTop) = cellNone
        gridcajas.Range(0, 0, gridcajas.Rows - 1, gridcajas.Cols - 1).Borders(cellInsideVertical) = cellNone
        gridcajas.AllowUserResizing = False
        gridcajas.DisplayFocusRect = False
        gridcajas.ExtendLastCol = True
        gridcajas.BoldFixedCell = False
        gridcajas.DrawMode = cellOwnerDraw
        gridcajas.Appearance = Flat
        gridcajas.ScrollBarStyle = Flat
        gridcajas.FixedRowColStyle = Flat
        gridcajas.BackColorFixed = RGB(90, 158, 214)
        gridcajas.BackColorFixedSel = RGB(110, 180, 230)
        gridcajas.BackColorBkg = RGB(90, 158, 214)
        gridcajas.BackColorScrollBar = RGB(231, 235, 247)
        gridcajas.BackColor1 = RGB(231, 235, 247)
        gridcajas.BackColor2 = RGB(239, 243, 255)
        gridcajas.GridColor = RGB(148, 190, 231)
        
        gridcajas.Column(0).Width = 0
        gridcajas.RowHeight(0) = gridcajas.DefaultRowHeight * 1.75
        gridcajas.Range(0, 1, 0, gridcajas.Cols - 1).WrapText = True
        
        For i = 1 To gridcajas.Cols - 1
        gridcajas.FrozenCols = 2
        gridcajas.FrozenRows = 1
            gridcajas.Cell(0, i).text = formatogrilla(1, i)
            gridcajas.Column(i).Width = Val(formatogrilla(8, i)) * (gridcajas.Cell(0, i).Font.Size + 1.25)
            gridcajas.Column(i).MaxLength = Val(formatogrilla(2, i))
            gridcajas.Column(i).FormatString = formatogrilla(4, i)
            gridcajas.Column(i).Locked = formatogrilla(5, i)
            gridcajas.Column(i).Alignment = cellRightCenter
        Next i
        gridcajas.Range(0, 1, 0, gridcajas.Cols - 1).Alignment = cellCenterCenter
        gridcajas.Range(0, 1, 0, gridcajas.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
End Sub
Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas "
        csql.sql = csql.sql + "  WHERE CODIGO < '50' ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                combolocal.AddItem (resultados(0))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
                
        combolocal.text = combolocal.List(CDbl(empresaActiva))
        End If
End Sub
Public Sub LeerConfiguracion(gridcajas As Grid, ByRef loc)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
'On Error Resume Next
Set csql.ActiveConnection = ServidorCaja

    csql.sql = "select empresa,caja,ipcaja,servidor,usuario,impresoraboleta,impresorafactura"
    csql.sql = csql.sql & " ,impresoranotacredito,impresorapago,sincronizalocal,ruta"
    csql.sql = csql.sql & " ,fechaversion from configura_caja where empresa = '" & loc & "' order by caja asc"
    
csql.Execute
'----
If csql.RowsAffected > 0 Then
On Error Resume Next
    Set resultados = csql.OpenResultset
    progreso.Visible = True
    progreso.Max = csql.RowsAffected
    Dim r As Long
    While Not resultados.EOF
With gridcajas
   .Rows = .Rows + 1
   r = .Rows - 1
    progreso.Value = r - 1
    .Cell(r, 1).text = resultados(0) '-local
    .Cell(r, 2).text = resultados(1) '-caja
    .Cell(r, 3).text = resultados(2) '-ip caja
    .Cell(r, 4).text = resultados(3) '-servidor
    .Cell(r, 5).text = resultados(4) '-usuario sql
    .Cell(r, 6).text = resultados(5) '-imp boleta
    .Cell(r, 7).text = resultados(6) '-imp factura
    .Cell(r, 8).text = resultados(7) '-imp nc
    .Cell(r, 9).text = resultados(8) '-pago
  If resultados(9) = "N" Then .Cell(r, 10).BackColor = vbRed
    .Cell(r, 10).text = resultados(9) '-sincroniza local
    .Cell(r, 11).text = resultados(10) '-ruta act
    .Cell(r, 12).text = resultados(11) '-FECHA VERSION
  If VERIFICAPING(resultados(2)) = True Then
    .Cell(r, 13).BackColor = vbRed
    .Cell(r, 13).text = "ON LINE"
    .Cell(r, 13).BackColor = vbGreen
  Else
    .Cell(r, 13).text = "OFF LINE"
    .Cell(r, 13).BackColor = vbRed
  End If
  If VentasLocales.Value = 1 And .Cell(r, 10).text = "S" Then .Cell(r, 14).text = BuscarVentasLocales(.Cell(r, 3).text)
    End With
resultados.MoveNext
Wend
resultados.Close
progreso.Visible = False
End If
End Sub
Sub ConectarServerCaja(ByVal servidor As String, ByVal usuariodb As String, ByVal passworddb As String)
     TxtServidor = servidor
     If VERIFICAPING(servidor) = True Then
     EstadoConexion.BackColor = vbGreen
     On Error GoTo error
        Dim cadena_conexion As String
        Dim bd As String
        bd = cliente_sql & "ventas"
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & servidor & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set ServidorCaja = New rdoConnection
        ServidorCaja.Connect = cadena_conexion
        ServidorCaja.CursorDriver = rdUseServer
        ServidorCaja.EstablishConnection
    Else
        GoTo error:
    End If
    Exit Sub
error:
EstadoConexion.BackColor = vbRed
MsgBox "NO SE PUDO CONECTAR AL SERVIDOR : " & servidor & " ", vbInformation, "ATENCION"
    Exit Sub
End Sub
Sub ConectarLocalHostCaja(ByVal IpCaja As String, ByVal usuariodb As String, ByVal passworddb As String)
        Dim cadena_conexion As String
        Dim bd As String
        bd = cliente_sql & "sincroniza"
        cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & IpCaja & "; DATABASE=" & bd & "; PWD=" & passworddb & "; UID=" & usuariodb & "; OPTION=3"
        Set LocalHostCaja = New rdoConnection
        LocalHostCaja.Connect = cadena_conexion
        LocalHostCaja.CursorDriver = rdUseServer
        LocalHostCaja.EstablishConnection
End Sub
Public Function BuscarVentasLocales(ByRef IpCaja) As String
On Error GoTo error
Call ConectarLocalHostCaja(IpCaja, "root", "123")
    Dim resultados1 As rdoResultset
    Dim cSql1 As New rdoQuery
Set cSql1.ActiveConnection = LocalHostCaja
        cSql1.sql = "SELECT count(id) FROM sincronizador "
        cSql1.sql = cSql1.sql & "where  fecha = '0000-00-00'  and consulta like '%" & clientesistema & "ventas.sv_documento_cabeza%'"
        cSql1.Execute
        Set resultados1 = cSql1.OpenResultset
    If cSql1.RowsAffected > 0 Then
            BuscarVentasLocales = resultados1(0) & " Rows" 'cSql1.RowsAffected & " Rows "
    Else
            BuscarVentasLocales = "0 Rows"
    End If
1:
        Set resultados1 = Nothing
        cSql1.Close
        LocalHostCaja.Close
Exit Function
error:
BuscarVentasLocales = "error en la tabla"
GoTo 1
End Function
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub GridPrincipal_DblClick()
detalle.Visible = True
detalle.Caption = "DETALLE CAJA :" & GridPrincipal.Cell(GridPrincipal.ActiveCell.row, 2).text & " DEL LOCAL " & GridPrincipal.Cell(GridPrincipal.ActiveCell.row, 1).text
Call GRILLADETALLE(empresaActiva, GridPrincipal.ActiveCell.row)
End Sub

Private Sub Label2_Click()
detalle.Visible = False
End Sub

Private Sub GRILLADETALLE(ByRef loc, caja As String)
Dim resultados As rdoResultset
Dim csql As New rdoQuery
With GridDetalle
.Rows = 32
.Cols = 3
.Column(1).Locked = True
.Column(2).Locked = True
.Column(1).Width = 200
.Column(2).Width = 200
.RowHeight(0) = 0

End With

Set csql.ActiveConnection = ServidorCaja
 csql.sql = "select * from configura_caja where empresa = '" & loc
 If caja < 10 Then
 csql.sql = csql.sql & "' AND CAJA = '0" & caja & "'"
 Else
 csql.sql = csql.sql & "' AND CAJA = '" & caja & " '"
 End If
 csql.Execute
Set resultados = csql.OpenResultset
If csql.RowsAffected > 0 Then
 With GridDetalle
    

.Cell(1, 1).text = "CAJA"
.Cell(2, 1).text = "IpCaja"
.Cell(3, 1).text = "empresa"
.Cell(4, 1).text = "servidor"
.Cell(5, 1).text = "usuario"
.Cell(6, 1).text = "password"
.Cell(7, 1).text = "basedatos"
.Cell(8, 1).text = "baseVentas"
.Cell(9, 1).text = "iva"
.Cell(10, 1).text = "com"
.Cell(11, 1).text = "estado"
.Cell(12, 1).text = "Ruta"
.Cell(13, 1).text = "factura"
.Cell(14, 1).text = "balanza"
.Cell(15, 1).text = "bodega"
.Cell(16, 1).text = "IMPRESORABOLETA"
.Cell(17, 1).text = "impresorafactura"
.Cell(18, 1).text = "impresoranotacredito"
.Cell(19, 1).text = "IMPRESORAPAGO"
.Cell(20, 1).text = "comfiscal"
.Cell(21, 1).text = "facturaelectronica"
.Cell(22, 1).text = "notadebitoelectronica"
.Cell(23, 1).text = "notacreditoelectronica"
.Cell(24, 1).text = "pos300"
.Cell(25, 1).text = "consultacheque"
.Cell(26, 1).text = "temporizadorcheques"
.Cell(27, 1).text = "baseteso"
.Cell(28, 1).text = "cajerovendedor"
.Cell(29, 1).text = "sincronizalocal"
.Cell(30, 1).text = "fechaversion"
.Cell(31, 1).text = "datosporsincrizar"
   
   For n = 0 To 30
   
      .Cell(n + 1, 2).text = resultados(n)
    Next
    End With

 End If
End Sub
