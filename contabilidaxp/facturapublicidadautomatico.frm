VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form publi0011 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PUBLICIDAD X FACTURAR"
   ClientHeight    =   9900
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   660
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   12000
      TabIndex        =   6
      Top             =   360
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   8760
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   15452
      BackColor       =   16761024
      Caption         =   "Listado de Publicidad por Facturar"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ColorBarraArriba=   4194304
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
      Begin XPFrame.FrameXp frmporce 
         Height          =   2412
         Left            =   4200
         TabIndex        =   13
         Top             =   2160
         Visible         =   0   'False
         Width           =   7692
         _ExtentX        =   13573
         _ExtentY        =   4260
         BackColor       =   16744576
         Caption         =   "Porcentaje"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton Command3 
            Caption         =   "GRABAR"
            Height          =   372
            Left            =   3240
            TabIndex        =   15
            Top             =   1800
            Width           =   1212
         End
         Begin VB.TextBox TXT_APORTE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   28.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   852
            Left            =   5640
            MaxLength       =   5
            TabIndex        =   14
            Text            =   "0"
            Top             =   840
            Width           =   1692
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FF8080&
            Caption         =   "PORCENTAJE"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   852
            Left            =   360
            TabIndex        =   17
            Top             =   840
            Width           =   4932
         End
         Begin VB.Label LBLEMPRESA 
            BackColor       =   &H00000000&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   492
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   7092
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "GENERAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   8160
         Width           =   2220
      End
      Begin VB.CommandButton CmdFacturar 
         BackColor       =   &H0080FF80&
         Caption         =   "FACTURAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8160
         Width           =   2220
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8160
         Width           =   2220
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7740
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   13653
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   900
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   1588
      BackColor       =   16744576
      Caption         =   "DATOS "
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      ColorBarraArriba=   4194304
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
      Begin VB.OptionButton Option2 
         Caption         =   "No Aportan"
         Height          =   252
         Left            =   2760
         TabIndex        =   12
         Top             =   600
         Width           =   1932
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aportan"
         Height          =   252
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   1932
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "DOBLE CLIK EN NUMERO DE FACTURA PARA VER PDF"
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
         Left            =   6000
         TabIndex        =   18
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15210
      TabIndex        =   1
      Top             =   9900
      Width           =   15240
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "publi0011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private localfiltro As String

Private MODIFI As Integer

'Private Sub codigo_Click()
'    Call dato1_KeyDown(vbKeyF2, 0)
'End Sub
 Private Sub imprimir()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE DEVOLUCIONES PENDIENTES ")
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.TopMargin = 3
Grid1.PageSetup.LeftMargin = 0.1
Grid1.PageSetup.RightMargin = 0.1
Grid1.PageSetup.BottomMargin = 3
Grid1.PageSetup.FooterMargin = 2
Grid1.PageSetup.BlackAndWhite = True

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid1.PrintPreview
End If
End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellLandscape
    Grid1.PageSetup.PrintTitleRows = 0
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "EMITIDO  :  " & Format(fechasistema, "dd-MM-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    

    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "P?g &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub


 

Private Sub CmdFacturar_Click()
Dim campos(20, 3) As String
Dim op As Double
Dim condicion As String

Dim n As Double
Dim tipo As String
Dim numero As String
Dim fecha As String
Dim vencimiento As String

Dim NETO As Double
Dim iva As Double
Dim EXENTO As Double
Dim total As Double

Dim rut_proveedor As String
Dim nombre_proveedor As String
Dim direccion_prove As String
Dim comuna_prove As String
Dim ciudad_prove As String
Dim giro_prove As String

Dim CRCC As String
Dim itemdte As String
Dim cuentapublicidad As String
cuentapublicidad = leerdatos(conta, "maestroempresas", "cuentapublicidad", "codigoempresa='" + empresaactiva + "' ")
For n = 1 To Grid1.Rows - 1
        If Grid1.Cell(n, 10).text <> "1" Then GoTo no:
        
        rut_proveedor = Grid1.Cell(n, 1).text
        If rut_proveedor = "" Then GoTo no
        
        
        
        Grid1.Cell(n, 8).text = PublicidadFacturada("FV", "98", CONFI_EMPRESAFAE, "", Format(fechasistema, "yyyy-mm-dd"), rut_proveedor, Grid1.Cell(n, 7).text)
        If Grid1.Cell(n, 8).text <> "" Then GoTo no
        
        
        
        total = Grid1.Cell(n, 7).text
        If total < 1 Then GoTo no
        
            NETO = Round(total / 1.19)
            iva = total - NETO
            EXENTO = 0
            
            Grid1.Range(n, 1, n, Grid1.Cols - 1).FontBold = True
            Grid1.Cell(n, 1).EnsureVisible
        
        
            campos(0, 0) = "rut"
            campos(1, 0) = "nombre"
            campos(2, 0) = "direccion"
            campos(3, 0) = "comuna"
            campos(4, 0) = "ciudad"
            campos(5, 0) = "giro"
            campos(6, 0) = ""
            campos(0, 2) = "cuentascorrientes"
            condicion = "tipo='" & cuentapublicidad & "' and rut='" & rut_proveedor & "' and a?o='" & Format(fechasistema, "yyyy") & "'"
            op = 5
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
            If sqlconta.status = 4 Then
                        MsgBox "CUENTA CORRIENTE PUBLICIDAD NO CREADA"
                        maestro02.dato1.Enabled = True
                        maestro02.dato2.Enabled = True
                        maestro02.DV.Caption = True
                        maestro02.dato1.text = cuentapublicidad
                        maestro02.dato2.text = rut_proveedor
                        maestro02.DV.Caption = Right(rut_proveedor, 1)
                        cierrect = "S"
                        maestro02.Show 1
                        GoTo no:
            End If
            
            nombre_proveedor = sqlconta.response(1, 3)
            direccion_prove = sqlconta.response(2, 3)
            comuna_prove = sqlconta.response(3, 3)
            ciudad_prove = sqlconta.response(4, 3)
            giro_prove = sqlconta.response(5, 3)
            'CREA EN MAESTRO DE CLIENTES
            If comuna_prove = "" Then
                MsgBox "SIN COMUNA EN CUENTA CORRIENTE"
                        maestro02.dato1.Enabled = True
                        maestro02.dato2.Enabled = True
                        maestro02.DV.Caption = True
                        maestro02.dato1.text = cuentapublicidad
                        maestro02.dato2.text = rut_proveedor
                        maestro02.DV.Caption = Right(rut_proveedor, 1)
                        cierrect = "S"
                        maestro02.Show 1
                        
                        GoTo no:
            End If
            If ciudad_prove = "" Then
                MsgBox "SIN CIUDAD EN CUENTA CORRIENTE"
                        maestro02.dato1.Enabled = True
                        maestro02.dato2.Enabled = True
                        maestro02.DV.Caption = True
                        maestro02.dato1.text = cuentapublicidad
                        maestro02.dato2.text = rut_proveedor
                        maestro02.DV.Caption = Right(rut_proveedor, 1)
                        cierrect = "S"
                        maestro02.Show 1
                        
                        GoTo no:
            End If
            
            If ciudad_prove = "" Then
                MsgBox "SIN GIRO EN CUENTA CORRIENTE"
                        maestro02.dato1.Enabled = True
                        maestro02.dato2.Enabled = True
                        maestro02.DV.Caption = True
                        maestro02.dato1.text = cuentapublicidad
                        maestro02.dato2.text = rut_proveedor
                        maestro02.DV.Caption = Right(rut_proveedor, 1)
                        cierrect = "S"
                        maestro02.Show 1
                        
                        GoTo no:
            End If
            Call crearcliente(rut_proveedor, "0", nombre_proveedor, direccion_prove, comuna_prove, ciudad_prove, giro_prove)
           
        '    fecha = Format(fechasistema, "yyyy-mm-dd")
            fecha = Format(Date, "yyyy-mm-dd")
'           FECHA = Format(PublicidadFacturadaFecha("FV", "98", CONFI_EMPRESAFAE, "", Format(fechasistema, "yyyy-mm-dd"), rut_proveedor, Grid1.Cell(n, 7).text), "yyyy-mm-dd")

            vencimiento = fecha
            
            
            itemdte = UCase("APORTE PUBLICITARIO " & MonthName(Format(fechasistema, "MM")) & " " & Format(fechasistema, "YYYY"))
            CRCC = "0101"
            tipo = 2
            numero = LeerNumeroDte("FV", "98", CONFI_EMPRESAFAE)
        
            'documento fue generado
              Call grabafacturaElectronica(tipo, numero, fecha, vencimiento, rut_proveedor, NETO, iva, EXENTO, total, CRCC, nombre_proveedor, direccion_prove, comuna_prove, ciudad_prove, giro_prove, itemdte)
            
  
'             Call grabafacturaElectronica(tipo, Grid1.Cell(n, 8).text, FECHA, vencimiento, rut_proveedor, neto, iva, exento, total, CRCC, nombre_proveedor, direccion_prove, comuna_prove, ciudad_prove, giro_prove, itemdte)
             Grid1.Range(n, 1, n, Grid1.Cols - 1).BackColor = vbGreen
             Grid1.Cell(n, 8).text = Format(numero, "0000000000")
             GoTo SIGUIENTE
no:
            'no genero documento por algun error
            Grid1.Range(n, 1, n, Grid1.Cols - 1).BackColor = vbRed
            
SIGUIENTE:
Next n
MsgBox "PROCESO TERMINADO, POR FAVOR ESPERAR A QUE SE GENEREN LOS DOCUMENTOS ELECTRONICOS."
End Sub

'Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then Call ayudactacte(dato2)
'    Call flechas(dato1, dato4, KeyCode)
'End Sub
 

Private Sub Command1_Click()
imprimir

End Sub

Private Sub COMMAND2_Click()
Call LEERGUIAS
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF2 Then Call ayudactacte(dato3)

End Sub
'
'Private Sub dato3_KeyPress(KeyAscii As Integer)
'KeyAscii = esNumero(KeyAscii)
'If KeyAscii = 13 Then
'Call ceros(dato3)
'DV.Caption = rut(dato3)
'lblnombreproveedor.Caption = leerdatos(contadb, "cuentascorrientes", "nombre", "tipo='" + CUENTAPROVEEDOR + "' and rut='" + dato3.text + DV.Caption + "' ")
'If lblnombreproveedor.Caption = "" Then
'dato3.SetFocus
'Else
'LEERGUIAS
'
'
'End If
'
'
'
'End If
'
'End Sub
 
Private Sub Command3_Click()

Call grabar_APORTE(Grid1.Cell(Grid1.ActiveCell.row, 1).text, Grid1.Cell(Grid1.ActiveCell.row, 2).text, TXT_APORTE.text)
frmporce.Visible = False

Grid1.Cell(Grid1.ActiveCell.row, 6).text = Format(TXT_APORTE.text, "##.#0")
If TXT_APORTE.text > 0 And TXT_APORTE.text < 30 Then
Grid1.Cell(Grid1.ActiveCell.row, 7).text = CDbl(Grid1.Cell(Grid1.ActiveCell.row, 5).text) * (TXT_APORTE.text / 100)
Else
Grid1.Cell(Grid1.ActiveCell.row, 7).text = "0"
End If

Grid1.Refresh
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

    Call Conectar_BD
    sc = 0
  
Call CARGAPERMISO(Me.Name)
 
 CARGAGRILLADETALLE


End Sub




Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub


'Sub ayudactacte(ByRef caja As TextBox)
'    Dim CAMPOS As Variant
'    Dim cfijo As Variant
'    Dim largo As Variant
'    CAMPOS = Array("rut", "nombre")
'    largo = Array("12n", "40s")
'    cfijo = "tipo='" & CUENTAPROVEEDOR & "' and a?o='" + Format(fechasistema, "yyyy") + "'"
'    cabezas = Array("rut", "nombre")
'    mensajeAyuda = "Ayuda Cuentas Corrientes"
'
'    Call cargaAyudaT(servidor, basebus, usuario, password, "cuentascorrientes", pivote, CAMPOS, cfijo, largo, 2)
'
'    If Val(pivote.text) = 0 Then dato3.SetFocus: GoTo no
'    dato3.text = Mid(pivote.text, 1, 9)
'    dv.Caption = Mid(pivote.text, 10, 1)
'    caja.Enabled = True
'    caja.SetFocus
'no:
'
'End Sub
'
' Sub ayudactacte(ByRef caja As TextBox)
'    Dim CAMPOS As Variant
'    Dim cfijo As Variant
'    Dim largo As Variant
'    CAMPOS = Array("cc.rut", "cc.nombre")
'    largo = Array("12n", "40s")
'    cfijo = "cc.tipo='" & CUENTAPROVEEDOR & "' and cc.a?o='" + Format(fechasistema, "yyyy") + "' and cp.fechatermino >='" & Format(fechasistema, "yyyy-mm-dd") & "'"
'    cabezas = Array("rut", "nombre")
'    mensajeAyuda = "Ayuda Cuentas Corrientes"
'
'    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes as cc inner join contratopublicidad as cp on cc.rut=cp.rut", pivote, CAMPOS, cfijo, largo, 2)
'
'    If Val(pivote.text) = 0 Then dato3.SetFocus: GoTo no
'    dato3.text = Mid(pivote.text, 1, 9)
'    DV.Caption = Mid(pivote.text, 10, 1)
'    caja.Enabled = True
'    caja.SetFocus
'no:
'
'End Sub
Sub CARGAGRILLADETALLE()
    Dim formatogrilla2(50, 50)
    formatogrilla2(1, 1) = "RUT"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "COMPRA"
    formatogrilla2(1, 4) = "N.CREDITO"
    formatogrilla2(1, 5) = "LIQUIDO "
    formatogrilla2(1, 6) = "RAPEL"
    formatogrilla2(1, 7) = "FACTURAR C/IVA"
    formatogrilla2(1, 8) = "FACTURA"
    formatogrilla2(1, 9) = "NO APORTA"
    formatogrilla2(1, 10) = "REVISADO"
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "9"
    formatogrilla2(2, 2) = "35"
    formatogrilla2(2, 3) = "11"
    formatogrilla2(2, 4) = "11"
    formatogrilla2(2, 5) = "11"
    formatogrilla2(2, 6) = "8"
    formatogrilla2(2, 7) = "10"
    formatogrilla2(2, 8) = "10"
    formatogrilla2(2, 9) = "8"
    formatogrilla2(2, 10) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    formatogrilla2(3, 8) = "N"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
      
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 3) = " ###,###,##0"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
    formatogrilla2(4, 6) = " ##0.00"
    formatogrilla2(4, 7) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "FALSE"
    
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 11
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
'    Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    Grid1.Column(9).CellType = cellCheckBox
    Grid1.Column(10).CellType = cellCheckBox
    
    
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 8
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
    Next k
 
 
    End Sub
 
Sub verdetalle(loc, numero)
'Dim cSql As New rdoQuery
'Dim resultados As rdoResultset
'Dim tipo As String
'tipo = "DM"
'
'Set cSql.ActiveConnection = contadb
'
'cSql.sql = "select linea,codigo,descripcion,cantidad,uxc,unidades,precio,descuento,total "
'cSql.sql = cSql.sql & "from " & clientesistema & "gestion" & leerrubro(dato1.text) & ".l_movimientos_detalle_" & loc & " where tipo='" & tipo & "' and numero='" & numero & "' order by linea"
'cSql.Execute
'
'If cSql.RowsAffected > 0 Then
'    Grid1.Rows = cSql.RowsAffected + 1
'    Set resultados = cSql.OpenResultset
'
'    While Not resultados.EOF
'        Grid1.Cell(resultados(0), 1).text = resultados(1)
'        Grid1.Cell(resultados(0), 2).text = resultados(2)
'        Grid1.Cell(resultados(0), 3).text = resultados(3)
'        Grid1.Cell(resultados(0), 4).text = resultados(4)
'        Grid1.Cell(resultados(0), 5).text = resultados(5)
'        Grid1.Cell(resultados(0), 6).text = resultados(6)
'        Grid1.Cell(resultados(0), 7).text = resultados(7)
'        Grid1.Cell(resultados(0), 8).text = resultados(8)
'        resultados.MoveNext
'    Wend
'End If
'
'cSql.Close
'Set cSql = Nothing
'Set resultados = Nothing
 
End Sub
Function leerrubro(loc) As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select rubro from " & clientesistema & "gestion.g_maestroempresas where "
    csql.sql = csql.sql & "codigo='" & loc & "' "
    csql.Execute
    
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    leerrubro = resultado(0)
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function


Sub LEERGUIAS()
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
Dim rutpaso As String
Dim totales(2) As Double
Dim totales2(2) As Double
Dim cuentapublicidad As String
Dim porce As Double
Dim TOTALAPORTE As Double
Dim TOTALCOMPRA As Double
Dim aporte As Double

Call CARGAGRILLADETALLE
tipo = "DM"
cuentapublicidad = leerdatos(conta, "maestroempresas", "cuentapublicidad", "codigoempresa='" + empresaactiva + "' ")

Set csql.ActiveConnection = contadb

csql.sql = "SELECT fc.rut,ifnull((select nombre from cuentascorrientes as cc where cc.rut=fc.rut and a?o='" + Format(fechasistema, "yyyy") + "' and tipo='23100026'),'** no creado **') as nombre,"
csql.sql = csql.sql + "SUM(IF(fc.tipo='1' OR fc.tipo='4',neto,0)) ,SUM(IF(fc.tipo='3' OR fc.tipo='6',neto,0)), "
csql.sql = csql.sql + "ifnull((select dato3 from eltit_conta.rapel as rp where rp.dato1=fc.rut),0) as rapel, "
csql.sql = csql.sql + "ifnull((select noaporta from cuentascorrientes as cc where cc.rut=fc.rut and a?o='" + Format(fechasistema, "yyyy") + "' and tipo='23100026'),'0') as noaporta "
csql.sql = csql.sql + " From facturasdecompras as fc "
csql.sql = csql.sql + " WHERE fecha like '" + Format(fechasistema, "yyyy-mm") + "%' "
csql.sql = csql.sql + " GROUP BY fc.rut "
If Option1.Value = True Then
csql.sql = csql.sql + " having noaporta<>'1' "
Else
csql.sql = csql.sql + " having noaporta='1' "
End If

csql.sql = csql.sql + " order by nombre "

csql.Execute
Grid1.Rows = 1
TOTALCOMPRA = 0
TOTALAPORTE = 0
aporte = 0
'FACTURA 8
Grid1.AutoRedraw = False
If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
  
  
    While Not resultados.EOF
        Grid1.Rows = Grid1.Rows + 1
        aporte = 0
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
        porce = resultados(4) * 100
'        If porce > 0 Then Stop
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultados(2), "###,###,###,##0")
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultados(3), "###,###,###,##0")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultados(2) - resultados(3), "###,###,###,##0")
        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(porce, "#0.00")
        
        aporte = (resultados(2) - resultados(3)) * resultados(4)
        aporte = Round(aporte * 1.19, 0)
        
        Grid1.Cell(Grid1.Rows - 1, 7).text = Format(aporte, "###,###,###,##0")
        
        Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(5)
        TOTALCOMPRA = TOTALCOMPRA + (resultados(2) - resultados(3))
        TOTALAPORTE = TOTALAPORTE + Round(((resultados(2) - resultados(3)) * resultados(4)), 0)
        
        If Grid1.Cell(Grid1.Rows - 1, 7).text <> 0 Then Grid1.Cell(Grid1.Rows - 1, 8).text = PublicidadFacturada("FV", "98", CONFI_EMPRESAFAE, "", Format(fechasistema, "yyyy-mm-dd"), resultados(0), Grid1.Cell(Grid1.Rows - 1, 7).text)
        Grid1.Cell(Grid1.Rows - 1, 10).text = leerrevisado(empresaactiva, resultados(0), Format(fechasistema, "yyyy-mm"))
        resultados.MoveNext
    Wend
        
        
End If
Grid1.Rows = Grid1.Rows + 1
Grid1.Cell(Grid1.Rows - 1, 5).text = Format(TOTALCOMPRA, "###,###,###,##0")
Grid1.Cell(Grid1.Rows - 1, 7).text = Format(TOTALAPORTE, "###,###,###,##0")
If TOTALCOMPRA <> 0 Then
Grid1.Cell(Grid1.Rows - 1, 6).text = Format(((TOTALAPORTE / TOTALCOMPRA) * 100) * 1.19, "#0.00")
End If
               
Grid1.AutoRedraw = True
Grid1.Refresh
csql.Close
Set csql = Nothing
Set resultados = Nothing
 
End Sub

Function leerrevisado(empr, rutprove, periodo) As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = conta
    csql.sql = "select revisado from rapel_revisado where "
    csql.sql = csql.sql & "empresa='" & empr & "' and rut='" & rutprove & "' and periodo='" & periodo & "'"
    csql.Execute
        leerrevisado = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerrevisado = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
End Function

Sub grabarrevisado(empr, rutprove, periodo, revisado)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim csql2 As New rdoQuery
    
    Set csql.ActiveConnection = conta
    Set csql2.ActiveConnection = conta
    
    csql.sql = "select revisado from rapel_revisado where "
    csql.sql = csql.sql & "empresa='" & empr & "' and rut='" & rutprove & "' and periodo='" & periodo & "'"
    csql.Execute
        
    If csql.RowsAffected > 0 Then
        csql2.sql = "update rapel_revisado set revisado='" & revisado & "' "
        csql2.sql = csql2.sql & " where empresa='" & empr & "' and rut='" & rutprove & "' and periodo='" & periodo & "'"
        csql2.Execute
    Else
        csql2.sql = "insert into  rapel_revisado (empresa,rut,periodo,revisado) "
        csql2.sql = csql2.sql & " value('" & empr & "','" & rutprove & "','" & periodo & "','" & revisado & "') "
        csql2.Execute
    End If
    
    
    csql.Close
    Set csql = Nothing
    
End Sub
Private Sub Grid1_Click()
If Grid1.ActiveCell.col = 9 Then
    If Grid1.Cell(Grid1.ActiveCell.row, 6).text = "0,00" Then
        If Grid1.Cell(Grid1.ActiveCell.row, 9).text = "0" Or Grid1.Cell(Grid1.ActiveCell.row, 9).text = "" Then
         Grid1.Cell(Grid1.ActiveCell.row, 9).text = "1"
        Else
         Grid1.Cell(Grid1.ActiveCell.row, 9).text = "0"
        End If
        
        Call modifica_proveedor("noaporta", Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text, Grid1.Cell(Grid1.ActiveCell.row, 1).text)
    Else
        MsgBox "ESTE PROVEEDOR TIENE APORTES COMPROMETIDOS NO PUEDE BORRAR "
        Grid1.Cell(Grid1.ActiveCell.row, 9).text = "0"
    End If

End If

If Grid1.ActiveCell.col = 6 Then
    TXT_APORTE.text = Grid1.Cell(Grid1.ActiveCell.row, 6).text
    LBLEMPRESA.Caption = Grid1.Cell(Grid1.ActiveCell.row, 2).text
    frmporce.Visible = True
    TXT_APORTE.SetFocus
End If
If Grid1.ActiveCell.col = 10 Then
    Call grabarrevisado(empresaactiva, Grid1.Cell(Grid1.ActiveCell.row, 1).text, Format(fechasistema, "yyyy-mm"), Grid1.Cell(Grid1.ActiveCell.row, 10).text)
End If


End Sub



Private Sub Grid1_DblClick()
    If Grid1.ActiveCell.col = 8 Then
        If Grid1.Cell(Grid1.ActiveCell.row, 8).text <> "" Then
            Call Cargarpdf("33", Grid1.Cell(Grid1.ActiveCell.row, 8).text, Grid1.Cell(Grid1.ActiveCell.row, 1).text, 0)
            Call Sleep(10000)
            Call Cargarpdf("33", Grid1.Cell(Grid1.ActiveCell.row, 8).text, Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1)
        End If
    End If
End Sub

Public Function Cargarpdf(tipo, numero, RUTCLIENTE, hoja) As String
Dim Tama?o As Double
Dim cn As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim mstream As ADODB.Stream
Dim pdfpath, pdfpath1 As String
Dim pdffile As ADODB.Stream

If tipo = "1" Then
    tipo = "33"
End If
If tipo = "4" Then
    tipo = "61"
End If

Dim ImgTemporal As String
ImgTemporal = "C:\tmp_pdf" & hoja & ".pdf"
If ExisteArchivo(ImgTemporal) = True Then Kill ImgTemporal

Set cn = New ADODB.Connection
cn.Open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & clientesistema & "ventas" & ";PWD=" & password & "; UID=" & Usuario & ";OPTION=3"
cn.CursorLocation = adUseClient


Set Rs = New ADODB.Recordset
'Rs.Open " select * from pdf where pdfid='" & txtid.text & "' and pdfname='" & txtname.text & "'", cn, adOpenKeyset, adLockOptimistic
Rs.Open "Select * from " & clientesistema & "fae" & CONFI_EMPRESAFAE & ".sv_dtepdf_" & CONFI_EMPRESAFAE & " where tipo='" & tipo & "' and numero='" & Val(numero) & "' and rut = '" & Val(Mid(RUTCLIENTE, 1, 9)) & Mid(RUTCLIENTE, 10, 1) & "' and cedible='" & hoja & "' limit 0,1 ", cn, adOpenKeyset, adLockOptimistic

If Not Rs.EOF Then
Set pdffile = New ADODB.Stream
pdffile.Type = adTypeBinary
pdffile.Open
If IsNull(Rs.Fields("pdf")) = False Then
pdffile.Write Rs.Fields("pdf").Value
'Dim pdfnme As String
'pdfnme = txtid.text & txtname.text
'pdffile.SaveToFile "" & App.Path & "\reports\" & pdfnme & ".pdf", adSaveCreateOverWrite
pdffile.SaveToFile ImgTemporal, adSaveCreateOverWrite
'pdffile.SaveToFile ImgTemporal, adSaveCreateOverWrite
pdffile.Close
Set pdffile = Nothing
'ShellExecute publi0006.hwnd, "print", ImgTemporal, vbNullString, App.path, 0
ShellExecute Me.hwnd, "open", ImgTemporal, vbNullString, App.path, 0
'Shell "C:\Archivos de programa\Adobe\Reader 10.0\Reader\AcroRd32.exe " & ImgTemporal
'MsgBox "pdf file downloaded"
Else
MsgBox "NO SE HA ENCONTRADO EL ARCHIVO", vbCritical, "ATENCION"
Rs.Close
Set Rs = Nothing
End If
End If
Rem If ExisteArchivo(ImgTemporal) = True Then Kill ImgTemporal
End Function

Private Sub Option1_Click()
COMMAND2_Click

End Sub

Private Sub Option2_Click()
COMMAND2_Click
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Sub modifica_proveedor(campo, monto, rut)
    campos(0, 0) = campo
    campos(1, 0) = ""
    campos(0, 1) = monto
    
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".cuentascorrientes "
    condicion = "rut='" + rut + "' and tipo='23100026' and a?o='" + Format(fechasistema, "yyyy") + "' "
    
    sqlconta.response = campos
    op = 3
        
    Call sqlconta.sqlconta(op, condicion)
    
End Sub



Sub grabafacturaElectronica(tipo, numero, fecha, vencimiento, rutproveedor, _
                            NETO, iva, EXENTO, total, CRCC, NOMBRE, direccion, comuna, ciudad, giro, glosadte)
    Dim netos As Double
    Dim DH As String
    Dim loc As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "total"
    campos(9, 0) = "fechadigitacion"
    campos(10, 0) = "crcc"
    campos(11, 0) = "nombre"
    campos(12, 0) = "direccion"
    campos(13, 0) = "comuna"
    campos(14, 0) = "ciudad"
    campos(15, 0) = "giro"
    campos(16, 0) = "itemdte"
    campos(17, 0) = "caja"
    campos(18, 0) = ""
    
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = fecha
    campos(3, 1) = vencimiento
    campos(4, 1) = rutproveedor
    campos(5, 1) = Replace(NETO, ".", "")
    campos(6, 1) = Replace(iva, ".", "")
    campos(7, 1) = Replace(EXENTO, ".", "")
    campos(8, 1) = Replace(total, ".", "")
    campos(9, 1) = Format(Now, "yyyy-mm-dd")
    campos(10, 1) = CRCC
    campos(11, 1) = NOMBRE
    campos(12, 1) = direccion
    campos(13, 1) = comuna
    campos(14, 1) = ciudad
    campos(15, 1) = giro
    campos(16, 1) = glosadte
    campos(17, 1) = "98"
    condicion = ""
    campos(0, 2) = "facturasdepublicidad"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "glosa"
    campos(4, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = 1
    campos(3, 1) = glosadte
    
    condicion = ""
    campos(0, 2) = "facturasdepublicidad_glosa"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
   loc = CONFI_EMPRESAFAE
   
   tipo = "FV"

    Call grabardte(tipo, numero, 1, fecha, fecha, rutproveedor, "0000000000100", glosadte, 1, total, 0, total, _
                    "", 0, "00", "0", "98", 0, "", "", NETO, iva, EXENTO, loc)
    
    
'    Call grabarcontable("6", numero, fecha, fecha, rutproveedor, NETO, iva, exento, total, fechasistema, "0101", glosadte)
 
'        If documentocreadopubli("FV", "98", CONFI_EMPRESAFAE, numero, Format(fecha, "yyyy-mm-dd"), rutproveedor) = True Then
''           NUMERODOCUMENTO_DTE
'        Call grabarcontable("6", NUMERODOCUMENTO_DTE, fecha, fecha, rutproveedor, NETO, iva, exento, total, fechasistema, "0101", glosadte)
'        Call actualizapubli(Format(fecha, "yyyy-mm-dd"), NUMERODOCUMENTO_DTE, rutproveedor, total, numero)
'       Else
'           Call grabarcontable("6", numero, fecha, fecha, rutproveedor, NETO, iva, exento, total, fechasistema, "0101", glosadte)
'
'        End If
        
End Sub
'Sub grabarcontable()
'    Dim campos(50, 3) As String
'    Dim condicion As String
'    Dim op As Integer
'
'    campos(0, 0) = "tipo"
'    campos(1, 0) = "numero"
'    campos(2, 0) = "fecha"
'    campos(3, 0) = "fechavencimiento"
'    campos(4, 0) = "rut"
'    campos(5, 0) = "neto"
'    campos(6, 0) = "iva"
'    campos(7, 0) = "exento"
'    campos(8, 0) = "total"
'    campos(9, 0) = "fechadigitacion"
'    campos(10, 0) = "crcc"
'    campos(11, 0) = "itemdte"
'    campos(12, 0) = ""
'
'    Rem campos(0, 1) = dato1.text
'
'   If dato1.text = "2" Then
'        campos(0, 1) = "6"
'        campos(1, 1) = Format(txtfolio.text, "0000000000")
'    Else
'        campos(0, 1) = dato1.text
'        campos(1, 1) = dato2.text
'    End If
'
'    campos(2, 1) = DATO5.text + "-" + dato4.text + "-" + dato3.text
'    campos(3, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
'    campos(4, 1) = dato9.text + DV.Caption
'    campos(5, 1) = Replace(dato11.text, ".", "")
'    campos(6, 1) = Replace(dato12.text, ".", "")
'    campos(7, 1) = Replace(dato13.text, ".", "")
'    campos(8, 1) = Replace(total.text, ".", "")
'    campos(9, 1) = fechasistema
'    campos(10, 1) = DATO21.text & DATO22.text
'    campos(11, 1) = txtitemfactura.text
'
'
'
'    condicion = ""
'    campos(0, 2) = "facturasdeventas"
'    op = 2
'    sqlconta.response = campos
'    Set sqlconta.conexion = contadb
'
'
'    Call sqlconta.sqlconta(op, condicion)
'
'Rem GRABADETALLEIMPUESTOS
'  grabardetallefactura
'
'Rem If dato1.text <> "2" Then
'Rem 12456 grabar2
'Rem End If
'
'
'
'End Sub

Sub grabarcontable(tipo, numero, fecha, fechavencimiento, rutcli, NETO, iva, EXENTO, total, fechadigitacion, CRCC, itemdte)
    Dim campos(50, 3) As String
    Dim condicion As String
    Dim op As Integer
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "total"
    campos(9, 0) = "fechadigitacion"
    campos(10, 0) = "crcc"
    campos(11, 0) = "itemdte"
    campos(12, 0) = "caja"
    campos(13, 0) = ""
    
'    campos(0, 1) = dato1.text
    
   
    campos(0, 1) = tipo
    campos(1, 1) = Format(numero, "0000000000")
    
    campos(2, 1) = Format(fecha, "yyyy-mm-dd")
    campos(3, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(4, 1) = rutcli
    campos(5, 1) = Replace(NETO, ".", "")
    campos(6, 1) = Replace(iva, ".", "")
    campos(7, 1) = Replace(EXENTO, ".", "")
    campos(8, 1) = Replace(total, ".", "")
    campos(9, 1) = Format(fechadigitacion, "yyyy-mm-dd")
    campos(10, 1) = CRCC
    campos(11, 1) = itemdte
    campos(12, 1) = "98"
    
    
    condicion = ""
    campos(0, 2) = "facturasdeventas"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
    
   Call grabardetallefactura(tipo, numero, "001", rutcli, "", "", total, "", "0101", "")
'    grabar2
 
 
End Sub
 Sub grabardetallefactura(tipo, numero, LINEA, rutcli, cuentadelmayor, glosa, monto, DH, centrodecosto, rutctacte)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = ""
   
    cuentadelmayor = "35150001"
    glosa = "INGRESOS POR PUBLICIDAD"
    tipo = "6"
    DH = "H"
    
    campos(0, 1) = tipo
    campos(1, 1) = Format(numero, "0000000000")
    campos(2, 1) = Format(LINEA, "000")
    campos(3, 1) = rutcli
    campos(4, 1) = cuentadelmayor
    campos(5, 1) = glosa
    campos(6, 1) = Replace(monto, ".", "")
    campos(7, 1) = DH
    campos(8, 1) = centrodecosto
    campos(9, 1) = ""
    
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
'    Call modificafactura(tipo, Format(numero, "0000000000"), "98")
    
    
End Sub
Sub modificafactura(tipo, numero, caja)
    Dim campos(10, 10) As String
    Dim condicion As String
    
    Dim netos As Double
    Dim DH As String
    campos(0, 0) = "caja"
    campos(1, 0) = ""
    campos(0, 1) = caja
    
    If tipo = "2" Then tipo = "6"
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' "
    campos(0, 2) = "facturasdeventas"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
End Sub

Sub grabardte(tipo, numero, LINEA, fecha, vencimiento, rut, codigo, descripcion, Cantidad, Precio, descuento, total, _
            Vendedor, pcosto, bodega, SUCURSAL, caja, descuentopesos, tipodespacho, despacho, NETO, iva, EXENTO, loc)
      'detalles
      campos(0, 0) = "local"
      campos(1, 0) = "tipo"
      campos(2, 0) = "numero"
      campos(3, 0) = "linea"
      campos(4, 0) = "fecha"
      campos(5, 0) = "rut"
      campos(6, 0) = "codigo"
      campos(7, 0) = "descripcion"
      campos(8, 0) = "cantidad"
      campos(9, 0) = "precio"
      campos(10, 0) = "descuento"
      campos(11, 0) = "total"
      campos(12, 0) = "vendedor"
      campos(13, 0) = "pcosto"
      campos(14, 0) = "bodega"
      campos(15, 0) = "sucursal"
      campos(16, 0) = "caja"
      campos(17, 0) = "descuentopesos"
      campos(18, 0) = "tipodespacho"
      campos(19, 0) = "despachado"
      campos(20, 0) = ""
      
      campos(0, 1) = loc
      campos(1, 1) = tipo
      campos(2, 1) = Format(numero, "0000000000")
      campos(3, 1) = Format(LINEA, "000")
      campos(4, 1) = Format(fecha, "yyyy-mm-dd")
      campos(5, 1) = rut
      campos(6, 1) = codigo
      campos(7, 1) = descripcion
      campos(8, 1) = Cantidad
      campos(9, 1) = Replace(Replace(Precio, ".", ""), ",", ".")
      campos(10, 1) = descuento
      campos(11, 1) = Replace(Replace(total, ".", ""), ",", ".")
      campos(12, 1) = Vendedor
      campos(13, 1) = pcosto
      campos(14, 1) = bodega
      campos(15, 1) = SUCURSAL
      campos(16, 1) = caja
      campos(17, 1) = descuentopesos
      campos(18, 1) = tipodespacho
      campos(19, 1) = 1
      
      campos(0, 2) = clientesistema & "ventas" & loc & ".sv_otros_documento_detalle_" & loc
      condicion = ""
      op = 2
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
          
      
      
      'cabeza
      campos(0, 0) = "local"
      campos(1, 0) = "tipo"
      campos(2, 0) = "numero"
      campos(3, 0) = "fecha"
      campos(4, 0) = "plazo"
      campos(5, 0) = "vencimiento"
      campos(6, 0) = "rut"
      campos(7, 0) = "cajera"
      campos(8, 0) = "notapedido"
      campos(9, 0) = "notaventa"
      campos(10, 0) = "ordencompra"
      campos(11, 0) = "neto"
      campos(12, 0) = "iva"
      campos(13, 0) = "impuestoharina"
      campos(14, 0) = "impuestoila"
      campos(15, 0) = "impuestoespecifico"
      campos(16, 0) = "exento"
      campos(17, 0) = "retencionparcial"
      campos(18, 0) = "retenciontotal"
      campos(19, 0) = "total"
      campos(20, 0) = "abono"
      campos(21, 0) = "pagado"
      campos(22, 0) = "caja"
      campos(23, 0) = "horaventas"
      campos(24, 0) = "subtotal"
      campos(25, 0) = "descuento"
      campos(26, 0) = "foliosii"
      campos(27, 0) = "vendedor"
      campos(28, 0) = "contabilizado"
      campos(29, 0) = "sucursal"
      campos(30, 0) = "glosafactura"
      campos(31, 0) = "fechacreacion"
      campos(32, 0) = ""
      
      
      
      
      campos(0, 1) = loc
      campos(1, 1) = tipo
      campos(2, 1) = Format(numero, "0000000000")
      campos(3, 1) = Format(fecha, "yyyy-mm-dd")
      campos(4, 1) = "000"
      campos(5, 1) = Format(fecha, "yyyy-mm-dd")
      campos(6, 1) = rut
      campos(7, 1) = "000000019"
      campos(8, 1) = "0000000000"
      campos(9, 1) = "0000000000"
      campos(10, 1) = "0000000000"
      campos(11, 1) = Replace(Replace(NETO, ".", ""), ",", ".")
      campos(12, 1) = Replace(Replace(iva, ".", ""), ",", ".")
      campos(13, 1) = "0"
      campos(14, 1) = "0"
      campos(15, 1) = "0"
      campos(16, 1) = Replace(Replace(EXENTO, ".", ""), ",", ".")
      campos(17, 1) = "0"
      campos(18, 1) = "0"
      campos(19, 1) = Replace(Replace(total, ".", ""), ",", ".")
      campos(20, 1) = Replace(Replace(total, ".", ""), ",", ".")
      campos(21, 1) = "S"
      campos(22, 1) = caja
      campos(23, 1) = Time
      campos(24, 1) = Replace(Replace(total, ".", ""), ",", ".")
      campos(25, 1) = "0"
      campos(26, 1) = Format(numero, "0000000000")
      campos(27, 1) = ""
      campos(28, 1) = "E"
      campos(29, 1) = SUCURSAL
      campos(30, 1) = descripcion
      campos(31, 1) = Format(fechasistema, "yyyy-mm-dd")
      
      
      campos(0, 2) = clientesistema & "ventas" & loc & ".sv_otros_documento_cabeza_" & loc
      condicion = ""
      op = 2
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
      
      'pagos
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "lineapago"
        campos(4, 0) = "fecha"
        campos(5, 0) = "tipopago"
        campos(6, 0) = "cuentacorriente"
        campos(7, 0) = "banco"
        campos(8, 0) = "plaza"
        campos(9, 0) = "numerodocumento"
        campos(10, 0) = "monto"
        campos(11, 0) = "vencimiento"
        campos(12, 0) = "rut"
        campos(13, 0) = "glosa"
        campos(14, 0) = "pagoenlazado"
        campos(15, 0) = "localdocumento"
        campos(16, 0) = "foliofiscal"
        campos(17, 0) = "cuotas"
        campos(18, 0) = "montocuotas"
        campos(19, 0) = "rutcredito"
        campos(20, 0) = "primervencimiento"
        campos(21, 0) = "caja"
        campos(22, 0) = "rutadicional"
        campos(23, 0) = ""
        
        campos(0, 1) = loc
        campos(1, 1) = tipo
        campos(2, 1) = Format(numero, "0000000000")
        campos(3, 1) = Format(LINEA, "000")
        campos(4, 1) = Format(fecha, "yyyy-mm-dd")
        campos(5, 1) = "1"
        campos(6, 1) = ""
        campos(7, 1) = ""
        campos(8, 1) = ""
        campos(9, 1) = ""
        campos(10, 1) = Replace(Replace(total, ".", ""), ",", ".")
        campos(11, 1) = Format(fecha, "yyyy-mm-dd")
        campos(12, 1) = rut
        campos(13, 1) = ""
        campos(14, 1) = ""
        campos(15, 1) = ""
        campos(16, 1) = Format(numero, "0000000000")
        campos(17, 1) = ""
        campos(18, 1) = ""
        campos(19, 1) = ""
        campos(20, 1) = ""
        campos(22, 1) = ""
        campos(21, 1) = caja
        
       
        campos(0, 2) = clientesistema & "ventas" & loc & ".sv_otros_documento_pagos_" & loc
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        
                
End Sub





 
Sub crearcliente(RUTCLIENTE, suc, NOMBRE, direccion, comuna, ciudad, giro)
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = contadb
    
            csql.sql = "replace INTO " & clientesistema & "ventas.sv_maestroclientes   "
            csql.sql = csql.sql & "(rut,sucursal,nombre,direccion,comuna,ciudad,giro) "
            csql.sql = csql.sql & "value ('" + RUTCLIENTE + "','" + suc + "','" & NOMBRE & "','" & direccion & "','" & comuna & "','" & ciudad & "','" & giro & "') "
            csql.Execute
            
            csql.Close
            Set csql = Nothing
            
End Sub
Public Function LeerNumeroDte(tipo, caja, loc) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    Set csql.ActiveConnection = contadb

    csql.sql = "select IFNULL(max(numero),0) from " & clientesistema & "ventas" & loc & ".sv_otros_documento_cabeza_" & loc
    csql.sql = csql.sql & " where tipo='FV' AND caja='98' GROUP BY tipo  "
            
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LeerNumeroDte = Format(resultados(0) + 1, "0000000000")
    Else
        LeerNumeroDte = Format(1, "0000000000")
    
    End If
    
End Function


Public Function PublicidadFacturada(tipo, caja, loc, FOLIO, fecha, proveedor, monto) As String
        Dim csql As rdoQuery
        Dim resultados As rdoResultset
        
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = contadb
        
        csql.sql = "select foliosii from " + cliente_sql + "ventas" + loc + ".sv_otros_documento_cabeza_" + loc & ""
        csql.sql = csql.sql & " WHERE tipo='" & tipo & "'"
        csql.sql = csql.sql & " and caja='" & caja & "'"
        csql.sql = csql.sql & " and local='" & loc & "'"
        csql.sql = csql.sql & " AND rut = '" & proveedor & "'"
        csql.sql = csql.sql & " AND fechacreacion like '" & Format(fecha, "yyyy-mm") & "%'"
        csql.sql = csql.sql & " AND total = '" & CDbl(monto) & "'"
        csql.sql = csql.sql & " AND glosafactura<>'' "
        
        csql.Execute
     
        PublicidadFacturada = ""
        
        If csql.RowsAffected > 0 Then
          Set resultados = csql.OpenResultset
            PublicidadFacturada = resultados(0)
             
        End If
        
        Set csql = Nothing
    End Function
    
    Public Function PublicidadFacturadaFecha(tipo, caja, loc, FOLIO, fecha, proveedor, monto) As String
        Dim csql As rdoQuery
        Dim resultados As rdoResultset
        
        Set csql = New rdoQuery
        
        Set csql.ActiveConnection = contadb
        
        csql.sql = "select fecha from " + cliente_sql + "ventas" + loc + ".sv_otros_documento_cabeza_" + loc & ""
        csql.sql = csql.sql & " WHERE tipo='" & tipo & "'"
        csql.sql = csql.sql & " and caja='" & caja & "'"
        csql.sql = csql.sql & " and local='" & loc & "'"
        csql.sql = csql.sql & " AND rut = '" & proveedor & "'"
        csql.sql = csql.sql & " AND fechacreacion like '" & Format(fecha, "yyyy-mm") & "%'"
        csql.sql = csql.sql & " AND total = '" & CDbl(monto) & "'"
        csql.sql = csql.sql & " AND glosafactura<>'' "
        
        csql.Execute
     
        PublicidadFacturadaFecha = Format(fecha, "yyyy-mm-dd")
        
        If csql.RowsAffected > 0 Then
          Set resultados = csql.OpenResultset
            PublicidadFacturadaFecha = resultados(0)
             
        End If
        
        Set csql = Nothing
    End Function

Public Sub existerut(a?o, tipo, rut, empresa)
 
    campos(0, 0) = "a?o"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = ""
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' and a?o='" + a?o + "'  "
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    Else
    
Rem Call grabar(a?o, tipo, rut, NOMBRE, empresa)
    
    End If

    
    End Sub
Public Function nombretraba(rut, empresa) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' "
    campos(0, 2) = clientesistema + "remu" + empresa + ".mt_fijo "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    nombretraba = sqlconta.response(0, 3)
    Else
    nombretraba = ""
    End If
    End Function

Sub grabar(a?o, tipo, rut, NOMBRE, empresa)
    campos(0, 0) = "a?o"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = NOMBRE
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
     Call grabar2(a?o, tipo, rut, empresa)
    
    End Sub
Sub grabar2(a?o, tipo, rut, empresa)
      
    campos(0, 0) = "a?o"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = ""
    
    campos(0, 1) = a?o
    campos(1, 1) = tipo
    campos(2, 1) = rut
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".saldosctacte"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub


Sub grabar_APORTE(rut, nombre2, monto)
      
    campos(0, 0) = "dato1"
    campos(1, 0) = "dato2"
    campos(2, 0) = "dato3"
    campos(3, 0) = ""
    
    campos(0, 1) = rut
    campos(1, 1) = nombre2
    campos(2, 1) = Round(monto / 100, 3)
    
    campos(0, 2) = clientesistema + "conta.rapel "
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    condicion = "dato1='" + rut + "' "
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    
End Sub

Private Sub TXT_APORTE_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)


End Sub
