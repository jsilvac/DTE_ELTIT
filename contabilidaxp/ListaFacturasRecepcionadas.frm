VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0014 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Ordenes Recepcionadas"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   867
   Begin VB.TextBox ORDEN 
      BackColor       =   &H00FFC0C0&
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
      Left            =   10965
      MaxLength       =   10
      TabIndex        =   23
      Top             =   8520
      Width           =   1500
   End
   Begin VB.CommandButton BUSCAR 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Busca Orden"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8520
      Width           =   1320
   End
   Begin VB.TextBox txtfactura 
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
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   20
      Tag             =   "rut"
      Top             =   8520
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "RECEPCIONAR FACTURAS"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9120
      Width           =   2535
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   9840
      TabIndex        =   15
      Top             =   9120
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
         TabIndex        =   17
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.CommandButton EXPORTAR 
      BackColor       =   &H00FF8080&
      Caption         =   "EXPORTAR A EXCEL"
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
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8400
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   14817
      BackColor       =   16761024
      Caption         =   "Listado Facturas Recepcionadas"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtfolio 
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
         Height          =   285
         Left            =   7485
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "rut"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   11625
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "fecha"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   11265
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
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
         Left            =   10905
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Buscar"
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
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1815
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
         Left            =   7500
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7860
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8220
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   600
         Width           =   615
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7140
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   12450
         _ExtentX        =   21960
         _ExtentY        =   12594
         Appearance      =   0
         BackColorBkg    =   16761024
         BackColorFixed  =   16777215
         BackColorScrollBar=   14737632
         BackColorSel    =   9567211
         Cols            =   8
         DefaultFontSize =   8.25
         Rows            =   1
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   675
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "LOCAL"
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
         Begin VB.ComboBox ComboLOCAL 
            Height          =   315
            Left            =   45
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   270
            Width           =   5715
         End
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA RECEPCION"
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
         Left            =   9000
         TabIndex        =   25
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA ENVIO"
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
         Left            =   6120
         TabIndex        =   24
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO ENVIO"
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
         Left            =   6120
         TabIndex        =   18
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta"
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
         Left            =   2460
         TabIndex        =   14
         Top             =   8490
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde"
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
         Left            =   255
         TabIndex        =   9
         Top             =   8490
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton btn_imprimir 
      BackColor       =   &H00FF8080&
      Caption         =   "IMPRIMIR"
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
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9120
      Width           =   2175
   End
   Begin VB.CommandButton btn_buscar 
      BackColor       =   &H00FF8080&
      Caption         =   "OTRA BUSQUEDA"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9120
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F5C9B1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CODIGO FACTURA"
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
      Left            =   120
      TabIndex        =   21
      Top             =   8520
      Width           =   1905
   End
End
Attribute VB_Name = "prove0014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean

Private Sub btn_buscar_Click()
CARGAGRILLA
txtfolio.text = ""
dato2.text = ""
dato3.text = ""
dato4.text = ""
DATO5.text = ""
dato6.text = ""
dato7.text = ""

txtfolio.SetFocus

End Sub

Private Sub COMMAND2_Click()
    Dim k As Double
    Dim FOLIO As String
    
    FOLIO = txtfolio.text
    For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 8).text = "1" Then
            Call marcafacturas(Grid1.Cell(k, 0).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 2).text, FOLIO)
        End If
    Next k
  
    Command1_Click
End Sub
Function LEERULTIMOFOLIO(loc) As String
    Dim csql  As New rdoQuery
    Dim resultado As rdoResultset
    Set csql.ActiveConnection = gestionrubro
    csql.sql = "select max(folioenvio)+1 from l_ordendecompra_detalle_facturas_" & loc
    csql.Execute
    LEERULTIMOFOLIO = "0000000001"
    
    If csql.RowsAffected > 0 Then
        Set resultado = csql.OpenResultset
        LEERULTIMOFOLIO = Format(resultado(0), "0000000000")
    End If
    csql.Close
    Set csql = Nothing
    
End Function
'Private Sub dato5_GotFocus()
'    dato5.SelStart = 0
'    dato5.SelLength = Len(dato5.text)
'    dato5.SetFocus
'End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyF2 Then 'Call ayudaProveedor(dato5)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO5): Call Pregunta(DATO5, dato6)

End Sub

Sub leeproveedor()
'    campos(0, 0) = "rut"
'    campos(1, 0) = "nombre"
'    campos(2, 0) = ""
'    campos(0, 2) = "r_maestroproveedores_" + rubro
'    'condicion = "rut='" & dato5.text + dv.caption & "' "
'    condicion = ""
'    op = 5
'    Set sqlconta.conexion = GESTIONrubro
'    sqlconta.response = campos
'    Call sqlconta.sqlconta(op, condicion)
    'nombreproveedor.caption = sqlconta.response(1, 3)
    Call leeOrdenes
    If sw = False Then
        'dato5.SelStart = 0
        'dato5.SelLength = Len(dato5)
        'dato5.SetFocus
    Else
        Grid1.SetFocus
    End If
End Sub

Sub ayudaProveedor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("13n", "30s")
    cfijo = "no"
    mensajeAyuda = "Ayuda Proveedores"
    cabezas = Array("rut", "nombre")
    'Call cargaAyudaT(servidor, basedatos & rubro, usuario, password, "r_maestroproveedores_" + rubro, dato5, campos, cfijo, largo, 2)
    'Call ceros(dato5)
    'dv.caption = rut(dato5)
    Call leeproveedor
End Sub

Function Busca_Proveedor(rutproveedor As String) As String
    campos(0, 0) = "ucase(nombre)"
    campos(1, 0) = ""
    campos(0, 2) = cliente_sql & "conta" & empresaactiva & ".cuentascorrientes"
    condicion = "rut='" & rutproveedor & "' and año='" & Format(fechasistema, "yyyy") & "' and tipo='23100026' "
    op = 5
    Set sqlconta.conexion = contadb
    sqlconta.response = campos
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then Busca_Proveedor = sqlconta.response(0, 3)
End Function

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Private Sub btn_imprimir_Click()
    Call Titulos
    
    
    Grid1.Column(7).Width = 0
    Grid1.PrintPreview
    Grid1.Column(7).Width = 80
    
End Sub

Private Sub Command1_Click()
    ORDEN.text = ""
    If txtfolio.text <> "" Then
        leeOrdenes
    Else
        MsgBox "DEBE INGRESAR UN NUMERO DE ENVIO A CONSULTAR"
    End If

End Sub

Private Sub exportar_Click()
 Grid1.ExportToExcel ("")

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
    
    Call cargatexto(DATO5)
End Sub

Private Sub dato6_GotFocus()
    Call cargatexto(dato6)
End Sub

Private Sub dato7_GotFocus()
    Call cargatexto(dato7)
End Sub


Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato2)
        Call Pregunta(dato2, dato3)
        If dato2.text = dia Then
            dato3.Enabled = True
            dato4.Enabled = True
            
            
        End If
    End If
    
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato7): Command1.SetFocus
    
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato6)
        Call Pregunta(dato6, dato7)
        If dato6.text = dia Then
            dato6.Enabled = True
            dato7.Enabled = True
            
            
        End If
    End If
    
End Sub


Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato4):: Call Pregunta(dato4, DATO5)
    
End Sub


Private Sub Form_Load()
    CARGAGRILLA
    LEErlocales

End Sub
    Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " & cliente_sql & "gestion" & ".g_maestroempresas where codigocontable='" & empresaactiva & "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        ComboLOCAL.Clear
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            ComboLOCAL.text = ComboLOCAL.List(0)
        End If
       
        
End Sub
Sub CARGAGRILLA()
    Grid1.Cols = 9
    
    Grid1.Column(0).Width = 30
    Grid1.Column(1).Width = 70
    Grid1.Column(2).Width = 70
    Grid1.Column(3).Width = 200
    Grid1.Column(4).Width = 90
    Grid1.Column(5).Width = 80
    Grid1.Column(6).Width = 80
    Grid1.Column(7).Width = 80
    Grid1.Column(8).Width = 80
    
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    Grid1.Column(8).Locked = False
    
    Grid1.Cell(0, 0).text = "TIPO"
    Grid1.Cell(0, 1).text = "NUMERO"
    Grid1.Cell(0, 2).text = "O.C."
    Grid1.Cell(0, 3).text = "PROVEEDOR"
    Grid1.Cell(0, 4).text = "FECHA"
    Grid1.Cell(0, 5).text = "MONTO"
    Grid1.Cell(0, 6).text = "F.RECEPCION"
    Grid1.Cell(0, 7).text = "ENTREGAR"
    Grid1.Cell(0, 8).text = "RECIBIDA"
    
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterGeneral
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).FontSize = 7
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).FontBold = True
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
    
    Grid1.Column(7).CellType = cellCheckBox
    Grid1.Column(8).CellType = cellCheckBox
    Grid1.Column(5).FormatString = "###,###,##0"
    Grid1.Column(4).Alignment = cellRightGeneral
    Grid1.Column(5).Alignment = cellRightGeneral
    Grid1.Column(6).Alignment = cellRightGeneral
    


    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid1.Rows = 1
End Sub

Sub leeOrdenes()
    Dim CONSULTA As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim total_comprado As Double
    Dim total_recepcionado As Double
    Dim color As String
    Dim bases As String
    Dim loc As String
    loc = Mid(ComboLOCAL.text, 1, 2)
    
        color = &HFFFFFF
        
 
         Set csql.ActiveConnection = contadb
 
            
           
        csql.sql = "SELECT lof.tipo,lof.numero,lof.ordendecompra,lof.rut,lof.fecha,lof.total,loc.fecharecepcion,"
        csql.sql = csql.sql & "if(lof.fechaentrega='0000-00-00','0','1'),if(lof.fecharecibe='0000-00-00','0','1')"
        csql.sql = csql.sql & ",ifnull(lof.fechaentrega,'0000-00-00'),ifnull(lof.fecharecibe,'0000-00-00') "
        csql.sql = csql.sql + "FROM " & cliente_sql & "gestion" & rubro & ".l_ordendecompra_detalle_facturas_" + loc + " as lof," & cliente_sql & "gestion" & rubro & ".l_ordendecompra_cabeza_" + loc + " as loc "
        csql.sql = csql.sql + "where lof.ordendecompra=loc.numero and loc.fecharecepcion>='2016-11-01' "
       
       If txtfolio.text = "" Then
         csql.sql = csql.sql + " and lof.folioenvio='' "
       Else
        csql.sql = csql.sql + " and lof.folioenvio='" & txtfolio.text & "' "
       End If
      
        
        csql.sql = csql.sql + "and (lof.tipo='FA' or lof.tipo='NC' or lof.tipo='ND' or lof.tipo='FAE' or lof.tipo='NCE' or lof.tipo='NDE') order by loc.fecharecepcion,lof.ordendecompra "
        csql.sql = csql.sql + ""
        csql.Execute
            
  
        sw = False
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             sw = True
            Grid1.AutoRedraw = False
            Grid1.Rows = 1
            Grid1.Cols = 9
            While Not resultados.EOF
           
        
'                     Grid1.AddItem resultados(0) & vbTab & resultados(1) & " " & Busca_Proveedor(resultados(1).Value) & vbTab & resultados(2) & vbTab & resultados(3) & vbTab & resultados(4) & vbTab & Format(resultados(5), "###,###,##0") & vbTab & Format(resultados(6), "###,###,##0"), False
                     Grid1.Rows = Grid1.Rows + 1
                     Grid1.Cell(Grid1.Rows - 1, 0).text = resultados(0)
                     Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
                     Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(2)
                     Grid1.Cell(Grid1.Rows - 1, 3).text = Busca_Proveedor(resultados(3))
                     Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(4)
                     Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(5)
                     Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(6)
                     Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(7)
                     Grid1.Cell(Grid1.Rows - 1, 8).text = resultados(8)
                     If txtfolio.text <> "" Then
                        dato2.text = Format(resultados(9), "dd")
                        dato3.text = Format(resultados(9), "mm")
                        dato4.text = Format(resultados(9), "yyyy")
                     End If
                     
                     If txtfolio.text <> "" Then
                        DATO5.text = Format(resultados(10), "dd")
                        dato6.text = Format(resultados(10), "mm")
                        dato7.text = Format(resultados(10), "yyyy")
                     End If
                     
                     
 
                
            
      
          
               
                resultados.MoveNext
            Wend
            
'            Grid1.Range(Grid1.rows - 1, 5, Grid1.rows - 1, 7).Borders(cellEdgeBottom) = cellThick
'            Grid1.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTALES" & vbTab & Format(total_comprado, "$ ###,###,##0") & vbTab & Format(total_recepcionado, "$ ###,###,##0"), False
'            Grid1.Range(Grid1.rows - 1, 1, Grid1.rows - 1, 7).FontSize = 8
'            Grid1.Range(Grid1.rows - 1, 1, Grid1.rows - 1, 7).FontBold = True
'            Grid1.Cell(Grid1.rows - 1, 5).Alignment = cellLeftGeneral
'            Grid1.Cell(Grid1.rows - 1, 6).Alignment = cellRightGeneral
'            Grid1.Cell(Grid1.rows - 1, 7).Alignment = cellRightGeneral
            
            resultados.Close
            Set resultados = Nothing
            Grid1.AutoRedraw = True
            Grid1.Refresh
            If dato7.text <> "0000" Then
                btn_imprimir.Enabled = True
            Else
                btn_imprimir.Enabled = False
            End If
        Else
            Grid1.Rows = 1
            MsgBox "No se encontraron resultados para la búsqueda, elija otro criterio e intente nuevamente.", vbInformation + vbOKOnly, "Consulta sin Resultados"
            dato2.SetFocus
        End If

End Sub

Sub Titulos()

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.PageSetup.Orientation = cellPortrait
    
    Grid1.PageSetup.PrintFixedRow = True
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.PrintTitleRows = 0
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE FACTURAS RECIBIDAS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FOLIO ENVIO : " & txtfolio.text & " FECHA RECEPCION: " & DATO5.text & "-" & dato6.text & "-" & dato7.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    Grid1.PageSetup.LeftMargin = 1
    Grid1.PageSetup.RightMargin = 1

    
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub
Private Function leerbonificacion(numero) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String
            Set csql.ActiveConnection = contadb
            bases = cliente_sql + "gestion" + rubro + "."

csql.sql = "select bonificacion from " & bases & "l_ordendecompra_detalle_facturas_" & empresaactiva & " "
csql.sql = csql.sql & "where ordendecompra='" & numero & "' "
csql.Execute
leerbonificacion = False

If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leerbonificacion = resultados(0)

End If

End Function

Private Sub Option1_Click()
Call Command1_Click
End Sub

Private Sub Option2_Click()
Call Command1_Click
End Sub

Private Sub Option3_Click()
Call Command1_Click
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub
Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
Sub marcafacturas(tipo, numero, oc, FOLIO)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim loc As String
    Dim rubroloc As String
    
    loc = Mid(ComboLOCAL.text, 1, 2)
    rubroloc = leerrubroloc(loc)
    Set csql.ActiveConnection = contadb
    csql.sql = "update " & cliente_sql & "gestion" & rubroloc & ".l_ordendecompra_detalle_facturas_" & loc
    csql.sql = csql.sql & " set usuariorecibe='" & USUARIOSISTEMA & "',"
    csql.sql = csql.sql & "fecharecibe='" & Format(fechasistema, "yyyy-mm-dd") & "' "
    csql.sql = csql.sql & " where tipo='" & tipo & "' and numero='" & numero & "' and ordendecompra='" & oc & "' and folioenvio='" & FOLIO & "'"
    csql.Execute
    Call sincronizadatos(csql.sql, conta, "")
    csql.Close
    Set csql = Nothing
    
End Sub
Private Function leerrubroloc(loc) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
       
        csql.sql = "select rubro "
        csql.sql = csql.sql + "FROM " & cliente_sql & "gestion" & ".g_maestroempresas where codigo='" & loc & "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        leerrubroloc = ""
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            leerrubroloc = resultados(0)
            resultados.Close
            Set resultados = Nothing
        End If
        csql.Close
        Set csql = Nothing
        
    
End Function

Private Sub txtfactura_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Call buscafactura(txtfactura.text)
    End If
End Sub

Private Sub txtfolio_GotFocus()
    Call cargatexto(txtfolio)
End Sub

Private Sub txtfolio_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(txtfolio)
        Command1.SetFocus
    End If
End Sub
Private Sub ORDEN_GotFocus()
Call cargatexto(ORDEN)
End Sub

Private Sub ORDEN_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(ORDEN)
Call BUSCAR_Click


End If

End Sub
Private Sub BUSCAR_Click()
 Dim i As Integer
 
  For i = 1 To Grid1.Rows - 1
            If Grid1.Cell(i, 2).text = ORDEN.text Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
            End If
        Next i
End Sub

Sub buscafactura(TEXTO)
    Dim tipo As String
    Dim numero As String
    Dim datos As Variant
    Dim dato2 As Variant
    Dim dato3 As Variant
    Dim i As Double
    On Error Resume Next
    
'    lof.tipo='FAE' or lof.tipo='NCE' or lof.tipo='NDE'
    
    
    datos = Split(TEXTO, ";")
    
    dato2 = Split(datos(5), ":")
    tipo = dato2(1)
    dato3 = Split(datos(7), ":")
    numero = Format(dato3(1), "0000000000")
    
    Select Case tipo
    Case "61"
        tipo = "NCE"
    Case "33"
        tipo = "FAE"
    Case "56"
        tipo = "NDE"
    End Select
    
    
    For i = 1 To Grid1.Rows - 1
            If Grid1.Cell(i, 0).text = tipo And Grid1.Cell(i, 1).text = numero Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
         End If
    Next i
    
    
End Sub
