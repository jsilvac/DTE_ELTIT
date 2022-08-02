VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form compra03 
   Caption         =   "Libro de Compra"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   15690
      BackColor       =   16761024
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPFrame.FrameXp FrameXp9 
         Height          =   855
         Left            =   0
         TabIndex        =   17
         Top             =   9600
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "Proporcionalidad"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtpropo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FF80&
            Height          =   375
            Left            =   1680
            TabIndex        =   20
            Text            =   "100"
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FF8080&
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FF8080&
            Caption         =   "Solo gastos"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TODOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   16
         Top             =   8040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CARACTERIZACION TIPO DE COMPRA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   15
         Top             =   8040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DOCUMENTOS NO ELECTRONICOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   8040
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Generar Archivo Doc Manuales"
         Height          =   735
         Left            =   11400
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdbuscar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Buscar"
         Height          =   735
         Left            =   9000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8040
         Width           =   2295
      End
      Begin XPFrame.FrameXp frmdetalle 
         Height          =   6495
         Left            =   0
         TabIndex        =   9
         Top             =   1440
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   11456
         BackColor       =   16761024
         Caption         =   "DETALLE LIBRO"
         BackColor       =   16761024
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
         Begin MSComctlLib.ProgressBar barra 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   6200
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin FlexCell.Grid Grid1 
            Height          =   5895
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   10398
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   1335
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   2355
         BackColor       =   16761024
         Caption         =   "Configuracion"
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
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   855
            Left            =   0
            TabIndex        =   2
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "EMPRESA"
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
            Begin VB.TextBox DATO1 
               Height          =   285
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Width           =   375
            End
            Begin VB.Label empresanombre 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   840
               TabIndex        =   4
               Top             =   360
               Width           =   3255
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   855
            Left            =   4560
            TabIndex        =   5
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "MES"
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
            Begin VB.ComboBox COMBOMES 
               Height          =   315
               Left            =   240
               TabIndex        =   6
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   855
            Left            =   9120
            TabIndex        =   7
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "AÑO"
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
            Begin VB.ComboBox COMBOAÑO 
               Height          =   315
               Left            =   240
               TabIndex        =   8
               Top             =   360
               Width           =   3855
            End
         End
      End
   End
End
Attribute VB_Name = "compra03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
    Private FORMATOGRILLA(20, 30)
    Private lin As Double
    Private tipoprove As String
    Private plan(2000, 3) As Variant
    Private canplan As Double
    Private total(21) As Double
    Private detalle(21, 21) As Double
    Private TIPOS(21) As String
    Private totaldocumentos As Double
    Private refrescos As String
    Private licores As String
    Private vinos As String
    Private cerveza As String
    Private HARINA As String
    Private CARNE As String
   
Private Sub cmdbuscar_Click()
    Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)
    For k = 1 To 2000
     plan(k, 3) = 0
    Next k
    For k = 1 To 20
        detalle(k, 1) = 0
        detalle(k, 2) = 0
        detalle(k, 3) = 0
        detalle(k, 4) = 0
        detalle(k, 5) = 0
        detalle(k, 6) = 0
        detalle(k, 7) = 0
        detalle(k, 8) = 0
        detalle(k, 9) = 0
        detalle(k, 10) = 0
        detalle(k, 11) = 0
        detalle(k, 12) = 0
        detalle(k, 13) = 0
        detalle(k, 14) = 0
        detalle(k, 15) = 0
        detalle(k, 16) = 0
        detalle(k, 17) = 0
        detalle(k, 18) = 0
        detalle(k, 19) = 0
        detalle(k, 20) = 0
        
    Next k


    Call Consulta_Informe2
End Sub

Private Sub Command1_Click()
    If opt1.Value = True Then
        Call ArchivoDocManuales
    Else
        MsgBox "DEBE TENER SELECCIONADO OPCION DOCUMENTOS NO ELECTRONICOS"
    End If
End Sub


 
Private Sub dato1_GotFocus()
     Call cargatexto(dato1)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato1.text <> "" Then
            leer
        End If
    End If
    
End Sub

Private Sub Form_Load()
     Dim disco As String
    For k = 1 To 12
    COMBOMES.AddItem MonthName(k)
    Next k
    COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
    For k = 2000 To Val(Format(fechasistema, "yyyy"))
    COMBOAÑO.AddItem k
    Next k
    COMBOAÑO.ListIndex = k - 2001
    dato1.text = empresaactiva
    empresanombre.Caption = nombreempresa
    Call CARGAGRILLA
    
    TIPOS(1) = "FACTURAS "
    TIPOS(2) = "NOTAS DE DEBITO"
    TIPOS(3) = "NOTAS DE CREDITO"
    TIPOS(4) = "FACTURAS ELECTRONICAS"
    TIPOS(5) = "NOTAS DE DEBITO ELECTRONICAS"
    TIPOS(6) = "NOTAS DE CREDITO ELECTRONICAS"
    TIPOS(7) = "FACTURAS ACTIVO FIJO ELECTRONICAS"
    TIPOS(8) = "FACTURAS COMPRAS PROPIAS"
    TIPOS(9) = "IMPORTACIONES."
    TIPOS(10) = "EXENTAS NORMALES"
    TIPOS(11) = "EXENTAS ELECTRONICAS"
    TIPOS(12) = "FACTURAS SUPERMERCADO "
    TIPOS(13) = "LIQUIDACION-FACTURAS ELECTRONICAS"
    TIPOS(14) = "FACTURAS ACTIVO FIJO NORMALES"
    
    


End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub

Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    COMBOMES.SetFocus
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "FOLIO"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "RUT"
    FORMATOGRILLA(1, 6) = "PROVEEDOR"
    FORMATOGRILLA(1, 7) = "NETO"
    FORMATOGRILLA(1, 8) = "IVA"
    FORMATOGRILLA(1, 9) = "EXENTO"
    FORMATOGRILLA(1, 10) = "IMPTO DIESEL"
    FORMATOGRILLA(1, 11) = "RETENCION"
    FORMATOGRILLA(1, 12) = "TOTAL"
    FORMATOGRILLA(1, 13) = "R.AZUCAR"
    FORMATOGRILLA(1, 14) = "LICORES"
    FORMATOGRILLA(1, 15) = "VINOS"
    FORMATOGRILLA(1, 16) = "CERVEZAS"
    FORMATOGRILLA(1, 17) = "HARINA"
    FORMATOGRILLA(1, 18) = "CARNE"
    FORMATOGRILLA(1, 19) = "R.N/AZUC"
    FORMATOGRILLA(1, 20) = "IVA/N/R"
    FORMATOGRILLA(1, 21) = "USO COMUN"
    FORMATOGRILLA(1, 22) = "A/F"
    FORMATOGRILLA(1, 23) = "DIESEL RECU"
    
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "4"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "30"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "9"
    FORMATOGRILLA(2, 12) = "9"
    FORMATOGRILLA(2, 13) = "9"
    FORMATOGRILLA(2, 14) = "9"
    FORMATOGRILLA(2, 15) = "9"
    FORMATOGRILLA(2, 16) = "9"
    FORMATOGRILLA(2, 17) = "9"
    FORMATOGRILLA(2, 18) = "9"
    FORMATOGRILLA(2, 19) = "9"
    
    FORMATOGRILLA(2, 20) = "9"
    FORMATOGRILLA(2, 21) = "9"
    FORMATOGRILLA(2, 22) = "3"
    FORMATOGRILLA(2, 23) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "N"
    FORMATOGRILLA(3, 16) = "N"
    FORMATOGRILLA(3, 17) = "N"
    FORMATOGRILLA(3, 18) = "N"
    FORMATOGRILLA(3, 19) = "N"
    FORMATOGRILLA(3, 20) = "N"
    FORMATOGRILLA(3, 21) = "N"
    FORMATOGRILLA(3, 22) = "N"
    FORMATOGRILLA(3, 23) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    FORMATOGRILLA(4, 10) = "###,###,###"
    FORMATOGRILLA(4, 11) = "###,###,###"
    FORMATOGRILLA(4, 12) = "###,###,###"
    FORMATOGRILLA(4, 13) = "###,###,###"
    FORMATOGRILLA(4, 14) = "###,###,###"
    FORMATOGRILLA(4, 15) = "###,###,###"
    FORMATOGRILLA(4, 16) = "###,###,###"
    FORMATOGRILLA(4, 17) = "###,###,###"
    FORMATOGRILLA(4, 18) = "###,###,###"
    FORMATOGRILLA(4, 19) = "###,###,###"
    FORMATOGRILLA(4, 20) = "###,###,###"
    FORMATOGRILLA(4, 21) = "###,###,###"
    FORMATOGRILLA(4, 23) = "###,###,###"
    
    Rem LOCCKED
    For k = 1 To 23
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    Grid1.Cols = 24
    Grid1.Rows = 2
    
    
    Grid1.DisplayFocusRect = False
   
    Grid1.BoldFixedCell = False
    
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

 
Sub Consulta_Informe2()
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim refresco As Double
    Dim licores As Double
    Dim vinos As Double
    Dim cerveza As Double
    Dim HARINA As Double
    Dim CARNE As Double
    Dim EXENTO As Double
    Dim proporcion As Double
    Dim noazucar As Double
    proporcional = ""
    If txtpropo.text = "" Then txtpropo.text = "0"
    proporcion = CDbl(Replace(txtpropo.text, ".", ","))
    proporcional = proporcion
    Dim norecu As Double
    Dim USOCOMUN As Double
    Dim PASO As String
    Dim k As Integer
    
    año = COMBOAÑO.text
    MES = COMBOMES.ListIndex + 1
    If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
    tipoprove = CUENTAPROVEEDOR

        totaldocumentos = 0
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT folio,fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,if(ivanorecuperable=1,0,iva) as iva,exento,impuestoespecifico,retencion,"
        csql.sql = csql.sql & "total,fc.electronica,fc.activo,fc.comprasuper,IF(ivanorecuperable=1,iva,0) AS ivanorecuperable,dieselrecuperado "
        csql.sql = csql.sql + "FROM facturasdecompras as fc,cuentascorrientes as cc "
        csql.sql = csql.sql + "WHERE fc.tipo<>'' and "
        If opt1.Value = True Then
            csql.sql = csql.sql & " (fc.tipo='1' or fc.tipo='2' or fc.tipo='3' or fc.tipo='8' or fc.tipo='9') and "
        End If
        
        If opt2.Value = True Then
            csql.sql = csql.sql & " (fc.tipo='4' or fc.tipo='0' or fc.tipo='L' or fc.tipo='5' or fc.tipo='6') and "
          
        End If
        
        
         csql.sql = csql.sql + "fc.rut=cc.rut and cc.año='" + COMBOAÑO.text + "' and cc.tipo='" + tipoprove + "' and añocontable='" + año + "' and mescontable='" + MES + "' "
        
      
        csql.sql = csql.sql + " order by fecha "
            
        csql.Execute
        Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
        total(7) = 0
        total(8) = 0
        total(9) = 0
        total(10) = 0
        total(11) = 0
        total(12) = 0
        total(13) = 0
        total(14) = 0
        total(15) = 0
        total(16) = 0
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected
        barra.Value = 0
        Set resultados = csql.OpenResultset
        lin = 0
         While Not resultados.EOF
            If opt2.Value = True Then
'                If resultados(2) = "0000000039" Then Stop
                    If ESGASTO(resultados(1), resultados(2), resultados(4), "") = False Then
                        If resultados(13) <> "S" Then
                            If resultados(14) <> "1" Then
                                If resultados(15) = 0 Then
                                    GoTo no:
                                End If
                            End If
                        End If
                    End If
'                End If
            End If
             barra.Value = lin
             lin = lin + 1
             Grid1.Rows = Grid1.Rows + 2
             For k = 0 To 11
             Grid1.Cell(lin, k + 1).text = resultados(k)
             
             Next k
             multi = 1
                totaldocumentos = totaldocumentos + 1
                If resultados(1) = "1" Then Grid1.Cell(lin, 2).text = "FA": Grid1.Cell(lin, 0).text = "30"
                If resultados(1) = "2" Then Grid1.Cell(lin, 2).text = "ND": Grid1.Cell(lin, 0).text = "55"
                If resultados(1) = "3" Then Grid1.Cell(lin, 2).text = "NC": multi = 1: Grid1.Cell(lin, 0).text = "60"
                If resultados(1) = "4" Then Grid1.Cell(lin, 2).text = "FAE": Grid1.Cell(lin, 0).text = "33"
                If resultados(1) = "5" Then Grid1.Cell(lin, 2).text = "NDE": Grid1.Cell(lin, 0).text = "56"
                If resultados(1) = "6" Then Grid1.Cell(lin, 2).text = "NCE": multi = 1: Grid1.Cell(lin, 0).text = "61"
                If resultados(1) = "7" Then Grid1.Cell(lin, 2).text = "FC": Grid1.Cell(lin, 0).text = "46"
                If resultados(1) = "8" Then Grid1.Cell(lin, 2).text = "IM": Grid1.Cell(lin, 0).text = "914"
                If resultados(1) = "9" Then Grid1.Cell(lin, 2).text = "FE": Grid1.Cell(lin, 0).text = "32"
                If resultados(1) = "0" Then Grid1.Cell(lin, 2).text = "FEE": Grid1.Cell(lin, 0).text = "34"
                If resultados(1) = "L" Then Grid1.Cell(lin, 2).text = "LFE": Grid1.Cell(lin, 0).text = "43"
                
             refrescos = leerimpuesto(resultados(1), resultados(2), resultados(4), "11400010")
             licores = leerimpuesto(resultados(1), resultados(2), resultados(4), "11400013")
             vinos = leerimpuesto(resultados(1), resultados(2), resultados(4), "11400011")
             cerveza = leerimpuesto(resultados(1), resultados(2), resultados(4), "11400014")
             HARINA = leerimpuesto(resultados(1), resultados(2), resultados(4), "11400005")
             CARNE = leerimpuesto(resultados(1), resultados(2), resultados(4), "11400012")
             noazucar = leerimpuesto(resultados(1), resultados(2), resultados(4), "11400017")
                
            Grid1.Cell(lin, 7).text = resultados(6) * multi
            Grid1.Cell(lin, 8).text = resultados(7) * multi
            Grid1.Cell(lin, 9).text = (resultados(8) - refrescos - licores - vinos - cerveza - HARINA - CARNE - noazucar) * multi
            Grid1.Cell(lin, 10).text = resultados(9) * multi
            Grid1.Cell(lin, 11).text = resultados(10) * multi
            Grid1.Cell(lin, 12).text = resultados(11) * multi
            Grid1.Cell(lin, 13).text = refrescos * multi
            Grid1.Cell(lin, 14).text = licores * multi
            Grid1.Cell(lin, 15).text = vinos * multi
            Grid1.Cell(lin, 16).text = cerveza * multi
            Grid1.Cell(lin, 17).text = HARINA * multi
            Grid1.Cell(lin, 18).text = CARNE * multi
            Grid1.Cell(lin, 19).text = noazucar * multi
            norecu = 0
            USOCOMUN = 0
             If opt2.Value = True Then proporcion = 90
            If proporcion <> 100 Then
                If ESGASTO(resultados(1), resultados(2), resultados(4), "") = True Then
                    norecu = resultados(7) - Round(resultados(7) * proporcion / 100)
                    USOCOMUN = resultados(7)
'                    Grid1.Cell(lin, 20).text = norecu * multi
                    Grid1.Cell(lin, 21).text = USOCOMUN * multi
                    norecu = 0
                End If
            End If
            
            If resultados("ivanorecuperable") > 0 Then
            
                    norecu = resultados("ivanorecuperable")
                    
                    Grid1.Cell(lin, 20).text = norecu * multi
            
            End If
            
            Grid1.Cell(lin, 22).text = resultados(13)
'            If resultados(16) > 0 Then Stop
            Grid1.Cell(lin, 23).text = resultados(16)
            
            
            Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 9) + "-" + Mid(resultados(4), 10, 1)
                
             If resultados(1) = "3" Or resultados(1) = "6" Then multi = -1 Else multi = 1
             total(1) = total(1) + resultados(6) * multi
             total(2) = total(2) + resultados(7) * multi
             total(3) = total(3) + (resultados(8) - refrescos - licores - vinos - cerveza - HARINA - CARNE - noazucar) * multi
             total(4) = total(4) + resultados(9) * multi
             total(5) = total(5) + resultados(10) * multi
             total(6) = total(6) + resultados(11) * multi
             total(7) = total(7) + refrescos * multi
             total(8) = total(8) + licores * multi
             total(9) = total(9) + vinos * multi
             total(10) = total(10) + cerveza * multi
             total(11) = total(11) + HARINA * multi
             total(12) = total(12) + CARNE * multi
             total(13) = total(13) + noazucar * multi
             total(14) = total(14) + norecu * multi
             total(15) = total(15) + USOCOMUN * multi
             total(16) = total(16) + resultados(16) * multi
             
             EXENTO = resultados(8) - refrescos - licores - vinos - cerveza - HARINA - CARNE - noazucar
                          
                          
                          If resultados(1) = "1" And resultados(13) <> "S" And resultados(14) <> "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + EXENTO: detalle(1, 5) = detalle(1, 5) + resultados(9): detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 7) = detalle(1, 7) + resultados(11): detalle(1, 8) = detalle(1, 8) + refrescos: detalle(1, 9) = detalle(1, 9) + licores: detalle(1, 10) = detalle(1, 10) + vinos: detalle(1, 11) = detalle(1, 11) + cerveza: detalle(1, 12) = detalle(1, 12) + HARINA: detalle(1, 13) = detalle(1, 13) + CARNE: detalle(1, 14) = detalle(1, 14) + noazucar: detalle(1, 15) = detalle(1, 15) + norecu: detalle(1, 16) = detalle(1, 16) + USOCOMUN: detalle(1, 17) = detalle(1, 17) + resultados(16)
                          If resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + EXENTO: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 7) = detalle(2, 7) + resultados(11): detalle(2, 8) = detalle(2, 8) + refrescos: detalle(2, 9) = detalle(2, 9) + licores: detalle(2, 10) = detalle(2, 10) + vinos: detalle(2, 11) = detalle(2, 11) + cerveza: detalle(2, 12) = detalle(2, 12) + HARINA: detalle(2, 13) = detalle(2, 13) + CARNE:  detalle(2, 14) = detalle(2, 14) + noazucar: detalle(2, 15) = detalle(2, 15) + norecu: detalle(2, 16) = detalle(2, 16) + USOCOMUN: detalle(2, 17) = detalle(2, 17) + resultados(16)
                          If resultados(1) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + EXENTO: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 7) = detalle(3, 7) + resultados(11): detalle(3, 8) = detalle(3, 8) + refrescos: detalle(3, 9) = detalle(3, 9) + licores: detalle(3, 10) = detalle(3, 10) + vinos: detalle(3, 11) = detalle(3, 11) + cerveza: detalle(3, 12) = detalle(3, 12) + HARINA: detalle(3, 13) = detalle(3, 13) + CARNE:  detalle(3, 14) = detalle(3, 14) + noazucar: detalle(3, 15) = detalle(3, 15) + norecu: detalle(3, 16) = detalle(3, 16) + USOCOMUN: detalle(3, 17) = detalle(3, 17) + resultados(16)
                          If resultados(1) = "4" And resultados(13) <> "S" And resultados(14) <> "1" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + EXENTO: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 7) = detalle(4, 7) + resultados(11): detalle(4, 8) = detalle(4, 8) + refrescos: detalle(4, 9) = detalle(4, 9) + licores: detalle(4, 10) = detalle(4, 10) + vinos: detalle(4, 11) = detalle(4, 11) + cerveza: detalle(4, 12) = detalle(4, 12) + HARINA: detalle(4, 13) = detalle(4, 13) + CARNE: detalle(4, 14) = detalle(4, 14) + noazucar: detalle(4, 15) = detalle(4, 15) + norecu: detalle(4, 16) = detalle(4, 16) + USOCOMUN: detalle(4, 17) = detalle(4, 17) + resultados(16)
                          If resultados(1) = "5" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(6): detalle(5, 3) = detalle(5, 3) + resultados(7):: detalle(5, 4) = detalle(5, 4) + EXENTO: detalle(5, 5) = detalle(5, 5) + resultados(9): detalle(5, 6) = detalle(5, 6) + resultados(10): detalle(5, 7) = detalle(5, 7) + resultados(11): detalle(5, 8) = detalle(5, 8) + refrescos: detalle(5, 9) = detalle(5, 9) + licores: detalle(5, 10) = detalle(5, 10) + vinos: detalle(5, 11) = detalle(5, 11) + cerveza: detalle(5, 12) = detalle(5, 12) + HARINA: detalle(5, 13) = detalle(5, 13) + CARNE:  detalle(5, 14) = detalle(5, 14) + noazucar: detalle(5, 15) = detalle(5, 15) + norecu: detalle(5, 16) = detalle(5, 16) + USOCOMUN: detalle(5, 17) = detalle(5, 17) + resultados(16)
                          If resultados(1) = "6" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(6): detalle(6, 3) = detalle(6, 3) + resultados(7):: detalle(6, 4) = detalle(6, 4) + EXENTO: detalle(6, 5) = detalle(6, 5) + resultados(9): detalle(6, 6) = detalle(6, 6) + resultados(10): detalle(6, 7) = detalle(6, 7) + resultados(11): detalle(6, 8) = detalle(6, 8) + refrescos: detalle(6, 9) = detalle(6, 9) + licores: detalle(6, 10) = detalle(6, 10) + vinos: detalle(6, 11) = detalle(6, 11) + cerveza: detalle(6, 12) = detalle(6, 12) + HARINA: detalle(6, 13) = detalle(6, 13) + CARNE:  detalle(6, 14) = detalle(6, 14) + noazucar: detalle(6, 15) = detalle(6, 15) + norecu: detalle(6, 16) = detalle(6, 16) + USOCOMUN: detalle(6, 17) = detalle(6, 17) + resultados(16)
                          If resultados(13) = "S" And resultados(1) <> "1" And resultados(1) <> "3" And resultados(1) <> "6" Then detalle(7, 1) = detalle(7, 1) + 1: detalle(7, 2) = detalle(7, 2) + resultados(6): detalle(7, 3) = detalle(7, 3) + resultados(7): detalle(7, 4) = detalle(7, 4) + EXENTO:: detalle(7, 5) = detalle(7, 5) + resultados(9): detalle(7, 6) = detalle(7, 6) + resultados(10): detalle(7, 7) = detalle(7, 7) + resultados(11): detalle(7, 8) = detalle(7, 8) + refrescos: detalle(7, 9) = detalle(7, 9) + licores: detalle(7, 10) = detalle(7, 10) + vinos: detalle(7, 11) = detalle(7, 11) + cerveza: detalle(7, 12) = detalle(7, 12) + HARINA: detalle(7, 13) = detalle(7, 13) + CARNE:   detalle(7, 14) = detalle(7, 14) + noazucar: detalle(7, 15) = detalle(7, 15) + norecu: detalle(7, 16) = detalle(7, 16) + USOCOMUN: detalle(7, 17) = detalle(7, 17) + resultados(16)
                          If resultados(1) = "7" Then detalle(8, 1) = detalle(8, 1) + 1: detalle(8, 2) = detalle(8, 2) + resultados(6): detalle(8, 3) = detalle(8, 3) + resultados(7):: detalle(8, 4) = detalle(8, 4) + EXENTO: detalle(8, 5) = detalle(8, 5) + resultados(9): detalle(8, 6) = detalle(8, 6) + resultados(10): detalle(8, 7) = detalle(8, 7) + resultados(11): detalle(8, 8) = detalle(8, 8) + refrescos: detalle(8, 9) = detalle(8, 9) + licores: detalle(8, 10) = detalle(8, 10) + vinos: detalle(8, 11) = detalle(8, 11) + cerveza: detalle(8, 12) = detalle(8, 12) + HARINA: detalle(8, 13) = detalle(8, 13) + CARNE:  detalle(8, 14) = detalle(8, 14) + noazucar: detalle(8, 15) = detalle(8, 15) + norecu: detalle(8, 16) = detalle(8, 16) + USOCOMUN: detalle(8, 17) = detalle(8, 17) + resultados(16)
                          If resultados(1) = "8" Then detalle(9, 1) = detalle(9, 1) + 1: detalle(9, 2) = detalle(9, 2) + resultados(6): detalle(9, 3) = detalle(9, 3) + resultados(7):: detalle(9, 4) = detalle(9, 4) + EXENTO: detalle(9, 5) = detalle(9, 5) + resultados(9): detalle(9, 6) = detalle(9, 6) + resultados(10): detalle(9, 7) = detalle(9, 7) + resultados(11): detalle(9, 8) = detalle(9, 8) + refrescos: detalle(9, 9) = detalle(9, 9) + licores: detalle(9, 10) = detalle(9, 10) + vinos: detalle(9, 11) = detalle(9, 11) + cerveza: detalle(9, 12) = detalle(9, 12) + HARINA: detalle(9, 13) = detalle(9, 13) + CARNE:  detalle(9, 14) = detalle(9, 14) + noazucar: detalle(9, 15) = detalle(9, 15) + norecu: detalle(9, 16) = detalle(9, 16) + USOCOMUN: detalle(9, 17) = detalle(9, 17) + resultados(16)
                          If resultados(1) = "9" Then detalle(10, 1) = detalle(10, 1) + 1: detalle(10, 2) = detalle(10, 2) + resultados(6): detalle(10, 3) = detalle(10, 3) + resultados(7):: detalle(10, 4) = detalle(10, 4) + EXENTO: detalle(10, 5) = detalle(10, 5) + resultados(9): detalle(10, 6) = detalle(10, 6) + resultados(10): detalle(10, 7) = detalle(10, 7) + resultados(11): detalle(10, 8) = detalle(10, 8) + refrescos: detalle(10, 9) = detalle(10, 9) + licores: detalle(10, 10) = detalle(10, 10) + vinos: detalle(10, 11) = detalle(10, 11) + cerveza: detalle(10, 12) = detalle(10, 12) + HARINA: detalle(10, 13) = detalle(10, 13) + CARNE:  detalle(10, 14) = detalle(10, 14) + noazucar: detalle(10, 15) = detalle(10, 15) + norecu: detalle(10, 16) = detalle(10, 16) + USOCOMUN: detalle(10, 17) = detalle(10, 17) + resultados(16)
                          If resultados(1) = "0" Then detalle(11, 1) = detalle(11, 1) + 1: detalle(11, 2) = detalle(11, 2) + resultados(6): detalle(11, 3) = detalle(11, 3) + resultados(7):: detalle(11, 4) = detalle(11, 4) + EXENTO: detalle(11, 5) = detalle(11, 5) + resultados(9): detalle(11, 6) = detalle(11, 6) + resultados(10): detalle(11, 7) = detalle(11, 7) + resultados(11): detalle(11, 8) = detalle(11, 8) + refrescos: detalle(11, 9) = detalle(11, 9) + licores: detalle(11, 10) = detalle(11, 10) + vinos: detalle(11, 11) = detalle(11, 11) + cerveza: detalle(11, 12) = detalle(11, 12) + HARINA: detalle(11, 13) = detalle(11, 13) + CARNE:  detalle(11, 14) = detalle(11, 14) + noazucar: detalle(11, 15) = detalle(11, 15) + norecu: detalle(11, 16) = detalle(11, 16) + USOCOMUN: detalle(11, 17) = detalle(11, 17) + resultados(16)
                          If resultados(14) = "1" And (resultados(1) = "1" Or resultados(1) = "4") Then detalle(12, 1) = detalle(12, 1) + 1: detalle(12, 2) = detalle(12, 2) + resultados(6): detalle(12, 3) = detalle(12, 3) + resultados(7):: detalle(12, 4) = detalle(12, 4) + EXENTO: detalle(12, 5) = detalle(12, 5) + resultados(9): detalle(12, 6) = detalle(12, 6) + resultados(10): detalle(12, 7) = detalle(12, 7) + resultados(11): detalle(12, 8) = detalle(12, 8) + refrescos: detalle(12, 9) = detalle(12, 9) + licores: detalle(12, 10) = detalle(12, 10) + vinos: detalle(12, 11) = detalle(12, 11) + cerveza: detalle(12, 12) = detalle(12, 12) + HARINA: detalle(12, 13) = detalle(12, 13) + CARNE: detalle(12, 14) = detalle(12, 14) + noazucar: detalle(12, 15) = detalle(12, 15) + norecu: detalle(12, 16) = detalle(12, 16) + USOCOMUN: detalle(12, 17) = detalle(12, 17) + resultados(16)
                          
                          If resultados(1) = "L" Then detalle(13, 1) = detalle(13, 1) + 1: detalle(13, 2) = detalle(13, 2) + resultados(6): detalle(13, 3) = detalle(13, 3) + resultados(7):: detalle(13, 4) = detalle(13, 4) + EXENTO: detalle(13, 5) = detalle(13, 5) + resultados(9): detalle(13, 6) = detalle(13, 6) + resultados(10): detalle(13, 7) = detalle(13, 7) + resultados(11): detalle(13, 8) = detalle(13, 8) + refrescos: detalle(13, 9) = detalle(13, 9) + licores: detalle(13, 10) = detalle(13, 10) + vinos: detalle(11, 11) = detalle(13, 11) + cerveza: detalle(13, 12) = detalle(13, 12) + HARINA: detalle(13, 13) = detalle(13, 13) + CARNE:  detalle(13, 14) = detalle(13, 14) + noazucar: detalle(13, 15) = detalle(13, 15) + norecu: detalle(13, 16) = detalle(13, 16) + USOCOMUN: detalle(13, 17) = detalle(13, 17) + resultados(16)
                          If resultados(13) = "S" And resultados(1) = "1" Then detalle(14, 1) = detalle(14, 1) + 1: detalle(14, 2) = detalle(14, 2) + resultados(6): detalle(14, 3) = detalle(14, 3) + resultados(7): detalle(14, 4) = detalle(14, 4) + EXENTO:: detalle(14, 5) = detalle(14, 5) + resultados(9): detalle(14, 6) = detalle(14, 6) + resultados(10): detalle(14, 7) = detalle(14, 7) + resultados(11): detalle(14, 8) = detalle(14, 8) + refrescos: detalle(14, 9) = detalle(14, 9) + licores: detalle(14, 10) = detalle(14, 10) + vinos: detalle(14, 11) = detalle(14, 11) + cerveza: detalle(14, 12) = detalle(14, 12) + HARINA: detalle(14, 13) = detalle(14, 13) + CARNE:   detalle(14, 14) = detalle(14, 14) + noazucar: detalle(14, 15) = detalle(14, 15) + norecu: detalle(14, 16) = detalle(14, 16) + USOCOMUN: detalle(14, 17) = detalle(14, 17) + resultados(16)
             
                 If opt2.Value = True Or opt3.Value = True Then
                     If ESTAENSII(resultados(1), resultados(2), resultados(4), resultados(11)) = True Then
                         Grid1.Range(lin, 1, lin, Grid1.Cols - 1).BackColor = vbGreen
                     End If
                 End If
              

'            If (resultados(1) = "4" Or resultados(1) = "5" Or resultados(1) = "6" Or resultados(1) = "0") Then
'                Grid1.Range(lin, 1, lin, Grid1.Cols - 1).BackColor = vbRed
'
'            End If
            
            
no:

PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
     
Call totallibro2
barra.Max = 1
Grid1.AutoRedraw = True
Grid1.Refresh
 

End Sub

Sub totallibro2()
    Dim totales40(1, 21) As Double
    
    Dim TOTALge As Double
      lin = lin + 1
        Grid1.Rows = lin + 1
        Grid1.Range(lin, 7, lin, 22).Borders(cellEdgeTop) = cellThin
        Grid1.Cell(lin, 6).text = "TOTAL DOCUMENTOS  " & Format(totaldocumentos, "###,###,###")
        Grid1.Cell(lin, 7).text = total(1)
        Grid1.Cell(lin, 8).text = total(2)
        Grid1.Cell(lin, 9).text = total(3)
        Grid1.Cell(lin, 10).text = total(4)
        Grid1.Cell(lin, 11).text = total(5)
        Grid1.Cell(lin, 12).text = total(6)
        Grid1.Cell(lin, 13).text = total(7)
        Grid1.Cell(lin, 14).text = total(8)
        Grid1.Cell(lin, 15).text = total(9)
        Grid1.Cell(lin, 16).text = total(10)
        Grid1.Cell(lin, 17).text = total(11)
        Grid1.Cell(lin, 18).text = total(12)
        Grid1.Cell(lin, 19).text = total(13)
        Grid1.Cell(lin, 20).text = total(14)
        Grid1.Cell(lin, 21).text = total(15)
        Grid1.Cell(lin, 23).text = total(16)
        
    
    TOTALge = 0
    lin = lin + 2
    Grid1.Rows = Grid1.Rows + 17
    Grid1.Range(lin, 5, lin + 14, 23).Borders(cellEdgeTop) = cellThin
    Grid1.Range(lin, 5, lin + 14, 23).Borders(cellEdgeLeft) = cellThin
    Grid1.Range(lin, 5, lin + 14, 23).Borders(cellEdgeRight) = cellThin
    Grid1.Range(lin, 5, lin + 14, 23).Borders(cellEdgeBottom) = cellThin
    Grid1.Range(lin, 5, lin + 14, 23).Borders(cellInsideHorizontal) = cellThin
    Grid1.Range(lin, 5, lin + 14, 23).Borders(cellInsideVertical) = cellThin
    
    Grid1.Cell(lin, 5).text = "Cant."
    Grid1.Cell(lin, 6).text = "Documentos"
    Grid1.Cell(lin, 7).text = "Neto"
    Grid1.Cell(lin, 8).text = "i.v.a"
    Grid1.Cell(lin, 9).text = "exento"
    Grid1.Cell(lin, 10).text = "diesel"
    Grid1.Cell(lin, 11).text = "retencion"
    Grid1.Cell(lin, 12).text = "total"
    Grid1.Cell(lin, 13).text = "R.Azuc"
    Grid1.Cell(lin, 14).text = "Licores"
    Grid1.Cell(lin, 15).text = "Vinos"
    Grid1.Cell(lin, 16).text = "Cerveza"
    Grid1.Cell(lin, 17).text = "Harina"
    Grid1.Cell(lin, 18).text = "Carne"
    Grid1.Cell(lin, 19).text = "R.N/Azuc"
    Grid1.Cell(lin, 20).text = "Iva N/R"
    Grid1.Cell(lin, 21).text = "Iva comun"
    Grid1.Cell(lin, 23).text = "Diesel Recu."
    Dim T As Double
    
    For T = 1 To 17
    totales40(1, T) = 0
    Next T
    
    For k = 1 To 14
    lin = lin + 1
    
    Grid1.Cell(lin, 6).text = TIPOS(k)
    Grid1.Cell(lin, 5).text = Format(detalle(k, 1), "###,###,##0")
    Grid1.Cell(lin, 7).text = Format(detalle(k, 2), "###,###,##0")
    Grid1.Cell(lin, 8).text = Format(detalle(k, 3), "###,###,##0")
    Grid1.Cell(lin, 9).text = Format(detalle(k, 4), "###,###,##0")
    Grid1.Cell(lin, 10).text = Format(detalle(k, 5), "###,###,##0")
    Grid1.Cell(lin, 11).text = Format(detalle(k, 6), "###,###,##0")
    Grid1.Cell(lin, 12).text = Format(detalle(k, 7), "###,###,##0")
    Grid1.Cell(lin, 13).text = Format(detalle(k, 8), "###,###,##0")
    Grid1.Cell(lin, 14).text = Format(detalle(k, 9), "###,###,##0")
    Grid1.Cell(lin, 15).text = Format(detalle(k, 10), "###,###,##0")
    Grid1.Cell(lin, 16).text = Format(detalle(k, 11), "###,###,##0")
    Grid1.Cell(lin, 17).text = Format(detalle(k, 12), "###,###,##0")
    Grid1.Cell(lin, 18).text = Format(detalle(k, 13), "###,###,##0")
    Grid1.Cell(lin, 19).text = Format(detalle(k, 14), "###,###,##0")
    Grid1.Cell(lin, 20).text = Format(detalle(k, 15), "###,###,##0")
    Grid1.Cell(lin, 21).text = Format(detalle(k, 16), "###,###,##0")
    Grid1.Cell(lin, 23).text = Format(detalle(k, 17), "###,###,##0")
    
    For T = 1 To 17
    If k = 3 Or k = 6 Then
    totales40(1, T) = totales40(1, T) - detalle(k, T)
    Else
    totales40(1, T) = totales40(1, T) + detalle(k, T)
    
    End If
    
    Next T
    Next k
    
    
    
    
    
    Grid1.Rows = Grid1.Rows + 1
    lin = lin + 1
    Grid1.Cell(Grid1.Rows - 1, 6).text = "TOTALES CUADRATURA "
    Rem Grid1.Cell(Grid1.Rows - 1, 5).text = Format(totales40(1, 1), "###,###,##0")
    
    Grid1.Cell(Grid1.Rows - 1, 7).text = Format(totales40(1, 2), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 8).text = Format(totales40(1, 3), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 9).text = Format(totales40(1, 4), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 10).text = Format(totales40(1, 5), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 11).text = Format(totales40(1, 6), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 12).text = Format(totales40(1, 7), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 13).text = Format(totales40(1, 8), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 14).text = Format(totales40(1, 9), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 15).text = Format(totales40(1, 10), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 16).text = Format(totales40(1, 11), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 17).text = Format(totales40(1, 12), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 18).text = Format(totales40(1, 13), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 19).text = Format(totales40(1, 14), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 20).text = Format(totales40(1, 15), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 21).text = Format(totales40(1, 16), "###,###,##0")
    Grid1.Cell(Grid1.Rows - 1, 23).text = Format(totales40(1, 17), "###,###,##0")
    
    Grid1.Rows = Grid1.Rows + 1
    lin = lin + 1
    
    
    For k = 1 To canplan
    If plan(k, 3) <> 0 Then
             lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(lin, 5).text = plan(k, 1)
        Grid1.Cell(lin, 6).text = plan(k, 2)
        Grid1.Cell(lin, 7).text = plan(k, 3)
        TOTALge = TOTALge + plan(k, 3)
        End If
    Next k
        lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(lin, 6, lin, 7).Borders(cellEdgeTop) = cellThin
        
        
        
        
        
        Grid1.Cell(lin, 6).text = "TOTAL DETALLE"
         Grid1.Cell(lin, 7).text = TOTALge
               
    End Sub
Public Function leerimpuesto(tipo, numero, rut, cuenta)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT monto "
        csql2.sql = csql2.sql + "FROM facturasdecompras_detalle "
        csql2.sql = csql2.sql + "where tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' and cuentadelmayor='" + cuenta + "' "
        csql2.Execute
        leerimpuesto = 0
        If csql2.RowsAffected > 0 Then
        
        Set resultados2 = csql2.OpenResultset
        linpaso = 0
        While Not resultados2.EOF
          
        leerimpuesto = resultados2(0)
        resultados2.MoveNext
        Wend

          resultados2.Close

        End If

End Function

Public Function ESGASTO(tipo, numero, rut, cuenta) As Boolean
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT cuentadelmayor "
        csql2.sql = csql2.sql + "FROM facturasdecompras_detalle "
        csql2.sql = csql2.sql + "where tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' and cuentadelmayor like '" + "4%" + "' "
        csql2.Execute
        ESGASTO = False
        If csql2.RowsAffected > 0 Then
        ESGASTO = True
        End If

End Function
Sub ArchivoDocManuales()
    Dim k As Double
    Dim ARCHIVO As String
    Dim cadena As String
    Dim i As Double
    Dim o As Double
    Dim tipotrans As String
    
    
    
    
    año = COMBOAÑO.text
    MES = COMBOMES.ListIndex + 1
    If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
    
    Close 20
    ARCHIVO = "u:\LIBROS\Doc_manuales_libro_compras_" + empresaactiva + "_" + año & "_" & MES + ".csv"


        Open ARCHIVO For Output As #20
    Print #20, "titulos"
    For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 0).text <> "" Then
            o = 0
              For i = 13 To 19
                If Grid1.Cell(k, i).text > 0 Then
                    o = o + 1
                End If
            Next i
            If o > 0 Then
                For i = 13 To 19
                    If Grid1.Cell(k, i).text > 0 Then
                    
                        cadena = Grid1.Cell(k, 0).text & ";" 'Tipo Doc;
                        cadena = cadena & Grid1.Cell(k, 3).text & ";" 'Folio;
                        cadena = cadena & Val(Mid(Grid1.Cell(k, 5).text, 1, 9)) & "-" & Right(Grid1.Cell(k, 5).text, 1) & ";" 'Rut Contraparte;
                        cadena = cadena & "19" & ";" 'Tasa Impuesto;
                        cadena = cadena & Grid1.Cell(k, 6).text & ";" 'Razon Social Contraparte;
                        cadena = cadena & "1" & ";" 'Tipo Impuesto[1=IVA:2=LEY 18211];
                        cadena = cadena & Format(Grid1.Cell(k, 4).text, "dd-mm-yyyy") & ";" 'Fecha Emision;
                        cadena = cadena & Grid1.Cell(k, 9).text & ";" 'Monto Exento;
                        cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Neto;
                        cadena = cadena & Val(Grid1.Cell(k, 8).text) - Val(Grid1.Cell(k, 21).text) & ";" 'Monto IVA (Recuperable);
                        If Val(Grid1.Cell(k, 20).text) > 0 Then
                            cadena = cadena & "1" & ";" 'Cod IVA no Rec;
                            cadena = cadena & Grid1.Cell(k, 20).text & ";" 'Monto IVA no Rec;
                        Else
                            cadena = cadena & "" & ";" 'Cod IVA no Rec;
                            cadena = cadena & "" & ";" 'Monto IVA no Rec;
                        End If
                        
                        cadena = cadena & Grid1.Cell(k, 21).text & ";" 'IVA Uso Comun;
            
          


                        If i = 13 Then
                            cadena = cadena & 271 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 18 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                        If i = 14 Then
                            cadena = cadena & 24 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & "31.5" & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 15 Then
                            cadena = cadena & 25 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & "20.5" & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 16 Then
                            cadena = cadena & 26 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & "20.5" & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 17 Then
                            cadena = cadena & 19 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 12 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 18 Then
                            cadena = cadena & 18 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 5 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 19 Then
                            cadena = cadena & 27 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 10 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                        
                        
                        cadena = cadena & Grid1.Cell(k, i).text & ";" 'Monto Otro Imp (Con Credito);
                        
                        cadena = cadena & "" & ";" 'Monto Otro Imp Sin Credito;
                        
                        'FormatoGrilla(1, 21) = "USO COMUN"
                        'FormatoGrilla(1, 22) = "A/F"
                        
'                        FormatoGrilla(1, 20) = "IVA/N/R"
'                        FormatoGrilla(1, 21) = "USO COMUN"
'                        FormatoGrilla(1, 22) = "A/F"
    
    
                        If Grid1.Cell(k, 22).text = "S" Then
                            cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Activo Fijo;
                            cadena = cadena & Grid1.Cell(k, 8).text & ";" 'Monto IVA Activo Fijo;
                        Else
                            cadena = cadena & "" & ";" 'Monto Activo Fijo;
                            cadena = cadena & "" & ";" 'Monto IVA Activo Fijo;
                        End If
                        
                        cadena = cadena & "" & ";" 'IVA No Retenido;
                        
                        
                        
                        cadena = cadena & "" & ";" 'Tabacos - Puros;
                        cadena = cadena & "" & ";" 'Tabacos - Cigarrillos;
                        cadena = cadena & "" & ";" 'Tabacos - Elaborados;
                        cadena = cadena & "" & ";" 'Codigo sucursal SII;
                        cadena = cadena & "" & ";" 'Numero Interno;
                        cadena = cadena & "" & ";" 'Emisor/Receptor;
                        cadena = cadena & Grid1.Cell(k, 12).text & ";" 'Monto Total;
                        tipotrans = "1"
                        
'                        If Grid1.Cell(k, 20).text = "" And Grid1.Cell(k, 21).text = "" And Grid1.Cell(k, 22).text = "N" Then
'                            tipotrans = "2"
'                        End If
                        
                        If Grid1.Cell(k, 22).text = "S" Then
                           tipotrans = "4" 'Tipo Transaccion
                        End If
                        
                        If Grid1.Cell(k, 20).text <> "" Then
                           tipotrans = "6" 'Tipo Transaccion
                        End If
                        If Grid1.Cell(k, 21).text <> "" Then
                           tipotrans = "5" 'Tipo Transaccion
                        End If
                        
                        
                        
                         cadena = cadena & tipotrans
                         
                         
                      
                        
                        
                         Print #20, cadena
                    End If
                Next i
            Else
                        cadena = Grid1.Cell(k, 0).text & ";" 'Tipo Doc;
                        cadena = cadena & Grid1.Cell(k, 3).text & ";" 'Folio;
                        cadena = cadena & Val(Mid(Grid1.Cell(k, 5).text, 1, 9)) & "-" & Right(Grid1.Cell(k, 5).text, 1) & ";" 'Rut Contraparte;
                        cadena = cadena & "19" & ";" 'Tasa Impuesto;
                        cadena = cadena & Grid1.Cell(k, 6).text & ";" 'Razon Social Contraparte;
                        cadena = cadena & "1" & ";" 'Tipo Impuesto[1=IVA:2=LEY 18211];
                        cadena = cadena & Format(Grid1.Cell(k, 4).text, "dd-mm-yyyy") & ";" 'Fecha Emision;
                        cadena = cadena & Grid1.Cell(k, 9).text & ";" 'Monto Exento;
                        cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Neto;
                        cadena = cadena & Val(Grid1.Cell(k, 8).text) - Val(Grid1.Cell(k, 21).text) & ";" 'Monto IVA (Recuperable);
                        If Val(Grid1.Cell(k, 20).text) > 0 Then
                            cadena = cadena & "1" & ";" 'Cod IVA no Rec;
                            cadena = cadena & Grid1.Cell(k, 20).text & ";" 'Monto IVA no Rec;
                        Else
                            cadena = cadena & "" & ";" 'Cod IVA no Rec;
                            cadena = cadena & "" & ";" 'Monto IVA no Rec;
                        End If
                        
                        cadena = cadena & Grid1.Cell(k, 21).text & ";" 'IVA Uso Comun;

                        cadena = cadena & "" & ";" 'Cod Otro Imp (Con Credito);
                        cadena = cadena & "" & ";" 'Tasa Otro Imp (Con Credito);
                        cadena = cadena & "" & ";" 'Monto Otro Imp (Con Credito);
                        
                        cadena = cadena & "" & ";" 'Monto Otro Imp Sin Credito;
                         If Grid1.Cell(k, 22).text = "S" Then
                            cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Activo Fijo;
                            cadena = cadena & Grid1.Cell(k, 8).text & ";" 'Monto IVA Activo Fijo;
                        Else
                            cadena = cadena & "" & ";" 'Monto Activo Fijo;
                            cadena = cadena & "" & ";" 'Monto IVA Activo Fijo;
                        End If
                        cadena = cadena & "" & ";" 'IVA No Retenido;
                        
                        
                        
                        cadena = cadena & "" & ";" 'Tabacos - Puros;
                        cadena = cadena & "" & ";" 'Tabacos - Cigarrillos;
                        cadena = cadena & "" & ";" 'Tabacos - Elaborados;
                        cadena = cadena & "" & ";" 'Codigo sucursal SII;
                        cadena = cadena & "" & ";" 'Numero Interno;
                        cadena = cadena & "" & ";" 'Emisor/Receptor;
                        cadena = cadena & Grid1.Cell(k, 12).text & ";" 'Monto Total;
                        
                        
                        tipotrans = "1"
                        
'                        If Grid1.Cell(k, 20).text = "" And Grid1.Cell(k, 21).text = "" And Grid1.Cell(k, 22).text = "N" Then
'                            tipotrans = "2"
'                        End If
                        
                        If Grid1.Cell(k, 22).text = "S" Then
                           tipotrans = "4" 'Tipo Transaccion
                        End If
                        
                        If Grid1.Cell(k, 20).text <> "" Then
                           tipotrans = "6" 'Tipo Transaccion
                        End If
                        If Grid1.Cell(k, 21).text <> "" Then
                           tipotrans = "5" 'Tipo Transaccion
                        End If
                        
                        
                        
                         cadena = cadena & tipotrans
                        
                        Print #20, cadena
            End If
            
        End If
    Next k
    
    Close 20
    Shell "NOTEPAD " + ARCHIVO

End Sub
Sub ArchivoComplementos()
  Dim ARCHIVO As String
  Dim contador As Double
  Dim codigo_iva As Double
  Dim tipodoc As String
  Dim cadena As String
  Dim TpoTranCompra As Double
  
     año = COMBOAÑO.text
    MES = COMBOMES.ListIndex + 1
    If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
    
    Close 20
    ARCHIVO = "u:\LIBROS\Caracterizacion_" + año & "-" & MES + ".csv"


        Open ARCHIVO For Output As #20
        contador = 0
    For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 0).text <> "" Then
        If Grid1.Cell(k, 1).BackColor <> vbGreen Then GoTo no:
        
            contador = contador + 1
            If contador = 1 Then
                cadena = "RUT-DV;Codigo_Tipo_Doc;Folio_Doc;TpoTranCompra;Codigo_IVA_e_Impuestos"
                Print #20, cadena
            End If
            cadena = Val(Mid(Grid1.Cell(k, 5).text, 1, 9)) & "-" & Right(Grid1.Cell(k, 5).text, 1) & ";"  'rut-dv
            cadena = cadena & Grid1.Cell(k, 0).text & ";" 'Codigo_Tipo_Doc;
            cadena = cadena & Grid1.Cell(k, 3).text & ";" 'Folio_Doc;
            'TpoTranCompra;
            '1 compras del giro
            '2 compra supermercados o compercios similares
            '3 adquisicion bienes raices
            '4 activo fijo
            '5 compras con IVA uso Comun
            '6 Compras sin Derecho  a Credito(IVA no Recuperable)
            '7 compras que no corresponde incluir
'            FormatoGrilla(1, 20) = "IVA/N/R"
'            FormatoGrilla(1, 21) = "USO COMUN"
'            FormatoGrilla(1, 22) = "A/F"
            If Grid1.Cell(k, 3).text = "0000000039" Then Stop
            TpoTranCompra = 1
            codigo_iva = 1
            If Grid1.Cell(k, 22).text = "S" Then ' activo fijo
                TpoTranCompra = 4
                
            End If
            If Grid1.Cell(k, 21).text <> "" And TpoTranCompra <> "4" Then '  iva uso comun
                TpoTranCompra = 5
                codigo_iva = 2
            End If
            If Grid1.Cell(k, 21).text <> "" Then
                codigo_iva = 2
            End If
            
            If Grid1.Cell(k, 20).text <> "" Then '  iva no recuperado
                TpoTranCompra = 6
                Select Case Grid1.Cell(k, 0).text
                    Case "33"
                        tipodoc = "4"
                    Case "34"
                        tipodoc = "0"
                    Case "43"
                        tipodoc = "L"
                    Case "56"
                        tipodoc = "5"
                    Case "61"
                        tipodoc = "6"
                End Select
                
                codigo_iva = buscamotivo(tipodoc, Grid1.Cell(k, 3).text, Format(Grid1.Cell(k, 4).text, "yyy-mm-dd"), Grid1.Cell(k, 5).text)
            End If
            
            
            
            cadena = cadena & TpoTranCompra & ";" 'TpoTranCompra;
            
            
            cadena = cadena & codigo_iva  'Codigo_IVA_e_Impuestos;
            Print #20, cadena
                
no:
        End If
    Next k
    
    Close 20
    Shell "NOTEPAD " + ARCHIVO
End Sub
Function buscamotivo(tipo, numero, fecha, rutprove) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select motivo from " & cliente_sql & "conta" & dato1.text & ".facturasdecompras_norecuperable "
    csql.sql = csql.sql & " where tipo='" & tipo & "' and numero='" & numero & "' and rut='" & Replace(rutprove, "-", "") & "' "
    csql.Execute
    buscamotivo = 1
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        buscamotivo = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
    
End Function

 
Sub GRABACARTOLALIBROS(tipo, numero, rutconsu, fecha, EXENTO, iva, NETO, total, loc, periodo, tipolibro)
    Dim tipo_dte2 As String
    Dim rut3 As String
    Dim dato As Variant
    If InStr(1, rutconsu, "-") = 0 Then GoTo no:
    dato = Split(rutconsu, "-")
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "rut"
    campos(3, 0) = "fecha"
    campos(4, 0) = "exento"
    campos(5, 0) = "iva"
    campos(6, 0) = "neto"
    campos(7, 0) = "total"
    campos(8, 0) = "mescontable"
    campos(9, 0) = "añocontable"
    campos(10, 0) = ""
    
'    If numero = "1419" Then Stop
    campos(0, 1) = tipo
    campos(1, 1) = Format(numero, "0000000000")
    campos(2, 1) = Format(dato(0), "000000000") & dato(1)
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = EXENTO
    campos(5, 1) = iva
    campos(6, 1) = NETO
    campos(7, 1) = total
    campos(8, 1) = Mid(periodo, 5, 2)
    campos(9, 1) = Mid(periodo, 1, 4)
    
    
    
    If Mid(tipolibro, 1, 5) = "VENTA" Then
        campos(0, 2) = clientesistema + "conta" + loc + ".sv_dte_libros_sii_ventas"
    Else
        campos(0, 2) = clientesistema + "conta" + loc + ".sv_dte_aceptados_sii_compras"
     End If
           
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
no:
End Sub

Public Function leercodigocontable2(rut) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = contadb

        csql.sql = "SELECT codigocontable  "
        csql.sql = csql.sql & "FROM " & clientesistema & "gestion.g_maestroempresas "
        csql.sql = csql.sql & "WHERE rut='" & rut & "' "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
        leercodigocontable2 = resultado(0)
        Else
        leercodigocontable2 = ""
        End If
End Function
 Public Function leercodigofae(codigo) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = contadb

        csql.sql = "SELECT empresafae  "
        csql.sql = csql.sql & "FROM " & clientesistema & "conta.maestroempresas "
        csql.sql = csql.sql & "WHERE codigoempresa='" & codigo & "' "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
        leercodigofae = resultado(0)
        Else
        leercodigofae = ""
        End If
End Function
Public Function ESTAENSII(tipo, numero, rut, monto) As Boolean
Dim multi As Integer
Dim empresafae As String

If tipo = 4 Then
 tipo = "33"
End If
If tipo = 1 Then
 tipo = "30"
End If
If tipo = 3 Then
 tipo = "60"
End If


If tipo = 5 Then
 tipo = "56"
End If
If tipo = 6 Then
 tipo = "61"
End If
If tipo = 0 Then
 tipo = "34"
End If

'numero = Val(numero)
'empresafae = CONFI_EMPRESAFAE
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT rut "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & dato1.text & ".sv_dte_aceptados_sii_compras "
        csql2.sql = csql2.sql + "Where tipo ='" + tipo + "' and numero = '" & numero & "' and rut = '" + rut + "' "
        csql2.sql = csql2.sql + "and total = '" & monto & "' "
        
        
        csql2.Execute
        ESTAENSII = False
        If csql2.RowsAffected > 0 Then
            ESTAENSII = True
        Else
            ESTAENSII = False
        End If
        
        

End Function
