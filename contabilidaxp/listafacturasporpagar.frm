VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form infoge02 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas Por Pagar"
   ClientHeight    =   5445
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   8325
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   4920
      TabIndex        =   17
      Top             =   4680
      Width           =   3255
      _ExtentX        =   5741
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
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp fechas 
      Height          =   1935
      Left            =   1440
      TabIndex        =   6
      Top             =   3360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      BackColor       =   14737632
      Caption         =   "Rangos de Fecha"
      CaptionEstilo3D =   1
      BackColor       =   14737632
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
      Begin CoolButtons.cool_Button command8 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         SkinId          =   "13"
         Caption         =   "Cambia Fecha"
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
         Left            =   2520
         TabIndex        =   11
         Top             =   720
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
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin XPFrame.FrameXp OPCIONES 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4948
      BackColor       =   16761024
      Caption         =   "Lista Facturas por Pagar"
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
      Begin VB.CommandButton Command2 
         Caption         =   "Genera Informe"
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   2280
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   1335
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
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
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
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
               Left            =   240
               TabIndex        =   5
               Text            =   "01"
               Top             =   360
               Width           =   375
            End
            Begin VB.Label empresanombre 
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
               Height          =   375
               Left            =   840
               TabIndex        =   4
               Top             =   360
               Width           =   6135
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   855
         Left            =   1170
         TabIndex        =   12
         Top             =   4725
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "TIPO DE IMPRESION"
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
         Begin VB.TextBox FOLIO 
            Height          =   285
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton timbrado 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Timbrado"
            Height          =   255
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton original 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Original"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "infoge02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private tipoprove As String
Private plan(2000, 3) As Variant
Private canplan As Double
Private total(10) As Double
Private detalle(10, 10) As Double
Private TIPOS(7) As String
Private MES As String
Private año As String




Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub



Private Sub COMMAND2_Click()
Dim TIMBRA As String

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)

Call CARGAGRILLA(infogrilla)

Call Consulta_Informe(infogrilla)


infogrilla.Visible = True
infogrilla.Caption = "FACTURAS POR PAGAR ": grillainformes.Tag = "infoge02" & TIMBRA & FOLIO.text

infogrilla.Show


End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)


End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
    
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
    
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub

Private Sub datos1_Click()

End Sub

Private Sub datos2_Click()

End Sub

Private Sub Form_Load()

CENTRAR Me

Dim i As Integer
Dim k As Integer

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO"
TIPOS(4) = "FACTURAS ELECTRONICAS"
TIPOS(5) = "NOTAS DE DEBITO ELECTRONICAS"
TIPOS(6) = "NOTAS DE CREDITO ELECTRONICAS"
TIPOS(7) = "FACTURAS ACTIVO FIJO"
    
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
For i = 1 To 10
For k = 1 To 10
detalle(k, i) = 0
Next k

Next i
opciones.Visible = True

original.Value = True

dato1.text = empresaactiva
empresanombre.Caption = nombreempresa
fechas.Visible = False

End Sub


    
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim tip As String

    Dim PASO As String
    tip = "1"
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fc.tipo,numero,fecha,fc.rut,ifnull(cc.nombre,''),fechavencimiento,total,abono,total-abono "
        csql.sql = csql.sql + "FROM facturasdecompras as fc left join cuentascorrientes as cc on (cc.tipo='" + CUENTAPROVEEDOR + "' and cc.rut=fc.rut and cc.año='" + Format(fechasistema, "yyyy") + "') "
        csql.sql = csql.sql + "where total<>abono order by cc.nombre,fc.tipo,fc.fecha "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        
        Set resultados = csql.OpenResultset
        infogrilla.Grid1.Rows = 1
        
         While Not resultados.EOF
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            barra.Value = barra.Value + 1
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 1).text = resultados(0)
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 2).text = resultados(1)
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 3).text = resultados(2)
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 4).text = resultados(3)
            If IsNull(resultados(4)) = False Then
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 5).text = resultados(4)
            Else
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 5).text = "** no creado **"
            End If
            If IsNull(resultados(5)) = False Then
            
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 6).text = resultados(5)
            End If
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 7).text = Format(resultados(6), "###,###,###,###")
            
             
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
     
Call totallibro(infogrilla)
barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub

Sub totallibro(infogrilla As grillainformes)
    
    Dim TOTALge As Double

        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        lin = infogrilla.Grid1.Rows - 1
        infogrilla.Grid1.Range(lin, 6, lin, 9).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 6).text = "TOTALES"
        infogrilla.Grid1.Cell(lin, 7).text = total(1)
        infogrilla.Grid1.Cell(lin, 8).text = total(2)
        infogrilla.Grid1.Cell(lin, 9).text = total(3)
End Sub
'    TOTALge = 0
'    lin = lin + 2
'    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 9
'    infogrilla.Grid1.Range(lin, 5, lin + 7, 12).Borders(cellEdgeTop) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 7, 12).Borders(cellEdgeLeft) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 7, 12).Borders(cellEdgeRight) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 7, 12).Borders(cellEdgeBottom) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 7, 12).Borders(cellInsideHorizontal) = cellThin
'    infogrilla.Grid1.Range(lin, 5, lin + 7, 12).Borders(cellInsideVertical) = cellThin
'
'    infogrilla.Grid1.Cell(lin, 5).text = "Cant."
'    infogrilla.Grid1.Cell(lin, 6).text = "Documentos"
'    infogrilla.Grid1.Cell(lin, 7).text = "Neto"
'    infogrilla.Grid1.Cell(lin, 8).text = "i.v.a"
'    infogrilla.Grid1.Cell(lin, 9).text = "exento"
'    infogrilla.Grid1.Cell(lin, 10).text = "diesel"
'    infogrilla.Grid1.Cell(lin, 11).text = "retencion"
'    infogrilla.Grid1.Cell(lin, 12).text = "total"
'
'
'
'    For K = 1 To 7
'    lin = lin + 1
'
'    infogrilla.Grid1.Cell(lin, 6).text = TIPOS(K)
'    infogrilla.Grid1.Cell(lin, 5).text = Format(detalle(K, 1), "###,###,##0")
'    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(K, 2), "###,###,##0")
'    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(K, 3), "###,###,##0")
'    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(K, 4), "###,###,##0")
'    infogrilla.Grid1.Cell(lin, 10).text = Format(detalle(K, 5), "###,###,##0")
'    infogrilla.Grid1.Cell(lin, 11).text = Format(detalle(K, 6), "###,###,##0")
'    infogrilla.Grid1.Cell(lin, 12).text = Format(detalle(K, 7), "###,###,##0")
'
'    Next K
'    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
'    lin = lin + 2
'    For K = 1 To canplan
'    If plan(K, 3) <> 0 Then
'             lin = lin + 1
'             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
'        infogrilla.Grid1.Cell(lin, 5).text = plan(K, 1)
'        infogrilla.Grid1.Cell(lin, 6).text = plan(K, 2)
'        infogrilla.Grid1.Cell(lin, 7).text = plan(K, 3)
'        TOTALge = TOTALge + plan(K, 3)
'        End If
'    Next K
'        lin = lin + 1
'             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
'        infogrilla.Grid1.Range(lin, 6, lin, 7).Borders(cellEdgeTop) = cellThin
'
'
'
'
'
'        infogrilla.Grid1.Cell(lin, 6).text = "TOTAL DETALLE"
'         infogrilla.Grid1.Cell(lin, 7).text = TOTALge
               
   





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "RUT"
    FORMATOGRILLA(1, 5) = "PROVEEDOR"
    FORMATOGRILLA(1, 6) = "VENCIMIENTO"
    FORMATOGRILLA(1, 7) = "TOTAL"
    FORMATOGRILLA(1, 8) = "ABONO"
    FORMATOGRILLA(1, 9) = "SALDO"
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "4"
    FORMATOGRILLA(2, 2) = "8"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "20"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "D"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "D"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    Rem LOCCKED
    For k = 1 To 14
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 10
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub leermayor()
    tipoprove = CUENTAPROVEEDOR
    

    
End Sub

'Sub Consultadetalle(MES As String, año As String)
'Dim multi As Integer
'
'Dim resultados2 As rdoResultset
'    Dim cSql2 As New rdoQuery
'        Set cSql2.ActiveConnection = contadb
'        cSql2.SQL = "SELECT cuentadelmayor,dfc.tipo,sum(dfc.monto)"
'        cSql2.SQL = cSql2.SQL + "FROM facturasdecompras as fc,detallefacturasdecompra as dfc "
'        cSql2.SQL = cSql2.SQL + "where añocontable='" + año + "' and mescontable='" + MES + "'"
'        cSql2.SQL = cSql2.SQL + " and fc.tipo=dfc.tipo and fc.numero=dfc.numero and fc.rut=dfc.rut"
'        cSql2.SQL = cSql2.SQL + " group by cuentadelmayor,dfc.tipo "
'
'        cSql2.Execute
'
'
'        If cSql2.RowsAffected > 0 Then
'        Set resultados2 = cSql2.OpenResultset
'
'         While Not resultados2.EOF
'         For K = 1 To canplan
'         If resultados2(1) = "3" Then multi = -1 Else multi = 1
'         If resultados2(0) = plan(K, 1) Then plan(K, 3) = plan(K, 3) + (resultados2(2) * multi): infogrilla.Grid1.Cell(lin, 11).text = plan(K, 2): K = canplan + 1
'         Next K
'          resultados2.MoveNext
'
'
'         Wend
'
'          resultados2.Close
'
'        End If
'
'End Sub
Sub CARGAmayor()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo "
        csql.sql = csql.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             plan(LINEA, 1) = resultados(0)
             plan(LINEA, 2) = resultados(1)
             plan(LINEA, 3) = 0

            resultados.MoveNext
            Wend
        End If
canplan = LINEA
   

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

