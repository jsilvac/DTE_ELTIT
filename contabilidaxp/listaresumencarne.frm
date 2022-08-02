VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form infocarne 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Resumen Carnes"
   ClientHeight    =   5685
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   5700
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   45
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   675
      Visible         =   0   'False
      Width           =   555
   End
   Begin XPFrame.FrameXp OPCIONES 
      Height          =   5430
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9578
      BackColor       =   16761024
      Caption         =   "Lista Resumen harinas"
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Con Totales"
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
         Left            =   135
         TabIndex        =   12
         Top             =   4095
         Width           =   2310
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sin Totales"
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
         Left            =   135
         TabIndex        =   11
         Top             =   3780
         Width           =   1995
      End
      Begin CoolButtons.cool_Button COMMAND2 
         Height          =   495
         Left            =   2520
         TabIndex        =   8
         Top             =   3870
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Genera Informe"
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   5040
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   3315
         Left            =   450
         TabIndex        =   2
         Top             =   315
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5847
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
               TabIndex        =   10
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
               TabIndex        =   9
               Top             =   360
               Width           =   3255
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   855
            Left            =   120
            TabIndex        =   4
            Top             =   1320
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
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   6
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   855
            Left            =   120
            TabIndex        =   5
            Top             =   2280
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
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   7
               Top             =   360
               Width           =   3855
            End
         End
      End
      Begin CoolButtons.cool_Button ARCHIVO 
         Height          =   495
         Left            =   2520
         TabIndex        =   13
         Top             =   4455
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Genera Archivo"
      End
   End
End
Attribute VB_Name = "infocarne"
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
Private TIPOS(6) As String
Private MES As String
Private año As String
Private centro As String



Private Sub ARCHIVO_Click()

Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)

año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)


generaarchivo

End Sub

Private Sub COMMAND2_Click()
Dim TIMBRA As String



Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)

año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

CARGAmayor
leermayor
Call CARGAGRILLA(infogrilla)
For k = 1 To 10
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
Next k
Call Consulta_Informe(infogrilla)
infogrilla.Visible = True
infogrilla.Caption = "ANEXO INFORME MENSUAL VENDEDORES DE CARNE": grillainformes.Tag = "INFOCARNE"
infogrilla.Show
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
    COMBOMES.SetFocus
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub



Private Sub Form_Load()

CENTRAR Me

Dim i As Integer
Dim k As Integer

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO FACTURAS"
TIPOS(4) = "NOTAS DE CREDITO BOLETAS"
    
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
For i = 1 To 10
For k = 1 To 10
detalle(k, i) = 0
Next k

Next i
opciones.Visible = True
Option1.Value = True


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


    
    

fechas.Visible = False


End Sub


    
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim total As Double
    

    Dim PASO As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT cc.rut,smc.nombre,'1',round(((sum(if(fc.tipo='1' or fc.tipo='6',fd.monto,fd.monto*-1))/2300)/.05),0),round((SUM(fd.monto)/.05),0),round(((SUM(if(fc.tipo='1' or fc.tipo='6',fd.monto,fd.monto*-1))/.05))*.05,0) "
        csql.sql = csql.sql + "FROM facturasdeventas as fc,facturasdeventas_detalle as fd,cuentascorrientes as cc , " + clientesistema + "ventas.sv_maestroclientes as smc "
        csql.sql = csql.sql + "where fc.rut=smc.rut and fc.rut=cc.rut and cc.tipo='" + cuentacliente + "' and fecha >= '" + año + "/" + MES + "/" + "01" + "' and fecha <= '" + año + "/" + MES + "/" + "31' "
        csql.sql = csql.sql + "and fc.tipo=fd.tipo and fc.numero=fd.numero and fd.cuentadelmayor='23200009' and cc.año='" + COMBOAÑO.text + "' and (fc.tipo='1' or fc.tipo='6') "
        csql.sql = csql.sql + "group by cc.rut order by cc.rut"
        
        csql.Execute
        
        infogrilla.Grid1.AutoRedraw = False
        total = 0
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        infogrilla.Grid1.Rows = csql.RowsAffected + 1
        
        Set resultados = csql.OpenResultset
        lin = 0
         While Not resultados.EOF
    
         
             barra.Value = lin
             lin = lin + 1
             
             infogrilla.Grid1.Cell(lin, 1).text = resultados(0)
             infogrilla.Grid1.Cell(lin, 2).text = resultados(1)
             infogrilla.Grid1.Cell(lin, 3).text = resultados(2)
             infogrilla.Grid1.Cell(lin, 4).text = resultados(3)
             infogrilla.Grid1.Cell(lin, 5).text = resultados(4)
             infogrilla.Grid1.Cell(lin, 6).text = resultados(5)
             
             total = total + resultados(5)
            
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
If Option1.Value = False Then

infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
infogrilla.Grid1.Range(infogrilla.Grid1.Rows - 1, 3, infogrilla.Grid1.Rows - 1, 4).Borders(cellEdgeTop) = cellThin




infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 3).text = "TOTAL RETENCIONES "
infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 6).text = total
End If


barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub



Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 9
    
    
    FORMATOGRILLA(1, 1) = "RUT"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = " "
    FORMATOGRILLA(1, 4) = "KILOS"
    FORMATOGRILLA(1, 5) = "NETO"
    FORMATOGRILLA(1, 6) = "RETENCION"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "25"
    FORMATOGRILLA(2, 3) = "12"
    FORMATOGRILLA(2, 4) = "12"
    FORMATOGRILLA(2, 5) = "12"
    FORMATOGRILLA(2, 6) = "12"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 4) = "###,###.00"
    FORMATOGRILLA(4, 5) = "###,###,###"
    FORMATOGRILLA(4, 6) = "###,###,###"
   
    Rem LOCCKED
    For k = 1 To 6
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 7
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
        csql.sql = csql.sql + "FROM cuentasdelmayor"
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


Sub generaarchivo()
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim total As Double
    Dim n1 As String * 4
    Dim n2 As String * 4
    Dim n3 As String * 9
    Dim n4 As String * 35
    Dim n5 As String * 35
    Dim n6 As String * 18
    Dim n7 As String * 7
    Dim n8 As String * 9
    Dim n9 As String * 35
    Dim n10 As String * 1
    Dim n11 As String * 10
    Dim n12 As String * 12
    Dim n13 As String * 12
    Dim n14 As String * 14
    Dim CARNE As String
    
    n1 = "3260"
    n2 = MES + Mid(año, 3, 2)
    n3 = Mid(rutempresa, 1, 8) + Mid(rutempresa, 10, 1)
    n4 = nombreempresa
    n5 = direccionempresa
    n6 = comunaempresa
    n7 = codigosii
    
    
    
    
    Dim PASO As String
        Set csql.ActiveConnection = contadb
'       cSql.sql = "SELECT cc.rut,smc.nombre,smc.direccion,smc.ciudad,sum(fd.monto) "
'        cSql.sql = cSql.sql + "FROM facturasdeventas as fc,facturasdeventas_detalle as fd,cuentascorrientes as cc , " + clientesistema + "ventas.sv_maestroclientes as smc "
'        cSql.sql = cSql.sql + "where fc.rut=smc.rut and fc.rut=cc.rut and cc.tipo='" + cuentacliente + "' and fecha >= '" + año + "/" + mes + "/" + "01" + "' and fecha <= '" + año + "/" + mes + "/" + "31' "
'        cSql.sql = cSql.sql + "and fc.tipo=fd.tipo and fc.numero=fd.numero and fd.cuentadelmayor='11400005' "
'        cSql.sql = cSql.sql + "group by cc.rut order by cc.rut"
'
        
        csql.sql = "SELECT cc.rut,smc.nombre,'1',round(((sum(fd.monto)/2300)/.05),0),round((SUM(fd.monto)/.05),0),round(((SUM(fd.monto)/.05))*.05,0) "
        csql.sql = csql.sql + "FROM facturasdeventas as fc,facturasdeventas_detalle as fd,cuentascorrientes as cc , " + clientesistema + "ventas.sv_maestroclientes as smc "
        csql.sql = csql.sql + "where fc.rut=smc.rut and fc.rut=cc.rut and cc.tipo='" + cuentacliente + "' and fecha >= '" + año + "/" + MES + "/" + "01" + "' and fecha <= '" + año + "/" + MES + "/" + "31' "
        csql.sql = csql.sql + "and fc.tipo=fd.tipo and fc.numero=fd.numero and fd.cuentadelmayor='23200009' "
        csql.sql = csql.sql + "group by cc.rut order by cc.rut"
        csql.Execute
        
        total = 0
       Close 20
       Open App.path + "\carnes.txt" For Output As #20
        If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
        
        n8 = Mid(resultados(0), 2, 9)
        n9 = resultados(1)
        n10 = "1"
        pivote.MaxLength = 10
        pivote.text = resultados(3)
        Call ceros(pivote)
        
        n11 = pivote.text
        pivote.MaxLength = 12
        pivote.text = resultados(4)
        Call ceros(pivote)
        n12 = pivote.text
        pivote.MaxLength = 12
        pivote.text = resultados(5)
        Call ceros(pivote)
        n13 = pivote.text
        n14 = String(14, 32)
        
        CARNE = n1 + n2 + n3 + n4 + n5 + n6 + n7 + n8 + n9 + n10 + n11 + n12 + n13 + n14
        Print #20, CARNE
             
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
Close 20

Shell "NOTEPAD " + App.path + "\CARNES.TXT"
End Sub


