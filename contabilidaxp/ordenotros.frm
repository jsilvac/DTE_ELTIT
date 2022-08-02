VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ordenotros 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7755
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmDatos 
      Height          =   7545
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   13309
      BackColor       =   12648384
      Caption         =   "OTROS DESCUENTOS O HABERES"
      BackColor       =   12648384
      ColorBarraArriba=   16576
      ColorBarraAbajo =   16576
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
      Begin VB.TextBox LBLDV 
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
         Height          =   300
         Left            =   12120
         MaxLength       =   1
         TabIndex        =   24
         Top             =   6435
         Width           =   435
      End
      Begin VB.TextBox DATO5 
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
         Height          =   300
         Left            =   10320
         MaxLength       =   9
         TabIndex        =   23
         Top             =   6435
         Width           =   1650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CARGA PUBLICIDAD"
         Height          =   465
         Left            =   3360
         TabIndex        =   22
         Top             =   6960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CARGA DEVOLUCIONES"
         Height          =   465
         Left            =   900
         TabIndex        =   19
         Top             =   6960
         Width           =   2175
      End
      Begin VB.TextBox dato2 
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
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   14
         Top             =   6435
         Width           =   6135
      End
      Begin VB.TextBox numero 
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
         Left            =   2565
         MaxLength       =   10
         TabIndex        =   12
         Top             =   585
         Width           =   1365
      End
      Begin VB.TextBox dato3 
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
         Left            =   7965
         MaxLength       =   9
         TabIndex        =   10
         Top             =   6435
         Width           =   1650
      End
      Begin VB.TextBox dato4 
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
         Left            =   9675
         MaxLength       =   1
         TabIndex        =   2
         Top             =   6435
         Width           =   435
      End
      Begin VB.TextBox dato1 
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
         Left            =   135
         MaxLength       =   8
         TabIndex        =   0
         Top             =   6435
         Width           =   1635
      End
      Begin XPFrame.FrameXp frmLista 
         Height          =   4875
         Left            =   45
         TabIndex        =   4
         Top             =   1080
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   8599
         BackColor       =   12648447
         Caption         =   "Lista de Otros"
         CaptionEstilo3D =   1
         BackColor       =   12648447
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   32896
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
            Height          =   4395
            Left            =   45
            TabIndex        =   3
            Top             =   360
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   7752
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
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   13005
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
      Begin VB.Label lblcuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9480
         TabIndex        =   25
         Top             =   6840
         Width           =   3735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto ORDEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10665
         TabIndex        =   21
         Top             =   405
         Width           =   2460
      End
      Begin VB.Label montoorden 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   10665
         TabIndex        =   20
         Top             =   720
         Width           =   2445
      End
      Begin VB.Label MONTOOTROS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   18
         Top             =   7110
         Width           =   2985
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6120
         TabIndex        =   17
         Top             =   6795
         Width           =   3000
      End
      Begin VB.Label lblNOMBRE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9480
         TabIndex        =   16
         Top             =   7200
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CUENTA CORRIENTE"
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
         Left            =   10215
         TabIndex        =   15
         Top             =   6075
         Width           =   3015
      End
      Begin VB.Label LBLPROVEEDOR 
         BackStyle       =   0  'Transparent
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
         Height          =   420
         Left            =   4140
         TabIndex        =   13
         Top             =   585
         Width           =   6315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO ORDEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   630
         Width           =   2715
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D/H"
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
         Left            =   9675
         TabIndex        =   9
         Top             =   6075
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
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
         Left            =   7965
         TabIndex        =   8
         Top             =   6075
         Width           =   1650
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Glosa"
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
         Left            =   1815
         TabIndex        =   7
         Top             =   6075
         Width           =   6120
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
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
         Left            =   135
         TabIndex        =   6
         Top             =   6075
         Width           =   1650
      End
   End
End
Attribute VB_Name = "ordenotros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private FORMATOGRILLA(10, 10) As String
    Private modifica As Boolean
    Private LINEAS As Double
    
Private Sub Command1_Click()
grabardevoluciones
End Sub

Private Sub COMMAND2_Click()
grabarpublicidad
End Sub

Sub grabardevoluciones()
        Dim suma As Double
        Dim rutpro As String
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        
        Dim tabla As String
Set csql.ActiveConnection = contadb
rutpro = Mid(prove0002.Grid1.Cell(prove0002.Grid1.ActiveCell.row, 3).text, 1, 9) + Mid(prove0002.Grid1.Cell(prove0002.Grid1.ActiveCell.row, 3).text, 11, 1)
csql.sql = "select dp.fecha,dp.numero,dp.monto,dp.tipo,dp.numero "
csql.sql = csql.sql & "from devoluciones_proveedores as dp left join cuentascorrientes as cc on (dp.rut=cc.rut and cc.tipo='" + CUENTAPROVEEDOR + "' AND cc.año='" + Format(fechasistema, "yyyy") + "') "
csql.sql = csql.sql & "where dp.rut='" & rutpro & "' and dp.montoco='0' and dp.local='" + localorden + "' "
csql.sql = csql.sql & "order by cc.nombre "
csql.Execute

If csql.RowsAffected > 0 Then
  
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        If guiarebajada(resultados(3), resultados(4), clientesistema + "gestion" + leerdatoslocal(localorden, "rubro") + ".l_ordendecompra_anexopagos_" + localorden) = False Then
        LINEA = LINEA + 1
          ' agregado por el numero fiscal 03-10-2017
            Dim numerosii As String
            If resultados(3) = "D1" Then
            numerosii = Format(leerFOLIOSIIDTE(localorden, "G4", resultados(1), resultados(0), "99", localorden), "0000000000")
            If numerosii = "" Then numerosii = "SIN FOLIO FISCAL"
            
            Call grabarEspecialesguias(numero, LINEA, "11200044", "GUIA DEVO.ELEC. " & numerosii & " DEL " & Format(resultados(0), "dd-mm-yyyy"), resultados(2), "H", resultados(3), resultados(4), rutpro)
            End If
            If resultados(3) = "DM" Then
            Call grabarEspecialesguias(numero, LINEA, "11200044", "GUIA DEVOLUCION " & resultados(1) & " DEL " & Format(resultados(0), "dd-mm-yyyy"), resultados(2), "H", resultados(3), resultados(4), rutpro)
            End If
            
'             original 03-10-2017
'             Call grabarEspecialesguias(numero, LINEA, "11200044", "GUIA DEVOLUCION " & resultados(1) & " DEL " & Format(resultados(0), "dd-mm-yyyy"), resultados(2), "H", resultados(3), resultados(4), rutpro)
       End If
        resultados.MoveNext
    Wend
End If

csql.Close
Set csql = Nothing
Set resultados = Nothing
leerEspeciales

End Sub
Public Function leerFOLIOSIIDTE(empre, tipo, numero, fecha, caja, loc) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Set csql.ActiveConnection = contadb
    csql.sql = "select numero "
'    If loc = "05" Then
'        csql.sql = csql.sql & " from " & clientesistema & "fae" & loc & ".sv_dte" & loc & "_res where tipodocumento='" & tipo & "' and "
'    Else
        csql.sql = csql.sql & " from " & clientesistema & "fae" & loc & ".sv_dte" & loc & " where tipodocumento='" & tipo & "' and "

'    End If
    csql.sql = csql.sql & "localdocumento='" & loc & "' and numerodocumento='" & numero & "' and "
    csql.sql = csql.sql & " fechadocumento='" & Format(fecha, "yyyy-mm-dd") & "' and "
    csql.sql = csql.sql & " cajadocumento='" & caja & "' "
    csql.Execute
    leerFOLIOSIIDTE = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerFOLIOSIIDTE = resultados(0)
        resultados.Close
        Set resultados = Nothing
    End If
    csql.Close
    Set csql = Nothing
    
End Function
Sub grabarpublicidad()
        Dim suma As Double
        Dim rutpro As String
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        
        Dim tabla As String
Set csql.ActiveConnection = contadb
rutpro = Mid(prove0002.Grid1.Cell(prove0002.Grid1.ActiveCell.row, 3).text, 1, 9) + Mid(prove0002.Grid1.Cell(prove0002.Grid1.ActiveCell.row, 3).text, 11, 1)
csql.sql = "select dp.fecha,if(dp.tipo='1',dp.numero,dp.foliosii),dp.total-dp.abono,dp.tipo,if(dp.tipo='1',dp.numero,dp.foliosii) "
csql.sql = csql.sql & "from facturasdepublicidad as dp left join cuentascorrientes as cc on (dp.rut=cc.rut and cc.tipo='" + CUENTAPROVEEDOR + "' AND cc.año='" + Format(fechasistema, "yyyy") + "') "
csql.sql = csql.sql & "where dp.rut='" & rutpro & "' and dp.abono<dp.total and dp.fecha>'2014-01-01' "
csql.sql = csql.sql & "order by cc.nombre "
csql.Execute

If csql.RowsAffected > 0 Then
  
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        If facturarebajada(resultados(3), resultados(4), clientesistema + "gestion" + leerdatoslocal(localorden, "rubro") + ".l_ordendecompra_anexopagos_" + localorden) = False Then
    
        LINEA = LINEA + 1
        If resultados(3) = "1" Then
        Call grabarEspecialespublicidad(numero, LINEA, "11200028", "FACTURA MANUAL " & resultados(1) & " DEL " & Format(resultados(0), "dd-mm-yyyy"), resultados(2), "H", resultados(3), resultados(4), rutpro)
        Else
        Call grabarEspecialespublicidad(numero, LINEA, "11200028", "FACTURA PUBLICIDAD " & resultados(1) & " DEL " & Format(resultados(0), "dd-mm-yyyy"), resultados(2), "H", resultados(3), resultados(4), rutpro)
        
        End If
        
        End If
        
        resultados.MoveNext
    Wend
End If

csql.Close
Set csql = Nothing
Set resultados = Nothing
leerEspeciales

End Sub


Public Sub grabarEspecialesguias(numero, LINEA, cuenta, glosa, monto, DH, TIPODO, numerodo, rut)
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "numero"
        campos(1, 0) = "linea"
        campos(2, 0) = "cuenta"
        campos(3, 0) = "glosa"
        campos(4, 0) = "monto"
        campos(5, 0) = "dh"
        campos(6, 0) = "tipodo"
        campos(7, 0) = "numerodo"
        campos(8, 0) = "rut"
        
        campos(9, 0) = ""
        
        campos(0, 1) = numero
        campos(1, 1) = LINEA
        campos(2, 1) = cuenta
        campos(3, 1) = glosa
        campos(4, 1) = monto
        campos(5, 1) = DH
        campos(6, 1) = TIPODO
        campos(7, 1) = numerodo
        campos(8, 1) = rut
        
        campos(0, 2) = clientesistema + "gestion" + rubro + ".l_ordendecompra_anexopagos_" & localorden
        
        
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
    End Sub
 Public Sub grabarEspecialespublicidad(numero, LINEA, cuenta, glosa, monto, DH, TIPODO, numerodo, rut)
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "numero"
        campos(1, 0) = "linea"
        campos(2, 0) = "cuenta"
        campos(3, 0) = "glosa"
        campos(4, 0) = "monto"
        campos(5, 0) = "dh"
        campos(6, 0) = "tipodo"
        campos(7, 0) = "numerodo"
        campos(8, 0) = "rut"
        campos(9, 0) = ""
        
        campos(0, 1) = numero
        campos(1, 1) = LINEA
        campos(2, 1) = cuenta
        campos(3, 1) = glosa
        campos(4, 1) = monto
        campos(5, 1) = DH
        campos(6, 1) = TIPODO
        campos(7, 1) = numerodo
        campos(8, 1) = rut
        
        campos(0, 2) = clientesistema & "gestion" & rubro & ".l_ordendecompra_anexopagos_" & localorden
                
        
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
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
    
Private Sub DATO2_LostFocus()
If dato2.text = "" And dato1.text = "" Then
dato2.SetFocus
End If

End Sub

    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub

Private Sub DATO3_LostFocus()
If dato3.text = "" And dato2.text <> "" Then
dato3.SetFocus

End If

End Sub


    Private Sub dato4_GotFocus()
        Call cargatexto(dato4)
    End Sub
    
    
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            
            Call ayudamayor(dato2)
            
        Else
            Call flechas(dato1, dato2, KeyCode)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato1, dato3, KeyCode)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato2, dato3, KeyCode)
    End Sub
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato3, dato4, KeyCode)
    End Sub
    
    
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    lblcuenta.Caption = leerdatos(contadb, "cuentasdelmayor", "nombre", "codigo='" + dato1.text + "' and año='" + año + "' ")
    If lblcuenta.Caption <> "" Then
    dato2.SetFocus
       
    Else
    MsgBox ("codigo de cuenta no existe")
    dato1.SetFocus
    
    End If
    
    End If
    
    End Sub
    Private Sub dato2_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
    dato3.SetFocus
    End If
    
    End Sub
    Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    dato4.SetFocus
    End If
    
    End Sub
    Private Sub dato4_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 45 Then KeyAscii = 68
    If KeyAscii = 43 Then KeyAscii = 72
    If KeyAscii = 13 And dato4.text = "D" Or dato4.text = "H" Then
    If modifica = False Then
    
    
    Call grabarEspeciales(LINEAS)
    Else
    Call modificaEspeciales(lista.ActiveCell.row)
    modifica = False
    
    End If
    
    leerEspeciales
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    LBLDV.text = ""
    
    dato1.SetFocus
    lblcuenta.Caption = ""
    
    
    End If
    End Sub
    
    
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
            If leerNombreCuentaMayor(dato1.text, "1") <> "" Then
            dato5.SetFocus
            
            End If
            
            
    End Sub
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
        col = 9
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, 1) = "CUENTA"
        FORMATOGRILLA(1, 2) = "GLOSA"
        FORMATOGRILLA(1, 3) = "MONTO"
        FORMATOGRILLA(1, 4) = "D/H"
        FORMATOGRILLA(1, 5) = "NOMBRE CUENTA"
        FORMATOGRILLA(1, 6) = "RUT"
        FORMATOGRILLA(1, 7) = "TIPO"
        FORMATOGRILLA(1, 8) = "DOCUMENTO"
        
        Rem LARGO DE LOS DATOS
        FORMATOGRILLA(2, 1) = "8"
        FORMATOGRILLA(2, 2) = "60"
        FORMATOGRILLA(2, 3) = "10"
        FORMATOGRILLA(2, 4) = "1"
        FORMATOGRILLA(2, 5) = "20"
        FORMATOGRILLA(2, 6) = "10"
        FORMATOGRILLA(2, 7) = "2"
        FORMATOGRILLA(2, 8) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        FORMATOGRILLA(3, 1) = "N"
        FORMATOGRILLA(3, 2) = "S"
        FORMATOGRILLA(3, 3) = "N"
        FORMATOGRILLA(3, 4) = "N"
        FORMATOGRILLA(3, 5) = "S"
        FORMATOGRILLA(3, 6) = "S"
        
        Rem FORMATO GRILLA
        FORMATOGRILLA(4, 1) = ""
        FORMATOGRILLA(4, 2) = ""
        FORMATOGRILLA(4, 3) = "########0"
        FORMATOGRILLA(4, 4) = ""
        FORMATOGRILLA(4, 5) = ""
        FORMATOGRILLA(4, 6) = ""
        
        Rem LOCCKED
        FORMATOGRILLA(5, 1) = "TRUE"
        FORMATOGRILLA(5, 2) = "TRUE"
        FORMATOGRILLA(5, 3) = "TRUE"
        FORMATOGRILLA(5, 4) = "TRUE"
        FORMATOGRILLA(5, 5) = "TRUE"
        FORMATOGRILLA(5, 6) = "TRUE"
        FORMATOGRILLA(5, 7) = "TRUE"
        FORMATOGRILLA(5, 8) = "TRUE"
        
        Rem VALOR MINIMO
        FORMATOGRILLA(6, 1) = ""
        FORMATOGRILLA(6, 2) = ""
        FORMATOGRILLA(6, 3) = ""
        FORMATOGRILLA(6, 4) = ""
        
        Rem VALOR MAXIMO
        FORMATOGRILLA(7, 1) = ""
        FORMATOGRILLA(7, 2) = ""
        FORMATOGRILLA(7, 3) = ""
        FORMATOGRILLA(7, 4) = ""
        
        Rem ANCHO
        FORMATOGRILLA(8, 1) = "8"
        FORMATOGRILLA(8, 2) = "25"
        FORMATOGRILLA(8, 3) = "10"
        FORMATOGRILLA(8, 4) = "5"
        FORMATOGRILLA(8, 5) = "15"
        FORMATOGRILLA(8, 6) = "15"
        FORMATOGRILLA(8, 7) = "2"
        FORMATOGRILLA(8, 8) = "15"
            
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
    Private Sub leerEspeciales()
        Dim suma As Double
        
        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        LINEAS = 0
        
        tabla = "SELECT cuenta,glosa,monto,dh,linea,rut,tipodo,numerodo "
        tabla = tabla & "FROM " + clientesistema + "gestion" + rubro + ".l_ordendecompra_anexopagos_" + localorden + " "
        tabla = tabla & "WHERE numero= '" & numero.text & "' ORDER BY linea asc "
        sql.sql = tabla
        sql.Execute
        
        suma = 0
        lista.Rows = 1
        lista.AutoRedraw = False
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            
            While Not resultados.EOF
                LINEAS = LINEAS + 1
                lista.Rows = lista.Rows + 1
                lista.Cell(lista.Rows - 1, 1).text = resultados(0)
                lista.Cell(lista.Rows - 1, 2).text = resultados(1)
                lista.Cell(lista.Rows - 1, 3).text = resultados(2)
                lista.Cell(lista.Rows - 1, 4).text = resultados(3)
                lista.Cell(lista.Rows - 1, 5).text = leerNombreMayor(resultados(0))
                lista.Cell(lista.Rows - 1, 6).text = resultados(5)
                lista.Cell(lista.Rows - 1, 7).text = resultados(6)
                lista.Cell(lista.Rows - 1, 8).text = resultados(7)
                
                
                If resultados(3) = "D" Then
                suma = suma + resultados(2)
                Else
                suma = suma - resultados(2)
                End If
                
                If LINEAS <> resultados(4) Then
                Call modificalineas(numero.text, resultados(4), LINEAS)
                End If
                resultados.MoveNext
            Wend
        lista.AutoRedraw = True
        lista.Refresh
        End If
    
    MONTOOTROS.Caption = Format(suma, "###,###,###")
    prove0002.Grid1.Cell(prove0002.Grid1.ActiveCell.row, prove0002.Grid1.ActiveCell.col).text = Format(suma, "###,###,###")
    End Sub
'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub grabarEspeciales(LINEA)
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "numero"
        campos(1, 0) = "linea"
        campos(2, 0) = "cuenta"
        campos(3, 0) = "glosa"
        campos(4, 0) = "monto"
        campos(5, 0) = "dh"
        campos(6, 0) = "rut"
        campos(7, 0) = ""
        campos(0, 1) = numero
        campos(1, 1) = LINEA + 1
        campos(2, 1) = dato1.text
        campos(3, 1) = dato2.text
        campos(4, 1) = dato3.text
        campos(5, 1) = dato4.text
        campos(6, 1) = dato5.text + LBLDV.text
        
        campos(0, 2) = clientesistema + "gestion" + rubro + ".l_ordendecompra_anexopagos_" & localorden
        
        
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
    End Sub
 
Public Sub modificalineas(numero, LINEA, lineanueva)
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "linea"
        campos(1, 0) = ""
        
        campos(0, 1) = lineanueva
        
        campos(0, 2) = clientesistema & "gestion" & rubro & ".l_ordendecompra_anexopagos_" & localorden
        
        
        
        condicion = "numero='" + numero + "' and linea='" & LINEA & "' "
        op = 3
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
    
    End Sub

'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub modificaEspeciales(LINEA)
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "cuenta"
        campos(1, 0) = "glosa"
        campos(2, 0) = "monto"
        campos(3, 0) = "dh"
        campos(4, 0) = "rut"
        
        campos(5, 0) = ""
        
        campos(0, 1) = dato1.text
        campos(1, 1) = dato2.text
        campos(2, 1) = dato3.text
        campos(3, 1) = dato4.text
        campos(4, 1) = dato5.text + LBLDV.text
        
        campos(0, 2) = clientesistema & "gestion" & rubro & ".l_ordendecompra_anexopagos_" & localorden
        
        
        
        condicion = "numero = '" & numero.text & "' AND linea = '" & LINEA & "'"
        op = 3
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
        modifica = False
    End Sub
'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================
    Private Sub eliminarEspeciales(LINEA)
        Dim condicion As String
        Dim campos(1, 3) As String
        Dim op As Integer
        campos(0, 2) = clientesistema + "gestion" + rubro + ".l_ordendecompra_anexopagos_" & localorden
        
        
        condicion = "numero = '" & numero.text & "' AND linea = '" & LINEA & "'"
        
        op = 4
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
    End Sub


Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato5.text = Format(dato5.text, "000000000")
LBLDV.text = rut(dato5.text)


If leerNombrerut(dato1.text, dato5.text + LBLDV.text) <> "" Then
lblnombre.Caption = leerNombrerut(dato1.text, dato5.text + LBLDV.text)

dato2.SetFocus
Else
MsgBox "numero de cuenta corriente no existe "
dato5.SetFocus

End If


End If

End Sub

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================

Private Sub Form_Activate()
cargaLista

End Sub

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
        Call CargaGrillaLista(1, 6)
        modifica = False
        leerEspeciales
    End Sub
    
Private Sub Form_Unload(Cancel As Integer)
If MONTOOTROS.Caption = "" Then MONTOOTROS.Caption = "0"
If montoorden.Caption = "" Then montoorden.Caption = "0"

If CDbl(MONTOOTROS.Caption) * -1 > CDbl(montoorden.Caption) Then
MsgBox "IMPOSIBLE REBAJAR MONTOS MAYORES AL TOTAL DE LA ORDEN "
Cancel = 1
Else
prove0002.leer

End If
End Sub

    Private Sub frmCerrar_BarClick()
        
        frmCerrar.CaptionEstilo3D = Inserted
    
    
        
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        frmCerrar.CaptionEstilo3D = RAISED
    End Sub

    Private Sub lista_DblClick()
        dato1.text = lista.Cell(lista.ActiveCell.row, 1).text
        dato2.text = lista.Cell(lista.ActiveCell.row, 2).text
        dato3.text = lista.Cell(lista.ActiveCell.row, 3).text
        dato4.text = lista.Cell(lista.ActiveCell.row, 4).text
        dato5.text = Mid(lista.Cell(lista.ActiveCell.row, 6).text, 1, 9)
        LBLDV.text = Mid(lista.Cell(lista.ActiveCell.row, 6).text, 10, 1)
        
        modifica = True
        dato3.SetFocus
    End Sub

    Private Sub lista_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        Select Case KeyCode
            Case 46
                If lista.ActiveCell.row > 0 Then
                    Call eliminarEspeciales(lista.ActiveCell.row)
                    leerEspeciales
                    
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

    
    Public Sub cargaLista()
        Call leerEspeciales
    End Sub












Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("cuenta", "nombre")
    mensajeAyuda = "Ayuda tipo de Cuentas Corrientes"
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

