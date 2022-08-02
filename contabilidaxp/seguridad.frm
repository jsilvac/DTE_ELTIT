VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form seguridad 
   BackColor       =   &H00D9EFFF&
   Caption         =   "MODULO DE SEGURIDAD"
   ClientHeight    =   8535
   ClientLeft      =   1260
   ClientTop       =   750
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   821
   Begin VB.Frame eliminapermiso 
      BackColor       =   &H000040C0&
      Caption         =   "Eliminando Permiso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   4455
      Left            =   4320
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton sieliminalinea 
         BackColor       =   &H00D9EFFF&
         Caption         =   "Eliminar"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton noeliminalinea 
         BackColor       =   &H00D9EFFF&
         Caption         =   "No eliminar"
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
         Left            =   2160
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3960
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid ELI 
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5530
         _Version        =   393216
         BackColor       =   16576
         ForeColor       =   8454143
         Rows            =   10
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   0
         BackColorBkg    =   16576
         GridColor       =   4210816
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame GRILLASEGURIDAD 
      BackColor       =   &H00C0E0FF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   12135
      Begin VB.TextBox USUARIO 
         BackColor       =   &H00E1FFFD&
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   8281
         _Version        =   393216
         BackColor       =   14282751
         ForeColor       =   16711680
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   12640511
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   14282751
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO"
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
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label nombreusuario 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Menu Maestro 
      Caption         =   "&INGRESOS"
      Begin VB.Menu mmaestro01 
         Caption         =   "Maestro de Cuentas del Mayor"
      End
      Begin VB.Menu mmaestro02 
         Caption         =   "Maestro de Cuentas Corrientes"
      End
      Begin VB.Menu mmaestro03 
         Caption         =   "Maestro de Centros de Costos"
      End
      Begin VB.Menu mmaestro04 
         Caption         =   "Ingreso de Comprobantes Contables"
      End
      Begin VB.Menu mmaestro05 
         Caption         =   "Ingreso de Facturas de Compra"
      End
      Begin VB.Menu mmaestro06 
         Caption         =   "Ingreso de Facturas de Ventas"
      End
      Begin VB.Menu mmaestro07 
         Caption         =   "Ingreso de Boletas de Honorarios"
      End
      Begin VB.Menu mmaestro08 
         Caption         =   "Ingreso Libro de Boletas o Zetas"
      End
      Begin VB.Menu mmaestro09 
         Caption         =   "Ingreso Facturas de Compras Propias"
      End
      Begin VB.Menu BANCO00 
         Caption         =   "SISTEMA DE BANCO"
         Begin VB.Menu BANCO01 
            Caption         =   "Cancelacion de Cheques "
         End
         Begin VB.Menu BANCO02 
            Caption         =   "Listado Cartola de Banco"
         End
         Begin VB.Menu BANCO03 
            Caption         =   "Listado de Cheques A fecha"
         End
         Begin VB.Menu BANCO04 
            Caption         =   "Listado de Cheques depositados Antes"
         End
      End
      Begin VB.Menu ACTIVO00 
         Caption         =   "SISTEMA DE ACTIVO FIJO"
      End
   End
   Begin VB.Menu proce 
      Caption         =   "&PROCESOS"
      Begin VB.Menu proce01 
         Caption         =   "Actualizacion de Movimientos"
      End
      Begin VB.Menu proce02 
         Caption         =   "Cierre Anual"
      End
   End
   Begin VB.Menu info 
      Caption         =   "&INFORMES"
      Begin VB.Menu info01 
         Caption         =   "Listado de cuentas del Mayor"
      End
      Begin VB.Menu info02 
         Caption         =   "Listado de Cuentas Corrientes"
      End
      Begin VB.Menu info03 
         Caption         =   "Listado de Centros de Costo"
      End
      Begin VB.Menu info04 
         Caption         =   "Listado Cartolas Cuentas del mayor"
      End
      Begin VB.Menu info05 
         Caption         =   "Listado Cartolas Cuentas Corrientes"
      End
      Begin VB.Menu info06 
         Caption         =   "Listado Cartolas Centros de Costo"
      End
      Begin VB.Menu infoaux 
         Caption         =   "MODULO INFORME LIBROS AUXILIARES"
         Begin VB.Menu infoau01 
            Caption         =   "Balance Tributario"
         End
         Begin VB.Menu infoau02 
            Caption         =   "Balance Analitico"
         End
         Begin VB.Menu infoau03 
            Caption         =   "Libro Mayor Analitico"
         End
         Begin VB.Menu infoau04 
            Caption         =   "Libro Diario"
         End
         Begin VB.Menu infoau05 
            Caption         =   "Libro de Compras"
         End
         Begin VB.Menu infoau06 
            Caption         =   "Libro de Ventas"
         End
         Begin VB.Menu infoau07 
            Caption         =   "Libro de Boletas"
         End
         Begin VB.Menu infoau08 
            Caption         =   "Libro de Honorarios"
         End
         Begin VB.Menu infoau09 
            Caption         =   "Libro de Compras Internas"
         End
      End
      Begin VB.Menu infosi00 
         Caption         =   "MODULO INFORMES Y CERTIFICADOS SII"
         Begin VB.Menu infosi01 
            Caption         =   "Formulario 3323"
         End
      End
      Begin VB.Menu INFOGE00 
         Caption         =   "MODULO INFORMES DE GESTION"
         Begin VB.Menu infoge01 
            Caption         =   "Lista Balance Acumulado"
         End
         Begin VB.Menu infoge02 
            Caption         =   "Lista Informe Comparativo Cuentas de Resultado"
         End
         Begin VB.Menu infoge03 
            Caption         =   "Lista Estado de Resultado"
         End
         Begin VB.Menu infoge04 
            Caption         =   "Lista Estado de Resultado x Centro de Costo"
         End
         Begin VB.Menu infoge05 
            Caption         =   "Lista Facturas x Pagar a Proveedores"
         End
         Begin VB.Menu infoge06 
            Caption         =   "Lista Cheques en Cartera"
         End
      End
   End
   Begin VB.Menu confi00 
      Caption         =   "&CONFIGURACION"
      Begin VB.Menu confi01 
         Caption         =   "Configura Sistema"
      End
      Begin VB.Menu confi02 
         Caption         =   "Maestro de Empresas"
      End
      Begin VB.Menu confi03 
         Caption         =   "Cambia Clave"
      End
      Begin VB.Menu confi04 
         Caption         =   "MODULO DE SEGURIDAD"
      End
   End
   Begin VB.Menu salir 
      Caption         =   "&SALIR"
   End
End
Attribute VB_Name = "seguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
servidor = servidor
usuariosql = USUARIO
password = password

Set SQLUTIL.conexion = db
Call Conectarconta(servidor, clientesistema + "conta", USUARIO, password)

USUARIO.text = "RODRIGO"
LEEPERMISOS


End Sub



Sub empresa()
    empre = "00"
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + empre + "'" + " ORDER BY codigoempresa"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    codigoempresa = SQLUTIL.datos(0, 3)
    nombreempresa = SQLUTIL.datos(1, 3)
    
    PRINCIPAL.Caption = "SISTEMA DE contaBILIDAD                      EMPRESA :" + codigoempresa + "  " + nombreempresa + "  " + Date$ + "  " + Time$
    status = SQLUTIL.estado

End Sub

Sub DATOSGRILLA()

grilla.Cols = 8
grilla.Rows = 2

grilla.ColWidth(0) = 120 * 10
grilla.ColWidth(1) = 120 * 30
grilla.ColWidth(2) = 120 * 9
grilla.ColWidth(3) = 120 * 9
grilla.ColWidth(4) = 120 * 9
grilla.ColWidth(5) = 120 * 9
grilla.ColWidth(6) = 120 * 9
grilla.ColWidth(7) = 120 * 9
Rem TITULOS
grilla.TextMatrix(0, 0) = "programa"
grilla.TextMatrix(0, 1) = "glosa"
grilla.TextMatrix(0, 2) = "ingresa"
grilla.TextMatrix(0, 3) = "agrega"
grilla.TextMatrix(0, 4) = "modifica"
grilla.TextMatrix(0, 5) = "elimina"
grilla.TextMatrix(0, 6) = "autoriza"
grilla.TextMatrix(0, 7) = "todas"
grilla.TextMatrix(1, 0) = ""
grilla.TextMatrix(1, 1) = ""
grilla.TextMatrix(1, 2) = ""
grilla.TextMatrix(1, 3) = ""
grilla.TextMatrix(1, 4) = ""
grilla.TextMatrix(1, 5) = ""
grilla.TextMatrix(1, 6) = ""
grilla.TextMatrix(1, 7) = ""

End Sub

Sub LEEPERMISOS()
    DATOSGRILLA
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = conta
        cSql.SQL = "SELECT programa,glosa,ingresa,agrega,modifica,elimina,autoriza,todas "
        cSql.SQL = cSql.SQL + "FROM permisos"
        cSql.SQL = cSql.SQL + " where usuario=" + "'" + USUARIO.text + "' order by programa"
      ' cSql.SQL = cSql.SQL + " where tipo=1 and numero=0000000005 order by linea"
        cSql.Execute
        linea = 0: SUMADOR = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                linea = linea + 1
                grilla.Rows = linea + 1
                For K = 0 To 7
                grilla.TextMatrix(linea, K) = resultados(K)
                Next K
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
    End With

End Sub

Sub grabapermiso(PROGRAMA, nombre)
    For K = 1 To grilla.Rows - 1
    If PROGRAMA = grilla.TextMatrix(K, 0) Then GoTo no:

    Next K
    campos(0, 0) = "usuario"
    campos(1, 0) = "programa"
    campos(2, 0) = "glosa"
    campos(3, 0) = "ingresa"
    campos(4, 0) = "agrega"
    campos(5, 0) = "modifica"
    campos(6, 0) = "elimina"
    campos(7, 0) = "autoriza"
    campos(8, 0) = "todas"
    campos(9, 0) = ""
    campos(0, 1) = USUARIO.text
    campos(1, 1) = PROGRAMA
    campos(2, 1) = nombre
    campos(3, 1) = "S"
    campos(4, 1) = "N"
    campos(5, 1) = "N"
    campos(6, 1) = "N"
    campos(7, 1) = "N"
    campos(8, 1) = "N"
    
    campos(0, 2) = "permisos"
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado <> 0 Then Stop
LEEPERMISOS
no:

End Sub

Private Sub grilla_DBLClick()
If grilla.Col = 0 Then elimina
If grilla.Row = 0 Then GoTo no:

If grilla.TextMatrix(grilla.Row, grilla.Col) <> "N" And grilla.TextMatrix(grilla.Row, grilla.Col) <> "S" Then GoTo no:
If grilla.TextMatrix(grilla.Row, grilla.Col) = "N" Then grilla.TextMatrix(grilla.Row, grilla.Col) = "S": GoTo PASO:
If grilla.TextMatrix(grilla.Row, grilla.Col) = "S" Then grilla.TextMatrix(grilla.Row, grilla.Col) = "N": GoTo PASO:
PASO:
Call MODIFICAPERMISO(grilla.TextMatrix(grilla.Row, 0), grilla.TextMatrix(0, grilla.Col), grilla.TextMatrix(grilla.Row, grilla.Col), grilla.Row)

no:

End Sub
Private Sub grilla_KeyPress(KeyAscii As Integer)

If KeyAscii = 78 And grilla.Col > 1 And grilla.Row > 0 Then grilla.TextMatrix(grilla.Row, grilla.Col) = "N": Call MODIFICAPERMISO(grilla.TextMatrix(grilla.Row, 0), grilla.TextMatrix(0, grilla.Col), grilla.TextMatrix(grilla.Row, grilla.Col), grilla.Row)
If KeyAscii = 83 And grilla.Col > 1 And grilla.Row > 0 Then grilla.TextMatrix(grilla.Row, grilla.Col) = "S": Call MODIFICAPERMISO(grilla.TextMatrix(grilla.Row, 0), grilla.TextMatrix(0, grilla.Col), grilla.TextMatrix(grilla.Row, grilla.Col), grilla.Row)

End Sub



Sub MODIFICAPERMISO(PROGRAMA, PERMISO, estado, ubicacion)
    campos(0, 0) = "usuario"
    campos(1, 0) = "programa"
    campos(2, 0) = PERMISO
    campos(3, 0) = ""
    campos(0, 1) = USUARIO.text
    campos(1, 1) = PROGRAMA
    campos(2, 1) = estado
    campos(0, 2) = "permisos"
    condicion = "usuario=" + "'" + USUARIO.text + "' and programa = '" + PROGRAMA + "'"
   
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    If SQLUTIL.estado <> 0 Then Stop
    linea = ubicacion
    grilla.Row = ubicacion
End Sub


Private Sub mmaestro01_Click()
Call grabapermiso("maestro01", mmaestro01.Caption)
End Sub
Private Sub mmaestro02_Click()
Call grabapermiso("maestro02", mmaestro02.Caption)
End Sub
Private Sub mmaestro03_Click()
Call grabapermiso("maestro03", mmaestro03.Caption)
End Sub
Private Sub mmaestro04_Click()
Call grabapermiso("maestro04", mmaestro04.Caption)
End Sub
Private Sub mmaestro05_Click()
Call grabapermiso("maestro05", mmaestro05.Caption)
End Sub
Private Sub mmaestro06_Click()
Call grabapermiso("maestro06", mmaestro06.Caption)
End Sub
Private Sub mmaestro07_Click()
Call grabapermiso("maestro07", mmaestro07.Caption)
End Sub
Private Sub mmaestro08_Click()
Call grabapermiso("maestro08", mmaestro08.Caption)
End Sub
Private Sub mmaestro09_Click()
Call grabapermiso("maestro09", mmaestro09.Caption)
End Sub

    

Private Sub noeliminalinea_Click()
GRILLASEGURIDAD.Enabled = True
grilla.SetFocus
eliminapermiso.Visible = False

End Sub

Private Sub sieliminalinea_Click()
eliminadato
eliminapermiso.Visible = False
GRILLASEGURIDAD.Enabled = True
End Sub

Private Sub USUARIO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudausuario(USUARIO)
End Sub
Sub ayudausuario(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("usuario", "nombre")
    largo = Array("18s", "40s")
    cfijo = "no"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "usuarios", USUARIO, campos, cfijo, largo, 2)
    grilla.SetFocus
    
End Sub

Private Sub USUARIO_KeyPress(KeyAscii As Integer)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 13 Then grilla.SetFocus
      End Sub

Private Sub USUARIO_LostFocus()

    campos(0, 0) = "usuario"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "usuarios"
    condicion = "usuario=" + "'" + USUARIO.text + "'"
    op = 5
    Rem If PERMISO(op) <> "S" Then NOPERMISO
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then USUARIO.SetFocus: GoTo no:
    
    nombreusuario = SQLUTIL.datos(1, 3)
    LEEPERMISOS
no:

End Sub
Sub elimina()
GRILLASEGURIDAD.Enabled = False
CARGADATOSAELIMINAR

eliminapermiso.Visible = True

End Sub
Sub CARGADATOSAELIMINAR()
ELI.Clear
ELI.Cols = 2
ELI.Rows = 10
ELI.ColWidth(0) = 120 * 12
ELI.ColWidth(1) = 120 * 20

Rem TITULOS
For VARINUM = 0 To 7
K = grilla.Row
ELI.TextMatrix(VARINUM, 0) = grilla.TextMatrix(0, VARINUM)
ELI.TextMatrix(VARINUM, 1) = grilla.TextMatrix(K, VARINUM)
Next VARINUM

End Sub
Sub eliminadato()
    campos(0, 0) = "usuario"
    campos(1, 0) = "programa"
    campos(2, 0) = ""
    campos(0, 1) = USUARIO.text
    campos(1, 1) = grilla.TextMatrix(grilla.Row, 0)
    campos(0, 2) = "permisos"
    condicion = "usuario=" + "'" + USUARIO.text + "' and programa = '" + campos(1, 1) + "'"
   
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado <> 0 Then Stop
LEEPERMISOS
End Sub
