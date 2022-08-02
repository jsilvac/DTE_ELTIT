VERSION 5.00
Begin VB.Form Principal2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEMA DE CONTABILIDAD"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "principal2.frx":0000
   ScaleHeight     =   705
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   Begin VB.Frame CONFIGU 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   9720
      Width           =   8655
      Begin VB.Label HORAFORM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7200
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label FECHAFORM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label NOMBREEMPRESAFORM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label USUARIOFORM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9EFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USUARIO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9EFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EMPRESA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9EFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9EFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HORA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7200
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Menu ingreso 
      Caption         =   "&INGRESOS"
      Begin VB.Menu ingresos 
         Caption         =   "Maestro de Cuentas del Mayor"
         Index           =   1
      End
      Begin VB.Menu ingresos 
         Caption         =   "Maestro de Cuentas Corrientes"
         Index           =   2
      End
      Begin VB.Menu ingresos 
         Caption         =   "Maestro de Centros de Costos"
         Index           =   3
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso de Comprobantes Contables"
         Index           =   4
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso de Facturas de Compra"
         Index           =   5
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso de Facturas de Ventas"
         Index           =   6
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso de Boletas de Honorarios"
         Index           =   7
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso Libro de Boletas o Zetas"
         Index           =   8
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso Facturas de Compras Propias"
         Index           =   9
      End
      Begin VB.Menu banco 
         Caption         =   "SISTEMA DE BANCO"
         Index           =   0
         Begin VB.Menu bancos 
            Caption         =   "Cancelacion de Cheques "
            Index           =   1
         End
         Begin VB.Menu bancos 
            Caption         =   "Listado Cartola de Banco"
            Index           =   2
         End
         Begin VB.Menu bancos 
            Caption         =   "Listado de Cheques A fecha"
            Index           =   3
         End
         Begin VB.Menu bancos 
            Caption         =   "Listado de Cheques depositados Antes"
            Index           =   4
         End
      End
      Begin VB.Menu activo 
         Caption         =   "SISTEMA DE ACTIVO FIJO"
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "&PROCESOS"
      Begin VB.Menu proceso 
         Caption         =   "Actualizacion de Movimientos"
         Index           =   1
      End
      Begin VB.Menu proceso 
         Caption         =   "Cierre Anual"
         Index           =   2
      End
   End
   Begin VB.Menu informes 
      Caption         =   "&INFORMES"
      Begin VB.Menu informe 
         Caption         =   "Listado de cuentas del Mayor"
         Index           =   1
      End
      Begin VB.Menu informe 
         Caption         =   "Listado de Cuentas Corrientes"
         Index           =   2
      End
      Begin VB.Menu informe 
         Caption         =   "Listado de Centros de Costo"
         Index           =   3
      End
      Begin VB.Menu informe 
         Caption         =   "Listado Cartolas Cuentas del mayor"
         Index           =   4
      End
      Begin VB.Menu informe 
         Caption         =   "Listado Cartolas Cuentas Corrientes"
         Index           =   5
      End
      Begin VB.Menu informe 
         Caption         =   "Listado Cartolas Centros de Costo"
         Index           =   6
      End
      Begin VB.Menu auxiliar 
         Caption         =   "MODULO INFORME LIBROS AUXILIARES"
         Begin VB.Menu auxiliares 
            Caption         =   "Balance Tributario"
            Index           =   1
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Balance Analitico"
            Index           =   2
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Libro Mayor Analitico"
            Index           =   3
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Libro Diario"
            Index           =   4
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Libro de Compras"
            Index           =   5
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Libro de Ventas"
            Index           =   6
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Libro de Boletas"
            Index           =   7
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Libro de Honorarios"
            Index           =   8
         End
         Begin VB.Menu auxiliares 
            Caption         =   "Libro de Compras Internas"
            Index           =   9
         End
      End
      Begin VB.Menu certificados 
         Caption         =   "MODULO INFORMES Y CERTIFICADOS SII"
         Begin VB.Menu sii 
            Caption         =   "Formulario 3323"
            Index           =   1
         End
      End
      Begin VB.Menu GESTIONES 
         Caption         =   "MODULO INFORMES DE GESTION"
         Begin VB.Menu gestion 
            Caption         =   "Lista Balance Acumulado"
            Index           =   1
         End
         Begin VB.Menu gestion 
            Caption         =   "Lista Informe Comparativo Cuentas de Resultado"
            Index           =   2
         End
         Begin VB.Menu gestion 
            Caption         =   "Lista Estado de Resultado"
            Index           =   3
         End
         Begin VB.Menu gestion 
            Caption         =   "Lista Estado de Resultado x Centro de Costo"
            Index           =   4
         End
         Begin VB.Menu gestion 
            Caption         =   "Lista Facturas x Pagar a Proveedores"
            Index           =   5
         End
         Begin VB.Menu gestion 
            Caption         =   "Lista Cheques en Cartera"
            Index           =   6
         End
      End
   End
   Begin VB.Menu configuraciones 
      Caption         =   "&CONFIGURACION"
      Begin VB.Menu configuracion 
         Caption         =   "Configura Sistema"
         Index           =   1
      End
      Begin VB.Menu configuracion 
         Caption         =   "Maestro de Empresas"
         Index           =   2
      End
      Begin VB.Menu configuracion 
         Caption         =   "Cambia Clave"
         Index           =   3
      End
      Begin VB.Menu seguridades 
         Caption         =   "MODULO DE SEGURIDAD"
         Begin VB.Menu seguridad 
            Caption         =   "Mantencion de usuarios"
            Index           =   1
         End
         Begin VB.Menu seguridad 
            Caption         =   "Historico de Eventos"
            Index           =   2
         End
      End
   End
   Begin VB.Menu salir 
      Caption         =   "&SALIR"
   End
End
Attribute VB_Name = "Principal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub confi04_Click()
seguridad.Show

End Sub




Private Sub Form_Load()

leeractivo
empresa
desabilitamenus
LEEPERMISOS

End Sub



Sub empresa()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = "direccion"
    campos(3, 0) = "comuna"
    campos(4, 0) = "ciudad"
    campos(5, 0) = "rut"
    campos(6, 0) = "cuentaproveedor"
    campos(7, 0) = "cuentahonorarios"
    campos(8, 0) = "cuentaclientes"
    campos(9, 0) = "ivacredito"
    campos(10, 0) = "ivadebito"
    campos(11, 0) = "retencionhonorarios"
    campos(12, 0) = "cuentaperdida"
    campos(13, 0) = "cuentaganancia"
    campos(14, 0) = "auditoria"
    campos(15, 0) = "actividadeconomica"
    campos(16, 0) = "representantelegal"
    campos(17, 0) = "rutrepresentante"
    campos(18, 0) = "emailcontable"
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + empresaactiva + "'" + " ORDER BY codigoempresa"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    codigoempresa = SQLUTIL.datos(0, 3)
    nombreempresa = SQLUTIL.datos(1, 3)
    NOMBREEMPRESAFORM = nombreempresa
    USUARIOFORM.Caption = USUARIOSISTEMA
    FECHAFORM.Caption = Mid(Date$, 4, 2) + "-" + Mid(Date$, 1, 2) + "-" + Mid(Date$, 7, 4)
    HORAFORM = Time$
    status = SQLUTIL.ESTADO
    For K = 0 To 18
    DATOSEMPRESA(K) = SQLUTIL.datos(K, 3)
    
    Next K
End Sub




Private Sub ingresos_Click(Index As Integer)
If Index = 1 Then maestro01.Show
If Index = 2 Then maestro02.Show
If Index = 3 Then maestro03.Show
If Index = 4 Then maestro04.Show
If Index = 5 Then maestro05.Show
If Index = 6 Then maestro06.Show
If Index = 7 Then maestro07.Show
If Index = 8 Then maestro08.Show
If Index = 9 Then maestro09.Show
End Sub

Private Sub segu01_Click()
seguridad.Show

End Sub
Sub LEEPERMISOS()

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = conta
        cSql.SQL = "SELECT programa,glosa,ingresa,agrega,modifica,elimina,autoriza,todas "
        cSql.SQL = cSql.SQL + "FROM permisos"
        cSql.SQL = cSql.SQL + " where usuario=" + "'" + USUARIOSISTEMA + "' order by programa"
      ' cSql.SQL = cSql.SQL + " where tipo=1 and numero=0000000005 order by linea"
        cSql.Execute
        linea = 0: SUMADOR = 0
        
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
            linea = linea + 1
           
            Call ACTIVAMENU(resultados(0))
            
            PERMISOS(linea, 1) = resultados(0)
            For K = 2 To 7
            PERMISOS(linea, K) = resultados(K)
            Next K
            If resultados(7) = "S" Then For K = 2 To 7: PERMISOS(linea, K) = "S": Next K
            resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
    End With

End Sub
Sub desactivamenu()

End Sub

Sub ACTIVAMENU(opcion)
If opcion = "maestro01" Then ingresos(1).Visible = True: ejecuta(1) = opcion
If opcion = "maestro02" Then ingresos(2).Visible = True: ejecuta(2) = opcion
If opcion = "maestro03" Then ingresos(3).Visible = True: ejecuta(3) = opcion
If opcion = "maestro04" Then ingresos(4).Visible = True: ejecuta(4) = opcion
If opcion = "maestro05" Then ingresos(5).Visible = True: ejecuta(5) = opcion
If opcion = "maestro06" Then ingresos(6).Visible = True: ejecuta(6) = opcion

End Sub
Sub desabilitamenus()
For K = 1 To 9
ingresos(K).Visible = False
Next K

End Sub
Sub grabaseguridad()


End Sub
Sub leeractivo()
    campos(0, 0) = "usuario"
    campos(1, 0) = "nombreprogramaactivo"
    campos(2, 0) = "empresaactiva"
    campos(3, 0) = ""
    campos(0, 1) = USUARIOSISTEMA
    campos(1, 1) = "MENU PRINCIPAL "
    campos(2, 1) = "00"
    campos(0, 2) = "usuariosactivos"
    condicion = "usuario=" + "'" + USUARIOSISTEMA + "' and nombreprogramaactivo='" + "MENU PRINCIPAL" + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
empresaactiva = SQLUTIL.datos(2, 3)

End Sub

Private Sub salir_Click()
Unload Me

End Sub
