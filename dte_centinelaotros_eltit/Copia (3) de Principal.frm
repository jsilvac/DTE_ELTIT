VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FF9A514-A943-11D2-8D43-F90F0D71B6F6}#1.0#0"; "changeres.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Ventas"
   ClientHeight    =   8085
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   11400
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Principal.frx":0ECA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ChangeResProject.ChangeRes ChangeRes1 
      Left            =   60
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1085
   End
   Begin MSComctlLib.StatusBar barraEstado 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7710
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   14446
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MMaestros 
      Caption         =   "&MAESTROS"
      Index           =   99
      Begin VB.Menu MMClientes 
         Caption         =   "Maestro de &Clientes"
      End
      Begin VB.Menu MMVendedores 
         Caption         =   "Maestro de &Vendedores"
      End
      Begin VB.Menu MMCajas 
         Caption         =   "Maestro de C&ajas"
      End
      Begin VB.Menu MMCajeros 
         Caption         =   "Maestro de Cajeros"
      End
      Begin VB.Menu Mtipovivienda 
         Caption         =   "Maestro Tipo Viviendas"
      End
   End
   Begin VB.Menu Mventas 
      Caption         =   "GESTION DE &VENTAS"
      Index           =   99
      Begin VB.Menu MVVentas 
         Caption         =   "Pantalla de &Ventas"
         Shortcut        =   ^V
      End
      Begin VB.Menu MVNotas 
         Caption         =   "Pantalla de &Notas de Crédito"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu MVPagos 
         Caption         =   "Pantalla de &Pagos de Clientes"
         Enabled         =   0   'False
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu MVEgresos 
         Caption         =   "Pantalla de &Egresos de Caja"
         Enabled         =   0   'False
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu MVAuditoria 
         Caption         =   "Pantalla de &Auditoria de Ventas"
         Shortcut        =   ^A
      End
      Begin VB.Menu MVCoptizaciones 
         Caption         =   "Pantalla de &Cotizaciones y Notas de Pedido"
         Enabled         =   0   'False
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu MVDepositos 
         Caption         =   "Pantalla de &Depósitos"
         Enabled         =   0   'False
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu CDImpresora 
         Caption         =   "Pantalla de &Configuracion de Impresora fiscal"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu MVCuponera 
         Caption         =   "&Imprimir Cuponera Crédito"
         Visible         =   0   'False
      End
      Begin VB.Menu MVPagoCuponera 
         Caption         =   "Pago con C&uponera"
         Visible         =   0   'False
      End
      Begin VB.Menu MVSeparador2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MVCambio 
         Caption         =   "Consulta Precio de P&roductos"
         Visible         =   0   'False
      End
      Begin VB.Menu Cventas 
         Caption         =   "Pantalla Comparación Ventas"
      End
   End
   Begin VB.Menu ds 
      Caption         =   "DESPACHOS Y GARANTIAS"
      Index           =   99
      Begin VB.Menu Mserviciotecnico 
         Caption         =   "Maestro Servicio Tecnico"
      End
      Begin VB.Menu MGarantias 
         Caption         =   "Pantalla de  Servicio Tecnico Garantias"
      End
      Begin VB.Menu STPendientes 
         Caption         =   "Listado Servicios Tecnicos Pendientes"
      End
      Begin VB.Menu MVSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu MFletes 
         Caption         =   "Generar Orden de Flete"
      End
      Begin VB.Menu Guiafletes 
         Caption         =   "Pantalla Guia Despacho Mercaderias"
      End
      Begin VB.Menu Guiadespachofletes 
         Caption         =   "Guia Despacho de Mercaderias"
      End
      Begin VB.Menu MPdespachar 
         Caption         =   "Listado de Control de Despachos"
      End
      Begin VB.Menu Dremoto 
         Caption         =   "Listado Despacho Remoto"
      End
   End
   Begin VB.Menu coti 
      Caption         =   "COTIZACIONES"
      Index           =   99
      Begin VB.Menu cotiza 
         Caption         =   "Pantalla Cotizaciones"
         Index           =   1
      End
      Begin VB.Menu cotiza 
         Caption         =   "Listado Cotizaciones Pendientes"
         Index           =   3
      End
   End
   Begin VB.Menu CMR2 
      Caption         =   "TARJETA DE CREDITO"
      Index           =   99
      Begin VB.Menu CMR 
         Caption         =   "PANTALLA GENERA CUOTAS"
         Index           =   1
      End
      Begin VB.Menu CMR 
         Caption         =   "PANTALLA PAGO CUOTAS"
         Index           =   2
      End
      Begin VB.Menu CMR 
         Caption         =   "LISTADO DE CREDITOS OTORGADOS"
         Index           =   3
      End
      Begin VB.Menu CMR 
         Caption         =   "LISTADO DE PAGOS RECIBIDOS"
         Index           =   4
      End
      Begin VB.Menu CMR 
         Caption         =   "LISTADO DE CLIENTES CREDITO"
         Index           =   5
      End
      Begin VB.Menu CMR 
         Caption         =   "LISTADO DE DEUDAS POR CUOTAS"
         Index           =   6
      End
      Begin VB.Menu CMR 
         Caption         =   "EMISION ESTADO DE CUENTA"
         Index           =   7
      End
      Begin VB.Menu CMR 
         Caption         =   "LISTA COMPORTAMIENTO DE PAGO"
         Index           =   9
      End
      Begin VB.Menu CMR 
         Caption         =   "LIBRO IMPUESTO LETRAS"
         Index           =   10
      End
      Begin VB.Menu CMR 
         Caption         =   "CONTRATOS"
         Index           =   13
      End
      Begin VB.Menu CMR 
         Caption         =   "LISTADO DE REPACTACIONES"
         Index           =   14
      End
      Begin VB.Menu CMR 
         Caption         =   "INFORME SEGURO DEGRAVAMEM"
         Index           =   15
      End
      Begin VB.Menu CMR 
         Caption         =   "LISTADO COMPRA CREDITOS CLIENTES"
         Index           =   16
      End
      Begin VB.Menu COTMP 
         Caption         =   "MODULO COBRANZA TMP"
         Index           =   99
         Begin VB.Menu COBRA 
            Caption         =   "PANTALLA AUDITORIA CREDITOS"
            Index           =   1
         End
         Begin VB.Menu COBRA 
            Caption         =   "PANTALLA GESTION COBRANZA CREDITOS"
            Index           =   2
         End
         Begin VB.Menu COBRA 
            Caption         =   "PANTALLA GENERA CARTAS COBRANZAS"
            Index           =   3
         End
         Begin VB.Menu COBRA 
            Caption         =   "PANTALLA GENERA ENVIOS A DICOM"
            Index           =   4
         End
      End
      Begin VB.Menu PRE 
         Caption         =   "MODULO PRESTAMO "
         Index           =   99
         Begin VB.Menu PREUF 
            Caption         =   "PANTALLA DE PRESTAMOS EN U.F"
            Index           =   1
         End
         Begin VB.Menu PREUF 
            Caption         =   "LISTADO PRESTAMOS A CARGAR"
            Index           =   2
         End
         Begin VB.Menu PREUF 
            Caption         =   "LISTADO PRESTAMOS VIGENTES"
            Index           =   3
         End
      End
   End
   Begin VB.Menu MCreditos 
      Caption         =   "GESTION DE &CREDITOS"
      Enabled         =   0   'False
      Index           =   99
      Visible         =   0   'False
      Begin VB.Menu MCCreditos 
         Caption         =   "Cartola de &Clientes"
      End
      Begin VB.Menu MCCheques 
         Caption         =   "Cartola de Clientes C&heques"
      End
      Begin VB.Menu MCSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu MCCartera 
         Caption         =   "&Listado de Cheques en Cartera"
      End
      Begin VB.Menu MLCobranza 
         Caption         =   "Li&stado de Cobranza"
      End
      Begin VB.Menu MCCupos 
         Caption         =   "Listado de C&upos, Saldos, Autorización"
      End
      Begin VB.Menu MCSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu MCDocumentos 
         Caption         =   "&Documentos Manuales a Cobranza"
      End
      Begin VB.Menu MCSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu MCProrroga 
         Caption         =   "Comprobantes de &Prorroga"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MCProtestados 
         Caption         =   "&Ingreso de Cheques Protestados"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MInformes 
      Caption         =   "&INFORMES"
      Index           =   99
      Begin VB.Menu MIClientes 
         Caption         =   "Listado de &Clientes"
      End
      Begin VB.Menu MIVendedores 
         Caption         =   "Listado de &Vendedores"
      End
      Begin VB.Menu MItiposdepago 
         Caption         =   "Listado Ventas x Tipos De Pago"
      End
      Begin VB.Menu mmcajeras 
         Caption         =   "Resumen de Ventas x Cajeras"
      End
      Begin VB.Menu MIRango 
         Caption         =   "Resumen de Ventas x Productos"
      End
      Begin VB.Menu MIUnidad 
         Caption         =   "Resumen de Ventas por Vendedores"
      End
      Begin VB.Menu MIDinero 
         Caption         =   "Resumen de Ventas Por Clientes"
      End
      Begin VB.Menu MIdocumentos 
         Caption         =   "Listado de Documentos Emitidos"
      End
      Begin VB.Menu MIseparador2 
         Caption         =   "-"
      End
      Begin VB.Menu MIventas 
         Caption         =   "Libro de V&entas"
      End
      Begin VB.Menu Lagotados 
         Caption         =   "Listado Faltantes"
      End
      Begin VB.Menu promecli 
         Caption         =   "Listado de Promedio Ventas x Cliente"
      End
   End
   Begin VB.Menu MConfiguracion 
      Caption         =   "&CONFIGURACION"
      Index           =   99
      Begin VB.Menu MCLocal 
         Caption         =   "Cambiar &Local Activo"
         Shortcut        =   ^L
      End
      Begin VB.Menu MCFecha 
         Caption         =   "Cambiar &Fecha Sistema"
         Shortcut        =   ^F
      End
      Begin VB.Menu MCCSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu MCPuesto 
         Caption         =   "Configurar Puesto de &Trabajo"
         Shortcut        =   ^T
      End
      Begin VB.Menu MCCSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu MCActualizar 
         Caption         =   "&Actualizar"
      End
      Begin VB.Menu MCCSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu MCPermisos 
         Caption         =   "&Permisos de Usuario"
      End
      Begin VB.Menu Mauditoria 
         Caption         =   "Modulo de Auditoria de Usuarios"
      End
      Begin VB.Menu MCClave 
         Caption         =   "&Cambio Clave"
      End
   End
   Begin VB.Menu motor0 
      Caption         =   "CONSULTA CHEQUES"
      Index           =   99
      Begin VB.Menu motor 
         Caption         =   "Motor de Consultas"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu motor 
         Caption         =   "Consultar Cheque"
         Index           =   2
      End
   End
   Begin VB.Menu harinas 
      Caption         =   "DESPACHO &HARINAS"
      Index           =   99
      Begin VB.Menu DespachoHarinas 
         Caption         =   "Despachar Harinas"
      End
   End
   Begin VB.Menu MSalir 
      Caption         =   "&SALIR"
      Index           =   99
      Begin VB.Menu MSSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu MSSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu MSAcerca 
         Caption         =   "&Acerca de..."
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Public PASO As Boolean

Private Sub CUOTAS_Click()
creditoTMP.Show
Call grabaprincipal(creditoTMP.Caption)
End Sub

Private Sub CMR_Click(Index As Integer)
If Index = 1 Then
AUTORIZASISTEMA = False

LOGINCAJERA.Show vbModal
If AUTORIZASISTEMA = True Then
creditoTMP.Caption = CMR(Index).Caption

creditoTMP.Show
End If

End If

If Index = 2 Then
AUTORIZASISTEMA = False

LOGINCAJERA.Show vbModal
If AUTORIZASISTEMA = True Then

creditoPAGOSTMP.Show
End If

End If

If Index = 3 Then tmplistado3.Show
If Index = 4 Then tmplistado4.Show
If Index = 5 Then tmplistado5.Show
If Index = 6 Then tmplistado6.Show
If Index = 7 Then tmplistado7.Show
If Index = 8 Then tmplistado8.Show
If Index = 9 Then tmplistado10.Show
If Index = 10 Then tmplistado9.Show
If Index = 13 Then contratos.Show
If Index = 14 Then tmplistado13.Show
If Index = 15 Then tmplistado14.Show
If Index = 16 Then ListadoComprasPorClientes.Show
Call grabaprincipal(CMR(Index).Caption)
End Sub



Private Sub COBRA_Click(Index As Integer)
If Index = 1 Then Modificarcuotas.Show: Modificarcuotas.Caption = COBRA(1).Caption
If Index = 2 Then tmpcobranza.Show: tmpcobranza.Caption = COBRA(2).Caption
If Index = 3 Then tmplistado11.Show: tmplistado11.Caption = COBRA(3).Caption
If Index = 4 Then tmplistado12.Show: tmplistado12.Caption = COBRA(4).Caption

End Sub

Private Sub cotiza_Click(Index As Integer)
If Index = 1 Then cotiza01.Show: cotiza01.Caption = cotiza(1).Caption
Call grabaprincipal(CMR(Index).Caption)

End Sub

Private Sub Cventas_Click()
        titCaption = Replace(Cventas.Caption, "&", "")
        Load LibroVentascomparacion
        LibroVentascomparacion.Caption = titCaption
        LibroVentascomparacion.Show
        Call grabaprincipal(LibroVentascomparacion.Caption)
End Sub

Private Sub DespachoHarinas_Click(Index As Integer)
DespachoHarina.Show
End Sub

Private Sub Dremoto_Click()
        titCaption = Replace(Dremoto.Caption, "&", "")
        Load Listadodespachootrolocal
        Listadodespachootrolocal.Caption = titCaption
        Listadodespachootrolocal.Show
        Call grabaprincipal(Listadodespachootrolocal.Caption)
End Sub

Private Sub Guiadespachofletes_Click()
        titCaption = Replace(Guiadespachofletes.Caption, "&", "")
        Load Pdespachoflete
        Pdespachoflete.Caption = titCaption
        Pdespachoflete.Show
        Call grabaprincipal(Pdespachoflete.Caption)
End Sub

Private Sub Guiafletes_Click()
        titCaption = Replace(Guiafletes.Caption, "&", "")
        Load Pguiadespacho
        Pguiadespacho.Caption = titCaption
        Pguiadespacho.Show
        Call grabaprincipal(Pguiadespacho.Caption)
End Sub

Private Sub Lagotados_Click()
        titCaption = Replace(Lagotados.Caption, "&", "")
        Load Agotados
        Agotados.Caption = titCaption
        Agotados.Show
        Call grabaprincipal(Agotados.Caption)
End Sub

Private Sub Mauditoria_Click()
        titCaption = Replace(Mauditoria.Caption, "&", "")
        Load moduloseguridad2
        moduloseguridad2.Caption = titCaption
        moduloseguridad2.Show
        Call grabaprincipal(moduloseguridad2.Caption)
End Sub

Private Sub MCClave_Click()
        titCaption = Replace(MFletes.Caption, "&", "")
        Load maestro15
        maestro15.Caption = titCaption
        maestro15.Show
        Call grabaprincipal(maestro15.Caption)
End Sub

Private Sub MFletes_Click()
        titCaption = Replace(MFletes.Caption, "&", "")
        Load Maestrofletes
        Maestrofletes.Caption = titCaption
        Maestrofletes.Show
        Call grabaprincipal(Maestrofletes.Caption)
End Sub

Private Sub MGarantias_Click()
        titCaption = Replace(MMGarantias.Caption, "&", "")
        Load MMGarantias
        MMGarantias.Caption = titCaption
        MMGarantias.Show
        Call grabaprincipal(MMGarantias.Caption)
End Sub

Private Sub MIdocumentos_Click()
        titCaption = Replace(MIdocumentos.Caption, "&", "")
        Load DocumentosEmitidos
        DocumentosEmitidos.Caption = titCaption
        DocumentosEmitidos.Show
        Call grabaprincipal(DocumentosEmitidos.Caption)
End Sub

Private Sub MItiposdepago_Click()
        titCaption = Replace(MItiposdepago.Caption, "&", "")
        Load listadoformapago
        listadoformapago.Caption = titCaption
        listadoformapago.Show
        Call grabaprincipal(listadoformapago.Caption)
End Sub

Private Sub mmcajeras_Click()
LibroVentascajera.Show
Call grabaprincipal(LibroVentascajera.Caption)
End Sub

Private Sub MMCajeros_Click()
        titCaption = Replace(MCajeras.Caption, "&", "")
        Load MCajeras
        MCajeras.Caption = titCaption
        MCajeras.Show
        Call grabaprincipal(MCajeras.Caption)
End Sub

'==================================================================================
'MAESTROS
'==================================================================================
    Private Sub MMClientes_Click()
        titCaption = Replace(MMClientes.Caption, "&", "")
        Load MClientes
        MClientes.Caption = titCaption
        MClientes.Show
        Call grabaprincipal(MClientes.Caption)
    End Sub
'
'Private Sub MMUsarios_Click()
'    MUsuarios.Show
'    MUsuarios.Caption = Replace(MUsuarios.Caption, "&", "")
'End Sub

    Private Sub MMVendedores_Click()
        titCaption = Replace(MMVendedores.Caption, "&", "")
        Load Mvendedores
        Mvendedores.Caption = titCaption
        Mvendedores.Show
        Call grabaprincipal(Mvendedores.Caption)
    End Sub
    
   
    
    Private Sub MMCajas_Click()
        titCaption = Replace(MMCajas.Caption, "&", "")
        Load MCajas
        MCajas.Caption = titCaption
        MCajas.Show
        Call grabaprincipal(MCajas.Caption)
        
    End Sub

Private Sub motor_Click(Index As Integer)

If Index = 2 Then
AUTORIZASISTEMA = False

LOGINCAJERA.Show vbModal
If AUTORIZASISTEMA = True Then
consultacheques.Caption = CMR(Index).Caption

consultacheques.Show
End If

End If

'If Index = 1 Then motor01.Show
Call grabaprincipal(motor01.Caption)
End Sub

Private Sub MPdespachar_Click()
        titCaption = Replace(MPdespachar.Caption, "&", "")
        Load MercaderiaPorDespachar
        MercaderiaPorDespachar.Caption = titCaption
        MercaderiaPorDespachar.Show
        Call grabaprincipal(MPdespachar.Caption)
End Sub

Private Sub Mserviciotecnico_Click()
        titCaption = Replace(MServicio.Caption, "&", "")
        Load MServicio
        MServicio.Caption = titCaption
        MServicio.Show
        Call grabaprincipal(MServicio.Caption)
End Sub

Private Sub MSSalir_Click()
End

End Sub

'==================================================================================
'MAESTROS
'==================================================================================
Private Sub Mtipovivienda_Click()
 titCaption = Replace(MVivienda.Caption, "&", "")
        Load MVivienda
        MVivienda.Caption = titCaption
        MVivienda.Show
        Call grabaprincipal(MVivienda.Caption)
End Sub

'==================================================================================
'GESTION DE VENTAS
'==================================================================================
    Private Sub MVVentas_Click()
        titCaption = Replace(MVVentas.Caption, "&", "")
        Load PVentas
        PVentas.Caption = titCaption
        PVentas.Show
        Call grabaprincipal(PVentas.Caption)
    End Sub
    
    Private Sub MVNotas_Click()
        titCaption = Replace(MVNotas.Caption, "&", "")
        Load PNotasCredito
        PNotasCredito.Caption = titCaption
        PNotasCredito.Show
        Call grabaprincipal(PNotasCredito.Caption)
    End Sub
    
    Private Sub MVPagos_Click()
        titCaption = Replace(MVPagos.Caption, "&", "")
        Load PClientes
        PClientes.Caption = titCaption
        PClientes.Show
         Call grabaprincipal(PClientes.Caption)
    End Sub
    
    Private Sub MVEgresos_Click()
        titCaption = Replace(MVEgresos.Caption, "&", "")
        Load ECaja
        ECaja.Caption = titCaption
        ECaja.Show
        Call grabaprincipal(ECaja.Caption)
    End Sub
    
    Private Sub MVAuditoria_Click()
        titCaption = Replace(MVAuditoria.Caption, "&", "")
        Load PAuditoriaVentas
        PAuditoriaVentas.Caption = titCaption
        PAuditoriaVentas.Show
        Call grabaprincipal(PAuditoriaVentas.Caption)
    End Sub
    
    Private Sub MVDepositos_Click()
        titCaption = Replace(MVDepositos.Caption, "&", "")
        Load PDepositos
        PDepositos.Caption = titCaption
        PDepositos.Show
        Call grabaprincipal(PDepositos.Caption)
    End Sub
    Private Sub MVCambio_Click()
        titCaption = Replace(MVCambio.Caption, "&", "")
        Load CambioPrecio
        CambioPrecio.Caption = titCaption
        CambioPrecio.Show
        Call grabaprincipal(CambioPrecio.Caption)
    End Sub
    Private Sub CDImpresora_Click()
        titCaption = Replace(CDImpresora.Caption, "&", "")
        Load ventaEspecial
        ventaEspecial.Caption = titCaption
        ventaEspecial.Show
        Call grabaprincipal(ventaEspecial.Caption)
End Sub
'==================================================================================
'GESTION DE VENTAS
'==================================================================================

'==================================================================================
'GESTION DE CREDITOS
'==================================================================================
    Private Sub MCCreditos_Click()
        titCaption = Replace(MCCreditos.Caption, "&", "")
        Load LCCliente
        LCCliente.Caption = titCaption
        LCCliente.Show
         Call grabaprincipal(LCCliente.Caption)
    End Sub
    
    Private Sub MCCheques_Click()
        titCaption = Replace(LChCliente.Caption, "&", "")
        Load LChCliente
        LChCliente.Caption = titCaption
        LChCliente.Show
        Call grabaprincipal(LChCliente.Caption)
    End Sub
    
    Private Sub MCCartera_Click()
        titCaption = Replace(MCCartera.Caption, "&", "")
        Load LChCartera
        LChCartera.Caption = titCaption
        LChCartera.Show
        Call grabaprincipal(LChCartera.Caption)
    End Sub
    
    Private Sub MLCobranza_Click()
        titCaption = Replace(MLCobranza.Caption, "&", "")
        Load LCobranza
        LCobranza.Caption = titCaption
        LCobranza.Show
        Call grabaprincipal(LCobranza.Caption)
    End Sub
    
    Private Sub MCCupos_Click()
        titCaption = Replace(MCCupos.Caption, "&", "")
        Load LCupSalAut
        LCupSalAut.Caption = titCaption
        LCupSalAut.Show
        Call grabaprincipal(LCupSalAut.Caption)
    End Sub
    
    Private Sub MCDocumentos_Click()
        titCaption = Replace(MCDocumentos.Caption, "&", "")
        Load docManuales
        docManuales.Caption = titCaption
        docManuales.Show
        Call grabaprincipal(docManuales.Caption)
    End Sub
    
    Private Sub MCProrroga_Click()
        titCaption = Replace(MCProrroga.Caption, "&", "")
        Load CProrroga
        CProrroga.Caption = titCaption
        CProrroga.Show
        Call grabaprincipal(CProrroga.Caption)
    End Sub
    
    Private Sub MCProtestados_Click()
        titCaption = Replace(MCProtestados.Caption, "&", "")
        Load CProtestados
        CProtestados.Caption = titCaption
        CProtestados.Show
        Call grabaprincipal(CProtestados.Caption)
    End Sub
'==================================================================================
'GESTION DE CREDITOS
'==================================================================================

'==================================================================================
'GESTION DE PRODUCCION
'==================================================================================
      
'GESTION DE PRODUCCION
'==================================================================================

'==================================================================================
'INFORMES
'==================================================================================
Private Sub MIClientes_Click()
    titCaption = Replace(MIClientes.Caption, "&", "")
    Load LClientes
    LClientes.Caption = titCaption
    LClientes.Show
    Call grabaprincipal(LClientes.Caption)
End Sub

Private Sub MIVendedores_Click()
    titCaption = Replace(MIVendedores.Caption, "&", "")
    Load LVendedores
    LVendedores.Caption = titCaption
    LVendedores.Show
    Call grabaprincipal(LVendedores.Caption)
End Sub

Private Sub MIRango_Click()
    titCaption = Replace(MIRango.Caption, "&", "")
    Load VentasMontosLocal
    VentasMontosLocal.Caption = titCaption
    VentasMontosLocal.Show
    Call grabaprincipal(VentasMontosLocal.Caption)
End Sub

Private Sub MIUnidad_Click()
    titCaption = Replace(LibroVentasvendedores.Caption, "&", "")
    Load LibroVentasvendedores
    LibroVentasvendedores.Caption = titCaption
    LibroVentasvendedores.Show
    Call grabaprincipal(LibroVentasvendedores.Caption)
End Sub

Private Sub MIDinero_Click()
    LibroVentasclientes.Show
    LibroVentasclientes.Caption = Replace(MIDinero.Caption, "&", "")
    Call grabaprincipal(LibroVentasclientes.Caption)
End Sub

Private Sub MIComparativas_Click()
    ListadoVentasComparativas.Show
    Call grabaprincipal(ListadoVentasComparativas.Caption)
    
End Sub

Private Sub MIMolienda_Click()
    LibroMolienda.Show
    Call grabaprincipal(LibroMolienda.Caption)
End Sub

Private Sub MIPromedios_Click()
    PreciosPromedio.Show
    Call grabaprincipal(PreciosPromedio.Caption)
End Sub


Private Sub MISII_Click()
    ListadoRetencion.Show
    Call grabaprincipal(ListadoRetencion.Caption)
End Sub



Private Sub MIVentas_Click()
    LibroVentas.Show
    LibroVentas.Caption = Replace(MIventas.Caption, "&", "")
    Call grabaprincipal(LibroVentas.Caption)
End Sub

Private Sub MIVentasCliente_Click()
    VentasCliente.Show
    Call grabaprincipal(VentasCliente.Caption)
    
End Sub

Private Sub MIVentasKilos_Click()
    VentasKilosDia.Show
    Call grabaprincipal(VentasKilosDia.Caption)
    
End Sub

Private Sub MIVentasVendedores_Click()
    ventasVendedor.Show
    Call grabaprincipal(ventasVendedor.Caption)
End Sub
'==================================================================================
'INFORMES
'==================================================================================

Private Sub MCActualizar_Click()
    Call escribeArchivoRuta("SISTEMA", App.Path, "C:\UPDATE.TXT")
    Call escribeArchivoRuta("UPDATE", rutaUpdate & "\" & App.EXEName & ".exe", "C:\UPDATE.TXT")
    
    If ExisteArchivo(App.Path & "\Update.exe") = True Then
        Call Shell(App.Path & "\Update.exe", vbNormalFocus)
    'Else
        'Call MsgBox("No se ha encontrado el archivo de actualizacion" & vbCrLf & "Contactese con el proveedor de su sistema", vbOKOnly, "Error")
    End If
    
    If comparaArchivos(App.Path & "\Update.exe", rutaUpdate & "\Update.exe") = False Then
        Call VisualFileCopy(rutaUpdate & "\Update.exe", App.Path & "\Update.exe")
    End If
    
    Call Shell(App.Path & "\Update.exe", vbNormalFocus)
    
End Sub

Private Sub MCFecha_Click()
    cambioFecha.Show vbModal
    cambioFecha.Caption = Replace(MCFecha.Caption, "&", "")
    Call grabaprincipal(cambioFecha.Caption)
End Sub

Private Sub MCLocal_Click()
        cambioLocal.Show vbModal
        cambioLocal.Caption = Replace(MCLocal.Caption, "&", "")
        Call grabaprincipal(cambioLocal.Caption)
End Sub

Private Sub MCPermisos_Click()
'    permisosUsuario.Show
'    permisosUsuario.Caption = Replace(MCPermisos.Caption, "&", "")
seguridad2.Show
Call grabaprincipal(seguridad2.Caption)

End Sub






Private Sub MDIForm_Load()
'    Dim saveTitle$
'    If App.PrevInstance Then
'        saveTitle$ = App.Title
'        App.Title = "... duplicate instance."
'        Me.Caption = "... duplicate instance. "
'        AppActivate saveTitle$
'        SendKeys "% R", True
'        End
'    End If
  Me.barraEstado.Panels(1).text = UCase(Me.Caption)
    If PASO = False Then
        barraEstado.Panels(3).text = ""
'        If segu = True Then
'        Seguridad.Show vbModal, Me

             Call revisarmenus(Principal)
            'Call cargaMenuPermisos
'        End If
'        segurity = False
    End If
    sqlventas.audit = False

    cajera = "0000"
'    ChangeRes1.GetMonitorInfo = True
'    resX = ChangeRes1.Xpixels
'    resY = ChangeRes1.Ypixels
'    ChangeRes1.Xpixels = 1024
'    ChangeRes1.Ypixels = 768
'    ChangeRes1.ChangeResolution = True
'    If ChangeRes1.Error = True Then
'        MsgBox "La resolucion de su monitor no puede cambiarse a " & ChangeRes1.Xpixels & " X " & ChangeRes1.Ypixels
'    End If
    Rem Me.Picture = LoadPicture(App.Path & "\trigo.jpg")
    'Call TranslucentForm(Me, 200)
   
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If PVentas.Visible = True Then
        If PVentas.imprimio = False Then
            Unload PVentas
        End If
    End If
'    ChangeRes1.Xpixels = resX
'    ChangeRes1.Ypixels = resY
'    ChangeRes1.ChangeResolution = True
    End
End Sub
Private Sub MSCerrar_Click()
    PASO = False
    usuarioSistema = ""
    passwordSistema = ""
    Call UnloadHijos(Me)
     
End Sub

Private Sub verificarUpdate()
    If rutaUpdate <> "" Then
        If comparaArchivos(App.Path & "\" & App.EXEName & ".exe", rutaUpdate & "\" & App.EXEName & ".exe") = False Then
                Call MCActualizar_Click
            
        End If
    End If
End Sub

Private Sub pagos_Click()
creditoPAGOSTMP.Show

Call grabaprincipal(creditoPAGOSTMP.Caption)
End Sub

Private Sub PREUF_Click(Index As Integer)
If Index = 1 Then prestamouf.Show: prestamouf.Caption = PREUF(Index).Caption

If Index = 2 Then prestamosacargar.Show: prestamosacargar.Caption = PREUF(Index).Caption
If Index = 3 Then prestamosvigentes.Show: prestamosvigentes.Caption = PREUF(Index).Caption
Call grabaprincipal(PREUF(Index).Caption)
End Sub

Private Sub promecli_Click()
        titCaption = Replace(promecli.Caption, "&", "")
        Load ListadoVentasclientes
        ListadoVentasclientes.Caption = titCaption
        ListadoVentasclientes.Show
        Call grabaprincipal(ListadoVentasclientes.Caption)

End Sub

Private Sub STPendientes_Click()
        titCaption = Replace(STPendientes.Caption, "&", "")
        Load ListadoServiciosPendientes
        ListadoServiciosPendientes.Caption = titCaption
        ListadoServiciosPendientes.Show
        Call grabaprincipal(ListadoServiciosPendientes.Caption)
End Sub
