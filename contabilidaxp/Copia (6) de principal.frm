VERSION 5.00
Begin VB.MDIForm PRINCIPAL 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   15195
   Icon            =   "principal.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "principal.frx":5A4A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu ingreso 
      Caption         =   "INGRESOS"
      Index           =   99
      WindowList      =   -1  'True
      Begin VB.Menu ingresos 
         Caption         =   "Maestro de Cuentas del Mayor"
         Index           =   1
      End
      Begin VB.Menu ingresos 
         Caption         =   "Maestro de Cuentas Corrientes"
         Index           =   2
      End
      Begin VB.Menu ingresos 
         Caption         =   "Maestro de Centros de Costo"
         Index           =   3
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso de Comprobantes Contables"
         Index           =   4
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso de Factura de Compra"
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
         Caption         =   "Ingreso de Boletas o Zetas"
         Index           =   8
      End
      Begin VB.Menu ingresos 
         Caption         =   "Ingreso U.F"
         Index           =   9
      End
      Begin VB.Menu bancos 
         Caption         =   "SISTEMA CONTROL BANCARIO"
         Index           =   99
         Begin VB.Menu banco 
            Caption         =   "Rebaja Cartola de Banco"
            Index           =   1
         End
         Begin VB.Menu banco 
            Caption         =   "Lista Cartola de banco"
            Index           =   2
         End
         Begin VB.Menu banco 
            Caption         =   "Lista cheques emitidos Sin Cobrar"
            Index           =   3
         End
         Begin VB.Menu banco 
            Caption         =   "Lista Cheques Cobrados Antes del Vencimiento"
            Index           =   4
         End
         Begin VB.Menu banco 
            Caption         =   "Distribucion de Cheques"
            Index           =   5
         End
         Begin VB.Menu banco 
            Caption         =   "SISTEMA AUTOMATICO BANCO SANTANDER"
            Index           =   6
         End
         Begin VB.Menu banco 
            Caption         =   "Informes Para Distribucion de Cheques"
            Index           =   7
         End
         Begin VB.Menu banco 
            Caption         =   "Maestro Codigos Bancarios"
            Index           =   8
         End
      End
      Begin VB.Menu proved 
         Caption         =   "SISTEMA PAGO PROVEEDORES"
         Index           =   99
         Begin VB.Menu prove 
            Caption         =   "Traspasa Facturas de Compra Recibida x Recepcion"
            Index           =   1
         End
         Begin VB.Menu prove 
            Caption         =   "Traspasa Facturas de Compra Recibida Empr.Relacionada"
            Index           =   2
         End
         Begin VB.Menu prove 
            Caption         =   "Pantalla Pago Proveedores x  Ordenes de Compra"
            Index           =   3
         End
         Begin VB.Menu prove 
            Caption         =   "Pantalla de Nominas Bancarias de Pago Proveedores"
            Index           =   4
         End
         Begin VB.Menu prove 
            Caption         =   "Listado de Guias de Devolucion Pendiente de Rebajar"
            Index           =   5
         End
         Begin VB.Menu prove 
            Caption         =   "Pantalla de Guias  de Devolucion"
            Index           =   6
         End
         Begin VB.Menu prove 
            Caption         =   "Pantalla Pago Mercaderia Entre Locales"
            Index           =   7
         End
         Begin VB.Menu prove 
            Caption         =   "Pantalla Ingresa Presupuesto de Pago"
            Index           =   8
         End
      End
      Begin VB.Menu publi 
         Caption         =   "SISTEMA DE PUBLICIDAD"
         Index           =   99
         Begin VB.Menu publicidad 
            Caption         =   "Contratos de Publicidad"
            Index           =   1
         End
         Begin VB.Menu publicidad 
            Caption         =   "Genera Facturas de Publicidad"
            Index           =   2
         End
         Begin VB.Menu publicidad 
            Caption         =   "Lista Publicidades Pendientes"
            Index           =   3
         End
         Begin VB.Menu publicidad 
            Caption         =   "Informa Resumen de Compras x Proveedor"
            Index           =   4
         End
         Begin VB.Menu publicidad 
            Caption         =   "Lista de Contratos Por Proveedor"
            Index           =   5
         End
      End
      Begin VB.Menu acti1 
         Caption         =   "SISTEMA DE ACTIVO FIJO"
         Index           =   99
         Begin VB.Menu activo 
            Caption         =   "Maestro Tabla de IPC"
            Index           =   1
         End
         Begin VB.Menu activo 
            Caption         =   "Maestro de Activos Fijo "
            Index           =   2
         End
         Begin VB.Menu activo 
            Caption         =   "Listado de Activo Fijo"
            Index           =   3
         End
         Begin VB.Menu activo 
            Caption         =   "Cierre Anual"
            Index           =   4
         End
      End
      Begin VB.Menu SCAR 
         Caption         =   "SISTEMA CONTROL DE ARRIENDOS"
         Index           =   99
         Begin VB.Menu arriendo 
            Caption         =   "Maestro de Propiedades"
            Index           =   1
         End
         Begin VB.Menu arriendo 
            Caption         =   "Maestro de Arrendadores"
            Index           =   2
         End
         Begin VB.Menu arriendo 
            Caption         =   "Maestro de Arrendatarios"
            Index           =   3
         End
         Begin VB.Menu arriendo 
            Caption         =   "Maestro de Monedas"
            Index           =   4
         End
         Begin VB.Menu arriendo 
            Caption         =   "Pantalla Control de Arriendos"
            Index           =   5
         End
         Begin VB.Menu arriendo 
            Caption         =   "Comprobante de Pago Arriendo"
            Index           =   6
         End
         Begin VB.Menu arriendo 
            Caption         =   "Listado de Propiedades en Arriendo y su Estado"
            Index           =   7
         End
      End
      Begin VB.Menu SCIN 
         Caption         =   "SISTEMA CONTROL DE INVERSIONES"
         Index           =   99
         Begin VB.Menu inver 
            Caption         =   "Maestro de Bancos"
            Index           =   1
         End
         Begin VB.Menu inver 
            Caption         =   "Maestro de Documentos de Inversion"
            Index           =   2
         End
         Begin VB.Menu inver 
            Caption         =   "Pantalla de Inversiones Depositos a Plazo"
            Index           =   3
         End
         Begin VB.Menu inver 
            Caption         =   "Pantalla de Inversiones Fondos Mutuos"
            Index           =   4
         End
         Begin VB.Menu inver 
            Caption         =   "Listado de Inversiones"
            Index           =   5
         End
      End
      Begin VB.Menu SCCB 
         Caption         =   "SISTEMA CONTROL DE CONSUMOS BASICOS"
         Index           =   99
         Begin VB.Menu consumo 
            Caption         =   "Maestro de Tipos de Consumo Basico"
            Index           =   1
         End
         Begin VB.Menu consumo 
            Caption         =   "Maestro de Proveedores de Consumos"
            Index           =   2
         End
         Begin VB.Menu consumo 
            Caption         =   "Maestro de Unidades de Consumo Basico"
            Index           =   3
         End
         Begin VB.Menu consumo 
            Caption         =   "Ingresa Boleta o Factura de Gastos"
            Index           =   4
         End
         Begin VB.Menu consumo 
            Caption         =   "Informe estado de Consumos Basicos"
            Index           =   5
         End
         Begin VB.Menu consumo 
            Caption         =   "Estadistica de Consumos Basicos"
            Index           =   6
         End
      End
      Begin VB.Menu SCBASI 
         Caption         =   "SISTEMA CONTROL DE COMPROMISOS FINANCIEROS"
         Index           =   99
         Begin VB.Menu prestamo 
            Caption         =   "Maestro Tipos de Compromiso Bancario"
            Index           =   1
         End
         Begin VB.Menu prestamo 
            Caption         =   "Maestro de Compromisos Bancarios"
            Index           =   2
         End
         Begin VB.Menu prestamo 
            Caption         =   "Informe de Compromisos Bancarios"
            Index           =   3
         End
         Begin VB.Menu prestamo 
            Caption         =   "Informe Resumen Compromisos Bancarios"
            Index           =   4
         End
      End
      Begin VB.Menu flu 
         Caption         =   "SISTEMA  FLUJO DE CAJA"
         Index           =   99
         Begin VB.Menu flujo 
            Caption         =   "Pantalla Flujo de Caja"
            Index           =   1
         End
      End
      Begin VB.Menu cp 
         Caption         =   "SISTEMA DE GASTOS V/S PRESUPUESTO"
         Index           =   99
         Begin VB.Menu costo 
            Caption         =   "Maestro Centros de Gastos y Produccion"
            Index           =   1
         End
         Begin VB.Menu costo 
            Caption         =   "Maestro detalle Cuentas Resultado"
            Index           =   2
         End
         Begin VB.Menu costo 
            Caption         =   "Ingreso de Presupuestos"
            Index           =   3
         End
         Begin VB.Menu costo 
            Caption         =   "Informe Comparacion Presupuesto"
            Index           =   4
         End
         Begin VB.Menu costo 
            Caption         =   "Informe Comparacion de Gastos"
            Enabled         =   0   'False
            Index           =   5
            Visible         =   0   'False
         End
      End
      Begin VB.Menu internos 
         Caption         =   "SISTEMA CONTROL CONSUMOS INTERNOS"
         Index           =   99
         Begin VB.Menu interno 
            Caption         =   "Genera Vales de Credito"
            Index           =   1
         End
         Begin VB.Menu interno 
            Caption         =   "Listado de Vales de Credito"
            Index           =   2
         End
         Begin VB.Menu interno 
            Caption         =   "Cartola de Vales de Credito"
            Enabled         =   0   'False
            Index           =   3
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "PROCESOS "
      Index           =   99
      Begin VB.Menu proceso 
         Caption         =   "Cierre Anual"
         Index           =   1
      End
      Begin VB.Menu proceso 
         Caption         =   "Actualizacion Mensual de Movimientos"
         Index           =   2
      End
      Begin VB.Menu proceso 
         Caption         =   "Traspasa Facturas de Venta"
         Index           =   3
      End
      Begin VB.Menu proceso 
         Caption         =   "Traspasa Boletas de Venta"
         Index           =   4
      End
      Begin VB.Menu proceso 
         Caption         =   "Centralizacion de  Ventas"
         Index           =   5
      End
      Begin VB.Menu proceso 
         Caption         =   "Revisa Contabilizacion Libro Compra"
         Index           =   6
      End
      Begin VB.Menu proceso 
         Caption         =   "Revisa Contabilizacion Libro de Ventas"
         Index           =   7
      End
      Begin VB.Menu proceso 
         Caption         =   "Revisa Contabilizacion Libro de Honorarios"
         Index           =   8
      End
      Begin VB.Menu proceso 
         Caption         =   "Contabilizacion Tesoreria"
         Index           =   9
      End
      Begin VB.Menu proceso 
         Caption         =   "Contabilizacion Promotora palguin"
         Index           =   10
      End
      Begin VB.Menu proceso 
         Caption         =   "Contabilizacion Movimientos Inventario"
         Index           =   11
      End
      Begin VB.Menu proceso 
         Caption         =   "Contabilizacion Remuneraciones"
         Index           =   12
      End
   End
   Begin VB.Menu info 
      Caption         =   "INFORMES"
      Index           =   99
      Begin VB.Menu informe 
         Caption         =   "Lista Archivos Maestros"
         Index           =   1
      End
      Begin VB.Menu informe 
         Caption         =   "Lista Cartolas Contables"
         Index           =   2
      End
      Begin VB.Menu informesauxiliares 
         Caption         =   "MODULO DE LIBROS AUXILIARES"
         Index           =   99
         Begin VB.Menu infoaux 
            Caption         =   "Lista Balance Tributario"
            Index           =   1
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista balance Analitico"
            Index           =   2
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Libro Mayor Analitico"
            Index           =   3
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Libro Diario"
            Index           =   4
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Libro de Ventas"
            Index           =   5
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Libro de Compras"
            Index           =   6
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Libro de Honorarios"
            Index           =   7
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Libro de Boletas"
            Index           =   8
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Determinacion Capital Propio"
            Index           =   9
         End
         Begin VB.Menu infoaux 
            Caption         =   "Lista Balance Clasificado"
            Index           =   10
         End
      End
      Begin VB.Menu informessii 
         Caption         =   "MODULO INFORMES Y CERTIFICADOS SII"
         Index           =   99
         Begin VB.Menu infosii 
            Caption         =   "formulario Resumen Iva F3323"
            Index           =   1
         End
         Begin VB.Menu infosii 
            Caption         =   "certificado honorarios F1879 "
            Index           =   2
         End
         Begin VB.Menu infosii 
            Caption         =   "Informe Retencion Mensual Harina"
            Index           =   3
         End
         Begin VB.Menu infosii 
            Caption         =   "Informe Retencion Mensual Carne"
            Index           =   4
         End
         Begin VB.Menu infosii 
            Caption         =   "Resumen Ventas Con I.L.A"
            Index           =   5
         End
         Begin VB.Menu infosii 
            Caption         =   "Archivos Planos I.V.A"
            Index           =   6
         End
      End
      Begin VB.Menu informesgestion 
         Caption         =   "MODULO INFORMES DE GESTION "
         Index           =   99
         Begin VB.Menu infoge 
            Caption         =   "Lista Estado de Resultados"
            Index           =   1
         End
         Begin VB.Menu infoge 
            Caption         =   "Lista Facturas Por Pagar"
            Index           =   2
         End
         Begin VB.Menu infoge 
            Caption         =   "Lista Honorarios por Pagar"
            Index           =   3
         End
         Begin VB.Menu infoge 
            Caption         =   "Lista Facturas por Cobrar"
            Index           =   4
         End
      End
      Begin VB.Menu infocontrol 
         Caption         =   "MODULO INFORMES DE CONTROL"
         Index           =   99
         Begin VB.Menu infoco 
            Caption         =   "Lista Comprobantes Descuadrados"
            Index           =   1
         End
         Begin VB.Menu infoco 
            Caption         =   "Lista Hojas Para Timbrar"
            Index           =   2
         End
         Begin VB.Menu infoco 
            Caption         =   "Buscador de Montos"
            Index           =   3
         End
         Begin VB.Menu infoco 
            Caption         =   "Busca Codigos Eliminados"
            Index           =   4
         End
      End
      Begin VB.Menu bala 
         Caption         =   "MODULO LIBRO INVENTARIO BALANCE"
         Index           =   99
         Begin VB.Menu balance 
            Caption         =   "Informe Inventario Valorizado"
            Index           =   1
         End
         Begin VB.Menu balance 
            Caption         =   "Informe Inventario Valorizado"
            Enabled         =   0   'False
            Index           =   2
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu configuracion 
      Caption         =   "CONFIGURACION "
      Index           =   99
      Begin VB.Menu confi 
         Caption         =   "Cambia Fecha Sistema"
         Index           =   1
      End
      Begin VB.Menu confi 
         Caption         =   "Configura empresa a utilizar"
         Index           =   2
      End
      Begin VB.Menu confi 
         Caption         =   "Maestro de Empresas"
         Index           =   3
      End
      Begin VB.Menu confi 
         Caption         =   "Cambiar Fecha Cierre"
         Index           =   4
      End
      Begin VB.Menu confi 
         Caption         =   "Cambia Clave"
         Index           =   5
      End
      Begin VB.Menu confi 
         Caption         =   "Traspasa Datos Version D.O.S"
         Index           =   6
      End
      Begin VB.Menu confi 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu confi 
         Caption         =   "&Actualizar"
         Index           =   8
      End
      Begin VB.Menu confi 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu seguridades 
         Caption         =   "MODULO DE SEGURIDAD Y PERMISOS"
         Index           =   99
         Begin VB.Menu seguri 
            Caption         =   "Mantencion de Usuarios"
            Index           =   1
         End
         Begin VB.Menu seguri 
            Caption         =   "Modulo de Auditoria de Usuarios"
            Index           =   2
         End
         Begin VB.Menu seguri 
            Caption         =   "Analisis de Usuarios"
            Index           =   3
         End
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "PRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub activo_Click(Index As Integer)
If Index = 1 Then maestro04.Show: maestro04.Caption = activo(Index).Caption
End Sub

Private Sub arriendo_Click(Index As Integer)
If Index = 1 Then arriendo01.Show: arriendo01.Caption = arriendo(Index).Caption
If Index = 2 Then arriendo02.Show: arriendo02.Caption = arriendo(Index).Caption
If Index = 3 Then arriendo03.Show: arriendo03.Caption = arriendo(Index).Caption
If Index = 4 Then arriendo04.Show: arriendo04.Caption = arriendo(Index).Caption
If Index = 5 Then arriendo05.Show: arriendo05.Caption = arriendo(Index).Caption
If Index = 6 Then arriendo06.Show: arriendo06.Caption = arriendo(Index).Caption
If Index = 7 Then arriendo07.Show: arriendo07.Caption = arriendo(Index).Caption

End Sub

Private Sub balance_Click(Index As Integer)
If Index = 1 Then balance01.Show: balance01.Caption = balance(Index).Caption
If Index = 2 Then balance02.Show: balance02.Caption = balance(Index).Caption
End Sub

Private Sub banco_Click(Index As Integer)
If Index = 1 Then banco01.Show: banco01.Caption = banco(Index).Caption
If Index = 2 Then banco02.Show: banco02.Caption = banco(Index).Caption
If Index = 3 Then banco03.Show: banco03.Caption = banco(Index).Caption
If Index = 4 Then banco04.Show: banco04.Caption = banco(Index).Caption
If Index = 5 Then banco05.Show: banco05.Caption = banco(Index).Caption
If Index = 6 Then banco06.Show: banco06.Caption = banco(Index).Caption
If Index = 7 Then banco07.Show: banco07.Caption = banco(Index).Caption
If Index = 8 Then Maestrocodigosbancarios.Show: Maestrocodigosbancarios.Caption = banco(Index).Caption

Call grabaprincipal(banco(Index).Caption)
End Sub

Private Sub confi_Click(Index As Integer)
Dim programas As Double
If Index = 1 Then confi00.Show vbModal
If Index = 2 Then
'  programas = Forms.Count
'  If programas = 1 Then
    confi01.Show
'  Else
'    MsgBox "DEBE CERRAR TODOS LOS PROGRAMAS ANTES DE CAMBIAR LOCAL", vbCritical, "ATENCION"
'  End If
End If


 
If Index = 3 Then confi02.Show
If Index = 4 Then confi08.Show
If Index = 6 Then TRASPASA.Show
If Index = 5 Then maestro15.Show vbModal
If Index = 8 Then Call actualizar
Call grabaprincipal(confi(Index).Caption)


End Sub


Private Sub consumo_Click(Index As Integer)
If Index = 1 Then consumo01.Show: consumo01.Caption = consumo(Index).Caption
If Index = 2 Then consumo02.Show: consumo02.Caption = consumo(Index).Caption
If Index = 3 Then consumo03.Show: consumo03.Caption = consumo(Index).Caption
If Index = 4 Then consumo04.Show: consumo04.Caption = consumo(Index).Caption
If Index = 5 Then consumo05.Show: consumo05.Caption = consumo(Index).Caption
If Index = 6 Then consumo06.Show: consumo06.Caption = consumo(Index).Caption

Call grabaprincipal(consumo(Index).Caption)

End Sub

Private Sub costo_Click(Index As Integer)
If Index = 1 Then presu00.Show: presu00.Caption = costo(Index).Caption
If Index = 2 Then presu01.Show: presu01.Caption = costo(Index).Caption
If Index = 3 Then presu02.Show: presu02.Caption = costo(Index).Caption
If Index = 4 Then presu03.Show: presu03.Caption = costo(Index).Caption
If Index = 5 Then presu04.Show: presu04.Caption = costo(Index).Caption

End Sub

Private Sub flujo_Click(Index As Integer)
If Index = 1 Then flujocajamaster.Show: flujocajamaster.Caption = flujo(Index).Caption

End Sub

Private Sub infoaux_Click(Index As Integer)
If Index = 1 Then auxiliar01.Show
If Index = 2 Then auxiliar02.Show
If Index = 3 Then auxiliar03.Show
If Index = 4 Then auxiliar04.Show
If Index = 5 Then auxiliar44.Show
If Index = 6 Then auxiliar05.Show
If Index = 7 Then auxiliar06.Show
If Index = 8 Then auxiliar07.Show
If Index = 9 Then MAESTRO20.Show
If Index = 10 Then maestro21.Show

Call grabaprincipal(infoaux(Index).Caption)


End Sub


Private Sub infoco_Click(Index As Integer)
If Index = 1 Then control01.Show
If Index = 2 Then control02.Show
If Index = 3 Then control03.Show
If Index = 4 Then control04.Show
Call grabaprincipal(infoco(Index).Caption)


End Sub

Private Sub infoge_Click(Index As Integer)
If Index = 1 Then infoge01.Show
If Index = 2 Then infoge02.Show
If Index = 3 Then infoge03.Show
If Index = 4 Then infoge04.Show
If Index = 5 Then infoge05.Show
If Index = 6 Then infoge06.Show
If Index = 7 Then infoge07.Show
If Index = 8 Then infoge08.Show
Call grabaprincipal(infoge(Index).Caption)

End Sub

Private Sub informe_Click(Index As Integer)
If Index = 1 Then informa01.Show
If Index = 2 Then informa04.Show
Call grabaprincipal(Informe(Index).Caption)

End Sub

Private Sub infosii_Click(Index As Integer)
If Index = 1 Then form3323.Show
If Index = 2 Then form1879.Show
If Index = 3 Then infoharina.Show
If Index = 4 Then infocarne.Show
If Index = 5 Then infoilas.Show
If Index = 6 Then planosiva.Show
Call grabaprincipal(infosii(Index).Caption)

End Sub

Private Sub interno_Click(Index As Integer)
If Index = 1 Then interno01.Show: interno01.Caption = inver(Index).Caption
If Index = 2 Then interno02.Show: interno02.Caption = inver(Index).Caption
If Index = 3 Then interno03.Show: interno03.Caption = inver(Index).Caption

End Sub

Private Sub inver_Click(Index As Integer)
If Index = 1 Then inver01.Show: inver01.Caption = inver(Index).Caption
If Index = 2 Then inver02.Show: inver02.Caption = inver(Index).Caption
If Index = 3 Then inver03.Show: inver03.Caption = inver(Index).Caption
If Index = 4 Then inver04.Show: inver04.Caption = inver(Index).Caption
If Index = 5 Then inver05.Show: inver05.Caption = inver(Index).Caption

End Sub

Private Sub MDIForm_Activate()
PRINCIPAL.Caption = "SISTEMA DE CONTABILIDAD             Usuario:" + USUARIOSISTEMA + "     Empresa: " + nombreempresa + "                 Fecha :" + Str(fechasistema)
sincronizarFechaHora

End Sub
Public Sub sincronizarFechaHora()
        
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim FECHA As String
        Dim HORA As String
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = db
        csql.sql = "SELECT DATE_FORMAT(CURRENT_TIMESTAMP(),'%d-%m-%Y') AS fecha, TIME_FORMAT(CURRENT_TIMESTAMP(),'%T') AS hora "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
               FECHA = resultado("fecha")
            HORA = resultado("hora")
         
            Date = DateValue(FECHA)
            'Para la hora:
            Time = TimeValue(HORA)
        End If
        csql.Close
        Set csql = Nothing
   
   
   
    Rem fechasistema = "2009-04-06"
    End Sub

Private Sub MDIForm_Load()
 Call Conectar_BD
 Call revisarmenus(PRINCIPAL)
End Sub
Public Sub revisarmenus(ByRef frm As Form)
    Dim ctlControl As Object
    Dim cad As String
    Dim cadindex As String
    Dim tipovariable As String
    
    On Error Resume Next
    For Each ctlControl In frm.Controls
           cad = ctlControl.Name
           cadindex = ctlControl.Index
           tipovariable = TypeName(ctlControl)
           
           
            If tipovariable = "Menu" And cadindex <> "99" Then
                 If existepermiso(USUARIOSISTEMA, ctlControl.Caption) = True Then
                    ctlControl.Enabled = True
                    Else
                    ctlControl.Enabled = False
                    End If
            End If
       cadindex = "0"
       ' DoEvents
    Next ctlControl
End Sub

Private Function existepermiso(usuario, programa) As Boolean

    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
        existepermiso = False
      
        Set cSql2.ActiveConnection = conta
        cSql2.sql = "SELECT todas,ingresa "
        cSql2.sql = cSql2.sql + "FROM " + clientesistema + "conta.segu_permisos "
        cSql2.sql = cSql2.sql + "where usuario='" + usuario + "' and programa='" + programa + "'"
        cSql2.Execute
        
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        If resultados2(0) = 1 Or resultados2(1) = 1 Then
            existepermiso = True
        End If
End If
End Function




Private Sub ingresos_Click(Index As Integer)
If Index = 1 Then maestro01.Show: maestro01.Caption = ingresos(Index).Caption
If Index = 2 Then maestro02.Show: maestro02.Caption = ingresos(Index).Caption
If Index = 3 Then maestro03.Show: maestro03.Caption = ingresos(Index).Caption
If Index = 4 Then ingreso01.Show: ingreso01.Caption = ingresos(Index).Caption
If Index = 5 Then ingreso02.Show: ingreso02.Caption = ingresos(Index).Caption
If Index = 6 Then ingreso03.Show: ingreso03.Caption = ingresos(Index).Caption
If Index = 7 Then ingreso04.Show: ingreso04.Caption = ingresos(Index).Caption
If Index = 8 Then ingreso05.Show: ingreso05.Caption = ingresos(Index).Caption
If Index = 9 Then maestro05.Show: maestro05.Caption = ingresos(Index).Caption
Call grabaprincipal(ingresos(Index).Caption)
End Sub

Private Sub segu01_Click()
seguridad.Show

End Sub
Sub desactivamenu()

End Sub

Sub desabilitamenus()
For k = 1 To 1
ingresos(k).Visible = False
Next k

End Sub
 
Private Sub prestamo_Click(Index As Integer)
If Index = 1 Then prestamo01.Show: prestamo01.Caption = prestamo(Index).Caption
If Index = 2 Then prestamo02.Show: prestamo02.Caption = prestamo(Index).Caption
If Index = 3 Then prestamo03.Show: prestamo03.Caption = prestamo(Index).Caption
If Index = 4 Then prestamo04.Show: prestamo04.Caption = prestamo(Index).Caption
If Index = 5 Then prestamo05.Show: prestamo05.Caption = prestamo(Index).Caption

Call grabaprincipal(prestamo(Index).Caption)
End Sub

Private Sub proceso_Click(Index As Integer)
If Index = 1 Then proceso01.Show
If Index = 2 Then proceso02.Show
If Index = 3 Then proceso03.Show
If Index = 4 Then proceso04.Show
If Index = 5 Then proceso05.Show
If Index = 6 Then revisacompras.Show
If Index = 7 Then revisaventas.Show
If Index = 8 Then revisahonorarios.Show
If Index = 9 Then contabilizatesoreria.Show
If Index = 10 Then
        If empresaactiva = "28" Then
            contabilizapromotora.Show
            Else
            MsgBox ("DEBE INGRESAR CON LA EMPRESA 28")
        End If
End If
If Index = 11 Then contabilizainventario.Show
If Index = 12 Then proceso06.Show
Call grabaprincipal(PROCESO(Index).Caption)

End Sub

Private Sub prove_Click(Index As Integer)
If Index = 1 Then prove0001.Show
If Index = 2 Then prove0003.Show
If Index = 3 Then prove0002.Show
If Index = 4 Then prove0004.Show
If Index = 5 Then prove0005.Show
If Index = 6 Then prove0006.Show
If Index = 7 Then prove0007.Show
If Index = 8 Then prove0008.Show: prove0008.Caption = prove(Index).Caption


Call grabaprincipal(prove(Index).Caption)


End Sub

Private Sub publicidad_Click(Index As Integer)
If Index = 1 Then publi0001.Show: publi0001.Caption = publicidad(Index).Caption
If Index = 2 Then publi0002.Show: publi0002.Caption = publicidad(Index).Caption
If Index = 3 Then publi0003.Show: publi0003.Caption = publicidad(Index).Caption
If Index = 4 Then publi0004.Show
If Index = 5 Then publi0005.Show: publi0005.Caption = publicidad(Index).Caption
Call grabaprincipal(publicidad(Index).Caption)

End Sub

Private Sub salir_Click()
Unload Me

End Sub

Private Sub seguri_Click(Index As Integer)
If Index = 1 Then moduloseguridad.Show: moduloseguridad.Caption = seguri(Index).Caption
If Index = 2 Then moduloseguridad2.Show: moduloseguridad2.Caption = seguri(Index).Caption
If Index = 3 Then moduloseguridad3.Show: moduloseguridad3.Caption = seguri(Index).Caption

Call grabaprincipal(seguri(Index).Caption)



End Sub

 
Private Sub seguri2_Click()

End Sub
