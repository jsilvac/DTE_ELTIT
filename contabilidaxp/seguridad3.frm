VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form moduloseguridadconta 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Mantencion de Usuarios"
   ClientHeight    =   10065
   ClientLeft      =   1260
   ClientTop       =   750
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   671
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1010
   WindowState     =   2  'Maximized
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   1905
      Left            =   2160
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   3360
      BackColor       =   8454016
      Caption         =   "MODULO DE COPIA"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   15
         Tag             =   "proveedor"
         Top             =   810
         Width           =   3045
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   14
         Tag             =   "proveedor"
         Top             =   450
         Width           =   3045
      End
      Begin VB.CommandButton CANCELAR 
         BackColor       =   &H0000FF00&
         Caption         =   "CANCELAR"
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
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1350
         Width           =   1635
      End
      Begin VB.CommandButton COPIAR 
         BackColor       =   &H0000FF00&
         Caption         =   "ACEPTAR"
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
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1350
         Width           =   1680
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Usuario Destino"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   17
         Top             =   810
         Width           =   1395
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Usuario Origen"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   450
         Width           =   1425
      End
   End
   Begin XPFrame.FrameXp MENU 
      Height          =   5325
      Left            =   45
      TabIndex        =   7
      Top             =   45
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   9393
      BackColor       =   49344
      Caption         =   "MENU"
      CaptionEstilo3D =   1
      BackColor       =   49344
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
      Begin VB.CommandButton MENU1 
         Appearance      =   0  'Flat
         Caption         =   "INGRESOS"
         Height          =   240
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin FlexCell.Grid Grid1 
         Height          =   4875
         Left            =   0
         TabIndex        =   9
         Top             =   225
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8599
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FRMUSUARIO 
      Height          =   4380
      Left            =   135
      TabIndex        =   0
      Top             =   5625
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   7726
      BackColor       =   8454016
      Caption         =   "USUARIOS"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "COPIAR PERMISOS"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3645
         Width           =   2355
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "AGREGAR USUARIO"
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
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3645
         Width           =   2355
      End
      Begin FlexCell.Grid Grid2 
         Height          =   3255
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supr - Eliminar Usuario"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   495
         TabIndex        =   18
         Top             =   4050
         Width           =   1740
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4335
      Left            =   7020
      TabIndex        =   2
      Top             =   5625
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7646
      BackColor       =   8454016
      Caption         =   "DATOS USUARIOS"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "MODIFICAR"
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3780
         Width           =   2355
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "GRABAR"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3780
         Width           =   2355
      End
      Begin FlexCell.Grid Grid3 
         Height          =   3165
         Left            =   45
         TabIndex        =   3
         Top             =   315
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   5583
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin FlexCell.Grid Grid4 
         Height          =   3165
         Left            =   4050
         TabIndex        =   5
         Top             =   315
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   5583
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.Menu maestro0 
      Caption         =   "&MAESTROS"
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Productos"
         Index           =   1
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Secciones"
         Index           =   2
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Departamentos"
         Index           =   3
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Lineas"
         Index           =   4
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Marcas de Productos"
         Index           =   5
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Temporadas"
         Index           =   6
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Impuestos"
         Index           =   7
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Locales"
         Index           =   8
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Bodegas"
         Index           =   9
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Proveedores"
         Index           =   10
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro Tipos de Precios"
         Index           =   11
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro de Rubros"
         Index           =   12
      End
      Begin VB.Menu maestro 
         Caption         =   "Maestro Alias"
         Index           =   15
      End
   End
   Begin VB.Menu VENTAS 
      Caption         =   "&VENTAS    "
      Begin VB.Menu sistemaventas 
         Caption         =   "Sistema de Ventas"
         Index           =   1
      End
   End
   Begin VB.Menu comprasp 
      Caption         =   "&COMPRAS"
      Begin VB.Menu compra1 
         Caption         =   "Ordenes de Compra"
         Index           =   1
      End
      Begin VB.Menu compra1 
         Caption         =   "Recepcion de Ordenes de Compra"
         Index           =   2
      End
      Begin VB.Menu compra1 
         Caption         =   "Busca Ordenes por Proveedor"
         Index           =   3
      End
      Begin VB.Menu compra1 
         Caption         =   "Listado de Ordenes Recepcionadas"
         Index           =   4
      End
      Begin VB.Menu compra1 
         Caption         =   "Listado de Mercaderias No Llegadas"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu inve 
      Caption         =   "&MOVIMIENTOS"
      Begin VB.Menu guia 
         Caption         =   "Traspaso Entre Bodegas"
         Index           =   1
      End
      Begin VB.Menu guia 
         Caption         =   "Módulo Control Mercaderias Entre Locales"
         HelpContextID   =   1
         Index           =   2
         Begin VB.Menu inve_10 
            Caption         =   "Guía Envio de Mercaderias entre Locales"
            Index           =   1
         End
         Begin VB.Menu inve_10 
            Caption         =   "Guía Recepcion de Mercaderias entre Locales"
            Enabled         =   0   'False
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu inve_10 
            Caption         =   "Informe Mercaderías entre Locales"
            Index           =   3
         End
      End
      Begin VB.Menu guia 
         Caption         =   "Módulo Control de Mermas y Hurtos"
         HelpContextID   =   1
         Index           =   3
         Begin VB.Menu inve1 
            Caption         =   "Comprobante Rebaja de Mermas"
            Index           =   1
         End
         Begin VB.Menu inve1 
            Caption         =   "Comprobante Rebaja de Hurtos"
            Index           =   2
         End
         Begin VB.Menu inve1 
            Caption         =   "Listado de Mermas"
            Index           =   3
         End
         Begin VB.Menu inve1 
            Caption         =   "Listado de Hurtos"
            Index           =   4
         End
      End
      Begin VB.Menu guia 
         Caption         =   "Módulo Control de Consumo Interno"
         HelpContextID   =   1
         Index           =   4
         Begin VB.Menu inve2 
            Caption         =   "Maestro de Centros de Consumo Interno"
            Index           =   1
         End
         Begin VB.Menu inve2 
            Caption         =   "Comprobante de Rebaja de Consumos Internos"
            Index           =   2
         End
         Begin VB.Menu inve2 
            Caption         =   "Listado de Consumos Internos  de Todos los Centros de Consumo"
            Index           =   3
         End
         Begin VB.Menu inve2 
            Caption         =   "Listado de Consumos Internos por Centro de Consumo"
            Index           =   4
         End
      End
      Begin VB.Menu guia 
         Caption         =   "Modulo Devoluciones a Proveedor"
         Index           =   6
         Begin VB.Menu devo 
            Caption         =   "Guia de Devolucion de Mercaderias"
            Index           =   1
         End
         Begin VB.Menu devo 
            Caption         =   "Listado de Devolucion de Mercaderias"
            Index           =   2
         End
      End
   End
   Begin VB.Menu PROCESOS 
      Caption         =   "&PROCESOS"
      Begin VB.Menu proceso 
         Caption         =   "Actualiza Stock"
         Index           =   3
      End
      Begin VB.Menu proceso 
         Caption         =   "Cierre Anual"
         Index           =   4
      End
   End
   Begin VB.Menu info 
      Caption         =   "&INFORMES"
      Begin VB.Menu informesmaestros 
         Caption         =   "Menu Informes de Archivos Maestros"
      End
      Begin VB.Menu infoestadistica 
         Caption         =   "Menu Informes Estadistica de Venta"
      End
      Begin VB.Menu estacompra 
         Caption         =   "Menu Informes Estadistica de Compra"
      End
      Begin VB.Menu uti 
         Caption         =   "Menu Informes Estadistica de Utilidades"
      End
      Begin VB.Menu vs 
         Caption         =   "Menu Informes Estadistica de Compras v/s Ventas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu infocostos 
         Caption         =   "Menu Informes de  Utilidades de Venta"
      End
      Begin VB.Menu infogestion 
         Caption         =   "Menu Informes Gestion Comercial"
         Index           =   0
         Begin VB.Menu infoestages 
            Caption         =   "Listado de Productos Bajo Stock Critico"
            Index           =   1
         End
         Begin VB.Menu infoestages 
            Caption         =   "Listado de Productos No Presentan Venta"
            Index           =   2
         End
         Begin VB.Menu infoestages 
            Caption         =   "Listado de Stock Valorizado"
            Index           =   3
         End
         Begin VB.Menu infoestages 
            Caption         =   "Listado Comparacion Bodega v/s Sala Venta"
            Index           =   4
         End
      End
      Begin VB.Menu informes 
         Caption         =   "Cartola  de Movimientos de Productos"
         Index           =   1
      End
      Begin VB.Menu informes 
         Caption         =   "Informe de Reposicion"
         Index           =   3
      End
   End
   Begin VB.Menu tomas 
      Caption         =   "&TOMA INVENTARIO"
      Begin VB.Menu inventario 
         Caption         =   "Ingreso Comprobantes de Ajuste"
         HelpContextID   =   1
         Index           =   1
         Begin VB.Menu ajuste 
            Caption         =   "Comprobante Ajuste de Ingreso Manual"
            Index           =   1
         End
         Begin VB.Menu ajuste 
            Caption         =   "Comprobante Ajuste de Egreso Manual"
            Index           =   2
         End
         Begin VB.Menu ajuste 
            Caption         =   "Comprobante Ajuste de Ingreso Automático"
            Index           =   3
         End
         Begin VB.Menu ajuste 
            Caption         =   "Comprobante Ajuste de Egreso Automático"
            Index           =   4
         End
      End
      Begin VB.Menu inventario 
         Caption         =   "Ingreso Comprobantes Toma Inventario"
         Index           =   2
      End
      Begin VB.Menu inventario 
         Caption         =   "Procesos"
         HelpContextID   =   1
         Index           =   3
         Begin VB.Menu stock 
            Caption         =   "Cuadra Stock Automaticos a cero"
            Index           =   1
         End
         Begin VB.Menu stock 
            Caption         =   "Toma Foto a Inventario"
            Index           =   2
         End
         Begin VB.Menu stock 
            Caption         =   "Compara Inventarios"
            Index           =   3
         End
      End
      Begin VB.Menu inventario 
         Caption         =   "Informes"
         Index           =   4
         Begin VB.Menu invein 
            Caption         =   "Listado de articulos no Contados"
            Index           =   1
         End
         Begin VB.Menu invein 
            Caption         =   "Planillas Para tomar Inventario"
            Index           =   2
         End
         Begin VB.Menu invein 
            Caption         =   "Listado de Perdidas Validadas"
            Enabled         =   0   'False
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu invein 
            Caption         =   "Cartola Tomas de Inventario"
            Index           =   4
         End
      End
   End
   Begin VB.Menu seguridad 
      Caption         =   "&SEGURIDAD"
      Begin VB.Menu permisos 
         Caption         =   "Módulo de Permisos de Usuario"
         Index           =   1
      End
   End
   Begin VB.Menu conf 
      Caption         =   "&CONFIGURACION"
      Begin VB.Menu confi 
         Caption         =   "Configura Sistema"
         Index           =   1
      End
      Begin VB.Menu confi 
         Caption         =   "Cambia Fecha"
         Index           =   2
      End
      Begin VB.Menu confi 
         Caption         =   "Mantencion de Usuarios"
         Index           =   4
      End
      Begin VB.Menu confi 
         Caption         =   "Historico Eventos de Seguridad"
         Index           =   5
      End
      Begin VB.Menu confi 
         Caption         =   "Cerrar Sesión"
         Index           =   6
      End
      Begin VB.Menu confi 
         Caption         =   "TRASPASA DATOS"
         Index           =   7
      End
      Begin VB.Menu confi 
         Caption         =   "Salir"
         Index           =   8
      End
   End
End
Attribute VB_Name = "moduloseguridadconta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FORMATOGRILLA(12, 12)
Private VARIABLE As String
Private USUARIOSELECCIONADO As String
Private menuseleccion As String
Private modifo As Double
Private eli As Boolean

Private Sub CANCELAR_Click()
dato3.text = ""
dato4.text = ""
FrameXp3.Visible = False
End Sub

Private Sub Command1_Click()
If Grid3.Cell(1, 2).text <> "" Then
grabarusuario
Command1.Enabled = False
Command4.Enabled = False
LEERUSUARIOS
Else
Grid3.Cell(1, 2).SetFocus
End If
End Sub

Private Sub COMMAND2_Click()
Grid3.Cell(1, 2).text = ""
Grid3.Cell(2, 2).text = ""
Grid3.Cell(3, 2).text = ""
Grid3.Cell(4, 2).text = ""
Grid3.Cell(5, 2).text = ""
Grid3.Cell(6, 2).text = ""
Grid3.Cell(1, 2).SetFocus
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
FrameXp3.Visible = True
If Grid2.Cell(Grid2.ActiveCell.row, 1).text = "USUARIOS" Then
dato3.text = ""
Else

dato3.text = Grid2.Cell(Grid2.ActiveCell.row, 1).text
End If

dato3.SetFocus

End Sub

Private Sub Command4_Click()
Grid3.Cell(1, 2).SetFocus
Command1.Enabled = True
End Sub

Private Sub COPIAR_Click()
 Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim LINEAS As Double
    
If MsgBox("ESTA SEGURO QUE DESEA COPIAR", vbOKCancel, "ADVERTENCIA") = vbOK Then
        Set csql2.ActiveConnection = conta
        csql2.sql = "DELETE "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta.segu_permisos "
        csql2.sql = csql2.sql + "where usuario='" + dato4.text + "' "
        csql2.Execute
        Call sincronizadatos(csql2.sql, conta, "")
        
        Call copiarpermisos(dato3.text, dato4.text)
        dato4.text = ""
        
        dato4.SetFocus
        
End If

End Sub

 Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
 End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas2(KeyCode, dato3)
End Sub
Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           dato4.SetFocus
        End If
End Sub

Private Sub dato4_GotFocus()
        Call cargatexto(dato4)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas2(KeyCode, dato3)
End Sub
 Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           COPIAR_Click
        End If
    End Sub

Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption


End Sub

Private Sub Form_Load()
Dim k As Integer

  '==================================
    'PERMITE UNA INSTANCIA DEL SISTEMA
    '==================================
    Dim saveTitle$
    If App.PrevInstance Then
        saveTitle$ = App.Title
        App.Title = "... duplicate instance."
        Me.Caption = "... duplicate instance."
        AppActivate saveTitle$
        Sendkeys "% R", True
        End
    End If
'    ==================================
'    PERMITE UNA INSTANCIA DEL SISTEMA
'    ==================================
''
'Close 20
'Open "c:\configu.txt" For Input As #20
'Input #20, SS
'    servidor = SS
'Input #20, SS
'
'    USUARIO = SS
'Input #20, SS
'    password = SS
'Input #20, SS
'    clientesistema = SS


'servidor = "localhost"
'USUARIO = "root"
'password = "123"

'servidor = "164.77.237.204"
'USUARIO = "prueba"
'password = ""
'If Format(fechasistema, "yyyy") > "2007" Then clientesistema = "molino2_"

    
'    basedatos = clientesistema + "conta"
'    Call Conectarconta(servidor, basedatos, USUARIO, password)
'
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda General"
    rutaUpdate = "i:\Actualizaciones"
    'Call verificarUpdate


Call CARGAGRILLAPERMISOS(6, 7)
Call CARGAGRILLAUSUARIOS(2, 2)
Call CARGAGRILLADATOS(7, 3)
Call CARGAGRILLAEMPRESA(10, 4)
leerempresa2
'For K = 1 To ingresos.Count
'ingresos(K).Checked = False
'
'Next K
LEERUSUARIOS
Call MENU1_Click




End Sub

Sub CARGAGRILLADATOS(row, col)
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "DATOS  "
    FORMATOGRILLA(1, 2) = "INGRESAR"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    
    Grid3.Cols = col
    Grid3.Rows = row
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    Grid3.Column(0).Width = 0
    
    Grid3.Column(1).Width = 10 * 10
    Grid3.Column(2).Width = 10 * 10
    
    Grid3.Cell(1, 1).text = "RUT"
    Grid3.Cell(2, 1).text = "USUARIO"
    Grid3.Cell(3, 1).text = "CLAVE"
    Grid3.Cell(4, 1).text = "NOMBRE"
    Grid3.Cell(5, 1).text = "LABOR"
    Grid3.Cell(6, 1).text = "EMAIL"
    
    Grid3.Column(1).Locked = True
    
    
    
    
End Sub

Sub CARGAGRILLAPERMISOS(row, col)
    Rem DATOS DE LA COLUMNA
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "40"
    FORMATOGRILLA(2, 2) = "2"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    
    Grid1.Cols = col
    Grid1.Rows = row
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.Column(0).Width = 0
    
    Grid1.Column(1).Width = 50 * 10
    Grid1.Column(1).Locked = True
    
    For k = 2 To Grid1.Cols - 1
    Grid1.Column(k).Width = 9 * 10
    Grid1.Column(k).CellType = cellCheckBox
   Next k
   Grid1.Cell(0, 1).text = "MODULO DEL SISTEMA"
   Grid1.Cell(0, 2).text = "INGRESAR"
   Grid1.Cell(0, 3).text = "AGREGAR"
   Grid1.Cell(0, 4).text = "MODIFICAR"
   Grid1.Cell(0, 5).text = "ELIMINAR"
   Grid1.Cell(0, 6).text = "SUPERVISOR"
   
    
End Sub

Sub CARGAGRILLAEMPRESA(row, col)
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "EMPRESA"
    FORMATOGRILLA(1, 3) = "ACTIVO"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    
    Grid4.Cols = col
    Grid4.Rows = row
    Grid4.AllowUserResizing = False
    Grid4.DisplayFocusRect = False
    Grid4.ExtendLastCol = True
    Grid4.BoldFixedCell = False
    Grid4.DrawMode = cellOwnerDraw
    Grid4.Appearance = Flat
    Grid4.ScrollBarStyle = Flat
    Grid4.FixedRowColStyle = Flat
    Grid4.Column(0).Width = 0
    
    Grid4.Column(1).Width = 4 * 10
    Grid4.Column(2).Width = 15 * 10
    
    Grid4.Cell(0, 1).text = "CODIGO"
    Grid4.Cell(0, 2).text = "NOMBRE"
 

    Grid4.Column(3).Width = 2 * 10
    Grid4.Column(3).CellType = cellCheckBox
    

    
    
End Sub


Sub CARGAGRILLAUSUARIOS(row, col)
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "NOMBRE"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    
    Grid2.Cols = col
    Grid2.Rows = row
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.Column(0).Width = 0
    
    Grid2.Column(1).Width = 10 * 10
    Grid2.Cell(0, 1).text = "USUARIOS"
    
    Grid2.Column(1).Locked = True
    
    
    
    
    
End Sub




Private Sub Grid1_Click()
If Grid2.Cell(Grid2.ActiveCell.row, 1).text <> "" Then

Call GrabarPermiso(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
End If


End Sub

Private Sub Grid2_Click()
Dim o As Integer

moduloseguridad.Caption = "MODULO DE SEGURIDAD USUARIO ACTIVO =" + Grid2.Cell(Grid2.ActiveCell.row, 1).text
Call LEERUSUARIOindividual(Grid2.Cell(Grid2.ActiveCell.row, 1).text)
USUARIOSELECCIONADO = Grid2.Cell(Grid2.ActiveCell.row, 1).text
MENU1_Click

For o = 1 To Grid1.Rows - 1
Call leerpermisos2(USUARIOSELECCIONADO, Grid1.Cell(o, 1).text, o)

Next o
Command4.Enabled = True
'Call leerpermisos(USUARIOSELECCIONADO)
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
Dim csql2 As New rdoQuery
    Dim csql As New rdoQuery
   
If KeyCode = 46 Then
If MsgBox("ESTA SEGURO QUE DESEA ELIMINAR A " + Grid2.Cell(Grid2.ActiveCell.row, 1).text + " Y SUS PERMISOS", vbOKCancel, "ATENCION") = vbOK Then

        Set csql2.ActiveConnection = conta
        csql2.sql = "DELETE "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta.segu_permisos "
        csql2.sql = csql2.sql + "where usuario='" + Grid2.Cell(Grid2.ActiveCell.row, 1).text + "' "
        csql2.Execute
        Call sincronizadatos(csql2.sql, conta, "")
        
        Set csql.ActiveConnection = conta
        csql.sql = "DELETE "
        csql.sql = csql.sql + "FROM " + clientesistema + "auditoria.segu_usuarios "
        csql.sql = csql.sql + "where usuario='" + Grid2.Cell(Grid2.ActiveCell.row, 1).text + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, conta, "")
        
        
        
End If
End If
LEERUSUARIOS
MENU1_Click

End Sub


Private Sub Grid3_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Grid3_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
NewCol = 2
End Sub



Private Sub ingresos_Click(Index As Integer)
Dim VARIABLE As String
VARIABLE = ingresos.Count



Permiso.Caption = ingresos(Index).Caption
VARIABLE = ingresos(Index).Caption


ingresos(Index).Checked = True
menuseleccion = "ingresos(" & Index & ")"
eli = False
Command3.Visible = True

'Call grabarpermiso("ingresos(" & Index & ")", VARIABLE, False)
'Call leerpermisos(USUARIOSELECCIONADO)






End Sub
Sub GrabarPermiso(nombreprograma As String)
    nombreprograma = achica(nombreprograma)
    campos(0, 0) = "usuario"
    campos(1, 0) = "empresa"
    campos(2, 0) = "programa"
    campos(3, 0) = "ingresa"
    campos(4, 0) = "modifica"
    campos(5, 0) = "elimina"
    campos(6, 0) = "agrega"
    campos(7, 0) = "todas"
    campos(8, 0) = "menu"
    campos(9, 0) = ""
  
    campos(0, 1) = USUARIOSELECCIONADO
    campos(1, 1) = ""
    campos(2, 1) = nombreprograma
    campos(3, 1) = Grid1.Cell(Grid1.ActiveCell.row, 2).text 'ingresa
    campos(4, 1) = Grid1.Cell(Grid1.ActiveCell.row, 4).text 'modificar
    campos(5, 1) = Grid1.Cell(Grid1.ActiveCell.row, 5).text 'eliminar
    campos(6, 1) = Grid1.Cell(Grid1.ActiveCell.row, 3).text 'agregar
    campos(7, 1) = Grid1.Cell(Grid1.ActiveCell.row, 6).text 'supervisor
    campos(8, 1) = ""
    
    campos(0, 2) = "segu_permisos"
    condicion = "usuario=" + "'" + USUARIOSELECCIONADO + "' and programa='" + nombreprograma + "'"
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    sqlconta.audit = True
    sqlconta.programaactivo = Me.Caption
    If ELIMINA = False Then

    op = 5
    Call sqlconta.sqlconta(op, condicion)
  
  
  If sqlconta.status = 4 Then
  op = 2
  condicion = ""
  Else
  op = 3
  End If
  Call sqlconta.sqlconta(op, condicion)
Else
  op = 4
  Call sqlconta.sqlconta(op, condicion)


End If


     
End Sub
Sub LEERUSUARIOS()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim LINEAS As Double
    

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT * "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "auditoria.segu_usuarios "
        csql2.sql = csql2.sql + "order by usuario "
        csql2.Execute
        Grid2.Rows = csql2.RowsAffected + 1
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        LINEAS = 0
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        Grid2.Cell(LINEAS, 1).text = resultados2(1)
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
      
 
  

End Sub
Sub LEERUSUARIOindividual(Usuario)
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim LINEAS As Double
    Dim inicio As Double
    
    

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT * "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "auditoria.segu_usuarios where usuario='" + Usuario + "' "
        csql2.Execute
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        LINEAS = 1
        While Not resultados2.EOF
        
        Grid3.Cell(1, 2).text = resultados2(0)
        Grid3.Cell(2, 2).text = resultados2(1)
        Grid3.Cell(3, 2).text = resultados2(2)
        Grid3.Cell(4, 2).text = resultados2(3)
        Grid3.Cell(5, 2).text = resultados2(4)
        Grid3.Cell(6, 2).text = resultados2(5)
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
       leerempresa2
       
    
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT * "
        csql2.sql = csql2.sql + "FROM segu_empresas where usuario='" + Usuario + "' "
        csql2.sql = csql2.sql + "order by empresa "
        
        csql2.Execute
        
        
        If csql2.RowsAffected > 0 Then
        
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        For inicio = 1 To Grid4.Rows - 1
        If resultados2(1) = Grid4.Cell(inicio, 1).text Then
            Grid4.Cell(inicio, 3).text = resultados2(2)
        End If
        Next inicio
        
        
        resultados2.MoveNext
        LINEAS = LINEAS + 1
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
 
    For k = 1 To 5
    Grid1.Cell(k, 2).text = 0
    Next k

End Sub

Private Function leerempresa2()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim LINEAS As Double
    

        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT * "
        csql2.sql = csql2.sql + "FROM maestroempresas  "
        csql2.Execute
        
        Grid4.Rows = csql2.RowsAffected + 1
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        LINEAS = 0
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        Grid4.Cell(LINEAS, 1).text = resultados2(0)
        Grid4.Cell(LINEAS, 2).text = resultados2(1)
        Grid4.Cell(LINEAS, 3).text = 0
        
        resultados2.MoveNext
       
        Wend
        resultados2.Close
        Set resultados2 = Nothing
        End If
End Function


Private Function leerpermisos2(Usuario, MENU, LINEA)
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim LINEAS As Double
    Dim FINAL As Double
    MENU = achica(MENU)

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT * "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta.segu_permisos "
        csql2.sql = csql2.sql + "where usuario='" + Usuario + "' and programa='" + MENU + "'"
        csql2.Execute
       
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
         
        Grid1.AutoRedraw = False
   
        While Not resultados2.EOF
      
        Grid1.Cell(LINEA, 2).text = resultados2(3)
        Grid1.Cell(LINEA, 3).text = resultados2(4)
        Grid1.Cell(LINEA, 4).text = resultados2(5)
        Grid1.Cell(LINEA, 5).text = resultados2(6)
        Grid1.Cell(LINEA, 6).text = resultados2(8)
        
        resultados2.MoveNext
       
        Wend
        resultados2.Close
        Set resultados2 = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
  
        
       
        End If
End Function
Private Function achica(palabra) As String
Dim inicio As Double
Dim FINAL As Double
For k = 1 To Len(palabra)
If Mid(palabra, k, 1) <> Chr(32) Then inicio = k: Exit For

Next k

achica = Mid(palabra, inicio, Len(palabra) - inicio)

End Function

Sub ACTIVAMENU(ByVal opcion As String)
 
End Sub



Private Sub MENU1_Click()
Dim contador As Double
Dim inicio As Double
Dim FINAL As Double
Dim pasar As Double
Dim NIVEL As String
Dim NIVELBANDERA As String
Call CARGAGRILLAPERMISOS(6, 7)

Close 20

Open App.path + "\principal.txt" For Input As #20
Grid1.Rows = 1
pasar = 0
While Not EOF(20)
Line Input #20, varipaso
If contador = 1 Then
For k = 1 To Len(varipaso)

If Mid(varipaso, k, 1) = Chr(34) Then
varipaso = Mid(varipaso, k + 1, 80)
k = Len(varipaso) + 1
End If

Next k
For k = 1 To Len(varipaso)

If Mid(varipaso, k, 1) = Chr(34) Then
varipaso = Mid(varipaso, 1, k)
k = Len(varipaso) + 1
End If

Next k
varipaso = Replace(varipaso, Chr(34), " ")
Grid1.Rows = Grid1.Rows + 1
If NIVELBANDERA = "0" Then

Rem Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 7).Merge
Grid1.Cell(Grid1.Rows - 1, 1).Font.Bold = True



End If



Grid1.Cell(Grid1.Rows - 1, 1).text = NIVEL + varipaso
contador = 0
End If

For k = 1 To Len(varipaso) - 13
If UCase(Mid(varipaso, k, 13)) = "BEGIN VB.MENU" Then
        
        If k = 4 Then
        NIVEL = ""
        NIVELBANDERA = "0"
        contador = 1
        End If
        
        
        
        If k = 7 Then
        NIVELBANDERA = "1"
        NIVEL = "       "
        contador = 1
        End If
        
        If k = 10 Then
        NIVELBANDERA = "3"
        NIVEL = "               "
        contador = 1
        End If

Exit For
Else
contador = 0

End If



Next k

'

Wend
 
End Sub

Sub grabarusuario()
    
    campos(0, 0) = "rut"
    campos(1, 0) = "usuario"
    campos(2, 0) = "clave"
    campos(3, 0) = "nombre"
    campos(4, 0) = "labor"
    campos(5, 0) = "email"
    campos(6, 0) = ""
  
    campos(0, 1) = Grid3.Cell(1, 2).text
    campos(1, 1) = Grid3.Cell(2, 2).text
    campos(2, 1) = Grid3.Cell(3, 2).text
    campos(3, 1) = Grid3.Cell(4, 2).text
    campos(4, 1) = Grid3.Cell(5, 2).text
    campos(5, 1) = Grid3.Cell(6, 2).text
    
   
    
    campos(0, 2) = clientesistema + "auditoria.segu_usuarios"
    condicion = "usuario=" + "'" + Grid3.Cell(2, 2).text + "' "
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
   

    op = 5
    Call sqlconta.sqlconta(op, condicion)
  
  
  If sqlconta.status = 4 Then
  op = 2
  condicion = ""
  Else
  op = 3
  End If
  Call sqlconta.sqlconta(op, condicion)
    
End Sub
Sub copiarpermisos(usuarioorigen, usuariodestino)
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery

        Set csql2.ActiveConnection = contadb
        csql2.sql = "INSERT INTO " & clientesistema & "conta.segu_permisos (usuario,programa,ingresa,agrega,modifica,elimina,autoriza,todas) "
        csql2.sql = csql2.sql + "SELECT  '" + usuariodestino + "',programa,ingresa,agrega,modifica,elimina,autoriza,todas "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta.segu_permisos "
        csql2.sql = csql2.sql + "where usuario='" + usuarioorigen + "'"
        csql2.Execute
        Call sincronizadatos(csql2.sql, conta, "")
         
End Sub
