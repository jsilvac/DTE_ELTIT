VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LCobranza 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cobranza"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   15270
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   240
      Top             =   7665
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
      LockType        =   -1
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1980
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   3493
      BackColor       =   16744576
      Caption         =   "Tipo de Listado"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin XPFrame.FrameXp frmGeneral 
         Height          =   1455
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2566
         BackColor       =   16744576
         Caption         =   "General"
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
         Alignment       =   1
         Begin VB.CheckBox chkSucursalVendedor 
            BackColor       =   &H00FF8080&
            Caption         =   "Sucursales Separadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2220
            TabIndex        =   23
            Top             =   720
            Value           =   1  'Checked
            Width           =   2250
         End
         Begin VB.TextBox dato1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   2
            Tag             =   "proveedor"
            Top             =   720
            Width           =   375
         End
         Begin VB.ComboBox cmbVendedores 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "LCobranza.frx":0000
            Left            =   1680
            List            =   "LCobranza.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   2715
         End
         Begin VB.Label lblNombreV 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   4320
         End
         Begin VB.Label lblDVV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1710
            TabIndex        =   20
            Top             =   720
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label lbl3 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Vendedor"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbl2 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Listado Por"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Documentos Vencidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos los Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.ComboBox cmbTipoListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "LCobranza.frx":002E
         Left            =   1815
         List            =   "LCobranza.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin XPFrame.FrameXp frmIndividual 
         Height          =   1455
         Left            =   8880
         TabIndex        =   13
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   2566
         BackColor       =   16744576
         Caption         =   "Individual"
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
         Alignment       =   1
         Begin VB.TextBox dato3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   4
            Tag             =   "proveedor"
            Top             =   720
            Width           =   420
         End
         Begin VB.CheckBox chkSucursal 
            BackColor       =   &H00FF8080&
            Caption         =   "Sucursal Individual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3720
            TabIndex        =   16
            Top             =   360
            Value           =   1  'Checked
            Width           =   1950
         End
         Begin VB.TextBox dato2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   3
            Tag             =   "proveedor"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblNombreC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   5520
         End
         Begin VB.Label lbl5 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Sucursal"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblDVC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   3150
            TabIndex        =   15
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lbl4 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Rut"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo de Listado"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   330
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   5400
      Left            =   60
      TabIndex        =   9
      Top             =   2160
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   9525
      BackColor       =   16744576
      Caption         =   "Listado"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin FlexCell.Grid Documentos 
         Height          =   4830
         Left            =   120
         TabIndex        =   7
         Top             =   405
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   8520
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   11820
      TabIndex        =   22
      Top             =   7680
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "I   M   P   R   I   M   I   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "LCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 11) As String
    Private tipoListado As String
    Private tipoVendedor As String
    Private TIPOCLIENTE As String
    Private CAMPOS(5, 5) As String
Private saldog As Double
Private abonog As Double
Private montog As Double
Private acumuladog As Double

Private Sub chkSucursal_Click()
    If chkSucursal.Value = 0 Then
        TIPOCLIENTE = "SIN"
        dato3.text = ""
        If opt1.Value = True Then
            Call cargaInformeClienteSinTodos
        Else
            Call cargaInformeClienteSinVencidos
        End If
    Else
        TIPOCLIENTE = "CON"
        dato3.Enabled = True
        dato3.Locked = False
        dato3.SetFocus
    End If
End Sub

Private Sub cmbTipoListado_Click()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    cmbTipoListado_KeyPress (13)
End Sub

Private Sub cmbTipoListado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then    'arriba
        If cmbTipoListado.ListIndex <= 0 Then
            Unload Me
        End If
    End If
    'If KeyCode = 40 Then    'abajo
    '    If cmbTipoListado.ListIndex = cmbTipoListado.ListCount - 1 Then
    '        SendKeys "{Tab}"
    '    End If
    'End If
End Sub

Private Sub cmbTipoListado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbTipoListado.ListIndex = 0 Then
            frmGeneral.Enabled = True
            frmIndividual.Enabled = False
            'cmbVendedores.Enabled = True
            'dato1.Enabled = True
            'dato2.text = ""
            'dato3.text = ""
            'dato2.Enabled = False
            'dato3.Enabled = False
            SendKeys "{Tab}"
        End If
        If cmbTipoListado.ListIndex = 1 Then
            frmIndividual.Enabled = True
            frmGeneral.Enabled = False
            'cmbVendedores.Enabled = False
            'dato1.text = ""
            'dato1.Enabled = False
            'dato2.Enabled = True
            'dato3.Enabled = True
            SendKeys "{Tab}"
        End If
        tipoListado = cmbTipoListado.List(cmbTipoListado.ListIndex)
    End If
End Sub

Private Sub cmbVendedores_Click()
    dato1.text = ""
    Call cmbVendedores_KeyPress(13)
End Sub

Private Sub cmbVendedores_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then    'arriba
        If cmbTipoListado.ListIndex < 0 Then
            Call Flechas(KeyCode, cmbTipoListado)
        End If
    End If
End Sub

Private Sub cmbVendedores_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbVendedores.ListIndex = 0 Then
            dato1.Enabled = False
            If opt1.Value = True Then
                Call cargaInformeVendedoresTodos
            Else
                Call cargaInformeVendedoresVencidos
            End If
            'SendKeys "{Tab}"
        End If
        If cmbVendedores.ListIndex = 1 Then
            dato1.Enabled = True
            SendKeys "{Tab}"
        End If
        tipoVendedor = cmbVendedores.List(cmbVendedores.ListIndex)
    End If
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Vendedor"
    End Sub
    
    Private Sub dato2_GotFocus()
        Call selecciona(dato2)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    
    Private Sub dato3_GotFocus()
        Call selecciona(dato3)
        'Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaVendedores(dato1)
        Else
            Call Flechas(KeyCode, cmbTipoListado)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCliente(dato2, dato3, lblDVC)
        Else
            Call Flechas(KeyCode, cmbTipoListado)
        End If
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
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
            dato1.text = ceros(dato1)
            lblDVV.Caption = rut(dato1.text)
            lblNombreV.Caption = leerNombreVendedor(dato1.text)
            If lblNombreV.Caption <> "" Then
                If opt1.Value = True Then
                    Call cargaInformeVendedorTodos
                Else
                    Call cargaInformeVendedorVencidos
                End If
                'SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            lblDVC.Caption = rut(dato2.text)
            lblNombreC.Caption = leerNombreCliente(dato2.text & lblDVC.Caption)
            If lblNombreC.Caption <> "" Then
                If chkSucursal.Value = 0 Then
                    If opt1.Value = True Then
                        Call cargaInformeClienteSinTodos
                    Else
                        Call cargaInformeClienteSinVencidos
                    End If
                Else
                    If opt1.Value = True Then
                        Call cargaInformeClienteConTodos
                    Else
                        Call cargaInformeClienteConVencidos
                    End If
                    SendKeys "{Tab}"
                End If
            End If
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        Dim cad As String
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            cad = leerNombreClienteSucursal(dato2.text & lblDVC.Caption, dato3.text)
            If cad <> "" Then
                If opt1.Value = True Then
                    Call cargaInformeClienteConTodos
                Else
                    Call cargaInformeClienteConVencidos
                End If
            End If
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato2_LostFocus()
        Call limpiaBarra(2)
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        'If KeyCode = 38 Then
        '    If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
        '        Unload Me
        '    End If
        'End If
    End Sub
    
    Private Sub Form_Load()
        Call Centrar(Me)
        Call CargaGrillaDocumentos(1, 11)
        TIPOCLIENTE = "CON"
        frmGeneral.Enabled = False
        frmIndividual.Enabled = False
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaDocumentos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "DOCUMENTO"
        formatogrilla(1, 2) = "FECHA"
        formatogrilla(1, 3) = "VEND."
        formatogrilla(1, 4) = "RAZON SOCIAL"
        formatogrilla(1, 5) = "VENCE"
        formatogrilla(1, 6) = "MONTO"
        formatogrilla(1, 7) = "ABONO"
        formatogrilla(1, 8) = "SALDO"
        formatogrilla(1, 9) = "ACUMULADO"
        formatogrilla(1, 10) = "OBS"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "13"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "5"
        formatogrilla(2, 4) = "50"
        formatogrilla(2, 5) = "10"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        formatogrilla(2, 8) = "9"
        formatogrilla(2, 9) = "9"
        formatogrilla(2, 10) = "50"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "C"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "S"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = ""
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
        formatogrilla(4, 10) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "FALSE"
        formatogrilla(5, 9) = "FALSE"
        formatogrilla(5, 10) = "FALSE"
                
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        formatogrilla(6, 10) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
        formatogrilla(7, 10) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "9"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "5"
        formatogrilla(8, 4) = "12"
        formatogrilla(8, 5) = "7"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "9"
        formatogrilla(8, 10) = "3"
            
        Documentos.Cols = col
        Documentos.Rows = row
        Documentos.AllowUserResizing = False
        Documentos.DisplayFocusRect = False
        Documentos.ExtendLastCol = True
        Documentos.BoldFixedCell = False
        Documentos.DrawMode = cellOwnerDraw
        Documentos.Appearance = Flat
        Documentos.ScrollBarStyle = Flat
        Documentos.FixedRowColStyle = Flat
        Documentos.BackColorFixed = RGB(90, 158, 214)
        Documentos.BackColorFixedSel = RGB(110, 180, 230)
        Documentos.BackColorBkg = RGB(90, 158, 214)
        Documentos.BackColorScrollBar = RGB(231, 235, 247)
        Documentos.BackColor1 = RGB(231, 235, 247)
        Documentos.BackColor2 = RGB(239, 243, 255)
        Documentos.GridColor = RGB(148, 190, 231)
        
        Documentos.Column(0).Width = 0
        For i = 1 To col - 1
            Documentos.Cell(0, i).text = formatogrilla(1, i)
            Documentos.Column(i).Width = Val(formatogrilla(8, i)) * (Documentos.Cell(0, i).Font.Size + 1.25)
            Documentos.Column(i).MaxLength = Val(formatogrilla(2, i))
            Documentos.Column(i).FormatString = formatogrilla(4, i)
            Documentos.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Documentos.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                Documentos.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                Documentos.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Alignment = cellCenterCenter
        Documentos.Enabled = True
        
        'Documentos.AddItem "FV 1234567890" & vbTab & "2007-01-01" & vbTab & "0135744387" & vbTab & "RODRIGO CAMPOS" & vbTab & "2007-02-01" & vbTab & "123456789" & vbTab & "123456789" & vbTab & "123456789" & vbTab & "123456789" & vbTab & "hola", True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'============================================================
'CARGA INFORME VENDEDORES TODOS LOS DOCUMENTOS
'============================================================
    Private Sub cargaInformeVendedoresTodos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As String
        Dim filaIni As Double
        Dim cad As String
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
                
        Documentos.Range(0, 3, 0, 4).MergeCells = False
        Documentos.Cell(0, 3).text = "VEND."
        Documentos.Cell(0, 4).text = "RAZON SOCIAL"
                
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo, dc.vendedor, dc.sucursal "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono <> dc.monto AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal, dc.vencimiento ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        sucursal = "0"
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If chkSucursalVendedor.Value = 0 Then
                    sucursal = data.Recordset.Fields("sucursal")
                End If
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    cad = "V: " & data.Recordset.Fields("vendedor")
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
        sumatotales
        Documentos.AutoRedraw = True
        
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME VENDEDORES TODOS LOS DOCUMENTOS
'============================================================

'============================================================
'CARGA INFORME VENDEDOR TODOS LOS DOCUMENTOS
'============================================================
    Private Sub cargaInformeVendedorTodos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As String
        Dim filaIni As Double
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
        Dim cad As String
        
        Documentos.Range(0, 3, 0, 4).Merge
        Documentos.Cell(0, 3).text = "RAZON SOCIAL"
        
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.sucursal, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono < dc.monto AND dc.vendedor = '" & dato1.text & "' AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal, dc.vencimiento ASC"
        
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If chkSucursalVendedor.Value = 0 Then
                    sucursal = data.Recordset.Fields("sucursal")
                End If
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                    
                    If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & leerNombreClienteSucursal(rut, sucursal) & vbTab & Replace(data.Recordset.Fields("item2"), ",", "."), True
                        Documentos.Range(Documentos.Rows - 1, 3, Documentos.Rows - 1, 4).Merge
                    End If
                    If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & leerNombreClienteSucursal(rut, sucursal) & vbTab & Replace(data.Recordset.Fields("item2"), ",", "."), True
                        Documentos.Range(Documentos.Rows - 1, 3, Documentos.Rows - 1, 4).Merge
                    End If
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
        sumatotales
        Documentos.AutoRedraw = True
        
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME VENDEDOR TODOS LOS DOCUMENTOS
'============================================================

'============================================================
'CARGA INFORME VENDEDORES DOCUMENTOS VENCIDOS
'============================================================
    Private Sub cargaInformeVendedoresVencidos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As Double
        Dim filaIni As Double
        Dim cad As String
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
        
        Documentos.Range(0, 3, 0, 4).MergeCells = False
        Documentos.Cell(0, 3).text = "VEND."
        Documentos.Cell(0, 4).text = "RAZON SOCIAL"
        
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo, dc.vendedor,dc.sucursal "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono < dc.monto AND dc.vencimiento < '" & Format(fechasistema, "yyyy-mm-dd") & "' AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal, dc.vencimiento ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        sucursal = "0"
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If chkSucursalVendedor.Value = 0 Then
                    sucursal = data.Recordset.Fields("sucursal")
                End If
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    cad = "V: " & data.Recordset.Fields("vendedor")
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                    If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                    If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
  
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
        sumatotales
        
        Documentos.AutoRedraw = True
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME VENDEDORES DOCUMENTOS VENCIDOS
'============================================================

'============================================================
'CARGA INFORME VENDEDOR DOCUMENTOS VENCIDOS
'============================================================
    Private Sub cargaInformeVendedorVencidos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As String
        Dim filaIni As Double
        Dim cad As String
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
        
        Documentos.Range(0, 3, 0, 4).MergeCells = False
        Documentos.Cell(0, 3).text = "RAZON SOCIAL"
        
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.sucursal, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo  "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono < dc.monto AND dc.vendedor = '" & dato1.text & "' AND dc.vencimiento < '" & Format(fechasistema, "yyyy-mm-dd") & "' AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal, dc.vencimiento ASC"
        
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If chkSucursalVendedor.Value = 0 Then
                    sucursal = data.Recordset.Fields("sucursal")
                End If
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                    If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & leerNombreClienteSucursal(rut, sucursal) & vbTab & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                    If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & leerNombreClienteSucursal(rut, sucursal) & vbTab & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
        
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
        sumatotales

        Documentos.AutoRedraw = True
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME VENDEDOR DOCUMENTOS VENCIDOS
'============================================================

'============================================================
'CARGA INFORME CLIENTE SIN SUCURSAL TODOS LOS DOCUMENTOS
'============================================================
    Private Sub cargaInformeClienteSinTodos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As String
        Dim filaIni As Double
        Dim cad As String
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
        
        Documentos.Range(0, 3, 0, 4).MergeCells = False
        Documentos.Cell(0, 3).text = "VEND."
        Documentos.Cell(0, 4).text = "RAZON SOCIAL"
        
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.sucursal, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo, dc.vendedor "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono < dc.monto AND dc.rut = '" & dato2.text & lblDVC.Caption & "' AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal, dc.vencimiento ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    cad = data.Recordset.Fields("vendedor")
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                    
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
        sumatotales
        Documentos.AutoRedraw = True
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME CLIENTE SIN SUCURSAL TODOS LOS DOCUMENTOS
'============================================================

'============================================================
'CARGA INFORME CLIENTE SIN SUCURSAL DOCUMENTOS VENCIDOS
'============================================================
    Private Sub cargaInformeClienteSinVencidos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As String
        Dim filaIni As Double
        Dim cad As String
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
        
        Documentos.Range(0, 3, 0, 4).MergeCells = False
        Documentos.Cell(0, 3).text = "VEND."
        Documentos.Cell(0, 4).text = "RAZON SOCIAL"
        
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.sucursal, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo, dc.vendedor "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono < dc.monto AND dc.rut = '" & dato2.text & lblDVC.Caption & "' AND dc.vencimiento < '" & Format(fechasistema, "yyyy-mm-dd") & "' AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal ,dc.vencimiento ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    cad = data.Recordset.Fields("vendedor")
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                           If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
        sumatotales
        Documentos.AutoRedraw = True
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME CLIENTE SIN SUCURSAL DOCUMENTOS VENCIDOS
'============================================================

'============================================================
'CARGA INFORME CLIENTE CON SUCURSAL TODOS LOS DOCUMENTOS
'============================================================
    Private Sub cargaInformeClienteConTodos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As String
        Dim filaIni As Double
        Dim cad As String
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
        
        Documentos.Range(0, 3, 0, 4).MergeCells = False
        Documentos.Cell(0, 3).text = "VEND."
        Documentos.Cell(0, 4).text = "RAZON SOCIAL"
        
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.sucursal, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo, dc.vendedor "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono < dc.monto AND dc.rut = '" & dato2.text & lblDVC.Caption & "' AND dc.sucursal = '" & dato3.text & "' AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal , dc.vencimiento ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    cad = data.Recordset.Fields("vendedor")
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                    If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
sumatotales
        Documentos.AutoRedraw = True
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME CLIENTE CON SUCURSAL TODOS LOS DOCUMENTOS
'============================================================

'============================================================
'CARGA INFORME CLIENTE CON SUCURSAL DOCUMENTOS VENCIDOS
'============================================================
    Private Sub cargaInformeClienteConVencidos()
        Dim tabla As String
        Dim rut As String
        Dim sucursal As String
        Dim filaIni As Double
        Dim cad As String
        Dim numerofactura As String
        Dim saldo As Double
        Dim cont As Integer
        
        Documentos.Range(0, 3, 0, 4).MergeCells = False
        Documentos.Cell(0, 3).text = "VEND."
        Documentos.Cell(0, 4).text = "RAZON SOCIAL"
        
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & vbTab & "') AS item1, CONCAT('" & vbTab & " ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', '$ ', FORMAT(dc.monto,0), '" & vbTab & "', '$ ', FORMAT(dc.abono,0), '" & vbTab & "', '$ ', FORMAT(dc.monto - dc.abono,0), '" & vbTab & "', '" & vbTab & "', '------------------------------') AS item2, dc.rut, dc.sucursal, dc.numero, dc.vencimiento, dc.monto - dc.abono AS saldo, dc.vendedor "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.abono < dc.monto AND dc.rut = '" & dato2.text & lblDVC.Caption & "' AND dc.sucursal = '" & dato3.text & "' AND dc.vencimiento < '" & Format(fechasistema, "yyyy-mm-dd") & "' AND dc.tipo <> 'GD' "
        tabla = tabla & "ORDER BY dc.rut, dc.sucursal, dc.vencimiento ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            rut = data.Recordset.Fields("rut")
            sucursal = data.Recordset.Fields("sucursal")
            filaIni = 1
            cont = 0
            While Not data.Recordset.EOF
                If rut = data.Recordset.Fields("rut") And sucursal = data.Recordset.Fields("sucursal") Then
                    cad = data.Recordset.Fields("vendedor")
                    saldo = CDbl(data.Recordset.Fields("saldo"))
                    numerofactura = data.Recordset.Fields("numero")
                    If Mid(data.Recordset.Fields("item1"), 1, 2) <> "NV" And saldo + leerNotaCreditoFactura(numerofactura) > 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
                    If Mid(data.Recordset.Fields("item1"), 1, 2) = "NV" And saldo <> 0 Then
                        cont = cont + 1
                        Documentos.AddItem data.Recordset.Fields("item1") & cad & vbTab & leerNombreClienteSucursal(rut, sucursal) & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    End If
         
                Else
                    If cont > 0 Then
                        Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                        Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                        Call sumaSaldos(filaIni, Documentos.Rows - 2)
                        Documentos.AddItem "", True
                        Call leerDatosCliente(rut, sucursal)
                    End If
                    cont = 0
                    filaIni = Documentos.Rows
                    rut = data.Recordset.Fields("rut")
                    sucursal = data.Recordset.Fields("sucursal")
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            If cont > 0 Then
                Documentos.AddItem vbTab & vbTab & vbTab & "TOTAL CLIENTE", True
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Merge
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 5).Alignment = cellCenterCenter
                Documentos.Range(Documentos.Rows - 1, 4, Documentos.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
                Call sumaSaldos(filaIni, Documentos.Rows - 2)
                Documentos.AddItem "", True
                Call leerDatosCliente(rut, sucursal)
            End If
            
        End If
        sumatotales
        Documentos.AutoRedraw = True
        Documentos.Refresh
    End Sub
'============================================================
'CARGA INFORME CLIENTE CON SUCURSAL DOCUMENTOS VENCIDOS
'============================================================

'============================================================
'SUMA SALDOS
'============================================================
Private Sub sumaSaldos(ByVal filaIni As Double, ByVal filaFin As Double)
        Dim MONTO As Double
        Dim abono As Double
        Dim saldo As Double
        Dim ACUMULADO As Double
        Dim i As Long
        MONTO = 0
        abono = 0
        saldo = 0
        ACUMULADO = 0
        For i = filaIni To filaFin
            MONTO = MONTO + CDbl(Documentos.Cell(i, 6).text)
            abono = abono + CDbl(Documentos.Cell(i, 7).text)
            saldo = saldo + CDbl(Documentos.Cell(i, 8).text)
            ACUMULADO = ACUMULADO + CDbl(Documentos.Cell(i, 8).text)
            Documentos.Cell(i, 9).text = Format(ACUMULADO, "$ ###,###,##0")
        Next i
        montog = montog + MONTO
        abonog = abonog + abono
        saldog = saldog + saldo
        acumuladog = acumuladog + ACUMULADO
        
        Documentos.Cell(i, 6).text = Format(MONTO, "$ ###,###,##0")
        Documentos.Cell(i, 7).text = Format(abono, "$ ###,###,##0")
        Documentos.Cell(i, 8).text = Format(saldo, "$ ###,###,##0")
        Documentos.Cell(i, 9).text = Format(ACUMULADO, "$ ###,###,##0")
    End Sub
Private Sub sumatotales()
        Dim MONTO As Double
        Dim abono As Double
        Dim saldo As Double
        Dim ACUMULADO As Double
        Dim i As Long
        
        Documentos.Rows = Documentos.Rows + 2
        i = Documentos.Rows - 1
        
        Documentos.Cell(i, 6).text = "_____________"
        Documentos.Cell(i, 7).text = "_____________"
        Documentos.Cell(i, 8).text = "_____________"
        Documentos.Cell(i, 9).text = "_____________"
        Documentos.Rows = Documentos.Rows + 1
        i = Documentos.Rows - 1
        Documentos.Cell(i, 3).text = "TOTALES "
        Documentos.Cell(i, 6).text = Format(montog, "$ ###,###,##0")
        Documentos.Cell(i, 7).text = Format(abonog, "$ ###,###,##0")
        Documentos.Cell(i, 8).text = Format(saldog, "$ ###,###,##0")
        Documentos.Cell(i, 9).text = Format(acumuladog, "$ ###,###,##0")
        montog = 0
        abonog = 0
        saldog = 0
        acumuladog = 0
        
    
    End Sub

'============================================================
'SUMA SALDOS
'============================================================

'============================================================
'LEER DATOS CLIENTE
'============================================================
    Private Sub leerDatosCliente(ByVal rut As String, ByVal sucursal As String)
        
        Dim op As Integer
        Dim cad As String
        
        cad = Left(rut, 9)
        cad = Format(cad, "###,###,##0") & "-" & Right(rut, 1)
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = "direccion"
        CAMPOS(2, 0) = "ciudad"
        CAMPOS(3, 0) = "fono1"
        CAMPOS(4, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        
        Documentos.AddItem "NOMBRE", True
        Documentos.AddItem "DIRECCION", True
        Documentos.AddItem "RUT" & vbTab & vbTab & vbTab & "CIUDAD" & vbTab & vbTab & vbTab & vbTab & "FONO", True
        Documentos.Cell(Documentos.Rows - 1, 5).Alignment = cellLeftCenter
        Documentos.AddItem "", True
        Documentos.AddItem "", True
        If sql.Status = 0 Then
            Documentos.Cell(Documentos.Rows - 5, 2).text = sql.response(0, 3)
            Documentos.Range(Documentos.Rows - 5, 2, Documentos.Rows - 5, 4).Merge
            Documentos.Range(Documentos.Rows - 5, 2, Documentos.Rows - 5, 4).Alignment = cellLeftCenter
            
            Documentos.Cell(Documentos.Rows - 4, 2).text = sql.response(1, 3)
            Documentos.Range(Documentos.Rows - 4, 2, Documentos.Rows - 4, 4).Merge
            Documentos.Range(Documentos.Rows - 4, 2, Documentos.Rows - 4, 4).Alignment = cellLeftCenter
            
            Documentos.Cell(Documentos.Rows - 3, 2).text = cad
            Documentos.Range(Documentos.Rows - 3, 2, Documentos.Rows - 3, 3).Merge
            Documentos.Range(Documentos.Rows - 3, 2, Documentos.Rows - 3, 3).Alignment = cellCenterCenter
            Documentos.Range(Documentos.Rows - 3, 1, Documentos.Rows - 3, 3).FontBold = True
            
            Documentos.Cell(Documentos.Rows - 3, 5).text = sql.response(2, 3)
            Documentos.Range(Documentos.Rows - 3, 5, Documentos.Rows - 3, 7).Merge
            Documentos.Range(Documentos.Rows - 3, 5, Documentos.Rows - 3, 7).Alignment = cellLeftCenter
            
            Documentos.Cell(Documentos.Rows - 3, 9).text = Format(sql.response(3, 3), "########0")
        End If
    End Sub
'============================================================
'LEER DATOS CLIENTE
'============================================================

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        
        frmImprimir.ColorBarraAbajo = &HFFC0C0
        frmImprimir.ColorBarraArriba = &H800000
        frmImprimir.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &H800000
        frmImprimir.ColorBarraArriba = &HFFC0C0
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub

    Private Sub imprimir()
        Dim defecto As Integer
        Call cabezaInforme(dato1.text, Documentos, "LISTADO DE COBRANZA", 1)
        Documentos.PageSetup.HeaderMargin = 1
        Documentos.PageSetup.TopMargin = 1
        Documentos.PageSetup.LeftMargin = 0.5
        Documentos.PageSetup.RightMargin = 0.5
        Documentos.PageSetup.PrintFixedRow = True
        Documentos.PageSetup.BlackAndWhite = True
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Documentos.PageSetup.Footer = "Pág &P de &N"
        Documentos.PageSetup.FooterAlignment = cellCenter
        
        defecto = Documentos.DefaultFont.Size
        Documentos.DefaultFont.Size = 7
        Call verificaImpresora(5, Documentos)
        Documentos.DefaultFont.Size = defecto
        Documentos.Refresh
    End Sub

Private Sub opt1_Click() 'TODOS LOS DOCUMENTOS
    opt1.Value = True
    If tipoListado = "GENERAL" Then
        If tipoVendedor = "VENDEDOR" Then
            Call cargaInformeVendedorTodos
        Else
            Call cargaInformeVendedoresTodos
        End If
    Else
        If TIPOCLIENTE = "SIN" Then
            Call cargaInformeClienteSinTodos
        End If
        If TIPOCLIENTE = "CON" Then
            Call cargaInformeClienteConTodos
        End If
    End If
End Sub

Private Sub opt2_Click() 'DOCUMENTOS VENCIDOS
    opt2.Value = True
    If tipoListado = "GENERAL" Then
        If tipoVendedor = "VENDEDOR" Then
            Call cargaInformeVendedorVencidos
        Else
            Call cargaInformeVendedoresVencidos
        End If
    Else
        If TIPOCLIENTE = "SIN" Then
            Call cargaInformeClienteSinVencidos
        End If
        If TIPOCLIENTE = "CON" Then
            Call cargaInformeClienteConVencidos
        End If
    End If
End Sub

    




