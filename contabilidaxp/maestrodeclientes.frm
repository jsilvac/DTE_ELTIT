VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ventas01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10710
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   706
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   714
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc mc 
      Height          =   375
      Left            =   9480
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2415
      Left            =   480
      TabIndex        =   17
      Top             =   6360
      Width           =   9615
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   59
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   15
         Left            =   1440
         TabIndex        =   58
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   57
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Usado"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Prorrogas"
         Height          =   255
         Left            =   3600
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   9615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Protestados"
         Height          =   255
         Left            =   3600
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Cupo"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Boletas"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Prorrogas"
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Boletas"
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   27
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   6360
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label40 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5775
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   13
         Left            =   1680
         TabIndex        =   56
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   12
         Left            =   1680
         TabIndex        =   55
         Top             =   4800
         Width           =   255
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   11
         Left            =   1680
         TabIndex        =   54
         Top             =   4440
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   53
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   9
         Left            =   1680
         TabIndex        =   52
         Top             =   3720
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   8
         Left            =   1680
         TabIndex        =   51
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   50
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   6
         Left            =   3960
         TabIndex        =   49
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   48
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   47
         Top             =   2280
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   46
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   45
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   44
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txtmclientes 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   43
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtRut 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Fono (2)"
         Height          =   255
         Left            =   3240
         TabIndex        =   42
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Giro"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Fono (1)"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Comuna"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   5775
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RUT"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso Maestro de Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Credito"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Dscuento"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   5160
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   600
      TabIndex        =   41
      Top             =   9120
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\remuxp\barra_opciones.swf"
      Src             =   "c:\remuxp\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   2415
      Left            =   600
      Top             =   6480
      Width           =   9615
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5775
      Left            =   1800
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "ventas01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

    'Call Limpia_Formulario
    txtRut.SetFocus

End Sub

Private Sub Form_Load()
    Call Conecta_Maestro_Clientes
    Call Conectar_BD



End Sub

Private Sub txtRut_GotFocus()

'-----------------------------------------------------------------
'      COLOR CUANDO TOMA EL FOCO CODIGO
'-----------------------------------------------------------------

    txtRut.BackColor = &HFFFF&

'-----------------------------------------------------------------
End Sub

Private Sub txtRut_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then ayudamaestrodeclientes.Show


End Sub
