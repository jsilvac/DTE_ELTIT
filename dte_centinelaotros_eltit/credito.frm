VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form credito 
   BackColor       =   &H00AE1118&
   BorderStyle     =   0  'None
   Caption         =   "Crédito"
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9763
      BackColor       =   16744576
      Caption         =   "Venta a Crédito"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.CommandButton cmdF1 
         BackColor       =   &H00FFC0C0&
         Caption         =   " F1"
         Height          =   285
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox txtPie 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   1
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtCuotas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7440
         MaxLength       =   9
         TabIndex        =   2
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtRut 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2040
         MaxLength       =   9
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image Image16 
         Height          =   285
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   285
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Opciones de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   29
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2160
         TabIndex        =   27
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   3840
         Width           =   2040
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto de Pie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   1920
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   11
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   1680
      End
      Begin VB.Label lbl3Cuotas 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   420
         Left            =   4680
         TabIndex        =   23
         Top             =   2880
         Width           =   4920
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2160
         TabIndex        =   22
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Cuotas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   4680
         TabIndex        =   21
         Top             =   3360
         Width           =   2640
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "% Factorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   4680
         TabIndex        =   20
         Top             =   3840
         Width           =   2640
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   7440
         TabIndex        =   19
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Cuota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   8
         Left            =   600
         TabIndex        =   18
         Top             =   4320
         Width           =   3720
      End
      Begin VB.Label lblCuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   705
         Left            =   4680
         TabIndex        =   17
         Top             =   4320
         Width           =   4935
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Rut Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1680
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3480
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3960
         TabIndex        =   14
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   7575
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Día de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label lblDiaPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2040
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Autorizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblCupo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Utilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   4
         Left            =   3240
         TabIndex        =   7
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label lblUtilizado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3240
         TabIndex        =   6
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   6480
         TabIndex        =   5
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   6480
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
      End
   End
End
Attribute VB_Name = "credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private color(1) As Variant
    Private trescuotas As Double
    Private saldo As Double
    Public cuotascontado As Double
    Private c As Cliente
    Private cr As Creditos
    Private modifica As Boolean
    
Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
    If KeyCode = vbKeyF1 Then
        opcionesCredito.Show vbModal
    End If
    If KeyCode = vbKeyF2 Then
        Call ayudaCliente(txtRut)
        Call actualizaCuota
    End If
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    color(0) = &HC0&        'rojo
    color(1) = &HFFC0C0     'celeste
    modifica = False
    cuotascontado = 2
    lblMonto.Caption = Format(detallePagos.Pagos.ActiveCell.text, "$ ###,###,##0")
    lbl3Cuotas.Caption = "*** 33% de Pie "
    trescuotas = CDbl(lblMonto.Caption) / 3
    If trescuotas < 100 Then
        trescuotas = 1
    Else
        trescuotas = Round(trescuotas / 100 + 1, 0)
    End If
    trescuotas = trescuotas * 100
    lbl3Cuotas.Caption = lbl3Cuotas.Caption & Format(trescuotas, "$ ###,###,##0") & " ***"
    saldo = CDbl(lblMonto.Caption) - trescuotas
    txtPie.text = Format(trescuotas, "$ ###,###,##0")
    Me.lblCuota = Format(saldo / 2, "$ ###,###,##0")
    txtCuotas.text = "2"
End Sub

Private Sub Label32_Click()
End Sub

Private Sub tmrBlink_Timer()
    Static estado As Integer
    estado = 1 - estado
    lbl3Cuotas.ForeColor = color(estado)
End Sub

Private Sub txtCuotas_Change()
    Call actualizaCuota
End Sub

Private Sub txtCuotas_GotFocus()
    Call selecciona(txtCuotas)
End Sub

Private Sub txtCuotas_KeyDown(KeyCode As Integer, shift As Integer)
    Call Flechas(KeyCode, txtPie)
End Sub

Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And txtCuotas.text <> "" Then
        Call ctrltostruct
        Unload Me
    End If
End Sub

Private Sub txtPie_Change()
    If txtCuotas.text <> "" Then
        Call actualizaCuota
    End If
End Sub

Private Sub txtPie_GotFocus()
    Call selecciona(txtPie)
End Sub

Private Sub txtPie_KeyDown(KeyCode As Integer, shift As Integer)
    Call Flechas(KeyCode, txtRut)
End Sub

Private Sub txtPie_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        txtPie.text = Format(txtPie.text, "$ ###,###,##0")
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtRut_GotFocus()
    Call selecciona(txtRut)
End Sub

Private Sub txtRut_KeyDown(KeyCode As Integer, shift As Integer)
    Call Flechas(KeyCode, txtRut)
End Sub

Private Sub txtRut_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        txtRut.text = ceros(txtRut)
        lblDV.Caption = rut(txtRut.text)
        If leerCliente(c, txtRut.text & lblDV.Caption, PVentas.dato7.text, "=") = True Then
            Call structtoctrl
            txtPie.text = Format(trescuotas, "$ ###,###,##0")
            Call actualizaCuota
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub actualizaCuota()
    Dim i As Integer
    Dim cuota As Double
    Dim suma As Double
    If txtPie.text = "" Or txtPie.text = "$ " Or txtPie.text = "$" Then
        txtPie.text = 0
    End If
    If txtCuotas.text = "" Or txtCuotas.text = "0" Then
        txtCuotas.text = 1
    End If
    saldo = CDbl(lblMonto.Caption) - CDbl(txtPie.text)
    Me.lblCuota.Caption = Format(Round(saldo / CDbl(txtCuotas.text) + 0.5, 0), "$ ###,###,##0")
    Me.lblSaldo.Caption = Format(CDbl(lblDisponible.Caption) - CDbl(lblMonto.Caption), "$ ###,###,##0")
    If CDbl(txtPie.text) < trescuotas Or CDbl(txtCuotas.text) > cuotascontado Then
    'interes
        suma = 0
        cuota = CDbl(lblCuota.Caption)
        For i = 1 To CDbl(txtCuotas.text)
            suma = suma + cuota + cuota * 0.03 * i
        Next i
        cuota = Round(suma / CDbl(txtCuotas.text) + 0.5, 0)
        Me.lblCuota.Caption = Format(cuota, "$ ###,###,##0")
        Me.lblSaldo.Caption = Format(CDbl(lblDisponible.Caption) - suma, "$ ###,###,##0")
    End If
    
    
   
    
    
    
End Sub


'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        Dim cad As String
        txtRut.text = c.rut
        lblDV.Caption = rut(txtRut.text)
        lblNombre.Caption = c.nombre
        lblDireccion.Caption = c.direccion
        lblDiaPago.Caption = c.diapago
        lblCupo.Caption = c.cupo
        lblUtilizado.Caption = c.cupoutilizado
        lblDisponible.Caption = c.cupo - c.cupoutilizado
        txtRut.Locked = True
        'Call DeshabilitarCajas(Me)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        'CABEZA
        cr.cabeza.loc = empresaactiva
        cr.cabeza.tipo = PVentas.dato1.text
        cr.cabeza.numero = PVentas.dato2.text
        cr.cabeza.rut = txtRut.text & lblDV.Caption
        cr.cabeza.montocompra = Format(lblMonto.Caption, "########0")
        cr.cabeza.piecompra = Format(txtPie.text, "########0")
        cr.cabeza.fecha = PVentas.dato5.text & "-" & PVentas.dato4.text & "-" & PVentas.dato3.text
        cr.cabeza.numerocuotas = txtCuotas.text
        cr.cabeza.montocuotas = Format(lblCuota.Caption, "########0")
        cr.cabeza.cajera = ""
        
        'DETALLE
        cr.Detalle.loc = cr.cabeza.loc
        cr.Detalle.tipo = cr.cabeza.tipo
        cr.Detalle.numero = cr.cabeza.numero
        cr.Detalle.rut = cr.cabeza.rut
        cr.Detalle.montocuota = Format(lblCuota.Caption, "########0")
        
        Call grabarCredito(cr, modifica, lblDiaPago.Caption)
        'Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================



