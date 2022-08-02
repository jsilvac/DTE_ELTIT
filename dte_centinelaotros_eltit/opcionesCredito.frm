VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form opcionesCredito 
   BackColor       =   &H00AE1118&
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   13996
      BackColor       =   16744576
      Caption         =   "Opciones Crédito"
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
         Left            =   2280
         MaxLength       =   9
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin FlexCell.Grid opciones 
         Height          =   6855
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   12091
         Cols            =   3
         DefaultFontSize =   8.25
         Rows            =   25
         SelectionMode   =   1
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1920
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
         TabIndex        =   1
         Top             =   2880
         Width           =   4920
      End
   End
End
Attribute VB_Name = "opcionesCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatoGrilla(10, 10) As String
    Private monto As Double
    Private cuotascontado As Double
    Private disponible As Double
    
Private Sub Form_Load()
    txtPie.text = Format(Credito.txtPie.text, "########0")
    monto = CDbl(Credito.lblMonto.Caption)
    cuotascontado = Credito.cuotascontado
    disponible = CDbl(Credito.lblDisponible.Caption)
    Call cargaGrilla
    Call llenaGrilla(CDbl(txtPie.text))
End Sub

Private Sub opciones_DblClick()
    Dim saldo As Double
    If opciones.Cell(opciones.ActiveCell.row, 2).text <> "NO DISPONIBLE" Then
        Credito.txtPie.text = txtPie.text
        Credito.txtCuotas.text = opciones.Cell(opciones.ActiveCell.row, 1).text
        Credito.lblCuota.Caption = opciones.Cell(opciones.ActiveCell.row, 2).text
        Credito.lblSaldo.Caption = Format(calculaSaldo(CDbl(Credito.txtCuotas.text), CDbl(Credito.lblCuota.Caption)), "$ ###,###,##0")
        Unload Me
        Credito.txtCuotas.SetFocus
    Else
        MsgBox "No tiene cupo suficiente para tomar esta oppcion de crédito", vbOKOnly, "Error"
    End If
End Sub

Private Sub txtPie_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And txtPie.text <> "" Then
        Call llenaGrilla(CDbl(txtPie.text))
        opciones.Cell(1, 1).SetFocus
    End If
End Sub

Private Sub cargaGrilla()
    Dim i As Integer
    Dim row As Long
    Dim col As Long
        row = 1
        col = 3
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "N° CUOTAS"
        formatoGrilla(1, 2) = "MONTO CUOTA"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "2"
        formatoGrilla(2, 2) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "10"
            
        opciones.Cols = col
        opciones.Rows = row
        opciones.AllowUserResizing = False
        opciones.DisplayFocusRect = False
        opciones.ExtendLastCol = True
        opciones.BoldFixedCell = False
        opciones.DrawMode = cellOwnerDraw
        opciones.Appearance = Flat
        opciones.ScrollBarStyle = Flat
        opciones.FixedRowColStyle = Flat
        opciones.BackColorFixed = RGB(90, 158, 214)
        opciones.BackColorFixedSel = RGB(110, 180, 230)
        opciones.BackColorBkg = RGB(90, 158, 214)
        opciones.BackColorScrollBar = RGB(231, 235, 247)
        opciones.BackColor1 = RGB(231, 235, 247)
        opciones.BackColor2 = RGB(239, 243, 255)
        opciones.GridColor = RGB(148, 190, 231)
        
        opciones.Column(0).Width = 0
        For i = 1 To col - 1
            opciones.Cell(0, i).text = formatoGrilla(1, i)
            opciones.Column(i).Width = Val(formatoGrilla(8, i)) * (opciones.Cell(0, i).Font.Size + 1.25)
            opciones.Column(i).MaxLength = Val(formatoGrilla(2, i))
            opciones.Column(i).FormatString = formatoGrilla(4, i)
            opciones.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                opciones.Column(i).Alignment = cellRightCenter
            Else
                opciones.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        opciones.Range(0, 0, 0, opciones.Cols - 1).Alignment = cellCenterCenter
End Sub

Private Sub llenaGrilla(ByVal pie As Double)
    Dim i As Integer
    Dim j As Integer
    Dim cuota As String
    Dim saldo As Double
    Dim trescuotas As Double
    Dim suma As Double
    opciones.Rows = 1
    opciones.AutoRedraw = False
    
    trescuotas = CDbl(monto) / 3
    If trescuotas < 100 Then
        trescuotas = 1
    Else
        trescuotas = Round(trescuotas / 100 + 1, 0)
    End If
    
    saldo = monto - pie
    For i = 1 To 24
        
        cuota = Format(Round(saldo / CDbl(i) + 0.5, 0), "########0")
        'Me.lblSaldo.Caption = Format(CDbl(lblDisponible.Caption) - CDbl(lblMonto.Caption), "$ ###,###,##0")
        If pie < trescuotas Or CDbl(i) > cuotascontado Then
        'interes
            suma = 0
            'cuota = CDbl(lblCuota.Caption)
            For j = 1 To CDbl(i)
                suma = suma + cuota + cuota * 0.03 * j
            Next j
            cuota = Format(Round(suma / CDbl(i) + 0.5, 0), "########0")
            'Me.lblCuota.Caption = Format(cuota, "$ ###,###,##0")
            'Me.lblSaldo.Caption = Format(CDbl(lblDisponible.Caption) - suma - CDbl(txtPie.text), "$ ###,###,##0")
        End If
        If calculaSaldo(CDbl(i), CDbl(cuota)) < 0 Then
            cuota = "NO DISPONIBLE"
        End If
        opciones.AddItem i & vbTab & cuota, True
    Next i
    
    opciones.AutoRedraw = True
    opciones.Refresh
End Sub

Private Function calculaSaldo(ByVal cuotas As Double, ByVal cuota As Double) As Double
    Dim saldo As Double
    saldo = disponible
    saldo = saldo - CDbl(cuotas) * CDbl(cuota)
    calculaSaldo = saldo
End Function




