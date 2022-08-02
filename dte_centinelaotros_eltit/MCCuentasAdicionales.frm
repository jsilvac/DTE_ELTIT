VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form MCCuentasAdicionales 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1875
      Left            =   60
      TabIndex        =   13
      Top             =   1980
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   3307
      BackColor       =   12648384
      Caption         =   "Cuentas Adicionales"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
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
      Begin FlexCell.Grid adicionales 
         Height          =   1455
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   2566
         DefaultFontSize =   8.25
         Rows            =   5
         SelectionMode   =   1
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   375
         Left            =   5940
         Top             =   1500
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
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   1875
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   3307
      BackColor       =   16744576
      Caption         =   "Información Cuenta Adicional"
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
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   1500
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox dato2 
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
         Left            =   1980
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   780
         Width           =   5100
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   1560
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   285
         Left            =   6840
         TabIndex        =   4
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Porcentaje Adicionales"
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
         TabIndex        =   12
         Top             =   1140
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblPorcentaje 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   6360
         TabIndex        =   11
         Top             =   1140
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lblCupo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1980
         TabIndex        =   9
         Top             =   1140
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Porcentaje"
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
         Left            =   60
         TabIndex        =   8
         Top             =   1500
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre Adic."
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
         Left            =   60
         TabIndex        =   7
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Adic."
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
         Left            =   60
         TabIndex        =   6
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cupo Titular"
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
         Left            =   60
         TabIndex        =   5
         Top             =   1140
         Visible         =   0   'False
         Width           =   1815
      End
   End
End
Attribute VB_Name = "MCCuentasAdicionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 10) As String
    Private ca As adicionales
    Private modificar As Boolean
    
    Private Sub adicionales_DblClick()
        modificar = True
        dato1.text = adicionales.Cell(adicionales.ActiveCell.row, 1).text
        lbldv.Caption = rut(dato1.text)
        dato2.text = adicionales.Cell(adicionales.ActiveCell.row, 2).text
'        dato3.text = adicionales.Cell(adicionales.ActiveCell.Row, 3).text
        dato2.SetFocus
    End Sub

    Private Sub adicionales_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        If adicionales.ActiveCell.row >= 1 Then
            Select Case KeyCode
                Case 46
                If MsgBox("Esta Seguro Que Desea Eliminar", vbOKCancel, "Atencion") = vbOK Then
                    Call eliminarCuentaAdicional(MClientes.dato1.text & MClientes.lbldv.Caption, MClientes.dato2.text, adicionales.Cell(adicionales.ActiveCell.row, 1).text)
                    adicionales.RemoveItem (adicionales.ActiveCell.row)
                    Call CargaGrillaAdicionales(1, 3)
                    Call leerCuentasAdicionales
                 End If
            End Select
        End If
    End Sub

'==============================================================
'GOTFOCUS
'==============================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
'==============================================================
'GOTFOCUS
'==============================================================

'==============================================================
'KEYDOWN
'==============================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
'==============================================================
'KEYDOWN
'==============================================================

'==============================================================
'KEYPRESS
'==============================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            lbldv.Caption = rut(dato1.text)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato2.text <> "" Then
             Call ctrltostruct
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato3.text <> "" Then
            Call ctrltostruct
        End If
    End Sub
'==============================================================
'KEYPRESS
'==============================================================
    
    Private Sub Form_Load()
        modificar = False
        ca.rut = MClientes.dato1.text & MClientes.lbldv.Caption
        ca.sucursal = MClientes.dato2.text
        Call CargaGrillaAdicionales(1, 3)
        Call leerCuentasAdicionales
    End Sub

    Private Sub frmCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmCerrar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub

'****************************************************************************
'Formato de la Grilla adicionales
'****************************************************************************
    Private Sub CargaGrillaAdicionales(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "RUT"
        formatogrilla(1, 2) = "NOMBRE"
'        formatogrilla(1, 3) = "% CUPO"
'        formatogrilla(1, 4) = "$ CUPO"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = "25"
'        formatogrilla(2, 3) = "5"
'        formatogrilla(2, 4) = "9"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
'        formatogrilla(3, 3) = "N"
'        formatogrilla(3, 4) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000"
        formatogrilla(4, 2) = ""
'        formatogrilla(4, 3) = "#0.0"
'        formatogrilla(4, 4) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
'        formatogrilla(5, 3) = "FALSE"
'        formatogrilla(5, 4) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
'        formatogrilla(6, 3) = ""
'        formatogrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
'        formatogrilla(7, 3) = ""
'        formatogrilla(7, 4) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "8"
        formatogrilla(8, 2) = "30"
'        formatogrilla(8, 3) = "5"
'        formatogrilla(8, 4) = "10"
            
        adicionales.Cols = col
        adicionales.Rows = row
        adicionales.AllowUserResizing = False
        adicionales.DisplayFocusRect = False
        adicionales.ExtendLastCol = True
        adicionales.BoldFixedCell = False
        adicionales.DrawMode = cellOwnerDraw
        adicionales.Appearance = Flat
        adicionales.ScrollBarStyle = Flat
        adicionales.FixedRowColStyle = Flat
        'TONOS VERDES
        adicionales.BackColorFixed = RGB(90, 214, 158)
        adicionales.BackColorFixedSel = RGB(110, 230, 180)
        adicionales.BackColorBkg = RGB(90, 214, 158)
        adicionales.BackColorScrollBar = RGB(231, 247, 235)
        adicionales.BackColor1 = RGB(231, 247, 235)
        adicionales.BackColor2 = RGB(239, 255, 243)
        adicionales.GridColor = RGB(148, 231, 190)
        
        adicionales.Column(0).Width = 0
        For i = 1 To col - 1
            adicionales.Cell(0, i).text = formatogrilla(1, i)
            adicionales.Column(i).Width = Val(formatogrilla(8, i)) * (adicionales.Cell(0, i).Font.Size + 1.25)
            adicionales.Column(i).MaxLength = Val(formatogrilla(2, i))
            adicionales.Column(i).FormatString = formatogrilla(4, i)
            adicionales.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                adicionales.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                adicionales.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        adicionales.Range(0, 0, 0, adicionales.Cols - 1).Alignment = cellCenterCenter
        adicionales.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla adicionales
'****************************************************************************

    Private Sub leerCuentasAdicionales()
        Dim tabla As String
        Dim suma As Double
        suma = 0
        tabla = "SELECT CONCAT(mca.rutadicional, '" & vbTab & "', mca.nombre ) AS item "
        tabla = tabla & "FROM sv_maestroclientes AS mc INNER JOIN sv_maestroclientes_adicionales AS mca ON mc.rut = mca.rut AND mc.sucursal = mca.sucursal "
        tabla = tabla & "WHERE mc.rut = '" & ca.rut & "' AND mc.sucursal = '" & ca.sucursal & "' "
        tabla = tabla & "ORDER BY rutadicional ASC"
        Call ConectarControlData(data, servidor, baseVentas, usuario, password, tabla)
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            adicionales.Rows = 1
            adicionales.AutoRedraw = False
            While Not data.Recordset.EOF
'                suma = suma + CDbl(data.Recordset.Fields("porcentajecupo"))
                adicionales.AddItem data.Recordset.Fields("item"), True
                data.Recordset.MoveNext
            Wend
            adicionales.AutoRedraw = True
            adicionales.Refresh
        End If
'        lblPorcentaje.Caption = suma
    End Sub

'==============================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'==============================================================
    Private Sub ctrltostruct()
    Dim K As Double
    Dim existe As Boolean
    
        ca.rutadicional = dato1.text & lbldv.Caption
        ca.nombre = dato2.text
'        ca.porcentajecupo = dato3.text
        existe = False
        For K = 1 To adicionales.Rows - 1
        If adicionales.Cell(K, 1).text = ca.rutadicional And modificar = False Then
        existe = True
        If MsgBox("Cuenta Adicional Ya Existe", vbOKOnly, "Atencion") = vbOK Then
        End If
        
        End If
        
        Next K
        If existe = False Then
        Call grabarClienteAdicionales(ca, modificar)
        End If
        
        Call leerCuentasAdicionales
        Call retorno
        
    End Sub
'==============================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'==============================================================

Sub retorno()
dato1.text = ""
lbldv.Caption = ""
dato2.text = ""
'lblCupo.Caption = ""
'lblPorcentaje.Caption = ""
'dato3.text = ""
dato1.SetFocus
modificar = False



End Sub
Private Sub eliminarCuentaAdicional(ByVal rut As String, ByVal sucursal As String, ByVal rut2 As String)
        
        Dim op As Integer
        Dim CAMPOS(2, 3) As String
        
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "' AND rutadicional= '" & rut2 & "' "
        op = 4
        Set sql.conexion = ventas
        CAMPOS(0, 2) = "sv_maestroclientes_adicionales"
        sql.response = CAMPOS
        Call sql.sqlventas(op, condicion)
    End Sub
