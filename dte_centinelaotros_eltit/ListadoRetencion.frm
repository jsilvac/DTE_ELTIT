VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ListadoRetencion 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Retencion SII"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11970
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   180
      Top             =   7980
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   5955
      Left            =   60
      TabIndex        =   10
      Top             =   1920
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   10504
      BackColor       =   16744576
      Caption         =   "Informe"
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
      Begin FlexCell.Grid impresion 
         Height          =   5475
         Left            =   60
         TabIndex        =   1
         Top             =   420
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   9657
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1755
      Left            =   1740
      TabIndex        =   11
      Top             =   60
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3096
      BackColor       =   16744576
      Caption         =   "Ingreso de Información"
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
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   420
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4560
         TabIndex        =   3
         Top             =   420
         Width           =   2595
      End
      Begin XPFrame.FrameXp frmMensual 
         Height          =   975
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1720
         BackColor       =   16744576
         Caption         =   "Informe Mensual"
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
         Begin VB.ComboBox cmbMeses 
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
            Height          =   315
            Left            =   3540
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   420
            Width           =   2175
         End
         Begin VB.Label lbl1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Seleccione Mes"
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
            Height          =   315
            Left            =   1380
            TabIndex        =   14
            Top             =   420
            Width           =   1875
         End
      End
      Begin XPFrame.FrameXp frmFechas 
         Height          =   975
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1720
         BackColor       =   16744576
         Caption         =   "Informe por Rango de Fechas"
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
         Begin VB.TextBox dato5 
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
            Height          =   315
            Left            =   5700
            MaxLength       =   2
            TabIndex        =   8
            Tag             =   "proveedor"
            Top             =   420
            Width           =   435
         End
         Begin VB.TextBox dato4 
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
            Height          =   315
            Left            =   5220
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "proveedor"
            Top             =   420
            Width           =   435
         End
         Begin VB.TextBox dato6 
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
            Height          =   315
            Left            =   6180
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "proveedor"
            Top             =   420
            Width           =   795
         End
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
            Height          =   315
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   6
            Tag             =   "proveedor"
            Top             =   420
            Width           =   795
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
            Height          =   315
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   5
            Tag             =   "proveedor"
            Top             =   420
            Width           =   435
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
            Height          =   315
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   4
            Tag             =   "proveedor"
            Top             =   420
            Width           =   435
         End
         Begin VB.Label lbl2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Desde"
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
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label lbl3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hasta"
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
            Height          =   315
            Left            =   3780
            TabIndex        =   16
            Top             =   420
            Width           =   1335
         End
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   7980
      Width           =   3375
      _ExtentX        =   5953
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
Attribute VB_Name = "ListadoRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private fecha1 As String
    Private fecha2 As String

Private Sub cmbMeses_Click()
    Call cmbMeses_KeyPress(13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
    Dim ultimo As String
    'Primero = DateSerial(Year(Now), Month(Now) + 0, 1)
    If KeyAscii = 13 Then
        fecha1 = Format(fechasistema, "yyyy") & "-" & Left(cmbMeses.List(cmbMeses.ListIndex), 2) & "-01"
        ultimo = DateSerial(Year(fecha1), Month(fecha1) + 1, 0)
        fecha2 = Format(fechasistema, "yyyy") & "-" & Left(cmbMeses.List(cmbMeses.ListIndex), 2) & "-" & Format(ultimo, "dd")
        Call generaInformeLR(data, impresion, fecha1, fecha2)
        SendKeys "{Tab}"
    End If
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
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
    
    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Tipo Pago"
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato5)
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
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            SendKeys "{Tab}"
            Call generaInformeLR(data, impresion, fecha1, fecha2)
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
'    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato1.text) = dato1.MaxLength Then
'            Call dato1_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato2.text) = dato2.MaxLength Then
'            Call dato2_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato3.text) = dato3.MaxLength Then
'            Call dato3_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato4_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato4.text) = dato4.MaxLength Then
'            Call dato4_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato5_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato5.text) = dato5.MaxLength Then
'            Call dato5_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato6_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato6.text) = dato6.MaxLength Then
'            Call dato6_KeyPress(13)
'        End If
'    End Sub
    '========================================================
    'KeyUp
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, 38
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call Centrar(Me)
    Call cargaMeses
    Call CargaGrillaInforme(1, 5)
End Sub

Private Sub opt1_Click()
    If opt1.Value = True Then
        frmMensual.Visible = True
        frmFechas.Visible = False
        cmbMeses.SetFocus
    End If
End Sub

Private Sub opt1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If opt1.Value = True And frmMensual.Visible = True Then
        frmMensual.Visible = True
        frmFechas.Visible = False
        cmbMeses.SetFocus
    End If
End Sub

Private Sub opt2_Click()
    If opt2.Value = True Then
        frmMensual.Visible = False
        frmFechas.Visible = True
        dato1.SetFocus
    End If
End Sub

Private Sub opt2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If opt2.Value = True And frmFechas.Visible = True Then
        frmMensual.Visible = False
        frmFechas.Visible = True
        dato1.SetFocus
    End If
End Sub

Private Sub cargaMeses()
    Dim i As Integer
    Dim fecha As String
    Dim cad As String
    For i = 1 To 12
        cad = Format(i, "00")
        fecha = "01-" & cad & "-" & Format(fechasistema, "yyyy")
        cmbMeses.AddItem cad & " - " & Format(fecha, "mmmm")
    Next i
End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 12) As String
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "RUT" & vbCrLf & "COMPRADOR"
        formatoGrilla(1, 2) = "NOMBRE O RAZON SOCIAL"
        formatoGrilla(1, 3) = "DOMICILIO" & vbCrLf & "(Calle, Nro, Comuna)"
        formatoGrilla(1, 4) = iha & "% RETENIDO" & vbCrLf & "POR ANTICIPO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "10"
        formatoGrilla(2, 2) = "50"
        formatoGrilla(2, 3) = "50"
        formatoGrilla(2, 4) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "C"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "S"
        formatoGrilla(3, 4) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = ""
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "23"
        formatoGrilla(8, 3) = "25"
        formatoGrilla(8, 4) = "12"
            
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        
        For i = 1 To col - 1
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        
        impresion.RowHeight(0) = 36
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).FontBold = True
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
                
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        impresion.AutoRedraw = False
        impresion.PageSetup.HeaderMargin = 2
    
        impresion.PageSetup.TopMargin = 2
        impresion.PageSetup.LeftMargin = 2
        impresion.PageSetup.RightMargin = 1
        impresion.PageSetup.BottomMargin = 2
        
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.PrintFixedRow = True
        'impresion.PageSetup.Orientation = cellLandscape
        
        Call verificaImpresora(5, impresion)
        
        impresion.AutoRedraw = True
    End Sub

