VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form activos_maestro_ubicaciones 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MAESTRO UBICACIONES"
   ClientHeight    =   7890
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   8025
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrmOpciones 
      Height          =   1140
      Left            =   75
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2011
      BackColor       =   16761024
      Caption         =   "OPCIONES"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   192
      ColorBarraArriba=   255
      ColorBarraAbajo =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   192
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Modificar"
         PicturePosition =   0
         Picture         =   "activos_maestro_ubicaciones.frx":0000
         PictureHover    =   "activos_maestro_ubicaciones.frx":0CB6
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16761024
      End
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   1
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Eliminar"
         PicturePosition =   0
         Picture         =   "activos_maestro_ubicaciones.frx":1A17
         PictureHover    =   "activos_maestro_ubicaciones.frx":2734
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16761024
      End
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   4
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Retorno"
         PicturePosition =   0
         Picture         =   "activos_maestro_ubicaciones.frx":34CD
         PictureHover    =   "activos_maestro_ubicaciones.frx":41F6
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16761024
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6735
      Left            =   75
      TabIndex        =   6
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11880
      BackColor       =   16744576
      Caption         =   "MAESTRO DE UBICACIONES DE ACTIVOS "
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   5295
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9340
         BackColor       =   16744576
         Caption         =   "UBICACIONES CREADAS"
         CaptionEstilo3D =   2
         BackColor       =   16744576
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         ColorBarraArriba=   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin FlexCell.Grid Grid1 
            Height          =   4905
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   8652
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
            SelectionMode   =   1
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   1931
         BackColor       =   16744576
         Caption         =   ""
         CaptionEstilo3D =   2
         BackColor       =   16744576
         ForeColor       =   8438015
         BordeColor      =   8388608
         ColorBarraArriba=   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox dato1 
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
            MaxLength       =   5
            TabIndex        =   0
            Tag             =   "fechauf"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox dato2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            MaxLength       =   80
            TabIndex        =   1
            Tag             =   "monto"
            Top             =   720
            Width           =   5895
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DESCRIPCION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CODIGO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1530
         End
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   7995
      TabIndex        =   5
      Top             =   7890
      Width           =   8025
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3735
      Left            =   8400
      TabIndex        =   2
      Top             =   240
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3495
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   12582912
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   16107953
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   16776436
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3735
         Left            =   0
         Top             =   0
         Width           =   4695
      End
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   1215
      Left            =   6360
      TabIndex        =   7
      Top             =   6720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2143
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   3735
      Left            =   8520
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "activos_maestro_ubicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private MODIFI As Integer


Private Sub dato1_GotFocus()
Call cargatexto(dato1)
End Sub
Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub
'Private Sub dato3_GotFocus()
'Call cargatexto(dato3)
'End Sub
'
'Private Sub dato4_GotFocus()
'If MODIFI = 0 Then Call leer
'Call cargatexto(dato4)
'End Sub
  
Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato2, KeyCode)
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
       Call flechas(dato1, dato2, KeyCode)
End Sub
'
' Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
'       Call flechas(dato3, dato4, KeyCode)
'End Sub
'Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
'       Call flechas(dato3, dato4, KeyCode)
'End Sub
'
'Private Sub dato3_KeyPress(KeyAscii As Integer)
'    snum = 0: KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 Then
'        Call ceros(dato3)
'        If dato3.text = "0000" Then dato3.text = Format(fechasistema, "yyyy")
'        If IsDate(dato3.text & "/" & dato2.text & "/" & dato1.text) = True Then
'            Call Pregunta(dato3, dato4)
'        Else
'            MsgBox "La fecha no es Valida", vbCritical, "Atención"
'        End If
'    End If
'End Sub
'
'Private Sub dato4_KeyPress(KeyAscii As Integer)
'    snum = 0: KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 And dato4.text <> "" Then
'        grabar
'        retorno
'    End If
'
'End Sub

Private Sub Grid1_DblClick()
Dim row As Double
row = Grid1.ActiveCell.row
dato1.text = Grid1.Cell(row, 1).text
Call dato1_KeyPress(13)
End Sub

 Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)
    Call Conectar_BD
    sc = 0
    FrmOpciones.Visible = False
 
Call CARGAPERMISO(Me.Name)
Call CARGAGRILLA

Call GenerarListado
dato1.text = LeerUltimo

End Sub
Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 2) = ""
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 3
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato1)
        
        If leer(dato1) = False Then Call Pregunta(dato1, dato2)
        
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        
        Call grabar
        Call GenerarListado
        Call retorno
    End If
   
End Sub

Function leer(codigo) As Boolean
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "af_maestro_ubicaciones"
    condicion = "codigo = '" & dato1 & "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then leer = False: Exit Function
    carga
    FrmOpciones.Visible = True
    disponible (True)
    habilita (True)
    MANUAL.SetFocus
    leer = True
        
no:
End Function
 
Function LeerUltimo() As String
    campos(0, 0) = "lpad(ifnull(max(codigo),0)+1,5,0)"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "af_maestro_ubicaciones"
    condicion = "codigo<>''"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LeerUltimo = sqlconta.response(0, 3)
    End If
no:
   
    
End Function

Sub carga()
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
 
  
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    dato1.Locked = condicion
    dato2.Locked = condicion
  '  dato3.Locked = condicion
 '   dato4.Locked = condicion
  
 
End Sub
Sub disponible(ByVal condicion As Boolean)
    dato1.Enabled = condicion
    dato2.Enabled = condicion
   ' dato3.Enabled = condicion
   ' dato4.Enabled = condicion
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudamonedasarriendo(ByRef caja As TextBox)
'    Dim CAMPOS As Variant
'    Dim cfijo As Variant
'    Dim largo As Variant
'    CAMPOS = Array("codigomoneda", "nombremoneda")
'    largo = Array("11s", "40s")
'    cfijo = "codigomoneda like '%%'"
'    cabezas = Array("Codigo", "Nombre")
'    mensajeAyuda = "Ayuda de Tipos de Monedas"
'
'    Call cargaAyudaT(servidor, clientesistema & "arriendos", usuario, password, ".maestro_monedas", caja, CAMPOS, cfijo, largo, 2)
'    caja.Enabled = True
'    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
   
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
  
    campos(0, 2) = "af_maestro_ubicaciones"
    If MODIFI = 1 Then condicion = "codigo = '" & dato1.text & "' "

    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
Sub ELIMINAR()
    campos(0, 2) = "af_maestro_ubicaciones"
    condicion = "codigo = '" & dato1.text & "' "

    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)

    
End Sub
  

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then modifica
If command = "elimina" Then

    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        If MsgBox("SE VA A ELIMNINAR ESTA UBICACION" & vbNewLine & " DESEA CONTINUAR?", vbYesNo, "ATENCION") = vbYes Then
            ELIMINA
        End If
    End If
End If
'If command = "siguiente" Then leersiguiente
'If command = "anterior" Then leeranterior

End Sub
Sub ELIMINA()
 
disponible (True)
habilita (False)
ELIMINAR
Call retorno
 
End Sub

Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
'dato2.Enabled = False
'dato3.Enabled = False
'dato4.SetFocus
MODIFI = 1

End Sub
Sub retorno()

disponible (True)
habilita (False)
limpia
FrmOpciones.Visible = False
dato1.text = LeerUltimo
dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
Call GenerarListado

 
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
  '  dato3.text = ""
  '  dato4.text = ""
  
End Sub
 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub
 

Sub GenerarListado()
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = conta

csql.sql = "select codigo,nombre from af_maestro_ubicaciones "
csql.Execute
Grid1.Rows = 1
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
    Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
    
    resultados.MoveNext
    Wend
End If

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub opcion_Click(Index As Integer)
Select Case Index
    Case 0
        Call opciones_FSCommand("modifica", "")
    Case 1
        Call opciones_FSCommand("elimina", "")
    Case 4
        Call opciones_FSCommand("retorno", "")
End Select
End Sub
