VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form publi0004 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informa Resumen de Compras x Proveedor"
   ClientHeight    =   8160
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10785
   Begin MSComctlLib.ProgressBar barra 
      Height          =   240
      Left            =   765
      TabIndex        =   7
      Top             =   7650
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   9120
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin XPFrame.FrameXp FrameXp8 
      Height          =   1695
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2990
      BackColor       =   14737632
      Caption         =   "Rangos de Fecha"
      CaptionEstilo3D =   1
      BackColor       =   14737632
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
      Alignment       =   1
      Begin CoolButtons.cool_Button cool_Button3 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         SkinId          =   "13"
         Caption         =   "Cambia Fecha"
      End
      Begin VB.Label hastafecha 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label desdefecha 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin XPFrame.FrameXp opcion2 
      Height          =   3375
      Left            =   900
      TabIndex        =   8
      Top             =   2070
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   5953
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
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
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   1815
         Left            =   270
         TabIndex        =   9
         Top             =   405
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3201
         BackColor       =   16761024
         Caption         =   "INGRESO DE RUT"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.TextBox ctdato1 
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
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   16
            Tag             =   "tipo"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox dv 
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
            Left            =   2700
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "tipo"
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox ctdato2 
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
            Left            =   1575
            MaxLength       =   9
            TabIndex        =   10
            Tag             =   "rut"
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   225
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label nombrectacte 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2700
            TabIndex        =   17
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   270
            TabIndex        =   14
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rut"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   270
            TabIndex        =   13
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label ctnombre 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1575
            TabIndex        =   12
            Top             =   1080
            Width           =   5865
         End
      End
      Begin CoolButtons.cool_Button command2 
         Height          =   495
         Left            =   3105
         TabIndex        =   15
         Top             =   2700
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "GENERA INFORME"
      End
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   7680
      TabIndex        =   19
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "publi0004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private saldo As Double
Private dedonde As Integer
Private tipoctacte As String





Private Sub busca_Click()

End Sub

Private Sub cmde01_Change()

End Sub

Private Sub cmde01_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub



Private Sub Command10_Click()
fechas.Visible = False



End Sub


Private Sub COMMAND2_Click()
lin = 0
dedonde = 2
Call ACEPTA(dedonde)
grillainformes.Tag = "publi0004"
End Sub

Private Function leercuenta(cuenta) As String

      campos(0, 0) = "nombre"
      campos(1, 0) = ""
      campos(0, 2) = "cuentasdelmayor"
      condicion = "codigo=" + "'" + cuenta + "' and año='" + Format(fechasistema, "yyyy") + "'"
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        leercuenta = sqlconta.response(0, 3)
        
        Else
        leercuenta = ""
        
        End If
        

End Function
Private Function leerctacte(tipo, rut) As String

      campos(0, 0) = "nombre"
      campos(1, 0) = ""
      campos(0, 2) = "cuentascorrientes"
      condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año='" + Format(fechasistema, "yyyy") + "'"
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        leerctacte = sqlconta.response(0, 3)
        
        Else
        leerctacte = ""
        
        End If
        

End Function

Sub ACEPTA(opcion)
Dim fecha1 As String
Dim fecha2 As String
Dim infogrilla As grillainformes
Set infogrilla = New grillainformes
Call CARGAGRILLA(infogrilla)
infogrilla.Caption = "COMPRAS A PROVEEDOR"

fecha1 = Format(desdefecha.Caption, "dd-mm-yyyy")
fecha2 = Format(hastafecha.Caption, "dd-mm-yyyy")
Call LEERMOVIMIENTOS(infogrilla, ctdato2.text + DV.text)


infogrilla.cabeza.Caption = "COMPRAS DESDE " & fecha1 & " al " & fecha2 & " "
infogrilla.Grid1.Visible = True
infogrilla.Show
End Sub

Private Sub Command6_Click()
lin = 0
dedonde = 3
Call ACEPTA(dedonde)
End Sub


Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub cool_Button3_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub ctdato1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If leercuenta(ctdato1.text) <> "" Then
nombrectacte.Caption = leercuenta(ctdato1.text)


Else
MsgBox ("codigo de cuenta no existe")
ctdato1.text = ""
ctdato1.SetFocus

End If
End If

End Sub

Private Sub ctdato2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudactacte(ctdato2)
End Sub

Private Sub ctdato2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(ctdato2)
DV.text = rut(ctdato2.text)
If leerctacte(ctdato1.text, ctdato2.text + DV.text) <> "" Then
ctnombre.Caption = leerctacte(ctdato1.text, ctdato2.text + DV.text)
Command2.SetFocus

Else

MsgBox ("rut no existe")
ctdato2.text = ""
DV.text = ""

ctdato2.SetFocus

End If


End If


End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
desdefecha.Caption = "01-" + Format(fechasistema, "mm-yyyy")
hastafecha.Caption = fechasistema

lin = 0
ctdato1.text = CUENTAPROVEEDOR
Call ctdato1_KeyPress(13)
DOCU(1) = "FA "
DOCU(2) = "ND "
DOCU(3) = "NC "
DOCU(4) = "FAE "
DOCU(5) = "NDE "
DOCU(6) = "NCE "
DOCU(7) = "FP"



End Sub


    
Sub LEERMOVIMIENTOS(infogrilla As grillainformes, rut)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
  
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim monto As Double
    Dim total As Double
    
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
        fecha2 = Mid(hastafecha.Caption, 7, 4) + "-" + Mid(hastafecha.Caption, 4, 2) + "-" + Mid(hastafecha.Caption, 1, 2)
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,tipo,numero,rut,neto "
        csql.sql = csql.sql + "FROM facturasdecompras where rut='" + rut + " ' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        csql.sql = csql.sql + "order by fecha,tipo,numero "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
        
        total = 0
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 1).text = resultados(0)
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 2).text = DOCU(resultados(1))
            
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 3).text = resultados(2)
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 4).text = Mid(resultados(3), 1, 9) + "-" + Mid(resultados(3), 10, 1)
            
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 5).text = ctnombre.Caption
            
            monto = resultados(4)
            If resultados(1) = "3" Or resultados(1) = "6" Then
            monto = monto * -1
            End If
            
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 6).text = monto
            total = total + monto
             resultados.MoveNext
          
         Wend
          resultados.Close
            Set resultados = Nothing
            
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            infogrilla.Grid1.Range(infogrilla.Grid1.Rows - 1, 1, infogrilla.Grid1.Rows - 1, infogrilla.Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
            infogrilla.Grid1.Range(infogrilla.Grid1.Rows - 1, 1, infogrilla.Grid1.Rows - 1, infogrilla.Grid1.Cols - 1).FontBold = True
            
            
            
            
            
            
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 5).text = "TOTAL COMPRAS PERIODO"
            infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 6).text = total
            
        End If
 
End Sub

Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "RUT"
    FORMATOGRILLA(1, 5) = "PROVEEDOR"
    FORMATOGRILLA(1, 6) = "NETO"
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "5"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "40"
    FORMATOGRILLA(2, 6) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,###,##0"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    
    infogrilla.Grid1.Cols = 7
    infogrilla.Grid1.Rows = 2
    
    'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    infogrilla.Grid1.DefaultFont.Size = 8
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub


   
Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & ctdato1.text & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", ctdato2, campos, cfijo, largo, 2)

    If Val(caja.text) = 0 Then ctdato2.SetFocus: GoTo no
   
    
    DV.text = rut(caja)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
