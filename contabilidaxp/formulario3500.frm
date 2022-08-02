VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form3500 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generacion Formulario 3323"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11850
      TabIndex        =   31
      Top             =   45
      Width           =   3255
      _ExtentX        =   5741
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
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp VIGENCIA 
      Height          =   8895
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   15690
      BackColor       =   16744576
      Caption         =   "PROCESO VIGENTE EN BASE DE DATOS"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HabilitarArrastre=   -1  'True
      Begin FlexCell.Grid Grid2 
         Height          =   2295
         Left            =   10200
         TabIndex        =   30
         Top             =   6480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4048
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generar archivo Para SII"
         Height          =   495
         Left            =   6840
         TabIndex        =   27
         Top             =   8280
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "PROCESAR EL FORMULARIO3500 "
         Height          =   240
         Left            =   11115
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   405
         Width           =   3660
      End
      Begin FlexCell.Grid Grid1 
         Height          =   5655
         Left            =   90
         TabIndex        =   9
         Top             =   720
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   9975
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL LINEAS SECCION 3"
         Height          =   285
         Left            =   5205
         TabIndex        =   25
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL LINEAS SECCION 2"
         Height          =   285
         Left            =   5205
         TabIndex        =   24
         Top             =   6525
         Width           =   2895
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   7
         Left            =   8175
         TabIndex        =   23
         Top             =   6840
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   6
         Left            =   8175
         TabIndex        =   22
         Top             =   6525
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   5
         Left            =   3180
         TabIndex        =   21
         Top             =   7785
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   4
         Left            =   3180
         TabIndex        =   20
         Top             =   7470
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   3
         Left            =   3180
         TabIndex        =   19
         Top             =   7155
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   2
         Left            =   3180
         TabIndex        =   18
         Top             =   6840
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   3180
         TabIndex        =   17
         Top             =   6525
         Width           =   1905
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA RETENIDO O ANTICIPO"
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   7785
         Width           =   2895
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA ANTICIPADO"
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   7470
         Width           =   2895
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA RETENIDO TOTAL"
         Height          =   285
         Left            =   210
         TabIndex        =   14
         Top             =   7155
         Width           =   2895
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA RETENIDO PARCIAL"
         Height          =   285
         Left            =   210
         TabIndex        =   13
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL REGISTROS A TRANSMITIR"
         Height          =   285
         Left            =   210
         TabIndex        =   12
         Top             =   6525
         Width           =   2895
      End
      Begin VB.Label Label2 
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
         Height          =   330
         Left            =   3735
         TabIndex        =   11
         Top             =   360
         Width           =   3570
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   3570
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   900
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   1588
      BackColor       =   16744576
      Caption         =   "PROCESO FORMULARIO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   65535
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
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   375
         Left            =   135
         TabIndex        =   2
         Top             =   360
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   315
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   556
         BackColor       =   8421504
         Caption         =   "LOCAL"
         CaptionEstilo3D =   1
         BackColor       =   8421504
         ForeColor       =   65535
         ColorBarraArriba=   12632256
         ColorBarraAbajo =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox ComboLOCAL 
            Height          =   315
            Left            =   45
            TabIndex        =   29
            Top             =   270
            Width           =   4395
         End
      End
   End
   Begin XPFrame.FrameXp FRMPROCESO 
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   688
      BackColor       =   16744576
      Caption         =   "FORMULARIO 3500"
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
      Begin VB.TextBox varipaso 
         Height          =   420
         Left            =   8865
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   630
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "SEGUNDO SEMESTRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3375
         TabIndex        =   6
         Top             =   495
         Width           =   3705
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "PRIMER SEMESTRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   45
         TabIndex        =   5
         Top             =   495
         Width           =   4830
      End
      Begin CoolButtons.cool_Button command1 
         Height          =   495
         Left            =   12690
         TabIndex        =   3
         Top             =   540
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "COMIENZA ACTUALIZACION"
      End
      Begin VB.Label actualiza 
         BackColor       =   &H00FF8080&
         Height          =   465
         Left            =   1350
         TabIndex        =   4
         Top             =   630
         Width           =   3750
      End
   End
End
Attribute VB_Name = "form3500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim debe(12) As Double
Dim haber(12) As Double
Dim totales(20) As Double
Dim año3323 As String
Dim periodo As String
Dim semestre As String

Dim FORMATOGRILLA(20, 20) As String



 

Private Sub COMMAND2_Click()
MsgBox "TIEMPO ESTIMADO 2 MINUTOS DE PROCESO,PRESIONE ENTER PARA CONTINUAR"

CARGAGRILLA
genera3500

LEERFORM3500

End Sub
Sub genera3500()
Call borrar3500("2")
Call borrar3500("3")

For k = 0 To ComboLOCAL.ListCount - 1
If Mid(ComboLOCAL.List(k), 1, 2) < "50" Then
Call LEERventascarne("00", Mid(ComboLOCAL.List(k), 1, 2))
Call LEERventasharina("00", Mid(ComboLOCAL.List(k), 1, 2))
Call LEERventascarne_retenedores("00", Mid(ComboLOCAL.List(k), 1, 2))
Call LEERventasharina_retenedores("00", Mid(ComboLOCAL.List(k), 1, 2))
Call LEERcompracarne_retenedores("00", Mid(ComboLOCAL.List(k), 1, 2))
Call LEERcompraharina_retenedores("00", Mid(ComboLOCAL.List(k), 1, 2))

End If
Next k




End Sub
Sub generaform3500()
Dim cadena As String
Dim vari1 As String
Dim AÑOPROCESO As String
Dim MESPROCESO As String
Dim CORRELATIVO As String
Dim RUTEMPRESAPROCESO As String
Dim TIPOINFORME As String
Dim DIGITOVERIFICADOR As String
Dim NOMBREEMPRESAPROCESO As String * 80
Dim monto As String
Dim RUTCLIENTE As String
Dim i As Double
Dim TIPOMO As String

Close 20
Open App.path + "\" + Mid(rutempresa, 1, 8) + ".500" For Output As #20
AÑOPROCESO = Format(fechasistema, "YYYY")
RUTEMPRESAPROCESO = Mid(rutempresa, 1, 8)
DIGITOVERIFICADOR = Mid(rutempresa, 10, 1)
TIPOINFORME = "O"
NOMBREEMPRESAPROCESO = nombreempresa

CORRELATIVO = "00"
Rem cargacantidadderegistros
varipaso.MaxLength = 10

varipaso.text = Replace(TOTAL(1).Caption + 3, ".", "")
Call ESPACIOS(varipaso)
cadena = "0" + Format(fechasistema, "yyyymm") + "3500" + RUTEMPRESAPROCESO + DIGITOVERIFICADOR + varipaso.text + "010000          0         0I00000000000000000100000000000000" + String(186, 32)
Print #20, cadena

cadena = "11" + Format(fechasistema, "yyyymm") + "3500U" + RUTEMPRESAPROCESO + DIGITOVERIFICADOR + NOMBREEMPRESAPROCESO + String(174, 32)



Print #20, cadena
For k = 1 To Grid1.Rows - 1

If Grid1.Cell(k, 1).text = "2" Then
    
    cadena = "23500" + Format(fechasistema, "yyyymm") + RUTEMPRESAPROCESO + DIGITOVERIFICADOR + "2"
    
    
    RUTCLIENTE = Mid(Grid1.Cell(k, 3).text, 2, 8)
    RUTCLIENTE = RUTCLIENTE + rut("0" + RUTCLIENTE)
    
    varipaso.MaxLength = 5
    varipaso.text = Grid1.Cell(k, 2).text
    Call ESPACIOS(varipaso)
    
    
    cadena = cadena + Format(k, "00000") + varipaso.text + RUTCLIENTE
    For i = 5 To 13
    varipaso.MaxLength = 15
    If i = 5 Then varipaso.MaxLength = 3
    If i = 6 Then varipaso.MaxLength = 12
    If i = 8 Then varipaso.MaxLength = 3

    varipaso.text = Grid1.Cell(k, i).text
    If varipaso.text = "0" Then varipaso.text = ""
    Call ESPACIOS(varipaso)
    If i = 7 Then cadena = cadena + String(15, 32)
    cadena = cadena + varipaso.text
    Next i
    cadena = Replace(cadena, ",", ".") + String(113, 32)

    Print #20, cadena
    s = Len(cadena)

    End If
    
    If Grid1.Cell(k, 1).text = "3" Then
    cadena = "23500" + Format(fechasistema, "yyyymm") + RUTEMPRESAPROCESO + DIGITOVERIFICADOR + "3" + String(157, 32)
    
    RUTCLIENTE = Mid(Grid1.Cell(k, 3).text, 2, 8)
    RUTCLIENTE = RUTCLIENTE + rut("0" + RUTCLIENTE)
    varipaso.MaxLength = 5
    varipaso.text = Grid1.Cell(k, 2).text
    Call ESPACIOS(varipaso)

    cadena = cadena + Format(k, "00000") + varipaso.text + RUTCLIENTE
    Rem TIPO DOCUMENTO
    
    varipaso.MaxLength = 3
    varipaso.text = Grid1.Cell(k, 5).text
    If varipaso.text = "0" Then varipaso.text = ""
    Call ESPACIOS(varipaso)
    cadena = cadena + varipaso.text
    
    Rem DOCUMENTO
    If Grid1.Cell(k, 16).text = "VR" Then TIPOMO = "1"
    If Grid1.Cell(k, 16).text = "CR" Then TIPOMO = "2"
    cadena = cadena + TIPOMO
    
    Rem FOLIO + UNIDES
    varipaso.MaxLength = 12
    varipaso.text = Grid1.Cell(k, 6).text
    If varipaso.text = "0" Then varipaso.text = ""
    Call ESPACIOS(varipaso)
    cadena = cadena + varipaso.text + String(15, 32)
    
    Rem CANTIDAD PRODUCTO
    varipaso.MaxLength = 15
    varipaso.text = Replace(Grid1.Cell(k, 7).text, ".", "")
    If varipaso.text = "0" Then varipaso.text = ""
    Call ESPACIOS(varipaso)
    cadena = cadena + varipaso.text
    
    Rem CODIGO UNIDAD MEDIDA
    varipaso.MaxLength = 3
    varipaso.text = Grid1.Cell(k, 8).text
    If varipaso.text = "0" Then varipaso.text = ""
    Call ESPACIOS(varipaso)
    cadena = cadena + varipaso.text
    
    Rem CODIGO NETO OPERACION
    varipaso.MaxLength = 15
    varipaso.text = Grid1.Cell(k, 9).text
    If varipaso.text = "0" Then varipaso.text = ""
    Call ESPACIOS(varipaso)
    cadena = cadena + varipaso.text
    
    Rem MONTO IVA ANTICIPADO
    varipaso.MaxLength = 15
    varipaso.text = Grid1.Cell(k, 13).text
    If varipaso.text = "0" Then varipaso.text = ""
    Call ESPACIOS(varipaso)
    cadena = cadena + varipaso.text
    cadena = Replace(cadena, ",", ".")

    
    Print #20, cadena
    s = Len(cadena)

    End If


Next k


Rem totales


cadena = "313500" + Format(fechasistema, "yyyymm") + RUTEMPRESAPROCESO + DIGITOVERIFICADOR

varipaso.MaxLength = 15
varipaso.text = Replace(TOTAL(2).Caption, ".", "")
Call ESPACIOS(varipaso)

cadena = cadena + varipaso.text

varipaso.MaxLength = 15
varipaso.text = Replace(TOTAL(3).Caption, ".", "")
Call ESPACIOS(varipaso)

cadena = cadena + varipaso.text

varipaso.MaxLength = 15
varipaso.text = Replace(TOTAL(4).Caption, ".", "")
Call ESPACIOS(varipaso)

cadena = cadena + varipaso.text

varipaso.MaxLength = 15
varipaso.text = "0"
Call ESPACIOS(varipaso)

cadena = cadena + varipaso.text

varipaso.MaxLength = 15
varipaso.text = Replace(TOTAL(5).Caption, ".", "")
Call ESPACIOS(varipaso)

cadena = cadena + varipaso.text

varipaso.MaxLength = 8

varipaso.text = Replace(TOTAL(6).Caption, ".", "")
Call ESPACIOS(varipaso)

cadena = cadena + varipaso.text

varipaso.MaxLength = 8
varipaso.text = Replace(TOTAL(7).Caption, ".", "")
Call ESPACIOS(varipaso)

cadena = cadena + varipaso.text

varipaso.MaxLength = 9
varipaso.text = Replace(rut_representante, "-", "")
Call ceros(varipaso)



cadena = cadena + varipaso.text

cadena = cadena + String(155, 32)

Print #20, cadena
Close 20


Shell "NOTEPAD " + App.path + "\" + Mid(rutempresa, 1, 8) + ".500"

End Sub

Private Sub Command3_Click()
generaform3500

End Sub

Private Sub Form_Load()
   
     
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
   Rem  Call Conectarventas(servidor, "molino_" + "ventas00", usuario, password)
    
    FRMPROCESO.Caption = "GENERA FORMULARIO 3500 AÑO:" + Format(fechasistema, "dd-mm-yyyy") + " " + nombreempresa
    
    CARGAGRILLA
   LEErlocales
   
    
    LEERFORM3500

End Sub
Sub LEERFORM3500()
Dim SUMAR As Double
Dim TIPODO As String
Dim tipo2 As Double
Dim TIPO3 As Double

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
       
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT * from form3500 where fecha like '" + Format(fechasistema, "yyyy-mm") + "%' and cantidad<>'0' " & "  order by seccion "
        csql.Execute
        If csql.RowsAffected > 0 Then
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        Set resultados = csql.OpenResultset
        For k = 1 To 20
        totales(k) = 0
        Next k
        tipo2 = 0
        TIPO3 = 0
         While Not resultados.EOF
          
          Grid1.Rows = Grid1.Rows + 1
          If resultados(0) = "2" Then
          tipo2 = tipo2 + 1
          Else
          TIPO3 = TIPO3 + 1
          End If
          
          Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
          Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
          Grid1.Cell(Grid1.Rows - 1, 3).text = Replace(resultados(2), ".", "0")
          Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
          If resultados(4) = "FV" Then TIPODO = "30"
          If resultados(4) = "NF" Then TIPODO = "60"
          If resultados(4) = "ND" Then TIPODO = "55"
          
          If resultados(4) = "FV" And resultados(13) = "E" Then
            TIPODO = "33"
          End If
          If resultados(4) = "NF" And resultados(13) = "E" Then
            TIPODO = "61"
          End If
          If resultados(4) = "ND" And resultados(13) = "E" Then
            TIPODO = "56"
          End If
          
          Grid1.Cell(Grid1.Rows - 1, 5).text = TIPODO
          Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
          Grid1.Cell(Grid1.Rows - 1, 7).text = Format(resultados(6), "###,##0.00")
          Grid1.Cell(Grid1.Rows - 1, 8).text = resultados(7)
          Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(8)
          Grid1.Cell(Grid1.Rows - 1, 10).text = resultados(9)
          Grid1.Cell(Grid1.Rows - 1, 11).text = resultados(10)
          Grid1.Cell(Grid1.Rows - 1, 12).text = resultados(11)
          Grid1.Cell(Grid1.Rows - 1, 13).text = resultados(12)
          Grid1.Cell(Grid1.Rows - 1, 14).text = resultados(13)
          Grid1.Cell(Grid1.Rows - 1, 15).text = Format(leerimpuestoFACTURA(resultados(1), resultados(4), resultados(5), resultados(13), empresaactiva), "###,###,##0")
          Grid1.Cell(Grid1.Rows - 1, 16).text = resultados(14)
          If resultados(14) <> "CR" Then
          Grid1.Cell(Grid1.Rows - 1, 17).text = leerNombrerut(cuentacliente, resultados(2))
          Else
          Grid1.Cell(Grid1.Rows - 1, 17).text = leerNombrerut(CUENTAPROVEEDOR, resultados(2))
          
          End If
          
          If resultados(14) = "VN" Then
          If resultados(4) = "NF" Then
          totales(1) = totales(1) - resultados(10)
          totales(2) = totales(2) - resultados(11)
          totales(3) = totales(3) - resultados(12)
          
          Else
          totales(1) = totales(1) + resultados(10)
          totales(2) = totales(2) + resultados(11)
          totales(3) = totales(3) + resultados(12)
          
          End If
          End If
          If (CDbl(Grid1.Cell(Grid1.Rows - 1, 15).text) - resultados(12) > 1 Or CDbl(Grid1.Cell(Grid1.Rows - 1, 15).text) - resultados(12) < -1) And resultados(14) = "VN" Then
          Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 16).ForeColor = &HFF&
          End If
          
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
       totales(4) = totales(1) + totales(2) + totales(3)
    TOTAL(1).Caption = Format(Grid1.Rows - 1, "###,###,###,##0")
    
    TOTAL(2).Caption = Format(totales(1), "###,###,###,##0")
    TOTAL(3).Caption = Format(totales(2), "###,###,###,##0")
    TOTAL(4).Caption = Format(totales(3), "###,###,###,##0")
    TOTAL(5).Caption = Format(totales(4), "###,###,###,##0")
    TOTAL(6).Caption = Format(tipo2, "###,###,###,##0")
    TOTAL(7).Caption = Format(TIPO3, "###,###,###,##0")
    
    
    Grid1.AutoRedraw = True
    Grid1.Refresh
LEERFORM3500_tipo

End Sub

Sub LEERFORM3500_tipo()
Dim SUMAR As Double
Dim TIPODO As String
Dim tipo2 As Double
Dim TIPO3 As Double

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
       
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT codigo,tipo,sum(iva_anticipado) from form3500 where fecha like '" + Format(fechasistema, "yyyy-mm") + "%'" & " and documento<>'c' group by codigo,tipo "
        csql.Execute
        If csql.RowsAffected > 0 Then
        Grid2.Cols = 4
        Grid2.Rows = 1
        Grid2.AutoRedraw = False
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
          
          Grid2.Rows = Grid2.Rows + 1
          Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
          Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
          Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
       
    Grid2.AutoRedraw = True
    Grid2.Refresh
        
End Sub


Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7
    FORMATOGRILLA(1, 1) = "SECCION"
    FORMATOGRILLA(1, 2) = "CODIGO"
    FORMATOGRILLA(1, 3) = "RUT"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "TIPO"
    FORMATOGRILLA(1, 6) = "FOLIO"
    FORMATOGRILLA(1, 7) = "CANTIDAD"
    FORMATOGRILLA(1, 8) = "MEDIDA"
    FORMATOGRILLA(1, 9) = "NETO"
    FORMATOGRILLA(1, 10) = "BASE"
    FORMATOGRILLA(1, 11) = "IVA RET PARC"
    FORMATOGRILLA(1, 12) = "IVA RET TOTAL"
    FORMATOGRILLA(1, 13) = "IVA ANTICIPADO"
    FORMATOGRILLA(1, 14) = "DOCUMENTO"
    FORMATOGRILLA(1, 15) = "CONTA"
    FORMATOGRILLA(1, 16) = "TIPO_MOV"
    FORMATOGRILLA(1, 17) = "NOMBRE"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"
    FORMATOGRILLA(2, 10) = "10"
    FORMATOGRILLA(2, 11) = "10"
    FORMATOGRILLA(2, 12) = "10"
    FORMATOGRILLA(2, 13) = "10"
    FORMATOGRILLA(2, 14) = "10"
    FORMATOGRILLA(2, 15) = "10"
    FORMATOGRILLA(2, 16) = "10"
    FORMATOGRILLA(2, 17) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "D"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "N"
    FORMATOGRILLA(3, 16) = "N"
    FORMATOGRILLA(3, 17) = "S"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    
    For k = 7 To 16
    
    FORMATOGRILLA(4, k) = "###,###,###"
    Next k
    FORMATOGRILLA(4, 7) = "#####0.00"
    
    
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    FORMATOGRILLA(5, 16) = "TRUE"
    FORMATOGRILLA(5, 17) = "TRUE"
    
    Grid1.Cols = 18
    Grid1.Rows = 1
    
     'infogrilla.grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
       
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Private Function LEERNOMBREPROVEEDOR(rut) As String



    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "rut=" + "'" + rut + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LEERNOMBREPROVEEDOR = sqlconta.response(0, 3)
    End If
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    

End Function

Sub LEERventascarne(rubro, loc)
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    fecha1 = Format(fechasistema, "yyyy-mm") + "-01"
    fecha2 = Format(fechasistema, "yyyy-mm-dd")
        Set csql.ActiveConnection = contadb
        csql.sql = "insert ignore into form3500 "
        csql.sql = csql.sql + "select '2',ml.codigosii,dc.rut,dc.fecha,dc.tipo,dc.foliosii,sum(dd.cantidad),'1',round(sum(dd.total)/1.24,0) as neto,'0','0','0',round(sum((dd.total/1.24)*porcentajeimpuesto),0) as impuesto,dc.contabilizado,'VN' "
        csql.sql = csql.sql + "from " + clientesistema + "ventas" + loc + ".sv_documento_detalle_" + loc + " as dd "
        csql.sql = csql.sql + "inner join " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " as dc on dd.tipo=dc.tipo and dd.numero=dc.numero and dd.fecha=dc.fecha and dd.caja=dc.caja and dc.nula<>'S' "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestroproductos_fijo_" + rubro + " as mpf on dd.codigo=mpf.codigobarra "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestrolineas_" + rubro + " as ml on mpf.codigoseccion=ml.codigoseccion and mpf.codigodepto=ml.codigodepto and mpf.codigolinea=ml.codigolinea "
        csql.sql = csql.sql + " where (dc.tipo='FV' or dc.tipo='NF' OR dc.tipo='ND') "
        csql.sql = csql.sql + "and dd.porcentajeimpuesto='0.05' and dc.fecha like '" + Format(fechasistema, "yyyy-mm") + "%' and dc.caja<'90' group by dc.tipo,dc.foliosii,ml.codigosii order by dc.rut,dc.tipo,dc.foliosii;"
        
        
        csql.Execute
     
End Sub
Sub LEERventasharina(rubro, loc)
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    fecha1 = Format(fechasistema, "yyyy-mm") + "-01"
    fecha2 = Format(fechasistema, "yyyy-mm-dd")
        Set csql.ActiveConnection = contadb
        csql.sql = "insert ignore into form3500 "
        csql.sql = csql.sql + "select '2','1900',dc.rut,dc.fecha,dc.tipo,dc.foliosii,sum(dd.cantidad),'1',round(sum(dd.total)/1.31,0) as neto,'0','0','0',round(sum((dd.total/1.31)*porcentajeimpuesto),0) as impuesto,dc.contabilizado,'VN' "
        csql.sql = csql.sql + "from " + clientesistema + "ventas" + loc + ".sv_documento_detalle_" + loc + " as dd "
        csql.sql = csql.sql + "inner join " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " as dc on dd.tipo=dc.tipo and dd.numero=dc.numero and dd.fecha=dc.fecha and dd.caja=dc.caja and dc.nula<>'S' "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestroproductos_fijo_" + rubro + " as mpf on dd.codigo=mpf.codigobarra "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestrolineas_" + rubro + " as ml on mpf.codigoseccion=ml.codigoseccion and mpf.codigodepto=ml.codigodepto and mpf.codigolinea=ml.codigolinea "
        csql.sql = csql.sql + " where (dc.tipo='FV' or dc.tipo='NF' OR dc.tipo='ND') "
        csql.sql = csql.sql + "and dd.porcentajeimpuesto='0.12' and dc.fecha like '" + Format(fechasistema, "yyyy-mm") + "%' and dc.caja<='99' group by dc.tipo,dc.foliosii,ml.codigosii order by dc.rut,dc.tipo,dc.foliosii;"
        
        
        csql.Execute
     
End Sub
Sub LEERventascarne_retenedores(rubro, loc)
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    fecha1 = Format(fechasistema, "yyyy-mm") + "-01"
    fecha2 = Format(fechasistema, "yyyy-mm-dd")
        Set csql.ActiveConnection = contadb
        csql.sql = "insert ignore into form3500 "
        csql.sql = csql.sql + "select '3',ml.codigosii,dc.rut,dc.fecha,dc.tipo,dc.foliosii,sum(dd.cantidad),'1',round(sum(dd.total)/1.19,0) as neto,'0','0','0',round(sum((dd.total/1.19)*porcentajeimpuesto),0) as impuesto,dc.contabilizado,'VR' "
        csql.sql = csql.sql + "from " + clientesistema + "ventas" + loc + ".sv_documento_detalle_" + loc + " as dd "
        csql.sql = csql.sql + "inner join " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " as dc on dd.tipo=dc.tipo and dd.numero=dc.numero and dd.fecha=dc.fecha and dd.caja=dc.caja and dc.nula<>'S' "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestroproductos_fijo_" + rubro + " as mpf on dd.codigo=mpf.codigobarra "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestrolineas_" + rubro + " as ml on mpf.codigoseccion=ml.codigoseccion and mpf.codigodepto=ml.codigodepto and mpf.codigolinea=ml.codigolinea "
        csql.sql = csql.sql + " where (dc.tipo='FV' or dc.tipo='NF' OR dc.tipo='ND') "
        csql.sql = csql.sql + "and mpf.codigoimpuesto='00005' and dd.porcentajeimpuesto='0' and dc.fecha like '" + Format(fechasistema, "yyyy-mm") + "%' and dc.caja>'80' group by dc.tipo,dc.foliosii,ml.codigosii order by dc.rut,dc.tipo,dc.foliosii "
        
        
        csql.Execute
     
End Sub
Sub LEERventasharina_retenedores(rubro, loc)
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    fecha1 = Format(fechasistema, "yyyy-mm") + "-01"
    fecha2 = Format(fechasistema, "yyyy-mm-dd")
        Set csql.ActiveConnection = contadb
        csql.sql = "insert ignore into form3500 "
        csql.sql = csql.sql + "select '3',ml.codigosii,dc.rut,dc.fecha,dc.tipo,dc.foliosii,sum(dd.cantidad),'1',round(sum(dd.total)/1.19,0) as neto,'0','0','0',round(sum((dd.total/1.19)*porcentajeimpuesto),0) as impuesto,dc.contabilizado,'VR' "
        csql.sql = csql.sql + "from " + clientesistema + "ventas" + loc + ".sv_documento_detalle_" + loc + " as dd "
        csql.sql = csql.sql + "inner join " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " as dc on dd.tipo=dc.tipo and dd.numero=dc.numero and dd.fecha=dc.fecha and dd.caja=dc.caja and dc.nula<>'S' "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestroproductos_fijo_" + rubro + " as mpf on dd.codigo=mpf.codigobarra "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestrolineas_" + rubro + " as ml on mpf.codigoseccion=ml.codigoseccion and mpf.codigodepto=ml.codigodepto and mpf.codigolinea=ml.codigolinea "
        csql.sql = csql.sql + " where (dc.tipo='FV' or dc.tipo='NF' OR dc.tipo='ND') "
        csql.sql = csql.sql + "and mpf.codigoimpuesto='00004' and dd.porcentajeimpuesto='0' and dc.fecha like '" + Format(fechasistema, "yyyy-mm") + "%' and dc.caja>'80' group by dc.tipo,dc.foliosii,ml.codigosii order by dc.rut,dc.tipo,dc.foliosii;"
        
        
        csql.Execute
     
End Sub
Sub LEERcompracarne_retenedores(rubro, loc)
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    fecha1 = Format(fechasistema, "yyyy-mm") + "-01"
    fecha2 = Format(fechasistema, "yyyy-mm-dd")
        Set csql.ActiveConnection = contadb
        csql.sql = "insert ignore into form3500 "

 
        
        csql.sql = csql.sql + "select '3',ml.codigosii,dd.rut,df.fecha,df.tipo,df.numero,sum(dd.cantidad),'1',round(sum(dd.total)/1.19,0) as neto,'0','0','0',round(sum((dd.total/1.19)*'0'),0) as impuesto,'C','CR' "
        csql.sql = csql.sql + "from " + clientesistema + "gestion" + rubro + ".l_movimientos_detalle_" + loc + " as dd "
        csql.sql = csql.sql + "left join " + clientesistema + "gestion" + rubro + ".l_ordendecompra_detalle_facturas_" + loc + " as df on df.ordendecompra = dd.numero "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestroproductos_fijo_" + rubro + " as mpf on dd.codigo=mpf.codigobarra "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestrolineas_" + rubro + " as ml on mpf.codigoseccion=ml.codigoseccion and mpf.codigodepto=ml.codigodepto and mpf.codigolinea=ml.codigolinea "
        csql.sql = csql.sql + " where (dd.tipo='OC') "
        csql.sql = csql.sql + "and mpf.codigoimpuesto='00005' and df.fecha like '" + Format(fechasistema, "yyyy-mm") + "%' group by dd.tipo,dd.numero,ml.codigosii order by dd.rut,dd.tipo,dd.numero "
        
        
        csql.Execute
     
End Sub
Sub LEERcompraharina_retenedores(rubro, loc)
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    fecha1 = Format(fechasistema, "yyyy-mm") + "-01"
    fecha2 = Format(fechasistema, "yyyy-mm-dd")
        Set csql.ActiveConnection = contadb
        csql.sql = "insert ignore into form3500 "

        csql.sql = csql.sql + "select '3','1900',dd.rut,df.fecha,df.tipo,df.numero,sum(dd.cantidad),'1',round(sum(dd.total)/1.19,0) as neto,'0','0','0',round(sum((dd.total/1.19)*'0.12'),0) as impuesto,'C','CR' "
        csql.sql = csql.sql + "from " + clientesistema + "gestion" + rubro + ".l_movimientos_detalle_" + loc + " as dd "
        csql.sql = csql.sql + "left join " + clientesistema + "gestion" + rubro + ".l_ordendecompra_detalle_facturas_" + loc + " as df on df.ordendecompra = dd.numero "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestroproductos_fijo_" + rubro + " as mpf on dd.codigo=mpf.codigobarra "
        csql.sql = csql.sql + "inner join " + clientesistema + "gestion" + rubro + ".r_maestrolineas_" + rubro + " as ml on mpf.codigoseccion=ml.codigoseccion and mpf.codigodepto=ml.codigodepto and mpf.codigolinea=ml.codigolinea "
        csql.sql = csql.sql + " where (dd.tipo='OC') "
        csql.sql = csql.sql + "and mpf.codigoimpuesto='00004' and df.fecha like '" + Format(fechasistema, "yyyy-mm") + "%' group by dd.tipo,dd.numero,ml.codigosii order by dd.rut,dd.tipo,dd.numero "
        
        
        csql.Execute
     
End Sub

Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion.g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub
Sub borrar3500(seccion)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "delete FROM " + clientesistema + "conta" + empresaactiva + ".form3500 WHERE seccion='" + seccion + "' and fecha like '" + Format(fechasistema, "yyyy-mm") + "%' "
        csql.Execute
        Call sincronizadatos(csql.sql, contadb, "")
        
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
