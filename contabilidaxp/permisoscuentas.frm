VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form PermisosCuentas 
   BorderStyle     =   0  'None
   ClientHeight    =   7785
   ClientLeft      =   600
   ClientTop       =   795
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp CABEZA 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   14631
      BackColor       =   16761024
      Caption         =   "Maestro Cuentas del Mayor"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   65535
      ColorBarraArriba=   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   16711680
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   855
         Left            =   0
         TabIndex        =   2
         Top             =   6960
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   1508
         BackColor       =   16777152
         Caption         =   "OPCIONES"
         CaptionEstilo3D =   1
         BackColor       =   16777152
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
         Begin VB.CheckBox atodo 
            BackColor       =   &H00FFFFC0&
            Caption         =   "ACCESO A TODAS LAS CUENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   10
            Top             =   360
            Width           =   3255
         End
         Begin CoolButtons.cool_Button cmd_xml 
            Height          =   495
            Left            =   12360
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "Genera XML SII"
         End
         Begin CoolButtons.cool_Button CmdSalir 
            Height          =   375
            Left            =   7080
            TabIndex        =   9
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Caption         =   "Salir"
         End
         Begin VB.Label Usuario 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label registros 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   10800
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label LETRA 
            Height          =   255
            Left            =   6600
            TabIndex        =   3
            Top             =   -2160
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6615
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   11668
         Cols            =   5
         DefaultFontSize =   8.25
         DefaultRowHeight=   15
         Rows            =   30
      End
      Begin VB.TextBox PIVOTE 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label titulofinal 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "PermisosCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub COMMAND2_Click()
Grid1.ExportToExcel (""), True
End Sub
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub atodo_Click()
Dim n As Integer
Dim Permiso As String
Dim cuenta As String
Grid1.AutoRedraw = False
For n = 1 To Grid1.Rows - 1
Grid1.Cell(n, 3).text = atodo.Value

Call GrabarPermiso(Grid1.Cell(n, 1).text, atodo.Value)
Next n
Grid1.AutoRedraw = True
Grid1.Refresh
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
Call CENTRAR(Me)

Call FORMATOGRILLA
'Call CargarCuentas


End Sub


Private Sub CargarCuentas()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT "
        csql2.sql = csql2.sql + " codigo,nombre,ifnull(permiso,0) as permiso "
        csql2.sql = csql2.sql + " FROM " & clientesistema & "conta" & empresaactiva & ".cuentasdelmayor LEFT join " & clientesistema & "conta.permisos_cuentas on "
        csql2.sql = csql2.sql & " usuario='" & Usuario.Caption & "' and cuenta=codigo where año='" & Format(fechasistema, "yyyy") & "' and  right(codigo,4)<>'0000'  order by codigo"
        csql2.Execute
        Grid1.Rows = csql2.RowsAffected + 1
        
        If csql2.RowsAffected > 0 Then
            Set resultados2 = csql2.OpenResultset
            LINEAS = 0
            While Not resultados2.EOF
            LINEAS = LINEAS + 1
            Grid1.Cell(LINEAS, 1).text = resultados2(0)
            Grid1.Cell(LINEAS, 2).text = resultados2(1)
            Grid1.Cell(LINEAS, 3).text = resultados2(2)
            resultados2.MoveNext
            Wend
              
              resultados2.Close
                Set resultados2 = Nothing

        End If
 End Sub

Private Sub FORMATOGRILLA()
With Grid1
.Cols = 4
.Column(0).Width = 0
.Column(1).Width = 80
.Column(2).Width = 300
.Column(3).Width = 100
.Cell(0, 0).text = ""
.Cell(0, 1).text = "CUENTA"
.Cell(0, 3).text = "PERMISO"
.Range(0, 1, 0, 2).Merge

.Column(3).CellType = cellCheckBox
.Column(3).Locked = False
End With

End Sub

Private Sub Grid1_Click()
Dim Permiso As String
Dim cuenta As String
cuenta = Grid1.Cell(Grid1.ActiveCell.row, 1).text
Permiso = Grid1.Cell(Grid1.ActiveCell.row, 3).text
Call GrabarPermiso(cuenta, Permiso)


End Sub
Private Sub Usuario_Change()
CargarCuentas
End Sub

Private Sub GrabarPermiso(cuenta, Permiso)
Dim csql As New rdoQuery
Set csql.ActiveConnection = conta
    
If Permiso = 1 Then
csql.sql = "insert ignore into permisos_cuentas (usuario,cuenta,permiso) values('" & Usuario.Caption & "','"
csql.sql = csql.sql & cuenta & "','" & Permiso & "')"
End If

If Permiso = 0 Then
csql.sql = "delete from permisos_cuentas where cuenta ='" & cuenta & "' and usuario = '" & Usuario.Caption & "'"

End If
csql.Execute
'Call sincronizadatos(csql1.sql, conta, "")


End Sub
