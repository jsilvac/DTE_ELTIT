VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form moduloseguridad2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MODULO DE SEGURIDAD"
   ClientHeight    =   10065
   ClientLeft      =   1260
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   671
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   WindowState     =   2  'Maximized
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   555
      Left            =   9180
      TabIndex        =   31
      Top             =   9180
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   979
      BackColor       =   16744576
      Caption         =   "ESTADO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   240
         Left            =   45
         TabIndex        =   32
         Top             =   270
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00800000&
      Caption         =   "BUSCAR"
      Height          =   330
      Left            =   12330
      TabIndex        =   14
      Top             =   7920
      Width           =   1320
   End
   Begin XPFrame.FrameXp MENU 
      Height          =   7665
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   15840
      _ExtentX        =   27940
      _ExtentY        =   13520
      BackColor       =   16744576
      Caption         =   "MENU"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.CommandButton Command2 
         Caption         =   "Borrar FIltros"
         Height          =   510
         Left            =   90
         TabIndex        =   30
         Top             =   7065
         Width           =   555
      End
      Begin VB.OptionButton FILTRO3 
         Caption         =   "FILTRO3"
         Height          =   195
         Left            =   10485
         TabIndex        =   29
         Top             =   7065
         Width           =   4605
      End
      Begin VB.OptionButton FILTRO2 
         Caption         =   "FILTRO2"
         Height          =   195
         Left            =   5760
         TabIndex        =   28
         Top             =   7065
         Width           =   4650
      End
      Begin VB.OptionButton FILTRO1 
         Caption         =   "FILTRO1"
         Height          =   195
         Left            =   765
         TabIndex        =   27
         Top             =   7065
         Width           =   4875
      End
      Begin VB.TextBox FIL1 
         Height          =   330
         Left            =   720
         TabIndex        =   26
         Top             =   7245
         Width           =   4920
      End
      Begin VB.TextBox FIL3 
         Height          =   330
         Left            =   10440
         TabIndex        =   25
         Top             =   7245
         Width           =   4650
      End
      Begin VB.TextBox FIL2 
         Height          =   330
         Left            =   5760
         TabIndex        =   4
         Top             =   7245
         Width           =   4650
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6000
         Left            =   90
         TabIndex        =   3
         Top             =   945
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   10583
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   885
         Left            =   1980
         TabIndex        =   5
         Top             =   45
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1561
         BackColor       =   16761024
         Caption         =   "Fecha Consultar"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.TextBox HASTA1 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1425
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "fecha"
            Top             =   525
            Width           =   375
         End
         Begin VB.TextBox HASTA2 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1785
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "fecha"
            Top             =   525
            Width           =   375
         End
         Begin VB.TextBox HASTA3 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2145
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "fecha"
            Top             =   525
            Width           =   615
         End
         Begin VB.TextBox DESDE1 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   30
            MaxLength       =   2
            TabIndex        =   8
            Tag             =   "fecha"
            Top             =   525
            Width           =   375
         End
         Begin VB.TextBox DESDE2 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   390
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "fecha"
            Top             =   525
            Width           =   375
         End
         Begin VB.TextBox DESDE3 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   750
            MaxLength       =   4
            TabIndex        =   6
            Tag             =   "fecha"
            Top             =   525
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DESDE"
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
            Left            =   30
            TabIndex        =   13
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HASTA"
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
            Left            =   1425
            TabIndex        =   12
            Top             =   285
            Width           =   1335
         End
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   0
         Top             =   0
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
      Begin VB.Label usua 
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
         Height          =   375
         Left            =   5445
         TabIndex        =   24
         Top             =   495
         Width           =   5595
      End
   End
   Begin XPFrame.FrameXp FRMUSUARIO 
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   7785
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   3836
      BackColor       =   8454016
      Caption         =   "USUARIOS"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid2 
         Height          =   1725
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   3043
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1200
      Left            =   9135
      TabIndex        =   15
      Top             =   7785
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   2117
      BackColor       =   16761024
      Caption         =   "USUARIOS"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Individuales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   23
         Top             =   675
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   22
         Top             =   315
         Width           =   1005
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   2145
      Left            =   7380
      TabIndex        =   16
      Top             =   7785
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   3784
      BackColor       =   16761024
      Caption         =   "Ver Eventos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   21
         Top             =   1755
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Elimino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   20
         Top             =   1395
         Width           =   1005
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   19
         Top             =   990
         Width           =   1320
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Creo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   18
         Top             =   585
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Visito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   17
         Top             =   225
         Width           =   1005
      End
   End
End
Attribute VB_Name = "moduloseguridad2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FORMATOGRILLA(12, 20)
Private VARIABLE As String
Private USUARIOSELECCIONADO As String
Private menuseleccion As String
Private modifo As Double
Private eli As Boolean
Private paso1 As String

Private Sub busqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click

End If

End Sub

Private Sub Command1_Click()

Call CARGAGRILLauditoria(1, 13)

leerauditoria

End Sub

Private Sub COMMAND2_Click()
FIL1.text = ""
FIL2.text = ""
FIL3.text = ""
FILTRO1.Value = False
FILTRO2.Value = False
FILTRO3.Value = False
FILTRO1.Caption = "FILTRO1"
FILTRO2.Caption = "FILTRO2"
FILTRO3.Caption = "FILTRO3"


End Sub

Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption
usua.Caption = USUARIOSISTEMA
USUARIOSELECCIONADO = USUARIOSISTEMA

End Sub

Private Sub Form_Load()
Dim k As Integer

  '==================================
    'PERMITE UNA INSTANCIA DEL SISTEMA
    '==================================
DESDE1.text = Format(fechasistema, "dd")
DESDE2.text = Format(fechasistema, "mm")
DESDE3.text = Format(fechasistema, "yyyy")
HASTA1.text = Format(fechasistema, "dd")
HASTA2.text = Format(fechasistema, "mm")
HASTA3.text = Format(fechasistema, "yyyy")

Call CARGAGRILLauditoria(1, 13)
Call CARGAGRILLAUSUARIOS(1, 3)
LEERUSUARIOS
leerauditoria



End Sub




Sub LEERUSUARIOS()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim LINEAS As Double
    

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT * "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "auditoria.segu_usuarios "
        csql2.sql = csql2.sql + "order by usuario "
        csql2.Execute
        Grid2.Rows = csql2.RowsAffected + 1
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        LINEAS = 0
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        Grid2.Cell(LINEAS, 1).text = resultados2(1)
        Grid2.Cell(LINEAS, 2).text = resultados2(3)
        Grid2.Cell(LINEAS, 3).text = resultados2(4)
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
      
 
  

End Sub
Sub LEERUSUARIOindividual(Usuario)
  
End Sub

Private Function achica(palabra) As String
Dim inicio As Double
Dim FINAL As Double
For k = 1 To Len(palabra)
If Mid(palabra, k, 1) <> Chr(32) Then inicio = k: Exit For

Next k

achica = Mid(palabra, inicio, Len(palabra) - inicio)

End Function

Sub ACTIVAMENU(ByVal opcion As String)


'For K = 1 To ingresos.Count
'
'
'If ingresos(K).caption = Opcion Then ingresos(K).Checked = True
'
'
'Next K
'

End Sub

Private Sub Grid4_Click()
If Grid2.Cell(Grid2.ActiveCell.row, 1).text <> "" Then

End If

End Sub

Private Sub MENU1_Click()
Dim contador As Double
Dim inicio As Double
Dim FINAL As Double
Dim pasar As Double
Dim NIVEL As String
Dim NIVELBANDERA As String
End Sub

'Sub LEERauditoria()
'    Dim resultados2 As rdoResultset
'    Dim cSql2 As New rdoQuery
'    Dim saldodebe As String
'    Dim saldohaber As String
'    Dim lineas As Double
'    Dim evento As String
'
'
'        Set cSql2.ActiveConnection = conAuditoria
'
'
'
'        cSql2.sql = "SELECT usuario,fecha,hora,evento,programa,glosa,solicitado,basedatos,tabla,campos,datosoriginales,datosmodificados "
'        cSql2.sql = cSql2.sql + "FROM auditoriagestion  "
'        USUARIOSISTEMA = "MKRAUSE"
'        cSql2.sql = cSql2.sql + "where fecha between '" + DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text + "' and '" + HASTA3.text + "-" + HASTA2.text + "-" + HASTA1.text + "' AND usuario='" + USUARIOSISTEMA + "' "
'        cSql2.sql = cSql2.sql + " order by fecha,hora "
'        cSql2.Execute
'        Grid1.Rows = 1
'
'        If cSql2.RowsAffected > 0 Then
'        Set resultados2 = cSql2.OpenResultset(3, , 1)
'        lineas = 0
'        While Not resultados2.EOF
'        Grid1.Rows = Grid1.Rows + 1
'
'        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
'        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados2(1)
'        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados2(2)
'        If resultados2(3) = "0" Then evento = "VISITO"
'        If resultados2(3) = "2" Then evento = "CREO"
'        If resultados2(3) = "3" Then evento = "MODIFICO"
'        If resultados2(3) = "4" Then evento = "ELIMINO"
'        Grid1.Cell(Grid1.Rows - 1, 4).text = evento
'       Grid1.Cell(Grid1.Rows - 1, 5).text = resultados2(4)
'
'
'
'
'
''
'        Grid1.Cell(Grid1.Rows - 1, 7).text = resultados2(6)
'        Grid1.Cell(Grid1.Rows - 1, 8).text = resultados2(7)
''        If IsNull(resultados2(8)) = False Then
''
'        Grid1.Cell(Grid1.Rows - 1, 9).text = resultados2(8)
''        End If
''        Grid1.Cell(Grid1.Rows - 1, 10).text = resultados2(9)
''        If IsNull(resultados2(10)) = False Then
''        Grid1.Cell(Grid1.Rows - 1, 11).text = resultados2(10)
''        End If
''
''        Grid1.Cell(Grid1.Rows - 1, 12).text = resultados2(11)
''
''
'        If evento = "ELIMINO" Then
'        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HFF&
'
'        End If
'        If evento = "MODIFICO" Then
'
'        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &H80FF80
'
'        End If
'
'
'
'
'        resultados2.MoveNext
'
'        Wend
'
'          resultados2.Close
'            Set resultados2 = Nothing
'
'        End If
'
'
'
'
'End Sub
Private Sub leerauditoria()
        Dim tabla As String
        Dim EVENTO As String
        
        tabla = "SELECT usuario,fecha,ifnull(hora,'00:00:00'),evento,programa,glosa,solicitado,basedatos,tabla,campos,datosoriginales,datosmodificados "
        tabla = tabla + "FROM auditoriacontabilidad  "
        tabla = tabla + "where fecha between '" + DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text + "' and '" + HASTA3.text + "-" + HASTA2.text + "-" + HASTA1.text + "' "
        If Option7.Value = True Then
        tabla = tabla + " AND usuario='" + USUARIOSELECCIONADO + "' "
        End If
        
        If Option1.Value = True Then
        tabla = tabla + " and evento='0' "
        End If
        If Option2.Value = True Then
        tabla = tabla + " and evento='2' "
        End If
        If Option3.Value = True Then
        tabla = tabla + " and evento='3' "
        End If
        If Option4.Value = True Then
        tabla = tabla + " and evento='4' "
        End If
        
        If FIL1.text <> "" Then
    
        tabla = tabla + " and " + FILTRO1.Caption + " like '%" + FIL1.text + "%' "
        End If
        If FIL2.text <> "" Then
        tabla = tabla + " and " + FILTRO2.Caption + " like '%" + FIL2.text + "%' "
        End If
        If FIL3.text <> "" Then
        tabla = tabla + " and " + FILTRO3.Caption + " like '%" + FIL3.text + "%' "
        End If

        tabla = tabla + " order by fecha,hora "
        
        Call ConectarControlData(data, Servidor, clientesistema + "auditoria", Usuario, password, tabla)
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        barra.Max = data.Recordset.RecordCount + 1
        barra.Value = 0
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
        Grid1.Rows = Grid1.Rows + 1
        barra.Value = barra.Value + 1
        barra.Refresh
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = data.Recordset(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = data.Recordset(1)
        Grid1.Cell(Grid1.Rows - 1, 3).text = data.Recordset(2)
        If data.Recordset(3) = "0" Then EVENTO = "VISITO"
        If data.Recordset(3) = "2" Then EVENTO = "CREO"
        If data.Recordset(3) = "3" Then EVENTO = "MODIFICO"
        If data.Recordset(3) = "4" Then EVENTO = "ELIMINO"
        Grid1.Cell(Grid1.Rows - 1, 4).text = EVENTO
        Grid1.Cell(Grid1.Rows - 1, 5).text = data.Recordset(4)
        Grid1.Cell(Grid1.Rows - 1, 6).text = data.Recordset(5)
        Grid1.Cell(Grid1.Rows - 1, 7).text = data.Recordset(6)
        Grid1.Cell(Grid1.Rows - 1, 8).text = data.Recordset(7)
        Grid1.Cell(Grid1.Rows - 1, 9).text = data.Recordset(8)
        Grid1.Cell(Grid1.Rows - 1, 10).text = data.Recordset(9)
        Grid1.Cell(Grid1.Rows - 1, 11).text = data.Recordset(10)
        Grid1.Cell(Grid1.Rows - 1, 12).text = data.Recordset(11)
        
        If EVENTO = "ELIMINO" Then
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).ForeColor = &HFF&

        End If
        If EVENTO = "MODIFICO" Then

        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).ForeColor = &H800000

        End If
'
             
                data.Recordset.MoveNext
            Wend
        Grid1.AutoRedraw = True
        Grid1.Refresh
        End If
    End Sub
Sub CARGAGRILLauditoria(row, col)
    Rem DATOS DE LA COLUMNA
    col = 14
    row = 1
    FORMATOGRILLA(1, 1) = "usuario"
    FORMATOGRILLA(1, 2) = "fecha"
    FORMATOGRILLA(1, 3) = "hora"
    FORMATOGRILLA(1, 4) = "evento"
    FORMATOGRILLA(1, 5) = "programa"
    FORMATOGRILLA(1, 6) = "glosa"
    FORMATOGRILLA(1, 7) = "solicitado"
    FORMATOGRILLA(1, 8) = "basedatos"
    FORMATOGRILLA(1, 9) = "tabla"
    FORMATOGRILLA(1, 10) = "campos"
    FORMATOGRILLA(1, 11) = "datosoriginales"
    FORMATOGRILLA(1, 12) = "datosmodificados"
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "6"
    FORMATOGRILLA(2, 3) = "6"
    FORMATOGRILLA(2, 4) = "6"
    FORMATOGRILLA(2, 5) = "20"
    FORMATOGRILLA(2, 6) = "20"
    FORMATOGRILLA(2, 7) = "8"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "20"
    FORMATOGRILLA(2, 10) = "20"
    FORMATOGRILLA(2, 11) = "20"
    FORMATOGRILLA(2, 12) = "20"
    FORMATOGRILLA(2, 13) = "4"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    Rem LOCCKED
    For k = 1 To 13
    FORMATOGRILLA(5, k) = "true"
    Next k
    
    
    Grid1.Cols = col
    Grid1.Rows = row
    Grid1.AllowUserResizing = True
    Grid1.DisplayFocusRect = False
    Grid1.AllowUserSort = True
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
    Grid1.Column(0).Width = 0
    
    
            For k = 1 To col - 1
            Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
            Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * 10
            Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
            Rem Grid1.Column(k).FormatString = formatoGrilla(4, k)
            Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
            If FORMATOGRILLA(3, k) = "S" Then
                Grid1.Column(k).Alignment = cellLeftCenter
            Else
                
                Grid1.Column(k).Alignment = cellRightCenter
            End If
            Grid1.Cell(0, k).Alignment = cellCenterCenter
        Next k
    
   ' Grid1.Column(7).CellType = cellTextBox
   ' Grid1.Column(8).CellType = cellTextBox
   ' Grid1.Column(9).CellType = cellTextBox
   ' Grid1.Column(11).CellType = cellTextBox

    If Option1.Value = True Then
        Grid1.Column(6).Width = 0
        Grid1.Column(7).Width = 0
        Grid1.Column(8).Width = 0
       
        Grid1.Cell(0, 9).text = "EMPRESA"
        Grid1.Column(10).Width = 0
        Grid1.Column(11).Width = 0
        Grid1.Column(12).Width = 0
       
    End If
    If Option2.Value = True Or Option3.Value = True Then
        Grid1.Column(6).Width = 0
        Grid1.Column(7).Width = 0
       
    End If
    
    
    
    
End Sub

Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col <> 4 Then
If FILTRO1.Value = True Then
FIL1.text = Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text
FILTRO1.Caption = Grid1.Cell(0, Grid1.ActiveCell.col).text

End If
If FILTRO2.Value = True Then
FIL2.text = Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text
FILTRO2.Caption = Grid1.Cell(0, Grid1.ActiveCell.col).text

End If
If FILTRO3.Value = True Then
FIL3.text = Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text
FILTRO3.Caption = Grid1.Cell(0, Grid1.ActiveCell.col).text

End If
End If

End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
detalleauditoria.Show vbModal


End If

End Sub

Private Sub Grid2_DblClick()
USUARIOSELECCIONADO = Grid2.Cell(Grid2.ActiveCell.row, 1).text
usua.Caption = usuarioseleecionado

Command1_Click

End Sub
Sub CARGAGRILLAUSUARIOS(row, col)
    Rem DATOS DE LA COLUMNA
    col = 4
    row = 1
    FORMATOGRILLA(1, 1) = "USUARIO"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "LABOR"
    FORMATOGRILLA(1, 4) = "EVENTO"
    FORMATOGRILLA(1, 5) = "PROGRAMA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "SOLICITADO"
    FORMATOGRILLA(1, 8) = "BASEDATOS"
    FORMATOGRILLA(1, 9) = "TABLA"
    FORMATOGRILLA(1, 10) = "CAMPOS"
    FORMATOGRILLA(1, 11) = "ORIGINALES"
    FORMATOGRILLA(1, 12) = "MODIFICADOS"
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "30"
    FORMATOGRILLA(2, 3) = "6"
    FORMATOGRILLA(2, 4) = "6"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "20"
    FORMATOGRILLA(2, 7) = "20"
    FORMATOGRILLA(2, 8) = "20"
    FORMATOGRILLA(2, 9) = "20"
    FORMATOGRILLA(2, 10) = "20"
    FORMATOGRILLA(2, 11) = "20"
    FORMATOGRILLA(2, 12) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    Rem LOCCKED
    For k = 1 To 12
    FORMATOGRILLA(5, k) = "true"
    Next k
    
    
    Grid2.Cols = col
    Grid2.Rows = row
    Grid2.AllowUserResizing = True
    Grid2.DisplayFocusRect = False
    Grid2.AllowUserSort = True
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    
    Grid2.Column(0).Width = 0
    
    
            For k = 1 To col - 1
            Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
            Grid2.Column(k).Width = Val(FORMATOGRILLA(2, k)) * 10
            Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
            Rem GRID2.Column(k).FormatString = formatoGrilla(4, k)
            Grid2.Column(k).Locked = FORMATOGRILLA(5, k)
            If FORMATOGRILLA(3, k) = "S" Then
                Grid2.Column(k).Alignment = cellLeftCenter
            Else
                
                Grid2.Column(k).Alignment = cellRightCenter
            End If
            Grid2.Cell(0, k).Alignment = cellCenterCenter
        Next k
    
   ' GRID2.Column(7).CellType = cellTextBox
   ' GRID2.Column(8).CellType = cellTextBox
   ' GRID2.Column(9).CellType = cellTextBox
   ' GRID2.Column(11).CellType = cellTextBox

    
    
    
End Sub

Private Sub Option1_Click()
Command1_Click

End Sub

Private Sub Option2_Click()
Command1_Click

End Sub

Private Sub Option3_Click()
Command1_Click

End Sub

Private Sub Option4_Click()
Command1_Click

End Sub

Private Sub Option5_Click()
Command1_Click

End Sub

Private Sub DESDE1_GotFocus()
    Call cargatexto(DESDE1)
End Sub
'
Private Sub DESDE1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE1, DESDE2, KeyCode)
End Sub

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And DESDE1.text <> "" Then Call ceros(DESDE1): DESDE2.SetFocus
End Sub

Private Sub DESDE2_GotFocus()
    Call cargatexto(DESDE2)
End Sub

Private Sub DESDE2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE1, DESDE3, KeyCode)
End Sub

Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And DESDE2.text <> "" Then Call ceros(DESDE2): DESDE3.SetFocus
End Sub

Private Sub DESDE3_GotFocus()
    Call cargatexto(DESDE3)
End Sub

Private Sub DESDE3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE2, HASTA1, KeyCode)
End Sub

Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And DESDE3.text <> "" Then Call ceros(DESDE3): HASTA1.SetFocus
End Sub

 

Private Sub HASTA1_GotFocus()
    Call cargatexto(HASTA1)
End Sub

Private Sub HASTA1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE3, HASTA2, KeyCode)
End Sub

Private Sub HASTA1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And HASTA1.text <> "" Then Call ceros(HASTA1): HASTA2.SetFocus
End Sub

Private Sub HASTA2_GotFocus()
    Call cargatexto(HASTA2)
End Sub

Private Sub HASTA2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(HASTA1, HASTA3, KeyCode)
End Sub

Private Sub HASTA2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And HASTA2.text <> "" Then Call ceros(HASTA2): HASTA3.SetFocus
End Sub

Private Sub HASTA3_GotFocus()
    Call cargatexto(HASTA3)
End Sub

Private Sub HASTA3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(HASTA2, HASTA3, KeyCode)
End Sub

Private Sub HASTA3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And HASTA3.text Then Call ceros(HASTA3): Command1.SetFocus
End Sub


