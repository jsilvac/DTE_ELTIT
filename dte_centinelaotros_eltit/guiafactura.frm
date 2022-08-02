VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form cartoladespacho 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GUIA DE DESPACHO"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   13996
      BackColor       =   16744576
      Caption         =   "DETALLES GUIAS"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox numerointerno 
         Height          =   285
         Left            =   5760
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox foliofiscal 
         Height          =   285
         Left            =   5760
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox dato9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Tag             =   "proveedor"
         Top             =   1370
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Retorno"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7200
         Width           =   1815
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "proveedor"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "proveedor"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3585
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3945
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4305
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   1080
         Width           =   1815
      End
      Begin FlexCell.Grid detalle 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8493
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   3195
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox dato1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2475
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "cartoladespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        Dim formatogrilla(20, 20) As String
        
        formatogrilla(1, 0) = "LN"
        formatogrilla(1, 1) = "GUIA"
        formatogrilla(1, 2) = "FECHA"
        formatogrilla(1, 3) = "LOCAL"
        formatogrilla(1, 4) = "CANTIDAD"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "D"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = "###,###,##0"
       
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
 
       
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
 
 
     
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
 
    
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "10"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "6"
 
 
       
            
        detalle.Cols = col
        detalle.Rows = row
        detalle.AllowUserResizing = False
        detalle.DisplayFocusRect = False
        detalle.ExtendLastCol = False
        detalle.BoldFixedCell = False
        detalle.DisplayRowIndex = True
        detalle.DrawMode = cellOwnerDraw
        detalle.Appearance = Flat
        detalle.ScrollBarStyle = Flat
        detalle.FixedRowColStyle = Flat
        detalle.BackColorFixed = RGB(90, 158, 214)
        detalle.BackColorFixedSel = RGB(110, 180, 230)
        detalle.BackColorBkg = RGB(90, 158, 214)
        detalle.BackColorScrollBar = RGB(231, 235, 247)
        detalle.BackColor1 = RGB(231, 235, 247)
        detalle.BackColor2 = RGB(239, 243, 255)
        detalle.GridColor = RGB(148, 190, 231)
        'detalle.DefaultFont.Size = 8
        
        
        
        detalle.Cell(0, 0).text = formatogrilla(1, 0)
        For i = 1 To col - 1
            detalle.Cell(0, i).text = formatogrilla(1, i)
            detalle.Column(i).Width = Val(formatogrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatogrilla(2, i))
            detalle.Column(i).FormatString = formatogrilla(4, i)
            detalle.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
                If i <> 5 And i <> 3 Then
                    detalle.Column(i).Mask = cellNumeric
                End If
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        
    
    End Sub

 

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub detalle_DblClick()
    If detalle.Rows > 1 Then
        Load Pdespachoflete
        Pdespachoflete.dato2.text = detalle.Cell(detalle.ActiveCell.row, 1).text
        Pdespachoflete.dato23.text = detalle.Cell(detalle.ActiveCell.row, 3).text
        Pdespachoflete.Show
        Pdespachoflete.leerguiadeafuera
        
    End If
End Sub

Private Sub Form_Load()
    Call CARGAGRILLA(1, 5)
End Sub
Sub cargadeafueraguias(localleer)
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select numero,fecha,cantidad from "
csql.sql = csql.sql & baseVentas & localleer & ".sv_guia_despacho_entrega_" & localleer & " "
csql.sql = csql.sql & "where tipodocumento='" & dato1.text & "' and (numerodocumento='" & foliofiscal.text & "' or numerodocumento='" & numerointerno.text & "') and codigo='" & dato8.text & "' "
csql.sql = csql.sql & "and cajadocumento='" & dato3.text & "' and localdocumento='" & dato4.text & "'  order by numero "
csql.Execute
If csql.RowsAffected > 0 Then
    
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        detalle.Rows = detalle.Rows + 1
        detalle.Cell(detalle.Rows - 1, 1).text = resultados(0)
        detalle.Cell(detalle.Rows - 1, 2).text = Format(resultados(1), "dd-mm-yyyy")
        detalle.Cell(detalle.Rows - 1, 3).text = localleer
        detalle.Cell(detalle.Rows - 1, 4).text = resultados(2)
        resultados.MoveNext
        
    Wend
    
    
End If

End Sub
 
