VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmpgestion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESA GESTION COBRANZA"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7215
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   12726
      BackColor       =   16761024
      Caption         =   "Gestión Cobranza"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin FlexCell.Grid Grid2 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   5760
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   661
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.TextBox dato1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   0
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   6840
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   150
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   3015
         Left            =   240
         TabIndex        =   14
         Top             =   3600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5318
         BackColor       =   16744576
         Caption         =   "Eventos"
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
         Begin FlexCell.Grid Grid1 
            Height          =   2655
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   4683
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.CommandButton cmdgrabar 
         BackColor       =   &H00FF8080&
         Caption         =   "GRABAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtglosa 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3000
         Width           =   6375
      End
      Begin VB.Label lblnombreevento 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA MOROSIDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblfechamoroso 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblrut 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblmonto 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblhora 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MONTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblnombre 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE"
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label lblfecha 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GLOSA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   6375
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EVENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HORA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
   End
End
Attribute VB_Name = "tmpgestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdgrabar_Click()
If dato1.text <> "" And txtglosa.text <> "" Then
    grabar
    leer (lblrut.Caption)
    limpia
End If
End Sub

Private Sub dato1_GotFocus()
FrameXP1.Caption = FrameXP1.Caption & " F2 Ayuda Evento "
End Sub

Private Sub dato1_LostFocus()
FrameXP1.Caption = "Gestion Cobranza"
End Sub

Private Sub Form_Load()
Call CargaGrillaGRID1(10, 5)
 cmdgrabar.Visible = False
 
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudagastoscobranza(dato1)

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato1.text <> "" Then
    dato1.text = ceros(dato1)
    lblnombreevento.Caption = leernombreevento(dato1.text)
    If InStr(lblnombreevento.Caption, "DICOM") = 0 Then
        If lblnombreevento.Caption <> "" Then
            txtglosa.Locked = False
            txtglosa.SetFocus
        End If
    Else
        MsgBox "EN ESTE MODULO NO SE PERMITEN ESTOS MOVIMIENTOS", vbCritical, "ATENCION"
    End If
End If

End Sub

'Private Sub Form_Unload(Cancel As Integer)
'tmpcobranza.LEErclientes
'End Sub

Private Sub txtglosa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If txtglosa.text <> "" And dato1.text <> "" And KeyAscii = 13 Then
    cmdgrabar.Visible = True
    cmdgrabar.SetFocus
    
End If
End Sub
 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "FECHA"
        formatogrilla(1, 2) = "HORA"
        formatogrilla(1, 3) = "EVENTO"
        formatogrilla(1, 4) = "GLOSA"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "D"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "S"
 
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""

        Rem LOCCKED
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
        formatogrilla(8, 1) = "7"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "20"
 

            
        Grid1.Cols = col
        Grid1.Rows = row
        Grid1.AllowUserResizing = True
        Grid1.DisplayFocusRect = False
        Grid1.ExtendLastCol = True
        Grid1.BoldFixedCell = False
        Grid1.DrawMode = cellOwnerDraw
        Grid1.Appearance = Flat
        Grid1.ScrollBarStyle = Flat
        Grid1.FixedRowColStyle = Flat
        Grid1.BackColorFixed = RGB(90, 158, 214)
        Grid1.BackColorFixedSel = RGB(110, 180, 230)
        Grid1.BackColorBkg = RGB(90, 158, 214)
        Grid1.BackColorScrollBar = RGB(231, 235, 247)
        Grid1.BackColor1 = RGB(231, 235, 247)
        Grid1.BackColor2 = RGB(239, 243, 255)
        Grid1.GridColor = RGB(148, 190, 231)
        
        Grid1.Column(0).Width = 0
        For i = 1 To col - 1
            Grid1.Cell(0, i).text = formatogrilla(1, i)
            Grid1.Column(i).Width = Val(formatogrilla(8, i)) * (Grid1.Cell(0, i).Font.Size + 1.25)
            Grid1.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid1.Column(i).FormatString = formatogrilla(4, i)
            Grid1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid1.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                Grid1.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                Grid1.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
        Grid1.Enabled = True
    
    
    End Sub
Sub leer(rut)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
rut = Replace(rut, "-", "")
rut = Replace(rut, ".", "")
pivote.MaxLength = 10
pivote.text = rut
pivote.text = ceros(pivote)

rut = pivote.text

Set csql.ActiveConnection = ventas
csql.sql = "select fecha,hora,evento,glosa from "
csql.sql = csql.sql & "sv_cobranza_gestion "
csql.sql = csql.sql & "where rut='" & rut & "' "
csql.Execute
If csql.RowsAffected > 0 Then
Grid1.Rows = 1
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
        Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
        resultados.MoveNext
    Wend
End If
End Sub

Sub grabar()
    Dim CAMPOS(10, 10) As String
    Dim rut As String
    Dim op As Integer
    

    rut = Replace(lblrut.Caption, "-", "")
    pivote.MaxLength = 10
    rut = Replace(rut, ".", "")
    pivote.text = ceros(pivote)
    rut = pivote.text

    CAMPOS(0, 0) = "rut"
    CAMPOS(1, 0) = "fecha"
    CAMPOS(2, 0) = "hora"
    CAMPOS(3, 0) = "fechamorosidad"
    CAMPOS(4, 0) = "evento"
    CAMPOS(5, 0) = "monto"
    CAMPOS(6, 0) = "glosa"
    CAMPOS(7, 0) = ""
  
    CAMPOS(0, 1) = rut
    CAMPOS(1, 1) = Format(lblfecha.Caption, "yyyy-mm-dd")
    CAMPOS(2, 1) = lblhora.Caption
    CAMPOS(3, 1) = Format(lblfechamoroso.Caption, "yyyy-mm-dd")
    CAMPOS(4, 1) = dato1.text
    CAMPOS(5, 1) = CDbl(lblmonto.Caption)
    CAMPOS(6, 1) = txtglosa.text
 
    
    CAMPOS(0, 2) = "sv_cobranza_gestion"
      
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    condicion = "rut='" & rut & "' and fecha='" & Format(lblfecha.Caption, "yyyy-mm-dd") & "' and evento ='" & dato1.text & "' "
    op = 5
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 4 Then
        op = 2
        condicion = ""
        Call sqlventas.sqlventas(op, condicion)
        If leercargo(dato1.text) > 0 Then
            Call grabarcuotacobranza(rut, Format(lblfecha.Caption, "yyyy-mm-dd"), dato1.text, leernombreevento(dato1.text))
        End If

        If dato1.text = "01" Then
            Call CargaGrillacobranza(1, 11)
            FOLIOINTERNO = 0
            If dato1.text = "01" Then
                FOLIOINTERNO = FOLIOINTERNO + 1
                With tmpcobranza
                    Call LEErCREDITOS(rut, .dato3.text + "-" + .dato2.text + "-" + .dato1.text, .dato6.text + "-" + .dato5.text + "-" + .dato4.text, FOLIOINTERNO)
                End With
            End If
        End If
    Else
        MsgBox "NO PUEDE CARGAR DOBLE CARGO AL MISMO RUT Y FECHA  ", vbInformation, "ATENCION"
    End If
    
End Sub

 Private Sub CargaGrillacobranza(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "LO"
        formatogrilla(1, 2) = "F.COMPRA"
        formatogrilla(1, 3) = "TD"
        formatogrilla(1, 4) = "NUMERO"
        formatogrilla(1, 5) = "COMPRA"
        formatogrilla(1, 6) = "VENCIMIENTO"
        formatogrilla(1, 7) = "CUOTAS VENCIDAS"
        formatogrilla(1, 8) = "TOTAL CUOTAS"
        formatogrilla(1, 9) = "INT.MORA"
        formatogrilla(1, 10) = "TOTAL"
        
        Rem ANCHO
        formatogrilla(8, 1) = "0"
        formatogrilla(8, 2) = "0"
        formatogrilla(8, 3) = "0"
        formatogrilla(8, 4) = "0"
        formatogrilla(8, 5) = "30"
        formatogrilla(8, 6) = "15"
        formatogrilla(8, 7) = "20"
        formatogrilla(8, 8) = "20"
        formatogrilla(8, 9) = "0"
        formatogrilla(8, 10) = "0"
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 8) = "$ ###,###,###"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "D"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "D"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""

        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"

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
        
            
        Grid2.Cols = col
        Grid2.Rows = row
        Grid2.AllowUserResizing = False
        Grid2.DisplayFocusRect = False
        Grid2.ExtendLastCol = True
        Grid2.BoldFixedCell = False
        Grid2.DrawMode = cellOwnerDraw
        Grid2.Appearance = Flat
        Grid2.ScrollBarStyle = Flat
        Grid2.FixedRowColStyle = Flat
        Grid2.BackColorFixed = RGB(90, 158, 214)
        Grid2.BackColorFixedSel = RGB(110, 180, 230)
        Grid2.BackColorBkg = RGB(90, 158, 214)
        Grid2.BackColorScrollBar = RGB(231, 235, 247)
        Grid2.BackColor1 = RGB(231, 235, 247)
        Grid2.BackColor2 = RGB(239, 243, 255)
        Grid2.GridColor = RGB(148, 190, 231)
        Grid2.DefaultFont.Bold = True
        
        Grid2.Column(0).Width = 0
        For i = 1 To col - 1
            Grid2.Cell(0, i).text = formatogrilla(1, i)
            Grid2.Column(i).Width = Val(formatogrilla(8, i)) * (Grid2.Cell(0, i).Font.Size)
            Grid2.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid2.Column(i).FormatString = formatogrilla(4, i)
            Grid2.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid2.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                Grid2.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                Grid2.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Grid2.Range(0, 0, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
        Grid2.Enabled = True
    End Sub
    
    Sub LEErCREDITOS(rut, desde, hasta, FOLIO)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim CREDITO As Double
        Dim usado As Double
        Dim disponible As Double
        Dim mora As Double
        Dim total1 As Double
        Dim total2 As Double
        Dim total3 As Double
        Dim total4 As Double
        Dim total5 As Double
        Dim ACUMULADO As Double
        Dim FECHAMORA As String
        Dim MESMORA As String
        Dim AÑOMORA As String
        Dim linea As Double
        
        
        
       
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT cd.local,cd.fechacompra,cd.tipo,cd.numero,cd.glosacompra,cd.vencimientoactual,count(cd.numerocuota),cd.cantidadcuotas,sum(cd.montocuota-cd.abono )"
        csql.sql = csql.sql & "FROM sv_cuotas_detalle as cd "
        csql.sql = csql.sql & "WHERE cd.rut='" + rut + "' and  cd.vencimientoactual between '" + desde + "' and '" + hasta + "' and montocuota>abono  "
        
        
        csql.sql = csql.sql & "group by cd.vencimientoactual order by cd.vencimientoactual  "
        
        
        
        csql.Execute
        Grid2.Rows = 1
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
'        If Option1.Value = True Then separador = resultado(4)
'        If Option2.Value = True Then separador = resultado(6)
        
        
       
        
        Grid2.Rows = 1
        Grid2.AutoRedraw = False
        
        total1 = 0
        total2 = 0
        total3 = 0
        total4 = 0
        total5 = 0
        
        While Not resultado.EOF
        
        
        
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = resultado(0)
        Grid2.Cell(Grid2.Rows - 1, 2).text = Format(resultado(1), "dd-mm-yyyy")
        Grid2.Cell(Grid2.Rows - 1, 3).text = resultado(2)
        Grid2.Cell(Grid2.Rows - 1, 4).text = resultado(3)
        Grid2.Cell(Grid2.Rows - 1, 5).text = "TOTAL ACUMULADO DE LA FECHA"
        Grid2.Cell(Grid2.Rows - 1, 6).text = Format(resultado(5), "dd-mm-yyyy")
        Grid2.Cell(Grid2.Rows - 1, 7).text = Format(resultado(7), "###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 8).text = Format(resultado(8), "$ ###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 9).text = Format(mora, "###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 10).text = Format(resultado(8) + mora, "###,###,###")
     
        total1 = total1 + resultado(8)
        total2 = total2 + mora
        total3 = total3 + (resultado(8) + mora)
        
        total11 = total11 + resultado(8)
        total12 = total12 + mora
        total13 = total13 + (resultado(8) + mora)
        
        

        
    
            resultado.MoveNext
            Wend
        Else
       
        End If
         
        
        
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Range(Grid2.Rows - 1, 6, Grid2.Rows - 1, 10).Borders(cellEdgeTop) = cellThick
          
        Grid2.Cell(Grid2.Rows - 1, 6).text = "TOTAL DEUDA"
        Grid2.Cell(Grid2.Rows - 1, 8).text = Format(total11, "$ ###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 9).text = Format(total12, "###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 10).text = Format(total13, "$ ###,###,###")
       
    Grid2.Column(1).Locked = False
    Grid2.Column(2).Locked = False
    Grid2.Column(3).Locked = False
    Grid2.Column(4).Locked = False
    Grid2.Column(5).Locked = False
    Grid2.Column(6).Locked = False
    Grid2.Column(7).Locked = False
    Grid2.Column(8).Locked = False
    Grid2.Column(9).Locked = False
    Grid2.Column(10).Locked = False
     For K = 1 To 30 - Grid2.Rows - 1
        Grid2.Rows = Grid2.Rows + 1
     Next K
    
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 8
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid2.Cell(Grid2.Rows - 1, 1).text = " "
    
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 9
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid2.Cell(Grid2.Rows - 1, 1).text = "        Se le solicita tenga a bien regularizar cuanto antes esta situación. Al valor total de cuotas  "
   
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 9
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid2.Cell(Grid2.Rows - 1, 1).text = " se deben agregar los gastos de cobranza mas el interes por morosidad el que sera calculado al momento de "
   
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 9
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid2.Cell(Grid2.Rows - 1, 1).text = " de su regularizacion "
   
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 9
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid2.Cell(Grid2.Rows - 1, 1).text = "         Saluda a usted cordialmente,  "
   
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 8
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid2.Cell(Grid2.Rows - 1, 1).text = " "
    
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 8
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellCenterCenter
    Grid2.Cell(Grid2.Rows - 1, 1).text = "Departamento de Cobranzas  "
    
 
    
     Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 50
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellLeftGeneral
    Grid2.Cell(Grid2.Rows - 1, 1).text = " "
    
    Grid2.Rows = Grid2.Rows + 2
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).FontSize = 9
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).Alignment = cellCenterCenter
    
    Grid2.Cell(Grid2.Rows - 1, 1).text = "Nota: en caso de haber recibido esta cobranza con posterioridad a la regularización de esta deuda, "
    
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).FontSize = 9
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 9).Alignment = cellCenterCenter
    
    Grid2.Cell(Grid2.Rows - 1, 1).text = "por favor, sírvase dejar sin efecto esta comunicación. "
    
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 9
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Alignment = cellCenterCenter
    Grid2.Cell(Grid2.Rows - 1, 1).text = " **** el detalle de su deuda puede ser solicitado en oficina comercial ****"
        
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid2.AutoRedraw = True
  Call Titulos2("CARTA COBRANZA", leerNombreCliente(rut), leerDireccionCliente(rut, "0"), leerComunaCliente(rut, "0"), leerFonoCliente(rut, "0"), FOLIO)
  
  
  


Grid2.PageSetup.HeaderMargin = 0
Grid2.PageSetup.PrintFixedRow = True

Grid2.PageSetup.TopMargin = 0.5
Grid2.PageSetup.LeftMargin = 1
Grid2.PageSetup.RightMargin = 0.5
Grid2.PageSetup.BottomMargin = 3
Grid2.PageSetup.FooterMargin = 2
Grid2.PageSetup.BlackAndWhite = True
 Grid2.Refresh
 
        Grid2.PrintPreview
 
    
    Call CargaGrillacobranza(1, 12)
    End Sub
    Sub Titulos2(titulo1, nombre, direccion, comuna, fono, FOLIO)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    Dim K As Integer
    
    Grid2.FixedRowColStyle = Fixed3D
    Grid2.CellBorderColorFixed = vbButtonShadow
    Grid2.ShowResizeTips = False
    Grid2.ReportTitles.Clear
    Grid2.PageSetup.CenterHorizontally = True
    Grid2.PageSetup.Orientation = cellPortrait

    
      
    Grid2.PageSetup.PrintTitleRows = 0
    
    'Logo
    
'    grid2.Images.Add App.Path & "\logo.jpg", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = cellLeft
'    objReportTitle.Height = 60
'    grid2.ReportTitles.Add objReportTitle
    
    
''    'ENCABEZADO DE PAGINA
'    grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & leerRutEmpresa(empresaActiva)
'    grid2.PageSetup.HeaderAlignment = cellLeft
'    grid2.PageSetup.HeaderFont.Name = "Verdana"
'    grid2.PageSetup.HeaderFont.Size = 8

  'ENCABEZADO DE PAGINA
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "EMPRESAS CREDITOS ELTIT " & vbCrLf & "PROMOTORA PALGUIN FRESIA 289 Sdo Piso " & vbCrLf & "FONO :441349 ANEXO 350"
'    grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & leerRutEmpresa(empresaActiva)
    objReportTitle.Font.Name = "Verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Align = cellLeft
    Grid2.ReportTitles.Add objReportTitle
'    grid2.PageSetup.HeaderAlignment = cellLeft
'    grid2.PageSetup.HeaderFont.Name = "Verdana"
'    grid2.PageSetup.HeaderFont.Size = 8
''
    'TITULOS DEL REPORTE
    
   
'    If Option1.Value = True Then tipoListado = "CLIENTES MAAT"
'    If Option2.Value = True Then tipoListado = "CLIENTES SKORPIOS"
'    If Option3.Value = True Then tipoListado = "CLIENTES TODOS"
'
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = comunaempresa & "," & Format(fechasistema, ("dd")) & " de " & MonthName(Month(Now)) & " " & Format(fechasistema, ("yyyy")) & "   "
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold
    objReportTitle.Align = cellRight
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "SEÑOR(A)." & vbCrLf & nombre & vbCrLf & direccion & vbCrLf & comuna & vbCrLf & comuna
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "PRESENTE"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Italic = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = False
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "            Estima Cliente , por medio de la presente informamos a ud. que en nuestros registros de cobranza " & vbCrLf & "aun se encuentran la siguientes cuotas impagas "
    

    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 9
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
   
    'PIE DE PAGINA
  Rem   Grid2.PageSetup.Footer = " ESTIMADOS SEÑORES " & vbCrLf & "PRUEBA"
    
    
    Grid2.PageSetup.Footer = "FOLIO INTERNO :" + Format(FOLIO, "00000")
    
    
    Grid2.PageSetup.FooterAlignment = cellCenter
    Grid2.PageSetup.FooterFont.Name = "Verdana"
    Grid2.PageSetup.FooterFont.Size = 7
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThin
     Grid2.Range(0, 1, 0, Grid2.Cols - 1).FontBold = True
       
End Sub


Sub limpia()
dato1.text = ""
txtglosa.text = ""
lblnombreevento.Caption = ""
txtglosa.Locked = True
cmdgrabar.Visible = False
dato1.SetFocus

End Sub
 
