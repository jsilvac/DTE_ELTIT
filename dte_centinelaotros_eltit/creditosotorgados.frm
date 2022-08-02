VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado1 
   Caption         =   "LISTADO DE CREDITOS OTORGADOS"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1005
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   1773
      BackColor       =   16761024
      Caption         =   ""
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
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8025
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   14155
      BackColor       =   16761024
      Caption         =   ""
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
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   5805
         TabIndex        =   3
         Top             =   7470
         Width           =   2760
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7080
         Left            =   45
         TabIndex        =   2
         Top             =   315
         Width           =   13650
         _ExtentX        =   24077
         _ExtentY        =   12488
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "tmplistado1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call CargaGrillaDocumentos(1, 7)
End Sub

 Private Sub CargaGrillaDocumentos(ByVal Row As Integer, ByVal Col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, 1) = "NUMERO"
        FORMATOGRILLA(1, 2) = "FECHA"
        FORMATOGRILLA(1, 3) = "SIT.COMERCIAL"
        FORMATOGRILLA(1, 4) = "CRÉDITO"
        
        Rem LARGO DE LOS DATOS
        FORMATOGRILLA(2, 1) = "10"
        FORMATOGRILLA(2, 2) = ""
        FORMATOGRILLA(2, 3) = ""
        FORMATOGRILLA(2, 4) = "9"
        
        Rem TIPO DE DATOS
        FORMATOGRILLA(3, 1) = "N"
        FORMATOGRILLA(3, 2) = "S"
        FORMATOGRILLA(3, 3) = "C"
        FORMATOGRILLA(3, 4) = "N"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        FORMATOGRILLA(4, 1) = "0000000000"
        FORMATOGRILLA(4, 2) = ""
        FORMATOGRILLA(4, 3) = ""
        FORMATOGRILLA(4, 4) = "$ ###,###,##0"
        
        Rem LOCCKED
        FORMATOGRILLA(5, 1) = "TRUE"
        FORMATOGRILLA(5, 2) = "TRUE"
        FORMATOGRILLA(5, 3) = "TRUE"
        FORMATOGRILLA(5, 4) = "TRUE"
        
        Rem VALOR MINIMO
        FORMATOGRILLA(6, 1) = ""
        FORMATOGRILLA(6, 2) = ""
        FORMATOGRILLA(6, 3) = ""
        FORMATOGRILLA(6, 4) = ""
        
        Rem VALOR MAXIMO
        FORMATOGRILLA(7, 1) = ""
        FORMATOGRILLA(7, 2) = ""
        FORMATOGRILLA(7, 3) = ""
        FORMATOGRILLA(7, 4) = ""
        
        Rem ANCHO
        FORMATOGRILLA(8, 1) = "10"
        FORMATOGRILLA(8, 2) = "30"
        FORMATOGRILLA(8, 3) = "10"
        FORMATOGRILLA(8, 4) = "9"
            
        Documentos.Cols = Col
        Documentos.Rows = Row
        Documentos.AllowUserResizing = False
        Documentos.DisplayFocusRect = False
        Documentos.ExtendLastCol = True
        Documentos.BoldFixedCell = False
        Documentos.DrawMode = cellOwnerDraw
        Documentos.Appearance = Flat
        Documentos.ScrollBarStyle = Flat
        Documentos.FixedRowColStyle = Flat
        Documentos.BackColorFixed = RGB(90, 158, 214)
        Documentos.BackColorFixedSel = RGB(110, 180, 230)
        Documentos.BackColorBkg = RGB(90, 158, 214)
        Documentos.BackColorScrollBar = RGB(231, 235, 247)
        Documentos.BackColor1 = RGB(231, 235, 247)
        Documentos.BackColor2 = RGB(239, 243, 255)
        Documentos.GridColor = RGB(148, 190, 231)
        
        Documentos.Column(0).Width = 0
        For i = 1 To Col - 1
            Documentos.Cell(0, i).text = FORMATOGRILLA(1, i)
            Documentos.Column(i).Width = Val(FORMATOGRILLA(8, i)) * (Documentos.Cell(0, i).Font.Size + 1.25)
            Documentos.Column(i).MaxLength = Val(FORMATOGRILLA(2, i))
            Documentos.Column(i).FormatString = FORMATOGRILLA(4, i)
            Documentos.Column(i).Locked = FORMATOGRILLA(5, i)
            If FORMATOGRILLA(3, i) = "N" Then
                Documentos.Column(i).Alignment = cellRightCenter
            End If
            If FORMATOGRILLA(3, i) = "S" Then
                Documentos.Column(i).Alignment = cellLeftCenter
            End If
            If FORMATOGRILLA(3, i) = "C" Then
                Documentos.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Alignment = cellCenterCenter
        Documentos.Enabled = True
    End Sub
'**
