VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form contratos 
   Caption         =   "Contratos"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "IMPRIMIR PAGARÉ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2400
      Width           =   1455
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   6540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11536
      BackColor       =   16761024
      Caption         =   "Datos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox montodeuda 
         Height          =   375
         Left            =   6360
         TabIndex        =   36
         Text            =   "0"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1815
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
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   31
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1455
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
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   30
         Tag             =   "proveedor"
         Top             =   765
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIMIR CON AVALISTA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4080
         TabIndex        =   20
         Top             =   2400
         Width           =   1940
      End
      Begin VB.CommandButton BTNIMPRIME 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR CONTRATO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton BTNMANDATO 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR MANDATO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2400
         Width           =   1455
      End
      Begin FlexCell.Grid impresion 
         Height          =   735
         Left            =   6600
         TabIndex        =   18
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         Cols            =   5
         DefaultFontName =   "Courier"
         DefaultFontSize =   9.75
         Rows            =   30
      End
      Begin VB.TextBox dato27 
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
         Height          =   285
         Left            =   10635
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   0
         Width           =   1455
      End
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
         Height          =   285
         Left            =   6600
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   360
         Width           =   615
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
         Height          =   285
         Left            =   6210
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   360
         Width           =   375
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
         Height          =   285
         Left            =   5850
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   360
         Width           =   375
      End
      Begin XPFrame.FrameXp frmavalista 
         Height          =   2415
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4260
         BackColor       =   16744576
         Caption         =   "Datos Avalista"
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
         Enabled         =   0   'False
         Begin VB.TextBox dato7 
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
            Height          =   285
            Left            =   1560
            MaxLength       =   9
            TabIndex        =   35
            Tag             =   "proveedor"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text1 
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
            Left            =   1560
            TabIndex        =   34
            Tag             =   "proveedor"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox dato10 
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
            Left            =   1560
            TabIndex        =   29
            Tag             =   "proveedor"
            Top             =   1605
            Width           =   6255
         End
         Begin VB.TextBox dato9 
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
            Left            =   1560
            TabIndex        =   28
            Tag             =   "proveedor"
            Top             =   1245
            Width           =   6255
         End
         Begin VB.TextBox dato8 
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
            Left            =   1560
            TabIndex        =   27
            Tag             =   "proveedor"
            Top             =   885
            Width           =   6255
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Monto Avalado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   1395
         End
         Begin VB.Label lbldv1 
            BackStyle       =   0  'Transparent
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
            Height          =   285
            Left            =   3075
            TabIndex        =   26
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
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
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   885
            Width           =   1395
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CIUDAD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   1605
            Width           =   1395
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DIRECCION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   1245
            Width           =   1395
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
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
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1395
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   4500
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   765
         Width           =   1275
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DIRECCION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   13
         Top             =   1530
         Width           =   1275
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CIUDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FONO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4500
         TabIndex        =   10
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label LBLNOMBRE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1485
         TabIndex        =   9
         Top             =   1170
         Width           =   6450
      End
      Begin VB.Label LBLFONO 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5850
         TabIndex        =   8
         Top             =   1890
         Width           =   2085
      End
      Begin VB.Label LBLCIUDAD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1485
         TabIndex        =   7
         Top             =   1890
         Width           =   2940
      End
      Begin VB.Label LBLDIRECCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1485
         TabIndex        =   6
         Top             =   1530
         Width           =   6450
      End
      Begin VB.Label lbldv 
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   2970
         TabIndex        =   5
         Top             =   765
         Width           =   330
      End
   End
End
Attribute VB_Name = "contratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CARGAGRILLA()
        Dim col As Integer
        Dim row As Integer
        col = 2
        row = 2
        Dim formatogrilla(10, 10) As String
        Rem DATOS DE LA COLUMNA
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
'        FORMATOGRILLA(7, 9) = ""
'        Rem ANCHO DE LA COLUMNA
        formatogrilla(8, 1) = "1000"
        impresion.Cols = col
        impresion.Rows = row
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
        Rem Asigna Valores a la Grilla
        impresion.Cell(0, 0).text = formatogrilla(1, 0)
        For K = 1 To col - 1
            impresion.Cell(0, K).text = formatogrilla(1, K)
            impresion.Column(K).Width = Val(formatogrilla(8, K))
            impresion.Column(K).MaxLength = Val(formatogrilla(2, K))
            impresion.Column(K).FormatString = formatogrilla(4, K)
            impresion.Column(K).Locked = formatogrilla(5, K)
'            If formatogrilla(3, K) = "S" Then
'                impresion.Column(K).Alignment = cellLeftCenter
'            Else
'                impresion.Column(K).Alignment = cellRightCenter
'            End If
'            If formatogrilla(3, K) = "D" Then impresion.Column(K).CellType = cellCalendar
            impresion.Cell(0, K).Alignment = cellCenterCenter
        Next K
    impresion.Column(0).Width = 0
    End Sub
 
Private Sub BTNIMPRIME_Click()
Dim Word As Word.Application
Dim documento As Word.Documents

Dim MiDoc As String
If dato6.text <> "" Then
  'On Error GoTo controlerror
        Set Word = CreateObject("Word.Application")
        Word.Documents.Open rutaUpdate & "\contrato.doc"
        Word.Visible = True
        Word.Application.WindowState = wdWindowStateMinimize
        Word.Selection.Font.Size = 10
        
        Word.Documents(1).Bookmarks("numerodocumento").Range = dato1.text
        Word.Documents(1).Bookmarks("nombreusuario").Range = LBLNOMBRE.Caption
        Word.Documents(1).Bookmarks("domiciliousuario").Range = LBLDIRECCION.Caption
        Word.Documents(1).Bookmarks("ciusuario").Range = Format(dato6.text, "###,###,###") & "-" & lbldv.Caption
        Word.Documents(1).Bookmarks("ciudadusuario").Range = ciudadempresa & ", a " & Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " del año " & Format(fechasistema, "yyyy")
       
        Word.ActiveDocument.SaveAs rutaUpdate & "\contrato6.doc"
        Word.ActiveDocument.Close savechanges:=wdDoNotSaveChanges
        MiDoc = rutaUpdate & "\contrato6.doc"
        Word.Application.Documents.Open MiDoc
        
        Call grabafoliocontrato(dato1.text, dato6.text & lbldv.Caption, "CONTRATO")
        Exit Sub
controlerror:
  MsgBox "DEBE TENER INSTALADO MICROSOFT OFFICE EN SU PC O " & vbCrLf & "DOCUMENTOS NO SE ENCUENTRAN EN " & rutaUpdate, vbCritical, "ATENCION"
Else
    MsgBox "DEBE INGRESAR UN CLIENTE ", vbCritical, "ATENCION"
End If

End Sub
 
 
Private Sub BTNMANDATO_Click()
Dim Word As Word.Application
Dim Word2 As Word.Application
Dim MiDoc As String

 On Error GoTo controlerror
        Set Word = CreateObject("Word.Application")
        Word.Documents.Open "Z:\RESPALDO\promotora\MANDATO.doc"
        Word.Visible = True
        Word.Application.WindowState = wdWindowStateMinimize
        Word.Selection.Font.Size = 10
        
        Word.Documents(1).Bookmarks("folio").Range = dato1.text
        Word.Documents(1).Bookmarks("fecha").Range = "En la ciudad de " & ciudadempresa & ",a " & Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " del año " & Format(fechasistema, "yyyy")
        Word.Documents(1).Bookmarks("nombreusuario").Range = LBLNOMBRE.Caption
        Word.Documents(1).Bookmarks("domusuario").Range = LBLDIRECCION.Caption
        Word.Documents(1).Bookmarks("ciuusuario").Range = LBLCIUDAD.Caption
        Word.Documents(1).Bookmarks("nombreusuario2").Range = LBLNOMBRE.Caption
        Word.Documents(1).Bookmarks("rutusuario").Range = Format(dato6.text, "###,###,###") & "-" & lbldv.Caption
        Word.Documents(1).Bookmarks("rutusuario2").Range = Format(dato6.text, "###,###,###") & "-" & lbldv.Caption
        Word.Documents(1).Bookmarks("rutusuario3").Range = Format(dato6.text, "###,###,###") & "-" & lbldv.Caption
        
       
        Word.ActiveDocument.SaveAs "Z:\RESPALDO\promotora\" & dato6.text & "MANDATODEUDOR.doc"
        Word.ActiveDocument.Close savechanges:=wdDoNotSaveChanges
        
        MiDoc = "Z:\RESPALDO\promotora\" & dato6.text & "MANDATODEUDOR.doc"
        Word.Application.Documents.Open MiDoc
        
  If Check1.Value = 1 Then
        Set Word2 = CreateObject("Word.Application")
        Word2.Documents.Open "Z:\RESPALDO\promotora\MANDATOAVALISTA.doc"
        Word2.Visible = True
        Word2.Application.WindowState = wdWindowStateMinimize
        Word2.Selection.Font.Size = 10
 
         Word2.Documents(1).Bookmarks("numerodocumento").Range = dato1.text
         Word2.Documents(1).Bookmarks("fechadocumento").Range = "En la ciudad de " & ciudadempresa & ",a " & Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " del año " & Format(fechasistema, "yyyy")
         Word2.Documents(1).Bookmarks("nombrecliente").Range = dato8.text
         Word2.Documents(1).Bookmarks("rutcliente").Range = Format(dato7.text, "###,###,###") & "-" & lbldv1.Caption
         Word2.Documents(1).Bookmarks("direccioncliente").Range = dato9.text
         Word2.Documents(1).Bookmarks("ciudadcliente").Range = dato10.text
         Word2.Documents(1).Bookmarks("rutcliente2").Range = Format(dato7.text, "###,###,###") & "-" & lbldv1.Caption
 
         Word2.ActiveDocument.SaveAs "Z:\RESPALDO\promotora\" & dato7.text & "MANDATOAVALISTA.doc"
         Word2.ActiveDocument.Close savechanges:=wdDoNotSaveChanges
  
         MiDoc = "Z:\RESPALDO\promotora\" & dato7.text & "MANDATOAVALISTA.doc"
         Word2.Application.Documents.Open MiDoc
     End If
         Exit Sub
controlerror:
  MsgBox "DEBE TENER INSTALADO MICROSOFT OFFICE EN SU PC O " & vbCrLf & " DOCUMENTOS NO SE ENCUENTRAN EN " & rutaUpdate, vbCritical, "ATENCION"
  
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        frmavalista.Visible = True
        frmavalista.Enabled = True
        
        BTNMANDATO.Enabled = False
        dato7.SetFocus
    End If
    If Check1.Value = 0 Then
        dato7.text = ""
        dato8.text = ""
        dato9.text = ""
        dato10.text = ""
        BTNMANDATO.Enabled = True
        frmavalista.Visible = False
    End If
End Sub

Private Sub Command1_Click()
Dim Word As Word.Application
Dim documento As Word.Documents
Dim MiDoc As String
Dim FOLIO As String

FOLIO = leefoliocontrato("PAGARE")
If FOLIO = "0000000000" Then FOLIO = "0000000001"

If (Format(obtienefoliocontrato(dato6.text & lbldv.Caption, "PAGARE"), "0000000000")) <> "0000000000" Then
FOLIO = (Format(obtienefoliocontrato(dato6.text & lbldv.Caption, "PAGARE"), "0000000000"))
Else
Call grabafoliocontrato(FOLIO, dato6.text & lbldv.Caption, "PAGARE")

End If


montodeuda.Visible = True


If CDbl(montodeuda.text) > 0 Then
    If dato6.text <> "" Then
        Set Word = CreateObject("Word.Application")
        Word.Documents.Open "Z:\RESPALDO\promotora\pagare.doc"
        Word.Visible = True
        Word.Application.WindowState = wdWindowStateMinimize
        Word.Selection.Font.Size = 10
        Word.Documents(1).Bookmarks("folio").Range = FOLIO
        Word.Documents(1).Bookmarks("monto").Range = Format(CDbl(montodeuda.text), "###,###,###")
        Word.Documents(1).Bookmarks("montoletras").Range = WORDNUM(Format(CDbl(montodeuda.text), "#.###.###0"), "PESO", "PESOS", 0)
        Word.Documents(1).Bookmarks("nombretitular").Range = LBLNOMBRE.Caption
        Word.Documents(1).Bookmarks("domiciliotitular").Range = LBLDIRECCION.Caption
        Word.Documents(1).Bookmarks("ciudadtitular").Range = LBLCIUDAD.Caption
        Word.Documents(1).Bookmarks("ruttitular").Range = Format(dato6.text, "###,###,###") & "-" & lbldv.Caption
        Word.Documents(1).Bookmarks("nombretitular2").Range = LBLNOMBRE.Caption
        Word.Documents(1).Bookmarks("fechapagare").Range = Format(fechasistema, "yyyy-mm-dd")
        Word.Documents(1).Bookmarks("fechareal").Range = Format(Date, "yyyy-mm-dd")
        
        
        Word.ActiveDocument.SaveAs "Z:\RESPALDO\promotora\" & dato6.text & "pagare.doc"
        Word.ActiveDocument.Close savechanges:=wdDoNotSaveChanges
        MiDoc = "Z:\RESPALDO\promotora\" & dato6.text & "pagare.doc"
        Word.Application.Documents.Open MiDoc
        
    End If
Else

MsgBox "Debe ingresar el Monto del Pagare "

End If

End Sub



Private Sub dato1_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            dato3.SetFocus
        End If
End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And dato10.text <> "" Then
        BTNMANDATO.Enabled = True
        BTNMANDATO.SetFocus
    End If
End Sub

 Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "dd")
            End If
            dato4.SetFocus
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "mm")
            End If
          dato5.SetFocus
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "0000" Then
                dato5.text = Format(fechasistema, "yyyy")
            End If
           dato6.SetFocus
        End If
    End Sub
Private Sub dato3_LostFocus()
Call esfecha(dato3, dato4, dato5, "dd")
End Sub
Private Sub dato4_LostFocus()
Call esfecha(dato3, dato4, dato5, "mm")
End Sub
Private Sub dato5_LostFocus()
Call esfecha(dato3, dato4, dato5, "yyyy")
End Sub

Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    Call ayudaClienteSin(dato6, lbldv)
End If
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
Dim FOLIO As String

KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato6.text <> "" Then
    dato6.text = ceros(dato6)
    lbldv.Caption = rut(dato6.text)
    
    FOLIO = Format(obtienefoliocontrato(dato6.text & lbldv.Caption, "CONTRATO"), "0000000000")
    If FOLIO <> "0000000000" Then
        If MsgBox("CLIENTE YA TIENE NUMERO DE CONTRATO " & FOLIO & " DESEA UTILIZAR EL MISMO ? ", vbYesNo, "ATENCION") = vbYes Then
        dato1.text = FOLIO
        End If
    End If
    
    
    
    LBLNOMBRE.Caption = leerNombreCliente(dato6.text & lbldv.Caption)
    LBLDIRECCION.Caption = leerDireccionCliente(dato6.text & lbldv.Caption, "0")
    LBLFONO.Caption = leerFonoCliente(dato6.text & lbldv.Caption, "0")
    LBLCIUDAD.Caption = Replace(leerCiudadCliente(dato6.text & lbldv.Caption, "0"), " ", "")
    BTNIMPRIME.Enabled = True
    BTNIMPRIME.SetFocus
    Check1.Enabled = True
End If
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato7.text <> "" Then
    dato7.text = ceros(dato7)
    lbldv1.Caption = rut(dato7.text)
    dato8.SetFocus
End If
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And dato8.text <> "" Then
        dato9.SetFocus
    End If
End Sub

Private Sub dato9_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And dato9.text <> "" Then
        dato10.SetFocus
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
dato3.text = Format(fechasistema, "dd")
dato4.text = Format(fechasistema, "mm")
dato5.text = Format(fechasistema, "yyyy")
dato1.text = leefoliocontrato("CONTRATO")
CARGAGRILLA
  
End Sub

Private Sub Label15_Click()

End Sub
Private Function leefoliocontrato(TIPO) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    Set csql.ActiveConnection = ventas
    csql.sql = "select lpad(max(id)+1,10,'0') from " + clientesistema + "ventas.sv_documentos_credito WHERE TIPO='" & TIPO & "'"
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leefoliocontrato = resultados(0)
        resultados.Close
   Else
    leefoliocontrato = "0000000001"
    End If
    Set resultados = Nothing
   End Function
   
Private Function grabafoliocontrato(FOLIO, rut, TIPO)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    Set csql.ActiveConnection = ventas
    csql.sql = "insert IGNORE  into sv_documentos_credito values ('" + FOLIO + "','" + rut + "','" + TIPO + "','" + Format(fechasistema, "yyyy-mm-dd") + "','" + usuarioSistema + "')"
    csql.Execute
   End Function

Private Function obtienefoliocontrato(rut, TIPO) As Double
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    Set csql.ActiveConnection = ventas
    csql.sql = "SELECT id FROM sv_documentos_credito WHERE rut='" & rut & "' AND TIPO='" & TIPO & "'"
    csql.Execute
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    obtienefoliocontrato = resultados(0)
    resultados.Close
    Else
    obtienefoliocontrato = 0
    End If
    Set resultados = Nothing
End Function

