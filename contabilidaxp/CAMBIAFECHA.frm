VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form confi00 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fecha de Sistema"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00F5C9B1&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   1455
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "Calendario"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F5C9B1&
         Caption         =   "Confirma Fecha"
         Height          =   255
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox DATO3 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "codigoempresa"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox DATO2 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "codigoempresa"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox dato1 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "codigoempresa"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AÑO"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MES"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   1455
         Left            =   -240
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DIA"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   495
      End
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
      _Version        =   524288
      _ExtentX        =   8070
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   12648447
      Year            =   2006
      Month           =   6
      Day             =   5
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   12582912
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   8388608
      GridLinesColor  =   16744576
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      FillColor       =   &H00FF8080&
      Height          =   1695
      Left            =   1200
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "confi00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_DblClick()
dato1.text = Calendar1.Day
Call ceros(dato1)
dato2.text = Calendar1.Month
Call ceros(dato2)
dato3.text = Calendar1.Year
Calendar1.Visible = False
dato1.SetFocus

End Sub

Private Sub Command1_Click()
Dim PASO As String
fechasistema = dato1.text & "-" & dato2.text & "-" & dato3.text
PASO = dato1.text + "-" + dato2.text + "-" + dato3.text
If IsDate(PASO) = False Then dato1.SetFocus: GoTo no:
fechasistema = PASO
mes = dato2.text
año = dato3.text
dia = dato1.text



Unload Me
no:
End Sub

Private Sub Command2_Click()
Calendar1.Visible = True

End Sub

Private Sub dato1_GotFocus()
Call cargatexto(dato1)

End Sub
Private Sub dato2_GotFocus()
Call cargatexto(dato2)

End Sub
Private Sub dato3_GotFocus()
Call cargatexto(dato3)

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
    
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub
Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato3)
    Call Pregunta(dato3, dato1)
    If IsDate(dato1.text & "-" & dato2.text & "-" & dato3.text) = True Then
        Command1_Click
    Else
        MsgBox "FECHA NO ES VALIDA", vbCritical, "ATENCION"
        dato1.text = ""
        dato2.text = ""
        dato3.text = ""
        dato1.SetFocus
    End If
    End If

End Sub


Private Sub Form_Load()
dato1.text = Mid(fechasistema, 1, 2)
dato2.text = Mid(fechasistema, 4, 2)
dato3.text = Mid(fechasistema, 7, 4)
End Sub
Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

