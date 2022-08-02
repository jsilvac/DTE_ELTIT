VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form confi07 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fecha de Cierre"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Calendario"
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Confirma Fecha"
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
         BackColor       =   &H0000C000&
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   4695
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
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13322
         SubFormatType   =   0
      EndProperty
      Height          =   2295
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   4665
      _Version        =   524288
      _ExtentX        =   8229
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   33023
      Year            =   2008
      Month           =   6
      Day             =   2
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   4194304
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   4194304
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
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      FillColor       =   &H00C0FFC0&
      Height          =   1695
      Left            =   840
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "confi07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calendar1_DblClick()
    DATO1.text = Calendar1.Day
    Call ceros(DATO1)
    DATO2.text = Calendar1.Month
    Call ceros(DATO2)
    DATO3.text = Calendar1.Year
    Calendar1.Visible = False
    DATO1.SetFocus
End Sub

Private Sub Command1_Click()
    Dim PASO As String
    fechacierre = DATO1.text + "-" + DATO2.text + "-" + DATO3.text
'    PASO = dato1.text + "-" + dato2.text + "-" + dato3.text
   Call actualizafechacierre(fechacierre)
 
    Unload Me

End Sub

Private Sub Command2_Click()
    Calendar1.Visible = True
    Calendar1.Today
End Sub

Private Sub dato1_GotFocus()
    Call cargatexto(DATO1)
End Sub

Private Sub dato1_LostFocus()
'If DATO1.text <> "" Then
' Call esfechareal(DATO1, DATO2, DATO3, "dd")
'Else
'DATO1.SetFocus
'End If
End Sub

Private Sub dato2_GotFocus()
    Call cargatexto(DATO2)
End Sub

Private Sub DATO2_LostFocus()
'If DATO2.text <> "" Then
'Call esfechareal(DATO1, DATO2, DATO3, "mm")
'Else
'DATO2.SetFocus
'End If
End Sub

Private Sub dato3_GotFocus()
    Call cargatexto(DATO3)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO1): Call Pregunta(DATO1, DATO2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO2): Call Pregunta(DATO2, DATO3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(DATO3)
    If IsDate(DATO3.text + "/" + DATO2.text + "/" + DATO1.text) Then
   Command1.SetFocus
   
    
    Else
    MsgBox ("fecha digitada incorrecta")
    
    End If
    
    
    End If
    
End Sub

Private Sub DATO3_LostFocus()
'If DATO3.text <> "" Then
'Call esfechareal(DATO1, DATO2, DATO3, "yyyy")
'Else
'DATO3.SetFocus
'End If
End Sub


Private Sub Form_Load()
    DATO1.text = Mid(fechacierre, 1, 2)
    DATO2.text = Mid(fechacierre, 4, 2)
    DATO3.text = Mid(fechacierre, 7, 4)
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub
