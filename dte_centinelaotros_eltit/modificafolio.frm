VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form modificafolio 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "CONFIRMA FOLIO"
   ClientHeight    =   2565
   ClientLeft      =   5115
   ClientTop       =   3510
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4471
      BackColor       =   16744576
      Caption         =   "NUMERO DE FOLIO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
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
      Begin VB.TextBox tipodocu 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton nada 
         Caption         =   "OK"
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
         Left            =   840
         TabIndex        =   0
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cambio 
         Caption         =   "Modificar"
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
         Left            =   2760
         TabIndex        =   2
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox nuevofolio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox FOLIOFACTURA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CAMBIE FOLIO EN ""NUEVO FOLIO "" Y PRESIONE MODIFICAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4080
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "OK SI FOLIO ESTA BIEN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nuevo Folio"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Folio Actual"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "modificafolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cambio_Click()
 If nuevofolio.text <> "" Then
         nuevofolio.text = ceros(nuevofolio)
         If leerDocumentoexiste(nuevofolio.text) = False Then
            NUEVONUMEROGUIA = nuevofolio.text
            Unload Me
         Else
            MsgBox ("NUMERO DE " & tipodocu.text & " YA EMITIDO")
            nuevofolio.SetFocus
         End If
     Else
     nuevofolio.SetFocus
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub
Public Function leerDocumentoexiste(ByVal NUMERO As String) As Boolean
        
        Dim op As Integer
        Dim CAMPOS(10, 10)
        CAMPOS(0, 0) = "numero"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = "sv_guia_despacho_flete_" + empresaActiva
        condicion = " numero = '" & NUMERO & "'"
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventasRubro
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
            leerDocumentoexiste = True
         Else
            leerDocumentoexiste = False
        End If
    End Function

Private Sub FOLIOFACTURA_GotFocus()
 Call selecciona(FOLIOFACTURA)
End Sub

Private Sub FOLIOFACTURA_KeyPress(KeyAscii As Integer)
       Dim condon As String
End Sub

Private Sub Form_Load()
FOLIOFACTURA.text = NUEVONUMEROGUIA
End Sub

Private Sub Form_Unload(Cancel As Integer)
            nuevofolio.text = ""
            tipodocu.text = ""
            FOLIOFACTURA.text = ""
End Sub

Private Sub nada_Click()
 NUEVONUMEROGUIA = FOLIOFACTURA.text
 Unload Me
End Sub

Private Sub nada_GotFocus()
 FOLIOFACTURA.text = ceros(FOLIOFACTURA)
End Sub

Private Sub nuevofolio_GotFocus()
 FOLIOFACTURA.text = ceros(FOLIOFACTURA)
End Sub

Private Sub nuevofolio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And FOLIOFACTURA.text <> "" And nuevofolio.text <> "" Then
        nuevofolio.text = ceros(nuevofolio)
        cambio.SetFocus
       End If
End Sub
