VERSION 5.00
Begin VB.Form cargaFoto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione una Foto"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmbAceptar 
         Caption         =   "&Aceptar"
         Height          =   255
         Left            =   6120
         TabIndex        =   5
         Top             =   3120
         Width           =   1455
      End
      Begin VB.FileListBox Archivos 
         Appearance      =   0  'Flat
         Height          =   3150
         Left            =   3480
         Pattern         =   "*.jpg; *.jpeg"
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.DirListBox Directorios 
         Appearance      =   0  'Flat
         Height          =   2790
         Left            =   240
         TabIndex        =   3
         Top             =   650
         Width           =   3015
      End
      Begin VB.DriveListBox Discos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image Preview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad de Disco:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1230
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3615
         Left            =   0
         Top             =   0
         Width           =   7815
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   3615
      Left            =   240
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "cargaFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private dir As String
    Private manejoArchivo As Guga

Private Sub Archivos_Click()
    Me.Preview.Picture = LoadPicture(Me.Directorios.path & "\" & Me.Archivos.FileName)
End Sub

Private Sub Archivos_DblClick()
    cmbAceptar_Click
End Sub

Private Sub cmbAceptar_Click()
    Dim ficheroOri As String
    Dim ficheroDes As String
    ficheroOri = Me.Directorios.path & "\" & Me.Archivos.FileName
    ficheroDes = dir & maestro01.dato(0).Text & ".jpg"
    Call manejoArchivo.CopiarArchivo(ficheroOri, ficheroDes)
    maestro01.foto.Picture = LoadPicture(ficheroDes)
    Unload Me
End Sub

Private Sub Directorios_Change()
    Me.Archivos.path = Me.Directorios
    If Me.Archivos.ListCount > 0 Then
        Me.Archivos.ListIndex = 0
    End If
End Sub

Private Sub Discos_Change()
    Me.Directorios.path = Me.Discos
    If Me.Archivos.ListCount > 0 Then
        Me.Archivos.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()
    Set manejoArchivo = New Guga
    dir = App.path & "\Fotos\"
    If Me.Archivos.ListCount > 0 Then
        Me.Archivos.ListIndex = 0
    End If
End Sub
