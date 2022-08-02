VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Descuento 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Agregar Pre-Venta"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2990
      BackColor       =   8421631
      Caption         =   "Descuento"
      CaptionEstilo3D =   1
      BackColor       =   8421631
      ColorBarraArriba=   12632319
      ColorBarraAbajo =   128
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
      Begin VB.TextBox txtPesos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2700
         MaxLength       =   9
         TabIndex        =   1
         Top             =   1140
         Width           =   1725
      End
      Begin VB.TextBox txtPorcen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2700
         MaxLength       =   5
         TabIndex        =   0
         Top             =   630
         Width           =   1725
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento ($)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   2400
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento (%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2400
      End
   End
End
Attribute VB_Name = "Descuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public MONTO As Double
    Private desc As Double
    
Private Sub Form_Activate()
 If PVentas.dato12.text > PVentas.lblSub Then
    txtPorcen.text = 0
    End If
    
    txtPorcen.text = PVentas.dato11.text
    txtPesos.text = PVentas.dato12.text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Or KeyCode = 38 Then
        Unload Me
    End If
    PVentas.dire = KeyCode
End Sub

Private Sub txtPorcen_GotFocus()
    Call selecciona(txtPorcen)
End Sub

Private Sub txtPesos_GotFocus()
    Call selecciona(txtPesos)
End Sub

Private Sub txtPorcen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtPorcen)
End Sub

Private Sub txtPesos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtPorcen)
End Sub

Private Sub txtPorcen_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumeroDecimal(txtPorcen, KeyAscii)
    If txtPorcen.text = "" Then
    txtPorcen.text = 0
    End If
    
    If KeyAscii = 13 Then
        If CDbl(txtPorcen.text) > 80 Then
            Call selecciona(txtPorcen)
        Else
            desc = Int((MONTO * CDbl(txtPorcen.text) / 100) + 0.5)
            txtPesos.text = Format(desc, "###,###,##0")
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub txtPesos_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        PVentas.dato11.text = txtPorcen.text
        PVentas.dato12.text = Format(txtPesos.text, "########0")
        Unload Me
        PVentas.dato12.SetFocus
    End If
End Sub

