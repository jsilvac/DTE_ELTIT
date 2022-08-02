VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form electro88 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manager de DTE"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cliente 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   3240
      Width           =   4935
   End
   Begin VB.TextBox tipo 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox folio 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2640
      Width           =   1815
   End
   Begin XPFrame.FrameXp frmenvios 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4471
      BackColor       =   16761024
      Caption         =   "Opciones"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.CommandButton Command4 
         Caption         =   "ENVIAR XML CORREO ELECTRONICO"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "VISTA DOCUMENTO RECIBIDO"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ENVIA PDF CORREO ELECTRONICO"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   1720
         BackColor       =   16744576
         Caption         =   "Correo Electronico"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   6975
         End
      End
   End
End
Attribute VB_Name = "electro88"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Command5_Click()

ARCHIVO = leerxmldterecibido(CONFI_EMPRESAFAE, Mid(tipo.text, 1, 2), FOLIO.text)
If ARCHIVO = "0" Then
    MsgBox "NO HAY RESPUESTA PARA ESTE ARCHIVO", vbInformation, "ATENCION"
    Exit Sub
End If
comi = Chr(34)
detalle = "<?xml-stylesheet type=" + comi + "text/xsl" + comi + " href=" + comi + "visualizador4.xsl" + comi + "?>" + ARCHIVO
cadena = detalle
For k = 1 To Len(detalle)
If Asc(Mid(detalle, k, 1)) > 128 And Mid(detalle, k, 1) <> "Ñ" Then
cadena = Replace(cadena, Mid(detalle, k, 1), "")
End If

Next k
detalle = cadena
detalle = Replace(detalle, "¥", "N")
detalle = Replace(detalle, "Ñ", "#209;")
detalle = Replace(detalle, "§", " ")
detalle = Replace(detalle, "º", " ")
detalle = Replace(detalle, "°", " ")
detalle = Replace(detalle, "&", "&amp;")
detalle = Replace(detalle, "ø", " ")
detalle = Replace(detalle, ",", ".")
detalle = Replace(detalle, "*", "x")
detalle = Replace(detalle, "´", "")
detalle = Replace(detalle, "Ç", "")
detalle = Replace(detalle, "ï", "")
detalle = Replace(detalle, "ï", "")


Close 20
        Open App.path + "\vistadte.xml" For Output As #20
        Print #20, detalle
        Close 20
     
'Call Shell("c:\archivos de programa\internet explorer\iexplore.exe " + App.path + "\vistadte.xml", vbMaximizedFocus)

Call Shell(Environ("programfiles") & "\internet explorer\iexplore.exe " + App.path + "\vistadte.xml", vbMaximizedFocus)


Unload Me


End Sub

Public Function leerxmldterecibido(loc, tipo, numero) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim VARIABLE As Variant
Set csql.ActiveConnection = db
csql.sql = "select xml from " + clientesistema + "fae" + loc + ".sv_dte" + loc + "_recibidos "
csql.sql = csql.sql & " where tipo='" + tipo + "' and numero='" + numero + "' "
csql.Execute
leerxmldterecibido = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
     VARIABLE = Mid(resultados(0), 1, 200000)
    
    If VARIABLE <> "" Then
   ' variable = Mid(resultados(0), 1, Len(resultados(0)))
    
        leerxmldterecibido = VARIABLE 'resultados(0)
    End If
End If
csql.Close
Set csql = Nothing
End Function

