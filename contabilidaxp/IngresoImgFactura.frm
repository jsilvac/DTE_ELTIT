VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form IngresoImgFactura 
   BackColor       =   &H00FF8080&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   10275
   ClientLeft      =   765
   ClientTop       =   570
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox Listado 
      BackColor       =   &H00FFC0C0&
      Height          =   4350
      Left            =   8520
      TabIndex        =   4
      Top             =   600
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   8520
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdEliminaImagen 
      Caption         =   "ELIMINAR"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12360
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton CmdGuardaImagen 
      Caption         =   "GUARDAR"
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton CmdNuevaImagen 
      Caption         =   "CARGAR IMAGEN"
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame frmOpciones 
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   10440
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "revision de imagenes"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ingreso de imagenes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Image imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   9735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PROVEEDOR"
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
      Index           =   3
      Left            =   8520
      TabIndex        =   8
      Top             =   6120
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUT PROVEEDOR :"
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
      Index           =   2
      Left            =   8520
      TabIndex        =   7
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NUMERO :"
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
      Index           =   1
      Left            =   8520
      TabIndex        =   6
      Top             =   5400
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO :"
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
      Index           =   0
      Left            =   8520
      TabIndex        =   5
      Top             =   5040
      Width           =   5655
   End
   Begin VB.Label kb 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   9960
      Width           =   975
   End
End
Attribute VB_Name = "IngresoImgFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ****************datos tamaño
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long
Option Explicit
Public cnn As ADODB.Connection
Public rst As ADODB.Recordset
Dim FOLIO, rut, Conex, CONSULTA, csql As String
Dim UserSQL, PassSQL, ServerSQL, BdSQL, TablaSQL, ImgNombre As String

' ××××××××××××××××DATOS FOTOS
Dim conn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim mystream As New ADODB.Stream
Dim Img As String
Dim KbImagen As Long
Dim loadStr, ImgTemporal, Imgtemporal2 As String

' ****************datos scan
Dim iX As Integer
Dim iY As Integer
Dim clrHashForeColor
Dim clrHashBackColor
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
      lpPictDesc As PictDesc, _
      riid As Guid, _
      ByVal fPictureOwnsHandle As Long, _
      ipic As IPicture _
    ) As Long
Private Sub CmdEliminaImagen_Click()
If Listado <> "" Then
    If MsgBox("Se va a eliminar la imagen de la factura :" & Mid(Listado, 3, 10) & " esta seguro ?", vbYesNo, "Atencion!!") = vbYes _
    Then CONSULTA = "DELETE FROM " & TablaSQL & " WHERE RUT =  '" & rut & "' AND NUMERO = '" & FOLIO & "'"
    Call cons(CONSULTA, 0)
    If Listado.SelCount = 1 Then
        Listado.Clear
    Else
        Listado.RemoveItem Listado.SelCount
    End If
Else
    MsgBox "Debe seleccionar una imagen del listado", vbExclamation, "Atencion"
End If
End Sub
Private Sub CmdGuardaImagen_Click()
If imagen.Picture = 0 Then Exit Sub
If kb > 62 Then
    RESIZE
Else
        ConexionImg (1)
        BuscaImg
Exit Sub
End If
End Sub
Private Sub CmdNuevaImagen_Click()
'ImgNombre = InputBox("Escriba el folio de la Imagen")
escaneo.Show vbModal
End Sub
Public Sub cons(CONSULTA, OPERACION)
 On Error GoTo error
Set cnn = Nothing
Set rst = Nothing
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
cnn.Open Conex
Set rst = Nothing
Set rst = New ADODB.Recordset
rst.Open CONSULTA, cnn, adLockOptimistic
If OPERACION = 1 Then       'CARGA DATOS AL LISTADO
        'Dim ITEM As String
    While rst.EOF = False
        With Listado
        'If rst(0) = 1 Then ITEM = "FACTURA"
            .AddItem rst(0) & vbTab & rst(1) & vbTab & rst(2)
            rst.MoveNext
        End With
    Wend
End If

Exit Sub
If OPERACION = 2 Then       'CARGA LA IMAGEN DESDE EL LISTADO
If rst.EOF = False Then
    Call ConexionImg(2)
End If
End If
Exit Sub
error:
    MsgBox Err.Description
Exit Sub
End Sub
Private Sub imagen_Click()
If imagen.Picture = 0 Then
    If MsgBox("No se puede guardar; debe tener una imagen cargada; desea cargarla ahora?", vbYesNo, "ATENCION") = vbYes Then
        CmdNuevaImagen_Click
        If MsgBox("Desea guardar la imagen?", vbYesNo, "ATENCION") = vbYes Then GoSub SubGuardaImagen
    Else
        Exit Sub
    End If
Else
    GoSub SubGuardaImagen
End If
Exit Sub
SubGuardaImagen:
With dialogo
    .FileName = ""
    .DialogTitle = " Guardar imagen "
    .Filter = "Archivos Jpg|*.jpg|"
    .ShowSave
                     '
    If dialogo.FileName = "" Then Exit Sub
    SavePicture imagen, .FileName
    MsgBox " archivo guardado correctamente ", vbInformation
End With
End Sub

Private Sub Listado_Click()
If Option1(0).Value = True Then
rut = ingreso02.dato9
FOLIO = Mid(Listado, 3, 10)
End If
If Option1(1).Value = True Then
rut = Mid(Label1(2), 16, 24)
FOLIO = Mid(Listado, 3, 10)
End If
Call ConexionImg(2)
End Sub
Private Sub Form_Load()
'ingreso02
With ingreso02
    Label1(0) = "TIPO :" & .dato1 & " " & .tipodocumento
    Label1(1) = "NUMERO :" & .dato2
    Label1(2) = "RUT PROVEEDOR :" & .dato9 & "-" & .dv
    Label1(3) = "PROVEEDOR :" & .nombreproveedor
End With
ImgTemporal = "c:\tmp.bmp"
Imgtemporal2 = "c:\tmp.jpg"
ServerSQL = servidor
UserSQL = usuario
PassSQL = password
BdSQL = clientesistema & "conta" & empresaactiva
TablaSQL = "facturasdecompra_imagen"

Conex = "driver={MySQL ODBC 3.51 Driver};server=" & ServerSQL & ";uid=" & _
UserSQL & ";pwd=" & PassSQL & ";database=" & BdSQL & ";connection=adUseClient"
BuscaImg
End Sub
Private Sub RESIZE()
kb = Empty
KbImagen = Empty
Dim retval As Long
MousePointer = vbHourglass 'CAMBIO EL PUNTERO A OCUPADO
loadStr = dialogo.FileName
SavePicture imagen.Picture, ImgTemporal
 On Error GoTo error
retval = DIWriteJpg(loadStr, 80, 1)
If retval = 1 Then  'correcto
   imagen.Picture = LoadPicture(loadStr)
   KbImagen = Mid(Str(FileLen(loadStr)), 1, Len(Str(FileLen(loadStr))) - 3)
Else                'ocurrió un error
   MsgBox "La conversión NO fue exitosa. Intentelo de nuevo."
   Exit Sub
End If
    Kill ImgTemporal
    MousePointer = vbNormal
    imagen = LoadPicture(dialogo.FileName)
    kb = KbImagen
    CmdGuardaImagen_Click
    Exit Sub
error:
    MousePointer = vbNormal
    MsgBox Err.Description
End Sub
Public Sub ConexionImg(OPERACION)
On Error GoTo error
mystream.Type = adTypeBinary
conn.ConnectionString = Conex
conn.CursorLocation = adUseClient
conn.Open
If OPERACION = 1 Then
    Rs.Open "Select * From " & BdSQL & "." & TablaSQL, conn, adOpenStatic, adLockOptimistic
    Rs.AddNew
    mystream.Open
    mystream.LoadFromFile dialogo.FileName
    With ingreso02
        Rs("TIPO") = .dato1
        Rs("NUMERO") = .dato2
        Rs("RUT") = .dato9
        Rs("FOTO") = mystream.Read
        Rs.Update
    End With
    mystream.Close
MsgBox "Se ha agregado la imagen satisfactoriamente", vbInformation, "Agregada"
Unload Me
End If
If OPERACION = 2 Then
Rs.Open "Select * From " & TablaSQL & " WHERE RUT = '" & rut & "' AND numero like '" & FOLIO & "'", conn, adOpenStatic, adLockOptimistic   '
On Local Error Resume Next
mystream.Type = adTypeBinary
mystream.Open
mystream.Write Rs.Fields("foto")
mystream.SaveToFile Imgtemporal2, adSaveCreateOverWrite
mystream.Close
imagen.Picture = LoadPicture(Imgtemporal2)
Shell ("rundll32.exe C:\WINDOWS\System32\shimgvw.dll,ImageView_Fullscreen " & Imgtemporal2), vbMaximizedFocus
KbImagen = ""
KbImagen = Mid(Str(FileLen(Imgtemporal2)), 1, Len(Str(FileLen(Imgtemporal2))) - 3)
kb = KbImagen & " KB"
End If
Rs.Close
conn.Close
Exit Sub
error:
MsgBox Err.Description
Exit Sub
End Sub
Sub BuscaImg()
Listado.Clear
CONSULTA = "SELECT tipo,numero,rut FROM " & BdSQL & "." & TablaSQL & " where rut = '" & ingreso02.dato9 & "'"
Call cons(CONSULTA, 1)
End Sub

Private Sub Option1_Click(Index As Integer)
'If Option1(0).Value = True Then 'ingresa
If Option1(1).Value = True Then 'revisa
    Listado.Clear
    CONSULTA = "SELECT tipo,numero,rut FROM "
    CONSULTA = CONSULTA + TablaSQL & " where rut = '" & Mid(Label1(2), 16, 24) & "'"
    Call cons(CONSULTA, 1)
End If

End Sub
