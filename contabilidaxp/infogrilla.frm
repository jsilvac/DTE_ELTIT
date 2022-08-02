VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form grillainformes 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9735
   ClientLeft      =   645
   ClientTop       =   1110
   ClientWidth     =   14925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp CABEZA 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   18653
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   65535
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   16711680
      Begin VB.TextBox PIVOTE 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1095
         Left            =   240
         TabIndex        =   2
         Top             =   8640
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   1931
         BackColor       =   16777152
         Caption         =   "OPCIONES"
         CaptionEstilo3D =   1
         BackColor       =   16777152
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton Command5 
            Caption         =   "GENERA ARCHIVO DOCUMENTOS NO ELECTRONICOS"
            Height          =   255
            Left            =   7800
            TabIndex        =   18
            Top             =   720
            Width           =   5415
         End
         Begin VB.CommandButton Command3 
            Caption         =   "CARACTERIZACION LIBRO DE COMPRAS"
            Height          =   255
            Left            =   3600
            TabIndex        =   17
            Top             =   720
            Width           =   4095
         End
         Begin FlexCell.Grid exporta 
            Height          =   135
            Left            =   8040
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   238
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin CoolButtons.cool_Button command1 
            Height          =   495
            Left            =   4200
            TabIndex        =   3
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            Caption         =   "Imprimir"
         End
         Begin CoolButtons.cool_Button COMMAND2 
            Height          =   495
            Left            =   6120
            TabIndex        =   4
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            Caption         =   "Exportar Excel"
         End
         Begin CoolButtons.cool_Button command4 
            Height          =   495
            Left            =   10200
            TabIndex        =   5
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            Caption         =   "Salir"
         End
         Begin MSComctlLib.Slider TAMAÑOS 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
         End
         Begin CoolButtons.cool_Button cmd_xml 
            Height          =   495
            Left            =   12120
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "Genera XML SII"
         End
         Begin CoolButtons.cool_Button ascii 
            Height          =   495
            Left            =   2400
            TabIndex        =   11
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            Caption         =   "Genera Ascii SII"
         End
         Begin CoolButtons.cool_Button CmdExportaSeleccion 
            Height          =   495
            Left            =   8160
            TabIndex        =   13
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            Caption         =   "Exportar CSV"
         End
         Begin CoolButtons.cool_Button cmdcomprobante 
            Height          =   495
            Left            =   2400
            TabIndex        =   15
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            Caption         =   "Genera Comprobante"
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Doble Click en rut para ver en factura compra"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label registros 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   10800
            TabIndex        =   10
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label LETRA 
            Height          =   255
            Left            =   1920
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   8415
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14843
         Cols            =   5
         DefaultFontSize =   8.25
         DefaultRowHeight=   15
         Rows            =   30
      End
      Begin VB.Label titulofinal 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "grillainformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 

Private Sub ascii_Click()
Dim datos1 As String * 9
Dim datos2 As String * 6
Dim datos3 As String * 10
Dim DATOS4 As String * 1
Dim DATOS5 As String * 3
Dim DATOS6 As String * 10
Dim DATOS7 As String * 8
Dim DATOS8 As String * 9
Dim DATOS9 As String * 50
Dim DATOS10 As String * 13
Dim DATOS11 As String * 13
Dim DATOS12 As String * 13
Dim DATOS13 As String * 13
Dim DATOS14 As String * 13
Dim DATOS15 As String * 13
Dim DATOS16 As String
On Error GoTo no:



Dim cadena As String
Dim ARCHIVO As String
Dim tipo As String

If xmlcompra = True Then
Close 20
ARCHIVO = "C:\SII\C" + FECHALC + ".TXT"

datos1 = String(8, 32) + "1"
datos2 = String(5, 32) + "2"
datos3 = String(9, 32) + "3"
DATOS4 = "4"
DATOS5 = String(2, 32) + "5"
DATOS6 = String(9, 32) + "6"
DATOS7 = String(7, 32) + "7"
DATOS8 = String(8, 32) + "8"
DATOS9 = String(49, 32) + "9"
DATOS10 = String(11, 32) + "10"
DATOS11 = String(11, 32) + "11"
DATOS12 = String(11, 32) + "12"
DATOS13 = String(11, 32) + "13"
DATOS14 = String(11, 32) + "14"
DATOS15 = String(11, 32) + "15"
DATOS16 = "180"


Open ARCHIVO For Output As #20

cadena = datos1 + datos2 + datos3 + DATOS4 + DATOS5 + DATOS6 + DATOS7 + DATOS8 + DATOS9 + DATOS10 + DATOS11 + DATOS12 + DATOS13 + DATOS14 + DATOS15 + DATOS16
Print #20, cadena

For k = 1 To Grid1.Rows - 1
If Val(Grid1.Cell(k, 1).text) <> 0 Then
Rem RUT
pivote.MaxLength = 9
pivote.Alignment = 1
pivote.text = Replace(rutempresa, "-", "")
Call ESPACIOS(pivote)
datos1 = pivote.text
Rem FECHA
datos2 = FECHALC
Rem FOLIO
pivote.MaxLength = 10
pivote.Alignment = 1
pivote.text = Val(Grid1.Cell(k, 1).text)
Call ESPACIOS(pivote)

datos3 = pivote.text
DATOS4 = "C"
If Grid1.Cell(k, 2).text = "FA" Then tipo = "030"
If Grid1.Cell(k, 2).text = "FAE" Then tipo = "033"
If Grid1.Cell(k, 2).text = "NC" Then tipo = "060"
If Grid1.Cell(k, 2).text = "NCE" Then tipo = "061"
If Grid1.Cell(k, 2).text = "ND" Then tipo = "055"
If Grid1.Cell(k, 2).text = "NDE" Then tipo = "056"
DATOS5 = tipo
DATOS6 = Grid1.Cell(k, 3).text
DATOS7 = Format(Grid1.Cell(k, 4).text, "DDMMYYYY")
Rem RUT ASOCIADO

pivote.MaxLength = 9
pivote.Alignment = 1
pivote.text = Replace(Grid1.Cell(k, 5).text, "-", "")

Call ESPACIOS(pivote)
DATOS8 = pivote.text



Rem NOMBRE ASOCIADO

pivote.MaxLength = 50
pivote.Alignment = 1
pivote.text = Grid1.Cell(k, 6).text
Call ESPACIOS(pivote)

DATOS9 = pivote.text

Rem EXENTO

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Replace(Val(Grid1.Cell(k, 9).text) + Val(Grid1.Cell(k, 10).text), "-", "")
Call ESPACIOS(pivote)

DATOS10 = pivote.text

Rem NETO

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 7).text, "-", ""))
Call ESPACIOS(pivote)

DATOS11 = pivote.text

Rem IVA

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 8).text, "-", ""))
Call ESPACIOS(pivote)

DATOS12 = pivote.text

Rem TOTAL

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 12).text, "-", ""))
Call ESPACIOS(pivote)

DATOS13 = pivote.text

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 19).text, "-", ""))
Call ESPACIOS(pivote)

DATOS14 = pivote.text

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 20).text, "-", ""))
Call ESPACIOS(pivote)

DATOS15 = pivote.text
DATOS16 = ""
If Grid1.Cell(k, 21).text = "S" Then
DATOS16 = "x"
End If




cadena = datos1 + datos2 + datos3 + DATOS4 + DATOS5 + DATOS6 + DATOS7 + DATOS8 + DATOS9 + DATOS10 + DATOS11 + DATOS12 + DATOS13 + DATOS14 + DATOS15 + DATOS16
Print #20, cadena
End If


Next k
Close 20
Shell "NOTEPAD " + ARCHIVO


End If

If xmlventa = True Then
Close 20
ARCHIVO = "C:\SII\V" + FECHALV + ".TXT"
datos1 = String(8, 32) + "1"
datos2 = String(5, 32) + "2"
datos3 = String(9, 32) + "3"
DATOS4 = "4"
DATOS5 = String(2, 32) + "5"
DATOS6 = String(9, 32) + "6"
DATOS7 = String(7, 32) + "7"
DATOS8 = String(8, 32) + "8"
DATOS9 = String(49, 32) + "9"
DATOS10 = String(11, 32) + "10"
DATOS11 = String(11, 32) + "11"
DATOS12 = String(11, 32) + "12"
DATOS13 = String(11, 32) + "13"
DATOS14 = String(11, 32) + "14"
DATOS15 = String(11, 32) + "15"
DATOS16 = "180"


Open ARCHIVO For Output As #20

cadena = datos1 + datos2 + datos3 + DATOS4 + DATOS5 + DATOS6 + DATOS7 + DATOS8 + DATOS9 + DATOS10 + DATOS11 + DATOS12 + DATOS13
Print #20, cadena

For k = 1 To Grid1.Rows - 1
If Val(Grid1.Cell(k, 2).text) <> 0 Then
Rem RUT
pivote.MaxLength = 9
pivote.Alignment = 1
pivote.text = Replace(rutempresa, "-", "")
Call ESPACIOS(pivote)
datos1 = pivote.text
Rem FECHA
datos2 = FECHALV
Rem FOLIO
pivote.MaxLength = 10
pivote.Alignment = 1
pivote.text = "999"
Call ESPACIOS(pivote)

datos3 = pivote.text
DATOS4 = "V"
If Grid1.Cell(k, 1).text = "FA" Then tipo = "030"
If Grid1.Cell(k, 1).text = "NC" Then tipo = "060"
If Grid1.Cell(k, 1).text = "ND" Then tipo = "055"

DATOS5 = tipo
DATOS6 = Grid1.Cell(k, 2).text
DATOS7 = Format(Grid1.Cell(k, 3).text, "DDMMYYYY")
Rem RUT ASOCIADO

pivote.MaxLength = 9
pivote.Alignment = 1
pivote.text = Replace(Grid1.Cell(k, 4).text, "-", "")

Call ESPACIOS(pivote)
DATOS8 = pivote.text



Rem NOMBRE ASOCIADO

pivote.MaxLength = 50
pivote.Alignment = 1
pivote.text = Grid1.Cell(k, 5).text
Call ESPACIOS(pivote)

DATOS9 = pivote.text

Rem EXENTO

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Replace(Val(Grid1.Cell(k, 8).text), "-", "")
Call ESPACIOS(pivote)

DATOS10 = pivote.text

Rem NETO

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 6).text, "-", ""))
Call ESPACIOS(pivote)

DATOS11 = pivote.text

Rem IVA

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 7).text, "-", ""))
Call ESPACIOS(pivote)

DATOS12 = pivote.text

Rem TOTAL

pivote.MaxLength = 13
pivote.Alignment = 1
pivote.text = Val(Replace(Grid1.Cell(k, 9).text, "-", ""))
Call ESPACIOS(pivote)

DATOS13 = pivote.text




cadena = datos1 + datos2 + datos3 + DATOS4 + DATOS5 + DATOS6 + DATOS7 + DATOS8 + DATOS9 + DATOS10 + DATOS11 + DATOS12 + DATOS13
Print #20, cadena
End If

Next k
Close 20
Shell "NOTEPAD " + ARCHIVO



End If

Exit Sub
no:
MsgBox "ERROR NO ESTA LA RUTA C:\SII"
End Sub

Private Sub cmd_xml_Click()
Dim sinmovimientos As Double

Dim totalcantsuper As Double
Dim totalexentosuper As Double
Dim totalivasuper As Double
Dim totalnetosuper As Double
Dim totalmontosuper As Double
Dim ivanorecusuper As Double
Dim totalrefrescosuper As Double
Dim totallicoressuper As Double
Dim totalvinossuper As Double
Dim totalcervezasuper As Double
Dim totalharinasuper As Double
Dim totalcarnesuper As Double
Dim TotalDieselSuper As Double
Dim proporcionsuper As Double
Dim ivausocomunsuper As Double
Dim totalnoazucarsuper As Double
Dim TotalDieselRecuSuper As Double


'ACTIVO FIJO NORMAL

Dim totalcantactivo As Double
Dim totalexentoactivo As Double
Dim totalivaactivo As Double
Dim totalnetoactivo As Double
Dim totalmontoactivo As Double
Dim ivanorecuactivo As Double
Dim totalrefrescoactivo As Double
Dim totallicoresactivo As Double
Dim totalvinosactivo As Double
Dim totalcervezaactivo As Double
Dim totalharinaactivo As Double
Dim totalcarneactivo As Double
Dim TotalDieselActivo As Double
Dim proporcionactivo As Double
Dim ivausocomunactivo As Double
Dim totalnoazucaractivo As Double
Dim TotalDieselRecuActivo As Double

' ACTIVO FIJO NORMAL


Dim i As Double
If nogenerar = True Then
    MsgBox "LIBRO CON ANOMALIAS POR FAVOR REGULARIZAR, NO SE GENERARA EL ARCHIVO", vbCritical, "ATENCION"
    nogenerar = False
    Exit Sub
End If

rutempresa = leerdatos(conta, "maestroempresas", "rut", "codigoempresa='" + empresaactiva + "' ")

Dim nombrearchivo As String
'On Error GoTo erro:
If xmllibrodiario = True Then

Call generacabezalibrodiario("DIARIO_RES_2012_12", rutempresa, "2012-12", "2012-12", "", "")
End If

 
If xmlcompra = True Then
        If f3328 = True Then
        
        nombrearchivo = "LC3328" + Format(fechasistema, "mm") + Format(fechasistema, "yyyy") + ".xml"
        
        Else
        
        nombrearchivo = "LC" + Format(fechasistema, "mm") + Format(fechasistema, "yyyy") + ".xml"
        End If
        
        Call generacaratula(rutempresa, rut_representante, Format(fechasistema, "yyyy-mm"), fecharesolucion, numeroresolucion, "COMPRA", "MENSUAL", "TOTAL", "", "1")
        sinmovimientos = 0
        ' para leer si tiene facturas de supermercado
        
            totalcantsuper = 0
            totalexentosuper = 0
            totalnetosuper = 0
            totalivasuper = 0
            totalmontosuper = 0
            ivanorecusuper = 0
            totalrefrescosuper = 0
            totallicoressuper = 0
            totalvinossuper = 0
            totalcervezasuper = 0
            totalharinasuper = 0
            totalcarnesuper = 0
            TotalDieselSuper = 0
            proporcionsuper = 0
            ivausocomunsuper = 0
            totalnoazucarsuper = 0
            TotalDieselSuper = 0
            TotalDieselRecuSuper = 0
            
            totalexentoactivo = 0
            totalnetoactivo = 0
            totalivaactivo = 0
            totalmontoactivo = 0
            ivanorecuactivo = 0
            totalrefrescoactivo = 0
            totallicoresactivo = 0
            totalvinosactivo = 0
            totalcervezaactivo = 0
            totalharinaactivo = 0
            totalcarneactivo = 0
            TotalDieselActivo = 0
            proporcionactivo = 0
            ivausocomunactivo = 0
            totalnoazucaractivo = 0
            TotalDieselActivo = 0
            TotalDieselRecuActivo = 0
            
            
        For i = 1 To Grid1.Rows - 1
            If Grid1.Cell(i, 6).text = "FACTURAS SUPERMERCADO " And Grid1.Cell(i, 5).text <> "0" Then
                totalcantsuper = Grid1.Cell(i, 5).text
                totalexentosuper = CDbl(Grid1.Cell(i, 9).text)
                totalnetosuper = Grid1.Cell(i, 7).text
                totalivasuper = Grid1.Cell(i, 8).text
                totalmontosuper = Grid1.Cell(i, 12).text
                ivanorecusuper = "0"
                totalrefrescosuper = Grid1.Cell(i, 13).text
                totallicoressuper = Grid1.Cell(i, 14).text
                totalvinossuper = Grid1.Cell(i, 15).text
                totalcervezasuper = Grid1.Cell(i, 16).text
                totalharinasuper = Grid1.Cell(i, 17).text
                totalcarnesuper = Grid1.Cell(i, 18).text
                TotalDieselSuper = Grid1.Cell(i, 10).text
                proporcionsuper = Grid1.Cell(i, 20).text
                ivausocomunsuper = Grid1.Cell(i, 21).text
                totalnoazucarsuper = Grid1.Cell(i, 19).text
                TotalDieselRecuSuper = Grid1.Cell(i, 23).text
                Exit For
            End If
        Next i
        
         For i = 1 To Grid1.Rows - 1
            If Grid1.Cell(i, 6).text = "FACTURAS ACTIVO FIJO ELECTRONICAS" And Grid1.Cell(i, 5).text <> "0" Then
                totalcantsuper = totalcantsuper + Grid1.Cell(i, 5).text
                totalexentosuper = totalexentosuper + CDbl(Grid1.Cell(i, 9).text)
                totalnetosuper = totalnetosuper + Grid1.Cell(i, 7).text
                totalivasuper = totalivasuper + Grid1.Cell(i, 8).text
                totalmontosuper = totalmontosuper + Grid1.Cell(i, 12).text
                ivanorecusuper = "0"
                totalrefrescosuper = totalrefrescosuper + Grid1.Cell(i, 13).text
                totallicoressuper = totallicoressuper + Grid1.Cell(i, 14).text
                totalvinossuper = totalvinossuper + Grid1.Cell(i, 15).text
                totalcervezasuper = totalcervezasuper + Grid1.Cell(i, 16).text
                totalharinasuper = totalharinasuper + Grid1.Cell(i, 17).text
                totalcarnesuper = totalcarnesuper + Grid1.Cell(i, 18).text
                TotalDieselSuper = TotalDieselSuper + Grid1.Cell(i, 10).text
                proporcionsuper = proporcionsuper + Grid1.Cell(i, 20).text
                ivausocomunsuper = ivausocomunsuper + Grid1.Cell(i, 21).text
                totalnoazucarsuper = totalnoazucarsuper + Grid1.Cell(i, 19).text
                TotalDieselRecuSuper = TotalDieselSuper + Grid1.Cell(i, 23).text
                Exit For
            End If
        Next i
        
     For i = 1 To Grid1.Rows - 1
        
        If Grid1.Cell(i, 6).text = "FACTURAS ACTIVO FIJO NORMALES" And Grid1.Cell(i, 5).text <> "0" Then
                totalcantactivo = Grid1.Cell(i, 5).text
                totalexentoactivo = CDbl(Grid1.Cell(i, 9).text)
                totalnetoactivo = Grid1.Cell(i, 7).text
                totalivaactivo = Grid1.Cell(i, 8).text
                totalmontoactivo = Grid1.Cell(i, 12).text
                ivanorecuactivo = "0"
                totalrefrescoactivo = Grid1.Cell(i, 13).text
                totallicoresactivo = Grid1.Cell(i, 14).text
                totalvinosactivo = Grid1.Cell(i, 15).text
                totalcervezaactivo = Grid1.Cell(i, 16).text
                totalharinaactivo = Grid1.Cell(i, 17).text
                totalcarneactivo = Grid1.Cell(i, 18).text
                TotalDieselActivo = Grid1.Cell(i, 10).text
                proporcionactivo = Grid1.Cell(i, 20).text
                ivausocomunactivo = Grid1.Cell(i, 21).text
                totalnoazucaractivo = Grid1.Cell(i, 19).text
                TotalDieselRecuActivo = Grid1.Cell(i, 23).text
                Exit For
            End If
        Next i
        
        For k = 1 To Grid1.Rows - 1
        
            If Grid1.Cell(k, 6).text = "FACTURAS " And (Grid1.Cell(k, 5).text <> "0" Or totalmontoactivo > 0) Then
            
            
            
                
                If totalmontoactivo > 0 Then
                    If Grid1.Cell(k, 5).text <> "0" Then
                        totalcantactivo = totalcantactivo + Grid1.Cell(k, 5).text
                        totalexentoactivo = totalexentoactivo + CDbl(Grid1.Cell(k, 9).text)
                        totalnetoactivo = totalnetoactivo + Grid1.Cell(k, 7).text
                        totalivaactivo = totalivaactivo + Grid1.Cell(k, 8).text
                        totalmontoactivo = totalmontoactivo + Grid1.Cell(k, 12).text
                        ivanorecuactivo = "0"
                        totalrefrescoactivo = totalrefrescoactivo + Grid1.Cell(k, 13).text
                        totallicoresactivo = totallicoresactivo + Grid1.Cell(k, 14).text
                        totalvinosactivo = totalvinosactivo + Grid1.Cell(k, 15).text
                        totalcervezaactivo = totalcervezaactivo + Grid1.Cell(k, 16).text
                        totalharinaactivo = totalharinaactivo + Grid1.Cell(k, 17).text
                        totalcarneactivo = totalcarneactivo + Grid1.Cell(k, 18).text
                        TotalDieselActivo = TotalDieselActivo + Grid1.Cell(k, 10).text
                        proporcionactivo = proporcionactivo + Grid1.Cell(k, 20).text
                        ivausocomunactivo = ivausocomunactivo + Grid1.Cell(k, 21).text
                        totalnoazucaractivo = totalnoazucaractivo + Grid1.Cell(k, 19).text
                        TotalDieselRecuActivo = TotalDieselRecuActivo + Grid1.Cell(k, 23).text
                        
                        
                         Call GENERATOTALTIPO("30", totalcantactivo, totalexentoactivo, totalnetoactivo, totalivaactivo, totalmontoactivo, "0", totalrefrescoactivo, totallicoresactivo, totalvinosactivo, totalcervezaactivo, totalharinaactivo, totalcarneactivo, TotalDieselActivo, proporcionactivo, ivausocomunactivo, totalnoazucaractivo, TotalDieselRecuActivo)
                    Else
                        Call GENERATOTALTIPO("30", totalcantactivo, totalexentoactivo, totalnetoactivo, totalivaactivo, totalmontoactivo, "0", totalrefrescoactivo, totallicoresactivo, totalvinosactivo, totalcervezaactivo, totalharinaactivo, totalcarneactivo, TotalDieselActivo, proporcionactivo, ivausocomunactivo, totalnoazucaractivo, TotalDieselRecuActivo)
                    End If
                Else
                     Call GENERATOTALTIPO("30", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                End If
                sinmovimientos = sinmovimientos + 1
            
            End If
            
            If Grid1.Cell(k, 6).text = "NOTAS DE DEBITO" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("55", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 6).text = "NOTAS DE CREDITO" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("60", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 6).text = "FACTURAS ELECTRONICAS" And (Grid1.Cell(k, 5).text <> "0" Or totalmontosuper > 0) Then
            
                If totalmontosuper > 0 Then
                    If Grid1.Cell(k, 5).text <> "0" Then
                        totalcantsuper = totalcantsuper + Grid1.Cell(k, 5).text
                        totalexentosuper = totalexentosuper + CDbl(Grid1.Cell(k, 9).text)
                        totalnetosuper = totalnetosuper + Grid1.Cell(k, 7).text
                        totalivasuper = totalivasuper + Grid1.Cell(k, 8).text
                        totalmontosuper = totalmontosuper + Grid1.Cell(k, 12).text
                        ivanorecusuper = "0"
                        totalrefrescosuper = totalrefrescosuper + Grid1.Cell(k, 13).text
                        totallicoressuper = totallicoressuper + Grid1.Cell(k, 14).text
                        totalvinossuper = totalvinossuper + Grid1.Cell(k, 15).text
                        totalcervezasuper = totalcervezasuper + Grid1.Cell(k, 16).text
                        totalharinasuper = totalharinasuper + Grid1.Cell(k, 17).text
                        totalcarnesuper = totalcarnesuper + Grid1.Cell(k, 18).text
                        TotalDieselSuper = TotalDieselSuper + Grid1.Cell(k, 10).text
                        proporcionsuper = proporcionsuper + Grid1.Cell(k, 20).text
                        ivausocomunsuper = ivausocomunsuper + Grid1.Cell(k, 21).text
                        totalnoazucarsuper = totalnoazucarsuper + Grid1.Cell(k, 19).text
                        TotalDieselRecuSuper = TotalDieselRecuSuper + Grid1.Cell(k, 23).text
                        
                        
                         Call GENERATOTALTIPO("33", totalcantsuper, totalexentosuper, totalnetosuper, totalivasuper, totalmontosuper, "0", totalrefrescosuper, totallicoressuper, totalvinossuper, totalcervezasuper, totalharinasuper, totalcarnesuper, TotalDieselSuper, proporcionsuper, ivausocomunsuper, totalnoazucarsuper, TotalDieselRecuSuper)
                    Else
                         Call GENERATOTALTIPO("33", totalcantsuper, totalexentosuper, totalnetosuper, totalivasuper, totalmontosuper, "0", totalrefrescosuper, totallicoressuper, totalvinossuper, totalcervezasuper, totalharinasuper, totalcarnesuper, TotalDieselSuper, proporcionsuper, ivausocomunsuper, totalnoazucarsuper, TotalDieselRecuSuper)
               
                    End If
                Else
                    Call GENERATOTALTIPO("33", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                End If
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 6).text = "NOTAS DE DEBITO ELECTRONICAS" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("56", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
                
            End If
            If Grid1.Cell(k, 6).text = "NOTAS DE CREDITO ELECTRONICAS" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("61", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
            
            
            If Grid1.Cell(k, 6).text = "FACTURAS COMPRAS PROPIAS" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("46", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
            If Grid1.Cell(k, 6).text = "IMPORTACIONES." And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("914", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 6).text = "EXENTAS NORMALES" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("32", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
            If Grid1.Cell(k, 6).text = "EXENTAS ELECTRONICAS" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("34", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 6).text = "LIQUIDACION-FACTURAS ELECTRONICAS" And Grid1.Cell(k, 5).text <> "0" Then
                Call GENERATOTALTIPO("43", Grid1.Cell(k, 5).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
                sinmovimientos = sinmovimientos + 1
            End If
        Next k
        If sinmovimientos > 0 Then
        lce = lce + "</ResumenPeriodo>"
        Else
            lce = Replace(lce, "<ResumenPeriodo>", "")
        End If
        
        For k = 1 To Grid1.Rows - 1
        If Val(Grid1.Cell(k, 1).text) <> 0 Then
            ACTIVO = "0"
            If Grid1.Cell(k, 13).text = "S" Then
                ACTIVO = Grid1.Cell(k, 7).text
            End If
        
        Call generalc(Grid1.Cell(k, 2).text, Grid1.Cell(k, 3).text, Grid1.Cell(k, 4).text, Grid1.Cell(k, 5).text, Grid1.Cell(k, 6).text, CDbl(Grid1.Cell(k, 9).text), Grid1.Cell(k, 7).text, Grid1.Cell(k, 8).text, ACTIVO, "0", Grid1.Cell(k, 12).text, "0", Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 20).text, Grid1.Cell(k, 21).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 23).text)
        registros.Caption = k & " / " & Grid1.Rows - 1
        registros.Refresh
        
        End If
        Next k
End If
        

If xmlventa = True Then
        If f3327 = True Then
        
        nombrearchivo = "LV3327" + Format(fechasistema, "mm") + Format(fechasistema, "yyyy") + ".xml"
        
        Else
        
        nombrearchivo = "LV" + Format(fechasistema, "mm") + Format(fechasistema, "yyyy") + ".xml"
        
        End If
        
        Call generacaratula(rutempresa, rut_enviasii, Format(fechasistema, "yyyy-mm"), fecharesolucion, numeroresolucion, "VENTA", "MENSUAL", "TOTAL", "", "0")
       sinmovimientos = 0
        For k = 1 To Grid1.Rows - 1
            If Grid1.Cell(k, 5).text = "FACTURAS " And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("30", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 5).text = "NOTAS DE DEBITO" And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("55", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If (Grid1.Cell(k, 5).text = "NOTAS DE CREDITO FACTURAS" Or Grid1.Cell(k, 5).text = "NOTAS DE CREDITO FACTURAS") And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("60", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 5).text = "FACTURAS ELECTRONICAS" And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("33", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 5).text = "NOTAS DE DEBITO ELECTRONICAS" And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("56", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 5).text = "N.CREDITO ELECTRONICAS FACTURA" And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("61", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 5).text = "FACTURAS EXENTAS" And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("32", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 5).text = "FACTURAS EXENTAS ELECTRONICAS" And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("34", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
           If Grid1.Cell(k, 5).text = "FACTURAS EXPORTACION" And Grid1.Cell(k, 4).text <> "0" Then
                Call GENERATOTALTIPOV("101", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, "0", "0", "0")
                sinmovimientos = sinmovimientos + 1
            End If
            
            If Grid1.Cell(k, 5).text = "LIQUIDACION-FACTURAS ELECTRONICAS" And Grid1.Cell(k, 4).text <> "0" Then
                 Call GENERATOTALTIPOV("43", Grid1.Cell(k, 4).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text, Grid1.Cell(k, 19).text, Grid1.Cell(k, 20).text)
                sinmovimientos = sinmovimientos + 1
            End If
        
        Next k
                If D35_cantidad <> 0 Then
                    Call GENERATOTALTIPOV("35", D35_cantidad, "0", D35_neto, D35_iva, D35_total, "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0")
                     sinmovimientos = sinmovimientos + 1
                End If
                
                If D38_cantidad <> 0 Then
                    Call GENERATOTALTIPOV("38", D38_cantidad, D38_neto, 0, 0, 0, "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0")
                     sinmovimientos = sinmovimientos + 1
                End If
                
                If D48_cantidad <> 0 Then
                    Call GENERATOTALTIPOV("48", D48_cantidad, "0", D48_neto, D48_iva, D48_total, "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0")
                     sinmovimientos = sinmovimientos + 1
                End If
                
        If sinmovimientos > 0 Then
            lce = lce + "</ResumenPeriodo>"
        Else
            lce = Replace(lce, "<ResumenPeriodo>", "")
        End If
        
        
        
        For k = 1 To Grid1.Rows - 1
        If Val(Grid1.Cell(k, 2).text) <> 0 Then
        
        If Grid1.Cell(k, 1).text <> "FAE" And Grid1.Cell(k, 1).text <> "NCE" And Grid1.Cell(k, 1).text <> "NDE" And Grid1.Cell(k, 1).text <> "LFE" Then
        Call generalv(Grid1.Cell(k, 1).text, Grid1.Cell(k, 2).text, Grid1.Cell(k, 3).text, Grid1.Cell(k, 4).text, Grid1.Cell(k, 5).text, Grid1.Cell(k, 8).text, Grid1.Cell(k, 6).text, Grid1.Cell(k, 7).text, "0", "0", Grid1.Cell(k, 9).text, "0", Grid1.Cell(k, 0).text, Grid1.Cell(k, 11).text, Grid1.Cell(k, 12).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 14).text, Grid1.Cell(k, 15).text, Grid1.Cell(k, 16).text, Grid1.Cell(k, 17).text)
        End If
        registros.Caption = k & " / " & Grid1.Rows - 1
        registros.Refresh
        End If
        Next k
End If
        If f3328 = True Then
        lce = lce + "<TmstFirma>" & Format(Date, "yyyy-mm-dd") & "T" & Time & "</TmstFirma>"
        lce = lce + "</EnvioLibro>"
        lce = lce + "</LibroCompraVenta>"
End If

        Rem If f3328 = False Then
        Call xml.LoadXml(lce)
        Rem End If
        
        If f3328 = True Then
         raiz = "C:\FORM33\"
        
        Else
        
         raiz = "u:\fae\" + CONFI_EMPRESAFAE + "\libros\"
        End If
        Rem If f3328 = False Then
        Call xml.SaveXml(raiz + nombrearchivo)
        Rem Else
        If f3328 = True Then
        Close 21
        Open raiz + "final_" + nombrearchivo For Output As #21
        
        Close 20
        Open raiz + nombrearchivo For Input As #20
        While EOF(20) = False
        Input #20, ll
        If ll <> "" Then
        Print #21, ll
        End If
        
        Wend
        
            
        Close 20: Close 21
        Kill (raiz + nombrearchivo)
        End If
        
If MsgBox("DESEA PRE VISUALIZAR INFORME ", vbYesNo, "ATENCION") = vbYes Then
    If f3328 = True Then
    Shell "notepad " + raiz + "final_" + nombrearchivo
Else
    Shell "notepad " + raiz + nombrearchivo

End If

End If

xmlcompra = False
xmlventa = False
Exit Sub
erro:
MsgBox error$
End Sub

Private Sub cmdcomprobante_Click()
    Dim k As Double
    Dim numero As String
    numero = ultimo("FM")
    
    For k = 1 To Grid1.Rows - 1
        Call generacomprobantecontable(numero, k, Format(fechasistema, "yyyy-mm-dd"), Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text, Grid1.Cell(k, 1).text)
    Next k
End Sub

Private Sub CmdExportaSeleccion_Click()
Call ExportaSeleccion
End Sub

Private Sub Command1_Click()
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
    Grid1.PageSetup.PaperWidth = 21.59
    Grid1.PageSetup.PaperHeight = 27.94
    
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar01" Then Call imprime_balancetributario(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar02" Then Call imprime_balanceanalitico(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar03" Then Call imprime_mayoranalitico(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10), titulofinal.Caption)
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar04" Then Call imprime_librodiario(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10), titulofinal.Caption)
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar05" Then Call imprime_librocompras(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "prove0010" Then Call imprime_FacturasDigitadas(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar06" Then Call imprime_librohonorarios(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "publi0004" Then Call imprime_publicidad(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar44" Then Call imprime_libroventas(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar07" Then Call imprime_libroboletas(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 7) = "banco02" Then Call imprime_cartolamayor("N", Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 12) = "CARTOLAMAYOR" Then Call imprime_cartolamayor("N", Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 13) = "CARTOLACTACTE" Then Call imprime_cartolaCTACTE("N", Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 11) = "informa01_1" Then Call imprime_NORMALES("PLAN DE CUENTAS")
    If Mid(grillainformes.Tag, 1, 11) = "informa01_2" Then Call imprime_NORMALES("CUENTAS CORRIENTES")
    If Mid(grillainformes.Tag, 1, 11) = "informa01_3" Then Call imprime_NORMALES("CENTROS DE COSTO")
    If Mid(grillainformes.Tag, 1, 8) = "infoge01" Then Call imprime_estadoresultado(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 8) = "infoge02" Then Call imprime_facturasporpagar(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 8) = "infoge03" Then Call imprime_honorariosporpagar(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 8) = "infoge04" Then Call imprime_ventasporpagar(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "control03" Then Call imprime_buscapormonto(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "control01" Then Call imprime_descuadrados(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "control04" Then Call imprime_buscacuentaseliminadas(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "INFOHARINA" Then Call imprime_harina(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "INFOCARNE" Then Call imprime_CARNE(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 7) = "PRESU04" Then Call imprime_estadoresultado(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    
    Unload Me
End Sub

Private Sub COMMAND2_Click()
Grid1.ExportToExcel (""), True



End Sub

Private Sub Command3_Click()
Call ArchivoComplementos
End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub cool_Button1_Click()

End Sub

Private Sub Command5_Click()
ArchivoDocManuales
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_Load()


Call CENTRAR(Me)
LETRA.Caption = Grid1.DefaultFont.Name
TAMAÑOS.Min = Grid1.DefaultFont.Size - 5
TAMAÑOS.Max = Grid1.DefaultFont.Size + 10
TAMAÑOS.Value = Grid1.DefaultFont.Size
Grid1.AutoRedraw = False
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Locked = True


Next k
Grid1.AutoRedraw = True
If xmlcompra = True Then
cmd_xml.Visible = True
ascii.Visible = True

End If
If xmlventa = True Then
cmd_xml.Visible = True
ascii.Visible = True

End If
If xmllibrodiario = True Then
cmd_xml.Visible = True

End If



End Sub

Sub imprime_balancetributario(tipo, FOLIO)
Dim titulo As String
Dim subtitulo As String

subtitulo = Mid(CABEZA.Caption, 20, ((InStr(CABEZA.Caption, "empresa") - 6) - 20))
titulo = "BALANCE TRIBUTARIO"
Call cabezas(titulo, tipo, FOLIO, subtitulo)
Grid1.DefaultFont.Size = 7

If Grid1.Column(1).Width <> "0" Then
    Grid1.PageSetup.Orientation = cellLandscape
'    For k = 1 To 11 - 1
'        Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
'    Next k
Else
    Grid1.PageSetup.Orientation = cellPortrait
    Grid1.Column(1).Width = 0
    For k = 2 To 11 - 1
        Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
    Next k
End If

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PrintPreview 120

End Sub
Sub imprime_balanceanalitico(tipo, FOLIO)
Dim titulo As String
titulo = "BALANCE ANALITICO"
Call cabezas(titulo, tipo, FOLIO, "")
'grid1.DefaultFont.Size = 6
Grid1.Column(2).Width = 20 * 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub
Sub imprime_mayoranalitico(tipo, FOLIO, TITULOalfinal)
Dim titulo As String
titulo = "MAYOR ANALITICO"
titulo = TITULOalfinal
Grid1.DefaultFont.Size = 8
For k = 1 To 10 - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 8
Next k
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub
Sub imprime_cartolamayor(tipo, FOLIO)

Dim titulo As String
titulo = "CARTOLA DEL MAYOR"
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
For k = 1 To 15 - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
Next k
Grid1.Column(6).Width = 200
Grid1.Column(5).Width = 0

Grid1.Column(14).Width = 0
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_cartolaCTACTE(tipo, FOLIO)
Dim titulo As String
titulo = "CARTOLA DEL CUENTAS CORRIENTES"
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
For k = 1 To 10 - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
Next k

Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub



Sub imprime_librodiario(tipo, FOLIO, titulocabeza)
Dim titulo As String
titulo = "LIBRO DIARIO"
titulo = titulocabeza
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0

For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 8
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub

Sub imprime_librocompras(tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE COMPRAS " + auxiliar05.COMBOMES.text + " de " + auxiliar05.COMBOAÑO.text


Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape


If tipo <> "N" Then
Grid1.PageSetup.Orientation = cellPortrait

Grid1.Cols = 13
End If
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_FacturasDigitadas(tipo, FOLIO)
Dim titulo As String
titulo = "DOCUMENTOS INGRESADOS POR " & USUARIOSISTEMA & " EL " + prove0010.DESDE1.text + "-" + prove0010.DESDE2.text & "-" & prove0010.DESDE3.text


Call cabezas(titulo, tipo, FOLIO, "ENTRE  " + prove0010.txtHoraDesde.text + " Y " + prove0010.txtHoraHasta.text & " HRS.")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape


If tipo <> "N" Then
Grid1.PageSetup.Orientation = cellPortrait

Grid1.Cols = 13
End If
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_publicidad(tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE COMPRAS a PROVEEDORES PERIODO " + publi0004.desdefecha.Caption + " HASTA " + publi0004.hastafecha.Caption

tipo = "N"
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub

Sub imprime_libroventas(tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE VENTAS " + auxiliar44.Combocrcc.text + " " + auxiliar44.COMBOMES.text + " de " + auxiliar44.COMBOAÑO.text

Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_libroboletas(tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE BOLETAS " + auxiliar07.Combocrcc.text + " " + auxiliar07.COMBOMES.text + " de " + auxiliar07.COMBOAÑO.text

Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_facturasporpagar(tipo, FOLIO)
Dim titulo As String
titulo = "FACTURAS POR PAGAR"
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_estadoresultado(tipo, FOLIO)
Dim titulo As String
titulo = "ESTADO DE RESULTADO "
Call cabezas(titulo, "N", FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

 For i = 1 To Grid1.PageSetup.PaperSizes.Count
            If UCase(Grid1.PageSetup.PaperSizes.item(i).PaperName) = "OFICIO" Then
                Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.item(i).Kind
                Exit For
            End If
        Next i
        


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 0
Grid1.PageSetup.TopMargin = 0.5
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k
Grid1.SelectionMode = cellSelectionFree


Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_honorariosporpagar(tipo, FOLIO)
Dim titulo As String
titulo = "HONORARIOS POR PAGAR"
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_ventasporpagar(tipo, FOLIO)
Dim titulo As String
titulo = "VENTAS POR PAGAR"
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_buscapormonto(tipo, FOLIO)
Dim titulo As String
titulo = "BUSCA POR MONTO "
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_buscacuentaseliminadas(tipo, FOLIO)
Dim titulo As String
titulo = "LISTADO DE CUENTAS ELIMINADAS "
Call cabezas(titulo, tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellPortrait

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_librohonorarios(tipo, FOLIO)
    Dim titulo As String
    titulo = "LIBRO DE HONORARIOS " + auxiliar06.COMBOMES.text + " de " + auxiliar06.COMBOAÑO.text
    Call cabezas(titulo, tipo, FOLIO, "")
    Grid1.DefaultFont.Size = 6
    Grid1.PageSetup.Orientation = cellPortrait
    Grid1.PageSetup.PrintFixedRow = True
    Grid1.PageSetup.BottomMargin = 2
    Grid1.PageSetup.TopMargin = 1
    Grid1.PageSetup.LeftMargin = 0.5
    Grid1.PageSetup.RightMargin = 0
    
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
    Next k
    Grid1.DisablePrintButton = False
    Grid1.Refresh
    Grid1.PrintPreview 120
   
End Sub


Sub cabezas(titulo, tipo, FOLIO, subtitulo)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    If subtitulo <> "" Then
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = UCase(subtitulo)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 9
        objReportTitle.Font.Bold = True
        objReportTitle.PrintOnAllPages = True
        Grid1.ReportTitles.Add objReportTitle
    End If
    
    'Report Title 1
    
    If tipo = "N" Then
        For k = 1 To 5
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload informa04
nogenerar = False
End Sub

Private Sub Grid1_DblClick()
Dim dia As String
If grillainformes.Tag = "auxiliar05N" Then
Rem electro88.Text1.text = impresion.Cell(impresion.ActiveCell.row, 0).text
If Grid1.ActiveCell.col = 5 Then
    Unload ingreso02
    If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "FAE" Then ingreso02.dato1.text = 4
    If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NCE" Then ingreso02.dato1.text = 6
    If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NDE" Then ingreso02.dato1.text = 5
    If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "FC" Then ingreso02.dato1.text = 1
    ingreso02.dato2.text = Grid1.Cell(Grid1.ActiveCell.row, 3).text
    ingreso02.dato9.text = Replace(Grid1.Cell(Grid1.ActiveCell.row, 5).text, "-", "")
    ingreso02.DV.Caption = Right(Replace(Grid1.Cell(Grid1.ActiveCell.row, 5).text, "-", ""), 1)
    
    ingreso02.Show
    Call ingreso02.CargarDeLibro
        'MsgBox "carga ingreso libro compra"
    Else
    If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "FAE" Or Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NCE" Or Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NCD" Then
        If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "FAE" Then electro88.tipo.text = "33"
        If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NCE" Then electro88.tipo.text = "61"
        If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "NDE" Then electro88.tipo.text = "56"
            electro88.cliente.text = Replace(Grid1.Cell(Grid1.ActiveCell.row, 5).text, "-", "")
        electro88.FOLIO.text = Grid1.Cell(Grid1.ActiveCell.row, 3).text
        
        electro88.Show vbModal
    End If
End If

End If
If Mid(CABEZA.Caption, 1, 7) = "CARTOLA" Then
da0 = Grid1.Cell(Grid1.ActiveCell.row, 2).text
da1 = Grid1.Cell(Grid1.ActiveCell.row, 3).text
da2 = Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "dd")
da3 = Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "mm")
da4 = Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "yyyy")
muestracomprobantes.Show vbModal
End If
If Mid(grillainformes.Tag, 1, 7) = "PRESU04" Then

End If
If Mid(CABEZA.Caption, 1, 6) = "ESTADO" Or Mid(CABEZA.Caption, 1, 18) = "BALANCE TRIBUTARIO" Then
    Load informa04
    informa04.cmdato1.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 2)
    informa04.cmdato2.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 3, 2)
    informa04.cmdato3.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 5, 4)
    informa04.desdefecha.Caption = "01-01-" + Format(fechasistema, "yyyy")
    informa04.cmnombre.Caption = Grid1.Cell(Grid1.ActiveCell.row, 2).text
    informa04.hastafecha.Caption = Format(fechasistema, "dd-mm-yyyy")
    informa04.Show



End If
If leertiene(Grid1.Cell(1, 1).text, 1) = True And Mid(CABEZA.Caption, 1, 18) <> "BALANCE TRIBUTARIO" Then

informa44.ctdato1.text = Grid1.Cell(1, 1).text
informa44.ctdato2.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 9)
informa44.DV.text = Right(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1)
informa44.nombrectacte = Grid1.Cell(1, 2).text



informa44.ctnombre = Grid1.Cell(Grid1.ActiveCell.row, 2).text
informa44.sbtab1.Tab = 1
informa44.ctindi = True


informa44.Show

End If


End Sub

Private Sub TAMAÑOS_Click()

Grid1.DefaultFont.Size = TAMAÑOS.Value
'For K = 1 To Grid1.Cols - 1
'Grid1.Column(K).Width = Len(Grid1.Cell(2, K).text) * TAMAÑOS.Value


'Next K



Grid1.Refresh

End Sub
Sub imprime_NORMALES(Titulos)
Dim titulo As String
titulo = Titulos
Call cabezas(titulo, "N", "000000000", "")
'grid1.DefaultFont.Size = 6

Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub

Sub imprime_harina(tipo, FOLIO)
Dim titulo As String
Dim titulo2 As String

titulo = "ANEXO INFORME MENSUAL VENDEDORES DE HARINA "
titulo2 = "INFORMACION DEL MES DE " + infoharina.COMBOMES.text + " AÑO " + infoharina.COMBOAÑO.text



Call CABEZAS2(titulo, titulo2)
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait



Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_CARNE(tipo, FOLIO)
Dim titulo As String
Dim titulo2 As String

titulo = "ANEXO INFORME MENSUAL RETENCION CARNE "
titulo2 = "INFORMACION DEL MES DE " + infocarne.COMBOMES.text + " AÑO " + infocarne.COMBOAÑO.text



Call CABEZAS2(titulo, titulo2)
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait



Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub CABEZAS2(titulo, titulo2)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo2
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub
Sub imprime_descuadrados(tipo, FOLIO)
Dim titulo As String
titulo = "LISTADOS DESCUADRADOS "
Call cabezas(titulo, "N", FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellPortrait


 For i = 1 To Grid1.PageSetup.PaperSizes.Count
            If UCase(Grid1.PageSetup.PaperSizes.item(i).PaperName) = "OFICIO" Then
                Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.item(i).Kind
                Exit For
            End If
        Next i
        


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 0
Grid1.PageSetup.TopMargin = 0.5
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Private Sub ExportaSeleccion()
Dim k As Double
Dim i As Double
Dim rI As Double
Dim rF As Double
Dim datos(99) As String
rI = Grid1.Selection.FirstRow
rF = Grid1.Selection.LastRow

exporta.Cols = 1
exporta.Cols = 9
exporta.Rows = 1

exporta.Cell(0, 1).text = "CUENTA"
exporta.Cell(0, 2).text = "GLOSA"
exporta.Cell(0, 3).text = "TP"
exporta.Cell(0, 4).text = "NUMERO"
exporta.Cell(0, 5).text = "F.VENCIMIENTO"
exporta.Cell(0, 6).text = "MONTO"
exporta.Cell(0, 7).text = "DH"
exporta.Cell(0, 8).text = "RUT CTACTE"

exporta.Column(4).Mask = cellNumeric
exporta.Column(6).Mask = cellNumeric

For k = rI To rF
    datos(0) = Grid1.Cell(k, 5).text 'cuenta
    datos(1) = Grid1.Cell(k, 6).text 'glosa
    datos(2) = Grid1.Cell(k, 2).text 'tp
    datos(3) = Grid1.Cell(k, 3).text 'numero
    datos(4) = Grid1.Cell(k, 1).text 'vencimiento
 If Val(Grid1.Cell(k, 11).text) > 0 Then
    datos(5) = Grid1.Cell(k, 11).text 'monto
    datos(6) = "D" 'Grid1.Cell(k, 12).text 'D
 End If
 If Val(Grid1.Cell(k, 12).text) > 0 Then
    datos(5) = Grid1.Cell(k, 12).text 'monto
    datos(6) = "H" 'Grid1.Cell(k, 12).text 'D
 End If
    
    
    datos(7) = Grid1.Cell(k, 0).text 'CTACTE
    
    If Val(datos(5)) > 0 And IsDate(datos(4)) = True Then
        exporta.AddItem "", True
        exporta.Cell(exporta.Rows - 1, 1).text = datos(0)
        exporta.Cell(exporta.Rows - 1, 2).text = datos(1)
        exporta.Cell(exporta.Rows - 1, 3).text = datos(2)
        exporta.Cell(exporta.Rows - 1, 4).text = datos(3)
        exporta.Cell(exporta.Rows - 1, 5).text = datos(4)
        exporta.Cell(exporta.Rows - 1, 6).text = datos(5)
        exporta.Cell(exporta.Rows - 1, 7).text = datos(6)
        exporta.Cell(exporta.Rows - 1, 8).text = datos(7)
    End If
'    exporta.Cell(exporta.Rows - 1, 8).text = Datos(7)
Next k

exporta.Refresh

If exporta.Rows - 1 > 1 Then
    Call exporta.ExportToCSV("", True, False)
End If
End Sub
Sub generacomprobantecontable(numero, LINEA, fecha, cuenta, glosacontable, tipodoc, NUMERODOC, monto, DH, rutctate, CRCC, rutprove, cuentapresupuesto, centrogastos)
    Dim w As Long
    Dim tipo2 As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "glosacontable"
    campos(6, 0) = "tipodocumento"
    campos(7, 0) = "numerodocumento"
    campos(8, 0) = "fechadocumento"
    campos(9, 0) = "fechavencimiento"
    campos(10, 0) = "monto"
    campos(11, 0) = "dh"
    campos(12, 0) = "creadopor"
    campos(13, 0) = "mes"
    campos(14, 0) = "año"
    campos(15, 0) = "rutctacte"
    campos(16, 0) = "centrocosto"
    campos(17, 0) = "fechacreacion"
    campos(18, 0) = "horacreacion"
    campos(19, 0) = "rutproveedor"
    campos(20, 0) = "cuenta_presupuesto"
    campos(21, 0) = "centro_gastos"
    campos(22, 0) = ""
    
    campos(0, 1) = "FM"
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = cuenta
    campos(5, 1) = glosacontable
    campos(6, 1) = tipodoc
    campos(7, 1) = NUMERODOC
    campos(8, 1) = campos(3, 1)
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Replace(monto, ",", ".")
    campos(11, 1) = DH
    campos(12, 1) = USUARIOSISTEMA
    campos(13, 1) = Format(fechasistema, "mm")
    campos(14, 1) = Format(fechasistema, "yyyy")
    campos(15, 1) = rutctacte
    campos(16, 1) = CRCC
    campos(17, 1) = Format(Date$, "yyyy-mm-dd")
    campos(18, 1) = Time$
    campos(19, 1) = rutprove
    campos(20, 1) = cuentapresupuesto
    campos(21, 1) = centrogastos
    
    campos(0, 2) = "movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
     
End Sub
Function ultimo(tipo) As String
    campos(0, 0) = "tipo"
    campos(1, 0) = "MAX(numero)"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + tipo + "' and año='" + Format(fechasistema, "yyyy") + "' and mes='" + Format(fechasistema, "mm") + "' "
   op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    ultimo = Format(sqlconta.response(1, 3) + 1, "0000000000")
    
      
    Rem aca
    End Function
Sub ArchivoComplementos()
  Dim ARCHIVO As String
  Dim contador As Double
  Dim codigo_iva As Double
  Dim tipodoc As String
  Dim cadena As String
  Dim TpoTranCompra As Double
  
     
  
    
    Close 20
    ARCHIVO = "C:\LIBROS\Caracterizacion_" + empresaactiva + "_" & año_lc & "_" & mes_lc + ".csv"


        Open ARCHIVO For Output As #20
        contador = 0
    For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 1).text <> "" Then
        Rem If Grid1.Cell(k, 1).BackColor <> vbGreen Then GoTo no:
        
            contador = contador + 1
            If contador = 1 Then
                cadena = "RUT-DV;Codigo_Tipo_Doc;Folio_Doc;TpoTranCompra;Codigo_IVA_e_Impuestos"
                Print #20, cadena
            End If
            cadena = Val(Mid(Grid1.Cell(k, 5).text, 1, 9)) & "-" & Right(Grid1.Cell(k, 5).text, 1) & ";"  'rut-dv
       tipo = Grid1.Cell(k, 2).text
  If tipo = "FA" Then tipo = "1"
 If tipo = "ND" Then tipo = "2"
If tipo = "NC" Then tipo = "3"
If tipo = "FAE" Then tipo = "4"
If tipo = "NDE" Then tipo = "5"
If tipo = "NCE" Then tipo = "6"
If tipo = "FC" Then tipo = "7"
If tipo = "IM" Then tipo = "8"
If tipo = "FE" Then tipo = "9"
If tipo = "FEE" Then tipo = "0"
If tipo = "LFE" Then tipo = "L"
If tipo = 4 Then
 tipo = "33"
End If
If tipo = 5 Then
 tipo = "56"
End If
If tipo = 6 Then
 tipo = "61"
End If
If tipo = 0 Then
 tipo = "34"
End If
If tipo = 1 Then
 tipo = "30"
End If
If tipo = 7 Then
 tipo = "46"
End If
If tipo = 3 Then
 tipo = "60"
End If
If tipo = 8 Then
 tipo = "914"

End If

            
            cadena = cadena & tipo & ";" 'Codigo_Tipo_Doc;
            cadena = cadena & Grid1.Cell(k, 3).text & ";" 'Folio_Doc;
            'TpoTranCompra;
            '1 compras del giro
            '2 compra supermercados o compercios similares
            '3 adquisicion bienes raices
            '4 activo fijo
            '5 compras con IVA uso Comun
            '6 Compras sin Derecho  a Credito(IVA no Recuperable)
            '7 compras que no corresponde incluir
'            FormatoGrilla(1, 20) = "IVA/N/R"
'            FormatoGrilla(1, 21) = "USO COMUN"
'            FormatoGrilla(1, 22) = "A/F"
           Rem  If Grid1.Cell(k, 3).text = "0000000039" Then Stop
            TpoTranCompra = 1
            codigo_iva = 1
            If Grid1.Cell(k, 22).text = "S" Then ' activo fijo
                TpoTranCompra = 4
                
            End If
            If Grid1.Cell(k, 21).text <> "" And Grid1.Cell(k, 20).text <> "" Then  '  iva uso comun
                TpoTranCompra = 5
                codigo_iva = 2
            End If
           
            
            If Grid1.Cell(k, 20).text <> "" And Grid1.Cell(k, 21).text = "" Then '  iva no recuperado
                TpoTranCompra = 6
                tipodoc = Grid1.Cell(k, 2).text
                codigo_iva = buscamotivo(tipodoc, Grid1.Cell(k, 3).text, Format(Grid1.Cell(k, 4).text, "yyy-mm-dd"), Grid1.Cell(k, 5).text)
            End If
            
            
            
            cadena = cadena & TpoTranCompra & ";" 'TpoTranCompra;
            
            
            cadena = cadena & codigo_iva  'Codigo_IVA_e_Impuestos;
            
            If TpoTranCompra <> 1 Then
            
            Print #20, cadena
            
            End If
                
no:
        End If
    Next k
    
    Close 20
    Shell "NOTEPAD " + ARCHIVO
End Sub

Function buscamotivo(tipo, numero, fecha, rutprove) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    If tipo = "FA" Then tipo = "1"
If tipo = "ND" Then tipo = "2"
If tipo = "NC" Then tipo = "3"
If tipo = "FAE" Then tipo = "4"
If tipo = "NDE" Then tipo = "5"
If tipo = "NCE" Then tipo = "6"
If tipo = "FC" Then tipo = "7"
If tipo = "IM" Then tipo = "8"
If tipo = "FE" Then tipo = "9"
If tipo = "FEE" Then tipo = "0"
If tipo = "LFE" Then tipo = "L"

    csql.sql = "select motivo from " & cliente_sql & "conta" & empresaactiva & ".facturasdecompras_norecuperable "
    csql.sql = csql.sql & " where tipo='" & tipo & "' and numero='" & numero & "' and rut='" & Replace(rutprove, "-", "") & "' "
    csql.Execute
    buscamotivo = 1
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        buscamotivo = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
    
End Function

Sub ArchivoDocManuales()
    Dim k As Double
    Dim ARCHIVO As String
    Dim cadena As String
    Dim i As Double
    Dim o As Double
    Dim tipotrans As String
    
    
    
    
  
    
    Close 20
    ARCHIVO = "u:\LIBROS\Doc_manuales_libro_compras_" + empresaactiva + "_" + año_lc & "_" & mes_lc + ".csv"


        Open ARCHIVO For Output As #20
    Print #20, "titulos"
    For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 2).text <> "" Then
            o = 0
              For i = 13 To 19
                If Grid1.Cell(k, i).text > 0 Then
                    o = o + 1
                End If
            Next i
            If o > 0 Then
                For i = 13 To 19
                    If Grid1.Cell(k, i).text > 0 Then
                       If Grid1.Cell(k, 2).text = "FA" Then tipo = "30"
                       If Grid1.Cell(k, 2).text = "ND" Then tipo = "55"
                       If Grid1.Cell(k, 2).text = "NC" Then tipo = "60"
                       If Grid1.Cell(k, 2).text = "IM" Then tipo = "914"
                        cadena = tipo & ";" 'Tipo Doc;
                        cadena = cadena & Grid1.Cell(k, 3).text & ";" 'Folio;
                        cadena = cadena & Val(Mid(Grid1.Cell(k, 5).text, 1, 9)) & "-" & Right(Grid1.Cell(k, 5).text, 1) & ";" 'Rut Contraparte;
                        cadena = cadena & "19" & ";" 'Tasa Impuesto;
                        cadena = cadena & Grid1.Cell(k, 6).text & ";" 'Razon Social Contraparte;
                        cadena = cadena & "1" & ";" 'Tipo Impuesto[1=IVA:2=LEY 18211];
                        cadena = cadena & Format(Grid1.Cell(k, 4).text, "dd-mm-yyyy") & ";" 'Fecha Emision;
                        If tipo = "60" Then
                       For j = 7 To 22
                       Grid1.Cell(k, j).text = Replace(Grid1.Cell(k, j).text, "-", "")
                       
                       Next j
                       End If
                        
                        cadena = cadena & Grid1.Cell(k, 9).text & ";" 'Monto Exento;
                        cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Neto;
                        cadena = cadena & Val(Grid1.Cell(k, 8).text) - Val(Grid1.Cell(k, 21).text) & ";" 'Monto IVA (Recuperable);
                        If Val(Grid1.Cell(k, 20).text) > 0 Then
                            cadena = cadena & "1" & ";" 'Cod IVA no Rec;
                            cadena = cadena & Grid1.Cell(k, 20).text & ";" 'Monto IVA no Rec;
                        Else
                            cadena = cadena & "" & ";" 'Cod IVA no Rec;
                            cadena = cadena & "" & ";" 'Monto IVA no Rec;
                        End If
                        
                        cadena = cadena & Grid1.Cell(k, 21).text & ";" 'IVA Uso Comun;
            
          


                        If i = 13 Then
                            cadena = cadena & 271 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 18 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                        If i = 14 Then
                            cadena = cadena & 24 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & "31.5" & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 15 Then
                            cadena = cadena & 25 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & "20.5" & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 16 Then
                            cadena = cadena & 26 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & "20.5" & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 17 Then
                            cadena = cadena & 19 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 12 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 18 Then
                            cadena = cadena & 18 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 5 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                         If i = 19 Then
                            cadena = cadena & 27 & ";" 'Cod Otro Imp (Con Credito);
                            cadena = cadena & 10 & ";" 'Tasa Otro Imp (Con Credito);
                        End If
                        
                       
                        cadena = cadena & Grid1.Cell(k, i).text & ";" 'Monto Otro Imp (Con Credito);
                        
                        cadena = cadena & "" & ";" 'Monto Otro Imp Sin Credito;
                        
                        'FormatoGrilla(1, 21) = "USO COMUN"
                        'FormatoGrilla(1, 22) = "A/F"
                        
'                        FormatoGrilla(1, 20) = "IVA/N/R"
'                        FormatoGrilla(1, 21) = "USO COMUN"
'                        FormatoGrilla(1, 22) = "A/F"
    
    
                        If Grid1.Cell(k, 22).text = "S" Then
                            cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Activo Fijo;
                            cadena = cadena & Grid1.Cell(k, 8).text & ";" 'Monto IVA Activo Fijo;
                        Else
                            cadena = cadena & "" & ";" 'Monto Activo Fijo;
                            cadena = cadena & "" & ";" 'Monto IVA Activo Fijo;
                        End If
                        
                        cadena = cadena & "" & ";" 'IVA No Retenido;
                        
                        
                        
                        cadena = cadena & "" & ";" 'Tabacos - Puros;
                        cadena = cadena & "" & ";" 'Tabacos - Cigarrillos;
                        cadena = cadena & "" & ";" 'Tabacos - Elaborados;
                        cadena = cadena & "" & ";" 'Codigo sucursal SII;
                        cadena = cadena & "" & ";" 'Numero Interno;
                        cadena = cadena & "" & ";" 'Emisor/Receptor;
                        cadena = cadena & Grid1.Cell(k, 12).text & ";" 'Monto Total;
                        tipotrans = "1"
                        
'                        If Grid1.Cell(k, 20).text = "" And Grid1.Cell(k, 21).text = "" And Grid1.Cell(k, 22).text = "N" Then
'                            tipotrans = "2"
'                        End If
                        
                        If Grid1.Cell(k, 22).text = "S" Then
                           tipotrans = "4" 'Tipo Transaccion
                        End If
                        
                        If Grid1.Cell(k, 20).text <> "" Then
                           tipotrans = "6" 'Tipo Transaccion
                        End If
                        If Grid1.Cell(k, 21).text <> "" Then
                           tipotrans = "5" 'Tipo Transaccion
                        End If
                        
                        
                        
                         cadena = cadena & tipotrans
                         
                         
                      
                        
                        
                         Print #20, cadena
                    End If
                Next i
            Else
                        
                       If Grid1.Cell(k, 2).text = "FA" Then tipo = "30"
                       If Grid1.Cell(k, 2).text = "ND" Then tipo = "55"
                       If Grid1.Cell(k, 2).text = "NC" Then tipo = "60"
                       If Grid1.Cell(k, 2).text = "IM" Then tipo = "914"
                        cadena = tipo & ";" 'Tipo Doc;
                        cadena = cadena & Grid1.Cell(k, 3).text & ";" 'Folio;
                        cadena = cadena & Val(Mid(Grid1.Cell(k, 5).text, 1, 9)) & "-" & Right(Grid1.Cell(k, 5).text, 1) & ";" 'Rut Contraparte;
                        cadena = cadena & "19" & ";" 'Tasa Impuesto;
                        cadena = cadena & Grid1.Cell(k, 6).text & ";" 'Razon Social Contraparte;
                        cadena = cadena & "1" & ";" 'Tipo Impuesto[1=IVA:2=LEY 18211];
                        cadena = cadena & Format(Grid1.Cell(k, 4).text, "dd-mm-yyyy") & ";" 'Fecha Emision;
                       If tipo = "60" Then
                       For j = 7 To 22
                       Grid1.Cell(k, j).text = Replace(Grid1.Cell(k, j).text, "-", "")
                       
                       Next j
                       End If
                        
                        cadena = cadena & Grid1.Cell(k, 9).text & ";" 'Monto Exento;
                        cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Neto;
                        cadena = cadena & Val(Grid1.Cell(k, 8).text) - Val(Grid1.Cell(k, 21).text) & ";" 'Monto IVA (Recuperable);
                        If Val(Grid1.Cell(k, 20).text) > 0 Then
                            cadena = cadena & "1" & ";" 'Cod IVA no Rec;
                            cadena = cadena & Grid1.Cell(k, 20).text & ";" 'Monto IVA no Rec;
                        Else
                            cadena = cadena & "" & ";" 'Cod IVA no Rec;
                            cadena = cadena & "" & ";" 'Monto IVA no Rec;
                        End If
                        
                        cadena = cadena & Grid1.Cell(k, 21).text & ";" 'IVA Uso Comun;

                        cadena = cadena & "" & ";" 'Cod Otro Imp (Con Credito);
                        cadena = cadena & "" & ";" 'Tasa Otro Imp (Con Credito);
                        cadena = cadena & "" & ";" 'Monto Otro Imp (Con Credito);
                        
                        cadena = cadena & "" & ";" 'Monto Otro Imp Sin Credito;
                         If Grid1.Cell(k, 22).text = "S" Then
                            cadena = cadena & Grid1.Cell(k, 7).text & ";" 'Monto Activo Fijo;
                            cadena = cadena & Grid1.Cell(k, 8).text & ";" 'Monto IVA Activo Fijo;
                        Else
                            cadena = cadena & "" & ";" 'Monto Activo Fijo;
                            cadena = cadena & "" & ";" 'Monto IVA Activo Fijo;
                        End If
                        cadena = cadena & "" & ";" 'IVA No Retenido;
                        
                        
                        
                        cadena = cadena & "" & ";" 'Tabacos - Puros;
                        cadena = cadena & "" & ";" 'Tabacos - Cigarrillos;
                        cadena = cadena & "" & ";" 'Tabacos - Elaborados;
                        cadena = cadena & "" & ";" 'Codigo sucursal SII;
                        cadena = cadena & "" & ";" 'Numero Interno;
                        cadena = cadena & "" & ";" 'Emisor/Receptor;
                        cadena = cadena & Grid1.Cell(k, 12).text & ";" 'Monto Total;
                        
                        
                        tipotrans = "1"
                        
'                        If Grid1.Cell(k, 20).text = "" And Grid1.Cell(k, 21).text = "" And Grid1.Cell(k, 22).text = "N" Then
'                            tipotrans = "2"
'                        End If
                        
                        If Grid1.Cell(k, 22).text = "S" Then
                           tipotrans = "4" 'Tipo Transaccion
                        End If
                        
                        If Grid1.Cell(k, 20).text <> "" Then
                           tipotrans = "6" 'Tipo Transaccion
                        End If
                        If Grid1.Cell(k, 21).text <> "" Then
                           tipotrans = "5" 'Tipo Transaccion
                        End If
                        
                        
                        
                         cadena = cadena & tipotrans
                        
                        Print #20, cadena
            End If
            
        End If
    Next k
    
    Close 20
    Shell "NOTEPAD " + ARCHIVO

End Sub

