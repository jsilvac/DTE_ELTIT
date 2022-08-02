Attribute VB_Name = "libroelec"
Public lce As String
Public xmllibrodiario As Boolean
Public confi_servermail As String
Public confi_mailsalida As String
Public confi_clavemail As String
Public archivopdf As String
Public confi_localempresa As String


Public TOTAL33 As Double
Public TOTAL30 As Double
Public texento As Double

Public xml As New ChilkatXml

Public Sub generalc(TpoDoc, NroDoc, FchDoc, RUTDoc, NOMBRE, MntExe, MntNeto, MntIVA, MntActivoFijo, MntSinCred, MntTotal, TabCigarrillos, refresco, licores, vinos, cerveza, HARINA, CARNE, Diesel, proporcion, ivausocomun, SINAZUCAR, dieselrecuperado)
Dim IVANUEVO2 As Double

If TpoDoc = "FA" Then TpoDoc = "30"
If TpoDoc = "ND" Then TpoDoc = "55"
If TpoDoc = "NC" Then TpoDoc = "60"
If TpoDoc = "FAE" Then TpoDoc = "33"
If TpoDoc = "NDE" Then TpoDoc = "56"
If TpoDoc = "NCE" Then TpoDoc = "61"
If TpoDoc = "FC" Then TpoDoc = "46"
If TpoDoc = "IM" Then TpoDoc = "914"
If TpoDoc = "FE" Then TpoDoc = "32"
If TpoDoc = "FEE" Then TpoDoc = "34"
If TpoDoc = "LFE" Then TpoDoc = "43"
'If diesel <> 0 Then Stop
'linea original
MntExe = MntExe + CDbl(Diesel) - CDbl(dieselrecuperado)
MntExe = Replace(MntExe, "-", "")
MntNeto = Replace(MntNeto, "-", "")
MntIVA = Replace(MntIVA, "-", "")
MntActivoFijo = Replace(MntActivoFijo, "-", "")
MntSinCred = Replace(MntSinCred, "-", "")
MntTotal = Replace(MntTotal, "-", "")

NOMBRE = Replace(NOMBRE, "¥", "N")
RUTDoc = Format(Mid(RUTDoc, 2, 8), "########") + "-" + Mid(RUTDoc, 11, 1)

lce = lce + "<Detalle>"
lce = lce + "<TpoDoc>" + TpoDoc + "</TpoDoc>"
lce = lce + "<NroDoc>" & Val(NroDoc) & "</NroDoc>"
lce = lce + "<FchDoc>" + Format(FchDoc, "yyyy-mm-dd") + "</FchDoc>"
lce = lce + "<RUTDoc>" + RUTDoc + "</RUTDoc>"
lce = lce + "<RznSoc>" + Replace(NOMBRE, "Ñ", "N") + "</RznSoc>"
lce = lce + "<MntExe>" + Replace(MntExe, ".", "") + "</MntExe>"
lce = lce + "<MntNeto>" + Replace(MntNeto, ".", "") + "</MntNeto>"
lce = lce + "<MntIVA>" + Replace(MntIVA, ".", "") + "</MntIVA>"
lce = lce + "<MntActivoFijo>" + Replace(MntActivoFijo, ".", "") + "</MntActivoFijo>"
'If TpoDoc = "33" Then
'
'texento = texento + Val(MntExe)
'End If


'If Val(NroDoc) = 375 Then Stop
If ivausocomun = "" Then ivausocomun = "0"
If proporcion = "" Then proporcion = "0"

If Val(ivausocomun) <> 0 Then
proporcion = Replace(proporcion, "-", "")
ivausocomun = Replace(ivausocomun, "-", "")
IVANUEVO2 = CDbl(MntIVA) - CDbl(proporcion)
lce = lce + "<IVANoRec>"
lce = lce + "<CodIVANoRec>" + "1" + "</CodIVANoRec"
lce = lce + "<MntIVANoRec>" + Replace(proporcion, ".", "") + "</MntIVANoRec>"
lce = lce + "</IVANoRec>"
lce = lce + "<IVAUsoComun>" & ivausocomun & "</IVAUsoComun>"
If TpoDoc = "30" Then TOTAL30 = TOTAL30 + CDbl(ivausocomun)
If TpoDoc = "33" Then TOTAL33 = TOTAL33 + CDbl(ivausocomun)

End If
If Val(ivausocomun) = 0 And proporcion <> 0 Then
proporcion = Replace(proporcion, "-", "")
ivausocomun = Replace(ivausocomun, "-", "")
IVANUEVO2 = CDbl(MntIVA) - CDbl(proporcion)
lce = lce + "<IVANoRec>"
lce = lce + "<CodIVANoRec>" + "2" + "</CodIVANoRec"
lce = lce + "<MntIVANoRec>" + Replace(proporcion, ".", "") + "</MntIVANoRec>"
lce = lce + "</IVANoRec>"


End If

Rem lce = lce + "<MntSinCred>" + Replace(MntSinCred, ".", "") + "</MntSinCred>"
'If MntSinCred <> 0 Then
'lce = lce + "<IVANoRec>"
'lce = lce + "<CodIVANoRec>" + "4" + "</CodIVANoRec>"
'lce = lce + "<MntIVANoRec>" + Replace(MntSinCred, ".", "") + "</MntIVANoRec>"
'lce = lce + "</IVANoRec>"
'End If
If TpoDoc = "46" Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "15" + "</CodImp>"
lce = lce + "<TasaImp>" + "19" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(MntIVA, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"
End If

refresco = Replace(refresco, "-", "")

If refresco <> 0 Then
lce = lce + "<OtrosImp>"
If Format(fechasistema, "yyyy-mm-dd") > "2014-09-30" Then
lce = lce + "<CodImp>" + "271" + "</CodImp>"
Else
lce = lce + "<CodImp>" + "27" + "</CodImp>"
End If

lce = lce + "<TasaImp>" + "18" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(refresco, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If

SINAZUCAR = Replace(SINAZUCAR, "-", "")

If SINAZUCAR <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "27" + "</CodImp>"
lce = lce + "<TasaImp>" + "10" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(SINAZUCAR, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If


licores = Replace(licores, "-", "")

If licores <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "24" + "</CodImp>"
lce = lce + "<TasaImp>" + "31.5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(licores, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
vinos = Replace(vinos, "-", "")

If vinos <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "25" + "</CodImp>"
lce = lce + "<TasaImp>" + "20.5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(vinos, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
cerveza = Replace(cerveza, "-", "")

If cerveza <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "26" + "</CodImp>"
lce = lce + "<TasaImp>" + "20.5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(cerveza, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
HARINA = Replace(HARINA, "-", "")

If HARINA <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "19" + "</CodImp>"
lce = lce + "<TasaImp>" + "12" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(HARINA, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
CARNE = Replace(CARNE, "-", "")

If CARNE <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "18" + "</CodImp>"
lce = lce + "<TasaImp>" + "5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(CARNE, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
dieselrecuperado = Replace(dieselrecuperado, "-", "")
If dieselrecuperado <> 0 Then
    lce = lce + "<OtrosImp>"
    lce = lce + "<CodImp>" + "28" + "</CodImp>"
    lce = lce + "<TasaImp>" + "1.5" + "</TasaImp>"
    lce = lce + "<MntImp>" + Replace(dieselrecuperado, ".", "") + "</TasaImp>"
    lce = lce + "</OtrosImp>"
End If


'If Val(NroDoc) = 781 Then
'lce = lce + "<IVANoRec>"
'lce = lce + "<CodIVANoRec>" + "1" + "</CodIVANoRec>"
'lce = lce + "<MntIVANoRec>" + "2264" + "</MntIVANoRec>"
'lce = lce + "</IVANoRec>"
'lce = lce + "<IVAUsoComun>" + "5661" + "</IVAUsoComun>"
'
'End If


lce = lce + "<MntTotal>" + Replace(MntTotal, ".", "") + "</MntTotal>"
lce = lce + "</Detalle>"
If TpoDoc = "33" Then
If MntExe <> "0" Or Diesel <> "0" Then
EXENTO = EXENTO + MntExe + CDbl(Diesel) - CDbl(dieselrecuperado)
diesel2 = diesel2 + CDbl(Diesel) - CDbl(dieselrecuperado)
End If
End If

End Sub
Public Sub generalv(TpoDoc, NroDoc, FchDoc, RUTDoc, NOMBRE, MntExe, MntNeto, MntIVA, MntActivoFijo, MntSinCred, MntTotal, TabCigarrillos, nula, refresco, licores, vinos, cerveza, HARINA, CARNE, SINAZUCAR)
If TpoDoc = "FX" Then TpoDoc = "32"
If TpoDoc = "FA" Then TpoDoc = "30"
If TpoDoc = "ND" Then TpoDoc = "55"
If TpoDoc = "NB" Then Exit Sub
If TpoDoc = "NF" Then TpoDoc = "60"
If TpoDoc = "FE" Then TpoDoc = "101"

If TpoDoc = "FAE" Then TpoDoc = "33"
If TpoDoc = "NDE" Then TpoDoc = "56"
If TpoDoc = "NCE" Then TpoDoc = "61": Stop
If TpoDoc = "FEE" Then TpoDoc = "34"

MntExe = Replace(MntExe, "-", "")
Rem If MntNeto = 0 Then TpoDoc = "32"
MntNeto = Replace(MntNeto, "-", "")
MntIVA = Replace(MntIVA, "-", "")
MntActivoFijo = Replace(MntActivoFijo, "-", "")
MntSinCred = Replace(MntSinCred, "-", "")
MntTotal = Replace(MntTotal, "-", "")
Rem If Mid(RUTDoc, 1, 1) = "." Then Stop
RUTDoc = Replace(RUTDoc, ".", "0")
RUTDoc = Format(Mid(RUTDoc, 1, 9), "########") + "-" + Mid(RUTDoc, 11, 1)
refresco = Replace(refresco, "-", "")
licores = Replace(licores, "-", "")
vinos = Replace(vinos, "-", "")
cerveza = Replace(cerveza, "-", "")
HARINA = Replace(HARINA, "-", "")
CARNE = Replace(CARNE, "-", "")
SINAZUCAR = Replace(SINAZUCAR, "-", "")

NOMBRE = Replace(NOMBRE, "¥", "N")
lce = lce + "<Detalle>"
lce = lce + "<TpoDoc>" + TpoDoc + "</TpoDoc>"
lce = lce + "<NroDoc>" & Val(NroDoc) & "</NroDoc>"
If Mid(RUTDoc, 1, 9) = "888888888" Then RUTDoc = "88888888-8"
If RUTDoc = "88888888-8" Then nula = "A" Else nula = ""

If nula = "A" Then
lce = lce + "<Anulado>A</Anulado>"
GoTo 20
End If

lce = lce + " <TasaImp>19.0</TasaImp>"
lce = lce + "<FchDoc>" + Format(FchDoc, "yyyy-mm-dd") + "</FchDoc>"
lce = lce + "<RUTDoc>" + RUTDoc + "</RUTDoc>"
lce = lce + "<RznSoc>" + Replace(NOMBRE, "Ñ", "N") + "</RznSoc>"
If MntExe <> 0 Then
lce = lce + "<MntExe>" + Replace(MntExe, ".", "") + "</MntExe>"
End If
lce = lce + "<MntNeto>" + Replace(MntNeto, ".", "") + "</MntNeto>"
lce = lce + "<MntIVA>" + Replace(MntIVA, ".", "") + "</MntIVA>"
If refresco <> 0 Then
lce = lce + "<OtrosImp>"
If Format(fechasistema, "yyyy-mm-dd") > "2014-09-30" Then
lce = lce + "<CodImp>" + "271" + "</CodImp>"
Else
lce = lce + "<CodImp>" + "27" + "</CodImp>"
End If
lce = lce + "<TasaImp>" + "18" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(refresco, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
If SINAZUCAR <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "27" + "</CodImp>"
lce = lce + "<TasaImp>" + "10" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(SINAZUCAR, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If

If licores <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "24" + "</CodImp>"
lce = lce + "<TasaImp>" + "31.5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(licores, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
If vinos <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "25" + "</CodImp>"
lce = lce + "<TasaImp>" + "20.5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(vinos, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
If cerveza <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "26" + "</CodImp>"
lce = lce + "<TasaImp>" + "20.5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(cerveza, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
If HARINA <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "19" + "</CodImp>"
lce = lce + "<TasaImp>" + "12" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(HARINA, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If
If CARNE <> 0 Then
lce = lce + "<OtrosImp>"
lce = lce + "<CodImp>" + "18" + "</CodImp>"
lce = lce + "<TasaImp>" + "5" + "</TasaImp>"
lce = lce + "<MntImp>" + Replace(CARNE, ".", "") + "</TasaImp>"
lce = lce + "</OtrosImp>"

End If


lce = lce + "<MntTotal>" + Replace(MntTotal, ".", "") + "</MntTotal>"

20:
lce = lce + "</Detalle>"

End Sub


Public Sub generacaratula(rutempresa, rutenviasii, periodo, fecharesolucion, numeroresulocion, tipooperacion, tipolibro, tipoenvio, numerosegmento, folionotificacion)
Dim comi As String
Dim NOMBRE As String

comi = Chr(34)
NOMBRE = periodo
lce = " "
lce = lce + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
If f3328 = True Then
lce = lce + "<LibroCompraVenta xmlns=" + comi + "http://www.sii.cl/SiiDte" + comi + " xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiDte LibroCVS_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
Else
lce = lce + "<LibroCompraVenta xmlns=" + comi + "http://www.sii.cl/SiiDte" + comi + " xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiDte LibroCV_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
End If
lce = lce + "<EnvioLibro ID=" + comi + "LCV-" + NOMBRE + comi + ">"
lce = lce + "<Caratula>"
rutempresa = Format(Mid(rutempresa, 1, 8), "########") + "-" + Mid(rutempresa, 10, 1)
Rem rutenviasii = Format(Mid(rutenviasii, 1, 8), "########") + "-" + Mid(rutenviasii, 10, 1)
'fecharesolucion = "2010-12-29"
Rem numeroresolucion = "0"
lce = lce + "<RutEmisorLibro>" + rutempresa + "</RutEmisorLibro>"
If f3328 = True Then
lce = lce + "<RutEnvia>" + rutempresa + "</RutEnvia>"
Else
lce = lce + "<RutEnvia>" + rutenviasii + "</RutEnvia>"
End If

lce = lce + "<PeriodoTributario>" + periodo + "</PeriodoTributario>"
If f3328 = True Then
lce = lce + "<FchResol>" + "2006-01-20" + "</FchResol>"
lce = lce + "<NroResol>102006</NroResol>"
Else
lce = lce + "<FchResol>" + Format(fecharesolucion, "yyyy-mm-dd") + "</FchResol>"
lce = lce + "<NroResol>" + numeroresolucion + "</NroResol>"
End If

lce = lce + "<TipoOperacion>" + tipooperacion + "</TipoOperacion>"

'LINEAS ORIGINALES
'-----------------
    ''If f3328 = True Then
    ''lce = lce + "<TipoLibro>ESPECIAL</TipoLibro>"
    ''Else
    ''lce = lce + "<TipoLibro>" + tipolibro + "</TipoLibro>"
    ''
    ''End If
    ''
    ''If f3328 = True Then
    ''lce = lce + "<TipoEnvio>" + tipoenvio + "</TipoEnvio>"
    ''
    ''lce = lce + "<FolioNotificacion>102006</FolioNotificacion>"
    ''Else
    ''
    ''lce = lce + "<TipoEnvio>" + tipoenvio + "</TipoEnvio>"
    ''End If
'-----------------

If f3328 = True Then
    lce = lce + "<TipoLibro>ESPECIAL</TipoLibro>"
Else
    If CODAUTREC = "" Then
        lce = lce + "<TipoLibro>" + tipolibro + "</TipoLibro>"
    Else
        lce = lce + "<TipoLibro>RECTIFICA</TipoLibro>"
        lce = lce + "<TipoEnvio>TOTAL</TipoEnvio>"
        lce = lce + "<CodAutRec>" + CODAUTREC + "</CodAutRec>"
    End If
End If

If f3328 = True Then
    lce = lce + "<TipoEnvio>" + tipoenvio + "</TipoEnvio>"
    lce = lce + "<FolioNotificacion>102006</FolioNotificacion>"
Else
     If CODAUTREC = "" Then
        lce = lce + "<TipoEnvio>" + tipoenvio + "</TipoEnvio>"
     End If
End If


lce = lce + "</Caratula>"
lce = lce + "<ResumenPeriodo>"



End Sub

Public Sub GENERATOTALTIPO(tipo, totaldocumentos, TOTALEXENTO, TOTALNETO, TOTALIVA, total, ivanorecu, totalrefresco, totallicores, totalvinos, totalcerveza, totalharina, totalcarne, totaldiesel, proporcion, ivausocomun, totalnoazucar, totaldieselrecuperado)
Dim IVANUEVO2 As Double

lce = lce + "<TotalesPeriodo>"
lce = lce + "<TpoDoc>" + tipo + "</TpoDoc>"
lce = lce + "<TotDoc>" + Replace(totaldocumentos, ".", "") + "</TotDoc>"

'linea original
TOTALEXENTO = TOTALEXENTO + CDbl(totaldiesel) - CDbl(totaldieselrecuperado)
'diesel va por separado del exento

lce = lce + "<TotMntExe>" + Replace(TOTALEXENTO, ".", "") + "</TotMntExe>"
lce = lce + "<TotMntNeto>" + Replace(TOTALNETO, ".", "") + "</TotMntNeto>"
lce = lce + "<TotMntIVA>" + Replace(TOTALIVA, ".", "") + "</TotMntIVA>"
If ivausocomun = "" Then ivausocomun = "0"
If proporcion = "" Then proporcion = "0"
If CDbl(ivausocomun) <> 0 Then
IVANUEVO2 = CDbl(ivausocomun) - CDbl(proporcion)
lce = lce + "<TotIVANoRec>"
lce = lce + "<CodIVANoRec>" + "1" + "</CodIVANoRec"
Rem lce = lce + "<TotOpIVANoRec>" + "1" + "</TotOpIVANoRec>"
lce = lce + "<TotMntIVANoRec>" + Replace(proporcion, ".", "") + "</TotMntIVANoRec>"
lce = lce + "</TotIVANoRec>"
Rem lce = lce + "<TotOpIVAUsoComun>" + "1" + "</TotOpIVAUsoComun>"
lce = lce + "<TotIVAUsoComun>" & Replace(ivausocomun, ".", "") & "</TotIVAUsoComun>"
lce = lce + "<FctProp>" + Replace(proporcional, ",", ".") + "</FctProp>"
lce = lce + "<TotCredIVAUsoComun>" & IVANUEVO2 & "</TotCredIVAUsoComun>"

End If



If proporcion <> 0 And ivausocomun = 0 Then
lce = lce + "<TotIVANoRec>"
lce = lce + "<CodIVANoRec>" + "4" + "</CodIVANoRec"

lce = lce + "<TotMntIVANoRec>" + Replace(proporcion, ".", "") + "</TotMntIVANoRec>"
lce = lce + "</TotIVANoRec>"
End If
If tipo = "46" Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "15" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(TOTALIVA, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"
End If

If totalrefresco <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "271" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalrefresco, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"
End If
If totallicores <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "24" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totallicores, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalvinos <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "25" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalvinos, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalcerveza <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "26" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalcerveza, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalharina <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "19" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalharina, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalcarne <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "18" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalcarne, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalnoazucar <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "27" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalnoazucar, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If


'nuevo por rz 2017-07-06
If totaldieselrecuperado <> 0 Then
    lce = lce + "<TotOtrosImp>"
    lce = lce + "<CodImp>" + "28" + "</CodImp>"
    lce = lce + "<TotMntImp>" + Replace(totaldieselrecuperado, ".", "") + "</TasaImp>"
    lce = lce + "</TotOtrosImp>"
End If

lce = lce + "<TotMntTotal>" + Replace(total, ".", "") + "</TotMntTotal>"
lce = lce + "</TotalesPeriodo>"

End Sub
Public Sub GENERATOTALTIPOV(tipo, totaldocumentos, TOTALEXENTO, TOTALNETO, TOTALIVA, total, ivanorecu, totalrefresco, totallicores, totalvinos, totalcerveza, totalharina, totalcarne, totalnoazucar, netocom, ivacom, totcom)
lce = lce + "<TotalesPeriodo>"
lce = lce + "<TpoDoc>" + tipo + "</TpoDoc>"
lce = lce + "<TotDoc>" + Replace(totaldocumentos, ".", "") + "</TotDoc>"
If tipo = "38" Then
lce = lce + "<TotMntExe>" + Replace(TOTALEXENTO, ".", "") + "</TotMntExe>"
lce = lce + "<TotMntNeto>0</TotMntNeto>"
lce = lce + "<TotMntIVA>0</TotMntIVA>"

Else
lce = lce + "<TotMntExe>" + Replace(TOTALEXENTO, ".", "") + "</TotMntExe>"
lce = lce + "<TotMntNeto>" + Replace(TOTALNETO, ".", "") + "</TotMntNeto>"
lce = lce + "<TotMntIVA>" + Replace(TOTALIVA, ".", "") + "</TotMntIVA>"

End If

totalrefresco = Replace(totalrefresco, "-", "")
totallicores = Replace(totallicores, "-", "")
totalvinos = Replace(totalvinos, "-", "")
totalcerveza = Replace(totalcerveza, "-", "")
totalharina = Replace(totalharina, "-", "")
totalcarne = Replace(totalcarne, "-", "")
totalnoazucar = Replace(totalnoazucar, "-", "")

If totalrefresco <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "271" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalrefresco, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"
End If
If totalnoazucar <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "27" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalnoazucar, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"
End If


If totallicores <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "24" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totallicores, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalvinos <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "25" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalvinos, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalcerveza <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "26" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalcerveza, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalharina <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "19" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalharina, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If
If totalcarne <> 0 Then
lce = lce + "<TotOtrosImp>"
lce = lce + "<CodImp>" + "18" + "</CodImp>"
lce = lce + "<TotMntImp>" + Replace(totalcarne, ".", "") + "</TasaImp>"
lce = lce + "</TotOtrosImp>"

End If

If tipo = "43" Then
    'tabla liquidacion
    lce = lce & "<TotLiquidaciones>"
    lce = lce & "<TotValComNeto>" & Replace(netocom, ".", "") & "</TotValComNeto>"
    lce = lce & "<TotValComExe>0</TotValComExe>"
    lce = lce & "<TotValComIVA>" & Replace(ivacom, ".", "") & "</TotValComIVA>"
    lce = lce & "</TotLiquidaciones>"
End If

If tipo = "38" Then total = TOTALEXENTO
lce = lce + "<TotMntTotal>" + Replace(total, ".", "") + "</TotMntTotal>"
lce = lce + "</TotalesPeriodo>"

End Sub

Public Sub generacabezalibrodiario(NOMBRE, rutempresa, periodoinicial, periodofinal, moneda, rectificatoria)
Dim comi As String


comi = Chr(34)

lce = " "
lce = lce + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
lce = lce + "<LceDiario xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " xmlns:ds=" + comi + "http://www.w3.org/2000/09/xmldsig#" + comi + " "
lce = lce + " xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiDte LceDiario_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
lce = lce + "<LceDiarioRes xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " xmlns:ds=" + comi + "http://www.w3.org/2000/09/xmldsig#" + comi + " "
lce = lce + " xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiDte LceDiarioRes_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
lce = lce + "<DocumentoDiarioRes ID=" + comi + NOMBRE + comi + ">"
lce = lce + "<Identificacion>"
rutempresa = Format(Mid(rutempresa, 1, 8), "########") + "-" + Mid(rutempresa, 10, 1)
lce = lce + "<RutContribuyente>" + rutempresa + "</RutContribuyente>"
lce = lce + "<Inicial>" + periodoinicial + "</Inicial> "
lce = lce + "<Final>" + periodofinal + "</Final> "



End Sub

'funcion original comentada por rZ 05-12-2017
''
''Public Sub EnviarMail(ByRef Asunto, ByRef MENSAJE, ByRef Servidor, ByRef MailDestinatario, ByVal NombreDestinatario As String, ByRef ArchivAdjunto, ByVal archivadjunto2 As String)
''Dim enviados As String
''
''Set email = New clsSendMail
''Screen.MousePointer = vbHourglass
''Rem MailDestinatario = "rodrigogranadino@adminerp.cl"
''
''MailDestinatario = LCase(MailDestinatario)
''
'''MailDestinatario = "cesarsandoval@adminerp.cl"
''Call empresadte(empresaactiva)
'' With email
''
''     ' **************************************************************************
''        .SMTPHostValidation = VALIDATE_NONE
''        .EmailAddressValidation = VALIDATE_SYNTAX
''        .Delimiter = ";"
''     ' **************************************************************************
''        .AsHTML = False                             'ENVIAR COMO HTML O TEXTO PLANO
''        .SMTPHost = confi_servermail     'Servidor                        ' servidor smtp
''        .From = confi_mailsalida 'usuario                  'email de quien envia
''        .FromDisplayName = nombreempresa
''        .Recipient = MailDestinatario           'email para
''        .RecipientDisplayName = NombreDestinatario  'NOMBRE PARA
''        '.CcRecipient = usuario                           'EMAIL COPIA
''        .CcDisplayName = ""                         'NOMBRE COPIA
''        .BccRecipient = MailDestinatario                                     'EMAIL COPIA OCULTA
''        .ReplyToAddress = ""                        'responder a otro email
''        .Subject = Asunto                           'ASUNTO
''        .Message = MENSAJE 'txtemail.text                    'CUERPO MAIL
''
''If ArchivAdjunto <> "" And archivadjunto2 = "" Then enviados = ArchivAdjunto
''If archivadjunto2 <> "" And ArchivAdjunto = "" Then enviados = archivadjunto2
''If archivadjunto2 <> "" And ArchivAdjunto <> "" Then enviados = ArchivAdjunto + ";" + archivadjunto2
''
''If enviados <> "" Then .Attachment = enviados 'RUTA ARCHIVO ADJUNTO Trim(txtAttach.Text)
''
''     ' **************************************************************************
''        .ContentBase = ""
''        .EncodeType = 0 'CODIFICACION ARCHIVOS ADJUNTOS
''        .Priority = HIGH_PRIORITY                   'PRIORIDAD
''        .Receipt = False                            '
''        .UseAuthentication = True                   'servidor requiere autenticacion
''        .UsePopAuthentication = False               'usar autenticacion del pop
''        .UserName = confi_mailsalida 'usuario                         'usuario cuenta correo
''        .password = confi_clavemail    'CLAVE                           'clave cuenta correo
''        .POP3Host = confi_servermail 'Servidor                        'servidor pop
''        .MaxRecipients = 100                        '
''         Rem Call Sleep(10000)
''        .Send                                       'envia el correo
''
''    End With
''
''    Screen.MousePointer = vbDefault
''End Sub





Public Sub EnviarMail(ByRef Asunto, ByRef MENSAJE, ByRef Servidor, ByRef MailDestinatario, _
                        ByVal NombreDestinatario As String, ByRef ArchivAdjunto, _
                        ByVal archivadjunto2 As String)
Dim enviados As String
On Error GoTo error
Screen.MousePointer = vbHourglass
MailDestinatario = LCase(MailDestinatario)
Call empresadte(empresaactiva)
'rz 05-12-2017 '
'funcion adaptada para enviar con cuenta de eltit

If LeerCuentaAlternativa = True Then
        confi_mailsalida = email_cuenta_usuario
        confi_clavemail = email_cuenta_clave
        confi_servermail = email_cuenta_server
End If
 
If ArchivAdjunto <> "" And archivadjunto2 = "" Then enviados = ArchivAdjunto
If archivadjunto2 <> "" And ArchivAdjunto = "" Then enviados = archivadjunto2
If archivadjunto2 <> "" And ArchivAdjunto <> "" Then enviados = ArchivAdjunto + ";" + archivadjunto2



Dim iMsg As Object
Dim iConf As Object
Dim strbody  As String
Dim Flds As Variant
Dim comi As String
Dim puertosalida As Double
Dim destinatario As String
'destinatario = "aalarcon@eltit.cl; rlzurita@gmail.com"
 
comi = Chr(34)
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
destinatario = Replace(destinatario, "<", "")
destinatario = Replace(destinatario, ">", "")
iConf.Load -1
 puertosalida = 465

Set Flds = iConf.Fields
With Flds
    .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = email_cuenta_usuario
    .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = email_cuenta_clave
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = email_cuenta_server
    .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = puertosalida
    .Update
End With

With iMsg
Set .Configuration = iConf
.To = MailDestinatario
.cc = ""
.BCC = "" 'aalarcon@eltit.cl; raulzurita@adminerp.cl"
 
.From = comi & comi & comi & nombreempresa & comi & " <" & email_cuenta_usuario & ">"
.Subject = Asunto
.TextBody = MENSAJE

If enviados <> "" Then .AddAttachment enviados


.Send

End With


Call grabar_envio_correo("", email_cuenta_usuario, destinatario)
Exit Sub


error:
MsgBox "NO SE PUDO ENVIAR EL CORREO" & vbNewLine & err.Description
Screen.MousePointer = vbDefault
End Sub


Public Sub empresadte(empresa)
   On Error GoTo PASA:
   Dim op As Integer
   Dim k As Integer
   Dim campos(20, 10) As String
   
    campos(0, 0) = "nombre"
    campos(1, 0) = "direccion"
    campos(2, 0) = "comuna"
    campos(3, 0) = "ciudad"
    campos(4, 0) = "rut"
    campos(5, 0) = "girodte"
    campos(6, 0) = "codactividadeconomica"
    campos(7, 0) = "certificado"
    campos(8, 0) = "rutenviasii"
    campos(9, 0) = "clave_certificado"
    campos(10, 0) = "fecharesolucion"
    campos(11, 0) = "numeroresolucion"
    campos(12, 0) = "servermail"
    campos(13, 0) = "mailsalida"
    campos(14, 0) = "clavemail"
    campos(15, 0) = "empresafae"
    campos(16, 0) = ""
    campos(0, 2) = clientesistema + "conta.maestroempresas"
  
    condicion = "codigoempresa = '" + empresa + "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    dte_e_nombre = sqlconta.response(0, 3)
    dte_e_direccion = sqlconta.response(1, 3)
    dte_e_comuna = sqlconta.response(2, 3)
    dte_e_ciudad = sqlconta.response(3, 3)
    dte_e_rut = sqlconta.response(4, 3)
    dte_e_giro = sqlconta.response(5, 3)
    dte_e_acti = sqlconta.response(6, 3)
    certificado = sqlconta.response(7, 3)
    dte_rutenvia = sqlconta.response(8, 3)
    clavecertificado = sqlconta.response(9, 3)
    fechacertificacion = sqlconta.response(10, 3)
    resolucion = sqlconta.response(11, 3)
    confi_servermail = sqlconta.response(12, 3)
    confi_mailsalida = sqlconta.response(13, 3)
    confi_clavemail = sqlconta.response(14, 3)
    confi_localempresa = sqlconta.response(15, 3)
    
    
    End If
Exit Sub
PASA:
    MsgBox "empresa  no es electronica"
End Sub

