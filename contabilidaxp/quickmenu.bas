Attribute VB_Name = "quickmenu"
 

Public Function HayFavoritos() As Boolean
   Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim Aplicacion As Form
    Set csql.ActiveConnection = contadb
    csql.sql = "select sistema,glosa,aplicacion from " & clientesistema & "menu.quick_menu where sistema ='" & LCase(App.EXEName) & "' and usuario = '" & USUARIOSISTEMA & "' and activado = '1' order by sistema"
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        HayFavoritos = True
        Set Aplicacion = Forms.Add(resultados(2))
        Aplicacion.Show ' vbModal
        Call DesactivaFavorito(USUARIOSISTEMA, App.EXEName, resultados(1))
    Else
        HayFavoritos = False
    End If
End Function
Public Sub DesactivaFavorito(Usuario, sistema, programa)
Dim csql1 As New rdoQuery
        Set csql1.ActiveConnection = contadb
        csql1.sql = "update " & clientesistema & "menu.quick_menu set activado='0' where usuario = '" & Usuario & "' and sistema = '" & sistema
        csql1.sql = csql1.sql & "' and glosa= '" & programa & "'"
        csql1.Execute
'        Call SincronizaDatos(csql1.sql, conta)
csql1.Close
End Sub

