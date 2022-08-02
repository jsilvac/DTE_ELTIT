Attribute VB_Name = "Funciones_Forms_Venta"
Public Sub limpiaProducto()
    With ventas04
        .txtProducto(0).text = ""
        .txtProducto(1).text = ""
        .txtProducto(2).text = ""
        .lblProducto(0).Caption = ""
        .lblProducto(1).Caption = ""
        .lblProducto(2).Caption = ""
        .txtProducto(0).SetFocus
    End With
End Sub
