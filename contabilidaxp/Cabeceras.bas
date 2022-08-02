Attribute VB_Name = "Cabeceras"
'====================================================================================
'                            CABECERAS GRILLAS
'====================================================================================
        
    '=========================================================================
    'Cabeceras del formulario de ventas
    '=========================================================================
        Sub ventas()
            With ventas04
                t$ = ">NL  |>CODIGO     |<DESCRIPCION          |>UNIDADES|>PRECIO          |>TOTAL               "
                .listaProductos.FormatString = t$
            End With
        End Sub

    '=========================================================================
    '           CABECERAS PAGO CLIENTES
    '=========================================================================
        Sub Cabeceras_Pago_Clientes()
            With ventas06
                t$ = ">TIPO|<DOCUMENTO                     |>Nº                        |>MONTO                       "
                .DocumentosPendientes.FormatString = t$
                .DocumentosCancelar.FormatString = t$
                t$ = "<BANCO                                                            |>NUMERO            |>MONTO                   |>FECHA         "
                .Cheques.FormatString = t$
            End With
        End Sub


