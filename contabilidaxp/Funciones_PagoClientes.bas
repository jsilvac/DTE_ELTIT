Attribute VB_Name = "Funciones_PagoClientes"
Sub Carga_Documentos_Pendientes()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With ventas06
        Set cSql.ActiveConnection = db
        rut = .dato(6).text
        cSql.SQL = "SELECT tipo,numero,total "
        cSql.SQL = cSql.SQL + "FROM datodocumento "
        cSql.SQL = cSql.SQL + "WHERE rut='" & rut & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            .DocumentosPendientes.Clear
            .DocumentosPendientes.Rows = 2
            Cabeceras.Cabeceras_Pago_Clientes
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .DocumentosPendientes.AddItem resultados(0) & vbTab & resultados(1) & vbTab & resultados(2), .DocumentosPendientes.Rows - 1
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            '.DocumentosPendientes.Rows = .DocumentosPendientes.Rows - 1
        End If
    End With

End Sub

Sub Graba_Pago_Cliente()
    Dim op As Integer
    Dim i As Integer
    Dim condicion As String
    
        With ventas06
        '================================================================
        '         GRABA EN TABLA "COMPROBANTEDEPAGO"
        '================================================================
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'NUMERO PAGO
            campos(1, 0) = .dato(1).Tag 'FECHA
            campos(2, 0) = .dato(4).Tag 'TIPO PAGO
            campos(3, 0) = .dato(5).Tag 'MONTO
            campos(4, 0) = .dato(6).Tag 'RUT
            campos(5, 0) = .lbl(3).Tag 'BANCO DEPOSITO
            campos(6, 0) = .dato(8).Tag 'DATOS DEPOSITO
            
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).text 'NUMERO PAGO
            campos(1, 1) = Format(.dato(1).text & "-" & .dato(2).text & "-" & .dato(3).text, "yyyy-mm-dd") 'FECHA
            campos(2, 1) = .dato(4).text 'TIPO PAGO
            campos(3, 1) = .dato(5).text 'MONTO
            campos(4, 1) = .dato(6).text 'RUT
            campos(5, 1) = .lbl(3).Caption 'BANCO DEPOSITO
            campos(6, 1) = .dato(8).text 'DATOS DEPOSITO
            
            'NOMBRE DE TABLA
            campos(0, 2) = "comprobantedepago"
            'OPCION CON QUE SE LLAMA LA FUNCION (2=INSERTA).
            op = 2
            SQLUTIL.datos = campos
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, "0")
        '================================================================
        End With
            
            
'==========================================================================================================

'           RUTINAS SIGUIENTES AUN NO ESTAN PROBADAS

'==========================================================================================================

        With ventas06
        '================================================================
        '         GRABA EN TABLA "CARTERACHEQUES"
        '================================================================
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = rut 'RUT
            campos(1, 0) = numerocheque 'NUMERO CHEQUE
            campos(2, 0) = banco 'BANCO
            campos(3, 0) = monto 'MONTO
            campos(4, 0) = fechavencimiento 'FECHA VENCIMIENTO
            
            'NOMBRE DE DATOS EN PROGRAMA
            i = 1
            While i <> .Cheques.Rows - 1
                campos(0, 1) = .dato(6).text 'RUT
                campos(1, 1) = .Cheques.TextMatrix(i, 1) 'NUMERO CHEQUE
                campos(2, 1) = .Cheques.TextMatrix(i, 0) 'BANCO
                campos(3, 1) = .Cheques.TextMatrix(i, 2) 'MONTO
                campos(4, 1) = .Cheques.TextMatrix(i, 3) 'FECHA VENCIMIENTO
                
                'NOMBRE DE TABLA
                campos(0, 2) = "carteracheques"
                'OPCION CON QUE SE LLAMA LA FUNCION (2=INSERTA).
                op = 2
                SQLUTIL.datos = campos
                Set SQLUTIL.conexion = db
                Call SQLUTIL.SQLUTIL(op, condicion)
                i = i + 1
            Wend
        '================================================================
        End With
            
        With ventas06
        '================================================================
        '         ACTUALIZA EN TABLA "DATODOCUMENTO"
        '================================================================
            'NOMBRES DE CAMPOS EN TABLA
            'campos(0, 0) = rut 'RUT
            campos(1, 0) = numerocheque 'NUMERO CHEQUE
            campos(2, 0) = banco 'BANCO
            campos(3, 0) = monto 'MONTO
            campos(4, 0) = fechavencimiento 'FECHA VENCIMIENTO
            
            'NOMBRE DE DATOS EN PROGRAMA
            i = 1
            While i <> .Cheques.Rows - 1
                'campos(0, 1) = .dato(6).text 'RUT
                campos(1, 1) = .Cheques.TextMatrix(i, 1) 'NUMERO CHEQUE
                campos(2, 1) = .Cheques.TextMatrix(i, 0) 'BANCO
                campos(3, 1) = .Cheques.TextMatrix(i, 2) 'MONTO
                campos(4, 1) = .Cheques.TextMatrix(i, 3) 'FECHA VENCIMIENTO
                
                'NOMBRE DE TABLA
                campos(0, 2) = "datodocumento"
                'OPCION CON QUE SE LLAMA LA FUNCION (3=ACTUALIZA).
                op = 3
                'CONDICION DEL SQL
                
'       =================================================================
'       ========================    D U D A    ==========================
'       =================================================================

                condicion = "rut=" + "'" + .dato(6).text + "'" + " AND " + "numero=" + "'" + codigo_cuenta + "'"
                SQLUTIL.datos = campos
                Set SQLUTIL.conexion = db
                Call SQLUTIL.SQLUTIL(op, "0")
                i = i + 1
            Wend
        '================================================================
        End With
            
            
End Sub

Sub Agrega_Cheque()
'AGREGA EL CHEQUE INGRESADO A LA GRILLA CORRESPONDIENTE.
    Dim banco As String
    Dim numero_cheque As String
    Dim monto As String
    Dim fecha_cheque As Date
    Dim i As Integer

    With ventas06
        banco = .dato(9).text & Space(8) & .lbl(4).Caption
        numero_cheque = .dato(10).text
        monto = .dato(11).text
        fecha_cheque = Format(.dato(12) & "-" & .dato(13) & "-" & .dato(14), "yyyy-mm-dd")
        .Cheques.AddItem banco & vbTab & numero_cheque & vbTab & monto & vbTab & fecha_cheque, .Cheques.Rows - 1
        For i = 9 To 14
            .dato(i).text = ""
        Next i
        .lbl(4).Caption = ""
        .Cheques.Enabled = False
        .Cheques.Enabled = True
        .dato(9).SetFocus
    End With
End Sub
