Attribute VB_Name = "sqlutil"
Option Explicit
    Public datos As Variant
    Public conexion As rdoConnection
    Public conAuditoria As rdoConnection
    Public estado As Integer
    Public area As String
    Public area2 As String
    Public rubro_trabajo As String
    Public usuarioauditoria As String
    Public codigolocal As String
    Public campos1 As String
    Public campos2 As String
    Public audit As Boolean
'==================================================================================
'          RUTINAS OPERACIONES CON MYSQL
'==================================================================================
    
    Public Sub SQLUTIL(ByVal Opcion As Integer, ByVal condicion As String)
    
        Dim cadena As String
        Dim i As Integer
        Dim resultados As rdoResultset
        Dim cSql As rdoQuery
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = conexion

        Select Case Opcion
        
        Case 2:    '<<<<<   INSERTA   >>>>>>
            
                    cadena = "INSERT INTO " & datos(0, 2) & " ("
                    campos1 = ""
                    'NOMBRE DE datos.
                    i = 0
                    While datos(i, 0) <> ""
                        cadena = cadena & datos(i, 0) + ","
                        campos1 = campos1 & "[" & datos(i, 0) & "]"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") VALUES ("
                    
                    'VALORES ASIGNADOS A CADA CAMPO.
                    i = 0
                    While datos(i, 0) <> ""
                        cadena = cadena & "'" & datos(i, 1) & "',"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") "
                    cadena = cadena & "ON DUPLICATE KEY UPDATE " & datos(0, 0) & " = " & datos(0, 0)
                    cSql.sql = cadena
                    cSql.Execute
                    If audit = True Then
                        Call Auditoria(Opcion, condicion)
                    End If
        Case 3:    '<<<<<   ACTUALIZA   >>>>>>
                    
                    cadena = "UPDATE " & datos(0, 2) & " SET "
                    i = 0
                    campos1 = ""
                    campos2 = ""
                    While datos(i, 0) <> ""
                        cadena = cadena & datos(i, 0) & "= '" & datos(i, 1) & "',"
                        campos1 = campos1 & "[" & datos(i, 0) & "]"
                        campos2 = campos2 & "[" & datos(i, 1) & "]"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " WHERE " & condicion
                    cSql.sql = cadena
                    If audit = True Then
                        Call Auditoria(Opcion, condicion)
                    End If
                    cSql.Execute
                    
            
        Case 4:    '<<<<<   ELIMINA   >>>>>>
                    If audit = True Then
                        Call Auditoria(Opcion, condicion)
                    End If
                    cadena = "DELETE FROM " & datos(0, 2) & " WHERE " & condicion
                    cSql.sql = cadena
                    cSql.Execute
        
        Case 5:    '<<<<<   LEE   >>>>>>
                    cadena = "SELECT "
                    i = 0
                    While datos(i, 0) <> ""
                        cadena = cadena & datos(i, 0) & ","
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " FROM " & datos(0, 2) & " WHERE " & condicion
                    cSql.sql = cadena
                    cSql.Execute
                    If cSql.RowsAffected > 0 Then
                        Set resultados = cSql.OpenResultset
                        estado = 0 'SI ENCONTRO DATOS.
                        i = 0
                        'TRASPASA LOS DATOS A LA COLUMNA 3 DE LA MATRIZ GLOBAL "datos".
                        While datos(i, 0) <> ""
                            If IsNull(resultados(0)) = True Then
                                datos(i, 3) = "0"
                            Else
                                datos(i, 3) = resultados(i)
                            End If
                            i = i + 1
                        Wend
                        resultados.Close
                        Set resultados = Nothing
                    Else
                        estado = 4 'NO ENCONTRO DATOS.
                    End If
        End Select
        cSql.Close
        Set cSql = Nothing
    End Sub
    
    
    Sub Auditoria(ByVal Evento As Integer, ByVal condicion As String)
        Dim cadena As String
        Dim resultados2 As rdoResultset
        Dim cSql2 As New rdoQuery
        Dim resultados As rdoResultset
        Dim cSql As New rdoQuery
        Dim columnas As String
        Dim registros As String
        Dim i As Integer
        Dim j As Integer
        
        Set cSql2.ActiveConnection = conAuditoria
        Set cSql.ActiveConnection = conexion
        
        If Left(datos(0, 2), 1) = "g" Then
            area2 = area
        Else
            If Left(datos(0, 2), 2) = "sv" Then
                cadena = conexion.Connect
                cadena = Left(cadena, InStr(1, cadena, ";") - 1)
                If Right(cadena, 2) = "as" Then
                    area2 = "ventas"
                Else
                    area2 = "ventas" & rubro_trabajo
                End If
            Else
                area2 = area & rubro_trabajo
            End If
        End If
        Select Case Evento
        
                Case 2: '<<<<<   AGREGA   >>>>>>
                
                        cadena = "INSERT INTO auditoria" & area & " ("
                        cadena = cadena + "usuario,evento,local,tabla,camposoriginales,datosoriginales) VALUES ( "
                        cadena = cadena & "'" & usuarioauditoria & "','" & Evento & "','" & codigolocal & "','" & datos(0, 2) & "','" & campos1 & "','" & ""
                        'VALORES ASIGNADOS A CADA CAMPO.
                        i = 0
                        While datos(i, 0) <> ""
                            cadena = cadena & "[" & datos(i, 1) & "]"
                            i = i + 1
                        Wend
                        cadena = cadena & "')"
                        cSql2.sql = cadena
                        cSql2.Execute
                        
                Case 3:  '<<<<<   ACTUALIZA   >>>>>>
                        'RESCATA LOS CAMPOS DE LA TABLA EN CUESTION DE LA BD DETERMINADA POR "area".
                        cSql.sql = "SHOW COLUMNS FROM " & datos(0, 2) & " FROM " & area2
                        cSql.Execute
                        If cSql.RowsAffected > 0 Then
                            Set resultados = cSql.OpenResultset
                            i = 0
                            While Not resultados.EOF
                                columnas = columnas & "[" & resultados(0) & "]"
                                i = i + 1
                                resultados.MoveNext
                            Wend
                            resultados.Close
                            Set resultados = Nothing
                        End If
                        'RESCATA LOS DATOS DE LOS CAMPOS A ACTUALIZAR.
                        cSql.sql = "SELECT * FROM " & datos(0, 2) & " WHERE " & condicion
                        cSql.Execute
                        If cSql.RowsAffected > 0 Then
                            Set resultados = cSql.OpenResultset
                            j = 0
                            While Not resultados.EOF
                                For j = 0 To i - 1
                                    registros = registros & "[" & resultados(j) & "]"
                                Next j
                                resultados.MoveNext
                            Wend
                            resultados.Close
                            Set resultados = Nothing
                        End If
                        
                        cadena = "INSERT INTO auditoria" & area & " ("
                        cadena = cadena + "usuario,evento,local,tabla,camposoriginales,datosoriginales,camposmodificados,datosmodificados) VALUES ("
                        cadena = cadena & "'" & usuarioauditoria & "','" & Evento & "','" & codigolocal & "','" & datos(0, 2) & "','" & columnas & "','" & registros & "','" & campos1 & "','" & campos2 & "')"
                        cSql2.sql = cadena
                        cSql2.Execute
                        
                Case 4:  '<<<<<   ELIMINA   >>>>>>
                        'RESCATA LOS CAMPOS DE LA TABLA EN CUESTION.
                        cSql.sql = "SHOW COLUMNS FROM " & datos(0, 2) & " FROM " & area2
                        cSql.Execute
                        If cSql.RowsAffected > 0 Then
                            Set resultados = cSql.OpenResultset
                            i = 0
                            While Not resultados.EOF
                                columnas = columnas & "[" & resultados(0) & "]"
                                i = i + 1
                                resultados.MoveNext
                            Wend
                            resultados.Close
                            Set resultados = Nothing
                        End If
                        'RESCATA LOS DATOS DE LOS CAMPOS A ELIMINAR.
                        cSql.sql = "SELECT * FROM " & datos(0, 2) & " WHERE " & condicion
                        cSql.Execute
                        If cSql.RowsAffected > 0 Then
                            Set resultados = cSql.OpenResultset
                            j = 0
                            While Not resultados.EOF
                                registros = ""
                                For j = 0 To i - 1
                                    registros = registros & "[" & resultados(j) & "]"
                                Next j
                                cadena = "INSERT INTO auditoria" & area & " ("
                                cadena = cadena + "usuario,evento,local,tabla,camposoriginales,datosoriginales) VALUES ("
                                cadena = cadena & "'" & usuarioauditoria & "','" & Evento & "','" & codigolocal & "','" & datos(0, 2) & "','" & columnas & "','" & registros & "')"
                                cSql2.sql = cadena
                                cSql2.Execute
                                resultados.MoveNext
                            Wend
                            resultados.Close
                            Set resultados = Nothing
                        End If
        End Select
    End Sub






