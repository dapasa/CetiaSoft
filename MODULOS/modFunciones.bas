Attribute VB_Name = "modFunciones"
Public Sub CERRAR_TABLA(tabla As ADODB.Recordset)
    On Error Resume Next
    If tabla.State = adStateOpen Then
        If tabla.State = adEditInProgress Then
            tabla.Update
        End If
        tabla.Close
    End If
End Sub

Public Sub CERRAR_TODO()
    CERRAR_TABLA adoAlumnos
    CERRAR_TABLA adoAlumnosXCurso
    CERRAR_TABLA adoAulas
    CERRAR_TABLA adoComoLlego
    CERRAR_TABLA adoCompaniasCelular
    CERRAR_TABLA adoCondIva
    CERRAR_TABLA adoCursos
    CERRAR_TABLA adoCursosXProfesor
    CERRAR_TABLA adoDuraciones
    CERRAR_TABLA adoHorarios
    CERRAR_TABLA adoItemsXMov
    CERRAR_TABLA adoListaEspera
    CERRAR_TABLA adoLocalidades
    CERRAR_TABLA adoMovimientos
    CERRAR_TABLA adoProfesores
    CERRAR_TABLA adoTiposCurso
    CERRAR_TABLA adoTiposDoc
    CERRAR_TABLA adoUsuarios
    
    CERRAR_TABLA adoTabla
    
    'adoConnection.Close
    
    'Set adoConnection = Nothing
    
    'Set adoAlumnos = Nothing
    'Set adoAlumnosXCurso = Nothing
    'Set adoAulas = Nothing
    'Set adoComoLlego = Nothing
    'Set adoCompaniasCelular = Nothing
    'Set adoCondIva = Nothing
    'Set adoCursos = Nothing
    'Set adoCursosXProfesor = Nothing
    'Set adoDuracion = Nothing
    'Set adoHorarios = Nothing
    'Set adoItemsXMov = Nothing
    'Set adoListaEspera = Nothing
    'Set adoLocalidades = Nothing
    'Set adoModalidades = Nothing
    'Set adoMovimientos = Nothing
    'Set adoProfesores = Nothing
    'Set adoTiposCurso = Nothing
    'Set adoTiposDoc = Nothing
    'Set adoUsuarios = Nothing
    
    'Set adoTabla = Nothing
    
End Sub

Public Function ENCRIPTAR(Clave As String) As String
    ENCRIPTAR = Clave
End Function

Public Function DESENCRIPTAR(Clave As String) As String
    DESENCRIPTAR = Clave
End Function

Public Sub CARGAR_COMBO(NombreCombo As String, tabla As ADODB.Recordset, NombreTabla As String, NombreCampo As String, Formulario As Form)
    CERRAR_TABLA tabla
    sSql = "SELECT " & NombreCampo & " FROM " & NombreTabla
    tabla.Open sSql, adoConnection, adOpenStatic, adLockOptimistic
    
    If Not tabla.EOF Then
        Do While Not tabla.EOF
            Formulario.Controls(NombreCombo).AddItem tabla.Fields(NombreCampo).Value
            tabla.MoveNext
        Loop
        
        Formulario.Controls(NombreCombo).AddItem "(No disponible)"
        
        Formulario.Controls(NombreCombo).ListIndex = 0
    End If
    
    tabla.Close
End Sub

Public Function DEVOLVER_ID(dato As String, tabla As ADODB.Recordset, NombreTabla As String, NombreCampo As String) As Long
    CERRAR_TABLA tabla
    sSql = "SELECT id FROM " & NombreTabla & " WHERE " & NombreCampo & " = '" & dato & "'"
    tabla.Open sSql, adoConnection, adOpenStatic, adLockOptimistic
    
    If Not tabla.EOF Then
        DEVOLVER_ID = tabla!id
    Else
        DEVOLVER_ID = 0
    End If
    
    tabla.Close
End Function

Public Function DEVOLVER_CAMPO(id As Long, tabla As ADODB.Recordset, NombreTabla As String, NombreCampo As String) As String
    CERRAR_TABLA tabla
    sSql = "SELECT " & NombreCampo & " FROM " & NombreTabla & " WHERE id = " & id
    tabla.Open sSql, adoConnection, adOpenStatic, adLockOptimistic
    
    If Not tabla.EOF Then
        DEVOLVER_CAMPO = tabla.Fields(NombreCampo).Value
    Else
        DEVOLVER_CAMPO = "(No disponible)"
    End If
    
    tabla.Close
End Function

Public Sub BOTONES(Boton As String, Formulario As Form)
    With Formulario
        Select Case Boton
            Case "Nuevo"
                .cmdNuevo.Enabled = False
                .cmdBuscar.Enabled = True
                .cmdEliminar.Enabled = False
                .cmdGuardar.Enabled = True
            Case "Buscar"
                .cmdNuevo.Enabled = True
                .cmdBuscar.Enabled = True
                .cmdEliminar.Enabled = True
                .cmdGuardar.Enabled = True
            Case "Eliminar"
                .cmdNuevo.Enabled = True
                .cmdBuscar.Enabled = True
                .cmdEliminar.Enabled = False
                .cmdGuardar.Enabled = False
            Case "Guardar"
                .cmdNuevo.Enabled = True
                .cmdBuscar.Enabled = True
                .cmdEliminar.Enabled = False
                .cmdGuardar.Enabled = False
        End Select
    End With
End Sub

Public Sub ENABLED_TODO(Estado As Boolean, Formulario As Form)
    On Error Resume Next
    
    With Formulario
        For k = 0 To .Controls.Count - 1
            If Left(.Controls(k).Name, 3) <> "cmd" Then
                .Controls(k).Enabled = Estado
            End If
        Next
    End With
End Sub

Public Sub PASAR_CAMPO(Formulario As Form)
    Dim Saltar As Boolean
    
    Saltar = True
    If Left(Formulario.ActiveControl.Name, 3) = "txt" Then
        If Formulario.ActiveControl.MultiLine = True Then
            Saltar = False
        End If
    End If
    If Saltar Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Public Function ULTIMO_NUMERO(Docu As String) As Integer
    CERRAR_TABLA adoNumeracion
    sSql = "SELECT * FROM Numeracion"
    adoNumeracion.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    ULTIMO_NUMERO = adoNumeracion.Fields(Docu).Value
    
    adoNumeracion.Fields(Docu).Value = adoNumeracion.Fields(Docu).Value + 1
    adoNumeracion.Update
    
    adoNumeracion.Close
End Function

Public Sub NUMERO_ATRAS(Docu As String)
    CERRAR_TABLA adoNumeracion
    sSql = "SELECT * FROM Numeracion"
    adoNumeracion.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    adoNumeracion.Fields(Docu).Value = adoNumeracion.Fields(Docu).Value - 1
    adoNumeracion.Update
    
    adoNumeracion.Close
End Sub

Public Sub IMPRIMIR(Fila As Byte, Columna As Byte, dato As String)
    Printer.CurrentY = Fila
    Printer.CurrentX = Columna
    Printer.Print dato
End Sub

Public Function PUEDE_BORRAR(tabla As String, x As Long) As Boolean
    sSql = "SELECT Usado FROM " & tabla & " WHERE id = " & x
    CERRAR_TABLA adoTemp
    adoTemp.Open sSql, adoConnection, adOpenStatic, adLockOptimistic
    If adoTemp!Usado Then
        MsgBox "No es posible eliminar este dato." & vbCrLf & "Se perdería la consistencia.", vbCritical, "CONSISTENCIA DE DATOS"
        PUEDE_BORRAR = False
    Else
        PUEDE_BORRAR = True
    End If
End Function

Public Sub GENERAR_LISTADO()
    Dim x As Integer
    caption_anterior = frmPrincipal.Caption
    For k = 1 To 3000
        frmPrincipal.Caption = k
    Next
    frmPrincipal.Caption = caption_anterior
End Sub

Public Sub CALCULAR_ACUMULADO_MENSUAL()
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(SubTotal) AS t FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NombreEmisor = '1-Fabru S.A.' AND Month(Fecha) = " & Month(Date) & " AND Year(Fecha) = " & Year(Date)
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS FABRU S.A. " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(SubTotal) AS t FROM Movimientos WHERE TipoDoc = 'REC' AND NombreEmisor = '1-Fabru S.A.' AND Month(Fecha) = " & Month(Date) & " AND Year(Fecha) = " & Year(Date)
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL RECIBOS FABRU S.A. " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NombreEmisor = '2-Sergio Fasano' AND Month(Fecha) = " & Month(Date) & " AND Year(Fecha) = " & Year(Date)
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS SERGIO FASANO " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE TipoDoc = 'REC' AND NombreEmisor = '2-Sergio Fasano' AND Month(Fecha) = " & Month(Date) & " AND Year(Fecha) = " & Year(Date)
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL RECIBOS SERGIO FASANO " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NombreEmisor = '3-Néstor Russaz' AND Month(Fecha) = " & Month(Date) & " AND Year(Fecha) = " & Year(Date)
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL FACTURAS NESTOR RUSSAZ " & Space(5) & adoTabla!t
    adoTabla.Close
    
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Total) AS t FROM Movimientos WHERE TipoDoc = 'REC' AND NombreEmisor = '3-Néstor Russaz' AND Month(Fecha) = " & Month(Date) & " AND Year(Fecha) = " & Year(Date)
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    mensaje = mensaje & vbCrLf & "TOTAL RECIBOS NESTOR RUSSAZ " & Space(5) & adoTabla!t
    adoTabla.Close
    
    MsgBox mensaje, , "ACUMULADO MENSUAL"
End Sub

Public Sub CHEQUEAR_A_CUENTA()
    sSql = "SELECT Movimientos.idAlumno, Movimientos.idCurso, ItemsXMov.Importe " & _
           "FROM Movimientos, ItemsXMov " & _
           "WHERE Movimientos.id = ItemsXMov.idMovimiento AND Movimientos.idAlumno =  " & id_Alumno & _
           "  AND Left(ItemsXMov.Detalle, 9) = ' A CUENTA' AND Movimientos.Saldo <> 0"
    
    CERRAR_TABLA adoTemp
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoTemp.EOF Then
        total_a_cuenta = 0
        adoTemp.MoveFirst
        Do While Not adoTemp.EOF
            total_a_cuenta = total_a_cuenta + adoTemp!Importe
            adoTemp.MoveNext
        Loop
        adoTemp.Close
        
        'total_a_cuenta = total_a_cuenta
        
        If MsgBox("Alumno con saldo a favor." & vbCrLf & "$" & total_a_cuenta & ".-" & vbCrLf & "¿Cancelar automáticamente?", vbYesNo + vbInformation, "ALUMNO CON SALDO A FAVOR") = vbYes Then
            
            sSql = "SELECT * FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso & " AND TipoDoc = 'MOD' ORDER BY Cuota"
            adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            adoTemp.MoveFirst
            
            Do While Not adoTemp.EOF
                If total_a_cuenta >= adoTemp!Saldo Then
                    total_a_cuenta = total_a_cuenta - adoTemp!Saldo
                    adoTemp!Paga = adoTemp!Saldo
                    adoTemp!Saldo = 0
                    adoTemp.Update
                    
                    sSql = "UPDATE Movimientos SET Saldo = Saldo - " & adoTemp!Saldo & " WHERE id = " & adoTemp!id
                    adoConnection.Execute sSql
                    
                    sSql = "UPDATE ItemsXMov SET Saldo = 0 WHERE idMovimiento = " & adoTemp!id
                    adoConnection.Execute sSql
                Else
                    adoTemp!Paga = adoTemp!Saldo - total_a_cuenta
                    adoTemp!Saldo = adoTemp!Saldo - total_a_cuenta
                    adoTemp.Update
                    
                    sSql = "UPDATE Movimientos SET Saldo = 0 WHERE id = " & adoTemp!id
                    adoConnection.Execute sSql
                    
                    sSql = "UPDATE ItemsXMov SET Saldo = Saldo - " & total_a_cuenta & " WHERE idMovimiento = " & adoTemp!id
                    adoConnection.Execute sSql
                
                    total_a_cuenta = 0
                End If
                
                adoTemp.MoveNext
            Loop
            
        End If
    End If
   
    CERRAR_TABLA adoTemp
End Sub


Public Sub CHEQUEAR_A_CUENTA_FACTURA()
    sSql = "SELECT Movimientos.idAlumno, Movimientos.idCurso, ItemsXMov.Importe " & _
           "FROM Movimientos, ItemsXMov " & _
           "WHERE Movimientos.id = ItemsXMov.idMovimiento AND Movimientos.idAlumno =  " & id_Alumno & _
           "  AND Left(ItemsXMov.Detalle, 9) = ' A CUENTA' AND ItemsXMov.Saldo <> 0"
    
    CERRAR_TABLA adoTemp
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoTemp.EOF Then
        total_a_cuenta = 0
        adoTemp.MoveFirst
        Do While Not adoTemp.EOF
            total_a_cuenta = total_a_cuenta + adoTemp!Importe
            adoTemp.MoveNext
        Loop
        adoTemp.Close
        
        total_a_cuenta = total_a_cuenta
        
        If MsgBox("Alumno con saldo a favor." & vbCrLf & "$" & total_a_cuenta & ".-" & vbCrLf & "¿Cancelar automáticamente?", vbYesNo + vbInformation, "ALUMNO CON SALDO A FAVOR") = vbYes Then
            
            'sSql = "SELECT * FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso & " AND TipoDoc = 'MOD' ORDER BY Cuota"
            'adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            frmFactura.lblTotalFactura.Caption = total_a_cuenta
            
            adoTempFactura.MoveFirst
            
            pago_a_cuenta = 0
            Do While Not adoTempFactura.EOF
                If total_a_cuenta >= adoTempFactura!Saldo Then
                    pago_a_cuenta = pago_a_cuenta + adoTempFactura!Saldo
                    total_a_cuenta = total_a_cuenta - adoTempFactura!Saldo
                    adoTempFactura!Paga = adoTempFactura!Saldo
                    adoTempFactura!Saldo = 0
                    adoTempFactura.Update
                    
                    sSql = "UPDATE ItemsXMov SET Saldo = 0 WHERE id = " & adoTempFactura!idItem
                    adoConnection.Execute sSql
                
                    'frmFactura.lblTotalFactura.Caption = frmFactura.lblTotalFactura.Caption + adoTempFactura!Saldo
                Else
                    pago_a_cuenta = pago_a_cuenta + total_a_cuenta
                    adoTempFactura!Paga = total_a_cuenta
                    adoTempFactura!Saldo = adoTempFactura!Saldo - total_a_cuenta
                    adoTempFactura.Update
                    
                    sSql = "UPDATE ItemsXMov SET Saldo = Saldo - " & total_a_cuenta & " WHERE id = " & adoTempFactura!idItem
                    adoConnection.Execute sSql
                
                    'frmFactura.lblTotalFactura.Caption = frmFactura.lblTotalFactura.Caption + adoTempFactura!Saldo
                    
                    total_a_cuenta = 0
                End If
                
                adoTempFactura.MoveNext
            Loop
        
        '*********
        sSql = "SELECT ItemsXMov.id, Movimientos.idAlumno, Movimientos.idCurso, ItemsXMov.Saldo " & _
           "FROM Movimientos, ItemsXMov " & _
           "WHERE Movimientos.id = ItemsXMov.idMovimiento AND Movimientos.idAlumno =  " & id_Alumno & _
           "  AND Left(ItemsXMov.Detalle, 9) = ' A CUENTA'"
    
        CERRAR_TABLA adoTemp
        adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        adoTemp.MoveFirst
        Do While Not adoTemp.EOF
            If pago_a_cuenta >= adoTemp!Saldo Then
                pago_a_cuenta = pago_a_cuenta - adoTemp!Saldo
                
                sSql = "UPDATE ItemsXMov SET Saldo = 0 WHERE id = " & adoTemp!id
                adoConnection.Execute sSql
            Else
                sSql = "UPDATE ItemsXMov SET Saldo = Saldo - " & pago_a_cuenta
                adoConnection.Execute sSql
                pago_a_cuenta = 0
            End If

            adoTemp.MoveNext
        Loop
        adoTemp.Close

        '*********
        
        End If
    End If
   
    'CERRAR_TABLA adoTemp
End Sub


Public Sub GUARDAR_LOG(usuario As String, fecha As String, hora As String, detalle As String)
    CERRAR_TABLA adoLog
    
    With adoLog
        .Open "Log", adoConnection, adOpenDynamic, adLockOptimistic
    
        .AddNew
        
        !usuario = usuario
        !fecha = fecha
        !hora = hora
        !detalle = detalle
        
        .Update
        
        .Close
    End With
End Sub

