VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBusDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUSCADOR DE DOCUMENTOS"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   Icon            =   "frmBusDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDato 
      Caption         =   "Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtDato 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   5880
      Picture         =   "frmBusDoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   6600
      Picture         =   "frmBusDoc.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancelar"
      Top             =   4440
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dbgTempDocus 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSeleccione 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione un documento y haga click en Aceptar."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3630
   End
End
Attribute VB_Name = "frmBusDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    Select Case TipoBusDoc
        Case "BorrarModelo"
            If MsgBox("¿Confirma que desea eliminar el MOD Nº" & adoTabla!NumDoc & " del alumno " & DEVOLVER_CAMPO(adoTabla!idAlumno, adoAlumnos, "Alumnos", "Nombre") & "?", vbQuestion + vbYesNo, "ELIMINAR MODELO") = vbYes Then
                sSql = "DELETE FROM ItemsXMov WHERE idMovimiento = " & adoTabla!id
                adoConnection.Execute sSql
                
                id_borrar = adoTabla!id
                adoTabla.Close
                sSql = "DELETE FROM Movimientos WHERE id = " & id_borrar
                adoConnection.Execute sSql
                
                ARMAR_GRILLA
            End If
        Case "AnularFactura"
            If MsgBox("¿Confirma que desea anular la FC Nº" & adoTabla!NumDoc & " del alumno " & DEVOLVER_CAMPO(adoTabla!idAlumno, adoAlumnos, "Alumnos", "Nombre") & "?", vbQuestion + vbYesNo, "ELIMINAR MODELO") = vbYes Then
                id_factura = adoTabla!id
                sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & id_factura
                With adoTemp
                    CERRAR_TABLA adoTemp
                    .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                    .MoveFirst
                    Do While Not .EOF
                        sSql = "SELECT * FROM Movimientos WHERE id = " & !idMod
                        CERRAR_TABLA adoTemp2
                        adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                        adoTemp2!Saldo = adoTemp2!Saldo + !Importe
                        adoTemp2.Update
                        adoTemp2.Close
                        
                        sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & !idMod
                        adoTemp3.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                        adoTemp3!Saldo = adoTemp3!Saldo + !Importe
                        adoTemp3.Update
                        adoTemp3.Close
                        
                        .MoveNext
                    Loop
                    .Close
                End With
            
                adoTabla!Anulado = True
                adoTabla.Update
                
                ARMAR_GRILLA
                
                'MOSTRAR_ACUMULADO_MENSUAL
                
                Unload Me
            End If
        Case "AnularRecibo"
            If MsgBox("¿Confirma que desea anular el REC Nº" & adoTabla!NumDoc & " del alumno " & DEVOLVER_CAMPO(adoTabla!idAlumno, adoAlumnos, "Alumnos", "Nombre") & "?", vbQuestion + vbYesNo, "ELIMINAR MODELO") = vbYes Then
                'INI - Guardo datos de la anulación.
                    With adoAnulados
                        .Open "Anulados", adoConnection, adOpenDynamic, adLockOptimistic
                        .AddNew
                        
                        !Sucursal = adoTabla!Sucursal
                        !TipoDoc = adoTabla!TipoDoc
                        !NumDoc = adoTabla!NumDoc
                        !FechaAnulacion = Date
                        !usuario = x_usuario
                        
                        .Update
                        
                        .Close
                    End With
                'FIN - Guardo datos de la anulación.
            
                GUARDAR_LOG x_usuario, Date, Time, "ANULA RECIBO " & adoTabla!Sucursal & "-" & adoTabla!NumDoc
            
                id_recibo = adoTabla!id
                sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & id_recibo
                With adoTemp
                    CERRAR_TABLA adoTemp
                    .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                    .MoveFirst
                    Do While Not .EOF
                        sSql = "SELECT * FROM Movimientos WHERE id = " & id_recibo
                        adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                        
                        es_curso_presencial = False
                        
                        If adoTemp2!UnidadNegocio = "CURSO PRESENCIAL" Then
                            es_curso_presencial = True
                        End If
                        
                        adoTemp2.Close
                    
                        If es_curso_presencial Then
                            sSql = "SELECT * FROM Movimientos WHERE id = " & !idMod
                            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                            adoTemp2!Saldo = adoTemp2!Saldo + !Importe
                            adoTemp2.Update
                            adoTemp2.Close
                            
                            sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & !idMod
                            adoTemp3.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                            adoTemp3!Saldo = adoTemp3!Saldo + !Importe
                            adoTemp3.Update
                            adoTemp3.Close
                        End If
                        
                        .MoveNext
                    Loop
                    .Close
                End With
            
                adoTabla!Anulado = True
                adoTabla.Update
                
                ARMAR_GRILLA
                
                Unload Me
            End If
        Case "Reimprimir"
            If MsgBox("¿Confirma que desea reimprimir el " & adoTabla!TipoDoc & " Nº " & adoTabla!NumDoc & "?", vbQuestion + vbYesNo, "REIMPRIMIR COMPROBANTE") = vbYes Then
                REIMPRIMIR
            End If
        Case "Consultar"
            id_Movimiento_Consultar = adoTabla!id
            Unload Me
            frmFacturaVer.Show
    End Select
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    CERRAR_TABLA adoTabla
    
    Unload Me
End Sub

Private Sub Form_Load()
    Select Case TipoBusDoc
        Case "BorrarModelo"
            Me.Caption = Me.Caption & " - Eliminar modelo"
            sSql = "SELECT Alumnos.Nombre, TiposCurso.Detalle AS TipoCurso, Movimientos.* FROM Movimientos, Alumnos, TiposCurso, Cursos WHERE TipoDoc = 'MOD' AND Saldo > 0 AND Alumnos.id = Movimientos.idAlumno AND Cursos.id = Movimientos.idCurso AND TiposCurso.id = Cursos.idTipoCurso"
        Case "AnularFactura"
            Me.Caption = Me.Caption & " - Anular factura"
            sSql = "SELECT * FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NOT Anulado"
        Case "AnularRecibo"
            Me.Caption = Me.Caption & " - Anular recibo"
            sSql = "SELECT * FROM Movimientos WHERE TipoDoc = 'REC' AND NOT Anulado"
        Case "Reimprimir"
            Me.Caption = Me.Caption & " - Reimprimir comprobante"
            sSql = "SELECT * FROM Movimientos WHERE TipoDoc <> 'MOD' AND TipoDoc <> 'ESP' AND NOT Anulado"
        Case "Consultar"
            Me.Caption = Me.Caption & " - Consultar comprobante"
            sSql = "SELECT * FROM Movimientos WHERE TipoDoc <> 'MOD' AND TipoDoc <> 'ESP' AND NOT Anulado"
    End Select

    ARMAR_GRILLA
End Sub

Private Sub ARMAR_GRILLA()
    CERRAR_TABLA adoTabla
    adoTabla.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
    
    Set dbgTempDocus.DataSource = adoTabla
    FORMATO_GRILLA
End Sub

Private Sub REIMPRIMIR()
    Dim Letras As New clsNumeros
    Dim cant_copias As Byte
    Dim Fila As Byte
    
    cant_copias = 1 'GetSetting("Gestion", "Documentos", "cantCopias")
        
    Printer.ScaleMode = 4
    Printer.FontSize = 10
    Printer.Font = "Courier"
    
    'Imprimo el encabezado
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos WHERE id = " & adoTabla!id
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    For k = 1 To cant_copias
        With adoMovimientos
            'Tacho factura y pongo nota de crédito
            If sMenu = "NotaCreditoPresenciales" Then
                IMPRIMIR 3, 55, "NOTA DE CRÉDITO"
                IMPRIMIR 4, 55, "XXXXXXXXXXXXXXXXXXX"
            End If
            
            IMPRIMIR 5, 60, !NumDoc
            IMPRIMIR 9, 80, !fecha
            IMPRIMIR 14, 10, !NombreCliente
            IMPRIMIR 15, 10, !Direccion
            IMPRIMIR 17, 10, !CondIva
            IMPRIMIR 17, 55, IIf(!Cuit = "", "----------", !Cuit)
            IMPRIMIR 19, 15, !FormaPago
        End With
        
        'Imprimo títulos
        IMPRIMIR 22, 67, "Unitario"
        IMPRIMIR 22, 77, "% Dto."
        
        'Imprimo el cuerpo
        CERRAR_TABLA adoItemsXMov
        sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & adoTabla!id
        adoItemsXMov.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        With adoItemsXMov
            .MoveFirst
            Fila = 24
            Do While Not .EOF
                IMPRIMIR Fila, 5, !Cantidad
                IMPRIMIR Fila, 9, !Detalle
                IMPRIMIR Fila, 67, Format(!Unitario, "Fixed")
                IMPRIMIR Fila, 77, Format(!Descuento, "Fixed")
                IMPRIMIR Fila, 86, Format(!Importe, "Fixed")
                
                .MoveNext
                Fila = Fila + 1
            Loop
        End With
        
        'Si pagó con tarjeta, imprimo el recargo.
        If adoMovimientos!Recargo > 0 Then
            IMPRIMIR Fila, 5, "1"
            IMPRIMIR Fila, 9, "RECARGO PAGO EN CUOTAS CON TARJETA DE CRÉDITO"
            IMPRIMIR Fila, 86, Format(adoMovimientos!Recargo, "Fixed")
        End If
        
        'Imprimo próximos vencimientos.
        CERRAR_TABLA adoTemp2
        sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & adoMovimientos!idAlumno & " AND TipoDoc = 'MOD' AND Saldo <> 0 ORDER BY fecha"
        adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If Not adoTemp2.EOF Then
            IMPRIMIR 29, 55, "Próximos vencimientos:"
            
            Fila = 30
            
            Printer.FontSize = 8
                
            
            Do While Not adoTemp2.EOF
                IMPRIMIR Fila, 60, Format(adoTemp2!fecha, "dd/mm/yyyy")
                adoTemp2.MoveNext
                
                Fila = Fila + 1
            Loop
                
            Printer.FontSize = 10
        End If
        
        adoTemp2.Close
        
        'Imprimo próximos vencimientos
        If sMenu = "FacturaPresenciales" Then
            IMPRIMIR 29, 55, "Próximos vencimientos:"
            
            Fila = 30
            
            Printer.FontSize = 8
                
            CERRAR_TABLA adoTemp2
            sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo <> 0 ORDER BY fecha"
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            Do While Not adoTemp2.EOF
                IMPRIMIR Fila, 60, adoTemp2!fecha
                    
                adoTemp2.MoveNext
                
                Fila = Fila + 1
            Loop
                
            Printer.FontSize = 10
        End If
        
        'If sMenu = "FacturaPresenciales" Then
        '    With adoTempFactura
        '        .MoveFirst
        '
        '        IMPRIMIR 29, 55, "Próximos vencimientos:"
        '        Fila = 30
        '        Printer.FontSize = 8
        '
        '        Do While Not .EOF
        '            If !Paga <> 0 Then
        '                sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo <> 0"
        '                CERRAR_TABLA adoTemp2
        '                adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        '                IMPRIMIR Fila, 60, adoTemp2!Fecha
        '                adoTemp2.Close
        '                Fila = Fila + 1
        '            End If
        '            .MoveNext
        '        Loop
        '
        '        Printer.FontSize = 10
        '    End With
        'End If
        
        'Imprimo próximos inicios
        With adoProximosInicios
            .Open "ProximosInicios", adoConnection, adOpenDynamic, adLockOptimistic
            
            IMPRIMIR 29, 5, "Próximos inicios:"
            Fila = 30
            Printer.FontSize = 8
            
            For j = 1 To 10
                campo = "proximo" & j
                IMPRIMIR Fila, 15, .Fields(campo) & ""
                Fila = Fila + 1
            Next
            .Close
            
            Printer.FontSize = 10
        End With
  
        'Imprimo otros datos
        IMPRIMIR 38, 5, adoMovimientos!Linea1
        IMPRIMIR 39, 5, adoMovimientos!Linea2
        IMPRIMIR 40, 5, adoMovimientos!Linea3
        
        'Discrimino el IVA
        If Right(adoMovimientos!TipoDoc, 1) = "A" Then
            IMPRIMIR 41, 88, Format(adoMovimientos!Subtotal, "Fixed")
            IMPRIMIR 43, 88, Format(adoMovimientos!Iva, "Fixed")
        End If
        
        IMPRIMIR 45, 88, Format(adoMovimientos!Total, "Fixed")
        
        IMPRIMIR 41, 5, "SON PESOS " & Letras.NroEnLetras(adoMovimientos!Total)

        'Imprimo
        Printer.EndDoc
    
    Next
    
    'Cierro las tablas
    adoMovimientos.Close
    'adoItemsXMov.Close
    CERRAR_TABLA adoItemsXMov

End Sub

Private Sub REIMPRIMIR_()
    Dim Fila As Byte

    Printer.ScaleMode = 4
    Printer.FontSize = 10
    Printer.Font = "Courier"
    
    'Imprimo el encabezado
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos WHERE id = " & adoTabla!id
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    With adoMovimientos
        'Tacho factura y pongo nota de crédito
        If sMenu = "NotaCreditoPresenciales" Then
            IMPRIMIR 3, 55, "NOTA DE CRÉDITO"
            IMPRIMIR 4, 55, "XXXXXXXXXXXXXXXXXXX"
        End If
        
        IMPRIMIR 5, 60, !NumDoc
        IMPRIMIR 9, 80, !fecha
        IMPRIMIR 14, 10, !NombreCliente
        IMPRIMIR 15, 10, !Direccion
        IMPRIMIR 17, 10, !CondIva
        IMPRIMIR 17, 55, IIf(!Cuit = "", "----------", !Cuit)
        IMPRIMIR 19, 15, !FormaPago
    End With
        
    'Imprimo títulos
    IMPRIMIR 22, 67, "Unitario"
    IMPRIMIR 22, 77, "% Dto."
    
    'Imprimo el cuerpo
    CERRAR_TABLA adoItemsXMov
    sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & adoTabla!id
    adoItemsXMov.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    With adoItemsXMov
        .MoveFirst
        Fila = 24
        Do While Not .EOF
            IMPRIMIR Fila, 5, !Cantidad
            IMPRIMIR Fila, 9, !Detalle
            IMPRIMIR Fila, 67, Format(!Unitario, "Fixed")
            IMPRIMIR Fila, 77, Format(!Descuento, "Fixed")
            IMPRIMIR Fila, 86, Format(!Importe, "Fixed")
            
            .MoveNext
            Fila = Fila + 1
        Loop
    End With
        
    'Imprimo otros datos
    IMPRIMIR 38, 5, adoMovimientos!Linea1
    IMPRIMIR 39, 5, adoMovimientos!Linea2
    IMPRIMIR 40, 5, adoMovimientos!Linea3
    
    'Discrimino el IVA
    If Right(adoMovimientos!TipoDoc, 1) = "A" Then
        IMPRIMIR 41, 88, Format(adoMovimientos!Subtotal, "Fixed")
        IMPRIMIR 43, 88, Format(adoMovimientos!Iva, "Fixed")
    End If
    
    IMPRIMIR 45, 88, Format(adoMovimientos!Total, "Fixed")
    
    'Imprimo
    Printer.EndDoc
    
    'Cierro las tablas
    adoMovimientos.Close
    adoItemsXMov.Close
End Sub

Private Sub REIMPRIMIR2()
    Dim Fila As Byte
    
    Printer.ScaleMode = 4
    Printer.FontSize = 10
    Printer.Font = "Courier"
    
    'Imprimo el encabezado
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos WHERE id = " & adoTabla!id
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    With adoMovimientos
        IMPRIMIR 5, 60, !NumDoc
        IMPRIMIR 9, 80, !fecha
        IMPRIMIR 14, 10, !NombreCliente
        IMPRIMIR 15, 10, !Direccion
        IMPRIMIR 17, 10, !CondIva
        IMPRIMIR 17, 55, IIf(!Cuit = "", "----------", !Cuit)
        IMPRIMIR 19, 15, !FormaPago
        
        IMPRIMIR 47, 86, Format(!Total, "Fixed")
    End With
    
    'Imprimo títulos
    IMPRIMIR 22, 67, "Unitario"
    IMPRIMIR 22, 77, "% Dto."
    
    'Imprimo el cuerpo
    CERRAR_TABLA adoItemsXMov
    sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & adoTabla!id
    adoItemsXMov.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    With adoItemsXMov
        .MoveFirst
        Fila = 24
        Do While Not .EOF
            IMPRIMIR Fila, 5, !Cantidad
            IMPRIMIR Fila, 9, !Detalle
            IMPRIMIR Fila, 67, Format(!Unitario, "Fixed")
            IMPRIMIR Fila, 77, Format(!Descuento, "Fixed")
            IMPRIMIR Fila, 86, Format(!Importe, "Fixed")
            
            .MoveNext
            Fila = Fila + 1
        Loop
    End With
    
    'Imprimo otros datos
    Fila = Fila + 3
    IMPRIMIR Fila, 5, adoMovimientos!Linea1
    Fila = Fila + 1
    IMPRIMIR Fila, 5, adoMovimientos!Linea2
    Fila = Fila + 1
    IMPRIMIR Fila, 5, adoMovimientos!Linea3
    
    'Discrimino el IVA
    If Right(adoMovimientos!TipoDoc, 1) = "A" Then
        IMPRIMIR Fila, 43, "SUBTOTAL " & Format(adoMovimientos!Subtotal, "Fixed")
        IMPRIMIR Fila, 45, "IVA      " & Format(adoMovimientos!Iva, "Fixed")
    End If
    
    'Imprimo
    Printer.EndDoc
    
    'Cierro las tablas
    adoMovimientos.Close
    adoItemsXMov.Close
End Sub

Private Sub FORMATO_GRILLA()
    Select Case TipoBusDoc
        Case "BorrarModelo"
        Case "AnularFactura"
        Case "Reimprimir"
            dbgTempDocus.Columns(0).Visible = False
            dbgTempDocus.Columns(1).Visible = False
            
            For k = 5 To 10
                dbgTempDocus.Columns(k).Visible = False
            Next
            For k = 13 To 23
                dbgTempDocus.Columns(k).Visible = False
            Next
            
            
    End Select
End Sub

Private Sub txtDato_Change()
    'filtro = "((NumDoc LIKE ('%" & txtDato.Text & "%')) OR (NombreCliente LIKE ('" & txtDato.Text & "%')))"
    filtro = "(NumDoc = '" & txtDato.Text & "')"
    
    Select Case TipoBusDoc
        Case "BorrarModelo"
            sSql = "SELECT * FROM Movimientos WHERE TipoDoc = 'MOD' AND Saldo > 0 AND " & filtro
        Case "AnularFactura"
            sSql = "SELECT * FROM Movimientos WHERE Left(TipoDoc, 2) = 'FC' AND NOT Anulado AND " & filtro
        Case "AnularRecibo"
            sSql = "SELECT * FROM Movimientos WHERE TipoDoc = 'REC' AND NOT Anulado AND " & filtro
        Case "Reimprimir"
            sSql = "SELECT * FROM Movimientos WHERE TipoDoc <> 'MOD' AND NOT Anulado AND " & filtro
        Case "Consultar"
            sSql = "SELECT * FROM Movimientos WHERE TipoDoc <> 'MOD' AND NOT Anulado AND " & filtro
    End Select

    ARMAR_GRILLA
End Sub
