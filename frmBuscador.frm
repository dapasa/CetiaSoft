VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUSCAR - Profesores"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "frmBuscador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   9120
      Picture         =   "frmBuscador.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   8400
      Picture         =   "frmBuscador.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   6600
      Width           =   615
   End
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
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtDato 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9375
      End
   End
   Begin MSDataGridLib.DataGrid dbgTabla 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Seleccione un profesor y haga click en Aceptar."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3405
   End
End
Attribute VB_Name = "frmBuscador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' MÓDULO: Buscador genérico           FECHA: Ago / 2007
'******************************************************
' RESUMEN: único buscador para todas las opciones del
'          sistema.
'******************************************************
' ÚLTIMA MODIFICACIÓN IMPORTANTE: 16/08/2007
'******************************************************
' ETAPA: release candidate.
'******************************************************
' AUTOR: Pablo Adrián Langholz
' CONTACTO: elmaildepablo@gmail.com
'******************************************************

Dim itemSeleccionado As Boolean
Dim busTabla As String
Dim busCampo As String
Dim busFiltro As String
Dim busFrom As String

Private Sub cmdAceptar_Click()
'    On Error GoTo ErrorHandle
    
    If itemSeleccionado Then
        Select Case sMenu
            Case "Alumnos"
                ActualizarPantalla "Alumnos", frmAlumnos, adoTablaBus
            Case "Empresas"
                ActualizarPantalla "Empresas", frmEmpresas, adoTablaBus
            Case "Profesores"
                ActualizarPantalla "Profesores", frmProfesores, adoTablaBus
            Case "Cursos"
                ActualizarPantalla "Cursos", frmCursos, adoTablaBus
            Case "Inscripcion"
                ActualizarPantalla "Inscripcion", frmInscripcionBis, adoTablaBus, "Alumno"
            Case "FacturaPresenciales", "FacturaDistancia", "FacturaServicioTecnico", "FacturaVentaHardware", "FacturaAnticipo", "NotaCreditoPresenciales", "NotaCreditoOtros"
                ActualizarPantalla "Factura", frmFactura, adoTablaBus
            Case "CobranzaCuentaCorriente"
                ActualizarPantalla "Cobranza", frmRecibo, adoTablaBus
        End Select
        Unload Me
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - cmdAceptar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
        
    CanceloBuscador = True
    
    If EstiloBuscador = "FacturaEmpresas" Then
        CancelaBusEmpresa = True
    End If
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - Form.Unload", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub dbgTabla_Click()
    On Error GoTo ErrorHandle
    
    itemSeleccionado = True

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - dbgTabla.Click", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub dbgTabla_DblClick()
    On Error GoTo ErrorHandle
    
    itemSeleccionado = True
    cmdAceptar_Click

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - dbgTabla.DblClick", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
        
    CanceloBuscador = False
    CancelaBusEmpresa = False
    
    CONFIGURAR_BUSCADOR
    
    ACTUALIZAR_GRILLA
    
    If adoTablaBus.RecordCount > 0 Then
        txtDato.Text = itemBuscado
        txtDato.SelStart = Len(txtDato.Text)
    End If

    If x_ficha_alumno_desde_factura <> "" Then
        txtDato.Text = x_ficha_alumno_desde_factura
        opcion_menu = sMenu
        sMenu = "Alumnos"
        cmdAceptar_Click
        sMenu = opcion_menu
        x_ficha_alumno_desde_factura = ""
    End If
        
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - Form.Load", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    CERRAR_TABLA adoTablaBus

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - Form.Unload", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub txtDato_Change()
    On Error GoTo ErrorHandle
    
    If EstiloBuscador = "Cursos" Then
        sSql = "SELECT Cursos.id, Cursos.Numero, TiposCurso.Detalle AS TipoCurso, Horarios.Detalle AS Horario, Aulas.Detalle AS Aula, Cursos.FechaIni, Cursos.FechaFin, Profesores.Nombre AS Profesor, Cursos.Vacantes, Cursos.Inscriptos, Duraciones.Detalle AS Duracion, Cursos.Abierto, Cursos.CantCuotas, Cursos.Cuota1, Cursos.Cuota2, Cursos.Cuota3, Cursos.Cuota4, Cursos.Cuota5, Cursos.Observaciones " & _
               "FROM Cursos, Duraciones, Profesores, TiposCurso, Aulas, Horarios " & _
               "WHERE Cursos.idTipoCurso = TiposCurso.id AND Cursos.idDuracion = Duraciones.id AND Cursos.idProfesor = Profesores.id AND Cursos.idAula = Aulas.id AND Cursos.idHorario = Horarios.id " & _
               "      AND ((TiposCurso.Detalle LIKE '%" & txtDato.Text & "%') OR (Cursos.Numero LIKE '%" & txtDato.Text & "%')) " & _
               "ORDER BY Cursos.Numero DESC"
    ElseIf EstiloBuscador = "Alumnos" Then
        sSql = "SELECT * FROM alumnos " & _
               "WHERE nombre LIKE '%" & txtDato.Text & "%' " & _
               "      OR numDoc = '" & txtDato.Text & "' " & _
               "ORDER BY nombre"
    Else
        sSql = "SELECT * FROM " & busTabla & " WHERE " & busCampo & " LIKE '%" & txtDato.Text & "%' ORDER BY " & busCampo
    End If
    
    ACTUALIZAR_GRILLA
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - txtDato.Change", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If KeyAscii = 13 Then
        itemSeleccionado = True
        cmdAceptar_Click
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmBuscador - txtDato.KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub ACTUALIZAR_GRILLA()
    On Error GoTo ErrorHandle
    
    CERRAR_TABLA adoTablaBus
    adoTablaBus.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
    Set dbgTabla.DataSource = adoTablaBus
    
    With dbgTabla
        Select Case EstiloBuscador
            Case "Alumnos", "FacturaAlumnos"
                'Oculto columnas
                .Columns(0).Visible = False
                .Columns(3).Visible = False
                .Columns(4).Visible = False
                .Columns(8).Visible = False
                .Columns(10).Visible = False
                .Columns(11).Visible = False
                .Columns(12).Visible = False
                .Columns(13).Visible = False
                .Columns(14).Visible = False
                .Columns(15).Visible = False
                .Columns(16).Visible = False
                
                'Modifico formatos
                .Columns(1).Width = 2000
                .Columns(2).Width = 1000
                .Columns(5).Width = 1500
                .Columns(6).Width = 1500
                .Columns(7).Width = 1500
                .Columns(9).Width = 1500
            Case "Empresas", "FacturaEmpresas"
                'Oculto columnas
                .Columns(0).Visible = False
                .Columns(3).Visible = False
                .Columns(4).Visible = False
                .Columns(5).Visible = False
                .Columns(6).Visible = False
                .Columns(9).Visible = False
                .Columns(11).Visible = False
                
                'Modifico formatos
                .Columns(1).Width = 2000
                .Columns(2).Width = 1000
                .Columns(7).Width = 1500
                .Columns(8).Width = 1500
                .Columns(10).Width = 2000
            Case "Profesores"
                'Oculto columnas
                .Columns(0).Visible = False
                .Columns(2).Visible = False
                .Columns(5).Visible = False
                .Columns(7).Visible = False
                
                'Modifico formatos
                .Columns(1).Width = 2000
                .Columns(3).Width = 1500
                .Columns(4).Width = 1500
                .Columns(6).Width = 1500
            Case "Cursos"
                .Columns(18).Visible = False
        End Select
    End With

    Exit Sub
ErrorHandle:
    'MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "ACTUALIZAR_GRILLA", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub CONFIGURAR_BUSCADOR()
    On Error GoTo ErrorHandle
    
    Select Case EstiloBuscador
        Case "Alumnos", "FacturaAlumnos"
            Me.Caption = "BUSCAR - Alumnos"
            lblSeleccione.Caption = "Seleccione un alumno y haga click en Aceptar."
            
            If itemBuscado = "" Then
                sSql = "SELECT * FROM Alumnos ORDER BY Nombre"
            Else
                sSql = "SELECT * FROM Alumnos WHERE Nombre LIKE '%" & itemBuscado & "%' ORDER BY Nombre"
            End If
            
            busTabla = "Alumnos"
            busCampo = "Nombre"
            
        Case "Empresas", "FacturaEmpresas"
            Me.Caption = "BUSCAR - Empresas"
            lblSeleccione.Caption = "Seleccione una empresa y haga click en Aceptar."
            If itemBuscado = "" Then
                sSql = "SELECT * FROM Empresas ORDER BY Nombre"
            Else
                sSql = "SELECT * FROM Empresas WHERE Nombre LIKE '%" & itemBuscado & "%' ORDER BY Nombre"
            End If
            
            busTabla = "Empresas"
            busCampo = "Nombre"
            
        Case "Profesores"
            Me.Caption = "BUSCAR - Profesores"
            lblSeleccione.Caption = "Seleccione un profesor y haga click en Aceptar."
            
            If itemBuscado = "" Then
                sSql = "SELECT * FROM Profesores ORDER BY Nombre"
            Else
                sSql = "SELECT * FROM Profesores WHERE Nombre LIKE '%" & itemBuscado & "%' ORDER BY Nombre"
            End If

            busTabla = "Profesores"
            busCampo = "Nombre"
            
        Case "Cursos"
            Me.Caption = "BUSCAR - Cursos"
            lblSeleccione.Caption = "Seleccione un curso y haga click en Aceptar."
            
            Me.Width = 12000
            dbgTabla.Width = 11670
            fraDato.Width = 11655
            'txtDato.Text = 11415
            cmdAceptar.Left = 10440
            cmdCancelar.Left = 11160
            
            If itemBuscado = "" Then
                'sSql = "SELECT Cursos.id, Cursos.Numero, Cursos.idDuracion, Cursos.FechaIni, Cursos.FechaFin, Cursos.idHorario, Cursos.idProfesor, Cursos.idTipoCurso, Cursos.idModalidad, Cursos.idAula, Cursos.CantCuotas, Cursos.ValorCuota, Cursos.Vacantes, Cursos.Inscriptos, Cursos.Abierto, Duraciones.Detalle, Horarios.Detalle, TiposCurso.Detalle FROM Cursos, Duraciones, Horarios WHERE Cursos.idDuracion = Duraciones.id AND Cursos.idHorario = Horarios.id AND Cursos.idTipoCurso = TiposCurso.id ORDER BY Cursos.Numero"
                sSql = "SELECT Cursos.id, Cursos.Numero, TiposCurso.Detalle AS TipoCurso, Horarios.Detalle AS Horario, Aulas.Detalle AS Aula, Cursos.FechaIni, Cursos.FechaFin, Profesores.Nombre AS Profesor, Cursos.Vacantes, Cursos.Inscriptos, Duraciones.Detalle AS Duracion, Cursos.Abierto, Cursos.CantCuotas, Cursos.Cuota1, Cursos.Cuota2, Cursos.Cuota3, Cursos.Cuota4, Cursos.Cuota5, Cursos.Observaciones FROM (((((Cursos INNER JOIN Duraciones ON Cursos.idDuracion = Duraciones.id) INNER JOIN Profesores ON Cursos.idProfesor = Profesores.id) INNER JOIN TiposCurso ON Cursos.idTipoCurso = TiposCurso.id) INNER JOIN Aulas ON Cursos.idAula = Aulas.id) INNER JOIN Horarios ON Cursos.idHorario = Horarios.id) ORDER BY Cursos.Numero"
            Else
                'sSql = "SELECT * FROM Cursos WHERE Numero LIKE '%" & itemBuscado & "%' ORDER BY Numero"
                sSql = "SELECT Cursos.id, Cursos.Numero, TiposCurso.Detalle AS TipoCurso, Horarios.Detalle AS Horario, Aulas.Detalle AS Aula, Cursos.FechaIni, Cursos.FechaFin, Profesores.Nombre AS Profesor, Cursos.Vacantes, Cursos.Inscriptos, Duraciones.Detalle AS Duracion, Cursos.Abierto, Cursos.CantCuotas, Cursos.Cuota1, Cursos.Cuota2, Cursos.Cuota3, Cursos.Cuota4, Cursos.Cuota5, Cursos.Observaciones  FROM (((((Cursos INNER JOIN Duraciones ON Cursos.idDuracion = Duraciones.id) INNER JOIN Profesores ON Cursos.idProfesor = Profesores.id) INNER JOIN TiposCurso ON Cursos.idTipoCurso = TiposCurso.id) INNER JOIN Aulas ON Cursos.idAula = Aulas.id) INNER JOIN Horarios ON Cursos.idHorario = Horarios.id) WHERE Numero LIKE '%" & itemBuscado & "%'  ORDER BY Cursos.Numero"
            End If
            
            busTabla = "Cursos"
            busCampo = "Numero"
    End Select

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "CONFIGURAR_BUSCADOR", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub
