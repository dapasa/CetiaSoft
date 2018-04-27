VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaAlumnosCursoNumero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTADO - Alumnos por curso (número)"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   Icon            =   "frmListaAlumnosCursoNumero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Actualizar"
      Height          =   855
      Left            =   7560
      Picture         =   "frmListaAlumnosCursoNumero.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Acceder"
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
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
      TabIndex        =   13
      Top             =   840
      Width           =   4095
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCerrado 
         Caption         =   "Cerrado"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAbierto 
         Caption         =   "Abierto"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fecha de inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Width           =   3015
      Begin MSComCtl2.DTPicker dtpFechaIniHasta 
         Height          =   330
         Left            =   960
         TabIndex        =   9
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   17629185
         CurrentDate     =   40147
      End
      Begin MSComCtl2.DTPicker dtpFechaIniDesde 
         Height          =   330
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   17629185
         CurrentDate     =   40147
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   6960
      Picture         =   "frmListaAlumnosCursoNumero.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Acceder"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7920
      Picture         =   "frmListaAlumnosCursoNumero.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancelar"
      Top             =   6360
      Width           =   855
   End
   Begin VB.Frame fraCursos 
      Caption         =   "Cursos disponibles"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   8655
      Begin MSDataGridLib.DataGrid dbgCursos 
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
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
   End
   Begin VB.Frame fraInscriptos 
      Caption         =   "Inscriptos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   8655
      Begin MSDataGridLib.DataGrid dbgInscriptos 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
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
   End
   Begin VB.Frame Frame7 
      Caption         =   "Tipo de curso"
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
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cboTipoCurso 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   6360
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "frmListaAlumnosCursoNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EventoLoad As Boolean
Dim id_ As String, FechaIni_ As String, FechaFin_ As String
Dim Detalle_ As String, Vacantes_ As String, Inscriptos_ As String, Espera_ As String
Dim Cuotas As Byte, ValorCuota As Single
Dim curso_actual As Byte
Dim mensaje_esta_seguro As String
Dim paso_a_espera As Boolean
Dim paso_a_inscriptos As Boolean
Dim se_va_a_espera As Boolean
Dim se_va_a_inscriptos As Boolean

Private Sub cboTipoCurso_Click()
    If Not EventoLoad Then
        fraCursos.Enabled = True
        fraInscriptos.Enabled = True
        
        CARGAR_CURSOS "Todas"
        
        VACIAR_GRILLA_INSCRIPTOS
    End If
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    sSql = "DELETE FROM zzAlumnosXCursoNumero"
    adoConnection.Execute sSql
    
    'Inscriptos
    sSql = "INSERT INTO zzAlumnosXCursoNumero (Nombre, Teléfono, Celular, [e-Mail], Matrícula, Cuotas, FechaIni, FechaFin, Curso, Horario) " & _
           "SELECT Nombre, Teléfono, Celular, [e-Mail], Matrícula, Cuotas, '" & adoTempCursos!Inicio & "', '" & adoTempCursos!Fin & "', '" & cboTipoCurso.Text & "', '" & adoTempCursos!Horario & "' " & _
           "FROM TempInscriptos"

    adoConnection.Execute sSql
   
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    
    rptListado.ReportFileName = App.Path & "\reportes\rptAlumnosXCursoNumero.rpt"
    
    rptListado.ReportTitle = "LISTADO DE ALUMNOS POR CURSO"
    rptListado.Action = 1
End Sub

Private Sub Command1_Click()
    CARGAR_CURSOS_FILTRO "Todas"
End Sub

Private Sub dbgCursos_Click()
    If Not adoTempCursos.EOF Then
        VER_INSCRIPTOS
    End If
End Sub

Private Sub Form_Load()
    EventoLoad = True
    CARGAR_COMBO "cboTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    EventoLoad = False

    dtpFechaIniDesde.Value = Date
    dtpFechaIniHasta.Value = Date
    
    VACIAR_GRILLA_CURSOS
    VACIAR_GRILLA_INSCRIPTOS
End Sub

Private Sub VACIAR_GRILLA_CURSOS()
    CERRAR_TABLA adoTempCursos
    sSql = "DELETE FROM TempCursos"
    adoConnection.Execute sSql
    adoTempCursos.Open "TempCursos", adoConnection, adOpenKeyset, adLockOptimistic
    Set dbgCursos.DataSource = adoTempCursos
    
    FORMATO_GRILLA_CURSOS
End Sub

Private Sub VACIAR_GRILLA_INSCRIPTOS()
    CERRAR_TABLA adoTempInscriptos
    sSql = "DELETE FROM TempInscriptos"
    adoConnection.Execute sSql
    adoTempInscriptos.Open "TempInscriptos", adoConnection, adOpenKeyset, adLockOptimistic
    Set dbgInscriptos.DataSource = adoTempInscriptos
    
    FORMATO_GRILLA_INSCRIPTOS
End Sub

Private Sub FORMATO_GRILLA_CURSOS()
    With dbgCursos
        .Columns(0).Visible = False
        .Columns(1).Width = 500
        .Columns(1).Caption = "Nº"
        .Columns(2).Width = 1100
        .Columns(3).Width = 1100
        .Columns(4).Width = 2400
        .Columns(5).Width = 900
        .Columns(6).Width = 900
        .Columns(7).Width = 900
    End With
End Sub

Private Sub FORMATO_GRILLA_INSCRIPTOS()
    With dbgInscriptos
        .Columns(0).Visible = False
        .Columns(1).Width = 2000
        .Columns(2).Width = 1000
        .Columns(3).Width = 1000
        .Columns(4).Width = 2000
        .Columns(5).Width = 800
        .Columns(6).Width = 800
    End With
End Sub

Private Sub CARGAR_CURSOS(Estado As String)
    Dim filtroEstado As String
    
    Select Case Estado
        Case "Disponibles"
            filtroEstado = "Cursos.Vacantes > 0"
        Case "EnEspera"
            filtroEstado = "Cursos.Vacantes = 0"
        Case "Todas"
            filtroEstado = "1 = 1"
    End Select
    
    VACIAR_GRILLA_CURSOS
    
    If optAbierto.Value = True Then
        filtroCurso = "(Abierto = True)"
    ElseIf optCerrado.Value = True Then
        filtroCurso = "(Abierto = False)"
    Else
        filtroCurso = "(1 = 1)"
    End If

    With adoCursos
        CERRAR_TABLA adoCursos
        sSql = "SELECT Cursos.id, Cursos.Numero, Cursos.FechaIni, Cursos.FechaFin, Horarios.Detalle, Cursos.Vacantes, Cursos.Inscriptos " & _
               "FROM Cursos, Horarios " & _
               "WHERE Cursos.idTipoCurso = " & DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "Detalle") & " " & _
               "AND Cursos.idHorario = Horarios.id " & _
               "AND " & filtroEstado & " " & _
               "AND Cursos.Abierto " & _
               "ORDER BY Cursos.Numero"
        .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
        Do While Not .EOF
            adoTempCursos.AddNew
            adoTempCursos!id = !id
            adoTempCursos!Numero = !Numero
            adoTempCursos!Inicio = !FechaIni
            adoTempCursos!Fin = !FechaFin
            adoTempCursos!Horario = !Detalle
            adoTempCursos!Vacantes = !Vacantes
            adoTempCursos!inscriptos = !inscriptos
            
            CERRAR_TABLA adoListaEspera
            sSql = "SELECT COUNT(id) AS CantEspera FROM ListaEspera WHERE idCurso = " & !id
            adoListaEspera.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            adoTempCursos!Espera = adoListaEspera!CantEspera
            adoListaEspera.Close
            
            adoTempCursos.Update
            
            .MoveNext
        Loop
        .Close
        
        If Not adoTempCursos.EOF Then
            adoTempCursos.MoveLast
        End If
    End With
End Sub

Private Sub CARGAR_CURSOS_FILTRO(Estado As String)
    Dim filtroEstado As String
    
    Select Case Estado
        Case "Disponibles"
            filtroEstado = "Cursos.Vacantes > 0"
        Case "EnEspera"
            filtroEstado = "Cursos.Vacantes = 0"
        Case "Todas"
            filtroEstado = "1 = 1"
    End Select
    
    VACIAR_GRILLA_CURSOS
    
    If optAbierto.Value = True Then
        filtroCurso = "(Cursos.Abierto = True)"
    ElseIf optCerrado.Value = True Then
        filtroCurso = "(Cursos.Abierto = False)"
    Else
        filtroCurso = "(1 = 1)"
    End If

    With adoCursos
        CERRAR_TABLA adoCursos
        sSql = "SELECT Cursos.id, Cursos.Numero, Cursos.FechaIni, Cursos.FechaFin, Horarios.Detalle, Cursos.Vacantes, Cursos.Inscriptos " & _
               "FROM Cursos, Horarios " & _
               "WHERE Cursos.idTipoCurso = " & DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "Detalle") & " " & _
               "AND Cursos.idHorario = Horarios.id " & _
               "AND (Cursos.FechaIni >= DateValue('" & dtpFechaIniDesde.Value & "') AND Cursos.FechaIni <= DateValue('" & dtpFechaIniHasta.Value & "')) " & _
               "AND " & filtroEstado & " " & _
               "AND " & filtroCurso & " " & _
               "ORDER BY Cursos.Numero"
        .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        Do While Not .EOF
            adoTempCursos.AddNew
            adoTempCursos!id = !id
            adoTempCursos!Numero = !Numero
            adoTempCursos!Inicio = !FechaIni
            adoTempCursos!Fin = !FechaFin
            adoTempCursos!Horario = !Detalle
            adoTempCursos!Vacantes = !Vacantes
            adoTempCursos!inscriptos = !inscriptos
            
            CERRAR_TABLA adoListaEspera
            sSql = "SELECT COUNT(id) AS CantEspera FROM ListaEspera WHERE idCurso = " & !id
            adoListaEspera.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            adoTempCursos!Espera = adoListaEspera!CantEspera
            adoListaEspera.Close
            
            adoTempCursos.Update
            
            .MoveNext
        Loop
        .Close
        
        If Not adoTempCursos.EOF Then
            adoTempCursos.MoveLast
        End If
    End With
End Sub

Private Sub VER_INSCRIPTOS()
        If adoTempCursos.EOF Then
            Exit Sub
        End If
        
        VACIAR_GRILLA_INSCRIPTOS
        
        TipoLista = "Inscriptos"
        
        id_ = adoTempCursos!id
        
        CERRAR_TABLA adoTabla
        sSql = "SELECT Alumnos.id, Alumnos.Nombre, Alumnos.Telefono, Alumnos.Celular, Alumnos.Mail, Alumnos.Celular FROM Alumnos, AlumnosXCurso WHERE Alumnos.id = AlumnosXCurso.idAlumno AND AlumnosXCurso.idCurso = " & adoTempCursos!id & " ORDER BY Nombre"
        adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
            
        Do While Not adoTabla.EOF
            adoTempInscriptos.AddNew
            adoTempInscriptos!id = adoTabla!id
            adoTempInscriptos!Nombre = adoTabla!Nombre
            adoTempInscriptos.Fields("Teléfono").Value = adoTabla!Telefono
            adoTempInscriptos!Celular = adoTabla!Celular
            adoTempInscriptos.Fields("e-Mail").Value = adoTabla!Mail
            
            a_curso = adoTempCursos!id
            a_alumno = adoTabla!id
            
            CERRAR_TABLA adoTemp2
            'sSql = "SELECT SUM(Total) AS Importe FROM Movimientos WHERE idAlumno = " & a_alumno & " AND idCurso = " & a_curso & " AND Cuota = 0 AND Left(TipoDoc, 2) = 'FC'"
            sSql = "SELECT SUM(Total) AS Importe FROM Movimientos WHERE idAlumno = " & a_alumno & " AND idCurso = " & a_curso & " AND Saldo = 0 AND Cuota = 0 AND TipoDoc = 'MOD'"
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            adoTempInscriptos.Fields("Matrícula").Value = adoTemp2!Importe
            
            CERRAR_TABLA adoTemp2
            sSql = "SELECT SUM(Total) AS Importe FROM Movimientos WHERE idAlumno = " & a_alumno & " AND idCurso = " & a_curso & " AND Saldo = 0 AND Cuota <> 0 AND TipoDoc = 'MOD'"
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            adoTempInscriptos!Cuotas = adoTemp2!Importe
            
            adoTempInscriptos.Update
            
            adoTabla.MoveNext
        Loop
        adoTabla.Close
        
        If Not adoTempInscriptos.EOF Then
            adoTempInscriptos.MoveFirst
        End If
End Sub

Private Sub VOLVER_CURSO_ACTUAL()
    With adoTempCursos
        .MoveFirst
        Do While !id <> id_Curso
            .MoveNext
        Loop
    End With
End Sub
