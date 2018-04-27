VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInscripcionBis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INSCRIPCIÓN"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   Icon            =   "frmInscripcionBis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbandono 
      Caption         =   "Abandono"
      Height          =   855
      Left            =   10800
      Picture         =   "frmInscripcionBis.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Abandonar el curso"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdInscribirRecursante 
      Caption         =   "Recursar"
      Height          =   855
      Left            =   10800
      Picture         =   "frmInscripcionBis.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Inscribir recursante"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "C&ambiar"
      Height          =   855
      Left            =   9840
      Picture         =   "frmInscripcionBis.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cambiar de curso"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdPasarAInscriptos 
      Caption         =   "Inscriptos"
      Height          =   855
      Left            =   8880
      Picture         =   "frmInscripcionBis.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Pasar a lista de inscriptos"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdPasarAEspera 
      Caption         =   "Espera"
      Height          =   855
      Left            =   8880
      Picture         =   "frmInscripcionBis.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Pasar a lista de espera"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminarEspera 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8880
      Picture         =   "frmInscripcionBis.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Eliminar"
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminarInscriptos 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8880
      Picture         =   "frmInscripcionBis.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Eliminar"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   9840
      Picture         =   "frmInscripcionBis.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancelar"
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton cmdSeleccionarAlumno 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   8880
      Picture         =   "frmInscripcionBis.frx":2652
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Seleccionar alumno"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdInscribir 
      Caption         =   "&Inscribir"
      Height          =   855
      Left            =   8880
      Picture         =   "frmInscripcionBis.frx":2A94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Inscribir al curso"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame fraAlumno 
      Caption         =   "Alumno"
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
      Left            =   8880
      TabIndex        =   24
      Top             =   1080
      Width           =   1815
      Begin VB.TextBox txtAlumno 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdNuevoAlumno 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   9840
      Picture         =   "frmInscripcionBis.frx":2ED6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo alumno"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdVerListaEspera 
      Caption         =   "&Espera"
      Height          =   855
      Left            =   9840
      Picture         =   "frmInscripcionBis.frx":3318
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ver lista de espera"
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdVerInscriptos 
      Caption         =   "&Inscriptos"
      Height          =   855
      Left            =   9840
      Picture         =   "frmInscripcionBis.frx":375A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ver inscriptos"
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdFactura 
      Caption         =   "&Factura"
      Height          =   855
      Left            =   10800
      Picture         =   "frmInscripcionBis.frx":3B9C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Generar factura"
      Top             =   1200
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
      TabIndex        =   22
      Top             =   840
      Width           =   8655
      Begin MSDataGridLib.DataGrid dbgCursos 
         Height          =   1935
         Left            =   120
         TabIndex        =   23
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
   Begin VB.Frame fraEspera 
      Caption         =   "En espera"
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
      TabIndex        =   20
      Top             =   5760
      Width           =   8655
      Begin MSDataGridLib.DataGrid dbgEspera 
         Height          =   1935
         Left            =   120
         TabIndex        =   21
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
      TabIndex        =   18
      Top             =   3360
      Width           =   8655
      Begin MSDataGridLib.DataGrid dbgInscriptos 
         Height          =   1935
         Left            =   120
         TabIndex        =   19
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
      TabIndex        =   17
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cboTipoCurso 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vacantes"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton optVacantesDisponibles 
         Caption         =   "Disponibles"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1120
      End
      Begin VB.OptionButton optVacantesEnEspera 
         Caption         =   "En espera"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optVacantesTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   8760
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmInscripcionBis"
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
        fraEspera.Enabled = True
        
        CARGAR_CURSOS "Todas"
        
        VACIAR_GRILLA_INSCRIPTOS
        VACIAR_GRILLA_ESPERA
    End If
End Sub

Private Sub cmdAbandono_Click()
    Dim tipo_curso As Long
    
    'Tengo que escribir código para ver si hay un ítem seleccionado
        'If mensaje_esta_seguro = "" Then
            mensaje_esta_seguro = "¿Confirma que el alumno abandona el curso?"
        'End If
        
        If MsgBox(mensaje_esta_seguro, vbQuestion + vbYesNo, "ABANDONAR CURSO") = vbYes Then
            'paso_a_espera = True
            
            id_Alumno = DEVOLVER_ID(adoTempInscriptos!Nombre, adoAlumnos, "Alumnos", "Nombre")
            id_Curso = adoTempCursos!id
            
            num_curso = DEVOLVER_CAMPO(id_Curso, adoCursos, "cursos", "numero")
            
            tipo_curso = DEVOLVER_CAMPO(id_Curso, adoCursos, "cursos", "idTipoCurso")
            nombre_curso = DEVOLVER_CAMPO(tipo_curso, adoTiposCurso, "TiposCurso", "detalle")
            
            
            'INI - Verifico cuantas cuotas tiene pagas.
            sSql = "SELECT * FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo = 0 AND idCurso = " & id_Curso
            
            adoTemp3.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            cuotas_pagas = 0
            If Not adoTemp3.EOF Then
                adoTemp3.MoveFirst
                Do While Not adoTemp3.EOF
                    cuotas_pagas = cuotas_pagas + 1
                    adoTemp3.MoveNext
                Loop
            End If
            
            adoTemp3.Close
            
            If cuotas_pagas > 0 Then
                cuotas_pagas = cuotas_pagas - 1
            End If
            'FIN - Verifico cuantas cuotas tiene pagas.
            
            
            'INI - Renombro los MOD a ABA para posible futuro uso.
                sSql = "UPDATE Movimientos SET TipoDoc = 'ABA' WHERE TipoDoc = 'MOD' AND idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso & " AND Saldo <> 0"
                adoConnection.Execute sSql
            'FIN - Renombro los MOD a ABA para posible futuro uso.
            
            sSql = "INSERT INTO abandonos (idAlumno, idCurso, fecha, cuotas_pagas) VALUES (" & id_Alumno & ", " & id_Curso & ", '" & Date & "', " & cuotas_pagas & ")"
            adoConnection.Execute sSql
            
            
            sSql = "UPDATE alumnos SET observaciones = observaciones + '**********ABANDONÓ EL CURSO Nº" & num_curso & " DE " & nombre_curso & "' WHERE id = " & id_Alumno
            adoConnection.Execute sSql
            
            adoTempInscriptos.Delete
            
            sSql = "DELETE FROM AlumnosXCurso WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
            adoConnection.Execute sSql
            
            sSql = "UPDATE Cursos SET Vacantes = Vacantes + 1, Inscriptos = Inscriptos - 1 WHERE id = " & id_Curso
            adoConnection.Execute sSql
                                    
            CARGAR_CURSOS "Todas"
            VOLVER_CURSO_ACTUAL
        End If
        
        mensaje_esta_seguro = ""
    'Else
    '    MsgBox "Debe seleccionar un alumno.", vbExclamation, "LISTA DE ALUMNOS"
    'End If
End Sub

Private Sub cmdCambiar_Click()
    'INICIO - Cambiar alumno de curso.
    If mensaje_esta_seguro = "" Then
        mensaje_esta_seguro = "¿Confirma que desea cambiar de curso al alumno?"
    End If
    
    If MsgBox(mensaje_esta_seguro, vbQuestion + vbYesNo, "CAMBIAR ALUMNO") = vbYes Then
        frmCursosDisponibles.Show vbModal
        
        If id_Curso_Nuevo = 0 Then
            Exit Sub
        End If
        
        id_Alumno = DEVOLVER_ID(adoTempInscriptos!Nombre, adoAlumnos, "Alumnos", "Nombre")
        id_Curso = adoTempCursos!id
        
        adoTempInscriptos.Delete
        
        'Saco al alumno del curso.
        sSql = "DELETE FROM AlumnosXCurso WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
        adoConnection.Execute sSql
        
        'Actualizo las vacantes del curso que deja.
        sSql = "UPDATE Cursos SET Vacantes = Vacantes + 1, Inscriptos = Inscriptos - 1 WHERE id = " & id_Curso
        adoConnection.Execute sSql
                                
        'Actualizo AlumnosXCurso
        CERRAR_TABLA adoAlumnosXCurso
        With adoAlumnosXCurso
            .Open "AlumnosXCurso", adoConnection, adOpenDynamic, adLockOptimistic
            .AddNew
            !idAlumno = id_Alumno
            !idCurso = id_Curso_Nuevo
            .Update
            .Close
        End With
        
        'La función que sigue no está terminada.
        'INSCRIBIR_CAMBIO
        
        'Actualizo las vacantes del curso al que pasa.
        sSql = "UPDATE Cursos SET Vacantes = Vacantes - 1, Inscriptos = Inscriptos + 1 WHERE id = " & id_Curso_Nuevo
        adoConnection.Execute sSql
                                
        'Modifico los comprobantes para que sean válidos para el nuevo curso.
        sSql = "UPDATE Movimientos SET idCurso = " & id_Curso_Nuevo & " WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
        adoConnection.Execute sSql
        
        'Actualizo las fechas de vencimiento.
        
        frmCambiarVencimientos.Show vbModal
        

        sSql = "SELECT * FROM Movimientos WHERE idCurso = " & id_Curso_Nuevo & " AND idAlumno = " & id_Alumno
        CERRAR_TABLA adoTemp
        adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If Not adoTemp.EOF Then
            adoTemp.MoveFirst
            
            Do While Not adoTemp.EOF
                sSql = "UPDATE ItemsXMov SET idCurso = " & id_Curso_Nuevo & ", Detalle = Left(Detalle, 12) +  '" & detalle_Curso_Nuevo & "' WHERE idMovimiento = " & adoTemp!id
                adoConnection.Execute sSql
        
                adoTemp.MoveNext
            Loop
        End If
        
        'sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso_Nuevo
        'CERRAR_TABLA adoTemp
        'adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        'For k = 0 To 4
        '    sSql = "SELECT * FROM Movimientos WHERE idCurso = " & id_Curso_Nuevo & " AND idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo > 0 AND Cuota = " & k
        '    CERRAR_TABLA adoTemp2
        '    adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
        '    If Not adoTemp2.EOF Then
        '        adoTemp2!Fecha = adoTemp!Cuota1
        '        adoTemp2.Update
        '    End If
        'Next
    End If
    
    'FIN - Cambiar alumno de curso.
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEliminarEspera_Click()
    'Tengo que escribir código para ver si hay un ítem seleccionado
        'If mensaje_esta_seguro = "" Then
            mensaje_esta_seguro = "¿Confirma que desea quitar al alumno de la lista de espera?"
        'End If
        
        If MsgBox(mensaje_esta_seguro, vbQuestion + vbYesNo, "LISTA DE ESPERA - QUITAR ALUMNO") = vbYes Then
            paso_a_inscriptos = True
            
            id_Alumno = DEVOLVER_ID(adoTempEspera!Nombre, adoAlumnos, "Alumnos", "Nombre")
            id_Curso = adoTempCursos!id
            
            adoTempEspera.Delete
            
            sSql = "DELETE FROM ListaEspera WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
            adoConnection.Execute sSql
        
            If Not se_va_a_inscriptos Then
                ELIMINAR_DOCUMENTOS_INSCRIPTO
            Else
                CONVERTIR_DOCUMENTOS_INSCRIPTO
            End If

            CARGAR_CURSOS "Todas"
            VOLVER_CURSO_ACTUAL
        End If
    'Else
    '    MsgBox "Debe seleccionar un alumno.", vbExclamation, "LISTA DE ALUMNOS"
    'End If
End Sub

Private Sub cmdEliminarInscriptos_Click()
        'If mensaje_esta_seguro = "" Then
            mensaje_esta_seguro = "¿Confirma que desea quitar al alumno de la lista de inscriptos?"
        'End If
        
        If MsgBox(mensaje_esta_seguro, vbQuestion + vbYesNo, "LISTA DE INSCRIPTOS - QUITAR ALUMNO") = vbYes Then
            paso_a_espera = True
            
            id_Alumno = DEVOLVER_ID(adoTempInscriptos!Nombre, adoAlumnos, "Alumnos", "Nombre")
            id_Curso = adoTempCursos!id
            
            adoTempInscriptos.Delete
            
            sSql = "DELETE FROM AlumnosXCurso WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
            adoConnection.Execute sSql
            
            sSql = "UPDATE Cursos SET Vacantes = Vacantes + 1, Inscriptos = Inscriptos - 1 WHERE id = " & id_Curso
            adoConnection.Execute sSql
                                    
            If Not se_va_a_espera Then
                ELIMINAR_DOCUMENTOS
            Else
                CONVERTIR_DOCUMENTOS
            End If
        
            CARGAR_CURSOS "Todas"
            VOLVER_CURSO_ACTUAL
        End If
        
        mensaje_esta_seguro = ""
    'Else
    '    MsgBox "Debe seleccionar un alumno.", vbExclamation, "LISTA DE ALUMNOS"
    'End If
End Sub

Private Sub cmdFactura_Click()
    Unload Me
    sMenu = "FacturaPresenciales"
    frmFactura.Show
End Sub

Private Sub cmdInscribir_Click()
    'Tengo que escribir código para ver si hay un ítem seleccionado
        
        'Obtengo el ID del curso seleccionado
        id_Curso = adoTempCursos!id
        
        'Me fijo que el alumno NO esté inscripto en el curso
        CERRAR_TABLA adoAlumnosXCurso
        sSql = "SELECT * FROM AlumnosXCurso WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
        adoAlumnosXCurso.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        If adoAlumnosXCurso.EOF Then
            
            'Me fijo que TAMPOCO esté en lista de espera
            CERRAR_TABLA adoListaEspera
            sSql = "SELECT * FROM ListaEspera WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
            adoListaEspera.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            If adoListaEspera.EOF Then
            
                'Me fijo si hay vacantes en el curso
                CERRAR_TABLA adoCursos
                sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso
                adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                
                AVISAR_MODS_PENDIENTES
                
                AVISAR_ABANDONO
                
                If adoCursos!Vacantes > 0 Then
                    'INSCRIBIR
                    If Not se_va_a_inscriptos Then
                        frmCuotas.Show vbModal
                        If Not CancelaInscripcion Then
                            INSCRIBIR
                            GENERAR_DOCUMENTOS

                            'CHEQUEAR_A_CUENTA Este no funciona, CHEQUEAR_A_CUENTA_FACTURA si.
                        End If
                    Else
                        INSCRIBIR
        
                        'GENERAR_DOCUMENTOS
                    End If
                    If Not CancelaInscripcion Then
                        se_va_a_espera = False
                    
                        MsgBox "El alumno " & txtAlumno.Text & " ha sido inscripto en el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbInformation, "INSCRIPCIÓN"
                                                
                        'Marco el alumno como USADO.
                        sSql = "UPDATE alumnos SET usado = true WHERE id = " & id_Alumno
                        adoConnection.Execute sSql

                    End If
                Else
                    'INSCRIBIR_LISTA_ESPERA
                    frmCuotas.Show vbModal
                    If Not CancelaInscripcion Then
                        se_va_a_espera = True
                        INSCRIBIR_LISTA_ESPERA
                        GENERAR_DOCUMENTOS
                        MsgBox "El alumno " & txtAlumno.Text & " ha sido puesto en la lista de espera para el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbInformation, "INSCRIPCIÓN"
                    
                                                
                        'Marco el alumno como USADO.
                        sSql = "UPDATE alumnos SET usado = true WHERE id = " & id_Alumno
                        adoConnection.Execute sSql
                    
                    End If
                End If
                
                adoCursos.Close
                
                If optVacantesDisponibles Then
                    CARGAR_CURSOS "Disponibles"
                ElseIf optVacantesEnEspera Then
                    CARGAR_CURSOS "EnEspera"
                ElseIf optVacantesTodas Then
                    CARGAR_CURSOS "Todas"
                End If
                                    
                VOLVER_CURSO_ACTUAL
                
                VER_INSCRIPTOS
                VER_ESPERA
                
                'AVISAR_MODS_PENDIENTES
            Else
                MsgBox "El alumno " & txtAlumno.Text & " está en lista de espera para el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbCritical, "INSCRIPCIÓN"
            End If
        Else
            MsgBox "El alumno " & txtAlumno.Text & " ya está inscripto en el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbCritical, "INSCRIPCIÓN"
        End If
    'Else
    '    MsgBox "Debe seleccionar un curso.", vbExclamation, "INSCRIPCIÓN"
    'End If
End Sub

Private Sub cmdInscribirRecursante_Click()
    'Tengo que escribir código para ver si hay un ítem seleccionado
        
        'Obtengo el ID del curso seleccionado
        id_Curso = adoTempCursos!id
        
        'Me fijo que el alumno NO esté inscripto en el curso
        CERRAR_TABLA adoAlumnosXCurso
        sSql = "SELECT * FROM AlumnosXCurso WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
        adoAlumnosXCurso.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        If adoAlumnosXCurso.EOF Then
            
            'Me fijo que TAMPOCO esté en lista de espera
            CERRAR_TABLA adoListaEspera
            sSql = "SELECT * FROM ListaEspera WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
            adoListaEspera.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            If adoListaEspera.EOF Then
            
                'Me fijo si hay vacantes en el curso
                CERRAR_TABLA adoCursos
                sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso
                adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                
                AVISAR_MODS_PENDIENTES
                
                If adoCursos!Vacantes > 0 Then
                    'INSCRIBIR
                    If Not se_va_a_inscriptos Then
                        'frmCuotas.Show vbModal
                        If Not CancelaInscripcion Then
                            INSCRIBIR
                            'GENERAR_DOCUMENTOS
                            
                            'CHEQUEAR_A_CUENTA Este no funciona, CHEQUEAR_A_CUENTA_FACTURA si.
                        End If
                    Else
                        INSCRIBIR
                        'GENERAR_DOCUMENTOS
                    End If
                    If Not CancelaInscripcion Then
                        se_va_a_espera = False
                    
                        MsgBox "El alumno " & txtAlumno.Text & " ha sido inscripto en el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbInformation, "INSCRIPCIÓN"
                    End If
                Else
                    'INSCRIBIR_LISTA_ESPERA
                    'frmCuotas.Show vbModal
                    If Not CancelaInscripcion Then
                        se_va_a_espera = True
                        INSCRIBIR_LISTA_ESPERA
                        'GENERAR_DOCUMENTOS
                        MsgBox "El alumno " & txtAlumno.Text & " ha sido puesto en la lista de espera para el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbInformation, "INSCRIPCIÓN"
                    End If
                End If
                
                adoCursos.Close
                
                If optVacantesDisponibles Then
                    CARGAR_CURSOS "Disponibles"
                ElseIf optVacantesEnEspera Then
                    CARGAR_CURSOS "EnEspera"
                ElseIf optVacantesTodas Then
                    CARGAR_CURSOS "Todas"
                End If
                                    
                VOLVER_CURSO_ACTUAL
                
                VER_INSCRIPTOS
                VER_ESPERA
                
                'AVISAR_MODS_PENDIENTES
            Else
                MsgBox "El alumno " & txtAlumno.Text & " está en lista de espera para el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbCritical, "INSCRIPCIÓN"
            End If
        Else
            MsgBox "El alumno " & txtAlumno.Text & " ya está inscripto en el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbCritical, "INSCRIPCIÓN"
        End If
    'Else
    '    MsgBox "Debe seleccionar un curso.", vbExclamation, "INSCRIPCIÓN"
    'End If
End Sub

Private Sub cmdNuevoAlumno_Click()
    frmAlumnos.Show vbModal
End Sub

Private Sub cmdPasarAEspera_Click()
    se_va_a_espera = True
    mensaje_esta_seguro = "¿Confirma que desea enviar a este alumno a la lista de espera?"
    cmdEliminarInscriptos_Click
    If paso_a_espera Then
        INSCRIBIR_LISTA_ESPERA
        VER_ESPERA
        CARGAR_CURSOS "Todas"
        VOLVER_CURSO_ACTUAL
    End If
    paso_a_espera = False
End Sub

Private Sub cmdPasarAInscriptos_Click()
    se_va_a_inscriptos = True
    mensaje_esta_seguro = "¿Confirma que desea enviar a este alumno a la lista de inscriptos?"
    cmdEliminarEspera_Click
    If paso_a_inscriptos Then
        cmdInscribir_Click
    End If
    paso_a_inscriptos = False
    se_va_a_inscriptos = False
End Sub

Private Sub cmdSeleccionarAlumno_Click()
    EstiloBuscador = "Alumnos"
    frmBuscador.Show vbModal
    CHEQUEAR_DEUDA
End Sub

Private Sub cmdVerInscriptos_Click()
    VER_INSCRIPTOS
End Sub

Private Sub cmdVerListaEspera_Click()
    VER_ESPERA
End Sub

Private Sub dbgCursos_Click()
    If Not adoTempCursos.EOF Then
        VER_INSCRIPTOS
        
        VER_ESPERA
        
        If adoTempCursos!Vacantes = 0 Then
            cmdPasarAInscriptos.Enabled = False
        Else
            cmdPasarAInscriptos.Enabled = True
        End If
    End If
End Sub

Private Sub dbgEspera_Click()
    cmdEliminarInscriptos.Enabled = False
    cmdEliminarEspera.Enabled = True
    
    If Not adoTempEspera.EOF Then
        txtAlumno.Text = adoTempEspera!Nombre
    End If
End Sub

Private Sub dbgInscriptos_Click()
    cmdEliminarEspera.Enabled = False
    cmdEliminarInscriptos.Enabled = True
    
    If Not adoTempInscriptos.EOF Then
        txtAlumno.Text = adoTempInscriptos!Nombre
        x_alumno_form_inscrip_factura = adoTempInscriptos!Nombre
    End If
End Sub

Private Sub Form_Load()
    EventoLoad = True
    CARGAR_COMBO "cboTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    EventoLoad = False

    x_hay_cuotas_pagas = False

    VACIAR_GRILLA_CURSOS
    VACIAR_GRILLA_INSCRIPTOS
    VACIAR_GRILLA_ESPERA
    
    If xVieneDe = "Alumnos" Then
        xVieneDe = ""
        
        txtAlumno.Text = DEVOLVER_CAMPO(id_Alumno, adoAlumnos, "Alumnos", "Nombre")
        CHEQUEAR_DEUDA
    End If
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

Private Sub VACIAR_GRILLA_ESPERA()
    CERRAR_TABLA adoTempEspera
    sSql = "DELETE FROM TempEspera"
    adoConnection.Execute sSql
    adoTempEspera.Open "TempEspera", adoConnection, adOpenKeyset, adLockOptimistic
    Set dbgEspera.DataSource = adoTempEspera
    
    FORMATO_GRILLA_ESPERA
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

Private Sub FORMATO_GRILLA_ESPERA()
    With dbgEspera
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

Private Sub INSCRIBIR()
    'Actualizo inscriptos y vacantes del curso
    With adoCursos
        !inscriptos = adoCursos!inscriptos + 1
        !Vacantes = adoCursos!Vacantes - 1
    
        .Update
    End With
    
    'Actualizo AlumnosXCurso
    CERRAR_TABLA adoAlumnosXCurso
    With adoAlumnosXCurso
        .Open "AlumnosXCurso", adoConnection, adOpenDynamic, adLockOptimistic
        .AddNew
        !idAlumno = id_Alumno
        !idCurso = id_Curso
        .Update
        .Close
    End With
End Sub

Private Sub INSCRIBIR_LISTA_ESPERA()
    'Actualizo ListaEspera
    CERRAR_TABLA adoListaEspera
    With adoListaEspera
        .Open "ListaEspera", adoConnection, adOpenDynamic, adLockOptimistic
        .AddNew
        !idAlumno = id_Alumno
        !idCurso = id_Curso
        !FechaDesde = Date
        !Observaciones = "" 'Hay que armar un TextBox para este campo
        .Update
        .Close
    End With
End Sub

Private Sub GENERAR_DOCUMENTOS()
    Dim ultimo_id As Long
'   Dim ultimo_id As Integer
    
    For k = 0 To x_cuotas
        'Genero el encabezado
        With adoMovimientos
            CERRAR_TABLA adoMovimientos
            .Open "Movimientos", adoConnection, adOpenDynamic, adLockOptimistic
            
            .AddNew
            
            !Sucursal = "0000" 'TODAVÍA NO SE A QUE SUCURSAL VA
            
            If se_va_a_espera Then
                !TipoDoc = "ESP"
                !NumDoc = ULTIMO_NUMERO("ESP")
            Else
                !TipoDoc = "MOD"
                !NumDoc = ULTIMO_NUMERO("MOD")
            End If
            
            
            !idAlumno = id_Alumno
            !idCurso = id_Curso
            
            If k = 0 Then 'Es matrícula
                If Not x_hay_cuotas_pagas Then
                    !fecha = Date
                    '************##################Calcular aca el descuento##########***********
                    '!Subtotal = x_valorMatricula
                    '!Iva = 0
                    '!Total = x_valorMatricula
                    '!Saldo = x_valorMatricula
                    '!Descuento = x_DetalleDescuento
                                        
                    !Subtotal = x_valorMatriculaReal
                    !Iva = 0
                    !Total = x_valorMatriculaReal
                    !Saldo = x_valorMatriculaReal
                    !Descuento = x_DetalleDescuento
                End If
            Else
                CERRAR_TABLA adoTemp
                sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso
                adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                
                campo = "Cuota" & k
                '!fecha = adoTemp.Fields(campo).Value
                !fecha = Date
                
                adoTemp.Close
                
                '!Subtotal = x_valorCuota
                '!Iva = 0
                '!Total = x_valorCuota
                '!Saldo = x_valorCuota
                '!Descuento = x_DetalleDescuento
                '!cuota = k
                
                !Subtotal = x_valorCuotaReal
                !Iva = 0
                !Total = x_valorCuotaReal
                !Saldo = x_valorCuotaReal
                !Descuento = x_DetalleDescuento
                !cuota = k
                
                
                
                
            End If
            
            .Update
            
            .MoveLast
            ultimo_id = !id
            
            .Close
        End With
        
        'Genero el cuerpo
        With adoItemsXMov
            CERRAR_TABLA adoItemsXMov
            .Open "ItemsXMov", adoConnection, adOpenDynamic, adLockOptimistic
            
            .AddNew
            
            !idMovimiento = ultimo_id
            !idCurso = id_Curso
            !Cantidad = 1
            
            If k = 0 Then 'Es matrícula
                !Detalle = "Matrícula - " & cboTipoCurso.Text & " " & Detalle_
                
                'If x_valorMatriculaReal = x_valorMatricula Then
                    !Unitario = x_valorMatriculaReal
                'Else
                '    !Unitario = x_valorMatricula
                'End If
                
                If x_DetalleDescuento = "EX ALUMNO" Then
                    !Descuento = 15
                    x_valorMatricula = x_valorMatriculaReal - (x_valorMatriculaReal * 15 / 100)
                End If
                
                !Importe = x_valorMatricula
                !Saldo = x_valorMatricula
            Else
                !Detalle = "Cuota " & k & "/" & x_cuotas & " - " & cboTipoCurso.Text & " " & Detalle_
                
                'If x_valorCuotaReal = x_valorCuota Then
                    !Unitario = x_valorCuotaReal
                'Else
                '    !Unitario = x_valorCuota
                'End If
                
                If x_DetalleDescuento = "EX ALUMNO" Then
                    !Descuento = 15
                    x_valorCuota = x_valorCuotaReal - (x_valorCuotaReal * 15 / 100)
                End If
                
                !Importe = x_valorCuota
                !Saldo = x_valorCuota
            End If
            
            .Update

            .Close
        End With
    Next
End Sub

Private Sub VER_INSCRIPTOS()
    'Tengo que escribir código para ver si hay un ítem seleccionado
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
        
        'If banVolverAListaAlumnos Then
        '    banVolverAListaAlumnos = False
        '    cmdVerListaEspera_Click
        'End If
    'Else
    '    MsgBox "Debe seleccionar un curso.", vbExclamation, "INSCRIPCIÓN"
    'End If
End Sub

Private Sub VER_ESPERA()
    'Tengo que escribir código para ver si hay un ítem seleccionado
        If adoTempCursos.EOF Then
            Exit Sub
        End If
        
        VACIAR_GRILLA_ESPERA
        
        TipoLista = "Espera"
        
        id_ = adoTempCursos!id
        
        CERRAR_TABLA adoTabla
        sSql = "SELECT ListaEspera.id, Alumnos.Nombre, Alumnos.Telefono, Alumnos.Celular, Alumnos.Mail FROM Alumnos, ListaEspera WHERE Alumnos.id = ListaEspera.idAlumno AND ListaEspera.idCurso = " & adoTempCursos!id & " ORDER BY FechaDesde"
        adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        Do While Not adoTabla.EOF
            adoTempEspera.AddNew
            adoTempEspera!id = adoTabla!id
            adoTempEspera!Nombre = adoTabla!Nombre
            adoTempEspera.Fields("Teléfono").Value = adoTabla!Telefono
            adoTempEspera!Celular = adoTabla!Celular
            adoTempEspera.Fields("e-Mail").Value = adoTabla!Mail
            
            a_curso = adoTempCursos!id
            a_alumno = adoTabla!id
            
            CERRAR_TABLA adoTemp2
            sSql = "SELECT SUM(Total) AS Importe FROM Movimientos WHERE idAlumno = " & a_alumno & " AND idCurso = " & a_curso & " AND Cuota = 0 AND Left(TipoDoc, 2) = 'FC'"
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            adoTempEspera.Fields("Matrícula").Value = adoTemp2!Importe
            
            CERRAR_TABLA adoTemp2
            sSql = "SELECT SUM(Total) AS Importe FROM Movimientos WHERE idAlumno = " & a_alumno & " AND idCurso = " & a_curso & " AND Cuota <> 0 AND Left(TipoDoc, 2) = 'FC'"
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            adoTempEspera!Cuotas = adoTemp2!Importe
            
            adoTempEspera.Update
            
            adoTabla.MoveNext
        Loop
        adoTabla.Close
        
        If Not adoTempEspera.EOF Then
            adoTempEspera.MoveFirst
        End If
        
        'If banVolverAListaAlumnos Then
        '    banVolverAListaAlumnos = False
        '    cmdVerListaEspera_Click
        'End If
    'Else
    '    MsgBox "Debe seleccionar un curso.", vbExclamation, "INSCRIPCIÓN"
    'End If
End Sub

Private Sub ELIMINAR_DOCUMENTOS()
    With adoMovimientos
        CERRAR_TABLA adoMovimientos
        sSql = "SELECT * FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
        .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                sSql = "DELETE FROM ItemsXMov WHERE idMovimiento = " & !id
                adoConnection.Execute sSql
                
                .MoveNext
            Loop
            
            .Close
        End If
        
        sSql = "DELETE FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
        adoConnection.Execute sSql
    End With
End Sub

Private Sub ELIMINAR_DOCUMENTOS_INSCRIPTO()
    With adoMovimientos
        CERRAR_TABLA adoMovimientos
        sSql = "SELECT * FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
        .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                sSql = "DELETE FROM ItemsXMov WHERE idMovimiento = " & !id
                adoConnection.Execute sSql
                
                .MoveNext
            Loop
            
            .Close
        End If
        
        sSql = "DELETE FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
        adoConnection.Execute sSql
    End With
End Sub

Private Sub VOLVER_CURSO_ACTUAL()
    With adoTempCursos
        .MoveFirst
        Do While !id <> id_Curso
            .MoveNext
        Loop
    End With
End Sub

Private Sub CONVERTIR_DOCUMENTOS()
    sSql = "UPDATE Movimientos SET TipoDoc = 'ESP' WHERE TipoDoc = 'MOD' AND idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
    adoConnection.Execute sSql
End Sub

Private Sub CONVERTIR_DOCUMENTOS_INSCRIPTO()
    sSql = "UPDATE Movimientos SET TipoDoc = 'MOD' WHERE TipoDoc = 'ESP' AND idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso
    adoConnection.Execute sSql
End Sub

Private Sub CHEQUEAR_DEUDA()
    sSql = "SELECT Movimientos.*, ItemsXMov.Detalle FROM Movimientos, ItemsXMov WHERE Movimientos.idAlumno = " & id_Alumno & " AND Left(Movimientos.TipoDoc, 2) = 'FC' AND Movimientos.Saldo > 0 AND Movimientos.idCurso <> 0 AND Left(ItemsXMov.Detalle, 9) <> ' A CUENTA' AND Movimientos.id = ItemsXMov.idMovimiento"
    CERRAR_TABLA adoTemp
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoTemp.EOF Then
        MsgBox "Alumno con deuda anterior.", vbInformation, "ALUMNO CON DEUDA"
    End If
    
    adoTemp.Close
End Sub


Private Sub AVISAR_MODS_PENDIENTES()
    sSql = "SELECT * FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo > 0"
    CERRAR_TABLA adoTemp
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoTemp.EOF Then
        MsgBox "ATENCIÓN!" & vbCrLf & "El alumno " & txtAlumno.Text & " tiene MODs pendientes.", vbExclamation, "ATENCIÓN !"
    End If
    
    adoTemp.Close
End Sub


Private Sub AVISAR_ABANDONO()
    sSql = "SELECT * FROM Abandonos WHERE idAlumno = " & id_Alumno
    CERRAR_TABLA adoTemp
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    curso_abandonado = ""
    mensaje_abandono = ""
    
    x_abandono_1 = ""
    x_abandono_2 = ""
    x_abandono_3 = ""
    
    If Not adoTemp.EOF Then
        adoTemp.MoveFirst
        
        cant = 1
        Do While Not adoTemp.EOF
            sSql = "SELECT TiposCurso.Detalle FROM Cursos, TiposCurso WHERE Cursos.id = " & adoTemp!idCurso & " AND Cursos.idTipoCurso = TiposCurso.id"
            CERRAR_TABLA adoTemp2
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
            curso_abandonado = adoTemp2!Detalle
            adoTemp2.Close
    
            If adoTemp!cuotas_pagas <> 0 Then
                mensaje_abandono = mensaje_abandono & vbCrLf & adoTemp!cuotas_pagas & " cuotas pagas del curso " & curso_abandonado & "."
            Else
                mensaje_abandono = ""
            End If
            
            For i = 1 To adoTemp!cuotas_pagas
                If cant = 1 Then
                    x_abandono_1 = "Cuota 1 de " & adoTemp!cuotas_pagas & " del curso " & curso_abandonado & "."
                    cant = cant + 1
                ElseIf cant = 2 Then
                    x_abandono_2 = "Cuota 2 de " & adoTemp!cuotas_pagas & " del curso " & curso_abandonado & "."
                    cant = cant + 1
                ElseIf cant = 3 Then
                    x_abandono_3 = "Cuota 3 de " & adoTemp!cuotas_pagas & " del curso " & curso_abandonado & "."
                    cant = cant + 1
                Else
                    cant = cant + 1
                End If
            Next
            
            adoTemp.MoveNext
        Loop
    End If
    
    If mensaje_abandono <> "" Then
        MsgBox "ATENCIÓN!" & vbCrLf & vbCrLf & "El alumno " & txtAlumno.Text & " tiene: " & vbCrLf & mensaje_abandono, vbExclamation, "ATENCIÓN !"
    End If
    
    adoTemp.Close
End Sub

Private Sub INSCRIBIR_CAMBIOS()
    id_Curso = id_Curso_Nuevo
    
    'Me fijo que el alumno NO esté inscripto en el curso
    CERRAR_TABLA adoAlumnosXCurso
    sSql = "SELECT * FROM AlumnosXCurso WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
    adoAlumnosXCurso.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If adoAlumnosXCurso.EOF Then
        
        'Me fijo que TAMPOCO esté en lista de espera
        CERRAR_TABLA adoListaEspera
        sSql = "SELECT * FROM ListaEspera WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
        adoListaEspera.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        If adoListaEspera.EOF Then
        
            AVISAR_MODS_PENDIENTES
            
            If adoCursos!Vacantes > 0 Then
                'INSCRIBIR
                If Not se_va_a_inscriptos Then
                    frmCuotas.Show vbModal
                    If Not CancelaInscripcion Then
                        INSCRIBIR
                        GENERAR_DOCUMENTOS
                        
                        'CHEQUEAR_A_CUENTA Este no funciona, CHEQUEAR_A_CUENTA_FACTURA si.
                    End If
                Else
                    INSCRIBIR
                    'GENERAR_DOCUMENTOS
                End If
                If Not CancelaInscripcion Then
                    se_va_a_espera = False
                
                    MsgBox "El alumno " & txtAlumno.Text & " ha sido inscripto en el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbInformation, "INSCRIPCIÓN"
                End If
            End If
            
            adoCursos.Close
        Else
            MsgBox "El alumno " & txtAlumno.Text & " está en lista de espera para el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbCritical, "INSCRIPCIÓN"
        End If
    Else
        MsgBox "El alumno " & txtAlumno.Text & " ya está inscripto en el curso " & cboTipoCurso.Text & " los " & adoTempCursos!Horario, vbCritical, "INSCRIPCIÓN"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    x_alumno_form_inscrip_factura = ""
End Sub
