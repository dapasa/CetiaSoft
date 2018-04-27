VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaAlumnosXCursosRealizados 
   Caption         =   "LISTADO - Alumnos por curso realizado"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   4575
      Begin VB.ComboBox cboHastaAlumno 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cboDesdeAlumno 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "cboDesdeAlumno"
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Horario"
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
      Left            =   4800
      TabIndex        =   15
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboDesdeHorario 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "cboDesdeHorario"
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cboHastaHorario 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cboHastaHorario"
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Curso"
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cboDesdeTipoCurso 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "cboDesdeTipoCurso"
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cboHastaTipoCurso 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cboHastaTipoCurso"
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   420
      End
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
      Left            =   4800
      TabIndex        =   10
      Top             =   1200
      Width           =   3855
      Begin VB.OptionButton optAbierto 
         Caption         =   "Abierto"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optCerrado 
         Caption         =   "Cerrado"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   7320
      Picture         =   "frmListaAlumnosXCursosRealizados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Acceder"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   8040
      Picture         =   "frmListaAlumnosXCursosRealizados.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancelar"
      Top             =   2400
      Width           =   615
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   6600
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowTitle     =   "SISTEMA DE GESTIÓN DE INSTITUTOS EDUCATIVOS v1.0 - LISTADO"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   8640
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "frmListaAlumnosXCursosRealizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    sSql = "DELETE FROM zzAlumnosXCursoRealizado"
    adoConnection.Execute sSql
    
    If optAbierto.Value = True Then
        filtroCurso = "(Abierto = True)"
    ElseIf optCerrado.Value = True Then
        filtroCurso = "(Abierto = False)"
    Else
        filtroCurso = "(1 = 1)"
    End If
    
    sSql = "INSERT INTO zzAlumnosXCursoRealizado (TipoCurso, NumCurso, Horario, Alumno, Celular, Telefono, TelefonoLaboral, Mail) " & _
           "SELECT TiposCurso.Detalle, Cursos.Numero, Horarios.Detalle, Alumnos.Nombre, Alumnos.Celular, Alumnos.Telefono, Alumnos.TelefonoLaboral, Alumnos.Mail " & _
           "FROM Alumnos, TiposCurso, Horarios, Cursos, AlumnosXCurso " & _
           "WHERE TiposCurso.id = Cursos.idTipoCurso AND Horarios.id = Cursos.idHorario " & _
           "      AND Alumnos.id = AlumnosXCurso.idAlumno AND Cursos.id = AlumnosXCurso.idCurso " & _
           "      AND (TiposCurso.Detalle >= '" & cboDesdeTipoCurso.Text & "' AND TiposCurso.Detalle <= '" & cboHastaTipoCurso.Text & "') " & _
           "      AND (Horarios.Detalle >= '" & cboDesdeHorario.Text & "' AND Horarios.Detalle <= '" & cboHastaHorario.Text & "') " & _
           "      AND (Alumnos.Nombre >= '" & cboDesdeAlumno.Text & "' AND Alumnos.Nombre <= '" & cboHastaAlumno.Text & "') " & _
           "      AND " & filtroCurso & _
           "ORDER BY TiposCurso.Detalle, Cursos.Numero, Alumnos.Nombre"
    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    rptListado.ReportFileName = App.Path & "\reportes\rptAlumnosXCursoRealizado.rpt"
    rptListado.ReportTitle = "LISTADO DE ALUMNOS POR CURSO REALIZADO"
    rptListado.Action = 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CARGAR_COMBO "cboDesdeTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    CARGAR_COMBO "cboHastaTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    cboHastaTipoCurso.ListIndex = cboHastaTipoCurso.ListCount - 1
    
    CARGAR_COMBO "cboDesdeHorario", adoHorarios, "Horarios", "Detalle", Me
    CARGAR_COMBO "cboHastaHorario", adoHorarios, "Horarios", "Detalle", Me
    cboHastaHorario.ListIndex = cboHastaHorario.ListCount - 1

    CARGAR_COMBO "cboDesdeAlumno", adoAlumnos, "Alumnos", "Nombre", Me
    CARGAR_COMBO "cboHastaAlumno", adoAlumnos, "Alumnos", "Nombre", Me
    cboHastaAlumno.ListIndex = cboHastaAlumno.ListCount - 1
End Sub

