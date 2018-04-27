VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmListaAlumnosCursosXIniciar 
   Caption         =   "LISTADO - Cursos por iniciar"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cboDesde 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "cboDesde"
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cboHasta 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cboHasta"
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   4080
      Picture         =   "frmListaAlumnosCursosXIniciar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   3360
      Picture         =   "frmListaAlumnosCursosXIniciar.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Acceder"
      Top             =   1320
      Width           =   615
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   2640
      Top             =   1440
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4680
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmListaAlumnosCursosXIniciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    sSql = "DELETE FROM zzCursosXIniciar"
    adoConnection.Execute sSql
    
    'Inscriptos
    sSql = "INSERT INTO zzCursosXIniciar (Alumno, Celular, Telefono, FechaIni, Horario, TipoCurso, Espera, NumCurso, Grupo) " & _
           "SELECT Alumnos.Nombre, Alumnos.Celular, Alumnos.Telefono, Cursos.FechaIni, Horarios.Detalle, TiposCurso.Detalle, 'Inscriptos', Cursos.Numero, TiposCurso.Detalle + '(' + Str(Cursos.Numero) + ')' " & _
           "FROM Alumnos, Cursos, Horarios, TiposCurso, AlumnosXCurso " & _
           "WHERE Alumnos.id = AlumnosXCurso.idAlumno AND Cursos.id = AlumnosXCurso.idCurso AND Cursos.idHorario = Horarios.id AND Cursos.idTipoCurso = TiposCurso.id AND Cursos.Abierto " & _
           "      AND (TiposCurso.Detalle >= '" & cboDesde.Text & "' AND TiposCurso.Detalle <= '" & cboHasta.Text & "') " & _
           "      AND FechaIni > DateValue('" & Date & "')"

    adoConnection.Execute sSql
    
    'En espera
    sSql = "INSERT INTO zzCursosXIniciar (Alumno, Celular, Telefono, FechaIni, Horario, TipoCurso, Espera, NumCurso, Grupo) " & _
           "SELECT Alumnos.Nombre, Alumnos.Celular, Alumnos.Telefono, Cursos.FechaIni, Horarios.Detalle, TiposCurso.Detalle, 'En espera', Cursos.Numero, TiposCurso.Detalle + '(' + Str(Cursos.Numero) + ')' " & _
           "FROM Alumnos, Cursos, Horarios, TiposCurso, ListaEspera " & _
           "WHERE Alumnos.id = ListaEspera.idAlumno AND Cursos.id = ListaEspera.idCurso AND Cursos.idHorario = Horarios.id AND Cursos.idTipoCurso = TiposCurso.id AND Cursos.Abierto " & _
           "      AND (TiposCurso.Detalle >= '" & cboDesde.Text & "' AND TiposCurso.Detalle <= '" & cboHasta.Text & "') " & _
           "      AND FechaIni > DateValue('" & Date & "')"
           
    adoConnection.Execute sSql
   
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    
    rptListado.ReportFileName = App.Path & "\reportes\rptAlumnosXIniciar.rpt"
    
    rptListado.ReportTitle = "LISTADO DE ALUMNOS DE CURSOS POR INICIAR"
    rptListado.Action = 1

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CARGAR_COMBO "cboDesde", adoTiposCurso, "TiposCurso", "Detalle", Me
    CARGAR_COMBO "cboHasta", adoTiposCurso, "TiposCurso", "Detalle", Me
    cboHasta.ListIndex = cboHasta.ListCount - 1
End Sub

