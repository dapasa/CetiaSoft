VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaAbandonosXCurso 
   Caption         =   "LISTADO - Abandonos por curso"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Nº de curso"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtNumCurso 
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   4800
      Picture         =   "frmListaAbandonosXCurso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   4080
      Picture         =   "frmListaAbandonosXCurso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Acceder"
      Top             =   960
      Width           =   615
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboTipoCurso 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   3360
      Top             =   1080
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
      X2              =   5400
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmListaAbandonosXCurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If txtNumCurso.Text = "" Then
        MsgBox "Debe indicar un Nº de curso.", vbCritical, "ERROR"
        Exit Sub
    End If
    
    sSql = "DELETE FROM zzAbandonosXCurso"
    adoConnection.Execute sSql
    
    id_tipo_curso = DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "detalle")
    
    sSql = "INSERT INTO zzAbandonosXCurso (TipoCurso, NumCurso, Horario, Alumno, YaPago, NoPago) " & _
           "SELECT TiposCurso.Detalle, Cursos.Numero, Horarios.Detalle, Alumnos.Nombre, 0, 0 " & _
           "FROM Alumnos, TiposCurso, Horarios, Cursos, AlumnosXCurso " & _
           "WHERE Cursos.Numero = " & txtNumCurso.Text & " AND " & _
           "  Cursos.idTipoCurso = " & id_tipo_curso & _
           " ORDER BY TiposCurso.Detalle, Cursos.Numero, Alumnos.Nombre"
    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    'rptListado.Connect = "PWD=FiatIdea"
    'rptListado.ReportFileName = App.Path & "\reportes\rptAlumnosXCursoRealizado.rpt"
    'rptListado.ReportTitle = "LISTADO DE ALUMNOS POR CURSO REALIZADO"
    'rptListado.Action = 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CARGAR_COMBO "cboTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    cboTipoCurso.ListIndex = cboTipoCurso.ListCount - 1
End Sub



