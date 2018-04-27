VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaMatriculas 
   Caption         =   "LISTADO - Matrículas"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Fecha"
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
      TabIndex        =   12
      Top             =   120
      Width           =   3135
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   330
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Format          =   17563649
         CurrentDate     =   40147
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   330
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Format          =   17563649
         CurrentDate     =   40147
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   465
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
      TabIndex        =   9
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cboHastaTipoCurso 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "cboHastaTipoCurso"
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cboDesdeTipoCurso 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cboDesdeTipoCurso"
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   6600
      Picture         =   "frmListaMatriculas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Acceder"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   7320
      Picture         =   "frmListaMatriculas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancelar"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agrupar por ¿cómo llegó?"
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
      Top             =   1200
      Width           =   2535
      Begin VB.OptionButton optAgruparSi 
         Caption         =   "Si"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optAgruparNo 
         Caption         =   "No"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   5880
      Top             =   2160
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
      X2              =   7920
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmListaMatriculas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    sSql = "DELETE FROM zzMatriculas"
    adoConnection.Execute sSql
    
    sSql = "INSERT INTO zzMatriculas (Fecha, ComoLlego, TipoCurso, Alumno) " & _
           "SELECT Movimientos.Fecha, ComoLlego.Detalle, TiposCurso.Detalle, Alumnos.id " & _
           "FROM Movimientos, TiposCurso, ComoLlego, Cursos, Alumnos " & _
           "WHERE (Movimientos.Fecha >= DateValue('" & dtpFechaDesde.Value & "') AND Movimientos.Fecha <= DateValue('" & dtpFechaHasta.Value & "')) " & _
           "      AND (Movimientos.Cuota = 0) " & _
           "      AND (Movimientos.idAlumno = Alumnos.id) " & _
           "      AND (Alumnos.idComoLlego = ComoLlego.id) " & _
           "      AND (Movimientos.idCurso = Cursos.id) " & _
           "      AND (Cursos.idTipoCurso = TiposCurso.id) " & _
           "      AND (Movimientos.TipoDoc = 'MOD') " & _
           "      AND (TiposCurso.Detalle >= '" & cboDesdeTipoCurso.Text & "' AND TiposCurso.Detalle <= '" & cboHastaTipoCurso.Text & "') " & _
           "ORDER BY Movimientos.Fecha, TiposCurso.Detalle, ComoLlego.Detalle"
    adoConnection.Execute sSql
    
    '"      AND (Movimientos.TipoDoc = 'REC' OR Movimientos.TipoDoc = 'FCA' OR Movimientos.TipoDoc = 'FCB' OR Movimientos.TipoDoc = 'FCC') "
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    
    If optAgruparSi.Value Then
        rptListado.ReportFileName = App.Path & "\reportes\rptMatriculas.rpt"
    Else
        rptListado.ReportFileName = App.Path & "\reportes\rptMatriculas2.rpt"
    End If
    
    rptListado.ReportTitle = "LISTADO DE INSCRIPTOS DESDE " & dtpFechaDesde.Value & " HASTA " & dtpFechaHasta.Value
    rptListado.Action = 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtpFechaDesde.Value = Date
    dtpFechaHasta.Value = Date
    
    CARGAR_COMBO "cboDesdeTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    CARGAR_COMBO "cboHastaTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    cboHastaTipoCurso.ListIndex = cboHastaTipoCurso.ListCount - 1
End Sub


