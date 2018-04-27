VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTiempo 
      Interval        =   1
      Left            =   240
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar barTiempo 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "C.E.T.I.A."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   4800
      TabIndex        =   2
      Top             =   2880
      Width           =   2250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "de Institutos de Enseñanza v6.8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sistema de Gestión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4170
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    On Error GoTo ErrorHandle
    
    Set adoConnection = New ADODB.Connection
    
    Set adoAlumnos = New ADODB.Recordset
    Set adoAlumnosXCurso = New ADODB.Recordset
    Set adoAnulados = New ADODB.Recordset
    Set adoAulas = New ADODB.Recordset
    Set adoClasesXCurso = New ADODB.Recordset
    Set adoClasesXProfe = New ADODB.Recordset
    Set adoComoLlego = New ADODB.Recordset
    Set adoCompaniasCelular = New ADODB.Recordset
    Set adoCondIva = New ADODB.Recordset
    Set adoCursos = New ADODB.Recordset
    Set adoCursosDisponibles = New ADODB.Recordset
    Set adoCursosXProfesor = New ADODB.Recordset
    Set adoDuraciones = New ADODB.Recordset
    Set adoEmisores = New ADODB.Recordset
    Set adoEmpresas = New ADODB.Recordset
    Set adoEstadoCursosXAlumno = New ADODB.Recordset
    Set adoFormasPago = New ADODB.Recordset
    Set adoHorarios = New ADODB.Recordset
    Set adoItemsXMov = New ADODB.Recordset
    Set adoListaEspera = New ADODB.Recordset
    Set adoLocalidades = New ADODB.Recordset
    Set adoLog = New ADODB.Recordset
    Set adoMovimientos = New ADODB.Recordset
    Set adoNumeracion = New ADODB.Recordset
    Set adoNumeracionCursos = New ADODB.Recordset
    Set adoPrecios = New ADODB.Recordset
    Set adoProfesores = New ADODB.Recordset
    Set adoProximosInicios = New ADODB.Recordset
    Set adoRecargosTarjeta = New ADODB.Recordset
    Set adoSucursales = New ADODB.Recordset
    Set adoTiposComprobante = New ADODB.Recordset
    Set adoTiposCurso = New ADODB.Recordset
    Set adoTiposDoc = New ADODB.Recordset
    Set adoUnidadesNegocio = New ADODB.Recordset
    Set adoUsuarios = New ADODB.Recordset

    
    Set adoTabla = New ADODB.Recordset
    Set adoTablaBus = New ADODB.Recordset
    Set adoTablaValidacion = New ADODB.Recordset
    
    Set adoTemp = New ADODB.Recordset
    Set adoTemp2 = New ADODB.Recordset
    Set adoTemp3 = New ADODB.Recordset
    
    Set adoTempAlumnos = New ADODB.Recordset
    Set adoTempClases = New ADODB.Recordset
    Set adoTempCursos = New ADODB.Recordset
    Set adoTempEspera = New ADODB.Recordset
    Set adoTempEstadoCursos = New ADODB.Recordset
    Set adoTempFactura = New ADODB.Recordset
    Set adoTempInscriptos = New ADODB.Recordset
    Set adoTempProfesores = New ADODB.Recordset
    
    connectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=MS Access Database;Initial Catalog=" & App.Path & "\Cetia_Dani.mdb;Password=FiatIdea"
    'connectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=MS Access Database;Initial Catalog=" & App.Path & "\Cetia.mdb;Password=FiatIdea"
    adoConnection.Open connectionString
    
    Exit Sub
ErrorHandle:
    MsgBox "No es posible acceder a la Base de Datos.", vbCritical, "SE HA PRODUCIDO UN ERROR"
    End
End Sub

Private Sub tmrTiempo_Timer()
    barTiempo.Value = barTiempo.Value + 1
    If barTiempo.Value = 30 Then
        Unload Me
        frmAcceso.Show vbModal
    End If
End Sub
