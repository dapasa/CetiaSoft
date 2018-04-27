VERSION 5.00
Begin VB.Form frmFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FILTRO"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   Icon            =   "frmFiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   3840
      Picture         =   "frmFiltro.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancelar"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   3120
      Picture         =   "frmFiltro.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Acceder"
      Top             =   1080
      Width           =   615
   End
   Begin VB.ComboBox cboHasta 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.ComboBox cboDesde 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblHasta 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lblDesde 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If tipoFiltro = "Alumnos" Then
        sqlFiltro = "Alumnos.Nombre >= '" & cboDesde.Text & "' AND Alumnos.Nombre <= '" & cboHasta.Text & "'"
    ElseIf tipoFiltro = "Cursos" Then
        sqlFiltro = "TiposCurso.Detalle >= '" & cboDesde.Text & "' AND TiposCurso.Detalle <= '" & cboHasta.Text & "'"
    ElseIf tipoFiltro = "Empresas" Then
        sqlFiltro = "Empresas.Nombre >= '" & cboDesde.Text & "' AND Empresas.Nombre <= '" & cboHasta.Text & "'"
    End If
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    sqlFiltro = ""
    Unload Me
End Sub

Private Sub Form_Load()
    If tipoFiltro = "Alumnos" Then
        CARGAR_COMBO "cboDesde", adoAlumnos, "Alumnos", "Nombre", Me
        CARGAR_COMBO "cboHasta", adoAlumnos, "Alumnos", "Nombre", Me
        
        cboHasta.ListIndex = cboHasta.ListCount - 1
    ElseIf tipoFiltro = "Cursos" Then
        CARGAR_COMBO "cboDesde", adoTiposCurso, "TiposCurso", "Detalle", Me
        CARGAR_COMBO "cboHasta", adoTiposCurso, "TiposCurso", "Detalle", Me
        
        cboHasta.ListIndex = cboHasta.ListCount - 1
    ElseIf tipoFiltro = "Empresas" Then
        CARGAR_COMBO "cboDesde", adoEmpresas, "Empresas", "Nombre", Me
        CARGAR_COMBO "cboHasta", adoEmpresas, "Empresas", "Nombre", Me
        
        cboHasta.ListIndex = cboHasta.ListCount - 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tipoFiltro = ""
End Sub
