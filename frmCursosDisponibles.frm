VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCursosDisponibles 
   Caption         =   "CURSOS DISPONIBLES"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   Icon            =   "frmCursosDisponibles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Cursos de"
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
      Width           =   8415
      Begin VB.ComboBox cboTiposCurso 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   7920
      Picture         =   "frmCursosDisponibles.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   7200
      Picture         =   "frmCursosDisponibles.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Acceder"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cursos disponibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8415
      Begin MSDataGridLib.DataGrid dbgCursosDisponibles 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5530
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
               LCID            =   3082
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
               LCID            =   3082
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
End
Attribute VB_Name = "frmCursosDisponibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboTiposCurso_Click()
    ARMAR_GRILLA cboTiposCurso.Text
End Sub

Private Sub cmdAceptar_Click()
    id_Curso_Nuevo = adoCursosDisponibles!id
    detalle_Curso_Nuevo = adoCursosDisponibles.Fields(4).Value
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    id_Curso_Nuevo = 0
    
    CARGAR_COMBO "cboTiposCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    
    ARMAR_GRILLA
End Sub

Private Sub ARMAR_GRILLA(Optional curso As String = "TODOS")
    CERRAR_TABLA adoCursosDisponibles
    
    If curso = "TODOS" Then
        sSql = "SELECT Cursos.id, Cursos.Numero, Cursos.FechaIni, Cursos.FechaFin, TiposCurso.Detalle, Horarios.Detalle FROM Cursos, Horarios, TiposCurso WHERE Cursos.idHorario = Horarios.id AND Cursos.idTipoCurso = TiposCurso.id AND Cursos.Vacantes >= 1 AND Cursos.Abierto ORDER BY TiposCurso.Detalle, Cursos.Numero DESC"
    Else
        sSql = "SELECT Cursos.id, Cursos.Numero, Cursos.FechaIni, Cursos.FechaFin, TiposCurso.Detalle, Horarios.Detalle FROM Cursos, Horarios, TiposCurso WHERE Cursos.idHorario = Horarios.id AND Cursos.idTipoCurso = TiposCurso.id AND Cursos.Vacantes >= 1 AND Cursos.Abierto AND TiposCurso.Detalle = '" & curso & "' ORDER BY Cursos.Numero DESC"
    End If
    
    adoCursosDisponibles.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
    
    Set dbgCursosDisponibles.DataSource = adoCursosDisponibles
    
    'Formato de la grilla.
    With dbgCursosDisponibles
        .Columns(0).Visible = False
        .Columns(1).Caption = "Nº"
        .Columns(2).Caption = "Inicio"
        .Columns(3).Caption = "Fin"
        .Columns(4).Caption = "Curso"
        .Columns(5).Caption = "Horario"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CERRAR_TABLA adoCursosDisponibles
End Sub
