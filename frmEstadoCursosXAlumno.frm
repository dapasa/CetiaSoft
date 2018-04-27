VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEstadoCursosXAlumno 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame16 
      Caption         =   "Cursos del alumno"
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
      Top             =   120
      Width           =   5535
      Begin MSDataGridLib.DataGrid dbgClases 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
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
   Begin VB.Frame Frame18 
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
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdCAMBIAR_ESTADO 
      Height          =   375
      Left            =   8880
      Picture         =   "frmEstadoCursosXAlumno.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cambiar situación"
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmEstadoCursosXAlumno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCAMBIAR_ESTADO_Click()
    adoTempCursos!Estado = cboEstado.Text
    adoTempCursos.Update
End Sub

Private Sub Form_Load()
    sSql = "SELECT Cursos.Detalle FROM Cursos, AlumnosXCurso WHERE Cursos.id = AlumnosXCurso.idCurso AND AlumnosXCurso.idAlumno = " & id_Alumno
    
End Sub
