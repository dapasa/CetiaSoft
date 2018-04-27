VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCursos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CURSOS"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   Icon            =   "frmCursos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOcultarClases 
      Caption         =   "Ocultar"
      Height          =   855
      Left            =   7800
      Picture         =   "frmCursos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Clases dictadas"
      Top             =   5640
      Width           =   855
   End
   Begin VB.Frame Frame19 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   10200
      TabIndex        =   62
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtObservaciones 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdVerClases 
      Caption         =   "Ver clases"
      Height          =   855
      Left            =   11400
      Picture         =   "frmCursos.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Clases dictadas"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdGuardarClases 
      Caption         =   "&Guardar"
      Height          =   855
      Left            =   5760
      Picture         =   "frmCursos.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Guardar"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdAGREGAR_CLASE 
      Height          =   375
      Left            =   4320
      Picture         =   "frmCursos.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Agregar clase"
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdBORRAR_CLASE 
      Height          =   375
      Left            =   4800
      Picture         =   "frmCursos.frx":11F2
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Borrar clase"
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdCAMBIAR_PROFESOR 
      Height          =   375
      Left            =   8880
      Picture         =   "frmCursos.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cambiar profesor"
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdCAMBIAR_SITUACION 
      Height          =   375
      Left            =   8880
      Picture         =   "frmCursos.frx":180E
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cambiar situación"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdCARGAR_CLASES 
      Caption         =   "Cargar clases"
      Height          =   495
      Left            =   9480
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdIMPRIMIR_CLASES 
      Height          =   375
      Left            =   5280
      Picture         =   "frmCursos.frx":1D40
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Imprimir clases"
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdTomarLista 
      Caption         =   "Tomar Lista"
      Height          =   855
      Left            =   6720
      Picture         =   "frmCursos.frx":1E42
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame Frame18 
      Caption         =   "Situación"
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
      TabIndex        =   59
      Top             =   4200
      Width           =   3015
      Begin VB.ComboBox cboSituacion 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "Profesor"
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
      TabIndex        =   58
      Top             =   4920
      Width           =   3015
      Begin VB.ComboBox cboProfesorClase 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraClases 
      Caption         =   "Clases del curso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   57
      Top             =   3120
      Visible         =   0   'False
      Width           =   5535
      Begin MSDataGridLib.DataGrid dbgClases 
         Height          =   3015
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5318
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
   Begin VB.Frame Frame3 
      Caption         =   "Vencimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   7680
      TabIndex        =   51
      Top             =   120
      Width           =   2415
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   4
         Left            =   960
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   5
         Left            =   960
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   3
         Left            =   960
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   39304
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   56
         Top             =   420
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 5"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   55
         Top             =   2340
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 4"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   54
         Top             =   1860
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 3"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   1380
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 2"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   900
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   7560
      Picture         =   "frmCursos.frx":2284
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Buscar"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   12360
      Picture         =   "frmCursos.frx":26C6
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancelar"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   10440
      Picture         =   "frmCursos.frx":2B08
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Guardar"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8520
      Picture         =   "frmCursos.frx":2F4A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Eliminar"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   6600
      Picture         =   "frmCursos.frx":338C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Nuevo"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdListado 
      Caption         =   "&Listado"
      Enabled         =   0   'False
      Height          =   855
      Left            =   9480
      Picture         =   "frmCursos.frx":37CE
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Lista de alumnos"
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame15 
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
      TabIndex        =   49
      Top             =   840
      Width           =   2055
      Begin VB.OptionButton optCerrado 
         Caption         =   "Cerrado"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAbierto 
         Caption         =   "Abierto"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Valor cuota"
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
      Left            =   2280
      TabIndex        =   47
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
      Begin VB.TextBox txtValorCuota 
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Left            =   480
         TabIndex        =   48
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Inscriptos"
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
      Left            =   2280
      TabIndex        =   46
      Top             =   1560
      Width           =   1095
      Begin VB.Label lblInscriptos 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame12 
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
      Left            =   1080
      TabIndex        =   45
      Top             =   1560
      Width           =   1095
      Begin VB.TextBox txtVacantes 
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Cuotas"
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
      TabIndex        =   44
      Top             =   1560
      Width           =   855
      Begin VB.TextBox txtCantCuotas 
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Inicio"
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
      TabIndex        =   43
      Top             =   840
      Width           =   1575
      Begin MSComCtl2.DTPicker dtpFechaIni 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   39304
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Finalización"
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
      Left            =   6000
      TabIndex        =   42
      Top             =   840
      Width           =   1575
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   39304
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Aula"
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
      Left            =   4560
      TabIndex        =   41
      Top             =   2280
      Width           =   3015
      Begin VB.ComboBox cboAula 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Modalidad"
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
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ComboBox cboModalidad 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1815
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
      Left            =   1560
      TabIndex        =   39
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox cboTipoCurso 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Profesor"
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
      Left            =   4560
      TabIndex        =   38
      Top             =   1560
      Width           =   3015
      Begin VB.ComboBox cboProfesor 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   2775
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
      Height          =   615
      Left            =   2280
      TabIndex        =   37
      Top             =   840
      Width           =   1935
      Begin VB.ComboBox cboHorario 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Duración"
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
      Left            =   5280
      TabIndex        =   36
      Top             =   120
      Width           =   2295
      Begin VB.ComboBox cboDuracion 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Número"
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
      TabIndex        =   34
      Top             =   120
      Width           =   1335
      Begin VB.Label lblNumero 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport rptInforme 
      Left            =   120
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   600
      WindowHeight    =   350
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   13200
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmCursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' MÓDULO: Mantenimiento de cursos     FECHA: Ago / 2007
'******************************************************
' RESUMEN:
'******************************************************
' ÚLTIMA MODIFICACIÓN IMPORTANTE: 11/08/2007
'******************************************************
' ETAPA: release candidate.
'******************************************************
' AUTOR: Pablo Adrián Langholz
' CONTACTO: elmaildepablo@gmail.com
'******************************************************

Dim ModoABM As String
Dim EventoLoad As Boolean
Dim Fila As Byte
Dim fechasRenglon As Byte
Dim inicializando_pantalla As Boolean

Private Sub cmdARMAR_GRILLA_CLASES_Click()
    sSql = "DELETE FROM TempClases"
    adoConnection.Execute sSql
    
    With adoTempClases
        CERRAR_TABLA adoTempClases
        .Open "TempClases", adoConnection, adOpenKeyset, adLockOptimistic

        seguir = True
        fecha_clase = dtpFechaIni.Value
        num_clase = 1
        Do While seguir
            .AddNew
            !Numero = num_clase
            !fecha = fecha_clase
            !Situacion = ""
            !Profesor = cboProfesor.Text
            .Update
            
            fecha_clase = dtpFechaIni.Value + (7 * num_clase)
            num_clase = num_clase + 1
            
            If fecha_clase > dtpFechaFin.Value Then
                seguir = False
            End If
        Loop
    End With
    
    adoTempClases.MoveFirst
    
    Set dbgClases.DataSource = adoTempClases

    FORMATO_GRILLA_CLASES
End Sub

Private Sub cmdCAMBIAR_PROFESOR_Click()
    adoTempClases!Profesor = cboProfesorClase.Text
    adoTempClases.Update
End Sub

Private Sub cmdCAMBIAR_SITUACION_Click()
    adoTempClases!Situacion = cboSituacion.Text
    adoTempClases.Update
End Sub

Private Sub cboTipoCurso_Click()
    If Not EventoLoad Then
        lblNumero.Caption = NUMERO_CURSO
    End If
End Sub

Private Sub cmdAGREGAR_CLASE_Click()
    With adoTempClases
        If Not adoTempClases.EOF Then
            .MoveLast
            
            ult_numero = !Numero
            ult_fecha = !fecha
        Else
            ult_numero = 0
            ult_fecha = dtpFechaIni.Value - 7
        End If
        
        .AddNew
        !Numero = ult_numero + 1
        !fecha = ult_fecha + 7
        !Situacion = ""
        !Profesor = cboProfesor.Text
        .Update
        
        .MoveFirst
        .MoveLast
    End With
End Sub

Private Sub cmdBORRAR_CLASE_Click()
    If MsgBox("¿Confirma que desea eliminar la clase Nº " & adoTempClases!Numero & "?", vbQuestion + vbYesNo, "ELIMINAR CLASE") = vbYes Then
        adoTempClases.Delete
        RENUMERAR_CLASES
    End If
End Sub

Private Sub cmdBuscar_Click()
'    On Error GoTo ErrorHandle
    
    BOTONES "Buscar", Me

    ENABLED_TODO True, Me
    
    EstiloBuscador = "Cursos"
    
    EventoLoad = True
    frmBuscador.Show vbModal
    
    'cmdARMAR_GRILLA_CLASES_Click
    CARGAR_GRILLA_CLASES
    
    EventoLoad = False
    
    cboProfesorClase.Text = cboProfesor.Text
    
    InicioAnterior = dtpFechaIni.Value
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - cmdBuscar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - cmdCancelar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCARGAR_CLASES_Click()
    CERRAR_TABLA adoClasesXCurso
    sSql = "SELECT * FROM ClasesXCurso WHERE idCurso = " & id_Curso
    adoClasesXCurso.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If adoClasesXCurso.EOF Then
        adoClasesXCurso.Close
        Exit Sub
    End If
    
    sSql = "DELETE FROM TempClases"
    adoConnection.Execute sSql
    
    With adoClasesXCurso
        .MoveFirst
        
        CERRAR_TABLA adoTempClases
        adoTempClases.Open "TempClases", adoConnection, adOpenKeyset, adLockOptimistic
        Do While Not .EOF
            adoTempClases.AddNew
            
            adoTempClases!Numero = !Numero
            adoTempClases!fecha = !fecha
            adoTempClases!Situacion = !Situacion
            adoTempClases!Profesor = !Profesor
                        
            adoTempClases.Update
        
            .MoveNext
        Loop
    End With
    
    adoClasesXCurso.Close
    
    adoTempClases.MoveFirst
    
    Set dbgClases.DataSource = adoTempClases

    FORMATO_GRILLA_CLASES
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo ErrorHandle
    BOTONES "Eliminar", Me
    
    If MsgBox("¿Confirma que desea eliminar el curso Nº " & lblNumero.Caption & " de " & cboTipoCurso.Text & "?", vbYesNo + vbQuestion, "Cursos") = vbYes Then
        frmClave.Show vbModal
        If Not AccesoPermitido Then
            MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
            Exit Sub
        End If

        If PUEDE_BORRAR("Cursos", id_Curso) Then
            ModoABM = "B"
            
            If NO_HAY_ALUMNOS(id_Curso) Then
                'Eliminar
                ELIMINAR_CURSO
            Else
                MsgBox "No es posible eliminar el curso." & vbCrLf & "Hay alumnos inscriptos.", vbCritical, "ERROR"
            End If
        End If
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - cmdEliminar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo ErrorHandle
    BOTONES "Guardar", Me
    
    If MsgBox("¿Confirma que desea guardar los cambios?", vbYesNo + vbQuestion, "CURSOS") = vbYes Then
        'Guardar
        If id_Curso = 0 Then 'Curso nuevo
            If VALIDAR_AULA Then
                ModoABM = "A"
                NUEVO_CURSO
                
                cmdARMAR_GRILLA_CLASES_Click
                
                ModoABM = ""
            Else
                MsgBox "El curso ya existe o el aula está ocupada.", vbCritical, "ERROR"
                BOTONES "Nuevo", Me
            End If
        Else 'Modificación de curso
            sMenu = "Cursos"
            
            AccesoPublico = True
                frmClave.Show vbModal
            AccesoPublico = False

            If Not AccesoPermitido Then
                MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
                Exit Sub
            End If
            
            GUARDAR_LOG x_usuario, Date, Time, "MODIFICA CURSO " & lblNumero.Caption & " " & cboTipoCurso.Text
            
            With adoCursos
                ModoABM = "M"
                MODIFICAR_CURSO
                ACTUALIZAR_VENCIMIENTOS
            End With
        End If
            
        'GUARDAR_CLASES
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - cmdGuardar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdGuardarClases_Click()
    sSql = "DELETE FROM ClasesXCurso WHERE idCurso = " & id_Curso
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoClasesXCurso
    
    With adoClasesXCurso
        .Open "ClasesXCurso", adoConnection, adOpenDynamic, adLockOptimistic
    
        adoTempClases.MoveFirst
        Do While Not adoTempClases.EOF
            .AddNew
            
            !idCurso = id_Curso
            !Numero = adoTempClases!Numero
            !fecha = adoTempClases!fecha
            !Situacion = adoTempClases!Situacion
            !Profesor = adoTempClases!Profesor
            !Pagada = "NO"
            !FechaPago = ""
            
            .Update
            
            adoTempClases.MoveNext
        Loop
    
        .Close
    End With
    
    adoTempClases.MoveFirst
End Sub

Private Sub cmdIMPRIMIR_CLASES_Click()
    Dim Fila As Byte
    
    Printer.ScaleMode = 4
    
    With adoTempClases
        .MoveFirst
        
        IMPRIMIR 1, 5, "CLASES DEL CURSO - " & cboTipoCurso.Text
        IMPRIMIR 2, 5, cboHorario.Text
        IMPRIMIR 3, 5, "Inicio: " & dtpFechaIni.Value & "   Fin: " & dtpFechaFin.Value
        
        Fila = 5
        Do While Not .EOF
            IMPRIMIR Fila, 5, !Numero
            IMPRIMIR Fila, 10, !fecha
            IMPRIMIR Fila, 23, !Situacion
            IMPRIMIR Fila, 37, !Profesor
            
            Fila = Fila + 1
            .MoveNext
        Loop
    End With
    
    Printer.EndDoc
End Sub

Private Sub cmdListado_Click()
    rptInforme.ReportFileName = App.Path & "\reportes\rptCursos.rpt"
    rptInforme.ReportTitle = "LISTADO DE CURSOS"
    rptInforme.Action = 1
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo ErrorHandle
    
    BOTONES "Nuevo", Me
    
    id_Curso = 0
    
    ENABLED_TODO True, Me
    
    cboTipoCurso.SetFocus
    
    INICIALIZAR_PANTALLA
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - cmdNuevo", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdOcultarClases_Click()
    fraClases.Visible = False
    Me.Height = 4560
End Sub

Private Sub cmdTomarLista_Click()
    Dim Columna As Byte
    Dim contFechas As Byte
    Dim imprimirMasAlumnos As Boolean
    
    imprimirMasAlumnos = False
    fechasRenglon = 13
    
    Printer.ScaleMode = 4
    Printer.Orientation = vbPRORLandscape
    
    With adoTempClases
        .MoveFirst
        
        IMPRIMIR 1, 2, "CURSO    " & cboTipoCurso.Text & " - " & lblNumero.Caption
        IMPRIMIR 2, 2, "PROFESOR " & cboProfesor.Text
        IMPRIMIR 3, 2, "HORARIO  " & cboHorario.Text
        
        Printer.FontSize = 7
        Fila = 5
        Columna = 15
        contFechas = 1
        Do While Not .EOF
            IMPRIMIR Fila, Columna, Left(!fecha, 10)
            
            Columna = Columna + 8
            
            contFechas = contFechas + 1
            If contFechas > fechasRenglon Then
                contFechas = 1
                IMPRIMIR_ALUMNOS
                            
                imprimirMasAlumnos = True
                
                Fila = Fila + 3
                Columna = 15
            End If
            
            .MoveNext
        Loop
        
        If imprimirMasAlumnos = True Then
            IMPRIMIR_ALUMNOS
        End If
    End With
    
    Printer.EndDoc
End Sub



Private Sub cmdVerClases_Click()
    Me.Height = 7545
    fraClases.Visible = True

    CARGAR_GRILLA_CLASES
End Sub

Private Sub dbgClases_DblClick()
    If adoTempClases!Situacion = "" Then
        adoTempClases!Situacion = "OK"
    ElseIf adoTempClases!Situacion = "OK" Then
        adoTempClases!Situacion = "NO CETIA"
    ElseIf adoTempClases!Situacion = "NO CETIA" Then
        adoTempClases!Situacion = "NO PROFESOR"
    ElseIf adoTempClases!Situacion = "NO PROFESOR" Then
        adoTempClases!Situacion = "NO FERIADO"
    ElseIf adoTempClases!Situacion = "NO FERIADO" Then
        adoTempClases!Situacion = "OK"
    End If
End Sub



Private Sub dtpFechaIni_Change()
    cant_clases = Val(Left(cboDuracion.Text, 1)) * 4
    
    fecha_fin = dtpFechaIni.Value
    
    For i = 2 To cant_clases
        fecha_fin = fecha_fin + 7
    Next
    
    dtpFechaFin.Value = fecha_fin
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    error_cant_cuotas = True
    
    If KeyAscii = 13 Then
        PASAR_CAMPO Me
    End If

    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - Form.KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Me.Height = 4560
    
    id_Curso = 0
    
    ENABLED_TODO False, Me
    
    CERRAR_TABLA adoCursos
    sSql = "SELECT * FROM Cursos"
    adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If adoCursos.EOF Then
        cmdBuscar.Enabled = False
    End If
    
    EventoLoad = True
    CARGAR_COMBO "cboTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    CARGAR_COMBO "cboDuracion", adoDuraciones, "Duraciones", "Detalle", Me
    CARGAR_COMBO "cboHorario", adoHorarios, "Horarios", "Detalle", Me
    'CARGAR_COMBO "cboModalidad", adoModalidades, "Modalidades", "Detalle", Me
    CARGAR_COMBO "cboProfesor", adoProfesores, "Profesores", "Nombre", Me
    CARGAR_COMBO "cboAula", adoAulas, "Aulas", "Detalle", Me
    CARGAR_COMBO "cboProfesorClase", adoProfesores, "Profesores", "Nombre", Me
    'cboProfesorClase.RemoveItem (0) 'Elimino la opción (No disponible)
    
    With cboSituacion
        .AddItem "OK"
        .AddItem "NO CETIA"
        .AddItem "NO PROFESOR"
        .AddItem "NO FERIADO"
        .ListIndex = 0
    End With
    
    EventoLoad = False
    
    dtpFechaIni.Value = Date
    dtpFechaFin.Value = Date
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - Form.Load", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    CERRAR_TABLA adoCursos

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - Form.Unload", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub NUEVO_CURSO()
    On Error GoTo ErrorHandle
    
    With adoCursos
        'Verifico que los datos sean correctos
        If VALIDAR_CURSO = True Then
            'Agrego el registro
            .AddNew
            
            ASIGNAR_DATOS
            'Guardo los cambios
            .Update
            
            .MoveLast
            id_Curso = !id
            
            'Actualizo la tabla de los últimos números de cada curso
            ACTUALIZAR_NUMERO_CURSO
        Else
            MsgBox MensajeValidacion, vbCritical, "ERROR"
            MensajeValidacion = ""
        End If
    End With

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - NUEVO_CURSO", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub MODIFICAR_CURSO()
    On Error GoTo ErrorHandle
    
    With adoCursos
        'Verifico que los datos sean correctos
        If VALIDAR_CURSO = True Then
            CERRAR_TABLA adoCursos
            sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso
            adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic

            ASIGNAR_DATOS
            'Guardo los cambios
            .Update
        Else
            MsgBox MensajeValidacion, vbCritical, "ERROR"
            MensajeValidacion = ""
        End If
    End With

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - MODIFICAR_CURSO", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub ELIMINAR_CURSO()
    On Error GoTo ErrorHandle
    
    With adoCursos
        CERRAR_TABLA adoCursos
        sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso
        adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
         
        .Delete
        
        adoCursos.Close
        
        INICIALIZAR_PANTALLA
    End With

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - ELIMINAR_CURSO", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Function VALIDAR_CURSO() As Boolean
    On Error GoTo ErrorHandle
    
    Dim HuboError As Boolean
    HuboError = False
    
    MensajeValidacion = ""
    
    With adoTablaValidacion
        'Por el momento no hay validación que hacer
    End With
    
    If HuboError Then
        VALIDAR_CURSO = False
    Else
        VALIDAR_CURSO = True
    End If

    Exit Function
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - VALIDAR_Curso", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Function

Private Sub ASIGNAR_DATOS()
    On Error GoTo ErrorHandle
    
    'Asigno datos
    With adoCursos
        !Numero = lblNumero.Caption
        !idTipoCurso = DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "Detalle")
        !idDuracion = DEVOLVER_ID(cboDuracion.Text, adoDuraciones, "Duraciones", "Detalle")
        !Abierto = optAbierto.Value 'Si no está seleccionado Abierto, está seleccionado Cerrado, entonce guarda FALSE
        !idHorario = DEVOLVER_ID(cboHorario.Text, adoHorarios, "Horarios", "Detalle")
        !FechaIni = dtpFechaIni.Value
        !FechaFin = dtpFechaFin.Value
        '!ValorCuota = Val(txtValorCuota.Text)
        !CantCuotas = Val(txtCantCuotas.Text)
        !Vacantes = Val(txtVacantes.Text)
        '!idModalidad = DEVOLVER_ID(cboModalidad, adoModalidades, "Modalidades", "Detalle")
        !idProfesor = DEVOLVER_ID(cboProfesor, adoProfesores, "Profesores", "Nombre")
        !idAula = DEVOLVER_ID(cboAula, adoAulas, "Aulas", "Detalle")
        
        If dtpCuota(1).Visible = True Then
            !Cuota1 = dtpCuota(1).Value
        End If
        If dtpCuota(2).Visible = True Then
            !Cuota2 = dtpCuota(2).Value
        End If
        If dtpCuota(3).Visible = True Then
            !Cuota3 = dtpCuota(3).Value
        End If
        If dtpCuota(4).Visible = True Then
            !Cuota4 = dtpCuota(4).Value
        End If
        If dtpCuota(5).Visible = True Then
            !Cuota5 = dtpCuota(5).Value
        End If
        
        !Observaciones = txtObservaciones.Text & ""
        
        'INI - Actualizo el campo USADO de las tablas relacionadas.
        sSql = "UPDATE Aulas SET usado = true WHERE id = " & DEVOLVER_ID(cboAula, adoAulas, "Aulas", "Detalle")
        adoConnection.Execute sSql
        
        sSql = "UPDATE Duraciones SET usado = true WHERE id = " & DEVOLVER_ID(cboDuracion.Text, adoDuraciones, "Duraciones", "Detalle")
        adoConnection.Execute sSql
        
        sSql = "UPDATE Horarios SET usado = true WHERE id = " & DEVOLVER_ID(cboHorario.Text, adoHorarios, "Horarios", "Detalle")
        adoConnection.Execute sSql
        
        sSql = "UPDATE Profesores SET usado = true WHERE id = " & DEVOLVER_ID(cboProfesor, adoProfesores, "Profesores", "Nombre")
        adoConnection.Execute sSql
        
        sSql = "UPDATE TiposCurso SET usado = true WHERE id = " & DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "Detalle")
        adoConnection.Execute sSql
        'FIN - Actualizo el campo USADO de las tablas relacionadas.
    End With
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - ASIGNAR_DATOS", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub INICIALIZAR_PANTALLA()
    On Error GoTo ErrorHandle
    
    inicializando_pantalla = True
    
    For Each Control In Me.Controls
        Select Case Left(Control.Name, 3)
            Case "txt"
                Control.Text = ""
            Case "cbo"
                If Control.ListCount > 0 Then
                    Control.ListIndex = 0
                Else
                    Control.ListIndex = -1
                End If
            Case "dtp"
                Control.Value = Date
        End Select
    Next
    
    lblInscriptos.Caption = "0"
    'cboModalidad.Text = "PRESENCIAL"
    inicializando_pantalla = False
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmCursos - INICIALIZAR_PANTALLA", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Function NUMERO_CURSO() As String
    sSql = "SELECT NumeracionCursos.UltimoNumero FROM NumeracionCursos, TiposCurso " & _
           "WHERE (TiposCurso.Detalle = '" & cboTipoCurso.Text & "') " & _
           "AND (TiposCurso.id = NumeracionCursos.idTipoCurso)"
    adoNumeracionCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoNumeracionCursos.EOF Then
        NUMERO_CURSO = adoNumeracionCursos!UltimoNumero + 1
    Else
        NUMERO_CURSO = "1"
    End If
    
    adoNumeracionCursos.Close
End Function

Private Sub ACTUALIZAR_NUMERO_CURSO()
    id_TipoCurso = DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "Detalle")
    
    sSql = "SELECT NumeracionCursos.UltimoNumero FROM NumeracionCursos " & _
           "WHERE (NumeracionCursos.idTipoCurso = " & id_TipoCurso & ")"
    adoNumeracionCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoNumeracionCursos.EOF Then
        adoNumeracionCursos!UltimoNumero = adoNumeracionCursos!UltimoNumero + 1
    Else
        adoNumeracionCursos.AddNew
        adoNumeracionCursos!UltimoNumero = 1
    End If
    
    adoNumeracionCursos.Update
    
    adoNumeracionCursos.Close
End Sub

Private Function VALIDAR_AULA() As Boolean
    v_TipoCurso = DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "Detalle")
    v_Horario = DEVOLVER_ID(cboHorario.Text, adoHorarios, "Horarios", "Detalle")
    v_FechaIni = Format(dtpFechaIni.Value, "yyyy/mm/dd")
    v_FechaFin = Format(dtpFechaFin.Value, "yyyy/mm/dd")
    v_Aula = DEVOLVER_ID(cboAula, adoAulas, "Aulas", "Detalle")
    
    v_fecha1 = "(#" & v_FechaIni & "# < FechaIni AND #" & v_FechaFin & "# > FechaIni)"
    v_fecha2 = "(#" & v_FechaIni & "# > FechaIni AND #" & v_FechaIni & "# < FechaFin)"
    v_fecha3 = "(#" & v_FechaIni & "# = FechaIni OR #" & v_FechaFin & "# = FechaFin)"
    
    validacion_fecha = "(" & v_fecha1 & " OR " & v_fecha2 & " OR " & v_fecha3 & ")"
    
    sSql = "SELECT * FROM Cursos WHERE idHorario = " & v_Horario & " AND " & validacion_fecha & " AND idAula = " & v_Aula
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If Not adoTabla.EOF Then
        VALIDAR_AULA = False
        adoTabla.Close
        Exit Function
    End If
    adoTabla.Close
    
    sSql = "SELECT * FROM Cursos WHERE idHorario = " & v_Horario & " AND FechaIni <= #" & v_FechaIni & "# AND FechaFin >= #" & v_FechaIni & "# AND idAula = " & v_Aula
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If Not adoTabla.EOF Then
        VALIDAR_AULA = False
        adoTabla.Close
        Exit Function
    End If
    adoTabla.Close
    
    VALIDAR_AULA = True
End Function

Private Sub ACTUALIZAR_VENCIMIENTOS()
    sSql = "UPDATE Cursos SET Cuota1 = " & dtpCuota(1).Value & " WHERE id = " & id_Curso & " AND Cuota1 = " & dtpCuota(1).Tag
    adoConnection.Execute sSql

    sSql = "UPDATE Cursos SET Cuota2 = " & dtpCuota(2).Value & " WHERE id = " & id_Curso & " AND Cuota2 = " & dtpCuota(2).Tag
    adoConnection.Execute sSql

    sSql = "UPDATE Cursos SET Cuota3 = " & dtpCuota(3).Value & " WHERE id = " & id_Curso & " AND Cuota3 = " & dtpCuota(3).Tag
    adoConnection.Execute sSql

    If dtpCuota(4).Visible = True Then
        sSql = "UPDATE Cursos SET Cuota4 = " & dtpCuota(4).Value & " WHERE id = " & id_Curso & " AND Cuota4 = " & dtpCuota(4).Tag
        adoConnection.Execute sSql
    End If
    
    If dtpCuota(5).Visible = True Then
        sSql = "UPDATE Cursos SET Cuota5 = " & dtpCuota(5).Value & " WHERE id = " & id_Curso & " AND Cuota5 = " & dtpCuota(5).Tag
        adoConnection.Execute sSql
    End If

    sSql = "UPDATE Movimientos SET Fecha = #" & dtpCuota(1).Value & "# WHERE idCurso = " & id_Curso & " AND Cuota = 1"
    adoConnection.Execute sSql

    sSql = "UPDATE Movimientos SET Fecha = #" & dtpCuota(2).Value & "# WHERE idCurso = " & id_Curso & " AND Cuota = 2"
    adoConnection.Execute sSql

    sSql = "UPDATE Movimientos SET Fecha = #" & dtpCuota(3).Value & "# WHERE idCurso = " & id_Curso & " AND Cuota = 3"
    adoConnection.Execute sSql

    If dtpCuota(4).Visible = True Then
        sSql = "UPDATE Movimientos SET Fecha = #" & dtpCuota(4).Value & "# WHERE idCurso = " & id_Curso & " AND Cuota = 4"
        adoConnection.Execute sSql
    End If
    
    If dtpCuota(5).Visible = True Then
        sSql = "UPDATE Movimientos SET Fecha = #" & dtpCuota(5).Value & "# WHERE idCurso = " & id_Curso & " AND Cuota = 5"
        adoConnection.Execute sSql
    End If
End Sub

Private Sub txtCantCuotas_Change()
    If Not inicializando_pantalla Then
        If Val(txtCantCuotas.Text) >= 1 And Val(txtCantCuotas.Text) <= 5 Then
            ARMAR_VENCIMIENTOS Val(txtCantCuotas.Text)
            
            cant_clases = Val(Left(cboDuracion.Text, 1)) * 4

            fecha_fin = dtpFechaIni.Value
            
            For i = 2 To cant_clases
                fecha_fin = fecha_fin + 7
            Next
            
            dtpFechaFin.Value = fecha_fin

        Else
            MsgBox "La cantidad de cuotas debe ser un número entre 1 y 5.", vbCritical, "ERROR - Cursos"
        End If
    End If
End Sub

Private Sub ARMAR_VENCIMIENTOS(x As Byte)
    For k = 1 To x
        lblCuota.Item(k).Visible = True
        dtpCuota.Item(k).Visible = True
        dtpCuota.Item(k).Value = dtpFechaIni.Value + (28 * (k - 1))
    Next
    If x + 1 <= 5 Then
        For k = x + 1 To 5
            lblCuota.Item(k).Visible = False
            dtpCuota.Item(k).Visible = False
        Next
    End If
End Sub

Private Sub RENUMERAR_CLASES()
    With adoTempClases
        .MoveFirst
        
        n = 1
        Do While Not .EOF
            !Numero = n
            .Update
            n = n + 1
            .MoveNext
        Loop
    End With
End Sub

Private Sub FORMATO_GRILLA_CLASES()
    dbgClases.Columns(0).Caption = "Nº"
    dbgClases.Columns(0).Width = 400
    dbgClases.Columns(1).Width = 1000
    dbgClases.Columns(2).Caption = "Situación"
    dbgClases.Columns(2).Width = 1300
    dbgClases.Columns(3).Width = 2000
End Sub

Private Sub GUARDAR_CLASES()
    sSql = "DELETE FROM ClasesXCurso WHERE idCurso = " & id_Curso
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoClasesXCurso
    adoClasesXCurso.Open "ClasesXCurso", adoConnection, adOpenDynamic, adLockOptimistic
    
    With adoTempClases
        If .State = adStateOpen Then
            .MoveFirst
            
            Do While Not .EOF
                adoClasesXCurso.AddNew
                
                adoClasesXCurso!idCurso = id_Curso
                adoClasesXCurso!Numero = !Numero
                adoClasesXCurso!fecha = !fecha
                adoClasesXCurso!Situacion = !Situacion & ""
                adoClasesXCurso!Profesor = !Profesor
                            
                adoClasesXCurso.Update
                
                .MoveNext
            Loop
        End If
    End With
    
    adoClasesXCurso.Close
End Sub

Private Sub IMPRIMIR_ALUMNOS()
    Dim k As Byte
    
    sSql = "SELECT Alumnos.Nombre FROM Alumnos, AlumnosXCurso WHERE AlumnosXCurso.idCurso = " & id_Curso & " AND Alumnos.id = AlumnosXCurso.idAlumno"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    Fila = Fila + 1
    Do While Not adoTabla.EOF
        'IMPRIMIR Fila, 1, "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Line (1, Fila - 0.5)-(120, Fila - 0.5)
        'Fila = Fila + 1
        IMPRIMIR Fila, 1, Left(adoTabla!Nombre, 30)
                
        For k = 22 To fechasRenglon * 10 Step 8
            IMPRIMIR Fila, k, "|"
        Next
        
        Fila = Fila + 1
        
        adoTabla.MoveNext
    Loop
    
    'IMPRIMIR Fila, 1, "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Line (1, Fila - 0.5)-(120, Fila - 0.5)
    adoTabla.Close
End Sub

Private Function NO_HAY_ALUMNOS(curso As Long) As Boolean
    CERRAR_TABLA adoTemp
    
    sSql = "SELECT * FROM AlumnosXCurso WHERE idCurso = " & curso
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If Not adoTemp.EOF Then
        NO_HAY_ALUMNOS = False
    Else
        NO_HAY_ALUMNOS = True
    End If
    adoTemp.Close
End Function

Private Sub CARGAR_GRILLA_CLASES()
    CERRAR_TABLA adoTempClases

    sSql = "DELETE FROM TempClases"
    adoConnection.Execute sSql
    
    With adoTempClases
        .Open "TempClases", adoConnection, adOpenDynamic, adLockOptimistic
        
        CERRAR_TABLA adoClasesXCurso
        sSql = "SELECT * FROM ClasesXCurso WHERE idCurso = " & id_Curso
        adoClasesXCurso.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic

        If Not adoClasesXCurso.EOF Then
            adoClasesXCurso.MoveFirst
            Do While Not adoClasesXCurso.EOF
                .AddNew
                !Numero = adoClasesXCurso!Numero
                !fecha = adoClasesXCurso!fecha
                !Situacion = adoClasesXCurso!Situacion
                !Profesor = adoClasesXCurso!Profesor
                .Update
                        
                adoClasesXCurso.MoveNext
            Loop
        Else
            For i = 1 To 16
                .AddNew
                !Numero = i
                !fecha = dtpFechaIni.Value + (i * 7)
                !Situacion = ""
                !Profesor = cboProfesor.Text
                .Update
            Next
        End If
    End With
    
    adoClasesXCurso.Close
    
    adoTempClases.Close
    
    adoTempClases.Open "TempClases", adoConnection, adOpenKeyset, adLockOptimistic
    
    If Not adoTempClases.EOF Then
        adoTempClases.MoveFirst
    End If
    
    Set dbgClases.DataSource = adoTempClases

    FORMATO_GRILLA_CLASES
End Sub

