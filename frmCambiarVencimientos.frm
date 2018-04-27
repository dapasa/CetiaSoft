VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCambiarVencimientos 
   Caption         =   "Nuevos vencimientos"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   ControlBox      =   0   'False
   Icon            =   "frmCambiarVencimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3600
      Picture         =   "frmCambiarVencimientos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancelar"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   855
      Left            =   2640
      Picture         =   "frmCambiarVencimientos.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Guardar"
      Top             =   3960
      Width           =   855
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
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   4
         Left            =   960
         TabIndex        =   3
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   5
         Left            =   960
         TabIndex        =   4
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39304
      End
      Begin MSComCtl2.DTPicker dtpCuota 
         Height          =   315
         Index           =   3
         Left            =   960
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39304
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 2"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 3"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1380
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 4"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1860
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 5"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2340
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCuota 
         AutoSize        =   -1  'True
         Caption         =   "Cuota 1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.Label lblNuevoCurso 
      Caption         =   "Nuevo curso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblCurso 
      Caption         =   "Curso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblAlumno 
      Caption         =   "Alumno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmCambiarVencimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardar_Click()
    sSql = "SELECT * FROM movimientos WHERE TipoDoc = 'MOD' AND Saldo <> 0 AND idAlumno = " & id_Alumno & " AND idCurso = " & id_Curso_Nuevo & " ORDER BY id"
    CERRAR_TABLA adoMovimientos
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic

    If Not adoMovimientos.EOF Then
        adoMovimientos.MoveFirst
        Do While Not adoMovimientos.EOF
            If adoMovimientos!cuota <> 0 Then
                adoMovimientos!fecha = dtpCuota(adoMovimientos!cuota).Value
                adoMovimientos.Update
            End If
            adoMovimientos.MoveNext
        Loop
    Else
        MsgBox "No se modifican las fechas de vencimiento." & vbCrLf & "El curso está pago por completo.", vbInformation, "CAMBIAR CURSO"
    End If
    
    adoMovimientos.Close
    
    'detalle_Curso_Nuevo = lblNuevoCurso.Caption
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim id_tipo_curso As Long
    
    lblAlumno.Caption = DEVOLVER_CAMPO(id_Alumno, adoAlumnos, "alumnos", "Nombre")
    
    id_tipo_curso = DEVOLVER_CAMPO(id_Curso, adoCursos, "cursos", "idTipoCurso")
    lblCurso.Caption = DEVOLVER_CAMPO(id_tipo_curso, adoTiposCurso, "TiposCurso", "Detalle")
        
    id_tipo_curso = DEVOLVER_CAMPO(id_Curso_Nuevo, adoCursos, "cursos", "idTipoCurso")
    lblNuevoCurso.Caption = DEVOLVER_CAMPO(id_tipo_curso, adoTiposCurso, "TiposCurso", "Detalle")
        
    sSql = "SELECT CantCuotas FROM Cursos WHERE id = " & id_Curso_Nuevo
    CERRAR_TABLA adoCursos
    adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    cant_cuotas = adoCursos!CantCuotas
    
    adoCursos.Close
    
    For k = 1 To cant_cuotas
        lblCuota(k).Visible = True
        dtpCuota(k).Visible = True
        dtpCuota(k).Value = Date
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CERRAR_TODO
End Sub
