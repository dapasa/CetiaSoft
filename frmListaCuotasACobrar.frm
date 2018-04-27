VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaCuotasACobrar 
   Caption         =   "LISTADO - Cuotas a cobrar"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4800
      TabIndex        =   17
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox cboHorarioExtra 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   18
         Text            =   "cboHorarioExtra"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cboHorario 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "cboHorario"
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraFechaHasta 
      Caption         =   "Fecha tope"
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
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   325
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   40147
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
      TabIndex        =   13
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox cboHasta 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cboHasta"
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cboDesde 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "cboDesde"
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   420
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
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
      Top             =   840
      Width           =   3135
      Begin VB.OptionButton optListado2 
         Caption         =   "Sin números de teléfono"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optListado1 
         Caption         =   "Con números de teléfono"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2100
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
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   4575
      Begin VB.OptionButton optTodo 
         Caption         =   "Todo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optAVencer 
         Caption         =   "A vencer"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton optVencido 
         Caption         =   "Vencido"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   6600
      Picture         =   "frmListaCuotasACobrar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Acceder"
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   7320
      Picture         =   "frmListaCuotasACobrar.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancelar"
      Top             =   2760
      Width           =   615
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   6120
      Top             =   2880
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
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmListaCuotasACobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    sSql = "DELETE FROM zzACobrar"
    adoConnection.Execute sSql
    
    If cboHorario.Text = "(No disponible)" Then
        cboHorario.ListIndex = 1
        cboHorarioExtra.ListIndex = cboHorarioExtra.ListCount - 1
    Else
        cboHorarioExtra.Text = cboHorario.Text
    End If
    
    sSql = "INSERT INTO zzACobrar (nombre, telefono_particular, celular, mail_celular, mail, vencimiento, importe, cuota, tipo_curso, numero_curso, horario, tipo_comprobante, descuento) " & _
           "SELECT Alumnos.Nombre, Alumnos.Telefono, Alumnos.Celular, '11'+Right(Alumnos.Celular, Len(Alumnos.Celular)-2)+CompaniasCelular.Dominio, Alumnos.Mail, Movimientos.Fecha, Movimientos.Saldo, Movimientos.Cuota, Left(TiposCurso.Detalle, 3), Cursos.Numero, Horarios.Detalle, Movimientos.TipoDoc, Movimientos.Descuento " & _
           "FROM Alumnos, CompaniasCelular, Movimientos, TiposCurso, Cursos, Horarios " & _
           "WHERE Alumnos.id = Movimientos.idAlumno AND CompaniasCelular.id = Alumnos.idCompaniaCelular AND TiposCurso.id = Cursos.idTipoCurso AND Cursos.id = Movimientos.idCurso AND Horarios.id = Cursos.idHorario " & _
           "      AND Movimientos.Fecha <= DateValue('" & dtpFechaHasta.Value & "') " & _
           "      AND Movimientos.TipoDoc <> 'ABA' " & _
           "      AND Movimientos.Saldo > 0 " & _
           "      AND (TiposCurso.Detalle >= '" & cboDesde.Text & "' AND TiposCurso.Detalle <= '" & cboHasta.Text & "') " & _
           "      AND (Horarios.Detalle >= '" & cboHorario.Text & "' AND Horarios.Detalle <= '" & cboHorarioExtra.Text & "') " & _
           "ORDER BY TiposCurso.Detalle, Cursos.Numero, Movimientos.Fecha"
           
           
    'sSql = "INSERT INTO zzACobrar (nombre, telefono_particular, celular, mail_celular, mail, vencimiento, importe, cuota, tipo_curso, numero_curso, horario, tipo_comprobante) " & _
    '       "SELECT Alumnos.Nombre, Alumnos.Telefono, Alumnos.Celular, '11'+Right(Alumnos.Celular, Len(Alumnos.Celular)-2)+CompaniasCelular.Dominio, Alumnos.Mail, Movimientos.Fecha, Movimientos.Saldo, Movimientos.Cuota, Left(TiposCurso.Detalle, 3), Cursos.Numero, Horarios.Detalle, Movimientos.TipoDoc " & _
    '       "FROM Alumnos, CompaniasCelular, Movimientos, TiposCurso, Cursos, Horarios " & _
    '       "WHERE Alumnos.id = Movimientos.idAlumno AND CompaniasCelular.id = Alumnos.idCompaniaCelular AND TiposCurso.id = Cursos.idTipoCurso AND Cursos.id = Movimientos.idCurso AND Horarios.id = Cursos.idHorario " & _
    '       "      AND Movimientos.Fecha <= DateValue('" & dtpFechaHasta.Value & "') " & _
    '       "      AND Movimientos.TipoDoc <> 'ABA' " & _
    '       "      AND Movimientos.Saldo > 0 " & _
    '       "      AND (TiposCurso.Detalle >= '" & cboDesde.Text & "' AND TiposCurso.Detalle <= '" & cboHasta.Text & "') " & _
    '       "      AND (Horarios.Detalle >= '" & cboHorario.Text & "' AND Horarios.Detalle <= '" & cboHorarioExtra.Text & "') " & _
    '       "ORDER BY TiposCurso.Detalle, Cursos.Numero, Movimientos.Fecha"
    
    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    
    If optListado1.Value = True Then
        rptListado.ReportFileName = App.Path & "\reportes\rptACobrar_new.rpt"
        'rptListado.ReportFileName = App.Path & "\reportes\rptACobrar.rpt"
    Else
        rptListado.ReportFileName = App.Path & "\reportes\rptACobrar2_new.rpt"
        'rptListado.ReportFileName = App.Path & "\reportes\rptACobrar2.rpt"
    End If
    
    rptListado.ReportTitle = "LISTADO DE CUOTAS PENDIENTES DE COBRO AL " & dtpFechaHasta.Value
    rptListado.Action = 1
    
    cboHorario.ListIndex = 0
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtpFechaHasta.Value = Date
    CARGAR_COMBO "cboDesde", adoTiposCurso, "TiposCurso", "Detalle", Me
    CARGAR_COMBO "cboHasta", adoTiposCurso, "TiposCurso", "Detalle", Me
    cboHasta.ListIndex = cboHasta.ListCount - 1

    CARGAR_COMBO "cboHorario", adoHorarios, "Horarios", "Detalle", Me
    CARGAR_COMBO "cboHorarioExtra", adoHorarios, "Horarios", "Detalle", Me
End Sub

Private Sub optAVencer_Click()
    fraFechaHasta.Visible = True
    'dtpFechaHasta.Visible = True
End Sub

Private Sub optTodo_Click()
    fraFechaHasta.Visible = True
    'dtpFechaHasta.Visible = True
End Sub

Private Sub optVencido_Click()
    dtpFechaHasta.Value = Date
    fraFechaHasta.Visible = False
    'dtpFechaHasta.Visible = False
End Sub
