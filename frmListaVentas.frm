VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaVentas 
   Caption         =   "LISTADO - Ventas"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Alumno"
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
      TabIndex        =   28
      Top             =   2280
      Width           =   4575
      Begin VB.ComboBox cboDesdeAlumno 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   30
         Text            =   "cboDesdeAlumno"
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cboHastaAlumno 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Emisor"
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
      TabIndex        =   25
      Top             =   1200
      Width           =   3015
      Begin VB.ComboBox cboDesdeEmisor 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "cboDesdeEmisor"
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboHastaEmisor 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "cboHastaEmisor"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Unidad de negocio"
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
      Left            =   3240
      TabIndex        =   22
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox cboDesdeUnidad 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "cboDesdeUnidad"
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboHastaUnidad 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cboHastaUnidad"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Forma de pago"
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
      Left            =   3240
      TabIndex        =   19
      Top             =   1200
      Width           =   3015
      Begin VB.ComboBox cboHastaFormaPago 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "cboHastaFormaPago"
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cboDesdeFormaPago 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   6
         Text            =   "cboDesdeFormaPago"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   465
      End
   End
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3015
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   330
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   40147
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   330
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   40147
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detallado"
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
      TabIndex        =   15
      Top             =   3360
      Width           =   3015
      Begin VB.OptionButton optDetalleSi 
         Caption         =   "Si"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optDetalleNo 
         Caption         =   "No"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agrupar por forma de pago"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Si"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   5640
      Picture         =   "frmListaVentas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cancelar"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   4920
      Picture         =   "frmListaVentas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Acceder"
      Top             =   4200
      Width           =   615
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   4440
      Top             =   4320
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
      X2              =   6240
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmListaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    sSql = "DELETE FROM zzVentas"
    adoConnection.Execute sSql
    

    sSql = "INSERT INTO zzVentas (Fecha, TipoDoc, NumDoc, Subtotal, Iva, Total, NombreCliente, NombreEmisor, UnidadNegocio, FormaPago, Anulado) " & _
           "SELECT Movimientos.Fecha, Movimientos.TipoDoc, Movimientos.NumDoc, Movimientos.Subtotal, Movimientos.Iva, Movimientos.Total, Movimientos.NombreCliente, Movimientos.NombreEmisor, Movimientos.UnidadNegocio, Movimientos.FormaPago, Movimientos.Anulado " & _
           "FROM Movimientos " & _
           "WHERE (Movimientos.Fecha >= DateValue('" & dtpFechaDesde.Value & "') AND Movimientos.Fecha <= DateValue('" & dtpFechaHasta.Value & "')) " & _
           "      AND (Movimientos.NombreEmisor >= '" & cboDesdeEmisor.Text & "' AND Movimientos.NombreEmisor <= '" & cboHastaEmisor.Text & "') " & _
           "      AND (Movimientos.UnidadNegocio >= '" & cboDesdeUnidad.Text & "' AND Movimientos.UnidadNegocio <= '" & cboHastaUnidad.Text & "') " & _
           "      AND (Movimientos.FormaPago >= '" & cboDesdeFormaPago.Text & "' AND Movimientos.FormaPago <= '" & cboHastaFormaPago.Text & "') " & _
           "      AND (Movimientos.NombreCliente >= '" & cboDesdeAlumno.Text & "' AND Movimientos.NombreCliente <= '" & cboHastaAlumno.Text & "') " & _
           "ORDER BY Movimientos.Fecha, Movimientos.TipoDoc, Movimientos.NumDoc"
    adoConnection.Execute sSql
    
    sSql = "UPDATE zzVentas SET NombreCliente = 'ANULADO' WHERE Anulado"
    adoConnection.Execute sSql
    
    sSql = "UPDATE zzVentas SET Subtotal = 0 WHERE Anulado"
    adoConnection.Execute sSql
    
    sSql = "UPDATE zzVentas SET Iva = 0 WHERE Anulado"
    adoConnection.Execute sSql
    
    sSql = "UPDATE zzVentas SET Total = 0 WHERE Anulado"
    adoConnection.Execute sSql
    
    sSql = "UPDATE zzVentas SET Total = Total * -1 WHERE TipoDoc = 'NCA' OR TipoDoc = 'NCB' OR TipoDoc = 'NCC'"
    adoConnection.Execute sSql
    
    sSql = "UPDATE zzVentas SET Subtotal = Total WHERE Left(NombreEmisor,1) = '2' OR Left(NombreEmisor,1) = '3'"
    adoConnection.Execute sSql
    
    sSql = "UPDATE zzVentas SET Iva = 0 WHERE Left(NombreEmisor,1) = '2' OR Left(NombreEmisor,1) = '3'"
    adoConnection.Execute sSql
    
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    
    'Hay un rptVentas3.rpt en stand by
    If optDetalleNo Then
        rptListado.ReportFileName = App.Path & "\reportes\rptVentas.rpt"
    Else
        rptListado.ReportFileName = App.Path & "\reportes\rptVentas2.rpt"
    End If
    
    rptListado.ReportTitle = "LISTADO DE VENTAS DESDE " & dtpFechaDesde.Value & " HASTA " & dtpFechaHasta.Value
    rptListado.Action = 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CARGAR_COMBO "cboDesdeEmisor", adoEmisores, "Emisores", "Detalle", Me
    CARGAR_COMBO "cboHastaEmisor", adoEmisores, "Emisores", "Detalle", Me
    cboHastaEmisor.ListIndex = cboHastaEmisor.ListCount - 1
    
    CARGAR_COMBO "cboDesdeUnidad", adoUnidadesNegocio, "UnidadesNegocio", "Detalle", Me
    CARGAR_COMBO "cboHastaUnidad", adoUnidadesNegocio, "UnidadesNegocio", "Detalle", Me
    cboHastaUnidad.ListIndex = cboHastaUnidad.ListCount - 1

    CARGAR_COMBO "cboDesdeFormaPago", adoFormasPago, "FormasPago", "Detalle", Me
    CARGAR_COMBO "cboHastaFormaPago", adoFormasPago, "FormasPago", "Detalle", Me
    cboHastaFormaPago.ListIndex = cboHastaFormaPago.ListCount - 1

    CARGAR_COMBO "cboDesdeAlumno", adoAlumnos, "Alumnos", "Nombre", Me
    CARGAR_COMBO "cboHastaAlumno", adoAlumnos, "Alumnos", "Nombre", Me
    cboHastaAlumno.ListIndex = cboHastaAlumno.ListCount - 1

    dtpFechaDesde.Value = Date
    dtpFechaHasta.Value = Date
End Sub
