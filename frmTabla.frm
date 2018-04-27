VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTabla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABLAS - Titulo"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmTabla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Height          =   615
      Left            =   4560
      Picture         =   "frmTabla.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3655
      Begin VB.TextBox txtDetalle 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   3465
      End
   End
   Begin MSDataGridLib.DataGrid dbgTabla 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin Crystal.CrystalReport rptInforme 
      Left            =   3960
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   800
      WindowHeight    =   465
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Modifique directamente en la tabla"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2430
   End
End
Attribute VB_Name = "frmTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tTabla As clsTabla

Private Sub cmdAgregar_Click()
    On Error GoTo ErrorHandle
    
    tTabla.Agregar
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTabla - cmdAgregar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCerrar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTabla - cmdCerrar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo ErrorHandle
    
    frmClave.Show vbModal
    If Not AccesoPermitido Then
        MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
        Exit Sub
    End If
    
    If PUEDE_BORRAR(tTabla.NombreTabla, tTabla.adoTabla.Fields(0).Value) Then
        tTabla.Eliminar
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTabla - cmdEliminar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdImprimir_Click()
    'tTabla.IMPRIMIR_TABLA
    sSql = "DELETE FROM zTabla"
    adoConnection.Execute sSql
    adoConnection.Execute tTabla.SqlListado
    
    rptInforme.ReportFileName = App.Path & "\reportes\rptTablas.rpt"
    rptInforme.ReportTitle = "LISTADO DE " & UCase(tTabla.TituloPlural)
    y = ((Screen.Height - 8800) / 2) / 11
    x = ((Screen.Width - 5115) / 2) / 11
    rptInforme.WindowTop = y
    rptInforme.WindowLeft = x
    rptInforme.Action = 1
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    Set tTabla = New clsTabla

    Select Case strTabla
        Case "Aulas"
            tTabla.adoTabla = adoAulas
            tTabla.NombreTabla = "Aulas"
            tTabla.TituloPlural = "Aulas"
            tTabla.TituloSingular = "Aula"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM Aulas"
        Case "Horarios"
            tTabla.adoTabla = adoHorarios
            tTabla.NombreTabla = "Horarios"
            tTabla.TituloPlural = "Horarios"
            tTabla.TituloSingular = "Horario"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM Horarios"
        Case "Duraciones"
            tTabla.adoTabla = adoDuraciones
            tTabla.NombreTabla = "Duraciones"
            tTabla.TituloPlural = "Duraciones"
            tTabla.TituloSingular = "Duración"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM Duraciones"
        Case "Modalidades"
            tTabla.adoTabla = adoModalidades
            tTabla.NombreTabla = "Modalidades"
            tTabla.TituloPlural = "Modalidades"
            tTabla.TituloSingular = "Modalidad"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM Modalidades"
        Case "TiposCurso"
            tTabla.adoTabla = adoTiposCurso
            tTabla.NombreTabla = "TiposCurso"
            tTabla.TituloPlural = "Tipos de curso"
            tTabla.TituloSingular = "Tipo de curso"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM TiposCurso"
        Case "CondIva"
            tTabla.adoTabla = adoCondIva
            tTabla.NombreTabla = "CondIva"
            tTabla.TituloPlural = "Condiciones de IVA"
            tTabla.TituloSingular = "Condición de IVA"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM CondIva"
        Case "TiposDoc"
            tTabla.adoTabla = adoTiposDoc
            tTabla.NombreTabla = "TiposDoc"
            tTabla.TituloPlural = "Tipos de documento"
            tTabla.TituloSingular = "Tipo de documento"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM TiposDoc"
        Case "Localidades"
            tTabla.adoTabla = adoLocalidades
            tTabla.NombreTabla = "Localidades"
            tTabla.TituloPlural = "Localidades"
            tTabla.TituloSingular = "Localidad"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM Localidades"
        Case "ComoLlego"
            tTabla.adoTabla = adoComoLlego
            tTabla.NombreTabla = "ComoLlego"
            tTabla.TituloPlural = "Formas de llegar"
            tTabla.TituloSingular = "Forma de llegar"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM ComoLlego"
        Case "Sucursales"
            tTabla.adoTabla = adoSucursales
            tTabla.NombreTabla = "Sucursales"
            tTabla.TituloPlural = "Sucursales"
            tTabla.TituloSingular = "Sucursal"
            tTabla.SqlListado = "INSERT INTO zTabla (id, Detalle) SELECT * FROM Sucursales"
    End Select
    
    tTabla.Grilla = dbgTabla
    tTabla.adoConexion = adoConnection
    tTabla.Detalle = txtDetalle
    tTabla.botonAgregar = cmdAgregar
    tTabla.botonEliminar = cmdEliminar
    tTabla.botonCerrar = cmdCerrar
    
    Me.Caption = "Tabla - " & tTabla.TituloPlural
    
    tTabla.Inicializar

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTabla - Form_Load", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    tTabla.Terminar

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTabla - Form_Unload", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub txtDetalle_Change()
    If Len(txtDetalle.Text) > 0 Then
        cmdAgregar.Enabled = True
    Else
        cmdAgregar.Enabled = False
    End If
End Sub

Private Sub txtDetalle_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If KeyAscii = 13 Then
        cmdAgregar_Click
    End If

    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTabla - txtDetalle_KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub
