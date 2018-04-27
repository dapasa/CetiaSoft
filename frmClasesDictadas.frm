VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmClasesDictadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLASES DICTADAS"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   Icon            =   "frmClasesDictadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Calcular"
      Height          =   855
      Left            =   9240
      Picture         =   "frmClasesDictadas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Guardar"
      Top             =   5160
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Caption         =   "Total a pagar"
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
      Left            =   8520
      TabIndex        =   19
      Top             =   4440
      Width           =   1575
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00.-"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   990
         TabIndex        =   21
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Valor clase"
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
      Left            =   6840
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
      Begin VB.TextBox txtValorClase 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Docente"
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
      Left            =   6840
      TabIndex        =   14
      Top             =   3720
      Width           =   3255
      Begin VB.ComboBox cboDocente 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hasta"
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
      Left            =   8520
      TabIndex        =   12
      Top             =   3000
      Width           =   1575
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60948481
         CurrentDate     =   39371
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Desde"
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
      Left            =   6840
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60948481
         CurrentDate     =   39371
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   615
      Left            =   4680
      Picture         =   "frmClasesDictadas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar"
      Top             =   3840
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dbgClasesDictadas 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
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
   Begin VB.CommandButton cmdGuardar 
      Height          =   615
      Left            =   4680
      Picture         =   "frmClasesDictadas.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Guardar"
      Top             =   3000
      Width           =   615
   End
   Begin VB.Frame Frame2 
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
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60948481
         CurrentDate     =   39371
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Docente"
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
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
      Begin VB.ComboBox cboProfesor 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione un curso"
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
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.ListBox lstCursos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         ItemData        =   "frmClasesDictadas.frx":1108
         Left            =   120
         List            =   "frmClasesDictadas.frx":110F
         TabIndex        =   1
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NºInt.|Detalle                       |Horario                       |Inicio    |Fin"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8715
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6120
      X2              =   6120
      Y1              =   3000
      Y2              =   7200
   End
End
Attribute VB_Name = "frmClasesDictadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id_curso_ As String

Private Sub cmdCalcular_Click()
    x_profe = DEVOLVER_ID(cboDocente.Text, adoProfesores, "Profesores", "Nombre")
    filtroFecha = "(Fecha BETWEEN #" & Format(dtpDesde, "yyyy-mm-dd") & "# AND #" & Format(dtpHasta, "yyyy-mm-dd") & "#)"
    
    CERRAR_TABLA adoTemp
    sSql = "SELECT COUNT(id) AS CantClases FROM ClasesXProfe WHERE " & _
           "idCurso = " & id_Curso & " AND " & _
           "idProfesor = " & x_profe & " AND " & filtroFecha
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    lblTotal.Caption = Val(txtValorClase.Text) * adoTemp!CantClases
    adoTemp.Close
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("¿Confirma la eliminación?", vbYesNo, "Eliminar - Clase dictada") = vbYes Then
        sSql = "DELETE FROM ClasesXProfe WHERE id = " & adoTabla!id
        adoConnection.Execute sSql
    End If
    
    ARMAR_TABLA
    ARMAR_GRILLA
End Sub

Private Sub cmdGuardar_Click()
    With adoClasesXProfe
        CERRAR_TABLA adoClasesXProfe
        .Open "ClasesXProfe", adoConnection, adOpenDynamic, adLockOptimistic
        .AddNew
        
        !idCurso = Val(Left(lstCursos.List(lstCursos.ListIndex), 5))
        !idProfesor = DEVOLVER_ID(cboProfesor.Text, adoProfesores, "Profesores", "Nombre")
        !Fecha = dtpFecha.Value
        
        .Update
        .Close
    End With
    
    ARMAR_TABLA
    ARMAR_GRILLA
End Sub

Private Sub Form_Load()
    CARGAR_COMBO "cboProfesor", adoProfesores, "Profesores", "Nombre", Me
    CARGAR_COMBO "cboDocente", adoProfesores, "Profesores", "Nombre", Me
    
    CARGAR_CURSOS

    dtpFecha.Value = Date
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub

Private Sub CARGAR_CURSOS()
    lstCursos.Clear
    
    With adoCursos
        CERRAR_TABLA adoCursos
        sSql = "SELECT Cursos.id, Cursos.FechaIni, Cursos.FechaFin, Horarios.Detalle AS Horario, TiposCurso.Detalle, Cursos.idProfesor " & _
               "FROM Cursos, Horarios, TiposCurso " & _
               "WHERE Cursos.idTipoCurso = TiposCurso.id " & _
               "AND Cursos.idHorario = Horarios.id " & _
               "AND Cursos.Abierto " & _
               "ORDER BY TiposCurso.Detalle, Cursos.id"
        .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
        Do While Not .EOF
            id_Curso = Left(!id & Space(5), 5)
            Detalle_ = Left(!Detalle & Space(30), 30)
            Horario_ = Left(!Horario & Space(30), 30)
            FechaIni_ = !FechaIni
            FechaFin_ = !FechaFin
                        
            lstCursos.AddItem id_Curso & " " & Detalle_ & " " & Horario_ & " " & FechaIni_ & " " & FechaFin_
            
            .MoveNext
        Loop
        
        .Close
        
        lstCursos.ListIndex = 0
    End With
End Sub


Private Sub lstCursos_Click()
    id_Curso = Val(Left(lstCursos.List(lstCursos.ListIndex), 5))
    sSql = "SELECT Nombre FROM Profesores, Cursos WHERE Cursos.id = " & id_Curso & " AND Profesores.id = Cursos.idProfesor"
    adoProfesores.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If Not adoProfesores.EOF Then
        cboProfesor.Text = adoProfesores!Nombre
        cboDocente.Text = adoProfesores!Nombre
    Else
        cboProfesor.Text = "(No disponible)"
        cboDocente.Text = "(No disponible)"
    End If
    adoProfesores.Close
    
    'Muestro las clases dictadas en el curso seleccionado.
    ARMAR_TABLA
    ARMAR_GRILLA
End Sub

Private Sub ARMAR_TABLA()
    With adoTabla
        CERRAR_TABLA adoTabla
        sSql = "SELECT ClasesXProfe.id, ClasesXProfe.Fecha, Profesores.Nombre FROM ClasesXProfe, Profesores " & _
               "WHERE ClasesXProfe.idCurso = " & Val(id_Curso) & " AND " & _
               "Profesores.id = ClasesXProfe.idProfesor " & _
               "ORDER BY ClasesXProfe.Fecha"
        .Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
    End With
End Sub

Private Sub ARMAR_GRILLA()
        Set dbgClasesDictadas.DataSource = adoTabla
        
        dbgClasesDictadas.Columns(0).Visible = False
        dbgClasesDictadas.Columns(1).Width = 1200
        dbgClasesDictadas.Columns(2).Caption = "Docente"
        dbgClasesDictadas.Columns(2).Width = 2600
        
End Sub
