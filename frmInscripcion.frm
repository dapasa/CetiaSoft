VERSION 5.00
Begin VB.Form frmInscripcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INSCRIPCIÓN"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "frmInscripcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFactura 
      Caption         =   "&Factura"
      Height          =   855
      Left            =   6960
      Picture         =   "frmInscripcion.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Ver lista de espera"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdVerInscriptos 
      Caption         =   "&Inscriptos"
      Height          =   855
      Left            =   5040
      Picture         =   "frmInscripcion.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Ver inscriptos"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdVerListaEspera 
      Caption         =   "&Espera"
      Height          =   855
      Left            =   6000
      Picture         =   "frmInscripcion.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ver lista de espera"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevoAlumno 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   3120
      Picture         =   "frmInscripcion.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Nuevo alumno"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame3 
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
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
      Begin VB.TextBox txtAlumno 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdInscribir 
      Caption         =   "&Inscribir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4080
      Picture         =   "frmInscripcion.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Inscribir al curso"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame2 
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
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton optVacantesTodas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optVacantesEnEspera 
         Caption         =   "En espera"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optVacantesDisponibles 
         Caption         =   "Disponibles"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1120
      End
   End
   Begin VB.CommandButton cmdSeleccionarAlumno 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   2160
      Picture         =   "frmInscripcion.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Seleccionar alumno"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7920
      Picture         =   "frmInscripcion.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   3960
      Width           =   855
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
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8655
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
         Height          =   2370
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   8415
      End
      Begin VB.Label Label1 
         Caption         =   "NºInt.|Inicio    |Fin       |Duración                      |Vac |Insc |Espera"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8415
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cboTipoCurso 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   8760
      Y1              =   3840
      Y2              =   3840
   End
End
Attribute VB_Name = "frmInscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' MÓDULO: Inscripción a cursos        FECHA: Ago / 2007
'******************************************************
' RESUMEN:
'******************************************************
' ÚLTIMA MODIFICACIÓN IMPORTANTE: 16/08/2007
'******************************************************
' ETAPA: desarrollo.
'******************************************************
' AUTOR: Pablo Adrián Langholz
' CONTACTO: elmaildepablo@gmail.com
'******************************************************

Dim EventoLoad As Boolean
Dim id_ As String, FechaIni_ As String, FechaFin_ As String
Dim Detalle_ As String, Vacantes_ As String, Inscriptos_ As String, Espera_ As String
Dim Cuotas As Byte, ValorCuota As Single
Dim curso_actual As Byte

Private Sub cboTipoCurso_Click()
    If Not EventoLoad Then
        CARGAR_CURSOS "Todas"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFactura_Click()
    Unload Me
    sMenu = "FacturaPresenciales"
    frmFactura.Show
End Sub

Private Sub cmdInscribir_Click()
    If lstCursos.ListIndex <> -1 Then
        'Obtengo el ID del curso seleccionado
        id_Curso = Val(Left(lstCursos.List(lstCursos.ListIndex), 5))
        
        CERRAR_TABLA adoAlumnosXCurso
        sSql = "SELECT * FROM AlumnosXCurso WHERE idCurso = " & id_Curso & " AND idAlumno = " & id_Alumno
        adoAlumnosXCurso.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        If adoAlumnosXCurso.EOF Then
            CERRAR_TABLA adoCursos
            sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso
            adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            If adoCursos!Vacantes > 0 Then
                'Cuotas = adoCursos!CantCuotas
                'ValorCuota = adoCursos!ValorCuota
                
                INSCRIBIR
                frmCuotas.Show vbModal
                GENERAR_DOCUMENTOS
                
                MsgBox "El alumno " & txtAlumno.Text & " ha sido inscripto en el curso " & cboTipoCurso.Text & " los " & Detalle_, vbInformation, "INSCRIPCIÓN"
            Else
                INSCRIBIR_LISTA_ESPERA
                MsgBox "El alumno " & txtAlumno.Text & " ha sido puesto en la lista de espera para el curso " & cboTipoCurso.Text & " los " & Detalle_, vbInformation, "INSCRIPCIÓN"
            End If
            
            adoCursos.Close
            
            If optVacantesDisponibles Then
                CARGAR_CURSOS "Disponibles"
            ElseIf optVacantesEnEspera Then
                CARGAR_CURSOS "EnEspera"
            ElseIf optVacantesTodas Then
                CARGAR_CURSOS "Todas"
            End If
        Else
            MsgBox "El alumno " & txtAlumno.Text & " ya está inscripto en el curso " & cboTipoCurso.Text & " los " & Detalle_, vbCritical, "INSCRIPCIÓN"
        End If
    Else
        MsgBox "Debe seleccionar un curso.", vbExclamation, "INSCRIPCIÓN"
    End If
End Sub

Private Sub cmdNuevoAlumno_Click()
    frmAlumnos.Show vbModal
End Sub

Private Sub cmdSeleccionarAlumno_Click()
    EstiloBuscador = "Alumnos"
    frmBuscador.Show vbModal
End Sub

Private Sub cmdVerInscriptos_Click()
    If lstCursos.ListIndex <> -1 Then
        curso_actual = lstCursos.ListIndex
        
        TipoLista = "Inscriptos"
        
        id_ = Left(lstCursos.List(lstCursos.ListIndex), 5)
        
        CERRAR_TABLA adoTabla
        sSql = "SELECT Alumnos.id, Alumnos.Nombre, Alumnos.Telefono, Alumnos.Mail, Alumnos.Celular FROM Alumnos, AlumnosXCurso WHERE Alumnos.id = AlumnosXCurso.idAlumno AND AlumnosXCurso.idCurso = " & Val(id_) & " ORDER BY Nombre"
        adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        With frmListaAlumnos
            .Caption = .Caption & " INSCRIPTOS"
            .fraListaAlumnos.Caption = "Alumnos inscriptos en el curso"
            .lblListaAlumnos.Caption = "Nº Int. |Nombre                                 |Teléfono           |e-Mail                                 |Celular            "
            
            Do While Not adoTabla.EOF
                .lstListaAlumnos.AddItem Right(Space(6) & adoTabla!id, 6) & "  " & Left(adoTabla!Nombre & Space(40), 40) & Left(adoTabla!Telefono & Space(20), 20) & Left(adoTabla!Mail & Space(40), 40) & Left(adoTabla!Celular & Space(20), 20)
                
                adoTabla.MoveNext
            Loop
            adoTabla.Close
            
            .Show vbModal
            CARGAR_CURSOS "Disponibles"
            lstCursos.ListIndex = curso_actual
        End With
        
        If banVolverAListaAlumnos Then
            banVolverAListaAlumnos = False
            cmdVerListaEspera_Click
        End If
    Else
        MsgBox "Debe seleccionar un curso.", vbExclamation, "INSCRIPCIÓN"
    End If
End Sub

Private Sub cmdVerListaEspera_Click()
    If lstCursos.ListIndex <> -1 Then
        TipoLista = "Espera"
        
        CERRAR_TABLA adoTabla
        sSql = "SELECT ListaEspera.id, Alumnos.Nombre, ListaEspera.FechaDesde, Alumnos.Telefono FROM Alumnos, ListaEspera WHERE Alumnos.id = ListaEspera.idAlumno AND ListaEspera.idCurso = " & Val(id_) & " ORDER BY FechaDesde"
        adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        With frmListaAlumnos
            .Caption = .Caption & " EN ESPERA"
            .fraListaAlumnos.Caption = "Alumnos en espera para el curso"
            .lblListaAlumnos.Caption = "Nº Int. |Nombre                                 |Desde     |Teléfono"
            
            Do While Not adoTabla.EOF
                .lstListaAlumnos.AddItem Right(Space(6) & adoTabla!id, 6) & "  " & Left(adoTabla!Nombre & Space(40), 40) & adoTabla!FechaDesde & " " & adoTabla!Telefono
                
                adoTabla.MoveNext
            Loop
            adoTabla.Close
            
            BOTON_INSCRIBIR
            
            .Show vbModal
        End With
    Else
        MsgBox "Debe seleccionar un curso.", vbExclamation, "INSCRIPCIÓN"
    End If
End Sub

Private Sub Form_Load()
    EventoLoad = True
    
    CARGAR_COMBO "cboTipoCurso", adoTiposCurso, "TiposCurso", "Detalle", Me
    
    EventoLoad = False
End Sub

Private Sub CARGAR_CURSOS(Estado As String)
    Dim filtroEstado As String
    
    Select Case Estado
        Case "Disponibles"
            filtroEstado = "Cursos.Vacantes > 0"
        Case "EnEspera"
            filtroEstado = "Cursos.Vacantes = 0"
        Case "Todas"
            filtroEstado = "1 = 1"
    End Select
    
    lstCursos.Clear
    
    With adoCursos
        CERRAR_TABLA adoCursos
        sSql = "SELECT Cursos.id, Cursos.FechaIni, Cursos.FechaFin, Horarios.Detalle, Cursos.Vacantes, Cursos.Inscriptos " & _
               "FROM Cursos, Horarios " & _
               "WHERE Cursos.idTipoCurso = " & DEVOLVER_ID(cboTipoCurso.Text, adoTiposCurso, "TiposCurso", "Detalle") & " " & _
               "AND Cursos.idHorario = Horarios.id " & _
               "AND " & filtroEstado & " " & _
               "ORDER BY Cursos.FechaIni"
        .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
        Do While Not .EOF
            id_ = Left(!id & Space(5), 5)
            FechaIni_ = !FechaIni
            FechaFin_ = !FechaFin
            Detalle_ = Left(!Detalle & Space(30), 30)
            Vacantes_ = Left(!Vacantes & Space(6), 4)
            Inscriptos_ = Left(!inscriptos & Space(2), 2)
            
            CERRAR_TABLA adoListaEspera
            sSql = "SELECT COUNT(id) AS CantEspera FROM ListaEspera WHERE idCurso = " & id_
            adoListaEspera.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            Espera_ = adoListaEspera!CantEspera
            adoListaEspera.Close
            
            lstCursos.AddItem id_ & " " & FechaIni_ & " " & FechaFin_ & " " & Detalle_ & " " & Vacantes_ & " " & Inscriptos_ & " " & Espera_
            
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CERRAR_TODO
End Sub

Private Sub optVacantesDisponibles_Click()
    CARGAR_CURSOS "Disponibles"
End Sub

Private Sub optVacantesEnEspera_Click()
    CARGAR_CURSOS "EnEspera"
End Sub

Private Sub optVacantesTodas_Click()
    CARGAR_CURSOS "Todas"
End Sub

Private Sub txtAlumno_Change()
    If txtAlumno.Text <> "" Then
        cmdInscribir.Enabled = True
    Else
        cmdInscribir.Enabled = False
    End If
End Sub

Private Sub INSCRIBIR()
    'Actualizo inscriptos y vacantes del curso
    With adoCursos
        !inscriptos = adoCursos!inscriptos + 1
        !Vacantes = adoCursos!Vacantes - 1
    
        .Update
    End With
    
    'Actualizo AlumnosXCurso
    CERRAR_TABLA adoAlumnosXCurso
    With adoAlumnosXCurso
        .Open "AlumnosXCurso", adoConnection, adOpenDynamic, adLockOptimistic
        .AddNew
        !idAlumno = id_Alumno
        !idCurso = id_Curso
        .Update
        .Close
    End With
End Sub

Private Sub INSCRIBIR_LISTA_ESPERA()
    'Actualizo ListaEspera
    CERRAR_TABLA adoListaEspera
    With adoListaEspera
        .Open "ListaEspera", adoConnection, adOpenDynamic, adLockOptimistic
        .AddNew
        !idAlumno = id_Alumno
        !idCurso = id_Curso
        !FechaDesde = Date
        !Observaciones = "" 'Hay que armar un TextBox para este campo
        .Update
        .Close
    End With
End Sub

Private Sub BOTON_INSCRIBIR()
    CERRAR_TABLA adoTabla
    sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso & " AND Vacantes > 0"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoTabla.EOF Then
        frmListaAlumnos.cmdInscribir.Enabled = True
        frmListaAlumnos.cmdInscribir.Visible = True
    End If
    
    adoTabla.Close
End Sub

Private Sub GENERAR_DOCUMENTOS()
    Dim ultimo_id As Long
'   Dim ultimo_id As Integer
    
    For k = 0 To x_cuotas
        'Genero el encabezado
        With adoMovimientos
            CERRAR_TABLA adoMovimientos
            .Open "Movimientos", adoConnection, adOpenDynamic, adLockOptimistic
            
            .AddNew
            
            !Sucursal = "0000" 'TODAVÍA NO SE A QUE SUCURSAL VA
            !TipoDoc = "MOD"
            !NumDoc = ULTIMO_NUMERO("MOD")
            !idAlumno = id_Alumno
            !idCurso = id_Curso
            
            If k = 0 Then 'Es matrícula
                !fecha = Date
                
                !Subtotal = x_valorMatricula
                !Iva = 0
                !Total = x_valorMatricula
                !Saldo = x_valorMatricula
                !Descuento = x_DetalleDescuento
            Else
                CERRAR_TABLA adoTemp
                sSql = "SELECT * FROM Cursos WHERE id = " & id_Curso
                adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                
                campo = "Cuota" & k
                !fecha = adoTemp.Fields(campo).Value
                
                adoTemp.Close
                
                !Subtotal = x_valorCuota
                !Iva = 0
                !Total = x_valorCuota
                !Saldo = x_valorCuota
                !Descuento = x_DetalleDescuento
                !cuota = k
            End If
            
            .Update
            
            .MoveLast
            ultimo_id = !id
            
            .Close
        End With
        
        'Genero el cuerpo
        With adoItemsXMov
            CERRAR_TABLA adoItemsXMov
            .Open "ItemsXMov", adoConnection, adOpenDynamic, adLockOptimistic
            
            .AddNew
            
            !idMovimiento = ultimo_id
            !idCurso = id_Curso
            !Cantidad = 1
            
            If k = 0 Then 'Es matrícula
                !Detalle = "Matrícula - " & cboTipoCurso.Text & " " & Detalle_
                !Unitario = x_valorMatriculaReal
                
                If x_DetalleDescuento = "EX ALUMNO" Then
                    !Descuento = 15
                End If
                
                !Importe = x_valorMatricula
                !Saldo = x_valorMatricula
            Else
                !Detalle = "Cuota " & k & "/" & x_cuotas & " - " & cboTipoCurso.Text & " " & Detalle_
                !Unitario = x_valorCuotaReal
                
                If x_DetalleDescuento = "EX ALUMNO" Then
                    !Descuento = 15
                End If
                
                !Importe = x_valorCuota
                !Saldo = x_valorCuota
            End If
            
            .Update

            'If x_DetalleDescuento = "EX ALUMNO" Then
            '    .AddNew
            '    !idMovimiento = ultimo_id
            '    !idCurso = id_Curso
            '    !Cantidad = 1
            '
            '    !Detalle = "BONIFICACION " & x_DetalleDescuento
            '
            '    If k = 0 Then
            '        !Unitario = x_valorMatricula - x_valorMatriculaReal
            '        !Importe = x_valorMatricula - x_valorMatriculaReal
            '        !Saldo = x_valorMatricula - x_valorMatriculaReal
            '    Else
            '        !Unitario = x_valorCuota - x_valorCuotaReal
            '        !Importe = x_valorCuota - x_valorCuotaReal
            '        !Saldo = x_valorCuota - x_valorCuotaReal
            '    End If
            '    .Update
            'End If
                        
                        
            .Close
        End With
    Next
End Sub
