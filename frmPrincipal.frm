VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Gestión de Institutos de Enseñanza v 5.8"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15360
   Icon            =   "frmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   15360
   Begin VB.Frame Frame1 
      Caption         =   "Facturación acumulada del mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11160
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   2400
         TabIndex        =   16
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   2400
         TabIndex        =   15
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   2400
         TabIndex        =   14
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label lblAcumuladoNestor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   2520
         TabIndex        =   13
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label lblAcumuladoSergio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label lblAcumuladoFabru 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C000&
         Caption         =   "Néstor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2370
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         Caption         =   "Sergio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   "Fabru S.A."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdActualizarSucursal 
      Caption         =   "Actualizar Sucursal"
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerarFotos 
      Caption         =   "Generar fotos"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   4080
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowTitle     =   "SISTEMA DE GESTIÓN DE INSTITUTOS EDUCATIVOS v1.0 - LISTADO"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label lblCobranza 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Cobranza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblFacturacionBis 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "(cursos presenciales)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblInscripcion 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Inscripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblFacturacion 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblAlumnos 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Alumnos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Archivo"
      Index           =   10
      Begin VB.Menu mnuArchivo 
         Caption         =   "&Alumnos"
         Index           =   10
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "&Empresas"
         Index           =   25
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   27
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "&Profesores"
         Index           =   30
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   40
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "&Cursos"
         Index           =   50
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   60
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "&Tablas"
         Index           =   70
         Begin VB.Menu mnuTablas 
            Caption         =   "&Cursos"
            Index           =   10
            Begin VB.Menu mnuTablasCursos 
               Caption         =   "&Aulas"
               Index           =   10
            End
            Begin VB.Menu mnuTablasCursos 
               Caption         =   "-"
               Index           =   20
            End
            Begin VB.Menu mnuTablasCursos 
               Caption         =   "&Horarios"
               Index           =   30
            End
            Begin VB.Menu mnuTablasCursos 
               Caption         =   "-"
               Index           =   40
            End
            Begin VB.Menu mnuTablasCursos 
               Caption         =   "&Duraciones"
               Index           =   50
            End
            Begin VB.Menu mnuTablasCursos 
               Caption         =   "-"
               Index           =   60
            End
            Begin VB.Menu mnuTablasCursos 
               Caption         =   "&Tipos de curso"
               Index           =   90
            End
         End
         Begin VB.Menu mnuTablas 
            Caption         =   "-"
            Index           =   20
         End
         Begin VB.Menu mnuTablas 
            Caption         =   "&Gestión"
            Index           =   30
            Begin VB.Menu mnuTablasGestion 
               Caption         =   "&Condiciones de IVA"
               Index           =   10
            End
         End
         Begin VB.Menu mnuTablas 
            Caption         =   "-"
            Index           =   40
         End
         Begin VB.Menu mnuTablas 
            Caption         =   "&Lugares"
            Index           =   45
            Begin VB.Menu mnuTablasLugares 
               Caption         =   "&Sucursales"
               Index           =   10
            End
            Begin VB.Menu mnuTablasLugares 
               Caption         =   "-"
               Index           =   20
            End
            Begin VB.Menu mnuTablasLugares 
               Caption         =   "&Localidades"
               Index           =   30
            End
         End
         Begin VB.Menu mnuTablas 
            Caption         =   "-"
            Index           =   47
         End
         Begin VB.Menu mnuTablas 
            Caption         =   "&Otras"
            Index           =   50
            Begin VB.Menu mnuTablasOtras 
               Caption         =   "&Tipos de documento"
               Index           =   10
            End
            Begin VB.Menu mnuTablasOtras 
               Caption         =   "-"
               Index           =   20
            End
            Begin VB.Menu mnuTablasOtras 
               Caption         =   "¿&Cómo llegó?"
               Index           =   30
            End
            Begin VB.Menu mnuTablasOtras 
               Caption         =   "-"
               Index           =   40
            End
            Begin VB.Menu mnuTablasOtras 
               Caption         =   "C&ompanías de telefonía celular"
               Index           =   50
            End
         End
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   80
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "&Salir"
         Index           =   90
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Gestion"
      Index           =   20
      Begin VB.Menu mnuInscripcion 
         Caption         =   "Inscripción"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFacturaCursosPresenciales 
         Caption         =   "Factura - Cursos presenciales"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFacturaAnticipo 
         Caption         =   "Factura - Anticipo"
      End
      Begin VB.Menu mnuCobranzaCuentaCorriente 
         Caption         =   "Cobranza - Cuenta corriente"
      End
      Begin VB.Menu mnuNCCursosPresenciales 
         Caption         =   "Nota de crédito - Cursos presenciales"
      End
      Begin VB.Menu mnuLista4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFacturaCursosDistancia 
         Caption         =   "Factura - Cursos a distancia"
      End
      Begin VB.Menu mnuFacturaServicioTecnico 
         Caption         =   "Factura - Servicio técnico"
      End
      Begin VB.Menu mnuFacturaVentaHardware 
         Caption         =   "Factura - Venta de hardware"
      End
      Begin VB.Menu mnuNCOtros 
         Caption         =   "Nota de crédito - Otros"
      End
      Begin VB.Menu mnuLinea5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGestionOtras 
         Caption         =   "Otras opciones"
         Begin VB.Menu mnuGestionOtrasEliminarModelo 
            Caption         =   "Eliminar modelo"
         End
         Begin VB.Menu mnuLinea51 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGestionOtrasAnularFactura 
            Caption         =   "Anular factura"
         End
         Begin VB.Menu mnuLinea52 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGestionOtrasAnularRecibo 
            Caption         =   "Anular recibo"
         End
         Begin VB.Menu mnuLinea53 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGestionOtrasReimprimir 
            Caption         =   "Reimprimir comprobante"
         End
         Begin VB.Menu mnuLinea54 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGestionOtrasConsultar 
            Caption         =   "Consultar comprobante"
         End
      End
      Begin VB.Menu mnuLinea6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiquidacionProfesor 
         Caption         =   "Liquidación profesor"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Listados"
      Index           =   30
      Begin VB.Menu mnuListAlumnos 
         Caption         =   "Alumnos"
         Begin VB.Menu mnuListadoAlumnos 
            Caption         =   "Alumnos"
         End
         Begin VB.Menu mnuListadoAlumnosXEmpresa 
            Caption         =   "Alumnos por empresa"
         End
         Begin VB.Menu mnuListadoAlumnosCurso 
            Caption         =   "Alumnos por curso (número)"
         End
         Begin VB.Menu mnuListadoAlumnosXTipoCurso 
            Caption         =   "Alumnos por &tipo de curso"
         End
         Begin VB.Menu mnuListadoAlumnosXCursoRealizado 
            Caption         =   "Alumnos por curso realizado"
         End
         Begin VB.Menu mnuListadoAlumnosCursosXIniciar 
            Caption         =   "Alumnos en cursos por iniciar"
         End
      End
      Begin VB.Menu mnuListCursos 
         Caption         =   "Cursos"
         Begin VB.Menu mnuListadoCursosXIniciar 
            Caption         =   "Cursos abiertos por iniciar"
         End
         Begin VB.Menu mnuListadoCursosIniciados 
            Caption         =   "Cursos abiertos iniciados"
         End
         Begin VB.Menu mnuListadoAbandonosCurso 
            Caption         =   "Abandonos por curso"
         End
         Begin VB.Menu mnuListadoCursosAbiertos 
            Caption         =   "Cursos abiertos"
         End
         Begin VB.Menu mnuListadoCuotasACobrar 
            Caption         =   "Cuotas a cobrar"
         End
      End
      Begin VB.Menu mnuListadoVentas 
         Caption         =   "Ventas"
      End
      Begin VB.Menu mnuListadoInscriptos 
         Caption         =   "Inscriptos"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Herramientas"
      Index           =   40
      Begin VB.Menu mnuHerramientasConfiguracion 
         Caption         =   "Configuración"
         Begin VB.Menu mnuConfiguracionProximosInicios 
            Caption         =   "Próximos inicios"
         End
      End
      Begin VB.Menu cmdHerramientasAdministracionContrasenas 
         Caption         =   "Administración de contraseñas"
      End
      Begin VB.Menu mnuHerramientasArchivoConfiguracion 
         Caption         =   "Archivo de configuración"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdActualizarSucursal_Click()

    If (MsgBox("¿Confirma que desea actualizar los números de sucursal?", vbQuestion + vbYesNo, "ACTUALIZAR SUCURSALES")) = vbYes Then
    
        sSql = "UPDATE Movimientos SET Sucursal = '0004' WHERE Sucursal = '0001'"
        adoConnection.Execute sSql
        
        sSql = "UPDATE Movimientos SET Sucursal = '0005' WHERE Sucursal = '0002'"
        adoConnection.Execute sSql
        
        cmdActualizarSucursal.Visible = False

    End If
End Sub

Private Sub cmdGenerarFotos_Click()
    For k = 62 To 4305
        Me.Caption = k
        FileCopy App.Path & "\fotos\sinfoto.jpg", App.Path & "\fotos\" & k & ".jpg"
    Next
End Sub

Private Sub cmdHerramientasAdministracionContrasenas_Click()
    frmClave.Show vbModal
    If Not AccesoPermitido Then
        MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
        Exit Sub
    End If
    
    frmAdministrarContrasenas.Show vbModal

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        CALCULAR_ACUMULADO_MENSUAL
    End If
    
    If KeyCode = vbKeyF7 Then
        cmdGenerarFotos_Click
    End If
    
    If KeyCode = vbKeyF9 Then
        cmdActualizarSucursal.Visible = True
    End If
End Sub

Private Sub Form_Load()
    'MOSTRAR_ACUMULADO_MENSUAL
End Sub

Private Sub lblAlumnos_Click()
    sMenu = "Alumnos"
    frmAlumnos.Show vbModal
End Sub

Private Sub lblCobranza_Click()
    sMenu = "CobranzaCuentaCorriente"
    frmRecibo.Show
End Sub

Private Sub lblFacturacion_Click()
    sMenu = "FacturaPresenciales"
    frmFactura.Show
End Sub

Private Sub lblFacturacionBis_Click()
    sMenu = "FacturaPresenciales"
    frmFactura.Show
End Sub

Private Sub lblInscripcion_Click()
    sMenu = "Inscripcion"
    frmInscripcionBis.Show vbModal
End Sub

Private Sub mnuArchivo_Click(Index As Integer)
    Select Case Index
        Case 10: 'Alumnos
            sMenu = "Alumnos"
            frmAlumnos.Show vbModal
        Case 25: 'Empresas
            sMenu = "Empresas"
            frmEmpresas.Show vbModal
        Case 30: 'Profesores
            sMenu = "Profesores"
            frmProfesores.Show vbModal
        Case 50: 'Cursos
            sMenu = "Cursos"
            frmCursos.Show vbModal
        Case 90: 'Salir
            CERRAR_TODO
            End
    End Select
End Sub


Private Sub mnuClasesDictadas_Click()
    frmClave.Show vbModal
    If Not AccesoPermitido Then
        MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
        Exit Sub
    End If
    
    sMenu = "ClasesDictadas"
    frmClasesDictadas.Show vbModal
End Sub

Private Sub mnuCobranzaCuentaCorriente_Click()
    sMenu = "CobranzaCuentaCorriente"
    frmRecibo.Show
End Sub

Private Sub mnuConfiguracionProximosInicios_Click()
    frmProximosInicios.Show vbModal
End Sub

Private Sub mnuFacturaAnticipo_Click()
    sMenu = "FacturaAnticipo"
    frmFactura.Show
End Sub

Private Sub mnuFacturaCursosDistancia_Click()
    sMenu = "FacturaDistancia"
    frmFactura.Show
End Sub

Private Sub mnuFacturaCursosPresenciales_Click()
    sMenu = "FacturaPresenciales"
    frmFactura.Show
End Sub

Private Sub mnuFacturaServicioTecnico_Click()
    sMenu = "FacturaServicioTecnico"
    frmFactura.Show
End Sub

Private Sub mnuFacturaVentaHardware_Click()
    sMenu = "FacturaVentaHardware"
    frmFactura.Show
End Sub

Private Sub mnuGestionOtrasAnularFactura_Click()
    
    sMenu = "GestionOtrasAnulaFactura"
    
    AccesoPublico = True
        frmClave.Show vbModal
    AccesoPublico = False
       
    If Not AccesoPermitido Then
        MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
        Exit Sub
    End If
    
    TipoBusDoc = "AnularFactura"
    frmBusDoc.Show vbModal
End Sub

Private Sub mnuGestionOtrasAnularRecibo_Click()
    sMenu = "GestionOtrasAnulaRecibo"
    
    AccesoPublico = True
        frmClave.Show vbModal
    AccesoPublico = False
    
    If Not AccesoPermitido Then
        MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
        Exit Sub
    End If
    
    TipoBusDoc = "AnularRecibo"
    frmBusDoc.Show vbModal
End Sub

Private Sub mnuGestionOtrasConsultar_Click()
    TipoBusDoc = "Consultar"
    frmBusDoc.Show vbModal
End Sub

Private Sub mnuGestionOtrasEliminarModelo_Click()
    frmClave.Show vbModal
    If Not AccesoPermitido Then
        MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
        Exit Sub
    End If
    
    TipoBusDoc = "BorrarModelo"
    frmBusDoc.Show vbModal
End Sub

Private Sub mnuGestionOtrasReimprimir_Click()
    TipoBusDoc = "Reimprimir"
    frmBusDoc.Show vbModal
End Sub

Private Sub mnuHerramientasArchivoConfiguracion_Click()
    frmConfiguracion.Show
End Sub

Private Sub mnuInscripcion_Click()
    sMenu = "Inscripcion"
    frmInscripcionBis.Show
End Sub


Private Sub mnuLiquidacionProfesor_Click()
    frmLiquidacionProfesor.Show
End Sub

Private Sub mnuListadoAbandonosCurso_Click()
    frmListaAbandonosXCurso.Show
End Sub

Public Sub mnuListadoAlumnos_Click()
    frmAlumnos.cmdListado_Click
End Sub

Private Sub mnuListadoAlumnosCurso_Click()
    frmListaAlumnosCursoNumero.Show
End Sub

Private Sub mnuListadoAlumnosCursosXIniciar_Click()
    frmListaAlumnosCursosXIniciar.Show
End Sub

Private Sub mnuListadoAlumnosXTipoCurso_Click()
    On Error GoTo ErrorHandle
        
    tipoFiltro = "Cursos"
    frmFiltro.Show vbModal
    
    If sqlFiltro = "" Then
        Exit Sub
    End If
    
    sSql = "DELETE FROM zzAlumnosXCurso"
    adoConnection.Execute sSql
    
    sSql = "INSERT INTO zzAlumnosXCurso (Nombre, Telefono, TelefonoLaboral, Celular, Mail, Curso, Horario) " & _
           "SELECT Alumnos.Nombre, Alumnos.Telefono, Alumnos.TelefonoLaboral, Alumnos.Celular, Alumnos.Mail, TiposCurso.Detalle, Horarios.Detalle " & _
           "FROM Alumnos, TiposCurso, Horarios, Cursos, AlumnosXCurso " & _
           "WHERE " & sqlFiltro & " AND TiposCurso.id = Cursos.idTipoCurso AND Horarios.id = Cursos.idHorario " & _
           "      AND Alumnos.id = AlumnosXCurso.idAlumno AND Cursos.id = AlumnosXCurso.idCurso " & _
           "ORDER BY TiposCurso.Detalle, Cursos.Numero, Alumnos.Nombre"
    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    rptListado.ReportFileName = App.Path & "\reportes\rptAlumnosXCurso.rpt"
    rptListado.ReportTitle = "LISTADO DE ALUMNOS POR CURSO"
    rptListado.Action = 1

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmPrincipal - mnuListadoAlumnosXCurso", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub mnuListadoAlumnosXCursoRealizado_Click()
    On Error GoTo ErrorHandle
        
    frmListaAlumnosXCursosRealizados.Show
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmPrincipal - mnuListadoAlumnosXCursoRealizado", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub mnuListadoAlumnosXEmpresa_Click()
    On Error GoTo ErrorHandle
        
    tipoFiltro = "Empresas"
    frmFiltro.Show vbModal
    
    If sqlFiltro = "" Then
        Exit Sub
    End If


    sSql = "DELETE FROM zzAlumnosXEmpresa"
    adoConnection.Execute sSql
    
    sSql = "INSERT INTO zzAlumnosXEmpresa (Empresa, Alumno) " & _
           "SELECT Empresas.Nombre, Alumnos.Nombre " & _
           "FROM Empresas, Alumnos " & _
           "WHERE " & sqlFiltro & " AND Empresas.id = Alumnos.idEmpresa " & _
           "ORDER BY Empresas.Nombre, Alumnos.Nombre"

    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    rptListado.ReportFileName = App.Path & "\reportes\rptAlumnosXEmpresa.rpt"
    rptListado.ReportTitle = "LISTADO DE ALUMNOS POR EMPRESA"
    rptListado.Action = 1

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmPrincipal - mnuListadoAlumnosXEmpresa", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub mnuListadoCuotasACobrar_Click()
    frmListaCuotasACobrar.Show
End Sub

Private Sub mnuListadoCursosAbiertos_Click()
    On Error GoTo ErrorHandle
        
    tipoFiltro = "Cursos"
    frmFiltro.Show vbModal
    
    If sqlFiltro = "" Then
        Exit Sub
    End If

    sSql = "DELETE FROM zzCursosAbiertos"
    adoConnection.Execute sSql
    
    sSql = "INSERT INTO zzCursosAbiertos (NumCurso, FechaIni, FechaFin, Horario, TipoCurso) " & _
           "SELECT Cursos.Numero, Cursos.FechaIni, Cursos.FechaFin, Horarios.Detalle, TiposCurso.Detalle " & _
           "FROM Cursos, TiposCurso, Horarios " & _
           "WHERE " & sqlFiltro & " AND TiposCurso.id = Cursos.idTipoCurso AND Horarios.id = Cursos.idHorario AND Cursos.Abierto " & _
           "ORDER BY TiposCurso.Detalle, Cursos.Numero"

    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    rptListado.ReportFileName = App.Path & "\reportes\rptCursosAbiertos.rpt"
    rptListado.ReportTitle = "LISTADO DE CURSOS ABIERTOS"
    rptListado.Action = 1

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmPrincipal - mnuListadoCursosAbiertos", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub mnuListadoCursosIniciados_Click()
    On Error GoTo ErrorHandle
        
    tipoFiltro = "Cursos"
    frmFiltro.Show vbModal
    
    If sqlFiltro = "" Then
        Exit Sub
    End If

    sSql = "DELETE FROM zzCursosXIniciar"
    adoConnection.Execute sSql
    
    sSql = "INSERT INTO zzCursosXIniciar (NumCurso, FechaIni, Horario, TipoCurso) " & _
           "SELECT Cursos.Numero, Cursos.FechaIni, Horarios.Detalle, TiposCurso.Detalle " & _
           "FROM Cursos, TiposCurso, Horarios " & _
           "WHERE " & sqlFiltro & " AND Cursos.FechaIni < DateValue('" & Date & "') AND TiposCurso.id = Cursos.idTipoCurso AND Horarios.id = Cursos.idHorario AND Cursos.Abierto " & _
           "ORDER BY TiposCurso.Detalle, Cursos.Numero"

    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    rptListado.ReportFileName = App.Path & "\reportes\rptCursosXIniciar.rpt"
    rptListado.ReportTitle = "LISTADO DE ABIERTOS INICIADOS"
    rptListado.Action = 1

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmPrincipal - mnuListadoCursosXIniciar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub mnuListadoCursosXIniciar_Click()
    On Error GoTo ErrorHandle
        
    tipoFiltro = "Cursos"
    frmFiltro.Show vbModal
    
    If sqlFiltro = "" Then
        Exit Sub
    End If

    sSql = "DELETE FROM zzCursosXIniciar"
    adoConnection.Execute sSql
    
    sSql = "INSERT INTO zzCursosXIniciar (NumCurso, FechaIni, Horario, TipoCurso) " & _
           "SELECT Cursos.Numero, Cursos.FechaIni, Horarios.Detalle, TiposCurso.Detalle " & _
           "FROM Cursos, TiposCurso, Horarios " & _
           "WHERE " & sqlFiltro & " AND Cursos.FechaIni > DateValue('" & Date & "') AND TiposCurso.id = Cursos.idTipoCurso AND Horarios.id = Cursos.idHorario AND Cursos.Abierto " & _
           "ORDER BY TiposCurso.Detalle, Cursos.Numero"

    adoConnection.Execute sSql
    
    GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    rptListado.ReportFileName = App.Path & "\reportes\rptCursosXIniciar.rpt"
    rptListado.ReportTitle = "LISTADO DE CURSOS ABIERTOS POR INICIAR"
    rptListado.Action = 1

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmPrincipal - mnuListadoCursosXIniciar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub


Private Sub mnuListadoInscriptos_Click()
    frmListaMatriculas.Show
End Sub

Private Sub mnuListadoVentas_Click()
    frmListaVentas.Show
End Sub

Private Sub mnuNCCursosPresenciales_Click()
    sMenu = "NotaCreditoPresenciales"
    frmFactura.Show
End Sub

Private Sub mnuNCOtros_Click()
    sMenu = "NotaCreditoOtros"
    frmFactura.Show
End Sub

Private Sub mnuTablasCursos_Click(Index As Integer)
    Select Case Index
        Case 10: 'Aulas
            strTabla = "Aulas"
            frmTabla.Show vbModal
        Case 30: 'Horarios
            strTabla = "Horarios"
            frmTabla.Show vbModal
        Case 50: 'Duraciones
            strTabla = "Duraciones"
            frmTabla.Show vbModal
        Case 70: 'Modalidades
            strTabla = "Modalidades"
            frmTabla.Show vbModal
        Case 90: 'Tipos de curso
            strTabla = "TiposCurso"
            frmTabla.Show vbModal
    End Select
End Sub

Private Sub mnuTablasGestion_Click(Index As Integer)
    Select Case Index
        Case 10: 'Condiciones de IVA
            strTabla = "CondIva"
            frmTabla.Show vbModal
    End Select
End Sub

Private Sub mnuTablasLugares_Click(Index As Integer)
    Select Case Index
        Case 10: 'Sucursales
            strTabla = "Sucursales"
            frmTabla.Show vbModal
        Case 30: 'Localidades
            strTabla = "Localidades"
            frmTabla.Show vbModal
    End Select
End Sub

Private Sub mnuTablasOtras_Click(Index As Integer)
    Select Case Index
        Case 10: 'Tipos de documento
            strTabla = "TiposDoc"
            frmTabla.Show vbModal
        Case 30: '¿Cómo llegó?
            strTabla = "ComoLlego"
            frmTabla.Show vbModal
        Case 50: 'Companías de celular
            frmTablaCompaniasCelular.Show vbModal
    End Select
End Sub

