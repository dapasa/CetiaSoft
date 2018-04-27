VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAlumnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALUMNOS"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "frmAlumnos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInscripcion 
      Caption         =   "&Inscribir"
      Height          =   855
      Left            =   7320
      Picture         =   "frmAlumnos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Inscribir al curso"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir ficha"
      Height          =   855
      Left            =   9240
      Picture         =   "frmAlumnos.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Generar factura"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdFactura 
      Caption         =   "&Factura"
      Height          =   855
      Left            =   8280
      Picture         =   "frmAlumnos.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Generar factura"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdCambiarFoto 
      Caption         =   "Cambiar Foto"
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame13 
      Caption         =   "Foto"
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
      Left            =   6720
      TabIndex        =   39
      Top             =   120
      Width           =   3375
      Begin VB.Image imgFoto 
         Height          =   2580
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de documento"
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
      TabIndex        =   38
      Top             =   840
      Width           =   2415
      Begin VB.ComboBox cboTipoDoc 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Nº de documento"
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
      Left            =   2640
      TabIndex        =   37
      Top             =   840
      Width           =   3975
      Begin VB.TextBox txtNumDoc 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Fecha de alta"
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
      TabIndex        =   36
      Top             =   3120
      Width           =   1815
      Begin VB.Label lblFechaAlta 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   120
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame5 
      Caption         =   "Empresa"
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
      Left            =   3480
      TabIndex        =   34
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdListado 
      Caption         =   "&Listado"
      Height          =   855
      Left            =   3840
      Picture         =   "frmAlumnos.frx":1330
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Lista de alumnos"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   960
      Picture         =   "frmAlumnos.frx":1772
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Nuevo"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2880
      Picture         =   "frmAlumnos.frx":1BB4
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Eliminar"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4800
      Picture         =   "frmAlumnos.frx":1FF6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Guardar"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5760
      Picture         =   "frmAlumnos.frx":2438
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancelar"
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   1920
      Picture         =   "frmAlumnos.frx":287A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Buscar"
      Top             =   8280
      Width           =   855
   End
   Begin VB.Frame Frame9 
      Caption         =   "Fecha nacimiento"
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
      Left            =   2040
      TabIndex        =   33
      Top             =   3120
      Width           =   1815
      Begin MSComCtl2.DTPicker dtpFechaNac 
         Height          =   290
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   17498113
         CurrentDate     =   40147
      End
   End
   Begin VB.Frame Frame11 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   32
      Top             =   6720
      Width           =   6495
      Begin VB.TextBox txtObservaciones 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "CP"
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
      TabIndex        =   31
      Top             =   5280
      Width           =   6495
      Begin VB.TextBox txtCodPostal 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Nombre"
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
      TabIndex        =   29
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtNombre 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Left            =   120
         MaxLength       =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Dirección"
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
      TabIndex        =   28
      Top             =   3840
      Width           =   6495
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   9
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comunicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   1560
      Width           =   6495
      Begin VB.TextBox txtTelefonoLaboral 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtCelular 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtMail 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1080
         Width           =   5175
      End
      Begin VB.ComboBox cboCompaniaCelular 
         Height          =   315
         Left            =   3600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tel. laboral"
         Height          =   195
         Left            =   3240
         TabIndex        =   35
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tel. particular"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "e-Mail"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cía."
         Height          =   195
         Left            =   3240
         TabIndex        =   24
         Top             =   780
         Width           =   300
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "¿Cómo llegó?"
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
      TabIndex        =   22
      Top             =   6000
      Width           =   6495
      Begin VB.ComboBox cboComoLlego 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localidad"
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
      TabIndex        =   21
      Top             =   4560
      Width           =   6495
      Begin VB.ComboBox cboLocalidad 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   6600
      Y1              =   8160
      Y2              =   8160
   End
End
Attribute VB_Name = "frmAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' MÓDULO: Mantenimiento de alumnos    FECHA: Ago / 2007
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
Dim cambiar_foto As Boolean

Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long


Private Sub cboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Then
        frmEmpresas.Show vbModal
        CARGAR_COMBO "cboEmpresa", adoEmpresas, "Empresas", "Nombre", Me
    End If
End Sub

Private Sub cboLocalidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Then
        strTabla = "Localidades"
        frmTabla.Show vbModal
        CARGAR_COMBO "cboLocalidad", adoLocalidades, "Localidades", "Detalle", Me
    End If
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo ErrorHandle
    
    BOTONES "Buscar", Me

    ENABLED_TODO True, Me
        
    EstiloBuscador = "Alumnos"
    frmBuscador.Show vbModal
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - cmdBuscar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCambiarFoto_Click()
    cambiar_foto = True
    
    frmSacarFoto.Show vbModal
    
    If id_Alumno <> 0 Then
        Kill App.Path & "\fotos\" & id_Alumno & ".jpg"
                
        'Convertir a JPG
        FileCopy App.Path & "\fotos\imagen.bmp", "c:\tmp.bmp"

        ret = DIWriteJpg(App.Path & "\fotos\" & id_Alumno & ".jpg", 75, True)
        
        Kill "c:\tmp.bmp"
        
        imgFoto.Picture = LoadPicture(App.Path & "\fotos\" & id_Alumno & ".jpg")
    Else
        'Convertir a JPG
        FileCopy App.Path & "\fotos\imagen.bmp", "c:\tmp.bmp"
        
        'Kill App.Path & "\fotos\imagen.jpg"
        
        ret = DIWriteJpg(App.Path & "\fotos\imagen.jpg", 75, True)
        
        'Kill App.Path & "c:\tmp.bmp"
    
        imgFoto.Picture = LoadPicture(App.Path & "\fotos\imagen.jpg")
    End If
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - cmdCancelar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo ErrorHandle
    
    BOTONES "Eliminar", Me
    
    If MsgBox("¿Confirma que desea eliminar al alumno " & txtNombre.Text & "?", vbYesNo + vbQuestion, "ALUMNOS") = vbYes Then
        frmClave.Show vbModal
        If Not AccesoPermitido Then
            MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
            Exit Sub
        End If
        
        If PUEDE_BORRAR("Alumnos", id_Alumno) Then
            ModoABM = "B"
            'Eliminar
            ELIMINAR_ALUMNO
        End If
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - cmdEliminar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdFactura_Click()
    x_alumno_form_inscrip_factura = txtNombre.Text
    
    Unload Me
    sMenu = "FacturaPresenciales"
    frmFactura.Show
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo ErrorHandle
    
    If MsgBox("¿Confirma que desea guardar los cambios?", vbYesNo + vbQuestion, "ALUMNOS") = vbYes Then
        BOTONES "Guardar", Me
        
        'Guardar
        If id_Alumno = 0 Then 'Alumno nuevo
            ModoABM = "A"
            NUEVO_ALUMNO
            
            CERRAR_TABLA adoTempAlumnos
            sSql = "SELECT * FROM alumnos ORDER BY id DESC"
            adoTempAlumnos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            adoTempAlumnos.MoveFirst
            
            If cambio_foto Then
                cambio_foto = False
                
                FileCopy App.Path & "\fotos\imagen.jpg", App.Path & "\fotos\" & adoTempAlumnos!id & ".jpg"
            Else
                FileCopy App.Path & "\fotos\sinfoto.jpg", App.Path & "\fotos\" & adoTempAlumnos!id & ".jpg"
            End If
            
            adoTempAlumnos.Close
            
            ModoABM = ""
        Else 'Modificación de alumno
            With adoAlumnos
                ModoABM = "M"
                MODIFICAR_ALUMNO
            End With
        End If
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - cmdGuardar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdImprimir_Click()
    Me.PrintForm
End Sub

Private Sub cmdInscripcion_Click()
    Unload Me
    
    sMenu = "Inscripcion"
    xVieneDe = "Alumnos"
    frmInscripcionBis.Show vbModal
End Sub

Public Sub cmdListado_Click()
    On Error GoTo ErrorHandle
        
    tipoFiltro = "Alumnos"
    frmFiltro.Show vbModal
    
    If sqlFiltro = "" Then
        Exit Sub
    End If
    
    sSql = "DELETE FROM zzAlumnos"
    adoConnection.Execute sSql
    
    sSql = "INSERT INTO zzAlumnos (Nombre, Telefono, TelefonoLaboral, Celular, Mail) " & _
           "SELECT Nombre, Telefono, TelefonoLaboral, Celular, Mail " & _
           "FROM Alumnos " & _
           "WHERE " & sqlFiltro & " " & _
           "ORDER BY Nombre"
    adoConnection.Execute sSql
    
    'GENERAR_LISTADO
    
    rptListado.Connect = "PWD=FiatIdea"
    rptListado.ReportFileName = App.Path & "\reportes\rptAlumnos.rpt"
    rptListado.ReportTitle = "LISTADO DE ALUMNOS"
    rptListado.Action = 1

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - cmdListado", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo ErrorHandle
    
    BOTONES "Nuevo", Me
        
    id_Alumno = 0
    
    ENABLED_TODO True, Me
    
    If sMenu <> "Inscripcion" Then
        txtNombre.SetFocus
    End If
    
    INICIALIZAR_PANTALLA

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - cmdNuevo", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub dtpFechaNac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Form_KeyPress 13
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle

    If KeyAscii = 13 Then
        PASAR_CAMPO Me
    End If

    If KeyAscii >= 97 And KeyAscii <= 122 And Me.ActiveControl.Name <> "txtMail" Then
        KeyAscii = KeyAscii - 32
    End If
    
    'CTRL + F
    If KeyAscii = 6 Then
        frmSacarFoto.Show vbModal
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - Form.KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    id_Alumno = 0
    
    ENABLED_TODO False, Me
        
    CERRAR_TABLA adoAlumnos
    sSql = "SELECT * FROM Alumnos"
    adoAlumnos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If adoAlumnos.EOF Then
        cmdBuscar.Enabled = False
    End If
    
    CARGAR_COMBO "cboEmpresa", adoEmpresas, "Empresas", "Nombre", Me
    CARGAR_COMBO "cboCompaniaCelular", adoCompaniasCelular, "CompaniasCelular", "Detalle", Me
    CARGAR_COMBO "cboLocalidad", adoLocalidades, "Localidades", "Detalle", Me
    CARGAR_COMBO "cboComoLlego", adoComoLlego, "ComoLlego", "Detalle", Me
    CARGAR_COMBO "cboTipoDoc", adoTiposDoc, "TiposDoc", "Detalle", Me
    
    dtpFechaNac.Value = Date - 8300
    
    If sMenu = "Inscripcion" Then
        cmdNuevo_Click
    End If
    
    If x_ficha_alumno_desde_factura <> "" Then
        cmdBuscar_Click
    End If
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - Form.Load", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    CERRAR_TABLA adoAlumnos

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - Form.Unload", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub NUEVO_ALUMNO()
    On Error GoTo ErrorHandle
    
    With adoAlumnos
        'Verifico que los datos sean correctos
        If VALIDAR_ALUMNO = True Then
            'Agrego el registro
            .AddNew
            
            ASIGNAR_DATOS
            'Guardo los cambios
            .Update
            
            If sMenu = "Inscripcion" Then
                adoAlumnos.MoveLast
                id_Alumno = adoAlumnos!id
                frmInscripcionBis.txtAlumno.Text = txtNombre.Text
                Unload Me
            End If
        Else
            MsgBox MensajeValidacion, vbCritical, "ERROR"
            MensajeValidacion = ""
        End If
    End With

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - NUEVO_ALUMNO", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub MODIFICAR_ALUMNO()
    On Error GoTo ErrorHandle
    
    With adoAlumnos
        'Verifico que los datos sean correctos
        If VALIDAR_ALUMNO = True Then
            CERRAR_TABLA adoAlumnos
            sSql = "SELECT * FROM Alumnos WHERE id = " & id_Alumno
            adoAlumnos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic

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
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - MODIFICAR_ALUMNO", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub ELIMINAR_ALUMNO()
    On Error GoTo ErrorHandle
    
    With adoAlumnos
        CERRAR_TABLA adoAlumnos
        sSql = "SELECT * FROM Alumnos WHERE id = " & id_Alumno
        adoAlumnos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
         
        .Delete
        
        adoAlumnos.Close
        
        INICIALIZAR_PANTALLA
    End With

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - ELIMINAR_VENDEDOR", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Function VALIDAR_ALUMNO() As Boolean
    On Error GoTo ErrorHandle
    
    Dim HuboError As Boolean
    Dim errorCuitDni As String
    
    HuboError = False
    
    MensajeValidacion = ""
    
    With adoTablaValidacion
        'CUIT / DNI existente
        'Solo lo valido si es un alumno nuevo
        If id_Alumno = 0 Then
            sSql = "SELECT * FROM Alumnos WHERE NumDoc = '" & txtNumDoc.Text & "'"
            errorCuitDni = "DNI"
            .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MensajeValidacion = MensajeValidacion & vbCrLf & "- El " & errorCuitDni & " ya existe."
                HuboError = True
            End If
            .Close
        End If
        
        'Existe nombre
        If txtNombre.Text = "" Then
            MensajeValidacion = MensajeValidacion & vbCrLf & "- Debe ingresar el nombre."
            HuboError = True
        End If
        
        'Fecha de nacimiento válida
        'If Year(Date) - 10 < Year(dtpFechaNac.Value) Then
        '    MensajeValidacion = MensajeValidacion & vbCrLf & "- Fecha de nacimiento incorrecta."
        '    HuboError = True
        'End If
        
        'DNI distinto de vacío
        If txtNumDoc.Text = "" Then
            MensajeValidacion = MensajeValidacion & vbCrLf & "- Debe ingresar el número de documento."
            HuboError = True
        End If
    End With
    
    If HuboError Then
        VALIDAR_ALUMNO = False
        BOTONES "Nuevo", Me
    Else
        VALIDAR_ALUMNO = True
    End If

    Exit Function
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - VALIDAR_ALUMNO", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Function

Private Sub ASIGNAR_DATOS()
    On Error GoTo ErrorHandle
    
    'Asigno datos
    With adoAlumnos
        !Nombre = UCase(txtNombre.Text) & ""
        !Direccion = UCase(txtDireccion.Text) & ""
        !CodPostal = txtCodPostal.Text & ""
        !idLocalidad = DEVOLVER_ID(cboLocalidad.Text, adoLocalidades, "Localidades", "Detalle")
        !Telefono = txtTelefono.Text & ""
        !TelefonoLaboral = txtTelefonoLaboral.Text & ""
        
        celularSinGuiones = Replace(txtCelular.Text, "-", "")
        If celularSinGuiones = "" Then
            celularSinGuiones = "15xxxxxxxx"
        End If
        !Celular = celularSinGuiones & ""
        
        !idCompaniaCelular = DEVOLVER_ID(cboCompaniaCelular.Text, adoCompaniasCelular, "CompaniasCelular", "Detalle")
        !Mail = txtMail.Text & ""
        !idComoLlego = DEVOLVER_ID(cboComoLlego.Text, adoComoLlego, "ComoLlego", "Detalle")
        !FechaAlta = lblFechaAlta.Caption
        !FechaNac = dtpFechaNac.Value
        !idTipoDoc = DEVOLVER_ID(cboTipoDoc.Text, adoTiposDoc, "TiposDoc", "Detalle")
        !NumDoc = txtNumDoc.Text & ""
        !Observaciones = UCase(txtObservaciones.Text) & ""
        !idEmpresa = DEVOLVER_ID(cboEmpresa.Text, adoEmpresas, "Empresas", "Nombre")
    End With
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - ASIGNAR_DATOS", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub INICIALIZAR_PANTALLA()
    On Error GoTo ErrorHandle
    
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
    
    lblFechaAlta.Caption = Date
    
    imgFoto.Picture = LoadPicture(App.Path & "\fotos\sinfoto.jpg")

    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmAlumnos - INICIALIZAR_PANTALLA", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub


Private Sub txtCelular_Change()
    If Len(txtCelular.Text) = 2 Or Len(txtCelular.Text) = 7 Then
        txtCelular.Text = txtCelular.Text & "-"
        txtCelular.SelStart = Len(txtCelular.Text)
    End If
End Sub

Private Sub txtNumDoc_Change()
    If cboTipoDoc.Text <> "CUIT" Then
        If Len(txtNumDoc.Text) = 2 Or Len(txtNumDoc.Text) = 6 Then
            txtNumDoc.Text = txtNumDoc.Text & "."
            txtNumDoc.SelStart = Len(txtNumDoc.Text)
        End If
    Else
        If Len(txtNumDoc.Text) = 2 Or Len(txtNumDoc.Text) = 11 Then
            txtNumDoc.Text = txtNumDoc.Text & "-"
            txtNumDoc.SelStart = Len(txtNumDoc.Text)
        End If
    End If
End Sub

Private Sub txtTelefono_Change()
    If Len(txtTelefono.Text) = 4 Then
        txtTelefono.Text = txtTelefono.Text & "-"
        txtTelefono.SelStart = Len(txtTelefono.Text)
    End If
End Sub

Private Sub txtTelefonoLaboral_Change()
    If Len(txtTelefonoLaboral.Text) = 4 Then
        txtTelefonoLaboral.Text = txtTelefonoLaboral.Text & "-"
        txtTelefonoLaboral.SelStart = Len(txtTelefonoLaboral.Text)
    End If
End Sub
