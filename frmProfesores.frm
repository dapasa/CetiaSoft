VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmProfesores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROFESORES"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   Icon            =   "frmProfesores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   1560
      Picture         =   "frmProfesores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Buscar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5400
      Picture         =   "frmProfesores.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancelar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   4440
      Picture         =   "frmProfesores.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Guardar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2520
      Picture         =   "frmProfesores.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   600
      Picture         =   "frmProfesores.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdListado 
      Caption         =   "&Listado"
      Height          =   855
      Left            =   3480
      Picture         =   "frmProfesores.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Lista de alumnos"
      Top             =   4320
      Width           =   855
   End
   Begin Crystal.CrystalReport rptInforme 
      Left            =   120
      Top             =   4440
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
   Begin VB.Frame Frame4 
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
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   6135
      Begin VB.TextBox txtObservaciones 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   15
      Top             =   840
      Width           =   6135
      Begin VB.ComboBox cboCompaniaCelular 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtMail 
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtCelular 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cía."
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   780
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "e-Mail"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   405
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   13
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   6240
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmProfesores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' MÓDULO: Mantenimiento de profesores FECHA: Ago / 2007
'******************************************************
' RESUMEN:
'******************************************************
' ÚLTIMA MODIFICACIÓN IMPORTANTE: 09/08/2007
'******************************************************
' ETAPA: release candidate.
'******************************************************
' AUTOR: Pablo Adrián Langholz
' CONTACTO: elmaildepablo@gmail.com
'******************************************************

Dim ModoABM As String

Private Sub cmdBuscar_Click()
    On Error GoTo ErrorHandle
    
    BOTONES "Buscar", Me

    ENABLED_TODO True, Me

    EstiloBuscador = "Profesores"
    frmBuscador.Show vbModal
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - cmdBuscar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - cmdCancelar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo ErrorHandle
    BOTONES "Eliminar", Me
    
    If MsgBox("¿Confirma que desea eliminar al profesor " & txtNombre.Text & "?", vbYesNo + vbQuestion, "Profesores") = vbYes Then
        frmClave.Show vbModal
        If Not AccesoPermitido Then
            MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
            Exit Sub
        End If
        
        If PUEDE_BORRAR("Profesores", id_Profesor) Then
            ModoABM = "B"
            'Eliminar
            ELIMINAR_PROFESOR
        End If
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - cmdEliminar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo ErrorHandle
    BOTONES "Guardar", Me
    
    If MsgBox("¿Confirma que desea guardar los cambios?", vbYesNo + vbQuestion, "PROFESORES") = vbYes Then
        'Guardar
        If id_Profesor = 0 Then 'Profesor nuevo
            ModoABM = "A"
            NUEVO_PROFESOR
            ModoABM = ""
        Else 'Modificación de profesor
            frmClave.Show vbModal
            If Not AccesoPermitido Then
                MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
                Exit Sub
            End If
        
            With adoProfesores
                ModoABM = "M"
                MODIFICAR_PROFESOR
            End With
        End If
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - cmdGuardar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdListado_Click()
    rptInforme.ReportFileName = App.Path & "\reportes\rptProfesores.rpt"
    rptInforme.ReportTitle = "LISTADO DE PROFESORES"
    rptInforme.Action = 1
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo ErrorHandle
    BOTONES "Nuevo", Me
    
    id_Profesor = 0
    
    ENABLED_TODO True, Me
    
    txtNombre.SetFocus
    
    INICIALIZAR_PANTALLA

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - cmdNuevo", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If KeyAscii = 13 Then
        PASAR_CAMPO Me
    End If
    
    If KeyAscii >= 97 And KeyAscii <= 122 And Me.ActiveControl.Name <> "txtMail" Then
        KeyAscii = KeyAscii - 32
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - Form.KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Load()
    'On Error GoTo ErrorHandle
    
    id_Profesor = 0
    
    ENABLED_TODO False, Me
    
    CERRAR_TABLA adoProfesores
    sSql = "SELECT * FROM Profesores"
    adoProfesores.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If adoProfesores.EOF Then
        cmdBuscar.Enabled = False
    End If
    
    CARGAR_COMBO "cboCompaniaCelular", adoCompaniasCelular, "CompaniasCelular", "Detalle", Me
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - Form.Load", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    CERRAR_TABLA adoProfesores

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - Form.Unload", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub NUEVO_PROFESOR()
    On Error GoTo ErrorHandle
    
    With adoProfesores
        'Verifico que los datos sean correctos
        If VALIDAR_PROFESOR = True Then
            'Agrego el registro
            .AddNew
            
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
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - NUEVO_PROFESOR", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub MODIFICAR_PROFESOR()
    On Error GoTo ErrorHandle
    
    With adoProfesores
        'Verifico que los datos sean correctos
        If VALIDAR_PROFESOR = True Then
            CERRAR_TABLA adoProfesores
            sSql = "SELECT * FROM Profesores WHERE id = " & id_Profesor
            adoProfesores.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic

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
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - MODIFICAR_PROFESOR", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub ELIMINAR_PROFESOR()
    On Error GoTo ErrorHandle
    
    With adoProfesores
        CERRAR_TABLA adoProfesores
        sSql = "SELECT * FROM Profesores WHERE id = " & id_Profesor
        adoProfesores.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
         
        .Delete
        
        adoProfesores.Close
        
        INICIALIZAR_PANTALLA
    End With

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - ELIMINAR_PROFESOR", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Function VALIDAR_PROFESOR() As Boolean
    On Error GoTo ErrorHandle
    
    Dim HuboError As Boolean
    HuboError = False
    
    MensajeValidacion = ""
    
    With adoTablaValidacion
        'Nombre existente
        'Solo lo valido si es un profesor nuevo
        If id_Profesor = 0 Then
            sSql = "SELECT * FROM Profesores WHERE Nombre = '" & txtNombre.Text & "'"
            .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MensajeValidacion = MensajeValidacion & vbCrLf & "- El nombre ya existe."
                HuboError = True
            End If
            .Close
        End If
        
        'Existe nombre
        If txtNombre.Text = "" Then
            MensajeValidacion = MensajeValidacion & vbCrLf & "- Debe ingresar el nombre."
            HuboError = True
        End If
    End With
    
    If HuboError Then
        VALIDAR_PROFESOR = False
    Else
        VALIDAR_PROFESOR = True
    End If

    Exit Function
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - VALIDAR_PROFESOR", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Function

Private Sub ASIGNAR_DATOS()
    On Error GoTo ErrorHandle
    
    'Asigno datos
    With adoProfesores
        !Nombre = UCase(txtNombre.Text)
        !Direccion = UCase(txtDireccion.Text)
        !Telefono = txtTelefono.Text
        !Celular = txtCelular.Text
        !idCompaniaCelular = DEVOLVER_ID(cboCompaniaCelular.Text, adoCompaniasCelular, "CompaniasCelular", "Detalle")
        !Mail = txtMail.Text
        !Observaciones = UCase(txtObservaciones.Text)
    End With
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - ASIGNAR_DATOS", vbCritical, "SE HA PRODUCIDO UN ERROR"
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
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmProfesores - INICIALIZAR_PANTALLA", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub txtCelular_Change()
    If Len(txtCelular.Text) = 2 Or Len(txtCelular.Text) = 7 Then
        txtCelular.Text = txtCelular.Text & "-"
        txtCelular.SelStart = Len(txtCelular.Text)
    End If
End Sub

Private Sub txtTelefono_Change()
    If Len(txtTelefono.Text) = 4 Then
        txtTelefono.Text = txtTelefono.Text & "-"
        txtTelefono.SelStart = Len(txtTelefono.Text)
    End If
End Sub
