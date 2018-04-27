VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMPRESAS"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   Icon            =   "frmEmpresas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdListado 
      Caption         =   "&Listado"
      Height          =   855
      Left            =   6360
      Picture         =   "frmEmpresas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Lista de alumnos"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   3480
      Picture         =   "frmEmpresas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nuevo"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5400
      Picture         =   "frmEmpresas.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Eliminar"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   7320
      Picture         =   "frmEmpresas.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Guardar"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   8280
      Picture         =   "frmEmpresas.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancelar"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   4440
      Picture         =   "frmEmpresas.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Buscar"
      Top             =   4680
      Width           =   855
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
      TabIndex        =   28
      Top             =   3120
      Width           =   6495
      Begin VB.TextBox txtObservaciones 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
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
      Left            =   5160
      TabIndex        =   27
      Top             =   840
      Width           =   1455
      Begin VB.TextBox txtCodPostal 
         Height          =   285
         Left            =   120
         TabIndex        =   4
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
      TabIndex        =   26
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2535
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
      TabIndex        =   25
      Top             =   840
      Width           =   4935
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   3
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
      TabIndex        =   20
      Top             =   1560
      Width           =   6495
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtCelular 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtMail 
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1080
         Width           =   5295
      End
      Begin VB.ComboBox cboCompaniaCelular 
         Height          =   315
         Left            =   3600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   405
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "e-Mail"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cía."
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   780
         Width           =   300
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "CUIT"
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
      TabIndex        =   19
      Top             =   120
      Width           =   1455
      Begin VB.TextBox txtCuit 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
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
      Left            =   6720
      TabIndex        =   18
      Top             =   840
      Width           =   2415
      Begin VB.ComboBox cboLocalidad 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Condición de IVA"
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
      TabIndex        =   17
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox cboCondIva 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin Crystal.CrystalReport rptInforme 
      Left            =   120
      Top             =   4800
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
      X2              =   9120
      Y1              =   4560
      Y2              =   4560
   End
End
Attribute VB_Name = "frmEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
' MÓDULO: Mantenimiento de Empresas    FECHA: Nov / 2008
'******************************************************
' RESUMEN:
'******************************************************
' ÚLTIMA MODIFICACIÓN IMPORTANTE: 08/11/2008
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
    
    txtCuit.Locked = True
    
    EstiloBuscador = "Empresas"
    frmBuscador.Show vbModal
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - cmdBuscar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - cmdCancelar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo ErrorHandle
    
    BOTONES "Eliminar", Me
    
    If MsgBox("¿Confirma que desea eliminar la empresa " & txtNombre.Text & "?", vbYesNo + vbQuestion, "Empresas") = vbYes Then
        frmClave.Show vbModal
        If Not AccesoPermitido Then
            MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
            Exit Sub
        End If
        
        If PUEDE_BORRAR("Empresas", id_Empresa) Then
            ModoABM = "B"
            'Eliminar
            ELIMINAR_EMPRESA
        End If
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - cmdEliminar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo ErrorHandle
    
    If MsgBox("¿Confirma que desea guardar los cambios?", vbYesNo + vbQuestion, "Empresas") = vbYes Then
        BOTONES "Guardar", Me
        
        'Guardar
        If id_Empresa = 0 Then 'empresa nuevo
            ModoABM = "A"
            NUEVA_EMPRESA
            ModoABM = ""
        Else 'Modificación de empresa
        
            AccesoPublico = True
                frmClave.Show vbModal
            AccesoPublico = False
            
            If Not AccesoPermitido Then
                MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
                Exit Sub
            End If
            
            With adoEmpresas
                ModoABM = "M"
                MODIFICAR_EMPRESA
            End With
            
            GUARDAR_LOG x_usuario, Date, Time, "MODIFICA EMPRESA - " & txtNombre.Text
        End If
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - cmdGuardar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdListado_Click()
    On Error GoTo ErrorHandle
    
    rptInforme.ReportFileName = App.Path & "\reportes\rptEmpresas.rpt"
    rptInforme.ReportTitle = "LISTADO DE Empresas"
    rptInforme.Action = 1

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - cmdListado", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo ErrorHandle
    
    BOTONES "Nuevo", Me
        
    id_Empresa = 0
    
    ENABLED_TODO True, Me
        
    txtCuit.Locked = False
    txtNombre.SetFocus
    
    INICIALIZAR_PANTALLA

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - cmdNuevo", vbCritical, "SE HA PRODUCIDO UN ERROR"
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
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - Form.KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    id_Empresa = 0
    
    ENABLED_TODO False, Me
        
    CERRAR_TABLA adoEmpresas
    sSql = "SELECT * FROM Empresas"
    adoEmpresas.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    If adoEmpresas.EOF Then
        cmdBuscar.Enabled = False
    End If
    
    CARGAR_COMBO "cboCompaniaCelular", adoCompaniasCelular, "CompaniasCelular", "Detalle", Me
    CARGAR_COMBO "cboCondIva", adoCondIva, "CondIva", "Detalle", Me
    CARGAR_COMBO "cboLocalidad", adoLocalidades, "Localidades", "Detalle", Me
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - Form.Load", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    CERRAR_TABLA adoEmpresas

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - Form.Unload", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub NUEVA_EMPRESA()
    On Error GoTo ErrorHandle
    
    With adoEmpresas
        'Verifico que los datos sean correctos
        If VALIDAR_EMPRESA = True Then
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
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - NUEVA_EMPRESA", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub MODIFICAR_EMPRESA()
    On Error GoTo ErrorHandle
    
    With adoEmpresas
        'Verifico que los datos sean correctos
        If VALIDAR_EMPRESA = True Then
            CERRAR_TABLA adoEmpresas
            sSql = "SELECT * FROM Empresas WHERE id = " & id_Empresa
            adoEmpresas.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic

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
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - MODIFICAR_empresa", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub ELIMINAR_EMPRESA()
    On Error GoTo ErrorHandle
    
    With adoEmpresas
        CERRAR_TABLA adoEmpresas
        sSql = "SELECT * FROM Empresas WHERE id = " & id_Empresa
        adoEmpresas.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
         
        .Delete
        
        adoEmpresas.Close
        
        INICIALIZAR_PANTALLA
    End With

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - ELIMINAR_VENDEDOR", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Function VALIDAR_EMPRESA() As Boolean
    On Error GoTo ErrorHandle
    
    Dim HuboError As Boolean
    Dim errorCuitDni As String
    
    HuboError = False
    
    MensajeValidacion = ""
    
    With adoTablaValidacion
        'CUIT
        'If txtCuit.Text = "" Then
        '    MensajeValidacion = MensajeValidacion & vbCrLf & "- Debe ingresar el número de CUIT."
        'End If
        
        If id_Empresa = 0 Then
            sSql = "SELECT * FROM Empresas WHERE Cuit = '" & txtCuit.Text & "'"
            errorCuitDni = "CUIT"
        
            .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            If Not .EOF Then
                MensajeValidacion = MensajeValidacion & vbCrLf & "- El " & errorCuitDni & " ya existe."
                HuboError = True
            End If
            .Close
            
            'Existe nombre
            If txtNombre.Text = "" Then
                MensajeValidacion = MensajeValidacion & vbCrLf & "- Debe ingresar el nombre."
                HuboError = True
            End If
        
            'IVA
            If cboCondIva.Text = "(No disponible)" Then
                MensajeValidacion = MensajeValidacion & vbCrLf & "- Debe seleccionar una Condición de IVA."
                HuboError = True
            End If
        End If
    End With
    
    If HuboError Then
        VALIDAR_EMPRESA = False
        BOTONES "Nuevo", Me
    Else
        VALIDAR_EMPRESA = True
    End If

    Exit Function
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - VALIDAR_empresa", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Function

Private Sub ASIGNAR_DATOS()
    On Error GoTo ErrorHandle
    
    'Asigno datos
    With adoEmpresas
        !Nombre = UCase(txtNombre.Text)
        !Cuit = txtCuit.Text
        !idCondIva = DEVOLVER_ID(cboCondIva.Text, adoCondIva, "CondIva", "Detalle")
        !Direccion = UCase(txtDireccion.Text)
        !CodPostal = txtCodPostal.Text
        !idLocalidad = DEVOLVER_ID(cboLocalidad.Text, adoLocalidades, "Localidades", "Detalle")
        !Telefono = txtTelefono.Text
        !Celular = txtCelular.Text
        !idCompaniaCelular = DEVOLVER_ID(cboCompaniaCelular.Text, adoCompaniasCelular, "CompaniasCelular", "Detalle")
        !Mail = txtMail.Text
        !Observaciones = UCase(txtObservaciones.Text)
    End With
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - ASIGNAR_DATOS", vbCritical, "SE HA PRODUCIDO UN ERROR"
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
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmEmpresas - INICIALIZAR_PANTALLA", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub


Private Sub txtCelular_Change()
    If Len(txtCelular.Text) = 2 Or Len(txtCelular.Text) = 7 Then
        txtCelular.Text = txtCelular.Text & "-"
        txtCelular.SelStart = Len(txtCelular.Text)
    End If
End Sub

Private Sub txtCuit_Change()
    If Len(txtCuit.Text) = 2 Or Len(txtCuit.Text) = 11 Then
        txtCuit.Text = txtCuit.Text & "-"
        txtCuit.SelStart = Len(txtCuit.Text)
    End If
End Sub

Private Sub txtTelefono_Change()
    If Len(txtTelefono.Text) = 4 Then
        txtTelefono.Text = txtTelefono.Text & "-"
        txtTelefono.SelStart = Len(txtTelefono.Text)
    End If
End Sub
