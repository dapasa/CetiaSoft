VERSION 5.00
Begin VB.Form frmCuotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPCIONES DE PAGO"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuotasACuenta 
      Caption         =   "Descontar cuotas (por abandono)"
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
      TabIndex        =   17
      Top             =   1680
      Width           =   5655
      Begin VB.CheckBox chkAbandono3 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   5415
      End
      Begin VB.CheckBox chkAbandono2 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   5415
      End
      Begin VB.CheckBox chkAbandono1 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   3720
      Picture         =   "frmCuotas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdClave 
      Caption         =   "Clave"
      Height          =   615
      Left            =   5160
      Picture         =   "frmCuotas.frx":0442
      TabIndex        =   6
      ToolTipText     =   "Aceptar"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   4440
      Picture         =   "frmCuotas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aceptar"
      Top             =   840
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "Matrícula"
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
      Top             =   840
      Width           =   1095
      Begin VB.TextBox txtValorMatricula 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Condición del alumno"
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
      TabIndex        =   12
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optNuevo 
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optEx 
         Caption         =   "Ex alumno"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de precios"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optAnterior 
         Caption         =   "Anterior"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optVigente 
         Caption         =   "Vigente"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Cuotas"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   855
      Begin VB.TextBox txtCantCuotas 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Valor cuota"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1215
      Begin VB.TextBox txtValorCuota 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5760
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    x_cuotas = Val(txtCantCuotas.Text)
    'x_valorCuota = Val(txtValorCuota.Text)
    'x_valorMatricula = Val(txtValorMatricula.Text)
    
    'INI - Chequeo si tiene cuotas pagas por abandono de otro curso.
    x_hay_cuotas_pagas = False
    
    If chkAbandono1.Value = vbChecked Then
        x_cuotas = x_cuotas - 1
        x_hay_cuotas_pagas = True
    End If
    
    If chkAbandono2.Value = vbChecked Then
        x_cuotas = x_cuotas - 1
        x_hay_cuotas_pagas = True
    End If
    
    If chkAbandono3.Value = vbChecked Then
        x_cuotas = x_cuotas - 1
        x_hay_cuotas_pagas = True
    End If
    
    
    'FIN - Chequeo si tiene cuotas pagas por abandono de otro curso.
    
    If optVigente.Value Then
        x_valorCuotaReal = adoPrecios!Vigente
        x_valorMatriculaReal = adoPrecios!Matricula_Vigente
    Else
        x_valorCuotaReal = adoPrecios!Anterior
        x_valorMatriculaReal = adoPrecios!Matricula_Anterior
    End If
        
    If optEx Then
        x_DetalleDescuento = "EX ALUMNO"
    Else
        x_DetalleDescuento = ""
    End If
        
    adoPrecios.Close
    
    GUARDAR_LOG x_usuario, Date, Time, "MODIFICA VALOR CUOTA DEL ALUMNO CON ID " & id_Alumno & " VALOR MATRÍCULA: $" & txtValorMatricula.Text & ".- VALOR CUOTA: $" & txtValorCuota.Text & ".-"
    
    Unload Me
End Sub


Private Sub cmdCancelar_Click()
    CancelaInscripcion = True
    Unload Me
End Sub

Private Sub cmdClave_Click()
    sMenu = "Inscripcion"
    
    AccesoPublico = True
        frmClave.Show vbModal
    AccesoPublico = False

    If Not AccesoPermitido Then
        MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
        Exit Sub
    End If
       
    txtValorMatricula.Locked = False
    txtValorCuota.Locked = False
End Sub

Private Sub Form_Load()
    CancelaInscripcion = False
    CERRAR_TABLA adoPrecios
    adoPrecios.Open "Precios", adoConnection, adOpenDynamic, adLockOptimistic
    txtValorMatricula.Text = adoPrecios!Matricula_Vigente
    txtValorCuota.Text = adoPrecios!Vigente
    txtCantCuotas.Text = Val(Left(DEVOLVER_CAMPO(adoCursos!idDuracion, adoDuraciones, "Duraciones", "Detalle"), 2))
    
    'Verifico si tiene cuotas a cuenta por abandonos.
    
    hay_cuotas = False
    
    If x_abandono_1 <> "" Then
        chkAbandono1.Caption = x_abandono_1
        hay_cuotas = True
    Else
        chkAbandono1.Visible = False
    End If
    
    If x_abandono_2 <> "" Then
        chkAbandono2.Caption = x_abandono_2
        hay_cuotas = True
    Else
        chkAbandono2.Visible = False
    End If
    
    If x_abandono_3 <> "" Then
        chkAbandono3.Caption = x_abandono_3
        hay_cuotas = True
    Else
        chkAbandono3.Visible = False
    End If
    
    If hay_cuotas Then
        Me.Height = 3690
        fraCuotasACuenta.Visible = True
    Else
        Me.Height = 1995
        fraCuotasACuenta.Visible = False
    End If
End Sub

Private Sub optAnterior_Click()
    If optNuevo.Value = True Then
        'por 0,9
        txtValorMatricula.Text = adoPrecios!Matricula_Anterior
        txtValorCuota.Text = adoPrecios!Anterior
    Else
        'por 0,8
        txtValorMatricula.Text = Round(adoPrecios!Matricula_Anterior * 0.85, 0)
        txtValorCuota.Text = Round(adoPrecios!Anterior * 0.85, 0)
    End If
End Sub

Private Sub optEx_Click()
    'por 0,8
    x_es_ex_alumno = True
    
    If optVigente.Value = True Then
        txtValorMatricula.Text = Round(adoPrecios!Matricula_Vigente * 0.85, 0)
        txtValorCuota.Text = Round(adoPrecios!Vigente * 0.85, 0)
        x_valorMatricula = adoPrecios!Matricula_Vigente
        x_valorCuota = adoPrecios!Vigente
    Else
        txtValorMatricula.Text = Round(adoPrecios!Matricula_Anterior * 0.85, 0)
        txtValorCuota.Text = Round(adoPrecios!Anterior * 0.85, 0)
        x_valorMatricula = adoPrecios!Matricula_Anterior
        x_valorCuota = adoPrecios!Anterior
        
    End If
    
End Sub

Private Sub optNuevo_Click()
    'por 0,9
    x_es_ex_alumno = False
    
    If optVigente.Value = True Then
        txtValorMatricula.Text = adoPrecios!Matricula_Vigente
        txtValorCuota.Text = adoPrecios!Vigente
        x_valorMatricula = adoPrecios!Matricula_Vigente
        x_valorCuota = adoPrecios!Vigente
    Else
        txtValorMatricula.Text = adoPrecios!Matricula_Anterior
        txtValorCuota.Text = adoPrecios!Anterior
        x_valorMatricula = adoPrecios!Matricula_Anterior
        x_valorCuota = adoPrecios!Anterior
    End If
End Sub

Private Sub optVigente_Click()
    If optNuevo.Value = True Then
        'por 0,9
        txtValorMatricula.Text = adoPrecios!Matricula_Vigente
        txtValorCuota.Text = adoPrecios!Vigente
    Else
        'por 0,8
        txtValorMatricula.Text = Round(adoPrecios!Matricula_Vigente * 0.85, 0)
        txtValorCuota.Text = Round(adoPrecios!Vigente * 0.85, 0)
    End If
End Sub

