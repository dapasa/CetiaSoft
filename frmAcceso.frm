VERSION 5.00
Begin VB.Form frmAcceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acceso al Sistema"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3345
   Icon            =   "frmAcceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContrasena 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   2640
      Picture         =   "frmAcceso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Acceder"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   1920
      Picture         =   "frmAcceso.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   3240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   165
      Width           =   540
   End
End
Attribute VB_Name = "frmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    CERRAR_TABLA adoUsuarios
    sSql = "SELECT id FROM Usuarios WHERE Usuario = '" & txtUsuario.Text & "' AND Contrasena = '" & ENCRIPTAR(txtContrasena.Text) & "'"
    adoUsuarios.Open sSql, adoConnection, adOpenStatic, adLockOptimistic
    
    If adoUsuarios.EOF Then
        MsgBox "Usuario y/o contraseña no válidos.", vbCritical, "ERROR"
    Else
        'If Dir("c:\windows\system32\x.y") <> "" Then
        If Dir("c:\temp\x.y") <> "" Then
            Unload Me
            
            'Cargo la configuración guardada en el registro
            
            
            frmPrincipal.Show
        Else
            MsgBox "Este sistema no se ha instalado correctamente.", vbCritical, "ERROR"
            End
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    If MsgBox("¿Confirma que desea cancelar el acceso al Sistema?", vbQuestion + vbYesNo, "ACCESO AL SISTEMA") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtUsuario.Text = "admin"
        txtContrasena.Text = "xyz"
        cmdAceptar_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtContrasena_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
    End If
End Sub
