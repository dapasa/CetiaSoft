VERSION 5.00
Begin VB.Form frmClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACCESO RESTRINGIDO"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3345
   Icon            =   "frmClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   2640
      Picture         =   "frmClave.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   1920
      Picture         =   "frmClave.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aceptar"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If AccesoPublico Then
        CERRAR_TABLA adoTemp
        sSql = "SELECT * FROM Usuarios WHERE contrasena = '" & txtClave.Text & "'"
        adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If Not adoTemp.EOF Then
            x_usuario = adoTemp!usuario
            
            adoTemp.Close
            
            AccesoPermitido = True
            Unload Me
            Exit Sub
        Else
            x_usuario = ""
        
            AccesoPermitido = False
            MsgBox "La clave no es correcta." & vbCrLf & "Vuelva a intentar.", vbCritical, "ACCESO DENEGADO"
        End If
    Else
        If txtClave.Text = "FiatIdea" Then
                AccesoPermitido = True
                Unload Me
        Else
            AccesoPermitido = False
            MsgBox "La clave no es correcta." & vbCrLf & "Vuelva a intentar.", vbCritical, "ACCESO DENEGADO"
        End If
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    AccesoPermitido = False
    
    Unload Me
End Sub

Private Sub Form_Load()
    AccesoPermitido = False
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
    End If
End Sub
