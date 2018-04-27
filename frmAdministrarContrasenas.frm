VERSION 5.00
Begin VB.Form frmAdministrarContrasenas 
   Caption         =   "Administrar contraseñas"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3465
   Icon            =   "frmAdministrarContrasenas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardarContrasena 
      Height          =   615
      Left            =   2760
      Picture         =   "frmAdministrarContrasenas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Guardar"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdEliminarUsuario 
      Height          =   615
      Left            =   2760
      Picture         =   "frmAdministrarContrasenas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar"
      Top             =   960
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Contraseña"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
      Begin VB.TextBox txtContrasena 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nuevo usuario"
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
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuario"
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
      Top             =   960
      Width           =   2535
      Begin VB.ComboBox cboUsuarios 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "cboUsuarios"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "Situación"
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
      Left            =   7440
      TabIndex        =   1
      Top             =   -120
      Width           =   3015
      Begin VB.ComboBox cboSituacion 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdAgregarUsuario 
      Height          =   615
      Left            =   2760
      Picture         =   "frmAdministrarContrasenas.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Agregar clase"
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   3240
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmAdministrarContrasenas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarUsuario_Click()
    If MsgBox("¿Confirma que desea agregar el usuario " & txtUsuario.Text & "?", vbQuestion + vbYesNo, "AGREGAR USUARIO") = vbYes Then
    
        With adoUsuarios
            CERRAR_TABLA adoUsuarios
            
            adoUsuarios.Open "Usuarios", adoConnection, adOpenDynamic, adLockOptimistic
            
            adoUsuarios.AddNew
            adoUsuarios!Usuario = txtUsuario.Text
            adoUsuarios.Update
            
            adoUsuarios.Close
        End With
        
        txtUsuario.Text = ""
        CARGAR_COMBO "cboUsuarios", adoUsuarios, "Usuarios", "Usuario", Me
    End If
End Sub

Private Sub cmdEliminarUsuario_Click()
    If MsgBox("¿Confirma que desea eliminar el usuario " & cboUsuarios.Text & "?", vbQuestion + vbYesNo, "ELIMINAR USUARIO") = vbYes Then
        sSql = "DELETE FROM Usuarios WHERE Usuario = '" & cboUsuarios.Text & "'"
        adoConnection.Execute sSql
        
        CARGAR_COMBO "cboUsuarios", adoUsuarios, "Usuarios", "Usuario", Me
    Else
        cboUsuarios.ListIndex = 0
    End If
End Sub

Private Sub cmdGuardarContrasena_Click()
    If MsgBox("¿Confirma que desea asignar una nueva contraseña al usuario " & cboUsuarios.Text & "?", vbQuestion + vbYesNo, "ASIGNAR CONTRASEÑA") = vbYes Then
        sSql = "UPDATE Usuarios SET contrasena = '" & txtContrasena.Text & "' WHERE Usuario = '" & cboUsuarios.Text & "'"
        adoConnection.Execute sSql
    End If
End Sub

Private Sub Form_Load()
    CARGAR_COMBO "cboUsuarios", adoUsuarios, "Usuarios", "Usuario", Me
End Sub
