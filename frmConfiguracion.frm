VERSION 5.00
Begin VB.Form frmConfiguracion 
   Caption         =   "Configuración"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmConfiguracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   3960
      Picture         =   "frmConfiguracion.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   3240
      Picture         =   "frmConfiguracion.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impresión de comprobantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtCopias 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Copias"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    SaveSetting "Gestion", "Documentos", "cantCopias", txtCopias.Text
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtCopias.Text = GetSetting("Gestion", "Documentos", "cantCopias")
End Sub
