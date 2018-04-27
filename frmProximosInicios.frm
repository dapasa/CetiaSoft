VERSION 5.00
Begin VB.Form frmProximosInicios 
   Caption         =   "Próximos inicios"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Height          =   615
      Left            =   3240
      Picture         =   "frmProximosInicios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Aceptar"
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   3960
      Picture         =   "frmProximosInicios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancelar"
      Top             =   4080
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtProximoCurso 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmProximosInicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    With adoProximosInicios
        .Open "ProximosInicios", adoConnection, adOpenDynamic, adLockOptimistic
        For k = 1 To 10
            campo = "proximo" & k
            .Fields(campo) = txtProximoCurso(k).Text & ""
            .Update
        Next
        .Close
    End With
End Sub

Private Sub Form_Load()
    With adoProximosInicios
        .Open "ProximosInicios", adoConnection, adOpenDynamic, adLockOptimistic
        For k = 1 To 10
            campo = "proximo" & k
            txtProximoCurso(k).Text = .Fields(campo) & ""
        Next
        .Close
    End With
End Sub

'Esto se imprime en todo Sergio y todo Néstor y B de Fabru
