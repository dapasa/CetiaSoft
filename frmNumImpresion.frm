VERSION 5.00
Begin VB.Form frmNumImpresion 
   BackColor       =   &H000080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblImporte 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Importe $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   60
      TabIndex        =   4
      Top             =   1440
      Width           =   8295
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   495
      Left            =   4320
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblCancelar 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2235
      Width           =   2175
   End
   Begin VB.Label lblContinuar 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "CONTINUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2235
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   495
      Left            =   1560
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   3000
      Left            =   0
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label lblNumImpresion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "0000-00000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   8460
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Imprimiendo el comprobante Nº"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmNumImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        lblContinuar_Click
    End If
End Sub

Private Sub lblCancelar_Click()
    x_imprimir = False
    NUMERO_ATRAS x_comprobante
    
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos ORDER BY id DESC"
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    num_id = adoMovimientos!id
    adoMovimientos.Delete
    adoMovimientos.Close
    
    sSql = "DELETE FROM itemsXMov WHERE idMovimiento = " & num_id
    adoConnection.Execute sSql
    
    Unload Me
End Sub

Private Sub lblContinuar_Click()
    x_imprimir = True
    Unload Me
End Sub
