VERSION 5.00
Begin VB.Form frmSacarFoto 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton cmdSacarFoto 
      Caption         =   "Sacar foto"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSacarFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaraciones:
Const ws_visible = &H10000000
Const ws_child = &H40000000
Const WM_USER = 1024
Const WM_CAP_EDIT_COPY = WM_USER + 30
Const wm_cap_driver_connect = WM_USER + 10
Const wm_cap_set_preview = WM_USER + 50
Const wm_cap_set_overlay = WM_USER + 51
Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Const WM_CAP_SEQUENCE = WM_USER + 62
Const WM_CAP_SINGLE_FRAME_OPEN = WM_USER + 70
Const WM_CAP_SINGLE_FRAME_CLOSE = WM_USER + 71
Const WM_CAP_SINGLE_FRAME = WM_USER + 72
Const DRV_USER = &H4000
Const DVM_DIALOG = DRV_USER + 100
Const PREVIEWRATE = 30

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal a As String, ByVal b As Long, ByVal c As Integer, ByVal d As Integer, ByVal e As Integer, ByVal f As Integer, ByVal g As Long, ByVal h As Integer) As Long

Dim hwndc As Long

Private Sub cmdSacarFoto_Click()
    cambio_foto = True
    
    'Código que realiza la captura de la imagen:
    
    temp = SendMessage(hwndc, WM_CAP_EDIT_COPY, 1, 0)
    Set Picture1.Picture = Clipboard.GetData
    SavePicture Picture1.Picture, App.Path & "\fotos\imagen.bmp"
    
    Unload Me
End Sub

Private Sub Form_Load()
'Código que activa la captura de imágenesse supone un formulario con 2 picture llamados "picture1" y "picture2")

hwndc = capCreateCaptureWindow("Ventana de Captura", ws_child Or ws_visible, 0, 0, Picture2.Width, Picture2.Height, Picture2.hwnd, 0)
If (hwndc <> 0) Then
temp = SendMessage(hwndc, wm_cap_driver_connect, 0, 0)
temp = SendMessage(hwndc, wm_cap_set_preview, 1, 0)
temp = SendMessage(hwndc, WM_CAP_SET_PREVIEWRATE, PREVIEWRATE, 0)
End If

End Sub
