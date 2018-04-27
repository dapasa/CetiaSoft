VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmcompara 
   AutoRedraw      =   -1  'True
   Caption         =   "Captura de Web Cam para enviar a Base  de Datos"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCompara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FormCompara.frx":08CA
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   735
   StartUpPosition =   2  'CenterScreen
   Begin KewlButtonz.KewlButtons Command2 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   6840
      WhatsThisHelpID =   1116
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cam.Opciones."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormCompara.frx":1231B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10500
      Top             =   1080
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   10560
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   4980
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   3570
      Left            =   360
      ScaleHeight     =   234
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   3570
      Left            =   5520
      ScaleHeight     =   234
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   4860
   End
   Begin MSComDlg.CommonDialog C_guardar 
      Left            =   10500
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin KewlButtonz.KewlButtons Salir1 
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   6840
      WhatsThisHelpID =   1118
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormCompara.frx":1231D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons PRENDER_CAMARA2 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   6840
      WhatsThisHelpID =   1113
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Encender Cam."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormCompara.frx":1231EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons PARAR_CAMARA2 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   6840
      WhatsThisHelpID =   1113
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Apagar Cam."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormCompara.frx":123208
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons KewlButtons1 
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   6840
      WhatsThisHelpID =   1113
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Capturar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   0
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormCompara.frx":123224
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   1275
      Left            =   6180
      TabIndex        =   8
      ToolTipText     =   "Presione aquí si entró en pánico."
      Top             =   1020
      WhatsThisHelpID =   1117
      Width           =   3435
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9900
      TabIndex        =   3
      Top             =   7620
      WhatsThisHelpID =   1114
      Width           =   75
   End
End
Attribute VB_Name = "frmcompara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Función Api
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054
'Variable para el hWnd de la ventana
Private hWndCap As Long
'Variable para contabilizar la cantidad de fotos capturadas
Dim fotos As Integer
'Variable para el directorio donde se guardaran las capturas
Dim directorio As String

Private Sub Form_Load()
On Error Resume Next
    'Seteamos controles en cuanto a su habilitación.
    habilitar
    'Seteamos el directorio path donde se ejecuta la aplicación.
    If Len(App.Path) = 3 Then
        directorio = App.Path
    Else
        directorio = App.Path & "\"
    End If
    'Verificamos si el directorio existe
    If Len(Dir(directorio & "WebCamPolice")) Then
        Else
            'Sino lo creamos.
            MkDir directorio & "WebCamPolice"
    End If
End Sub

    'Para la captura de la imagen hacemos clic
Private Sub KewlButtons1_Click()
    On Error Resume Next
    'Trasladamos la imagen del picture1 en estos momentos al picture2
    Picture2.Picture = Picture1.Picture
    'Aumentamos en uno el contador de fotos capturadas
    fotos = fotos + 1
    'Mostramos en el captión del formulario el resultado.
    Me.Caption = "       " & fotos & "   -----------     fotos capturadas.  ------  WebCam Control."
    'Llamamos al procedimiento para guardar la captura
    Call guardar_Click
End Sub

Private Sub PRENDER_CAMARA2_Click()
On Error Resume Next
    'Inhabilitamos controles que nos pueden hacer colgar el programa.
    inhabilitar
    'Mostramos los pictures donde se desarrollará la acción.
    Picture1.Visible = True
    Picture2.Visible = True
    'Habilitamos el temporizador que permite ver la imagen en movimiento en un pictube.
    Timer1.Enabled = True
    'función capCreateCaptureWindow, en el anteúltimo parámetro, se le envía el Hwnd de la ventana donde se capturará la webcam , por ejemplo un picture o formulario
    hWndCap = capCreateCaptureWindow("WebcamControl", 0, 0, 0, 0, 0, Me.hwnd, 0)
    'Permitir otros procesos y evitar que el programa se cuelgue
    DoEvents
    'Conectamos el dispositivo.
    SendMessage hWndCap, CONNECT, 0, 0
End Sub

Private Sub PARAR_CAMARA2_Click()
On Error Resume Next
    'Seteamos los controles para permitir salir del programa
    habilitar
    'Deshabilitamos el temporizador.
    Timer1.Enabled = False
    'Hacemos invisibles los picture
    Picture1.Visible = False
    Picture2.Visible = False
    'Los vaciamos de contenido.
    Picture1.Picture = Picture3.Picture
    Picture2.Picture = Picture1.Picture
    'Permitimos otros procesos y desconectamos el dispositivo.
    DoEvents: SendMessage hWndCap, DISCONNECT, 0, 0
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
    'Obtiene frames para Picture1
    SendMessage hWndCap, GET_FRAME, 0, 0
    'Lo copiamos al frame al Clipboard
    SendMessage hWndCap, COPY, 0, 0
    'Lo bajamos al Picture1 a ese frame
    Picture1.Picture = Clipboard.GetData
    'Vaciamos el clipboard
    Clipboard.Clear
End Sub

Private Sub guardar_Click()
    On Error Resume Next
    'Variable para la dirección y nombre del archivo a guardar
    Dim direccion As String
    'Determinamos con la fecha y hora el nombre del archivo a guardar
    fecha = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "-Hs-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time)
    'Finalmente creamos el resultado final
    direccion = directorio & "WebCamPolice\" & fecha & ".jpg"
    'Requerido por DIjpg.dll
    SavePicture Picture2.Image, "C:\tmp.bmp"
    'Grabamos en formato JPEG
    retval = DIWriteJpg(direccion, 75, 1)
    'Removemos el archivo temporario.
    Kill "C:\tmp.bmp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    DestroyWindow hWndCap
    End
End Sub

Private Sub Salir1_Click()
    DestroyWindow hWndCap
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If PRENDER_CAMARA2.Enabled = False Then
        MsgBox "No ha procedido a apagar la Web Cam. Este procedimiento no es válido en esta instancia!.", vbCritical, "Web Cam Control."
        Cancel = True
        Exit Sub
    End If
        DestroyWindow hWndCap
        End
End Sub

Private Sub Command2_Click()
capDlgVideoSource1 hWndCap
End Sub

Sub habilitar()
    PRENDER_CAMARA2.Enabled = True
    PARAR_CAMARA2.Enabled = False
    KewlButtons1.Enabled = False
    Command2.Enabled = False
    Salir1.Enabled = True
End Sub

Sub inhabilitar()
    PRENDER_CAMARA2.Enabled = False
    PARAR_CAMARA2.Enabled = True
    KewlButtons1.Enabled = True
    Command2.Enabled = False
    Salir1.Enabled = False
End Sub
