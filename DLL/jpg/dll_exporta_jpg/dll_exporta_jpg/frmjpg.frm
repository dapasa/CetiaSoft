VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmjpg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   7695
      Begin VB.HScrollBar hScroll 
         Height          =   255
         LargeChange     =   5
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   8
         Top             =   840
         Value           =   100
         Width           =   3375
      End
      Begin VB.CheckBox chkProg 
         Caption         =   "Progresivo"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar a Jpg"
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblcalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblQ 
         BackStyle       =   0  'Transparent
         Caption         =   "Calidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vista de las imágenes"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "Seleccionar imagen"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   1935
      End
      Begin VB.PictureBox PicOriginal 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1980
         Left            =   240
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   3
         Top             =   600
         Width           =   1980
      End
      Begin VB.PictureBox PicResultado 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1980
         Left            =   5280
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   1
         Top             =   600
         Width           =   1980
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   2880
      End
      Begin VB.Label lblSrc 
         BackStyle       =   0  'Transparent
         Caption         =   "Original"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblJpg 
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmjpg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long

Private Sub Command1_Click()
'Valor de retorno de la dll
Dim ret As Long



If PicOriginal.Picture = 0 Then MsgBox "No hay imagen para exportar. Seleccione una", vbExclamation: Exit Sub

'Guardamos en disco un temporal
SavePicture PicOriginal.Image, "c:\tmp.bmp"
'Ejecutamos la función de la dll que exporta la imagen en el app de la aplicacion
ret = DIWriteJpg(App.Path & "\resultado.jpg", hScroll.Value, chkProg.Value)

'Si devuelve un 1 esta todo Ok, otro numero es un error
If ret = 1 Then  'Success
    PicResultado.Picture = LoadPicture(App.Path & "\resultado.jpg")
    MsgBox "se exportó el archivo en el App de la aplicacion", vbInformation + vbOKOnly
Else
    MsgBox "No se pudo exportar"
End If
'Eliminamos el archivo temporal
Kill "c:\tmp.bmp"


End Sub

Private Sub Command2_Click()
On Error GoTo ehandle
CommonDialog1.ShowOpen 'Abrimos
If CommonDialog1.FileName = "" Then Exit Sub
PicOriginal.Picture = LoadPicture(CommonDialog1.FileName) 'Cargamos el Picture

Exit Sub

ehandle:
If Err.Number = 481 Then MsgBox "EL archivo que eligió no es una imagen válida": Exit Sub

End Sub

Private Sub hScroll_Change()
hScroll_Scroll
End Sub

Private Sub hScroll_Scroll()
lblcalidad = CByte(hScroll.Value)
End Sub
