VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl Calendario 
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   ScaleHeight     =   2940
   ScaleWidth      =   3450
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   2160
      TabIndex        =   44
      Top             =   120
      Width           =   975
   End
   Begin MSComCtl2.UpDown updAnio 
      Height          =   325
      Left            =   3120
      TabIndex        =   43
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   582
      _Version        =   393216
      Value           =   2007
      Max             =   2099
      Min             =   2007
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox cboMes 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "S"
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
      Left            =   3000
      TabIndex        =   51
      Top             =   570
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "V"
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
      Left            =   2520
      TabIndex        =   50
      Top             =   570
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "J"
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
      Left            =   2040
      TabIndex        =   49
      Top             =   570
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "M"
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
      Left            =   1560
      TabIndex        =   48
      Top             =   570
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "M"
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
      Left            =   1080
      TabIndex        =   47
      Top             =   570
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "L"
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
      Left            =   600
      TabIndex        =   46
      Top             =   570
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   570
      Width           =   255
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   41
      Left            =   3000
      TabIndex        =   42
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   40
      Left            =   2520
      TabIndex        =   41
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   39
      Left            =   2040
      TabIndex        =   40
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   38
      Left            =   1560
      TabIndex        =   39
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   37
      Left            =   1080
      TabIndex        =   38
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   36
      Left            =   600
      TabIndex        =   37
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   35
      Left            =   120
      TabIndex        =   36
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   34
      Left            =   3000
      TabIndex        =   35
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   33
      Left            =   2520
      TabIndex        =   34
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   32
      Left            =   2040
      TabIndex        =   33
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   31
      Left            =   1560
      TabIndex        =   32
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   30
      Left            =   1080
      TabIndex        =   31
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   29
      Left            =   600
      TabIndex        =   30
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   28
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   27
      Left            =   3000
      TabIndex        =   28
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   26
      Left            =   2520
      TabIndex        =   27
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   25
      Left            =   2040
      TabIndex        =   26
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   24
      Left            =   1560
      TabIndex        =   25
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   23
      Left            =   1080
      TabIndex        =   24
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   22
      Left            =   600
      TabIndex        =   23
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   21
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   20
      Left            =   3000
      TabIndex        =   21
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   19
      Left            =   2520
      TabIndex        =   20
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   2040
      TabIndex        =   19
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   1560
      TabIndex        =   18
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   1080
      TabIndex        =   17
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   600
      TabIndex        =   16
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   3000
      TabIndex        =   14
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   2520
      TabIndex        =   13
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   2040
      TabIndex        =   12
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   1080
      TabIndex        =   10
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   600
      TabIndex        =   9
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   240
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   240
   End
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True



Private Sub updAnio_Change()
    txtAnio.Text = updAnio.Value
End Sub

Private Sub UserControl_Initialize()
    With cboMes
        .AddItem "Enero"
        .AddItem "Febrero"
        .AddItem "Marzo"
        .AddItem "Abril"
        .AddItem "Mayo"
        .AddItem "Junio"
        .AddItem "Julio"
        .AddItem "Agosto"
        .AddItem "Septiembre"
        .AddItem "Octubre"
        .AddItem "Noviembre"
        .AddItem "Diciembre"
    End With
    
    cboMes.ListIndex = Month(Date) - 1
    txtAnio.Text = Year(Date)

    For k = 0 To 41
        lblDia.Item(k).Caption = k
    Next
    
    'Detecto el día 1 del mes.
    Dim diaUno As Byte
    diaUno = Weekday("1/" & cboMes.ListIndex + 1 & "/" & txtAnio.Text, vbSunday)

    dia = 1
    For k = diaUno - 1 To 41
        lblDia.Item(k).Caption = dia
        dia = dia + 1
    Next
    
    'Oculto/muestro días.
    For k = 0 To diaUno - 2
        lblDia.Item(k).Visible = False
    Next
    For k = diaUno - 1 To ULTIMO_DIA(cboMes.ListIndex + 1)
        
End Sub


