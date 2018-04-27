VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FACTURACIÓN"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "frmFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboNumCuotas 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   6510
      Width           =   615
   End
   Begin VB.TextBox txtNumTicket 
      Height          =   285
      Left            =   8880
      MaxLength       =   6
      TabIndex        =   55
      Top             =   6520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame11 
      Caption         =   "Cursos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   48
      Top             =   5520
      Width           =   1095
      Begin VB.ListBox lstCursos 
         Height          =   1035
         Left            =   120
         TabIndex        =   49
         Top             =   200
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Pago a cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   41
      Top             =   7200
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdPagoACuenta 
         Caption         =   "Pago"
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtDetallePago 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtImportePago 
         Height          =   285
         Left            =   840
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detalle"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   645
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cheques"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   40
      Top             =   5520
      Width           =   3735
      Begin VB.TextBox txtLinea3 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   15
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtLinea2 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   14
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtLinea1 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   13
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   9240
      Picture         =   "frmFactura.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cancelar"
      Top             =   7320
      Width           =   855
   End
   Begin VB.Frame Frame10 
      Caption         =   "Comprobantes anteriores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   37
      Top             =   6960
      Width           =   7935
      Begin VB.ListBox lstUltimosPagos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         ItemData        =   "frmFactura.frx":0884
         Left            =   120
         List            =   "frmFactura.frx":0886
         TabIndex        =   19
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Frame fraComprobantes 
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   7680
      TabIndex        =   35
      Top             =   840
      Width           =   2415
      Begin VB.OptionButton optFacturaC 
         Caption         =   "Factura C"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optReciboC 
         Caption         =   "Recibo"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraModoPago 
      Caption         =   "Modo de pago"
      Enabled         =   0   'False
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
      Left            =   5400
      TabIndex        =   34
      Top             =   1560
      Width           =   2175
      Begin VB.OptionButton optContado 
         Caption         =   "Contado"
         Height          =   255
         Left            =   980
         TabIndex        =   9
         Top             =   240
         Width           =   900
      End
      Begin VB.OptionButton optCuotas 
         Caption         =   "Cuotas"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   800
      End
   End
   Begin VB.Frame fraEmpresa 
      Caption         =   "Empresa"
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
      Left            =   3000
      TabIndex        =   33
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtEmpresa 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Unidad de negocio"
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
      Left            =   5400
      TabIndex        =   32
      Top             =   840
      Width           =   2175
      Begin VB.ComboBox cboUnidadNegocio 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Emisor"
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
      Left            =   7680
      TabIndex        =   31
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox cboEmisor 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Forma de pago"
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
      Left            =   2760
      TabIndex        =   30
      Top             =   1560
      Width           =   2535
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   855
      Left            =   8160
      Picture         =   "frmFactura.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Guardar"
      Top             =   7320
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Condición de IVA"
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
      TabIndex        =   28
      Top             =   1560
      Width           =   2535
      Begin VB.ComboBox cboCondIva 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "CUIT / DNI"
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
      Left            =   5880
      TabIndex        =   27
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Dirección"
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
      TabIndex        =   26
      Top             =   840
      Width           =   5175
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Alumno"
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
      TabIndex        =   25
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton cmdVerFichaAlumno 
         Height          =   375
         Left            =   2280
         Picture         =   "frmFactura.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid dbgFactura 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraComprobantesFabru 
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   7680
      TabIndex        =   36
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         TabIndex        =   24
         Text            =   "21"
         Top             =   915
         Width           =   615
      End
      Begin VB.OptionButton optFacturaB 
         Caption         =   "Factura B"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optFacturaA 
         Caption         =   "Factura A"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1200
         TabIndex        =   39
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.Label lblRecargo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   6360
      TabIndex        =   63
      Top             =   6075
      Width           =   855
   End
   Begin VB.Label lblLblRecargo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recargo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5235
      TabIndex        =   62
      Top             =   6075
      Width           =   915
   End
   Begin VB.Label lblPesosRecargo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   6240
      TabIndex        =   61
      Top             =   6075
      Width           =   135
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   6360
      TabIndex        =   60
      Top             =   5590
      Width           =   855
   End
   Begin VB.Label lblLblSubTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   59
      Top             =   5590
      Width           =   870
   End
   Begin VB.Label lblPesosSubTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   6240
      TabIndex        =   58
      Top             =   5590
      Width           =   135
   End
   Begin VB.Label lblNumCuotas 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuotas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   56
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNumTicket 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      TabIndex        =   54
      Top             =   6480
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblSignoPesos 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1440
      TabIndex        =   53
      Top             =   2280
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblAnticipo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1440
      TabIndex        =   52
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTextoAnticipo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anticipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   51
      Top             =   2280
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   8400
      TabIndex        =   47
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   8520
      TabIndex        =   46
      Top             =   6000
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7440
      TabIndex        =   45
      Top             =   6000
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7320
      Top             =   6000
      Width           =   2805
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      TabIndex        =   44
      Top             =   5520
      Width           =   705
   End
   Begin VB.Shape shpRectangulo 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   9360
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblTotalFactura 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   8640
      TabIndex        =   29
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7320
      Top             =   5520
      Width           =   2805
   End
   Begin VB.Shape shpAnticipo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   300
      Left            =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Shape shpNumTicket 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7320
      Top             =   6480
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Shape shpNumCuotas 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   5160
      Top             =   6480
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Shape shpSubTotal 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   5160
      Top             =   5520
      Width           =   2085
   End
   Begin VB.Shape shpRecargo 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   5160
      Top             =   6000
      Width           =   2085
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim evento_load As Boolean
Dim item_actual As Byte
Dim FacturaPorEmpresa As Boolean
Dim Descuento_Ex As Single
Dim Descuento_Contado As Single
Dim Descuento_Total As Single
Dim Total_Factura As Single
Dim Total_Descuento As Single
Dim ultimo_id As Long
'Dim ultimo_id As Integer
Dim aplico_anticipo As Boolean

Private Sub cboEmisor_Click()
    If Left(cboEmisor.Text, 1) = 1 Then
        fraComprobantes.Visible = False
        fraComprobantesFabru.Visible = True
        
        If cboCondIva.Text = "RESPONSABLE INSCRIPTO" Then
            optFacturaA.Enabled = True
        Else
            optFacturaA.Enabled = False
            optFacturaB.Value = True
        End If
    Else
        fraComprobantes.Visible = True
        fraComprobantesFabru.Visible = False
    End If
End Sub

Private Sub cboFormaPago_Click()
    'Si es TARJETA muestro txtNumTicket y cboNumCuotas
    If Left(cboFormaPago.Text, 2) = "T." Then
        cboNumCuotas.Text = 1
        
        CALCULAR_TOTAL
        CALCULAR_RECARGO
        
        lblNumTicket.Visible = True
        shpNumTicket.Visible = True
        txtNumTicket.Visible = True
        shpNumCuotas.Visible = True
        lblNumCuotas.Visible = True
        cboNumCuotas.Visible = True
        
        lblLblSubTotal.Visible = True
        lblPesosSubTotal.Visible = True
        lblSubTotal.Visible = True
        shpSubTotal.Visible = True
        lblLblRecargo.Visible = True
        lblPesosRecargo.Visible = True
        lblRecargo.Visible = True
        shpRecargo.Visible = True
    Else
        CALCULAR_TOTAL
        
        lblNumTicket.Visible = False
        shpNumTicket.Visible = False
        txtNumTicket.Visible = False
        shpNumCuotas.Visible = False
        lblNumCuotas.Visible = False
        cboNumCuotas.Visible = False
    
        lblLblSubTotal.Visible = False
        lblPesosSubTotal.Visible = False
        lblSubTotal.Visible = False
        shpSubTotal.Visible = False
        lblLblRecargo.Visible = False
        lblPesosRecargo.Visible = False
        lblRecargo.Visible = False
        shpRecargo.Visible = False
    End If
End Sub

Private Sub cboNumCuotas_Click()
    If Not evento_load Then
        CALCULAR_RECARGO
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim Coeficiente As Single
    Dim Comision As Single
    Dim TotFacturaSinTarjeta As Single
    Dim TotFacturaConTarjeta As Single
    
    If txtNombre.Text = "" Then
        MsgBox "El nombre del cliente/alumno no puede estar en blanco.", vbCritical, "ERROR"
        Exit Sub
    End If
    
    If cboEmisor = "(No disponible)" Then
        MsgBox "Debe seleccionar un emisor.", vbCritical, "ERROR"
        Exit Sub
    End If
    
    If cboFormaPago = "(No disponible)" Then
        MsgBox "Debe seleccionar una forma de pago.", vbCritical, "ERROR"
        Exit Sub
    ElseIf Left(cboFormaPago.Text, 2) = "T." Then
        If txtNumTicket.Text = "" Then
            MsgBox "Debe indicar el número de ticket", vbCritical, "ERROR"
            Exit Sub
        Else
            CALCULAR_RECARGO
        End If
    End If
    
    If Val(lblTotalFactura.Caption) = 0 Then
        MsgBox "El total del documento no puede ser 0.", vbCritical, "ERROR"
        Exit Sub
    End If
    
    If sMenu = "FacturaVentaHardware" And txtIva.Text = "" Then
        MsgBox "Debe ingresar el porcentaje de IVA.", vbCritical, "ERROR"
        Exit Sub
    End If
        
    
    'Genero el encabezado
    With adoMovimientos
        CERRAR_TABLA adoMovimientos
        .Open "Movimientos", adoConnection, adOpenDynamic, adLockOptimistic
        
        .AddNew
        
        !Sucursal = "000" & Left(cboEmisor.Text, 1)
        
        'EMISORES
        '1- FABRU
        '2- SERGIO
        '3- NESTOR
        If sMenu <> "NotaCreditoPresenciales" And sMenu <> "NotaCreditoOtros" Then
            If Left(cboEmisor.Text, 1) = 1 Then
                If optFacturaA.Value Then
                    !TipoDoc = "FCA"
                    !NumDoc = ULTIMO_NUMERO("UltA")
                    x_comprobante = "UltA"
                ElseIf optFacturaB.Value Then
                    !TipoDoc = "FCB"
                    !NumDoc = ULTIMO_NUMERO("UltB")
                    x_comprobante = "UltB"
                Else
                    !TipoDoc = "REC"
                    !NumDoc = ULTIMO_NUMERO("UltR")
                    x_comprobante = "UltR"
                End If
            Else
                If optFacturaC.Value Then
                    !TipoDoc = "FCC"
                    If Left(cboEmisor.Text, 1) = 2 Then
                        !NumDoc = ULTIMO_NUMERO("UltFS")
                        x_comprobante = "UltFS"
                    Else
                        !NumDoc = ULTIMO_NUMERO("UltFN")
                        x_comprobante = "UltFN"
                    End If
                Else
                    !TipoDoc = "REC"
                    If Left(cboEmisor.Text, 1) = 2 Then
                        !NumDoc = ULTIMO_NUMERO("UltRS")
                        x_comprobante = "UltRS"
                    Else
                        !NumDoc = ULTIMO_NUMERO("UltRN")
                        x_comprobante = "UltRN"
                    End If
                End If
            End If
        Else
            If Left(cboEmisor.Text, 1) = 1 Then
                If optFacturaA.Value Then
                    !TipoDoc = "NCA"
                    !NumDoc = ULTIMO_NUMERO("UltNCAF")
                    x_comprobante = "UltNCAF"
                ElseIf optFacturaB.Value Then
                    !TipoDoc = "NCB"
                    !NumDoc = ULTIMO_NUMERO("UltNCBF")
                    x_comprobante = "UltNCBF"
                Else
                    !TipoDoc = "REC"
                    !NumDoc = ULTIMO_NUMERO("UltR")
                    x_comprobante = "UltR"
                End If
            Else
                If optFacturaC.Value Then
                    !TipoDoc = "NCC"
                    If Left(cboEmisor.Text, 1) = 2 Then
                        !NumDoc = ULTIMO_NUMERO("UltFS")
                        x_comprobante = "UltFS"
                    Else
                        !NumDoc = ULTIMO_NUMERO("UltFN")
                        x_comprobante = "UltFN"
                    End If
                End If
            End If
        
        End If
        
        frmNumImpresion.lblNumImpresion.Caption = Right(cboEmisor.Text, Len(cboEmisor.Text) - 2) & " " & Format(adoMovimientos!NumDoc, "00000000")
        frmNumImpresion.lblImporte.Caption = "Importe: $" & Format(lblTotalFactura.Caption, "Fixed") & ".-"
        frmNumImpresion.Show vbModal
        
     End With
     
     If x_imprimir Then
        With adoMovimientos
            !fecha = Date
            !idAlumno = id_Alumno
            !idCurso = id_Curso
            
            If FacturaPorEmpresa Then
                !NombreCliente = txtEmpresa.Text
            Else
                !NombreCliente = txtNombre.Text
            End If
            !NombreEmisor = cboEmisor.Text
            !Subtotal = Val(lblTotalFactura.Caption) / (1 + (Val(txtIva.Text) / 100))
            !Iva = (Val(lblTotalFactura.Caption) / (1 + (Val(txtIva.Text) / 100))) * (Val(txtIva.Text) / 100)
            !Total = Val(lblTotalFactura.Caption)
            !Recargo = Val(lblRecargo.Caption)
            
            !Direccion = txtDireccion.Text & ""
            !CondIva = cboCondIva.Text & ""
            !Cuit = txtCuit.Text & ""
            !FormaPago = cboFormaPago.Text & ""
            
            If cboFormaPago.Text <> "CUENTA CORRIENTE" Then
                If sMenu <> "FacturaAnticipo" Then
                    !Saldo = 0
                Else
                    !Saldo = Val(lblTotalFactura.Caption)
                End If
            Else
                !Saldo = Val(lblTotalFactura.Caption)
            End If
            
            '!idUnidadNegocio = DEVOLVER_ID(cboUnidadNegocio.Text, adoUnidadesNegocio, "UnidadesNegocio", "Detalle")
            !UnidadNegocio = cboUnidadNegocio.Text
            
            !Linea1 = txtLinea1.Text & ""
            !Linea2 = txtLinea2.Text & ""
            !Linea3 = txtLinea3.Text & ""
            
            !NumTicket = txtNumTicket.Text & ""
            
            .Update
            
            .MoveLast
            ultimo_id = !id
            
            .Close
        End With
        
        'Genero el cuerpo
        With adoItemsXMov
            CERRAR_TABLA adoItemsXMov
            .Open "ItemsXMov", adoConnection, adOpenDynamic, adLockOptimistic
            
            adoTempFactura.MoveFirst
            Do While Not adoTempFactura.EOF
                If adoTempFactura!Cantidad <> 0 And adoTempFactura!Paga <> 0 Then
                    .AddNew
                    
                    !idMovimiento = ultimo_id
                    '!idCurso = id_Curso
                    !Cantidad = adoTempFactura!Cantidad
                    
                    If sMenu <> "NotaCreditoPresenciales" And sMenu <> "NotaCreditoOtros" Then
                        texto_adicional_detalle = ""
                        If sMenu = "FacturaPresenciales" Then
                            If adoTempFactura!Saldo < adoTempFactura!Unitario And adoTempFactura!Paga = adoTempFactura!Saldo Then
                                texto_adicional_detalle = "SALDO"
                            End If
                            If adoTempFactura!Paga <> adoTempFactura!Saldo And optCuotas.Value = True Then
                                texto_adicional_detalle = "PAGO A CUENTA"
                            End If
                        End If
                        
                        !Detalle = texto_adicional_detalle & " " & adoTempFactura!Alumno & " " & adoTempFactura!Detalle
                    Else
                        !Detalle = adoTempFactura!Alumno & " ANULA " & adoTempFactura!Detalle
                    End If
                    !Unitario = adoTempFactura!Unitario
                    !Descuento = adoTempFactura!Descuento
                    !Importe = adoTempFactura!Paga
                    
                    If cboEmisor.Text <> "1-Fabru S.A." And cboFormaPago.Text <> "CUENTA CORRIENTE" Then
                        !Saldo = adoTempFactura!Saldo - adoTempFactura!Paga
                    End If
                    
                    If sMenu = "FacturaAnticipo" Then
                        !Saldo = adoTempFactura!Paga
                    End If
                    
                    !idMod = adoTempFactura!idMovimiento
                    
                    sSql = "SELECT idCurso FROM Movimientos WHERE id = " & adoTempFactura!idMovimiento
                    CERRAR_TABLA adoTemp
                    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                    
                    If adoTemp.EOF Then
                        !idCurso = 0
                    Else
                        !idCurso = adoTemp!idCurso
                    End If
                    
                    adoTemp.Close
                    
                    .Update
                    
                    If sMenu = "FacturaPresenciales" Then
                            If Left(adoTempFactura!Detalle, 11) <> "A CUENTA - " Then
                                'Actualizo el saldo en el documento modelo - Movimientos
                                sSql = "SELECT * FROM Movimientos WHERE id = " & adoTempFactura!idMovimiento
                                CERRAR_TABLA adoTemp
                                adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                                If optContado Then
                                    adoTemp!Saldo = 0
                                Else
                                
                                    If adoTemp!Descuento = "EX ALUMNO" Then
                                        
                                        adoTemp!Saldo = adoTemp!Saldo - (adoTemp!Saldo * 15 / 100)
                                    
                                    End If
                                
                                    adoTemp!Saldo = adoTemp!Saldo - adoTempFactura!Paga
                                End If
                                adoTemp.Update
                                adoTemp.Close
                
                                'Actualizo el saldo en el documento modelo - itemsXMov
                                sSql = "SELECT * FROM ItemsXMov WHERE id = " & adoTempFactura!idItem
                                CERRAR_TABLA adoTemp
                                adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                                If optContado Then
                                    adoTemp!Saldo = 0
                                Else
                                    If adoTemp!Descuento = "EX ALUMNO" Then
                                        
                                        adoTemp!Saldo = adoTemp!Saldo - (adoTemp!Saldo * 15 / 100)
                                    
                                    End If
                                    
                                    adoTemp!Saldo = adoTemp!Saldo - adoTempFactura!Paga
                                    
                                End If
                                adoTemp.Update
                                adoTemp.Close

                        End If
                    ElseIf sMenu = "NotaCreditoPresenciales" Then
                        'Guardo el id del documento que se va a anular.
                        id_anulando = adoTempFactura!idMovimiento
                            
                        sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & id_anulando
                        With adoTemp
                            CERRAR_TABLA adoTemp
                            .Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                            .MoveFirst
                            Do While Not .EOF
                                sSql = "SELECT * FROM Movimientos WHERE id = " & !idMod
                                CERRAR_TABLA adoTemp2
                                adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                                adoTemp2!Saldo = adoTemp2!Saldo + !Importe
                                adoTemp2.Update
                                adoTemp2.Close
                                
                                sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & !idMod
                                adoTemp3.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                                adoTemp3!Saldo = adoTemp3!Saldo + !Importe
                                adoTemp3.Update
                                adoTemp3.Close
                                
                                .MoveNext
                            Loop
                            .Close
                        End With
                    
                    End If
                End If
                adoTempFactura.MoveNext
            Loop
            
            .Close
        End With
         
        If sMenu <> "FacturaAnticipo" Then
            CHEQUEAR_A_CUENTA_FACTURA
        End If
        
        IMPRIMIR_DOCUMENTO
        
        'MOSTRAR_ACUMULADO_MENSUAL
        
        Unload Me

    End If

    End Sub

Private Sub cmdPagoACuenta_Click()
    adoTempFactura.AddNew
    adoTempFactura!Cantidad = 1
    adoTempFactura!Detalle = "A CUENTA - " & txtDetallePago.Text
    adoTempFactura!Unitario = Val(txtImportePago.Text)
    adoTempFactura!Descuento = 0
    adoTempFactura!Importe = Val(txtImportePago.Text)
    adoTempFactura!Saldo = 0
    adoTempFactura!Paga = Val(txtImportePago.Text)
    adoTempFactura.Update
    adoTempFactura.MoveFirst
    
    CALCULAR_TOTAL
End Sub

Private Sub cmdVerFichaAlumno_Click()
    x_ficha_alumno_desde_factura = txtNombre.Text
    
    frmAlumnos.Show vbModal
End Sub

Private Sub dbgFactura_AfterColUpdate(ByVal ColIndex As Integer)
    'If Left(adoTempFactura!Detalle, 11) = "A CUENTA - " Then
    '    adoTempFactura.CancelUpdate
    '    Exit Sub
    'End If
    
    If dbgFactura.Col = 2 Then 'Cantidad
        If sMenu = "FacturaPresenciales" Then
            If adoTempFactura!Cantidad <> 0 And adoTempFactura!Cantidad <> 1 Then
                MsgBox "Solo se permite ingresar 0 ó 1", vbCritical, "ERROR"
                adoTempFactura!Cantidad = 1
            Else
                dbgFactura.Col = 3
            End If
        Else
            adoTempFactura!Importe = adoTempFactura!Cantidad * adoTempFactura!Unitario
            adoTempFactura!Saldo = adoTempFactura!Importe
            adoTempFactura!Paga = adoTempFactura!Importe
            adoTempFactura.MoveNext
            CALCULAR_TOTAL
            adoTempFactura.MovePrevious
            dbgFactura.Col = 3
        End If
    ElseIf dbgFactura.Col = 3 Then 'Detalle
        If sMenu = "FacturaAnticipo" Then
            adoTempFactura!Cantidad = 1
            adoTempFactura!Detalle = "A CUENTA - " & adoTempFactura!Detalle
            
            'CALCULAR_TOTAL
        
        End If

        If sMenu = "FacturaPresenciales" Then
            dbgFactura.Col = 8
        Else
            dbgFactura.Col = 4
        End If
    ElseIf dbgFactura.Col = 4 Then 'Unitario
        If sMenu = "FacturaAnticipo" Then
            adoTempFactura!Importe = adoTempFactura!Unitario
            adoTempFactura!Descuento = 0
            adoTempFactura!Saldo = 0
            adoTempFactura!Paga = adoTempFactura!Unitario
        Else
            adoTempFactura!Importe = adoTempFactura!Cantidad * adoTempFactura!Unitario
            adoTempFactura!Saldo = adoTempFactura!Importe
            adoTempFactura!Paga = adoTempFactura!Importe
        End If
        
        adoTempFactura.MoveNext
        CALCULAR_TOTAL
        dbgFactura.Col = 2
    ElseIf dbgFactura.Col = 8 Then 'Paga
        If adoTempFactura!Paga <= adoTempFactura!Saldo Then
            adoTempFactura.Update
            'adoTempFactura.MoveNext
            CALCULAR_TOTAL
            CALCULAR_RECARGO
            dbgFactura.Col = 8
        Else
            MsgBox "El pago no puede ser superior al saldo", vbCritical, "ERROR"
            adoTempFactura!Paga = 0
        End If
    End If
End Sub

Private Sub dbgFactura_Click()
    If sMenu = "NotaCreditoPresenciales" Then
        cboEmisor.Text = adoTempFactura!NombreEmisor
        cboFormaPago.Text = "EFECTIVO"
        
        sSql = "SELECT * FROM Movimientos WHERE id = " & adoTempFactura!idMovimiento
        CERRAR_TABLA adoTemp3
        adoTemp3.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                
        txtNombre.Text = adoTemp3!NombreCliente
        txtDireccion.Text = adoTemp3!Direccion
        cboCondIva.Text = adoTemp3!CondIva
        
        adoTemp3.Close
        
        If Left(adoTempFactura!NombreEmisor, 1) = "1" Then
            fraComprobantes.Visible = False
            fraComprobantesFabru.Visible = True
        
            If Left(adoTempFactura!Detalle, 3) = "FCA" Then
                optFacturaA.Enabled = True
                optFacturaB.Enabled = False
                optFacturaA.Value = True
            Else
                optFacturaA.Enabled = False
                optFacturaB.Enabled = True
                optFacturaB.Value = True
            End If
        Else
            fraComprobantes.Visible = True
            fraComprobantesFabru.Visible = False
        End If
    ElseIf sMenu = "NotaCreditoOtros" Then
            cboFormaPago.Text = "EFECTIVO"
    End If
End Sub

Private Sub dbgFactura_DblClick()
    If sMenu <> "NotaCreditoPresenciales" And sMenu <> "NotaCreditoOtros" Then
        If dbgFactura.Col = 7 Then 'Saldo
            adoTempFactura!Paga = adoTempFactura!Saldo
            adoTempFactura.Update
            CALCULAR_TOTAL
        End If
    Else
        cboEmisor.Text = adoTempFactura!NombreEmisor
        det_fc = adoTempFactura!Detalle
        uni_fc = adoTempFactura!Unitario
        tot_fc = adoTempFactura!Importe
        adoTempFactura.AddNew
        adoTempFactura!Cantidad = 1
        adoTempFactura!Detalle = "ANULA COMPROBANTE " & det_fc
        adoTempFactura!Unitario = uni_fc
        adoTempFactura!Importe = tot_fc
        adoTempFactura.Update
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        CALCULAR_ACUMULADO_MENSUAL
    End If
End Sub

Private Sub Form_Load()
    evento_load = True
    
    sSql = "DELETE FROM TempFactura"
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoTempFactura
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    adoTempFactura.AddNew
    
    ARMAR_GRILLA
    
    CARGAR_COMBO "cboCondIva", adoCondIva, "CondIva", "Detalle", Me
    cboCondIva.AddItem "(No disponible)"
    
    CARGAR_COMBO "cboFormaPago", adoFormasPago, "FormasPago", "Detalle", Me
    
    CARGAR_COMBO "cboEmisor", adoEmisores, "Emisores", "Detalle", Me

    CARGAR_COMBO "cboUnidadNegocio", adoUnidadesNegocio, "UnidadesNegocio", "Detalle", Me
    
    CARGAR_COMBO "cboNumCuotas", adoRecargosTarjeta, "RecargosTarjeta", "Cuotas", Me
    
    Select Case sMenu
        Case "FacturaPresenciales"
            cboUnidadNegocio.Text = "CURSO PRESENCIAL"
        Case "FacturaDistancia"
            cboUnidadNegocio.Text = "CURSO A DISTANCIA"
        Case "FacturaServicioTecnico"
            cboUnidadNegocio.Text = "SERVICIO TECNICO"
        Case "FacturaVentaHardware"
            txtIva.Text = ""
            cboUnidadNegocio.Text = "VENTA DE HARDWARE"
        Case "NotaCreditoPresenciales", "NotaCreditoOtros"
            optFacturaA.Caption = "Nota de crédito A"
            optFacturaB.Caption = "Nota de crédito B"
            optFacturaC.Caption = "Nota de crédito C"
            optReciboC.Visible = False
            optCuotas.Enabled = False
            optContado.Value = True
    End Select
    cboUnidadNegocio.Enabled = False
    
    lblFecha.Caption = Date
    
    If x_alumno_form_inscrip_factura <> "" Then
        txtNombre.Text = x_alumno_form_inscrip_factura
        FACTURAR_DESDE_INSCRIPCION
    End If
    
    evento_load = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    x_alumno_form_inscrip_factura = ""
    x_es_ex_alumno = False
End Sub




Private Sub optContado_Click()
    If txtNombre.Text = "" And txtEmpresa.Text = "" Then
        Exit Sub
    End If
    'If Descuento_Ex <> 0 Then
    '    sSql = "UPDATE TempFactura SET Descuento = 20"
    'Else
    '    sSql = "UPDATE TempFactura SET Descuento = 10"
    'End If
    
    sSql = "UPDATE TempFactura SET Descuento = 20 WHERE Descuento = 15"
    adoConnection.Execute sSql
    sSql = "UPDATE TempFactura SET Descuento = 10 WHERE Descuento = 0"
    adoConnection.Execute sSql
    
    
    If Not x_es_ex_alumno Then
        sSql = "UPDATE TempFactura SET Importe = Unitario - (Unitario * Descuento / 100)"
    Else
        'sSql = "UPDATE TempFactura SET Importe = Unitario - (Unitario * 5.93 / 100)"
        sSql = "UPDATE TempFactura SET Importe = Importe - (Importe * 5.93 / 100)"
    End If
    adoConnection.Execute sSql
    
    If x_es_ex_alumno Then
        With adoTempFactura
            .Close
            .Open "TempFactura", adoConnection, adOpenDynamic, adLockOptimistic
            .MoveFirst
            
            Do While Not .EOF
                !Importe = Round(!Importe)
                .Update
                .MoveNext
            Loop
            
            '.Close
        End With
    End If
    
    sSql = "UPDATE TempFactura SET Saldo = 0"
    adoConnection.Execute sSql
    
    sSql = "UPDATE TempFactura SET Paga = Importe"
    adoConnection.Execute sSql
    
    adoTempFactura.Close
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    ARMAR_GRILLA
    
    CALCULAR_TOTAL
End Sub

Private Sub optCuotas_Click()
    If txtNombre.Text = "" And txtEmpresa.Text = "" Then
        Exit Sub
    End If
    'If Descuento_Ex <> 0 Then
    '    sSql = "UPDATE TempFactura SET Descuento = 15"
    'Else
    '    sSql = "UPDATE TempFactura SET Descuento = 0"
    'End If
    
    sSql = "UPDATE TempFactura SET Descuento = 0 WHERE Descuento = 10"
    adoConnection.Execute sSql
    sSql = "UPDATE TempFactura SET Descuento = 15 WHERE Descuento = 20"
    adoConnection.Execute sSql
    
    sSql = "UPDATE TempFactura SET Importe = Unitario - (Unitario * Descuento / 100)"
    adoConnection.Execute sSql
    
    'sSql = "UPDATE TempFactura SET Importe = Unitario"
    'adoConnection.Execute sSql
    
    sSql = "UPDATE TempFactura SET Saldo = SaldoReal"
    adoConnection.Execute sSql
    
    sSql = "UPDATE TempFactura SET Paga = 0"
    adoConnection.Execute sSql
    
    
    adoTempFactura.Close
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    ARMAR_GRILLA
    
    lblTotalFactura.Caption = "0.00"
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If txtEmpresa.Text = "" Then
        '    Exit Sub
        'End If
        
        FacturaPorEmpresa = True
        
        cboEmisor.Enabled = True
        fraComprobantesFabru.Enabled = True
        fraEmpresa.Enabled = True
        
        CERRAR_TABLA adoEmpresas
        sSql = "SELECT * FROM Empresas WHERE Nombre = '" & txtEmpresa.Text & "' OR Cuit = '" & txtEmpresa.Text & "'"
        adoEmpresas.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If adoEmpresas.EOF Then
            EstiloBuscador = "FacturaEmpresas"
            Extra = "Empresas"
            frmBuscador.Show vbModal
            
            If CanceloBuscador Then
                Unload Me
            End If
            
            Extra = ""
        Else
            txtNombre.Text = adoEmpresas!Nombre
            txtCuit.Text = adoEmpresas!Cuit
            cboCondIva.Text = DEVOLVER_CAMPO(adoEmpresas!idCondIva, adoCondIva, "CondIva", "Detalle")
            txtDireccion.Text = adoEmpresas!Direccion
        End If
        
        If cboCondIva.Text = "RESPONSABLE INSCRIPTO" Then
            optFacturaA.Enabled = True
        Else
            optFacturaA.Enabled = False
            optFacturaB.Value = True
        End If
        
        If Not CancelaBusEmpresa Then
            If sMenu <> "FacturaServicioTecnico" Then
                DOCUMENTOS_PENDIENTES
                VER_OTROS_PAGOS
                fraModoPago.Enabled = True
            Else
                GENERAR_RENGLONES
            End If
        End If
    End If
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FacturaPorEmpresa = False
        
        cboEmisor.Enabled = True
        fraComprobantesFabru.Enabled = True
        fraEmpresa.Enabled = True
        
        CERRAR_TABLA adoAlumnos
        sSql = "SELECT * FROM Alumnos WHERE Nombre = '" & txtNombre.Text & "' OR NumDoc = '" & txtNombre.Text & "'"
        adoAlumnos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If adoAlumnos.EOF Then
            EstiloBuscador = "FacturaAlumnos"
            Extra = "Alumno"
            frmBuscador.Show vbModal
            
            If CanceloBuscador Then
                Unload Me
            End If
            
            Extra = ""
        Else
            txtNombre.Text = adoAlumnos!Nombre
            txtDireccion.Text = adoAlumnos!Direccion
        End If
        
        If sMenu = "FacturaPresenciales" Then
            DOCUMENTOS_PENDIENTES
            VER_OTROS_PAGOS
            ASOCIACION_CON_EMPRESA
        ElseIf sMenu = "NotaCreditoPresenciales" Then
            DOCUMENTOS_EMITIDOS
        Else
            GENERAR_RENGLONES
        End If
    
        If cboCondIva.Text = "RESPONSABLE INSCRIPTO" Then
            optFacturaA.Enabled = True
        Else
            optFacturaA.Enabled = False
            optFacturaB.Value = True
        End If
        
        fraModoPago.Enabled = True
        CHEQUEAR_PERMITIR_CONTADO
    
        MOSTRAR_ANTICIPO
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub CALCULAR_TOTAL()
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Paga) AS total FROM TempFactura"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    Total_Factura = adoTabla!Total
    
    adoTabla.Close
    
    lblTotalFactura.Caption = Format(Round(Total_Factura), "Fixed")
End Sub

Private Sub MOSTRAR_DOCUS_PENDIENTES()
    Dim alumno_en_espera As Boolean
    
    sSql = "DELETE FROM TempFactura"
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoTempFactura
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    lstCursos.Clear
    
    With adoMovimientos
        alumno_en_espera = False
        
        Do While Not .EOF
            If !TipoDoc = "ESP" Then
                alumno_en_espera = True
            End If
            
            CERRAR_TABLA adoItemsXMov
            sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & !id
            adoItemsXMov.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            Descuento_Ex = adoItemsXMov!Descuento
            
            If Descuento_Ex <> 0 Then
                x_es_ex_alumno = True
            End If
            
            Do While Not adoItemsXMov.EOF
                'If Left(adoItemsXMov!Detalle, 12) <> "BONIFICACION" Then
                    adoTempFactura.AddNew
    
                    adoTempFactura!idMovimiento = !id
                    
                    adoTempFactura!Cantidad = 1
                    If FacturaPorEmpresa Then
                        adoTempFactura!Alumno = adoMovimientos!Nombre
                    End If
                    adoTempFactura!Detalle = adoItemsXMov!Detalle
                    adoTempFactura!Unitario = adoItemsXMov!Unitario
                    adoTempFactura!Descuento = adoItemsXMov!Descuento
                    adoTempFactura!Importe = adoItemsXMov!Importe
                    adoTempFactura!Saldo = adoItemsXMov!Saldo
                    adoTempFactura!SaldoReal = adoItemsXMov!Saldo
                    
                    adoTempFactura!idItem = adoItemsXMov!id
                    
                    adoTempFactura.Update
                    
                    'Agrego los cursos en lstCursos
                    sSql = "SELECT Numero FROM Cursos WHERE id = " & adoItemsXMov!idCurso
                    CERRAR_TABLA adoCursos
                    adoCursos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                    
                    encontrado = False
                    If lstCursos.ListCount > 0 Then
                        For k = 0 To lstCursos.ListCount - 1
                            If lstCursos.List(k) = adoCursos!Numero Then
                                encontrado = True
                                Exit For
                            End If
                        Next
                    End If
                    If Not encontrado Then
                        If Not adoCursos.EOF Then
                            lstCursos.AddItem adoCursos!Numero
                        End If
                    End If
                    adoCursos.Close
                'Else
                '    total_bonificacion = total_bonificacion + adoItemsXMov!Unitario
                '    total_saldo_bonificacion = total_saldo_bonificacion + adoItemsXMov!Saldo
                'End If
                
                adoItemsXMov.MoveNext
            Loop
            
            adoItemsXMov.Close
            
            .MoveNext
        Loop
        
        'INI - Muestro la bonificación
        'adoTempFactura.AddNew

        ''adoTempFactura!idMovimiento = !id
        
        'adoTempFactura!Cantidad = 1
        ''If FacturaPorEmpresa Then
        ''    adoTempFactura!Alumno = !Nombre
        ''End If
        'adoTempFactura!Detalle = "BONIFICACION EX ALUMNO"
        'adoTempFactura!Unitario = total_bonificacion
        'adoTempFactura!Importe = total_bonificacion
        'adoTempFactura!Saldo = total_saldo_bonificacion
        
        ''adoTempFactura!idItem = adoItemsXMov!id
        
        'adoTempFactura.Update
        'FIN - Muestro la bonificación
        
        .Close
    End With
    
    ARMAR_GRILLA
    
    If alumno_en_espera Then
        alumno_en_espera = False
        MsgBox "ATENCIÓN!!!" & vbCrLf & "Este alumno está en lista de espera.", vbExclamation, "ALUMNO EN ESPERA"
    End If
End Sub

Private Sub DOCUMENTOS_PENDIENTES()
    CERRAR_TABLA adoMovimientos
    
    If FacturaPorEmpresa Then
        sSql = "SELECT Movimientos.id, Movimientos.Sucursal, Movimientos.TipoDoc, Movimientos.NumDoc, Movimientos.Fecha, Movimientos.idAlumno, Movimientos.SubTotal, Movimientos.Iva, Movimientos.Total, Movimientos.Saldo, Movimientos.Descuento, Alumnos.Nombre FROM Movimientos, Alumnos WHERE (TipoDoc = 'MOD' OR TipoDoc = 'ESP') AND Alumnos.idEmpresa = " & id_Empresa & " AND Movimientos.idAlumno = Alumnos.id AND Saldo > 1"
    Else
        sSql = "SELECT * FROM Movimientos WHERE (TipoDoc = 'MOD' OR TipoDoc = 'ESP') AND idAlumno = " & id_Alumno & " AND Saldo > 1"
    End If
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoMovimientos.EOF Then
        'If Not FacturaPorEmpresa Then
            MOSTRAR_DOCUS_PENDIENTES
        'End If
    Else
        adoMovimientos.Close
    End If
End Sub


Private Sub DOCUMENTOS_EMITIDOS()
    CERRAR_TABLA adoMovimientos
    
    If FacturaPorEmpresa Then
        sSql = "SELECT Movimientos.id, Movimientos.Sucursal, Movimientos.TipoDoc, Movimientos.NumDoc, Movimientos.Fecha, Movimientos.idAlumno, Movimientos.SubTotal, Movimientos.Iva, Movimientos.Total, Movimientos.Saldo, Movimientos.Descuento, Alumnos.Nombre FROM Movimientos, Alumnos WHERE (Left(TipoDoc, 2) = 'FC' OR Left(TipoDoc, 3) = 'REC') AND Alumnos.idEmpresa = " & id_Empresa & " AND Movimientos.idAlumno = Alumnos.id AND NOT Anulado"
    Else
        sSql = "SELECT Movimientos.* FROM Movimientos " & _
               "WHERE (Left(Movimientos.TipoDoc, 2) = 'FC' OR Left(Movimientos.TipoDoc, 3) = 'REC') " & _
               "AND Movimientos.idAlumno = " & id_Alumno & _
               " AND NOT Anulado"
    End If
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoMovimientos.EOF Then
        MOSTRAR_DOCUS_EMITIDOS
    Else
        adoMovimientos.Close
    End If
End Sub

Private Sub MOSTRAR_DOCUS_EMITIDOS()
    Dim tiene_nc As Boolean
    
    sSql = "DELETE FROM TempFactura"
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoTempFactura
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    With adoMovimientos
        Do While Not .EOF
            
            tiene_nc = False
            
            If sMenu = "NotaCreditoPresenciales" Then
                'VER SI LA FACTURA QUE VOY A MOSTRAR EN LA GRILLA
                'FUE OBJETO DE UNA NC, NO MOSTRARLA.
                
                sSql = "SELECT * FROM itemsXMov WHERE idMod = " & !id
                CERRAR_TABLA adoTemp3
                adoTemp3.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                
                If Not adoTemp3.EOF Then
                    tiene_nc = True
                End If
                
                adoTemp3.Close
                
            End If
            
            If Not tiene_nc Then
                adoTempFactura.AddNew
    
                adoTempFactura!idMovimiento = !id
                
                'adoTempFactura!Cantidad = 1
                If FacturaPorEmpresa Then
                    adoTempFactura!Alumno = !Nombre
                End If
                adoTempFactura!Detalle = !TipoDoc & " - " & !NumDoc
                adoTempFactura!Unitario = !Total 'Decía !Subtotal
                adoTempFactura!Descuento = !Descuento
                adoTempFactura!Importe = !Total
                adoTempFactura!Saldo = !Saldo
                adoTempFactura!SaldoReal = !Saldo
                adoTempFactura!NombreEmisor = !NombreEmisor
                
                adoTempFactura!idItem = !id
                
                adoTempFactura.Update
            End If
            
            .MoveNext
        Loop
        
        adoTempFactura.MoveFirst
        cboEmisor.Text = adoTempFactura!NombreEmisor
        cboFormaPago.Text = "EFECTIVO"
        
        .Close
    End With
    
    ARMAR_GRILLA
End Sub

Private Sub ACTUALIZAR_TOTAL()
    'Dim total As Single
    
    'total = 0
    'For k = 0 To txtDetalle.Count - 1
    '    total = total + Val(lblTotal(k).Caption)
    'Next
    
    'lblTotalFactura.Caption = Format(total, "Fixed")
End Sub

Private Sub ARMAR_GRILLA()
    CERRAR_TABLA adoTempFactura
    
    sSql = "SELECT * FROM TempFactura ORDER BY Detalle"
    adoTempFactura.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic

    Set dbgFactura.DataSource = adoTempFactura
    
    dbgFactura.Columns(0).Visible = False
    If Not FacturaPorEmpresa Then
        dbgFactura.Columns(1).Visible = False
        dbgFactura.Columns(1).Locked = True
    End If
    dbgFactura.Columns(2).Caption = "Cant."
    dbgFactura.Columns(2).Width = 600
    dbgFactura.Columns(2).Alignment = dbgRight
    dbgFactura.Columns(2).NumberFormat = "Fixed"
    dbgFactura.Columns(3).Width = 5000
    
    If sMenu = "FacturaPresenciales" Then
        dbgFactura.Columns(3).Locked = True
    End If
    
    dbgFactura.Columns(4).Caption = "Unit."
    dbgFactura.Columns(4).Width = 800
    dbgFactura.Columns(4).Alignment = dbgRight
    dbgFactura.Columns(4).NumberFormat = "Fixed"
    
    If sMenu = "FacturaPresenciales" Then
        dbgFactura.Columns(4).Locked = True
    End If
    
    dbgFactura.Columns(5).Caption = "% Dto."
    dbgFactura.Columns(5).Width = 800
    dbgFactura.Columns(5).Alignment = dbgRight
    dbgFactura.Columns(5).NumberFormat = "Fixed"
    
    If sMenu = "FacturaPresenciales" Then
        dbgFactura.Columns(5).Locked = True
    End If
    
    dbgFactura.Columns(6).Width = 800
    dbgFactura.Columns(6).Alignment = dbgRight
    dbgFactura.Columns(6).NumberFormat = "Fixed"
    
    If sMenu = "FacturaPresenciales" Then
        dbgFactura.Columns(6).Locked = True
    End If
    
    dbgFactura.Columns(7).Width = 800
    dbgFactura.Columns(7).Alignment = dbgRight
    dbgFactura.Columns(7).NumberFormat = "Fixed"
    
    If sMenu = "FacturaPresenciales" Then
        dbgFactura.Columns(7).Locked = True
    End If
    
    dbgFactura.Columns(8).Width = 800
    dbgFactura.Columns(8).Alignment = dbgRight
    dbgFactura.Columns(8).NumberFormat = "Fixed"
    dbgFactura.Columns(9).Visible = False
    dbgFactura.Columns(10).Visible = False
    dbgFactura.Columns(11).Visible = False
    
    adoTempFactura.MoveFirst
End Sub

Private Sub VER_OTROS_PAGOS()
    CERRAR_TABLA adoTabla
    If FacturaPorEmpresa Then
        'sSql = "SELECT Movimientos.*, TiposCurso.Detalle AS Curso, ItemsXMov.Detalle AS Descripcion FROM Movimientos, Cursos, TiposCurso, ItemsXMov WHERE NombreCliente = '" & txtEmpresa.Text & "' AND NOT Anulado AND Cursos.id = Movimientos.idCurso AND TiposCurso.id = Cursos.idTipoCurso AND ItemsXMov.idMovimiento = Movimientos.id ORDER BY Fecha DESC"
        sSql = "SELECT Movimientos.*, TiposCurso.Detalle AS Curso, ItemsXMov.Detalle AS Descripcion FROM Movimientos, Cursos, TiposCurso, ItemsXMov WHERE Movimientos.idAlumno = " & id_Alumno & " AND NOT Anulado AND Cursos.id = ItemsXMov.idCurso AND TiposCurso.id = Cursos.idTipoCurso AND ItemsXMov.idMovimiento = Movimientos.id AND Movimientos.Saldo = 0 AND (Left(Movimientos.TipoDoc, 2) = 'FC' OR Left(Movimientos.TipoDoc, 3) = 'REC') ORDER BY Fecha DESC"
    Else
        'sSql = "SELECT Movimientos.*, TiposCurso.Detalle AS Curso, ItemsXMov.Detalle AS Descripcion FROM Movimientos, Cursos, TiposCurso, ItemsXMov WHERE NombreCliente = '" & txtNombre.Text & "' AND NOT Anulado AND Cursos.id = Movimientos.idCurso AND TiposCurso.id = Cursos.idTipoCurso AND ItemsXMov.idMovimiento = Movimientos.id ORDER BY Fecha DESC"
        sSql = "SELECT Movimientos.*, TiposCurso.Detalle AS Curso, ItemsXMov.Detalle AS Descripcion FROM Movimientos, Cursos, TiposCurso, ItemsXMov WHERE Movimientos.idAlumno = " & id_Alumno & " AND NOT Anulado AND Cursos.id = ItemsXMov.idCurso AND TiposCurso.id = Cursos.idTipoCurso AND ItemsXMov.idMovimiento = Movimientos.id AND Movimientos.Saldo = 0 AND (Left(Movimientos.TipoDoc, 2) = 'FC' OR Left(Movimientos.TipoDoc, 3) = 'REC') ORDER BY Fecha DESC"
    End If
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    lstUltimosPagos.Clear
    Do While Not adoTabla.EOF
        lstUltimosPagos.AddItem adoTabla!fecha & "     " & adoTabla!TipoDoc & "     " & Right("00000000" & adoTabla!NumDoc, 8) & "     " & adoTabla!NombreEmisor & "     " & adoTabla!curso & "     " & adoTabla!Descripcion
        adoTabla.MoveNext
    Loop
    
    adoTabla.Close
End Sub

Private Sub ASOCIACION_CON_EMPRESA()
    CERRAR_TABLA adoTemp
    sSql = "SELECT idEmpresa FROM Alumnos WHERE id = " & id_Alumno
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If adoTemp.EOF Then
        adoTemp.Close
        Exit Sub
    End If
    
    If adoTemp!idEmpresa = 0 Then
        dbgFactura.Width = 10020
        shpRectangulo.Left = 9250
        Me.Width = 10320
        txtEmpresa.Text = ""
        'txtCuit.Text = ""
        fraEmpresa.Enabled = False
        cboCondIva.Text = "Consumidor final"
    Else
        dbgFactura.Width = 12060
        shpRectangulo.Left = 11000
        Me.Width = 12345
        fraEmpresa.Enabled = True
        If MsgBox("Alumno asociado a una empresa." & vbCrLf & "¿Facturar a la empresa?", vbQuestion + vbYesNo, "FACTURACION") = vbYes Then
            FacturaPorEmpresa = True
        
            CERRAR_TABLA adoEmpresas
            sSql = "SELECT * FROM Empresas WHERE id = " & adoTemp!idEmpresa
            adoEmpresas.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            If adoEmpresas.EOF Then
                EstiloBuscador = "FacturaEmpresas"
                Extra = "Empresas"
                frmBuscador.Show vbModal
                
                If CanceloBuscador Then
                    Unload Me
                End If
                
                Extra = ""
            Else
                id_Empresa = adoEmpresas!id
                txtNombre.Text = adoEmpresas!Nombre
                txtEmpresa.Text = adoEmpresas!Nombre
                txtCuit.Text = adoEmpresas!Cuit
                cboCondIva.Text = DEVOLVER_CAMPO(adoEmpresas!idCondIva, adoCondIva, "CondIva", "Detalle")
                txtDireccion.Text = adoEmpresas!Direccion
                
                If cboCondIva.Text = "RESPONSABLE INSCRIPTO" Then
                    cboEmisor.Text = "1-Fabru S.A."
                    cboEmisor.Enabled = False
                    fraComprobantesFabru.Enabled = False
                    fraEmpresa.Enabled = False
                End If
            End If
            
            If Not CancelaBusEmpresa Then
                DOCUMENTOS_PENDIENTES
                VER_OTROS_PAGOS
            End If
        End If
    End If
    
    adoTemp.Close
End Sub

Private Sub GENERAR_RENGLONES()
    If sMenu <> "NotaCreditoPresenciales" Then
        sSql = "DELETE FROM TempFactura"
        adoConnection.Execute sSql
    End If
    
    CERRAR_TABLA adoTempFactura
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    If sMenu = "FacturaAnticipo" Then
        adoTempFactura.AddNew
    Else
        For k = 1 To 10
            adoTempFactura.AddNew
        Next
    End If
    adoTempFactura.MoveFirst
    
    ARMAR_GRILLA
End Sub

Private Sub IMPRIMIR_DOCUMENTO()
    Dim Letras As New clsNumeros
    Dim cant_copias As Byte
    Dim Fila As Byte
    Dim email As String
    
    cant_copias = GetSetting("Gestion", "Documentos", "cantCopias")
    
    If cant_copias = 0 Then
        Exit Sub
    End If
    
    Printer.ScaleMode = 4
    Printer.FontSize = 10
    Printer.Font = "Courier"
    
    'Obtengo Mail de alumno
    CERRAR_TABLA adoAlumnos
    
    sSql = "SELECT Mail FROM Alumnos WHERE id = " & id_Alumno
    adoAlumnos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    With adoAlumnos
        email = !Mail
    End With
    
    'Imprimo el encabezado
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos WHERE id = " & ultimo_id
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    For k = 1 To cant_copias
        With adoMovimientos
            'Tacho factura y pongo nota de crédito
            If sMenu = "NotaCreditoPresenciales" Or sMenu = "NotaCreditoOtros" Then
                IMPRIMIR 3, 55, "NOTA DE CRÉDITO"
                IMPRIMIR 4, 55, "XXXXXXXXXXXXXXXXXXX"
            End If
            
            IMPRIMIR 5, 60, !NumDoc
            IMPRIMIR 9, 80, !fecha
            IMPRIMIR 14, 10, !NombreCliente
            IMPRIMIR 15, 10, txtDireccion.Text
            IMPRIMIR 17, 10, cboCondIva.Text
            IMPRIMIR 17, 55, IIf(txtCuit.Text = "", "----------", txtCuit.Text)
            IMPRIMIR 19, 15, cboFormaPago.Text
            IMPRIMIR 19, 55, email
        End With
        
        'Imprimo títulos
        IMPRIMIR 22, 67, "Unitario"
        IMPRIMIR 22, 77, "% Dto."
        
        'Imprimo el cuerpo
        CERRAR_TABLA adoItemsXMov
        sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & ultimo_id
        adoItemsXMov.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        With adoItemsXMov
            .MoveFirst
            Fila = 24
            Do While Not .EOF
                IMPRIMIR Fila, 5, !Cantidad
                IMPRIMIR Fila, 9, !Detalle
                IMPRIMIR Fila, 67, Format(!Unitario, "Fixed")
                IMPRIMIR Fila, 77, Format(!Descuento, "Fixed")
                IMPRIMIR Fila, 86, Format(!Importe, "Fixed")
                
                .MoveNext
                Fila = Fila + 1
            Loop
        End With
        
        'Si pagó con tarjeta, imprimo el recargo.
        If Val(lblRecargo.Caption) > 0 Then
            IMPRIMIR Fila, 5, "1"
            IMPRIMIR Fila, 9, "RECARGO PAGO EN CUOTAS CON TARJETA DE CRÉDITO"
            IMPRIMIR Fila, 86, lblRecargo.Caption
        End If
        
        'Imprimo próximos vencimientos
        If sMenu = "FacturaPresenciales" Then
            If Fila < 29 Then
                IMPRIMIR 29, 55, "Próximos vencimientos:"
                
                Fila = 30
                
                Printer.FontSize = 8
                    
                CERRAR_TABLA adoTemp2
                sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo <> 0 ORDER BY fecha"
                adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                
                Do While Not adoTemp2.EOF
                    IMPRIMIR Fila, 60, Format(adoTemp2!fecha, "dd/mm/yyyy")
                        
                    adoTemp2.MoveNext
                    
                    Fila = Fila + 1
                Loop
                    
                Printer.FontSize = 10
            End If
        End If
        
        'If sMenu = "FacturaPresenciales" Then
        '    With adoTempFactura
        '        .MoveFirst
        '
        '        IMPRIMIR 29, 55, "Próximos vencimientos:"
        '        Fila = 30
        '        Printer.FontSize = 8
        '
        '        Do While Not .EOF
        '            If !Paga <> 0 Then
        '                sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo <> 0"
        '                CERRAR_TABLA adoTemp2
        '                adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        '                IMPRIMIR Fila, 60, adoTemp2!Fecha
        '                adoTemp2.Close
        '                Fila = Fila + 1
        '            End If
        '            .MoveNext
        '        Loop
        '
        '        Printer.FontSize = 10
        '    End With
        'End If
        
        'Imprimo próximos inicios
        With adoProximosInicios
            If Fila < 29 Then
                .Open "ProximosInicios", adoConnection, adOpenDynamic, adLockOptimistic
                
                IMPRIMIR 29, 5, "Próximos inicios:"
                Fila = 30
                Printer.FontSize = 8
                
                For j = 1 To 10
                    campo = "proximo" & j
                    IMPRIMIR Fila, 15, .Fields(campo) & ""
                    Fila = Fila + 1
                Next
                .Close
                
                Printer.FontSize = 10
            End If
        End With
  
        'Imprimo otros datos
        If Fila < 38 Then
            IMPRIMIR 38, 5, adoMovimientos!Linea1
            IMPRIMIR 39, 5, adoMovimientos!Linea2
            IMPRIMIR 40, 5, adoMovimientos!Linea3
        Else
            IMPRIMIR 39, 5, adoMovimientos!Linea1
            IMPRIMIR 40, 5, adoMovimientos!Linea2
        End If
        
        'Discrimino el IVA
        If Right(adoMovimientos!TipoDoc, 1) = "A" Then
            IMPRIMIR 41, 88, Format(adoMovimientos!Subtotal, "Fixed")
            IMPRIMIR 43, 88, Format(adoMovimientos!Iva, "Fixed")
        End If
        
        IMPRIMIR 45, 88, Format(lblTotalFactura.Caption, "Fixed")
        
        IMPRIMIR 41, 5, "SON PESOS " & Letras.NroEnLetras(lblTotalFactura.Caption)

        'Imprimo
        Printer.EndDoc
    
    Next
    
    'Cierro las tablas
    adoMovimientos.Close
    'adoItemsXMov.Close
    CERRAR_TABLA adoItemsXMov
End Sub

Private Sub CHEQUEAR_PERMITIR_CONTADO()
    optContado.Enabled = True
    
    With adoTempFactura
        If .State = adStateClosed Then
            Exit Sub
        End If
                        
        pago_matricula = True
        pago_cuota_1 = True
        
        .MoveFirst
        Do While Not .EOF
            If Left(!Detalle, 9) = "Matrícula" Then
                pago_matricula = False
            End If
            If Left(!Detalle, 7) = "Cuota 1" Then
                pago_cuota_1 = False
            End If
            
            .MoveNext
        Loop
    End With
    
    If Not pago_matricula Or Not pago_cuota_1 Then
        optContado.Enabled = True
    Else
        optContado.Enabled = False
    End If
End Sub

Private Sub FACTURAR_DESDE_INSCRIPCION()
    FacturaPorEmpresa = False
    
    cboEmisor.Enabled = True
    fraComprobantesFabru.Enabled = True
    fraEmpresa.Enabled = True
    
    CERRAR_TABLA adoAlumnos
    sSql = "SELECT * FROM Alumnos WHERE Nombre = '" & txtNombre.Text & "'"
    adoAlumnos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    id_Alumno = adoAlumnos!id
    
    txtDireccion.Text = adoAlumnos!Direccion
    
    Me.txtCuit.Text = adoAlumnos!NumDoc
    
    
    
    DOCUMENTOS_PENDIENTES
    VER_OTROS_PAGOS
    ASOCIACION_CON_EMPRESA

    If cboCondIva.Text = "RESPONSABLE INSCRIPTO" Then
        optFacturaA.Enabled = True
    Else
        optFacturaA.Enabled = False
        optFacturaB.Value = True
    End If
    
    fraModoPago.Enabled = True
    CHEQUEAR_PERMITIR_CONTADO

    CHEQUEAR_A_CUENTA_FACTURA
End Sub

Private Sub MOSTRAR_ANTICIPO()
    sSql = "SELECT Movimientos.idAlumno, Movimientos.idCurso, ItemsXMov.Importe " & _
           "FROM Movimientos, ItemsXMov " & _
           "WHERE Movimientos.id = ItemsXMov.idMovimiento AND Movimientos.idAlumno =  " & id_Alumno & _
           "  AND Left(ItemsXMov.Detalle, 9) = ' A CUENTA' AND ItemsXMov.Saldo <> 0"
    
    CERRAR_TABLA adoTemp
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoTemp.EOF Then
        total_a_cuenta = 0
        adoTemp.MoveFirst
        Do While Not adoTemp.EOF
            total_a_cuenta = total_a_cuenta + adoTemp!Importe
            adoTemp.MoveNext
        Loop
        adoTemp.Close
    End If
    
    If total_a_cuenta <> 0 Then
        shpAnticipo.Visible = True
        lblTextoAnticipo.Visible = True
        lblAnticipo.Visible = True
        
        lblAnticipo.Caption = Format(total_a_cuenta, "Fixed")
    End If
    
End Sub

Private Sub CALCULAR_RECARGO()
    'Calculo el recargo por pago con tarjeta en cuotas
    CALCULAR_TOTAL
    
    cant_cuotas = Val(cboNumCuotas.Text)
    
    sSql = "SELECT * FROM RecargosTarjeta WHERE cuotas = " & cant_cuotas
    CERRAR_TABLA adoRecargosTarjeta
    adoRecargosTarjeta.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
    Coeficiente = adoRecargosTarjeta!Coeficiente
    Comision = adoRecargosTarjeta!Comision
    
    adoRecargosTarjeta.Close
    
    TotFacturaSinTarjeta = Val(lblTotalFactura.Caption)
    TotFacturaConTarjeta = TotFacturaSinTarjeta * Coeficiente * Comision
    
    lblSubTotal.Caption = Format(Round(TotFacturaSinTarjeta), "Fixed")
    lblRecargo.Caption = Format(Round(TotFacturaConTarjeta - TotFacturaSinTarjeta), "Fixed")
    lblTotalFactura.Caption = Format(Round(TotFacturaConTarjeta), "Fixed")
End Sub
