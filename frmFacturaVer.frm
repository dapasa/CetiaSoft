VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFacturaVer 
   Caption         =   "DOCUMENTO - Consultar"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReimprimir 
      Caption         =   "&Reimprimir"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8280
      Picture         =   "frmFacturaVer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Cancelar"
      Top             =   6960
      Width           =   855
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
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton optFacturaA 
         Caption         =   "Factura A"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optFacturaB 
         Caption         =   "Factura B"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         TabIndex        =   30
         Text            =   "21"
         Top             =   915
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1200
         TabIndex        =   33
         Top             =   960
         Width           =   120
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
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdVerFichaAlumno 
         Height          =   375
         Left            =   2280
         Picture         =   "frmFacturaVer.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   180
         Width           =   375
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
      TabIndex        =   23
      Top             =   840
      Width           =   5175
      Begin VB.TextBox txtDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   24
         Top             =   240
         Width           =   4935
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
      TabIndex        =   21
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
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
      TabIndex        =   19
      Top             =   1560
      Width           =   2535
      Begin VB.ComboBox cboCondIva 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Text            =   "cboCondIva"
         Top             =   240
         Width           =   2295
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
      TabIndex        =   17
      Top             =   1560
      Width           =   2535
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   18
         Text            =   "cboFormaPago"
         Top             =   240
         Width           =   2295
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
      TabIndex        =   15
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox cboEmisor 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   16
         Text            =   "cboEmisor"
         Top             =   240
         Width           =   2175
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
      TabIndex        =   13
      Top             =   840
      Width           =   2175
      Begin VB.ComboBox cboUnidadNegocio 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   14
         Text            =   "cboUnidadNegocio"
         Top             =   240
         Width           =   1935
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
      TabIndex        =   11
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtEmpresa 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   12
         Top             =   240
         Width           =   2535
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
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
      Begin VB.OptionButton optCuotas 
         Caption         =   "Cuotas"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   800
      End
      Begin VB.OptionButton optContado 
         Caption         =   "Contado"
         Height          =   255
         Left            =   980
         TabIndex        =   9
         Top             =   240
         Width           =   900
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
      TabIndex        =   5
      Top             =   840
      Width           =   2415
      Begin VB.OptionButton optReciboC 
         Caption         =   "Recibo"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optFacturaC 
         Caption         =   "Factura C"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cerrar"
      Height          =   855
      Left            =   9240
      Picture         =   "frmFacturaVer.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   6960
      Width           =   855
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
      TabIndex        =   0
      Top             =   5520
      Width           =   7095
      Begin VB.TextBox txtLinea1 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   6855
      End
      Begin VB.TextBox txtLinea2 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   6855
      End
      Begin VB.TextBox txtLinea3 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   6855
      End
   End
   Begin MSDataGridLib.DataGrid dbgFactura 
      Height          =   2775
      Left            =   120
      TabIndex        =   28
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
   Begin VB.Label lblNumTicket 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-----------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   8880
      TabIndex        =   41
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblNumeroTicket 
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
      TabIndex        =   40
      Top             =   6480
      Width           =   1290
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
      TabIndex        =   39
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Shape shpRectangulo 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   9360
      Top             =   2400
      Width           =   735
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
      TabIndex        =   38
      Top             =   5520
      Width           =   705
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
      TabIndex        =   37
      Top             =   6000
      Width           =   990
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
      TabIndex        =   36
      Top             =   6000
      Width           =   1515
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
      TabIndex        =   35
      Top             =   5520
      Width           =   180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7320
      Top             =   5520
      Width           =   2805
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7320
      Top             =   6000
      Width           =   2805
   End
   Begin VB.Shape shpNumTicket 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7320
      Top             =   6480
      Width           =   2805
   End
End
Attribute VB_Name = "frmFacturaVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim item_actual As Byte
Dim FacturaPorEmpresa As Boolean
Dim Descuento_Ex As Single
Dim Descuento_Contado As Single
Dim Descuento_Total As Single
Dim Total_Factura As Single
Dim Total_Descuento As Single
Dim ultimo_id As Long
' Dim ultimo_id As Integer

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdReimprimir_Click()
    REIMPRIMIR
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        cmdReimprimir.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    sSql = "DELETE FROM TempFactura"
    adoConnection.Execute sSql
    
    cboCondIva.AddItem "(No disponible)"
    
    ARMAR_DOCUMENTO
End Sub

Private Sub CALCULAR_TOTAL()
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Paga) AS total FROM TempFactura"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    Total_Factura = adoTabla!Total
    
    adoTabla.Close
    
    lblTotalFactura.Caption = Format(Total_Factura, "Fixed")
End Sub

Private Sub ARMAR_GRILLA()
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

Private Sub IMPRIMIR_DOCUMENTO()
    Dim Letras As New clsNumeros
    Dim cant_copias As Byte
    Dim Fila As Byte
    
    cant_copias = GetSetting("Gestion", "Documentos", "cantCopias")
        
    Printer.ScaleMode = 4
    Printer.FontSize = 10
    Printer.Font = "Courier"
    
    'Imprimo el encabezado
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos WHERE id = " & ultimo_id
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    For k = 1 To cant_copias
        With adoMovimientos
            'Tacho factura y pongo nota de crédito
            If sMenu = "NotaCreditoPresenciales" Then
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
        
        'Imprimo próximos vencimientos
        If sMenu = "FacturaPresenciales" Then
            IMPRIMIR 29, 55, "Próximos vencimientos:"
            
            Fila = 30
            
            Printer.FontSize = 8
                
            CERRAR_TABLA adoTemp2
            sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo <> 0 ORDER BY fecha"
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            Do While Not adoTemp2.EOF
                IMPRIMIR Fila, 60, adoTemp2!fecha
                    
                adoTemp2.MoveNext
                
                Fila = Fila + 1
            Loop
                
            Printer.FontSize = 10
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
        End With
  
        'Imprimo otros datos
        IMPRIMIR 38, 5, adoMovimientos!Linea1
        IMPRIMIR 39, 5, adoMovimientos!Linea2
        IMPRIMIR 40, 5, adoMovimientos!Linea3
        
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

Private Sub ARMAR_DOCUMENTO()
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos WHERE id = " & id_Movimiento_Consultar
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    With adoMovimientos
        'IMPRIMIR 5, 60, !NumDoc
        lblFecha.Caption = !fecha
        txtNombre.Text = !NombreCliente
        txtDireccion.Text = !Direccion
        
        cboCondIva.Text = !CondIva
        
        txtCuit.Text = !Cuit
        cboEmisor.Text = !NombreEmisor
        cboUnidadNegocio.Text = !UnidadNegocio
        
        If Left(!TipoDoc, 1) = "F" Then
            If Right(!TipoDoc, 1) = "A" Then
                fraComprobantesFabru.Visible = True
                fraComprobantes.Visible = False
                optFacturaA.Value = True
            ElseIf Right(!TipoDoc, 1) = "B" Then
                fraComprobantesFabru.Visible = True
                fraComprobantes.Visible = False
                optFacturaB.Value = True
            Else
                fraComprobantesFabru.Visible = False
                fraComprobantes.Visible = True
                optFacturaC.Value = True
            End If
        Else
            optReciboC.Value = True
        End If
        
        cboFormaPago.Text = !FormaPago
    End With
    
    CERRAR_TABLA adoItemsXMov
    sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & id_Movimiento_Consultar
    adoItemsXMov.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    sSql = "DELETE FROM TempFactura"
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoTempFactura
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
        
    With adoItemsXMov
        .MoveFirst
        
        Do While Not .EOF
            adoTempFactura.AddNew
            
            adoTempFactura!Cantidad = !Cantidad
            adoTempFactura!Detalle = !Detalle
            adoTempFactura!Unitario = Format(!Unitario, "Fixed")
            adoTempFactura!Descuento = Format(!Descuento, "Fixed")
            adoTempFactura!Importe = Format(!Importe, "Fixed")
            
            adoTempFactura.Update
            
            .MoveNext
        Loop
    End With
    
    'Si es pago con tarjeta muestro el renglón de recargo.
    If adoMovimientos!Recargo > 0 Then
        adoTempFactura.AddNew
        
        adoTempFactura!Cantidad = 1
        adoTempFactura!Detalle = "RECARGO PAGO EN CUOTAS CON TARJETA DE CRÉDITO"
        adoTempFactura!Unitario = Format(adoMovimientos!Recargo, "Fixed")
        adoTempFactura!Descuento = Format(0, "Fixed")
        adoTempFactura!Importe = Format(adoMovimientos!Recargo, "Fixed")
        
        adoTempFactura.Update
    End If
        
    ARMAR_GRILLA
        
    txtLinea1.Text = adoMovimientos!Linea1
    txtLinea2.Text = adoMovimientos!Linea2
    txtLinea3.Text = adoMovimientos!Linea3
    
    lblNumTicket.Caption = adoMovimientos!NumTicket & ""
    
    'Discrimino el IVA
    If Right(adoMovimientos!TipoDoc, 1) = "A" Then
        IMPRIMIR 41, 88, Format(adoMovimientos!Subtotal, "Fixed")
        IMPRIMIR 43, 88, Format(adoMovimientos!Iva, "Fixed")
    End If
    
    lblTotalFactura.Caption = Format(adoMovimientos!Total, "Fixed")
End Sub

Private Sub REIMPRIMIR()
    'VER DE UTILIZAR EL MISMO REIMPRIMIR QUE TIENE BUSDOC
    
    Dim Letras As New clsNumeros
    Dim cant_copias As Byte
    Dim Fila As Byte
    
    cant_copias = 1 'GetSetting("Gestion", "Documentos", "cantCopias")
        
    Printer.ScaleMode = 4
    Printer.FontSize = 10
    Printer.Font = "Courier"
    
    'Imprimo el encabezado
    CERRAR_TABLA adoMovimientos
    sSql = "SELECT * FROM Movimientos WHERE id = " & adoTabla!id
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    For k = 1 To cant_copias
        With adoMovimientos
            'Tacho factura y pongo nota de crédito
            If sMenu = "NotaCreditoPresenciales" Then
                IMPRIMIR 3, 55, "NOTA DE CRÉDITO"
                IMPRIMIR 4, 55, "XXXXXXXXXXXXXXXXXXX"
            End If
            
            IMPRIMIR 5, 60, !NumDoc
            IMPRIMIR 9, 80, !fecha
            IMPRIMIR 14, 10, !NombreCliente
            IMPRIMIR 15, 10, !Direccion
            IMPRIMIR 17, 10, !CondIva
            IMPRIMIR 17, 55, IIf(!Cuit = "", "----------", !Cuit)
            IMPRIMIR 19, 15, !FormaPago
        End With
        
        'Imprimo títulos
        IMPRIMIR 22, 67, "Unitario"
        IMPRIMIR 22, 77, "% Dto."
        
        'Imprimo el cuerpo
        CERRAR_TABLA adoItemsXMov
        sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & adoTabla!id
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
        
        'Imprimo próximos vencimientos.
        CERRAR_TABLA adoTemp2
        sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & adoMovimientos!idAlumno & " AND TipoDoc = 'MOD' AND Saldo <> 0 ORDER BY fecha"
        adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
        
        If Not adoTemp2.EOF Then
            IMPRIMIR 29, 55, "Próximos vencimientos:"
            
            Fila = 30
            
            Printer.FontSize = 8
                
            
            Do While Not adoTemp2.EOF
                IMPRIMIR Fila, 60, Format(adoTemp2!fecha, "dd/mm/yy")
                    
                adoTemp2.MoveNext
                
                Fila = Fila + 1
            Loop
                
            Printer.FontSize = 10
        End If
        
        adoTemp2.Close
        
        'Imprimo próximos vencimientos
        If sMenu = "FacturaPresenciales" Then
            IMPRIMIR 29, 55, "Próximos vencimientos:"
            
            Fila = 30
            
            Printer.FontSize = 8
                
            CERRAR_TABLA adoTemp2
            sSql = "SELECT Fecha FROM Movimientos WHERE idAlumno = " & id_Alumno & " AND TipoDoc = 'MOD' AND Saldo <> 0 ORDER BY fecha"
            adoTemp2.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            Do While Not adoTemp2.EOF
                IMPRIMIR Fila, 60, adoTemp2!fecha
                    
                adoTemp2.MoveNext
                
                Fila = Fila + 1
            Loop
                
            Printer.FontSize = 10
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
        End With
  
        'Imprimo otros datos
        IMPRIMIR 38, 5, adoMovimientos!Linea1
        IMPRIMIR 39, 5, adoMovimientos!Linea2
        IMPRIMIR 40, 5, adoMovimientos!Linea3
        
        'Discrimino el IVA
        If Right(adoMovimientos!TipoDoc, 1) = "A" Then
            IMPRIMIR 41, 88, Format(adoMovimientos!Subtotal, "Fixed")
            IMPRIMIR 43, 88, Format(adoMovimientos!Iva, "Fixed")
        End If
        
        IMPRIMIR 45, 88, Format(adoMovimientos!Total, "Fixed")
        
        IMPRIMIR 41, 5, "SON PESOS " & Letras.NroEnLetras(adoMovimientos!Total)

        'Imprimo
        Printer.EndDoc
    
    Next
    
    'Cierro las tablas
    adoMovimientos.Close
    'adoItemsXMov.Close
    CERRAR_TABLA adoItemsXMov

End Sub

