VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRecibo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COBRANZA CUENTA CORRIENTE"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "frmRecibo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRetenciones 
      Caption         =   "Retenciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   5280
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtRetIva 
         Height          =   285
         Left            =   1440
         TabIndex        =   45
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtRetGanancias 
         Height          =   285
         Left            =   1440
         TabIndex        =   44
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtRetIngresosBrutos 
         Height          =   285
         Left            =   1440
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ganancias"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ingresos Brutos"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
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
      TabIndex        =   37
      Top             =   840
      Width           =   2415
      Begin VB.OptionButton optReciboCtaCte 
         Caption         =   "Recibo"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   855
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
      TabIndex        =   36
      Top             =   5520
      Width           =   5055
      Begin VB.TextBox txtLinea3 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtLinea2 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtLinea1 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   9240
      Picture         =   "frmRecibo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
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
      TabIndex        =   32
      Top             =   6960
      Width           =   5055
      Begin VB.ListBox lstUltimosPagos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         ItemData        =   "frmRecibo.frx":0884
         Left            =   120
         List            =   "frmRecibo.frx":0886
         TabIndex        =   33
         Top             =   240
         Width           =   4815
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
      TabIndex        =   30
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
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
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
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtEmpresa 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   5295
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox cboEmisor 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
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
      TabIndex        =   25
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
      Picture         =   "frmRecibo.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   19
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   2535
      End
   End
   Begin MSDataGridLib.DataGrid dbgFactura 
      Height          =   2775
      Left            =   120
      TabIndex        =   15
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
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         TabIndex        =   14
         Text            =   "21"
         Top             =   915
         Width           =   615
      End
      Begin VB.OptionButton optFacturaB 
         Caption         =   "Factura B"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optFacturaA 
         Caption         =   "Factura A"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   1200
         TabIndex        =   35
         Top             =   960
         Width           =   120
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
   End
   Begin VB.Label lblTotalFactura 
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
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   9000
      TabIndex        =   50
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label8 
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
      Left            =   7920
      TabIndex        =   49
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label Label6 
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
      Left            =   7920
      TabIndex        =   48
      Top             =   6120
      Width           =   990
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   9000
      TabIndex        =   47
      Top             =   6200
      Width           =   1125
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
      Left            =   8760
      TabIndex        =   46
      Top             =   5640
      Width           =   180
   End
   Begin VB.Shape shpRectangulo 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   9360
      Top             =   2400
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7800
      Top             =   6120
      Width           =   2805
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   7800
      Top             =   5640
      Width           =   2805
   End
End
Attribute VB_Name = "frmRecibo"
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

Private Sub cboEmisor_Click()
    If Left(cboEmisor.Text, 1) = 1 Then
        fraRetenciones.Visible = True
        
        fraComprobantes.Visible = False
        fraComprobantesFabru.Visible = True
        
        If cboCondIva.Text = "RESPONSABLE INSCRIPTO" Then
            optFacturaA.Enabled = True
        Else
            optFacturaA.Enabled = False
            optFacturaB.Value = True
        End If
    Else
        fraRetenciones.Visible = False
        fraComprobantes.Visible = True
        fraComprobantesFabru.Visible = False
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If Val(lblTotalFactura.Caption) = 0 Then
        MsgBox "El total del documento no puede ser 0.", vbCritical, "ERROR"
        Exit Sub
    End If
    
    'Genero el encabezado
    With adoMovimientos
        CERRAR_TABLA adoMovimientos
        .Open "Movimientos", adoConnection, adOpenDynamic, adLockOptimistic
        
        .AddNew
        
        !Sucursal = "000" & Left(cboEmisor.Text, 1)
        
        !TipoDoc = "REC"
        !NumDoc = ULTIMO_NUMERO("UltR")
        x_comprobante = "UltR"
        
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
            
            !Direccion = txtDireccion.Text & ""
            !CondIva = cboCondIva.Text & ""
            !Cuit = txtCuit.Text & ""
            !FormaPago = cboFormaPago.Text & ""
            
            !Saldo = 0
            
            '!idUnidadNegocio = DEVOLVER_ID(cboUnidadNegocio.Text, adoUnidadesNegocio, "UnidadesNegocio", "Detalle")
            !UnidadNegocio = cboUnidadNegocio.Text
            
            !Linea1 = txtLinea1.Text & ""
            !Linea2 = txtLinea2.Text & ""
            !Linea3 = txtLinea3.Text & ""
                                    
            !RetIngresosBrutos = Val(txtRetIngresosBrutos.Text)
            !RetGanancias = Val(txtRetGanancias.Text)
            !RetIva = Val(txtRetIva.Text)
            
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
                If adoTempFactura!Cantidad = 1 Then '<> 0 And adoTempFactura!Paga <> 0 Then
                    .AddNew
                    
                    !idMovimiento = ultimo_id
                    '!idCurso = id_Curso
                    !Cantidad = adoTempFactura!Cantidad
                    
                    !Detalle = adoTempFactura!Alumno & " " & adoTempFactura!Detalle
                    
                    !Unitario = adoTempFactura!Unitario
                    !Descuento = adoTempFactura!Descuento
                    '!Importe = adoTempFactura!Paga
                    '!Saldo = adoTempFactura!Saldo - adoTempFactura!Paga
                    !Importe = adoTempFactura!Importe

                    .Update
                    
                    sSql = "UPDATE ItemsXMov SET Pagada = True WHERE id = " & adoTempFactura!idItem
                    adoConnection.Execute sSql
                    
                    If sMenu = "FacturaPresenciales" Then
                        'Actualizo el saldo en el documento modelo - Movimientos
                        sSql = "SELECT * FROM Movimientos WHERE id = " & adoTempFactura!idMovimiento
                        CERRAR_TABLA adoTemp
                        adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
                        If optContado Then
                            adoTemp!Saldo = 0
                        Else
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
                            adoTemp!Saldo = adoTemp!Saldo - adoTempFactura!Paga
                        End If
                        adoTemp.Update
                        adoTemp.Close
                    End If
                End If
                adoTempFactura.MoveNext
            Loop
            
            .Close
        End With
            
        IMPRIMIR_DOCUMENTO
    End If

    Unload Me
End Sub

Private Sub dbgFactura_AfterColUpdate(ByVal ColIndex As Integer)
    If dbgFactura.Col = 2 Then 'Cantidad
        If adoTempFactura!Cantidad <> 0 And adoTempFactura!Cantidad <> 1 Then
            MsgBox "Solo se permite ingresar 0 ó 1", vbCritical, "ERROR"
            adoTempFactura!Cantidad = 0
        Else
            adoTempFactura.MoveNext
            
            CALCULAR_TOTAL
            
            adoTempFactura.MovePrevious
            dbgFactura.Col = 3
        End If
    ElseIf dbgFactura.Col = 3 Then 'Detalle
        If sMenu = "FacturaPresenciales" Then
            dbgFactura.Col = 8
        Else
            dbgFactura.Col = 4
        End If
    ElseIf dbgFactura.Col = 4 Then 'Unitario
        adoTempFactura!Importe = adoTempFactura!Cantidad * adoTempFactura!Unitario
        adoTempFactura!Saldo = adoTempFactura!Importe
        adoTempFactura!Paga = adoTempFactura!Importe
        adoTempFactura.MoveNext
        CALCULAR_TOTAL
        dbgFactura.Col = 2
    ElseIf dbgFactura.Col = 8 Then 'Paga
        If adoTempFactura!Paga <= adoTempFactura!Saldo Then
            adoTempFactura.MoveNext
            'dbgFactura.Col = 2
            CALCULAR_TOTAL
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
    End If
End Sub

Private Sub dbgFactura_DblClick()
    If sMenu <> "NotaCreditoPresenciales" Then
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
        adoTempFactura!Cantidad = 0
        adoTempFactura!Detalle = "ANULA COMPROBANTE " & det_fc
        adoTempFactura!Unitario = uni_fc
        adoTempFactura!Importe = tot_fc
        adoTempFactura.Update
    End If
End Sub

Private Sub Form_Load()
    lblFecha.Caption = Date
    
    CARGAR_COMBO "cboCondIva", adoCondIva, "CondIva", "Detalle", Me
    cboCondIva.AddItem "(No disponible)"
    
    CARGAR_COMBO "cboFormaPago", adoFormasPago, "FormasPago", "Detalle", Me
    
    CARGAR_COMBO "cboEmisor", adoEmisores, "Emisores", "Detalle", Me
    cboEmisor.Text = "1-Fabru S.A."
    
    CARGAR_COMBO "cboUnidadNegocio", adoUnidadesNegocio, "UnidadesNegocio", "Detalle", Me
    
    Select Case sMenu
        Case "FacturaPresenciales"
            cboUnidadNegocio.Text = "CURSO PRESENCIAL"
        Case "FacturaDistancia"
            cboUnidadNegocio.Text = "CURSO A DISTANCIA"
        Case "FacturaServicioTecnico"
            cboUnidadNegocio.Text = "SERVICIO TECNICO"
        Case "FacturaVentaHardware"
            cboUnidadNegocio.Text = "VENTA DE HARDWARE"
        Case "NotaCreditoPresenciales"
            optFacturaA.Caption = "Nota de crédito A"
            optFacturaB.Caption = "Nota de crédito B"
            optFacturaC.Caption = "Nota de crédito C"
    End Select
    cboUnidadNegocio.Enabled = False
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
    
    
    sSql = "UPDATE TempFactura SET Importe = Unitario - (Unitario * Descuento / 100)"
    adoConnection.Execute sSql
    
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
        
        DOCUMENTOS_PENDIENTES
        VER_OTROS_PAGOS
        
        If adoTempFactura.State = adStateOpen Then
            If adoTempFactura.EOF Then
                If MsgBox("El cliente " & txtEmpresa.Text & " no tiene documentos pendientes." & vbCrLf & "¿Desea seleccionar otro cliente?", vbYesNo + vbCritical, "DOCUMENTOS PENDIENTES") = vbYes Then
                    txtEmpresa.Text = ""
                    SendKeys "{ENTER}"
                Else
                    Unload Me
                End If
            End If
        End If
        
        fraModoPago.Enabled = True
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
    End If
End Sub

Private Sub CALCULAR_TOTAL()
    CERRAR_TABLA adoTabla
    sSql = "SELECT SUM(Importe * Cantidad) AS total FROM TempFactura"
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    Total_Factura = adoTabla!Total
    
    adoTabla.Close
    
    lblTotalFactura.Caption = Format(Total_Factura, "Fixed")
End Sub

Private Sub MOSTRAR_DOCUS_PENDIENTES()
    sSql = "DELETE FROM TempFactura"
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoTempFactura
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    With adoMovimientos
        Do While Not .EOF
            CERRAR_TABLA adoItemsXMov
            sSql = "SELECT * FROM ItemsXMov WHERE idMovimiento = " & !id & " AND NOT Pagada"
            adoItemsXMov.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
            
            
            Do While Not adoItemsXMov.EOF
                'If Left(adoItemsXMov!Detalle, 12) <> "BONIFICACION" Then
                    adoTempFactura.AddNew
    
                    adoTempFactura!idMovimiento = !id
                    
                    adoTempFactura!Cantidad = 0
                    If FacturaPorEmpresa Then
                        'adoTempFactura!Alumno = adoMovimientos!Nombre
                    End If
                    adoTempFactura!Detalle = adoItemsXMov!Detalle
                    adoTempFactura!Unitario = adoItemsXMov!Unitario
                    adoTempFactura!Descuento = adoItemsXMov!Descuento
                    adoTempFactura!Importe = adoItemsXMov!Importe
                    adoTempFactura!Saldo = adoItemsXMov!Saldo
                    adoTempFactura!SaldoReal = adoItemsXMov!Saldo
                    
                    adoTempFactura!idItem = adoItemsXMov!id
                    
                    adoTempFactura.Update
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
    
    sSql = "SELECT Movimientos.id, Movimientos.Sucursal, Movimientos.TipoDoc, Movimientos.NumDoc, Movimientos.Fecha, Movimientos.idAlumno, Movimientos.SubTotal, Movimientos.Iva, Movimientos.Total, Movimientos.Saldo, Movimientos.Descuento FROM Movimientos WHERE (TipoDoc <> 'MOD' AND TipoDoc <> 'ESP') AND NombreCliente = '" & txtEmpresa.Text & "' AND Movimientos.FormaPago = 'CUENTA CORRIENTE'"
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoMovimientos.EOF Then
        MOSTRAR_DOCUS_PENDIENTES
    Else
        adoMovimientos.Close
    End If
End Sub


Private Sub DOCUMENTOS_EMITIDOS()
    CERRAR_TABLA adoMovimientos
    
    If FacturaPorEmpresa Then
        sSql = "SELECT Movimientos.id, Movimientos.Sucursal, Movimientos.TipoDoc, Movimientos.NumDoc, Movimientos.Fecha, Movimientos.idAlumno, Movimientos.SubTotal, Movimientos.Iva, Movimientos.Total, Movimientos.Saldo, Movimientos.Descuento, Alumnos.Nombre FROM Movimientos, Alumnos WHERE Left(TipoDoc, 2) = 'FC' AND Alumnos.idEmpresa = " & id_Empresa & " AND Movimientos.idAlumno = Alumnos.id"
    Else
        sSql = "SELECT * FROM Movimientos WHERE Left(TipoDoc,2) = 'FC' AND idAlumno = " & id_Alumno
    End If
    adoMovimientos.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If Not adoMovimientos.EOF Then
        MOSTRAR_DOCUS_EMITIDOS
    Else
        adoMovimientos.Close
    End If
End Sub

Private Sub MOSTRAR_DOCUS_EMITIDOS()
    sSql = "DELETE FROM TempFactura"
    adoConnection.Execute sSql
    
    CERRAR_TABLA adoTempFactura
    adoTempFactura.Open "TempFactura", adoConnection, adOpenKeyset, adLockOptimistic
    
    With adoMovimientos
        Do While Not .EOF
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
    dbgFactura.Columns(3).Locked = True
    
    dbgFactura.Columns(4).Caption = "Unit."
    dbgFactura.Columns(4).Width = 800
    dbgFactura.Columns(4).Alignment = dbgRight
    dbgFactura.Columns(4).NumberFormat = "Fixed"
    dbgFactura.Columns(4).Locked = True
    
    dbgFactura.Columns(5).Caption = "% Dto."
    dbgFactura.Columns(5).Width = 800
    dbgFactura.Columns(5).Alignment = dbgRight
    dbgFactura.Columns(5).NumberFormat = "Fixed"
    dbgFactura.Columns(5).Locked = True
    
    dbgFactura.Columns(6).Width = 800
    dbgFactura.Columns(6).Alignment = dbgRight
    dbgFactura.Columns(6).NumberFormat = "Fixed"
    dbgFactura.Columns(6).Locked = True
    
    dbgFactura.Columns(7).Visible = False
    dbgFactura.Columns(8).Visible = False
    dbgFactura.Columns(9).Visible = False
    dbgFactura.Columns(10).Visible = False
    dbgFactura.Columns(11).Visible = False
    
    If Not adoTempFactura.EOF Then
        adoTempFactura.MoveFirst
    End If
End Sub

Private Sub VER_OTROS_PAGOS()
    CERRAR_TABLA adoTabla
    If FacturaPorEmpresa Then
        sSql = "SELECT * FROM Movimientos WHERE NombreCliente = '" & txtEmpresa.Text & "' ORDER BY Fecha DESC"
    Else
        sSql = "SELECT * FROM Movimientos WHERE NombreCliente = '" & txtNombre.Text & "' ORDER BY Fecha DESC"
    End If
    adoTabla.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    lstUltimosPagos.Clear
    Do While Not adoTabla.EOF
        lstUltimosPagos.AddItem adoTabla!fecha & "     " & adoTabla!TipoDoc & "     " & adoTabla!NombreEmisor
        adoTabla.MoveNext
    Loop
    
    adoTabla.Close
End Sub

Private Sub ASOCIACION_CON_EMPRESA()
    CERRAR_TABLA adoTemp
    sSql = "SELECT idEmpresa FROM Alumnos WHERE id = " & id_Alumno
    adoTemp.Open sSql, adoConnection, adOpenDynamic, adLockOptimistic
    
    If adoTemp!idEmpresa = 0 Then
        dbgFactura.Width = 10020
        shpRectangulo.Left = 9250
        Me.Width = 10320
        txtEmpresa.Text = ""
        txtCuit.Text = ""
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
                Extra = ""
            Else
                id_Empresa = adoEmpresas!id
                'txtNombre.Text = adoEmpresas!Nombre
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

            DOCUMENTOS_PENDIENTES
            VER_OTROS_PAGOS
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
    
    For k = 1 To 10
        adoTempFactura.AddNew
    Next
    adoTempFactura.MoveFirst
    
    ARMAR_GRILLA
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
        
        'Imprimo retenciones
        IMPRIMIR Fila, 9, "Retención Ingresos Brutos"
        IMPRIMIR Fila, 86, Format(Val(txtRetIngresosBrutos.Text), "Fixed")
        Fila = Fila + 1
        IMPRIMIR Fila, 9, "Retención Ganancias"
        IMPRIMIR Fila, 86, Format(Val(txtRetGanancias.Text), "Fixed")
        Fila = Fila + 1
        IMPRIMIR Fila, 9, "Retención IVA"
        IMPRIMIR Fila, 86, Format(Val(txtRetIva.Text), "Fixed")
        
        
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
    CERRAR_TABLA adoMovimientos
    CERRAR_TABLA adoItemsXMov
End Sub

Private Sub CHEQUEAR_PERMITIR_CONTADO()
    optContado.Enabled = True
    
    With adoTempFactura
        .MoveFirst
        Do While Not .EOF
            If Left(!Detalle, 9) = "Matrícula" And !Saldo <> 0 And !Saldo <> !Importe Then
                optContado.Enabled = False
                Exit Do
            End If
            If Left(!Detalle, 7) = "Cuota 1" And !Saldo <> 0 And !Saldo <> !Importe Then
                optContado.Enabled = False
                Exit Do
            End If
            
            .MoveNext
        Loop
    End With
End Sub

Private Sub txtRetGanancias_Change()
    CALCULAR_TOTAL
    lblTotalFactura.Caption = Format(Val(lblTotalFactura.Caption) - Val(txtRetIngresosBrutos.Text) - Val(txtRetGanancias.Text) - Val(txtRetIva.Text), "Fixed")
End Sub

Private Sub txtRetIngresosBrutos_Change()
    CALCULAR_TOTAL
    lblTotalFactura.Caption = Format(Val(lblTotalFactura.Caption) - Val(txtRetIngresosBrutos.Text) - Val(txtRetGanancias.Text) - Val(txtRetIva.Text), "Fixed")
End Sub

Private Sub txtRetIva_Change()
    CALCULAR_TOTAL
    lblTotalFactura.Caption = Format(Val(lblTotalFactura.Caption) - Val(txtRetIngresosBrutos.Text) - Val(txtRetGanancias.Text) - Val(txtRetIva.Text), "Fixed")
End Sub
