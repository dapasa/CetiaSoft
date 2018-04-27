VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmLiquidacionProfesor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LIQUIDAR PAGO"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLiquidacionProfesor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir grilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      Picture         =   "frmLiquidacionProfesor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Lista de alumnos"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelarPagoClase 
      Caption         =   "&Cancelar pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      Picture         =   "frmLiquidacionProfesor.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Guardar"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdPagarClase 
      Caption         =   "&Pagar clase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10080
      Picture         =   "frmLiquidacionProfesor.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Guardar"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pagar a cuenta"
      Height          =   615
      Left            =   7320
      TabIndex        =   16
      Top             =   240
      Width           =   1575
      Begin VB.TextBox txtPago 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   570
         TabIndex        =   18
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Picture         =   "frmLiquidacionProfesor.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdPagar 
      BackColor       =   &H0000FF00&
      Caption         =   "  P  A  G  A  R  T  O  D  O"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Guardar"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Calcular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      Picture         =   "frmLiquidacionProfesor.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Guardar"
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Caption         =   "Total a pagar"
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   1575
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   510
         TabIndex        =   12
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Valor hora"
      Height          =   615
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtValorHora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hasta"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1575
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   17563649
         CurrentDate     =   39371
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Desde"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   17563649
         CurrentDate     =   39371
      End
   End
   Begin MSDataGridLib.DataGrid dbgClasesDictadas 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5953
      _Version        =   393216
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame8 
      Caption         =   "Profesor"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cboProfesor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Line Line2 
      X1              =   7200
      X2              =   7200
      Y1              =   1440
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   1440
      Y2              =   240
   End
End
Attribute VB_Name = "frmLiquidacionProfesor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id_curso_ As String

Private Sub cmdCalcular_Click()

    cant_clases = 0
    With adoTabla
        .MoveFirst
        
        Do While Not .EOF
            If !Situacion = "OK" And !Pagada = "NO" Then
                cant_clases = cant_clases + 1
            End If
            .MoveNext
        Loop
    End With

    lblTotal.Caption = Val(txtValorHora.Text) * cant_clases * 3
    
    txtPago.Text = lblTotal.Caption
End Sub

Private Sub cmdCancelarPagoClase_Click()
    If adoTabla!Situacion = "OK" And adoTabla!Pagada = "SI" Then
        id_clase = adoTabla!id
            
        adoTabla.Close
        
        sSql = "UPDATE ClasesXCurso SET Pagada = 'NO' WHERE id = " & id_clase
        adoConnection.Execute sSql
        
        sSql = "UPDATE ClasesXCurso SET FechaPago = '' WHERE id = " & id_clase
        adoConnection.Execute sSql
    End If
    
    ARMAR_TABLA
    ARMAR_GRILLA
End Sub

Private Sub cmdImprimir_Click()
    Dim Fila As Byte
    
    Printer.ScaleMode = 4
    
    With adoTabla
        .MoveFirst
        
        IMPRIMIR 1, 5, "LIQUIDACIÓN HORAS - " & cboProfesor.Text
        IMPRIMIR 2, 5, "DESDE: " & dtpDesde.Value & " HASTA: " & dtpHasta.Value
        IMPRIMIR 3, 5, "Valor hora: $" & txtValorHora.Text
        IMPRIMIR 4, 5, "Total a pagar: $" & txtPago.Text
        
        
        IMPRIMIR 6, 5, "Fecha"
        IMPRIMIR 6, 18, "Nº"
        IMPRIMIR 6, 23, "Curso"
        IMPRIMIR 6, 56, "Horario"
        IMPRIMIR 6, 79, "Situación"
        IMPRIMIR 6, 92, "Pagada"
        IMPRIMIR 6, 96, "Fecha pago"
        
        IMPRIMIR 7, 5, "===================================================================================================="
        
        Fila = 8
        Do While Not .EOF
            
            IMPRIMIR Fila, 5, .Fields(0).Value
            IMPRIMIR Fila, 14, .Fields(1).Value
            IMPRIMIR Fila, 20, Left(.Fields(3).Value, 23)
            IMPRIMIR Fila, 45, Left(.Fields(5).Value, 20)
            IMPRIMIR Fila, 60, .Fields(6).Value & ""
            IMPRIMIR Fila, 70, .Fields(7).Value & ""
            IMPRIMIR Fila, 80, .Fields(8).Value & ""
            
            Fila = Fila + 1
            .MoveNext
        Loop
    End With
    
    Printer.EndDoc

End Sub

Private Sub cmdPagar_Click()
    x_pesos = Val(txtPago.Text)
    
    With adoTabla
        .MoveFirst
        Do While Not .EOF
            If !Situacion = "OK" And !Pagada = "NO" Then
                
                If x_pesos >= Val(txtValorHora.Text) * 3 Then
            
                    id_clase = adoTabla!id
                
                    sSql = "UPDATE ClasesXCurso SET Pagada = 'SI' WHERE id = " & id_clase
                    adoConnection.Execute sSql
                
                    sSql = "UPDATE ClasesXCurso SET FechaPago = '" & Format(Date, "dd/mm/yyyy") & "' WHERE id = " & id_clase
                    adoConnection.Execute sSql
                
                    x_pesos = x_pesos - (Val(txtValorHora.Text) * 3)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    ARMAR_TABLA
    ARMAR_GRILLA
End Sub

Private Sub cmdPagarClase_Click()
    If adoTabla!Situacion = "OK" And adoTabla!Pagada = "NO" Then
        id_clase = adoTabla!id
            
        adoTabla.Close
        
        sSql = "UPDATE ClasesXCurso SET Pagada = 'SI' WHERE id = " & id_clase
        adoConnection.Execute sSql
        
        sSql = "UPDATE ClasesXCurso SET FechaPago = '" & Format(Date, "dd/mm/yyyy") & "' WHERE id = " & id_clase
        adoConnection.Execute sSql
    End If
    
    ARMAR_TABLA
    ARMAR_GRILLA
End Sub

Private Sub cmdVer_Click()
    ARMAR_TABLA
    ARMAR_GRILLA
End Sub


Private Sub Form_Load()
    CARGAR_COMBO "cboProfesor", adoProfesores, "Profesores", "Nombre", Me
    
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub

Private Sub ARMAR_TABLA()
    With adoTabla
        CERRAR_TABLA adoTabla
        sSql = "SELECT ClasesXCurso.Fecha, ClasesXCurso.Numero, Cursos.idTipoCurso, TiposCurso.Detalle, Cursos.idHorario, Horarios.Detalle, ClasesXCurso.Situacion, ClasesXCurso.Pagada, ClasesXCurso.FechaPago, ClasesXCurso.id FROM ClasesXCurso, Cursos, TiposCurso, Horarios " & _
               "   WHERE Cursos.id = ClasesXCurso.idCurso AND TiposCurso.id = Cursos.idTipoCurso AND " & _
               "         Cursos.id = ClasesXCurso.idCurso AND Horarios.id = Cursos.idHorario AND " & _
               "         ClasesXCurso.Fecha >= DateValue('" & dtpDesde.Value & "') AND ClasesXCurso.Fecha <= DateValue('" & dtpHasta.Value & "') AND " & _
               "         ClasesXCurso.Profesor = '" & cboProfesor.Text & "' " & _
               "   ORDER BY ClasesXCurso.Pagada, ClasesXCurso.Fecha"
        .Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
    End With
End Sub

Private Sub ARMAR_GRILLA()
        Set dbgClasesDictadas.DataSource = adoTabla
        
        dbgClasesDictadas.Columns(0).Caption = "Fecha"
        dbgClasesDictadas.Columns(0).Width = 1000
        
        dbgClasesDictadas.Columns(1).Caption = "Nº"
        dbgClasesDictadas.Columns(1).Width = 400
        
        dbgClasesDictadas.Columns(2).Visible = False
        
        dbgClasesDictadas.Columns(3).Caption = "Curso"
        dbgClasesDictadas.Columns(3).Width = 2900
        
        dbgClasesDictadas.Columns(4).Visible = False
        
        dbgClasesDictadas.Columns(5).Caption = "Horario"
        dbgClasesDictadas.Columns(5).Width = 1800
        
        dbgClasesDictadas.Columns(6).Caption = "Situación"
        dbgClasesDictadas.Columns(6).Width = 1200
        
        dbgClasesDictadas.Columns(7).Caption = "Pagada"
        dbgClasesDictadas.Columns(7).Width = 800
        
        dbgClasesDictadas.Columns(8).Caption = "Fecha pago"
        dbgClasesDictadas.Columns(8).Width = 1000
        
        dbgClasesDictadas.Columns(9).Visible = False
End Sub
