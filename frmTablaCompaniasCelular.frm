VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmTablaCompaniasCelular 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABLAS - Companías de telefonía celular"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   Icon            =   "frmTablaCompaniasCelular.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Detalle y dominio"
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
      TabIndex        =   4
      Top             =   120
      Width           =   3655
      Begin VB.TextBox txtDominio 
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDetalle 
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   615
      Left            =   4560
      Picture         =   "frmTablaCompaniasCelular.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir"
      Top             =   2280
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dbgTablaCompaniasCelular 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3201
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Modifique directamente en la tabla"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2430
   End
End
Attribute VB_Name = "frmTablaCompaniasCelular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    On Error GoTo ErrorHandle
    
    If txtDetalle.Text = "" Then
        MsgBox "Debe ingresar una companía de telefonía celular", vbCritical, "Ingresar - Companías de telefonía celular"
        Detalle.SetFocus
    Else
        CERRAR_TABLA adoCompaniasCelular
        sSql = "SELECT * FROM CompaniasCelular WHERE Detalle = '" & txtDetalle.Text & "'"
        adoCompaniasCelular.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
        
        If adoCompaniasCelular.EOF Then
            adoCompaniasCelular.AddNew
            adoCompaniasCelular!Detalle = txtDetalle.Text
            adoCompaniasCelular!Dominio = txtDominio.Text
            adoCompaniasCelular.Update
            adoCompaniasCelular.MoveFirst
            adoCompaniasCelular.MoveLast
            txtDetalle.Text = ""
            txtDominio.Text = ""
            txtDetalle.SetFocus
            cmdEliminar.Enabled = True
        Else
            MsgBox "El dato que intenta ingresar ya existe.", vbCritical, "ERROR - Dato existente"
        End If
        
        adoCompaniasCelular.Close
        
        INICIALIZAR_FORMULARIO
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTablaCompaniasCelular - cmdAgregar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdCerrar_Click()
    On Error GoTo ErrorHandle
    
    Unload Me
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTablaCompaniasCelular - cmdCerrar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo ErrorHandle
    
    If MsgBox("¿Confirma la eliminación?", vbYesNo, "Eliminar - Companía de telefonía celular") = vbYes Then
        frmClave.Show vbModal
        If Not AccesoPermitido Then
            MsgBox "La clave no es correcta." & vbCrLf & "La operación no ha sido realizada.", vbCritical, "ACCESO DENEGADO"
            Exit Sub
        End If
        
        If PUEDE_BORRAR("CompaniasCelular", adoCompaniasCelular!id) Then
            adoCompaniasCelular.Delete
        End If
    End If
    
    If adoCompaniasCelular.RecordCount = 0 Then
        cmdEliminar.Enabled = False
    Else
        cmdEliminar.Enabled = True
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTablaCompaniasCelular - cmdEliminar", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    
    INICIALIZAR_FORMULARIO
    
    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTablaCompaniasCelular - Form.Load", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If adoCompaniasCelular.EditMode = adEditInProgress Then
        adoCompaniasCelular.Update
    End If
    adoCompaniasCelular.Close
    strTabla = ""
End Sub

Private Sub txtDetalle_Change()
    If Len(txtDetalle.Text) > 0 And Len(txtDominio.Text) > 0 Then
        cmdAgregar.Enabled = True
    Else
        cmdAgregar.Enabled = False
    End If
End Sub

Private Sub txtDetalle_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If KeyAscii = 13 Then
        txtDominio.SetFocus
    End If

    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTablaCompaniasCelular - txtDetalle.KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub txtDominio_Change()
    If Len(txtDetalle.Text) > 0 And Len(txtDominio.Text) > 0 Then
        cmdAgregar.Enabled = True
    Else
        cmdAgregar.Enabled = False
    End If
End Sub

Private Sub txtDominio_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    
    If KeyAscii = 13 Then
        cmdAgregar_Click
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTablaCompaniasCelular - txtDominio.KeyPress", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub

Private Sub INICIALIZAR_FORMULARIO()
    On Error GoTo ErrorHandle
    
    sSql = "SELECT * FROM CompaniasCelular WHERE Detalle <> '(No disponible)' ORDER BY Detalle"
    adoCompaniasCelular.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
    Set dbgTablaCompaniasCelular.DataSource = adoCompaniasCelular
    
    With dbgTablaCompaniasCelular
        .Columns(0).Visible = False
        .Columns(1).Caption = "Detalle"
        .Columns(1).Width = 1500
        .Columns(1).Locked = True
        .Columns(2).Caption = "Dominio"
        .Columns(2).Width = 1500
        .Columns(3).Visible = False
    End With
    
    If adoCompaniasCelular.RecordCount = 0 Then
        cmdEliminar.Enabled = False
    Else
        cmdEliminar.Enabled = True
        adoCompaniasCelular.MoveLast
        adoCompaniasCelular.MoveFirst
    End If

    Exit Sub
ErrorHandle:
    MsgBox "Tome nota:" & vbCrLf & "ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & "frmTablaCompaniasCelular - INICIALIZAR_FORMULARIO", vbCritical, "SE HA PRODUCIDO UN ERROR"
End Sub
