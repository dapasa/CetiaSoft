VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mNombreTabla As String
Private madoTabla As ADODB.Recordset
Private madoConexion As ADODB.Connection
Private mTituloPlural As String
Private mTituloSingular As String
Private mSqlListado As String
Private mGrilla As DataGrid
Private mDetalle As TextBox
Private mbotonAgregar As CommandButton
Private mbotonEliminar As CommandButton
Private mbotonCerrar As CommandButton

Public Property Get NombreTabla() As String
    NombreTabla = mNombreTabla
End Property

Public Property Let NombreTabla(ByVal vNewValue As String)
    mNombreTabla = vNewValue
End Property

Public Property Get adoTabla() As ADODB.Recordset
    Set adoTabla = madoTabla
End Property

Public Property Let adoTabla(ByVal vNewValue As ADODB.Recordset)
    Set madoTabla = vNewValue
End Property

Public Property Get adoConexion() As ADODB.Connection
    Set adoConexion = madoConexion
End Property

Public Property Let adoConexion(ByVal vNewValue As ADODB.Connection)
    Set madoConexion = vNewValue
End Property

Public Property Get TituloPlural() As String
    TituloPlural = mTituloPlural
End Property

Public Property Let TituloPlural(ByVal vNewValue As String)
    mTituloPlural = vNewValue
End Property

Public Property Get TituloSingular() As String
    TituloSingular = mTituloSingular
End Property

Public Property Let TituloSingular(ByVal vNewValue As String)
    mTituloSingular = vNewValue
End Property

Public Property Get SqlListado() As String
    SqlListado = mSqlListado
End Property

Public Property Let SqlListado(ByVal vNewValue As String)
    mSqlListado = vNewValue
End Property

Public Property Get Grilla() As DataGrid
    Set Grilla = mGrilla
End Property

Public Property Let Grilla(ByVal vNewValue As DataGrid)
    Set mGrilla = vNewValue
End Property

Public Property Get Detalle() As TextBox
    Set Detalle = mDetalle
End Property

Public Property Let Detalle(ByVal vNewValue As TextBox)
    Set mDetalle = vNewValue
End Property

Public Property Get botonAgregar() As CommandButton
    Set botonAgregar = mbotonAgregar
End Property

Public Property Let botonAgregar(ByVal vNewValue As CommandButton)
    Set mbotonAgregar = vNewValue
End Property

Public Property Get botonEliminar() As CommandButton
    Set botonEliminar = mbotonEliminar
End Property

Public Property Let botonEliminar(ByVal vNewValue As CommandButton)
    Set mbotonEliminar = vNewValue
End Property

Public Property Get botonCerrar() As CommandButton
    Set botonCerrar = mbotonCerrar
End Property

Public Property Let botonCerrar(ByVal vNewValue As CommandButton)
    Set mbotonCerrar = vNewValue
End Property

Public Sub Inicializar()
    sSql = "SELECT * FROM " & NombreTabla & " WHERE Detalle <> '(No disponible)' ORDER BY Detalle"
    adoTabla.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
    Set Grilla.DataSource = adoTabla
    
    With Grilla
        .Columns(0).Visible = False
        .Columns(1).Caption = "Detalle"
        .Columns(1).Width = 3100
        .Columns(2).Visible = False
    End With
    
    If adoTabla.RecordCount = 0 Then
        botonEliminar.Enabled = False
    Else
        botonEliminar.Enabled = True
        adoTabla.MoveLast
    End If
End Sub

Public Sub Agregar()
    If Detalle.Text = "" Then
        MsgBox "Debe ingresar una " & TituloSingular, vbCritical, "Ingresar - " & TituloPlural
        Detalle.SetFocus
    Else
        CERRAR_TABLA adoTabla
        sSql = "SELECT * FROM " & NombreTabla & " WHERE Detalle = '" & Detalle & "'"
        adoTabla.Open sSql, adoConnection, adOpenKeyset, adLockOptimistic
        
        If adoTabla.EOF Then
            adoTabla.AddNew
            adoTabla!Detalle = Detalle.Text
            adoTabla.Update
            
            adoTabla.MoveFirst
            adoTabla.MoveLast
            
            If NombreTabla = "TiposCurso" Then
                CERRAR_TABLA adoNumeracionCursos
                adoNumeracionCursos.Open "NumeracionCursos", adoConnection, adOpenDynamic, adLockOptimistic
                adoNumeracionCursos.AddNew
                adoNumeracionCursos!idTipoCurso = adoTabla!id
                adoNumeracionCursos!UltimoNumero = 0
                adoNumeracionCursos.Update
                adoNumeracionCursos.Close
            End If
            
            Detalle.Text = ""
            Detalle.SetFocus
            botonEliminar.Enabled = True
        Else
            MsgBox "El dato que intenta ingresar ya existe.", vbCritical, "ERROR - Dato existente"
        End If
        
        adoTabla.Close
        
        Inicializar
    End If
End Sub

Public Sub Eliminar()
    'Valido antes de eliminar para no generar
    'inconsistencia de datos.
    If PuedeBorrar(madoTabla!id) Then
        If MsgBox("¿Confirma la eliminación?", vbYesNo, "Eliminar - " & TituloSingular) = vbYes Then
            adoTabla.Delete
            'adoTabla.MoveFirst
        End If
        
        If adoTabla.RecordCount = 0 Then
            botonEliminar.Enabled = False
        Else
            botonEliminar.Enabled = True
        End If
    Else
        MsgBox "No se puede eliminar este Item." & vbCrLf & "Hay un dato asociado a este/a " & TituloSingular, vbCritical, TituloPlural & " - Eliminar"
    End If
End Sub

Public Sub ValidarDetalle()
    adoTabla.Filter = "Detalle = '" & Detalle.Text & "'"
    If adoTabla.RecordCount > 0 Then
        MsgBox "Ese dato ya existe", vbCritical, "ERROR - " & TituloSingular
        Detalle.SelStart = 0
        Detalle.SelLength = Len(Detalle.Text)
        Detalle.SetFocus
    End If
    adoTabla.Filter = "Detalle <> ''"
End Sub

Public Sub Terminar()
    If Not adoTabla.EOF Then
        If adoTabla.EditMode = adEditInProgress Then
            adoTabla.Update
        End If
    End If
    adoTabla.Close
    strTabla = ""
End Sub

Private Function PuedeBorrar(idItem As String) As Boolean
    Select Case mNombreTabla
        Case "Rubros"
            sSql = "SELECT * FROM Articulos WHERE idRubro = '" & idItem & "'"
        Case "Familias"
            sSql = "SELECT * FROM Articulos WHERE idFamilia = '" & idItem & "'"
        Case "Marcas"
            sSql = "SELECT * FROM Articulos WHERE idMarca = '" & idItem & "'"
        Case "Tipos"
            sSql = "SELECT * FROM Articulos WHERE idTipo = '" & idItem & "'"
        Case "Lineas"
            sSql = "SELECT * FROM Articulos WHERE idLinea = '" & idItem & "'"
        Case "Colores"
            sSql = "SELECT * FROM Articulos WHERE idColor = '" & idItem & "'"
        Case "UnidadesMedida"
            sSql = "SELECT * FROM Articulos WHERE idUnidadMedida = '" & idItem & "'"
        Case "Categoria"
            sSql = "SELECT * FROM Clientes WHERE idCategoria = '" & idItem & "'"
        Case "CondIva"
            sSql = "SELECT * FROM Clientes WHERE idIva = '" & idItem & "'"
            sSql = "SELECT * FROM Proveedores WHERE idIva = '" & idItem & "'"
        Case "ListaPrecios"
            sSql = "SELECT * FROM Clientes WHERE idListaPrecios = '" & idItem & "'"
        Case "Zonas"
            sSql = "SELECT * FROM Clientes WHERE idZona = '" & idItem & "'"
        Case "Localidades"
            sSql = "SELECT * FROM Clientes WHERE idLocalidad = '" & idItem & "'"
        Case "Provincias"
            sSql = "SELECT * FROM Clientes WHERE idProvincia = '" & idItem & "'"
            sSql = "SELECT * FROM Proveedores WHERE idProvincia = '" & idItem & "'"
            sSql = "SELECT * FROM Vendedores WHERE idProvincia = '" & idItem & "'"
        Case "Paises"
            sSql = "SELECT * FROM Clientes WHERE idPais = '" & idItem & "'"
            sSql = "SELECT * FROM Proveedores WHERE idPais = '" & idItem & "'"
            sSql = "SELECT * FROM Vendedores WHERE idPais = '" & idItem & "'"
        Case "Categorias"
            sSql = "SELECT * FROM Clientes WHERE idCategoria = '" & idItem & "'"
        Case "CategoriasProveedor"
            sSql = "SELECT * FROM Proveedores WHERE idCategoria = '" & idItem & "'"
    End Select
        
    adoTablaValidacion.Open sSql, adoConnection, adOpenStatic, adLockOptimistic
    If adoTablaValidacion.EOF Then
        PuedeBorrar = True
    Else
        PuedeBorrar = False
    End If
    
    'Establezco la variable a TRUE hasta que la validación funcione.
    PuedeBorrar = True
    '===============================================================
    
    adoTablaValidacion.Close
End Function

Public Sub IMPRIMIR_TABLA()
    Dim nLinea As Byte
        
    frmPrinter.Show
    
    frmPrinter.FontName = "Courier"
    
    frmPrinter.CurrentX = 4
    frmPrinter.CurrentY = 2
    frmPrinter.FontUnderline = True
    frmPrinter.FontBold = True
    frmPrinter.FontSize = 14
    frmPrinter.Print "Listado de " & TituloPlural
    frmPrinter.FontSize = 10
    frmPrinter.FontBold = False
    frmPrinter.FontUnderline = False
    
    frmPrinter.CurrentX = 4
    frmPrinter.CurrentY = 4
    frmPrinter.Print "--------------------------------------------------------------------------------"
    frmPrinter.CurrentX = 4
    frmPrinter.CurrentY = 5
    frmPrinter.Print "DETALLE"
    frmPrinter.CurrentX = 4
    frmPrinter.CurrentY = 6
    frmPrinter.Print "--------------------------------------------------------------------------------"
    
    With adoTabla
        .MoveFirst
        nLinea = 7
        Do While Not .EOF
            frmPrinter.CurrentY = nLinea
            frmPrinter.CurrentX = 4
            frmPrinter.Print !Detalle
            nLinea = nLinea + 1
            .MoveNext
        Loop
        .MoveFirst
    End With
End Sub
