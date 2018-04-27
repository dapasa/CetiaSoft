Attribute VB_Name = "modVariables"
'Indica la opción de menú actual
Public sMenu As String
Public sMenuAdicional As String

'Indica el tipo de filtro para el frmFiltro de listados
Public tipoFiltro As String
Public sqlFiltro As String

'Indica la tabla con la que debe trabajar frmTabla
Public strTabla As String

'Indican el id del dato seleccionado en el buscador
Public id_Alumno As Long
Public id_Curso As Long
Public id_Curso_Nuevo As Long
Public id_Empresa As Long
Public id_Profesor As Long
Public id_Movimiento_Consultar As Long

Public detalle_Curso_Nuevo As String

'Indica el estilo de buscador a utilizar
Public EstiloBuscador As String

'Indica el tipo de buscador de documentos
Public TipoBusDoc As String

'Guarda el mensaje de validación de las tablas maestras
Public MensajeValidacion As String

'Indica el tipo de lista actual en frmListaAlumnos
Public TipoLista As String 'Inscriptos - Espera

'Bandera que permite volver a frmListaAlumnos cuando se elimina un
'inscripto mostrando la lista de espera
Public banVolverAListaAlumnos As Boolean

'Se utilizan para pasar la información entre frmCuotas y frmInscripcion
Public x_cuotas As Byte
Public x_valorCuota As Single
Public x_valorCuotaReal As Single
Public x_valorMatricula As Single
Public x_valorMatriculaReal As Single
Public x_DetalleDescuento As String
Public x_alumno_form_inscrip_factura As String

Public x_ficha_alumno_desde_factura As String

Public x_imprimir As Boolean

Public x_usuario As String

Public cambio_foto As Boolean

'Se utiliza para pasar la información entre frmCuotas y frmfactura
Public x_es_ex_alumno As Boolean

'Se utiliza para saber el comprobante que estoy generando para, posteriormente,
'poder cancelar y volver la numeración atrás.
Public x_comprobante As String

'Se utiliza para saber si el usuario puede o no hacer alguna de las operaciones
'restringidas.
Public AccesoPermitido As Boolean

'Se utiliza para que, si cancelo el frmCuotas, no inscriba al alumno.
Public CancelaInscripcion As Boolean

'Se utiliza para que, si cancelo el buscador de empresas desde factura, no intente documentos pendientes.
Public CancelaBusEmpresa As Boolean

