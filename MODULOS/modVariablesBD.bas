Attribute VB_Name = "modVariablesBD"
Public sSql As String

Public connectionString As String
Public adoConnection As ADODB.Connection

Public adoAlumnos As ADODB.Recordset
Public adoAlumnosXCurso As ADODB.Recordset
Public adoAulas As ADODB.Recordset
Public adoClasesXCurso As ADODB.Recordset
Public adoClasesXProfe As ADODB.Recordset
Public adoComoLlego As ADODB.Recordset
Public adoCompaniasCelular As ADODB.Recordset
Public adoCondIva As ADODB.Recordset
Public adoCursos As ADODB.Recordset
Public adoCursosDisponibles As ADODB.Recordset
Public adoCursosXProfesor As ADODB.Recordset
Public adoDuraciones As ADODB.Recordset
Public adoEmisores As ADODB.Recordset
Public adoEmpresas As ADODB.Recordset
Public adoEstadoCursosXAlumno As ADODB.Recordset
Public adoFormasPago As ADODB.Recordset
Public adoHorarios As ADODB.Recordset
Public adoItemsXMov As ADODB.Recordset
Public adoListaEspera As ADODB.Recordset
Public adoLocalidades As ADODB.Recordset
Public adoLog As ADODB.Recordset
Public adoMovimientos As ADODB.Recordset
Public adoNumeracion As ADODB.Recordset
Public adoNumeracionCursos As ADODB.Recordset
Public adoPrecios As ADODB.Recordset
Public adoProfesores As ADODB.Recordset
Public adoProximosInicios As ADODB.Recordset
Public adoRecargosTarjeta As ADODB.Recordset
Public adoSucursales As ADODB.Recordset
Public adoTiposComprobante As ADODB.Recordset
Public adoTiposCurso As ADODB.Recordset
Public adoTiposDoc As ADODB.Recordset
Public adoUnidadesNegocio As ADODB.Recordset
Public adoUsuarios As ADODB.Recordset


Public adoTabla As ADODB.Recordset
Public adoTablaBus As ADODB.Recordset
Public adoTablaValidacion As ADODB.Recordset

Public adoTemp As ADODB.Recordset
Public adoTemp2 As ADODB.Recordset
Public adoTemp3 As ADODB.Recordset

Public adoTempAlumnos As ADODB.Recordset
Public adoTempClases As ADODB.Recordset
Public adoTempCursos As ADODB.Recordset
Public adoTempEspera As ADODB.Recordset
Public adoTempEstadoCursos As ADODB.Recordset
Public adoTempFactura As ADODB.Recordset
Public adoTempInscriptos As ADODB.Recordset
Public adoTempProfesores As ADODB.Recordset

