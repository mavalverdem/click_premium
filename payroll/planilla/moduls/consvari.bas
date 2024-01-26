Attribute VB_Name = "mdlConsVar"
Option Explicit

Public ps_DataBase As String                ' Nombre de la base de datos del sistema
Public ps_DaBasCon As String                ' Nombre de la base de datos de contabilidad
Public s_Sql As String                      ' Cadena de sentencias de sql

Public ps_NomSistema As String              ' Nombre del sistema
Public ps_NomEmpresa As String              ' Nombre de empresa
Public ps_CodEmpresa As String * 3          ' Codigo de empresa
Public ps_EmpresaCon As String              ' Codigo de empresa contabilidad
Public ps_RucEmpresa As String * 11         ' Numero de RUC de empresa
Public ps_PathSystem As String              ' Directorio de ubicación del sistema
Public ps_Fecha_LimiteProc                  ' Fecha Limite para permitir procesos en el Sistema

Public pn_NivelCenCosto  As Integer         ' Nivel de movimiento de centro de costo

Public gsReportName As String               ' Nombre de archivo de reporte
Public gsReportTitle As String              ' Título de la ventana de reporte
Public gs_FechaHora As String               ' Fecha y hora del Sistema

Public ps_ClsPlanilla As String             ' Codigo clase planilla activa
Public ps_DesClsPlanilla As String          ' Descripcion de clase planilla activa
Public pn_HoroLaboraxDia As Double          ' Numero de horas laborables por dia
Public ps_Anyo As String * 4                ' Año activo
Public ps_Mes As String * 2                 ' Mes activo
Public ps_Usuario As String                 ' Nombre de usuario del sistema
Public ps_NivelUsr As String * 1            ' Nivel de usuario
Public pl_Salir As Boolean                  ' Swist de acceso o salida del sistema

Public a_Campos(), a_Where()                ' Arreglos para formar los datos de actualizacion
Public a_Valores(), a_Tipos()               ' Arreglos para formar los datos de actualizacion

' Deficions de tipos estandares
Public Enum TipoExecution                   ' Tipo de acciona realizar
  Inserta = 0: Modifica = 1: Elimina = 2: Seleccion = 3
End Enum

Public Enum TipoDato                        ' Tipos de datos caracter, numerico, fecha, boooleano
  Caracter = 0: Numero = 1: Fecha = 2: Logico = 3
End Enum
Public Enum DatoAdo
  dChar = adChar: dCaracter = adVarChar: dDecimal = adNumeric
  dFLoat = adDouble: dEntero = adInteger: dsEntero = adSmallInt
  dFecha = adDBTimeStamp: dLogico = adBinary
End Enum
Public Enum TypeParameter
  pEntrada = adParamInput: pSalida = adParamOutput
  pEntSal = adParamInputOutput: pRetorno = adParamReturnValue
End Enum
Public Enum NivelUsuario                   ' Nivel de usuario
  Administrador = 0: Supervisor = 1: Asistente = 2: Auxiliar = 3
End Enum

Public o_SwSelUbica As Form                 ' Objeto de Caso de selección de ubicación geografica
Public n_SwSelUbica As Integer              ' Indice de seleccion de ubicacion geofrafica

' Controles o modulos de clases
'Public gdl_Funcion As Object                ' Control o clase de libreria de funciones
'Public gdl_Procedure As Object              ' Control o clase de libreria de procedimientos
'Public gdl_Calculo As Object                ' Control o clase de libreria de proceso de Cálculo
Public gdl_Funcion As New syslink.Funciones     ' Control o clase de libreria de funciones
Public gdl_Procedure As New syslink.Procedure   ' Control o clase de libreria de procedimientos
'Public gdl_Calculo As New syslink.Calculo       ' Control o clase de libreria de proceso de Cálculo

' Objetos de instancia de formularios
Public s_SwRegistro As String                     ' Caso de la instancia del formulario
Public o_SelAsistencia As New fSelPersoxPeriodo   ' Registro de selección de de personal por periodo (asistencia)
Public o_SelExepcional As New fSelPersoxPeriodo   ' Registro de selección de de personal por periodo (exepcional)
Public o_SelDisCenCosto As New fSelPersoxDistribu ' Registro de selección de de personal por periodo (centro costo)
Public o_CalculoPersona As New fSelPersoxDetalle  ' Proceso de Cálculo de planilla por persona
Public o_DepuraCalculo As New fSelPeriCalculo     ' Proceso de inicialización de Cálculo de planilla
Public o_ContaPlanilla As New fContabilizacion    ' Proceso de contabilización de planillas de Cálculos
Public o_RepContaPlani As New fReporContabiliza   ' Proceso de contabilización de planillas de Cálculos
Public o_SelConsulxcpc As New fSelPersoxPeriodo   ' Selección de personal por periodo (consulta por concepto)
Public o_Consultaxcpc As New fConsultaCalculo     ' Consulta de Cálculo (concepto x personal)
Public o_Consultaxpsn As New fConsultaCalculo     ' Consulta de Cálculo (personal x concepto)
Public o_SelRentaQuinta As New fSelPersoxAnalisis ' Selección de personal por periodo (consulta renta quinta)
Public o_SelVacacionAna As New fSelPersoxEstado   ' Selección de personal por estado y analisis (consulta de vacaciones)
Public o_PvsVacaPeriodo As New fPvsPeriodo        ' Periodo de provisiones de vacaciones
Public o_PvsVacaciones As New fPvsPersonal        ' Provisiones de vacaciones
Public o_PvsVacaConsul As New fConsultaProvision  ' Consulta provisones de vacaciones - Cálculo
Public o_PvsVacaCalcul As New fPvsCalculo         ' Cálculo de provisiones de vacaciones
Public o_PvsVacaDepura As New fPvsCalculo         ' Depuración Cálculo provisiones de vacaciones
Public o_PvsGratiPeriod As New fPvsPeriodo        ' Periodo de provisiones de gratificaciones
Public o_PvsGratifica As New fPvsPersonal         ' Provisiones de gratificaciones
Public o_PvsGratiConsul As New fConsultaProvision ' Consulta provisiones de gratificacion - Cálculo
Public o_PvsGratiCalcul As New fPvsCalculo        ' Cálculo de provisones de gratificacion
Public o_PvsGratiDepura As New fPvsCalculo        ' Depuración Cálculo provisiones de gratificacion
Public o_PvsComxTieSer As New fPvsPersonal        ' Provisiones de compensacion x tiempo de servicio(cts)
Public o_RepComxTieSer As New fSelPersonalCts     ' Analisis de compensacion x tiempo de servicio(cts)
Public o_ContaProvision As New fContabilizacion   ' Proceso de contabilización de provisiones (vacación, gratificación y cts)
Public o_RptReciboPago As New fReciboPago         ' Reporte de recibo de pago
Public o_RepLiquidacion As New fSelPersoxLiquida  ' Reporte de liquidación de beneficios
Public o_CertifikdoLiqi As New fSelPersoxLiquida  ' Reporte de certificado de trabajo
Public o_PlanillaGnral As New fReporPlanillaGnral ' Reporte de planilla general(ministerio)
Public o_SelReporGnral As New fSelReporte         ' Selección de formato de reportes
Public o_RepPrePlanilla  As New fSelPeriodo       ' Reporte planilla de trabajo por centro de costo
Public o_ExportarSunat As New fExpInformacion     ' Generación de información sunat-pdt
Public o_Certifikdo5ta As New fSelPersoCertifik   ' Certificado de quinta categoria
Public o_CertifikdoSnp As New fSelPersoCertifik   ' Certificado de ONP
Public o_CertifikdoAfp As New fSelEntiAfpCertifik ' Certificado de Entidad de pensión(AFP)
Public o_CertifikdoUti As New fSelPersoCertifik   ' Certificado de distribución de utilidades
Public o_RptDisBillete As New fReciboPago         ' Reporte de distribución de billetaje

'MAYO 2015
Public o_SelConsulxpsn As New fSelConceptoxPersona 'Selección de concepto por persona (consulta de concepto x Persona)

' Formatos de visualización de Información
Public Const s_FormatoNum_0 As String = "#,###,###,##0.00"    ' Formato númerico general de importe
Public Const s_FormatoNum_1 As String = "#0.0000"             ' Formato númerico de tipo de cambio
Public Const s_FormatoNum_2 As String = "###,###,##0.000000"  ' Formato númerico general de costos
Public Const s_FormatoFecha As String = "dd/mm/yyyy"          ' Formato de fecha corta
Public Const s_FmtFechaHora As String = "dd/mm/yyyy hh:mm:ss" ' Formato de fecha y hora larga
Public Const s_FormatoHora_0 As String = "hh:mm am/pm"        ' Formato de hora 12 horas
Public Const s_FormatoHora_1 As String = "hh:mm"              ' Formato de hora 24 horas

Public Const s_FmtFechMysql_0 As String = "yyyy/mm/dd"          ' Formato de fecha corta para mysql visualizacion
Public Const s_FmtFechMysql_1 As String = "%Y/%m/%d"            ' Formato de fecha corta para mysql
Public Const s_FmtFeHoMysql_0 As String = "yyyy-mm-dd hh:mm:ss" ' Formato de fecha y hora larga mysql visualizacion
Public Const s_FmtFeHoMysql_1 As String = "%d/%m/%Y %H:%i:%s"   ' Formato de fecha y hora larga mysql

' Constantes de informacion
Public Const s_Estado_Ina As Byte = "0"                       ' Estado no activo
Public Const s_Estado_Act As Byte = "1"                       ' Estado activo
Public Const s_Estado_Blq As Byte = "2"                       ' Estado bloqueado

Public Const s_Centro_Ina As Byte = "0"                       ' Estado no activo
Public Const s_Centro_Act As Byte = "1"                       ' Estado activo

Public Const s_Codmon_mn As String = "N"                      ' Tipo de moneda nacional
Public Const s_Codmon_me As String = "E"                      ' Tipo de moneda extranjera
Public Const s_Codmon_mn_Txt As String = "S/."                ' Signo de moneda nacional
Public Const s_Codmon_me_Txt As String = "US$"                ' Signo de moneda extranjera
Public Const s_Codmon_mn_Nom As String = "Soles"              ' Nombre de moneda nacional
Public Const s_Codmon_me_Nom As String = "Dólares Americanos" ' Nombre de moneda extranjera

Public Const s_PeriodoRemAper As String = "00000000"          ' Periodo Remuneraciones Anteriores
Public Const s_ProcesoRemAper As String = "00"                ' Proceso de Cálculo de Remuneraciones Anteriores
Public Const s_EstadoRemAper As Byte = "9"                    ' Estado Remuneraciones Anteriores

Public Const n_FormatoReg As Byte = 0                         ' Formato de registro de información
Public Const n_FormatoPrc As Byte = 1                         ' Formato de procesos de información
Public Const n_FormatoRpt As Byte = 2                         ' Formato de reporte de información
Public Const n_FormatoLst As Byte = 3                         ' Formato de listado de información
Public Const n_FormatoCst As Byte = 4                         ' Formato de consulta de información
Public Const n_FormatoPvs As Byte = 5                         ' Formato de provisión de información
Public Const n_FormatoLbr As Byte = 6                         ' Formato de otros

Public Const s_MdoData_Ins As String = "A"                    ' Caso Insertar Información
Public Const s_MdoData_Del As String = "B"                    ' Caso Eliminar Información
Public Const s_MdoData_Upd As String = "C"                    ' Caso Actualizar Información
Public Const s_MdoData_Vis As String = "V"                    ' Caso Visualizar Información

Public Const n_BackColorHelp As Double = &H80000018           ' Color de Grilla de Ayuda
Public Const n_BackColorMdf As Double = 13427690              ' Color de Grilla de Modificación

Public Const nNewBlankDocument As Integer = 0                 ' Nuevo documento de word en blanco
Public Const nFormLetters As Integer = 0                      ' Formato de documento carta
Public Const nOpenFormatAuto As Integer = 0                   ' Formato por defecto
Public Const nMergeSubTypeOther As Integer = 0                ' Subtipo de combinación

Public Const nMaxTime As Long = 10                            ' Segundos de espera
Public Const nSleepTime As Long = 250                         ' Milisegundos de espera

Public Const KEYEVENTF_KEYUP = &H2                            ' presiona tecla
Public Const KEYEVENTF_EXTENDEDKEY = &H1                      ' suelta tecla

' Constantes  de mensajes
Public Const s_Msg_ValDato_3000 As String = "Registro no Existe, Verificar"     ' Mensaje de registros
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

