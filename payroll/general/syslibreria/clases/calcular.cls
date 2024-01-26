VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                             ' Declarar variable antes de usarla

Public oConexion As Object                  ' Objeto de conexión
Public sCadenaConexion As String            ' Cadena de conexión
Public sClasePlanilla As String             ' Codigo de clase planilla
Public sTipoCalculo As String               ' Tipo de calculo
Public sTipoProceso As String               ' Tipo de proceso de calculo
Public sCodigoPeriodo As String             ' Codigo de periodo de pago
Public sCodigoEmpleado As String            ' Codigo de personal
Public sDiaProceso As String                ' Dia de proceso
Public sMesProceso As String                ' Mes de proceso
Public sAnyoProceso As String               ' Año de proceso de calculo
Public sFechaIngreso As String              ' Fecha de ingreso del personal
Public sFechaCese As String                 ' Fecha de cese del personal
Public sEstadoEmpleado As String            ' Estado del personal
Public sDesAusenciaBF As String             ' Descuento ausencias beneficios sociales

Private porstRecord As ADODB.Recordset      ' Recordset de Resultado(WithEvents)
Private sSQL As String                      ' Cadena de sentencia de seleccion de registros

'[
Function Acumulado(ByVal sConcepto As String, ByVal sDelMes As String) As Double
  '[ Retona el importe acumulado de un concepto ]
  Dim sAnoMesIngreso As String
  
  On Error GoTo ControlError
    
  ' Obtengo el año, mes de ingreso
  sAnoMesIngreso = Format(sFechaIngreso, "yyyymm")
  
  ' Instancio al recordset
  Set porstRecord = New ADODB.Recordset
  
  ' Obtengo importe de retenciones
  sSQL = "SELECT SUM(IFNULL(importe_mn, 0)) AS nRemunera_mn, SUM(IFNULL(importe_me, 0)) AS nRemunera_me "
  sSQL = sSQL & "FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND codcpc='" & sConcepto & "' "
  sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND pdomes" & IIf(sDelMes = "S", "=", "<=") & "'" & sMesProceso & "' "
  sSQL = sSQL & "AND (CONCAT(pdoano, pdomes)>='" & sAnoMesIngreso & "' "
  sSQL = sSQL & "OR CONCAT(pdoano, codpdo)='" & sAnyoProceso & s_PeriodoRemAper & "')"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  Acumulado = 0
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    Acumulado = CDec(porstRecord!nRemunera_mn)
  End If
  porstRecord.Close
  Set porstRecord = Nothing
    
Exit Function
ControlError:
  Err.Clear

End Function
Function BaseEssalud(ByVal sBaseImponible As String, ByVal sAporteEssalud As String, ByVal nSueldoMinimo As Double, ByVal sRemuSubsidio As String) As Double
  ' [ Calcula laBase Imponible para ESSALUD como Aporte del Empleador ]
  Dim nBaseImponible As Double, nAporteEssalud As Double, nRemuSubsidio As Double
  Dim aRemuneracion(2, 3) As Double, nBImponibleMinima As Double
  Dim nSecuencia As Integer, nDCalendario As Integer, nDSubsidio As Integer
  Dim nDBaseMes As Integer, nDLicencia As Integer, nDvacacion As Integer
  Dim sExpresion As String
  
  On Error GoTo ControlError
  ' Instancio al recordset
  Set porstRecord = New ADODB.Recordset
  Set gdl_Funcion = CreateObject("syslink.Funciones")
  
  ' Primer paso: Acumuladores del periodo
  For nSecuencia = 1 To 3
    sExpresion = Choose(nSecuencia, sBaseImponible, sAporteEssalud, sRemuSubsidio)
    sSQL = "SELECT ROUND(IFNULL(SUM(IFNULL(importe_mn, 0)),0), 2) nRemunera_mn, ROUND(IFNULL(SUM(IFNULL(importe_me, 0)),0), 2) nRemunera_me "
    sSQL = sSQL & "FROM plresultado "
    sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND codpdo='" & sCodigoPeriodo & "' "
    sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND codcpc='" & sExpresion & "' "
    sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
    sSQL = sSQL & "AND pdomes='" & sMesProceso & "'"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      aRemuneracion(1, nSecuencia) = CDec(porstRecord!nRemunera_mn)
    End If
    porstRecord.Close
  Next nSecuencia
  
  ' Segundo paso: Dias vacaciones y licencia sin goce
  sSQL = "SELECT SUM(asi.diatrabajo+asi.diamediotm+asi.diaparcial+asi.diaprepostnatal+asi.accidente+asi.enfermedad) DiasBase, "
  sSQL = sSQL & "SUM(IF(asi.codmdi_licen='05',asi.licencia,0)) DiaLicencia, SUM(asi.diavacaciones) DiaVacacion "
  sSQL = sSQL & "FROM plasistencia asi "
  sSQL = sSQL & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
  sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND pdo.anopdo='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND pdo.mespdo='" & sMesProceso & "'"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    nDBaseMes = CInt(porstRecord!DiasBase)
    nDLicencia = CInt(porstRecord!DiaLicencia)
    nDvacacion = CInt(porstRecord!DiaVacacion)
  End If
  porstRecord.Close

  ' Tercer paso: Acumuladores del mes
  For nSecuencia = 1 To 3
    sExpresion = Choose(nSecuencia, sBaseImponible, sAporteEssalud, sRemuSubsidio)
    sSQL = "SELECT ROUND(IFNULL(SUM(IFNULL(importe_mn, 0)),0), 2) nRemunera_mn, ROUND(IFNULL(SUM(IFNULL(importe_me, 0)),0), 2) nRemunera_me "
    sSQL = sSQL & "FROM plresultado "
    sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND codpdo<>'" & sCodigoPeriodo & "' "
    sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND codcpc='" & sExpresion & "' "
    sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
    sSQL = sSQL & "AND pdomes='" & sMesProceso & "'"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      aRemuneracion(2, nSecuencia) = CDec(porstRecord!nRemunera_mn)
    End If
    porstRecord.Close
  Next nSecuencia
  nBaseImponible = Round(aRemuneracion(1, 1) + aRemuneracion(2, 1), 2)
  nRemuSubsidio = Round(aRemuneracion(1, 3) + aRemuneracion(2, 3), 2)
  
  ' Determino dias calendario y subsidio
  nDCalendario = gdl_Funcion.NumeroDiasMes(CInt(sMesProceso), CInt(sAnyoProceso))
  nBImponibleMinima = Round(((1 - (nDSubsidio / nDCalendario)) * nSueldoMinimo), 2)
  
  ' Determino base imponible del mes no sea inferior al Salario Minimo
  nBaseImponible = IIf((nSueldoMinimo > nBaseImponible And nRemuSubsidio = 0), nSueldoMinimo, nBaseImponible)
  nBaseImponible = IIf(aRemuneracion(1, 1) = 0, 0, nBaseImponible)

'' Validaciones aporte acumulado
'AporteAcumu := IIF(BaseImponible <= 0, 0, AporteAcumu);
'AporteEssalud := RND(BaseImponible * FactorEssalud / 100);
'AporteEssalud := AporteEssalud - AporteAcumu;
  BaseEssalud = nBaseImponible
  
  Set porstRecord = Nothing
  Set gdl_Funcion = Nothing
  
  Exit Function
ControlError:
  Err.Clear

End Function
Function CtsTrunca() As Integer
  '[ Retorna el numero de dias compensación por tiempo de servicio(CTS) trunca]
  Dim nAnoIngreso As Integer, nMesIngreso As Integer, nDiasIngreso As Integer
  Dim nAnyoProceso  As Integer, nMesProceso As Integer
  Dim sFechaProceso As String, sFechaInicio As String, sFechaFinal As String
  Dim nDiasCalculo As Long, nDiasAusencia As Long
  
  ' Configuro los objetos de clases
  Set gdl_Funcion = CreateObject("syslink.Funciones")

  ' Valido estado y tiempo minimo de trabajo
  If Not (sEstadoEmpleado = "I" And sTipoProceso = "L") Then nDiasCalculo = 0: GoTo Finaliza
  ' Obtengo requerimiento minimo dias servicio (30)
  sFechaProceso = sDiaProceso & "/" & sMesProceso & "/" & sAnyoProceso
  sFechaProceso = IIf(sFechaCese = "", sFechaProceso, sFechaCese)
  nDiasCalculo = gdl_Funcion.NumeroDias360(CDate(sFechaProceso), CDate(sFechaIngreso), CDate(sFechaCese))
  If nDiasCalculo < 30 Then nDiasCalculo = 0: GoTo Finaliza
  
  ' Inicializo informarción de proceso
  nAnoIngreso = CInt(Right(sFechaIngreso, 4))
  nMesIngreso = CInt(Mid(sFechaIngreso, 4, 2))
  nDiasIngreso = CInt(Left(sFechaIngreso, 2))
  nAnyoProceso = CInt(sAnyoProceso)
  nMesProceso = CInt(sMesProceso)
  
  ' Segundo semestre de calculo
  sFechaInicio = "01/05/" & sAnyoProceso
  sFechaFinal = "31/10/" & sAnyoProceso
  sFechaInicio = IIf((Format(sFechaIngreso, "yyyymmdd") > Format(sFechaInicio, "yyyymmdd")), sFechaIngreso, sFechaInicio)
  sFechaFinal = IIf((Format(sFechaProceso, "yyyymmdd") < Format(sFechaFinal, "yyyymmdd")), sFechaProceso, sFechaFinal)
  sFechaInicio = IIf((nAnoIngreso = nAnyoProceso And nMesIngreso = 4 And nDiasIngreso >= 2), sFechaIngreso, sFechaInicio)
  
  ' Primer semestre de proceso
  If (nMesProceso <= 4 Or nMesProceso >= 11) Then
    sFechaInicio = "01/11/" & (nAnyoProceso - IIf(nMesProceso <= 10, 1, 0))
    sFechaFinal = "30/04/" & (nAnyoProceso + IIf(nMesProceso <= 10, 0, 1))
    If ((nAnyoProceso - nAnoIngreso) <= 1) Then
      sFechaInicio = IIf((Format(sFechaIngreso, "yyyymmdd") > Format(sFechaInicio, "yyyymmdd")), sFechaIngreso, sFechaInicio)
      sFechaInicio = IIf((nMesIngreso = 10 And nDiasIngreso >= 2), sFechaIngreso, sFechaInicio)
    End If
  End If
  sFechaFinal = IIf((Format(sFechaProceso, "yyyymmdd") < Format(sFechaFinal, "yyyymmdd")), sFechaProceso, sFechaFinal)
  
  nDiasCalculo = gdl_Funcion.NumeroDias360(CDate(sFechaProceso), CDate(sFechaInicio), CDate(sFechaFinal))
  nDiasAusencia = 0
  ' Resta Ausencias encontradas en el período (Si el parámetro así lo indica)
  If sDesAusenciaBF = s_Estado_Act Then
    nDiasAusencia = gdl_Funcion.DiasAusenciaBS(sCadenaConexion, sClasePlanilla, sCodigoEmpleado, sFechaInicio, sFechaProceso)
  End If
  nDiasCalculo = nDiasCalculo - nDiasAusencia
  nDiasCalculo = IIf(nDiasCalculo < 0, 0, nDiasCalculo)

Finaliza:
  CtsTrunca = CInt(nDiasCalculo)
  ' Destruyo los objetos
  Set gdl_Funcion = Nothing

End Function
Function DiasGratificacion(ByVal sMesPago As String, ByVal sMesExacto As String, Optional ByVal sProyecta As String = "N") As Long
  '[ Retona el numero de dias graticaciones por fecha ingreso o asistencia]
  Dim nAnyoCalculo As Integer, nMesCalculo As Integer, nDiasCalculo As Integer
  Dim nMesProceso As Integer, nDiasProceso As Integer, nDGratificacion As Integer
  Dim sFechaProceso As String, sFechaDesde As String, sFechaHasta As String
  Dim sFechaEgreso As String
  Dim nServicio As Long, nDiasAusencia As Long
    
  ' Configuro los objetos de clases
  Set gdl_Funcion = CreateObject("syslink.Funciones")
  
  ' Identifico las fechas de los limites de rango
  nMesProceso = CInt(sMesProceso)
  sFechaProceso = sDiaProceso & "/" & sMesProceso & "/" & sAnyoProceso
  sFechaDesde = "01" & "/" & sMesProceso & "/" & sAnyoProceso
  sFechaHasta = sFechaProceso
  DiasGratificacion = 0
  
  ' No considera proyección diciembre y liquidación
  If sProyecta = "S" And (nMesProceso = "12" Or (sEstadoEmpleado = "I" And sTipoProceso = "L")) Then GoTo Finaliza
  
   sFechaEgreso = sFechaProceso
  ' Incializa fecha cese proceso de liquidación
  If (IsDate(sFechaCese) And Mid(sFechaCese, 7, 4) = sAnyoProceso And Mid(sFechaCese, 4, 2) = sMesProceso And sEstadoEmpleado = "I" And sTipoProceso = "L") Then
   sFechaEgreso = sFechaCese
  End If
  
  ' Record minimo un mes de trabajoo 30 dias
  If sProyecta <> "S" Then
    nServicio = DateDiff("d", DateAdd("d", -1, DateAdd("m", 1, CDate(sFechaIngreso))), CDate(sFechaEgreso))
    If ((nServicio < 0 Or DateDiff("d", DateAdd("d", -1, CDate(sFechaIngreso)), CDate(sFechaEgreso)) < IIf(CInt(Mid(sFechaDesde, 4, 2)) <> 2, 30, 28)) And ((sMesPago = "S" And (nMesProceso = 7 Or nMesProceso = 12)) Or (sEstadoEmpleado = "I" And sTipoProceso = "L"))) Then GoTo Finaliza
  End If
  
  ' Determina fechas gratificación semestre pago o trunco
  sFechaDesde = "01/01/" & sAnyoProceso
  sFechaHasta = "30/06/" & sAnyoProceso
  If nMesProceso >= 7 Then
    sFechaDesde = "01/07/" & sAnyoProceso
    sFechaHasta = "31/12/" & sAnyoProceso
  End If
  ' Determino fechas gratificación proyectada
  If sProyecta = "S" Then
    sFechaHasta = "31/12/" & sAnyoProceso
    sFechaProceso = sFechaHasta
    sFechaEgreso = sFechaHasta
  End If
  ' Proceso pago gratificación julio
  If sMesPago = "S" And nMesProceso = 7 Then
    sFechaDesde = "01/01/" & sAnyoProceso
    sFechaHasta = "30/" & Format(nMesProceso - 1, "00") & "/" & sAnyoProceso
  End If
  ' Corrige períodos por Ingreso o Egreso
  sFechaDesde = IIf(Format(sFechaIngreso, "yyyymmdd") > Format(sFechaDesde, "yyyymmdd"), sFechaIngreso, sFechaDesde)
  sFechaHasta = IIf(Format(sFechaHasta, "yyyymmdd") > Format(sFechaEgreso, "yyyymmdd"), sFechaEgreso, sFechaHasta)
  ' Proceso mes exacto
  If sMesExacto = "S" Then
    ' Fecha incial
    nAnyoCalculo = CInt(Right(sFechaDesde, 4))
    nMesCalculo = CInt(Mid(sFechaDesde, 4, 2))
    nDiasProceso = CInt(Mid(sFechaDesde, 1, 2))
    nDiasCalculo = gdl_Funcion.NumeroDiasMes(nMesCalculo, nAnyoCalculo)
    If ((nDiasProceso > 1 And nDiasCalculo < 31) Or (nDiasProceso > 2 And nDiasCalculo = 31)) Then
      sFechaDesde = Format(DateAdd("m", 1, CDate("01/" & Mid(sFechaDesde, 4, 7))), "dd/mm/yyyy")
    End If
  ' Fecha final
    nAnyoCalculo = CInt(Right(sFechaHasta, 4))
    nMesCalculo = CInt(Mid(sFechaHasta, 4, 2))
    nDiasProceso = CInt(Mid(sFechaHasta, 1, 2))
    nDiasCalculo = gdl_Funcion.NumeroDiasMes(nMesCalculo, nAnyoCalculo)
    If (nDiasCalculo > nDiasProceso And nDiasProceso < 30) Then
      sFechaHasta = Format(DateAdd("d", -1, CDate("01/" & Mid(sFechaHasta, 4, 7))), "dd/mm/yyyy")
    End If
  End If
  
  nDGratificacion = 0: nDiasAusencia = 0
  ' Recupero los dias de gratificación
  nDGratificacion = gdl_Funcion.NumeroDias360(CDate(sFechaHasta), CDate(sFechaDesde), CDate(sFechaHasta))
  
  ' Resta Ausencias encontradas en el período (Si el parámetro así lo indica)
  If sDesAusenciaBF = s_Estado_Act Then
    nDiasAusencia = gdl_Funcion.DiasAusenciaBS(sCadenaConexion, sClasePlanilla, sCodigoEmpleado, sFechaDesde, sFechaHasta)
  End If
  nDGratificacion = nDGratificacion - nDiasAusencia
  nDGratificacion = IIf(nDGratificacion < 0, 0, nDGratificacion)
  
  DiasGratificacion = IIf((sMesPago = "N" And nMesProceso = 12), 0, nDGratificacion)

Finaliza:
  ' Destruyo los objetos
  Set gdl_Funcion = Nothing

End Function
Function DiasVacacionGozada(ByVal sPerVacacion As String, ByVal sFechaDesde As String, ByVal sFechaHasta As String) As Long
'[ Retona los dias vacaciones rango de fecha y perido]
  Dim nSecuencia As Integer

  On Error GoTo ControlError
    
  ' Instancio al recordset
  Set porstRecord = New ADODB.Recordset
  
  sSQL = "SELECT ("
  For nSecuencia = 1 To 5
    sSQL = sSQL & "IFNULL((SELECT SUM(IFNULL(TIMESTAMPDIFF(day, asi.fechainivaca" & nSecuencia & ", asi.fechafinvaca" & nSecuencia & "), 0) + (CASE WHEN asi.fechainivaca" & nSecuencia & "<>'' THEN 1 ELSE 0 END)) nDias "
    sSQL = sSQL & "FROM plasistencia asi "
    sSQL = sSQL & "INNER JOIN plperiodo pdo ON pdo.codcls=asi.codcls AND asi.codpdo=pdo.codpdo "
    sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND asi.pdovaca" & nSecuencia & "='" & sPerVacacion & "' "
    sSQL = sSQL & "AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & Format(sFechaDesde, "yyyymm") & "' "
    sSQL = sSQL & "AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & Format(sFechaHasta, "yyyymm") & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechainivaca" & nSecuencia & ",'%Y-%m-%d')>='" & Format(sFechaDesde, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechafinvaca" & nSecuencia & ",'%Y-%m-%d')<='" & Format(sFechaHasta, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "GROUP BY asi.codcls, asi.codpsn, asi.pdovaca" & nSecuencia & "),0)"
    sSQL = sSQL & IIf(nSecuencia = 5, ") nDias", " + ")
  Next nSecuencia
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not (porstRecord.EOF And porstRecord.BOF) Then
    DiasVacacionGozada = CDec(porstRecord!nDias)
  End If
  porstRecord.Close

ControlError:
  ' Destruyo los objetos
  Set porstRecord = Nothing
  Err.Clear
  
End Function
Function EvaluaValores(ByVal sConcepto As String, ByVal nValorAnalizar As Double, ByVal sLimite As String, ByVal sDelMes As String) As Double
  '[ Retona comparaciones de importe de un concepto y referencia ]
  Dim nImpAnterior As Double, nImpPeriodo As Double, nImpAcumulado As Double
  Dim nAnalizaOld As Double, nAnalizaFin As Double
  Dim sAnoMesIngreso As String
  
  On Error GoTo ControlError
    
  ' Obtengo el año, mes de ingreso
  sAnoMesIngreso = Right(sFechaIngreso, 4) & Mid(sFechaIngreso, 4, 2)
  EvaluaValores = 0: nImpAnterior = 0
  nImpPeriodo = 0: nImpAcumulado = 0
  
  ' Instancio al recordset
  Set porstRecord = New ADODB.Recordset
  
  ' Obtengo importe de retenciones
  sSQL = "SELECT SUM(IFNULL(nRemunera_mn, 0)) nRemunera_mn, SUM(IFNULL(nRemunera_me, 0)) nRemunera_me, "
  sSQL = sSQL & "SUM(IFNULL(nAcuRemunera_mn, 0)) nAcuRemunera_mn, SUM(IFNULL(nAcuRemunera_me, 0)) nAcuRemunera_me "
  sSQL = sSQL & "FROM("
  sSQL = sSQL & "SELECT res.codpsn, SUM(IFNULL(res.importe_mn, 0)) nRemunera_mn, SUM(IFNULL(res.importe_me, 0)) nRemunera_me, "
  sSQL = sSQL & "0.00 nAcuRemunera_mn, 0.00 nAcuRemunera_me "
  sSQL = sSQL & "FROM plresultado res "
  sSQL = sSQL & "WHERE res.codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND res.codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND res.codcpc='" & sConcepto & "' "
  sSQL = sSQL & "AND res.pdoano='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND res.pdomes" & IIf(sDelMes = "S", "=", "<=") & "'" & sMesProceso & "' "
  sSQL = sSQL & "AND res.codpdo='" & sCodigoPeriodo & "' "
  sSQL = sSQL & "AND (CONCAT(res.pdoano, res.pdomes)>='" & sAnoMesIngreso & "' "
  sSQL = sSQL & "OR CONCAT(res.pdoano, res.codpdo)='" & sAnyoProceso & Left(s_PeriodoRemAper, 2) & "') "
  sSQL = sSQL & "GROUP BY res.codpsn "
  sSQL = sSQL & "UNION "
  sSQL = sSQL & "SELECT res.codpsn, 0.00 nRemunera_mn, 0.00 nRemunera_me, "
  sSQL = sSQL & "SUM(IFNULL(res.importe_mn, 0)) nAcuRemunera_mn, SUM(IFNULL(res.importe_me, 0)) nAcuRemunera_me "
  sSQL = sSQL & "FROM plresultado res "
  sSQL = sSQL & "WHERE res.codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND res.codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND res.codcpc='" & sConcepto & "' "
  sSQL = sSQL & "AND res.pdoano='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND res.pdomes" & IIf(sDelMes = "S", "=", "<=") & "'" & sMesProceso & "' "
  sSQL = sSQL & "AND res.codpdo<>'" & sCodigoPeriodo & "' "
  sSQL = sSQL & "AND (CONCAT(res.pdoano, res.pdomes)>='" & sAnoMesIngreso & "' "
  sSQL = sSQL & "OR CONCAT(res.pdoano, res.codpdo)='" & sAnyoProceso & s_PeriodoRemAper & "')"
  sSQL = sSQL & "GROUP BY res.codpsn "
  sSQL = sSQL & ") tblanaliza "
  sSQL = sSQL & "GROUP BY codpsn"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    nImpAnterior = porstRecord!nAcuRemunera_mn
    nImpPeriodo = CDec(porstRecord!nRemunera_mn)
    nImpAcumulado = Round(nImpAnterior + nImpPeriodo, 2)
  End If
  
  ' Valido que importe segun limite(minimo o maximo), calculo anterior
  If sLimite = "MIN" Then
    EvaluaValores = IIf(nImpAcumulado >= nValorAnalizar, nImpAcumulado, nValorAnalizar)
    If nImpAnterior > 0 Then
      nAnalizaOld = IIf(nImpAnterior >= nValorAnalizar, nImpAnterior, nValorAnalizar)
      nAnalizaFin = IIf(nImpAcumulado >= nAnalizaOld, nImpAcumulado, nAnalizaOld)
      EvaluaValores = Round(nAnalizaFin - nAnalizaOld, 2)
    End If
  Else
    EvaluaValores = IIf(nImpAcumulado >= nValorAnalizar, nValorAnalizar, nImpAcumulado)
    If nImpAnterior > 0 Then
      nAnalizaOld = IIf(nImpAnterior >= nValorAnalizar, nValorAnalizar, nImpAnterior)
      nImpAcumulado = IIf(nImpAcumulado >= nValorAnalizar, nValorAnalizar, nImpAcumulado)
      nAnalizaFin = IIf((nImpAcumulado >= nAnalizaOld), nImpAcumulado, nAnalizaOld)
      EvaluaValores = Round(nAnalizaFin - nAnalizaOld, 2)
    End If
  End If
  EvaluaValores = IIf(EvaluaValores > 0, EvaluaValores, 0)

Exit Function
ControlError:
  Err.Clear
  ' Destruyo los objetos
  porstRecord.Close
  Set porstRecord = Nothing

End Function
Function Pendiente(ByVal sConcepto As String, ByVal nPendiente As Integer) As Double
  '[ Retona el importe pendiente del concepto ]
    
  Dim nConcepto As Double
  On Error GoTo ControlError
    
  ' Instancio al recordset
  Set porstRecord = New ADODB.Recordset
  
  ' Obtengo importe del concepto
  sSQL = "SELECT pdomes, IFNULL(res.importe_mn, 0) AS nRemunera_mn, IFNULL(res.importe_me, 0) AS nRemunera_me "
  sSQL = sSQL & "FROM plresultado res "
  sSQL = sSQL & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND IFNULL(DATE_FORMAT(psn.fecbaja, '%Y%m'), '')<>'" & sAnyoProceso & sMesProceso & "' "
  sSQL = sSQL & "WHERE res.codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND res.codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND res.codcpc='" & sConcepto & "' "
  sSQL = sSQL & "AND res.pdoano='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND res.pdomes='" & sMesProceso & "'"
  sSQL = sSQL & "ORDER BY pdomes DESC"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  nConcepto = 0
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    nConcepto = CDec(porstRecord!nRemunera_mn)
  End If
  porstRecord.Close
  Set porstRecord = Nothing
  Pendiente = CDec(nConcepto * nPendiente)
    
Exit Function
ControlError:
  Err.Clear

End Function
Function Promedio(ByVal sConcepto As String, ByVal nMeses As Integer, ByVal nIncidencia As Integer, ByVal sMesActual As String, ByVal sCalculo As String) As Double
  '[ Retona el importe de promedio de un concepto ]
    
  Dim nConcepto As Double, nMesesPromedio As Double
  Dim sPeriodoIni As String, sPeriodoFin As String, sFechaPro As String
  Dim dFechaFin As Date, dFechaIni As Date
  Dim nRegistro As Integer, nMesMenos As Integer
  
  On Error GoTo ControlError
  
  ' Configuro los objetos de clases
  Set gdl_Funcion = CreateObject("syslink.Funciones")
  ' Instancio al objeto
  Set porstRecord = New ADODB.Recordset
    
  ' Obtengo fecha ingreso si hay ocurrencias
  If nIncidencia > 0 Then
    sFechaPro = DateAdd("m", IIf(sMesActual = "S", 0, -1), CDate(sDiaProceso & "/" & sMesProceso & "/" & sAnyoProceso))
    nMesMenos = DateDiff("m", CDate("01/" & Format(sFechaIngreso, "mm/yyyy")), CDate("01" & Mid(sFechaPro, 3))) + 1
    nMesMenos = IIf(nMesMenos >= 0, nMesMenos, 0)
    nMeses = IIf((nMesMenos >= nMeses), nMeses, nMesMenos)
  End If
    
  ' Rango de periodos
  dFechaFin = DateAdd("m", IIf(sMesActual = "S", 0, -1), CDate(sDiaProceso & "/" & sMesProceso & "/" & sAnyoProceso))
  sPeriodoFin = Format(dFechaFin, "yyyymm")
  nMesMenos = nMeses - IIf(nMeses <> 0, 1, 0)    ' Primer parametro era 1
  
  dFechaIni = DateAdd("m", -nMesMenos, dFechaFin)
  sPeriodoIni = Format(dFechaIni, "yyyymm")
  Promedio = 0
  
  ' Obtengo el numero de incidencias
  sSQL = "SELECT IFNULL(COUNT(DISTINCT codcls, codpsn, codcpc, pdoano, pdomes), 0) AS nRegistros "
  sSQL = sSQL & "FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND codcpc='" & sConcepto & "' "
  sSQL = sSQL & "AND CONCAT(pdoano, pdomes)>='" & sPeriodoIni & "' "
  sSQL = sSQL & "AND CONCAT(pdoano, pdomes)<='" & sPeriodoFin & "' "
  sSQL = sSQL & "AND (IFNULL(importe_mn, 0)>0.00 OR IFNULL(importe_me, 0)>0.00)"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  nRegistro = 0
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    nRegistro = CInt(porstRecord!nRegistros)
  End If
  porstRecord.Close
  
  If nRegistro >= nIncidencia Then
    ' Obtengo importe del concepto
    sSQL = "SELECT IFNULL(SUM(importe_mn), 0) AS nRemunera_mn, IFNULL(SUM(importe_me), 0) AS nRemunera_me "
    sSQL = sSQL & "FROM plresultado "
    sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND codcpc='" & sConcepto & "' "
    sSQL = sSQL & "AND CONCAT(pdoano, pdomes)>='" & sPeriodoIni & "' "
    sSQL = sSQL & "AND CONCAT(pdoano, pdomes)<='" & sPeriodoFin & "' "
    sSQL = sSQL & "GROUP BY codcpc"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    nConcepto = 0
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      nConcepto = CDec(porstRecord!nRemunera_mn)
    End If
    porstRecord.Close
    
    ' Obtengo los dias de calculo de acuerdo al caso D:ias o M:eses
    nMesesPromedio = nMeses
    If sCalculo = "D" Then
      dFechaIni = CDate("01" & Mid(dFechaIni, 3))
      dFechaIni = IIf((Format(dFechaIni, "yyyymmdd") >= Format(sFechaIngreso, "yyyymmdd")), dFechaIni, sFechaIngreso)
      nMeses = gdl_Funcion.NumeroDias360(CDate(dFechaFin), CDate(dFechaIni), CDate(dFechaFin))
      nMesesPromedio = CDec(nMeses / 30)
    End If
    Promedio = Round(nConcepto / nMesesPromedio, 2)
  End If
  Set porstRecord = Nothing
  Set gdl_Funcion = Nothing
    
Exit Function
ControlError:
  Err.Clear

End Function
Function RentaQuinta(ByVal nValorUit As Double, ByVal nRetenido As Integer, ByVal nDivide As Integer, ByVal s_ConcepGanado As String, ByVal s_ConcepPendiente As String, ByVal s_ConcepExtraordinario As String, ByVal s_ConcepQuinta As String, ByVal s_Domiciliado As String) As Double
  '[ Retona el importe de renta de quinta ]
  Dim aEscala()
  Dim nRemuneraGanado As Double, nRemuneraExtraordinario As Double, nRemuExtraRetenida As Double
  Dim nRemuneraPendiente As Double, nRemuneraAcumulada As Double, nLimiteBase As Double
  Dim nIndex As Integer, nInicio As Integer
  Dim nRentaRetenida As Double, nRentaNeta As Double
  Dim nDeduccion As Double, nRentaNetaExtra As Double
  Dim nRemuneraAfecta As Double, nIngresoAfecto As Double
  Dim nEscalaAfecta As Double, nNoDomiciliado As Double

  On Error GoTo ControlError
  ' Instancio al recordset
  Set porstRecord = New ADODB.Recordset

  ' Obtengo el escala de factores x ejercicio
  sSQL = "SELECT orden, numerouit, factor FROM plescalaquinta WHERE pdoanyo='" & sAnyoProceso & "' ORDER BY orden"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  nIndex = 1: nLimiteBase = 0
  While Not porstRecord.EOF
    ' Incremento la dimensión del array
    ReDim Preserve aEscala(3, nIndex)
    ' Defino las escala y porcentajes
    aEscala(1, nIndex) = nLimiteBase
    aEscala(2, nIndex) = CDec(nValorUit * CInt(porstRecord!numerouit))
    aEscala(3, nIndex) = CDec(porstRecord!factor)
    nLimiteBase = CDec(aEscala(2, nIndex))
    porstRecord.MoveNext
    nIndex = UBound(aEscala, 2) + 1
  Wend
  porstRecord.Close
  
  nRemuneraGanado = 0
  ' Total general de ingresos brutos
  sSQL = "SELECT ROUND(IFNULL(SUM(importe_mn), 0), 2) AS nImporte_mn, ROUND(IFNULL(SUM(importe_me), 0), 2) AS nImporte_me "
  sSQL = sSQL & "FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND codcpc='" & s_ConcepGanado & "' "
  sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND codpdo='" & sCodigoPeriodo & "'"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  nRemuneraGanado = CDec(porstRecord!nImporte_mn)
  porstRecord.Close
  
  ' LBS Meses a retener mes de proceso
  nRetenido = IIf((sEstadoEmpleado = "I" And sTipoProceso = "L"), CInt(sMesProceso), nRetenido)
  nRemuExtraRetenida = 0
  ' Remuneracion extraordinaria (meses a retener)
  sSQL = "SELECT ROUND(IFNULL(SUM(importe_mn), 0), 2) AS nImporte_mn, ROUND(IFNULL(SUM(importe_me), 0), 2) AS nImporte_me "
  sSQL = sSQL & "FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND codcpc='" & s_ConcepExtraordinario & "' "
  sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND pdomes>'" & Format(nRetenido, "00") & "' "
  sSQL = sSQL & "AND pdomes<'" & sMesProceso & "'"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  nRemuExtraRetenida = CDec(porstRecord!nImporte_mn)
  porstRecord.Close
  ' Remuneracion ganada resta acumulado extraordinaria (meses a retener)
  nRemuneraGanado = Round(nRemuneraGanado - nRemuExtraRetenida, 2)

  nRemuneraExtraordinario = 0
  ' Total ingresos extraordinarios brutos periodo
  sSQL = "SELECT ROUND(IFNULL(SUM(importe_mn), 0), 2) AS nImporte_mn, ROUND(IFNULL(SUM(importe_me), 0), 2) AS nImporte_me "
  sSQL = sSQL & "FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND codcpc='" & s_ConcepExtraordinario & "' "
  sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
  sSQL = sSQL & "AND codpdo='" & sCodigoPeriodo & "'"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  nRemuneraExtraordinario = CDec(porstRecord!nImporte_mn)
  porstRecord.Close
  ' Resto ingresos extraordinarios
  nRemuneraGanado = Round(nRemuneraGanado - nRemuneraExtraordinario, 2)
  
  ' Establesco inicio de parametro si es domiciliado
  nInicio = IIf(s_Domiciliado = "D", 1, UBound(aEscala, 2))
  ' Obtengo el total de ingresos brutos anteriores
  If s_Domiciliado = "N" Then
    sSQL = "SELECT ROUND(IFNULL(SUM(importe_mn), 0), 2) AS nImporte_mn, ROUND(IFNULL(SUM(importe_me), 0), 2) AS nImporte_me "
    sSQL = sSQL & "FROM plresultado "
    sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND codcpc='" & s_ConcepGanado & "' "
    sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
    sSQL = sSQL & "AND codpdo<'" & sCodigoPeriodo & "' "
    sSQL = sSQL & "GROUP BY codpdo "
    sSQL = sSQL & "ORDER BY codpdo DESC"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (porstRecord.BOF And porstRecord.EOF) Then
      nNoDomiciliado = CDec(porstRecord!nImporte_mn)
    End If
    porstRecord.Close
  End If
  
  ' Calculo el importe de la renta para no domiciliados
  RentaQuinta = Round(nRemuneraGanado - nNoDomiciliado, 2) * aEscala(3, nInicio)
  If s_Domiciliado = "D" Then
    nRentaRetenida = 0
    ' Obtengo importe de retenciones
    sSQL = "SELECT ROUND(IFNULL(SUM(importe_mn), 0), 2) AS nRenta_mn, ROUND(IFNULL(SUM(importe_me), 0), 2) AS nRenta_me "
    sSQL = sSQL & "FROM plresultado "
    sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND codcpc='" & s_ConcepQuinta & "' "
    sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
    sSQL = sSQL & "AND pdomes<='" & Format(nRetenido, "00") & "'"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      nRentaRetenida = CDec(porstRecord!nRenta_mn)
    End If
    porstRecord.Close
    
    ' Obtengo concepto devolucion quinta anterior
    sSQL = "SELECT ROUND(IFNULL(SUM(res.importe_mn), 0), 2) AS nRenta_mn, ROUND(IFNULL(SUM(res.importe_me), 0), 2) AS nRenta_me "
    sSQL = sSQL & "FROM plresultado res, plcfgempresa cfg "
    sSQL = sSQL & "WHERE res.codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND res.codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND res.pdoano='" & sAnyoProceso & "' "
    sSQL = sSQL & "AND res.pdoano=cfg.pdoano "
    sSQL = sSQL & "AND res.codcpc=cfg.codcpc5ta_ing "
    sSQL = sSQL & "AND res.pdomes<='" & Format(nRetenido, "00") & "'"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      nRentaRetenida = nRentaRetenida - CDec(porstRecord!nRenta_mn)
    End If
    porstRecord.Close
    
    nRemuneraPendiente = 0
    ' Obtengo la remuneración pendiente
    sSQL = "SELECT ROUND(IFNULL(SUM(importe_mn), 0), 2) AS nImporte_mn, ROUND(IFNULL(SUM(importe_me), 0), 2) AS nImporte_me "
    sSQL = sSQL & "FROM plresultado "
    sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND codcpc='" & s_ConcepPendiente & "' "
    sSQL = sSQL & "AND pdoano='" & sAnyoProceso & "' "
    sSQL = sSQL & "AND codpdo='" & sCodigoPeriodo & "'"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    nRemuneraPendiente = CDec(porstRecord!nImporte_mn)
    porstRecord.Close
    
    ' Primer paso: Renta remuneracion ordinaria
    ' Total de remuneración acumulada
    nRemuneraAcumulada = Round(nRemuneraGanado + nRemuneraPendiente, 2)
    nRemuneraAfecta = Round(nRemuneraAcumulada - (nValorUit * 7), 2)
    nIngresoAfecto = nRemuneraAfecta
    
    RentaQuinta = 0: nRentaNeta = 0
    ' Obtengo el impuesto si remunera Afecta es mayor que inicio indice de escale
    If nRemuneraAfecta > aEscala(1, nInicio) Then
      ' Obtengo la renta neta de acuerdo a la escala
      For nIndex = nInicio To UBound(aEscala, 2)
        If nRemuneraAfecta > aEscala(1, nIndex) And nRemuneraAfecta <= aEscala(2, nIndex) Then
          nRentaNeta = nRentaNeta + (nIngresoAfecto * aEscala(3, nIndex))
          Exit For
        Else
          nEscalaAfecta = (aEscala(2, nIndex) - IIf(nIndex > 1, aEscala(2, nIndex - 1), 0))
          nRentaNeta = nRentaNeta + (nEscalaAfecta * aEscala(3, nIndex))
          nIngresoAfecto = nIngresoAfecto - nEscalaAfecta
        End If
      Next nIndex
    End If

    ' Monto de retencion de impuesto
    nDeduccion = CDec(nRentaNeta - nRentaRetenida)
    ' LBS Meses a dividir es 1
    nDivide = IIf((sEstadoEmpleado = "I" And sTipoProceso = "L"), 1, nDivide)
    RentaQuinta = CDec(nDeduccion / nDivide)
    
    ' Segundo paso: Obtengo el impuesto remuneracion extraordinaria
    nRentaNetaExtra = 0
    If nRemuneraExtraordinario > 0 Then
      ' Total de remuneración acumulada
      nRemuneraAcumulada = Round((nRemuneraGanado + nRemuneraPendiente + nRemuneraExtraordinario), 2)
      nRemuneraAfecta = Round(nRemuneraAcumulada - (nValorUit * 7), 2)
      nIngresoAfecto = nRemuneraAfecta
      
      nRentaNetaExtra = 0
      ' Obtengo el impuesto si remunera Afecta es mayor que inicio indice de escale
      If nRemuneraAfecta > aEscala(1, nInicio) Then
        ' Obtengo la renta neta de acuerdo a la escala
        For nIndex = nInicio To UBound(aEscala, 2)
          If nRemuneraAfecta > aEscala(1, nIndex) And nRemuneraAfecta <= aEscala(2, nIndex) Then
            nRentaNetaExtra = nRentaNetaExtra + (nIngresoAfecto * aEscala(3, nIndex))
            Exit For
          Else
            nEscalaAfecta = (aEscala(2, nIndex) - IIf(nIndex > 1, aEscala(2, nIndex - 1), 0))
            nRentaNetaExtra = nRentaNetaExtra + (nEscalaAfecta * aEscala(3, nIndex))
            nIngresoAfecto = nIngresoAfecto - nEscalaAfecta
          End If
        Next nIndex
      End If
      nRentaNetaExtra = Round(nRentaNetaExtra - nRentaNeta, 2)
    End If
    ' Monto de retencion de impuesto
    RentaQuinta = Round(RentaQuinta + nRentaNetaExtra, 2)
  End If
  If RentaQuinta < 0 Then
    RentaQuinta = IIf(((sMesProceso = 12) Or (sEstadoEmpleado = "I" And sTipoProceso = "L")), RentaQuinta, 0)
  End If
  
  ' Periodo de graticación lo divido entre dos
  If RentaQuinta > 0 And sTipoProceso = "G" Then ' Gratificación
    RentaQuinta = RentaQuinta / 2
  End If

Exit Function
ControlError:
    Err.Clear
End Function
Function SubsidioEnfermedad(ByVal sRemuneraFija As String, ByVal sRemuneraVariable1 As String, ByVal sRemuneraVariable2 As String, ByVal sRemuneraVariable3 As String, ByVal sRemuneraVariable4 As String, ByVal sRemuneraVariable5 As String, ByVal nDiasDMedico As Long) As Double
  '[ Retona el importe diario de subsidio enfermedad descanso medico mas de 20 dias  en el año ]
  
  Dim sFechaDesde As String, sFechaHasta As String, sFechaInicio As String, sFechaSubsidio As String
  Dim mDBaseSubsidio As Long, nMesIncidencia As Long, nMesesCalculo As Long
  Dim nDiferenciaMes As Long, nDSubsidio As Long, nDCalendario As Long
  Dim nRemuVariable1 As Double, nRemuVariable2 As Double, nRemuVariable3 As Double
  Dim nRemuVariable4 As Double, nRemuVariable5 As Double
  Dim nRemuneraFija As Double, nRemueraVaria As Double, nRemueraSubsidio As Double
  
  ' Configuro los objetos de clases, instancio recordset
  Set gdl_Funcion = CreateObject("syslink.Funciones")
  Set porstRecord = New ADODB.Recordset
  On Error GoTo Finaliza
  
  ' Determino periodo inicio subsidio
  sFechaInicio = ""
  sFechaDesde = DateAdd("m", -12, CDate("01/" & sMesProceso + "/" & sAnyoProceso))
  sFechaHasta = DateAdd("d", -1, CDate("01/" & sMesProceso + "/" & sAnyoProceso))
  ' Corrige fecha por ingreso
  sFechaDesde = IIf(CDate(sFechaIngreso) > CDate(sFechaDesde), sFechaIngreso, sFechaDesde)
  
  ' Determino inicio de subsidio
  If nDiasDMedico = 0 Then
    sSQL = "SELECT DATE_ADD(asi.fechaini_enfer, INTERVAL -1 day) FInicio "
    sSQL = sSQL & "FROM plasistencia asi "
    sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_enfer,'%Y-%m-%d')>='" & Format(sFechaDesde, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_enfer,'%Y-%m-%d')<='" & Format(sFechaHasta, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "AND NOT EXISTS (SELECT * FROM plasistencia ini "
    sSQL = sSQL & "WHERE ini.codcls=asi.codcls "
    sSQL = sSQL & "AND ini.codpsn=asi.codpsn "
    sSQL = sSQL & "AND ini.fechaini_enfer=DATE_ADD(asi.fechaini_enfer, INTERVAL -1 day)) "
    sSQL = sSQL & "ORDER BY 1 ASC"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      sFechaInicio = porstRecord!FInicio
    End If
    porstRecord.Close
  
    ' Determino fecha inicial de subsidio
    If (IsDate(sFechaInicio) And sFechaInicio <= sFechaHasta) Then
      sFechaDesde = "01/01/" & sAnyoProceso
      
      ' Selecciono registros de subsidio
      sSQL = "SELECT asi.enfermedad DSubsidio, asi.fechaini_enfer FSubsidio "
      sSQL = sSQL & "FROM plasistencia asi "
      sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
      sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
      sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_enfer,'%Y-%m-%d')>='" & Format(sFechaDesde, "yyyy-mm-dd") & "' "
      sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_enfer,'%Y-%m-%d')<='" & Format(sFechaHasta, "yyyy-mm-dd") & "' "
      sSQL = sSQL & "ORDER BY 2 ASC"
      porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
        nDSubsidio = 0
        ' Fecha de inicio subsidio
        While Not porstRecord.EOF
          sFechaSubsidio = porstRecord!FSubsidio
          nDSubsidio = nDSubsidio + CLng(porstRecord!DSubsidio)
          porstRecord.MoveNext
        Wend
      End If
      porstRecord.Close
    
      ' Fecha inicial ejercicio subsidio
      If Not IsDate(sFechaInicio) Then
        sFechaInicio = "01/" & sMesProceso + "/" & sAnyoProceso
      End If
      
      ' Licencias continua inicio subsidio y calendario
      nDSubsidio = 0
      sSQL = "SELECT SUM(asi.enfermedad) DSubsidio "
      sSQL = sSQL & "FROM plasistencia asi "
      sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
      sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
      sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_enfer,'%Y-%m-%d')>='" & Format(sFechaInicio, "yyyy-mm-dd") & "' "
      sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_enfer,'%Y-%m-%d')<='" & Format(sFechaHasta, "yyyy-mm-dd") & "' "
      sSQL = sSQL & "GROUP BY asi.codcls,asi.codpsn"
      porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
        nDSubsidio = CLng(porstRecord!DSubsidio)
      End If
      porstRecord.Close
      nDCalendario = DateDiff("d", CDate(sFechaInicio), CDate(sFechaHasta)) + 1
     
      ' Determinar fecha real incio subsidio
      If CDate(sFechaSubsidio) > CDate(sFechaInicio) Then
        sFechaInicio = sFechaSubsidio
      End If
      sFechaInicio = DateAdd("d", 1, CDate(sFechaInicio))
      
      ' Determino periodo base subsidio
      If nDSubsidio = nDCalendario Then
  '      PProceso = Copy(sFechaInicio, 7, 4) + "/" + Copy(sFechaInicio, 4, 2)
      End If
    End If
  End If
  
  ' Fecha inicial ejercicio subsidio
  If Not IsDate(sFechaInicio) Then
    sFechaInicio = "01/" & sMesProceso + "/" & sAnyoProceso
  End If
  ' Determino periodos base subsidido
  sFechaDesde = DateAdd("m", -12, CDate(sFechaInicio))
  sFechaHasta = DateAdd("d", -1, CDate(sFechaInicio))
  ' Corrige fecha por ingreso
  sFechaDesde = IIf(CDate(sFechaIngreso) > CDate(sFechaDesde), sFechaIngreso, sFechaDesde)

  nRemuneraFija = 0
  ' Obtengo remuneracion base de subsidio, rango periodo
  sSQL = "SELECT ROUND(SUM(IFNULL(importe_mn, 0)), 2) nRemunera_mn, ROUND(SUM(IFNULL(importe_me, 0)), 2) nRemunera_me "
  sSQL = sSQL & "FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND codcpc='" & sRemuneraFija & "' "
  sSQL = sSQL & "AND CONCAT(pdoano, pdomes)>='" & Format(sFechaDesde, "yyyymm") & "' "
  sSQL = sSQL & "AND CONCAT(pdoano, pdomes)<='" & Format(sFechaHasta, "yyyymm") & "'"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    nRemuneraFija = CDec(porstRecord!nRemunera_mn)
  End If
  porstRecord.Close
  
  ' Obtengo promedio de rermuneracion variable
  nMesesCalculo = (DatePart("m", CDate(sFechaHasta)) - DatePart("m", CDate(sFechaDesde))) + 1
  nMesesCalculo = nMesesCalculo + IIf(nMesesCalculo <= 0, 12, 0)

  mDBaseSubsidio = nMesesCalculo * 30
  nMesIncidencia = gdl_Funcion.Redondeo(nMesesCalculo / 2, 0)
  ' Diferencia periodo a Inicio Subsidio
  ' nDiferenciaMes = AMD(StrToDate("01" + Copy(PHasta, 5, 3) + "/" + Copy(PHasta, 1, 4)), FDESPRO, "M") - 1
'   If DiferenciaMes > 0 Then
'      nMesesCalculo = nMesesCalculo + nDiferenciaMes
  
  nRemuVariable1 = Promedio(sRemuneraVariable1, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable2 = Promedio(sRemuneraVariable2, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable3 = Promedio(sRemuneraVariable3, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable4 = Promedio(sRemuneraVariable4, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable5 = Promedio(sRemuneraVariable5, nMesesCalculo, nMesIncidencia, "N", "M")
  
  ' Base de remuneracion diaria
  nRemueraVaria = Round(nRemuVariable1 + nRemuVariable2 + nRemuVariable3 + nRemuVariable4 + nRemuVariable5, 2)
  nRemueraSubsidio = Round(nRemuneraFija + nRemueraVaria, 2)
  nRemueraSubsidio = Round(nRemueraSubsidio / mDBaseSubsidio, 2)

  SubsidioEnfermedad = nRemueraSubsidio
Finaliza:
  Err.Clear
  ' Destruyo los objetos
  Set gdl_Funcion = Nothing
  Set porstRecord = Nothing
    
End Function
Function SubsidioMaternidad(ByVal sRemuneraFija As String, ByVal sRemuneraVariable1 As String, ByVal sRemuneraVariable2 As String, ByVal sRemuneraVariable3 As String, ByVal sRemuneraVariable4 As String, ByVal sRemuneraVariable5 As String) As Double
  
  '[ Retorna remuneracion diaria de subsidio  de maternidad segun remuneracion fija y variable ]

  Dim sFechaDesde As String, sFechaHasta As String, sFechaInicio As String, sFechaSubsidio As String
  Dim mDBaseSubsidio As Long, nMesIncidencia As Long, nMesesCalculo As Long
  Dim nDiferenciaMes As Long, nDSubsidio As Long, nDCalendario As Long
  Dim nRemuVariable1 As Double, nRemuVariable2 As Double, nRemuVariable3 As Double
  Dim nRemuVariable4 As Double, nRemuVariable5 As Double
  Dim nRemuneraFija As Double, nRemueraVaria As Double, nRemueraSubsidio As Double
  
  ' Configuro los objetos de clases, instancio recordset
  Set gdl_Funcion = CreateObject("syslink.Funciones")
  Set porstRecord = New ADODB.Recordset
  On Error GoTo Finaliza
  
  ' Determino periodo inicio subsidio
  sFechaInicio = ""
  sFechaDesde = DateAdd("m", -12, CDate("01/" & sMesProceso + "/" & sAnyoProceso))
  sFechaHasta = DateAdd("d", -1, CDate("01/" & sMesProceso + "/" & sAnyoProceso))
  ' Corrige fecha por ingreso
  sFechaDesde = IIf(CDate(sFechaIngreso) > CDate(sFechaDesde), sFechaIngreso, sFechaDesde)
  
  ' Determino inicio de subsidio
  sSQL = "SELECT DATE_ADD(asi.fechaini_natal, INTERVAL -1 day) FInicio "
  sSQL = sSQL & "FROM plasistencia asi "
  sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_natal,'%Y-%m-%d')>='" & Format(sFechaDesde, "yyyy-mm-dd") & "' "
  sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_natal,'%Y-%m-%d')<='" & Format(sFechaHasta, "yyyy-mm-dd") & "' "
  sSQL = sSQL & "AND NOT EXISTS (SELECT * FROM plasistencia ini "
  sSQL = sSQL & "WHERE ini.codcls=asi.codcls "
  sSQL = sSQL & "AND ini.codpsn=asi.codpsn "
  sSQL = sSQL & "AND ini.fechaini_natal=DATE_ADD(asi.fechaini_natal, INTERVAL -1 day)) "
  sSQL = sSQL & "ORDER BY 1 ASC"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    sFechaInicio = porstRecord!FInicio
  End If
  porstRecord.Close

  ' Determino fecha inicial de subsidio
  If (IsDate(sFechaInicio) And sFechaInicio <= sFechaHasta) Then
    sFechaDesde = "01/01/" & sAnyoProceso
    
    ' Selecciono registros de subsidio
    sSQL = "SELECT asi.diaprepostnatal DSubsidio, asi.fechaini_natal FSubsidio "
    sSQL = sSQL & "FROM plasistencia asi "
    sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_natal,'%Y-%m-%d')>='" & Format(sFechaDesde, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_natal,'%Y-%m-%d')<='" & Format(sFechaHasta, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "ORDER BY 2 ASC"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      nDSubsidio = 0
      ' Fecha de inicio subsidio
      While Not porstRecord.EOF
        sFechaSubsidio = porstRecord!FSubsidio
        nDSubsidio = nDSubsidio + CLng(porstRecord!DSubsidio)
        porstRecord.MoveNext
      Wend
    End If
    porstRecord.Close
    
    ' Fecha inicial ejercicio subsidio
    If Not IsDate(sFechaInicio) Then
      sFechaInicio = "01/" & sMesProceso + "/" & sAnyoProceso
    End If
    
    ' Licencias continua inicio subsidio y calendario
    nDSubsidio = 0
    sSQL = "SELECT SUM(asi.diaprepostnatal) DSubsidio "
    sSQL = sSQL & "FROM plasistencia asi "
    sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
    sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_natal,'%Y-%m-%d')>='" & Format(sFechaInicio, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_natal,'%Y-%m-%d')<='" & Format(sFechaHasta, "yyyy-mm-dd") & "' "
    sSQL = sSQL & "GROUP BY asi.codcls,asi.codpsn"
    porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
      nDSubsidio = CLng(porstRecord!DSubsidio)
    End If
    porstRecord.Close
    nDCalendario = DateDiff("d", CDate(sFechaInicio), CDate(sFechaHasta)) + 1
   
    ' Determinar fecha real incio subsidio
    If CDate(sFechaSubsidio) > CDate(sFechaInicio) Then
      sFechaInicio = sFechaSubsidio
    End If
    sFechaInicio = DateAdd("d", 1, CDate(sFechaInicio))
    
    ' Determino periodo base subsidio
    If nDSubsidio = nDCalendario Then
  '      PProceso = Copy(sFechaInicio, 7, 4) + "/" + Copy(sFechaInicio, 4, 2)
    End If
  End If
  
  ' Fecha inicial ejercicio subsidio
  If Not IsDate(sFechaInicio) Then
    sFechaInicio = "01/" & sMesProceso + "/" & sAnyoProceso
  End If
  ' Determino periodos base subsidido
  sFechaDesde = DateAdd("m", -12, CDate(sFechaInicio))
  sFechaHasta = DateAdd("d", -1, CDate(sFechaInicio))
  ' Corrige fecha por ingreso
  sFechaDesde = IIf(CDate(sFechaIngreso) > CDate(sFechaDesde), sFechaIngreso, sFechaDesde)

  nRemuneraFija = 0
  ' Obtengo remuneracion base de subsidio, rango periodo
  sSQL = "SELECT ROUND(SUM(IFNULL(importe_mn, 0)), 2) nRemunera_mn, ROUND(SUM(IFNULL(importe_me, 0)), 2) nRemunera_me "
  sSQL = sSQL & "FROM plresultado "
  sSQL = sSQL & "WHERE codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND codcpc='" & sRemuneraFija & "' "
  sSQL = sSQL & "AND CONCAT(pdoano, pdomes)>='" & Format(sFechaDesde, "yyyymm") & "' "
  sSQL = sSQL & "AND CONCAT(pdoano, pdomes)<='" & Format(sFechaHasta, "yyyymm") & "'"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    nRemuneraFija = CDec(porstRecord!nRemunera_mn)
  End If
  porstRecord.Close
  
  ' Obtengo promedio de rermuneracion variable
  nMesesCalculo = (DatePart("m", CDate(sFechaHasta)) - DatePart("m", CDate(sFechaDesde))) + 1
  nMesesCalculo = nMesesCalculo + IIf(nMesesCalculo <= 0, 12, 0)

  mDBaseSubsidio = nMesesCalculo * 30
  nMesIncidencia = gdl_Funcion.Redondeo(nMesesCalculo / 2, 0)
  ' Diferencia periodo a Inicio Subsidio
  ' nDiferenciaMes = AMD(StrToDate("01" + Copy(PHasta, 5, 3) + "/" + Copy(PHasta, 1, 4)), FDESPRO, "M") - 1
'   If DiferenciaMes > 0 Then
'      nMesesCalculo = nMesesCalculo + nDiferenciaMes
  
  nRemuVariable1 = Promedio(sRemuneraVariable1, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable2 = Promedio(sRemuneraVariable2, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable3 = Promedio(sRemuneraVariable3, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable4 = Promedio(sRemuneraVariable4, nMesesCalculo, nMesIncidencia, "N", "M")
  nRemuVariable5 = Promedio(sRemuneraVariable5, nMesesCalculo, nMesIncidencia, "N", "M")
  
  ' Base de remuneracion diaria
  nRemueraVaria = Round(nRemuVariable1 + nRemuVariable2 + nRemuVariable3 + nRemuVariable4 + nRemuVariable5, 2)
  nRemueraSubsidio = Round(nRemuneraFija + nRemueraVaria, 2)
  nRemueraSubsidio = Round(nRemueraSubsidio / mDBaseSubsidio, 2)

  SubsidioMaternidad = nRemueraSubsidio
Finaliza:
  Err.Clear
  ' Destruyo los objetos
  Set gdl_Funcion = Nothing
  Set porstRecord = Nothing

End Function
Function UltimoMovimiento(ByVal sConcepto As String, ByVal sTablaResultado As String, ByVal nMesRetroceso As Integer, ByVal sTipoProceso As String) As Double
  '[ Retona el importe de ultimo movimiento del concepto segun los meses]
  '[ Tipo Resultado R:Resultado, C:CTS, G:Gratificación
  Dim sAnyoMesInicio As String
  Dim nImpConcepto As Double
  Dim nMesDiferencia As Integer
  
  On Error GoTo ControlError
    
  ' Obtengo periodo de inicio segun meses a retroceder
  sAnyoMesInicio = Format(sFechaIngreso, "yyyymm")
  If nMesRetroceso <> 0 Then
    nMesRetroceso = nMesRetroceso - 1
    sAnyoMesInicio = Format(DateAdd("m", -nMesRetroceso, CDate("01/" & sMesProceso & "/" & sAnyoProceso)), "yyyymm")
  End If
  
  ' Valido tipo proceso
  Select Case sTipoProceso
   Case "G"
    nMesDiferencia = IIf(CInt(sMesProceso) <= 6, 1, 7)
    nMesRetroceso = CInt(sMesProceso) - nMesDiferencia
    sAnyoMesInicio = Format(DateAdd("m", -nMesRetroceso, CDate("01/" & sMesProceso & "/" & sAnyoProceso)), "yyyymm")
   Case "C"
    If (CInt(sMesProceso) >= 1 And CInt(sMesProceso) <= 4) Then
      nMesDiferencia = 1
    ElseIf (CInt(sMesProceso) >= 5 And CInt(sMesProceso) <= 10) Then
      nMesDiferencia = -5
    ElseIf CInt(sMesProceso) >= 11 Then
      nMesDiferencia = -11
    End If
    nMesRetroceso = CInt(sMesProceso) + nMesDiferencia
    sAnyoMesInicio = Format(DateAdd("m", -nMesRetroceso, CDate("01/" & sMesProceso & "/" & sAnyoProceso)), "yyyymm")
  End Select
  
  ' Instancio al recordset
  Set porstRecord = New ADODB.Recordset
  
  ' Obtengo importe del concepto
  sSQL = "SELECT CONCAT(pdoano, pdomes) AS pdoanomes, IFNULL(res.importe_mn, 0) AS nRemunera_mn, IFNULL(res.importe_me, 0) AS nRemunera_me "
  sSQL = sSQL & "FROM " & sTablaResultado & " res "
  sSQL = sSQL & "WHERE res.codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND res.codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND res.codcpc='" & sConcepto & "' "
  sSQL = sSQL & "AND CONCAT(res.pdoano, res.pdomes)>='" & sAnyoMesInicio & "' "
  sSQL = sSQL & "AND CONCAT(res.pdoano, res.pdomes)>='" & Format(sFechaIngreso, "yyyymm") & "' "
  sSQL = sSQL & "AND CONCAT(res.pdoano, res.pdomes)<='" & sAnyoProceso & sMesProceso & "' "
  sSQL = sSQL & "ORDER BY pdoanomes DESC"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  nImpConcepto = 0
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    nImpConcepto = CDec(porstRecord!nRemunera_mn)
  End If
  porstRecord.Close
  Set porstRecord = Nothing
  UltimoMovimiento = CDec(nImpConcepto)
    
Exit Function
ControlError:
    Err.Clear

End Function
Function VacacionTrunca() As Double
  '[ Retona el numero dias de vacaciones truncas ]
  Dim nDiasCalculo As Long, nDiasAusencia As Long, nDiaVacGozada As Long
  Dim dFechaIngreso As String, sFechaProceso As String
  Dim sFechaInicio As String, sFechaFinal As String
  Dim sPerVacGozada As String
    
  ' Configuro los objetos de clases
  Set gdl_Funcion = CreateObject("syslink.Funciones")
  Set porstRecord = New ADODB.Recordset
    
  nDiasCalculo = 0
  ' Obtengo las fechas de inicio y fin de vacaciones
  sSQL = "SELECT asi.fechainiliqvaca AS fechainicio, asi.fechafinliqvaca AS fechafinal, "
  sSQL = sSQL & "asi.fechacese, dxr.fecingreso "
  sSQL = sSQL & "FROM plasistencia asi "
  sSQL = sSQL & "INNER JOIN pldatoresultado dxr ON asi.codcls=dxr.codcls AND asi.codpdo=dxr.codpdo AND asi.codpsn=dxr.codpsn "
  sSQL = sSQL & "WHERE asi.codcls='" & sClasePlanilla & "' "
  sSQL = sSQL & "AND asi.codpdo='" & sCodigoPeriodo & "' "
  sSQL = sSQL & "AND asi.codpsn='" & sCodigoEmpleado & "' "
  sSQL = sSQL & "AND dxr.estadopsn='I'"
  porstRecord.Open sSQL, sCadenaConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not ((porstRecord.EOF And porstRecord.BOF) Or porstRecord.RecordCount = 0) Then
    dFechaIngreso = Format(porstRecord!fecingreso, "dd/mm/yyyy")
    sFechaInicio = Format(porstRecord!fechainicio, "dd/mm/yyyy")
    sFechaFinal = Format(porstRecord!fechafinal, "dd/mm/yyyy")
    sFechaProceso = Format(porstRecord!fechacese, "dd/mm/yyyy")
    ' Cierro los objetos
    porstRecord.Close
    
    ' Record minimo un mes de trabajo
    If (DateDiff("d", DateAdd("d", -1, DateAdd("m", 1, CDate(dFechaIngreso))), CDate(sFechaProceso)) >= 0) Then
      ' Año de ingreso y procesos
      sFechaInicio = Left(dFechaIngreso, 6) & IIf((sAnyoProceso = Right(dFechaIngreso, 4)) Or (sAnyoProceso <> Right(dFechaIngreso, 4) And (sMesProceso > Mid(dFechaIngreso, 4, 2))), sAnyoProceso + 1, sAnyoProceso)
      sFechaInicio = Format(DateAdd("yyyy", -1, CDate(sFechaInicio)), "dd/mm/yyyy")
      sFechaFinal = sFechaProceso
      If sFechaInicio <> "" And sFechaFinal <> "" Then
        ' Obtengo dias vacaciones truncas
        nDiasCalculo = gdl_Funcion.NumeroDias360(CDate(sFechaProceso), CDate(sFechaInicio), CDate(sFechaFinal))
        nDiasAusencia = 0
        ' Resta Ausencias encontradas en el período (Si el parámetro así lo indica)
        If sDesAusenciaBF = s_Estado_Act Then
          nDiasAusencia = gdl_Funcion.DiasAusenciaBS(sCadenaConexion, sClasePlanilla, sCodigoEmpleado, sFechaInicio, sFechaProceso)
        End If
        nDiasCalculo = nDiasCalculo - nDiasAusencia
        ' Descuenta vacaciones adelantadas del periodo
        sPerVacGozada = Right(sFechaInicio, 4) & Right(sFechaFinal, 4)
        nDiaVacGozada = DiasVacacionGozada(sPerVacGozada, sFechaInicio, sFechaProceso)
        nDiaVacGozada = (nDiaVacGozada * 360) / 30
        
        nDiasCalculo = nDiasCalculo - nDiaVacGozada
        nDiasCalculo = IIf(nDiasCalculo < 0, 0, nDiasCalculo)
      End If
    End If
  End If
  VacacionTrunca = CDec(nDiasCalculo)

Finaliza:
  ' Destruyo los objetos
  Set porstRecord = Nothing
  Set gdl_Funcion = Nothing

End Function
