VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Funciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                     ' Declarar variable antes de usarla
Private rs_Consulta As New ADODB.Recordset          ' Recodset para efectos de consulta
Public Function ValidaEmail(ByVal strEmailAdd As String) As Boolean
  With CreateObject("VBScript.RegExp")
    .IgnoreCase = True
    .Global = True
    .Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    ValidaEmail = .Test(strEmailAdd)
  End With
End Function
Function aTexto(ByVal Expresion) As String
If IsNull(Expresion) Or IsEmpty(Expresion) Then
  aTexto = ""
Else
  aTexto = Trim$(Expresion)
End If
End Function
Function DameCuentaCorriente(ByVal s_Conexion As String, ByVal s_Planilla As String, ByVal s_Personal As String, ByVal s_Periodo As String, ByVal s_Moneda As String, ByRef n_Descuento As Double) As Double
' Elimino y genero la tabla temporal de cuenta corriente
sSQL = "DROP TABLE IF EXISTS tmpcuentacte"
rs_Consulta.Open sSQL, s_Conexion, adOpenForwardOnly, adLockReadOnly, adCmdText
sSQL = "CREATE TEMPORARY TABLE tmpcuentacte "
sSQL = sSQL & "SELECT DISTINCTROW cte.codcls, cte.codpsn, cte.numctacte, cte.numcuota, "
sSQL = sSQL & "cte.codpdoprv, cte.codcpc, cte.codmon, cte.estadoctacte, pdo.tipocambio, "
sSQL = sSQL & "ROUND(IFNULL(IF(cte.codmon='" & s_Codmon_mn & "', cte.cargo_mn, (cte.cargo_me * pdo.tipocambio)), 0), 2) AS cargo_mn, "
sSQL = sSQL & "ROUND(IFNULL(IF(cte.codmon='" & s_Codmon_me & "', cte.cargo_me, (cte.cargo_mn / pdo.tipocambio)), 0), 2) AS cargo_me "
sSQL = sSQL & "FROM plcuentacte cte "
sSQL = sSQL & "INNER JOIN plperiodo pdo ON cte.codcls=pdo.codcls AND cte.codpdoprv=pdo.codpdo "
sSQL = sSQL & "WHERE cte.codcls='" & s_Planilla & "' "
sSQL = sSQL & "AND cte.codpsn='" & s_Personal & "' "
sSQL = sSQL & "AND pdo.codpdo<='" & s_Periodo & "' "
sSQL = sSQL & "AND cte.numcuota=" & s_Estado_Ina & " "
sSQL = sSQL & "ORDER BY cte.codpsn, cte.numctacte, cte.numcuota"
rs_Consulta.Open sSQL, , adOpenForwardOnly, adLockReadOnly, adCmdText

' Obtengo el saldo inicial
sSQL = "SELECT DISTINCTROW cte.codpsn, "
sSQL = sSQL & "ROUND(SUM(IFNULL(IF(cte.codmon='" & s_Codmon_mn & "', cte.abono_mn, (cte.abono_me * tmp.tipocambio)), 0)), 2) AS abono_mn, "
sSQL = sSQL & "ROUND(SUM(IFNULL(IF(cte.codmon='" & s_Codmon_me & "', cte.abono_me, (cte.abono_mn / tmp.tipocambio)), 0)), 2) AS abono_me "
sSQL = sSQL & "FROM plcuentacte cte "
sSQL = sSQL & "INNER JOIN tmpcuentacte tmp ON cte.codcls=tmp.codcls AND cte.codpsn=tmp.codpsn AND cte.numctacte=tmp.numctacte "
sSQL = sSQL & "WHERE cte.codcls='" & s_Planilla & "' "
sSQL = sSQL & "AND cte.codpsn='" & s_Personal & "' "
sSQL = sSQL & "AND cte.codpdoprv>='" & s_Periodo & "' "
sSQL = sSQL & "AND cte.numcuota<>" & s_Estado_Ina & " "
sSQL = sSQL & "GROUP BY cte.codcls, cte.codpsn "
sSQL = sSQL & "ORDER BY cte.codpsn "
rs_Consulta.Open sSQL, , adOpenForwardOnly, adLockReadOnly, adCmdText
If Not (rs_Consulta.EOF Or rs_Consulta.BOF) Then
  DameCuentaCorriente = CDec(rs_Consulta("abono_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me")))
End If
rs_Consulta.Close

' Elimino y genero la tabla temporal de cuenta corriente por concepto
sSQL = "DROP TABLE IF EXISTS tmpcuentactex"
rs_Consulta.Open sSQL, , adOpenForwardOnly, adLockReadOnly, adCmdText
  
sSQL = "CREATE TEMPORARY TABLE tmpcuentactex "
sSQL = sSQL & "SELECT DISTINCTROW cte.codcls, cte.codpsn, cte.codcpc, "
sSQL = sSQL & "ROUND(SUM(IFNULL(IF(cte.codmon='" & s_Codmon_mn & "', cte.cargo_mn, (cte.cargo_me * pdo.tipocambio)), 0)), 2) AS cargo_mn, "
sSQL = sSQL & "ROUND(SUM(IFNULL(IF(cte.codmon='" & s_Codmon_me & "', cte.cargo_me, (cte.cargo_mn / pdo.tipocambio)), 0)), 2) AS cargo_me "
sSQL = sSQL & "FROM plcuentacte cte "
sSQL = sSQL & "INNER JOIN plperiodo pdo ON cte.codcls=pdo.codcls AND cte.codpdoprv=pdo.codpdo "
sSQL = sSQL & "WHERE cte.codcls='" & s_Planilla & "' "
sSQL = sSQL & "AND cte.codpsn='" & s_Personal & "' "
sSQL = sSQL & "AND pdo.codpdo<='" & s_Periodo & "' "
sSQL = sSQL & "AND cte.numcuota=" & s_Estado_Ina & " "
sSQL = sSQL & "GROUP BY cte.codcls, cte.codpsn, cte.codcpc "
sSQL = sSQL & "ORDER BY cte.codpsn"
rs_Consulta.Open sSQL, , adOpenForwardOnly, adLockReadOnly, adCmdText

' Obtengo el descuento del periodo
sSQL = "SELECT DISTINCTROW res.codpsn, "
sSQL = sSQL & "ROUND(SUM(IFNULL(tmp.cargo_mn, 0)), 2) AS cargo_mn, "
sSQL = sSQL & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS abono_mn, "
sSQL = sSQL & "ROUND(SUM(IFNULL(tmp.cargo_me, 0)), 2) AS cargo_me, "
sSQL = sSQL & "ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS abono_me "
sSQL = sSQL & "FROM plresultado res "
sSQL = sSQL & "INNER JOIN tmpcuentactex tmp ON res.codcls=tmp.codcls AND res.codpsn=tmp.codpsn AND res.codcpc=tmp.codcpc "
sSQL = sSQL & "WHERE res.codcls='" & s_Planilla & "' "
sSQL = sSQL & "AND res.codpsn='" & s_Personal & "' "
sSQL = sSQL & "AND res.codpdo='" & s_Periodo & "' "
sSQL = sSQL & "GROUP BY res.codcls, res.codpsn "
sSQL = sSQL & "ORDER BY res.codpsn"
rs_Consulta.Open sSQL, , adOpenForwardOnly, adLockReadOnly, adCmdText
If Not (rs_Consulta.EOF Or rs_Consulta.BOF) Then
  n_Descuento = CDec(rs_Consulta("abono_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me")))
End If
rs_Consulta.Close: Set rs_Consulta = Nothing

End Function
Function DameDescripcion(ByVal s_Conexion As String, ByVal s_Empresa As String, ByVal s_CodigoBus As String, ByVal s_Caso As String) As String
Dim l_ExistRecord As Boolean

' Verifico que se ingrese un valor de busqueda
s_CodigoBus = aTexto(s_CodigoBus)
If s_CodigoBus = "" Or s_CodigoBus = "," Then DameDescripcion = "": Exit Function
    
' Armo Cadena de Sql de Busqueda
Select Case s_Caso
 Case "AC": sSQL = "SELECT desact FROM plactividad WHERE codact='" & s_CodigoBus & "'"
 Case "BP": sSQL = "SELECT desboleta FROM plboletapago WHERE codcls='" & s_Empresa & "' AND codboleta='" & s_CodigoBus & "'"
 Case "CA": sSQL = "SELECT aliascpc FROM plconcepto WHERE codcpc = '" & s_CodigoBus & "'"
 Case "CC": sSQL = "SELECT detcco FROM cocco WHERE codcco='" & s_CodigoBus & "'"
 Case "CO": sSQL = "SELECT destic FROM pltipcom WHERE codtic = '" & s_CodigoBus & "'"
 Case "CP": sSQL = "SELECT descpc FROM plconcepto WHERE codcpc = '" & s_CodigoBus & "'"
 Case "CT": sSQL = "SELECT tipocpc FROM plconcepto WHERE codcpc = '" & s_CodigoBus & "'"
 Case "CU": sSQL = "SELECT detcta FROM cocta WHERE tpocta='" & s_Empresa & "' AND codcta='" & s_CodigoBus & "'"
 Case "CX": sSQL = "SELECT descon FROM plconcesunat WHERE codcon = '" & s_CodigoBus & "'"
 Case "CY": sSQL = "SELECT descao FROM plcatocu WHERE codcao = '" & s_CodigoBus & "'"
 Case "CZ": sSQL = "SELECT desctr FROM plconven WHERE codctr = '" & s_CodigoBus & "'"
 Case "DC": sSQL = "SELECT descgo FROM plcargo WHERE codcls='" & s_Empresa & "' AND codcgo='" & s_CodigoBus & "'"
 Case "DI": sSQL = "SELECT desdci FROM pldocidentidad WHERE coddci = '" & s_CodigoBus & "'"
 Case "DS": sSQL = "SELECT sigladci FROM pldocidentidad WHERE coddci = '" & s_CodigoBus & "'"
 Case "EB": sSQL = "SELECT desbco FROM plbanco WHERE codbco = '" & s_CodigoBus & "'"
 Case "EC": sSQL = "SELECT descricts FROM plctsperiodo WHERE codcls='" & s_Empresa & "' AND pdocts='" & s_CodigoBus & "'"
 Case "EM": sSQL = "SELECT razemp FROM tgemp WHERE codemp= '" & s_CodigoBus & "'"
 Case "EO": sSQL = "SELECT desepr FROM plestablecimientopropio WHERE codepr= '" & s_CodigoBus & "'"
 Case "EP": sSQL = "SELECT desafp FROM plentidadafp WHERE codafp='" & s_CodigoBus & "'"
 Case "ES": sSQL = "SELECT deseps FROM plentidadeps WHERE codeps='" & s_CodigoBus & "'"
 Case "ID": sSQL = "SELECT Format(now, 'yyyy-mm-dd hh:mm:ss') FROM sistemas WHERE sis_codigo = '" & s_CodigoBus & "'"
 Case "FP": sSQL = "SELECT desmof FROM plmotfin WHERE codmof = '" & s_CodigoBus$ & "'"
 Case "GR": sSQL = "SELECT desniv FROM plniveducativo WHERE codniv='" & s_CodigoBus & "'"
 Case "LD": sSQL = "SELECT desldn FROM plcodigoldn WHERE codldn='" & s_CodigoBus & "'"
 Case "MF": sSQL = "SELECT desmfo FROM plmodforma WHERE codmfo = '" & s_CodigoBus$ & "'"
 Case "MP": sSQL = "SELECT mespdo FROM plperiodo WHERE codcls='" & s_Empresa & "' AND codpdo='" & s_CodigoBus & "'"
 Case "NA": sSQL = "SELECT desnac FROM plnacionalidad WHERE codnac = '" & s_CodigoBus & "'"
 Case "NS": sSQL = "SELECT sis_descripcion FROM sistemas WHERE sis_codigo = '" & s_CodigoBus & "'"
 Case "NU": sSQL = "SELECT nomusr FROM sgusr WHERE codusr='" & s_CodigoBus & "'"
 Case "PA": sSQL = "SELECT destip FROM pltippago WHERE codtip='" & s_CodigoBus$ & "'"
 Case "PC": sSQL = "SELECT desproce FROM plproceso WHERE codcls='" & s_Empresa & "' AND codproce='" & s_CodigoBus & "'"
 Case "PD": sSQL = "SELECT despemi FROM plpaisemidocum WHERE codpemi='" & s_CodigoBus & "'"
 Case "PE": sSQL = "SELECT desprd FROM plperiodicidad WHERE codprd='" & s_CodigoBus$ & "'"
 Case "PF": sSQL = "SELECT despfs FROM plprofesion WHERE codpfs='" & s_CodigoBus & "'"
 Case "PL": sSQL = "SELECT despll FROM plplanilla WHERE codcls='" & s_Empresa & "' AND codpll='" & s_CodigoBus & "'"
 Case "PR": sSQL = "SELECT despdo FROM plperiodo WHERE codcls='" & s_Empresa & "' AND codpdo='" & s_CodigoBus & "'"
 Case "PV": sSQL = "SELECT descripvs FROM plpvsperiodovac WHERE codcls='" & s_Empresa & "' AND codpvs='" & s_CodigoBus & "'"
 Case "PW": sSQL = "SELECT clausr FROM sgusr WHERE codusr='" & s_CodigoBus & "'"
 Case "QM": sSQL = "SELECT desqmd FROM plempresasqmdes WHERE codqmd='" & s_CodigoBus & "'"
 Case "RA": sSQL = "SELECT gratixasis FROM plcfgempresa WHERE pdoano = '" & s_CodigoBus & "'"
 Case "RF": sSQL = "SELECT tre_descripcion FROM tipreferencia WHERE tre_empresa = '" & s_Empresa & "' AND tre_codigo = '" & s_CodigoBus & "'"
 Case "RH": sSQL = "SELECT formarpt FROM plgenreporte WHERE codcls='" & s_Empresa & "' AND codrpt='" & s_CodigoBus & "'"
 Case "RP": sSQL = "SELECT desrpt FROM plgenreporte WHERE codcls='" & s_Empresa & "' AND codrpt='" & s_CodigoBus & "'"
 Case "RT": sSQL = "SELECT titulorpt FROM plgenreporte WHERE codcls='" & s_Empresa & "' AND codrpt='" & s_CodigoBus & "'"
 Case "RX": sSQL = "SELECT deseqd FROM plempresasqdes WHERE codeqd='" & s_CodigoBus & "'"
 Case "RU": sSQL = "SELECT desepr as desepr,cdgepr as cdgepr,tasepr as tasepr FROM plestablecimientopropio WHERE codepr='" & s_CodigoBus & "' union all select deseed as desepr,esteed as cdgepr, taseed as tasepr from plempresaseqdes WHERE codeed='" & s_CodigoBus & "'"
 Case "SC": sSQL = "SELECT descrisub FROM plctsperiodosub WHERE codcls='" & s_Empresa & "' AND pdocts='" & Left(s_CodigoBus, InStr(1, s_CodigoBus, "|") - 1) & "' AND subcts='" & Mid(s_CodigoBus, InStr(1, s_CodigoBus, "|") + 1) & "'"
 Case "SE": sSQL = "SELECT dessec FROM plseccion WHERE codsec='" & s_CodigoBus & "'"
 Case "SI": sSQL = "SELECT desstp FROM plsitrapen WHERE codstp='" & s_CodigoBus & "'"
 Case "ST": sSQL = "SELECT descdt FROM plconditrabajo WHERE codcls='" & s_Empresa & "' AND codcdt='" & s_CodigoBus & "'"
 Case "TA": sSQL = "SELECT tan_descripcion FROM tipoanalisis WHERE tan_empresa = '" & s_Empresa & "' AND tan_codigo = '" & s_CodigoBus & "'"
 Case "TB": sSQL = "SELECT destbl FROM pltablabase WHERE codcls='" & s_Empresa & "' AND pdoano='" & Left(s_CodigoBus, 4) & "' AND codtbl='" & Mid(s_CodigoBus, 5) & "'"
 Case "TC": sSQL = "SELECT destco FROM pltipcontrato WHERE codtco='" & s_CodigoBus & "'"
 Case "TE": sSQL = "SELECT desest FROM plestablecimiento WHERE codest='" & s_CodigoBus & "'"
 Case "TP": sSQL = "SELECT tpopdo FROM plperiodo WHERE codcls='" & s_Empresa & "' AND codpdo='" & s_CodigoBus & "'"
 Case "TS": sSQL = "SELECT destsu FROM pltipsusp WHERE codtsu='" & s_CodigoBus & "'"
 Case "TT": sSQL = "SELECT destpt FROM pltpotrabajador WHERE codtpt='" & s_CodigoBus & "'"
 Case "TV": sSQL = "SELECT desvia FROM pltipovia WHERE codvia = '" & s_CodigoBus$ & "'"
 Case "TZ": sSQL = "SELECT deszona FROM pltipozona WHERE codzona = '" & s_CodigoBus$ & "'"
 Case "UB": sSQL = "SELECT desubg FROM tgubigeo WHERE nivelubg = " & Val(s_Empresa) & " AND codubg = '" & s_CodigoBus & "'"
 Case "UG": sSQL = "SELECT CONCAT(dp.desubg, '/', pr.desubg, '/', ds.desubg) AS desubg"
  sSQL = sSQL & " FROM ((tgubigeo ds"
  sSQL = sSQL & " LEFT JOIN tgubigeo pr ON LEFT(ds.codubg, 4)=pr.codubg)"
  sSQL = sSQL & " LEFT JOIN tgubigeo dp ON LEFT(ds.codubg, 2)=dp.codubg)"
  sSQL = sSQL & " WHERE ds.nivelubg='2'"
  sSQL = sSQL & " AND ds.codubg='" & s_CodigoBus & "'"
  sSQL = sSQL & " AND pr.nivelubg='1'"
  sSQL = sSQL & " AND dp.nivelubg='0'"
 Case "UL": sSQL = "SELECT desubica FROM plubicacion WHERE codubica='" & s_CodigoBus & "'"
 Case "VC": sSQL = "SELECT descripcion FROM plvarfunc WHERE tipo='" & Left(s_CodigoBus, 1) & "' AND codigo='" & Mid(s_CodigoBus, 2) & "'"
 Case "VF": sSQL = "SELECT descripcion FROM plvarfunc WHERE tipo='" & Left(s_CodigoBus, 1) & "' AND nombre='" & Mid(s_CodigoBus, 2) & "'"
End Select

' Aperturo el Recordset
rs_Consulta.Open sSQL, s_Conexion, adOpenForwardOnly, adLockReadOnly, adCmdText
l_ExistRecord = (rs_Consulta.EOF Or rs_Consulta.BOF)
' Obtengo la Descripcion
If Not l_ExistRecord Then
   DameDescripcion = aTexto(rs_Consulta(0))
Else
   DameDescripcion = "???"
End If
rs_Consulta.Close: Set rs_Consulta = Nothing
    
End Function
Function Desencripta(ByVal s_Expresion As String) As String
  Dim nLenExpresion As Long, nLenSigla As Long
  Dim nSecuencia As Long, nControl As Long, nFase As Long
  Dim sLetraExpresion As String, sLetraSigla As String, sEncripta As String
    
  nLenExpresion = Len(s_Expresion)
  nLenSigla = Len(s_Sigla)
  For nSecuencia = 1 To nLenExpresion
    sLetraExpresion = Mid(s_Expresion, nSecuencia, 1)
    If nControl = nLenSigla Then nControl = 0
    nControl = nControl + 1
    sLetraSigla = Mid(s_Sigla, nControl, 1)
    nFase = Asc(sLetraExpresion) - Asc(sLetraSigla)
    If nFase < 0 Then
      nFase = nFase + 255
    End If
    sEncripta = sEncripta & Chr(nFase)
  Next nSecuencia
  Desencripta = sEncripta

End Function
Function DiasAusenciaBS(ByVal s_Conexion As String, ByVal s_ClasePlanilla As String, ByVal s_Empleado As String, ByVal s_FechaDesde As String, ByVal s_FechaHasta As String) As Long
'[ Retona los dias de ausencia rango de fecha]
  
  sSQL = "SELECT ("
  sSQL = sSQL & "IFNULL((SELECT SUM(IFNULL((asi.diafalta + asi.diasuspension), 0)) nDias "
  sSQL = sSQL & "FROM plasistencia asi "
  sSQL = sSQL & "INNER JOIN plperiodo pdo ON pdo.codcls=asi.codcls AND asi.codpdo=pdo.codpdo "
  sSQL = sSQL & "WHERE asi.codcls='" & s_ClasePlanilla & "' "
  sSQL = sSQL & "AND asi.codpsn='" & s_Empleado & "' "
  sSQL = sSQL & "AND CONCAT(pdo.anopdo, pdo.mespdo)>='" & Format(s_FechaDesde, "yyyymm") & "' "
  sSQL = sSQL & "AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & Format(s_FechaHasta, "yyyymm") & "'),0) "
  sSQL = sSQL & "+ "
  sSQL = sSQL & "IFNULL((SELECT SUM(IFNULL(asi.licencia, 0)) nDias "
  sSQL = sSQL & "FROM plasistencia asi "
  sSQL = sSQL & "INNER JOIN plperiodo pdo ON pdo.codcls=asi.codcls AND asi.codpdo=pdo.codpdo "
  sSQL = sSQL & "WHERE asi.codcls='" & s_ClasePlanilla & "' "
  sSQL = sSQL & "AND asi.codpsn='" & s_Empleado & "' "
  sSQL = sSQL & "AND asi.codmdi_licen IN('01','05','07') "
  sSQL = sSQL & "AND DATE_FORMAT(asi.fechaini_licen,'%Y-%m-%d')>='" & Format(s_FechaDesde, "yyyy-mm-dd") & "' "
  sSQL = sSQL & "AND DATE_FORMAT(asi.fechafin_licen,'%Y-%m-%d')<='" & Format(s_FechaHasta, "yyyy-mm-dd") & "'),0)) nDias"
  rs_Consulta.Open sSQL, s_Conexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  
  If Not (rs_Consulta.EOF And rs_Consulta.BOF) Then
    DiasAusenciaBS = CLng(rs_Consulta!nDias)
  End If
  rs_Consulta.Close: Set rs_Consulta = Nothing

  sSQL = "call sf_Pla_DiasAusenciaBS("
  sSQL = sSQL & "'" & s_ClasePlanilla & "', '" & s_Empleado & "', "
  sSQL = sSQL & "'" & Format(s_FechaDesde, "yyyy-mm-dd") & "', '" & Format(s_FechaHasta, "yyyy-mm-dd") & "')"
'  Set porstBancaria = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)

End Function
Function DiferenciaFechas(ByVal d_FechaInicio As Date, ByVal d_FechaFinal As Date, s_Unidad As String) As Long
  ' MD: La diferencia entre los días en fecha_inicial y fecha_final. Los meses y años de las fechas no se consideran
  ' YM: La diferencia entre los meses de fecha_inicial y fecha_final. Los días y años de las fechas no se consideran
  ' YD: La diferencia entre los días de fecha_inicial y fecha_final. Los años de las fechas no se consideran
  Dim nDiaDesde As Integer, nMesDesde As Integer, nAnyoDesde As Integer
  Dim nDiaHasta As Integer, nMesHasta As Integer, nAnyoHasta As Integer
  Dim nDiasAjuste As Integer, nMesesAjuste As Integer, nAnyosAjuste As Integer
  Dim nDiasMes As Integer, nDesgloseFecha As Integer
  
  DiferenciaFechas = 0
  ' Fecha de proceso es menor i/o igual que la fecha final
  If (d_FechaFinal >= d_FechaInicio) Then
    ' Caso unidad analisis D: El número de días en el período
    If s_Unidad = "D" Then nDesgloseFecha = DateDiff("d", d_FechaInicio, d_FechaFinal) + 1: GoTo Finalizar
    
    '[ Primero : obtengo dias truncos e inicializo fecha inicial
    nDiaDesde = DatePart("d", d_FechaInicio)
    nMesDesde = DatePart("m", d_FechaInicio)
    nAnyoDesde = DatePart("yyyy", d_FechaInicio)
    ' Segundo : obtengo dias truncos e inicializo fecha final o proceso
    nDiaHasta = DatePart("d", d_FechaFinal)
    nMesHasta = DatePart("m", d_FechaFinal)
    nAnyoHasta = DatePart("yyyy", d_FechaFinal)
    nAnyoHasta = nAnyoHasta - IIf(nMesHasta = 12, 0, 1)
    ' Tercero : obtengo ajuste de meses y dias
    nDiasMes = NumeroDiasMes(nMesHasta, nAnyoHasta)
    nDiasAjuste = IIf((nDiaHasta = nDiasMes), 0, nDiaHasta) - nDiaDesde + 1
    nMesesAjuste = IIf((nMesHasta = 12), 0, nMesHasta) - nMesDesde + IIf((nDiaHasta = nDiasMes), 1, 0)
    nAnyosAjuste = (nAnyoHasta - nAnyoDesde) + 1
    ']
    
    ' Caso unidad analisis Y: El número de años completos en el período
    If s_Unidad = "Y" Then nDesgloseFecha = nAnyosAjuste: GoTo Finalizar
    ' Caso unidad analisis M: El número de meses completos en el período
    If s_Unidad = "M" Then
      nDesgloseFecha = (nAnyosAjuste * 12) + nMesesAjuste + IIf(nDiasAjuste >= 30, 1, 0)
      GoTo Finalizar
    End If

  End If
Finalizar:
  DiferenciaFechas = nDesgloseFecha
  
End Function
Function Encripta(ByVal s_Expresion As String) As String
  Dim nLenExpresion As Long, nLenSigla As Long
  Dim nSecuencia As Long, nControl As Long, nFase As Long
  Dim sLetraExpresion As String, sLetraSigla As String, sEncripta As String
    
    ' Cantidad de caracteres...
  nLenExpresion = Len(s_Expresion)
  nLenSigla = Len(s_Sigla)
  For nSecuencia = 1 To nLenExpresion
    sLetraExpresion = Mid(s_Expresion, nSecuencia, 1)
    If nControl = nLenSigla Then nControl = 0
    nControl = nControl + 1
    sLetraSigla = Mid(s_Sigla, nControl, 1)
    nFase = Asc(sLetraExpresion) + Asc(sLetraSigla)
    If nFase > 255 Then
      nFase = nFase - 255
    End If
    sEncripta = sEncripta & Chr(nFase)
  Next nSecuencia
  Encripta = sEncripta

End Function
Function Execution(ByVal s_Conexion As String, ByVal s_Sentencia As String) As Boolean
Dim pcn_Conexion As ADODB.Connection
Dim n_RecordAffected As Long

Execution = False
'[ Instancio e inicio la conexión a la base de datos ]
Set pcn_Conexion = New ADODB.Connection
pcn_Conexion.Open s_Conexion
pcn_Conexion.BeginTrans           ' Inicia transacción

On Error GoTo ErrorExecute
' Realizo el proceso de actualización de los registros
pcn_Conexion.Execute s_Sentencia, n_RecordAffected
If pcn_Conexion.Errors.Count > 0 Then
  pcn_Conexion.RollbackTrans      ' Cancelo la transacción
  GoTo ErrorExecute
End If
pcn_Conexion.CommitTrans          ' Confirma transacción

Execution = True
GoTo Salir

ErrorExecute:
  MsgBox "Ocurrio un error al tratar de Actualizar la BD" & vbCrLf & vbCrLf & _
  "Fuente: " & Err.Source & vbCrLf & _
  "Numero: " & Err.Number & vbCrLf & _
  "Descripcion: " & Err.Description & vbCrLf, vbCritical, "ERROR"
  pcn_Conexion.RollbackTrans

Salir:
'[ Finalizo la conexión a la base de datos ]
pcn_Conexion.Close: Set pcn_Conexion = Nothing
On Error GoTo 0

End Function
Function ExisteArchivo(ByVal s_Archivo As String) As Boolean
  Dim nAtributo As Integer
  On Error Resume Next
  nAtributo = GetAttr(s_Archivo)
  If Err.Number Then
    Err.Clear
    ExisteArchivo = False
  Else
    ExisteArchivo = True
  End If
End Function
Function HelpTablas(ByVal s_Tabla As String, ByVal s_Campo As String, Optional ByVal s_Condicion As String, Optional ByVal s_Filtro As String) As String
  Dim n_Longitud As Integer
  Dim s_Where As String
  n_Longitud = Len(s_Filtro)
  s_Where = " AND "
  sSQL = ""
  Select Case s_Tabla
   Case "act"           ' Tipo de Actividades
    sSQL = "SELECT codact, desact FROM plactividad"
    sSQL = sSQL & " WHERE estadoact='" & s_Estado_Act & "'"
    s_Campo = "desact"
   Case "afp"           ' Entidad de pensiones
    sSQL = "SELECT codafp, desafp FROM plentidadafp"
    sSQL = sSQL & " WHERE estadoafp='" & s_Estado_Act & "'"
   Case "bco"           ' Entidad bancaria
    sSQL = "SELECT codbco, desbco FROM plbanco"
    sSQL = sSQL & " WHERE estadobco='" & s_Estado_Act & "'"
   Case "bol"           ' Formato de boleta de pago
    sSQL = "SELECT codboleta, desboleta FROM plboletapago"
    sSQL = sSQL & " WHERE codcls='" & s_Condicion & "'"
   Case "cao"           ' Categoria Ocupacional
    sSQL = "SELECT codcao, descao "
    sSQL = sSQL & "FROM plcatocu "
    sSQL = sSQL & "WHERE estadocao='" & s_Estado_Act & "'"
   Case "cco"           ' Centro de costo por nivel
    sSQL = "SELECT codcco, detcco FROM cocco"
    sSQL = sSQL & " WHERE LENGTH(codcco)>=" & s_Condicion
   Case "cdt"           ' Condición de trabajo
    sSQL = "SELECT codcdt, descdt FROM plconditrabajo"
    sSQL = sSQL & " WHERE codcls='" & s_Condicion & "'"
    sSQL = sSQL & " AND estadocdt='" & s_Estado_Act & "'"
   Case "ced"           ' Periodos de provisión de cts estado diferente
    sSQL = "SELECT pdocts, descricts "
    sSQL = sSQL & "FROM plctsperiodo "
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 2) & "' "
    sSQL = sSQL & " AND estadocts<>'" & Left(s_Condicion, 1) & "'"
   Case "cgo"           ' Cargo de personal
    sSQL = "SELECT codcgo, descgo FROM plcargo"
    sSQL = sSQL & " WHERE codcls='" & s_Condicion & "'"
    sSQL = sSQL & " AND estadocgo='" & s_Estado_Act & "'"
   Case "cit"           ' Conceptos por clase de planilla imprimir y tipo
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc "
    sSQL = sSQL & "FROM plconcepto cpc "
    sSQL = sSQL & "INNER JOIN plconceplanilla cxp ON cpc.codcpc=cxp.codcpc AND cxp.codcls='" & Left(s_Condicion, 2) & "' "
    sSQL = sSQL & "WHERE cpc.estadocpc='" & s_Estado_Act & "' "
    sSQL = sSQL & "AND cpc.tipocpc IN('" & Mid(s_Condicion, 3) & "') "
    sSQL = sSQL & "AND cxp.impbolecpc='" & s_Estado_Act & "'"
   Case "con"           ' Conceptos Sunat
    sSQL = "SELECT codcon, descon,tipcon FROM plconcesunat"
    sSQL = sSQL & " WHERE estadocon='" & s_Estado_Act & "'"
    's_Campo = " codcon + 0 "
   Case "cpc"           ' Concepto de calculo por tipo
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc "
    sSQL = sSQL & "FROM plconcepto cpc "
    sSQL = sSQL & "WHERE cpc.tipocpc IN('" & s_Condicion & "') "
    sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "'"
   Case "cpp"           ' Conceptos por planilla por proceso de cálculo
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc"
    sSQL = sSQL & " FROM (plconceplanilla cxp"
    sSQL = sSQL & " LEFT JOIN plconcepto cpc ON cxp.codcpc=cpc.codcpc)"
    sSQL = sSQL & " WHERE cxp.codcls='" & Left(s_Condicion, 2) & "'"
    sSQL = sSQL & " AND NOT EXISTS(SELECT * FROM plconceproceso cpp"
    sSQL = sSQL & " WHERE cpp.codproce='" & Mid(s_Condicion, 3) & "'"
    sSQL = sSQL & " AND cpp.codcls=cxp.codcls AND cpp.codcpc=cxp.codcpc)"
    sSQL = sSQL & " AND cpc.estadocpc='" & s_Estado_Act & "'"
   Case "cpi"           ' Conceptos por clase de planilla imprimir
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc "
    sSQL = sSQL & "FROM plconcepto cpc "
    sSQL = sSQL & "INNER JOIN plconceplanilla cxp ON cpc.codcpc=cxp.codcpc AND cxp.codcls='" & s_Condicion & "' "
    sSQL = sSQL & "WHERE cpc.estadocpc='" & s_Estado_Act & "' "
    sSQL = sSQL & "AND cxp.impbolecpc='" & s_Estado_Act & "'"
   Case "cpt"           ' Conceptos por clase de planilla todos
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc "
    sSQL = sSQL & "FROM plconcepto cpc "
    sSQL = sSQL & "INNER JOIN plconceplanilla cxp ON cpc.codcpc=cxp.codcpc AND cxp.codcls='" & s_Condicion & "' "
    sSQL = sSQL & "WHERE cpc.estadocpc='" & s_Estado_Act & "'"
   Case "cta"           ' Cuenta contable
    sSQL = "SELECT codcta, detcta, tpocta "
    sSQL = sSQL & "FROM cocta "
    sSQL = sSQL & "WHERE tpocta='" & s_Estado_Act & "'"
    sSQL = sSQL & "AND estcta='A'"
   Case "ctp"           ' Conceptos por clase de planilla por tipo
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc "
    sSQL = sSQL & "FROM plconcepto cpc "
    sSQL = sSQL & "INNER JOIN plconceplanilla cxp ON cpc.codcpc=cxp.codcpc AND cxp.codcls='" & Left(s_Condicion, 2) & "' "
    sSQL = sSQL & "WHERE cpc.tipocpc IN('" & Mid(s_Condicion, 3) & "') "
    sSQL = sSQL & "AND cpc.estadocpc='" & s_Estado_Act & "'"
   Case "ctr"           ' Convenios Tributarios
    sSQL = "SELECT codctr, desctr "
    sSQL = sSQL & "FROM plconven "
    sSQL = sSQL & "WHERE estadoctr='" & s_Estado_Act & "'"
   Case "cxc"           ' Conceptos por planilla por centro de costo
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc"
    sSQL = sSQL & " FROM (plconceplanilla cxp"
    sSQL = sSQL & " LEFT JOIN plconcepto cpc ON cxp.codcpc=cpc.codcpc)"
    sSQL = sSQL & " WHERE cxp.codcls='" & Left(s_Condicion, 2) & "'"
    sSQL = sSQL & " AND NOT EXISTS(SELECT * FROM plctacencos cxc"
    sSQL = sSQL & " WHERE cxc.codcco='" & Mid(s_Condicion, 3) & "'"
    sSQL = sSQL & " AND cxc.codcls=cxp.codcls AND cxc.codcpc=cxp.codcpc)"
    sSQL = sSQL & " AND cpc.estadocpc='" & s_Estado_Act & "'"
    sSQL = sSQL & " AND cxp.impbolecpc='" & s_Estado_Act & "'"
   Case "cxe"           ' Periodos de provisión de cts por estado
    sSQL = "SELECT pdocts, descricts "
    sSQL = sSQL & "FROM plctsperiodo "
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 2) & "' "
    sSQL = sSQL & " AND estadocts='" & Left(s_Condicion, 1) & "'"
   Case "cxp"           ' Conceptos no existentes en las planilla
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc"
    sSQL = sSQL & " FROM plconcepto cpc"
    sSQL = sSQL & " WHERE cpc.estadocpc='" & s_Estado_Act & "'"
    sSQL = sSQL & " AND NOT EXISTS(SELECT * FROM plconceplanilla cxp"
    sSQL = sSQL & " WHERE cxp.codcls='" & s_Condicion & "' AND cxp.codcpc=cpc.codcpc)"
   Case "cxt"           ' Concepto planilla por tipo y clase
    sSQL = "SELECT cpc.codcpc, cpc.descpc, cpc.tipocpc"
    sSQL = sSQL & " FROM (plconceplanilla cxp"
    sSQL = sSQL & " LEFT JOIN plconcepto cpc ON cxp.codcpc=cpc.codcpc)"
    sSQL = sSQL & " WHERE cxp.codcls='" & Left(s_Condicion, 2) & "'"
    sSQL = sSQL & " AND cxp.clasecpc='" & Mid(s_Condicion, 3, 1) & "'"
    sSQL = sSQL & " AND cpc.tipocpc IN('" & Mid(s_Condicion, 4) & "')"
    sSQL = sSQL & " AND cpc.estadocpc='" & s_Estado_Act & "'"
   Case "dci"           ' Documento de identidad
    sSQL = "SELECT coddci, desdci, sigladci FROM pldocidentidad"
    sSQL = sSQL & " WHERE estadodci='" & s_Estado_Act & "'"
   Case "emp"           ' Empresas del sistema
    sSQL = "SELECT codemp, razemp, rucemp FROM tgemp"
    sSQL = sSQL & " WHERE estemp='A'"
   Case "epr"           ' Establecimiento Propio
    sSQL = "SELECT codepr, desepr FROM plestablecimientopropio"
   Case "eps"           ' Entidad EPS
    sSQL = "SELECT codeps, deseps FROM plentidadeps"
    sSQL = sSQL & " WHERE estadoeps='" & s_Estado_Act & "'"
   Case "eqX"           ' RUC de Empresas
    sSQL = "SELECT codeqd, deseqd FROM plempresasqdes "
   Case "eqd"           ' RUC de Empresas
    sSQL = "SELECT codepr AS codepr,'" & s_Condicion & "' rucepr, desepr, cdgepr, tasepr "
    sSQL = sSQL & "FROM plestablecimientopropio WHERE estadoepr=1 UNION ALL "
    sSQL = sSQL & "SELECT codeed AS codepr, ruceed AS rucepr,deseed AS desepr, esteed AS cdgepr,taseed AS tasepr"
    sSQL = sSQL & "FROM plempresaseqdes WHERE estadoeed='" & s_Estado_Act & "'"
    s_Campo = "1"
   Case "est"           ' Tipo de Establecimiento
    sSQL = "SELECT codest, desest FROM plestablecimiento"
    sSQL = sSQL & " WHERE estadoest='" & s_Estado_Act & "'"
   Case "ldn"           ' Larga distancia nacional
    sSQL = "SELECT codldn, desldn FROM plcodigoldn"
    sSQL = sSQL & " WHERE estadoldn='" & s_Estado_Act & "'"
   Case "mof"           ' Entidad EPS
    sSQL = "SELECT codmof, desmof FROM plmotfin"
    sSQL = sSQL & " WHERE estadomof='" & s_Estado_Act & "'"
   Case "mfo"           ' Entidad EPS
    sSQL = "SELECT codmfo, desmfo FROM plmodforma"
    sSQL = sSQL & " WHERE estadomfo='" & s_Estado_Act & "'"
   Case "niv"           ' Nivel Educativo
    sSQL = "SELECT codniv, desniv FROM plniveducativo"
    sSQL = sSQL & " WHERE estadoniv='" & s_Estado_Act & "'"
   Case "pdo"           ' Periodo de pago
    sSQL = "SELECT codpdo, despdo FROM plperiodo"
    sSQL = sSQL & " WHERE codcls='" & Left(s_Condicion, 2) & "'"
    sSQL = sSQL & " AND anopdo='" & Mid(s_Condicion, 3) & "'"
   Case "pmd"           ' Pais emisión de documento
    sSQL = "SELECT codpemi, despemi FROM plpaisemidocum"
    sSQL = sSQL & " WHERE estadopemi='" & s_Estado_Act & "'"
   Case "pfs"           ' Profesion de personal
    sSQL = "SELECT codpfs, despfs FROM plprofesion"
    sSQL = sSQL & " WHERE estadopfs='" & s_Estado_Act & "'"
   Case "pll"           ' Formato de planilla
    sSQL = "SELECT DISTINCTROW codpll, despll FROM plplanilla"
    sSQL = sSQL & " WHERE codcls='" & s_Condicion & "'"
   Case "pxe"           ' Periodo de pago por estado
    sSQL = "SELECT codpdo, despdo FROM plperiodo"
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 2, 2) & "'"
    sSQL = sSQL & " AND anopdo='" & Mid(s_Condicion, 4) & "'"
    sSQL = sSQL & " AND estadopdo='" & Left(s_Condicion, 1) & "'"
   Case "pad"           ' Periodo de pago 2 meses estado diferente
    sSQL = "SELECT codpdo, despdo FROM plperiodo"
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 2, 2) & "'"
    sSQL = sSQL & " AND CONCAT(anopdo, mespdo)>='" & (Mid(s_Condicion, 4) - 90) & "'"
    sSQL = sSQL & " AND CONCAT(anopdo, mespdo)<'" & Mid(s_Condicion, 4) & "'"
    sSQL = sSQL & " AND estadopdo<>'" & Left(s_Condicion, 1) & "'"
   Case "ped"           ' Periodo de pago estado diferente
    sSQL = "SELECT codpdo, despdo FROM plperiodo"
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 2, 2) & "'"
    sSQL = sSQL & " AND anopdo='" & Mid(s_Condicion, 4) & "'"
    sSQL = sSQL & " AND estadopdo<>'" & Left(s_Condicion, 1) & "'"
   Case "pet"           ' Periodo de pago estado diferente y tipo
    sSQL = "SELECT codpdo, despdo FROM plperiodo"
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 3, 2) & "'"
    sSQL = sSQL & " AND anopdo='" & Mid(s_Condicion, 5) & "'"
    sSQL = sSQL & " AND tpopdo='" & Mid(s_Condicion, 2, 1) & "'"
    sSQL = sSQL & " AND estadopdo<>'" & Left(s_Condicion, 1) & "'"
   Case "prd"           ' Periodicidad
    sSQL = "SELECT codprd, desprd FROM plperiodicidad "
    sSQL = sSQL & " WHERE estadoprd='" & s_Estado_Act & "'"
   Case "pro"           ' Procesos de calculo por estado
    sSQL = "SELECT codproce, desproce FROM plproceso "
    sSQL = sSQL & "WHERE codcls='" & Mid(s_Condicion, 2, 2) & "' "
    sSQL = sSQL & "AND estadoproce='" & Left(s_Condicion, 1) & "' "
   Case "pvt"           ' Ejercicio de provisión de vacaciones todos
    sSQL = "SELECT codpvs, descripvs "
    sSQL = sSQL & "FROM plpvsperiodovac "
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 2) & "' "
   Case "qmd"           ' Empresas quienes me Destacan
    sSQL = "SELECT codqmd, desqmd FROM plempresasqmdes "
    sSQL = sSQL & " WHERE estadoqmd='" & s_Estado_Act & "'"
   Case "rpt"           ' Formatos de reportes generados
    sSQL = "SELECT codrpt, desrpt FROM plgenreporte"
    sSQL = sSQL & " WHERE codcls='" & s_Condicion & "'"
   Case "sec"           ' Sección o Area de la empresa
    sSQL = "SELECT codsec, dessec FROM plseccion "
    sSQL = sSQL & "WHERE estadosec='" & s_Estado_Act & "'"
   Case "sed"           ' Sub Periodo de provisión de cts estado diferente
    sSQL = "SELECT subcts, descrisub "
    sSQL = sSQL & "FROM plctsperiodosub "
    sSQL = sSQL & "WHERE codcls='" & Mid(s_Condicion, 2, 2) & "' "
    sSQL = sSQL & "AND pdocts='" & Mid(s_Condicion, 4) & "' "
    sSQL = sSQL & "AND estadosub<>'" & Left(s_Condicion, 1) & "'"
   Case "sxe"           ' Sub Periodo de provisión de cts por estado
    sSQL = "SELECT subcts, descrisub "
    sSQL = sSQL & "FROM plctsperiodosub "
    sSQL = sSQL & "WHERE codcls='" & Left(s_Condicion, 2) & "' "
    sSQL = sSQL & "AND pdocts='" & Mid(s_Condicion, 3) & "' "
    sSQL = sSQL & "AND estadosub='" & s_Filtro & "'"
   Case "stp"           ' Situacion del Trabajador pensionista
    sSQL = "SELECT codstp, desstp FROM plsitrapen"
    sSQL = sSQL & " WHERE estadostp='" & s_Estado_Act & "'"
   Case "tbl"           ' Tablas de factores o valores generales
    sSQL = "SELECT codtbl, destbl FROM pltablabase "
    sSQL = sSQL & "WHERE codcls='" & Mid(s_Condicion, 5) & "' "
    sSQL = sSQL & "AND pdoano='" & Left(s_Condicion, 4) & "'"
   Case "tcc"           ' Todos los Centro de Costos
    sSQL = "SELECT ccb02cod, ccb02des FROM ccb02cc"
   Case "tdc"           ' Tipo de documento
    sSQL = "SELECT ccb02cod, ccb02des FROM ccb02tipd"
    s_Where = " WHERE "
   Case "tic"           ' Tipo de Comprobante
    sSQL = "SELECT codtic, destic FROM pltipcom"
    sSQL = sSQL & " WHERE estadotic='" & s_Estado_Act & "'"
   Case "tip"           ' Tipo de Pago
    sSQL = "SELECT codtip, destip FROM pltippago"
    sSQL = sSQL & " WHERE estadotip='" & s_Estado_Act & "'"
    Case "tco"           ' Tipo de contrato
    sSQL = "SELECT codtco, destco FROM pltipcontrato"
    sSQL = sSQL & " WHERE estadotco='" & s_Estado_Act & "'"
    Case "tsu"           ' Tipo de Suspension
    sSQL = "SELECT codtsu, destsu FROM pltipsusp"
    sSQL = sSQL & " WHERE estadotsu='" & s_Estado_Act & "'"
   Case "tpt"           ' Tipo de trabajador
    sSQL = "SELECT codtpt, destpt FROM pltpotrabajador"
    sSQL = sSQL & " WHERE estadotpt='" & s_Estado_Act & "'"
   Case "ubi"           ' Ubicación o Localidad
    sSQL = "SELECT codubica, desubica FROM plubicacion "
    sSQL = sSQL & "WHERE estadoubica='" & s_Estado_Act & "'"
   Case "ubg"           ' Ubicacion geografica
    sSQL = "SELECT codubg, desubg, postalubg FROM tgubigeo"
    sSQL = sSQL & " WHERE nivelubg=" & Left(s_Condicion, 1)
    sSQL = sSQL & " AND LEFT(codubg," & Len(Mid(s_Condicion, 2)) & ")='" & Mid(s_Condicion, 2) & "'"
   Case "usr"           ' Usuarios del sistema
    sSQL = "SELECT codusr, nomusr FROM sgusr"
    sSQL = sSQL & " WHERE estusr='A'"
   Case "via"           ' Documento de identidad
    sSQL = "SELECT codvia, desvia, abrevia FROM pltipovia"
    sSQL = sSQL & " WHERE estadovia='" & s_Estado_Act & "'"
   Case "vxe"           ' Ejercicio de provisión de vacacion por estado
    sSQL = "SELECT codpvs, descripvs "
    sSQL = sSQL & "FROM plpvsperiodovac "
    sSQL = sSQL & " WHERE codcls='" & Mid(s_Condicion, 2) & "' "
    sSQL = sSQL & " AND estadopvs='" & Left(s_Condicion, 1) & "'"
   Case "vxf"           ' Variables y funciones generales
    sSQL = "SELECT codigo, nombre, descripcion, valor FROM plvarfunc"
    sSQL = sSQL & " WHERE tipo='" & Mid(s_Condicion, 2) & "'"
    If Left(s_Condicion, 1) = "N" Then
      sSQL = sSQL & " AND IFNULL(valor, '')<>''"
    End If
   Case "zon"           ' Documento de identidad
    sSQL = "SELECT codzona, deszona, abrezona FROM pltipozona"
    sSQL = sSQL & " WHERE estadozona='" & s_Estado_Act & "'"
   Case "nac"           ' Nacionalidad
    sSQL = "SELECT codnac, desnac FROM plnacionalidad"
    sSQL = sSQL & " WHERE estadonac='" & s_Estado_Act & "'"
  End Select
  If s_Filtro <> "" Then
    sSQL = sSQL & s_Where & "LEFT(" & s_Campo & "," & n_Longitud & ")='" & s_Filtro & "'"
  End If
  sSQL = sSQL & " ORDER BY " & s_Campo
  HelpTablas = sSQL

End Function
Function NombreDiaSemana(ByVal n_DiaProceso As Integer) As String
  NombreDiaSemana = Choose(n_DiaProceso, "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo")
End Function
Function NombreMes(ByVal n_MesProceso As Integer) As String
  NombreMes = Choose(n_MesProceso + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre 1", "Cierre 2", "Cierre 3")
End Function
Function NumeroDias360(ByVal d_FechaProceso As Date, ByVal d_FechaInicio As Date, d_FechaFin As Date) As Long
  Dim nDiaDesde As Integer, nMesDesde As Integer, nAnyoDesde As Integer
  Dim nDiaHasta As Integer, nMesHasta As Integer, nAnyoHasta As Integer
  Dim nAnyosAjuste As Integer, nDiasAjuste As Integer, nMesesAjuste As Integer
  Dim nDiasMes As Integer
  
  NumeroDias360 = 0
  d_FechaProceso = IIf(d_FechaProceso > d_FechaFin, d_FechaProceso, d_FechaFin)
  ' Fecha de proceso es menor i/o igual que la fecha final
  If (d_FechaFin >= d_FechaInicio And d_FechaFin >= d_FechaProceso) Then
    ' Primero : obtengo dias truncos e inicializo fecha inicial
    nDiaDesde = DatePart("d", d_FechaInicio)
    nMesDesde = DatePart("m", d_FechaInicio)
    nAnyoDesde = DatePart("yyyy", d_FechaInicio)
    ' Segundo : obtengo dias truncos e inicializo fecha final o proceso
    nDiaHasta = DatePart("d", d_FechaProceso)
    nMesHasta = DatePart("m", d_FechaProceso)
    nAnyoHasta = DatePart("yyyy", d_FechaProceso)
    nAnyoHasta = nAnyoHasta - IIf(nMesHasta = 12, 0, 1)
    ' Tercero : obtengo ajuste de meses y dias
    nDiasMes = NumeroDiasMes(nMesHasta, nAnyoHasta)
    nDiasAjuste = IIf((nDiaHasta = nDiasMes), 0, nDiaHasta) - nDiaDesde + 1
    nMesesAjuste = IIf((nMesHasta = 12), 0, nMesHasta) - nMesDesde + IIf((nDiaHasta = nDiasMes), 1, 0)
    nAnyosAjuste = (nAnyoHasta - nAnyoDesde) + 1
    ' Ultimo : retorno los dias en proceso
    NumeroDias360 = (nAnyosAjuste * 360) + (nMesesAjuste * 30) + nDiasAjuste
  End If
  
End Function

Function NumeroDias360Nuevo(ByVal d_FechaProceso As Date, ByVal d_FechaInicio As Date, d_FechaFin As Date) As Long
  Dim dias As Integer
  If Month(d_FechaInicio) <> 2 Then
    If Day(d_FechaInicio) = 31 Or Day(d_FechaInicio) = 30 Then
      dias = 1
    Else
      dias = (30 - Day(d_FechaInicio)) + 1
    End If
  Else
    If Day(d_FechaInicio) = 29 Or Day(d_FechaInicio) = 28 Then
      dias = 1
    Else
      dias = (30 - Day(d_FechaInicio)) + 1
    End If
  End If
  NumeroDias360Nuevo = DateDiff("m", d_FechaInicio, d_FechaProceso) * 30 + dias
End Function

Function NumeroDiasMes(ByVal n_MesProceso As Integer, ByVal n_AnoProceso As Integer) As Integer
  Dim d_NewFecha As Date
  Select Case n_MesProceso
   Case 0, 13, 14, 15
      NumeroDiasMes = IIf(n_MesProceso = 0, 1, 31)
   Case Else
      d_NewFecha = DateAdd("m", 1, CDate("01/" & n_MesProceso & "/" & n_AnoProceso))
      NumeroDiasMes = DatePart("d", DateAdd("d", -1, d_NewFecha))
  End Select
End Function

Function NumeroEnLetras(ByVal n_Numero As Double, Optional l_CentimoFinal As Boolean = True, Optional l_CeroFinal As Boolean = True) As String
  ' Devuelve un número expresado en letras.
  ' El parámetro l_CeroFinal se utiliza en la recursión para saber si se debe colocar
  ' la "O" final cuando la palabra es UN(O)
  ' El parámetro l_CentimoFinal si al final se usa centimos/100
  Dim nCentavos As Double
  Dim lngContDec As Long, lngContCent As Long
  Dim lngContMil As Long, lngContMillon As Long
  Dim aNumero As Variant, aDecenas As Variant, aCentenas As Variant
  Dim lNegativo As Boolean, lPlural As Boolean
  Dim sNumeroLetras As String
    
  If Int(n_Numero) = 0# Then
    sNumeroLetras = "CERO"
  End If
    
  aNumero = Array(vbNullString, "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE")
  aDecenas = Array(vbNullString, vbNullString, "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA", "CIEN")
  aCentenas = Array(vbNullString, "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")

  If n_Numero < 0# Then
    lNegativo = True
    n_Numero = Abs(n_Numero)
  End If

  If Int(n_Numero) <> n_Numero Then
    nCentavos = Abs(n_Numero - Int(n_Numero))
    n_Numero = Int(n_Numero)
  End If

  Do While n_Numero >= 1000000#
    lngContMillon = lngContMillon + 1
    n_Numero = n_Numero - 1000000#
  Loop

  Do While n_Numero >= 1000#
    lngContMil = lngContMil + 1
    n_Numero = n_Numero - 1000#
  Loop
    
  Do While n_Numero >= 100#
    lngContCent = lngContCent + 1
    n_Numero = n_Numero - 100#
  Loop
    
  If Not (n_Numero > 10# And n_Numero <= 20#) Then
    Do While n_Numero >= 10#
      lngContDec = lngContDec + 1
      n_Numero = n_Numero - 10#
    Loop
  End If
    
  If lngContMillon > 0 Then
    If lngContMillon >= 1 Then   'si el número es >1000000 usa recursividad
      sNumeroLetras = NumeroEnLetras(lngContMillon, l_CentimoFinal, False)
      If Not lPlural Then lPlural = (lngContMillon > 1)
      lngContMillon = 0
    End If
    sNumeroLetras = Trim(sNumeroLetras) & aNumero(lngContMillon) & " MILLON" & IIf(lPlural, "ES ", " ")
  End If
    
  If lngContMil > 0 Then
    If lngContMil >= 1 Then   'si el número es >100000 usa recursividad
      sNumeroLetras = sNumeroLetras & NumeroEnLetras(lngContMil, l_CentimoFinal, False)
      lngContMil = 0
    End If
    sNumeroLetras = Trim(sNumeroLetras) & aNumero(lngContMil) & " MIL "
  End If
    
  If lngContCent > 0 Then
    If lngContCent = 1 And lngContDec = 0 And n_Numero = 0# Then
      sNumeroLetras = sNumeroLetras & "CIEN"
    Else
      sNumeroLetras = sNumeroLetras & aCentenas(lngContCent) & " "
    End If
  End If
    
  If lngContDec >= 1 Then
    If lngContDec = 1 Then
      sNumeroLetras = sNumeroLetras & aNumero(10)
    Else
      sNumeroLetras = sNumeroLetras & aDecenas(lngContDec)
    End If
    
    If lngContDec >= 3 And n_Numero > 0# Then
      sNumeroLetras = sNumeroLetras & " Y "
    End If
  Else
    If n_Numero >= 0# And n_Numero <= 20# Then
      sNumeroLetras = sNumeroLetras & aNumero(n_Numero)
      If n_Numero = 1# And l_CeroFinal Then
        sNumeroLetras = sNumeroLetras & "O"
      End If
      If nCentavos > 0# Or l_CeroFinal Then
        sNumeroLetras = Trim(sNumeroLetras) & IIf(l_CentimoFinal, " CON " & Format$(CInt(nCentavos * 100#), "00") & "/100", "")
      End If
      NumeroEnLetras = sNumeroLetras
      Exit Function
    End If
  End If
    
  If n_Numero > 0# Then
    sNumeroLetras = sNumeroLetras & aNumero(n_Numero)
    If n_Numero = 1# And l_CeroFinal Then
      sNumeroLetras = sNumeroLetras & "O"
    End If
  End If
  sNumeroLetras = sNumeroLetras & IIf(l_CentimoFinal, " CON " + Format$(CInt(nCentavos * 100#), "00") & "/100", "")
  NumeroEnLetras = IIf(lNegativo, "(" & sNumeroLetras & ")", sNumeroLetras)

End Function
Function PadC(ByVal Expresion, ByVal n_Longitud As Integer, ByVal s_Caracter As String) As String

Expresion = aTexto(Expresion)
PadC = String$((n_Longitud - Len(Expresion)) / 2, s_Caracter) + Expresion + String$((n_Longitud - Len(Expresion)) / 2, s_Caracter)

End Function
Function PadL(ByVal Expresion, ByVal n_Longitud As Integer, ByVal s_Caracter As String) As String

Expresion = aTexto(Expresion)
PadL = String$(n_Longitud - Len(Expresion), s_Caracter) & Expresion
    
End Function
Function PadR(ByVal Expresion, ByVal n_Longitud As Integer, ByVal s_Caracter As String) As String
Static s_Cadena As String

Expresion = aTexto(Expresion)
If n_Longitud > Len(Expresion) Then
   s_Cadena = Expresion & String$(n_Longitud - Len(Expresion), s_Caracter)
Else
   s_Cadena = Expresion
End If
PadR = Left$(s_Cadena, n_Longitud)
    
End Function
Function PathApp(ByVal s_Expresion As String) As String
  PathApp = s_Expresion & IIf(Right$(s_Expresion, 1) <> "\", "\", "")
End Function
Function RangoImpresion(ByVal s_Conexion As String, ByVal s_Proceso As String, ByVal s_Valor As String, ByVal s_Usuario As String, s_FechaHora As String, ByVal s_Accion As String) As Boolean
Dim sSQL As String

If s_Accion = "A" Then
  sSQL = "INSERT INTO rangoimpresion (proceso, valor, usrcre, fyhcre) "
  sSQL = sSQL & "VALUES("
  sSQL = sSQL & "'" & s_Proceso & "', "
  sSQL = sSQL & "'" & s_Valor & "', "
  sSQL = sSQL & "'" & s_Usuario & "', "
  sSQL = sSQL & "'" & s_FechaHora & "')"
Else
  sSQL = "DELETE FROM rangoimpresion "
  sSQL = sSQL & "WHERE proceso='" & s_Proceso & "' "
  sSQL = sSQL & "AND usrcre='" & s_Usuario & "' "
  sSQL = sSQL & "AND fyhcre='" & s_FechaHora & "'"
End If
RangoImpresion = Execution(s_Conexion, sSQL)
    
End Function
Function Redondea(ByVal n_Numero As Variant, ByVal n_Decimal As Integer) As Double
Static s_Numero As String, s_Picture As String
Static n_Entero As Integer

s_Numero = aTexto(n_Numero): n_Numero = Val(s_Numero)
n_Entero = InStr(s_Numero, ".") - 1
If InStr(s_Numero, ".") = 0 Or n_Decimal < 0 Then
    Redondea = Val(n_Numero)
Else
    If n_Entero > 0 Then
        s_Picture = String$(n_Entero, "#")
        If n_Decimal > 0 Then s_Picture = s_Picture & "." & String$(n_Decimal, "#")
        Redondea = Val(Format$(n_Numero, s_Picture))
    Else
        If n_Decimal = 0 Then
            s_Picture = String$(Len(s_Numero), "#")
            Redondea = Val(Format$(n_Numero, s_Picture))
        Else
            s_Picture = String$(Len(s_Numero), "#") & "." & String$(n_Decimal, "#")
            Redondea = Val(Format$(n_Numero, s_Picture))
        End If
    End If
End If

End Function
Function Redondeo(ByVal n_Numero As Double, ByVal n_Decimales As Integer) As Double
  Redondeo = Int(n_Numero * 10 ^ n_Decimales + 1 / 2) / 10 ^ n_Decimales
End Function
Function SacaEntRetApos(ByVal s_Expresion As String) As String

s_Expresion = aTexto(s_Expresion)
If s_Expresion <> "" Then
    ' saco los Enters de la cadena de caracteres
    While InStr(s_Expresion, Chr(13)) <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(13)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(13)) + 1)
    Wend
    ' saco los Retornos de la cadena de caracteres
    While InStr(s_Expresion, Chr(10)) <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(10)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(10)) + 1)
    Wend
    ' saco los Apostrofes de la cadena de caracteres
    While InStr(s_Expresion, "'") <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "'") - 1) & "´" & Mid$(s_Expresion, InStr(s_Expresion, "'") + 1)
    Wend
    ' saco los Rayas de la cadena de caracteres
    While InStr(s_Expresion, "|") <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "|") - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, "|") + 1)
    Wend
End If
SacaEntRetApos = Trim$(s_Expresion)

End Function
Function Tipocambio(ByVal s_Conexion As String, ByVal d_Fecha As String, ByVal s_Tipo As String) As Double

  Tipocambio = 0
  d_Fecha = Format(d_Fecha, "yyyy-mm-dd")
  sSQL = "SELECT IFNULL(imptcb_" & IIf(s_Tipo = "V", "vta", "cpr") & ", 0) AS nFactor "
  sSQL = sSQL & "FROM tgtcb "
  sSQL = sSQL & "WHERE fehtcb= DATE_FORMAT('" & d_Fecha & "', '%Y-%m-%d')"
  ' Aperturo el Recordset
  rs_Consulta.Open sSQL, s_Conexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not (rs_Consulta.EOF Or rs_Consulta.BOF) Then
     Tipocambio = CDec(rs_Consulta!nFactor)
  End If
  rs_Consulta.Close: Set rs_Consulta = Nothing

End Function
Function ValidaEstructura(ByVal s_Conexion As String, ByVal s_Empresa As String, ByVal s_Proceso As String, ByVal s_Expresion As String, ByVal s_Tabla As String, ByVal s_Campo As String, Optional ByVal s_Where As String) As String
  Static l_ExistRecord As Boolean
  Static n_Final As Integer, s_Nivel As String
  ValidaEstructura = "E"
  
  sSQL = "SELECT ecp_nivel, ecp_posinicio, ecp_posfinal, ecp_movimiento "
  sSQL = sSQL & "FROM estructura "
  sSQL = sSQL & "WHERE ecp_empresa = '" & s_Empresa & "' "
  sSQL = sSQL & "AND ecp_clase = '" & s_Proceso & "' "
  sSQL = sSQL & "AND ecp_posfinal = " & Len(s_Expresion) & " "
  
  ' Aperturo el Recordset
  rs_Consulta.Open sSQL, s_Conexion, adOpenForwardOnly, adLockReadOnly, adCmdText
  l_ExistRecord = (rs_Consulta.EOF Or rs_Consulta.BOF)
  If Not l_ExistRecord Then
      ValidaEstructura = aTexto(rs_Consulta!ecp_movimiento)
      s_Nivel = rs_Consulta!ecp_nivel
      If s_Nivel <> "1" Then
          rs_Consulta.Close
          n_Final% = DameDescripcion(s_Conexion, s_Empresa, s_Proceso & Trim$(s_Nivel - 1), "EF")
          sSQL = "SELECT COUNT(*) "
          sSQL = sSQL & "FROM " & s_Tabla & " "
          sSQL = sSQL & "WHERE " & Left(s_Campo, 4) & "empresa = '" & s_Empresa & "' "
          sSQL = sSQL & "AND " & s_Campo & " = '" & Mid$(s_Expresion, 1, n_Final%) & "'"
          sSQL = sSQL & IIf(s_Where <> "", " AND " & s_Where, "")
          ' Aperturo el Recordset
          rs_Consulta.Open sSQL, s_Conexion, adOpenForwardOnly, adLockReadOnly, adCmdText
          l_ExistRecord = (rs_Consulta.EOF Or rs_Consulta.BOF)
          If Not l_ExistRecord And rs_Consulta(0) = 0 Then
              ValidaEstructura = "E"
              MsgBox "Primero ingrese nivel anterior de la estructura del código; Verifique", vbCritical
          End If
      End If
  Else
     MsgBox "Longitud no corresponde a nivel de la estructura del código; Verifique", vbCritical
  End If
  rs_Consulta.Close: Set rs_Consulta = Nothing

End Function
Function ValidaFecha(ByVal d_Fecha As String, ByVal n_Ano As Integer) As Boolean

  ValidaFecha = True
  If IsDate(d_Fecha) Then
    If Val(Right$(d_Fecha, 4)) < n_Ano Then
      Beep
      MsgBox "No puede Ingresar Fecha Menor a " & Trim$(n_Ano), vbExclamation
      ValidaFecha = False
    Else
      If Val(Mid$(d_Fecha, 4, 2)) > 12 Then
        Beep
        MsgBox "Fecha Errónea", vbExclamation
        ValidaFecha = False
      End If
    End If
  Else
    Beep
    MsgBox "Fecha Errónea", vbExclamation
    ValidaFecha = False
  End If

End Function
Function ValidaRuc(s_NumeroRuc As String) As Boolean
  Static n_Acumulador%, i%, n_Residuo%
  
  ValidaRuc = False: n_Acumulador% = 0
  For i% = 1 To 10
      n_Acumulador% = n_Acumulador% + (Val(Mid$(s_NumeroRuc, i%, 1)) * Choose(i%, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2))
  Next i%
  n_Residuo% = n_Acumulador% Mod 11
  If (11 - n_Residuo% = Val(Right$(s_NumeroRuc, 1))) Or (11 - n_Residuo% = 10 And Val(Right$(s_NumeroRuc, 1)) = 0) Then
     ValidaRuc = True
  End If
End Function
