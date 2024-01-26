Attribute VB_Name = "fTraduDB"
Option Explicit
'ini 2015-08-27/09-02 ctr obligac sunat
Public Function fPunto() As String
'ps_Plataforma
'    sSql = sSql & "SELECT * FROM information_schema.tables WHERE TABLE_SCHEMA='" & xNomBDS & "' AND TABLE_TYPE='BASE TABLE' AND TABLE_NAME='"

    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = s_Expresion & "."
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = s_Expresion & ".."
    End If
    fPunto = s_Expresion
End Function

'fin 2015-08-27/09-02 ctr obligac sunat

'ini 2015-08-04 verificar estructura de la base de datos
'conversion de un dato entero int(6)=mysql int=sql
Public Function finformation_schema_tables(xNomBDS As String, pTABLE_NAME As String) As String
'ps_Plataforma
'    sSql = sSql & "SELECT * FROM information_schema.tables WHERE TABLE_SCHEMA='" & xNomBDS & "' AND TABLE_TYPE='BASE TABLE' AND TABLE_NAME='"

    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = s_Expresion & "SELECT * FROM information_schema.tables WHERE TABLE_SCHEMA='" & xNomBDS & "' AND TABLE_TYPE='BASE TABLE' AND TABLE_NAME='" & pTABLE_NAME & "'"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = s_Expresion & "SELECT * FROM information_schema.tables WHERE TABLE_CATALOG='" & xNomBDS & "' AND TABLE_TYPE='BASE TABLE' AND TABLE_NAME='" & pTABLE_NAME & "'"
    End If
    finformation_schema_tables = s_Expresion
End Function

Public Function finformation_schema_COLUMNS2(xNomBDS As String, pTABLE_NAME As String) As String
'ps_Plataforma
'    sSql2 = sSql2 & "SELECT * FROM information_schema.COLUMNS WHERE TABLE_SCHEMA='siscfg' AND TABLE_NAME='" '& pTablaTmp & "'"

    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = s_Expresion & "SELECT * FROM information_schema.COLUMNS WHERE TABLE_SCHEMA='" & xNomBDS & "' AND TABLE_NAME='" & pTABLE_NAME & "'"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = s_Expresion & "SELECT * FROM information_schema.COLUMNS WHERE TABLE_CATALOG='" & xNomBDS & "' AND TABLE_NAME='" & pTABLE_NAME & "'"
    End If
    finformation_schema_COLUMNS2 = s_Expresion
End Function

'recordset para buscar informacion
Public Function fRstOpenBuscar(pCnn As ADODB.Connection, pRst As ADODB.Recordset, _
pSource As String) As ADODB.Recordset
  On Error GoTo ErrorRs
    Set pRst = New ADODB.Recordset
    'pRst = fRstOpen(pCnn, pRst, adOpenDynamic, adLockReadOnly)
    Set fRstOpenBuscar = fRstOpen(pCnn, pRst, pSource, adOpenDynamic, adLockReadOnly)
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Function

'Public Function fRstOpen(pCnn As ADODB.Connection, pRst As ADODB.Recordset, pSource As String, _
'Optional pCursorType As Integer, Optional pLockType As Integer) As ADODB.Recordset
'  On Error GoTo ErrorRs
'
'    'Set pRst = New ADODB.Recordset
'    With pRst
'        .ActiveConnection = pCnn
'        If .State = adStateOpen Then .Close
'        .Source = pSource
'        .CursorType = pCursorType 'adOpenDynamic
'        .LockType = pLockType 'adLockReadOnly
'        .Open
'    End With
'    Set fRstOpen = pRst
'  Exit Function
'ErrorRs:
'  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
'End Function

'fin 2015-08-04 verificar estructura de la base de datos

'ini 2015-05-18 validacion frm
'++++++++++++++++++++++++++++++++++++++++
Public Function fRstOpen(pCnn As ADODB.Connection, pRst As ADODB.Recordset, pSource As String, _
Optional pCursorType As Integer, Optional pLockType As Integer) As ADODB.Recordset
  On Error GoTo ErrorRs
  
    'Set pRst = New ADODB.Recordset
    With pRst
        .ActiveConnection = pCnn
        If .State = adStateOpen Then .Close
        .Source = pSource
        .CursorType = pCursorType 'adOpenDynamic
        .LockType = pLockType 'adLockReadOnly
        .Open
    End With
    Set fRstOpen = pRst
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Function
Public Function fCnnOpen(pCnn As ADODB.Connection, Optional pTimeout As Integer) As ADODB.Connection
  On Error GoTo ErrorRs
   Set pCnn = New ADODB.Connection
   With pCnn
        If pTimeout > 0 Then
            '2014-09-11 error time out
            .CommandTimeout = pTimeout 'segundos de espera
        End If
        .CursorLocation = adUseClient
        .ConnectionString = CONNSTRG & gsNomBDS
        .Open
    End With

    Set fCnnOpen = pCnn
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Function

Public Sub fRstClose(pRst As ADODB.Recordset)
  On Error GoTo ErrorRs
  If Not pRst Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
    'If prst.State = adStateOpen Then prst.Close: Set prst = Nothing
    If pRst.State = adStateOpen Then pRst.Close
    Set pRst = Nothing
  End If
  Exit Sub
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub
Public Sub fCnnClose(pCnn As ADODB.Connection)
  On Error GoTo ErrorCnn
  If Not pCnn Is Nothing Then 'Is Nothing (no existe objeto) / if si xiste objeto aplica
    If pCnn.State = adStateOpen Then pCnn.Close
    Set pCnn = Nothing
  End If
  Exit Sub
    
ErrorCnn:

'''  pCnn.RollbackTrans              'RESTAURA TRANSACCION.
'''  If pCnn.State = adStateOpen Then pCnn.Close
'''  Set pCnn = Nothing
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
    
End Sub
'++++++++++++++++++++++++++++++++++++++++
'fin 2015-05-18 validacion frm
Public Function fMid(Optional fValor As String, Optional fStar As Integer, Optional fLen As Integer) As String
    Dim sSentencia As String
    sSentencia = ""
    If fStar = 0 Then
    fStar = 1
    End If
    If fLen = 0 Then
    fLen = Len(fValor)
    End If
    If Len(Trim(fValor)) = 0 Then
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "MID", "SUBSTRING") & "(" & fValor & "," & Str(fStar) & "," & Str(fLen)
    Else
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "MID", "SUBSTRING") & "(" & fValor & "," & Str(fStar) & "," & Str(fLen) & ")"
    End If
    fMid = sSentencia
End Function

Public Function fDistinct() As String
'ps_Plataforma
    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = s_Expresion + "DISTINCTROW"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = s_Expresion + "DISTINCT"
    End If
    fDistinct = s_Expresion
End Function
Public Function fConCat(ByRef arrVar1() As String) As String
'Public Function fConCat(arrVar1() As String) As String
'Public Function fConCat(ParamArray arrVar() As Variant) As String
    Dim s_Expresion As String
    Dim n As Integer
    n = 0
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "CONCAT("
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = ""
    End If
    For n = 0 To UBound(arrVar1)
        If ps_Plataforma = pSrvMySql Then
        s_Expresion = s_Expresion & arrVar1(n) & IIf(n = UBound(arrVar1), ")", ",")
        End If
        If ps_Plataforma = pSrvSql Then
        s_Expresion = s_Expresion & arrVar1(n) & IIf(n = UBound(arrVar1), " ", "+")
        End If
    Next
    fConCat = s_Expresion
End Function

Public Function fDropTable(pTable As String, pTipo As Integer) As String
'pTipo: 0=Tabla Fisica 1=Tabla Temporal
    Dim s_Expresion As String
    Dim sSentencia As String
    If pTipo = 1 Then
    sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & Len(pTable) + 2 & ")='#" & pTable & "_') DROP TABLE #" & pTable & ""
    Else
    sSentencia = "IF EXISTS (select * from sysobjects where name='" & pTable & "' and xtype='U') DROP TABLE " & pTable
    End If
    s_Expresion = IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & pTable & "", sSentencia)
    
    fDropTable = s_Expresion
End Function
Public Function fDropTable2(pTable As String, pTipo As Integer) As String
'ini 2014-12-15 error lectura tempdb..sysobjects bloquea
'se determino error en lectura del
'IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trpRngBceCpb') DROP TABLE #trpRngBceCpb
'por este motivo estamos cambiando a  OBJECT_ID
'IF OBJECT_ID ('tempdb..#trpRngBceCpb') IS NOT NULL DROP TABLE #trpRngBceCpb
'fin 2014-12-15 error lectura tempdb..sysobjects bloquea
'******************************
'pTipo: 0=Tabla Fisica 1=Tabla Temporal
    Dim s_Expresion As String
    Dim sSentencia As String
    If pTipo = 1 Then
'    sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & Len(pTable) + 2 & ")='#" & pTable & "_') DROP TABLE #" & pTable & ""
    sSentencia = "IF OBJECT_ID ('tempdb.." & pTable & "') IS NOT NULL DROP TABLE " & pTable & " "
    Else
    sSentencia = "IF EXISTS (select * from sysobjects where name='" & pTable & "' and xtype='U') DROP TABLE " & pTable
    End If
    s_Expresion = IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & pTable & "", sSentencia)
    
    fDropTable2 = s_Expresion
End Function

Public Function fCreateTable(pTable As String, pTipo As Integer) As String
'pTipo: 0=Tabla Fisica 1=Tabla Temporal
    'Dim s_Expresion As String
    Dim sSentencia As String
    
    
    If pTipo = 1 Then
         sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & pTable & "  ", "")
    Else
         sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS " & pTable & "  ", "")
    End If
    
    fCreateTable = sSentencia
End Function
Public Function fInto(pTable As String, pTipo As Integer) As String
'pTipo: 0=Tabla Fisica 1=Tabla Temporal
    'Dim s_Expresion As String
    Dim sSentencia As String
    sSentencia = ""
    
    If pTipo = 1 Then
         sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #" & pTable & "  ")
    Else
         sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO " & pTable & "  ")
    End If
    
    fInto = sSentencia
End Function

Public Function fCreateStruc(pTable As String, pSql As String) As String
'    Dim s_Expresion As String
    Dim sql As String
'    sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, " & Len(pTable) + 2 & ")='#" & pTable & "_') DROP TABLE #" & pTable & ""
'    s_Expresion = IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & pTable & "", sSentencia)
    
    
    If ps_Plataforma = pSrvMySql Then
        sql = "CREATE TABLE IF NOT EXISTS " & pTable & " "
        sql = sql & "( "
        sql = sql & pSql
        sql = sql & ") "
    End If
    'xtype='U' / U = Tabla de usuario
    'http://msdn.microsoft.com/es-es/library/ms177596.aspx
    If ps_Plataforma = pSrvSql Then
        sql = "if not exists (select * from sysobjects where name='" & pTable & "' and xtype='U')"
        sql = sql & "    create table " & pTable & " ("
        'Name varchar(64) not null
        sql = sql & pSql
        sql = sql & "    ) "
        'sql = sql & " go " 'sale error
    End If
    fCreateStruc = sql
  
End Function

Public Function fLen(fValor As String) As String
    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "LENGTH(" + fValor + ")"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "LEN(" + fValor + ")"
    End If
    
    fLen = s_Expresion
End Function
Public Function fIsNull(Optional fValor As String) As String
    Dim sSentencia As String
    sSentencia = ""
    If Len(Trim(fValor)) = 0 Then
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & fValor
    Else
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & fValor & ")"
    End If
    fIsNull = sSentencia
End Function
Public Function fLTrim(Optional fValor As String) As String
    Dim sSentencia As String
    sSentencia = ""
    If Len(Trim(fValor)) = 0 Then
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "TRIM", "LTRIM") & "(" & fValor
    Else
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "TRIM", "LTRIM") & "(" & fValor & ")"
    End If
    fLTrim = sSentencia
End Function

Public Function fFieldName(fValor As String) As String
    Dim sSentencia As String
    sSentencia = ""
    'sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(" & fValor & ")"
    sSentencia = sSentencia & Choose(gsIdioma, fValor, fValor & "x")
    fFieldName = sSentencia
End Function

Public Function fConvert103(fValor As String) As String
    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "DATE_FORMAT(" + fValor + ",'%Y%m%d')"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "CONVERT(smalldatetime," + fValor + ", 103)"
    End If
    
    fConvert103 = s_Expresion
'prueba compatibilidad 13/08/2012
' SELECT
'CONVERT(datetime, fehtcb, 120) f1,
'CONVERT(datetime, fehtcb, 103) f2,
'CONVERT(char(10), fehtcb, 120) f3,
'CONVERT(char(10), fehtcb, 103) f4,
'RIGHT(CONVERT(char(10), fehtcb, 103),4) f5,
'LEFT(CONVERT(char(10), fehtcb, 120),4) f6
'--RIGHT(CONVERT(char(10), fehtcb, 103),4) f7
'--*
'From tgtcb

'SELECT
'DATE_FORMAT(a.fehtcb, '%Y')
'#*
'FROM tgtcb a
End Function
Public Function fConvert103yyyy(fValor As String) As String
    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "DATE_FORMAT(" + fValor + ",'%Y')"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "RIGHT(CONVERT(char(10), fehtcb, 103),4)"
    End If
    
    fConvert103yyyy = s_Expresion
End Function
Public Function fConvert103ddmmyyySay(fValor As String) As String
    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "DATE_FORMAT(" + fValor + ",'%d/%m/%Y')"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "CONVERT(char(10)," + fValor + ", 103)"
    End If
    
    fConvert103ddmmyyySay = s_Expresion
'Sql
'****
'SELECT
'CONVERT(char(10), fehtcb, 103) f4
'From tgtcb
'
'Mysql
'******
'SELECT
'DATE_FORMAT(FehTCb,'%d/%m/%Y')
'From tgtcb
End Function

Public Function fDateFmt(fValor As Date) As String
' FehCpb   sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(dtpFehCpb.Value, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
    Dim s_Expresion As String
'sSentencia = sSentencia & _
'IIf(ps_Plataforma = pSrvMySql, _
'"DATE_FORMAT(", _
'"CONVERT(smalldatetime, ") _
'& "'" & Format(dtpFehCpb.Value, "yyyy-mm-dd") & _
'"', " & _
'IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "DATE_FORMAT(" & "'" & Format(fValor, "yyyy-mm-dd") & "'," & "'%Y-%m-%d'" & ")"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "CONVERT(smalldatetime," & "'" & Format(fValor, "yyyy-mm-dd") & "'," & " 120" & ")"
    End If
    
    If fValor = "01/01/1000" Then
    s_Expresion = "Null"
    End If

    fDateFmt = s_Expresion
End Function
Public Function fDateyyyymmdd(fValor As String) As String
  
    fDateyyyymmdd = fDateyyyymmdd2(fValor)
End Function
Public Function fDateyyyymmdd1(fValor As Date) As String
    Dim s_Expresion As String
    
    s_Expresion = "'" & Format(fValor, "yyyy-mm-dd") & "' "

    fDateyyyymmdd1 = s_Expresion
End Function

Public Function fDateyyyymmdd2(fValor As String) As String
    Dim s_Expresion As String
  
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "'" & Format(fValor, "yyyy-mm-dd") & "' "
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "'" & Format(fValor, "yyyy-mm-dd") & "'"
    End If
    
    If fValor = "01/01/1000" Then
    s_Expresion = "Null"
    End If

    fDateyyyymmdd2 = s_Expresion
End Function

Public Function fDateNow() As String
' FyHCre   sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
    Dim s_Expresion As String
' sSentencia = sSentencia & _
' IIf(ps_Plataforma = pSrvMySql, _
' "DATE_FORMAT(", _
' "CONVERT(datetime, ") _
' & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & _
' IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
    
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "DATE_FORMAT(" & "'" & Format(Now, "yyyy-mm-dd") & "'," & "'%Y-%m-%d'" & ")"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "CONVERT(smalldatetime," & "'" & Format(Now, "yyyy-mm-dd") & "'," & " 120" & ")"
    End If

    fDateNow = s_Expresion
End Function

'Public Function fVbDateFmt(Optional fValor As Date) As Date
Public Function fVbDateFmt() As Date
'            If IsNull(fValor) = True Then
            fVbDateFmt = Format("01/01/1000", "dd/mm/yyyy")
'            Else
'            fVbDateFmt = fValor
'            End If

End Function



Public Function fPrimaryKey(fValor As String, fFields As String) As String
    Dim s_Expresion As String
    If ps_Plataforma = pSrvMySql Then
    s_Expresion = "KEY " & fValor & "(" & fFields & ")"
    End If
    If ps_Plataforma = pSrvSql Then
    s_Expresion = "CONSTRAINT " & fValor & " PRIMARY KEY CLUSTERED "
    s_Expresion = s_Expresion & "(" & fFields
    s_Expresion = s_Expresion & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    s_Expresion = s_Expresion & ") ON [PRIMARY]"
    End If
    s_Expresion = ")"
    
    fPrimaryKey = s_Expresion
End Function

'ini 2015-07-27 Sdo. doc al mes pendiente
'Public Function fsdo_doc_pdte_hoy(pCnn As ADODB.Connection) As ADODB.Connection
Public Function fsdo_doc_pdte_hoy(pocnnTmp As ADODB.Connection) As ADODB.Connection
  On Error GoTo ErrorRs
 '  Set pCnn = New ADODB.Connection
 Dim sTabla As String
 Dim cCadReporte As String
'ini 2015-07-15 adicion pgo segun diario
    sTabla = "tmp_xls_pdte"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
'saldos de documento segun reporte historico
    cCadReporte = cCadReporte & "SELECT "
    cCadReporte = cCadReporte & "    a.pdoano AS cAno, a.MesPvs, a.CodCta, a.CodAux,a.CodTDc,"
    cCadReporte = cCadReporte & "    a.SerDoc, a.NroDoc, a.CodDro, a.NroCpb, Null AS codcco,"
    'cCadReporte = cCadReporte & "    Null AS detcco, CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocum, a.FehOpe, a.FeEDoc, a.FeVDoc,"
    cCadReporte = cCadReporte & "    Null AS detcco,"
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocum, "
    cCadReporte = cCadReporte & "a.FehOpe, a.FeEDoc, a.FeVDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, b.RazAux, "
    
    'cCadReporte = cCadReporte & "    a.RefDoc, a.GloIte AS GloIte, b.RazAux, (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
    cCadReporte = cCadReporte & "(CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) AS cDebeMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) AS cHaberMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END) AS cDebeME, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) AS cHaberME "
    
'    cCadReporte = cCadReporte & "    (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END) AS cDebeMN, (CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END) AS cHaberMN,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END) AS cDebeME, (CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END) AS cHaberME"
    cCadReporte = cCadReporte & "    ,a.TpoPvs"
    cCadReporte = cCadReporte & "    ,CONCAT(year(a.FehOpe),'-',LPAD(month(a.FehOpe),2,'0'),'-',LPAD(day(a.FehOpe),2,'0'),'-',a.CodDro,'-',a.NroCpb,'-',a.Nroite) AS x_clave "
    cCadReporte = cCadReporte & "FROM ((((COCpbDet a "
    cCadReporte = cCadReporte & "    LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    cCadReporte = cCadReporte & "    LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    cCadReporte = cCadReporte & "    LEFT JOIN Cocta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.Codcta=d.Codcta) "
    cCadReporte = cCadReporte & "    LEFT JOIN CoCCo e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.codcco=e.codcco) "
'    cCadReporte = cCadReporte & "WHERE a.codemp='010' "
'    cCadReporte = cCadReporte & "    AND a.pdoano='2014' "
    cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
    cCadReporte = cCadReporte & "    AND a.pdoano='" & gsAnoAct & "' "
    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta, 2)>='01' AND LEFT(a.codcta, 2)<='FF' "
    cCadReporte = cCadReporte & "    AND (a.ImpMN<> 0.00 OR a.ImpME<> 0.00) "
    'cCadReporte = cCadReporte & "    AND a.Mespvs <='03' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND a.Mespvs <='" & gsMesAct & "' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND IFNULL(a.NroDoc, '') <>'' AND d.inddoc='1' "
    'cCadReporte = cCadReporte & "    # AND a.CodAux='10097267265'"
    cCadReporte = cCadReporte & "    AND a.TpoPvs='" & TPOPVS_CAN & "' " 'TPOPVS_CAN
    'cCadReporte = cCadReporte & "    AND a.TpoPvs='C' " 'TPOPVS_CAN
    cCadReporte = cCadReporte & "ORDER BY a.codcta, a.codaux, a.codtdc, a.serdoc, a.NroDoc, a.TpoPvs, a.MesPvs, a.FehOpe "
    
    pocnnTmp.Execute cCadReporte

'*********************
    sTabla = "tmp_xls_pdte2"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
cCadReporte = cCadReporte & "SELECT "
cCadReporte = cCadReporte & "    codcta,codaux,cdocum,min(x_clave) x_clave "
cCadReporte = cCadReporte & "From tmp_xls_pdte "
'cCadReporte = cCadReporte & "GROUP BY codcta, codaux,cdocum # x_clave "
'cCadReporte = cCadReporte & "ORDER BY codcta, codaux,cdocum #x_clave "
cCadReporte = cCadReporte & "GROUP BY codcta, codaux,cdocum "
cCadReporte = cCadReporte & "ORDER BY codcta, codaux,cdocum "

    pocnnTmp.Execute cCadReporte

'*********************
    sTabla = "tmp_xls_pdte3"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
    cCadReporte = cCadReporte & "SELECT "
    cCadReporte = cCadReporte & "* "
    cCadReporte = cCadReporte & "From tmp_xls_pdte "
    cCadReporte = cCadReporte & "Where x_clave "
    cCadReporte = cCadReporte & "    IN (select x_clave from tmp_xls_pdte2) "
    pocnnTmp.Execute cCadReporte

    Set fsdo_doc_pdte_hoy = pocnnTmp
  Exit Function
ErrorRs:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Function
'fin 2015-07-27 Sdo. doc al mes pendiente


