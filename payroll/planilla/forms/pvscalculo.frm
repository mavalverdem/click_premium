VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fPvsCalculo 
   Caption         =   "Proceso de Cálculo"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   Icon            =   "pvscalculo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   6810
   Begin VB.CheckBox chkDiciembre 
      Alignment       =   1  'Right Justify
      Caption         =   "Diciembre"
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbPeriodo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2235
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   885
      Width           =   1860
   End
   Begin Threed.SSCommand cmdProceso 
      Height          =   315
      Left            =   4515
      TabIndex        =   6
      Top             =   1245
      Width           =   2025
      _Version        =   65536
      _ExtentX        =   3572
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "Procesar"
      Picture         =   "pvscalculo.frx":000C
   End
   Begin VB.ComboBox cmbEjercicio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   885
      Width           =   1740
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
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
      PromptChar      =   "_"
   End
   Begin Threed.SSFrame sfmProgreso 
      Height          =   480
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   1935
      Width           =   6480
      _Version        =   65536
      _ExtentX        =   11430
      _ExtentY        =   847
      _StockProps     =   14
      Caption         =   " Procesando personal : "
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      ShadowStyle     =   1
      Begin MSComctlLib.ProgressBar pgbProgreso 
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Top             =   225
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6810
      _Version        =   65536
      _ExtentX        =   12012
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   600
         TabIndex        =   8
         Top             =   120
         Width           =   5490
      End
   End
   Begin VB.Label lblDato 
      Caption         =   "Mes :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   2235
      TabIndex        =   10
      Top             =   615
      Width           =   1680
   End
   Begin VB.Label lblDato 
      Caption         =   "Fecha Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1290
      Width           =   1395
   End
   Begin VB.Label lblDato 
      Caption         =   "Ejercicio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   615
      Width           =   1680
   End
End
Attribute VB_Name = "fPvsCalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private o_Calculo As Object                             ' Objeto de clase de calculo
Private nContador As Integer                            ' Indice de rango
Private s_OptRegistro As String                         ' Instancia del formulario activo
Private s_Proceso As String                             ' Tipo de proceso
Private s_DsctoAusenciaBS As String                     ' Parametro descuento ausencia
'[
Private Function ppCalculoGratificacion(ByVal sFechaHora As String, ByVal s_Ejercicio As String, s_Mes As String, ByVal s_Fecha As String) As Boolean
  Dim nTotalRegistros As Long, nRegistro As Long, nNumeroDias As Long
  Dim sFechaIni As String, sFechaFin As String, sFechaPro As String
  Dim sGratixAsisistencia As String, sPersonal As String, sFechaIng As String
  Dim nRemunera_mn As Double, nRemuxDia_mn As Double, nRemuxAcu_mn As Double
  Dim nRemuxPro_mn As Double, nProviPen_mn As Double
  Dim nRemunera_me As Double, nRemuxDia_me As Double, nRemuxAcu_me As Double
  Dim nRemuxPro_me As Double, nProviPen_me As Double
  Dim s_Semestre As String, s_CuentaDebe_mn As String, s_CuentaHaber_mn As String
  Dim s_CuentaDebe_me As String, s_CuentaHaber_me As String
  Dim porstCalculo As ADODB.Recordset
  
  ' Obtengo si gratificación descuenta dias
  sGratixAsisistencia = s_Estado_Ina
  s_Sql = "SELECT cfg.pdoano, cfg.gratixasis "
  s_Sql = s_Sql & "FROM plcfgempresa cfg "
  s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "' "
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    sGratixAsisistencia = porstRecordset!gratixasis
  End If
  porstRecordset.Close
  
  ' Informacion de personal a procesar
  s_Sql = "SELECT DISTINCTROW psn.codpsn, "
  s_Sql = s_Sql & "CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
  s_Sql = s_Sql & "dxr.fecingreso, asi.fechacese, dxr.estadopsn, pvs.sempvs, pvs.mesini, pvs.mesfin, "
  s_Sql = s_Sql & "res.codcpc, res.codmon, res.importe_mn, res.importe_me, "
  s_Sql = s_Sql & "cta.codctagra_debmn, cta.codctagra_habmn, cta.codctagra_debme, cta.codctagra_habme, "
  s_Sql = s_Sql & "cta.codctagex_debmn, cta.codctagex_habmn, cta.codctagex_debme, cta.codctagex_habme "
  s_Sql = s_Sql & "FROM plpersonal psn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON psn.codcls=pdo.codcls "
  s_Sql = s_Sql & "INNER JOIN plresultado res ON pdo.codcls=res.codcls AND pdo.codpdo=res.codpdo AND psn.codpsn=res.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpvsperiodogra pvs ON res.codcls=pvs.codcls AND res.pdoano=pvs.pdoano AND res.codcpc=pvs.rembasbeneficio "
  s_Sql = s_Sql & "LEFT JOIN plctapvs cta ON dxr.codcls=cta.codcls AND dxr.codcco=cta.codcco "
  s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND psn.remintegralgrati='" & s_Estado_Ina & "' "
 ' s_Sql = s_Sql & "AND dxr.estadopsn<>'I' "
  s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & sFechaHora & "') "
  s_Sql = s_Sql & "AND CONCAT(pdo.anopdo, mespdo)='" & s_Ejercicio & s_Mes & "' "
  s_Sql = s_Sql & "AND pdo.tpopdo<>'L' "
  s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(dxr.fecingreso, '0000-00-00'), '%Y%m')<='" & s_Ejercicio & s_Mes & "' "
  s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(asi.fechacese, '9999-12-31'), '%Y%m')>'" & s_Ejercicio & s_Mes & "' "
  s_Sql = s_Sql & "AND (pvs.mesini<='" & s_Mes & "' AND pvs.mesfin>='" & s_Mes & "') "
  s_Sql = s_Sql & "AND pvs.estadopvs<>'" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY psn.codpsn"
  Set porstCalculo = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstCalculo.EOF And porstCalculo.BOF) And porstCalculo.RecordCount > 0 Then
    ' Obtengo los datos del proceso
    s_Semestre = Trim(porstCalculo!sempvs)
    sFechaFin = Format(gdl_Funcion.NumeroDiasMes(Trim(porstCalculo!mesfin), s_Ejercicio), "00")
    sFechaFin = Format(sFechaFin & "/" & porstCalculo!mesfin & "/" & s_Ejercicio, s_FormatoFecha)
    sFechaPro = IIf(Format(sFechaFin, "yyyymmdd") <= Format(s_Fecha, "yyyymmdd"), sFechaFin, s_Fecha)
    
    ' Instancio el objeto de calculo
    Set o_Calculo = CreateObject("syslink.calculo")
    o_Calculo.sCadenaConexion = ps_StrgConnec & ps_DataBase
    o_Calculo.sClasePlanilla = ps_ClsPlanilla
    o_Calculo.sDiaProceso = Left(sFechaPro, 2)
    o_Calculo.sMesProceso = Mid(sFechaPro, 4, 2)
    o_Calculo.sAnoProceso = Right(sFechaPro, 4)
    
    ' Creo los arreglos para la actualización
    a_Campos = Array("codcls", "pdoano", "sempvs", "codpsn", "pdomes", "fechaini", "fechafin", "numerodias", "codmon", "remunera_mn", "remunera_me", "imporpvsacu_mn", "imporpvsacu_me", "importepvs_mn", "importepvs_me", "codcta_debmn", "codcta_habmn", "codcta_debme", "codcta_habme", "fechacan", "estadogra", "usrcre", "fyhcre")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Fecha, TipoDato.Fecha, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Fecha, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    
    nTotalRegistros = CLng(porstCalculo.RecordCount)
    While Not porstCalculo.EOF
      'Cuentas contables de provisión
      s_CuentaDebe_mn = gdl_Funcion.aTexto(porstCalculo!codctag_debmn)
      s_CuentaHaber_mn = gdl_Funcion.aTexto(porstCalculo!codctag_habmn)
      s_CuentaDebe_me = gdl_Funcion.aTexto(porstCalculo!codctag_debme)
      s_CuentaHaber_me = gdl_Funcion.aTexto(porstCalculo!codctag_habme)
    
      sPersonal = Trim(porstCalculo!codpsn)
      
      'Actualizo el registro del personal
      
      sfmProgreso(0).Caption = " Procesando personal : " & sPersonal & " - " & Trim(porstCalculo!nombrepsn) & " "
      sfmProgreso(0).Refresh
      
      ' Obtengo los importes de provisones anteriores (pendientes)
      s_Sql = "SELECT DISTINCTROW MIN(fechaini)  AS fechainicio, "
      s_Sql = s_Sql & "ROUND(IFNULL(SUM(IFNULL(importepvs_mn, 0)), 0), 2) AS provision_mn, "
      s_Sql = s_Sql & "ROUND(IFNULL(SUM(IFNULL(importepvs_me, 0)), 0), 2) AS provision_me "
      s_Sql = s_Sql & "FROM plpvsgratifica "
      s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND codpsn='" & sPersonal & "' "
      s_Sql = s_Sql & "AND pdoano='" & s_Ejercicio & "' "
      s_Sql = s_Sql & "AND sempvs='" & Trim(porstCalculo!sempvs) & "' "
      s_Sql = s_Sql & "AND CONCAT(pdoano, pdomes)<'" & s_Ejercicio & s_Mes & "' "
      s_Sql = s_Sql & "AND estadogra<>'" & s_Estado_Blq & "'"
      Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstRecordset.EOF And porstRecordset.BOF) Then
        nProviPen_mn = CDec(porstRecordset!provision_mn)
        nProviPen_me = CDec(porstRecordset!provision_me)
      End If
      porstRecordset.Close
      
      ' Obtengo el numero de dias de provision(acumulados)
      sFechaIng = Format(porstCalculo!fecingreso, s_FormatoFecha)
      sFechaIni = Format("01/" & porstCalculo!mesini & "/" & s_Ejercicio, s_FormatoFecha)
      sFechaIni = IIf(Format(sFechaIni, "yyyymmdd") >= Format(sFechaIng, "yyyymmdd"), sFechaIni, sFechaIng)
      o_Calculo.sCodigoEmpleado = sPersonal
      o_Calculo.sFechaIngreso = sFechaIng
      o_Calculo.sFechaCese = Format(porstCalculo!fechacese, s_FormatoFecha)
      nNumeroDias = ppNumeroDias("N", sGratixAsisistencia)
      
      If chkDiciembre.Value = Checked Then
        nNumeroDias = ppNumeroDias("S", sGratixAsisistencia)
      End If
          
      ' Obtengo los importes de remuneración
      nRemunera_mn = CDec(porstCalculo!importe_mn)
      nRemuxDia_mn = nRemunera_mn / 180
      nRemuxAcu_mn = Round(nRemuxDia_mn * nNumeroDias, 2)
      nRemuxPro_mn = Round(nRemuxAcu_mn - nProviPen_mn, 2)
      nRemunera_me = CDec(porstCalculo!importe_me)
      nRemuxDia_me = nRemunera_me / 180
      nRemuxAcu_me = Round(nRemuxDia_me * nNumeroDias, 2)
      nRemuxPro_me = Round(nRemuxAcu_me - nProviPen_me, 2)
          
      ' Elimina la provision calculada
      s_Sql = "DELETE gra.* "
      s_Sql = s_Sql & "FROM plpvsgratifica gra "
      s_Sql = s_Sql & "WHERE gra.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND gra.pdoano='" & s_Ejercicio & "' "
      s_Sql = s_Sql & "AND gra.sempvs='" & Trim(porstCalculo!sempvs) & "' "
      s_Sql = s_Sql & "AND gra.codpsn='" & sPersonal & "' "
      s_Sql = s_Sql & "AND gra.pdomes='" & s_Mes & "'"
      If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
          
      ' Realizo el proceso de actualización de los registros
      a_Valores = Array(ps_ClsPlanilla, s_Ejercicio, Trim(porstCalculo!sempvs), sPersonal, s_Mes, Format(sFechaIni, s_FmtFechMysql_0), Format(sFechaPro, s_FmtFechMysql_0), CLng(nNumeroDias), porstCalculo!codmon, CDec(nRemunera_mn), CDec(nRemunera_me), CDec(nRemuxAcu_mn), CDec(nRemuxAcu_me), CDec(nRemuxPro_mn), CDec(nRemuxPro_me), s_CuentaDebe_mn, s_CuentaHaber_mn, s_CuentaDebe_me, s_CuentaHaber_me, Format("", s_FmtFechMysql_0), s_Estado_Act, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
      If Not Records_Ins("plpvsgratifica", a_Campos, a_Valores, a_Tipos) Then GoTo Error
      
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      pgbProgreso(0).Value = ((nRegistro \ nTotalRegistros) * 100)
      porstCalculo.MoveNext
    Wend
    ' Actualizo los periodos de gratificación
    s_Sql = "UPDATE plpvsperiodogra "
    s_Sql = s_Sql & "SET estadopvs='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdoano='" & s_Ejercicio & "' "
    s_Sql = s_Sql & "AND sempvs='" & s_Semestre & "'"
    If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
  End If
  ppCalculoGratificacion = True
  GoTo Finalizar

Error:
  ppCalculoGratificacion = False
Finalizar:
  Set porstCalculo = Nothing
  Set o_Calculo = Nothing

End Function
Private Function ppCalculoVacacion(ByVal sFechaHora As String, ByVal s_Ejercicio As String, s_Mes As String, ByVal s_Fecha As String) As Boolean
  Dim nTotalRegistros As Long, nRegistro As Long, nDiasAusencia As Long
  Dim sFechaIni As String, sFechaFin As String, sFechaPro As String, sFechaIniPvs As String
  Dim sPersonal As String, sFechaIng As String, sPeriodoRemunera As String
  Dim nRemunera_mn As Double, nRemuxDia_mn As Double, nRemuxAcu_mn As Double
  Dim nRemunera_me As Double, nRemuxDia_me As Double, nRemuxAcu_me As Double
  Dim nRemuxPvs_mn As Double, nProviAnt_mn As Double, nRemuxPer_mn As Double
  Dim nRemuxPvs_me As Double, nProviAnt_me As Double, nRemuxPer_me As Double
  Dim nDiasVacaAnt As Double, nDiasVacaFra As Double, nDiasVacaAcu As Double
  Dim nDiasVacaPer As Double, nDiasVacacion As Double
  Dim nRemuVaca_mn As Double, nRemuVaca_me As Double
  Dim s_CuentaDebe_mn As String, s_CuentaHaber_mn As String
  Dim s_CuentaDebe_me As String, s_CuentaHaber_me As String
  Dim nSecuencia As Long
  Dim porstCalculo As ADODB.Recordset, porstVacacion As ADODB.Recordset
  
  ' Informacion de personal a procesar
  s_Sql = "SELECT DISTINCTROW vac.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn "
  s_Sql = s_Sql & "FROM plpvsvacacion vac "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON vac.codcls=psn.codcls AND vac.codpsn=psn.codpsn "
  s_Sql = s_Sql & "WHERE vac.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND vac.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & sFechaHora & "') "
  s_Sql = s_Sql & "AND (DATE_FORMAT(IFNULL(vac.fechaini, '0000-00-01'), '%Y%m')<='" & s_Ejercicio & s_Mes & "' "
  s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(vac.fechafin, '0000-00-01'), '%Y%m')>='" & s_Ejercicio & s_Mes & "') "
  s_Sql = s_Sql & "AND vac.estadovac<>'" & s_Estado_Blq & "' "
  Set porstCalculo = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstCalculo.EOF And porstCalculo.BOF) And porstCalculo.RecordCount > 0 Then
    ' Creo los arreglos para la actualización
    a_Campos = Array("codcls", "codpvs", "codpsn", "pdopvs", "pdoano", "pdomes", "fechaini", "fechafin", "numerodias", "diasvacper", "diasvacacu", "codmon", "remunera_mn", "remunera_me", "importepvs_mn", "importepvs_me", "imporpvsper_mn", "imporpvsper_me", "imporpvsacu_mn", "imporpvsacu_me", "codcta_debmn", "codcta_habmn", "codcta_debme", "codcta_habme", "fechacan", "estadodet", "usrcre", "fyhcre")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Fecha, TipoDato.Fecha, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Fecha, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    
    nTotalRegistros = CLng(porstCalculo.RecordCount)
    While Not porstCalculo.EOF
      sPersonal = Trim(porstCalculo!codpsn)
      ' Actualizo el registro del personal
      sfmProgreso(0).Caption = " Procesando personal : " & sPersonal & " - " & Trim(porstCalculo!nombrepsn) & " "
      sfmProgreso(0).Refresh
      
      ' Obtengo los importes de provisones anteriores (pendientes)
      s_Sql = "SELECT DISTINCTROW MIN(vac.fechaini)  AS fechainicio, "
      s_Sql = s_Sql & "ROUND(IFNULL(SUM(IFNULL(det.importepvs_mn, 0)), 0), 2) AS provision_mn, "
      s_Sql = s_Sql & "ROUND(IFNULL(SUM(IFNULL(det.importepvs_me, 0)), 0), 2) AS provision_me, "
      s_Sql = s_Sql & "SUM(det.numerodias) AS dprovision "
      s_Sql = s_Sql & "FROM plpvsvacaciondet det "
      s_Sql = s_Sql & "INNER JOIN plpvsvacacion vac ON vac.codcls=det.codcls AND vac.codpsn=det.codpsn AND vac.codpvs=det.codpvs "
      s_Sql = s_Sql & "WHERE det.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND det.codpsn='" & sPersonal & "' "
      s_Sql = s_Sql & "AND CONCAT(det.pdoano, det.pdomes)<'" & s_Ejercicio & s_Mes & "' "
      s_Sql = s_Sql & "AND det.estadodet<>'" & s_Estado_Blq & "'"
      Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstRecordset.EOF And porstRecordset.BOF) Then
        sFechaIng = Format(porstRecordset!FechaInicio, s_FormatoFecha)
        nProviAnt_mn = CDec(porstRecordset!provision_mn)
        nProviAnt_me = CDec(porstRecordset!provision_me)
        nDiasVacaAnt = CDec(porstRecordset!dprovision)
      End If
      porstRecordset.Close
      
      ' Obtengo dias fisicos gozados, compra vacaciones, dias indemnizables
      s_Sql = ""
      For nSecuencia = 1 To 5
        s_Sql = s_Sql & "SELECT asi.codpsn, asi.codpdo, "
        s_Sql = s_Sql & "CONCAT(pdo.anopdo, pdo.mespdo) AS pdovacacion, pdovaca" & nSecuencia & " AS pdovaca, "
        s_Sql = s_Sql & "IFNULL(DateDiff(asi.fechafinvaca" & nSecuencia & ", asi.fechainivaca" & nSecuencia & ") + 1, 0) AS diasvaca "
        s_Sql = s_Sql & "FROM plasistencia asi "
        s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
        s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON pdo.codcls=asi.codcls AND pdo.codpdo=asi.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & s_Ejercicio & s_Mes & "' "
        s_Sql = s_Sql & "INNER JOIN plpvsvacacion vac ON vac.codcls=asi.codcls AND vac.codpsn=asi.codpsn AND vac.pdopvs=asi.pdovaca" & nSecuencia & " AND vac.estadovac<>'" & s_Estado_Blq & "' "
        s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND NOT ISNULL(asi.fechainivaca" & nSecuencia & ") "
        s_Sql = s_Sql & "AND NOT ISNULL(asi.fechafinvaca" & nSecuencia & ") "
        s_Sql = s_Sql & "AND asi.codpsn='" & sPersonal & "' "
        s_Sql = s_Sql & IIf(nSecuencia = 5, "", "UNION ")
      Next nSecuencia
      s_Sql = s_Sql & "ORDER BY codpsn, pdovaca, pdovacacion, codpdo"
      Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      
      nDiasVacacion = 0
      If Not (porstRecordset.EOF And porstRecordset.BOF) Then
        While Not porstRecordset.EOF
          nDiasVacacion = nDiasVacacion + CDec(porstRecordset!diasvaca)
          porstRecordset.MoveNext
        Wend
      End If
      porstRecordset.Close
      
      ' Obtengo ultimo periodo del mes calculado
      sPeriodoRemunera = ""
      s_Sql = "SELECT DISTINCT MAX(res.codpdo) AS pdofin "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON pdo.codcls=res.codcls AND pdo.codpdo=res.codpdo "
      s_Sql = s_Sql & "INNER JOIN plpvsvacacion vac ON res.codcls=vac.codcls AND res.codpsn=vac.codpsn "
      s_Sql = s_Sql & "INNER JOIN plpvsperiodovac pvs ON vac.codcls=pvs.codcls AND vac.codpvs=pvs.codpvs AND res.codcls=pvs.codcls AND res.codcpc=pvs.rembasbeneficio "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & s_Ejercicio & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & s_Mes & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & sPersonal & "' "
      s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(psn.fecbaja, '0000-00-01'), '%Y%m')<'" & s_Ejercicio & s_Mes & "' "
      s_Sql = s_Sql & "AND (DATE_FORMAT(IFNULL(vac.fechaini, '0000-00-01'), '%Y%m')<='" & s_Ejercicio & s_Mes & "' "
      s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(vac.fechafin, '0000-00-01'), '%Y%m')>='" & s_Ejercicio & s_Mes & "') "
      s_Sql = s_Sql & "AND vac.estadovac<>'" & s_Estado_Blq & "' "
      s_Sql = s_Sql & "AND pdo.tpopdo='N' "
      s_Sql = s_Sql & "ORDER BY pdofin"
      Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstRecordset.EOF And porstRecordset.BOF) Then
        sPeriodoRemunera = gdl_Funcion.aTexto(porstRecordset!pdofin)
      End If
      porstRecordset.Close
      
      ' Obtengo la informacion para el calculo
      s_Sql = "SELECT res.codpsn, dxr.fecingreso, psn.fecbaja, vac.codpvs, vac.pdopvs, res.codmon, "
      s_Sql = s_Sql & "vac.fechaini, vac.fechafin, res.codcpc, "
      s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS importe_mn, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS importe_me, "
      s_Sql = s_Sql & "cta.codctavac_debmn, cta.codctavac_habmn, cta.codctavac_debme, cta.codctavac_habme "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
      s_Sql = s_Sql & "INNER JOIN plpvsvacacion vac ON res.codcls=vac.codcls AND res.codpsn=vac.codpsn "
      s_Sql = s_Sql & "INNER JOIN plpvsperiodovac pvs ON vac.codcls=pvs.codcls AND vac.codpvs=pvs.codpvs AND res.codcls=pvs.codcls AND res.codcpc=pvs.rembasbeneficio "
      s_Sql = s_Sql & "LEFT JOIN plctapvs cta ON dxr.codcls=cta.codcls AND dxr.codcco=cta.codcco "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.pdoano='" & s_Ejercicio & "' "
      s_Sql = s_Sql & "AND res.pdomes='" & s_Mes & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & sPeriodoRemunera & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & sPersonal & "' "
      s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(psn.fecbaja, '0000-00-01'), '%Y%m')<'" & s_Ejercicio & s_Mes & "' "
      s_Sql = s_Sql & "AND (DATE_FORMAT(IFNULL(vac.fechaini, '0000-00-01'), '%Y%m')<='" & s_Ejercicio & s_Mes & "' "
      s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(vac.fechafin, '0000-00-01'), '%Y%m')>='" & s_Ejercicio & s_Mes & "') "
      s_Sql = s_Sql & "AND vac.estadovac<>'" & s_Estado_Blq & "' "
      s_Sql = s_Sql & "GROUP BY res.codpsn, dxr.fecingreso, psn.fecbaja, vac.codpvs, vac.pdopvs, res.codmon, vac.fechaini, vac.fechafin, res.codcpc, "
      s_Sql = s_Sql & "cta.codctavac_debmn, cta.codctavac_habmn, cta.codctavac_debme, cta.codctavac_habme "
      s_Sql = s_Sql & "ORDER BY codpsn, vac.codpvs"
      Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstRecordset.EOF And porstRecordset.BOF) And porstRecordset.RecordCount > 0 Then
        ' Cuentas contables de provisión
        s_CuentaDebe_mn = gdl_Funcion.aTexto(porstRecordset!codctavac_debmn)
        s_CuentaHaber_mn = gdl_Funcion.aTexto(porstRecordset!codctavac_habmn)
        s_CuentaDebe_me = gdl_Funcion.aTexto(porstRecordset!codctavac_debme)
        s_CuentaHaber_me = gdl_Funcion.aTexto(porstRecordset!codctavac_habme)
        ' Calculo de provisión base 30 dias mes y 360 dias año
        While Not porstRecordset.EOF
          ' Obtengo el numero de dias de provision mes
          sFechaIni = Format(porstRecordset!fechaini, s_FormatoFecha)
          sFechaFin = Format(porstRecordset!fechafin, s_FormatoFecha)
          sFechaPro = IIf(Format(sFechaFin, "yyyymmdd") <= Format(s_Fecha, "yyyymmdd"), sFechaFin, s_Fecha)
          sFechaIni = IIf(sFechaIng <> "", IIf(Format(sFechaIng, "yyyymmdd") >= Format(sFechaIni, "yyyymmdd"), sFechaIng, sFechaIni), sFechaIni)
          sFechaIniPvs = Format("01/" & s_Mes & "/" & s_Ejercicio, s_FormatoFecha)
          sFechaIniPvs = IIf(sFechaIng <> "", IIf(Format(sFechaIng, "yyyymmdd") >= Format(sFechaIniPvs, "yyyymmdd"), sFechaIng, sFechaIniPvs), sFechaIniPvs)
          
          ' Dias fración mes 30 dias
          nDiasVacaFra = gdl_Funcion.NumeroDias360(CDate(sFechaPro), CDate(sFechaIniPvs), CDate(sFechaPro))
          ' Resta Ausencias encontradas en el período (Si el parámetro así lo indica)
          If s_DsctoAusenciaBS = s_Estado_Act Then
            nDiasAusencia = gdl_Funcion.DiasAusenciaBS(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, sPersonal, sFechaIniPvs, sFechaPro)
          End If
          nDiasVacaFra = nDiasVacaFra - nDiasAusencia
          nDiasVacaFra = Round(nDiasVacaFra / 12, 3)
          nDiasVacaAcu = nDiasVacaAnt + nDiasVacaFra
          
          ' Disminuyo dias vacaciones
          nDiasVacaAcu = nDiasVacaAcu - nDiasVacacion
          
          nDiasVacaPer = 0: nRemuxPer_mn = 0: nRemuxPer_me = 0
          ' Obtengo información periodo provisión anterior
          s_Sql = "SELECT det.diasvacper, det.imporpvsper_mn, det.imporpvsper_me "
          s_Sql = s_Sql & "FROM plpvsvacaciondet det "
          s_Sql = s_Sql & "WHERE det.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND det.codpvs='" & porstRecordset!codpvs & "' "
          s_Sql = s_Sql & "AND det.pdopvs='" & porstRecordset!pdopvs & "' "
          s_Sql = s_Sql & "AND det.codpsn='" & sPersonal & "' "
          s_Sql = s_Sql & "AND CONCAT(det.pdoano, det.pdomes)<'" & s_Ejercicio & s_Mes & "' "
          s_Sql = s_Sql & "ORDER BY det.pdoano DESC, det.pdomes DESC"
          Set porstVacacion = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
          If Not (porstVacacion.EOF And porstVacacion.BOF) And porstVacacion.RecordCount > 0 Then
            nDiasVacaPer = CDec(porstVacacion!diasvacper)
            nRemuxPer_mn = CDec(porstVacacion!imporpvsper_mn)
            nRemuxPer_me = CDec(porstVacacion!imporpvsper_me)
          End If
          porstVacacion.Close
          nDiasVacaPer = nDiasVacaPer + nDiasVacaFra

          ' importes de remuneración MN
          nRemunera_mn = CDec(porstRecordset!importe_mn)
          nRemuxDia_mn = Round(nRemunera_mn / 30, 2)
          nRemuxAcu_mn = Round(nRemuxDia_mn * nDiasVacaAcu, 2)
          nRemuVaca_mn = Round(nRemuxDia_mn * nDiasVacacion, 2)
          nRemuxPvs_mn = Round(nRemuxAcu_mn - (nProviAnt_mn - nRemuVaca_mn), 2)
          
          ' importes de remuneración ME
          nRemunera_me = CDec(porstRecordset!importe_me)
          nRemuxDia_me = Round(nRemunera_me / 30, 2)
          nRemuxAcu_me = Round(nRemuxDia_me * nDiasVacaAcu, 2)
          nRemuVaca_me = Round(nRemuxDia_me * nDiasVacacion, 2)
          nRemuxPvs_me = Round(nRemuxAcu_me - (nProviAnt_me - nRemuVaca_me), 2)

          ' Elimina la provision calculada
          s_Sql = "DELETE det.* "
          s_Sql = s_Sql & "FROM plpvsvacaciondet det "
          s_Sql = s_Sql & "WHERE det.codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND det.codpsn='" & sPersonal & "' "
          s_Sql = s_Sql & "AND det.codpvs='" & porstRecordset!codpvs & "' "
          s_Sql = s_Sql & "AND det.pdoano='" & s_Ejercicio & "' "
          s_Sql = s_Sql & "AND det.pdomes='" & s_Mes & "'"
          If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
          
          ' Realizo el proceso de actualización de los registros
          a_Valores = Array(ps_ClsPlanilla, porstRecordset!codpvs, sPersonal, porstRecordset!pdopvs, s_Ejercicio, s_Mes, Format(sFechaIni, s_FmtFechMysql_0), Format(sFechaPro, s_FmtFechMysql_0), CDec(nDiasVacaFra), CDec(nDiasVacaPer), CDec(nDiasVacaAcu), porstRecordset!codmon, CDec(nRemunera_mn), CDec(nRemunera_me), CDec(nRemuxPvs_mn), CDec(nRemuxPvs_me), CDec(nRemuxPer_mn), CDec(nRemuxPer_me), CDec(nRemuxAcu_mn), CDec(nRemuxAcu_me), s_CuentaDebe_mn, s_CuentaHaber_mn, s_CuentaDebe_me, s_CuentaHaber_me, Format("", s_FmtFechMysql_0), s_Estado_Act, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
          If Not Records_Ins("plpvsvacaciondet", a_Campos, a_Valores, a_Tipos) Then GoTo Error
          porstRecordset.MoveNext
        Wend
        ' Actualizo los sub periodos de vacaciones
        s_Sql = "UPDATE plpvsvacacion "
        s_Sql = s_Sql & "SET estadovac='" & s_Estado_Act & "' "
        s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
        s_Sql = s_Sql & "AND codpsn='" & sPersonal & "' "
        s_Sql = s_Sql & "AND (DATE_FORMAT(IFNULL(fechaini, '0000-00-01'), '%Y%m')<='" & s_Ejercicio & s_Mes & "' "
        s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(fechafin, '0000-00-01'), '%Y%m')>='" & s_Ejercicio & s_Mes & "')"
        If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
      End If
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      pgbProgreso(0).Value = ((nRegistro \ nTotalRegistros) * 100)
      porstCalculo.MoveNext
    Wend
    ' Actualizo periodos de provision de vacaciones
    s_Sql = "UPDATE plpvsperiodovac pdo, plpvsvacacion vac "
    s_Sql = s_Sql & "SET pdo.estadopvs='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "WHERE pdo.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND (DATE_FORMAT(IFNULL(vac.fechaini, '0000-00-01'), '%Y%m')<='" & s_Ejercicio & s_Mes & "' "
    s_Sql = s_Sql & "AND DATE_FORMAT(IFNULL(vac.fechafin, '0000-00-01'), '%Y%m')>='" & s_Ejercicio & s_Mes & "') "
    s_Sql = s_Sql & "AND vac.codcls=pdo.codcls "
    s_Sql = s_Sql & "AND vac.codpvs=pdo.codpvs "
    s_Sql = s_Sql & "AND vac.estadovac='" & s_Estado_Act & "'"
    If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
  End If
  ppCalculoVacacion = True
  GoTo Finalizar

Error:
  ppCalculoVacacion = False
Finalizar:
  Set porstCalculo = Nothing

End Function
Private Function ppDepuraGratificacion(ByVal sFechaHora As String, ByVal s_Ejercicio As String, s_Mes As String, ByVal s_Fecha As String) As Boolean
  
  ' Elimina provisones de graificación
  s_Sql = "DELETE gra.* "
  s_Sql = s_Sql & "FROM plpvsgratifica gra "
  s_Sql = s_Sql & "WHERE gra.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND gra.pdoano='" & s_Ejercicio & "' "
  s_Sql = s_Sql & "AND gra.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & sFechaHora & "') "
  s_Sql = s_Sql & "AND gra.pdomes='" & s_Mes & "' "
  s_Sql = s_Sql & "AND gra.estadogra<>'" & s_Estado_Blq & "'"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
  
  ' Actualizo periodo de provisión de gratificación
  s_Sql = "UPDATE plpvsperiodogra pvs "
  s_Sql = s_Sql & "SET pvs.estadopvs='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE pvs.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND pvs.pdoano='" & s_Ejercicio & "' "
  s_Sql = s_Sql & "AND pvs.estadopvs='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND NOT EXISTS(SELECT * FROM plpvsgratifica gra "
  s_Sql = s_Sql & "WHERE gra.codcls=pvs.codcls "
  s_Sql = s_Sql & "AND gra.pdoano=pvs.pdoano "
  s_Sql = s_Sql & "AND gra.sempvs=pvs.sempvs)"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
  ppDepuraGratificacion = True
  GoTo Finalizar

Error:
  ppDepuraGratificacion = False
Finalizar:

End Function
Private Function ppDepuraVacacion(ByVal sFechaHora As String, ByVal s_Ejercicio As String, s_Mes As String, ByVal s_Fecha As String) As Boolean
  
  ' Elimina la provision calculada
  s_Sql = "DELETE det.* "
  s_Sql = s_Sql & "FROM plpvsvacaciondet det "
  s_Sql = s_Sql & "WHERE det.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND det.pdoano='" & s_Ejercicio & "' "
  s_Sql = s_Sql & "AND det.pdomes='" & s_Mes & "' "
  s_Sql = s_Sql & "AND det.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & sFechaHora & "') "
  s_Sql = s_Sql & "AND det.estadodet<>'" & s_Estado_Blq & "'"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
  
  ' Actualizo sub periodos de provisión
  s_Sql = "UPDATE plpvsvacacion vac "
  s_Sql = s_Sql & "SET vac.estadovac='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE vac.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND vac.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & sFechaHora & "') "
  s_Sql = s_Sql & "AND vac.estadovac='" & s_Estado_Act & "'"
  s_Sql = s_Sql & "AND NOT EXISTS(SELECT * FROM plpvsvacaciondet det "
  s_Sql = s_Sql & "WHERE det.codcls=vac.codcls "
  s_Sql = s_Sql & "AND det.codpvs=vac.codpvs "
  s_Sql = s_Sql & "AND det.codpsn=vac.codpsn "
  s_Sql = s_Sql & "AND det.pdopvs=vac.pdopvs)"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
  
  ' Actualizo ejercicios de provisión
  s_Sql = "UPDATE plpvsperiodovac pvs "
  s_Sql = s_Sql & "SET pvs.estadopvs='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE pvs.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND pvs.estadopvs='" & s_Estado_Act & "'"
  s_Sql = s_Sql & "AND NOT EXISTS(SELECT * FROM plpvsvacacion vac "
  s_Sql = s_Sql & "WHERE vac.codcls=pvs.codcls "
  s_Sql = s_Sql & "AND vac.codpvs=pvs.codpvs "
  s_Sql = s_Sql & "AND vac.estadovac='" & s_Estado_Act & "')"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
  ppDepuraVacacion = True
  GoTo Finalizar

Error:
  ppDepuraVacacion = False
Finalizar:

End Function
Private Function ppNumeroDias(ByVal sMesPago As String, ByVal sAsistencia As String) As Long
  ppNumeroDias = o_Calculo.DiasGratificacion(sMesPago, "N", sAsistencia)
End Function
']
Private Sub cmbPeriodo_Click()
  Dim sFecha As String
  
  ' Obtengo al fecha final del mes
  sFecha = gdl_Funcion.NumeroDiasMes(Left(Trim(cmbPeriodo), 2), Trim(cmbejercicio)) & "/" & Left(Trim(cmbPeriodo), 2) & "/" & Trim(cmbejercicio)
  gdl_Procedure.EditMask "AT", mskFecha, sFecha, s_MdoData_Upd, True, "##/##/####"

End Sub
Private Sub cmdProceso_Click()
  Dim s_OldMessage As String, s_FechaHora As String
  
  ' Realizo las validaciones de los parametros
  If s_OptRegistro = "pvsvacacio" Then
    If o_PvsVacaciones.tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Calculo", vbExclamation: Exit Sub
  Else
    If o_PvsGratifica.tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Calculo", vbExclamation: Exit Sub
  End If
  If Trim(cmbejercicio) = "" Then Beep: MsgBox "Debe indicar el ejercicio de proceso!", vbCritical + vbOKOnly: cmbejercicio.SetFocus: Exit Sub
  If Trim(cmbPeriodo) = "" Then Beep: MsgBox "Debe indicar el periodo de proceso!", vbCritical + vbOKOnly: cmbPeriodo.SetFocus: Exit Sub
  If Not gdl_Funcion.ValidaFecha(mskFecha, 1900) Then mskFecha.SetFocus: Exit Sub
  If Right(mskFecha.ClipText, 4) <> Trim(cmbejercicio) Then Beep: MsgBox "Fecha debe ser del periodo de Proceso", vbExclamation: mskFecha.SetFocus: Exit Sub
  If Mid(mskFecha.ClipText, 3, 2) <> Left(Trim(cmbPeriodo), 2) Then Beep: MsgBox "Fecha debe ser del mes de Proceso", vbExclamation: mskFecha.SetFocus: Exit Sub
  
  If MsgBox("Seguro de Procesar el Período " & Trim(cmbejercicio) & "-" & Trim(cmbPeriodo), vbQuestion + vbDefaultButton2 + vbYesNo) <> vbYes Then
    cmdProceso.SetFocus
    Exit Sub
  End If

  'Inactiva botón de proceso
  cmdProceso.Enabled = False
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Procesando Información ..."
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  s_DsctoAusenciaBS = s_Estado_Ina
  ' Parametro disminuye ausencias
  s_Sql = "SELECT cfg.pdoano, cfg.gratixasis "
  s_Sql = s_Sql & "FROM plcfgempresa cfg "
  s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    s_DsctoAusenciaBS = porstRecordset!gratixasis
  End If
  porstRecordset.Close
  
  s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
  If s_OptRegistro = "pvsvacacio" Then
    ' Barro el arreglo de registros marcadas (bookmarks)
    For nContador = 0 To o_PvsVacaciones.tdbRegistro.SelBookmarks.Count - 1
      o_PvsVacaciones.tdbRegistro.Bookmark = o_PvsVacaciones.tdbRegistro.SelBookmarks(nContador)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, o_PvsVacaciones.tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next nContador
    ' Verifico si no existen movimientos de provisión
    s_Sql = "SELECT COUNT(*) AS registros "
    s_Sql = s_Sql & "FROM plpvsvacaciondet "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codpsn IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    s_Sql = s_Sql & "AND CONCAT(pdoano, pdomes)>'" & Left(Trim(cmbejercicio), 4) & Left(Trim(cmbPeriodo), 2) & "'"
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    nContador = CLng(porstRecordset!registros)
    porstRecordset.Close
  Else
    ' Barro el arreglo de registros marcadas (bookmarks)
    For nContador = 0 To o_PvsGratifica.tdbRegistro.SelBookmarks.Count - 1
      o_PvsGratifica.tdbRegistro.Bookmark = o_PvsGratifica.tdbRegistro.SelBookmarks(nContador)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, o_PvsGratifica.tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next nContador
    ' Verifico si no existen movimientos provisiones
    s_Sql = "SELECT IFNULL(COUNT(*), 0) AS registros "
    s_Sql = s_Sql & "FROM plpvsgratifica gra "
    s_Sql = s_Sql & "INNER JOIN plpvsperiodogra pvs ON gra.codcls=pvs.codcls AND gra.pdoano=pvs.pdoano AND gra.sempvs=pvs.sempvs "
    s_Sql = s_Sql & "WHERE gra.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND gra.pdoano='" & Left(Trim(cmbejercicio), 4) & "' "
    s_Sql = s_Sql & "AND gra.codpsn IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    s_Sql = s_Sql & "AND CONCAT(gra.pdoano, gra.pdomes)>'" & Left(Trim(cmbejercicio), 4) & Left(Trim(cmbPeriodo), 2) & "' "
    s_Sql = s_Sql & "AND (pvs.mesini<='" & Left(Trim(cmbPeriodo), 2) & "' AND pvs.mesfin>='" & Left(Trim(cmbPeriodo), 2) & "')"
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    nContador = CLng(porstRecordset!registros)
    porstRecordset.Close
  End If
  If nContador > 0 Then Beep: MsgBox "Primero depurar provisiones de meses continuos", vbCritical: GoTo Finalizar
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  '[ Proceso de acuerdo a selección
  pgbProgreso(0).Value = 0
  If s_OptRegistro = "pvsvacacio" Then
    If s_Proceso = "Cálculo" Then
      If Not ppCalculoVacacion(s_FechaHora, Left(Trim(cmbejercicio), 4), Left(Trim(cmbPeriodo), 2), Format(mskFecha, s_FormatoFecha)) Then GoTo Error
    Else
      If Not ppDepuraVacacion(s_FechaHora, Left(Trim(cmbejercicio), 4), Left(Trim(cmbPeriodo), 2), Format(mskFecha, s_FormatoFecha)) Then GoTo Error
    End If
  ElseIf s_OptRegistro = "pvsgratifi" Then
    If s_Proceso = "Cálculo" Then
      If Not ppCalculoGratificacion(s_FechaHora, Left(Trim(cmbejercicio), 4), Left(Trim(cmbPeriodo), 2), Format(mskFecha, s_FormatoFecha)) Then GoTo Error
    Else
      If Not ppDepuraGratificacion(s_FechaHora, Left(Trim(cmbejercicio), 4), Left(Trim(cmbPeriodo), 2), Format(mskFecha, s_FormatoFecha)) Then GoTo Error
    End If
  End If
  sfmProgreso(0).Caption = " " & lblTitle & " Finalizado "
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  ']
  MsgBox "Proceso concluyó satisfactoriamente.", vbInformation + vbOKOnly
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Elimino el rango de selección
  gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  ' Reinicializo los mensajes
  MuestraMensaje s_OldMessage
  'Activo el botón de proceso
  cmdProceso.Enabled = True
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()

  ' Establece posición del formulario
  Me.Height = 3165: Me.Width = 7050
  Me.Left = 3000: Me.Top = 2000
  
  ' Verifico que exista y Cargo el Icono del Formulario
  Me.Icon = LoadPicture()
  s_Sql = gdl_Procedure.ps_PathImagen & "proceso.ico"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then
    Me.Icon = LoadPicture(s_Sql)
  End If
  s_OptRegistro = Left(Trim(fMenu.Tag), 10)
  s_Proceso = Choose(CInt(Right(Trim(fMenu.Tag), 1)), "Cálculo", "Depuración")
  Me.Caption = "Proceso de " & s_Proceso
  lblTitle = s_Proceso & " de Provisión de " & IIf(s_OptRegistro = "pvsvacacio", "Vacaciones", "Gratificaciones")
  
  chkDiciembre.Visible = (s_OptRegistro = "pvsgratifi")
  ' Configuro los listados, datos adicionales
  For nContador = (Val(ps_Anyo) - 2) To (Val(ps_Anyo) + 2): cmbejercicio.AddItem Format(nContador, "0000"): Next nContador
  For nContador = 1 To 12: cmbPeriodo.AddItem Choose(nContador, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next nContador
  
  ' Inicializo la fecha de proceso
  gdl_Procedure.EditMask "AT", mskFecha, "", s_MdoData_Ins, True, "##/##/####"
  gdl_Procedure.EditCombo "PK", cmbejercicio, 2, s_MdoData_Ins, True

End Sub
Private Sub Form_Unload(Cancel As Integer)
  fMenu.cmbejercicio.Enabled = Not Cancel
End Sub
Private Sub mskFecha_GotFocus()
  gdl_Procedure.MarcaGet mskFecha
End Sub
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub mskFecha_Validate(Cancel As Boolean)
  If mskFecha.ClipText <> "" Then
    If Not gdl_Funcion.ValidaFecha(mskFecha, 1900) Then mskFecha.SetFocus: Cancel = True: Exit Sub
  End If
  If Right(mskFecha.ClipText, 4) <> Trim(cmbejercicio) Then Beep: MsgBox "Fecha debe ser del periodo a Procesar", vbExclamation: mskFecha.SetFocus: Cancel = True: Exit Sub
  If Mid(mskFecha.ClipText, 3, 2) <> Left(Trim(cmbPeriodo), 2) Then Beep: MsgBox "Fecha debe ser del mes a Procesar", vbExclamation: mskFecha.SetFocus: Cancel = True: Exit Sub
End Sub
