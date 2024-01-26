VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPPDTHPrMes 
   Caption         =   "[título]"
   ClientHeight    =   2760
   ClientLeft      =   2640
   ClientTop       =   3960
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRptAfp 
      Caption         =   "&Reporte AFP"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmpProcAfp 
      Caption         =   "P&rocesar AFP"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CmnDlgUbica 
      Left            =   4920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar PDT"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbEtapa1 
      Height          =   345
      Left            =   225
      TabIndex        =   0
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label LblProces 
      Caption         =   "Procesando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   405
      Width           =   1635
   End
End
Attribute VB_Name = "frmPPDTHPrMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2015-01-14
'he revisado este frm 100% en fuentes de mysql y lo he pasado a sql.
'sin errores. luego he copiado el frm en  OCont_sql
'frm compatible con fuentes de OCont_sql

Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCOHPrDoc As ADODB.Recordset
Public pbNuevo As Boolean
Public pcNroCpb As String

Public unMargenIzquierdo As Integer

'Private Sub cmdRptAfp_Click()
Private Sub cmdRptAfp_Click_2015_07_27_inclu_sdo_igual_compra()
    On Error GoTo Err
    
    Set pocnnMain = New ADODB.Connection
    Set porstCOHPrDoc = New ADODB.Recordset
    With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
    End With
  
    Set pocnnMain = fsdo_doc_pdte_hoy(pocnnMain) 'ini 2015-07-27 Sdo. doc al mes pendiente
  
     Dim cCadReporte  As String
     Dim sTabla As String
     sTabla = "xlsHPrCab"
    'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & sTabla & " ", cCadReporte)
     pocnnMain.Execute fDropTable2(sTabla, 1)
 
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
    cCadReporte = cCadReporte & "SELECT "
    'ini 2015-01-13 error sum de bruto
    cCadReporte = cCadReporte & fConvert103ddmmyyySay("MIN(hpr.feedoc)") & " AS feedoc,"
    cCadReporte = cCadReporte & fConvert103ddmmyyySay("MIN(det.fehope)") & " AS fehpgo," '2014-09-08 adiciona fecha pago
    cCadReporte = cCadReporte & fConvert103ddmmyyySay("MIN(hpr.fehope)") & " AS fehope,"
    'fin 2015-01-13 error sum de bruto
    '**cCadReporte = cCadReporte & "    det.codtdc,"
    If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
        cCadReporte = cCadReporte & "    det.codtdc,"
    Else
        cCadReporte = cCadReporte & "    MAX(det.codtdc) codtdc,"
    End If
    cCadReporte = cCadReporte & "    det.serdoc,"
    cCadReporte = cCadReporte & "    det.nrodoc,"
    cCadReporte = cCadReporte & "    det.codaux,"
    If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
        cCadReporte = cCadReporte & "    aux.rucaux,"
        cCadReporte = cCadReporte & "    " & fIsNull("auxonp.numeroafp,''") & " numeroafp," 'cupss
        cCadReporte = cCadReporte & "    aux.razaux,"
    Else
        cCadReporte = cCadReporte & "    MAX(aux.rucaux) rucaux,"
        cCadReporte = cCadReporte & "    MAX(" & fIsNull("auxonp.numeroafp,''") & ") numeroafp," 'cupss
        cCadReporte = cCadReporte & "    MAX(aux.razaux) razaux,"
    End If
    'ini 2015-01-13 error sum de bruto
    cCadReporte = cCadReporte & "    MAX(ROUND(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_mn, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_mn, 0") & " END, 2)) AS impmnh,"
    cCadReporte = cCadReporte & "    MAX(ROUND(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_me, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_me, 0") & " END, 2)) AS impmeh,"
    'fin 2015-01-13 error sum de bruto
    If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
        cCadReporte = cCadReporte & "    hpr.impir4_mn,"
        cCadReporte = cCadReporte & "    hpr.impir4_me,"
        cCadReporte = cCadReporte & "    import_mn ," 'AFP_ONP_mn
        cCadReporte = cCadReporte & "    import_me ," 'AFP_ONP_me
        cCadReporte = cCadReporte & "    impnet_mn ," 'Neto_mn
        cCadReporte = cCadReporte & "    impnet_me ," 'Neto_me
        'cCadReporte = cCadReporte & "    IFNULL(afp.desafp,'') desafp,"
        cCadReporte = cCadReporte & "    " & fIsNull("afp.desafp,''") & " desafp,"
    Else
        cCadReporte = cCadReporte & "    MAX(hpr.impir4_mn)impir4_mn,"
        cCadReporte = cCadReporte & "    MAX(hpr.impir4_me)impir4_me,"
        cCadReporte = cCadReporte & "    MAX(import_mn)import_mn ," 'AFP_ONP_mn
        cCadReporte = cCadReporte & "    MAX(import_me)import_me," 'AFP_ONP_me
        cCadReporte = cCadReporte & "    MAX(impnet_mn)impnet_mn," 'Neto_mn
        cCadReporte = cCadReporte & "    MAX(impnet_me)impnet_me," 'Neto_me
        'cCadReporte = cCadReporte & "    IFNULL(afp.desafp,'') desafp,"
        cCadReporte = cCadReporte & "    MAX(" & fIsNull("afp.desafp,''") & ") desafp,"
    End If
    cCadReporte = cCadReporte & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN " & fIsNull("det.impmn, 0") & "*-1 ELSE " & fIsNull("det.impmn, 0") & " END), 2) AS impmn,"
    If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
        '2014-08-29 asi esta en la exportacion cCadReporte = cCadReporte & "    aux.codtdi,"
        cCadReporte = cCadReporte & "    nat.codtdi,"
        cCadReporte = cCadReporte & "    nat.numdci,"
        cCadReporte = cCadReporte & "    hpr.indafeir4,"
        cCadReporte = cCadReporte & "    hpr.tpomon,"
    Else
        '2014-08-29 asi esta en la exportacion cCadReporte = cCadReporte & "    MAX(aux.codtdi) codtdi,"
        cCadReporte = cCadReporte & "    MAX(nat.codtdi) codtdi,"
        cCadReporte = cCadReporte & "    MAX(nat.numdci) numdci,"
        cCadReporte = cCadReporte & "    MAX(hpr.indafeir4) indafeir4,"
        cCadReporte = cCadReporte & "    MAX(hpr.tpomon) tpomon,"
    End If
    cCadReporte = cCadReporte & "    ROUND(Avg(det.imptcb), 4) As imptcb "
    cCadReporte = cCadReporte & "FROM (((((cocpbdet det "
    cCadReporte = cCadReporte & "INNER JOIN cohprdoc hpr"
    cCadReporte = cCadReporte & "    ON det.codemp=hpr.codemp AND det.codaux=hpr.codaux AND det.serdoc=hpr.serdoc AND det.nrodoc=hpr.nrodoc) "
    cCadReporte = cCadReporte & "LEFT JOIN tgaux aux "
    cCadReporte = cCadReporte & "    ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
    cCadReporte = cCadReporte & "LEFT JOIN tgauxnat nat "
    cCadReporte = cCadReporte & "    ON aux.codemp=nat.codemp AND aux.codaux=nat.codaux) "
    cCadReporte = cCadReporte & "LEFT JOIN codonpafp auxonp "
    cCadReporte = cCadReporte & "    ON det.codemp=auxonp.codemp AND det.codaux=auxonp.codaux) "
    cCadReporte = cCadReporte & "LEFT JOIN coentidadpen afp "
    cCadReporte = cCadReporte & "    ON auxonp.codemp=afp.codemp AND auxonp.codafp=afp.codafp) "
    '2014-08-25 error where cCadReporte = cCadReporte & "WHERE det.codemp='010' AND det.pdoano='2014' AND det.mespvs='02' AND det.codtdc='02' AND det.tpopvs='C' "
    cCadReporte = cCadReporte & "WHERE det.codemp='" & gsCodEmp & "' "
    cCadReporte = cCadReporte & "AND det.pdoano='" & gsAnoAct & "' "
    cCadReporte = cCadReporte & "AND det.mespvs='" & gsMesAct & "' "
    cCadReporte = cCadReporte & "AND det.codtdc='" & CODTDC_HPR & "' "
    cCadReporte = cCadReporte & "AND det.tpopvs='" & TPOPVS_CAN & "' "
    cCadReporte = cCadReporte & "GROUP BY det.codaux, det.serdoc, det.nrodoc "
  '--------------------------
    pocnnMain.Execute cCadReporte
  
'    sTabla = "tmp_xls_pdte3"
'    pocnnTmp.Execute fDropTable2(sTabla, 1)
'
'    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
'    cCadReporte = cCadReporte & "SELECT "
'    cCadReporte = cCadReporte & "* "
'    cCadReporte = cCadReporte & "From tmp_xls_pdte "
'    cCadReporte = cCadReporte & "Where x_clave "
'    cCadReporte = cCadReporte & "    IN (select x_clave from tmp_xls_pdte2) "
'    pocnnTmp.Execute cCadReporte
    
  With porstCOHPrDoc
    .ActiveConnection = pocnnMain
    
    .Source = "SELECT "
    .Source = .Source & "    a.feedoc,fehpgo,a.fehope,"
    .Source = .Source & "    a.codtdc,a.serdoc,a.nrodoc,a.codaux,rucaux,"
    .Source = .Source & "    impmnh,impmeh,"
    .Source = .Source & "    impir4_mn,impir4_me,import_mn,import_me,impnet_mn,impnet_me,desafp,"
    .Source = .Source & "    impmn,codtdi,numdci,indafeir4,tpomon,"
    .Source = .Source & "    imptcb,"
    '#   a.*,,b.FehOpe
    .Source = .Source & "    b.FehOpe FehOpe1,"
    .Source = .Source & "    IFNULL(b.cDebeMN,0)-IFNULL(b.cHaberMN,0) PgoMN,"
    .Source = .Source & "    IFNULL(b.cDebeME,0)-IFNULL(b.cHaberME,0) PgoME "
    .Source = .Source & "FROM xlsHPrCab a "
    .Source = .Source & "LEFT JOIN tmp_xls_pdte3 b "
    .Source = .Source & "    ON a.CodAux=b.CodAux and a.codtdc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc "
    
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  
  Dim Index As Integer
  Index = 2
  '---reporte
    'gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    gpEncabezadoRpt frmMain.rptMain, "Reporte de AFP", Date, True, True, porstCOHPrDoc
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRRpteAfp.rpt"
      '.WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .WindowShowExportBtn = True
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  
  '--------------------------
  porstCOHPrDoc.Close
  pocnnMain.Close
  Set porstCOHPrDoc = Nothing
  Set pocnnMain = Nothing

  Exit Sub
Err:
'  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Sub

'Private Sub cmdRptAfp1()
'End Sub
'
'Private Sub cmpProcAfp1()
'End Sub

Private Sub ppEtapa_01A()   ' Generacion de Texto en File
  Dim dnContador As Integer
  Dim dsTexto, dsFile As String
  Dim sCaracter As String
  Dim sImporte As String
  
  On Error GoTo CancelaDialogo
  
  dnContador = 0
  pgbEtapa1.Min = 0
  dsFile = "0601" & gsAnoAct & gsMesAct & gsRUCEmp & ".afp"
  CmnDlgUbica.FileName = dsFile
  CmnDlgUbica.CancelError = True
  CmnDlgUbica.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then MsgBox error(Err.Number): Exit Sub
  On Error GoTo 0
  
  Open dsFile For Output As #1
  sCaracter = "|"
  Do
    With porstCOHPrDoc
      If .RecordCount = 0 Then Exit Do
      .MoveFirst
      pgbEtapa1.Max = .RecordCount
      pgbEtapa1.Value = pgbEtapa1.Min
      Dim NumSecu As Integer
      NumSecu = 1
      Do
        dsTexto = ""
        
        dsTexto = dsTexto & gfCeros(Trim(Str(NumSecu)), 5, 0, "0") & sCaracter
        dsTexto = dsTexto & !numeroafp & sCaracter
        dsTexto = dsTexto & IIf(!codtdi = "01", "0", IIf(!codtdi = "02", "1", "")) & sCaracter
        dsTexto = dsTexto & !numdci & sCaracter
        dsTexto = dsTexto & !NomAux & sCaracter
        dsTexto = dsTexto & !ApePatAux & sCaracter
        dsTexto = dsTexto & !ApeMatAux & sCaracter
        dsTexto = dsTexto & " " & sCaracter
        dsTexto = dsTexto & " " & sCaracter
        sImporte = Format(!impmnh, "############.00")
        dsTexto = dsTexto & sImporte & sCaracter
        dsTexto = dsTexto & "0" & sCaracter
        dsTexto = dsTexto & "0" & sCaracter
        dsTexto = dsTexto & "0" & sCaracter
        dsTexto = dsTexto & "I" & sCaracter

 '----------------------------------------
'''        'dsTexto = Trim(IIf(IsNull(!codtdi), "", !codtdi)) & sCaracter
'''
'''        dsTexto = Mid(!TpoDci, 2, 1) & sCaracter
'''
'''        dsTexto = dsTexto & IIf(!TpoDci = "06", !rucaux, !numdci) & sCaracter
'''
'''        'dsTexto = dsTexto & Trim(IIf(IsNull(!numdci), "", !numdci)) & sCaracter
'''        dsTexto = dsTexto & "R" & sCaracter
'''        dsTexto = dsTexto & Trim(IIf(IsNull(!serdoc), "", IIf(Left(!serdoc, 1) = "E", !serdoc, Right(!serdoc, 3)))) & sCaracter
'''        dsTexto = dsTexto & Trim(IIf(IsNull(!nrodoc), "", Mid(!nrodoc, 3, 8))) & sCaracter
'''        sImporte = Format(!ImpMNH, "############.00")
'''        If !tpomon = TPOMON_EXT Then
'''          sImporte = Format(Round(CDec(!ImpMEH) * CDec(!ImpTCb), 2), "############.00")
'''        End If
'''    '    sImporte = Replace(sImporte, ".", "")
'''        dsTexto = dsTexto & Trim(sImporte) & sCaracter
'''        dsTexto = dsTexto & Format(!feedoc, "dd/mm/yyyy") & sCaracter
'''        dsTexto = dsTexto & Format(!fehope, "dd/mm/yyyy") & sCaracter
'''
'''        dsTexto = dsTexto & IIf(!ImpIR4_MN + !ImpIR4_ME <> 0, "1", "0") & sCaracter & "3" & sCaracter & sCaracter
'''        'dsTexto = dsTexto & Trim(!IndAfeIR4) & sCaracter
 '----------------------------------------
        NumSecu = NumSecu + 1
        Print #1, dsTexto
        dnContador = dnContador + 1
        pgbEtapa1.Value = dnContador
        .MoveNext
      Loop Until .EOF
    End With
    Exit Do
  Loop
  Close #1
  MsgBox TEXT_8008, vbInformation

End Sub

Private Sub cmdRptAfp_Click()
  On Error GoTo Err
  
  Set pocnnMain = New ADODB.Connection
  Set porstCOHPrDoc = New ADODB.Recordset
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  
  'Set pocnnMain = fsdo_doc_pdte_hoy(pocnnMain) 'ini 2015-07-27 Sdo. doc al mes pendiente
  '--------------------------
  With porstCOHPrDoc
    .ActiveConnection = pocnnMain
    
'''.Source = "SELECT"
'''.Source = .Source & "    MIN(hpr.feedoc) AS feedoc,"
''''2014-08-26 erro fecha pago .Source = .Source & "    MAX(det.fehope) AS fehope,"
'''.Source = .Source & "    MIN(hpr.fehope) AS fehope,"
'''.Source = .Source & "    det.codtdc,"
'''.Source = .Source & "    det.serdoc,"
'''.Source = .Source & "    det.nrodoc,"
'''.Source = .Source & "    det.codaux,"
'''.Source = .Source & "    aux.rucaux,"
'''.Source = .Source & "    IFNULL(auxonp.numeroafp,'') numeroafp," 'cupss
'''.Source = .Source & "    aux.razaux,"
'''.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh,"
'''.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh,"
'''.Source = .Source & "    hpr.impir4_mn,"
'''.Source = .Source & "    hpr.impir4_me,"
'''.Source = .Source & "    import_mn ," 'AFP_ONP_mn
'''.Source = .Source & "    import_me ," 'AFP_ONP_me
'''.Source = .Source & "    impnet_mn ," 'Neto_mn
'''.Source = .Source & "    impnet_me ," 'Neto_me
'''.Source = .Source & "    IFNULL(afp.desafp,'') desafp,"
'''.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn,"
'''.Source = .Source & "    aux.tpodci,"
'''.Source = .Source & "    nat.numdci,"
'''.Source = .Source & "    hpr.indafeir4,"
'''.Source = .Source & "    hpr.tpomon,"
'''.Source = .Source & "    ROUND(Avg(det.imptcb), 4) As imptcb "
'''.Source = .Source & "FROM (((((cocpbdet det "
'''.Source = .Source & "INNER JOIN cohprdoc hpr"
'''.Source = .Source & "    ON det.codemp=hpr.codemp AND det.codaux=hpr.codaux AND det.serdoc=hpr.serdoc AND det.nrodoc=hpr.nrodoc) "
'''.Source = .Source & "LEFT JOIN tgaux aux "
'''.Source = .Source & "    ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
'''.Source = .Source & "LEFT JOIN tgauxnat nat "
'''.Source = .Source & "    ON aux.codemp=nat.codemp AND aux.codaux=nat.codaux) "
'''.Source = .Source & "LEFT JOIN codonpafp auxonp "
'''.Source = .Source & "    ON det.codemp=auxonp.codemp AND det.codaux=auxonp.codaux) "
'''.Source = .Source & "LEFT JOIN coentidadpen afp "
'''.Source = .Source & "    ON auxonp.codemp=afp.codemp AND auxonp.codafp=afp.codafp) "
''''2014-08-25 error where .Source = .Source & "WHERE det.codemp='010' AND det.pdoano='2014' AND det.mespvs='02' AND det.codtdc='02' AND det.tpopvs='C' "
'''.Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
'''.Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
'''.Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
'''.Source = .Source & "AND det.codtdc='" & CODTDC_HPR & "' "
'''.Source = .Source & "AND det.tpopvs='" & TPOPVS_CAN & "' "
'''
'''.Source = .Source & "GROUP BY det.codaux, det.serdoc, det.nrodoc "

.Source = "SELECT "
'ini 2015-01-13 error sum de bruto
'.Source = .Source & "    MIN(hpr.feedoc) AS feedoc,"
''2014-08-26 erro fecha pago .Source = .Source & "    MAX(det.fehope) AS fehope,"
'.Source = .Source & "    MIN(det.fehope) AS fehpgo," '2014-09-08 adiciona fecha pago
'.Source = .Source & "    MIN(hpr.fehope) AS fehope,"
.Source = .Source & fConvert103ddmmyyySay("MIN(hpr.feedoc)") & " AS feedoc,"
.Source = .Source & fConvert103ddmmyyySay("MIN(det.fehope)") & " AS fehpgo," '2014-09-08 adiciona fecha pago
.Source = .Source & fConvert103ddmmyyySay("MIN(hpr.fehope)") & " AS fehope,"

'fin 2015-01-13 error sum de bruto

'**.Source = .Source & "    det.codtdc,"
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    .Source = .Source & "    det.codtdc,"
Else
    .Source = .Source & "    MAX(det.codtdc) codtdc,"
End If

.Source = .Source & "    det.serdoc,"
.Source = .Source & "    det.nrodoc,"
.Source = .Source & "    det.codaux,"

'''.Source = .Source & "    aux.rucaux,"
'''.Source = .Source & "    IFNULL(auxonp.numeroafp,'') numeroafp," 'cupss
'''.Source = .Source & "    aux.razaux,"
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    .Source = .Source & "    aux.rucaux,"
    .Source = .Source & "    " & fIsNull("auxonp.numeroafp,''") & " numeroafp," 'cupss
    .Source = .Source & "    aux.razaux,"
Else
    .Source = .Source & "    MAX(aux.rucaux) rucaux,"
    .Source = .Source & "    MAX(" & fIsNull("auxonp.numeroafp,''") & ") numeroafp," 'cupss
    .Source = .Source & "    MAX(aux.razaux) razaux,"
End If

'2014-08-29.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh,"
'2014-08-29 .Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh,"
'ini 2015-01-13 error sum de bruto
'.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_mn, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_mn, 0") & " END), 2) AS impmnh,"
'.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_me, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_me, 0") & " END), 2) AS impmeh,"
.Source = .Source & "    MAX(ROUND(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_mn, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_mn, 0") & " END, 2)) AS impmnh,"
.Source = .Source & "    MAX(ROUND(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_me, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_me, 0") & " END, 2)) AS impmeh,"
'fin 2015-01-13 error sum de bruto

'''.Source = .Source & "    hpr.impir4_mn,"
'''.Source = .Source & "    hpr.impir4_me,"
'''.Source = .Source & "    import_mn ," 'AFP_ONP_mn
'''.Source = .Source & "    import_me ," 'AFP_ONP_me
'''.Source = .Source & "    impnet_mn ," 'Neto_mn
'''.Source = .Source & "    impnet_me ," 'Neto_me
'''.Source = .Source & "    IFNULL(afp.desafp,'') desafp,"
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    .Source = .Source & "    hpr.impir4_mn,"
    .Source = .Source & "    hpr.impir4_me,"
    .Source = .Source & "    import_mn ," 'AFP_ONP_mn
    .Source = .Source & "    import_me ," 'AFP_ONP_me
    .Source = .Source & "    impnet_mn ," 'Neto_mn
    .Source = .Source & "    impnet_me ," 'Neto_me
    '.Source = .Source & "    IFNULL(afp.desafp,'') desafp,"
    .Source = .Source & "    " & fIsNull("afp.desafp,''") & " desafp,"
Else
    .Source = .Source & "    MAX(hpr.impir4_mn)impir4_mn,"
    .Source = .Source & "    MAX(hpr.impir4_me)impir4_me,"
    .Source = .Source & "    MAX(import_mn)import_mn ," 'AFP_ONP_mn
    .Source = .Source & "    MAX(import_me)import_me," 'AFP_ONP_me
    .Source = .Source & "    MAX(impnet_mn)impnet_mn," 'Neto_mn
    .Source = .Source & "    MAX(impnet_me)impnet_me," 'Neto_me
    '.Source = .Source & "    IFNULL(afp.desafp,'') desafp,"
    .Source = .Source & "    MAX(" & fIsNull("afp.desafp,''") & ") desafp,"
End If

'2014-08-29 .Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn,"
.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN " & fIsNull("det.impmn, 0") & "*-1 ELSE " & fIsNull("det.impmn, 0") & " END), 2) AS impmn,"

'''.Source = .Source & "    aux.tpodci,"
'''.Source = .Source & "    nat.numdci,"
'''.Source = .Source & "    hpr.indafeir4,"
'''.Source = .Source & "    hpr.tpomon,"
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    '2014-08-29 asi esta en la exportacion .Source = .Source & "    aux.codtdi,"
    .Source = .Source & "    nat.codtdi,"
    .Source = .Source & "    nat.numdci,"
    .Source = .Source & "    hpr.indafeir4,"
    .Source = .Source & "    hpr.tpomon,"
Else
    '2014-08-29 asi esta en la exportacion .Source = .Source & "    MAX(aux.codtdi) codtdi,"
    .Source = .Source & "    MAX(nat.codtdi) codtdi,"
    .Source = .Source & "    MAX(nat.numdci) numdci,"
    .Source = .Source & "    MAX(hpr.indafeir4) indafeir4,"
    .Source = .Source & "    MAX(hpr.tpomon) tpomon,"
End If

.Source = .Source & "    ROUND(Avg(det.imptcb), 4) As imptcb, "
'ini 2015-07-27 Sdo. doc al mes pendiente
.Source = .Source & "    ROUND(MAX(hpr.imptcb), 4) As imptcb_prov "
'fin 2015-07-27 Sdo. doc al mes pendiente

.Source = .Source & "FROM (((((cocpbdet det "
.Source = .Source & "INNER JOIN cohprdoc hpr"
.Source = .Source & "    ON det.codemp=hpr.codemp AND det.codaux=hpr.codaux AND det.serdoc=hpr.serdoc AND det.nrodoc=hpr.nrodoc) "
.Source = .Source & "LEFT JOIN tgaux aux "
.Source = .Source & "    ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
.Source = .Source & "LEFT JOIN tgauxnat nat "
.Source = .Source & "    ON aux.codemp=nat.codemp AND aux.codaux=nat.codaux) "
.Source = .Source & "LEFT JOIN codonpafp auxonp "
.Source = .Source & "    ON det.codemp=auxonp.codemp AND det.codaux=auxonp.codaux) "
.Source = .Source & "LEFT JOIN coentidadpen afp "
.Source = .Source & "    ON auxonp.codemp=afp.codemp AND auxonp.codafp=afp.codafp) "
'2014-08-25 error where .Source = .Source & "WHERE det.codemp='010' AND det.pdoano='2014' AND det.mespvs='02' AND det.codtdc='02' AND det.tpopvs='C' "
.Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
.Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
.Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
.Source = .Source & "AND det.codtdc='" & CODTDC_HPR & "' "
.Source = .Source & "AND det.tpopvs='" & TPOPVS_CAN & "' "

.Source = .Source & "GROUP BY det.codaux, det.serdoc, det.nrodoc "

    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  Dim Index As Integer
  Index = 2
  '---reporte
    'gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
    gpEncabezadoRpt frmMain.rptMain, "Reporte de AFP", Date, True, True, porstCOHPrDoc
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRRpteAfp.rpt"
      '.WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .WindowShowExportBtn = True
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  
  '--------------------------
  porstCOHPrDoc.Close
  pocnnMain.Close
  Set porstCOHPrDoc = Nothing
  Set pocnnMain = Nothing

  Exit Sub
Err:
'  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description


End Sub

Private Sub cmpProcAfp_Click()
  On Error GoTo Err
  
  Dim dnContador As Integer
  
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  LblProces.Visible = True
  pgbEtapa1.Value = 0

  Set pocnnMain = New ADODB.Connection
  Set porstCOHPrDoc = New ADODB.Recordset
  
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
   
  With porstCOHPrDoc
    .ActiveConnection = pocnnMain
'If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
If ps_Plataforma = "xxx" Then '2014-08-28 conver mysql a sql
'ini 2015-01-13 error sum de bruto

''''esta es la version original en mysql
'''.Source = "SELECT"
'''.Source = .Source & "    MIN(hpr.feedoc) AS feedoc,"
''''2014-08-26 erro fecha pago .Source = .Source & "    MAX(det.fehope) AS fehope,"
'''.Source = .Source & "    MIN(hpr.fehope) AS fehope,"
'''.Source = .Source & "    det.codtdc,"
'''.Source = .Source & "    det.serdoc,"
'''.Source = .Source & "    det.nrodoc,"
'''.Source = .Source & "    det.codaux,"
'''.Source = .Source & "    aux.rucaux,"
'''.Source = .Source & "    IFNULL(auxonp.numeroafp,'') numeroafp," 'cupss
'''.Source = .Source & "    aux.razaux,"
''''ini 2015-01-13 error sum de bruto
''''.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh,"
''''.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh,"
'''.Source = .Source & "    MAX(ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2)) AS impmnh,"
'''.Source = .Source & "    MAX(ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2)) AS impmeh,"
''''fin 2015-01-13 error sum de bruto
'''.Source = .Source & "    hpr.impir4_mn,"
'''.Source = .Source & "    hpr.impir4_me,"
'''.Source = .Source & "    import_mn ," 'AFP_ONP_mn
'''.Source = .Source & "    import_me ," 'AFP_ONP_me
'''.Source = .Source & "    impnet_mn ," 'Neto_mn
'''.Source = .Source & "    impnet_me ," 'Neto_me
'''.Source = .Source & "    IFNULL(afp.desafp,'') desafp,"
'''.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn,"
''''.Source = .Source & "    aux.tpodci,"
'''.Source = .Source & "    nat.codtdi,"
'''.Source = .Source & "    nat.numdci,"
'''.Source = .Source & "    nat.nomaux,"
'''.Source = .Source & "    nat.apepataux,"
'''.Source = .Source & "    nat.apemataux,"
'''.Source = .Source & "    hpr.indafeir4,"
'''.Source = .Source & "    hpr.tpomon,"
'''.Source = .Source & "    ROUND(Avg(det.imptcb), 4) As imptcb "
'''.Source = .Source & "FROM (((((cocpbdet det "
'''.Source = .Source & "INNER JOIN cohprdoc hpr"
'''.Source = .Source & "    ON det.codemp=hpr.codemp AND det.codaux=hpr.codaux AND det.serdoc=hpr.serdoc AND det.nrodoc=hpr.nrodoc) "
'''.Source = .Source & "LEFT JOIN tgaux aux "
'''.Source = .Source & "    ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
'''.Source = .Source & "LEFT JOIN tgauxnat nat "
'''.Source = .Source & "    ON aux.codemp=nat.codemp AND aux.codaux=nat.codaux) "
'''.Source = .Source & "LEFT JOIN codonpafp auxonp "
'''.Source = .Source & "    ON det.codemp=auxonp.codemp AND det.codaux=auxonp.codaux) "
'''.Source = .Source & "LEFT JOIN coentidadpen afp "
'''.Source = .Source & "    ON auxonp.codemp=afp.codemp AND auxonp.codafp=afp.codafp) "
''''2014-08-25 error .Source = .Source & "WHERE det.codemp='010' AND det.pdoano='2014' AND det.mespvs='02' AND det.codtdc='02' AND det.tpopvs='C' "
'''.Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
'''.Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
'''.Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
'''.Source = .Source & "AND det.codtdc='" & CODTDC_HPR & "' "
'''.Source = .Source & "AND det.tpopvs='" & TPOPVS_CAN & "' "
'''.Source = .Source & "GROUP BY det.codaux, det.serdoc, det.nrodoc "

'fin 2015-01-13 error sum de bruto
Else
'2014-08-29 esta es la version corregida en mysql y sql
'ini 2015-01-13 error sum de bruto
'    .Source = "SELECT"
'    .Source = .Source & "    MIN(hpr.feedoc) AS feedoc,"
'    '2014-08-26 erro fecha pago .Source = .Source & "    MAX(det.fehope) AS fehope,"
'    .Source = .Source & "    MIN(hpr.fehope) AS fehope,"
    .Source = "SELECT "
    .Source = .Source & fConvert103ddmmyyySay("MIN(hpr.feedoc)") & " AS feedoc,"
    .Source = .Source & fConvert103ddmmyyySay("MIN(hpr.fehope)") & " AS fehope,"
'fin 2015-01-13 error sum de bruto
    
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    .Source = .Source & "    det.codtdc,"
Else
    .Source = .Source & "    MAX(det.codtdc) codtdc,"
End If
    .Source = .Source & "    det.serdoc,"
    .Source = .Source & "    det.nrodoc,"
    .Source = .Source & "    det.codaux,"
    
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    .Source = .Source & "    aux.rucaux,"
    '.Source = .Source & "    IFNULL(auxonp.numeroafp,'') numeroafp," 'cupss
    .Source = .Source & "    " & fIsNull("auxonp.numeroafp,''") & " numeroafp," 'cupss
    .Source = .Source & "    aux.razaux,"
Else
    .Source = .Source & "    MAX(aux.rucaux) rucaux,"
    '.Source = .Source & "    IFNULL(auxonp.numeroafp,'') numeroafp," 'cupss
    .Source = .Source & "    MAX(" & fIsNull("auxonp.numeroafp,''") & ") numeroafp," 'cupss
    .Source = .Source & "    MAX(aux.razaux) razaux,"
End If
'    .Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh,"
'    .Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh,"
'ini 2015-01-13 error sum de bruto
'    .Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_mn, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_mn, 0") & " END), 2) AS impmnh,"
'    .Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_me, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_me, 0") & " END), 2) AS impmeh,"
    .Source = .Source & "    MAX(ROUND(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_mn, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_mn, 0") & " END, 2)) AS impmnh,"
    .Source = .Source & "    MAX(ROUND(CASE det.tpoctb WHEN 'H' THEN  " & fIsNull("hpr.impbru_me, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_me, 0") & " END, 2)) AS impmeh,"
'fin 2015-01-13 error sum de bruto
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    .Source = .Source & "    hpr.impir4_mn,"
    .Source = .Source & "    hpr.impir4_me,"
    .Source = .Source & "    import_mn ," 'AFP_ONP_mn
    .Source = .Source & "    import_me ," 'AFP_ONP_me
    .Source = .Source & "    impnet_mn ," 'Neto_mn
    .Source = .Source & "    impnet_me ," 'Neto_me
    '.Source = .Source & "    IFNULL(afp.desafp,'') desafp,"
    .Source = .Source & "    " & fIsNull("afp.desafp,''") & " desafp,"
Else
    .Source = .Source & "    MAX(hpr.impir4_mn)impir4_mn,"
    .Source = .Source & "    MAX(hpr.impir4_me)impir4_me,"
    .Source = .Source & "    MAX(import_mn)import_mn ," 'AFP_ONP_mn
    .Source = .Source & "    MAX(import_me)import_me," 'AFP_ONP_me
    .Source = .Source & "    MAX(impnet_mn)impnet_mn," 'Neto_mn
    .Source = .Source & "    MAX(impnet_me)impnet_me," 'Neto_me
    '.Source = .Source & "    IFNULL(afp.desafp,'') desafp,"
    .Source = .Source & "    MAX(" & fIsNull("afp.desafp,''") & ") desafp,"
End If
    '.Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn,"
    .Source = .Source & "    ROUND(SUM(CASE det.tpoctb WHEN 'H' THEN " & fIsNull("det.impmn, 0") & "*-1 ELSE " & fIsNull("det.impmn, 0") & " END), 2) AS impmn,"
    
If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
    '.Source = .Source & "    aux.tpodci,"
    .Source = .Source & "    nat.codtdi,"
    .Source = .Source & "    nat.numdci,"
    .Source = .Source & "    nat.nomaux,"
    .Source = .Source & "    nat.apepataux,"
    .Source = .Source & "    nat.apemataux,"
    .Source = .Source & "    hpr.indafeir4,"
    .Source = .Source & "    hpr.tpomon,"
Else
    '.Source = .Source & "    aux.tpodci,"
    .Source = .Source & "    MAX(nat.codtdi) codtdi,"
    .Source = .Source & "    MAX(nat.numdci) numdci,"
    .Source = .Source & "    MAX(nat.nomaux) nomaux,"
    .Source = .Source & "    MAX(nat.apepataux) apepataux,"
    .Source = .Source & "    MAX(nat.apemataux) apemataux,"
    .Source = .Source & "    MAX(hpr.indafeir4) indafeir4,"
    .Source = .Source & "    MAX(hpr.tpomon) tpomon,"
End If
    .Source = .Source & "    ROUND(Avg(det.imptcb), 4) As imptcb "
    .Source = .Source & "FROM (((((cocpbdet det "
    .Source = .Source & "INNER JOIN cohprdoc hpr"
    .Source = .Source & "    ON det.codemp=hpr.codemp AND det.codaux=hpr.codaux AND det.serdoc=hpr.serdoc AND det.nrodoc=hpr.nrodoc) "
    .Source = .Source & "LEFT JOIN tgaux aux "
    .Source = .Source & "    ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
    .Source = .Source & "LEFT JOIN tgauxnat nat "
    .Source = .Source & "    ON aux.codemp=nat.codemp AND aux.codaux=nat.codaux) "
    .Source = .Source & "LEFT JOIN codonpafp auxonp "
    .Source = .Source & "    ON det.codemp=auxonp.codemp AND det.codaux=auxonp.codaux) "
    .Source = .Source & "LEFT JOIN coentidadpen afp "
    .Source = .Source & "    ON auxonp.codemp=afp.codemp AND auxonp.codafp=afp.codafp) "
    '2014-08-25 error .Source = .Source & "WHERE det.codemp='010' AND det.pdoano='2014' AND det.mespvs='02' AND det.codtdc='02' AND det.tpopvs='C' "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
    .Source = .Source & "AND det.codtdc='" & CODTDC_HPR & "' "
    .Source = .Source & "AND det.tpopvs='" & TPOPVS_CAN & "' "
    .Source = .Source & "GROUP BY det.codaux, det.serdoc, det.nrodoc "
End If
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  
  'Etapa1 : Generando Texto segun lectura de Tabla.
  dnContador = 0
  pgbEtapa1.Min = 0
  pgbEtapa1.Value = pgbEtapa1.Min
  ppEtapa_01A
  
  porstCOHPrDoc.Close
  pocnnMain.Close
  Set porstCOHPrDoc = Nothing
  Set pocnnMain = Nothing
  
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Sub

Private Sub Form_Activate()
  LblProces.Visible = False
  cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
  On Error GoTo Err
  
  Dim dnContador As Integer
  
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  LblProces.Visible = True
  pgbEtapa1.Value = 0

  Set pocnnMain = New ADODB.Connection
  Set porstCOHPrDoc = New ADODB.Recordset
  
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
   
  With porstCOHPrDoc
    .ActiveConnection = pocnnMain
'ini 2015-01-14 conver a sql
'''If ps_Plataforma = pSrvMySql Then '2014-08-28 conver mysql a sql
'''    .Source = "SELECT det.codaux, aux.tpodci, aux.rucaux, nat.numdci, det.codtdc, det.serdoc, det.nrodoc, "
'''    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn, "
'''    '2015-01-13 error sum de bruto.Source = .Source & "MIN(hpr.feedoc) AS feedoc, MAX(det.fehope) AS fehope, hpr.indafeir4, hpr.impir4_mn, hpr.impir4_me, import_mn, import_me, hpr.tpomon, "
'''    .Source = .Source & fConvert103ddmmyyySay("MIN(hpr.feedoc)") & "  AS feedoc,"
'''    .Source = .Source & fConvert103ddmmyyySay("MAX(det.fehope)") & "  AS fehope,"
'''    .Source = .Source & "hpr.indafeir4, hpr.impir4_mn, hpr.impir4_me, import_mn, import_me, hpr.tpomon, "
''''ini 2015-01-13 error sum de bruto
'''    '.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh, "
'''    '.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh, "
'''    .Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END, 2)) AS impmnh, "
'''    .Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END, 2)) AS impmeh, "
''''fin 2015-01-13 error sum de bruto
'''Else
'''    'sql8 11/08/12.Source = "SELECT det.codaux, aux.tpodci, aux.rucaux, nat.numdci, det.codtdc, det.serdoc, det.nrodoc, "
'''    .Source = "SELECT det.codaux, det.serdoc,det.nrodoc, MAX(aux.tpodci) tpodci, MAX(aux.rucaux) rucaux, MAX(nat.numdci) numdci, MAX(det.codtdc) codtdc,  "
'''    'sql8 11/08/12.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn, "
'''    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN " & fIsNull("det.impmn, 0") & "*-1 ELSE " & fIsNull("det.impmn, 0") & " END), 2) AS impmn, "
'''    'sql8 11/08/12.Source = .Source & "MIN(hpr.feedoc) AS feedoc, MAX(det.fehope) AS fehope, hpr.indafeir4, hpr.impir4_mn, hpr.impir4_me, hpr.tpomon, "
'''    '2015-01-13 error sum de bruto .Source = .Source & "MIN(hpr.feedoc) AS feedoc, MAX(det.fehope) AS fehope, MAX(hpr.indafeir4) indafeir4 , MAX(hpr.impir4_mn) impir4_mn, MAX(hpr.impir4_me) impir4_me, MAX(hpr.tpomon) tpomon, "
'''    .Source = .Source & fConvert103ddmmyyySay("MIN(hpr.feedoc)") & " AS feedoc, "
'''    .Source = .Source & fConvert103ddmmyyySay("MAX(det.fehope)") & " AS fehope,"
'''    .Source = .Source & "MAX(hpr.indafeir4) indafeir4 , MAX(hpr.impir4_mn) impir4_mn, MAX(hpr.impir4_me) impir4_me, MAX(hpr.tpomon) tpomon, "
'''    'sql8 11/08/12.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh, "
'''    'sql8 11/08/12.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh, "
''''ini 2015-01-13 error sum de bruto
'''    .Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN " & fIsNull("hpr.impbru_mn, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_mn, 0") & " END, 2)) AS impmnh, "
'''    .Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN " & fIsNull("hpr.impbru_me, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_me, 0") & " END, 2)) AS impmeh, "
''''fin 2015-01-13 error sum de bruto
'''End If

    'ini 2015-01-14 conver a sql
    '.Source = "SELECT det.codaux, aux.tpodci, aux.rucaux, nat.numdci, det.codtdc, det.serdoc, det.nrodoc, "
    .Source = "SELECT det.codaux,  "
    If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "aux.tpodci, aux.rucaux, nat.numdci, det.codtdc,  "
    Else
        .Source = .Source & "MAX(aux.tpodci) tpodci,"
        .Source = .Source & "MAX(aux.rucaux) rucaux,"
        .Source = .Source & "MAX(nat.numdci) numdci,"
        .Source = .Source & "MAX(det.codtdc) codtdc,"
    End If
    .Source = .Source & " det.serdoc, det.nrodoc, "
    'fin 2015-01-14 conver a sql
   
    '.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn, "
    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN " & fIsNull("det.impmn, 0") & "*-1 ELSE " & fIsNull("det.impmn, 0") & " END), 2) AS impmn, "
    
    '2015-01-13 error sum de bruto.Source = .Source & "MIN(hpr.feedoc) AS feedoc, MAX(det.fehope) AS fehope, hpr.indafeir4, hpr.impir4_mn, hpr.impir4_me, import_mn, import_me, hpr.tpomon, "
    .Source = .Source & fConvert103ddmmyyySay("MIN(hpr.feedoc)") & "  AS feedoc,"
    .Source = .Source & fConvert103ddmmyyySay("MAX(det.fehope)") & "  AS fehope,"
    'ini 2015-01-14 conver a sql
    '.Source = .Source & "hpr.indafeir4, hpr.impir4_mn, hpr.impir4_me, import_mn, import_me, hpr.tpomon, "
    If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "hpr.indafeir4, hpr.impir4_mn, hpr.impir4_me, import_mn, import_me, hpr.tpomon, "
    Else
        .Source = .Source & "MAX(hpr.indafeir4) indafeir4,"
        .Source = .Source & "MAX(hpr.impir4_mn) impir4_mn,"
        .Source = .Source & "MAX(hpr.impir4_me) impir4_me,"
        .Source = .Source & "MAX(import_mn) import_mn,"
        .Source = .Source & "MAX(import_me) import_me,"
        .Source = .Source & "MAX(hpr.tpomon) tpomon,"
    End If
    'fin 2015-01-14 conver a sql
    
'ini 2015-01-13 error sum de bruto
    '.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh, "
    '.Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh, "
    
    '.Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END, 2)) AS impmnh, "
    '.Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END, 2)) AS impmeh, "
    
    .Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN " & fIsNull("hpr.impbru_mn, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_mn, 0") & " END, 2)) AS impmnh, "
    .Source = .Source & "MAX(ROUND(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN " & fIsNull("hpr.impbru_me, 0") & "*-1 ELSE " & fIsNull("hpr.impbru_me, 0") & " END, 2)) AS impmeh, "
    
'fin 2015-01-13 error sum de bruto

'fin 2015-01-14 conver a sql


    .Source = .Source & "ROUND(AVG(det.imptcb), 4) AS imptcb "
    .Source = .Source & "FROM (((cocpbdet det "
    .Source = .Source & "INNER JOIN cohprdoc hpr ON det.codemp=hpr.codemp AND det.codaux=hpr.codaux AND det.serdoc=hpr.serdoc AND det.nrodoc=hpr.nrodoc) "
    .Source = .Source & "LEFT JOIN tgaux aux ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
    .Source = .Source & "LEFT JOIN tgauxnat nat ON aux.codemp=nat.codemp AND aux.codaux=nat.codaux) "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
    .Source = .Source & "AND det.codtdc='" & CODTDC_HPR & "' "
    .Source = .Source & "AND det.tpopvs='" & TPOPVS_CAN & "' "
    .Source = .Source & "GROUP BY det.codaux, det.serdoc, det.nrodoc "
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  
  'Etapa1 : Generando Texto segun lectura de Tabla.
  dnContador = 0
  pgbEtapa1.Min = 0
  pgbEtapa1.Value = pgbEtapa1.Min
  ppEtapa_01
  
  porstCOHPrDoc.Close
  pocnnMain.Close
  Set porstCOHPrDoc = Nothing
  Set pocnnMain = Nothing
  
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub ppEtapa_01()   ' Generacion de Texto en File
  Dim dnContador As Integer
  Dim dsTexto, dsFile As String
  Dim sCaracter As String
  Dim sImporte As String
  
  On Error GoTo CancelaDialogo
  
  dnContador = 0
  pgbEtapa1.Min = 0
  dsFile = "0601" & gsAnoAct & gsMesAct & gsRUCEmp & ".4ta"
  CmnDlgUbica.FileName = dsFile
  CmnDlgUbica.CancelError = True
  CmnDlgUbica.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then MsgBox error(Err.Number): Exit Sub
  On Error GoTo 0
  
  Open dsFile For Output As #1
  sCaracter = "|"
  Do
    With porstCOHPrDoc
      If .RecordCount = 0 Then Exit Do
      .MoveFirst
      pgbEtapa1.Max = .RecordCount
      pgbEtapa1.Value = pgbEtapa1.Min
      Do
        dsTexto = ""
        'dsTexto = Trim(IIf(IsNull(!codtdi), "", !codtdi)) & sCaracter
        
        dsTexto = Mid(!TpoDci, 2, 1) & sCaracter
        
        dsTexto = dsTexto & IIf(!TpoDci = "06", !rucaux, !numdci) & sCaracter
        
        'dsTexto = dsTexto & Trim(IIf(IsNull(!numdci), "", !numdci)) & sCaracter
        dsTexto = dsTexto & "R" & sCaracter
        dsTexto = dsTexto & Trim(IIf(IsNull(!serdoc), "", IIf(Left(!serdoc, 1) = "E", !serdoc, Right(!serdoc, 3)))) & sCaracter
        dsTexto = dsTexto & Trim(IIf(IsNull(!nrodoc), "", Mid(!nrodoc, 3, 8))) & sCaracter
        sImporte = Format(!impmnh, "############.00")
        If !tpomon = TPOMON_EXT Then
          sImporte = Format(Round(CDec(!ImpMEH) * CDec(!ImpTCb), 2), "############.00")
        End If
    '    sImporte = Replace(sImporte, ".", "")
        dsTexto = dsTexto & Trim(sImporte) & sCaracter
        dsTexto = dsTexto & Format(!feedoc, "dd/mm/yyyy") & sCaracter
        dsTexto = dsTexto & Format(!fehope, "dd/mm/yyyy") & sCaracter
                
        ' //TC 10-11-2014 dsTexto = dsTexto & IIf(!ImpIR4_MN + !ImpIR4_ME <> 0, "1", "0") & sCaracter & "3" & sCaracter & sCaracter //TC 10-11-2014
        dsTexto = dsTexto & IIf(!ImpIR4_MN + !ImpIR4_ME <> 0, "1", "0") & sCaracter & "" & sCaracter & sCaracter
        
        'dsTexto = dsTexto & Trim(!IndAfeIR4) & sCaracter
        Print #1, dsTexto
        dnContador = dnContador + 1
        pgbEtapa1.Value = dnContador
        .MoveNext
      Loop Until .EOF
    End With
    Exit Do
  Loop
  Close #1
  MsgBox TEXT_8008, vbInformation

End Sub

Private Sub Form_Load()
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  LblProces.Caption = Choose(gsIdioma, "Procesando", "Processing")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']
   unMargenIzquierdo = 240             'Margen izquierdo.

End Sub


