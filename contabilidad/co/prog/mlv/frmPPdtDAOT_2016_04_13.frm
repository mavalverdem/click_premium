VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPPDTDAOT 
   Caption         =   "[título]"
   ClientHeight    =   3240
   ClientLeft      =   2640
   ClientTop       =   3960
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmPPdtDAOT.frx":0000
      Left            =   3345
      List            =   "frmPPdtDAOT.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Vista Preliminar"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1725
      Picture         =   "frmPPdtDAOT.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Reporte de validación"
      Top             =   2565
      Width           =   1150
   End
   Begin MSComDlg.CommonDialog CmnDlgUbica 
      Left            =   165
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   375
      TabIndex        =   2
      Top             =   2565
      Width           =   1150
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3060
      TabIndex        =   1
      Top             =   2565
      Width           =   1150
   End
   Begin ComctlLib.ProgressBar pgbEtapa1 
      Height          =   345
      Left            =   225
      TabIndex        =   0
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar PgBEtapa2 
      Height          =   345
      Left            =   225
      TabIndex        =   5
      Top             =   1815
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   0
      Left            =   2490
      TabIndex        =   8
      Top             =   165
      Width           =   765
   End
   Begin VB.Label LblProces 
      Caption         =   "Procesando Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   6
      Top             =   1545
      Width           =   2355
   End
   Begin VB.Label LblProces 
      Caption         =   "Procesando Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   690
      Width           =   2355
   End
End
Attribute VB_Name = "frmPPDTDAOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private udFecha As Date
Private unCopias As Integer
Private unMargenIzquierdo As Integer
Private usDEstino As String
Private usOrientacionRpt As String
Private usOrientacionOri As String

Private pocnnMain As ADODB.Connection
Public pocnnConf As ADODB.Connection
Public porstCOCprDoc As ADODB.Recordset
Public porstCOVtaDoc As ADODB.Recordset
Public porstTGEMP As ADODB.Recordset
Public pbNuevo As Boolean
Public pcNroCpb As String

Private Sub cmdImprimir_Click(Index As Integer)
''2015-02-19 deberia general igual al filtro ejemplo
'.Source = .Source & "AND (concat(a.pdoano,a.mespvs)>='" & gsAnoAct & "01' " & " AND concat(a.pdoano,a.mespvs)<='201502') " & "
'(nose pone este contexto) AND year(a.feedoc)='" & gsAnoAct & "' "

  On Error GoTo Err
  
  Dim porstMRp As New ADODB.Recordset
  Dim sSentencia As String, sMoneda As String
  Dim nRegistros As Long
  
  cmdAceptar.Enabled = False
  cmdImprimir(0).Enabled = False
  cmdSalir.Enabled = False
  
  ' Aperturo la conexión
  Set pocnnMain = New ADODB.Connection
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  ' Instancio el recordset de reporte
  With porstMRp
    .ActiveConnection = pocnnMain
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  '[ Registro de gastos
  sSentencia = "SELECT c.RUCAux, c.RazAux, a.MesPvs, a.FeEDoc, a.FehOpe, a.CodDro, "
  sSentencia = sSentencia & "a.NroCpb, b.AbvTDc, a.SerDoc, a.NroDoc, a.RefDoc, a.NroCDt, a.FehCDt, "
  sSentencia = sSentencia & "(a.ImpOGr_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOgr, "
  sSentencia = sSentencia & "(a.ImpOGN_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOGN, "
  sSentencia = sSentencia & "(a.ImpONG_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpONG, "
  sSentencia = sSentencia & "(a.ImpExo_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpExo, "
  sSentencia = sSentencia & "(a.ImpIGV_OGr_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGr, "
  sSentencia = sSentencia & "(a.ImpIGV_OGN_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGN, "
  sSentencia = sSentencia & "(a.ImpIGV_ONG_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVONG, "
  sSentencia = sSentencia & "(a.ImpISC_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpISC, "
  sSentencia = sSentencia & "(a.ImpOIm_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOIm, "
  sSentencia = sSentencia & "(a.ImpTot_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot, b.CodTDc "
  sSentencia = sSentencia & "FROM ((COCprDoc a "
  sSentencia = sSentencia & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
  sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
'ini 2015-02-19 correccion segun teo
''  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
''  sSentencia = sSentencia & "AND a.MesPvs<='" & gsMesAct & "' "
  sSentencia = sSentencia & " AND (concat(a.pdoano,a.mespvs)>='" & gsAnoAct & "01' "
  sSentencia = sSentencia & " AND concat(a.pdoano,a.mespvs)<='" & Trim(Str(Val(gsAnoAct) + 1)) & "02') "
  '2015-02-23 cambios segun teo sSentencia = sSentencia & " AND year(a.feedoc)='" & gsAnoAct & "' "
  sSentencia = sSentencia & " AND year(a.feedoc)<='" & gsAnoAct & "' "
'ini 2015-02-19 correccion segun teo
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
  sSentencia = sSentencia & "ORDER BY c.RUCAux, a.MesPvs, a.FeEDoc, a.CodDro, a.NroCpb ASC"
  ' Aperturo el listado de registros
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
  End With
  gpEncabezadoRpt frmMain.rptMain, "Registro de Gastos" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, False, porstMRp
  With frmMain.rptMain
    .ReportFileName = gsRutRpt & "rptRGastos.rpt"
    .WindowState = crptMaximized
    .MarginLeft = unMargenIzquierdo
    .Destination = crptToWindow
    .Action = 1
  End With
  ']

  '[ Registro de ingresos
  sSentencia = "SELECT c.RucAux, c.RazAux, a.MesPvs, a.FeEDoc, a.FehOpe, "
  sSentencia = sSentencia & "a.CodDro, a.NroCpb, b.AbvTDc, a.SerDoc, a. NroDoc, "
  sSentencia = sSentencia & "a.SerDoc_Fin, a.NroDoc_Fin , a.RefDoc, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodDro, '-', a.NroCpb)", "(a.CodDro+'-'+a.NroCpb)") & " AS  cx1, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc, '-', a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS  cx2, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc_Fin, '-', a.NroDoc_Fin)", "(a.SerDoc_Fin+'-'+a.NroDoc_Fin)") & " AS cx3, "
  sSentencia = sSentencia & "(a.ImpOGr_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmOgr, "
  sSentencia = sSentencia & "(a.ImpExp_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmExp, "
  sSentencia = sSentencia & "(a.ImpExo_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmExo, "
  sSentencia = sSentencia & "(a.ImpIGV_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmIgv, "
  sSentencia = sSentencia & "(a.ImpISC_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmISC, "
  sSentencia = sSentencia & "(a.ImpOIm_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmOlm, "
  sSentencia = sSentencia & "(a.ImpTot_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmTot, "
  sSentencia = sSentencia & "b.CodTDc "
  sSentencia = sSentencia & "FROM ((COVtaDoc a "
  sSentencia = sSentencia & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
  sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
'ini 2015-02-19 correccion segun teo
'2015-02-20 para el ingreso se mantiene condicion original
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.Mespvs<='" & gsMesAct & "' "
'''  sSentencia = sSentencia & " AND (concat(a.pdoano,a.mespvs)>='" & gsAnoAct & "01' "
'''  sSentencia = sSentencia & " AND concat(a.pdoano,a.mespvs)<='" & Trim(Str(Val(gsAnoAct) + 1)) & "02') "
'''  sSentencia = sSentencia & " AND year(a.feedoc)='" & gsAnoAct & "' "
'fin 2015-02-19 correccion segun teo
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
  sSentencia = sSentencia & "ORDER BY c.RucAux, a.MesPvs, a.FeEDoc, a.CodTDc, a.SerDoc, a.NroDoc  ASC"
  ' Aperturo el listado de registros
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
  End With
  gpEncabezadoRpt frmMain.rptMain, "Registro de Ingresos" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, False, porstMRp
  With frmMain.rptMain
    .ReportFileName = gsRutRpt & "rptRIngresos.rpt"
    .WindowState = crptMaximized
    .MarginLeft = unMargenIzquierdo
    .Destination = crptToWindow
    .Action = 1
  End With
  ']
  cmdAceptar.Enabled = True
  cmdImprimir(0).Enabled = True
  pocnnMain.Close
  Set pocnnMain = Nothing
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  Exit Sub
  
Err:
  Set porstMRp = Nothing
  If pocnnMain.State = adStateOpen Then pocnnMain.Close
  Set pocnnMain = Nothing
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub Form_Activate()
   LblProces(0).Visible = False
   LblProces(1).Visible = False
   cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
  ' On Error GoTo Err
   
  Dim dnContador As Integer
 
  cmdImprimir(0).Enabled = False
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  LblProces(0).Visible = True
  LblProces(1).Visible = False
  pgbEtapa1.Value = 0
  PgBEtapa2.Value = 0

  'Declaración de Variables.
   
  'Abrir Tablas.
   Set pocnnMain = New ADODB.Connection
   Set pocnnConf = New ADODB.Connection
   Set porstTGEMP = New ADODB.Recordset
   Set porstCOCprDoc = New ADODB.Recordset
   Set porstCOVtaDoc = New ADODB.Recordset

   With pocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG  & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With pocnnConf
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With porstTGEMP
      .ActiveConnection = pocnnConf
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
   End With
   With porstCOCprDoc
      .ActiveConnection = pocnnMain
      .Source = "SELECT b.Tpoper, b.RucAux, "
      .Source = .Source & "SUM((CASE d.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_MN+a.ImpOGn_MN+a.ImpONG_MN+a.ImpExo_mn) * -1 ELSE (a.ImpOGr_MN+a.ImpOGn_MN+a.ImpONG_MN+a.ImpExo_mn) END)) AS Total, "
      .Source = .Source & "b.RazAux, c.ApePatAux, c.ApeMatAux, c.NomAux, a.CodAux "
      .Source = .Source & "FROM CoCprDoc a "
      .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux "
      .Source = .Source & "LEFT JOIN TGAuxNat c ON a.codemp=c.codemp AND a.CodAux=c.CodAux "
      .Source = .Source & "LEFT JOIN TgTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      
      '2015-02-19 .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      'cmabio seun teo .Source = .Source & "AND (concat(a.pdoano,a.mespvs)>='" & gsAnoAct & "01' " & " AND concat(a.pdoano,a.mespvs)<='201502') " & " AND year(a.feedoc)='" & gsAnoAct & "' "
      
      .Source = .Source & "AND (concat(a.pdoano,a.mespvs)>='" & gsAnoAct & "01' " & " AND concat(a.pdoano,a.mespvs)<='" & Str(Val(gsAnoAct) + 1) & "02') " & " AND year(a.feedoc)='" & gsAnoAct & "' "
      '"01' el 1er mes del año / pero año siguiente menor o igual  concat(a.pdoano,a.mespvs)<='201502') (sumar el año) y siempre mes=2
      
      .Source = .Source & "GROUP BY a.CodAux, b.RazAux, c.ApePatAux, c.ApeMatAux, c.NomAux, b.Tpoper, b.RucAux "
      If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "HAVING Total>" & (CDec(gnImpUIT) * 2) & " "
      ElseIf ps_Plataforma = pSrvSql Then
        .Source = .Source & "HAVING SUM((CASE d.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_MN+a.ImpOGn_MN+a.ImpONG_MN+a.ImpExo_mn) * -1 ELSE (a.ImpOGr_MN+a.ImpOGn_MN+a.ImpONG_MN+a.ImpExo_mn) END))>" & (CDec(gnImpUIT) * 2) & " "
      End If
      .Source = .Source & "ORDER BY a.CodAux"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
      .Open
      .Properties("Unique Table").Value = "COCprDoc"
   End With
   
'   pocnnMain.BeginTrans                'INICIA TRANSACCION.
 
  'Etapa1 : Generando Texto segun lectura de Tabla.
   
   dnContador = 0
   pgbEtapa1.Min = 0
''   pgbEtapa1.Max = 4
   pgbEtapa1.Value = pgbEtapa1.Min
   
   ppEtapa_01
   
   With porstCOVtaDoc
      .ActiveConnection = pocnnMain
      .Source = "SELECT b.Tpoper, b.RucAux, "
      .Source = .Source & "ROUND(SUM(CASE d.SgnTDc WHEN " & SGNTDC_NEG & " THEN ((a.ImpOGr_MN+a.ImpExp_MN+a.ImpExo_mn) * -1) ELSE (a.ImpOGr_MN+a.ImpExp_MN+a.ImpExo_mn) END), 2) AS Total, "
      .Source = .Source & "b.RazAux, c.ApePatAux, c.ApeMatAux, c.NomAux, a.CodAux "
      .Source = .Source & "FROM CoVtaDoc a "
      .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux "
      .Source = .Source & "LEFT JOIN TGAuxNat c ON a.codemp=c.codemp AND a.CodAux=c.CodAux "
      .Source = .Source & "LEFT JOIN TgTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "GROUP BY a.CodAux, b.RazAux, c.ApePatAux, c.ApeMatAux, c.NomAux, b.Tpoper, b.RucAux "
      If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "HAVING Total>" & (CDec(gnImpUIT) * 2) & " "
      ElseIf ps_Plataforma = pSrvSql Then
        .Source = .Source & "HAVING SUM((CASE d.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_MN+a.ImpExp_MN+a.ImpExo_mn) * -1 ELSE (a.ImpOGr_MN+a.ImpExp_MN+a.ImpExo_mn END)))>" & (CDec(gnImpUIT) * 3) & " "
      End If
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockBatchOptimistic
      .Open
      .Properties("Unique Table").Value = "COVtaDoc"
   End With
   LblProces(1).Visible = True
   
   ppEtapa_02
   
   porstCOCprDoc.Close
   porstCOVtaDoc.Close
   pocnnConf.Close
   pocnnMain.Close
   Set porstTGEMP = Nothing
   Set porstCOCprDoc = Nothing
   Set porstCOVtaDoc = Nothing
   Set pocnnConf = Nothing
   Set pocnnMain = Nothing
   
   MsgBox TEXT_8008, vbInformation
   cmdImprimir(0).Enabled = True
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

Private Sub ppEtapa_01()   ' Generacion de Texto en File Costos
   Dim dnContador As Integer, dnCaracter As Integer
   Dim dsTexto, dsFile As String
   
   dnContador = 0
   pgbEtapa1.Min = 0
    With porstTGEMP
      .Source = "Select RucEmp From TGEMP Where CodEmp='" & gsCodEmp & "'"
      .Open
   End With
   dsFile = "Costos.TXT"
   CmnDlgUbica.FileName = dsFile
   CmnDlgUbica.ShowSave
   Open dsFile For Output As #1
   Do
      With porstCOCprDoc
         If .RecordCount = 0 Then
            Exit Do
         End If
         .MoveFirst
         pgbEtapa1.Max = .RecordCount
         pgbEtapa1.Value = pgbEtapa1.Min
         Do
            dnContador = dnContador + 1
            dsTexto = Trim(Str(dnContador)) & "|"
            dsTexto = dsTexto & "6|" & porstTGEMP!RUCEmp & "|"
            dsTexto = dsTexto & gsAnoAct & "|"
            dsTexto = dsTexto & IIf(!TpoPer = TPOPER_JUR, "02", "01") & "|"
            dsTexto = dsTexto & "6|" & Trim(!rucaux) & "|"
            dsTexto = dsTexto & Trim(Str(gfRedond(!Total, 0))) & "|"
            dsTexto = dsTexto & Trim(!ApePatAux) & "|"
            dsTexto = dsTexto & Trim(!ApeMatAux) & "|"
            If Not IsNull(!NomAux) Then
              dnCaracter = InStr(1, Trim(!NomAux), " ")
              dnCaracter = IIf(dnCaracter <> 0, dnCaracter - 1, Len(Trim(!NomAux)))
              dsTexto = dsTexto & Mid(Trim(!NomAux), 1, dnCaracter) & "|"
              dnCaracter = InStr(1, Trim(!NomAux), " ")
              If dnCaracter <> 0 Then
                dnCaracter = IIf(dnCaracter <> 0, dnCaracter + 1, dnCaracter)
                dsTexto = dsTexto & Mid(Trim(!NomAux), dnCaracter) & "|"
              Else
                dsTexto = dsTexto & "|"
              End If
            Else
              dsTexto = dsTexto & "||"
            End If
            '2015-02-19 dsTexto = dsTexto & Trim(!razAux) & "|"
            dsTexto = dsTexto & IIf(Mid(!codaux, 1, 1) = "2", Trim(!razAux), "") & "|"
            Print #1, dsTexto
            'dnContador = dnContador + 1
            pgbEtapa1.Value = dnContador
            .MoveNext
         Loop Until .EOF
      End With
      Exit Do
   Loop
   Close #1
   porstTGEMP.Close
End Sub

Private Sub ppEtapa_02()   ' Generacion de Texto en File Ingresos
   Dim dnContador As Integer, dnCaracter As Integer
   Dim dsTexto, dsFile As String
   
   dnContador = 0
   PgBEtapa2.Min = 0
   With porstTGEMP
      .Source = "Select RucEmp From TGEMP Where CodEmp='" & gsCodEmp & "'"
      .Open
   End With
   'Open "C:\Owl-paqu\Angel.TXT" For Output As #1
   dsFile = "Ingresos.TXT"
   CmnDlgUbica.FileName = dsFile
   CmnDlgUbica.ShowSave
   Open dsFile For Output As #2
   Do
      With porstCOVtaDoc
         If .RecordCount = 0 Then
            Exit Do
         End If
         .MoveFirst
         PgBEtapa2.Max = .RecordCount
         PgBEtapa2.Value = PgBEtapa2.Min
         Do
            dnContador = dnContador + 1
            dsTexto = Trim(Str(dnContador)) & "|"
            dsTexto = dsTexto & "6|" & porstTGEMP!RUCEmp & "|"
            dsTexto = dsTexto & gsAnoAct & "|"
            dsTexto = dsTexto & IIf(!TpoPer = TPOPER_JUR, "02", "01") & "|"
            dsTexto = dsTexto & "6|" & Trim(!rucaux) & "|"
            dsTexto = dsTexto & Trim(Str(gfRedond(!Total, 0))) & "|"
            dsTexto = dsTexto & Trim(!ApePatAux) & "|"
            dsTexto = dsTexto & Trim(!ApeMatAux) & "|"
            If Not IsNull(!NomAux) Then
              dnCaracter = InStr(1, Trim(!NomAux), " ")
              dnCaracter = IIf(dnCaracter <> 0, dnCaracter - 1, Len(Trim(!NomAux)))
              dsTexto = dsTexto & Mid(Trim(!NomAux), 1, dnCaracter) & "|"
              dnCaracter = InStr(1, Trim(!NomAux), " ")
              If dnCaracter <> 0 Then
                dnCaracter = IIf(dnCaracter <> 0, dnCaracter + 1, dnCaracter)
                dsTexto = dsTexto & Mid(Trim(!NomAux), dnCaracter) & "|"
              Else
                dsTexto = dsTexto & "|"
              End If
            Else
              dsTexto = dsTexto & "||"
            End If
            dsTexto = dsTexto & Trim(!razAux) & "|"
            Print #2, dsTexto
            PgBEtapa2.Value = dnContador
            .MoveNext
         Loop Until .EOF
      End With
      Exit Do
   Loop
   Close #2
   porstTGEMP.Close
End Sub

Private Sub Form_Load()
  
 '[Parámetros.                         'Cambiar.
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :")
  Next nElemento
  LblProces(0).Caption = Choose(gsIdioma, "Procesando Compras", "Processing Purchases")
  LblProces(1).Caption = Choose(gsIdioma, "Procesando Ventas", "Processing Sales")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, True, False, False, False, False, False, True, aLabel
 ']
  
  'Características de impresión.
  udFecha = Date                      'Fecha en el encabezado.
  unCopias = 1                        'Cantidad de Copias.
  unMargenIzquierdo = 240             'Margen izquierdo.
  usDEstino = PRN_DEST_GRAF           'PRN_DEST_GRAF:ica _
                                       PRN_DEST_MATR:icial.
  usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                       PRN_ORIE_HORI:zontal.

End Sub
