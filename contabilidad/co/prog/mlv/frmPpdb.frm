VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPPDB 
   Caption         =   "[título]"
   ClientHeight    =   4425
   ClientLeft      =   2640
   ClientTop       =   3960
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkGeneral 
      Caption         =   "Todas"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   0
      Left            =   3255
      TabIndex        =   6
      Top             =   1215
      Value           =   1  'Checked
      Width           =   1080
   End
   Begin VB.CheckBox chkGeneral 
      Caption         =   "Todas"
      ForeColor       =   &H00800000&
      Height          =   200
      Index           =   1
      Left            =   3255
      TabIndex        =   13
      Top             =   2505
      Value           =   1  'Checked
      Width           =   1080
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "Comrpobante DUA Ventas"
      Height          =   200
      Index           =   4
      Left            =   150
      TabIndex        =   12
      Top             =   3120
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   3075
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   165
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "Tipo de cambio"
      Height          =   200
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   580
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "Comprobantes de Compras"
      Height          =   200
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   1215
      Value           =   1  'Checked
      Width           =   2800
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "Forma de pago Compras"
      Height          =   200
      Index           =   2
      Left            =   150
      TabIndex        =   8
      Top             =   1850
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "Comprobantes de Ventas"
      Height          =   200
      Index           =   3
      Left            =   150
      TabIndex        =   10
      Top             =   2505
      Value           =   1  'Checked
      Width           =   2800
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmPpdb.frx":0000
      Left            =   3225
      List            =   "frmPpdb.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
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
      Left            =   1560
      Picture         =   "frmPpdb.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Reporte de validación"
      Top             =   3840
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   1470
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   2090
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   11
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   4
      Left            =   150
      TabIndex        =   14
      Top             =   3370
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   830
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CdlUbicacion 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   2370
      TabIndex        =   2
      Top             =   165
      Width           =   765
   End
End
Attribute VB_Name = "frmPPDB"
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
Public porstProcesa As ADODB.Recordset
Public pbNuevo As Boolean
Public pcNroCpb As String

Private Sub chkProceso_Click(Index As Integer)
  If (Index = 1 And chkProceso(Index).Value = vbUnchecked) Then
    chkGeneral(0).Value = vbUnchecked
  ElseIf (Index = 3 And chkProceso(Index).Value = vbUnchecked) Then
    chkGeneral(1).Value = vbUnchecked
  End If
  chkGeneral(0).Enabled = (chkProceso(1).Value = vbChecked)
  chkGeneral(1).Enabled = (chkProceso(3).Value = vbChecked)
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
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
  sSentencia = sSentencia & "(a.ImpExo_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpExo, "
  sSentencia = sSentencia & "(a.ImpIGV_OGr_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGr, "
  sSentencia = sSentencia & "(a.ImpIGV_OGN_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGN, "
  sSentencia = sSentencia & "(a.ImpIGV_ONG_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVONG, "
  sSentencia = sSentencia & "(a.ImpISC_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpISC, "
  sSentencia = sSentencia & "(a.ImpOIm_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOIm, "
  sSentencia = sSentencia & "(a.ImpTot_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot, b.CodTDc "
  sSentencia = sSentencia & "FROM ((COCprDoc a "
  sSentencia = sSentencia & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
  sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<='" & gsMesAct & "' "
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
  sSentencia = sSentencia & "(a.ImpIGV_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmIgv, "
  sSentencia = sSentencia & "(a.ImpISC_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmISC, "
  sSentencia = sSentencia & "(a.ImpOIm_" & sMoneda & " * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmOlm, "
  sSentencia = sSentencia & "(a.ImpTot_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmTot, "
  sSentencia = sSentencia & "b.CodTDc "
  sSentencia = sSentencia & "FROM ((COVtaDoc a "
  sSentencia = sSentencia & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
  sSentencia = sSentencia & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.Mespvs<='" & gsMesAct & "' "
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
   chkProceso(0).Enabled = True
   chkProceso(1).Enabled = True
   chkProceso(2).Enabled = True
   chkProceso(3).Enabled = True
   chkProceso(4).Enabled = True
   cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
'  On Error GoTo Err
  
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  cmdImprimir(0).Enabled = False
  pgbProceso(0).Value = 0: pgbProceso(0).Min = 0
  pgbProceso(1).Value = 0: pgbProceso(1).Min = 0
  pgbProceso(2).Value = 0: pgbProceso(2).Min = 0
  pgbProceso(3).Value = 0: pgbProceso(3).Min = 0
  pgbProceso(4).Value = 0: pgbProceso(4).Min = 0
  
  'Abrir Tablas.
  Set pocnnMain = New ADODB.Connection
  Set porstProcesa = New ADODB.Recordset
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  With porstProcesa
    .ActiveConnection = pocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  
  pocnnMain.BeginTrans                'INICIA TRANSACCION.
  'Paso 1: Tipos de cambio
  If chkProceso(0).Value Then ppTipoCambio
  'Paso 2: Comprobantes de compras
  If chkProceso(1).Value Then ppCpbCompras
  'Paso 3: Forma de pago
  If chkProceso(2).Value Then ppCpbPagoCompras
  'Paso 4: Comprobantes de venta
  If chkProceso(3).Value Then ppCpbVentas
  'Paso 5: Comprobantes de DUA
  If chkProceso(4).Value Then ppCpbDuaVentas
  pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
  
  MsgBox TEXT_8008, vbInformation
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  cmdImprimir(0).Enabled = True
  cmdSalir.SetFocus
  
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  pocnnMain.Close
  Set porstProcesa = Nothing
  Set pocnnMain = Nothing
   
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub ppCpbCompras()
  Dim sArchivo As String, sCadena As String, sCadenaIni As String
  Dim sCaracter As String, sMoneda As String, sRegistro As String
  Dim nLongitud As Integer, nDestino As Integer
  Dim nInicio As Integer, nFinal As Integer, nSecuencia As Integer
  Dim nImporte As Double, nProgreso As Long
   
  ' Inicializo variables y nombre de archivo
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sArchivo = "c" & gsRUCEmp & gsAnoAct & gsMesAct & ".txt"
  sCaracter = "|"
  cdlUbicacion.FileName = sArchivo
  cdlUbicacion.ShowSave
  Open sArchivo For Output As #1
   
  With porstProcesa
    If .State = adStateOpen Then .Close
    .Source = "SELECT det.codtdc, det.feedoc, det.serdoc, det.nrodoc, aux.tpoper, aux.tpodci, aux.rucaux, "
    .Source = .Source & "aux.razaux, nat.apepataux, nat.apemataux, nat.nomaux, nat.numdci, det.tpomon, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impogr_mn ELSE det.impogr_me END) AS impogr, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impogn_mn ELSE det.impogn_me END) AS impogn, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impong_mn ELSE det.impong_me END) AS impong, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impexo_mn ELSE det.impexo_me END) AS impexo, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impisc_mn ELSE det.impisc_me END) AS impisc, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impigv_ogr_mn ELSE det.impigv_ogr_me END) AS impigv1, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impigv_ogn_mn ELSE det.impigv_ogn_me END) AS impigv2, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impigv_ong_mn ELSE det.impigv_ong_me END) AS impigv3, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impoim_mn ELSE det.impoim_me END) AS impoim, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impigv_mn ELSE det.impigv_me END) AS impigv, "
    .Source = .Source & "det.indcdt, det.nrocdt, "
    .Source = .Source & "det.indcprext, det.codaduana, det.annodua, det.nrodua, det.indreten, det.tsadetrac, "
    .Source = .Source & "det.codtdc_ref, det.serdoc_ref, det.nrodoc_ref, det.feedoc_ref, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impbasref_mn ELSE det.impbasref_me END) AS impbasref, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impigvref_mn ELSE det.impigvref_me END) AS impigvref "
    .Source = .Source & "FROM cocprdoc det "
    .Source = .Source & "INNER JOIN tgaux aux ON det.codemp=aux.codemp AND det.codaux=aux.codaux "
    .Source = .Source & "LEFT JOIN tgauxnat nat ON det.codemp=nat.codemp AND det.codaux=nat.codaux "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
    .Source = .Source & "ORDER BY det.feedoc, det.codaux"
    .Open
  End With
  ' Verifico si existe registros
  If porstProcesa.RecordCount > 0 Then
    porstProcesa.MoveFirst
    pgbProceso(1).Max = porstProcesa.RecordCount
    pgbProceso(1).Value = pgbProceso(1).Min
    nProgreso = 0
    Do While Not porstProcesa.EOF
      nInicio = 0: nFinal = 0
      If CDec(porstProcesa!impogr) <> 0 Then nInicio = 1: nFinal = 1
      If CDec(porstProcesa!impogn) <> 0 Then nInicio = IIf(nInicio = 0, 2, nInicio): nFinal = 2
      If CDec(porstProcesa!impong) <> 0 Then nInicio = IIf(nInicio = 0, 3, nInicio): nFinal = 3
      If chkGeneral(0).Value = vbChecked Then
        If CDec(porstProcesa!impexo) <> 0 Then nInicio = IIf(nInicio = 0, 4, nInicio): nFinal = 4
      End If
      If nInicio <> 0 Then
        ' Cadena inicial
        sCadenaIni = ""
        sRegistro = Trim(IIf(IsNull(porstProcesa!indcprext), "0", porstProcesa!indcprext))
        sCadenaIni = sCadenaIni & Format(Val(sRegistro) + 1, "00") & sCaracter                                    ' Tipo de compra
        sCadenaIni = sCadenaIni & Trim(porstProcesa!CodTDc) & sCaracter                                           ' Tipo de comprobante
        sCadenaIni = sCadenaIni & Format(porstProcesa!feedoc, "dd/mm/yyyy") & sCaracter                           ' Fecha de emisión / pago
        sRegistro = Trim(IIf(porstProcesa!CodTDc = "10" Or porstProcesa!CodTDc = "12", "", porstProcesa!SerDoc))
        If porstProcesa!CodTDc >= "52" And porstProcesa!CodTDc <= "55" Then
          sRegistro = Trim(IIf(IsNull(porstProcesa!codaduana), "", porstProcesa!codaduana))
          sRegistro = sRegistro & Trim(IIf(IsNull(porstProcesa!annodua), "", porstProcesa!annodua))
          sRegistro = sRegistro & Trim(IIf(IsNull(porstProcesa!nrodua), "", porstProcesa!nrodua))
        End If
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                                           ' Serie del comprobante de pago
        sRegistro = ""
        If Not (porstProcesa!CodTDc >= "52" And porstProcesa!CodTDc <= "55") Then
          sRegistro = Trim(porstProcesa!NroDoc)
        End If
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                                           ' Numero de comprobante
        sRegistro = IIf(porstProcesa!TpoPer = TPOPER_NAT, "01", IIf(porstProcesa!TpoPer = TPOPER_JUR, "02", "03"))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                                           ' Tipo de persona
        sRegistro = Right(Trim(porstProcesa!TpoDci), 1)
        sRegistro = IIf(sRegistro = "0", "-", IIf((sRegistro <= "7" Or sRegistro = "A"), sRegistro, ""))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                                           ' Tipo de documento identidad
        ' Numero de documento
        sRegistro = ""
        If porstProcesa!TpoPer <> TPOPER_DOM Then
          sRegistro = Trim(porstProcesa!rucaux)
          If Right(Trim(porstProcesa!TpoDci), 1) = "1" Then
            If Not IsNull(porstProcesa!numdci) Then
              sRegistro = Trim(porstProcesa!numdci)
            End If
            sRegistro = Right(sRegistro, 8)
          End If
        End If
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                            ' Numero de documento identidad
        sCadenaIni = sCadenaIni & Trim(porstProcesa!razaux) & sCaracter                                           ' Nombre o razon social
        sRegistro = Trim(IIf(IsNull(porstProcesa!ApePatAux), "", porstProcesa!ApePatAux))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                                           ' Apellido paterno
        sRegistro = Trim(IIf(IsNull(porstProcesa!ApeMatAux), "", porstProcesa!ApeMatAux))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                                           ' Apellido materno
        sRegistro = Trim(IIf(IsNull(porstProcesa!NomAux), "", porstProcesa!NomAux))
        If sRegistro <> "" Then
          nLongitud = InStr(1, sRegistro, " ")
          nLongitud = IIf(nLongitud <> 0, nLongitud - 1, Len(sRegistro))
          sCadenaIni = sCadenaIni & Mid(sRegistro, 1, nLongitud) & sCaracter                                      ' Primer nombre
          nLongitud = InStr(1, sRegistro, " ")
          If nLongitud <> 0 Then
            nLongitud = IIf(nLongitud <> 0, nLongitud + 1, nLongitud)
            sCadenaIni = sCadenaIni & Mid(sRegistro, nLongitud) & sCaracter                                       ' Segundo nombre
          Else
            sCadenaIni = sCadenaIni & sCaracter                                                                   ' Segundo nombre
          End If
        Else
            sCadenaIni = sCadenaIni & sCaracter & sCaracter                                                       ' Primer y segundo nombre
        End If
        sRegistro = IIf(porstProcesa!tpomon = TPOMON_NAC, "1", "2")
        sCadenaIni = sCadenaIni & sRegistro & sCaracter                                                           ' Tipo de moneda
        nSecuencia = 0
        For nDestino = nInicio To nFinal
          nImporte = CDec(porstProcesa(Choose(nDestino, "impogr", "impogn", "impong", "impexo")))
          ' Genero la cadena si importe es diferente de cero
          If nImporte <> 0 Then
            ' Inicializo la cadena del documento
            sCadena = sCadenaIni
            sRegistro = Trim(IIf(nInicio <> nFinal, "5", nDestino))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Codigo de destino
            nSecuencia = nSecuencia + 1
            sCadena = sCadena & Trim(IIf(nInicio <> nFinal, nDestino, nSecuencia)) & sCaracter                    ' Numero secuencia destino
            sCadena = sCadena & Format(nImporte, "#0.00") & sCaracter                                             ' Base imponible
            
            Select Case Trim(porstProcesa!CodTDc)
             Case "03"
              sCadena = sCadena & "" & sCaracter                                                                  ' Monto ISC - caso 1
             Case "04"
              sCadena = sCadena & "" & sCaracter                                                                  ' Monto ISC - caso 2
             Case "05"
              sCadena = sCadena & "" & sCaracter                                                                  ' Monto ISC - caso 3
             Case "10"
              sCadena = sCadena & "" & sCaracter                                                                  ' Monto ISC - caso 4
             Case Else
              sCadena = sCadena & Format(CDec(porstProcesa!impisc), "#0.00") & sCaracter                          ' Monto ISC - caso 5
            End Select
                        
            Select Case nDestino
             Case 1
              sCadena = sCadena & Format(CDec(porstProcesa!impigv1), "#0.00") & sCaracter                         ' Monto IGV - caso 1
             Case 2
              sCadena = sCadena & Format(CDec(porstProcesa!impigv2), "#0.00") & sCaracter                         ' Monto IGV - caso 2
             Case 3
              sCadena = sCadena & Format(CDec(porstProcesa!impigv3), "#0.00") & sCaracter                         ' Monto IGV - caso 3
             Case Else
              sCadena = sCadena & Format(CDec(Val("0.00")), "#0.00") & sCaracter                                  ' Monto IGV - caso 4
            End Select
            
            sCadena = sCadena & Format(IIf(nSecuencia = 1, CDec(porstProcesa!impoim), 0), "#0.00") & sCaracter    ' Otros importes
            ' Detraccion
            sRegistro = Trim(IIf(IsNull(porstProcesa!indcdt), "0", porstProcesa!indcdt))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Indicador de detracciones
            sRegistro = Trim(IIf(IsNull(porstProcesa!tsadetrac), "", porstProcesa!tsadetrac))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Codigo tasa detracción
            sRegistro = Trim(IIf(IsNull(porstProcesa!NroCDt), "", porstProcesa!NroCDt))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Numero constancia detracción
            ' retencion
            sRegistro = Trim(IIf(IsNull(porstProcesa!indreten), "0", porstProcesa!indreten))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Indicador retención
            ' Referencia
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Tipo documento referencia
            sRegistro = Trim(IIf(IsNull(porstProcesa!serdoc_ref), "", porstProcesa!serdoc_ref))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Serie documento referencia
            sRegistro = Trim(IIf(IsNull(porstProcesa!nrodoc_ref), "", porstProcesa!nrodoc_ref))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Numero documento referencia
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref))
            sRegistro = IIf(sRegistro = "", sRegistro, IIf(IsNull(porstProcesa!feedoc_ref), "", porstProcesa!feedoc_ref))
            sCadena = sCadena & Format(sRegistro, "dd/mm/yyyy") & sCaracter                                       ' Fecha emisión referencia
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref))
            sRegistro = IIf(sRegistro = "", sRegistro, Format(CDec(IIf(IsNull(porstProcesa!impbasref), 0, porstProcesa!impbasref)), "#0.00"))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Base imponible referencia
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref))
            sRegistro = IIf(sRegistro = "", sRegistro, Format(CDec(IIf(IsNull(porstProcesa!impigvref), 0, porstProcesa!impigvref)), "#0.00"))
            sCadena = sCadena & sRegistro & sCaracter                                                             ' Igv referencia
            Print #1, sCadena
          End If
        Next nDestino
      End If
      nProgreso = nProgreso + 1
      pgbProceso(1).Value = nProgreso
      porstProcesa.MoveNext
    Loop
  End If
  Close #1
  porstProcesa.Close
   
End Sub

Private Sub ppCpbDuaVentas()
  Dim sArchivo As String, sCadena As String, sCadenaIni As String
  Dim sCaracter As String, sMoneda As String, sRegistro As String
  Dim nImporte As Double, nProgreso As Long
   
  ' Inicializo variables y nombre de archivo
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sArchivo = gsRUCEmp & gsAnoAct & gsMesAct & ".dua"
  sCaracter = "|"
  cdlUbicacion.FileName = sArchivo
  cdlUbicacion.ShowSave
  Open sArchivo For Output As #1
   
  With porstProcesa
    If .State = adStateOpen Then .Close
    .Source = "SELECT det.codtdc, det.feedoc, det.serdoc, det.nrodoc, "
    .Source = .Source & "det.codaduana, det.annodua, det.nrodua, det.feembarq, det.feregula,  "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impfob_mn ELSE det.impfob_me END) AS impfob "
    .Source = .Source & "FROM covtadoc det "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
    .Source = .Source & "AND det.indvtaext='" & INDANU_VER & "' "
    .Source = .Source & "ORDER BY det.feembarq"
    .Open
  End With
  ' Verifico si existe registros
  If porstProcesa.RecordCount > 0 Then
    porstProcesa.MoveFirst
    pgbProceso(4).Max = porstProcesa.RecordCount
    pgbProceso(4).Value = pgbProceso(4).Min
    nProgreso = 0
    Do While Not porstProcesa.EOF
      nImporte = CDec(porstProcesa!impfob)
      ' Genero la cadena si importe es diferente de cero
      If nImporte <> 0 Then
        ' Inicializo la cadena
        sCadena = ""
        sRegistro = Trim(IIf(IsNull(porstProcesa!codaduana), "", porstProcesa!codaduana))
        sCadena = sCadena & sRegistro & sCaracter
        sRegistro = Trim(IIf(IsNull(porstProcesa!annodua), "", porstProcesa!annodua))
        sCadena = sCadena & sRegistro & sCaracter
        sRegistro = Trim(IIf(IsNull(porstProcesa!nrodua), "", porstProcesa!nrodua))
        sCadena = sCadena & sRegistro & sCaracter
        sRegistro = Trim(IIf(IsNull(porstProcesa!codaduana), "", porstProcesa!codaduana)) ' fecha embarque
        sRegistro = IIf(sRegistro = "", sRegistro, IIf(IsNull(porstProcesa!feembarq), "", porstProcesa!feembarq))
        sCadena = sCadena & Format(sRegistro, "dd/mm/yyyy") & sCaracter
        sRegistro = Trim(IIf(IsNull(porstProcesa!codaduana), "", porstProcesa!codaduana)) ' fecha regulariza
        sRegistro = IIf(sRegistro = "", sRegistro, IIf(IsNull(porstProcesa!feregula), "", porstProcesa!feregula))
        sCadena = sCadena & Format(sRegistro, "dd/mm/yyyy") & sCaracter
        sCadena = sCadena & Format(nImporte, "#0.00") & sCaracter
        Print #1, sCadena
      End If
      nProgreso = nProgreso + 1
      pgbProceso(4).Value = nProgreso
      porstProcesa.MoveNext
    Loop
  End If
  Close #1
  porstProcesa.Close
   
End Sub

Private Sub ppCpbPagoCompras()
  Dim sArchivo As String, sCadena As String
  Dim sCaracter As String, sMoneda As String, sRegistro As String
  Dim nImporte As Double, nProgreso As Long
   
  ' Inicializo variables y nombre de archivo
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sArchivo = "f" & gsRUCEmp & gsAnoAct & gsMesAct & ".txt"
  sCaracter = "|"
  cdlUbicacion.FileName = sArchivo
  cdlUbicacion.ShowSave
  Open sArchivo For Output As #1
   
  With porstProcesa
    If .State = adStateOpen Then .Close
    .Source = "SELECT cpr.indcprext, det.codtdc, det.serdoc, det.nrodoc, cpr.codaduana, cpr.annodua, cpr.nrodua, "
    .Source = .Source & "aux.tpoper, aux.tpodci, aux.rucaux, det.tpomon, ban.codbco, ban.tpodoc, ban.docban, "
    .Source = .Source & "det.fehope, bco.codent, "
    '.Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impmn ELSE det.impme END) AS impope "
    
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN (cpr.impogr_mn + cpr.impogn_mn + cpr.impong_mn + cpr.impigv_mn) ELSE (cpr.impogr_me + cpr.impogn_me + cpr.impong_me + cpr.impigv_me) END) AS impope "
    
    
    .Source = .Source & "FROM cocpbdet det "
    .Source = .Source & "INNER JOIN cocprdoc cpr ON det.codemp=cpr.codemp AND det.codaux=cpr.codaux AND det.codtdc=cpr.codtdc AND det.serdoc=cpr.serdoc AND det.nrodoc=cpr.nrodoc "
    .Source = .Source & "INNER JOIN tgaux aux ON det.codemp=aux.codemp AND det.codaux=aux.codaux "
    .Source = .Source & "LEFT JOIN cobancab ban ON det.codemp=ban.codemp AND det.pdoano=ban.pdoano AND det.mespvs=ban.mespvs AND det.coddro=ban.coddro AND det.nrocpb=ban.nroban "
    .Source = .Source & "LEFT JOIN cobco bco ON ban.codemp=bco.codemp AND bco.codbco=ban.codbco "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
    .Source = .Source & "AND det.tpopvs='" & TPOPVS_CAN & "' "
    .Source = .Source & "ORDER BY det.fehope"
    .Open
  End With
  ' Verifico si existe registros
  If porstProcesa.RecordCount > 0 Then
    porstProcesa.MoveFirst
    pgbProceso(2).Max = porstProcesa.RecordCount
    pgbProceso(2).Value = pgbProceso(2).Min
    nProgreso = 0
    Do While Not porstProcesa.EOF
      nImporte = CDec(porstProcesa!impope)
      ' Genero la cadena si importe es diferente de cero
      If nImporte <> 0 Then
        ' Inicializo la cadena
        sCadena = ""
        sRegistro = Trim(IIf(IsNull(porstProcesa!indcprext), "0", porstProcesa!indcprext))
        sCadena = sCadena & Format(Val(sRegistro) + 1, "00") & sCaracter
        sCadena = sCadena & Trim(porstProcesa!CodTDc) & sCaracter
        sRegistro = Trim(IIf(porstProcesa!CodTDc = "10" Or porstProcesa!CodTDc = "12", "", porstProcesa!SerDoc))
        If porstProcesa!CodTDc >= "52" And porstProcesa!CodTDc <= "55" Then
          sRegistro = Trim(IIf(IsNull(porstProcesa!codaduana), "", porstProcesa!codaduana))
          sRegistro = sRegistro & Trim(IIf(IsNull(porstProcesa!annodua), "", porstProcesa!annodua))
          sRegistro = sRegistro & Trim(IIf(IsNull(porstProcesa!nrodua), "", porstProcesa!nrodua))
        End If
        sCadena = sCadena & sRegistro & sCaracter
        sRegistro = ""
        If Not (porstProcesa!CodTDc >= "52" And porstProcesa!CodTDc <= "55") Then
          sRegistro = Trim(porstProcesa!NroDoc)
        End If
        sCadena = sCadena & sRegistro & sCaracter
        sRegistro = IIf(porstProcesa!TpoPer = TPOPER_JUR, "02", IIf(porstProcesa!TpoPer = TPOPER_NAT, "01", "03"))
        sCadena = sCadena & sRegistro & sCaracter
        sRegistro = Right(Trim(porstProcesa!TpoDci), 1)
        sRegistro = IIf(sRegistro = "0", "-", IIf((sRegistro <= "7" Or sRegistro = "A"), sRegistro, ""))
        sCadena = sCadena & sRegistro & sCaracter
        sCadena = sCadena & Trim(porstProcesa!rucaux) & sCaracter
        'sRegistro = Format(Trim(IIf(IsNull(porstProcesa!tpodoc), "8", porstProcesa!tpodoc)), "000")
        sRegistro = Format(Trim(IIf(IsNull(porstProcesa!tpodoc), "9", porstProcesa!tpodoc)), "000")
        sCadena = sCadena & sRegistro & sCaracter
        sMoneda = IIf(IsNull(porstProcesa!codent) Or porstProcesa!codent = "", 0, porstProcesa!codent)
        If (sMoneda <> 0 And Not (sRegistro = "009" Or sRegistro = "011" Or sRegistro = "013" Or sRegistro = "014" Or sRegistro = "098")) Then
        '  sRegistro = Format(Choose(porstProcesa!codent, "2", "3", "7", "8", "9", "11", "18", "23", "26", "29", "35", "37", "38", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "99"), "00")
          sRegistro = Format(porstProcesa!codent, "00")
        Else
          sRegistro = ""
        End If
        sCadena = sCadena & sRegistro & sCaracter
        
        'sRegistro = Trim(IIf(IsNull(porstProcesa!docban), "", porstProcesa!docban))
        'sCadena = sCadena & sRegistro & sCaracter
        
        If IsNull(porstProcesa!docban) Or porstProcesa!tpodoc = "009" Then
           sCadena = sCadena & "" & sCaracter
        Else
           sRegistro = Trim(porstProcesa!docban)
           sCadena = sCadena & sRegistro & sCaracter
        End If
        
        
        If IsNull(porstProcesa!tpodoc) Or porstProcesa!tpodoc = "009" Then
           sCadena = sCadena & "" & sCaracter
        Else
           sCadena = sCadena & Format(porstProcesa!fehope, "dd/mm/yyyy") & sCaracter
        End If
        
        sCadena = sCadena & Format(nImporte, "#0.00") & sCaracter
        Print #1, sCadena
      End If
      nProgreso = nProgreso + 1
      pgbProceso(2).Value = nProgreso
      porstProcesa.MoveNext
    Loop
  End If
  Close #1
  porstProcesa.Close

End Sub

Private Sub ppCpbVentas()
  Dim sArchivo As String, sCadena As String, sCadenaIni As String
  Dim sCaracter As String, sMoneda As String, sRegistro As String
  Dim nLongitud As Integer, nDestino As Integer
  Dim nInicio As Integer, nFinal As Integer, nSecuencia As Integer
  Dim nImporte As Double, nProgreso As Long
   
  ' Inicializo variables y nombre de archivo
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  sArchivo = "v" & gsRUCEmp & gsAnoAct & gsMesAct & ".txt"
  sCaracter = "|"
  cdlUbicacion.FileName = sArchivo
  cdlUbicacion.ShowSave
  Open sArchivo For Output As #1
   
  With porstProcesa
    If .State = adStateOpen Then .Close
    .Source = "SELECT det.codtdc, det.feedoc, det.serdoc, det.nrodoc, aux.tpoper, aux.tpodci, aux.rucaux, "
    .Source = .Source & "aux.razaux, nat.apepataux, nat.apemataux, nat.nomaux, nat.numdci, det.tpomon,  "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impogr_mn ELSE det.impogr_me END) AS impogr, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impexp_mn ELSE det.impexp_me END) AS impexp, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impexo_mn ELSE det.impexo_me END) AS impexo, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impisc_mn ELSE det.impisc_me END) AS impisc, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impigv_mn ELSE det.impigv_me END) AS impigv, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impoim_mn ELSE det.impoim_me END) AS impoim, "
    .Source = .Source & "det.indvtaext, det.indpercep, det.tsapercep, det.serpercep, det.nropercep,  "
    .Source = .Source & "det.codtdc_ref, det.serdoc_ref, det.nrodoc_ref, det.feedoc_ref, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impbasref_mn ELSE det.impbasref_me END) AS impbasref, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impigvref_mn ELSE det.impigvref_me END) AS impigvref "
    .Source = .Source & "FROM covtadoc det "
    .Source = .Source & "INNER JOIN tgaux aux ON det.codemp=aux.codemp AND det.codaux=aux.codaux "
    .Source = .Source & "LEFT JOIN tgauxnat nat ON det.codemp=nat.codemp AND det.codaux=nat.codaux "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
    If chkGeneral(1).Value = vbUnchecked Then
      .Source = .Source & "AND det.indvtaext='" & INDANU_VER & "' "
    End If
    .Source = .Source & "ORDER BY det.feedoc, det.codaux"
    .Open
  End With
  ' Verifico si existe registros
  If porstProcesa.RecordCount > 0 Then
    porstProcesa.MoveFirst
    pgbProceso(3).Max = porstProcesa.RecordCount
    pgbProceso(3).Value = pgbProceso(3).Min
    nProgreso = 0
    Do While Not porstProcesa.EOF
      nInicio = 0: nFinal = 0
      If CDec(porstProcesa!impogr) <> 0 Then nInicio = 1: nFinal = 1
      If CDec(porstProcesa!impexp) <> 0 Then nInicio = IIf(nInicio = 0, 2, nInicio): nFinal = 2
      If CDec(porstProcesa!impexo) <> 0 Then nInicio = IIf(nInicio = 0, 3, nInicio): nFinal = 3
      If nInicio <> 0 Then
        ' Cadena inicial
        sCadenaIni = ""
        sRegistro = Trim(IIf(IsNull(porstProcesa!indvtaext), "0", porstProcesa!indvtaext))
        sCadenaIni = sCadenaIni & Format(Val(sRegistro) + 1, "00") & sCaracter
        sCadenaIni = sCadenaIni & Trim(porstProcesa!CodTDc) & sCaracter
        sCadenaIni = sCadenaIni & Format(porstProcesa!feedoc, "dd/mm/yyyy") & sCaracter
        sCadenaIni = sCadenaIni & Trim(porstProcesa!SerDoc) & sCaracter
        sCadenaIni = sCadenaIni & Trim(porstProcesa!NroDoc) & sCaracter
        sRegistro = IIf(porstProcesa!TpoPer = TPOPER_NAT, "01", IIf(porstProcesa!TpoPer = TPOPER_JUR, "02", "03"))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter
        sRegistro = Right(Trim(porstProcesa!TpoDci), 1)
        sRegistro = IIf(sRegistro = "0", "-", IIf((sRegistro <= "7" Or sRegistro = "A"), sRegistro, ""))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter
        ' Numero de documento
        sRegistro = ""
        If porstProcesa!TpoPer <> TPOPER_DOM Then
          sRegistro = Trim(porstProcesa!rucaux)
          If Right(Trim(porstProcesa!TpoDci), 1) = "1" Then
            If Not IsNull(porstProcesa!numdci) Then
              sRegistro = Trim(porstProcesa!numdci)
            End If
            sRegistro = Right(sRegistro, 8)
          End If
        End If
        
        sCadenaIni = sCadenaIni & sRegistro & sCaracter
        sCadenaIni = sCadenaIni & Trim(porstProcesa!razaux) & sCaracter
        sRegistro = Trim(IIf(IsNull(porstProcesa!ApePatAux), "", porstProcesa!ApePatAux))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter
        sRegistro = Trim(IIf(IsNull(porstProcesa!ApeMatAux), "", porstProcesa!ApeMatAux))
        sCadenaIni = sCadenaIni & sRegistro & sCaracter
        sRegistro = Trim(IIf(IsNull(porstProcesa!NomAux), "", porstProcesa!NomAux))
        If sRegistro <> "" Then
          nLongitud = InStr(1, sRegistro, " ")
          nLongitud = IIf(nLongitud <> 0, nLongitud - 1, Len(sRegistro))
          sCadenaIni = sCadenaIni & Mid(sRegistro, 1, nLongitud) & sCaracter
          nLongitud = InStr(1, sRegistro, " ")
          If nLongitud <> 0 Then
            nLongitud = IIf(nLongitud <> 0, nLongitud + 1, nLongitud)
            sCadenaIni = sCadenaIni & Mid(sRegistro, nLongitud) & sCaracter
          Else
            sCadenaIni = sCadenaIni & sCaracter
          End If
        Else
            sCadenaIni = sCadenaIni & sCaracter & sCaracter
        End If
        sRegistro = IIf(porstProcesa!tpomon = TPOMON_NAC, "1", "2")
        sCadenaIni = sCadenaIni & sRegistro & sCaracter
        nSecuencia = 0
        For nDestino = nInicio To nFinal
          nImporte = CDec(porstProcesa(Choose(nDestino, "impogr", "impexp", "impexo")))
          ' Genero la cadena si importe es diferente de cero
          If nImporte <> 0 Then
            ' Inicializo la cadena del documento
            sCadena = sCadenaIni
            sRegistro = Trim(IIf(nInicio <> nFinal, "3", nDestino))
            sCadena = sCadena & sRegistro & sCaracter
            nSecuencia = nSecuencia + 1
            sCadena = sCadena & Trim(nSecuencia) & sCaracter
            sCadena = sCadena & Format(nImporte, "#0.00") & sCaracter
            sCadena = sCadena & Format(CDec(porstProcesa!impisc), "#0.00") & sCaracter
            Select Case nDestino
            Case 1
                 sCadena = sCadena & Format(CDec(porstProcesa!impigv), "#0.00") & sCaracter
            Case 2
                 sCadena = sCadena & Format(CDec(Val("0.00")), "#0.00") & sCaracter
            Case Else
                 sCadena = sCadena & Format(CDec(Val("0.00")), "#0.00") & sCaracter
            End Select
            
            'sCadena = sCadena & Format(CDec(porstProcesa!impigv), "#0.00") & sCaracter
            sCadena = sCadena & Format(CDec(porstProcesa!impoim), "#0.00") & sCaracter
            ' Percepcion
            sRegistro = Trim(IIf(IsNull(porstProcesa!indpercep), "0", porstProcesa!indpercep))
            sCadena = sCadena & sRegistro & sCaracter
            sRegistro = Format(Trim(IIf(IsNull(porstProcesa!tsapercep), "0", porstProcesa!tsapercep)), "00")
            sRegistro = IIf(sRegistro = "00", "", sRegistro)
            sCadena = sCadena & sRegistro & sCaracter
            sRegistro = Trim(IIf(IsNull(porstProcesa!serpercep), "", porstProcesa!serpercep))
            sCadena = sCadena & sRegistro & sCaracter
            sRegistro = Trim(IIf(IsNull(porstProcesa!nropercep), "", porstProcesa!nropercep))
            sCadena = sCadena & sRegistro & sCaracter
            ' Referencia
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref))
            sCadena = sCadena & sRegistro & sCaracter
            sRegistro = Trim(IIf(IsNull(porstProcesa!serdoc_ref), "", porstProcesa!serdoc_ref)) ' serie
            sCadena = sCadena & sRegistro & sCaracter
            sRegistro = Trim(IIf(IsNull(porstProcesa!nrodoc_ref), "", porstProcesa!nrodoc_ref)) ' nro
            sCadena = sCadena & sRegistro & sCaracter
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref)) ' fecha
            sRegistro = IIf(sRegistro = "", sRegistro, IIf(IsNull(porstProcesa!feedoc_ref), "", porstProcesa!feedoc_ref))
            sCadena = sCadena & Format(sRegistro, "dd/mm/yyyy") & sCaracter
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref)) ' base
            sRegistro = IIf(sRegistro = "", sRegistro, Format(CDec(IIf(IsNull(porstProcesa!impbasref), 0, porstProcesa!impbasref)), "#0.00"))
            sCadena = sCadena & sRegistro & sCaracter
            sRegistro = Trim(IIf(IsNull(porstProcesa!codtdc_ref), "", porstProcesa!codtdc_ref)) ' igv
            sRegistro = IIf(sRegistro = "", sRegistro, Format(CDec(IIf(IsNull(porstProcesa!impigvref), 0, porstProcesa!impigvref)), "#0.00"))
            sCadena = sCadena & sRegistro & sCaracter
            Print #1, sCadena
          End If
        Next nDestino
      End If
      nProgreso = nProgreso + 1
      pgbProceso(3).Value = nProgreso
      porstProcesa.MoveNext
    Loop
  End If
  Close #1
  porstProcesa.Close
   
End Sub

Private Sub ppTipoCambio()
  Dim sArchivo As String, sCaracter As String, sCadena As String
  Dim nProgreso As Long
   
  With porstProcesa
    If .State = adStateOpen Then .Close
    .Source = "SELECT fehtcb, imptcb_cpr, imptcb_vta FROM tgtcb "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    If ps_Plataforma = pSrvMySql Then
    .Source = .Source & "AND DATE_FORMAT(fehtcb, '%Y-%m')='" & gsAnoAct & "-" & gsMesAct & "' "
    ElseIf ps_Plataforma = pSrvSql Then
    .Source = .Source & "AND CONVERT(smalldatetime, fehtcb, 103)='" & gsAnoAct & "-" & gsMesAct & "' "
    End If
    .Source = .Source & "ORDER BY fehtcb"
    .Open
  End With
  sArchivo = gsRUCEmp & ".tc"
  sCaracter = "|"
  cdlUbicacion.FileName = sArchivo
  cdlUbicacion.ShowSave
  Open sArchivo For Output As #1
   
  If porstProcesa.RecordCount > 0 Then
    porstProcesa.MoveFirst
    pgbProceso(0).Max = porstProcesa.RecordCount
    pgbProceso(0).Value = pgbProceso(0).Min
    nProgreso = 0
    Do While Not porstProcesa.EOF
      sCadena = ""
      sCadena = sCadena & Format(porstProcesa!FehTCb, "dd/mm/yyyy") & sCaracter
      sCadena = sCadena & Format(CDec(porstProcesa!ImpTCb_Cpr), "#0.000") & sCaracter
      sCadena = sCadena & Format(CDec(porstProcesa!ImpTCb_Vta), "#0.000") & sCaracter
      Print #1, sCadena
      nProgreso = nProgreso + 1
      pgbProceso(0).Value = nProgreso
      porstProcesa.MoveNext
    Loop
  End If
  Close #1
  porstProcesa.Close
   
End Sub

'Private Sub ppEtapa_02()   ' Generacion de Texto en File Ingresos
'   Dim dnContador As Integer, dnCaracter As Integer
'   Dim sCadena , dsFile As String
'
'   dnContador = 0
'   PgBEtapa2.Min = 0
'   With porstTGEMP
'      .Source = "Select RucEmp From TGEMP Where CodEmp='" & gsCodEmp & "'"
'      .Open
'   End With
'   'Open "C:\Owl-paqu\Angel.TXT" For Output As #1
'   dsFile = "Ingresos.TXT"
'   CdlUbicacion.FileName = dsFile
'   CdlUbicacion.ShowSave
'   Open dsFile For Output As #2
'   Do
'      With porstCOVtaDoc
'         If .RecordCount = 0 Then
'            Exit Do
'         End If
'         .MoveFirst
'         PgBEtapa2.Max = .RecordCount
'         PgBEtapa2.Value = PgBEtapa2.Min
'         Do
'            dnContador = dnContador + 1
'            sCadena  = Trim(Str(dnContador)) & "|"
'            sCadena  = sCadena  & "6|" & porstTGEMP!RUCEmp & "|"
'            sCadena  = sCadena  & gsAnoAct & "|"
'            sCadena  = sCadena  & IIf(!TpoPer = TPOPER_JUR, "02", "01") & "|"
'            sCadena  = sCadena  & "6|" & Trim(!RucAux) & "|"
'            sCadena  = sCadena  & Trim(Str(gfRedond(!Total, 0))) & "|"
'            sCadena  = sCadena  & Trim(!ApePatAux) & "|"
'            sCadena  = sCadena  & Trim(!ApeMatAux) & "|"
'            If Not IsNull(!NomAux) Then
'              dnCaracter = InStr(1, Trim(!NomAux), " ")
'              dnCaracter = IIf(dnCaracter <> 0, dnCaracter - 1, Len(Trim(!NomAux)))
'              sCadena  = sCadena  & Mid(Trim(!NomAux), 1, dnCaracter) & "|"
'              dnCaracter = InStr(1, Trim(!NomAux), " ")
'              If dnCaracter <> 0 Then
'                dnCaracter = IIf(dnCaracter <> 0, dnCaracter + 1, dnCaracter)
'                sCadena  = sCadena  & Mid(Trim(!NomAux), dnCaracter) & "|"
'              Else
'                sCadena  = sCadena  & "|"
'              End If
'            Else
'              sCadena  = sCadena  & "||"
'            End If
'            sCadena  = sCadena  & Trim(!RazAux) & "|"
'            Print #2, sCadena
'            PgBEtapa2.Value = dnContador
'            .MoveNext
'         Loop Until .EOF
'      End With
'      Exit Do
'   Loop
'   Close #2
'   porstTGEMP.Close
'End Sub

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
  chkProceso(0).Caption = Choose(gsIdioma, "Tipo de Cambio", "Rate of Exchange")
  chkProceso(1).Caption = Choose(gsIdioma, "Comprobantes de Compras", "Vouchers of Purchases")
  chkProceso(2).Caption = Choose(gsIdioma, "Forma de Pago de Compras", "Mode of Payment Purchases")
  chkProceso(3).Caption = Choose(gsIdioma, "Comprobantes de Ventas", "Vouchers of Sales")
  chkProceso(4).Caption = Choose(gsIdioma, "Comprobantes de DUA Ventas", "Vouchers of DUA Sales")
  chkGeneral(0).Caption = Choose(gsIdioma, "Todos", "All")
  chkGeneral(1).Caption = Choose(gsIdioma, "Todos", "All")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, True, False, False, False, False, False, True, aLabel
 ']
 chkGeneral(0).Value = vbUnchecked
 chkGeneral(1).Value = vbUnchecked
  
  'Características de impresión.
  udFecha = Date                      'Fecha en el encabezado.
  unCopias = 1                        'Cantidad de Copias.
  unMargenIzquierdo = 240             'Margen izquierdo.
  usDEstino = PRN_DEST_GRAF           'PRN_DEST_GRAF:ica
  usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical

End Sub

