VERSION 5.00
Begin VB.Form frmRRegPcp 
   Caption         =   "[título]"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   12
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2535
      Width           =   7320
      Begin VB.CommandButton cmdConfig 
         Caption         =   "&Configuración de Impresora"
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
         Left            =   2355
         TabIndex        =   6
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
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
         Index           =   1
         Left            =   1245
         Picture         =   "frmRRegPcp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
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
         Left            =   0
         Picture         =   "frmRRegPcp.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
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
         Left            =   3720
         Picture         =   "frmRRegPcp.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   6060
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   900
      Width           =   1260
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7290
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6885
         Picture         =   "frmRRegPcp.frx":077E
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   325
         Width           =   255
      End
      Begin VB.Label lblDatoDeta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   3
         Top             =   315
         Width           =   5520
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   5250
      TabIndex        =   10
      Top             =   945
      Width           =   750
   End
End
Attribute VB_Name = "frmRRegPcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents MRViewer As MRViewerObject
Attribute MRViewer.VB_VarHelpID = -1

Public udFecha As Date
Public unCopias As Integer
Public unMargenIzquierdo As Integer
Public usDEstino As String
Public usOrientacionRpt As String
Public usOrientacionOri As String
Private paOpciones As Variant
Private pocnnMain As ADODB.Connection
Private porstMRp As ADODB.Recordset

'[Propio del formulario.
Private porstTGAux As ADODB.Recordset
Private pnSaldo As Double
Dim dnSaldo As Double
']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstTGAux = New ADODB.Recordset
   
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstMRp
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
   End With
   With porstTGAux
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodAux, RazAux "
      .Source = .Source & "FROM TGAux "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 0
         .Item(dnContador).DataField = "CodAux"
         .Item(dnContador).MaxLength = porstTGAux.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :")
  Next nElemento
  fraAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
   
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_1, 0
      .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND

 '[Datos predeterminados.              'Cambiar.
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
  
  'Otros.
   
  'Características de impresión.
   chkImpFecha.Value = vbChecked
   udFecha = Date                      'Fecha en el encabezado.
   unCopias = 1 'frmMain.rptMain.CopiesToPrinter  'Cantidad de Copias.
   unMargenIzquierdo = 240             'Margen izquierdo.
   usDEstino = PRN_DEST_MATR           'PRN_DEST_GRAF:ica _
                                        PRN_DEST_MATR:icial.
   usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                        PRN_ORIE_HORI:zontal.
 ']
   frmOPrnCfg.OrientacionPrn 0, Me
   frmOPrnCfg.lblOriPrn.Caption = Printer.Orientation
   
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub Form_Activate()
   'Orden: Vista Previa, Imprimir, Exportar.
   zaOpciones = Array(gbPms04, gbPms05, gbPms06)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   picOpciones.Width = Me.Width - 120
   cmdSalir.Left = picOpciones.Width - 1135
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Cambiar. Añadir recordsets.
   porstTGAux.Close
   pocnnMain.Close
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  ppHabilitacion False
  Static sSentencia As String, nRegistros As Double
  
  ' Tabla temporal de documentos por retencion
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 18)='#tmpCoCPbDetRPDoc_') DROP TABLE #tmpCoCPbDetRPDoc"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpCoCPbDetRPDoc", sSentencia)
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpCoCPbDetRPDoc ", "")
  sSentencia = sSentencia & "SELECT DISTINCT a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, "
  sSentencia = sSentencia & "SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0)) AS cImpDebe, "
  sSentencia = sSentencia & "SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0)) AS cImpHaber "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpCoCPbDetRPDoc ")
  sSentencia = sSentencia & "FROM (CoCPbDetRP a "
  sSentencia = sSentencia & "LEFT JOIN CoCPbDet b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND a.CodTDc_RtcPcp='" & gsCodTDc_Pcp & "' "
  sSentencia = sSentencia & "AND a.IndRtcPcp='S' "
  If Trim(txtDato(0).Text) <> "" Then
    sSentencia = sSentencia & "AND a.CodAux='" & Trim(txtDato(0).Text) & "' "
  End If
  sSentencia = sSentencia & "GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (cImpDebe <> 0.00) OR (cImpHaber <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (SUM(ISNULL((CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0))) <> 0.00 "
    sSentencia = sSentencia & "OR (SUM(ISNULL((CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0))) <> 0.00 "
  End If
  sSentencia = sSentencia & "ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
  pocnnMain.Execute sSentencia, nRegistros
  
  ' Tabla temporal de las percepciones
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 15)='#tmpCoCPbDetRP_') DROP TABLE #tmpCoCPbDetRP"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpCoCPbDetRP", sSentencia)
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TABLE IF NOT EXISTS tmpCoCPbDetRP ", "")
  sSentencia = sSentencia & "SELECT DISTINCT a.*, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0) AS cImpRetDeb, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0) AS cImpRetHab "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpCoCPbDetRP ")
  sSentencia = sSentencia & "FROM ((CoCPbDetRP a "
  sSentencia = sSentencia & "LEFT JOIN CoCPbDet b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte) "
  sSentencia = sSentencia & "INNER JOIN " & ps_Prefijo & "tmpCoCPbDetRPDoc c ON a.CodCta=c.CodCta AND a.CodAux=c.CodAux AND a.CodTDc=c.CodTDc AND a.SerDoc=c.SerDoc AND a.NroDoc=c.NroDoc) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND a.CodTDc_RtcPcp='" & gsCodTDc_Pcp & "' "
  sSentencia = sSentencia & "AND a.IndRtcPcp='S' "
  If Trim(txtDato(0).Text) <> "" Then
    sSentencia = sSentencia & "AND a.CodAux='" & Trim(txtDato(0).Text) & "' "
  End If
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (cImpRetDeb <> 0.00) OR (cImpRetHab <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "AND (ISNULL((CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0)) <> 0.00 "
    sSentencia = sSentencia & "OR (ISNULL((CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & "_RtcPcp ELSE 0 END), 0)) <> 0.00 "
  End If
  sSentencia = sSentencia & "ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
  pocnnMain.Execute sSentencia, nRegistros
  
  ' Generacion de la tabla del reporte
  sSentencia = "SELECT DISTINCT a.CodCta AS CodCta, a.CodAux AS CodAux, a.CodTDc AS CodTDc, a.SerDoc AS SerDoc, a.NroDoc AS NroDoc, "
  sSentencia = sSentencia & "a.MesPvs, a.FehOpe AS FehOpe, a.FeEDoc AS FeEDoc, a.TpoPvs, a.TpoGnr, "
  sSentencia = sSentencia & "(CASE a.TpoPvs WHEN '" & TPOPVS_PVS & "' THEN 'PROVISION' WHEN '" & TPOPVS_CAN & "' THEN 'CANCELACION' ELSE 'AJUSTE' END) AS cOperacion, "
  sSentencia = sSentencia & "(CASE a.TpoPvs WHEN '" & TPOPVS_CAN & "' THEN '  ' ELSE a.CodTDc END ) AS cTipoDoc, "
  sSentencia = sSentencia & "(CASE a.TpoPvs WHEN '" & TPOPVS_CAN & "' THEN a.RefDoc ELSE " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc, ' - ', a.NroDoc)", "(a.SerDoc+' - '+a.NroDoc)") & " END) AS cDocumento, "
  sSentencia = sSentencia & "(CASE a.TpoPvs WHEN '" & TPOPVS_PVS & "' THEN '0' WHEN '" & TPOPVS_CAN & "' THEN '4' ELSE '3' END) AS cOrden, c.RazAux, c.RUCAux, "
  sSentencia = sSentencia & "(CASE a.TpoPvs WHEN '" & TPOPVS_CAN & "' THEN 'CHEQUE/RECIBO' ELSE " & Choose(gsIdioma, "d.DetTDc", "d.DetTDcx") & " END) AS DetTDc, "
  sSentencia = sSentencia & "ROUND((" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & " ELSE 0 END), 0) + "
  sSentencia = sSentencia & "(CASE WHEN (a.TpoPvs='" & TPOPVS_CAN & "' AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte) THEN b.cImpRetHab ELSE 0 END)), 2) AS cDebe, "
  sSentencia = sSentencia & "ROUND((" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & " ELSE 0 END), 0) + "
  sSentencia = sSentencia & "(CASE WHEN (a.TpoPvs='" & TPOPVS_CAN & "' AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte) THEN b.cImpRetDeb ELSE 0 END)), 2) AS cHaber, "
  sSentencia = sSentencia & "ROUND(a.ImpMN - a.ImpMN, 2) AS cImpRetDeb, "
  sSentencia = sSentencia & "ROUND(a.ImpMN - a.ImpMN, 2) AS cImpRetHab, "
  sSentencia = sSentencia & "ROUND(a.ImpMN - a.ImpMN, 2) AS cSaldo "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmprptRegPerce ")
  sSentencia = sSentencia & "FROM (((CoCPbDet a "
  sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmpCoCPbDetRP b ON a.CodCta=b.CodCta AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc) "
  sSentencia = sSentencia & "LEFT JOIN TgAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
  sSentencia = sSentencia & "LEFT JOIN TgTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND b.MesPvs<='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND b.CodTDc_RtcPcp='" & gsCodTDc_Pcp & "' "
  If Trim(txtDato(0).Text) <> "" Then
    sSentencia = sSentencia & "AND a.CodAux='" & Trim(txtDato(0).Text) & "' "
  End If
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (cDebe <> 0.00) OR (cHaber <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "AND ((ROUND((ISNULL((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & " ELSE 0 END), 0) + "
    sSentencia = sSentencia & "(CASE WHEN (a.TpoPvs='" & TPOPVS_CAN & "' AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte) THEN b.cImpRetHab ELSE 0 END)), 2) <> 0.00) "
    sSentencia = sSentencia & "OR (ROUND((ISNULL((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, "MN", "ME") & " ELSE 0 END), 0) + "
    sSentencia = sSentencia & "(CASE WHEN (a.TpoPvs='" & TPOPVS_CAN & "' AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte) THEN b.cImpRetDeb ELSE 0 END)), 2) <> 0.00)) "
  End If
  sSentencia = sSentencia & "UNION "
  sSentencia = sSentencia & "SELECT DISTINCT b.CodCta AS CodCta, b.CodAux AS CodAux, b.CodTDc AS CodTDc, b.SerDoc AS SerDoc, b.NroDoc AS NroDoc, "
  sSentencia = sSentencia & "b.MesPvs, b.FeEDoc_RtcPcp AS FehOpe, b.FeEDoc_RtcPcp AS FeEDoc, '" & TPOPVS_PVS & "' AS TpoPvs, '" & TPOGNR_DRP & "' AS TpoGnr, 'PERCEPCION' AS cOperacion, "
  sSentencia = sSentencia & "b.CodTDc_RtcPcp, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(b.SerDoc_RtcPcp, ' - ', b.NroDoc_RtcPcp)", "(b.SerDoc_RtcPcp+' - '+b.NroDoc_RtcPcp)") & " AS cDocumento, "
  sSentencia = sSentencia & "'2' AS cOrden, c.RazAux, c.RUCAux, " & Choose(gsIdioma, "d.DetTDc", "d.DetTDcx") & " AS DetTDc, "
  sSentencia = sSentencia & "b.cImpRetDeb AS cDebe, b.cImpRetHab AS cHaber, "
  sSentencia = sSentencia & "(CASE WHEN b.MesPvs='" & gsMesAct & "' THEN b.cImpRetDeb ELSE 0 END) AS cImpRetDeb, "
  sSentencia = sSentencia & "(CASE WHEN b.MesPvs='" & gsMesAct & "' THEN b.cImpRetHab ELSE 0 END) AS cImpRetHab, ROUND(0, 2) AS cSaldo "
  sSentencia = sSentencia & "FROM ((" & ps_Prefijo & "tmpCoCPbDetRP b "
  sSentencia = sSentencia & "LEFT JOIN TgAux c ON c.codemp='" & gsCodEmp & "' AND b.CodAux=c.CodAux) "
  sSentencia = sSentencia & "LEFT JOIN TgTDc d ON d.codemp='" & gsCodEmp & "' AND b.CodTDc_RtcPcp=d.CodTDc) "
  sSentencia = sSentencia & "WHERE b.MesPvs<='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND b.CodTDc_RtcPcp='" & gsCodTDc_Pcp & "' "
  If Trim(txtDato(0).Text) <> "" Then
    sSentencia = sSentencia & "AND b.CodAux='" & Trim(txtDato(0).Text) & "' "
  End If
  sSentencia = sSentencia & "ORDER BY CodAux, CodTDc, SerDoc, NroDoc, FehOpe, FeEDoc, cOrden, cTipoDoc"
  ' Genero la tabla temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmprptRegPerce", "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 16)='#tmprptRegPerce_') DROP TABLE #tmprptRegPerce")
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmprptRegPerce ", "") & sSentencia
  pocnnMain.Execute sSentencia
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpCoCPbDetRPDoc", "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 18)='#tmpCoCPbDetRPDoc_') DROP TABLE #tmpCoCPbDetRPDoc")
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpCoCPbDetRP", "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 16)='#tmpCoCPbDetRP_') DROP TABLE #tmpCoCPbDetRP")
  
  ' Actualizo la tabla del reporte
  ppDatosPercepcion
  '[ Creacion de Temporales
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT * FROM " & ps_Prefijo & "tmprptRegPerce "
    .Source = .Source & "ORDER BY CodAux, CodTDc, SerDoc, NroDoc, FehOpe, FeEDoc, cOrden, cTipoDoc"
    .Open
  End With
     
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRRegPcp.rpt"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      '.LoadReport gsRutRpt & "rptRRegHPr.mrp"
      .LoadReport gsRutRpt & "rptRRegPcp.mrp"
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pSigMon") = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, gsTpoMon_Sgn_ME, gsTpoMon_Sgn_MN)
      ']
      If Index = 0 Then
        .PreviewReport
      Else
        '[ARREGLAR: Revisar el uso de los tres primeros parámetros de Print.
        .Print 1, 0, 0, unCopias
        ']ARREGLAR.
      End If
      .UnLoadReport
    End With
    Set MRViewer = Nothing
  End If
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmprptRegPerce", "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 16)='#tmprptRegPerce_') DROP TABLE #tmprptRegPerce")
  porstMRp.Close

  ppHabilitacion True

End Sub

Private Sub cmdConfig_Click()
   With frmOPrnCfg
      .ConfiguraPrn 0, Me
   
      .Show vbModal
    
      .ConfiguraPrn 1, Me
   End With
   
   cmdImprimir(1).SetFocus
End Sub

Private Sub cmdSalir_Click()
   frmOPrnCfg.OrientacionPrn 1, Me
   
   Unload Me
End Sub

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
'End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   txtDato(Index).SelStart = 0
   txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
'   Select Case Index    'Completa con ceros a la izquierda.
'   Case 0, 1                           'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0                              'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstTGAux
         .MoveFirst
         .Find "CodAux='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & !RazAux
         End If
      End With
   End Select
End Function

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
   optTipoImpresion(0).Enabled = tbHabilitar
   optTipoImpresion(1).Enabled = tbHabilitar
   cmdImprimir(0).Enabled = tbHabilitar
   cmdImprimir(1).Enabled = tbHabilitar
   cmdConfig.Enabled = tbHabilitar
   cmdSalir.Enabled = tbHabilitar

End Sub

'[Propio del formulario.

Private Sub ppDatosPercepcion()
   
  Static porstRpt As ADODB.Recordset
  Static nDebe As Double, nHaber As Double, nSaldo As Double
  Static sCodAux As String, sCodTDc As String
  Static sSerDoc As String, sNroDoc As String
  Static lRptJump As Boolean
  
  ' Seteo el recordset temporal para la grabacion
  Set porstRpt = New ADODB.Recordset
  With porstRpt
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = "SELECT * FROM " & ps_Prefijo & "tmprptRegPerce "
    .Source = .Source & "ORDER BY CodAux, CodTDc, SerDoc, NroDoc, FehOpe, FeEDoc, cOrden, cTipoDoc"
    .CursorType = adOpenDynamic
    .LockType = adLockBatchOptimistic
    .Open
  End With
  If Not (porstRpt.BOF And porstRpt.EOF) Then
    porstRpt.MoveFirst
    sCodAux = "": sCodTDc = "": sSerDoc = "": sNroDoc = ""
    nDebe = 0: nHaber = 0: nSaldo = 0
    Do While Not porstRpt.EOF
      lRptJump = Not (sCodAux = porstRpt!codaux And sCodTDc = porstRpt!CodTDc And sSerDoc = porstRpt!SerDoc And sNroDoc = porstRpt!NroDoc)
      sCodAux = porstRpt!codaux: sCodTDc = porstRpt!CodTDc
      sSerDoc = porstRpt!SerDoc: sNroDoc = porstRpt!NroDoc
      nDebe = IIf(lRptJump, 0, nDebe)
      nHaber = IIf(lRptJump, 0, nHaber)
      nDebe = Round(CDec(nDebe + porstRpt!cDebe), 2)
      nHaber = Round(CDec(nHaber + porstRpt!cHaber), 2)
      nSaldo = Round(CDec(nDebe - nHaber), 2)
      ' Modifico los saldos del documento
      porstRpt!cSaldo = nSaldo
      porstRpt.MoveNext
    Loop
    porstRpt.UpdateBatch
  End If
  porstRpt.Close
  Set porstRpt = Nothing

End Sub

']Propio del formulario.

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

