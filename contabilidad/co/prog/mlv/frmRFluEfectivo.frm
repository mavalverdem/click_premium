VERSION 5.00
Begin VB.Form frmRFluEfectivo 
   Caption         =   "[título]"
   ClientHeight    =   2295
   ClientLeft      =   3750
   ClientTop       =   2445
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4845
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox cmbMoneda 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3600
      TabIndex        =   13
      Text            =   "Moneda"
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cmbPeriodo 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   675
      TabIndex        =   11
      Text            =   "Periodo"
      Top             =   270
      Width           =   2000
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2640
      TabIndex        =   8
      Top             =   1020
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   9
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   990
      Width           =   2175
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   5
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1005
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
      ScaleWidth      =   4845
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1755
      Width           =   4845
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
         TabIndex        =   2
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
         Picture         =   "frmRFluEfectivo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmRFluEfectivo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmRFluEfectivo.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label LblTexto 
      Caption         =   "Periodo"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   30
      TabIndex        =   14
      Top             =   285
      Width           =   600
   End
   Begin VB.Label LblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   1
      Left            =   2775
      TabIndex        =   12
      Top             =   300
      Width           =   735
   End
End
Attribute VB_Name = "frmRFluEfectivo"
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

Public pocnnMain As ADODB.Connection

Private Sub Form_Activate()
   'Orden: Vista Previa, Imprimir, Exportar.
   zaOpciones = Array(gbPms04, gbPms05, gbPms06)
End Sub

Private Sub Form_Load()

   Dim dnContador As Integer
'[Parametros
   For dnContador = 0 To 13
    If gsIdioma = NvlUsr_Sup Then
      cmbPeriodo.AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre") & " " & gsAnoAct
    Else
      cmbPeriodo.AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing") & " " & gsAnoAct
    End If
   Next dnContador
   cmbPeriodo.ListIndex = Val(gsMesAct)
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Periodo :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Period :", "Currency :")
  Next nElemento
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
 
   With cmbmoneda
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
   End With
   cmbmoneda.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   
  ' Características de impresión.
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

End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   picOpciones.Width = Me.Width - 120
   cmdSalir.Left = picOpciones.Width - 1135
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    
  Dim sSentencia As String, sMoneda As String
  Dim porstMRp As ADODB.Recordset

  'On Error GoTo Err
    
  ppHabilitacion False
  
  ' Seteo y activo la coneccion
  Set pocnnMain = New ADODB.Connection
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  sMoneda = IIf(cmbmoneda.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  ' Genero la tabla temporal saldos anteriores
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpSaldos", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 10)='#tmpSaldos') DROP TABLE #tmpSaldos")
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpSaldos ", "")
  sSentencia = sSentencia & "SELECT DISTINCT "
  sSentencia = sSentencia & IIf(OptTipo(0).Value, " '" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "'", " (MesPvs +1 )") & " AS MesPvs, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mSalDebe, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mSalHaber "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmpSaldos ", "")
  sSentencia = sSentencia & "FROM ((CoCpbDetFjo a "
  sSentencia = sSentencia & "LEFT JOIN CoFjo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodFjo=b.CodFjo) "
  sSentencia = sSentencia & "LEFT JOIN CoCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "' AND c.IndFjo='" & INDFJO_ACT & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodFjo, '')<>'' "
  sSentencia = sSentencia & IIf(OptTipo(1).Value, "GROUP BY MesPvs ORDER BY MesPvs", "")
  pocnnMain.Execute sSentencia
  
  ' Creo la sentencia de seleccion del reporte
  sSentencia = "SELECT DISTINCT a.MesPvs, c.TpoEfe, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT('ACTIVIDAD DE ',(CASE c.TpoEfe WHEN '" & TPOEFE_OPE & "' THEN '" & UCase(TPOEFE_OPE_TXT) & "' WHEN '" & TPOEFE_INV & "' THEN '" & UCase(TPOEFE_INV_TXT) & "' ELSE '" & UCase(TPOEFE_FIN_TXT) & "' END))", "('ACTIVIDAD DE '+(CASE c.TpoEfe WHEN '" & TPOEFE_OPE & "' THEN '" & UCase(TPOEFE_OPE_TXT) & "' WHEN '" & TPOEFE_INV & "' THEN '" & UCase(TPOEFE_INV_TXT) & "' ELSE '" & UCase(TPOEFE_FIN_TXT) & "' END))") & " AS mDesTitulo, "
  sSentencia = sSentencia & "b.TpoFjo, LEFT(c.CodEfe, 2) As mCodNivel, "
  sSentencia = sSentencia & Choose(gsIdioma, "d.DetEfe", "d.DetEfex") & " AS mDetNivel, c.CodEfe, "
  sSentencia = sSentencia & Choose(gsIdioma, "c.DetEfe", "c.DetEfex") & " AS DetEfe, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mDebe, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mHaber, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AVG(f.mSalDebe), 0) AS mSalDebe, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AVG(f.mSalHaber), 0) AS mSalHaber "
  sSentencia = sSentencia & "FROM (((((CoCpbDetFjo a "
  sSentencia = sSentencia & "LEFT JOIN CoFjo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodFjo=b.CodFjo) "
  sSentencia = sSentencia & "LEFT JOIN CoEfe c ON b.codemp=c.codemp AND b.pdoano=c.pdoano AND b.CodEfe=c.CodEfe) "
  sSentencia = sSentencia & "LEFT JOIN CoEfe d ON c.codemp=d.codemp AND c.pdoano=d.pdoano AND LEFT(c.CodEfe, 2)=d.CodEfe) "
  sSentencia = sSentencia & "LEFT JOIN CoCta e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.CodCta=e.CodCta) "
  sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmpSaldos f ON a.MesPvs=f.MesPvs) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs" & IIf(OptTipo(0).Value, "=", "<=") & "'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodFjo, '')<>'' "
  sSentencia = sSentencia & "AND e.IndFjo='" & INDFJO_ACT & "' "
  sSentencia = sSentencia & "GROUP BY a.MesPvs, c.TpoEfe, c.CodEfe, b.TpoFjo, " & Choose(gsIdioma, "d.DetEfe, c.DetEfe ", "d.DetEfex, c.DetEfex ")
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (mDebe<>0.00 OR mHaber<>0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ISNULL(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2)<>0.00 "
    sSentencia = sSentencia & "OR ROUND(ISNULL(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2)<>0.00) "
  End If
  sSentencia = sSentencia & "ORDER BY a.MesPvs, c.TpoEfe, b.TpoFjo, c.CodEfe "
  If OptTipo(1).Value Then
    ' Genero tabla temporal de flujo anual
    ppGene_FlujoAnual sSentencia
    sSentencia = "SELECT * FROM tmpFlujoAnual ORDER BY TpoFjo, CodFjo"
  End If
  ' Seteo y activo el recordset
  Set porstMRp = New ADODB.Recordset
  With porstMRp
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
'    .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = sSentencia
    .Open
  End With
    
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, UCase(Me.Caption) & " (" & IIf(cmbmoneda.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRFluEfeM.rpt", "rptRFluEfeA.rpt")
      '         .WindowShowGroupTree = True
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  End If
  porstMRp.Close
  Set porstMRp = Nothing
  pocnnMain.Close
  Set pocnnMain = Nothing
  ppHabilitacion True
  Exit Sub
  
Err:
  pocnnMain.Close
  Set pocnnMain = Nothing
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

']

Private Sub ppGene_FlujoAnual(sSentencia As String)
  
  Dim porstReg As ADODB.Recordset
  Dim porstTmp As ADODB.Recordset
  Dim nDebeIni As Double, nHaberIni As Double
  Dim nSaldoDeb As Double, nSaldoHab As Double, nImporte As Double
  Dim sSaldoMes As String, sCommandText As String
  Dim sConversion As String

  sConversion = "CONVERT(" & IIf(ps_Plataforma = pSrvMySql, "0, decimal(18, 2)", "decimal(18, 2), 0") & ")"
  ' Creo la tabla temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpFlujoAnual", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#tmpFlujoAnual') DROP TABLE #tmpFlujoAnual")
  sCommandText = sCommandText & "SELECT DISTINCT b.TpoFjo, (CASE b.TpoFjo WHEN '" & TPOFJO_ING & "' THEN '" & UCase(TPOFJO_ING_TXT) & "' ELSE '" & UCase(TPOFJO_EGR_TXT) & "' END) AS mDesTitulo, "
  sCommandText = sCommandText & "LEFT(a.CodFjo, 2) As mCodNivel, c.DetFjo AS mDetNivel, a.CodFjo, " & Choose(gsIdioma, "b.DetFjo", "b.DetFjox") & " AS DetFjo, "
  sCommandText = sCommandText & sConversion & " AS mDeb00, " & sConversion & " AS mHab00, " & sConversion & " AS mDeb01, " & sConversion & " AS mHab01, "
  sCommandText = sCommandText & sConversion & " AS mDeb02, " & sConversion & " AS mHab02, " & sConversion & " AS mDeb03, " & sConversion & " AS mHab03, "
  sCommandText = sCommandText & sConversion & " AS mDeb04, " & sConversion & " AS mHab04, " & sConversion & " AS mDeb05, " & sConversion & " AS mHab05, "
  sCommandText = sCommandText & sConversion & " AS mDeb06, " & sConversion & " AS mHab06, " & sConversion & " AS mDeb07, " & sConversion & " AS mHab07, "
  sCommandText = sCommandText & sConversion & " AS mDeb08, " & sConversion & " AS mHab08, " & sConversion & " AS mDeb09, " & sConversion & " AS mHab09, "
  sCommandText = sCommandText & sConversion & " AS mDeb10, " & sConversion & " AS mHab10, " & sConversion & " AS mDeb11, " & sConversion & " AS mHab11, "
  sCommandText = sCommandText & sConversion & " AS mDeb12, " & sConversion & " AS mHab12, " & sConversion & " AS mDeb13, " & sConversion & " AS mHab13, "
  sCommandText = sCommandText & sConversion & " AS mDebTo, " & sConversion & " AS mHabTo, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb00, " & sConversion & " AS mSalHab00, " & sConversion & " AS mSalDeb01, " & sConversion & " AS mSalHab01, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb02, " & sConversion & " AS mSalHab02, " & sConversion & " AS mSalDeb03, " & sConversion & " AS mSalHab03, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb04, " & sConversion & " AS mSalHab04, " & sConversion & " AS mSalDeb05, " & sConversion & " AS mSalHab05, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb06, " & sConversion & " AS mSalHab06, " & sConversion & " AS mSalDeb07, " & sConversion & " AS mSalHab07, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb08, " & sConversion & " AS mSalHab08, " & sConversion & " AS mSalDeb09, " & sConversion & " AS mSalHab09, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb10, " & sConversion & " AS mSalHab10, " & sConversion & " AS mSalDeb11, " & sConversion & " AS mSalHab11, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb12, " & sConversion & " AS mSalHab12, " & sConversion & " AS mSalDeb13, " & sConversion & " AS mSalHab13, "
  sCommandText = sCommandText & sConversion & " AS mSalDebTo, " & sConversion & " AS mSalHabTo "
  sCommandText = sCommandText & IIf(ps_Plataforma = pSrvSql, "INTO #tmpFlujoAnual ", "")
  sCommandText = sCommandText & "FROM (((CoCpbDetFjo a "
  sCommandText = sCommandText & "LEFT JOIN CoFjo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodFjo=b.CodFjo) "
  sCommandText = sCommandText & "LEFT JOIN CoFjo c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND LEFT(a.CodFjo, 2)=c.CodFjo) "
  sCommandText = sCommandText & "LEFT JOIN CoCta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCta=d.CodCta) "
  sCommandText = sCommandText & "WHERE a.CodFjo='" & gsCodEmp & "' "
  sCommandText = sCommandText & "AND a.pdoano='" & gsAnoAct & "' "
  sCommandText = sCommandText & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodFjo, '')<>'' "
  sCommandText = sCommandText & "AND a.MesPvs" & IIf(OptTipo(0).Value, "=", "<=") & "'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sCommandText = sCommandText & "AND d.IndFjo='" & INDFJO_ACT & "' "
  sCommandText = sCommandText & "GROUP BY b.TpoFjo, a.CodFjo, b.DetFjo, c.DetFjo "
  sCommandText = sCommandText & "ORDER BY b.TpoFjo, a.CodFjo"
  pocnnMain.Execute sCommandText
  
  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  With porstTmp
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open
  End With
  If Not (porstTmp.BOF And porstTmp.EOF) Then
    porstTmp.MoveFirst
    sSaldoMes = "": nSaldoDeb = 0: nSaldoHab = 0
    nDebeIni = CDec(IIf(IsNull(porstTmp!mSalDebe), 0, porstTmp!mSalDebe))
    nHaberIni = CDec(IIf(IsNull(porstTmp!mSalHaber), 0, porstTmp!mSalHaber))
    Set porstReg = New ADODB.Recordset
    With porstReg
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = "SELECT * "
      .Source = .Source & "FROM " & ps_Prefijo & "tmpFlujoAnual "
      .Source = .Source & "ORDER BY CodFjo"
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
      
    Do While Not porstTmp.EOF
      nImporte = 0
      If sSaldoMes <> porstTmp!mespvs Then
        sSaldoMes = porstTmp!mespvs
        nImporte = CDec(IIf(IsNull(porstTmp!mSalDebe), 0, porstTmp!mSalDebe))
        nSaldoDeb = CDec(nSaldoDeb) + nImporte
        nImporte = CDec(IIf(IsNull(porstTmp!mSalHaber), 0, porstTmp!mSalHaber))
        nSaldoHab = CDec(nSaldoHab) + nImporte
        ' Actualizo el saldo inicial del mes
        sCommandText = "UPDATE tmpFlujoAnual SET"
        sCommandText = sCommandText & " mSalDeb" & porstTmp!mespvs & "=" & CDbl(nSaldoDeb) & ","
        sCommandText = sCommandText & " mSalHab" & porstTmp!mespvs & "=" & CDbl(nSaldoHab)
        pocnnMain.Execute sCommandText
        porstReg.Requery
      End If
      ' Modifico los saldos de los flujos
      With porstReg
        If .RecordCount <> 0 Then .MoveFirst
        .Find "CodFjo = '" & porstTmp!CodFjo & "'"
        !CodFjo = porstTmp!CodFjo
        nImporte = IIf(IsNull(.Fields("mDeb" & porstTmp!mespvs)), 0, .Fields("mDeb" & porstTmp!mespvs))
        .Fields("mDeb" & porstTmp!mespvs) = CDec(nImporte) + porstTmp!mDebe
        nImporte = IIf(IsNull(.Fields("mHab" & porstTmp!mespvs)), 0, .Fields("mHab" & porstTmp!mespvs))
        .Fields("mHab" & porstTmp!mespvs) = CDec(nImporte) + porstTmp!mHaber
        !mDebTo = CDec(!mDebTo) + porstTmp!mDebe
        !mHabTo = CDec(!mHabTo) + porstTmp!mHaber
        !mSalDebTo = nDebeIni
        !mSalHabTo = nHaberIni
        .UpdateBatch
        .Requery
      End With
      porstTmp.MoveNext
    Loop
    porstReg.Close
    Set porstReg = Nothing
  End If
  porstTmp.Close
  Set porstTmp = Nothing

End Sub

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   
   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
   optTipoImpresion(0).Enabled = tbHabilitar
   optTipoImpresion(1).Enabled = tbHabilitar
   cmdImprimir(0).Enabled = tbHabilitar
   cmdImprimir(1).Enabled = tbHabilitar
   cmdConfig.Enabled = tbHabilitar
   cmdSalir.Enabled = tbHabilitar

End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

