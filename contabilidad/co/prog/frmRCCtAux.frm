VERSION 5.00
Begin VB.Form frmRCCtAux 
   Caption         =   "[título]"
   ClientHeight    =   2355
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   5100
      TabIndex        =   12
      Top             =   840
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   14
         Top             =   225
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   13
         Top             =   225
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   945
      Width           =   2175
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   6
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Auxiliar"
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   0
      TabIndex        =   8
      Top             =   45
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
         TabIndex        =   4
         Top             =   315
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6885
         Picture         =   "frmRCCtAux.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   315
         Width           =   5520
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
      ScaleWidth      =   7290
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1815
      Width           =   7290
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
         Left            =   3690
         Picture         =   "frmRCCtAux.frx":01AA
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
         Picture         =   "frmRCCtAux.frx":02F4
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
         Picture         =   "frmRCCtAux.frx":0826
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRCCtAux"
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
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
    End With
    With porstTGAux
       .ActiveConnection = pocnnMain
       .Source = "SELECT CodAux, RazAux "
       .Source = .Source & "FROM TGAux "
       .Source = .Source & "WHERE codemp='" & gsCodEmp & "'"
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenDynamic
       .LockType = adLockReadOnly
       .Open
    End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 0  ' 1
         .Item(dnContador).DataField = "CodAux"
         .Item(dnContador).MaxLength = porstTGAux.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  fraAuxiliar.Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
'   With porstTgAux
'      .MoveLast
'      txtDato(1).Text = !CodAux
'      .MoveFirst
'      txtDato(0).Text = !CodAux
'   End With
  
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   'If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
   OptTipo(0).Value = True
   
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
   Case 0          ', 1
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim cCadReporte  As String
  
  ppHabilitacion False
    
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpRptCtaCte_') DROP TABLE #tmpRptCtaCte"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRptCtaCte", cCadReporte)

  cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmpRptCtaCte ", "")
  cCadReporte = cCadReporte & "SELECT a.CodAux, a.CodCta, c.AbvTDc, a.CodTDc, a.SerDoc, a.NroDoc, b.RazAux, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol "
  cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpRptCtaCte ")
  cCadReporte = cCadReporte & "FROM (((cocpbdet a "
  cCadReporte = cCadReporte & "LEFT JOIN tgaux b ON a.codemp=b.codemp AND a.codaux=b.codaux) "
  cCadReporte = cCadReporte & "LEFT JOIN tgtdc c ON a.codemp=c.codemp AND a.codtdc=c.codtdc) "
  cCadReporte = cCadReporte & "LEFT JOIN cocta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.codcta=d.codcta) "
  cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
  cCadReporte = cCadReporte & "AND a.pdoano='" & gsAnoAct & "' "
  cCadReporte = cCadReporte & "AND a.MesPvs<='" & gsMesAct & "' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodTDc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.SerDoc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.NroDoc, '') <>'' AND d.inddoc='1' "
  If Trim(txtDato(0).Text) <> "" Then
    cCadReporte = cCadReporte & "AND a.CodAux='" & txtDato(0).Text & "' "
  End If
  cCadReporte = cCadReporte & "GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, c.AbvTDc, a.CodTDc, b.RazAux "
  If ps_Plataforma = pSrvMySql Then
    cCadReporte = cCadReporte & "HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
  Else
    cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
    cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 "
    cCadReporte = cCadReporte & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) - "
    cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
  End If
  cCadReporte = cCadReporte & "ORDER BY a.CodAux, a.CodCta, a.CodTDc, a.SerDoc, a.NroDoc"
  pocnnMain.Execute cCadReporte
   
  If OptTipo(0).Value Then
    cCadReporte = "SELECT  Distinct a.CodAux, a.CodCta, a.RazAux, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(a.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocum, "
    cCadReporte = cCadReporte & "b.CodDro, b.NroCpb, b.FeEDoc, b.FeVDoc, b.RefDoc, " & Choose(gsIdioma, "b.GloIte", "b.GloItex") & " AS GloIte, a.DebeSol AS NumCol1, "
    cCadReporte = cCadReporte & "a.HaberSol AS NumCol2, a.DebeDol AS NumCol3, a.HaberDol AS NumCol4 "
    cCadReporte = cCadReporte & "FROM (" & ps_Prefijo & "tmpRptCtaCte a "
    cCadReporte = cCadReporte & "INNER JOIN COCpbDet b ON b.codemp='" & gsCodEmp & "' AND b.pdoano='" & gsAnoAct & "' AND a.CodAux=b.CodAux AND a.CodCta=b.CodCta AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc) "
    cCadReporte = cCadReporte & "WHERE b.TpoPvs='" & TPOPVS_PVS & "' "
    cCadReporte = cCadReporte & "ORDER BY a.CodAux, a.CodCta, a.CodTDc, a.SerDoc, a.NroDoc"
  Else
    cCadReporte = "SELECT Distinct a.CodAux, a.RazAux, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.DebeSol), 0), 2) AS NumCol1, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.HaberSol), 0), 2) AS NumCol2, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.DebeDol), 0), 2) AS NumCol3, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.HaberDol), 0), 2) AS NumCol4 "
    cCadReporte = cCadReporte & "FROM " & ps_Prefijo & "tmpRptCtaCte a "
    cCadReporte = cCadReporte & "GROUP BY a.CodAux, a.RazAux "
    If ps_Plataforma = pSrvMySql Then
      cCadReporte = cCadReporte & "HAVING (ROUND(NumCol1 - NumCol2, 2) <> 0.00 OR ROUND(NumCol3 - NumCol4, 2) <> 0.00) "
    Else
      cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM(a.DebeSol), 0), 2) - ROUND(ISNULL(SUM(a.HaberSol), 0), 2), 2) <> 0.00) "
      cCadReporte = cCadReporte & "OR (ROUND(ROUND(ISNULL(SUM(a.DebeDol), 0), 2) - ROUND(ISNULL(SUM(a.HaberDol), 0), 2), 2) <> 0.00) "
    End If
    cCadReporte = cCadReporte & "ORDER BY a.CodAux, a.RazAux"
  End If
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = cCadReporte
    .Open
  End With
    
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(OptTipo(0).Value = True, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRCCTAuxDet.rpt", "rptRCCtAuxRes.rpt")
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
      If OptTipo(0).Value = True Then
        .LoadReport gsRutRpt & "rptRCCTAuxDet.mrp"
      Else
        .LoadReport gsRutRpt & "rptRCCtAuxRes.mrp"
      End If
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(OptTipo(0).Value = True, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pPeriodoAdc") = "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
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
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpRptCtaCte_') DROP TABLE #tmpRptCtaCte"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRptCtaCte", cCadReporte)
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
   Select Case Index    'Completa con ceros a la izquierda.
   Case 0                              'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0     ', 1                           'Cambiar (añadir índices).
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

  'Controles del formulario.
'   cboTpoMon.Enabled = tbHabilitar
'   dtpFecha.Enabled = tbHabilitar
'   optTipo(0).Enabled = tbHabilitar
'   optTipo(1).Enabled = tbHabilitar
'   With txtDato
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With cmdDatoAyud
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With lblDatoDeta
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

