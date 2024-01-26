VERSION 5.00
Begin VB.Form frmRCpbNCu_1 
   Caption         =   "[título]"
   ClientHeight    =   2295
   ClientLeft      =   1620
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   2640
      TabIndex        =   17
      Top             =   990
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   19
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo "
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   0
      TabIndex        =   16
      Top             =   990
      Width           =   2175
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   315
         Width           =   1035
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Mes"
      ForeColor       =   &H80000002&
      Height          =   780
      Left            =   0
      TabIndex        =   11
      Top             =   90
      Width           =   3060
      Begin VB.ComboBox CmbMes 
         Height          =   315
         ItemData        =   "frmRCpbNCu_1.frx":0000
         Left            =   1035
         List            =   "frmRCpbNCu_1.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   315
         Width           =   1905
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Todos"
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4125
         Picture         =   "frmRCpbNCu_1.frx":00A8
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1170
         Visible         =   0   'False
         Width           =   255
      End
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
         Left            =   165
         TabIndex        =   8
         Top             =   1155
         Width           =   315
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   4125
         Picture         =   "frmRCpbNCu_1.frx":0252
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1530
         Visible         =   0   'False
         Width           =   255
      End
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
         Index           =   1
         Left            =   165
         TabIndex        =   9
         Top             =   1515
         Visible         =   0   'False
         Width           =   315
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
         Left            =   465
         TabIndex        =   15
         Top             =   1155
         Width           =   3675
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
         Index           =   1
         Left            =   465
         TabIndex        =   14
         Top             =   1515
         Visible         =   0   'False
         Width           =   3675
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
      TabIndex        =   10
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
         Picture         =   "frmRCpbNCu_1.frx":03FC
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
         Picture         =   "frmRCpbNCu_1.frx":0546
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
         Picture         =   "frmRCpbNCu_1.frx":0A78
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRCpbNCu_1"
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
Private porstCOCpbDet As ADODB.Recordset
Private porstCrystal As ADODB.Recordset
']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   
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
   With porstCOCpbDet
      .ActiveConnection = pocnnMain
      .Source = "SELECT MesPvs "
      .Source = .Source & "FROM CocpbDet "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "MesPvs"
         .Item(dnContador).MaxLength = porstCOCpbDet.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  fraRangos.Caption = Choose(gsIdioma, "Mes", "Month")
  chkMes.Caption = Choose(gsIdioma, "Todos", "All")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCpbDet
      .MoveLast
      txtDato(1).Text = !mespvs
      .MoveFirst
      txtDato(0).Text = !mespvs
   End With

  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
  CmbMes.ListIndex = gsMesAct
   
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
   porstCOCpbDet.Close
   pocnnMain.Close
   Set porstCOCpbDet = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim CadCrystal As String

  ppHabilitacion False
  
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  '[rcs 050604.
  'primer temporal, calcula los saldos de D/H
  If ps_Plataforma = pSrvMySql Then
    pocnnMain.Execute "DROP TABLE IF EXISTS tRptRCpbNCu_1A"
    CadCrystal = "CREATE TEMPORARY TABLE IF NOT EXISTS tRptRCpbNCu_1A "
  ElseIf ps_Plataforma = pSrvSql Then
    pocnnMain.Execute "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 15)='#tRptRCpbNCu_1A') DROP TABLE #tRptRCpbNCu_1A"
    CadCrystal = ""
  End If
'ini 2016-03-31 error rpt descudrados
''  CadCrystal = CadCrystal & "SELECT MesPvs,CodDro, NroCpb, "
''  CadCrystal = CadCrystal & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END)), 2) AS clmpDeb, "
''  CadCrystal = CadCrystal & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END)), 2) AS clmpHab, "
''  CadCrystal = CadCrystal & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END)), 2) AS clmpDebME, "
''  CadCrystal = CadCrystal & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END)), 2) AS clmpHabME "
''  CadCrystal = CadCrystal & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tRptRCpbNCu_1A ")
''  CadCrystal = CadCrystal & "FROM cocpbdet "
''  CadCrystal = CadCrystal & "WHERE codemp='" & gsCodEmp & "' "
''  CadCrystal = CadCrystal & "AND pdoano='" & gsAnoAct & "' "
''  CadCrystal = CadCrystal & "GROUP BY MesPvs, CodDro, NroCpb "
''  CadCrystal = CadCrystal & "HAVING "
''  CadCrystal = CadCrystal & "ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE ImpMN * -1 END)), 2) <> 0.00 "
''  CadCrystal = CadCrystal & " OR ROUND(SUM((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE ImpME * -1 END)), 2) <> 0.00 "
'------------------------------
CadCrystal = CadCrystal & "SELECT mespvs,coddro,nrocpb,"
CadCrystal = CadCrystal & "ROUND(SUM(CASE TPOCTB WHEN '" & TPOCTB_DEB & "' THEN impmn ELSE 0.00 END), 2) AS clmpDeb,"
CadCrystal = CadCrystal & "ROUND(SUM(CASE TPOCTB WHEN '" & TPOCTB_HAB & "' THEN impmn ELSE 0.00 END), 2) AS clmpHab,"
CadCrystal = CadCrystal & "ROUND(SUM(CASE TPOCTB WHEN '" & TPOCTB_DEB & "' THEN impmn ELSE 0.00 END)-SUM(CASE TPOCTB WHEN '" & TPOCTB_HAB & "' THEN impmn ELSE 0.00 END), 2) AS X,"
CadCrystal = CadCrystal & "ROUND(SUM(CASE TPOCTB WHEN '" & TPOCTB_DEB & "' THEN impme ELSE 0.00 END), 2) AS clmpDebME,"
CadCrystal = CadCrystal & "ROUND(SUM(CASE TPOCTB WHEN '" & TPOCTB_HAB & "' THEN impme ELSE 0.00 END), 2) AS clmpHabME,"
CadCrystal = CadCrystal & "ROUND(SUM(CASE TPOCTB WHEN '" & TPOCTB_DEB & "' THEN impme ELSE 0.00 END)-SUM(CASE TPOCTB WHEN '" & TPOCTB_HAB & "' THEN impme ELSE 0.00 END), 2) AS Y "
CadCrystal = CadCrystal & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tRptRCpbNCu_1A ")
'CadCrystal = CadCrystal & "FROM cocpbdet WHERE codemp='012' and pdoano='2012' and MESPVS<='12' "
CadCrystal = CadCrystal & "FROM cocpbdet "
CadCrystal = CadCrystal & "WHERE codemp='" & gsCodEmp & "' "
CadCrystal = CadCrystal & "AND pdoano='" & gsAnoAct & "' "
CadCrystal = CadCrystal & "AND MESPVS<='" & Format(CmbMes.ListIndex, "00") & "' "

CadCrystal = CadCrystal & "GROUP BY mespvs,coddro,nrocpb "
If ps_Plataforma = pSrvMySql Then
CadCrystal = CadCrystal & "HAVING x <> 0.00 Or Y <> 0.00 "
Else
  CadCrystal = CadCrystal & "HAVING "
  CadCrystal = CadCrystal & "ROUND(SUM(CASE TPOCTB WHEN 'D' THEN impmn ELSE 0.00 END)-SUM(CASE TPOCTB WHEN 'H' THEN impmn ELSE 0.00 END), 2) <> 0.00 "
  CadCrystal = CadCrystal & " OR ROUND(SUM(CASE TPOCTB WHEN 'D' THEN impme ELSE 0.00 END)-SUM(CASE TPOCTB WHEN 'H' THEN impme ELSE 0.00 END), 2) <> 0.00 "
End If
'fin 2016-03-31 error rpt descudrados
  pocnnMain.Execute CadCrystal
  ']rcs.
  
  With porstMRp
    If .State = adStateOpen Then .Close
    '[rcs 050604.
    If OptTipo(0).Value = True Then
      .Source = "SELECT b.IndNCu, x.clmpDeb AS D_MN, x.clmpHab AS H_MN, "
      .Source = .Source & "x.clmpDebME AS D_ME, x.clmpHabME AS H_ME, b.MesPvs, "
      .Source = .Source & "a.CodCta, a.CodDro, a.NroCpb, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
      .Source = .Source & "a.CodAux, d.RazAux, a.BlqIte, a.FehOpe, "
    Else
      .Source = "SELECT b.IndNCu, x.clmpDeb AS D_MN, x.clmpHab AS H_MN, "
      .Source = .Source & "x.clmpDebME AS D_ME, x.clmpHabME AS H_ME, b.MesPvs, "
      .Source = .Source & "a.CodCta, a.CodDro, a.NroCpb, " & Choose(gsIdioma, "b.GloCpb", "b.GloCpb") & " AS GloIte, "
      .Source = .Source & "a.CodAux, d.RazAux, a.BlqIte, b.FehCpb FehOpe, "
    End If
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cNroDoc, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END) AS clmpDeb, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END) AS clmpHab, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END) AS clmpDebME, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END) AS clmpHabME, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodDro, '-', a.NroCpb)", "(a.CodDro+'-'+a.NroCpb)") & " AS cDroCpb "
    .Source = .Source & "FROM ((((CocPbDet a "
    .Source = .Source & "LEFT JOIN CoCpbCab b ON  a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb "
    .Source = .Source & "LEFT JOIN " & ps_Prefijo & "tRptRCpbNCu_1A x ON  a.MesPvs=x.MesPvs AND a.CodDro=x.CodDro AND a.NroCpb=x.NroCpb "
    .Source = .Source & "LEFT JOIN TgTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc "
    .Source = .Source & "LEFT JOIN TgAux d ON a.codemp=d.codemp AND a.CodAux=d.CodAux )))) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND (x.clmpDeb <> 0 AND x.clmpHab <> 0 AND x.clmpDebME <> 0 AND x.clmpHabME <> 0) "
    .Source = .Source & "AND a.MesPvs" & IIf(chkMes.Value = vbChecked, "<=", "=") & "'" & Format(CmbMes.ListIndex, "00") & "' "
    .Source = .Source & "ORDER BY b.MesPvs, a.CodDro, a.NroCpb, a.NroIte"
    .Open
  End With
  ']rcs.
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & IIf(OptTipo(0).Value, Choose(gsIdioma, " (Detallado)", " (Detail)"), Choose(gsIdioma, " (Resumen)", " (Summary)")), udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRCpbNCu_1.rpt", "rptRCpbNCuRes_1.rpt")
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
        .LoadReport gsRutRpt & "rptRCpbNCu_1.mrp"
      Else
        .LoadReport gsRutRpt & "rptRCpbNCuRes_1.mrp"
      End If
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & IIf(OptTipo(0).Value = True, " (Detallado)", " (Resumen)"), udFecha, True, chkImpFecha.Value)
      
      '[Parámetros adicionales.
      If chkMes.Value = False Then
        .Parameters("pPeriodoAdc") = Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
      Else
        .Parameters("pPeriodoAdc") = "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
      End If
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
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tRptRCpbNCu_1A", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 15)='#tRptRCpbNCu_1A') DROP TABLE #tRptRCpbNCu_1A")
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
   Case 0, 1                           'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   
   Select Case tnIndex                 'Cambiar.
    Case 0, 1
      If Val(txtDato(0)) > 12 Or Val(txtDato(0)) = 0 Then
        txtDato(0).Text = "99"
        lblDatoDeta(0).Caption = "Todos los Meses"
      Else
        lblDatoDeta(0).Caption = Format(CVDate("01" & "/" & (txtDato(tnIndex) & "/" & Year(udFecha))), "mmmm")
      End If
'      If txtDato(tnIndex).Text = "" Then
'         lblDatoDeta(tnIndex).Caption = ""
'         Exit Function
'      End If
'      With porstCocpbDet
'         .MoveFirst
'         .Find "MesPvs='" & txtDato(tnIndex).Text & "'"
'         If .EOF Then
'            MsgBox TEXT_8006, vbExclamation
'            ppAyuDet = True
'         Else
'            lblDatoDeta(tnIndex).Caption = " " & !DetTDc
'         End If
'      End With
   End Select
End Function

'[Propio del formulario.

Private Sub ChkMes_Click()
    If chkMes.Value = 0 Then CmbMes.Enabled = True
    If chkMes.Value = 1 Then CmbMes.Enabled = False
End Sub
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
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

