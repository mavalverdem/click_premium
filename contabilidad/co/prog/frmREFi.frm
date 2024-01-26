VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmREFi 
   Caption         =   "[título]"
   ClientHeight    =   4785
   ClientLeft      =   1740
   ClientTop       =   1500
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6330
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   2685
      Width           =   2805
      Begin VB.OptionButton OptTipo 
         Caption         =   "Anual"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   14
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Al Mes"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   315
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Del Mes"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.Frame fraImpresion 
      Caption         =   " Tipo de Impresora "
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   3840
      TabIndex        =   7
      Top             =   3465
      Width           =   2430
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "&Gráfica"
         ForeColor       =   &H80000001&
         Height          =   240
         Index           =   1
         Left            =   1335
         TabIndex        =   9
         Top             =   270
         Width           =   900
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "&Matricial"
         ForeColor       =   &H80000001&
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2715
      Width           =   1260
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   6330
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4245
      Width           =   6330
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
         Picture         =   "frmREFi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmREFi.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmREFi.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
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
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Height          =   2490
      Left            =   45
      TabIndex        =   13
      Top             =   90
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   4392
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Estados Financieros"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   4380
      TabIndex        =   6
      Top             =   2760
      Width           =   600
   End
End
Attribute VB_Name = "frmREFi"
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
Private porstMRpRs  As ADODB.Recordset
Private porstCOEFi  As ADODB.Recordset
Private porstCOCtaAcu As ADODB.Recordset
Private porstCOCfg As ADODB.Recordset
Private aImpFormula(1) As Double
']

Private Sub Form_Load()
   On Error GoTo Err
  
 '[Recordsets.                         'Cambiar.
    Set pocnnMain = New ADODB.Connection
    Set porstMRp = New ADODB.Recordset
    Set porstCOEFi = New ADODB.Recordset
    Set porstMRpRs = New ADODB.Recordset
    Set porstCOCtaAcu = New ADODB.Recordset
    Set porstCOCfg = New ADODB.Recordset
    
    With pocnnMain
        .CursorLocation = adUseClient
        .ConnectionString = CONNSTRG & gsNomBDS
        .Open
    End With
    
    'Obtener simbolo de Moneda
    With porstCOCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT TpoMon_Sgn_MN, TpoMon_Sgn_ME " _
              & "FROM COCfg"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      gsTpoMon_Sgn_MN = .Fields(0)
      gsTpoMon_Sgn_ME = .Fields(1)
      .Close
    End With
        Set porstCOCfg.ActiveConnection = Nothing
        Set porstCOCfg = Nothing
    
    With porstCOCtaAcu
       .ActiveConnection = pocnnMain
       .Source = "SELECT a.*, b.NatCta FROM (COCtaAcu a LEFT JOIN CoCta b USING(CodCta)) ORDER BY CodCta"
'      .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenDynamic
       .LockType = adLockReadOnly
       .Open
    End With
    
    With porstMRp
        .ActiveConnection = pocnnMain
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
    End With
    With porstCOEFi
        .ActiveConnection = pocnnMain
        .Source = "SELECT CodEfi, DetEfi FROM CoEFi ORDER BY 1 "
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
    End With
    With porstMRpRs
        .ActiveConnection = pocnnMain
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Source = "SELECT * FROM trptREstFin ORDER BY NroLin"
'        .Open
    End With
 ']

 '[Parámetros.                         'Cambiar.
 ']
   
    With cboTpoMon
        .AddItem TPOMON_NAC_TXT_1, 0
        .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND
    
 '[Datos predeterminados.              'Cambiar.
  'Otros.
   
  'Características de impresión.
   udFecha = Date                      'Fecha en el encabezado.
'   unCopias = rptMain.CopiesToPrinter  'Cantidad de Copias.
   unMargenIzquierdo = 240             'Margen izquierdo.
   usDEstino = PRN_DEST_MATR           'PRN_DEST_GRAF:ica _
                                        PRN_DEST_MATR:icial.
   usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                        PRN_ORIE_HORI:zontal.
 ']
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = porstCOEFi
   
   frmOPrnCfg.OrientacionPrn 0, Me
   frmOPrnCfg.lblOriPrn.Caption = Printer.Orientation
      
   Exit Sub
Err:
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub Form_Activate()
   'Orden: Vista Previa, Imprimir, Exportar.
   zaOpciones = Array(gbPms04, gbPms05, gbPms06)
   ppDatosGrid
   
   
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   
   picOpciones.Width = Me.Width - 120
   cmdSalir.Left = picOpciones.Width - 1135
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Cambiar. Añadir recordsets.
   porstCOCtaAcu.Close
   porstCOEFi.Close
   pocnnMain.Close
   Set porstCOCtaAcu = Nothing
   Set porstCOEFi = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
   ppHabilitacion True
   
  If porstCOEFi.RecordCount = 0 Then
    MsgBox TEXT_8001, vbCritical
    Exit Sub
  End If

  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT CodEfi, NroLin, DetLin, TpoLin, 'b' AS SubLin, FmlLin, BsePct, GrpPct, "
    .Source = .Source & "0.00 AS ImpSaldoAnt, 0.00 AS ImpPorceAnt, "
    .Source = .Source & "0.00 AS ImpSaldoAct, 0.00 AS ImpPorceAct "
    .Source = .Source & "FROM CoEFiLin "
    .Source = .Source & "WHERE CodEfi='" & porstCOEFi!CodEFi & "' "
    .Source = .Source & "ORDER BY 2"
    .Open
  End With
  Llena_Temporal porstMRp.Source
         
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    Call gpEncabezadoRpt(frmMain.rptMain, porstCOEFi!DetEFi & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True)
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptREstFinm", "rptREstFins") & ".rpt"
      ' .WindowShowGroupTree = True
      ']
      .WindowState = crptMaximized
      .Connect = "Provider=MySqlProv;Extended Properties=" & CONNSTRG & gsNomBDS
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .WindowShowRefreshBtn = False
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
'    Set MRViewer = New MRViewerObject
'
'    With MRViewer
'      If OptTipo(0).Value Then
'        .DataRecordSet = porstMRpRs
'      Else
'        .DataRecordSet = porstMRp
'      End If
'
'      '.LoadReport gsRutRpt & IIf(OptTipo(0).Value, "rptRCtlPsp", "rptRCtlPspRes") & ".mrp"
'
'      .LoadReport gsRutRpt & IIf(OptTipo(0).Value, "rptRBceGrl", "rptRBceGrl") & ".mrp"
'
'
'      'Call gpEncabezadoMRp(MRViewer, Me.Caption & " -" & IIf(OptTipo(0).Value, "Detalle", "Resumen") & "-" & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True)
'      Call gpEncabezadoMRp(MRViewer, Me.dgrMain.Columns(1) & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True)
'
'       'Simbolo de la Moneda
'       '  Usando variables globales
'             .Parameters("cMoneda") = IIf((cboTpoMon.ListIndex = TPOMON_NAC_IND), gsTpoMon_Sgn_MN, gsTpoMon_Sgn_ME)
'       '.Parameters("cMoneda") = IIf((cboTpoMon.ListIndex = TPOMON_NAC_IND), "S/.", "US $.")
'
'
'      '[Parámetros adicionales.
'      'If optAlcance(0).Value = True Then
'      '    .Parameters("pPeriodoAdc") = "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
'      'Else
'      '    .Parameters("pPeriodoAdc") = Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
'      'End If
'      ']
'
'      If Index = 0 Then
'        .PreviewReport
'      Else
''[ARREGLAR: Revisar el uso de los tres primeros parámetros de Print.
'        .Print 1, 0, 0, unCopias
'']ARREGLAR.
'      End If
'        .UnLoadReport
'    End With
'    Set MRViewer = Nothing
  End If
  pocnnMain.Execute "DROP TABLE IF EXISTS trptREstFin"
  porstMRpRs.Close
   
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

Private Sub ppDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = "Código"
            .Item(dnNum).Width = 1000
         Case 1
            .Item(dnNum).Caption = "Descripción"
            .Item(dnNum).Width = 4680
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub
']
' ma
Public Sub Llena_Temporal(s_Source As String)

    Dim aImpSubTotLin(1) As Double, aImpTotaLin(1) As Double
    Dim aImpSaldo(1) As Double, aImpFormula(1) As Double
    Dim sCadExecute As String, nContador As Integer
    
    pocnnMain.Execute "DROP TABLE IF EXISTS trptREstFin"
    pocnnMain.Execute "CREATE TABLE IF NOT EXISTS trptREstFin " & s_Source
    pocnnMain.Execute "DELETE FROM trptREstFin"
    porstMRpRs.Open
    If porstMRp.RecordCount > 0 Then
        porstMRp.MoveFirst
        pocnnMain.BeginTrans    '[ INICIA TRANSACCION ]
        Do While Not porstMRp.EOF
            Select Case porstMRp.Fields!TpoLin
             Case 1         ' Total
               For nContador = 1 To 2
                porstMRpRs.AddNew
                porstMRpRs.Fields!CodEFi = porstMRp.Fields!CodEFi
                porstMRpRs.Fields!NroLin = porstMRp.Fields!NroLin
                porstMRpRs.Fields!TpoLin = porstMRp.Fields!TpoLin
                porstMRpRs.Fields!SubLin = Choose(nContador, "a", "c")
                ' Se incluye line en blanco
               Next nContador
               aImpTotaLin(0) = gfRedond(aImpTotaLin(0) + aImpSubTotLin(0), 2)
               aImpTotaLin(1) = gfRedond(aImpTotaLin(1) + aImpSubTotLin(1), 2)
               aImpSaldo(0) = aImpSubTotLin(0): aImpSaldo(1) = aImpSubTotLin(1)
               aImpSubTotLin(0) = 0: aImpSubTotLin(1) = 0
             Case 2         ' Sub Total
               porstMRpRs.AddNew
               porstMRpRs.Fields!CodEFi = porstMRp.Fields!CodEFi
               porstMRpRs.Fields!NroLin = porstMRp.Fields!NroLin
               porstMRpRs.Fields!TpoLin = porstMRp.Fields!TpoLin
               porstMRpRs.Fields!SubLin = "c"
               ' Se incluye line en blanco
               aImpSaldo(0) = aImpTotaLin(0)
               aImpSaldo(1) = aImpTotaLin(1)
               aImpTotaLin(0) = 0: aImpTotaLin(1) = 0
             Case 3         ' Formula
               aImpFormula(0) = 0: aImpFormula(1) = 0
               ResuelveFormula Trim$(porstMRp.Fields!FmlLin), aImpFormula
               aImpSaldo(0) = aImpFormula(0)
               aImpSaldo(1) = aImpFormula(1)
               aImpSubTotLin(0) = gfRedond(aImpSubTotLin(0) + aImpFormula(0), 2)
               aImpSubTotLin(1) = gfRedond(aImpSubTotLin(1) + aImpFormula(1), 2)
             Case 4         ' Mascara Formula
               aImpFormula(0) = 0: aImpFormula(1) = 0
               ResuelveFormula Trim$(porstMRp.Fields!FmlLin), aImpFormula
               aImpSaldo(0) = aImpFormula(0)
               aImpSaldo(1) = aImpFormula(1)
'            Case Else
            End Select
            porstMRpRs.AddNew
            porstMRpRs.Fields!CodEFi = porstMRp.Fields!CodEFi
            porstMRpRs.Fields!NroLin = porstMRp.Fields!NroLin
            porstMRpRs.Fields!DetLin = porstMRp.Fields!DetLin
            porstMRpRs.Fields!TpoLin = porstMRp.Fields!TpoLin
            porstMRpRs.Fields!SubLin = porstMRp.Fields!SubLin
            porstMRpRs.Fields!FmlLin = porstMRp.Fields!FmlLin
            porstMRpRs.Fields!BsePct = porstMRp.Fields!BsePct
            porstMRpRs.Fields!GrpPct = porstMRp.Fields!GrpPct
            porstMRpRs.Fields!ImpSaldoAnt = Format(aImpSaldo(0), "####0.00")
            porstMRpRs.Fields!ImpSaldoAct = Format(aImpSaldo(1), "####0.00")
            porstMRpRs.Fields!ImpPorceAnt = Format(porstMRp.Fields!BsePct * 100, "####0.00")
            porstMRpRs.Fields!ImpPorceAct = Format(porstMRp.Fields!BsePct * 100, "####0.00")
            aImpSaldo(0) = 0: aImpSaldo(1) = 0
            porstMRp.MoveNext
        Loop
        porstMRpRs.UpdateBatch
        pocnnMain.CommitTrans   '[ CONFIRMA TRANSACCION ]
    End If
    sCadExecute = "UPDATE trptREstFin a, trptREstFin b"
    sCadExecute = sCadExecute & " SET a.ImpPorceAnt=ROUND((ABS(a.ImpSaldoAnt) * b.ImpPorceAnt)/ABS(b.ImpSaldoAnt), 2),"
    sCadExecute = sCadExecute & " a.ImpPorceAct=ROUND((ABS(a.ImpSaldoAct) * b.ImpPorceAct)/ABS(b.ImpSaldoAct), 2)"
    sCadExecute = sCadExecute & " WHERE b.CodEfi=a.CodEfi"
    sCadExecute = sCadExecute & " AND b.BsePct=1 AND a.BsePct<>1"
    pocnnMain.Execute sCadExecute
    ' Elimino los registros que no se imprimen
    sCadExecute = "DELETE FROM trptREstFin WHERE (TpoLin='3' AND ROUND(ImpSaldoAnt+ImpSaldoAct, 2)=0.00) OR TpoLin='4'"
    pocnnMain.Execute sCadExecute

End Sub

Private Function ResuelveFormula(ByVal s_Cadena As String, ByRef aImpFormula) As Double
    
   Static sVariable As String, sSigno As String, sCaso As String
   Static nInicio As Integer, nFinal As Integer, nLen As Integer, nContador As Integer
   Static nImpDebe As Double, nImpHaber As Double, nImpSaldo(1) As Double
   Static aTipo(2) As String, aOperador(2) As String
    
   nInicio = 1: nFinal = 1: nLen = 0: nContador = 0
   nImpDebe = 0: nImpHaber = 0: nImpSaldo(0) = 0: nImpSaldo(1) = 0
   sCaso = Left(s_Cadena, 1)
   aOperador(0) = "+": aOperador(1) = "New"
   Do While nContador <= Len(s_Cadena)
     Select Case sCaso
       Case "["         ' Cuenta
         nInicio = (InStr(nInicio, s_Cadena, "[", vbTextCompare)) + 1
         nFinal = InStr(nInicio, s_Cadena, "]", vbTextCompare)
         nLen = (nFinal - nInicio)
         sVariable = Mid$(s_Cadena, nInicio, nLen)
         sSigno = Left(sVariable, 1)
         sVariable = IIf(IsNumeric(sSigno), sVariable, Mid(sVariable, 2))
         nImpDebe = 0: nImpHaber = 0
         nImpSaldo(0) = 0: nImpSaldo(1) = 0
         If porstCOCtaAcu.RecordCount > 0 Then porstCOCtaAcu.MoveFirst
         porstCOCtaAcu.Find "CodCta='" & sVariable & "'"
         If Not porstCOCtaAcu.EOF Then
            aTipo(0) = IIf(porstCOCtaAcu!NatCta = NATCTA_DEU, TPOCTB_DEB, TPOCTB_HAB)
            aTipo(1) = IIf(porstCOCtaAcu!NatCta = NATCTA_ACR, TPOCTB_DEB, TPOCTB_HAB)
            For nLen = 0 To Val(gsMesAct)
               nImpDebe = nImpDebe + porstCOCtaAcu("Acu" & aTipo(0) & Format(nLen, "00") & IIf(cboTpoMon.ListIndex = 0, "_MN", "_ME"))
               nImpHaber = nImpHaber + porstCOCtaAcu("Acu" & aTipo(1) & Format(nLen, "00") & IIf(cboTpoMon.ListIndex = 0, "_MN", "_ME"))
               If nLen = Val(gsMesAct) Then
                nImpSaldo(0) = gfRedond(porstCOCtaAcu("Acu" & aTipo(0) & Format(nLen, "00") & IIf(cboTpoMon.ListIndex = 0, "_MN", "_ME")) - porstCOCtaAcu("Acu" & aTipo(1) & Format(nLen, "00") & IIf(cboTpoMon.ListIndex = 0, "_MN", "_ME")), 2)
               End If
            Next nLen
            nImpSaldo(1) = gfRedond(nImpDebe - nImpHaber, 2)
         End If
         nImpSaldo(0) = IIf(sSigno = "-" And nImpSaldo(0) > 0, 0, IIf(sSigno = "+" And nImpSaldo(0) < 0, 0, nImpSaldo(0)))
         nImpSaldo(1) = IIf(sSigno = "-" And nImpSaldo(1) > 0, 0, IIf(sSigno = "+" And nImpSaldo(1) < 0, 0, nImpSaldo(1)))
         aOperador(1) = Mid(s_Cadena, nFinal + 1, 1)
         sCaso = "Cuenta"
         nContador = nFinal
       Case "{"         ' Linea
         nInicio = (InStr(nInicio, s_Cadena, "{", vbTextCompare)) + 1
         nFinal = InStr(nInicio, s_Cadena, "}", vbTextCompare)
         nLen = (nFinal - nInicio)
         sVariable = Mid$(s_Cadena, nInicio, nLen)
         nImpSaldo(0) = 0: nImpSaldo(1) = 0
         If porstMRpRs.RecordCount > 0 Then porstMRpRs.MoveFirst
         porstMRpRs.Find "NroLin='" & sVariable & "'"
         If Not porstMRpRs.EOF Then
            porstMRpRs.Move IIf(porstMRpRs!SubLin = "a", 2, IIf(porstMRpRs!SubLin = "c", 1, 0))
            nImpSaldo(0) = porstMRpRs!ImpSaldoAnt
            nImpSaldo(1) = porstMRpRs!ImpSaldoAct
         End If
         aOperador(1) = Mid(s_Cadena, nFinal + 1, 1)
         sCaso = "Linea"
         nContador = nFinal
       Case "("         ' Valores
         nInicio = (InStr(nInicio, s_Cadena, "(", vbTextCompare)) + 1
         nFinal = InStr(nInicio, s_Cadena, ")", vbTextCompare)
         nLen = (nFinal - nInicio)
         sVariable = Mid$(s_Cadena, nInicio, nLen)
         nImpSaldo(0) = CDec(Val(sVariable))
         nImpSaldo(1) = nImpSaldo(0)
         aOperador(1) = Mid(s_Cadena, nFinal + 1, 1)
         sCaso = "Valores"
         nContador = nFinal
       Case Else        ' Otro Caso
         Select Case aOperador(0)
           Case "+" ' Operación de suma
             aImpFormula(0) = gfRedond(aImpFormula(0) + nImpSaldo(0), 2)
             aImpFormula(1) = gfRedond(aImpFormula(1) + nImpSaldo(1), 2)
             sCaso = Mid(s_Cadena, nFinal + 2, 1)
             nContador = nFinal + 1
             aOperador(0) = aOperador(1)
           Case "-" ' Operación de resta
             aImpFormula(0) = gfRedond(aImpFormula(0) - nImpSaldo(0), 2)
             aImpFormula(1) = gfRedond(aImpFormula(1) - nImpSaldo(1), 2)
             sCaso = Mid(s_Cadena, nFinal + 2, 1)
             nContador = nFinal + 1
             aOperador(0) = aOperador(1)
           Case "*" ' Operación de multiplicación
             aImpFormula(0) = gfRedond(aImpFormula(0) * nImpSaldo(0), 2)
             aImpFormula(1) = gfRedond(aImpFormula(1) * nImpSaldo(1), 2)
             sCaso = Mid(s_Cadena, nFinal + 2, 1)
             nContador = nFinal + 1
             aOperador(0) = aOperador(1)
         End Select
     End Select
   Loop

End Function
' ma
']
