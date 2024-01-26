VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmREFi 
   Caption         =   "[título]"
   ClientHeight    =   6495
   ClientLeft      =   1935
   ClientTop       =   1290
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7815
   Begin MSFlexGridLib.MSFlexGrid mfgMain 
      Height          =   3450
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   6085
      _Version        =   393216
      BackColorFixed  =   16761024
      BackColorBkg    =   12632256
   End
   Begin VB.CheckBox chkTitulo 
      Caption         =   "Titulo Auxiliar"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6090
      TabIndex        =   24
      Top             =   4005
      Width           =   1335
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6090
      TabIndex        =   23
      Top             =   4365
      Width           =   1335
   End
   Begin VB.Frame frmTipoReporte 
      Caption         =   " Tipo de Reporte "
      ForeColor       =   &H00800000&
      Height          =   2000
      Left            =   45
      TabIndex        =   15
      Top             =   3840
      Width           =   4905
      Begin VB.OptionButton optProceso 
         Caption         =   "&Formato CONASEV (Estado Gan. Perd)"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   6
         Left            =   200
         TabIndex        =   22
         Top             =   1680
         Width           =   3250
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "&Resumen por meses"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   4
         Left            =   200
         TabIndex        =   21
         Top             =   1200
         Width           =   3250
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "&Formato CONASEV (Balance)"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   5
         Left            =   200
         TabIndex        =   20
         Top             =   1440
         Width           =   3250
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "&Mes / Acumulado (Año Actual)"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   2
         Left            =   200
         TabIndex        =   19
         Top             =   720
         Width           =   3250
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "&Año Anterior / Año Actual"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   3
         Left            =   200
         TabIndex        =   18
         Top             =   960
         Width           =   3250
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "&Formato General"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   0
         Left            =   200
         TabIndex        =   17
         Top             =   240
         Width           =   3250
      End
      Begin VB.OptionButton optProceso 
         Caption         =   "&Dos  monedas"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   1
         Left            =   200
         TabIndex        =   16
         Top             =   480
         Width           =   3250
      End
   End
   Begin VB.Frame frmSaldos 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   530
      Left            =   5130
      TabIndex        =   11
      Top             =   4680
      Width           =   2450
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
         Caption         =   "al mes"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   130
         TabIndex        =   13
         Top             =   200
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "del mes"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1215
         TabIndex        =   12
         Top             =   200
         Width           =   1080
      End
   End
   Begin VB.Frame fraImpresion 
      Caption         =   " Tipo de Impresora "
      ForeColor       =   &H80000002&
      Height          =   530
      Left            =   5130
      TabIndex        =   8
      Top             =   5310
      Width           =   2430
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "&Gráfica"
         ForeColor       =   &H80000001&
         Height          =   240
         Index           =   1
         Left            =   1185
         TabIndex        =   10
         Top             =   200
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "&Grafica"
         ForeColor       =   &H80000001&
         Height          =   240
         Index           =   0
         Left            =   130
         TabIndex        =   9
         Top             =   200
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   6330
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3600
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
      ScaleWidth      =   7815
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5955
      Width           =   7815
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
         Picture         =   "frmrefi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmrefi.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmrefi.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   5580
      TabIndex        =   7
      Top             =   3645
      Width           =   630
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
Private porstCOCCoAcu As ADODB.Recordset
Private porstCoCfg As ADODB.Recordset
Private psAnoOld As String
Private nLimite As Integer, aIndx(22) As Integer
Private nFilSele As Long
']

Private Sub Form_Load()
  Dim dnContador As Long
   
  '[ Verifico exitencia año anterior
  psAnoOld = Trim$(Val(gsAnoAct) - 1)
  ']
   
   'On Error GoTo Err
  
 '[Recordsets.                         'Cambiar.
    Set pocnnMain = New ADODB.Connection
    Set porstMRp = New ADODB.Recordset
    Set porstCOEFi = New ADODB.Recordset
    Set porstMRpRs = New ADODB.Recordset
    Set porstCOCtaAcu = New ADODB.Recordset
    Set porstCOCCoAcu = New ADODB.Recordset
    Set porstCoCfg = New ADODB.Recordset
    
    With pocnnMain
        .CursorLocation = adUseClient
        .ConnectionString = CONNSTRG & gsNomBDS
        .Open
    End With
    
    'Obtener simbolo de Moneda
    With porstCoCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT TpoMon_Sgn_MN, TpoMon_Sgn_ME "
      .Source = .Source & "FROM COCfg "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      gsTpoMon_Sgn_MN = .Fields(0)
      gsTpoMon_Sgn_ME = .Fields(1)
      .Close
    End With
    Set porstCoCfg.ActiveConnection = Nothing
    Set porstCoCfg = Nothing
    
    With porstCOCtaAcu
      .ActiveConnection = pocnnMain
      .Source = "SELECT a.*, b.NatCta, "
      .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.CodCta)", "(a.pdoano+a.CodCta)") & " AS cLlave "
      .Source = .Source & "FROM COCtaAcu a "
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano>='" & psAnoOld & "' "
      .Source = .Source & "AND a.pdoano<='" & gsAnoAct & "' "
      .Source = .Source & "ORDER BY a.pdoano, a.CodCta"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
    End With
    
    With porstCOCCoAcu
      .ActiveConnection = pocnnMain
      .Source = "SELECT a.*, b.NatCta, "
      .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.CodCta, a.CodCCo)", "(a.pdoano+a.CodCta+a.CodCCo)") & " AS cLlave "
      .Source = .Source & "FROM COCCoAcu a "
      .Source = .Source & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta "
      .Source = .Source & "LEFT JOIN CoCCo c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCCo=c.CodCCo "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano>='" & psAnoOld & "' "
      .Source = .Source & "AND a.pdoano<='" & gsAnoAct & "' "
      .Source = .Source & "ORDER BY a.pdoano, a.CodCta, a.CodCCo"
'     .CursorLocation = adUseClient   'Es el Default.
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
        .Source = "SELECT efi.codefi, efi." & Choose(gsIdioma, "detefi", "detefix") & " AS detefi, efi.coddpe, "
        If ps_Plataforma = pSrvMySql Then
          .Source = .Source & "Concat(dpe.coddpe," & Choose(gsIdioma, "dpe.detdpe", "dpe.detdpex") & ") AS despro "
        Else
          .Source = .Source & "Concat(dpe.coddpe+" & Choose(gsIdioma, "dpe.detdpe", "dpe.detdpex") & ") AS despro "
        End If
        .Source = .Source & "FROM coefi efi "
        .Source = .Source & "LEFT JOIN codpe dpe ON efi.codemp=dpe.codemp AND efi.pdoano='" & gsAnoAct & "' AND efi.coddpe=dpe.coddpe "
        .Source = .Source & "WHERE efi.codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND efi.pdoano='" & gsAnoAct & "' "
        .Source = .Source & "ORDER BY efi.coddpe, efi.codefi "
'        .CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
    End With
    With porstMRpRs
        .ActiveConnection = pocnnMain
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Source = "SELECT * "
        .Source = .Source & "FROM " & IIf(ps_Plataforma = pSrvSql, "#", "") & "trptREstFin "
        .Source = .Source & "ORDER BY NroLin"
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
    
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :")
  Next nElemento
  frmTipoReporte.Caption = Choose(gsIdioma, "Tipo de Reporte", "Type of Report")
  optProceso(0).Caption = Choose(gsIdioma, "&Formato General", "General &Format")
  optProceso(1).Caption = Choose(gsIdioma, "&Dos monedas", "Two Currencies")
  optProceso(2).Caption = Choose(gsIdioma, "&Mes / Acumulado (Año Actual)", "&Month / Accrued (Actual Year)")
  optProceso(3).Caption = Choose(gsIdioma, "&Año Anterior / Año Actual", "&Last Year / Actual Year")
  optProceso(4).Caption = Choose(gsIdioma, "&Resumen por meses", "&Summary for months")
  optProceso(5).Caption = Choose(gsIdioma, "Formato CONASEV (&Balance)", "Format CONASEV (&Balance)")
  optProceso(6).Caption = Choose(gsIdioma, "Formato CONASEV (&Estado Gan. Perd)", "Format CONASEV (&Profit Lost Statement)")
  chkTitulo.Caption = Choose(gsIdioma, "Titulo Auxiliar", "Auxiliary Title")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  frmSaldos.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "al mes", "to month")
  OptTipo(1).Caption = Choose(gsIdioma, "del mes", "from month")
  fraImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
  ']
 
  '[Datos predeterminados.              'Cambiar.
  optProceso(0).Value = True
  chkImpFecha.Value = Checked
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
 
   ' Inicializo la grilla
  mfgMain.Clear
  ppDatosGrid
  nFilSele = 1
  mfgMain.Rows = nFilSele
  While Not porstCOEFi.EOF
    dnContador = mfgMain.Rows
    With mfgMain
      .AddItem ""
      .TextMatrix(dnContador, 1) = IIf(IsNull(porstCOEFi!despro), "", porstCOEFi!despro)
      .TextMatrix(dnContador, 2) = porstCOEFi!CodEfi
      .TextMatrix(dnContador, 3) = porstCOEFi!DetEFi
    End With
    porstCOEFi.MoveNext
  Wend
  
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
   Set porstCOCCoAcu = Nothing
   Set porstCOEFi = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim sReporte As String, sTitulo As String
  Dim sNombreMes As String, sDias As String
  Dim sSubTitulo As String, sConversion As String
  
   ppHabilitacion True
  If porstCOEFi.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub
  sConversion = "CONVERT(" & IIf(ps_Plataforma = pSrvMySql, "0, decimal(18, 2)", "decimal(18, 2), 0") & ")"
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT CodEfi, NroLin, " & Choose(gsIdioma, "DetLin", "DetLinx") & " AS DetLin, TpoLin, 'b' AS SubLin, FmlLin, BsePct, GrpPct, "
    .Source = .Source & "IndBdeSup, IndBdeInf, IndFonDet, IndFonDet_Syd, IndFonImp, "
    .Source = .Source & sConversion & " AS ImpSaldoIni, " & sConversion & " AS ImpPorceIni, "
    .Source = .Source & sConversion & " AS ImpSaldoMes, " & sConversion & " AS ImpPorceMes, "
    .Source = .Source & sConversion & " AS ImpSaldoAcu, " & sConversion & " AS ImpPorceAcu, "
    .Source = .Source & sConversion & " AS ImpSalIniMN, " & sConversion & " AS ImpPorIniMN, "
    .Source = .Source & sConversion & " AS ImpSalMesMN, " & sConversion & " AS ImpPorMesMN, "
    .Source = .Source & sConversion & " AS ImpSalAcuMN, " & sConversion & " AS ImpPorAcuMN, "
    .Source = .Source & sConversion & " AS ImpSalIniME, " & sConversion & " AS ImpPorIniME, "
    .Source = .Source & sConversion & " AS ImpSalMesME, " & sConversion & " AS ImpPorMesME, "
    .Source = .Source & sConversion & " AS ImpSalAcuME, " & sConversion & " AS ImpPorAcuME, "
    .Source = .Source & sConversion & " AS ImpSaldo_00, " & sConversion & " AS ImpSaldo_01, "
    .Source = .Source & sConversion & " AS ImpSaldo_02, " & sConversion & " AS ImpSaldo_03, "
    .Source = .Source & sConversion & " AS ImpSaldo_04, " & sConversion & " AS ImpSaldo_05, "
    .Source = .Source & sConversion & " AS ImpSaldo_06, " & sConversion & " AS ImpSaldo_07, "
    .Source = .Source & sConversion & " AS ImpSaldo_08, " & sConversion & " AS ImpSaldo_09, "
    .Source = .Source & sConversion & " AS ImpSaldo_10, " & sConversion & " AS ImpSaldo_11, "
    .Source = .Source & sConversion & " AS ImpSaldo_12, " & sConversion & " AS ImpSaldo_13 "
    .Source = .Source & "FROM CoEFiLin "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND CodEfi='" & mfgMain.TextMatrix(nFilSele, 2) & "' "
    .Source = .Source & "ORDER BY NroLin"
    .Open
  End With
  
  '[Reporte, titulo y indices
  sReporte = "rptREFi": sTitulo = IIf(IsNull(mfgMain.TextMatrix(nFilSele, 3)), "", mfgMain.TextMatrix(nFilSele, 3))
  aIndx(0) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 0, 3)
  aIndx(3) = 0: aIndx(4) = 1: aIndx(5) = 2
  aIndx(6) = 3: aIndx(7) = 4: aIndx(8) = 5
  For nLimite = 0 To Val(gsMesAct)
    aIndx(9 + nLimite) = IIf(OptTipo(0).Value, Choose(cboTpoMon.ListIndex + 1, (12 + ((4 * nLimite) + 1)), (12 + ((4 * nLimite) + 3))), Choose(cboTpoMon.ListIndex + 1, (12 + (4 * nLimite)), (12 + ((4 * nLimite) + 2))))
  Next nLimite
  nLimite = 11
  If optProceso(0).Value Then
    sReporte = sReporte & "1" & IIf(OptTipo(0).Value, "a", "m")
    sTitulo = sTitulo & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")"
    aIndx(1) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, 4)
    aIndx(2) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 2, 5)
  ElseIf optProceso(1).Value Then
    OptTipo(0).Value = True
    sReporte = sReporte & "2g"
    aIndx(0) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 0, 3)
    aIndx(1) = IIf(OptTipo(0).Value, 2, 1)
    aIndx(2) = IIf(OptTipo(0).Value, 5, 4)
  ElseIf optProceso(2).Value Or optProceso(3).Value Then
    sReporte = sReporte & "3g"
    sTitulo = sTitulo & IIf(optProceso(3).Value, Choose(gsIdioma, " Comparativo", " Compared"), "")
    aIndx(4) = IIf(optProceso(3).Value, IIf(OptTipo(0).Value, 8, 7), 1)
    aIndx(5) = IIf(optProceso(3).Value, IIf(OptTipo(0).Value, 2, 1), 2)
    aIndx(7) = IIf(optProceso(3).Value, IIf(OptTipo(0).Value, 11, 10), 4)
    aIndx(8) = IIf(optProceso(3).Value, IIf(OptTipo(0).Value, 5, 4), 5)
  ElseIf optProceso(4).Value Then
    sReporte = sReporte & "4" & IIf(OptTipo(0).Value, "a", "m")
    sTitulo = sTitulo & IIf(OptTipo(0).Value, Choose(gsIdioma, " - Acumulado", " - Accrued"), Choose(gsIdioma, " - Mensual", " - Monthly")) & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")"
    nLimite = nLimite + ((Val(gsMesAct) + 1) * 4)
  ElseIf optProceso(5).Value Then
    sReporte = sReporte & "5g"
    sTitulo = sTitulo & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")"
    aIndx(1) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, 4)
    aIndx(2) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 2, 5)
  ElseIf optProceso(6).Value Then
    sReporte = sReporte & "6g"
    sTitulo = sTitulo & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")"
    aIndx(0) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 7, 10)
    aIndx(1) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, 4)
    aIndx(2) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 2, 5)
    aIndx(3) = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 8, 11)
  End If
  
  Llena_Temporal porstMRp.Source
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT * FROM " & ps_Prefijo & "trptrestfin"
    .Open
  End With
  ']
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_GRAF, PRN_DEST_MATR)
  If gsIdioma = NvlUsr_Sup Then
    sNombreMes = Choose(Val(gsMesAct) + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
  Else
    sNombreMes = Choose(Val(gsMesAct) + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
  End If
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, sTitulo, udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & sReporte & ".rpt"
      ' .WindowShowGroupTree = True
      ']
      '[ Formulas adicionales
      If optProceso(0).Value Or optProceso(1).Value Then
        sDias = Left(gfUltDia("01/" & gsMesAct & "/" & gsAnoAct), 2)
        sDias = sDias & a_Sufijo(sDias)
        sSubTitulo = IIf(OptTipo(0).Value, Choose(gsIdioma, "AL " & sDias, "TO " & UCase(sNombreMes)) & Choose(gsIdioma, "  DE " & UCase(sNombreMes), " " & sDias & ","), Choose(gsIdioma, "DEL MES DE ", "FROM MONTH OF ") & UCase(sNombreMes)) & Choose(gsIdioma, " DEL ", " ") & gsAnoAct
        .Formulas(5) = "mPeriodo='" & sSubTitulo & "'"
        
        If chkTitulo.Value = vbChecked Then
          If gsIdioma = NvlUsr_Sup Then
            sSubTitulo = "POR EL PERIODO TERMINADO " & IIf(OptTipo(0).Value, "EL " & sDias & "  DE ", "DEL MES DE ") & UCase(sNombreMes) & " DEL " & gsAnoAct
          Else
            sSubTitulo = "FOR THE FINISHED PERIOD " & IIf(OptTipo(0).Value, "TO ", "FROM MONTH OF ") & UCase(sNombreMes) & IIf(OptTipo(0).Value, " " & sDias & ", ", " ") & gsAnoAct
          End If
          .Formulas(5) = "mPeriodo='" & sSubTitulo & "'"
        End If
        
        If sReporte = "rptREFi1a" Then
          .Formulas(7) = "RepLegal='" & gsRepEmp & "'"
          .Formulas(8) = "Contador='" & gsConEmp & "'"
        End If
        
      ElseIf (optProceso(2).Value Or optProceso(3).Value) Then
        If gsIdioma = NvlUsr_Sup Then
          .Formulas(7) = "mTituloCola='" & IIf(optProceso(3).Value, IIf(OptTipo(0).Value, "AL ", "") & Left(gfUltDia("01/" & gsMesAct & "/" & psAnoOld), 2), "DEL MES") & "  DE " & UCase(Left(sNombreMes, 3)) & ".'"
          .Formulas(8) = "mTituloColb='" & IIf(optProceso(3).Value, IIf(OptTipo(0).Value, "AL ", "") & Left(gfUltDia("01/" & gsMesAct & "/" & gsAnoAct), 2), "AL MES") & "  DE " & UCase(Left(sNombreMes, 3)) & ".'"
        Else
          sDias = Left(gfUltDia("01/" & gsMesAct & "/" & psAnoOld), 2)
          sDias = sDias & a_Sufijo(sDias)
          .Formulas(7) = "mTituloCola='" & IIf(optProceso(3).Value, IIf(OptTipo(0).Value, "TO ", ""), "FROM MONTH ") & UCase(Left(sNombreMes, 3)) & "." & IIf(optProceso(3).Value, " " & sDias, "") & "'"
          sDias = Left(gfUltDia("01/" & gsMesAct & "/" & gsAnoAct), 2)
          sDias = sDias & a_Sufijo(sDias)
          .Formulas(8) = "mTituloColb='" & IIf(optProceso(3).Value, IIf(OptTipo(0).Value, "TO ", ""), "TO MONTH ") & UCase(Left(sNombreMes, 3)) & "." & IIf(optProceso(3).Value, " " & sDias, "") & "'"
        End If
        .Formulas(9) = "mTituloColc='" & Choose(gsIdioma, "DE ", "") & IIf(optProceso(2).Value, gsAnoAct, psAnoOld) & "'"
        .Formulas(10) = "mTituloCold='" & Choose(gsIdioma, "DE ", "") & gsAnoAct & "'"
      ElseIf optProceso(5).Value Then
        sDias = Left(gfUltDia("01/" & gsMesAct & "/" & Trim$(Val(gsAnoAct) + 1)), 2)
        sDias = sDias & a_Sufijo(sDias)
        .Formulas(7) = "mTituloCola='" & Choose(gsIdioma, "Al ", "To ") & Choose(gsIdioma, sDias & " de " & Left(sNombreMes, 3) & ".", Left(sNombreMes, 3) & ". " & sDias) & "'"
        .Formulas(8) = "mTituloColb='" & Choose(gsIdioma, "Al 31 de Dic.", "To Dec. 31st") & "'"
        .Formulas(9) = "mTituloColc='" & Choose(gsIdioma, "de ", "") & gsAnoAct & "'"
        .Formulas(10) = "mTituloCold='" & Choose(gsIdioma, "de ", "") & psAnoOld & "'"
      ElseIf optProceso(6).Value Then
        sDias = Left(gfUltDia("01/" & gsMesAct & "/" & Trim$(Val(gsAnoAct) + 1)), 2)
        sDias = sDias & a_Sufijo(sDias)
        .Formulas(7) = "mTituloCola='" & Choose(gsIdioma, "Del 01 de ", "From ") & Left(sNombreMes, 3) & Choose(gsIdioma, ". Al ", ". 01st To") & "'"
        .Formulas(8) = "mTituloColb='" & Choose(gsIdioma, "Del 01 de Ene. Al", "From Jan. 01st To") & "'"
        .Formulas(9) = "mTituloColc='" & Choose(gsIdioma, sDias & " de ", "") & Left(sNombreMes, 3) & Choose(gsIdioma, ". de ", ". " & sDias & ", ") & gsAnoAct & "'"
        .Formulas(10) = "mTituloCold='" & Choose(gsIdioma, sDias & " de ", "") & Left(sNombreMes, 3) & Choose(gsIdioma, ". de ", ". " & sDias & ", ") & psAnoOld & "'"
      End If
      If chkTitulo.Value = vbChecked Then
        .Formulas(0) = "mSistema=''"
      End If
      
      ']
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  End If
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptREstFin", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#trptREstFin') DROP TABLE #trptREstFin")
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
  Dim nIndice As Integer
  
  With mfgMain
    .cols = 4
    .FixedCols = 1
    .Rows = .Rows
    .FixedRows = IIf(mfgMain.row = 0, mfgMain.row, 1)
    .GridColor = vbBlack
    .GridColorFixed = vbBlue
    .GridLines = flexGridFlat
    .GridLinesFixed = flexGridInset
    .GridLineWidth = 1
    .SelectionMode = flexSelectionFree
    .BackColor = &H80000018
    .BackColorBkg = &H8000000F
    .BackColorFixed = &HFFC0C0
    .BackColorSel = &HE0E0E0
    .ForeColor = vbBlack
    .ForeColorFixed = vbWhite
    .ForeColorSel = vbBlue
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusHeavy
    .MergeCells = flexMergeRestrictColumns
    .MergeCol(1) = True
    .Font.Bold = False
  End With
  nIndice = mfgMain.row
  mfgMain.row = 0
  mfgMain.Col = 1: mfgMain.CellFontBold = True
  mfgMain.Col = 2: mfgMain.CellFontBold = True
  mfgMain.Col = 3: mfgMain.CellFontBold = True
  mfgMain.row = nIndice
  For nIndice = 0 To (mfgMain.cols - 1)
    mfgMain.Col = nIndice
    If gsIdioma = NvlUsr_Sup Then
      mfgMain.TextMatrix(0, nIndice) = Choose(nIndice + 1, "", "Proyecto", "Codigo", "Descripción")
    Else
      mfgMain.TextMatrix(0, nIndice) = Choose(nIndice + 1, "", "Project", "Code", "Description")
    End If
    mfgMain.ColAlignment(nIndice) = Choose(nIndice + 1, flexAlignLeftCenter, flexAlignLeftCenter, flexAlignLeftCenter, flexAlignLeftCenter)
    mfgMain.ColWidth(nIndice) = Choose(nIndice + 1, 300, 2010, 800, 4250)
  Next nIndice
  
End Sub

']
' ma
Private Sub Llena_Temporal(ByVal s_Source As String)

    Static aImpSubTotLin(67) As Double, aImpTotaLin(67) As Double
    Static aImpSaldo(67) As Double, aImpFormula(67) As Double
    Static sCadExecute As String, nContador As Integer
    
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptREstFin", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 12)='#trptREstFin') DROP TABLE #trptREstFin")
    s_Source = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptREstFin (", "CREATE TABLE #trptREstFin (")
    s_Source = s_Source & "CodEfi char(2) NOT NULL, NroLin varchar(4) NOT NULL, "
    s_Source = s_Source & "DetLin varchar(50) NULL, TpoLin char(1) NULL, "
    s_Source = s_Source & "SubLin char(1) DEFAULT 'b', FmlLin varchar(355) NULL, "
    s_Source = s_Source & "BsePct smallint DEFAULT 0, GrpPct char(1) NULL, "
    s_Source = s_Source & "IndBdeSup char(1) DEFAULT '0', IndBdeInf char(1) DEFAULT '0', "
    s_Source = s_Source & "IndFonDet smallint DEFAULT 0, IndFonDet_Syd smallint DEFAULT 0, "
    s_Source = s_Source & "IndFonImp smallint DEFAULT 0, "
    s_Source = s_Source & "ImpSaldoIni decimal(18, 2) DEFAULT 0.00, ImpPorceIni decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldoMes decimal(18, 2) DEFAULT 0.00, ImpPorceMes decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldoAcu decimal(18, 2) DEFAULT 0.00, ImpPorceAcu decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSalIniMN decimal(18, 2) DEFAULT 0.00, ImpPorIniMN decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSalMesMN decimal(18, 2) DEFAULT 0.00, ImpPorMesMN decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSalAcuMN decimal(18, 2) DEFAULT 0.00, ImpPorAcuMN decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSalIniME decimal(18, 2) DEFAULT 0.00, ImpPorIniME decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSalMesME decimal(18, 2) DEFAULT 0.00, ImpPorMesME decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSalAcuME decimal(18, 2) DEFAULT 0.00, ImpPorAcuME decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldo_00 decimal(18, 2) DEFAULT 0.00, ImpSaldo_01 decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldo_02 decimal(18, 2) DEFAULT 0.00, ImpSaldo_03 decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldo_04 decimal(18, 2) DEFAULT 0.00, ImpSaldo_05 decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldo_06 decimal(18, 2) DEFAULT 0.00, ImpSaldo_07 decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldo_08 decimal(18, 2) DEFAULT 0.00, ImpSaldo_09 decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldo_10 decimal(18, 2) DEFAULT 0.00, ImpSaldo_11 decimal(18, 2) DEFAULT 0.00, "
    s_Source = s_Source & "ImpSaldo_12 decimal(18, 2) DEFAULT 0.00, ImpSaldo_13 decimal(18, 2) DEFAULT 0.00)"
    pocnnMain.Execute s_Source
    pocnnMain.Execute "DELETE FROM " & ps_Prefijo & "trptREstFin"
        
    porstMRpRs.Open
    If porstMRp.RecordCount > 0 Then
        porstMRp.MoveFirst
        ' Inicializo las variables de importe
        For nContador = 0 To nLimite
          aImpSubTotLin(nContador) = 0: aImpTotaLin(nContador) = 0
          aImpSaldo(nContador) = 0: aImpFormula(nContador) = 0
        Next nContador
        pocnnMain.BeginTrans    '[ INICIA TRANSACCION ]
        Do While Not porstMRp.EOF
            Select Case porstMRp.Fields!TpoLin
             Case 1         ' Total
               For nContador = 1 To 2
                porstMRpRs.AddNew
                porstMRpRs.Fields!CodEfi = porstMRp.Fields!CodEfi
                porstMRpRs.Fields!NroLin = porstMRp.Fields!NroLin
                porstMRpRs.Fields!TpoLin = porstMRp.Fields!TpoLin
                porstMRpRs.Fields!SubLin = Choose(nContador, "a", "c")
                ' Se incluye line en blanco
               Next nContador
               For nContador = 0 To nLimite
                 aImpTotaLin(nContador) = Round(aImpTotaLin(nContador) + aImpSubTotLin(nContador), 2)
                 aImpSaldo(nContador) = aImpSubTotLin(nContador)
                 aImpSubTotLin(nContador) = 0
               Next nContador
             Case 2         ' Sub Total
               porstMRpRs.AddNew
               porstMRpRs.Fields!CodEfi = porstMRp.Fields!CodEfi
               porstMRpRs.Fields!NroLin = porstMRp.Fields!NroLin
               porstMRpRs.Fields!TpoLin = porstMRp.Fields!TpoLin
               porstMRpRs.Fields!SubLin = "c"
               ' Se incluye line en blanco
               For nContador = 0 To nLimite
                 aImpSaldo(nContador) = aImpTotaLin(nContador)
                 aImpTotaLin(nContador) = 0
               Next nContador
             Case 3         ' Formula
               For nContador = 0 To nLimite: aImpFormula(nContador) = 0: Next nContador
               ResuelveFormula IIf(optProceso(3).Value, 1, 0), Trim$(IIf(IsNull(porstMRp.Fields!FmlLin), "", porstMRp.Fields!FmlLin)), aImpFormula
               For nContador = 0 To nLimite
                 'aImpSaldo(nContador) = IIf(porstMRp.Fields!TpoLin = 3, aImpFormula(nContador), 0)
                 aImpSaldo(nContador) = aImpFormula(nContador)
                 aImpSubTotLin(nContador) = Round(aImpSubTotLin(nContador) + aImpFormula(nContador), 2)
               Next nContador
             Case 4          ' Formula
               For nContador = 0 To nLimite: aImpFormula(nContador) = 0: Next nContador
               ResuelveFormula IIf(optProceso(3).Value, 1, 0), Trim$(IIf(IsNull(porstMRp.Fields!FmlLin), "", porstMRp.Fields!FmlLin)), aImpFormula
               For nContador = 0 To nLimite
                 aImpSaldo(nContador) = aImpFormula(nContador)
                 'aImpSubTotLin(nContador) = round(aImpSubTotLin(nContador) + aImpFormula(nContador), 2)
               Next nContador
               
'            Case Else
            End Select
            porstMRpRs.AddNew
            porstMRpRs.Fields!CodEfi = porstMRp.Fields!CodEfi
            porstMRpRs.Fields!NroLin = porstMRp.Fields!NroLin
            porstMRpRs.Fields!DetLin = porstMRp.Fields!DetLin
            porstMRpRs.Fields!TpoLin = porstMRp.Fields!TpoLin
            porstMRpRs.Fields!SubLin = porstMRp.Fields!SubLin
            porstMRpRs.Fields!FmlLin = porstMRp.Fields!FmlLin
            porstMRpRs.Fields!BsePct = porstMRp.Fields!BsePct
            porstMRpRs.Fields!grppct = porstMRp.Fields!grppct
            porstMRpRs.Fields!IndBdeSup = porstMRp.Fields!IndBdeSup
            porstMRpRs.Fields!IndBdeInf = porstMRp.Fields!IndBdeInf
            porstMRpRs.Fields!IndFonDet = porstMRp.Fields!IndFonDet
            porstMRpRs.Fields!IndFonDet_Syd = porstMRp.Fields!IndFonDet_Syd
            porstMRpRs.Fields!IndFonImp = porstMRp.Fields!IndFonImp
            porstMRpRs.Fields!ImpSaldoIni = CDec(aImpSaldo(aIndx(0)))
            porstMRpRs.Fields!ImpSaldoMes = CDec(aImpSaldo(aIndx(1)))
            porstMRpRs.Fields!ImpSaldoAcu = CDec(aImpSaldo(aIndx(2)))
            porstMRpRs.Fields!ImpPorceIni = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpPorceMes = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpPorceAcu = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpSalIniMN = CDec(aImpSaldo(aIndx(3)))
            porstMRpRs.Fields!ImpSalMesMN = CDec(aImpSaldo(aIndx(4)))
            porstMRpRs.Fields!ImpSalAcuMN = CDec(aImpSaldo(aIndx(5)))
            porstMRpRs.Fields!ImpPorIniMN = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpPorMesMN = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpPorAcuMN = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpSalIniME = CDec(aImpSaldo(aIndx(6)))
            porstMRpRs.Fields!ImpSalMesME = CDec(aImpSaldo(aIndx(7)))
            porstMRpRs.Fields!ImpSalAcuME = CDec(aImpSaldo(aIndx(8)))
            porstMRpRs.Fields!ImpPorIniME = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpPorMesME = CDec(porstMRp.Fields!BsePct * 100)
            porstMRpRs.Fields!ImpPorAcuME = CDec(porstMRp.Fields!BsePct * 100)
            For nContador = 0 To Val(gsMesAct)
              porstMRpRs.Fields("ImpSaldo_" & Format(nContador, "00")) = CDec(aImpSaldo(aIndx(9 + nContador)))
            Next
            For nContador = 0 To nLimite: aImpSaldo(nContador) = 0: Next nContador
            porstMRp.MoveNext
        Loop
        porstMRpRs.UpdateBatch
        pocnnMain.CommitTrans   '[ CONFIRMA TRANSACCION ]
    End If
    ' Obtengo la base de porcentaje
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS basegrupo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 10)='#basegrupo') DROP TABLE #basegrupo")
    sCadExecute = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS basegrupo ", "")
    sCadExecute = sCadExecute & "SELECT DISTINCT grppct, bsepct, "
    sCadExecute = sCadExecute & "ImpPorceIni, ImpSaldoIni, ImpPorceMes, ImpSaldoMes, "
    sCadExecute = sCadExecute & "ImpPorceAcu, ImpSaldoAcu, ImpPorIniMN, ImpSalIniMN, "
    sCadExecute = sCadExecute & "ImpPorMesMN, ImpSalMesMN, ImpPorAcuMN, ImpSalAcuMN, "
    sCadExecute = sCadExecute & "ImpPorIniME, ImpSalIniME, ImpPorMesME, ImpSalMesME, "
    sCadExecute = sCadExecute & "ImpPorAcuME, ImpSalAcuME "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "", "INTO #basegrupo ")
    sCadExecute = sCadExecute & "FROM " & ps_Prefijo & "trptREstFin "
    sCadExecute = sCadExecute & "WHERE bsepct=1 "
    sCadExecute = sCadExecute & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(grppct, '')<>''"
    pocnnMain.Execute sCadExecute

    ' Actualizo importes y porcentajes de las lineas
    sCadExecute = IIf(ps_Plataforma = pSrvMySql, "UPDATE trptREstFin a, basegrupo b ", "UPDATE #trptREstFin ")
    sCadExecute = sCadExecute & "SET "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorceIni=ROUND((a.ImpSaldoIni * b.ImpPorceIni)/b.ImpSaldoIni, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorceMes=ROUND((a.ImpSaldoMes * b.ImpPorceMes)/b.ImpSaldoMes, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorceAcu=ROUND((a.ImpSaldoAcu * b.ImpPorceAcu)/b.ImpSaldoAcu, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorIniMN=ROUND((a.ImpSalIniMN * b.ImpPorIniMN)/b.ImpSalIniMN, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorMesMN=ROUND((a.ImpSalMesMN * b.ImpPorMesMN)/b.ImpSalMesMN, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorAcuMN=ROUND((a.ImpSalAcuMN * b.ImpPorAcuMN)/b.ImpSalAcuMN, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorIniME=ROUND((a.ImpSalIniME * b.ImpPorIniME)/b.ImpSalIniME, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorMesME=ROUND((a.ImpSalMesME * b.ImpPorMesME)/b.ImpSalMesME, 2), "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "a.", "") & "ImpPorAcuME=ROUND((a.ImpSalAcuME * b.ImpPorAcuME)/b.ImpSalAcuME, 2) "
    sCadExecute = sCadExecute & IIf(ps_Plataforma = pSrvMySql, "", "FROM #trptREstFin a, #basegrupo b ")
    sCadExecute = sCadExecute & "WHERE b.grppct=a.grppct "
    sCadExecute = sCadExecute & "AND a.BsePct<>1"
    pocnnMain.Execute sCadExecute
    
    ' Elimino la tabla temporal de porcentaje
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS basegrupo", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 10)='#basegrupo') DROP TABLE #basegrupo")
    
    ' Elimino las lineas con importes en cero !(CONASEV)
    sCadExecute = "DELETE FROM " & ps_Prefijo & "trptREstFin "
    sCadExecute = sCadExecute & "WHERE TpoLin='4' "
    If Not (optProceso(5).Value Or optProceso(5).Value) Then
      If Not (optProceso(0).Value) Then
         sCadExecute = sCadExecute & "OR (TpoLin='" & TPOLIN_OPE & "' "
         sCadExecute = sCadExecute & "AND ROUND((ImpSaldoIni + ImpSalIniMN + ImpSalIniME + "
      Else
         sCadExecute = sCadExecute & "OR (TpoLin='" & TPOLIN_OPE & "' "
         sCadExecute = sCadExecute & "AND ROUND(("
      End If
      If OptTipo(0).Value Then
        sCadExecute = sCadExecute & "ImpSaldoAcu + ImpSalAcuMN + ImpSalAcuME"
      Else
        sCadExecute = sCadExecute & "ImpSaldoMes + ImpSalMesMN + ImpSalMesME"
      End If
      For nContador = 0 To Val(gsMesAct)
        sCadExecute = sCadExecute & " + ImpSaldo_" & Format(nContador, "00")
      Next
      sCadExecute = sCadExecute & "), 2) = 0.00)"
    End If
    pocnnMain.Execute sCadExecute

End Sub

Private Function ResuelveFormula(ByVal nAnos As Integer, ByVal s_Cadena As String, ByRef aImpFormula) As Double
    
   Static sVariable As String, sSigno As String, sCaso As String
   Static nInicio As Integer, nFinal As Integer, nLen As Integer, nContador As Integer
   Static nImpDebe(1) As Double, nImpHaber(1) As Double, nImpSaldo(67) As Double
   Static aTipo(2) As String, aOperador(2) As String
   Static nIndex As Integer, nRngSaldos As Integer
    
   nInicio = 1: nFinal = 1: nLen = 0: nContador = 0
   For nIndex = 0 To 1: nImpDebe(nIndex) = 0: nImpHaber(nIndex) = 0: Next nIndex
   For nIndex = 0 To nLimite: nImpSaldo(nIndex) = 0: Next nIndex
   
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
         ' Inicializa los saldos
         For nIndex = 0 To nLimite: nImpSaldo(nIndex) = 0: Next nIndex
         ' Selelccion segun años
         For nIndex = 0 To nAnos
           nRngSaldos = (nIndex * 6)
           nImpDebe(0) = 0: nImpHaber(0) = 0: nImpDebe(1) = 0: nImpHaber(1) = 0
           ' Busco la cuenta
           If porstCOCtaAcu.RecordCount > 0 Then porstCOCtaAcu.MoveFirst
           porstCOCtaAcu.Find "cLlave='" & Choose(nIndex + 1, gsAnoAct, psAnoOld) & sVariable & "'"
           If Not porstCOCtaAcu.EOF Then
             aTipo(0) = IIf(porstCOCtaAcu!NatCta = NATCTA_DEU, TPOCTB_DEB, TPOCTB_HAB)
             aTipo(1) = IIf(porstCOCtaAcu!NatCta = NATCTA_ACR, TPOCTB_DEB, TPOCTB_HAB)
             ' Acumulo los importes hasta el mes actual
             For nLen = 0 To Val(gsMesAct)
               nImpDebe(0) = nImpDebe(0) + porstCOCtaAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT)
               nImpHaber(0) = nImpHaber(0) + porstCOCtaAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT)
               nImpDebe(1) = nImpDebe(1) + porstCOCtaAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT)
               nImpHaber(1) = nImpHaber(1) + porstCOCtaAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT)
               ' Saldos por meses y acumulado
               If nIndex = 0 Then
                 nImpSaldo(12 + (4 * nLen)) = Round(porstCOCtaAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT) - porstCOCtaAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT), 2)
                 nImpSaldo(12 + ((4 * nLen) + 1)) = Round(nImpDebe(0) - nImpHaber(0), 2)
                 nImpSaldo(12 + ((4 * nLen) + 2)) = Round(porstCOCtaAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT) - porstCOCtaAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT), 2)
                 nImpSaldo(12 + ((4 * nLen) + 3)) = Round(nImpDebe(1) - nImpHaber(1), 2)
               End If
             Next nLen
             ' Saldo inicial del año
             nImpSaldo(nRngSaldos) = Round(porstCOCtaAcu("Acu" & aTipo(0) & "00_" & TPOMON_NAC_TXT) - porstCOCtaAcu("Acu" & aTipo(1) & "00_" & TPOMON_NAC_TXT), 2)
             nImpSaldo(nRngSaldos + 3) = Round(porstCOCtaAcu("Acu" & aTipo(0) & "00_" & TPOMON_EXT_TXT) - porstCOCtaAcu("Acu" & aTipo(1) & "00_" & TPOMON_EXT_TXT), 2)
             ' Saldo del mes
             nImpSaldo(nRngSaldos + 1) = Round(porstCOCtaAcu("Acu" & aTipo(0) & gsMesAct & "_" & TPOMON_NAC_TXT) - porstCOCtaAcu("Acu" & aTipo(1) & gsMesAct & "_" & TPOMON_NAC_TXT), 2)
             nImpSaldo(nRngSaldos + 4) = Round(porstCOCtaAcu("Acu" & aTipo(0) & gsMesAct & "_" & TPOMON_EXT_TXT) - porstCOCtaAcu("Acu" & aTipo(1) & gsMesAct & "_" & TPOMON_EXT_TXT), 2)
             ' Saldo acumulado
             nImpSaldo(nRngSaldos + 2) = Round(nImpDebe(0) - nImpHaber(0), 2)
             nImpSaldo(nRngSaldos + 5) = Round(nImpDebe(1) - nImpHaber(1), 2)
           End If
         Next nIndex
         
         ' Determino el importe de la cuenta (naturaleza)
         For nIndex = 0 To nLimite
           nImpSaldo(nIndex) = IIf(sSigno = "-" And nImpSaldo(nIndex) > 0, 0, IIf(sSigno = "+" And nImpSaldo(nIndex) < 0, 0, nImpSaldo(nIndex)))
         Next nIndex
         aOperador(1) = Mid(s_Cadena, nFinal + 1, 1)
         sCaso = "Cuenta"
         nContador = nFinal
       Case "¡"         ' Centro de costo
         nInicio = (InStr(nInicio, s_Cadena, "¡", vbTextCompare)) + 1
         nFinal = InStr(nInicio, s_Cadena, "!", vbTextCompare)
         nLen = (nFinal - nInicio)
         sVariable = Mid$(s_Cadena, nInicio, nLen)
         sVariable = Replace(sVariable, "$", "")
         sSigno = Left(sVariable, 1)
         sVariable = IIf(IsNumeric(sSigno), sVariable, Mid(sVariable, 2))
         ' Inicializa los saldos
         For nIndex = 0 To nLimite: nImpSaldo(nIndex) = 0: Next nIndex
         ' Selelccion segun años
         For nIndex = 0 To nAnos
           nRngSaldos = (nIndex * 6)
           nImpDebe(0) = 0: nImpHaber(0) = 0: nImpDebe(1) = 0: nImpHaber(1) = 0
           ' Busco la cuenta
           If porstCOCCoAcu.RecordCount > 0 Then porstCOCCoAcu.MoveFirst
           porstCOCCoAcu.Find "cLlave='" & Choose(nIndex + 1, gsAnoAct, psAnoOld) & sVariable & "'"
           If Not porstCOCCoAcu.EOF Then
             aTipo(0) = IIf(porstCOCCoAcu!NatCta = NATCTA_DEU, TPOCTB_DEB, TPOCTB_HAB)
             aTipo(1) = IIf(porstCOCCoAcu!NatCta = NATCTA_ACR, TPOCTB_DEB, TPOCTB_HAB)
             ' Acumulo los importes hasta el mes actual
             For nLen = 0 To Val(gsMesAct)
               nImpDebe(0) = nImpDebe(0) + porstCOCCoAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT)
               nImpHaber(0) = nImpHaber(0) + porstCOCCoAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT)
               nImpDebe(1) = nImpDebe(1) + porstCOCCoAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT)
               nImpHaber(1) = nImpHaber(1) + porstCOCCoAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT)
               ' Saldos por meses y acumulado
               If nIndex = 0 Then
                 nImpSaldo(12 + (4 * nLen)) = Round(porstCOCCoAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT) - porstCOCCoAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_NAC_TXT), 2)
                 nImpSaldo(12 + ((4 * nLen) + 1)) = Round(nImpDebe(0) - nImpHaber(0), 2)
                 nImpSaldo(12 + ((4 * nLen) + 2)) = Round(porstCOCCoAcu("Acu" & aTipo(0) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT) - porstCOCCoAcu("Acu" & aTipo(1) & Format(nLen, "00") & "_" & TPOMON_EXT_TXT), 2)
                 nImpSaldo(12 + ((4 * nLen) + 3)) = Round(nImpDebe(1) - nImpHaber(1), 2)
               End If
             Next nLen
             ' Saldo inicial del año
             nImpSaldo(nRngSaldos) = Round(porstCOCCoAcu("Acu" & aTipo(0) & "00_" & TPOMON_NAC_TXT) - porstCOCCoAcu("Acu" & aTipo(1) & "00_" & TPOMON_NAC_TXT), 2)
             nImpSaldo(nRngSaldos + 3) = Round(porstCOCCoAcu("Acu" & aTipo(0) & "00_" & TPOMON_EXT_TXT) - porstCOCCoAcu("Acu" & aTipo(1) & "00_" & TPOMON_EXT_TXT), 2)
             ' Saldo del mes
             nImpSaldo(nRngSaldos + 1) = Round(porstCOCCoAcu("Acu" & aTipo(0) & gsMesAct & "_" & TPOMON_NAC_TXT) - porstCOCCoAcu("Acu" & aTipo(1) & gsMesAct & "_" & TPOMON_NAC_TXT), 2)
             nImpSaldo(nRngSaldos + 4) = Round(porstCOCCoAcu("Acu" & aTipo(0) & gsMesAct & "_" & TPOMON_EXT_TXT) - porstCOCCoAcu("Acu" & aTipo(1) & gsMesAct & "_" & TPOMON_EXT_TXT), 2)
             ' Saldo acumulado
             nImpSaldo(nRngSaldos + 2) = Round(nImpDebe(0) - nImpHaber(0), 2)
             nImpSaldo(nRngSaldos + 5) = Round(nImpDebe(1) - nImpHaber(1), 2)
           End If
         Next nIndex
         
         ' Determino el importe de la cuenta (naturaleza)
         For nIndex = 0 To nLimite
           nImpSaldo(nIndex) = IIf(sSigno = "-" And nImpSaldo(nIndex) > 0, 0, IIf(sSigno = "+" And nImpSaldo(nIndex) < 0, 0, nImpSaldo(nIndex)))
         Next nIndex
         aOperador(1) = Mid(s_Cadena, nFinal + 1, 1)
         sCaso = "Costo"
         nContador = nFinal
       Case "{"         ' Linea
         nInicio = (InStr(nInicio, s_Cadena, "{", vbTextCompare)) + 1
         nFinal = InStr(nInicio, s_Cadena, "}", vbTextCompare)
         nLen = (nFinal - nInicio)
         sVariable = Mid$(s_Cadena, nInicio, nLen)
         For nIndex = 0 To nLimite: nImpSaldo(nIndex) = 0: Next nIndex
         If porstMRpRs.RecordCount > 0 Then porstMRpRs.MoveFirst
         porstMRpRs.Find "NroLin='" & sVariable & "'"
         If Not porstMRpRs.EOF Then
            porstMRpRs.Move IIf(porstMRpRs!SubLin = "a", 2, IIf(porstMRpRs!SubLin = "c", 1, 0))
            nImpSaldo(aIndx(0)) = porstMRpRs!ImpSaldoIni
            nImpSaldo(aIndx(1)) = porstMRpRs!ImpSaldoMes
            nImpSaldo(aIndx(2)) = porstMRpRs!ImpSaldoAcu
            nImpSaldo(aIndx(3)) = porstMRpRs!ImpSalIniMN
            nImpSaldo(aIndx(4)) = porstMRpRs!ImpSalMesMN
            nImpSaldo(aIndx(5)) = porstMRpRs!ImpSalAcuMN
            nImpSaldo(aIndx(6)) = porstMRpRs!ImpSalIniME
            nImpSaldo(aIndx(7)) = porstMRpRs!ImpSalMesME
            nImpSaldo(aIndx(8)) = porstMRpRs!ImpSalAcuME
            For nIndex = 0 To Val(gsMesAct)
              nImpSaldo(aIndx(9 + nIndex)) = porstMRpRs.Fields("ImpSaldo_" & Format(nIndex, "00"))
            Next
         End If
         aOperador(1) = Mid(s_Cadena, nFinal + 1, 1)
         sCaso = "Linea"
         nContador = nFinal
       Case "("         ' Valores
         nInicio = (InStr(nInicio, s_Cadena, "(", vbTextCompare)) + 1
         nFinal = InStr(nInicio, s_Cadena, ")", vbTextCompare)
         nLen = (nFinal - nInicio)
         sVariable = Mid$(s_Cadena, nInicio, nLen)
         For nIndex = 0 To nLimite: nImpSaldo(nIndex) = CDec(Val(sVariable)): Next nIndex
         aOperador(1) = Mid(s_Cadena, nFinal + 1, 1)
         sCaso = "Valores"
         nContador = nFinal
       Case Else        ' Otro Caso
         Select Case aOperador(0)
           Case "+" ' Operación de suma
             For nIndex = 0 To nLimite
               aImpFormula(nIndex) = Round(aImpFormula(nIndex) + nImpSaldo(nIndex), 2)
             Next nIndex
             sCaso = Mid(s_Cadena, nFinal + 2, 1)
             nContador = nFinal + 1
             aOperador(0) = aOperador(1)
           Case "-" ' Operación de resta
             For nIndex = 0 To nLimite
               aImpFormula(nIndex) = Round(aImpFormula(nIndex) - nImpSaldo(nIndex), 2)
             Next nIndex
             sCaso = Mid(s_Cadena, nFinal + 2, 1)
             nContador = nFinal + 1
             aOperador(0) = aOperador(1)
           Case "*" ' Operación de multiplicación
             For nIndex = 0 To nLimite
               aImpFormula(nIndex) = Round(aImpFormula(nIndex) * nImpSaldo(nIndex), 2)
             Next nIndex
             sCaso = Mid(s_Cadena, nFinal + 2, 1)
             nContador = nFinal + 1
             aOperador(0) = aOperador(1)
         End Select
     End Select
   Loop

End Function

Private Sub mfgMain_SelChange()
  Dim nSeleccion As Long, nColumna As Long
  nSeleccion = mfgMain.RowSel
  nColumna = mfgMain.ColSel
  mfgMain.Redraw = False
  mfgMain.row = nFilSele
  mfgMain.Col = 2: mfgMain.CellForeColor = vbBlack
  mfgMain.Col = 3: mfgMain.CellForeColor = vbBlack
  nFilSele = nSeleccion
  mfgMain.row = nFilSele
  mfgMain.Col = 2: mfgMain.CellForeColor = vbBlue
  mfgMain.Col = 3: mfgMain.CellForeColor = vbBlue
  mfgMain.Col = nColumna
  mfgMain.Redraw = True
End Sub
' ma
']
Private Sub optProceso_Click(Index As Integer)

cboTpoMon.Enabled = (Index = 0 Or Index = 4 Or Index = 5 Or Index = 6)
frmSaldos.Enabled = (Index <> 1 And Index <> 2 And Index <> 5 And Index <> 6)

End Sub

