VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRMovBco 
   Caption         =   "[título]"
   ClientHeight    =   4425
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDocumento 
      Caption         =   " Documento "
      ForeColor       =   &H00C00000&
      Height          =   960
      Left            =   30
      TabIndex        =   13
      Top             =   2055
      Width           =   1905
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   570
         Width           =   1670
      End
      Begin VB.ComboBox cboTpoDoc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   225
         Width           =   1670
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   " Auxiliar "
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   1335
      Width           =   6975
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   2
         Left            =   6570
         Picture         =   "frmRMovBco.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   255
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
         Height          =   285
         Index           =   2
         Left            =   1365
         TabIndex        =   12
         Top             =   255
         Width           =   5205
      End
   End
   Begin VB.Frame fraMovimiento 
      Caption         =   " Movimiento "
      ForeColor       =   &H00C00000&
      Height          =   645
      Left            =   2040
      TabIndex        =   16
      Top             =   2055
      Width           =   2775
      Begin VB.CheckBox chkMovimiento 
         Caption         =   "Egreso"
         ForeColor       =   &H00800000&
         Height          =   190
         Index           =   1
         Left            =   1530
         TabIndex        =   18
         Top             =   255
         Width           =   1000
      End
      Begin VB.CheckBox chkMovimiento 
         Caption         =   "Ingreso"
         ForeColor       =   &H00800000&
         Height          =   190
         Index           =   0
         Left            =   195
         TabIndex        =   17
         Top             =   255
         Width           =   1000
      End
   End
   Begin VB.CheckBox chkFecha 
      Caption         =   " Rango Fecha "
      ForeColor       =   &H00800000&
      Height          =   190
      Left            =   2310
      TabIndex        =   22
      Top             =   3075
      Width           =   1350
   End
   Begin VB.Frame fraRngFecha 
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   30
      TabIndex        =   23
      Top             =   3075
      Width           =   3810
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   300
         Left            =   540
         TabIndex        =   25
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   2400
         TabIndex        =   27
         Top             =   255
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   37953
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   24
         Top             =   315
         Width           =   255
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "al"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   3
         Left            =   2025
         TabIndex        =   26
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5625
      TabIndex        =   19
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   28
      Top             =   3075
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   29
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   30
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   45
      Width           =   6975
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   645
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Width           =   645
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   0
         Left            =   6570
         Picture         =   "frmRMovBco.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   1
         Left            =   6570
         Picture         =   "frmRMovBco.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   780
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
         Height          =   285
         Index           =   1
         Left            =   750
         TabIndex        =   9
         Top             =   780
         Width           =   5820
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
         Height          =   285
         Index           =   0
         Left            =   750
         TabIndex        =   7
         Top             =   450
         Width           =   5820
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Bancos"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5730
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2520
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
      ScaleWidth      =   6975
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3885
      Width           =   6975
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
         Picture         =   "frmRMovBco.frx":04FE
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
         Picture         =   "frmRMovBco.frx":0648
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
         Picture         =   "frmRMovBco.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   4995
      TabIndex        =   20
      Top             =   2565
      Width           =   660
   End
End
Attribute VB_Name = "frmRMovBco"
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
Private porstCoBco As ADODB.Recordset
Private porstTGAux As ADODB.Recordset
']

Private Sub chkFecha_Click()
  fraRngFecha.Enabled = (chkFecha.Value = vbChecked)
End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCoBco = New ADODB.Recordset
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
   With porstCoBco
      .ActiveConnection = pocnnMain
      .Source = "SELECT codbco, " & Choose(gsIdioma, "detbco", "detbcox") & " AS detbco "
      .Source = .Source & "FROM cobco "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "ORDER BY codbco"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With porstTGAux
    .ActiveConnection = pocnnMain
    .Source = "SELECT codaux, RazAux "
    .Source = .Source & "FROM TGAux "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "ORDER BY codaux"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  With cboTpoDoc
    .AddItem "", 0
    .AddItem TPODOC_DPS_TXT, TPODOC_DPS_IND
    .AddItem TPODOC_GRO_TXT, TPODOC_GRO_IND
    .AddItem TPODOC_TRA_TXT, TPODOC_TRA_IND
    .AddItem TPODOC_ORD_TXT, TPODOC_ORD_IND
    .AddItem TPODOC_DEB_TXT, TPODOC_DEB_IND
    .AddItem TPODOC_CRE_TXT, TPODOC_CRE_IND
    .AddItem TPODOC_CHQ_TXT, TPODOC_CHQ_IND
    .AddItem TPODOC_OTR_TXT, TPODOC_OTR_IND
    .AddItem TPODOC_EFE_TXT, TPODOC_EFE_IND
    .AddItem TPODOC_PEX_TXT, TPODOC_PEX_IND
    .AddItem TPODOC_LTR_TXT, TPODOC_LTR_IND
  End With
   
  With txtDato
    For dnContador = 0 To 1
      .Item(dnContador).DataField = "codbco"
      .Item(dnContador).MaxLength = porstCoBco.Fields(.Item(dnContador).DataField).DefinedSize
    Next
  End With
  txtDato.Item(2).DataField = "CodAux"
  txtDato.Item(2).MaxLength = porstTGAux.Fields(txtDato.Item(2).DataField).DefinedSize
  txtDato(3).MaxLength = 10
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Bancos :", "Moneda :", "Del :", "Al :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Banks :", "Currency :", "From :", "To :")
  Next nElemento
  
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraAuxiliar.Caption = Choose(gsIdioma, " Auxiliar ", " Auxiliary ")
  fraDocumento.Caption = Choose(gsIdioma, " Documento ", " Document ")
  fraMovimiento.Caption = Choose(gsIdioma, "Movimiento", "Movement")
  chkMovimiento(0).Caption = TPOBAN_ING_TXT
  chkMovimiento(1).Caption = TPOBAN_EGR_TXT
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  chkFecha.Caption = Choose(gsIdioma, "Rango Fecha", "Range Date")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  chkMovimiento(0).Value = vbUnchecked
  chkMovimiento(1).Value = vbChecked
  fraRngFecha.Enabled = False
  dtpDesde.Value = CDate("01/" & gsMesAct & "/" & gsAnoAct)
  dtpHasta.Value = gfUltDia(dtpDesde.Value)
  'Límites de rangos.
   With porstCoBco
      .MoveLast
      txtDato(1).Text = !codbco
      .MoveFirst
      txtDato(0).Text = !codbco
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   cboTpoDoc.ListIndex = 0
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
   porstCoBco.Close
   porstTGAux.Close
   pocnnMain.Close
   Set porstCoBco = Nothing
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim sDocBanco As String
  ppHabilitacion False
  
  sDocBanco = "(CASE cab.tpodoc WHEN " & TPODOC_DPS_IND & " THEN 'DPS-' WHEN " & TPODOC_GRO_IND & " THEN 'GRO-' "
  sDocBanco = sDocBanco & "WHEN " & TPODOC_TRA_IND & " THEN 'TRF-' WHEN " & TPODOC_ORD_IND & " THEN 'ODP-' WHEN " & TPODOC_DEB_IND & " THEN 'TDE-' "
  sDocBanco = sDocBanco & "WHEN " & TPODOC_CRE_IND & " THEN 'TCR-' WHEN " & TPODOC_CHQ_IND & " THEN 'CHQ-' WHEN " & TPODOC_OTR_IND & " THEN 'OTR-' "
  sDocBanco = sDocBanco & "WHEN " & TPODOC_EFE_IND & " THEN 'EFE-' WHEN " & TPODOC_PEX_IND & " THEN 'PEX-' WHEN " & TPODOC_LTR_IND & " THEN 'LTR-' ELSE '' END)"
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT cab.fehban, cab.codbco, cab.codcta, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "Concat(" & sDocBanco & ", IFNULL(cab.docban, ''))", "(" & sDocBanco & "+ISNULL(cab.docban, ''))") & " AS cdocubanco, "
    .Source = .Source & Choose(gsIdioma, "cab.globan", "cab.globanx") & " AS globan, "
    .Source = .Source & "(CASE cab.tpomon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cabmoneda, det.codaux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cab.coddro, '-', cab.nroban)", "(cab.coddro+'-'+cab.nroban)") & " AS cnroban, "
    .Source = .Source & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "Concat(tdc.abvtdc,'-',det.serdoc,'-',det.nrodoc)", "(tdc.abvtdc+'-'+det.serdoc+'-'+det.nrodoc)") & " AS cdocumento, "
    .Source = .Source & "(CASE det.tpomon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS detmoneda, "
    .Source = .Source & Choose(gsIdioma, "bco.detbco", "bco.detbcox") & " AS detbco, "
    .Source = .Source & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, aux.razaux, "
    .Source = .Source & "(CASE det.tpoban WHEN '" & TPOBAN_ING & "' THEN det.impmn ELSE 0 END) AS cargomn, "
    .Source = .Source & "(CASE det.tpoban WHEN '" & TPOBAN_EGR & "' THEN det.impmn ELSE 0 END) AS abonomn, "
    .Source = .Source & "(CASE det.tpoban WHEN '" & TPOBAN_ING & "' THEN det.impme ELSE 0 END) AS cargome, "
    .Source = .Source & "(CASE det.tpoban WHEN '" & TPOBAN_EGR & "' THEN det.impme ELSE 0 END) AS abonome "
    .Source = .Source & "FROM cobancab cab "
    .Source = .Source & "INNER JOIN cobandet det ON cab.codemp=det.codemp AND cab.pdoano=det.pdoano AND cab.mespvs=det.mespvs AND cab.coddro=det.coddro AND cab.nroban=det.nroban "
    .Source = .Source & "INNER JOIN cocta cta ON cab.codemp=cta.codemp AND cab.pdoano=cta.pdoano AND cab.codcta=cta.codcta "
    .Source = .Source & "LEFT JOIN cobco bco ON cab.codemp=bco.codemp AND cab.codbco=bco.codbco "
    .Source = .Source & "LEFT JOIN tgaux aux ON det.codemp=aux.codemp AND det.codaux=aux.codaux "
    .Source = .Source & "LEFT JOIN tgtdc tdc ON det.codemp=tdc.codemp AND det.codtdc=tdc.codtdc "
    .Source = .Source & "WHERE cab.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND cab.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND cab.mespvs<>'00' "
    If chkFecha.Value = vbChecked Then
      If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "AND DATE_FORMAT(cab.fehban, '%Y-%m-%d') >='" & Format(dtpDesde, "yyyy-mm-dd") & "' "
        .Source = .Source & "AND DATE_FORMAT(cab.fehban, '%Y-%m-%d') <='" & Format(dtpHasta, "yyyy-mm-dd") & "' "
      Else
        .Source = .Source & "AND CONVERT(datetime, cab.fehban, 120) >='" & Format(dtpDesde, "yyyy-mm-dd") & "' "
        .Source = .Source & "AND CONVERT(datetime, cab.fehban, 120) <='" & Format(dtpHasta, "yyyy-mm-dd") & "' "
      End If
    Else
      .Source = .Source & "AND cab.mespvs='" & gsMesAct & "' "
    End If
    If Not (chkMovimiento(0).Value = vbChecked And chkMovimiento(1).Value = vbChecked) Then
      .Source = .Source & "AND cab.tpoban='" & IIf(chkMovimiento(0).Value = vbChecked, TPOBAN_ING, TPOBAN_EGR) & "' "
    End If
    .Source = .Source & "AND cab.tpomon='" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT) & "' "
    If cboTpoDoc.Text <> "" Then
      .Source = .Source & "AND cab.tpodoc='" & cboTpoDoc.ListIndex & "' "
    End If
    If txtDato(2).Text <> "" Then
      .Source = .Source & "AND det.codaux='" & txtDato(2).Text & "' "
    End If
    If txtDato(3).Text <> "" Then
      .Source = .Source & "AND cab.docban='" & txtDato(3).Text & "' "
    End If
    .Source = .Source & "AND cab.codbco BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "ORDER BY cab.fehban, cab.codbco, cdocubanco"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptrmovbco.rpt"
      '         .WindowShowGroupTree = True
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRMovCja.mrp"
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      '.Parameters("pTipoFecha") = IIf(optFecha(0).Value, "Emisión", "Cancelac.")
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
   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2                          'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
    modAyuBus.Bco_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 2                           'Cambiar (añadir índices).
    modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
   Case 0, 1
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstCoBco
      .MoveFirst
      .Find "codbco='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detbco), "", !detbco)
      End If
    End With
   Case 2
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstTGAux
      .MoveFirst
      .Find "codaux='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!RazAux), "", !RazAux)
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

