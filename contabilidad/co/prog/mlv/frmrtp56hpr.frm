VERSION 5.00
Begin VB.Form frmRTp56HPr 
   Caption         =   "[título]"
   ClientHeight    =   3405
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFolio 
      Caption         =   "Folio Inicial"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2385
      TabIndex        =   15
      Top             =   1770
      Width           =   1800
   End
   Begin VB.CheckBox chkCabecera 
      Caption         =   "Imprime Cabecera"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   45
      TabIndex        =   14
      Top             =   1770
      Width           =   1800
   End
   Begin VB.CheckBox chkDiario 
      Caption         =   "Totaliza Diario"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3180
      TabIndex        =   7
      Top             =   855
      Width           =   1335
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Diario"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   1005
      Width           =   4530
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
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   780
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4080
         Picture         =   "frmrtp56hpr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   255
         Width           =   3240
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   21
      Top             =   2115
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   22
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   23
         Top             =   315
         Width           =   1035
      End
   End
   Begin VB.Frame fraOrden 
      Caption         =   "Orden"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   2115
      Width           =   4995
      Begin VB.OptionButton OptOrden 
         Caption         =   "Comprobante"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   20
         Top             =   315
         Width           =   1260
      End
      Begin VB.OptionButton OptOrden 
         Caption         =   "Documento"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   2250
         TabIndex        =   19
         Top             =   315
         Width           =   1140
      End
      Begin VB.OptionButton OptOrden 
         Caption         =   "Proveedor"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   18
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton OptOrden 
         Caption         =   "Fecha"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   4
      Top             =   135
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6885
         Picture         =   "frmrtp56hpr.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   255
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
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   1260
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
         TabIndex        =   6
         Top             =   255
         Width           =   5520
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   6060
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1035
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
      ScaleWidth      =   7320
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2865
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
         Picture         =   "frmrtp56hpr.frx":0354
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
         Picture         =   "frmrtp56hpr.frx":049E
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
         Picture         =   "frmrtp56hpr.frx":09D0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   5250
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frmRTp56HPr"
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
Private porstCodro As ADODB.Recordset
']

Private Sub chkDiario_Click()
  fraRangos.Enabled = (chkDiario.Value = vbUnchecked)
  txtDato(1).Text = IIf(chkDiario.Value = vbChecked, "", txtDato(1).Text)
  lblDatoDeta(1).Caption = IIf(chkDiario.Value = vbChecked, "", lblDatoDeta(1).Caption)
End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstTGAux = New ADODB.Recordset
   Set porstCodro = New ADODB.Recordset
   
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
   
   With porstCodro
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
    .Source = .Source & "FROM CODro "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "ORDER BY CodDro"
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
   End With
 
 ']

 '[Parámetros.                         'Cambiar.
  txtDato.Item(0).DataField = "CodAux"
  txtDato.Item(0).MaxLength = porstTGAux.Fields(txtDato.Item(0).DataField).DefinedSize
 
  txtDato.Item(1).DataField = "CodDro"
  txtDato.Item(1).MaxLength = porstCodro.Fields(txtDato.Item(1).DataField).DefinedSize
 ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :")
  Next nElemento
  fraAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  chkDiario.Caption = Choose(gsIdioma, "Totaliza Diario", "Journal Totalizes")
  fraRangos.Caption = Choose(gsIdioma, "Diario", "Journal")
  fraOrden.Caption = Choose(gsIdioma, "Orden", "Order")
  OptOrden(0).Caption = Choose(gsIdioma, "Fecha", "Date")
  OptOrden(1).Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  OptOrden(2).Caption = Choose(gsIdioma, "Documento", "Document")
  OptOrden(3).Caption = Choose(gsIdioma, "Comprobante", "Voucher")
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
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
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
   Case 0, 1
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim sMoneda As String
  Dim sRegistro As String
     
  ppHabilitacion False
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.FeEDoc, a.RefDoc, b.RucAux, b.RazAux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodDro,'-',a.NroCpb)", "(a.CodDro+'-'+a.NroCpb)") & " AS cDroCpb, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc,'-',a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS cNroDoc, "
    .Source = .Source & "a.ImpBru_" & sMoneda & " AS clmBru, a.ImpIR4_" & sMoneda & " AS clmIR4, "
    .Source = .Source & "a.ImpIES_" & sMoneda & " AS clmIES, a.ImpORT_" & sMoneda & " AS clmORT, "
    .Source = .Source & "a.ImpNet_" & sMoneda & " AS clmNet, "
    .Source = .Source & "a.ImpNet_" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_TXT, TPOMON_NAC_TXT) & " AS clmpNet_OM, "
    ' Genero la agrupacion del detalle
    If chkDiario.Value = vbChecked Then
    End If
    .Source = .Source & "a.CodDro, c.DetDro, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "a.CodDro", IIf(Trim(txtDato(1).Text) <> "", "a.CodDro", "'drxx'")) & " AS grupo, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "'1'", IIf(Trim(txtDato(1).Text) <> "", "'2'", "'0'")) & " AS resumen "
    .Source = .Source & "FROM ((COHPrDoc a "
    .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    .Source = .Source & "LEFT JOIN CODro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodDro=c.CodDro) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs='" & gsMesAct & "' "
    If Trim(txtDato(0).Text) <> "" Then
      .Source = .Source & "AND a.CodAux = '" & Trim(txtDato(0).Text) & "' "
    End If
    If Trim(txtDato(1).Text) <> "" Then
      .Source = .Source & "AND Left(a.CodDro, " & Len(Trim(txtDato(1).Text)) & ")='" & Trim(txtDato(1).Text) & "' "
    End If
    If OptOrden(0).Value = True Then
      .Source = .Source & "ORDER BY grupo, a.FeEDoc, a.SerDoc, a.NroDoc"
    Else
      If OptOrden(1).Value = True Then
        .Source = .Source & "ORDER BY grupo, a.CodAux, a.FeEDoc, a.SerDoc, a.NroDoc"
      Else
        If OptOrden(2).Value = True Then
          .Source = .Source & "ORDER BY grupo, a.SerDoc, a.NroDoc"
        Else
          .Source = .Source & "ORDER BY grupo, a.CodDro, a.NroCpb"
        End If
      End If
    End If
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptr56reghpr.rpt"
      '[ Formulas adicionales del reporte
      sRegistro = Format(gsRUCEmp, "&&.&&&.&&&&&&")
      .Formulas(6) = "mRucEmpresa='" & sRegistro & "'"
      .Formulas(7) = "mDireccion='" & gsDirEmp & "'"
      .Formulas(8) = "mDistrito='" & gsLocEmp & "'"
      .Formulas(9) = "mActividad='" & gsGirEmp & "'"
      .Formulas(10) = "mRepresentante='" & gsRepEmp & "'"
      sRegistro = Format(gsRepDNIEmp, "&&.&&&.&&&&&&")
      .Formulas(11) = "mDniRepresentante='" & sRegistro & "'"
      .Formulas(12) = "pSigMon='" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, gsTpoMon_Sgn_ME, gsTpoMon_Sgn_MN) & "'"
      sRegistro = IIf(chkCabecera.Value = vbChecked, "S", "N")
      .ParameterFields(1) = "Cabecera;" & sRegistro & ";true"
      sRegistro = IIf(chkFolio.Value = vbChecked, "S", "N")
      .ParameterFields(2) = "FolioInicial;" & sRegistro & ";true"
      
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
      .LoadReport gsRutRpt & "rptRRegHPr.mrp"
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pSigMon") = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, gsTpoMon_Sgn_ME, gsTpoMon_Sgn_MN)
      
      .Parameters("pPagePrinter") = ""
      If porstMRp.RecordCount > 0 Then
        porstMRp.MoveLast
        .Parameters("pPagePrinter") = porstMRp!RucAux & porstMRp!cNroDoc & porstMRp!cDroCpb
        porstMRp.MoveFirst
      End If
      ']
      
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

  Select Case Index    'Busca el dato en su tabla principal.
   Case 0                              'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
   Case 1
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
    If Len(txtDato(Index)) <> 4 Then txtDato(Index).SetFocus: Exit Sub
  End Select

End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
    Case 0                              'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    Case 1
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
    Case 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstCodro
         .MoveFirst
         .Find "Coddro='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
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

