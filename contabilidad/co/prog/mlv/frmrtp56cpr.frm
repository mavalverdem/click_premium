VERSION 5.00
Begin VB.Form frmRTp56Cpr 
   Caption         =   "[título]"
   ClientHeight    =   3015
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCabecera 
      Caption         =   "Imprime Cabecera"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   75
      TabIndex        =   14
      Top             =   1710
      Width           =   1800
   End
   Begin VB.CheckBox chkFolio 
      Caption         =   "Folio Inicial"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2415
      TabIndex        =   15
      Top             =   1710
      Width           =   1800
   End
   Begin VB.CheckBox chkDiario 
      Caption         =   "Totaliza Diario"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2895
      TabIndex        =   7
      Top             =   765
      Width           =   1545
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Diario"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   930
      Width           =   4530
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4080
         Picture         =   "frmrtp56cpr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   780
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
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   16
      Top             =   1680
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   315
         Value           =   -1  'True
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
         Width           =   1020
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   4
      Top             =   45
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6885
         Picture         =   "frmrtp56cpr.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   20
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
      ItemData        =   "frmrtp56cpr.frx":0354
      Left            =   6180
      List            =   "frmrtp56cpr.frx":0356
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   900
      Width           =   1125
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7335
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2475
      Width           =   7335
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
         Picture         =   "frmrtp56cpr.frx":0358
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
         Picture         =   "frmrtp56cpr.frx":04A2
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
         Picture         =   "frmrtp56cpr.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
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
      Left            =   5325
      TabIndex        =   11
      Top             =   945
      Width           =   765
   End
End
Attribute VB_Name = "frmRTp56Cpr"
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
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND IndPrv=" & INDAUX_PRV_ACT & " "
      .Source = .Source & "ORDER BY CodAux"
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
   With cboTpoMon
      .AddItem TPOMON_NAC_TXT_1, 0
      .AddItem TPOMON_EXT_TXT_1, 1
   End With
   
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
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  chkCabecera.Caption = Choose(gsIdioma, "Imprime Cabecera", "Print Head")
  chkFolio.Caption = Choose(gsIdioma, "Folio Inicial", "Initial Folio")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
'   With porstTgAux
'      .MoveLast
'      'txtDato(1).Text = !CodAux
'      .MoveFirst
'      txtDato(0).Text = !CodAux
'   End With
  
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   
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
  Dim dnContador As Byte
  Dim sRegistro As String
  
  ppHabilitacion False
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.FeEDoc, a.FehOpe, a.CodDro, a.NroCpb, b.AbvTDc, "
    .Source = .Source & "a.SerDoc, a.NroDoc, a.RefDoc, c.RUCAux, c.RazAux, a.NroCDt, a.FehCDt, "
    '[ARREGLAR. Poder configurar el signo en Tipo de Documento. ImpIGV_OGr_MN
    If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
      .Source = .Source & "(a.ImpOGr_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOgr, "
      .Source = .Source & "(a.ImpOGN_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOGN, "
      .Source = .Source & "(a.ImpONG_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpONG, "
      .Source = .Source & "(a.ImpExo_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpExo, "
      .Source = .Source & "(a.ImpIGV_OGr_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGr, "
      .Source = .Source & "(a.ImpIGV_OGN_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGN, "
      .Source = .Source & "(a.ImpIGV_ONG_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVONG, "
      .Source = .Source & "(a.ImpISC_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpISC, "
      .Source = .Source & "(a.ImpOIm_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOIm, "
      .Source = .Source & "(a.ImpTot_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot, "
      .Source = .Source & "(a.ImpTot_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot_OM, "
    Else
      .Source = .Source & "(a.ImpOGr_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOgr, "
      .Source = .Source & "(a.ImpOGN_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOGN, "
      .Source = .Source & "(a.ImpONG_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpONG, "
      .Source = .Source & "(a.ImpExo_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpExo, "
      .Source = .Source & "(a.ImpIGV_OGr_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGr, "
      .Source = .Source & "(a.ImpIGV_OGN_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGN, "
      .Source = .Source & "(a.ImpIGV_ONG_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVONG, "
      .Source = .Source & "(a.ImpISC_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpISC, "
      .Source = .Source & "(a.ImpOIm_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOIm, "
      .Source = .Source & "(a.ImpTot_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot, "
      .Source = .Source & "(a.ImpTot_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot_OM, "
    End If
    ']ARREGLAR.
    .Source = .Source & "b.CodTDc, d.DetDro, b.dettdc, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "a.CodDro", IIf(Trim(txtDato(1).Text) <> "", "a.CodDro", "'drxx'")) & " AS grupo, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "'1'", IIf(Trim(txtDato(1).Text) <> "", "'2'", "'0'")) & " AS resumen "
    .Source = .Source & "FROM (((COCprDoc a "
    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    .Source = .Source & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
    .Source = .Source & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs='" & gsMesAct & "' "
    If Trim(txtDato(0).Text) <> "" Then
      .Source = .Source & "AND a.CodAux='" & Trim(txtDato(0).Text) & "' "
    End If
    If Trim(txtDato(1).Text) <> "" Then
      .Source = .Source & "AND Left(a.CodDro, " & Len(Trim(txtDato(1).Text)) & ")='" & Trim(txtDato(1).Text) & "' "
    End If
    .Source = .Source & "ORDER BY grupo, a.CodDro, a.NroCpb ASC"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptr56regcpr.rpt"
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
      .LoadReport gsRutRpt & "rptRRegCpr.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      
      '[Parámetros adicionales.
      If porstMRp.RecordCount > 0 Then
        porstMRp.MoveLast
        .Parameters("pPagePrinter") = porstMRp!coddro & porstMRp!NroCpb
        porstMRp.MoveFirst
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
      modAyuBus.Aux_Det "IndPrv=" & INDAUX_PRV_ACT & " ", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
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

End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

