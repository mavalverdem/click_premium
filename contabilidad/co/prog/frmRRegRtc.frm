VERSION 5.00
Begin VB.Form frmRRegRtc 
   Caption         =   "[título]"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   12
         Top             =   315
         Value           =   -1  'True
         Width           =   915
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
      Top             =   2205
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
         Picture         =   "frmRRegRtc.frx":0000
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
         Picture         =   "frmRRegRtc.frx":0102
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
         Picture         =   "frmRRegRtc.frx":0634
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
         Picture         =   "frmRRegRtc.frx":077E
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
   Begin VB.Label Label2 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   2
      Left            =   5400
      TabIndex        =   10
      Top             =   945
      Width           =   600
   End
End
Attribute VB_Name = "frmRRegRtc"
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
Private porstMRpRs As ADODB.Recordset

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
   Set porstMRpRs = New ADODB.Recordset
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
    With porstMRpRs
        .ActiveConnection = pocnnMain
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Source = "SELECT * FROM tmpRptRegReten"
'        .Open
    End With
   With porstTGAux
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodAux, RazAux " _
              & "FROM TGAux"
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
   
    With cboTpoMon
        .AddItem TPOMON_NAC_TXT_1, 0
        .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND

 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
'   With porstTGAux
'      .MoveLast
'      txtDato(1).Text = !CodAux
'      .MoveFirst
'      txtDato(0).Text = !CodAux
'   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
  
  'Otros.
   
  'Características de impresión.
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
   
   '[ Creacion de Temporales
  With porstMRp
    If .State = adStateOpen Then .Close
        .Source = "SELECT DISTINCTROW a.CodAux, b.RazAux, b.RucAux, a.FehOpe, a.NroIte, a.FehOpe as cFehOpeR, " _
                & "DetTDc AS cDocumento, DetTDc as cDenCpb, DetTDc as cTpoTra, "
        .Source = .Source & "a.ImpMN AS cDebe, a.ImpMN AS cHaber, a.CodTDc, a.NroIte as cNroGrp, a.ImpMN AS cSaldo "
        .Source = .Source & "FROM ((CoCpbDet a " _
                & "  LEFT JOIN TgAux b ON a.CodAux=b.CodAux) " _
                & "  LEFT JOIN TgTDc c ON a.CodTDc=c.CodTDc) " _
                & "WHERE (a.TpoPvs='" & TPOPVS_PVS & "' or a.TpoPvs='" & TPOPVS_CAN & "')" _
                & " AND a.TpoGnr<>'" & TPOGNR_DST & "' "
    .Open
  End With
  ppDatosRetenciones porstMRp.Source
     
    usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
    If usDEstino = PRN_DEST_GRAF Then
        Call gpEncabezadoRpt(frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True)
'
'        'Prepara_Crystal
'
        With frmMain.rptMain
        '[Datos y parámetros del reporte.  'Cambiar.
            '.ReportFileName = gsRutRpt & "rptRRegHPr.rpt"
            .ReportFileName = gsRutRpt & "rptRRegRet.rpt"
            .WindowShowExportBtn = IIf(paOpciones(2), True, False)
            '[ Formula para Simbolo de Moneda ]
'            .Formulas(6) = "pSigMon='" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, gsTpoMon_Sgn_ME, gsTpoMon_Sgn_MN) & "'"
            .MarginLeft = unMargenIzquierdo
            .WindowState = crptMaximized
            .Connect = "Provider=MySqlProv;Extended Properties=" & CONNSTRG & gsNomBDS
            .Action = 1
        End With
'        If Index = 0 Then frmMain.rptMain.Destination = crptToWindow
    Else
        Set MRViewer = New MRViewerObject
        With MRViewer
            .DataRecordSet = porstMRpRs
            '.LoadReport gsRutRpt & "rptRRegHPr.mrp"
            .LoadReport gsRutRpt & "rptRRegRet.mrp"
            Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True)
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
   pocnnMain.Execute "DROP TABLE IF EXISTS tmpRptRegReten"
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

'[Propio del formulario.

Public Sub ppDatosRetenciones(s_Source As String)
   Dim dorstRetencion As ADODB.Recordset
   Dim dorstCOCpbDetDR As ADODB.Recordset
   Dim dorstCOCpbDetI As ADODB.Recordset
   Dim dsNroIte As Integer, dsNroGrp As Integer
   
   Set dorstRetencion = New ADODB.Recordset
   Set dorstCOCpbDetDR = New ADODB.Recordset
   Set dorstCOCpbDetI = New ADODB.Recordset
   pocnnMain.Execute "DROP TABLE IF EXISTS tmpRptRegReten"
   pocnnMain.Execute "CREATE TABLE IF NOT EXISTS tmpRptRegReten " & s_Source
   pocnnMain.Execute "DELETE FROM tmpRptRegReten"
   porstMRpRs.Open
   pocnnMain.BeginTrans
   With dorstCOCpbDetDR
      .ActiveConnection = pocnnMain
      .Source = "SELECT DISTINCTROW CONCAT(a.CodAux, a.CodCta, a.CodTDc, a.SerDoc, a.NroDoc) as cLlaveD " _
              & "FROM (CoCpbDet a " _
              & "  LEFT JOIN COCpbDetRP b ON a.CodAux=b.CodAux AND a.CodCta=b.CodCta AND" _
              & "                            a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND" _
              & "                            a.NroDoc=b.NroDoc AND a.NroIte=b.NroIte) " _
              & "WHERE (a.TpoGnr<>'" & TPOGNR_DST & "') AND b.CodTDc_RP='" & gsCodTDc_Rtc & "' " _
              & "  AND (a.TpoPvs='" & TPOPVS_CAN & "') AND (b.SerDoc<>'') " _
              & IIf(txtDato(0).Text = "", " ", " AND a.CodAux='" & txtDato(0).Text & "' ") _
              & "ORDER BY a.CodAux, a.FehOpe, a.CodTDc, a.SerDoc, a.NroDoc "
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With dorstCOCpbDetI
      .ActiveConnection = pocnnMain
      .Source = "SELECT DISTINCTROW a.CodAux, c.RazAux, c.RUCAux, a.FehOpe, a.NroIte, " _
              & "  IF(a.CodTDc='01','Factura','Nta. Crédito')  as cDenCpb, " _
              & "  IF(a.TpoPvs='C', 'Pago', IF(a.TpoPvs='P', 'Compra', 'Ajuste')) as cTpoTra, " _
              & "  a.TpoCtb, a.ImpMN, a.ImpME, a.ImpTCb, a.CodTDc, a.TpoPvs, " _
              & "  CONCAT(a.CodAux, a.CodCta, a.CodTDc, a.SerDoc, a.NroDoc) as cLlaveD, " _
              & "  CONCAT(a.SerDoc, '-', a.NroDoc) AS cDocumento, " _
              & "  CONCAT(b.CodTDc_RP, b.SerDoc_RP, b.NroDoc_RP) AS cDocR " _
              & "FROM ((CoCpbDet a " _
              & "  LEFT JOIN COCpbDetRP b ON a.CodAux=b.CodAux AND a.CodCta=b.CodCta AND" _
              & "                            a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND" _
              & "                            a.NroDoc=b.NroDoc AND a.NroIte=b.NroIte) " _
              & "  LEFT JOIN TGAux c ON a.CodAux=c.CodAux) " _
              & "WHERE (a.TpoGnr<>'" & TPOGNR_DST & "') " _
              & "  AND a.MesPvs<='" & gsMesAct & "' " _
              & "  AND (IFNULL(a.CodAux, '')<>'' AND IFNULL(a.CodTDc, '')<>'' " _
              & "  AND IFNULL(a.SerDoc, '')<>'' AND IFNULL(a.NroDoc, '')<>'') " _
              & "ORDER BY a.TpoPvs DESC "
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With dorstRetencion
      .ActiveConnection = pocnnMain
      .Source = "SELECT DISTINCTROW a.FehOpe, " _
              & "  CONCAT(a.SerDoc, '-', a.NroDoc) AS cDocumento, 'C/R' as cDenCpb, 'Retención' as cTpoTra, " _
              & "  a.TpoCtb, a.ImpMN, a.ImpME, a.ImpTCb, a.CodTDc, " _
              & "  CONCAT(b.CodTDc_RP, b.SerDoc_RP, b.NroDoc_RP) AS cDocR " _
              & "FROM (CoCpbDet a " _
              & "  LEFT JOIN COCpbDetRP b ON a.CodTDc=b.CodTDc_RP AND a.SerDoc=b.SerDoc_RP AND" _
              & "                            a.NroDoc=b.NroDoc_RP) " _
              & "WHERE a.CodTDc='" & gsCodTDc_Rtc & "' " _
              & "  AND (a.TpoPvs='" & TPOPVS_PVS & "') " _
              & "  AND a.TpoGnr<>'" & TPOGNR_DST & "' "
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   dsNroGrp = 0
   With dorstCOCpbDetDR
      If .RecordCount > 0 Then .MoveFirst
      Do While Not .EOF
         dsNroGrp = dsNroGrp + 1
         pnSaldo = 0
         dorstCOCpbDetI.Filter = adFilterNone
         If dorstCOCpbDetI.RecordCount > 0 Then dorstCOCpbDetI.MoveFirst
         dorstCOCpbDetI.Filter = "cLlaveD='" & dorstCOCpbDetDR!cLlaveD & "'"
         If Not dorstCOCpbDetI.EOF Then
            dsNroIte = 1
            dorstCOCpbDetI.MoveFirst
            Do While Not dorstCOCpbDetI.EOF
               If dorstCOCpbDetI!TpoPvs = "P" Then
                  PPAddrstMRP dorstCOCpbDetI!CodAux, dorstCOCpbDetI!RazAux, dorstCOCpbDetI!RucAux, dsNroIte, dorstCOCpbDetI!FehOpe, dorstCOCpbDetI!FehOpe, dorstCOCpbDetI!cDocumento, _
                              dorstCOCpbDetI!cDenCpb, dorstCOCpbDetI!cTpoTra, dorstCOCpbDetI!TpoCtb, dorstCOCpbDetI!ImpMN, _
                              dorstCOCpbDetI!ImpME, dorstCOCpbDetI!CodTDc, dsNroGrp, dorstCOCpbDetI!ImpTCb
               Else
                  If dorstRetencion.RecordCount > 0 Then dorstRetencion.MoveFirst
                  dorstRetencion.Find "cDocR='" & dorstCOCpbDetI!cDocR & "'"
                  If Not dorstRetencion.EOF Then
                     PPAddrstMRP dorstCOCpbDetI!CodAux, dorstCOCpbDetI!RazAux, dorstCOCpbDetI!RucAux, dsNroIte, dorstCOCpbDetI!FehOpe, dorstCOCpbDetI!FehOpe, dorstCOCpbDetI!cDocumento, _
                                 dorstCOCpbDetI!cDenCpb, dorstCOCpbDetI!cTpoTra, dorstCOCpbDetI!TpoCtb, CDec(dorstCOCpbDetI!ImpMN) - CDec(dorstRetencion!ImpMN), _
                                 CDec(dorstCOCpbDetI!ImpME) - CDec(dorstRetencion!ImpME), dorstCOCpbDetI!CodTDc, dsNroGrp, dorstCOCpbDetI!ImpTCb
                     dsNroIte = dsNroIte + 1
                     PPAddrstMRP dorstCOCpbDetI!CodAux, dorstCOCpbDetI!RazAux, dorstCOCpbDetI!RucAux, dsNroIte, dorstRetencion!FehOpe, dorstRetencion!FehOpe, dorstRetencion!cDocumento, _
                                 dorstRetencion!cDenCpb, dorstRetencion!cTpoTra, dorstCOCpbDetI!TpoCtb, dorstRetencion!ImpMN, _
                                 dorstRetencion!ImpME, dorstRetencion!CodTDc, dsNroGrp, dorstRetencion!ImpTCb
                  Else
                     PPAddrstMRP dorstCOCpbDetI!CodAux, dorstCOCpbDetI!RazAux, dorstCOCpbDetI!RucAux, dsNroIte, dorstCOCpbDetI!FehOpe, dorstCOCpbDetI!FehOpe, dorstCOCpbDetI!cDocumento, _
                                 dorstCOCpbDetI!cDenCpb, dorstCOCpbDetI!cTpoTra, dorstCOCpbDetI!TpoCtb, CDec(dorstCOCpbDetI!ImpMN), _
                                 CDec(dorstCOCpbDetI!ImpME), dorstCOCpbDetI!CodTDc, dsNroGrp, dorstCOCpbDetI!ImpTCb
                  End If
               End If
               dsNroIte = dsNroIte + 1
               dorstCOCpbDetI.MoveNext
            Loop
         End If
         .MoveNext
      Loop
      porstMRpRs.UpdateBatch
      pocnnMain.CommitTrans
   End With
End Sub

Public Sub PPAddrstMRP(cCodAux As String, cRazAux As String, cRucAux As String, cNroIte As Integer, cFOpeR As String, _
                        cFOpe As String, cDoc As String, cDCpb As String, cTTra As String, cTpoCtb As String, _
                        cImpMN As Double, cImpME As Double, cCodTDc As String, cNroGrp As Integer, cImpTCb As Double)

   porstMRpRs.AddNew
   porstMRpRs!CodAux = Trim(cCodAux)
   porstMRpRs!RazAux = Trim(cRazAux)
   porstMRpRs!RucAux = Trim(cRucAux)
   porstMRpRs!NroIte = Trim(cNroIte)
   porstMRpRs!cFehOpeR = Format(cFOpeR, "yyyy-mm-dd")
   porstMRpRs!FehOpe = Format(cFOpe, "yyyy-mm-dd")
   porstMRpRs!cDocumento = Trim(cDoc)
   porstMRpRs!cDenCpb = Trim(cDCpb)
   porstMRpRs!cTpoTra = Trim(cTTra)
   If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
      porstMRpRs!cDebe = IIf(Trim(cTpoCtb) = TPOCTB_DEB, CDec(cImpMN), 0)
      porstMRpRs!cHaber = IIf(Trim(cTpoCtb) = TPOCTB_HAB, CDec(cImpMN), 0)
   Else
      porstMRpRs!cDebe = IIf(Trim(cTpoCtb) = TPOCTB_DEB, CDec(cImpME) * CDec(cImpTCb), 0)
      porstMRpRs!cHaber = IIf(Trim(cTpoCtb) = TPOCTB_HAB, CDec(cImpME) * CDec(cImpTCb), 0)
   End If
   porstMRpRs!CodTDc = cCodTDc
   porstMRpRs!cNroGrp = cNroGrp
   dnSaldo = CDec(dnSaldo) + CDec(porstMRpRs!cDebe) - CDec(porstMRpRs!cHaber)
   pnSaldo = CDec(Abs(dnSaldo))
   porstMRpRs!cSaldo = pnSaldo
End Sub
                                                                               
']Propio del formulario.

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

