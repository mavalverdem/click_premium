VERSION 5.00
Begin VB.Form frmRCCtDetra 
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
      Width           =   3045
      Begin VB.OptionButton OptTipo 
         Caption         =   "Cancelar"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   315
         Width           =   1290
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Pendiente"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   1200
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
         Picture         =   "frmRCCtDetra.frx":0000
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
         Picture         =   "frmRCCtDetra.frx":01AA
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
         Picture         =   "frmRCCtDetra.frx":02F4
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
         Picture         =   "frmRCCtDetra.frx":0826
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRCCtDetra"
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
  Dim dnContador As Integer
   
  On Error GoTo Err
  
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
  OptTipo(0).Caption = Choose(gsIdioma, "Pendiente", "Pending")
  OptTipo(1).Caption = Choose(gsIdioma, "Cancelar", "Liquidate")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
  ']
  
  '[Datos predeterminados.              'Cambiar.
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
  Dim cCadReporte  As String, sServicio As String
  Dim nDetraccion As Double
  
  ppHabilitacion False
    
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 15)='#tmpRptCteDetra_') DROP TABLE #tmpRptCteDetra"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRptCteDetra", cCadReporte)

  cCadReporte = "CREATE TABLE " & IIf(ps_Plataforma = pSrvMySql, "tmpRptCteDetra (", ps_Prefijo & "tmpRptCteDetra (")
  cCadReporte = cCadReporte & "codaux varchar(11) Default Null, "
  cCadReporte = cCadReporte & "razaux varchar(80) Default Null, "
  cCadReporte = cCadReporte & "rucaux varchar(11) Default Null, "
  cCadReporte = cCadReporte & "sdocumento varchar(20) Default Null, "
  cCadReporte = cCadReporte & "scomprobante varchar(11) Default Null, "
  cCadReporte = cCadReporte & "fehope date Default Null, "
  cCadReporte = cCadReporte & "feedoc date Default Null, "
  cCadReporte = cCadReporte & "bienserv char(3) Default Null, "
  cCadReporte = cCadReporte & "tipopera char(2) Default Null, "
  cCadReporte = cCadReporte & "smoneda varchar(3) Default Null, "
  cCadReporte = cCadReporte & "importot_mn decimal(12,2) Not Null Default '0.00', "
  cCadReporte = cCadReporte & "impordetra_mn decimal(12,2) Not Null Default '0.00', "
  cCadReporte = cCadReporte & "importot_me decimal(12,2) Not Null Default '0.00', "
  cCadReporte = cCadReporte & "impordetra_me decimal(12,2) Not Null Default '0.00', "
  cCadReporte = cCadReporte & "porceta decimal(5,2) Not Null Default '0.00')"
  pocnnMain.Execute cCadReporte
  
  With porstMRp
    If .State = adStateOpen Then .Close
    cCadReporte = "SELECT cpr.codaux, aux.razaux, aux.rucaux, CONCAT(tdc.abvtdc, '-', cpr.serdoc, '-', cpr.nrodoc) AS sDocumento, cpr.fehope, cpr.feedoc, "
    cCadReporte = cCadReporte & "CONCAT(cpr.coddro, '-', cpr.nrocpb) AS scomprobante, "
    cCadReporte = cCadReporte & "LEFT(cpr.tsadetrac, 3) AS bienserv, RIGHT(cpr.tsadetrac, 2) AS tipopera, "
    cCadReporte = cCadReporte & "(CASE WHEN cpr.tpomon='" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS smoneda, "
    cCadReporte = cCadReporte & "cpr.imptot_mn, cpr.imptot_me "
    cCadReporte = cCadReporte & "FROM cocprdoc cpr "
    cCadReporte = cCadReporte & "INNER JOIN tgaux aux ON aux.codemp=cpr.codemp AND aux.codaux=cpr.codaux "
    cCadReporte = cCadReporte & "LEFT JOIN tgtdc tdc ON tdc.codemp=cpr.codemp AND tdc.codtdc=cpr.codtdc "
    cCadReporte = cCadReporte & "WHERE cpr.codemp='" & gsCodEmp & "' "
    cCadReporte = cCadReporte & "AND CONCAT(cpr.pdoano, cpr.mespvs)='" & gsAnoAct & gsMesAct & "' "
    If txtDato(0).Text <> "" Then
      cCadReporte = cCadReporte & "AND cpr.codaux='" & txtDato(0).Text & "' "
    End If
    cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cpr.tsadetrac, '0')<>'" & INDCDT_INA & "' "
    cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cpr.nrocdt, '')='' "
    If OptTipo(1).Value Then
      cCadReporte = cCadReporte & "AND cpr.indcdt='" & INDCDT_ACT & "' "
    End If
    cCadReporte = cCadReporte & "ORDER BY cpr.codaux, cpr.codtdc, cpr.serdoc, cpr.nrodoc"
    .Source = cCadReporte
    .Open
  End With
  ' Genero información
  If Not (porstMRp.BOF And porstMRp.EOF) Then
  Dim nContador As Integer ' 2014-04-06 reclasificacion de cod detraccion
  
    While Not porstMRp.EOF
      cCadReporte = "INSERT INTO " & IIf(ps_Plataforma = pSrvMySql, "tmpRptCteDetra (", ps_Prefijo & "tmpRptCteDetra (")
      cCadReporte = cCadReporte & "codaux, razaux, rucaux, sdocumento, scomprobante, fehope, feedoc, bienserv, tipopera, smoneda, importot_mn, impordetra_mn, importot_me, impordetra_me, porceta) VALUES ("
      cCadReporte = cCadReporte & "'" & porstMRp!codaux & "', "
      cCadReporte = cCadReporte & "'" & porstMRp!razAux & "', "
      cCadReporte = cCadReporte & IIf(IsNull(porstMRp!rucaux), "Null", "'" & porstMRp!rucaux & "'") & ", "
      cCadReporte = cCadReporte & "'" & porstMRp!sDocumento & "', "
      cCadReporte = cCadReporte & "'" & porstMRp!sComprobante & "', "
      cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(porstMRp!fehope, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(porstMRp!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sServicio = porstMRp!bienserv
      cCadReporte = cCadReporte & "'" & sServicio & "', "
      cCadReporte = cCadReporte & "'" & porstMRp!tipopera & "', "
      
'ini 2015-07-02 adic tabla detrac
'********
'todo se puso en un funcioni para no repetir codigo
'********
'        Dim uorstCoDetrac As ADODB.Recordset
'        Set uorstCoDetrac = New ADODB.Recordset
'        Set uorstCoDetrac = fRstDetrac(pocnnMain, uorstCoDetrac)
'        With uorstCoDetrac
'            If .RecordCount > 0 Then .MoveFirst
'                .Find "coddetrac3='" & sServicio & "'"
'                If Not .EOF Then
'                    nDetraccion = !tsadetrac
'                End If
'        End With
'        uorstCoDetrac.Close
'        Set uorstCoDetrac = Nothing
      
         nDetraccion = fTsaDetrac(pocnnMain, sServicio)
     
''      'ini 2014-04-06 reclasificacion de cod detraccion
''      For nContador = 1 To UBound(aDtraccDet, 1)
''        If Left(aDtraccDet(nContador), 3) = sServicio Then
''            nDetraccion = aDtraccPor(nContador)
''            Exit For
''        End If
''      Next nContador
'fin 2015-07-02 adic tabla detrac
      
      'fin 2014-04-06 reclasificacion de cod detraccion
      
      cCadReporte = cCadReporte & "'" & porstMRp!sMoneda & "', "
      cCadReporte = cCadReporte & porstMRp!imptot_mn & ", "
      cCadReporte = cCadReporte & Round(porstMRp!imptot_mn * nDetraccion, 2) & ", "
      cCadReporte = cCadReporte & porstMRp!imptot_me & ", "
      cCadReporte = cCadReporte & Round(porstMRp!imptot_me * nDetraccion, 2) & ", "
      cCadReporte = cCadReporte & Round(nDetraccion * 100, 2) & ")"
      pocnnMain.Execute cCadReporte
      porstMRp.MoveNext
    Wend
  End If
  ' Selecciono la información del reporte
  With porstMRp
    If .State = adStateOpen Then .Close
    cCadReporte = "SELECT * "
    cCadReporte = cCadReporte & "FROM " & IIf(ps_Plataforma = pSrvMySql, "tmpRptCteDetra ", ps_Prefijo & "tmpRptCteDetra ")
    cCadReporte = cCadReporte & "ORDER BY codaux, sdocumento"
    .Source = cCadReporte
    .Open
  End With
    
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(OptTipo(0).Value = True, Choose(gsIdioma, "Pendiente", "Pending"), Choose(gsIdioma, "Por Cancelar", "For Liquidate")) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptrctactedetra.rpt"
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
      .LoadReport gsRutRpt & "rptrctactedetra.mrp"
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
            lblDatoDeta(tnIndex).Caption = " " & !razAux
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

