VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRCCtPHs 
   Caption         =   "[título]"
   ClientHeight    =   6105
   ClientLeft      =   2460
   ClientTop       =   1455
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6285
   Begin VB.CheckBox chkFecha 
      Caption         =   " Rango Fecha "
      ForeColor       =   &H00800000&
      Height          =   190
      Left            =   2310
      TabIndex        =   9
      Top             =   4845
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Frame fraRangos 
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   15
      TabIndex        =   10
      Top             =   4830
      Visible         =   0   'False
      Width           =   3810
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   300
         Left            =   540
         TabIndex        =   12
         Top             =   255
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   47513601
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   2400
         TabIndex        =   14
         Top             =   255
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   47513601
         CurrentDate     =   37953
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   315
         Width           =   240
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "al"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   13
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4590
      Width           =   1080
   End
   Begin VB.CheckBox chkAjuste 
      Caption         =   "Ajuste Diferencia de Cambio"
      ForeColor       =   &H00800000&
      Height          =   190
      Left            =   75
      TabIndex        =   5
      Top             =   4605
      Width           =   3390
   End
   Begin VB.PictureBox picToolBox 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   6255
      TabIndex        =   20
      Top             =   0
      Width           =   6285
      Begin VB.CommandButton cmdSelRango 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   405
         Index           =   2
         Left            =   4860
         Picture         =   "frmRCCtPHs.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Inicializa Rango"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton cmdSelRango 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   405
         Index           =   1
         Left            =   4410
         Picture         =   "frmRCCtPHs.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Establece Fin de Rango"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton cmdSelRango 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   405
         Index           =   0
         Left            =   3945
         Picture         =   "frmRCCtPHs.frx":0CE4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Establece Inicio de Rango"
         Top             =   30
         Width           =   420
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4860
      TabIndex        =   8
      Top             =   4955
      Width           =   1335
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Auxiliar"
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   3765
      Width           =   6255
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
         Left            =   5865
         Picture         =   "frmRCCtPHs.frx":1356
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Width           =   4500
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
      ScaleWidth      =   6285
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5565
      Width           =   6285
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
         TabIndex        =   18
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
         Picture         =   "frmRCCtPHs.frx":1500
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "frmRCCtPHs.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmRCCtPHs.frx":1B7C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   1125
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfgSeleccion 
      Bindings        =   "frmRCCtPHs.frx":1C7E
      Height          =   3210
      Left            =   30
      TabIndex        =   0
      Top             =   555
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5662
      _Version        =   393216
      BackColorFixed  =   16777152
      ForeColorFixed  =   16711680
      BackColorBkg    =   12632256
      AllowBigSelection=   -1  'True
      FocusRect       =   2
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   4335
      TabIndex        =   6
      Top             =   4635
      Width           =   630
   End
End
Attribute VB_Name = "frmRCCtPHs"
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
Private porstCOCta As ADODB.Recordset
Private porstTGAux As ADODB.Recordset
Private nSelInicial As Long, nSelFinal As Long, lSeleccion As Boolean
']
Private Sub ppCabeceraGrilla()
  Dim nIndice As Integer
  
  With mfgSeleccion
    .Cols = 3
    .FixedCols = 1
    .Rows = 2
    .FixedRows = 1
    .GridColor = vbBlack
    .GridColorFixed = vbBlue
    .GridLines = flexGridFlat
    .GridLinesFixed = flexGridInset
    .GridLineWidth = 1
    .SelectionMode = flexSelectionByRow
    .BackColor = &H80000018
    .BackColorBkg = &H8000000F
    .BackColorFixed = &H808040
    .BackColorSel = &HE0E0E0
    .ForeColor = vbBlack
    .ForeColorFixed = vbWhite
    .ForeColorSel = vbBlue
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusNone
    .Font.Bold = True
  End With
  For nIndice = 0 To (mfgSeleccion.Cols - 1)
    mfgSeleccion.Col = nIndice
    If gsIdioma = NvlUsr_Sup Then
      mfgSeleccion.TextMatrix(0, nIndice) = Choose(nIndice + 1, "", "Cuenta", "Descripción")
    Else
      mfgSeleccion.TextMatrix(0, nIndice) = Choose(nIndice + 1, "", "Account", "Description")
    End If
    mfgSeleccion.ColAlignment(nIndice) = Choose(nIndice + 1, flexAlignLeftCenter, flexAlignLeftCenter, flexAlignLeftCenter)
    mfgSeleccion.ColWidth(nIndice) = Choose(nIndice + 1, 300, 1000, 4530)
  Next nIndice
  
End Sub

Private Sub chkFecha_Click()
  fraRangos.Enabled = (chkFecha.Value = vbChecked)
End Sub

Private Sub cmdSelRango_Click(Index As Integer)
  nSelInicial = IIf(Index = 0, mfgSeleccion.RowSel, nSelInicial)
  nSelFinal = IIf(Index = 1, mfgSeleccion.RowSel, nSelFinal)
  ppSelRango mfgSeleccion, Index, nSelInicial, nSelFinal, lSeleccion
End Sub

Private Sub Form_Load()
'   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
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
   With porstCOCta
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
      .Source = .Source & "FROM CoCta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND tpocta='" & TPOCTA_TRA & "' "
      .Source = .Source & "AND inddoc='" & INDAUX_ACT & "' "
      .Source = .Source & "AND tpoanl='" & TPOANL_AUX & "' "
'TC   .Source = .Source & "AND inddoc='" & INDDOC_ACT & "' "
      .Source = .Source & "ORDER BY codcta"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
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

  ' Inicializo la grilla
  mfgSeleccion.Clear
  ppCabeceraGrilla
  mfgSeleccion.Font.Bold = False
  mfgSeleccion.Rows = 1
  While Not porstCOCta.EOF
    dnContador = mfgSeleccion.Rows
    With mfgSeleccion
      .AddItem ""
      .TextMatrix(dnContador, 1) = porstCOCta!codcta
      .TextMatrix(dnContador, 2) = porstCOCta!detcta
    End With
    porstCOCta.MoveNext
  Wend

 '[Parámetros.                         'Cambiar.
   With txtDato
    .Item(0).DataField = "CodAux"
    .Item(0).MaxLength = porstTGAux.Fields(.Item(0).DataField).DefinedSize
   End With
 ']
  
  '[ Cargo los mensajes de botones
  lblTexto(0).Caption = Choose(gsIdioma, "Del", "From")
  lblTexto(1).Caption = Choose(gsIdioma, "al", "to")
  lblTexto(2).Caption = Choose(gsIdioma, "Moneda :", "Currency :")
  fraAuxiliar.Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  chkAjuste.Caption = Choose(gsIdioma, "Ajuste Diferencia de Cambio", "Adjustment Defference of Exchange")
  chkFecha.Caption = Choose(gsIdioma, "Rango Fecha", "Range Date")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  chkAjuste.Value = vbChecked
  fraRangos.Enabled = False
  dtpDesde.Value = CDate("01/" & gsMesAct & "/" & gsAnoAct)
  dtpHasta.Value = gfUltDia(dtpDesde.Value)
 ']
   
 '[Datos predeterminados.              'Cambiar.
   
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
  
  'Otros.
   
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
   porstCOCta.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim s_Sentencia As String, s_Moneda As String
  Dim dnContador As Long, sFechaFin As String
       
  If Not lSeleccion Then MsgBox Choose(gsIdioma, "No ha Seleccionado Ningún Registro; Verificar", "Did not Select any Record; Verify"), vbExclamation: mfgSeleccion.SetFocus: Exit Sub
  ' Valido el rango de periodos
  If chkFecha.Value = vbChecked Then
    sFechaFin = Format(dtpHasta, "yyyymmdd")
    If Not (Left(sFechaFin, 6) <= gsAnoAct & gsMesAct) Then MsgBox Choose(gsIdioma, "Fecha Final debe ser menor o igual que del periodo", "End month must be equal or more than opening balance"), vbExclamation: dtpHasta.SetFocus: Exit Sub
    If Not (Format(dtpDesde, "yyyymmdd") <= sFechaFin) Then MsgBox Choose(gsIdioma, "Fecha Final debe ser mayor o igual que Fecha Inicial", "End date must be equal or more than opening balance"), vbExclamation: dtpDesde.SetFocus: Exit Sub
  End If
  
  ppHabilitacion False
  s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
    
  ' Elimino registros de selccion
  s_Sentencia = "DELETE FROM cotmprpt WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND usrcre='" & gsCodUsr & "'"
  pocnnMain.Execute s_Sentencia
  ' Selecciono la información de cuentas seleccionadas
  'For dnContador = 1 To mfgSeleccion.Rows - 2  Modif TC
  For dnContador = 1 To mfgSeleccion.Rows - 1
    mfgSeleccion.Row = dnContador
    If mfgSeleccion.CellBackColor = &H8000000D Then
      s_Sentencia = "INSERT INTO cotmprpt (codemp, pdoano, codcta, usrcre) "
      s_Sentencia = s_Sentencia & "VALUES ('" & gsCodEmp & "', '" & gsAnoAct & "', '" & Trim(mfgSeleccion.TextMatrix(dnContador, 1)) & "', '" & gsCodUsr & "')"
      pocnnMain.Execute s_Sentencia
    End If
  Next dnContador
  
  ' Elimino y genero el archivo temporal de documentos pendientes
  s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpdocumento_') DROP TABLE #tmpdocumento"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocumento", s_Sentencia)
  
  s_Sentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdocumento ", "")
  s_Sentencia = s_Sentencia & "SELECT distinct det.CodAux, det.CodCta, tdc.AbvTDc, det.CodTDc, det.SerDoc, det.NroDoc, aux.RazAux, "
  s_Sentencia = s_Sentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE det.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN det.Imp" & s_Moneda & " ELSE 0 END)), 0), 2) AS Debe, "
  s_Sentencia = s_Sentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE det.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN det.Imp" & s_Moneda & " ELSE 0 END)), 0), 2) AS Haber, "
  s_Sentencia = s_Sentencia & "det.pdoano,det.mespvs,det.coddro,det.nrocpb,det.nroite "
  s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpdocumento ")
  s_Sentencia = s_Sentencia & "FROM (((cocpbdet det "
  s_Sentencia = s_Sentencia & "LEFT JOIN tgaux aux ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
  s_Sentencia = s_Sentencia & "LEFT JOIN tgtdc tdc ON det.codemp=tdc.codemp AND det.codtdc=tdc.codtdc) "
  s_Sentencia = s_Sentencia & "LEFT JOIN cotmprpt tmp ON det.codemp=tmp.codemp AND det.pdoano=tmp.pdoano AND det.codcta=tmp.codcta) "
  s_Sentencia = s_Sentencia & "WHERE det.codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND det.pdoano='" & gsAnoAct & "' "
  If chkFecha.Value = vbChecked Then
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "AND DATE_FORMAT(det.FeEDoc, '%Y%m%d') >='" & Format(dtpDesde, "yyyymmdd") & "' "
      s_Sentencia = s_Sentencia & "AND DATE_FORMAT(det.FeEDoc, '%Y%m%d') <='" & Format(dtpHasta, "yyyymmdd") & "' "
    Else
      s_Sentencia = s_Sentencia & "AND CONVERT(smalldatetime, det.FeEDoc, 103) >='" & Format(dtpDesde, "dd/mm/yyyy") & "' "
      s_Sentencia = s_Sentencia & "AND CONVERT(smalldatetime, det.FeEDoc, 103) <='" & Format(dtpHasta, "dd/mm/yyyy") & "' "
    End If
  Else
    s_Sentencia = s_Sentencia & "AND det.MesPvs <='" & gsMesAct & "' "
  End If
  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.CodAux, '') <>'' "
  's_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.CodTDc, '') <>'' "
  's_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.SerDoc, '') <>'' "
  's_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.NroDoc, '') <>'' "
  s_Sentencia = s_Sentencia & "AND tmp.usrcre='" & gsCodUsr & "' "
  If Trim(txtDato(0).Text) <> "" Then
    s_Sentencia = s_Sentencia & "AND det.CodAux='" & txtDato(0).Text & "' "
  End If
  If chkAjuste.Value = vbUnchecked Then
    s_Sentencia = s_Sentencia & "AND det.tpognr<>" & TPOGNR_DCA & " "
  End If
  s_Sentencia = s_Sentencia & "GROUP BY det.Mespvs, det.Coddro, det.Nrocpb, det.Nroite "
  's_Sentencia = s_Sentencia & "GROUP BY det.CodCta, det.CodAux, det.CodTDc, det.SerDoc, det.NroDoc, tdc.AbvTDc, aux.RazAux "
  If ps_Plataforma = pSrvMySql Then
  '  s_Sentencia = s_Sentencia & "HAVING ROUND(Debe - Haber, 2) <> 0.00 "
  Else
  '  s_Sentencia = s_Sentencia & "HAVING ROUND(ROUND(ISNULL(SUM((CASE det.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN det.Imp" & s_Moneda & " ELSE 0 END)), 0), 2) - "
  '  s_Sentencia = s_Sentencia & "ROUND(ISNULL(SUM((CASE det.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN det.Imp" & s_Moneda & " ELSE 0 END)), 0), 2), 2) <> 0.00 "
  End If
  s_Sentencia = s_Sentencia & "ORDER BY det.CodAux, det.CodCta, det.CodTDc, det.SerDoc, det.NroDoc"
  pocnnMain.Execute s_Sentencia
  
  ' Obtengo la información para el reporte
  s_Sentencia = "SELECT distinct " & IIf(chkFecha.Value = vbChecked, "det.FeEDoc", "Null") & " AS quiebre, det.MesPvs, det.CodCta, det.CodAux, det.CodTDc, det.SerDoc, det.NroDoc, det.CodDro, det.NroCpb, "
  s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tmp.AbvTDc,'-', det.SerDoc,'-', det.NroDoc)", "(tmp.AbvTDc+'-'+det.SerDoc+'-'+det.NroDoc)") & " AS cDocum, "
  s_Sentencia = s_Sentencia & "det.FehOpe, det.FeEDoc, det.FeVDoc, det.RefDoc, " & Choose(gsIdioma, "det.GloIte", "det.GloItex") & " AS GloIte, tmp.RazAux, "
  s_Sentencia = s_Sentencia & "(CASE det.TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
  s_Sentencia = s_Sentencia & "(CASE det.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN det.Imp" & s_Moneda & " ELSE 0 END) AS cDebe, "
  s_Sentencia = s_Sentencia & "(CASE det.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN det.Imp" & s_Moneda & " ELSE 0 END) AS cHaber "
  s_Sentencia = s_Sentencia & "FROM COCpbDet det "
  s_Sentencia = s_Sentencia & "INNER JOIN " & ps_Prefijo & "tmpdocumento tmp ON det.CodAux=tmp.codaux "
'TC s_Sentencia = s_Sentencia & "AND det.CodCta=tmp.codcta AND det.CodTDc=tmp.codtdc "
'TC s_Sentencia = s_Sentencia & "AND det.SerDoc=tmp.serdoc AND det.NroDoc=tmp.nrodoc "
  s_Sentencia = s_Sentencia & "AND det.CodCta=tmp.codcta and det.mespvs=tmp.mespvs and det.coddro=tmp.coddro and det.nrocpb=tmp.nrocpb and det.nroite=tmp.nroite "
  s_Sentencia = s_Sentencia & "WHERE det.codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND det.pdoano='" & gsAnoAct & "' "
  If chkFecha.Value = vbChecked Then
    If ps_Plataforma = pSrvMySql Then
      s_Sentencia = s_Sentencia & "AND DATE_FORMAT(det.FeEDoc, '%Y%m%d') >='" & Format(dtpDesde, "yyyymmdd") & "' "
      s_Sentencia = s_Sentencia & "AND DATE_FORMAT(det.FeEDoc, '%Y%m%d') <='" & Format(dtpHasta, "yyyymmdd") & "' "
    Else
      s_Sentencia = s_Sentencia & "AND CONVERT(smalldatetime, det.FeEDoc, 103) >='" & Format(dtpDesde, "dd/mm/yyyy") & "' "
      s_Sentencia = s_Sentencia & "AND CONVERT(smalldatetime, det.FeEDoc, 103) <='" & Format(dtpHasta, "dd/mm/yyyy") & "' "
    End If
  Else
    s_Sentencia = s_Sentencia & "AND det.MesPvs <='" & gsMesAct & "' "
  End If
  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.Codcta, '') <>'' "
  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.Codaux, '') <>'' "
  's_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.SerDoc, '') <>'' "
  's_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.NroDoc, '') <>'' "
  If chkAjuste.Value = vbUnchecked Then
    s_Sentencia = s_Sentencia & "AND det.tpognr<>" & TPOGNR_DCA & " "
  End If
  s_Sentencia = s_Sentencia & "ORDER BY quiebre, det.CodAux, det.CodTDc, det.SerDoc, det.NroDoc, det.MesPvs, " & IIf(chkFecha.Value = vbChecked, "", "det.FeEDoc,") & " det.TpoPvs DESC, det.CodCta"
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = s_Sentencia
    .Open
  End With

  usDEstino = PRN_DEST_GRAF
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptrcctpehs.rpt"
      .ParameterFields(1) = "Quiebre;" & IIf(chkFecha.Value = vbChecked, "1", "0") & ";true"
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  End If
  ' Elimino registros de seleccion
  s_Sentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpdocumento_') DROP TABLE #tmpdocumento"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocumento", s_Sentencia)
  pocnnMain.Execute "DELETE FROM cotmprpt WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND usrcre='" & gsCodUsr & "'"
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
   'Select Case Index    'Completa con ceros a la izquierda.
   'Case 0, 1                           'Cambiar (añadir índices).
   '   If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
   '      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
   '   End If
   'End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2                          'Cambiar (añadir índices).
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
   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
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

