VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRCCtCta 
   Caption         =   "[título]"
   ClientHeight    =   4005
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkVerificar 
      Caption         =   "Comprobación"
      ForeColor       =   &H00800000&
      Height          =   190
      Left            =   30
      TabIndex        =   28
      Top             =   2400
      Width           =   1560
   End
   Begin VB.Frame fraFecha 
      ForeColor       =   &H00800000&
      Height          =   585
      Left            =   2850
      TabIndex        =   14
      Top             =   2700
      Width           =   2175
      Begin MSComCtl2.DTPicker dtpFechaVence 
         Height          =   300
         Left            =   540
         TabIndex        =   16
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   19005441
         CurrentDate     =   37953
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CheckBox chkFecha 
      Caption         =   "Fecha de Vencimiento"
      ForeColor       =   &H00800000&
      Height          =   190
      Left            =   2865
      TabIndex        =   13
      Top             =   2505
      Width           =   2085
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   18
      Top             =   2640
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   19
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   20
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   30
      TabIndex        =   10
      Top             =   2640
      Width           =   2760
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H80000004&
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1425
         TabIndex        =   12
         Top             =   315
         Width           =   1200
      End
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H80000004&
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rangos"
      ForeColor       =   &H80000002&
      Height          =   2130
      Left            =   0
      TabIndex        =   4
      Top             =   45
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   6900
         Picture         =   "frmRCCtCta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1680
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
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   1665
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6600
         Picture         =   "frmRCCtCta.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6600
         Picture         =   "frmRCCtCta.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   540
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
         Left            =   150
         TabIndex        =   7
         Top             =   885
         Width           =   945
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
         Left            =   150
         TabIndex        =   6
         Top             =   525
         Width           =   945
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Auxiliar"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   1440
         Width           =   495
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
         Index           =   2
         Left            =   1380
         TabIndex        =   27
         Top             =   1665
         Width           =   5520
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
         Left            =   1080
         TabIndex        =   26
         Top             =   885
         Width           =   5520
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
         Left            =   1080
         TabIndex        =   25
         Top             =   540
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   585
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3465
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
         Left            =   3720
         Picture         =   "frmRCCtCta.frx":04FE
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
         Picture         =   "frmRCCtCta.frx":0648
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
         Picture         =   "frmRCCtCta.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRCCtCta"
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
']

Private Sub chkFecha_Click()
  fraFecha.Enabled = (chkFecha.Value = vbChecked)
End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset   'Cuentas
   Set porstTGAux = New ADODB.Recordset   'Auxiliar
   
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
    .Source = .Source & "FROM COCta "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "ORDER BY CodCta"
  '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
   End With
 
   With porstTGAux
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodAux, RazAux "
      .Source = .Source & "FROM TgAux "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "ORDER BY CodAux"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 
 ']

 '[Parámetros.                         'Cambiar.
   
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
      For dnContador = 2 To 2
         .Item(dnContador).DataField = "CodAux"
         .Item(dnContador).MaxLength = porstTGAux.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuentas :", "Auxiliar :", "Del :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Accounts :", "Auxiliary :", "From :")
  Next nElemento
  
  dtpFechaVence.MinDate = CDate("01/" & gsMesAct & "/" & gsAnoAct)
  dtpFechaVence.MaxDate = gfUltDia(dtpFechaVence.MinDate)
  dtpFechaVence.Value = dtpFechaVence.MaxDate
  fraFecha.Enabled = False
  
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  chkVerificar.Caption = Choose(gsIdioma, "Comprobación", "Checking")
  chkFecha.Caption = Choose(gsIdioma, "Fecha de Vencimiento", "Range Date")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !codcta
      .MoveFirst
      txtDato(0).Text = !codcta
   End With
  
   With porstTGAux
      '.MoveLast
      'txtDato(1).Text = !CodCta
      'Beto
      '.MoveFirst
      'txtDato(2).Text = !CodAux
   End With
  
  
  
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
   If txtDato(2).Text <> "" Then ppAyuDet 2
  
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
   porstCOCta.Close
   porstTGAux.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
   Case 2     ', 3
      txtDato(Index).SetFocus
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim cCadReporte  As String, sTitulo As String
   
  ppHabilitacion False
  ' primero: Genero la información detallada
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpdocumento_') DROP TABLE #tmpdocumento"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocumento", cCadReporte)
  
  cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdocumento ", "")
  cCadReporte = cCadReporte & "SELECT a.codemp, a.pdoano, a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, MIN(a.FeEDoc) AS FeEDoc, MIN(a.FeVDoc) AS FeVDoc, "
  cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(d.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(d.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
  cCadReporte = cCadReporte & "(CASE b.TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS DebeSol, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) AS HaberSol, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS DebeDol, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2) AS HaberDol, "
  cCadReporte = cCadReporte & "c.RucAux, c.RazAux, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta "
  cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpdocumento ")
  cCadReporte = cCadReporte & "FROM (((COCpbDet a "
  cCadReporte = cCadReporte & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
  cCadReporte = cCadReporte & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
  cCadReporte = cCadReporte & "LEFT JOIN TGTDc d ON a.codemp=d.codemp AND a.CodTDc=d.CodTDc) "
  cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
  cCadReporte = cCadReporte & "AND a.pdoano='" & gsAnoAct & "' "
  cCadReporte = cCadReporte & "AND a.MesPvs <= '" & gsMesAct & "' "
  cCadReporte = cCadReporte & "AND LEFT(a.CodCta, " & Len(Trim(txtDato(0).Text)) & ")>='" & txtDato(0).Text & "' "
  cCadReporte = cCadReporte & "AND LEFT(a.CodCta, " & Len(Trim(txtDato(1).Text)) & ")<='" & txtDato(1).Text & "' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodTDc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.SerDoc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.NroDoc, '') <>'' AND b.inddoc=" & INDDOC_ACT & " "
  If (Trim(txtDato(2).Text) <> "" And chkVerificar.Value = Unchecked) Then
    cCadReporte = cCadReporte & "AND a.CodAux='" & txtDato(2).Text & "' "
  End If
  cCadReporte = cCadReporte & "GROUP BY a.codemp, a.pdoano, a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, d.AbvTDc, c.RucAux, c.RazAux, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & ", b.tpomon "
  If ps_Plataforma = pSrvMySql Then
    cCadReporte = cCadReporte & "HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
  Else
    cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 0), 2) - "
    cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 0), 2), 2) <> 0.00 "
    cCadReporte = cCadReporte & "OR ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END)), 0), 2) - "
    cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END)), 0), 2), 2) <> 0.00) "
  End If
  cCadReporte = cCadReporte & "ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
  pocnnMain.Execute cCadReporte
  
  ' segundo: Genero la información provisiones
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpdocuprv_') DROP TABLE #tmpdocuprv"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocuprv", cCadReporte)
  
  cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdocuprv ", "")
  cCadReporte = cCadReporte & "SELECT DISTINCT det.codemp, det.pdoano, det.mespvs, det.codcta, det.codaux, det.codtdc, det.serdoc, det.nrodoc, det.coddro, det.nrocpb, "
  cCadReporte = cCadReporte & "det.feedoc, det.fevdoc, det.refdoc, " & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
  cCadReporte = cCadReporte & "(CASE WHEN det.tpomon='" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS csigno "
  cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpdocuprv ")
  cCadReporte = cCadReporte & "FROM cocpbdet det "
  cCadReporte = cCadReporte & "INNER JOIN cocta cta ON det.codemp=cta.codemp AND det.pdoano=cta.pdoano AND det.codcta=cta.codcta "
  cCadReporte = cCadReporte & "WHERE det.codemp='" & gsCodEmp & "' "
  cCadReporte = cCadReporte & "AND det.pdoano='" & gsAnoAct & "' "
  cCadReporte = cCadReporte & "AND det.mespvs<= '" & gsMesAct & "' "
  cCadReporte = cCadReporte & "AND LEFT(det.codcta, " & Len(Trim(txtDato(0).Text)) & ")>='" & txtDato(0).Text & "' "
  cCadReporte = cCadReporte & "AND LEFT(det.codcta, " & Len(Trim(txtDato(1).Text)) & ")<='" & txtDato(1).Text & "' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.codaux, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.codtdc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.serdoc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.nrodoc, '') <>'' "
  cCadReporte = cCadReporte & "AND det.TpoPvs='" & TPOPVS_PVS & "' AND cta.inddoc=" & INDDOC_ACT & " "
  If (Trim(txtDato(2).Text) <> "" And chkVerificar.Value = Unchecked) Then
    cCadReporte = cCadReporte & "AND det.codaux='" & txtDato(2).Text & "' "
  End If
  cCadReporte = cCadReporte & "ORDER BY det.codcta, det.codaux, det.codtdc, det.serdoc, det.nrodoc"
  pocnnMain.Execute cCadReporte
  
  ' Obtengo las fechas iniciales
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpdetalle_') DROP TABLE #tmpdetalle"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdetalle", cCadReporte)
  
  cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdetalle ", "")
  cCadReporte = cCadReporte & "SELECT DISTINCT a.codemp, a.pdoano, a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, b.CodDro, b.NroCpb, "
  cCadReporte = cCadReporte & "(CASE WHEN b.FeEDoc IS NULL THEN a.FeEDoc ELSE b.FeEDoc END) AS FeEDoc, (CASE WHEN b.FeVDoc IS NULL THEN a.FeVDoc ELSE b.FeVDoc END) AS FeVDoc, "
  cCadReporte = cCadReporte & "b.RefDoc, b.GloIte, a.cDocume, b.cSigno, "
  cCadReporte = cCadReporte & "a.DebeSol, a.HaberSol, a.DebeDol, a.HaberDol, a.RucAux, a.RazAux, a.DetCta "
  cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpdetalle ")
  cCadReporte = cCadReporte & "FROM " & ps_Prefijo & "tmpdocumento a "
  cCadReporte = cCadReporte & "LEFT JOIN " & ps_Prefijo & "tmpdocuprv b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc "
  cCadReporte = cCadReporte & "WHERE b.codemp='" & gsCodEmp & "' "
  cCadReporte = cCadReporte & "AND b.pdoano='" & gsAnoAct & "' "
  cCadReporte = cCadReporte & "AND b.MesPvs <= '" & gsMesAct & "' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodAux, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodTDc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.SerDoc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.NroDoc, '') <>'' "
  cCadReporte = cCadReporte & "ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
  pocnnMain.Execute cCadReporte
  ' Elimino temporales
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpdocumento_') DROP TABLE #tmpdocumento"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocumento", cCadReporte)
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpdocuprv_') DROP TABLE #tmpdocuprv"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdocuprv", cCadReporte)
  If chkVerificar.Value = Unchecked Then
    If OptTipo(0).Value Then
      cCadReporte = "SELECT a.CodCta, a.CodAux, a.SerDoc, a.NroDoc, a.CodDro, a.NroCpb, a.FeEDoc, a.FeVDoc, "
      'cCadReporte = cCadReporte & "a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, a.cDocume, a.cSigno, a.DebeSol, a.HaberSol, a.DebeDol, a.HaberDol, "
      cCadReporte = cCadReporte & "a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloIte") & " AS GloIte, a.cDocume, a.cSigno, a.DebeSol, a.HaberSol, a.DebeDol, a.HaberDol, "
      cCadReporte = cCadReporte & "a.RucAux, a.RazAux, a.DetCta "
      cCadReporte = cCadReporte & "FROM " & ps_Prefijo & "tmpdetalle a "
      If chkFecha.Value = vbChecked Then
        If ps_Plataforma = pSrvMySql Then
          cCadReporte = cCadReporte & "WHERE DATE_FORMAT(a.FeVDoc, '%Y-%m-%d')='" & Format(dtpFechaVence.Value, "yyyy-mm-dd") & "' "
        Else
          cCadReporte = cCadReporte & "WHERE CONVERT(smalldatetime, a.FeVDoc, 103)='" & Format(dtpFechaVence.Value, "dd/mm/yyyy") & "' "
        End If
      End If
      cCadReporte = cCadReporte & "ORDER BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
    ElseIf OptTipo(1).Value Then
      cCadReporte = "SELECT a.CodCta, a.RUCAux, a.RazAux, a.DetCta, "
      cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.DebeSol), 0), 2) AS DebeSol, "
      cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.HaberSol), 0), 2) AS HaberSol, "
      cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.DebeDol), 0), 2) AS DebeDol, "
      cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(a.HaberDol), 0), 2) AS HaberDol "
      cCadReporte = cCadReporte & "FROM " & ps_Prefijo & "tmpdetalle a "
      If chkFecha.Value = vbChecked Then
        If ps_Plataforma = pSrvMySql Then
          cCadReporte = cCadReporte & "WHERE DATE_FORMAT(a.FeVDoc, '%Y-%m-%d')='" & Format(dtpFechaVence.Value, "yyyy-mm-dd") & "' "
        Else
          cCadReporte = cCadReporte & "WHERE CONVERT(smalldatetime, a.FeVDoc, 103)='" & Format(dtpFechaVence.Value, "dd/mm/yyyy") & "' "
        End If
      End If
      cCadReporte = cCadReporte & "GROUP BY a.CodCta, a.CodAux, a.RUCAux, a.RazAux, a.DetCta "
      If ps_Plataforma = pSrvMySql Then
        cCadReporte = cCadReporte & " HAVING (ROUND(DebeSol - HaberSol, 2) <> 0.00 OR ROUND(DebeDol - HaberDol, 2) <> 0.00) "
      Else
        cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM(a.DebeSol), 0), 2) - ROUND(ISNULL(SUM(a.HaberSol), 0), 2), 2) <> 0.00 "
        cCadReporte = cCadReporte & "OR ROUND(ROUND(ISNULL(SUM(a.DebeDol), 0), 2) - ROUND(ISNULL(SUM(a.HaberDol), 0), 2), 2) <> 0.00) "
      End If
      cCadReporte = cCadReporte & " ORDER BY a.CodCta, a.CodAux"
    End If
  Else    ' verificacion
    ' temporal acumulado
    cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpanaliza_') DROP TABLE #tmpanaliza"
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpanaliza", cCadReporte)
    
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpanaliza ", "")
    cCadReporte = cCadReporte & "SELECT DISTINCT det.codemp, det.pdoano, det.codcta, "
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & "det.codaux, det.codtdc, det.serdoc, det.nrodoc, "
    End If
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END)), 0), 2) AS debemn, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END)), 0), 2) AS habemn, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END)), 0), 2) AS debeme, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END)), 0), 2) AS habeme, "
    cCadReporte = cCadReporte & "000000000000.00 AS impctemn, 000000000000.00 AS impcteme "
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpanaliza ")
    cCadReporte = cCadReporte & "FROM cocpbdet det "
    cCadReporte = cCadReporte & "INNER JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    cCadReporte = cCadReporte & "WHERE det.codemp='" & gsCodEmp & "' "
    cCadReporte = cCadReporte & "AND det.pdoano='" & gsAnoAct & "' "
    cCadReporte = cCadReporte & "AND det.mespvs <= '" & gsMesAct & "' "
    cCadReporte = cCadReporte & "AND LEFT(det.codcta, " & Len(Trim(txtDato(0).Text)) & ")>='" & Trim(txtDato(0).Text) & "' "
    cCadReporte = cCadReporte & "AND LEFT(det.codcta, " & Len(Trim(txtDato(1).Text)) & ")<='" & Trim(txtDato(1).Text) & "' "
    cCadReporte = cCadReporte & "AND cta.inddoc=" & INDDOC_ACT & " "
    cCadReporte = cCadReporte & "GROUP BY det.codemp, det.pdoano, det.codcta "
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & ", det.codaux, det.codtdc, det.serdoc, det.nrodoc "
    End If
    If ps_Plataforma = pSrvMySql Then
      cCadReporte = cCadReporte & "HAVING (ROUND(debemn - habemn, 2) <> 0.00 OR ROUND(debeme- habeme, 2) <> 0.00) "
    Else
      cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END)), 0), 2) - "
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END)), 0), 2), 2) <> 0.00 "
      cCadReporte = cCadReporte & "OR ROUND(ROUND(ISNULL(SUM((CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END)), 0), 2) - "
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END)), 0), 2), 2) <> 0.00) "
    End If
    cCadReporte = cCadReporte & "ORDER BY det.codcta"
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & ", det.codaux, det.codtdc, det.serdoc, det.nrodoc"
    End If
    pocnnMain.Execute cCadReporte
    
    ' Inserto cuenta corriente
    cCadReporte = "INSERT INTO " & ps_Prefijo & "tmpanaliza "
    cCadReporte = cCadReporte & "SELECT DISTINCT det.codemp, det.pdoano, det.codcta, "
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & "det.codaux, det.codtdc, det.serdoc, det.nrodoc, "
    End If
    cCadReporte = cCadReporte & "0.00 AS debemn, 0.00 AS habemn, 0.00 AS debeme, 0.00 AS habeme, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(det.debesol-det.habersol), 0), 2) AS impctemn, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(det.debedol-det.haberdol), 0), 2) AS impcteme "
    cCadReporte = cCadReporte & "FROM " & ps_Prefijo & "tmpdetalle det "
    cCadReporte = cCadReporte & "GROUP BY det.codemp, det.pdoano, det.codcta "
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & ", det.codaux, det.codtdc, det.serdoc, det.nrodoc "
    End If
    If ps_Plataforma = pSrvMySql Then
      cCadReporte = cCadReporte & "HAVING (impctemn <> 0.00 OR impcteme <> 0.00) "
    Else
      cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM(det.DebeSol), 0), 2) - ROUND(ISNULL(SUM(det.HaberSol), 0), 2), 2) <> 0.00 "
      cCadReporte = cCadReporte & "OR ROUND(ROUND(ISNULL(SUM(det.DebeDol), 0), 2) - ROUND(ISNULL(SUM(det.HaberDol), 0), 2), 2) <> 0.00) "
    End If
    cCadReporte = cCadReporte & "ORDER BY det.codcta"
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & ", det.codaux, det.codtdc, det.serdoc, det.nrodoc"
    End If
    pocnnMain.Execute cCadReporte
    
    ' sentencia de seleccion
    cCadReporte = "SELECT DISTINCT det.codcta, cta.detcta,  "
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & "det.codaux, aux.razaux, "
      cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(tdc.abvtdc, '-', det.serdoc, '-', det.nrodoc)", "(tdc.abvtdc + '-' + det.serdoc + '-' + det.nrodoc)") & " AS cDocume, "
    End If
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(det.debemn-det.habemn), 0), 2) AS impbalmn, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(det.debeme-det.habeme), 0), 2) AS impbalme, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(det.impctemn), 0), 2) AS impctemn, "
    cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(det.impcteme), 0), 2) AS impcteme "
    cCadReporte = cCadReporte & "FROM " & ps_Prefijo & "tmpanaliza  det "
    cCadReporte = cCadReporte & "INNER JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & "LEFT JOIN TGAux aux ON aux.codemp=det.codemp AND aux.codaux=det.codaux "
      cCadReporte = cCadReporte & "LEFT JOIN TGTDc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc "
    End If
    cCadReporte = cCadReporte & "GROUP BY det.codemp, det.pdoano, det.codcta, cta.detcta "
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & ", det.codaux, aux.razaux, tdc.abvtdc, det.codtdc, det.serdoc, det.nrodoc "
    End If
    If ps_Plataforma = pSrvMySql Then
      cCadReporte = cCadReporte & "HAVING (impbalmn <> impctemn OR impbalme <> impcteme) "
    Else
      cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM(det.debemn), 0), 2) - ROUND(ISNULL(SUM(det.habemn), 0), 2), 2) <> 0.00 "
      cCadReporte = cCadReporte & "OR ROUND(ROUND(ISNULL(SUM(det.debeme), 0), 2) - ROUND(ISNULL(SUM(det.habeme), 0), 2), 2) <> 0.00) "
    End If
    cCadReporte = cCadReporte & "ORDER BY det.codcta"
    If OptTipo(0).Value Then
      cCadReporte = cCadReporte & ", det.codaux, det.codtdc, det.serdoc, det.nrodoc"
    End If
  End If
  
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = cCadReporte
    .Open
  End With
   
  sTitulo = IIf(chkFecha.Value = vbChecked, Choose(gsIdioma, " - del ", " - from ") & Format(dtpFechaVence, "dd/mm/yyyy"), "")
  sTitulo = IIf(chkVerificar.Value = Unchecked, sTitulo, "")
  sTitulo = Me.Caption & " (" & IIf(OptTipo(0).Value, Choose(gsIdioma, "Detalle", "Detail"), Choose(gsIdioma, "Resumen", "Summary")) & IIf(chkVerificar.Value = vbChecked, Choose(gsIdioma, " - Comprobación", " - Checking"), "") & sTitulo & ")"
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, sTitulo, udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRCCtCta" & IIf(OptTipo(0).Value, IIf(chkVerificar.Value = Unchecked, "Det", "Cpbde"), IIf(chkVerificar.Value = Unchecked, "Res", "Cpbre")) & ".rpt"
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
        .LoadReport gsRutRpt & "rptRCCtCtaDet.mrp"
      Else
        .LoadReport gsRutRpt & "rptRCCtCtaRes.mrp"
      End If
      gpEncabezadoMRp MRViewer, sTitulo, udFecha, True, chkImpFecha.Value
      '[Parámetros adicionales.
      .Parameters("pPeriodoAct") = "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
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
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpanaliza_') DROP TABLE #tmpanaliza"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpanaliza", cCadReporte)
  
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 12)='#tmpdetalle_') DROP TABLE #tmpdetalle"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdetalle", cCadReporte)
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
   If txtDato(2) = "" Then
       lblDatoDeta(2).Caption = ""
       'lblDatoDeta(0).Caption = "Todos..."
   End If

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
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   Case 2    ', 1                           'Cambiar (añadir índices). - Auxiliar
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 2         ', 1                           'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
      With porstCOCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
         End If
      End With
   
   Case 2   ', 1  - Auxiliar
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

