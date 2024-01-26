VERSION 5.00
Begin VB.Form frmRTp56Dro 
   Caption         =   "[título]"
   ClientHeight    =   3930
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMeses 
      Caption         =   " Rango de Meses "
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   30
      TabIndex        =   13
      Top             =   1710
      Width           =   4245
      Begin VB.ComboBox cboMeses 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   2670
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   280
         Width           =   1410
      End
      Begin VB.ComboBox cboMeses 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   280
         Width           =   1410
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Inicio : "
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   345
         Width           =   555
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Fin  : "
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   2235
         TabIndex        =   16
         Top             =   345
         Width           =   345
      End
   End
   Begin VB.CheckBox chkCabecera 
      Caption         =   "Imprime Cabecera"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   45
      TabIndex        =   19
      Top             =   2535
      Width           =   1695
   End
   Begin VB.CheckBox chkFolio 
      Caption         =   "Folio Inicial"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1965
      TabIndex        =   20
      Top             =   2535
      Width           =   1455
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4335
      TabIndex        =   18
      Top             =   1815
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   3615
      TabIndex        =   21
      Top             =   2640
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
         Width           =   1020
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
      ScaleWidth      =   5850
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3390
      Width           =   5850
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
         Picture         =   "frmrtp56dro.frx":0000
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
         Picture         =   "frmrtp56dro.frx":0102
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
         Picture         =   "frmrtp56dro.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkNuevaPagina 
      Caption         =   "Nueva página por cada Diario"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Top             =   1395
      Width           =   2535
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmrtp56dro.frx":077E
      Left            =   4575
      List            =   "frmrtp56dro.frx":0780
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1425
      Width           =   1140
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   15
      TabIndex        =   4
      Top             =   75
      Width           =   5790
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   5415
         Picture         =   "frmrtp56dro.frx":0782
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   5415
         Picture         =   "frmrtp56dro.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   495
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
         TabIndex        =   6
         Top             =   480
         Width           =   570
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
         TabIndex        =   8
         Top             =   840
         Width           =   570
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Diarios"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   480
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
         Left            =   720
         TabIndex        =   7
         Top             =   495
         Width           =   4695
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
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   4695
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
      Height          =   240
      Index           =   1
      Left            =   3900
      TabIndex        =   11
      Top             =   1470
      Width           =   645
   End
End
Attribute VB_Name = "frmRTp56Dro"
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
Private porstCodro As ADODB.Recordset
']


Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
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
   With porstCodro
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
      .Source = .Source & "FROM CODro "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "ORDER BY CodDro"
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

   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodDro"
         .Item(dnContador).MaxLength = porstCodro.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(3, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Diarios:", "Moneda:", "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Journals:", "Currency:", "Beginning :", "End :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraMeses.Caption = Choose(gsIdioma, "Rango de Meses", "Range of Months")
  chkNuevaPagina.Caption = Choose(gsIdioma, "Nueva página por cada Diario", "New page for each Journal")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCodro
      .MoveLast
      txtDato(1).Text = !coddro
      .MoveFirst
      txtDato(0).Text = !coddro
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   chkNuevaPagina.Value = 0
   
  For dnContador = 0 To 13
    If gsIdioma = NvlUsr_Sup Then
      cboMeses(0).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
      cboMeses(1).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
    Else
      cboMeses(0).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
      cboMeses(1).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
    End If
  Next dnContador
  cboMeses(0).ListIndex = Val(gsMesAct)
  cboMeses(1).ListIndex = Val(gsMesAct)
   
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
   porstCodro.Close
   pocnnMain.Close
   Set porstCodro = Nothing
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
  Dim sMoneda As String
  Dim sRegistro As String
  
  ppHabilitacion False
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  ' Genero el query para el reporte
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT LEFT(a.CodDro,2) AS cDiario, a.CodDro, a.NroCpb, a.FehOpe, a.CodTDc, a.SerDoc, a.NroDoc, "
    .Source = .Source & "a.CodCta, a.CodAux, b.RazAux, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, a.TpoCtb, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodAux, '-', b. RazAux)", "(a.CodAux+'-'+b. RazAux)") & " AS cx1, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    .Source = .Source & "c.AbvTDc, " & Choose(gsIdioma, "e.DetDro", "e.DetDrox") & " AS DetDro, "
    .Source = .Source & Choose(gsIdioma, "d.DetDro", "d.DetDrox") & " AS cDetSubDro "
    .Source = .Source & "FROM (((COCpbDet a "
    .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    .Source = .Source & "LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    .Source = .Source & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    .Source = .Source & "LEFT JOIN CODro e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND LEFT(a.CodDro, 2)=RTrim(e.CodDro) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.Mespvs >='" & Format(cboMeses(0).ListIndex, "00") & "' "
    .Source = .Source & "AND a.Mespvs <='" & Format(cboMeses(1).ListIndex, "00") & "' "
    .Source = .Source & "AND d.CodDro BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "ORDER BY a.CodDro, a.NroCpb, a.FehOpe "
    .Open
  End With
   
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptr56droaux.rpt"
      '[ Formulas adicionales del reporte
      sRegistro = Choose(gsIdioma, "Desde ", "From ") & "01/" & Format(cboMeses(0).ListIndex, "00") & "/" & gsAnoAct & Choose(gsIdioma, " al ", " to ") & gfUltDia("01/" & Format(cboMeses(1).ListIndex, "00") & "/" & gsAnoAct)
      .Formulas(5) = "mPeriodo='" & sRegistro & "'"
      sRegistro = FormatNumber(Left(Trim(gsRUCEmp), 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(6) = "mRucEmpresa='" & sRegistro & Mid(Trim(gsRUCEmp), 9) & "'"
      .Formulas(7) = "mDireccion='" & gsDirEmp & "'"
      .Formulas(8) = "mDistrito='" & gsLocEmp & "'"
      .Formulas(9) = "mActividad='" & gsGirEmp & "'"
      .Formulas(10) = "mRepresentante='" & gsRepEmp & "'"
      sRegistro = FormatNumber(Left(gsRepDNIEmp, 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(11) = "mDniRepresentante='" & sRegistro & Mid(gsRepDNIEmp, 9) & "'"
      .Formulas(12) = "sNuevaPagina='" & IIf(chkNuevaPagina.Value, "S", "N") & "'"
      sRegistro = IIf(chkCabecera.Value = vbChecked, "S", "N")
      .ParameterFields(1) = "Cabecera;" & sRegistro & ";true"
      sRegistro = IIf(chkFolio.Value = vbChecked, "S", "N")
      .ParameterFields(2) = "FolioInicial;" & sRegistro & ";true"
      ']
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
      .LoadReport gsRutRpt & "rptRDroAux" & IIf(chkNuevaPagina.Value, "s", "") & ".mrp"
      
      gpEncabezadoMRp MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value
      '[Parámetros adicionales.
      '         .Parameters("pTipoFecha") = IIf(optFecha(0).Value, "Emisión", "Cancelac.")
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
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
      With porstCodro
         .MoveFirst
         .Find "CoddRO='" & txtDato(tnIndex).Text & "'"
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

