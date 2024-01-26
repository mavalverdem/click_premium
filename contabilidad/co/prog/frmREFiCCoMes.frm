VERSION 5.00
Begin VB.Form frmREFiCCoMes 
   Caption         =   "[título]"
   ClientHeight    =   3555
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraNivelCenCos 
      Caption         =   " Nivel de Centro de Costo "
      ForeColor       =   &H80000002&
      Height          =   840
      Left            =   15
      TabIndex        =   8
      Top             =   1455
      Width           =   4335
      Begin VB.OptionButton optNivCCo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "4 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "3 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "2 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optNivCCo 
         Caption         =   "5 dígitos"
         ForeColor       =   &H80000001&
         Height          =   200
         Index           =   3
         Left            =   3000
         TabIndex        =   13
         Top             =   550
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4965
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4245
      TabIndex        =   17
      Top             =   2310
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
         Left            =   1020
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5175
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1485
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   4
      Top             =   90
      Width           =   6420
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   5940
         Picture         =   "frmREFiCCoMes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   135
         TabIndex        =   6
         Top             =   495
         Width           =   765
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
         Left            =   135
         TabIndex        =   7
         Top             =   855
         Width           =   765
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   5940
         Picture         =   "frmREFiCCoMes.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   855
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
         Left            =   900
         TabIndex        =   24
         Top             =   495
         Width           =   5040
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
         Left            =   900
         TabIndex        =   23
         Top             =   855
         Width           =   5040
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centros de Costo"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   1215
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
      ScaleWidth      =   6480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3015
      Width           =   6480
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
         Picture         =   "frmREFiCCoMes.frx":0354
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
         Left            =   4170
         Picture         =   "frmREFiCCoMes.frx":0886
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
         Left            =   2385
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
         Left            =   1275
         Picture         =   "frmREFiCCoMes.frx":09D0
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
      Index           =   1
      Left            =   4410
      TabIndex        =   14
      Top             =   1530
      Width           =   675
   End
End
Attribute VB_Name = "frmREFiCCoMes"
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
Private porstCoCCo          As ADODB.Recordset
']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCoCCo = New ADODB.Recordset
   
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
   With porstCoCCo
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCCo, " & Choose(gsIdioma, "DetCCo", "DetCCox") & " AS DetCCo "
      .Source = .Source & "FROM CoCCo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCCo"
         .Item(dnContador).MaxLength = porstCoCCo.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Centro de Costos :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Cost Center :", "Currency :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraNivelCenCos.Caption = Choose(gsIdioma, "Nivel Centro Costos", "Cost Center Level")
  optNivCCo(4).Caption = Choose(gsIdioma, "Detalle", "Detail")
  optNivCCo(0).Caption = Choose(gsIdioma, "2 dígitos", "2 digits")
  optNivCCo(1).Caption = Choose(gsIdioma, "3 dígitos", "3 digits")
  optNivCCo(2).Caption = Choose(gsIdioma, "4 dígitos", "4 digits")
  optNivCCo(3).Caption = Choose(gsIdioma, "5 dígitos", "5 digits")
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
  For dnContador = 1 To Len(gsNivCCo)
    optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Visible = True
    Select Case dnContador
     Case Is = 1
      optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 120
     Case Is = 2
      optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 1080
     Case Is = 3
      optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 2040
     Case Is = 4
      optNivCCo(Val(Mid(gsNivCCo, dnContador, 1)) - 2).Left = 3000
    End Select
  Next
  optNivCCo(4).Value = True
  fraNivelCenCos.Width = optNivCCo(Val(Mid(gsNivCCo, dnContador - 1, 1)) - 2).Left + 1300
 
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCoCCo
      .MoveLast
      txtDato(1).Text = !CodCCo
      .MoveFirst
      txtDato(0).Text = !CodCCo
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   
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
   
   porstCoCCo.Close
   pocnnMain.Close
   Set porstCoCCo = Nothing
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
  Dim nNivCoCCo As Integer
  
  ppHabilitacion False
    
  nNivCoCCo = IIf(optNivCCo(0).Value, 2, IIf(optNivCCo(1).Value, 3, 5))
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.CodCta, " & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta, a.CodCCo, " & Choose(gsIdioma, "b.DetCCo", "b.DetCCox") & " AS DetCCo, "
    If gsIdioma = NvlUsr_Sup Then
      .Source = .Source & "(CASE LEFT(a.CodCta,2) WHEN '70' THEN 'Total Ingresos por C.Costos' ELSE 'Total Gastos por C.Costos' END) AS cTitulo, "
    Else
      .Source = .Source & "(CASE LEFT(a.CodCta,2) WHEN '70' THEN 'Total Income for Cost Center' ELSE 'Total Expenses for Cost Center' END) AS cTitulo, "
    End If
    .Source = .Source & "(CASE LEFT(a.CodCta,2) WHEN '70' THEN 'cta70' ELSE 'cta9' END) AS cTipo, "
    .Source = .Source & "ROUND(a.AcuD01_" & sMoneda & "-a.AcuH01_" & sMoneda & ", 2) AS cAcu01, "
    .Source = .Source & "ROUND(a.AcuD02_" & sMoneda & "-a.AcuH02_" & sMoneda & ", 2) AS cAcu02, "
    .Source = .Source & "ROUND(a.AcuD03_" & sMoneda & "-a.AcuH03_" & sMoneda & ", 2) AS cAcu03, "
    .Source = .Source & "ROUND(a.AcuD04_" & sMoneda & "-a.AcuH04_" & sMoneda & ", 2) AS cAcu04, "
    .Source = .Source & "ROUND(a.AcuD05_" & sMoneda & "-a.AcuH05_" & sMoneda & ", 2) AS cAcu05, "
    .Source = .Source & "ROUND(a.AcuD06_" & sMoneda & "-a.AcuH06_" & sMoneda & ", 2) AS cAcu06, "
    .Source = .Source & "ROUND(a.AcuD07_" & sMoneda & "-a.AcuH07_" & sMoneda & ", 2) AS cAcu07, "
    .Source = .Source & "ROUND(a.AcuD08_" & sMoneda & "-a.AcuH08_" & sMoneda & ", 2) AS cAcu08, "
    .Source = .Source & "ROUND(a.AcuD09_" & sMoneda & "-a.AcuH09_" & sMoneda & ", 2) AS cAcu09, "
    .Source = .Source & "ROUND(a.AcuD10_" & sMoneda & "-a.AcuH10_" & sMoneda & ", 2) AS cAcu10, "
    .Source = .Source & "ROUND(a.AcuD11_" & sMoneda & "-a.AcuH11_" & sMoneda & ", 2) AS cAcu11, "
    .Source = .Source & "ROUND(a.AcuD12_" & sMoneda & "-a.AcuH12_" & sMoneda & ", 2) AS cAcu12 "
    .Source = .Source & "FROM ((COCCoAcu a "
    .Source = .Source & "LEFT JOIN CoCCo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCCo=b.CodCCo) "
    .Source = .Source & "LEFT JOIN CoCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.CodCCo BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "AND (LEFT(a.CodCta, 2)='70' OR LEFT(a.CodCta, 1)='9') "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(RTrim(a.CodCCo))=" & nNivCoCCo & " "
    .Source = .Source & "AND c.TpoCta='1' "
    .Source = .Source & "ORDER BY a.CodCta, a.CodCCo"
    .Open
  End With
   
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptREFiCCoMes.rpt"
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
      .LoadReport gsRutRpt & "rptREFiCCoMes.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pPeriodoAdc") = gsAnoAct
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
   'Select Case Index    'Completa con ceros a la izquierda.
   'Case 0, 1                           'Cambiar (añadir índices).
   '   If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
   '      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
   '   End If
   'End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1            'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.CCo_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
       With porstCoCCo
          .MoveFirst
          .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
          If .EOF Then
             MsgBox TEXT_8006, vbExclamation
             ppAyuDet = True
          Else
             lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetCCo), "", !DetCCo)
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

End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property
