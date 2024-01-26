VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPPDTVta 
   Caption         =   "[título]"
   ClientHeight    =   2205
   ClientLeft      =   2640
   ClientTop       =   3960
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmPPdtVta.frx":0000
      Left            =   3345
      List            =   "frmPPdtVta.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
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
      Left            =   1725
      Picture         =   "frmPPdtVta.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Reporte de validación"
      Top             =   1485
      Width           =   1150
   End
   Begin MSComDlg.CommonDialog CmnDlgUbica 
      Left            =   165
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   375
      TabIndex        =   2
      Top             =   1485
      Width           =   1150
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   3060
      TabIndex        =   1
      Top             =   1485
      Width           =   1150
   End
   Begin ComctlLib.ProgressBar pgbEtapa1 
      Height          =   345
      Left            =   225
      TabIndex        =   0
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
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
      Left            =   2490
      TabIndex        =   6
      Top             =   165
      Width           =   765
   End
   Begin VB.Label lblTexto 
      Caption         =   "Procesando Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   4
      Top             =   690
      Width           =   2355
   End
End
Attribute VB_Name = "frmPPDTVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private udFecha As Date
Private unCopias As Integer
Private unMargenIzquierdo As Integer
Private usDEstino As String
Private usOrientacionRpt As String
Private usOrientacionOri As String

Private pocnnMain As ADODB.Connection
Public pocnnConf As ADODB.Connection
Public porstCOVtaDoc As ADODB.Recordset
Public porstTGEMP As ADODB.Recordset
Public pbNuevo As Boolean

Private Sub cmdImprimir_Click(Index As Integer)
  On Error GoTo Err
  
  Dim porstMRp As New ADODB.Recordset
  Dim sSentencia As String, sMoneda As String
  Dim nRegistros As Long
  
  cmdAceptar.Enabled = False
  cmdImprimir(0).Enabled = False
  cmdSalir.Enabled = False
  
  ' Aperturo la conexión
  Set pocnnMain = New ADODB.Connection
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  ' Instancio el recordset de reporte
  With porstMRp
    .ActiveConnection = pocnnMain
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  '[ Registro de ingresos
  sSentencia = "SELECT a.mespvs, a.codtdc, a.serdoc, "
  sSentencia = sSentencia & "ROUND(SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN ((a.ImpOGr_" & sMoneda & "+a.ImpExo_" & sMoneda & ") * -1) ELSE (a.ImpOGr_" & sMoneda & "+a.ImpExo_" & sMoneda & ") END), 2) AS nVtaTotal, "
  sSentencia = sSentencia & "ROUND(SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_" & sMoneda & " * -1) ELSE a.ImpOGr_" & sMoneda & " END), 2) AS nVtaGrava "
  sSentencia = sSentencia & "FROM CoVtaDoc a "
  sSentencia = sSentencia & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux "
  sSentencia = sSentencia & "LEFT JOIN TgTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.mespvs NOT IN('00', '13') "
  sSentencia = sSentencia & "AND a.codtdc IN('01', '03', '06', '07', '08', '12') "
  sSentencia = sSentencia & "GROUP BY a.mespvs, a.codtdc, a.serdoc "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING nVtaTotal <> 0.00 "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_" & sMoneda & "+a.ImpExo_" & sMoneda & ") * -1 ELSE (a.ImpOGr_" & sMoneda & "+a.ImpExo_" & sMoneda & ") END) > 0.00 "
  End If
  sSentencia = sSentencia & "ORDER BY a.mespvs, a.codtdc, a.serdoc"
  ' Aperturo el listado de registros
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = sSentencia
    .Open
  End With
  gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "Ventas Anuales", "Sales Year") & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, False, False, porstMRp
  With frmMain.rptMain
    .ReportFileName = gsRutRpt & "rptRPDTVta.rpt"
    ' Formulas adicionales
    .Formulas(5) = "mPeriodo='" & Choose(gsIdioma, "Ejercicio - ", "Fiscal year - ") & gsAnoAct & "'"
    .WindowState = crptMaximized
    .MarginLeft = unMargenIzquierdo
    .Destination = crptToWindow
    .Action = 1
  End With
  ']
  cmdAceptar.Enabled = True
  cmdImprimir(0).Enabled = True
  pocnnMain.Close
  Set pocnnMain = Nothing
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  Exit Sub
  
Err:
  Set porstMRp = Nothing
  If pocnnMain.State = adStateOpen Then pocnnMain.Close
  Set pocnnMain = Nothing
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus

End Sub

Private Sub Form_Activate()
  cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
  ' On Error GoTo Err
   
  Dim dnContador As Integer
 
  cmdImprimir(0).Enabled = False
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  pgbEtapa1.Value = 0

  'Declaración de Variables.
   
  'Abrir Tablas.
   Set pocnnMain = New ADODB.Connection
   Set pocnnConf = New ADODB.Connection
   Set porstTGEMP = New ADODB.Recordset
   Set porstCOVtaDoc = New ADODB.Recordset

   With pocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG  & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With pocnnConf
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDC
      .Open
   End With
   With porstTGEMP
      .ActiveConnection = pocnnConf
      .CursorType = adOpenStatic
      .LockType = adLockReadOnly
   End With
'   pocnnMain.BeginTrans                'INICIA TRANSACCION.
 
  ' Generando Texto segun lectura de Tabla.
   dnContador = 0
   pgbEtapa1.Min = 0
   pgbEtapa1.Value = pgbEtapa1.Min
   
   With porstCOVtaDoc
    .ActiveConnection = pocnnMain
    .Source = "SELECT a.mespvs, a.codtdc, a.serdoc, "
    If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
      .Source = .Source & "ROUND(SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN ((a.ImpOGr_MN+a.ImpExo_MN) * -1) ELSE (a.ImpOGr_MN+a.ImpExo_MN) END), 2) AS nVtaTotal, "
      .Source = .Source & "ROUND(SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_MN * -1) ELSE a.ImpOGr_MN END), 2) AS nVtaGrava "
    Else
      .Source = .Source & "ROUND(SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN ((a.ImpOGr_ME+a.ImpExo_ME) * -1) ELSE (a.ImpOGr_ME+a.ImpExo_ME) END), 2) AS nVtaTotal, "
      .Source = .Source & "ROUND(SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_ME * -1) ELSE a.ImpOGr_ME END), 2) AS nVtaGrava "
    End If
    .Source = .Source & "FROM CoVtaDoc a "
    .Source = .Source & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux "
    .Source = .Source & "LEFT JOIN TgTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.mespvs NOT IN('00', '13') "
    .Source = .Source & "AND a.codtdc IN('01', '03', '06', '07', '08', '12') "
    .Source = .Source & "GROUP BY a.mespvs, a.codtdc, a.serdoc "
    If ps_Plataforma = pSrvMySql Then
      .Source = .Source & "HAVING nVtaTotal <> 0.00 "
    ElseIf ps_Plataforma = pSrvSql Then
      If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        .Source = .Source & "HAVING SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_MN+a.ImpExo_MN) * -1 ELSE (a.ImpOGr_MN+a.ImpExo_MN) END) > 0.00 "
      Else
        .Source = .Source & "HAVING SUM(CASE c.SgnTDc WHEN " & SGNTDC_NEG & " THEN (a.ImpOGr_ME+a.ImpExo_ME) * -1 ELSE (a.ImpOGr_ME+a.ImpExo_ME) END) > 0.00 "
      End If
    End If
    .Source = .Source & "ORDER BY a.mespvs, a.codtdc, a.serdoc"
    
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
   End With
   ppEtapa_01
   
   porstCOVtaDoc.Close
   pocnnConf.Close
   pocnnMain.Close
   Set porstTGEMP = Nothing
   Set porstCOVtaDoc = Nothing
   Set pocnnConf = Nothing
   Set pocnnMain = Nothing
   
   MsgBox TEXT_8008, vbInformation
   cmdImprimir(0).Enabled = True
   cmdAceptar.Enabled = True
   cmdSalir.Enabled = True
   cmdSalir.SetFocus
   
   Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub ppEtapa_01()   ' Generacion de Texto en File Ingresos
  Dim dnContador As Integer
  Dim dsTexto, dsFile As String
  
  dnContador = 0
  pgbEtapa1.Min = 0
  With porstTGEMP
    .Source = "Select RucEmp From TGEMP Where CodEmp='" & gsCodEmp & "'"
    .Open
  End With
  dsFile = "3550" & porstTGEMP!RUCEmp & gsAnoAct & ".txt"
  CmnDlgUbica.FileName = dsFile
  CmnDlgUbica.ShowSave
  Open dsFile For Output As #1
  Do
    With porstCOVtaDoc
      If .RecordCount = 0 Then
        Exit Do
      End If
      .MoveFirst
      pgbEtapa1.Max = .RecordCount
      pgbEtapa1.Value = pgbEtapa1.Min
      Do
        dnContador = dnContador + 1
        dsTexto = ""
        dsTexto = dsTexto & !CodTDc & "|"
        dsTexto = dsTexto & !SerDoc & "|"
        dsTexto = dsTexto & !mespvs & "|"
        dsTexto = dsTexto & Trim(Str(Abs(gfRedond(!nvtatotal, 2)))) & "|"
        dsTexto = dsTexto & Trim(Str(Abs(gfRedond(!nvtagrava, 2)))) & "|"
        Print #1, dsTexto
        pgbEtapa1.Value = dnContador
        .MoveNext
      Loop Until .EOF
    End With
    Exit Do
  Loop
  Close #1
  porstTGEMP.Close

End Sub

Private Sub Form_Load()
  
 '[Parámetros.                         'Cambiar.
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :", "Procesando Ventas")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :", "Processing Sales")
  Next nElemento
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, True, False, False, False, False, False, True, aLabel
 ']
  
  'Características de impresión.
  udFecha = Date                      'Fecha en el encabezado.
  unCopias = 1                        'Cantidad de Copias.
  unMargenIzquierdo = 240             'Margen izquierdo.
  usDEstino = PRN_DEST_GRAF           'PRN_DEST_GRAF:ica _
                                       PRN_DEST_MATR:icial.
  usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                       PRN_ORIE_HORI:zontal.

End Sub
