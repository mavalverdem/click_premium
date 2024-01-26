VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTPdoGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   1860
   ClientTop       =   2010
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   8775
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8775
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
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
         Height          =   560
         Index           =   1
         Left            =   2880
         Picture         =   "frmTPdoGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdRevisar 
         Caption         =   "&Revisar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   720
         Picture         =   "frmTPdoGrd.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdRefrescar 
         Caption         =   "Re&frescar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   2160
         Picture         =   "frmTPdoGrd.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   720
      End
      Begin VB.Frame fraBuscar 
         Caption         =   "&Buscar por [Columna]"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   3650
         TabIndex        =   0
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   200
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   8055
         Picture         =   "frmTPdoGrd.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   0
         Picture         =   "frmTPdoGrd.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   1440
         Picture         =   "frmTPdoGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTPdoGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public uocnnMain As ADODB.Connection
Public uocnnNoGrabable As ADODB.Connection
Public uorstMain As ADODB.Recordset
Public uorstMain_Grd As ADODB.Recordset
Private psConnStrgSele As String, _
        psConnStrgSele_Grd As String, _
        psConnStrgOrde As String
Private pnColumnaOrd As Integer

'[Propio del formulario.
Public uorstTGAux As ADODB.Recordset
Public uorstTGTCb As ADODB.Recordset
Public uorstCoDPe As ADODB.Recordset
Public uorstCoDPeCta As ADODB.Recordset
Public uorstCoCta As ADODB.Recordset
Public uorstCoCCo As ADODB.Recordset
Public uorstCoCprProd As ADODB.Recordset
Public uorstCoPdoCprProd As ADODB.Recordset
Public porstCancel As ADODB.Recordset

Public usConnStrgSele_CoPdoCprCta As String
Public usConnStrgWher_CoPdoCprCta As String
Public usConnStrgOrde_CoPdoCprCta As String

Public usConnStrgSele_CoPdoCprProd As String
Public usConnStrgWher_CoPdoCprProd As String
Public usConnStrgOrde_CoPdoCprProd As String
']

Private Sub cmdImprimir_Click(Index As Integer)
  If uorstMain.RecordCount = 0 Then
     MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
     Exit Sub
  End If
 '[Datos del formulario de impresión.  'Cambiar.
   frmLPdo.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLPdo.Show vbModal
 ']
End Sub

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
  psConnStrgSele_Grd = "SELECT copdocpr.coddpe, copdocpr.pdocpr, copdocpr.codaux, b.razaux, copdocpr.Fehpdo, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "copdocpr." & Choose(gsIdioma, "detpdo", "detpdox") & " AS cdetpdo, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "copdocpr.tpomon, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE copdocpr.tpomon WHEN '" & TPOMON_NAC & "' THEN copdocpr.impmn ELSE copdocpr.impme END) AS cImporte, copdocpr.nrointerno, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "Concat(copdocpr.coddpe, copdocpr.pdocpr, copdocpr.codaux)", "(copdocpr.coddpe+copdocpr.pdocpr+copdocpr.codaux)") & " AS cLlave "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "FROM CoPdoCpr "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGAux b ON copdocpr.codemp=b.codemp AND copdocpr.CodAux=b.CodAux "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "WHERE copdocpr.codemp='" & gsCodEmp & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND copdocpr.pdoano='" & gsAnoAct & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND copdocpr.mespvs='" & gsMesAct & "' "
  
  psConnStrgSele = "SELECT copdocpr.coddpe, copdocpr.pdocpr, copdocpr.codaux, "
  psConnStrgSele = psConnStrgSele & "copdocpr.fehpdo, copdocpr.detpdo, copdocpr.detpdox, "
  psConnStrgSele = psConnStrgSele & "copdocpr.tpomon, copdocpr.imptcb, copdocpr.impmn, copdocpr.impme, "
  psConnStrgSele = psConnStrgSele & "copdocpr.impdife, copdocpr.indcta, copdocpr.indext, copdocpr.nrointerno, "
  psConnStrgSele = psConnStrgSele & "copdocpr.UsrCre, copdocpr.FyHCre, copdocpr.UsrMdf, copdocpr.FyHMdf, "
  psConnStrgSele = psConnStrgSele & "copdocpr.codemp, copdocpr.pdoano, copdocpr.mespvs, "
  psConnStrgSele = psConnStrgSele & "copdocpr.tpoigv, " '2014-05-22
  psConnStrgSele = psConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "Concat(copdocpr.coddpe, copdocpr.pdocpr, copdocpr.codaux)", "(copdocpr.coddpe+copdocpr.pdocpr+copdocpr.codaux)") & " AS cLlave, "
  psConnStrgSele = psConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "Concat(copdocpr.coddpe, copdocpr.pdocpr)", "(copdocpr.coddpe+copdocpr.pdocpr)") & " AS cLlave1 "
  psConnStrgSele = psConnStrgSele & "FROM copdocpr "
  psConnStrgSele = psConnStrgSele & "WHERE copdocpr.codemp='" & gsCodEmp & "' "
  psConnStrgSele = psConnStrgSele & "AND copdocpr.pdoano='" & gsAnoAct & "' "
  psConnStrgSele = psConnStrgSele & "AND copdocpr.mespvs='" & gsMesAct & "' "
  psConnStrgOrde = "ORDER BY copdocpr.coddpe, copdocpr.pdocpr, copdocpr.codaux"
  
  usConnStrgSele_CoPdoCprCta = "SELECT codcta, codcco, impcta_mn, impcta_me, impctadif, "
  usConnStrgSele_CoPdoCprCta = usConnStrgSele_CoPdoCprCta & "mespvs, coddpe, pdocpr, "
  usConnStrgSele_CoPdoCprCta = usConnStrgSele_CoPdoCprCta & "usrcre, fyhcre, usrmdf, fyhmdf, "
  usConnStrgSele_CoPdoCprCta = usConnStrgSele_CoPdoCprCta & "codemp, pdoano "
  usConnStrgSele_CoPdoCprCta = usConnStrgSele_CoPdoCprCta & "FROM " & ps_Prefijo & "tmpcopdocprcta "
  usConnStrgWher_CoPdoCprCta = ""
  usConnStrgOrde_CoPdoCprCta = "ORDER BY 7, 8, 1"
  
  usConnStrgSele_CoPdoCprProd = "SELECT codprod, " & Choose(gsIdioma, "gloprod", "gloprodx") & ", cantiprod, impprod_mn, impprod_me, "
  usConnStrgSele_CoPdoCprProd = usConnStrgSele_CoPdoCprProd & "impouni_mn, impouni_me, " & Choose(gsIdioma, "gloprodx", "gloprod") & ", indigv, "
  usConnStrgSele_CoPdoCprProd = usConnStrgSele_CoPdoCprProd & "codcta, codcco, mespvs, coddpe, pdocpr, "
  usConnStrgSele_CoPdoCprProd = usConnStrgSele_CoPdoCprProd & "usrcre, fyhcre, usrmdf, fyhmdf, "
  usConnStrgSele_CoPdoCprProd = usConnStrgSele_CoPdoCprProd & "codemp, pdoano "
  usConnStrgSele_CoPdoCprProd = usConnStrgSele_CoPdoCprProd & "FROM " & ps_Prefijo & "tmpcopdocprprod "
  usConnStrgWher_CoPdoCprProd = ""
  usConnStrgOrde_CoPdoCprProd = "ORDER BY 13, 14, 1"
  
  Set uocnnMain = New ADODB.Connection
  Set uocnnNoGrabable = New ADODB.Connection
  Set uorstMain = New ADODB.Recordset
  Set uorstMain_Grd = New ADODB.Recordset
  Set uorstTGAux = New ADODB.Recordset
  Set uorstTGTCb = New ADODB.Recordset
  Set uorstCoDPe = New ADODB.Recordset
  Set uorstCoDPeCta = New ADODB.Recordset
  Set uorstCoCta = New ADODB.Recordset
  Set uorstCoCCo = New ADODB.Recordset
  Set uorstCoCprProd = New ADODB.Recordset
  Set uorstCoPdoCprProd = New ADODB.Recordset
  Set porstCancel = New ADODB.Recordset
  With uocnnMain
     .CursorLocation = adUseClient
     .ConnectionString = CONNSTRG & gsNomBDS
     .Open
  End With
  With uocnnNoGrabable
     .CursorLocation = adUseClient
     .ConnectionString = CONNSTRG & gsNomBDS
     .Open
  End With
  With uorstMain_Grd
     .ActiveConnection = uocnnMain
     .Source = psConnStrgSele_Grd & psConnStrgOrde
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic 'adLockReadOnly
     .Open
     .Properties("Unique Table").Value = "copdocpr"
  End With
  With uorstMain
     .ActiveConnection = uocnnMain
     .Source = psConnStrgSele & psConnStrgOrde
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic 'adLockReadOnly
     .Open
     .Properties("Unique Table").Value = "copdocpr"
  End With
  With uorstCoDPeCta
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
  End With
  With uorstTGTCb
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.FehTCb, a.ImpTCb_Cpr, a.ImpTCb_Vta "
     .Source = .Source & "FROM TGTCb a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
  With uorstCoDPe
    .ActiveConnection = uocnnMain
    .Source = "SELECT coddpe, " & Choose(gsIdioma, "detdpe", "detdpex") & " AS detdpe, codcco "
    .Source = .Source & "FROM codpe "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    '.Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(coddpe)=4"
''     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
  With uorstCoCta
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.CodCta, " & Choose(gsIdioma, "a.DetCta", "a.DetCtax") & " AS DetCta, a.TpoTCb, a.IndDoc, a.IndCCo, a.codcco_def "
     .Source = .Source & "FROM COCta a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND a.TpoCta=" & TPOCTA_TRA & " "
     .Source = .Source & "AND a.EstCta='" & ESTCTA_ACT & "'"
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  With uorstCoCCo
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.CodCCo, " & Choose(gsIdioma, "a.DetCCo", "a.DetCCox") & " AS DetCCo "
     .Source = .Source & "FROM COCCo a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND a.indpdocpr='" & INDCCT_ACT & "' "
     .Source = .Source & "AND a.EstCCo='" & ESTCCO_ACT & "' "
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodCCo)>2"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  ' productos de compras
  With uorstCoCprProd
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.codprod, " & Choose(gsIdioma, "a.detprod", "a.detprodx") & " AS detprod, a.unimed, "
    .Source = .Source & "a.codcta, a.tpomon, a.impcpr, b.codcco_def "
    .Source = .Source & "FROM cocprprod a "
    .Source = .Source & "INNER JOIN cocta b ON b.codemp=a.codemp AND b.pdoano=a.pdoano AND a.codcta=b.codcta "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "'"
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  With uorstCoPdoCprProd
    .ActiveConnection = uocnnMain
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
  End With
  
  With porstCancel
    .ActiveConnection = uocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockBatchOptimistic ' adLockOptimistic
  End With
  With uorstTGAux
     .ActiveConnection = uocnnNoGrabable
     .Source = "SELECT a.CodAux, a.RazAux "
     .Source = .Source & "FROM TGAux a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND a.IndPrv=1 "
     .Source = .Source & "AND a.EstAux='" & ESTAUX_ACT & "'"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  ']
  '[ Elimino y creo tabla temporal de cuentas contables
  If ps_Plataforma = pSrvMySql Then
    uocnnMain.Execute "DROP TABLE IF EXISTS tmpcopdocprcta"
    uocnnMain.Execute "CREATE TEMPORARY TABLE tmpcopdocprcta SELECT * FROM copdocprcta WHERE pdocpr='tmppedido'"
  ElseIf ps_Plataforma = pSrvSql Then
    ' Activo detector de errores
    On Error Resume Next
    uocnnMain.Execute "DROP TABLE " & ps_Prefijo & "tmpcopdocprcta"
    If Not (Err.Number = -2147217865 Or Err.Number = 0) Then
      MsgBox Err.Description, vbInformation
    End If
    On Error GoTo 0
    uocnnMain.Execute "SELECT * INTO #tmpcopdocprcta FROM copdocprcta WHERE pdocpr='tmppedido'"
  End If
  ']
  '[ Elimino y creo tabla temporal de productos pedido
  If ps_Plataforma = pSrvMySql Then
    uocnnMain.Execute "DROP TABLE IF EXISTS tmpcopdocprprod"
    uocnnMain.Execute "CREATE TEMPORARY TABLE tmpcopdocprprod SELECT * FROM copdocprprod WHERE pdocpr='tmpproducto'"
  ElseIf ps_Plataforma = pSrvSql Then
    ' Activo detector de errores
    On Error Resume Next
    uocnnMain.Execute "DROP TABLE " & ps_Prefijo & "tmpcopdocprprod"
    If Not (Err.Number = -2147217865 Or Err.Number = 0) Then
      MsgBox Err.Description, vbInformation
    End If
    On Error GoTo 0
    uocnnMain.Execute "SELECT * INTO #tmpcopdocprprod FROM copdocprprod WHERE pdocpr='tmpproducto'"
  End If
  ']
  
  dgrMain.MarqueeStyle = dbgHighlightRow
  Set dgrMain.DataSource = uorstMain_Grd
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
  ']
End Sub

Private Sub Form_Activate()
  'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
  zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
  upDatosGrid
  fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub
Private Sub Form_Resize()
   On Error Resume Next
  
   gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   uorstTGAux.Close
   uorstTGTCb.Close
   uorstCoDPe.Close
   uorstCoCta.Close
   uorstCoCCo.Close
   If uorstCoDPeCta.State = adStateOpen Then uorstCoDPeCta.Close
   uorstCoCprProd.Close
   If uorstCoPdoCprProd.State = adStateOpen Then uorstCoPdoCprProd.Close
   uorstMain_Grd.Close
   uorstMain.Close
   uocnnMain.Close
   Set porstCancel = Nothing
   Set uorstTGAux = Nothing
   Set uorstTGTCb = Nothing
   Set uorstCoDPe = Nothing
   Set uorstCoCta = Nothing
   Set uorstCoCCo = Nothing
   Set uorstCoDPeCta = Nothing
   Set uorstCoCprProd = Nothing
   Set uorstCoPdoCprProd = Nothing
   Set uorstMain_Grd = Nothing
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub

Private Sub cmdNuevo_Click()
  '[Propio del formulario.
  'Verificación de Mes Cerrado.
  If gbCieCpr Then MsgBox TEXT_9016, vbCritical: Exit Sub
  ']
  gpTUg_Nuevo Me, frmTPdo             'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
  On Error GoTo Err
  
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub

  '[Búsqueda del ítem.
  uorstMain.Requery
  uorstMain.MoveFirst
  uorstMain.Find "cLlave='" & uorstMain_Grd!coddpe & uorstMain_Grd!pdocpr & uorstMain_Grd!codaux & "'"
  ']

  With frmTPdo                        'Cambiar Formulario de Datos.
    .zbNuevo = False
    .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
    .txtLlave(0).Enabled = False
    .chkExtension.Enabled = False
    .txtDato(0).Enabled = False
    .cmdDatoAyud(0).Enabled = False
    ']
    .Caption = TEXT_MODIF & " " & Me.Caption
    .Show vbModal
  End With
  dgrMain.SetFocus
  
  Exit Sub
Err:
  gpErrores
End Sub

Public Sub cmdEliminar_Click()
  Dim dsLlaveSiguiente As String
  
  On Error GoTo Err
  
  'Verificación de Mes Cerrado.
  If gbCieCpr Then MsgBox TEXT_9016, vbCritical: Exit Sub
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub
  
 'ini 2016-05-27/28 nivel=asisten no elimin datos
   If gsNvlUsr = NVLUSR_ASIS Then
      MsgBox TEXT_9026, vbCritical
      Exit Sub
   End If
'fin 2016-05-27/28 nivel=asisten no elimin datos
  ' Mensaje de verificación            'Cambiar.
  If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & "-" & Trim(dgrMain.Columns(2)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
    With porstCancel
      .Source = "SELECT mespvs, pdocpr, codaux "
      .Source = .Source & "FROM cocprdoc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND mespvs='" & gsMesAct & "' AND codaux='" & uorstMain_Grd!codaux & "' "
      .Source = .Source & "AND pdocpr='" & uorstMain_Grd!coddpe & uorstMain_Grd!pdocpr & "' "
      .Source = .Source & "UNION "
      .Source = .Source & "SELECT mespvs, pdocpr, codaux "
      .Source = .Source & "FROM cohprdoc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND mespvs='" & gsMesAct & "' AND codaux='" & uorstMain_Grd!codaux & "' "
      .Source = .Source & "AND pdocpr='" & uorstMain_Grd!coddpe & uorstMain_Grd!pdocpr & "'"
      .Open
      If porstCancel.RecordCount = 0 Then
        uorstMain.MoveFirst
        uorstMain.Find "cLlave = '" & uorstMain_Grd!coddpe & uorstMain_Grd!pdocpr & uorstMain_Grd!codaux & "'"
        
        uocnnMain.BeginTrans       'INICIA TRANSACCION.
        uorstMain.Properties("Unique Table").Value = "copdocpr"
        uorstMain.Delete
        uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.

        'Busca siguiente ítem.
        With uorstMain_Grd
          .MoveNext
          If .EOF Then .MoveLast
          dsLlaveSiguiente = !coddpe & !pdocpr & !codaux
          .Requery
          If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
        End With
        upDatosGrid
        ' actualizo recordset principal
        uorstMain.Requery
        If uorstMain.RecordCount > 0 Then uorstMain.Find "cLlave = '" & dsLlaveSiguiente & "'"
      Else
        MsgBox Choose(gsIdioma, "Debe eliminar antes las Provisiones.", " The Provisions must be eliminated before."), vbExclamation
      End If
    End With
    porstCancel.Close
  End If
  dgrMain.SetFocus
  Exit Sub
Err:
  gpErrores
  
  uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
'[ARREGLAR. Usar gpTUg_Refrescar Me, pero se debe cambiar ppDatosGrid a upDatosGrid para todos los _
            formularios que lo usan (formularios de registro único).
''   gpTUg_Refrescar Me
   uorstMain_Grd.Requery
   upDatosGrid
   
   dgrMain.SetFocus
']ARREGLAR.
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
'[ARREGLAR. No acepta ordenar por columna de tablas secundarias en el recordset.
   If ColIndex = 3 Then Exit Sub
']ARREGLAR.

   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   psConnStrgOrde = "ORDER BY "
   Select Case pnColumnaOrd            'Cambiar.
'   Case 1
'      psConnStrgOrde = psConnStrgOrde & "2, 3, 4"
   Case Else
      psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
   End Select
   With uorstMain_Grd
      .Close
      .Properties("Unique Table").Value = "copdocpr"
      .Source = psConnStrgSele_Grd & psConnStrgOrde
      .Open
   End With
   Set dgrMain.DataSource = uorstMain_Grd
   upDatosGrid

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
  If uorstMain_Grd.RecordCount = 0 Then Exit Sub
  
  Select Case KeyCode
   Case vbKeyHome
    uorstMain_Grd.MoveFirst
   Case vbKeyEnd
    uorstMain_Grd.MoveLast
  End Select
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With uorstMain_Grd
      dvRegistroActual = .Bookmark
   
'[ARREGLAR: Búsqueda con distintos tipos de columna.
      Select Case VarType(.Fields(pnColumnaOrd))
      Case vbString
         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " LIKE '" & Trim(txtBuscar) & "*'"
      Case vbInteger, vbSingle, vbByte, vbDouble, vbLong, vbDecimal
         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
'     Case vbDate
'         dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
      End Select
      .Find dsCriterio, , , 1
      If .EOF = True Then
         .Bookmark = dvRegistroActual
      End If
   End With
']ARREGLAR.
   
   Exit Sub
Err:
   If Err.Number = 3001 Then   'Se produce al llegar a EOF de adcMain.
      uorstMain_Grd.Bookmark = dvRegistroActual
   Else
      gpErrores
   End If
End Sub

Public Sub upDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Proyecto", "Project")
            .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("coddpe").DefinedSize + 1)
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Pedido", "Order")
            .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("pdocpr").DefinedSize + 1)
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("codaux").DefinedSize + 0.5)
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "Razón Social", "Firm Name")
            .Item(dnNum).Width = 1600
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "F.Emisión", "Issue Date")
            .Item(dnNum).Width = 980
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Detalle", "Detail")
            .Item(dnNum).Width = 1600
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "Mon", "Cur")
            .Item(dnNum).Width = 250
         Case 7
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe", "Amount")
            .Item(dnNum).Width = 1200
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
            .Item(dnNum).Alignment = dbgRight
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'[Código propio del formulario.

']

Private Property Get znColumnaOrd() As Integer
   znColumnaOrd = pnColumnaOrd
End Property
Private Property Let znColumnaOrd(ByVal tnColumnaOrd As Integer)
   pnColumnaOrd = tnColumnaOrd
End Property

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdNuevo.Enabled = taOpciones(0)
   cmdEliminar.Enabled = taOpciones(1)
   cmdImprimir(1).Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property
