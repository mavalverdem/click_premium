VERSION 5.00
Begin VB.Form frmLVta 
   Caption         =   "[título]"
   ClientHeight    =   2880
   ClientLeft      =   2460
   ClientTop       =   3060
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5775
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   2115
      Left            =   0
      TabIndex        =   11
      Top             =   105
      Width           =   5775
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   5400
         Picture         =   "frmLVta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
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
         TabIndex        =   4
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
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1740
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
         TabIndex        =   6
         Top             =   1320
         Width           =   1755
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   5400
         Picture         =   "frmLVta.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   5400
         Picture         =   "frmLVta.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   5600
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de Documento"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1485
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
         Height          =   300
         Index           =   0
         Left            =   675
         TabIndex        =   15
         Top             =   480
         Width           =   4725
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
         Height          =   300
         Index           =   2
         Left            =   1755
         TabIndex        =   14
         Top             =   1680
         Width           =   3645
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
         Height          =   300
         Index           =   1
         Left            =   1755
         TabIndex        =   13
         Top             =   1320
         Width           =   3645
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Documento"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1650
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
      ScaleWidth      =   5775
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2340
      Width           =   5775
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
         Left            =   4620
         Picture         =   "frmLVta.frx":04FE
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
         Picture         =   "frmLVta.frx":0648
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
         Picture         =   "frmLVta.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmLVta"
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
Private porstFiltro As ADODB.Recordset
Private porstRango As ADODB.Recordset
Private sSelectSql As String, sWhereSql As String, sOrderSql As String
']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstFiltro = New ADODB.Recordset
   Set porstRango = New ADODB.Recordset
   
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
   With porstFiltro
      .ActiveConnection = pocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Source = "SELECT CodTDc, " & Choose(gsIdioma, "DetTDc", "DetTDcx") & " AS DetTDc "
      .Source = .Source & "FROM TGTDc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "ORDER BY CodTDc"
      .Open
   End With
  '[Datos predeterminados.              'Cambiar.
   txtDato(0).DataField = "CodTDc"
   txtDato(0).MaxLength = porstFiltro.Fields(txtDato(0).DataField).DefinedSize
   txtDato(0).Text = porstFiltro.Fields(txtDato(0).DataField)
  
  sSelectSql = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc, '-', NroDoc)", "(SerDoc+'-'+NroDoc)") & " AS cDocumento, GloDoc, SerDoc, NroDoc "
  sSelectSql = sSelectSql & "FROM CoVtaDoc "
  sSelectSql = sSelectSql & "WHERE codemp='" & gsCodEmp & "' "
  sSelectSql = sSelectSql & "AND pdoano='" & gsAnoAct & "' "
  sSelectSql = sSelectSql & "AND MesPvs='" & gsMesAct & "' "
  sWhereSql = "AND CodTDc='" & txtDato(0).Text & "' "
  sOrderSql = "ORDER BY SerDoc, NroDoc"
   With porstRango
    .ActiveConnection = pocnnMain
'    .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Source = sSelectSql & sWhereSql & sOrderSql
    .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
 ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Tipo de Documento :", "Documentos de Venta :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Type of Document :", "Documents of Sale :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
  
 '[Datos predeterminados.              'Cambiar.
  With txtDato
    For dnContador = 1 To 2
      .Item(dnContador).DataField = "cDocumento"
      .Item(dnContador).MaxLength = porstRango.Fields(.Item(dnContador).DataField).DefinedSize
      .Item(dnContador).Text = ""
      lblDatoDeta(dnContador) = ""
    Next
  End With
  'Límites de rangos.
  If Not (porstRango.EOF And porstRango.BOF) Then
    With porstRango
      .MoveLast
      txtDato(2).Text = .Fields(txtDato(2).DataField)
      .MoveFirst
      txtDato(1).Text = .Fields(txtDato(1).DataField)
    End With
  End If
  'Otros.
  If txtDato(0).Text <> "" Then ppAyuDet 0
  If txtDato(1).Text <> "" Then ppAyuDet 1
  If txtDato(2).Text <> "" Then ppAyuDet 2
   
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
   porstRango.Close
   pocnnMain.Close
   Set porstRango = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim s_Sentencia As String, sDocumento As String
  Dim sMoneda As String, sImporteLetras As String
  Dim nImporteTotal As Double, nRegistros As Long
  Dim nFormato As Integer, nRegistro As Integer
  Dim nContador As Integer, nDiferencia As Integer, nLen As Integer
  
  ppHabilitacion False
  
  ' Genero la tabla temporal de reporte
  If ps_Plataforma = pSrvMySql Then
    pocnnMain.Execute "DROP TABLE IF EXISTS trptdocventa"
    s_Sentencia = "CREATE TEMPORARY TABLE IF NOT EXISTS trptdocventa (documento varchar(16) NOT NULL, "
    s_Sentencia = s_Sentencia & "secuencia smallint(1) DEFAULT '0', codtdc char(2) NOT NULL, "
    s_Sentencia = s_Sentencia & "serdoc char(4) NOT NULL, nrodoc varchar(10) NOT NULL, "
    s_Sentencia = s_Sentencia & "emision date NULL, modifica date NULL, "
    s_Sentencia = s_Sentencia & "codaux varchar(11) NULL, razaux varchar(80) NULL, " '2014-07-29 aumeto rsocial
    s_Sentencia = s_Sentencia & "diraux varchar(80) NULL, rucaux varchar(11) NULL, "
    s_Sentencia = s_Sentencia & "tpomon char(1) NULL, signomon char(3) NULL, "
    s_Sentencia = s_Sentencia & "pctigv decimal(4,2) DEFAULT '0', refdoc varchar(20) NULL, "
    s_Sentencia = s_Sentencia & "glodet0 varchar(250) NULL, glodet1 varchar(250) NULL, "
    s_Sentencia = s_Sentencia & "glortc varchar(250) NULL, "
    s_Sentencia = s_Sentencia & "impbase decimal(12,2) DEFAULT '0', impigv decimal(12,2) DEFAULT '0', "
    s_Sentencia = s_Sentencia & "imptotal decimal(12,2) DEFAULT '0', dettdc varchar(40) NULL, "
    s_Sentencia = s_Sentencia & "forimp smallint(1) DEFAULT '0', importeletra varchar(250) NULL, "
    s_Sentencia = s_Sentencia & "PRIMARY KEY (documento, secuencia))"
  ElseIf ps_Plataforma = pSrvSql Then
    pocnnMain.Execute "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptdocventa') DROP TABLE #trptdocventa"
    s_Sentencia = "CREATE TABLE #trptdocventa (documento varchar(16) NOT NULL, "
    s_Sentencia = s_Sentencia & "secuencia smallint DEFAULT '0', codtdc char(2) NOT NULL, "
    s_Sentencia = s_Sentencia & "serdoc char(4) NOT NULL, nrodoc varchar(10) NOT NULL, "
    s_Sentencia = s_Sentencia & "emision smalldatetime NULL, modifica smalldatetime NULL, "
    s_Sentencia = s_Sentencia & "codaux varchar(11) NULL, razaux varchar(80) NULL, " '2014-07-29 aumeto rsocial
    s_Sentencia = s_Sentencia & "diraux varchar(80) NULL, rucaux varchar(11) NULL, "
    s_Sentencia = s_Sentencia & "tpomon char(1) NULL, signomon char(3) NULL, "
    s_Sentencia = s_Sentencia & "pctigv decimal(4,2) DEFAULT '0', refdoc varchar(20) NULL, "
    s_Sentencia = s_Sentencia & "glodet0 varchar(250) NULL, glodet1 varchar(250) NULL, "
    s_Sentencia = s_Sentencia & "glortc varchar(250) NULL, "
    s_Sentencia = s_Sentencia & "impbase decimal(12,2) DEFAULT '0', impigv decimal(12,2) DEFAULT '0', "
    s_Sentencia = s_Sentencia & "imptotal decimal(12,2) DEFAULT '0', dettdc varchar(40) NULL, "
    s_Sentencia = s_Sentencia & "forimp smallint DEFAULT '0',  importeletra varchar(250) NULL, "
    s_Sentencia = s_Sentencia & "PRIMARY KEY (documento, secuencia))"
  End If
  pocnnMain.Execute s_Sentencia
  
  s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.codtdc, det.serdoc, det.nrodoc)", "(det.codtdc+det.serdoc+det.nrodoc)") & " AS documento, det.codtdc, det.serdoc, det.nrodoc, vta.fehope AS emision, vta.feedoc AS modifica, "
  s_Sentencia = s_Sentencia & "vta.codaux, aux.razaux, aux.diraux, aux.rucaux, vta.tpomon, vta.pctigv, vta.refdoc, "
  s_Sentencia = s_Sentencia & "det.glodet0, det.glodet1, vta.glodoc_rtc AS glortc, "
  s_Sentencia = s_Sentencia & "(CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impcta_mn ELSE det.impcta_me END) AS impbase, "
  s_Sentencia = s_Sentencia & "(CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN vta.impigv_mn ELSE vta.impigv_me END) AS impigv, "
  s_Sentencia = s_Sentencia & "(CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN vta.imptot_mn ELSE vta.imptot_me END) AS imptotal, "
  s_Sentencia = s_Sentencia & "doc.dettdc , doc.forimp, (CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS signomon "
  s_Sentencia = s_Sentencia & "FROM covtadoc vta "
  s_Sentencia = s_Sentencia & "LEFT JOIN covtadoccta det ON vta.codemp=det.codemp AND vta.pdoano=det.pdoano AND vta.codtdc=det.codtdc AND vta.serdoc=det.serdoc AND vta.nrodoc=det.nrodoc "
  s_Sentencia = s_Sentencia & "INNER JOIN tgaux aux ON vta.codemp=aux.codemp AND vta.codaux=aux.codaux "
  s_Sentencia = s_Sentencia & "INNER JOIN tgtdc doc ON vta.codemp=doc.codemp AND vta.codtdc=doc.codtdc "
  s_Sentencia = s_Sentencia & "WHERE vta.codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND vta.pdoano='" & gsAnoAct & "' "
  s_Sentencia = s_Sentencia & "AND vta.mespvs='" & gsMesAct & "' "
  s_Sentencia = s_Sentencia & "AND vta.codtdc='" & txtDato(0).Text & "' "
  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(vta.serdoc, '-', vta.nrodoc)", "(vta.serdoc+'-'+vta.nrodoc)") & " BETWEEN '" & txtDato(1).Text & "' AND '" & txtDato(2).Text & "' "
  s_Sentencia = s_Sentencia & "AND det.tpocnc<= 3 "
  s_Sentencia = s_Sentencia & "ORDER BY vta.serdoc, vta.nrodoc, det.tpocnc, det.orden"
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = s_Sentencia
    .Open
  End With
  If porstMRp.RecordCount = 0 Then MsgBox Choose(gsIdioma, "No existe detalle del documento", "Not exist. records"), vbCritical: Exit Sub
  nRegistros = CLng(porstMRp.RecordCount)
  While Not porstMRp.EOF
    nFormato = CInt(porstMRp!forimp)
    If (nFormato <= 0 Or nFormato >= 99) Then MsgBox Choose(gsIdioma, "No existe formato de impresión", "Not exist. format print"), vbCritical: GoTo Finaliza
    sDocumento = porstMRp!documento
    nImporteTotal = CDec(porstMRp!imptotal)
    sImporteLetras = "SON : " & gfNumLet(nImporteTotal, porstMRp!tpomon)
    nRegistro = 0: nContador = 0
    ' Genero el detalle del documento
    Do
      nDiferencia = ppNumeroLinea(IIf(IsNull(porstMRp!glodet0), "", porstMRp!glodet0) & IIf(IsNull(porstMRp!glodet1), "", porstMRp!glodet1))
      nContador = nContador + nDiferencia
      nRegistro = nRegistro + 1
      s_Sentencia = "INSERT INTO " & ps_Prefijo & "trptdocventa "
      s_Sentencia = s_Sentencia & "(documento, secuencia, codtdc, serdoc, nrodoc, emision, modifica, codaux, razaux, "
      s_Sentencia = s_Sentencia & "diraux, rucaux, tpomon, signomon, pctigv, refdoc, glodet0, glodet1, glortc, "
      s_Sentencia = s_Sentencia & "impbase, impigv, imptotal, dettdc , forimp, importeletra) "
      s_Sentencia = s_Sentencia & "VALUES ('" & sDocumento & "', "
      s_Sentencia = s_Sentencia & nRegistro & ", "
      s_Sentencia = s_Sentencia & "'" & porstMRp!codtdc & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!serdoc & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!nrodoc & "', "
      If ps_Plataforma = pSrvMySql Then
        s_Sentencia = s_Sentencia & "DATE_FORMAT('" & Format(porstMRp!emision, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
        s_Sentencia = s_Sentencia & "DATE_FORMAT('" & Format(porstMRp!modifica, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
      Else
        s_Sentencia = s_Sentencia & "CONVERT(smalldatetime, '" & Format(porstMRp!emision, "yyyy-mm-dd") & "', 120), "
        s_Sentencia = s_Sentencia & "CONVERT(smalldatetime, '" & Format(porstMRp!modifica, "yyyy-mm-dd") & "', 120), "
      End If
      s_Sentencia = s_Sentencia & "'" & porstMRp!codaux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!razAux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!DirAux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!rucaux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!tpomon & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!signomon & "', "
      s_Sentencia = s_Sentencia & CDec(porstMRp!PctIGV) & ", "
      s_Sentencia = s_Sentencia & "'" & porstMRp!RefDoc & "', "
      s_Sentencia = s_Sentencia & IIf(IsNull(porstMRp!glodet0), "Null", "'" & porstMRp!glodet0 & "'") & ", "
      s_Sentencia = s_Sentencia & IIf(IsNull(porstMRp!glodet1), "Null", "'" & porstMRp!glodet1 & "'") & ", "
      s_Sentencia = s_Sentencia & IIf(IsNull(porstMRp!glortc), "Null", "'" & porstMRp!glortc & "'") & ", "
      s_Sentencia = s_Sentencia & CDec(porstMRp!impbase) & ", "
      s_Sentencia = s_Sentencia & CDec(porstMRp!impigv) & ", "
      s_Sentencia = s_Sentencia & CDec(porstMRp!imptotal) & ", "
      s_Sentencia = s_Sentencia & "'" & porstMRp!dettdc & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!forimp & "', "
      s_Sentencia = s_Sentencia & "'" & sImporteLetras & "')"
      pocnnMain.Execute s_Sentencia
      porstMRp.MoveNext
      If porstMRp.EOF Then Exit Do
    Loop While sDocumento = porstMRp("documento")
    porstMRp.MovePrevious
    ' Inserto los detalles adicionales
    nRegistro = nContador + 1
    For nContador = nRegistro To 7
      s_Sentencia = "INSERT INTO " & ps_Prefijo & "trptdocventa "
      s_Sentencia = s_Sentencia & "(documento, secuencia, codtdc, serdoc, nrodoc, emision, modifica, codaux, razaux, "
      s_Sentencia = s_Sentencia & "diraux, rucaux, tpomon, signomon, pctigv, refdoc, glodet0, glodet1, glortc, "
      s_Sentencia = s_Sentencia & "impbase, impigv, imptotal, dettdc , forimp, importeletra) "
      s_Sentencia = s_Sentencia & "VALUES ('" & sDocumento & "', "
      s_Sentencia = s_Sentencia & nContador & ", "
      s_Sentencia = s_Sentencia & "'" & porstMRp!codtdc & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!serdoc & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!nrodoc & "', "
      If ps_Plataforma = pSrvMySql Then
        s_Sentencia = s_Sentencia & "DATE_FORMAT('" & Format(porstMRp!emision, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
        s_Sentencia = s_Sentencia & "DATE_FORMAT('" & Format(porstMRp!modifica, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
      Else
        s_Sentencia = s_Sentencia & "CONVERT(smalldatetime, '" & Format(porstMRp!emision, "yyyy-mm-dd") & "', 120), "
        s_Sentencia = s_Sentencia & "CONVERT(smalldatetime, '" & Format(porstMRp!modifica, "yyyy-mm-dd") & "', 120), "
      End If
      s_Sentencia = s_Sentencia & "'" & porstMRp!codaux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!razAux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!DirAux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!rucaux & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!tpomon & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!signomon & "', "
      s_Sentencia = s_Sentencia & CDec(porstMRp!PctIGV) & ", "
      s_Sentencia = s_Sentencia & "'" & porstMRp!RefDoc & "', "
      s_Sentencia = s_Sentencia & "Null" & ", "
      s_Sentencia = s_Sentencia & "Null" & ", "
      s_Sentencia = s_Sentencia & IIf(IsNull(porstMRp!glortc), "Null", "'" & porstMRp!glortc & "'") & ", "
      s_Sentencia = s_Sentencia & "0" & ", "
      s_Sentencia = s_Sentencia & CDec(porstMRp!impigv) & ", "
      s_Sentencia = s_Sentencia & CDec(porstMRp!imptotal) & ", "
      s_Sentencia = s_Sentencia & "'" & porstMRp!dettdc & "', "
      s_Sentencia = s_Sentencia & "'" & porstMRp!forimp & "', "
      s_Sentencia = s_Sentencia & "'" & sImporteLetras & "')"
      pocnnMain.Execute s_Sentencia
    Next nContador
    porstMRp.MoveNext
  Wend
  
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT * FROM " & ps_Prefijo & "trptdocventa ORDER BY documento, secuencia"
    .Open
  End With
  ' Realizo la impresion
  gpEncabezadoRpt frmMain.rptMain, Me.Caption, udFecha, True, False, porstMRp
  With frmMain.rptMain
    '[Datos y parámetros del reporte
    .ReportFileName = gsRutRpt & "rptdocvcenta" & nFormato & ".rpt"
    .WindowState = crptMaximized
    .MarginLeft = unMargenIzquierdo
    .Destination = crptToWindow
    .Action = 1
  End With
Finaliza:
  ' Elimino la tabla temporal de impresion
  s_Sentencia = IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptdocventa", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptdocventa') DROP TABLE #trptdocventa")
  pocnnMain.Execute s_Sentencia
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
   Case 0                           'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
      With porstRango
        If .State = adStateOpen Then .Close
        sWhereSql = "AND CodTDc='" & txtDato(Index).Text & "' "
        .Source = sSelectSql & sWhereSql & sOrderSql
        .Open
        'Límites de rangos.
        If Not (.EOF And .BOF) Then
          .MoveLast
          txtDato(2).Text = .Fields(txtDato(2).DataField)
          .MoveFirst
          txtDato(1).Text = .Fields(txtDato(1).DataField)
        End If
      End With
      If txtDato(1).Text <> "" Then ppAyuDet 1
      If txtDato(2).Text <> "" Then ppAyuDet 2
   Case 1, 2                           'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2                          'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                          'Cambiar (añadir índices).
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1, 2                          'Cambiar (añadir índices).
      modAyuBus.Doc_Vta "CodTDc='" & txtDato(0).Text & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
      With porstFiltro
         .MoveFirst
         .Find "CodTDc='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & porstFiltro!dettdc
         End If
      End With
   Case 1, 2
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstRango
         .MoveFirst
         .Find "cDocumento='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & porstRango!GloDoc
         End If
      End With
   End Select

End Function

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

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
