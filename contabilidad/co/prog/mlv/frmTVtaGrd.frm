VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTVtaGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9270
      _ExtentX        =   16351
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
      ScaleWidth      =   9270
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   9270
      Begin VB.CommandButton cmdGenera 
         Caption         =   "&Generar"
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
         Left            =   4110
         Picture         =   "frmTVtaGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   700
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
         Height          =   560
         Index           =   1
         Left            =   3405
         Picture         =   "frmTVtaGrd.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   700
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
         Left            =   645
         Picture         =   "frmTVtaGrd.frx":0794
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   650
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
         Left            =   1980
         Picture         =   "frmTVtaGrd.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   700
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
         Left            =   4835
         TabIndex        =   0
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   9
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
         Left            =   8550
         Picture         =   "frmTVtaGrd.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmTVtaGrd.frx":0B2A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   650
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
         Left            =   1275
         Picture         =   "frmTVtaGrd.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   700
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
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
         Left            =   2700
         Picture         =   "frmTVtaGrd.frx":0D2E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   700
      End
   End
End
Attribute VB_Name = "frmTVtaGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2016-02-02.06 correccion ple

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
Public uorstTGTDc As ADODB.Recordset
Public uorstTGTCb As ADODB.Recordset
Public uorstCoCta As ADODB.Recordset
Public uorstCoCCo As ADODB.Recordset
Public uorstCODro As ADODB.Recordset
Public uorstCoAsiTipo As ADODB.Recordset
Public uorstCOVtaDocCta As ADODB.Recordset
Public uorstCOVtaDocCCo As ADODB.Recordset
Public uorstCOCpbCab As ADODB.Recordset
Public uorstCOCpbDet As ADODB.Recordset
Public uorstTemporal As ADODB.Recordset
Private porstCancel As ADODB.Recordset

Public uorstcodetrac As ADODB.Recordset '2015-07-08 adic tabla detrac

Public uorstCodMon As ADODB.Recordset '2016-02-02.06  correccion ple

Public usConnStrgSele_COVtaDocCta As String, _
       usConnStrgWher_COVtaDocCta As String, _
       usConnStrgOrde_COVtaDocCta As String
Public usConnStrgSele_COVtaDocCCo As String, _
       usConnStrgWher_COVtaDocCCo As String, _
       usConnStrgOrde_COVtaDocCCo As String
Public usConnStrgSele_COCpbDet As String, _
       usConnStrgWher_COCpbDet As String, _
       usConnStrgOrde_COCpbDet As String

Public ubGrabaMas As Byte
'[Repetir en frmTVta y frmTVtaMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
              
Private x_Validacion As Integer '2016-07-08 log de proceso
']
Private Sub cmdGenera_Click()
  Dim s_Sentencia As String
  
  'Verificación de Mes Cerrado.
  If gbCieVta Then
    MsgBox TEXT_9016, vbCritical
    Exit Sub
  End If
  ' Genero información
  With porstCancel
    .Source = "SELECT vta.CodDro, vta.NroCpb, vta.CodAux, vta.SerDoc, vta.NroDoc, "
    .Source = .Source & "vta.FeEDoc, vta.TpoMon, "
    .Source = .Source & "vta.CodTDc, vta.FehOpe, vta.FeVDoc, "
    .Source = .Source & "vta.ImpTCb, vta.PctIGV, vta.PctISC, "
    .Source = .Source & "vta.RefDoc, vta.GloDoc, vta.GloDocx, "
    .Source = .Source & "vta.MesPvs, vta.codasi, "
    .Source = .Source & "vta.ImpOGr_MN, vta.ImpExp_MN, vta.ImpExo_MN, "
    .Source = .Source & "vta.ImpIGV_MN, vta.ImpISC_MN, vta.ImpOIm_MN, vta.ImpTot_MN, "
    .Source = .Source & "vta.ImpOGr_ME, vta.ImpExp_ME, vta.ImpExo_ME, "
    .Source = .Source & "vta.ImpIGV_ME, vta.ImpISC_ME, vta.ImpOIm_ME, vta.ImpTot_ME, "
    .Source = .Source & "vta.IndCta_OGr, vta.IndCta_Exp, vta.IndCta_Exo, "
    .Source = .Source & "vta.IndCta_IGV, vta.IndCta_ISC, vta.IndCta_OIm, vta.IndCta_Tot, "
    .Source = .Source & "vta.IndPreGen, vta.IndGen, vta.IndAnu, "
    .Source = .Source & "vta.codemp, vta.pdoano, vta.codcon, "
    .Source = .Source & "vta.codmon " '2016-02-02.08  correccion ple
    .Source = .Source & "FROM CoVtaDoc vta "
    .Source = .Source & "LEFT JOIN CoCpbCab cab ON vta.codemp=cab.codemp AND vta.pdoano=cab.pdoano AND vta.mespvs=cab.MesPvs AND vta.CodDro=cab.CodDro AND vta.NroCpb=cab.NroCpb "
    .Source = .Source & "WHERE vta.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND vta.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND vta.MesPvs='" & gsMesAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(vta.IndGen", "ISNULL(vta.IndGen") & ", '0')='0' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(cab.CodDro, cab.NroCpb)", "ISNULL((cab.CodDro+cab.NroCpb)") & ", '')='' "
    .Source = .Source & "ORDER BY vta.CodDro, vta.NroDoc"
    .Open
  End With
 'ini 2016-07-08 log de proceso
    x_Validacion = 0
    Dim sSentencia As String
    sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#trptRPTraInf_') DROP TABLE #trptRPTraInf"
    uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRPTraInf", sSentencia)
    
    sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptRPTraInf (", "CREATE TABLE #trptRPTraInf (")
    sSentencia = sSentencia & "opcion char(1) Null, desopcion varchar(40) Null, caso char(2) Null, "
    sSentencia = sSentencia & "descripcion varchar(80) Null, registro varchar(6) DEFAULT '0')"
    uocnnMain.Execute sSentencia
 'fin 2016-07-08 log de proceso
  'Valido las Cuentas esten Correctas(llenas para todas los valores)
  If porstCancel.RecordCount > 0 Then
 Dim x As Integer
    While Not porstCancel.EOF
 'ini 2016-07-08 exporte excel
        If porstCancel!NroCpb = "000016" Then
            x = 0
        End If
 'fin 2016-07-08 exporte excel
      If VerificaCtaCCo(porstCancel) Then
        ' Genero el comprobante de diario
        ppGeneraCpbCab porstCancel
      End If
      porstCancel.MoveNext
    Wend
    ' Actualizo la grilla
    uorstMain.Requery
    uorstMain_Grd.Requery
    upDatosGrid
  End If
  porstCancel.Close
  
 'ini 2016-07-08 log de proceso
    If x_Validacion <> 0 Then
    'If x_Validacion <> 1 Then
      ' Obtengo los registros del reporte
        Dim porstMRp As ADODB.Recordset
        Set porstMRp = New ADODB.Recordset
     
      With porstMRp
        If .State = adStateOpen Then .Close
        .ActiveConnection = uocnnMain
        '.CursorLocation = adUseClient   'Es el Default.
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT * "
        .Source = .Source & "FROM " & ps_Prefijo & "trptRPTraInf "
        .Open
      End With
      ' Listado de Errores
      gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "Errores o Alertas de la Validación de Información (Ventas)", "Erros or Alerts of the Validation of Information (Sales)"), Date, True, False, porstMRp
      With frmMain.rptMain
        .ReportFileName = gsRutRpt & "rptLInfVal.rpt"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
      End With
      porstMRp.Close
    End If
    uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRPTraInf", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptRPTraInf') DROP TABLE #trptRPTraInf")
 'fin 2016-07-08 log de proceso

End Sub
Private Sub cmdImprimir_Click(Index As Integer)
 '[Datos del formulario de impresión.  'Cambiar.
   frmLVta.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLVta.Show vbModal
 ']
End Sub

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
  psConnStrgSele_Grd = "SELECT COVtaDoc.CodDro, COVtaDoc.NroCpb, c.AbvTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc, COVtaDoc.CodAux, b.RazAux, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "COVtaDoc.FeEDoc, COVtaDoc.TpoMon, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE COVtaDoc.TpoMon WHEN '" & TPOMON_NAC & "' THEN COVtaDoc.ImpTot_MN ELSE COVtaDoc.ImpTot_ME END) as cImpTot, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE COVtaDoc.IndGen WHEN -1 THEN 'x' ELSE ' ' END) as cIndGen, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "b.CodAux, c.CodTDc, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COVtaDoc.CodTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc)", "(COVtaDoc.CodTDc+COVtaDoc.SerDoc+COVtaDoc.NroDoc)") & " AS cLlave "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "FROM (COVtaDoc "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGAux b ON COVtaDoc.codemp = b.codemp AND COVtaDoc.CodAux = b.CodAux) "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGTDc c ON COVtaDoc.codemp = c.codemp AND COVtaDoc.CodTDc = c.CodTDc "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "WHERE COVtaDoc.codemp='" & gsCodEmp & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND COVtaDoc.pdoano='" & gsAnoAct & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND COVtaDoc.MesPvs='" & gsMesAct & "' "
  
  psConnStrgSele = "SELECT COVtaDoc.CodDro, COVtaDoc.NroCpb, COVtaDoc.SerDoc, COVtaDoc.NroDoc, COVtaDoc.CodAux, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.FeEDoc, COVtaDoc.TpoMon, "
  psConnStrgSele = psConnStrgSele & "(CASE COVtaDoc.TpoMon WHEN '" & TPOMON_NAC & "' THEN COVtaDoc.ImpTot_MN ELSE COVtaDoc.ImpTot_ME END) as cImpTot, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.CodTDc, COVtaDoc.FehOpe, COVtaDoc.SerDoc_Fin, COVtaDoc.NroDoc_Fin, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.FeVDoc, COVtaDoc.ImpTCb, COVtaDoc.PctIGV, COVtaDoc.PctISC, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.RefDoc, COVtaDoc.GloDoc, COVtaDoc.GloDocx, COVtaDoc.TpoGlo_Rtc, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.GloDoc_Rtc, COVtaDoc.MesPvs, COVtaDoc.codasi, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.ImpOGr_MN, COVtaDoc.ImpExp_MN, COVtaDoc.ImpExo_MN, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.ImpIGV_MN, COVtaDoc.ImpISC_MN, COVtaDoc.ImpOIm_MN, COVtaDoc.ImpTot_MN, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.ImpOGr_ME, COVtaDoc.ImpExp_ME, COVtaDoc.ImpExo_ME, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.ImpIGV_ME, COVtaDoc.ImpISC_ME, COVtaDoc.ImpOIm_ME, COVtaDoc.ImpTot_ME, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.IndCta_OGr, COVtaDoc.IndCta_Exp, COVtaDoc.IndCta_Exo, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.IndCta_IGV, COVtaDoc.IndCta_ISC, COVtaDoc.IndCta_OIm, COVtaDoc.IndCta_Tot, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.IndPreGen, COVtaDoc.IndGen, COVtaDoc.IndAnu, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.tpoimpuesto, COVtaDoc.categoriadoc, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.indvtaext, COVtaDoc.codaduana, COVtaDoc.annodua, COVtaDoc.nrodua, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.feembarq, COVtaDoc.feregula, COVtaDoc.impfob_mn, COVtaDoc.impfob_me, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.indpercep, COVtaDoc.tsapercep, COVtaDoc.serpercep, COVtaDoc.nropercep, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.codtdc_ref, COVtaDoc.serdoc_ref, COVtaDoc.nrodoc_ref, COVtaDoc.feedoc_ref, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.impbasref_mn, COVtaDoc.impigvref_mn, COVtaDoc.impbasref_me, COVtaDoc.impigvref_me, "
'ini 2015-07-07 detrac vtas
  psConnStrgSele = psConnStrgSele & "COVtaDoc.indcdt, COVtaDoc.fehcdt, COVtaDoc.nrocdt,"
  psConnStrgSele = psConnStrgSele & "COVtaDoc.tsadetrac, COVtaDoc.pctdetrac," '2015-07-08 adic tabla detrac
'fin 2015-07-07 detrac vtas
  psConnStrgSele = psConnStrgSele & "COVtaDoc.codmon, " '2016-02-02.06  correccion ple
  psConnStrgSele = psConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COVtaDoc.CodTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc)", "(COVtaDoc.CodTDc+COVtaDoc.SerDoc+COVtaDoc.NroDoc)") & " AS cLlave, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.UsrCre, COVtaDoc.FyHCre, COVtaDoc.UsrMdf, COVtaDoc.FyHMdf, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.codemp, COVtaDoc.pdoano, "
  psConnStrgSele = psConnStrgSele & "COVtaDoc.codcon "
  psConnStrgSele = psConnStrgSele & "FROM COVtaDoc "
  psConnStrgSele = psConnStrgSele & "WHERE COVtaDoc.codemp='" & gsCodEmp & "' "
  psConnStrgSele = psConnStrgSele & "AND COVtaDoc.pdoano='" & gsAnoAct & "' "
  psConnStrgOrde = "ORDER BY COVtaDoc.CodTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc"
  
  usConnStrgSele_COVtaDocCta = "SELECT COVtaDocCta.CodCta, COVtaDocCta.ImpCta_MN, COVtaDocCta.ImpCta_ME, "
  If gsIdioma = INDCCO_ACT Then
    usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(COVtaDocCta.GloDet0, ''), IFNULL(COVtaDocCta.GloDet1, ''))", "(ISNULL(COVtaDocCta.GloDet0, '')+ISNULL(COVtaDocCta.GloDet1, ''))") & " AS GloDet, "
  Else
    usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(COVtaDocCta.GloDet0x, ''), IFNULL(COVtaDocCta.GloDet1x, ''))", "(ISNULL(COVtaDocCta.GloDet0x, '')+ISNULL(COVtaDocCta.GloDet1x, ''))") & " AS GloDetx, "
  End If
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & "COVtaDocCta.CodRuc, COVtaDocCta.GloDet0, COVtaDocCta.GloDet1, COVtaDocCta.GloDet0x, COVtaDocCta.GloDet1x, "
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & "COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc, "
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & "COVtaDocCta.TpoCnc, COVtaDocCta.Orden, "
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc, COVtaDocCta.TpoCnc, COVtaDocCta.Orden)", "(COVtaDocCta.CodTDc+COVtaDocCta.SerDoc+COVtaDocCta.NroDoc+RTrim(COVtaDocCta.TpoCnc)+COVtaDocCta.Orden)") & " AS cLlave, "
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc, COVtaDocCta.TpoCnc, COVtaDocCta.Orden, COVtaDocCta.CodCta)", "(COVtaDocCta.CodTDc+COVtaDocCta.SerDoc+COVtaDocCta.NroDoc+RTrim(COVtaDocCta.TpoCnc)+COVtaDocCta.Orden+COVtaDocCta.CodCta)") & " AS cLlave2, "
  If gsIdioma = INDCCO_ACT Then
    usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(COVtaDocCta.GloDet0x, ''), IFNULL(COVtaDocCta.GloDet1x, ''))", "(ISNULL(COVtaDocCta.GloDet0x, '')+ISNULL(COVtaDocCta.GloDet1x, ''))") & " AS GloDetx, "
  Else
    usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(COVtaDocCta.GloDet0, ''), IFNULL(COVtaDocCta.GloDet1, ''))", "(ISNULL(COVtaDocCta.GloDet0, '')+ISNULL(COVtaDocCta.GloDet1, ''))") & " AS GloDet, "
  End If
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & "COVtaDocCta.UsrCre, COVtaDocCta.FyHCre, COVtaDocCta.UsrMdf, COVtaDocCta.FyHMdf, "
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & "COVtaDocCta.codemp, COVtaDocCta.pdoano "
  usConnStrgSele_COVtaDocCta = usConnStrgSele_COVtaDocCta & "FROM COVtaDocCta "
  usConnStrgWher_COVtaDocCta = ""
  usConnStrgOrde_COVtaDocCta = "ORDER BY 11, 12, 1" ' DESC"

  usConnStrgSele_COVtaDocCCo = "SELECT COVtaDocCCo.CodCCo, COVtaDocCCo.ImpCCo_MN, COVtaDocCCo.ImpCCo_ME, "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & "COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta, COVtaDocCCo.Orden, "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & "COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc, "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COVtaDocCCo.TpoCnc, COVtaDocCCo.Orden, COVtaDocCCo.CodCta)", "(RTrim(COVtaDocCCo.TpoCnc)+COVtaDocCCo.Orden+COVtaDocCCo.CodCta)") & " AS cLlave, "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc, COVtaDocCCo.TpoCnc, COVtaDocCCo.Orden, COVtaDocCCo.CodCta)", "(COVtaDocCCo.CodTDc+COVtaDocCCo.SerDoc+COVtaDocCCo.NroDoc+RTrim(COVtaDocCCo.TpoCnc)+COVtaDocCCo.Orden+COVtaDocCCo.CodCta)") & " AS cLlave1, "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc, COVtaDocCCo.TpoCnc, COVtaDocCCo.Orden, COVtaDocCCo.CodCta, COVtaDocCCo.CodCCo)", "(COVtaDocCCo.CodTDc+COVtaDocCCo.SerDoc+COVtaDocCCo.NroDoc+RTrim(COVtaDocCCo.TpoCnc)+COVtaDocCCo.Orden+COVtaDocCCo.CodCta+COVtaDocCCo.CodCCo)") & " AS cLlave2, "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & "COVtaDocCCo.UsrCre, COVtaDocCCo.FyHCre, COVtaDocCCo.UsrMdf, COVtaDocCCo.FyHMdf, "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & "COVtaDocCCo.codemp, COVtaDocCCo.pdoano "
  usConnStrgSele_COVtaDocCCo = usConnStrgSele_COVtaDocCCo & "FROM COVtaDocCCo "
  usConnStrgWher_COVtaDocCCo = ""
  usConnStrgOrde_COVtaDocCCo = "ORDER BY 4, 6, 5, 1"
    
  usConnStrgSele_COCpbDet = "SELECT COCpbDet.CodCta, COCpbDet.CodAux, COCpbDet.CodCCo, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & Choose(gsIdioma, "COCpbDet.GloIte, ", "COCpbDet.GloItex, ")
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN COCpbDet.ImpMN ELSE 0 END) AS cImpMN_Deb, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN 0 ELSE COCpbDet.ImpMN END) AS cImpMN_Hab, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN COCpbDet.ImpME ELSE 0 END) AS cImpME_Deb, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN 0 ELSE COCpbDet.ImpME END) AS cImpME_Hab, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "(CASE COCpbDet.TpoGnr WHEN " & TPOGNR_DST & " THEN '*' ELSE '' END) AS cTpoGnr, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.MesPvs, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte, COCpbDet.FehOpe, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.CodTDc, COCpbDet.SerDoc, COCpbDet.NroDoc, COCpbDet.FeEDoc, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.FeVDoc, COCpbDet.FeRDoc, COCpbDet.RefDoc, COCpbDet.TpoMon, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.ImpTCb, COCpbDet.ImpMN, COCpbDet.ImpME, COCpbDet.TpoCtb, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.TpoGnr, COCpbDet.tpopvs, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & IIf(ps_Plataforma = pSrvMySql, "Concat(COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte)", "(COCpbDet.CodDro+COCpbDet.NroCpb+COCpbDet.NroIte)") & " AS cLlave, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & Choose(gsIdioma, " COCpbDet.GloItex, ", " COCpbDet.GloIte, ")
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.UsrCre, COCpbDet.FyHCre, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.codemp, COCpbDet.pdoano, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.codcon, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.codmon " '2016-02-02.08  correccion ple
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "FROM COCpbDet "
  usConnStrgWher_COCpbDet = "WHERE COCpbDet.codemp='" & gsCodEmp & "' COCpbDet.pdoano='" & gsAnoAct & "' "
  usConnStrgWher_COCpbDet = usConnStrgWher_COCpbDet & "AND COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='' AND COCpbDet.NroCpb='' "
  usConnStrgOrde_COCpbDet = "ORDER BY COCpbDet.NroIte"
  
  Set uocnnMain = New ADODB.Connection
  Set uocnnNoGrabable = New ADODB.Connection
  Set uorstMain = New ADODB.Recordset
  Set uorstMain_Grd = New ADODB.Recordset
  Set uorstTGAux = New ADODB.Recordset
  Set uorstTGTDc = New ADODB.Recordset
  Set uorstTGTCb = New ADODB.Recordset
  Set uorstCoCta = New ADODB.Recordset
  Set uorstCoCCo = New ADODB.Recordset
  Set uorstCODro = New ADODB.Recordset
  Set uorstCoAsiTipo = New ADODB.Recordset
  Set uorstCOVtaDocCta = New ADODB.Recordset
  Set uorstCOVtaDocCCo = New ADODB.Recordset
  Set uorstCOCpbCab = New ADODB.Recordset
  Set uorstCOCpbDet = New ADODB.Recordset
  Set porstCancel = New ADODB.Recordset
    
  Set uorstcodetrac = New ADODB.Recordset '2015-07-08 adic tabla detrac
  
  Set uorstCodMon = New ADODB.Recordset '2016-02-02.06  correccion ple

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
    .Properties("Unique Table").Value = "COVtaDoc"
  End With
  With uorstMain
    .ActiveConnection = uocnnMain
    .Source = psConnStrgSele & psConnStrgOrde
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic 'adLockReadOnly
    .Open
    .Properties("Unique Table").Value = "COVtaDoc"
  End With
  With uorstTGTDc
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.CodTDc, " & Choose(gsIdioma, "a.DetTDc", "a.DetTDcx") & " AS DetTDc, a.SgnTDc "
    .Source = .Source & "FROM TGTDc a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
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
  With uorstCoCta
    .ActiveConnection = frmTVtaGrd.uocnnMain
    .Source = "SELECT a.CodCta, " & Choose(gsIdioma, "a.DetCta", "a.DetCtax") & " AS DetCta, a.TpoTCb, a.IndDoc, a.IndCCo, a.codcco_def "
    .Source = .Source & ",tpomon " '2015-06-30 correccion tipo mon cta
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
     .Source = .Source & "AND a.EstCCo='" & ESTCCO_ACT & "' "
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodCCo)>2"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  With uorstCODro
     .ActiveConnection = uocnnMain
     .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro, Cpb" & gsMesAct & " "
     .Source = .Source & "FROM CODro "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4 "
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  With uorstCoAsiTipo
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.CodAsi, " & Choose(gsIdioma, "a.DetAsi", "a.DetAsix") & " AS DetAsi, a.TpoAsi "
    .Source = .Source & "FROM CoAsiTipo a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.TpoAsi='" & TPOGNR_VTA & "'"
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  With uorstCOVtaDocCta
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
  End With
  With uorstCOVtaDocCCo
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
  End With
  With uorstCOCpbCab
    .ActiveConnection = uocnnMain
    .Source = "SELECT CodDro, NroCpb, FehCpb, GloCpb, GloCpbx, TpoGnr, IndNCu, MesPvs, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "Concat(CodDro, NroCpb)", "(CodDro+NroCpb)") & " AS cLlave, "
    .Source = .Source & "codemp, pdoano, UsrCre, FyHCre "
    .Source = .Source & "FROM COCpbCab "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND MesPvs='" & gsMesAct & "'"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
'2016-02-02.08  correccion ple
'aqui va con el selec de la venta,
'pero aqui recien retoma valores del cocpbdet
'en el proceso ppDatosWhere de frmTVta
  With uorstCOCpbDet
    .ActiveConnection = uocnnMain
    .Source = psConnStrgSele & psConnStrgOrde
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockBatchOptimistic ' adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "COCpbDet"
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
    .Source = .Source & "AND a.IndCli=1 "
    .Source = .Source & "AND a.EstAux='" & ESTAUX_ACT & "'"
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
']
   
'ini 2015-07-08 adic tabla detrac
  With uorstcodetrac
     .ActiveConnection = uocnnMain
     .Source = "SELECT coddetrac, " & Choose(gsIdioma, "detdetrac", "detdetracx") & " AS DetDetrac,pctdetrac ,  "
     .Source = .Source & "codemp "
     .Source = .Source & "FROM codetrac  "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND estdetrac ='" & ESTDETRAC_ACT & "' "
     '.Source = .Source & "AND pdoano='" & gsAnoAct & "' "
     '.Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
'fin 2015-07-08 adic tabla detrac
   
'ini 2016-02-02.06  correccion ple
  With uorstCodMon
     '.ActiveConnection = uocnnMain
     .ActiveConnection = CONNSTRG & gsNomBDC
     '.Source = fSqlTabla("004")
     '.Source = fSqlTabla(CODSUNAT_004) '2016-02-02.06  correccion ple
     .Source = gf_tb_sunat(CODSUNAT_004)
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
'fin 2016-02-02.06  correccion ple
  
  dgrMain.MarqueeStyle = dbgHighlightRow
  Set dgrMain.DataSource = uorstMain_Grd
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
  cmdVerificar.Caption = Choose(gsIdioma, "&Verificar", "&Check")
  cmdGenera.Caption = Choose(gsIdioma, "&Generar", "&Generate")
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
   uorstTGTDc.Close
   uorstTGTCb.Close
   uorstCoCta.Close
   uorstCoCCo.Close
   uorstCODro.Close
   uorstCoAsiTipo.Close
   
   uorstcodetrac.Close '2015-07-08 adic tabla detrac

   uorstCodMon.Close '2016-02-02.06  correccion ple

   
'[ARREGLAR. Genera demora al salir de la opción.
   If uorstCOVtaDocCta.State = adStateOpen Then uorstCOVtaDocCta.Close
   If uorstCOVtaDocCCo.State = adStateOpen Then uorstCOVtaDocCCo.Close
']ARREGLAR.
   uorstCOCpbCab.Close
   uorstCOCpbDet.Close
   uorstMain_Grd.Close
   uorstMain.Close
   uocnnMain.Close
   Set porstCancel = Nothing
   Set uorstTemporal = Nothing
   Set uorstTGAux = Nothing
   Set uorstTGTDc = Nothing
   Set uorstTGTCb = Nothing
   Set uorstCoCta = Nothing
   Set uorstCoCCo = Nothing
   Set uorstCODro = Nothing
   Set uorstCoAsiTipo = Nothing
   Set uorstCOVtaDocCta = Nothing
   Set uorstCOVtaDocCCo = Nothing
   Set uorstCOCpbCab = Nothing
   Set uorstCOCpbDet = Nothing
   Set uorstMain_Grd = Nothing
   Set uorstMain = Nothing
   
   Set uorstcodetrac = Nothing '2015-07-08 adic tabla detrac
   
   Set uorstCodMon = Nothing '2016-02-02.06  correccion ple
   
   Set uocnnMain = Nothing
End Sub

Private Sub cmdNuevo_Click()
 '[Propio del formulario.
   'Verificación de Mes Cerrado.
   If gbCieVta Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
 
   ubGrabaMas = INDMASCTA_INI
   uocnnMain.BeginTrans
 ']
   gpTUg_Nuevo Me, frmTVta             'Cambiar Formulario de Datos.
'///Angel 12/12/2003
'/// Agregado para eliminar el registro creado como cabecera al intentar registrar un dato de cuenta y luego cancelar el ingreso completo.
   cmdRefrescar_Click
'///
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err

   'Verificación de existencia de ítemes.
   If uorstMain_Grd.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If

 '[Propio del formulario.
   ubGrabaMas = INDMASCTA_CTA
 ']

 '[Búsqueda del ítem.
  uorstMain.Requery
  uorstMain.MoveFirst
  uorstMain.Find "cLlave='" & uorstMain_Grd!codtdc & uorstMain_Grd!SerDoc & uorstMain_Grd!NroDoc & "'"
 ']

   With frmTVta                        'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
      .txtLlave(0).Enabled = False
      .txtLlave(1).Enabled = False
      .txtLlave(2).Enabled = False
      .cmdLlaveAyud(0).Enabled = False
      .lblLlaveDeta(0).Enabled = False
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
   On Error GoTo Err
   
   Dim dsLlaveSiguiente As String
   
   'Verificación de Mes Cerrado.
   If gbCieVta Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   'Verificación de existencia de ítemes.
   If uorstMain_Grd.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If
'ini 2016-05-27/28 nivel=asisten no elimin datos
   If gsNvlUsr = NVLUSR_ASIS Then
      MsgBox TEXT_9026, vbCritical
      Exit Sub
   End If
'fin 2016-05-27/28 nivel=asisten no elimin datos
   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & "-" & Trim(dgrMain.Columns(2)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      With porstCancel
        .Source = "SELECT MesPvs, CodAux, CodTDc, SerDoc, NroDoc, TpoPvs "
        .Source = .Source & "FROM COCpbDet "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND MesPvs='" & gsMesAct & "' AND CodAux='" & uorstMain_Grd!codaux & "' "
        .Source = .Source & "AND CodTDc='" & uorstMain_Grd!codtdc & "' AND SerDoc='" & uorstMain_Grd!SerDoc & "'"
        .Source = .Source & "AND NroDoc='" & uorstMain_Grd!NroDoc & "' AND TpoPvs<>'" & TPOPVS_CAN & "'"
         .Open
         If porstCancel.RecordCount = 0 Then
            uorstMain.MoveFirst
            uorstMain.Find "cLlave = '" & uorstMain_Grd!codtdc & uorstMain_Grd!SerDoc & uorstMain_Grd!NroDoc & "'"

            uocnnMain.BeginTrans       'INICIA TRANSACCION.
            uocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND MesPvs='" & gsMesAct & "' AND CodDro='" & Trim(dgrMain.Columns(0)) & "' And NroCpb='" & Trim(dgrMain.Columns(1)) & "' And TpoGnr='" & TPOGNR_VTA & "'"
            uorstMain.Properties("Unique Table").Value = "COVtaDoc"
            uorstMain.Delete
            uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.

           'Busca siguiente ítem.
            With uorstMain_Grd
               .MoveNext
               If .EOF Then .MoveLast
               dsLlaveSiguiente = !codtdc & !SerDoc & !NroDoc
               .Requery
               If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
            End With
            'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
            fEstMayUpd
            'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
         Else
            MsgBox Choose(gsIdioma, "Debe eliminar antes las Cancelaciones.", " The Cancelations must be eliminated before."), vbExclamation
         End If
      End With
      porstCancel.Close
      upDatosGrid
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

Public Sub cmdVerificar_Click()
  '[Datos del formulario de impresión.  'Cambiar.
  Dim s_Sentencia As String
  Dim porstMRp As New ADODB.Recordset
 
  s_Sentencia = "SELECT CodTDc, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc, '-',NroDoc)", "(SerDoc+'-'+NroDoc)") & " AS cDocumento, FehOpe, FeEdoc, GloDoc, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cMoneda, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpOGr_MN ELSE ImpOGr_ME END) AS cImpBas, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpExo_MN ELSE ImpExo_ME END) AS cExonerado, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpIGV_MN ELSE ImpIGV_ME END) AS cImpIGV, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpTot_MN ELSE ImpTot_ME END) AS cImpTotal, "
  s_Sentencia = s_Sentencia & "a.CodDro, a.NroCpb "
  s_Sentencia = s_Sentencia & "FROM CoVtaDoc a "
  s_Sentencia = s_Sentencia & "LEFT JOIN CoCpbCab AS b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb "
  s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND a.pdoano='" & gsAnoAct & "' "
  s_Sentencia = s_Sentencia & "AND a.MesPvs='" & gsMesAct & "' "
  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(b.CodDro, b.NroCpb)", "ISNULL((b.CodDro+b.NroCpb)") & ", '')='' "
  s_Sentencia = s_Sentencia & "ORDER BY a.CodDro, a.NroCpb, CodTDc, SerDoc, NroDoc"
  With porstMRp
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = s_Sentencia
    .Open
  End With

  gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "DOCUMENTOS DE VENTAS NO CONTABILIZADOS", "NOT COUNTED DOCUMENTS OF SALES"), Date, True, False, porstMRp
  With frmMain.rptMain
    '[Datos y parámetros del reporte.  'Cambiar.
    .ReportFileName = gsRutRpt & "rptLVtaCpb.rpt"
    .WindowShowExportBtn = True
    .MarginLeft = 240
    .WindowState = crptMaximized
    .Destination = crptToWindow
    .Action = 1
  End With
  porstMRp.Close
  Set porstMRp = Nothing
 ']
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
'[ARREGLAR. No acepta ordenar por columna de tablas secundarias en el recordset.
   If ColIndex = 2 Or ColIndex = 6 Then Exit Sub
']ARREGLAR.

   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   psConnStrgOrde = "ORDER BY "
   Select Case pnColumnaOrd            'Cambiar.
   Case 3
      psConnStrgOrde = psConnStrgOrde & "4, 2, 3"
'   Case 4
'      psConnStrgOrde = psConnStrgOrde & "5, 1, 2, 3"
   Case Else
      psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
   End Select
   With uorstMain_Grd
      .Close
      .Properties("Unique Table").Value = "COVtaDoc"
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
            .Item(dnNum).Caption = Choose(gsIdioma, "Diario", "Journal")
            .Item(dnNum).Width = 500
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "NºComp.", "NºVouch")
            .Item(dnNum).Width = 700
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "TDc", "TDc")   ' Type of document
            .Item(dnNum).Width = 500
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "Ser", "Ser")
            .Item(dnNum).Width = 500
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "Número", "Number")
            .Item(dnNum).Width = 1000
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 1100
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "Razón Social", "Firm Name")
            .Item(dnNum).Width = 1720
         Case 7
            .Item(dnNum).Caption = Choose(gsIdioma, "F.Emisión", "Issue Date")
            .Item(dnNum).Width = 1000
         Case 8
            .Item(dnNum).Caption = Choose(gsIdioma, "Mon", "Cur")
            .Item(dnNum).Width = 250
         Case 9
            .Item(dnNum).Caption = Choose(gsIdioma, "Total", "Total")
            .Item(dnNum).Width = 1200
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
            .Item(dnNum).Alignment = dbgRight
         Case 10
            .Item(dnNum).Caption = "G"
            .Item(dnNum).Width = 230
            .Item(dnNum).Alignment = dbgCenter
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'[Código propio del formulario.
'solo sirve para el boton generar de la grilla
Private Sub ppGeneraCpbCab(ByVal oRecordset As ADODB.Recordset)
'2016-07-08 ..  On Error GoTo ErrGrabar
  Dim nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nNumRegistros As Long
  Dim sSentencia As String, sComprobante As String
  Dim sCodAux As String, sTpoCtb As String, sContrato As String
  Dim nIndCco As Byte
  Dim porstCprCta As ADODB.Recordset

  Set porstCprCta = New ADODB.Recordset
  With porstCprCta
    .ActiveConnection = uocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  
  uocnnMain.BeginTrans            'INICIA TRANSACCION.
      
  sComprobante = IIf(IsNull(oRecordset!NroCpb), "", oRecordset!NroCpb)
  ' Captura del siguiente numero de comprobante
  If sComprobante = "" Then
    sComprobante = gfNumComprobante(gsAnoAct, gsMesAct, oRecordset!coddro)
    sSentencia = "UPDATE codro SET cpb" & gsMesAct & "='" & sComprobante & "' "
    sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND codDro='" & oRecordset!coddro & "'"
    uocnnMain.Execute sSentencia, nNumRegistros
  End If
  
  ' Grabación de cabecera de comprobante
  sSentencia = "INSERT INTO cocpbcab(codemp, pdoano, mespvs, coddro, nrocpb, fehcpb, glocpb, glocpbx, tpognr, indncu, indanu, usrcre, fyhcre, usrmdf, fyhmdf)"
  sSentencia = sSentencia & " VALUES("
  sSentencia = sSentencia & "'" & gsCodEmp & "', "
  sSentencia = sSentencia & "'" & gsAnoAct & "', "
  sSentencia = sSentencia & "'" & gsMesAct & "', "
  sSentencia = sSentencia & "'" & oRecordset!coddro & "', "
  sSentencia = sSentencia & "'" & sComprobante & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!fehope, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(IsNull(oRecordset!GloDoc), "Null", "'" & oRecordset!GloDoc & "'") & ", "
  sSentencia = sSentencia & IIf(IsNull(oRecordset!glodocx), "Null", "'" & oRecordset!glodocx & "'") & ", "
  sSentencia = sSentencia & "'" & TPOGNR_VTA & "', "
  sSentencia = sSentencia & "'" & INDNCU_FAL & "', "
  sSentencia = sSentencia & "'" & INDANU_FAL & "', "
  sSentencia = sSentencia & "'" & gsAbvUsr & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
  sSentencia = sSentencia & "Null, Null)"
  uocnnMain.Execute sSentencia, nNumRegistros
  
  ' Información detalle cuentas
  sContrato = IIf(IsNull(oRecordset!codcon), "", oRecordset!codcon)
  With porstCprCta
    .Source = "SELECT vta.tpocnc, vta.orden, vta.codcta, cco.codcco, vta.glodet0, vta.glodet0x, vta.impcta_mn, vta.impcta_me, cco.impcco_mn, cco.impcco_me, vta.codruc, "
    .Source = .Source & "cta.indcco, cta.inddoc, cta.inddoc, cta.tpotcb, tdc.sgntdc "
    .Source = .Source & "FROM covtadoccta vta "
    .Source = .Source & "INNER JOIN cocta cta ON vta.codemp=cta.codemp AND vta.pdoano=cta.pdoano AND vta.codcta=cta.codcta "
    .Source = .Source & "INNER JOIN tgtdc tdc ON vta.codemp=tdc.codemp AND vta.codtdc=tdc.codtdc "
    .Source = .Source & "LEFT JOIN covtadoccco cco ON vta.codemp=cco.codemp AND vta.pdoano=cco.pdoano AND vta.codtdc=cco.codtdc "
    .Source = .Source & "AND vta.serdoc=cco.serdoc AND vta.nrodoc=cco.nrodoc AND vta.tpocnc=cco.tpocnc AND vta.orden=cco.orden AND vta.codcta=cco.codcta "
    .Source = .Source & "WHERE vta.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND vta.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND vta.codtdc='" & oRecordset!codtdc & "' "
    .Source = .Source & "AND vta.serdoc='" & oRecordset!SerDoc & "' "
    .Source = .Source & "AND vta.nrodoc='" & oRecordset!NroDoc & "' "
    .Source = .Source & "ORDER BY vta.tpocnc DESC, vta.orden"
    .Open
  End With
  If porstCprCta.RecordCount > 0 Then
    nRegistro = 0
    While Not porstCprCta.EOF
      nIndCco = porstCprCta!indcco
      nImporte_mn = CDec(porstCprCta(IIf(nIndCco = INDCCO_ACT, "impcco_mn", "impcta_mn")))
      nImporte_me = CDec(porstCprCta(IIf(nIndCco = INDCCO_ACT, "impcco_me", "impcta_me")))
      sCodAux = IIf(IsNull(porstCprCta!codruc), "", porstCprCta!codruc)
      sCodAux = IIf(porstCprCta!IndDoc = INDDOC_ACT, oRecordset!codaux, IIf(sCodAux = "", oRecordset!codaux, sCodAux))
      If (nImporte_me > 0) Or (nImporte_mn > 0) Then
        sTpoCtb = IIf(porstCprCta!tpocnc = TPOCNC_TOT_VTA, IIf(porstCprCta!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(porstCprCta!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        sTpoCtb = IIf(porstCprCta!tpocnc = TPOCNC_TOT_VTA, IIf(porstCprCta!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(porstCprCta!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
      
      nRegistro = nRegistro + 1
      ' Grabación de cabecera de comprobante
      sSentencia = "INSERT INTO CoCpbDet(codemp, pdoano, coddro, nrocpb, nroite, mespvs, blqite, codtdc, fehope, codcta, codcco, codaux, serdoc, nrodoc, feedoc, fevdoc, "
'ini 2016-02-02.08  correccion ple
      'sSentencia = sSentencia & "FeRDoc, RefDoc, GloIte, GloItex, codcon, TpoCtb, TpoPvs, TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, tpognr, UsrCre, FyHCre, UsrMdf, FyHMdf) "
      sSentencia = sSentencia & "FeRDoc, RefDoc, GloIte, GloItex, codcon, TpoCtb, TpoPvs, TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, tpognr, UsrCre, FyHCre, UsrMdf, FyHMdf "
      'sSentencia = sSentencia & "VALUES("
      sSentencia = sSentencia & ",codmon "
      sSentencia = sSentencia & ") VALUES("
'fin 2016-02-02.08  correccion ple
      sSentencia = sSentencia & "'" & gsCodEmp & "', "
      sSentencia = sSentencia & "'" & gsAnoAct & "', "
      sSentencia = sSentencia & "'" & oRecordset!coddro & "', "
      sSentencia = sSentencia & "'" & sComprobante & "', "
      sSentencia = sSentencia & "'" & nRegistro & "', "
      sSentencia = sSentencia & "'" & gsMesAct & "', "
      sSentencia = sSentencia & "'" & nRegistro & "', "
      sSentencia = sSentencia & "'" & oRecordset!codtdc & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!fehope, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & "'" & porstCprCta!CodCta & "', "
      sSentencia = sSentencia & IIf(IsNull(porstCprCta!codcco), "Null", "'" & porstCprCta!codcco & "'") & ", "
      sSentencia = sSentencia & IIf(sCodAux = "", "Null", "'" & sCodAux & "'") & ", "
      sSentencia = sSentencia & "'" & oRecordset!SerDoc & "', "
      sSentencia = sSentencia & "'" & oRecordset!NroDoc & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!fevdoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(IsNull(oRecordset!RefDoc), "Null", "'" & oRecordset!RefDoc & "'") & ", "
      sSentencia = sSentencia & IIf(IsNull(porstCprCta!glodet0), "Null", "'" & Left(porstCprCta!glodet0, 60) & "'") & ", "
      sSentencia = sSentencia & IIf(IsNull(porstCprCta!glodet0x), "Null", "'" & Left(porstCprCta!glodet0x, 60) & "'") & ", "
      sSentencia = sSentencia & IIf((CInt(porstCprCta!tpocnc) >= 4 Or sContrato = ""), "Null", "'" & sContrato & "'") & ", "
      sSentencia = sSentencia & "'" & sTpoCtb & "', "
      sSentencia = sSentencia & "'" & TPOPVS_PVS & "', "
      sSentencia = sSentencia & "'" & oRecordset!tpomon & "', "
      sSentencia = sSentencia & "'" & porstCprCta!TpoTcb & "', "
      sSentencia = sSentencia & CDec(oRecordset!ImpTCb) & ", "
      sSentencia = sSentencia & Abs(nImporte_mn) & ", "
      sSentencia = sSentencia & Abs(nImporte_me) & ", "
      sSentencia = sSentencia & "'" & TPOGNR_VTA & "', "
      sSentencia = sSentencia & "'" & gsAbvUsr & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
'ini 2016-02-02.08  correccion ple
      'sSentencia = sSentencia & "Null, Null)"
      sSentencia = sSentencia & "Null, Null"
      sSentencia = sSentencia & ",'" & oRecordset!codmon & "' "
      sSentencia = sSentencia & ")"
'fin 2016-02-02.08  correccion ple
      uocnnMain.Execute sSentencia, nNumRegistros
      porstCprCta.MoveNext
    Wend
  End If
  porstCprCta.Close
  'Si no está marcado para generar, marca el documento como no generado.
  sSentencia = "UPDATE CoVtaDoc SET indpregen=" & INDPREGEN_ACT & ", indgen=-1 "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND codaux='" & oRecordset!codaux & "' "
  sSentencia = sSentencia & "AND codtdc='" & oRecordset!codtdc & "' "
  sSentencia = sSentencia & "AND serdoc='" & oRecordset!SerDoc & "' "
  sSentencia = sSentencia & "AND nrodoc='" & oRecordset!NroDoc & "'"
  uocnnMain.Execute sSentencia, nNumRegistros
  uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  
  Exit Sub
ErrGrabar:
   gpErrores
   
   uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.

End Sub
Private Function VerificaCtaCCo(ByVal oRecordset As ADODB.Recordset) As Boolean
  Dim sSentencia As String '2016-07-08 log de proceso
  Static nNumRegistros As Double '2016-07-08 log de proceso

  Dim nContador As Integer, nIndCco As Byte
  Dim sRegistro As String, sIndicado As String, sSource As String
  Dim nImporteCpr_mn As Double, nImporteCpr_me As Double
  Dim nImporteCta_mn As Double, nImporteCta_me As Double
  Dim nImporteCCo_mn As Double, nImporteCCo_me As Double
  Dim porstCprCta As ADODB.Recordset
  Dim porstCprCco As ADODB.Recordset
   
  Set porstCprCta = New ADODB.Recordset
  Set porstCprCco = New ADODB.Recordset
  With porstCprCta
    .ActiveConnection = uocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  With porstCprCco
    .ActiveConnection = uocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  
  VerificaCtaCCo = False
  For nContador = 1 To 7
    sRegistro = Choose(nContador, "impogr", "impexp", "impexo", "impigv", "impisc", "impoim", "imptot")
    sIndicado = "indcta_" & Right(sRegistro, 3)
    nImporteCpr_mn = CDec(oRecordset(sRegistro & "_mn"))
    nImporteCpr_me = CDec(oRecordset(sRegistro & "_me"))
    nImporteCta_mn = 0
    nImporteCta_me = 0
    ' Verifico los importes de las cuentas
    If oRecordset(sIndicado) <> 0 Then
      With porstCprCta
        .Source = "SELECT vta.orden, vta.codcta, vta.impcta_mn, vta.impcta_me, cta.indcco "
        .Source = .Source & "FROM covtadoccta vta "
        .Source = .Source & "INNER JOIN cocta cta ON vta.codemp=cta.codemp AND vta.pdoano=cta.pdoano AND vta.codcta=cta.codcta "
        .Source = .Source & "WHERE vta.codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND vta.pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND vta.codtdc='" & oRecordset!codtdc & "' "
        .Source = .Source & "AND vta.serdoc='" & oRecordset!SerDoc & "' "
        .Source = .Source & "AND vta.nrodoc='" & oRecordset!NroDoc & "' "
        .Source = .Source & "AND vta.tpocnc='" & nContador & "' "
        .Source = .Source & "ORDER BY orden"
        .Open
      End With
      ' Valido los centro de costos
      If porstCprCta.RecordCount > 0 Then
        nImporteCta_mn = 0
        nImporteCta_me = 0
        While Not porstCprCta.EOF
          nImporteCta_mn = nImporteCta_mn + CDec(porstCprCta!impcta_mn)
          nImporteCta_me = nImporteCta_me + CDec(porstCprCta!impcta_me)
          nIndCco = porstCprCta!indcco
          nImporteCCo_mn = 0
          nImporteCCo_me = 0
          If nIndCco = INDCCO_ACT Then
            With porstCprCco
              .Source = "SELECT vta.codcta, ROUND(SUM(vta.impcco_mn), 2) AS impcco_mn, ROUND(SUM(vta.impcco_me), 2) AS impcco_me "
              .Source = .Source & "FROM covtadoccco vta "
              .Source = .Source & "INNER JOIN cocco cco ON vta.codemp=cco.codemp AND vta.pdoano=cco.pdoano AND vta.codcco=cco.codcco "
              .Source = .Source & "WHERE vta.codemp='" & gsCodEmp & "' "
              .Source = .Source & "AND vta.pdoano='" & gsAnoAct & "' "
              .Source = .Source & "AND vta.codtdc='" & oRecordset!codtdc & "' "
              .Source = .Source & "AND vta.serdoc='" & oRecordset!SerDoc & "' "
              .Source = .Source & "AND vta.nrodoc='" & oRecordset!NroDoc & "' "
              .Source = .Source & "AND vta.tpocnc='" & nContador & "' "
              .Source = .Source & "AND vta.orden='" & porstCprCta!orden & "' "
              .Source = .Source & "AND vta.codcta='" & porstCprCta!CodCta & "' "
              .Source = .Source & "GROUP BY vta.codcta "
              .Open
            End With
            ' Valido los centro de costos
            If porstCprCco.RecordCount > 0 Then
              nImporteCCo_mn = CDec(porstCprCco!impcco_mn)
              nImporteCCo_me = CDec(porstCprCco!impcco_me)
            End If
            porstCprCco.Close
            VerificaCtaCCo = (CDec(porstCprCta!impcta_mn) = nImporteCCo_mn)
            If Not VerificaCtaCCo Then GoTo ErrorVerifica
            VerificaCtaCCo = (CDec(porstCprCta!impcta_me) = nImporteCCo_me)
            If Not VerificaCtaCCo Then GoTo ErrorVerifica
          End If
          porstCprCta.MoveNext
        Wend
      End If
      porstCprCta.Close
    End If
    
    ' Verifico información de rubro
    VerificaCtaCCo = (nImporteCpr_mn = nImporteCta_mn)
'ini 2016-07-08 log de proceso
    'si el importe de la cabecar mn <> importe de detalle cuenta
    If Not VerificaCtaCCo Then
        x_Validacion = 1
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & " VALUES ('1','Importe <>','01','Doc:" & oRecordset!SerDoc & "-" & oRecordset!NroDoc & " MN Cabeza <> MN Detalle Cta.','.'"
        sSentencia = sSentencia & ")"
        uocnnMain.Execute sSentencia, nNumRegistros
    End If
'fin 2016-07-08 log de proceso
    If Not VerificaCtaCCo Then GoTo ErrorVerifica
    VerificaCtaCCo = (nImporteCpr_me = nImporteCta_me)
 'ini 2016-07-08 log de proceso
    If Not VerificaCtaCCo Then
        x_Validacion = 1
        sSentencia = "INSERT INTO " & ps_Prefijo & "trptRPTraInf (opcion, desopcion, caso, descripcion, registro) "
        sSentencia = sSentencia & " VALUES ('1','Importe <>','01','Doc:" & oRecordset!SerDoc & "-" & oRecordset!NroDoc & " ME Cabeza <> ME Detalle Cta.','.'"
        sSentencia = sSentencia & ")"
        uocnnMain.Execute sSentencia, nNumRegistros
    End If
 'fin 2016-07-08 log de proceso
    If Not VerificaCtaCCo Then GoTo ErrorVerifica
  Next nContador
  
ErrorVerifica:
  Set porstCprCco = Nothing
  Set porstCprCta = Nothing

End Function
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
   cmdVerificar.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
   cmdImprimir(1).Enabled = (taOpciones(2) Or taOpciones(3))
   cmdGenera.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property

