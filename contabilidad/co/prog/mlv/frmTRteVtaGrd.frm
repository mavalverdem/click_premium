VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTRteVtaGrd 
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
         Picture         =   "frmTRteVtaGrd.frx":0000
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
         Picture         =   "frmTRteVtaGrd.frx":0312
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
         Picture         =   "frmTRteVtaGrd.frx":0794
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
         Picture         =   "frmTRteVtaGrd.frx":0896
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
         Picture         =   "frmTRteVtaGrd.frx":09E0
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
         Picture         =   "frmTRteVtaGrd.frx":0B2A
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
         Picture         =   "frmTRteVtaGrd.frx":0C2C
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
         Picture         =   "frmTRteVtaGrd.frx":0D2E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   700
      End
   End
End
Attribute VB_Name = "frmTRteVtaGrd"
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
Public uorstTGTDc As ADODB.Recordset
Public uorstCoCta As ADODB.Recordset
Public uorstCoCCo As ADODB.Recordset
Public uorstCODro As ADODB.Recordset
Public uorstCoRteVtaCta As ADODB.Recordset
Public uorstCoRteVtaCCo As ADODB.Recordset
Public uorstCOCpbDet As ADODB.Recordset
Public uorstTemporal As ADODB.Recordset
Private porstCancel As ADODB.Recordset
Public usConnStrgSele_CoRteVtaCta As String, _
       usConnStrgWher_CoRteVtaCta As String, _
       usConnStrgOrde_CoRteVtaCta As String
Public usConnStrgSele_CoRteVtaCCo As String, _
       usConnStrgWher_CoRteVtaCCo As String, _
       usConnStrgOrde_CoRteVtaCCo As String
Public usConnStrgSele_COCpbDet As String, _
       usConnStrgWher_COCpbDet As String, _
       usConnStrgOrde_COCpbDet As String

Public ubGrabaMas As Byte
'[Repetir en frmTrteVta y frmTrteVtaMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
Private Sub cmdGenera_Click()
  Dim s_dExpresion As String, s_TipoDocumento As String
  Dim s_SerieDocumento As String, s_NroDocumento As String
  Dim s_DiarioDocumento As String, s_NroComprobante As String
  Dim s_Sentencia As String, s_Expresion As String
  Dim n_ImporteMN As Double, n_ImporteME As Double
  Dim n_TipoCambio As Double, nNumeRegistro As Long
  
  'Verificación de Mes Cerrado.
  If gbCieVta Then MsgBox TEXT_9016, vbCritical: Exit Sub
  If MsgBox(Choose(gsIdioma, "Generar proceso documentos ", "Generating document processing") & Me.Caption & "?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then Exit Sub
  
  ' fecha de proceso
  s_dExpresion = InputBox(Choose(gsIdioma, "Ingrese Fecha de Proceso ", "Enter Date Process ") & Me.Caption, Choose(gsIdioma, "Generación Documento de Ventas", "Sales Document Generation"), "01/" & gsMesAct & "/" & gsAnoAct)
  If Not IsDate(s_dExpresion) Then Beep: MsgBox TEXT_8010, vbCritical: Exit Sub
  ' Valida fecha de proceso
  If Month(s_dExpresion) <> Val(gsMesAct) Or Year(s_dExpresion) <> Val(gsAnoAct) Then
    MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
    Exit Sub
  End If
  
  ' Obtengo el tipo de cambio de la fecha
  With porstCancel
    If .State = adStateOpen Then .Close
    .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(ImpTCb_Vta, 1) AS nImpTCb_Vta "
    .Source = .Source & "FROM tgtcb "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    If ps_Plataforma = pSrvMySql Then
      .Source = .Source & "AND DATE_FORMAT(fehtcb,'%d/%m/%Y')=DATE_FORMAT('" & Format(s_dExpresion, "yyyy-mm-dd") & "', '%d/%m/%Y')"
    ElseIf ps_Plataforma = pSrvSql Then
      .Source = .Source & "AND CONVERT(smalldatetime, FehTCb, 103)=CONVERT(smalldatetime, '" & Format(s_dExpresion, "dd/mm/yyyy") & "', 103)"
    End If
    .Open
    n_TipoCambio = IIf(.EOF, 0, !nImpTCb_Vta)
    .Close
  End With
  ' valida tipo de cambio
  If n_TipoCambio = 0 Then Beep: MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical: Exit Sub
  
  ' Valido el presupuesto
  If Not ValidoPresupuesto(n_TipoCambio) Then Exit Sub
  
  ' Generacion informacion
  With porstCancel
    If .State = adStateOpen Then .Close
    .Source = "SELECT rte.sernegocio, rte.nronegocio, rte.codaux, rte.fehope, rte.feedoc, rte.fevdoc, rte.refdoc, rte.glodoc, rte.glodocx, "
    .Source = .Source & "rte.codtdc, rte.serdoc, rte.coddro, rte.tpoglo_rtc, rte.glodoc_rtc, rte.tpomon, rte.pctigv, rte.pctisc, rte.impogr, "
    .Source = .Source & "rte.impexp, rte.impexo, rte.impigv, rte.impisc, rte.impoim, rte.imptot "
    .Source = .Source & "FROM CoRteVta rte "
    .Source = .Source & "WHERE rte.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND rte.indestado='" & ESTCCO_ACT & "' "
    .Source = .Source & "AND NOT EXISTS(SELECT * FROM covtadoc vta WHERE vta.codemp=rte.codemp AND vta.codaux=rte.codaux AND vta.codtdc=rte.codtdc AND vta.serdoc=rte.serdoc "
    .Source = .Source & "AND vta.refdoc=CONCAT(rte.sernegocio,'-',rte.nronegocio) AND vta.pdoano='" & gsAnoAct & "' AND vta.mespvs='" & gsMesAct & "') "
    .Source = .Source & "ORDER BY rte.codtdc, rte.serdoc, rte.sernegocio, rte.nronegocio, rte.codaux"
    .Open
  End With
  'Valido las Cuentas esten Correctas(llenas para todas los valores)
  If porstCancel.RecordCount > 0 Then
    While Not porstCancel.EOF
      If VerificaCtaCCo(porstCancel) Then
        ' Numero documento de ventas
        If Not ((s_TipoDocumento = porstCancel!codtdc) And (s_SerieDocumento = porstCancel!serdoc)) Then
          s_TipoDocumento = porstCancel!codtdc
          s_SerieDocumento = porstCancel!serdoc
          s_NroDocumento = pfNumFacturaVenta(s_TipoDocumento, s_SerieDocumento)
        End If
        s_NroDocumento = gfCeros(s_NroDocumento, 10, 1, "0")
        ' Numero comprobante contable
        If Not (s_DiarioDocumento = porstCancel!coddro) Then
          s_DiarioDocumento = porstCancel!coddro
          s_NroComprobante = gfNumComprobante(gsAnoAct, gsMesAct, s_DiarioDocumento)
          s_NroComprobante = gfCeros(s_NroComprobante, 6, -1, "0")
        End If
        s_NroComprobante = gfCeros(s_NroComprobante, 6, 1, "0")
        ' Genero documentos de ventas
        ppGeneraVtaDoc s_NroDocumento, s_dExpresion, n_TipoCambio, s_DiarioDocumento, s_NroComprobante, porstCancel
      End If
      porstCancel.MoveNext
    Wend
    ' primer paso: actualizo sumatoria documento
    s_Sentencia = "UPDATE covtadoc SET "
    s_Sentencia = s_Sentencia & "imptot_mn=ROUND((impogr_mn+impexp_mn+impexo_mn+impigv_mn+impisc_mn+impoim_mn), 2), "
    s_Sentencia = s_Sentencia & "imptot_me=ROUND((impogr_me+impexp_me+impexo_me+impigv_me+impisc_me+impoim_me), 2) "
    s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "AND mespvs='" & gsMesAct & "' "
    s_Sentencia = s_Sentencia & "AND IFNULL(refdoc, '')<>'' "
    s_Sentencia = s_Sentencia & "AND indpregen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND indgen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND indanu='" & INDANU_FAL & "' "
    s_Sentencia = s_Sentencia & "AND ((ABS(ROUND(imptot_mn-ROUND((impogr_mn+impexp_mn+impexo_mn+impigv_mn+impisc_mn+impoim_mn), 2), 2))<=0.01) "
    s_Sentencia = s_Sentencia & "OR (ABS(ROUND(imptot_me-ROUND((impogr_me+impexp_me+impexo_me+impigv_me+impisc_me+impoim_me), 2), 2))<=0.01))"
    uocnnMain.Execute s_Sentencia, nNumeRegistro
    ' segundo paso: Sumatoria de cuentas
    uocnnMain.Execute "DROP TABLE IF EXISTS tmpventacta", nNumeRegistro
    s_Sentencia = "CREATE TEMPORARY TABLE IF NOT EXISTS tmpventacta "
    s_Sentencia = s_Sentencia & "SELECT cta.codemp, cta.pdoano, cta.codtdc, cta.serdoc, cta.nrodoc, cta.tpocnc, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(CASE WHEN cta.orden<>'01' THEN cta.impcta_mn ELSE 0 END), 2) AS impcta_mn, "
    s_Sentencia = s_Sentencia & "ROUND(SUM(CASE WHEN cta.orden<>'01' THEN cta.impcta_me ELSE 0 END), 2) AS impcta_me, "
    s_Sentencia = s_Sentencia & "ROUND(AVG(CASE cta.tpocnc WHEN '1' THEN vta.impogr_mn WHEN '2' THEN vta.impexp_mn WHEN '3' THEN vta.impexo_mn WHEN '4' THEN vta.impigv_mn WHEN '5' THEN vta.impoim_mn ELSE vta.imptot_mn END), 2) AS impvta_mn, "
    s_Sentencia = s_Sentencia & "ROUND(AVG(CASE cta.tpocnc WHEN '1' THEN vta.impogr_me WHEN '2' THEN vta.impexp_me WHEN '3' THEN vta.impexo_me WHEN '4' THEN vta.impigv_me WHEN '5' THEN vta.impoim_me ELSE vta.imptot_me END), 2) AS impvta_me "
    s_Sentencia = s_Sentencia & "FROM covtadoccta cta "
    s_Sentencia = s_Sentencia & "INNER JOIN covtadoc vta ON vta.codemp=cta.codemp AND vta.pdoano=cta.pdoano AND vta.codtdc=cta.codtdc AND vta.serdoc=cta.serdoc AND vta.nrodoc=cta.nrodoc "
    s_Sentencia = s_Sentencia & "AND vta.mespvs='" & gsMesAct & "' "
    s_Sentencia = s_Sentencia & "AND IFNULL(vta.refdoc, '')<>'' "
    s_Sentencia = s_Sentencia & "AND vta.indpregen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND vta.indgen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND vta.indanu='" & INDANU_FAL & "' "
    s_Sentencia = s_Sentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND cta.pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "GROUP BY cta.codemp, cta.pdoano, cta.codtdc, cta.serdoc, cta.nrodoc, cta.tpocnc "
    s_Sentencia = s_Sentencia & "ORDER BY cta.codemp, cta.pdoano, cta.codtdc, cta.serdoc, cta.nrodoc, cta.tpocnc"
    uocnnMain.Execute s_Sentencia, nNumeRegistro
    ' tercer paso: actualizo importe de cuenta inicial
    s_Sentencia = "UPDATE covtadoccta cta, covtadoc vta, tmpventacta tmp SET "
    s_Sentencia = s_Sentencia & "cta.impcta_mn=ROUND((tmp.impvta_mn-tmp.impcta_mn), 2), "
    s_Sentencia = s_Sentencia & "cta.impcta_me=ROUND((tmp.impvta_me-tmp.impcta_me), 2) "
    s_Sentencia = s_Sentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND cta.pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "AND cta.orden='01' "
    s_Sentencia = s_Sentencia & "AND vta.codemp=cta.codemp "
    s_Sentencia = s_Sentencia & "AND vta.pdoano=cta.pdoano "
    s_Sentencia = s_Sentencia & "AND vta.codtdc=cta.codtdc "
    s_Sentencia = s_Sentencia & "AND vta.serdoc=cta.serdoc "
    s_Sentencia = s_Sentencia & "AND vta.nrodoc=cta.nrodoc "
    s_Sentencia = s_Sentencia & "AND vta.mespvs='" & gsMesAct & "' "
    s_Sentencia = s_Sentencia & "AND IFNULL(vta.refdoc, '')<>'' "
    s_Sentencia = s_Sentencia & "AND vta.indpregen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND vta.indgen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND vta.indanu='" & INDANU_FAL & "' "
    s_Sentencia = s_Sentencia & "AND cta.codemp=tmp.codemp "
    s_Sentencia = s_Sentencia & "AND cta.pdoano=tmp.pdoano "
    s_Sentencia = s_Sentencia & "AND cta.codtdc=tmp.codtdc "
    s_Sentencia = s_Sentencia & "AND cta.serdoc=tmp.serdoc "
    s_Sentencia = s_Sentencia & "AND cta.nrodoc=tmp.nrodoc "
    s_Sentencia = s_Sentencia & "AND cta.tpocnc=tmp.tpocnc"
    uocnnMain.Execute s_Sentencia, nNumeRegistro
    uocnnMain.Execute "DROP TABLE IF EXISTS tmpventacta", nNumeRegistro
    
    ' cuarto paso: Sumatoria de centro costos
    With porstCancel
      If .State = adStateOpen Then .Close
      .Source = "SELECT cco.codemp, cco.pdoano, cco.codtdc, cco.serdoc, cco.nrodoc, cco.tpocnc, cco.orden, cco.codcta, cco.codcco, "
      .Source = .Source & "ROUND(SUM(cco.impcco_mn), 2) AS impcco_mn, ROUND(SUM(cco.impcco_me), 2) AS impcco_me, "
      .Source = .Source & "ROUND(AVG(cta.impcta_mn), 2) AS impcta_mn, ROUND(AVG(cta.impcta_me), 2) AS impcta_me, "
      .Source = .Source & "CONCAT(cco.codtdc, cco.serdoc, cco.nrodoc, cco.tpocnc, cco.orden, cco.codcta) AS sPrimaryKey "
      .Source = .Source & "FROM covtadoccco cco "
      .Source = .Source & "INNER JOIN covtadoccta cta ON cta.codemp=cco.codemp AND cta.pdoano=cco.pdoano AND cta.codtdc=cco.codtdc AND cta.serdoc=cco.serdoc AND cta.nrodoc=cco.nrodoc AND cta.tpocnc=cco.tpocnc AND cta.orden=cco.orden AND cta.codcta=cco.codcta "
      .Source = .Source & "INNER JOIN covtadoc vta ON vta.codemp=cta.codemp AND vta.pdoano=cta.pdoano AND vta.codtdc=cta.codtdc AND vta.serdoc=cta.serdoc AND vta.nrodoc=cta.nrodoc "
      .Source = .Source & "AND vta.mespvs='" & gsMesAct & "' "
      .Source = .Source & "AND IFNULL(vta.refdoc, '')<>'' "
      .Source = .Source & "AND vta.indpregen='" & INDPREGEN_INA & "' "
      .Source = .Source & "AND vta.indgen='" & INDPREGEN_INA & "' "
      .Source = .Source & "AND vta.indanu='" & INDANU_FAL & "' "
      .Source = .Source & "WHERE cco.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND cco.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND cco.orden='01' "
      .Source = .Source & "GROUP BY cco.codemp, cco.pdoano, cco.codtdc, cco.serdoc, cco.nrodoc, cco.tpocnc, cco.orden, cco.codcta "
      .Source = .Source & "ORDER BY cco.codemp, cco.pdoano, cco.codtdc, cco.serdoc, cco.nrodoc, cco.tpocnc, cco.orden, cco.codcta, cco.codcco"
      .Open
    End With
    
    s_Expresion = ""
    If porstCancel.RecordCount > 0 Then
      While Not porstCancel.EOF
        ' cuenta diferente
        If Not (s_Expresion = porstCancel!sPrimaryKey) Then
          s_Expresion = porstCancel!sPrimaryKey
          n_ImporteMN = Round((CDec(porstCancel!impcta_mn) - CDec(porstCancel!impcco_mn)), 2)
          n_ImporteME = Round((CDec(porstCancel!impcta_me) - CDec(porstCancel!impcco_me)), 2)
          ' actualizo importe centro costo inicial (diferencia)
          If Not (n_ImporteMN = 0 And n_ImporteME = 0) Then
            s_Sentencia = "UPDATE covtadoccco cco SET "
            s_Sentencia = s_Sentencia & "cco.impcco_mn=ROUND((cco.impcco_mn+" & CDec(n_ImporteMN) & "), 2), "
            s_Sentencia = s_Sentencia & "cco.impcco_me=ROUND((cco.impcco_me+" & CDec(n_ImporteME) & "), 2) "
            s_Sentencia = s_Sentencia & "WHERE cco.codemp='" & gsCodEmp & "' "
            s_Sentencia = s_Sentencia & "AND cco.pdoano='" & gsAnoAct & "' "
            s_Sentencia = s_Sentencia & "AND cco.codtdc='" & porstCancel!codtdc & "' "
            s_Sentencia = s_Sentencia & "AND cco.serdoc='" & porstCancel!serdoc & "' "
            s_Sentencia = s_Sentencia & "AND cco.nrodoc='" & porstCancel!nrodoc & "' "
            s_Sentencia = s_Sentencia & "AND cco.tpocnc='" & porstCancel!tpocnc & "' "
            s_Sentencia = s_Sentencia & "AND cco.orden='" & porstCancel!orden & "' "
            s_Sentencia = s_Sentencia & "AND cco.codcta='" & porstCancel!CodCta & "' "
            s_Sentencia = s_Sentencia & "AND cco.codcco='" & porstCancel!codcco & "'"
            uocnnMain.Execute s_Sentencia, nNumeRegistro
          End If
        End If
        porstCancel.MoveNext
      Wend
    End If
    ' quinto paso: actualizo indicadores
    s_Sentencia = "UPDATE covtadoc SET indanu='" & INDANU_VER & "' "
    s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND pdoano='" & gsAnoAct & "' "
    s_Sentencia = s_Sentencia & "AND mespvs='" & gsMesAct & "' "
    s_Sentencia = s_Sentencia & "AND IFNULL(refdoc, '')<>'' "
    s_Sentencia = s_Sentencia & "AND indpregen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND indgen='" & INDPREGEN_INA & "' "
    s_Sentencia = s_Sentencia & "AND indanu='" & INDANU_FAL & "'"
    uocnnMain.Execute s_Sentencia, nNumeRegistro
    
    ' Actualizo la grilla
    uorstMain.Requery
    uorstMain_Grd.Requery
    upDatosGrid
  End If
  porstCancel.Close
  MsgBox TEXT_8008, vbInformation

End Sub
Private Sub cmdImprimir_Click(Index As Integer)
  '[Datos del formulario de impresión.  'Cambiar.
  frmLVta.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
  frmLVta.Show vbModal
  ']
End Sub

Private Sub Form_Load()
  '[Recordsets                          'Cambiar.
  psConnStrgSele_Grd = "SELECT cortevta.codaux, b.razaux, cortevta.sernegocio, cortevta.nronegocio, cortevta.feedoc, cortevta.coddro, c.abvtdc, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "cortevta.serdoc, cortevta.tpomon, cortevta.imptot, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE WHEN " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(vta.nrodoc, '')<>'' THEN 'x' ELSE ' ' END) AS cIndGen, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE WHEN cortevta.indestado='" & ESTCCO_ACT & "' THEN '" & ESTCCO_ACT_TXT & "' ELSE '" & ESTCCO_INA_TXT & "' END) as cIndGen, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "cortevta.codtdc, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cortevta.sernegocio, cortevta.nronegocio)", "(cortevta.sernegocio+cortevta.nronegocio)") & " AS cLlave "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "FROM CoRteVta "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "INNER JOIN TGAux b ON cortevta.codemp = b.codemp AND cortevta.CodAux = b.CodAux "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGTDc c ON cortevta.codemp = c.codemp AND cortevta.CodTDc = c.CodTDc "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN CoVtaDoc vta ON vta.codemp=cortevta.codemp AND vta.codaux=cortevta.codaux AND vta.codtdc=cortevta.codtdc AND vta.serdoc=cortevta.serdoc AND vta.refdoc=CONCAT(cortevta.sernegocio,'-',cortevta.nronegocio) AND vta.pdoano='" & gsAnoAct & "' AND vta.mespvs='" & gsMesAct & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "WHERE cortevta.codemp='" & gsCodEmp & "' "
  
  psConnStrgSele = "SELECT cortevta.sernegocio, cortevta.nronegocio, cortevta.codaux, cortevta.fehope, cortevta.feedoc, cortevta.fevdoc, cortevta.refdoc, cortevta.glodoc, cortevta.glodocx, cortevta.codtdc, cortevta.serdoc, "
  psConnStrgSele = psConnStrgSele & "cortevta.coddro, cortevta.tpoglo_rtc, cortevta.glodoc_rtc, cortevta.tpomon, cortevta.pctigv, cortevta.pctisc, cortevta.impogr, cortevta.impexp, cortevta.impexo, "
  psConnStrgSele = psConnStrgSele & "cortevta.impigv, cortevta.impisc, cortevta.impoim, cortevta.imptot, cortevta.indestado, "
  psConnStrgSele = psConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cortevta.sernegocio, cortevta.nronegocio)", "(cortevta.sernegocio+cortevta.nronegocio)") & " AS cLlave, "
  psConnStrgSele = psConnStrgSele & "cortevta.UsrCre, cortevta.FyHCre, cortevta.UsrMdf, cortevta.FyHMdf, "
  psConnStrgSele = psConnStrgSele & "cortevta.codemp, cortevta.pdoano "
  psConnStrgSele = psConnStrgSele & "FROM CoRteVta "
  psConnStrgSele = psConnStrgSele & "WHERE cortevta.codemp='" & gsCodEmp & "' "
  psConnStrgOrde = "ORDER BY cortevta.indestado desc, cortevta.sernegocio, cortevta.nronegocio"
  
  usConnStrgSele_CoRteVtaCta = "SELECT CoRteVtaCta.CodCta, CoRteVtaCta.porimpcta, "
  If gsIdioma = INDCCO_ACT Then
    usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(CoRteVtaCta.GloDet0, ''), IFNULL(CoRteVtaCta.GloDet1, ''))", "(ISNULL(CoRteVtaCta.GloDet0, '')+ISNULL(CoRteVtaCta.GloDet1, ''))") & " AS GloDet, "
  Else
    usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(CoRteVtaCta.GloDet0x, ''), IFNULL(CoRteVtaCta.GloDet1x, ''))", "(ISNULL(CoRteVtaCta.GloDet0x, '')+ISNULL(CoRteVtaCta.GloDet1x, ''))") & " AS GloDetx, "
  End If
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & "CoRteVtaCta.CodRuc, CoRteVtaCta.GloDet0, CoRteVtaCta.GloDet1, CoRteVtaCta.GloDet0x, CoRteVtaCta.GloDet1x, "
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & "CoRteVtaCta.sernegocio, CoRteVtaCta.nronegocio, "
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & "CoRteVtaCta.TpoCnc, CoRteVtaCta.Orden, "
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(CoRteVtaCta.sernegocio, CoRteVtaCta.nronegocio, CoRteVtaCta.TpoCnc, CoRteVtaCta.Orden)", "(CoRteVtaCta.sernegocio+CoRteVtaCta.nronegocio+RTrim(CoRteVtaCta.TpoCnc)+CoRteVtaCta.Orden)") & " AS cLlave, "
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(CoRteVtaCta.sernegocio, CoRteVtaCta.nronegocio, CoRteVtaCta.TpoCnc, CoRteVtaCta.Orden, CoRteVtaCta.CodCta)", "(CoRteVtaCta.sernegocio+CoRteVtaCta.nronegocio+RTrim(CoRteVtaCta.TpoCnc)+CoRteVtaCta.Orden+CoRteVtaCta.CodCta)") & " AS cLlave2, "
  If gsIdioma = INDCCO_ACT Then
    usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(CoRteVtaCta.GloDet0x, ''), IFNULL(CoRteVtaCta.GloDet1x, ''))", "(ISNULL(CoRteVtaCta.GloDet0x, '')+ISNULL(CoRteVtaCta.GloDet1x, ''))") & " AS GloDetx, "
  Else
    usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & IIf(ps_Plataforma = pSrvMySql, "CONCAT(IFNULL(CoRteVtaCta.GloDet0, ''), IFNULL(CoRteVtaCta.GloDet1, ''))", "(ISNULL(CoRteVtaCta.GloDet0, '')+ISNULL(CoRteVtaCta.GloDet1, ''))") & " AS GloDet, "
  End If
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & "CoRteVtaCta.UsrCre, CoRteVtaCta.FyHCre, CoRteVtaCta.UsrMdf, CoRteVtaCta.FyHMdf, "
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & "CoRteVtaCta.codemp, CoRteVtaCta.pdoano "
  usConnStrgSele_CoRteVtaCta = usConnStrgSele_CoRteVtaCta & "FROM CoRteVtaCta "
  usConnStrgWher_CoRteVtaCta = ""
  usConnStrgOrde_CoRteVtaCta = "ORDER BY 11, 12, 1" ' DESC"

  usConnStrgSele_CoRteVtaCCo = "SELECT CoRteVtaCCo.CodCCo, CoRteVtaCCo.porimpcco, "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & "CoRteVtaCCo.TpoCnc, CoRteVtaCCo.CodCta, CoRteVtaCCo.Orden, "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & "CoRteVtaCCo.sernegocio, CoRteVtaCCo.nronegocio, "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & IIf(ps_Plataforma = pSrvMySql, "CONCAT(CoRteVtaCCo.TpoCnc, CoRteVtaCCo.Orden, CoRteVtaCCo.CodCta)", "(RTrim(CoRteVtaCCo.TpoCnc)+CoRteVtaCCo.Orden+CoRteVtaCCo.CodCta)") & " AS cLlave, "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & IIf(ps_Plataforma = pSrvMySql, "CONCAT(CoRteVtaCCo.sernegocio, CoRteVtaCCo.nronegocio, CoRteVtaCCo.TpoCnc, CoRteVtaCCo.Orden, CoRteVtaCCo.CodCta)", "(CoRteVtaCCo.sernegocio+CoRteVtaCCo.nronegocio+RTrim(CoRteVtaCCo.TpoCnc)+CoRteVtaCCo.Orden+CoRteVtaCCo.CodCta)") & " AS cLlave1, "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & IIf(ps_Plataforma = pSrvMySql, "CONCAT(CoRteVtaCCo.sernegocio, CoRteVtaCCo.nronegocio, CoRteVtaCCo.TpoCnc, CoRteVtaCCo.Orden, CoRteVtaCCo.CodCta, CoRteVtaCCo.CodCCo)", "(CoRteVtaCCo.sernegocio+CoRteVtaCCo.nronegocio+RTrim(CoRteVtaCCo.TpoCnc)+CoRteVtaCCo.Orden+CoRteVtaCCo.CodCta+CoRteVtaCCo.CodCCo)") & " AS cLlave2, "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & "CoRteVtaCCo.UsrCre, CoRteVtaCCo.FyHCre, CoRteVtaCCo.UsrMdf, CoRteVtaCCo.FyHMdf, "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & "CoRteVtaCCo.codemp, CoRteVtaCCo.pdoano "
  usConnStrgSele_CoRteVtaCCo = usConnStrgSele_CoRteVtaCCo & "FROM CoRteVtaCCo "
  usConnStrgWher_CoRteVtaCCo = ""
  usConnStrgOrde_CoRteVtaCCo = "ORDER BY 3, 5, 4, 1"
  
  usConnStrgSele_COCpbDet = "SELECT cta.codcta, cta.codruc, cco.codcco, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & Choose(gsIdioma, "cta.glodet0", "cta.glodet0x") & " AS gloite, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "vta.tpomon, "
'  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "(CASE WHEN cta.tpocnc='" & TPOCNC_TOT_VTA & "' THEN  tdc.sgntdc CASE WHEN (porimpcta>0 THEN  vta.tpomon, "
'  If (frmTRteVtaGrd.uorstCoRteVtaCCo!impcco_me > 0) And (frmTRteVtaGrd.uorstCoRteVtaCCo!impcco_mn > 0) Then
'    !TpoCtb = "(CASE WHEN cta.tpocnc='" & TPOCNC_TOT_VTA & "' THEN (CASE WHEN tdc.sgntdc='" & SGNTDC_POS & "' THEN importe ELSE  0 END) ELSE (CASE WHEN tdc.sgntdc='" & SGNTDC_NEG & "' THEN importe ELSE  0 END) END)"
'  Else
'    !TpoCtb = "(CASE WHEN cta.tpocnc='" & TPOCNC_TOT_VTA & "' THEN (CASE WHEN tdc.sgntdc='" & SGNTDC_NEG & "' THEN importe ELSE  0 END) ELSE (CASE WHEN tdc.sgntdc='" & SGNTDC_POS & "' THEN importe ELSE  0 END) END)"
'  End If
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "ROUND(("
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "(CASE cta.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim ELSE vta.imptot END) * "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "ROUND((cta.porimpcta*IFNULL(cco.porimpcco, 100))/100, 2))/100, 2) AS cImporDeb, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "cta.tpocnc, cta.orden "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "FROM cortevtacta cta "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "INNER JOIN cortevta vta ON vta.codemp=cta.codemp AND vta.pdoano=cta.pdoano AND vta.sernegocio=cta.sernegocio AND vta.nronegocio=cta.nronegocio "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "INNER JOIN tgtdc tdc ON tdc.codemp=vta.codemp AND tdc.codtdc=vta.codtdc "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "LEFT JOIN cortevtacco cco ON cco.codemp=cta.codemp AND cco.pdoano=cta.pdoano AND cco.sernegocio=cta.sernegocio AND cco.nronegocio=cta.nronegocio AND cco.tpocnc=cta.tpocnc AND cco.orden=cta.orden AND cco.codcta=cta.codcta "
  usConnStrgWher_COCpbDet = "WHERE cta.codemp='" & gsCodEmp & "' "
  usConnStrgWher_COCpbDet = usConnStrgWher_COCpbDet & "AND cta.sernegocio='' AND cta.nronegocio='' "
  usConnStrgOrde_COCpbDet = "ORDER BY cta.tpocnc, cta.codcta, cta.orden"
  
  Set uocnnMain = New ADODB.Connection
  Set uocnnNoGrabable = New ADODB.Connection
  Set uorstMain = New ADODB.Recordset
  Set uorstMain_Grd = New ADODB.Recordset
  Set uorstTGAux = New ADODB.Recordset
  Set uorstTGTDc = New ADODB.Recordset
  Set uorstCoCta = New ADODB.Recordset
  Set uorstCoCCo = New ADODB.Recordset
  Set uorstCODro = New ADODB.Recordset
  Set uorstCoRteVtaCta = New ADODB.Recordset
  Set uorstCoRteVtaCCo = New ADODB.Recordset
  Set uorstCOCpbDet = New ADODB.Recordset
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
    .Properties("Unique Table").Value = "CoRteVta"
  End With
  With uorstMain
    .ActiveConnection = uocnnMain
    .Source = psConnStrgSele & psConnStrgOrde
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic 'adLockReadOnly
    .Open
    .Properties("Unique Table").Value = "CoRteVta"
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
  With uorstCoRteVtaCta
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
  End With
  With uorstCoRteVtaCCo
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
  End With
  With uorstCOCpbDet
    .ActiveConnection = uocnnMain
    .Source = psConnStrgSele & psConnStrgOrde
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
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
  uorstCoCta.Close
  uorstCoCCo.Close
  uorstCODro.Close
  '[ARREGLAR. Genera demora al salir de la opción.
  If uorstCoRteVtaCta.State = adStateOpen Then uorstCoRteVtaCta.Close
  If uorstCoRteVtaCCo.State = adStateOpen Then uorstCoRteVtaCCo.Close
  ']ARREGLAR.
  uorstCOCpbDet.Close
  uorstMain_Grd.Close
  uorstMain.Close
  uocnnMain.Close
  Set porstCancel = Nothing
  Set uorstTemporal = Nothing
  Set uorstTGAux = Nothing
  Set uorstTGTDc = Nothing
  Set uorstCoCta = Nothing
  Set uorstCoCCo = Nothing
  Set uorstCODro = Nothing
  Set uorstCoRteVtaCta = Nothing
  Set uorstCoRteVtaCCo = Nothing
  Set uorstCOCpbDet = Nothing
  Set uorstMain_Grd = Nothing
  Set uorstMain = Nothing
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
   gpTUg_Nuevo Me, frmTRteVta             'Cambiar Formulario de Datos.
'///Angel 12/12/2003
'/// Agregado para eliminar el registro creado como cabecera al intentar registrar un dato de cuenta y luego cancelar el ingreso completo.
   cmdRefrescar_Click
'///
End Sub

Public Sub cmdRevisar_click()
  On Error GoTo Err
  
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub
  
  '[Propio del formulario.
  ubGrabaMas = INDMASCTA_CTA
  ']
  
  '[Búsqueda del ítem.
  uorstMain.Requery
  uorstMain.MoveFirst
  uorstMain.Find "cLlave='" & uorstMain_Grd!sernegocio & uorstMain_Grd!nronegocio & "'"
  ']
  
  With frmTRteVta                        'Cambiar Formulario de Datos.
    .zbNuevo = False
    .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
    .txtLlave(0).Enabled = False
    .txtLlave(1).Enabled = False
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
  If gbCieVta Then MsgBox TEXT_9016, vbCritical: Exit Sub
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub
 'ini 2016-05-27/28 nivel=asisten no elimin datos
   If gsNvlUsr = NVLUSR_ASIS Then
      MsgBox TEXT_9026, vbCritical
      Exit Sub
   End If
'fin 2016-05-27/28 nivel=asisten no elimin datos
  'Mensaje de verificación            'Cambiar.
  If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(2)) & "-" & Trim(dgrMain.Columns(3)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
    With porstCancel
      .Source = "SELECT mespvs, codaux, codtdc, serdoc, nrodoc "
      .Source = .Source & "FROM covtadoc "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND MesPvs='" & gsMesAct & "' AND CodAux='" & uorstMain_Grd!codaux & "' "
      .Source = .Source & "AND CodTDc='" & uorstMain_Grd!codtdc & "' AND SerDoc='" & uorstMain_Grd!serdoc & "' "
      .Source = .Source & "AND refdoc='" & uorstMain_Grd!sernegocio & "-" & uorstMain_Grd!nronegocio & "'"
      .Open
      If porstCancel.RecordCount = 0 Then
        uorstMain.MoveFirst
        uorstMain.Find "cLlave = '" & uorstMain_Grd!sernegocio & uorstMain_Grd!nronegocio & "'"
        
        uocnnMain.BeginTrans       'INICIA TRANSACCION.
        uocnnMain.Execute "DELETE FROM covtadoc WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND MesPvs='" & gsMesAct & "' AND codaux='" & Trim(dgrMain.Columns(0)) & "' AND refdoc='" & Trim(dgrMain.Columns(2)) & "-" & Trim(dgrMain.Columns(3)) & "' And codtdc='" & uorstMain_Grd!codtdc & "' And serdoc='" & Trim(dgrMain.Columns(7)) & "'"
        uorstMain.Properties("Unique Table").Value = "CoRteVta"
        uorstMain.Delete
        uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.
        
        'Busca siguiente ítem.
        With uorstMain_Grd
          .MoveNext
          If .EOF Then .MoveLast
          dsLlaveSiguiente = !sernegocio & !nronegocio
          .Requery
          If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
        End With
      Else
        MsgBox Choose(gsIdioma, "Debe eliminar antes Documentos de Ventas.", "The Documents  Sales must be eliminated before."), vbExclamation
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
  '[ARREGLAR. Usar gpTUg_Refrescar Me, pero se debe cambiar ppDatosGrid a upDatosGrid para todos los formularios que lo usan (formularios de registro único).
  ' gpTUg_Refrescar Me
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
  s_Sentencia = s_Sentencia & "FROM CoRteVta a "
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
   Case Else
    psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
  End Select
  With uorstMain_Grd
    .Close
    .Properties("Unique Table").Value = "CoRteVta"
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
  
'  psConnStrgSele_Grd = "SELECT cortevta.CodAux, b.RazAux, cortevta.diapvs, cortevta.secuencia, cortevta.FeEDoc, cortevta.coddro, c.abvtdc, "
'  psConnStrgSele_Grd = psConnStrgSele_Grd & "cortevta.SerDoc, cortevta.tpomon, cortevta.imptot, "
'  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE WHEN cortevta.indestado=1 THEN 'Act' ELSE 'Ina' END) as cIndGen, "
'  psConnStrgSele_Grd = psConnStrgSele_Grd & "c.codtdc, "
'  psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cortevta.mespvs, cortevta.diapvs, cortevta.secuencia)", "(cortevta.mespvs+cortevta.diapvs+cortevta.secuencia)") & " AS cLlave "
  
  
  With dgrMain.Columns
    For dnNum = 0 To .Count - 1
      Select Case dnNum
       Case 0
        .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
        .Item(dnNum).Width = 1100
       Case 1
        .Item(dnNum).Caption = Choose(gsIdioma, "Razón Social", "Firm Name")
        .Item(dnNum).Width = 1750
       Case 2
        .Item(dnNum).Caption = Choose(gsIdioma, "Sng", "Sng")
        .Item(dnNum).Width = 500
       Case 3
        .Item(dnNum).Caption = Choose(gsIdioma, "Número", "Number")
        .Item(dnNum).Width = 1000
       Case 4
        .Item(dnNum).Caption = Choose(gsIdioma, "F.Emisión", "Issue Date")
        .Item(dnNum).Width = 1000
       Case 5
        .Item(dnNum).Caption = Choose(gsIdioma, "Diario", "Journal")
        .Item(dnNum).Width = 500
       Case 6
        .Item(dnNum).Caption = Choose(gsIdioma, "TDc", "TDc")
        .Item(dnNum).Width = 500
       Case 7
        .Item(dnNum).Caption = Choose(gsIdioma, "Ser", "Ser")
        .Item(dnNum).Width = 470
       Case 8
        .Item(dnNum).Caption = Choose(gsIdioma, "Mon", "Cur")
        .Item(dnNum).Width = 250
       Case 9
        .Item(dnNum).Caption = Choose(gsIdioma, "Total", "Total")
        .Item(dnNum).Width = 1100
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
        .Item(dnNum).Alignment = dbgRight
       Case 10
        .Item(dnNum).Caption = "G"
        .Item(dnNum).Width = 230
        .Item(dnNum).Alignment = dbgCenter
       Case 11
        .Item(dnNum).Caption = "Ok"
        .Item(dnNum).Width = 300
        .Item(dnNum).Alignment = dbgLeft
       Case Else
        .Item(dnNum).Visible = False
      End Select
    Next
  End With
End Sub

'[Código propio del formulario.
Private Function pfNumFacturaVenta(ByVal s_TipoDocu As String, s_SerieDocu As String) As String
  Dim porstRetorno As ADODB.Recordset
  Dim s_Sentencia As String
  
  s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(nrodoc), '0000000000') AS cNumMaxDoc "
  s_Sentencia = s_Sentencia & "FROM covtadoc "
  s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND CONCAT(pdoano, mespvs)<='" & gsAnoAct & gsMesAct & "' "
  s_Sentencia = s_Sentencia & "AND codtdc='" & s_TipoDocu & "' "
  s_Sentencia = s_Sentencia & "AND serdoc='" & s_SerieDocu & "'"
  Set porstRetorno = New ADODB.Recordset
  With porstRetorno
    .ActiveConnection = uocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = s_Sentencia
    .Open
  End With
  pfNumFacturaVenta = gfCeros(porstRetorno!cNumMaxDoc, 10, 0, "0")
  porstRetorno.Close
  Set porstRetorno = Nothing

End Function
Private Sub ppGeneraVtaDoc(ByVal sDocumentoVenta As String, ByVal sFechaProceso As String, ByVal nTipoCambio As Double, ByVal sDiaroContable As String, ByVal sComprobante As String, ByVal oRecordset As ADODB.Recordset)
  Dim nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nNumRegistros As Long
  Dim sSentencia As String, sExpresion As String

  On Error GoTo ErrGrabar
  
  uocnnMain.BeginTrans            'INICIA TRANSACCION.
      
  ' Grabación de cabecera de comprobante
  sSentencia = "INSERT INTO covtadoc(codemp, pdoano, codtdc, serdoc, nrodoc, fehope, serdoc_fin, nrodoc_fin, codaux, feedoc, fevdoc, tpomon, imptcb, pctigv, pctisc, refdoc, glodoc, "
  sSentencia = sSentencia & "glodocx, tpoglo_rtc, glodoc_rtc, codcon, mespvs, coddro, nrocpb, codasi, impogr_mn, impogr_me, indcta_ogr, impexp_mn, impexp_me, indcta_exp, impexo_mn, impexo_me, "
  sSentencia = sSentencia & "indcta_exo, impigv_mn, impigv_me, indcta_igv, impisc_mn, impisc_me, indcta_isc, impoim_mn, impoim_me, indcta_oim, imptot_mn, imptot_me, indcta_tot, "
  sSentencia = sSentencia & "indpregen, indgen, indanu, tpoimpuesto, categoriadoc, indvtaext, codaduana, annodua, nrodua, feembarq, "
  sSentencia = sSentencia & "feregula, impfob_mn, impfob_me, indpercep, tsapercep, serpercep, nropercep, codtdc_ref, serdoc_ref, nrodoc_ref, feedoc_ref, impbasref_mn, "
  sSentencia = sSentencia & "impigvref_mn, impbasref_me, impigvref_me, usrcre, fyhcre, usrmdf, fyhmdf) "
  sSentencia = sSentencia & "VALUES("
  sSentencia = sSentencia & "'" & gsCodEmp & "', "
  sSentencia = sSentencia & "'" & gsAnoAct & "', "
  sSentencia = sSentencia & "'" & oRecordset!codtdc & "', "
  sSentencia = sSentencia & "'" & oRecordset!serdoc & "', "
  sSentencia = sSentencia & "'" & sDocumentoVenta & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(sFechaProceso, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & "Null, Null, "
  sSentencia = sSentencia & "'" & oRecordset!codaux & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(sFechaProceso, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(sFechaProceso, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & "'" & oRecordset!tpomon & "', "
  sSentencia = sSentencia & CDec(nTipoCambio) & ", "
  sSentencia = sSentencia & CDec(oRecordset!PctIGV) & ", "
  sSentencia = sSentencia & CDec(oRecordset!PctISC) & ", "
  ' referencia de negocio
  sExpresion = oRecordset!sernegocio & "-" & oRecordset!nronegocio
  sSentencia = sSentencia & "'" & sExpresion & "', "
  sSentencia = sSentencia & IIf(IsNull(oRecordset!GloDoc), "Null", "'" & oRecordset!GloDoc & "'") & ", "
  sSentencia = sSentencia & IIf(IsNull(oRecordset!glodocx), "Null", "'" & oRecordset!glodocx & "'") & ", "
  sSentencia = sSentencia & "'" & oRecordset!TpoGlo_Rtc & "', "
  sSentencia = sSentencia & IIf(IsNull(oRecordset!glodoc_rtc), "Null", "'" & oRecordset!glodoc_rtc & "'") & ", "
  sSentencia = sSentencia & "Null, "
  sSentencia = sSentencia & "'" & gsMesAct & "', "
  sSentencia = sSentencia & "'" & sDiaroContable & "', '" & sComprobante & "', Null, "
  ' operacion gravada
  nImporte_mn = CDec(oRecordset!impogr)
  nImporte_me = Round(nImporte_mn / CDec(nTipoCambio), 2)
  If (oRecordset!tpomon = TPOMON_EXT) Then
    nImporte_me = CDec(oRecordset!impogr)
    nImporte_mn = Round(nImporte_me * CDec(nTipoCambio), 2)
  End If
  sSentencia = sSentencia & CDec(nImporte_mn) & ", "
  sSentencia = sSentencia & CDec(nImporte_me) & ", "
  sSentencia = sSentencia & IIf(CDec(nImporte_mn) <> 0, INDMASCTA_MAS, INDMASCTA_INI) & ", "
  ' operacion exportacion
  nImporte_mn = CDec(oRecordset!impexp)
  nImporte_me = Round(nImporte_mn / CDec(nTipoCambio), 2)
  If (oRecordset!tpomon = TPOMON_EXT) Then
    nImporte_me = CDec(oRecordset!impexp)
    nImporte_mn = Round(nImporte_me * CDec(nTipoCambio), 2)
  End If
  sSentencia = sSentencia & CDec(nImporte_mn) & ", "
  sSentencia = sSentencia & CDec(nImporte_me) & ", "
  sSentencia = sSentencia & IIf(CDec(nImporte_mn) <> 0, INDMASCTA_MAS, INDMASCTA_INI) & ", "
  ' operacion exonerada
  nImporte_mn = CDec(oRecordset!impexo)
  nImporte_me = Round(nImporte_mn / CDec(nTipoCambio), 2)
  If (oRecordset!tpomon = TPOMON_EXT) Then
    nImporte_me = CDec(oRecordset!impexo)
    nImporte_mn = Round(nImporte_me * CDec(nTipoCambio), 2)
  End If
  sSentencia = sSentencia & CDec(nImporte_mn) & ", "
  sSentencia = sSentencia & CDec(nImporte_me) & ", "
  sSentencia = sSentencia & IIf(CDec(nImporte_mn) <> 0, INDMASCTA_MAS, INDMASCTA_INI) & ", "
  ' igv
  nImporte_mn = CDec(oRecordset!impigv)
  nImporte_me = Round(nImporte_mn / CDec(nTipoCambio), 2)
  If (oRecordset!tpomon = TPOMON_EXT) Then
    nImporte_me = CDec(oRecordset!impigv)
    nImporte_mn = Round(nImporte_me * CDec(nTipoCambio), 2)
  End If
  sSentencia = sSentencia & CDec(nImporte_mn) & ", "
  sSentencia = sSentencia & CDec(nImporte_me) & ", "
  sSentencia = sSentencia & IIf(CDec(nImporte_mn) <> 0, INDMASCTA_MAS, INDMASCTA_INI) & ", "
  ' isc
  nImporte_mn = CDec(oRecordset!impisc)
  nImporte_me = Round(nImporte_mn / CDec(nTipoCambio), 2)
  If (oRecordset!tpomon = TPOMON_EXT) Then
    nImporte_me = CDec(oRecordset!impisc)
    nImporte_mn = Round(nImporte_me * CDec(nTipoCambio), 2)
  End If
  sSentencia = sSentencia & CDec(nImporte_mn) & ", "
  sSentencia = sSentencia & CDec(nImporte_me) & ", "
  sSentencia = sSentencia & IIf(CDec(nImporte_mn) <> 0, INDMASCTA_MAS, INDMASCTA_INI) & ", "
  ' otros impuestos
  nImporte_mn = CDec(oRecordset!impoim)
  nImporte_me = Round(nImporte_mn / CDec(nTipoCambio), 2)
  If (oRecordset!tpomon = TPOMON_EXT) Then
    nImporte_me = CDec(oRecordset!impoim)
    nImporte_mn = Round(nImporte_me * CDec(nTipoCambio), 2)
  End If
  sSentencia = sSentencia & CDec(nImporte_mn) & ", "
  sSentencia = sSentencia & CDec(nImporte_me) & ", "
  sSentencia = sSentencia & IIf(CDec(nImporte_mn) <> 0, INDMASCTA_MAS, INDMASCTA_INI) & ", "
  ' importe total
  nImporte_mn = CDec(oRecordset!imptot)
  nImporte_me = Round(nImporte_mn / CDec(nTipoCambio), 2)
  If (oRecordset!tpomon = TPOMON_EXT) Then
    nImporte_me = CDec(oRecordset!imptot)
    nImporte_mn = Round(nImporte_me * CDec(nTipoCambio), 2)
  End If
  sSentencia = sSentencia & CDec(nImporte_mn) & ", "
  sSentencia = sSentencia & CDec(nImporte_me) & ", "
  sSentencia = sSentencia & IIf(CDec(nImporte_mn) <> 0, INDMASCTA_MAS, INDMASCTA_INI) & ", "
  ' indicadores  generación
  sSentencia = sSentencia & INDPREGEN_INA & ", " & INDPREGEN_INA & ", " & INDANU_FAL & ", "
  sSentencia = sSentencia & TipoImpuesto.Ninguno & ", " & CategoriaDocumento.Ninguno & ", " & INDPREGEN_INA & ", "
  sSentencia = sSentencia & "Null, Null, Null, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(sFechaProceso, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(sFechaProceso, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & "0, 0, "
  sSentencia = sSentencia & INDPREGEN_INA & ", " & TipoImpuesto.Ninguno & ", "
  sSentencia = sSentencia & "Null, Null, Null, Null, Null, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(sFechaProceso, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
  sSentencia = sSentencia & "0, 0, 0, 0, "
  sSentencia = sSentencia & "'" & gsAbvUsr & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
  sSentencia = sSentencia & "Null, Null)"
  uocnnMain.Execute sSentencia, nNumRegistros
  
  ' Información cuentas ventas
  sExpresion = " " & UCase(MonthName(gsMesAct)) & " " & gsAnoAct
  sSentencia = "INSERT INTO covtadoccta (codemp, pdoano, codtdc, serdoc, nrodoc, tpocnc, orden, codcta, glodet0, glodet1, glodet0x, glodet1x, codruc, impcta_mn, impcta_me, usrcre, fyhcre, usrmdf, fyhmdf) "
  sSentencia = sSentencia & "SELECT cta.codemp, '" & gsAnoAct & "' AS pdoano, vta.codtdc, vta.serdoc, '" & sDocumentoVenta & "' AS nrodoc, "
  sSentencia = sSentencia & "cta.tpocnc, cta.orden, cta.codcta, cta.glodet0, "
  sSentencia = sSentencia & "(CASE WHEN RIGHT(RTRIM(IFNULL(cta.glodet0, '')), 6)='MES DE' THEN CONCAT(RTRIM(IFNULL(cta.glodet1, '')), '" & sExpresion & "') ELSE cta.glodet1 END) AS glodet1, cta.glodet0x, "
  sSentencia = sSentencia & "(CASE WHEN RIGHT(RTRIM(IFNULL(cta.glodet0x, '')), 8)='MONTH OF' THEN CONCAT(RTRIM(IFNULL(cta.glodet1x, '')), '" & sExpresion & "') ELSE cta.glodet1x END) AS glodet1x, cta.codruc, "
  sSentencia = sSentencia & "ROUND(ROUND("
  sSentencia = sSentencia & "(CASE cta.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo "
  sSentencia = sSentencia & "WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim ELSE vta.imptot END) * "
  sSentencia = sSentencia & "(CASE WHEN vta.tpomon='" & TPOMON_EXT & "' THEN " & CDec(nTipoCambio) & " ELSE 1 END), 2) * ROUND(cta.porimpcta/100, 4), 2) AS impcta_mn, "
  sSentencia = sSentencia & "ROUND(ROUND("
  sSentencia = sSentencia & "(CASE cta.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo "
  sSentencia = sSentencia & "WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim ELSE vta.imptot END) / "
  sSentencia = sSentencia & "(CASE WHEN vta.tpomon='" & TPOMON_NAC & "' THEN " & CDec(nTipoCambio) & " ELSE 1 END), 2) * ROUND(cta.porimpcta/100, 4), 2) AS impcta_me, "
  sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ") AS fyhcre, "
  sSentencia = sSentencia & "Null AS usrmdf, Null AS fyhmdf "
  sSentencia = sSentencia & "FROM cortevtacta cta "
  sSentencia = sSentencia & "INNER JOIN cortevta vta ON vta.codemp=cta.codemp AND vta.sernegocio=cta.sernegocio AND vta.nronegocio=cta.nronegocio "
  sSentencia = sSentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND cta.sernegocio='" & oRecordset!sernegocio & "' "
  sSentencia = sSentencia & "AND cta.nronegocio='" & oRecordset!nronegocio & "' "
  sSentencia = sSentencia & "AND cta.porimpcta<>0.00 "
  sSentencia = sSentencia & "ORDER BY cta.tpocnc, cta.orden"
  uocnnMain.Execute sSentencia, nNumRegistros

  ' Información centro costo ventas
  sSentencia = "INSERT INTO covtadoccco (codemp, pdoano, codtdc, serdoc, nrodoc, tpocnc, orden, codcta, codcco, impcco_mn, impcco_me, usrcre, fyhcre, usrmdf, fyhmdf) "
  sSentencia = sSentencia & "SELECT cco.codemp, '" & gsAnoAct & "' AS pdoano, vta.codtdc, vta.serdoc, '" & sDocumentoVenta & "' AS nrodoc, "
  sSentencia = sSentencia & "cco.tpocnc, cco.orden, cco.codcta, cco.codcco, "
  sSentencia = sSentencia & "ROUND((ROUND("
  sSentencia = sSentencia & "(CASE cta.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo "
  sSentencia = sSentencia & "WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim ELSE vta.imptot END) * "
  sSentencia = sSentencia & "(CASE WHEN vta.tpomon='" & TPOMON_EXT & "' THEN " & CDec(nTipoCambio) & " ELSE 1 END), 2) * ROUND((cta.porimpcta*cco.porimpcco)/100, 4))/100, 2) AS impcco_mn, "
  sSentencia = sSentencia & "ROUND((ROUND("
  sSentencia = sSentencia & "(CASE cta.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo "
  sSentencia = sSentencia & "WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim ELSE vta.imptot END) / "
  sSentencia = sSentencia & "(CASE WHEN vta.tpomon='" & TPOMON_NAC & "' THEN " & CDec(nTipoCambio) & " ELSE 1 END), 2) * ROUND((cta.porimpcta*cco.porimpcco)/100, 4))/100, 2) AS impcco_me, "
  sSentencia = sSentencia & "'" & gsAbvUsr & "' AS usrcre, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ") AS fyhcre, "
  sSentencia = sSentencia & "Null AS usrmdf, Null AS fyhmdf "
  sSentencia = sSentencia & "FROM cortevtacco cco "
  sSentencia = sSentencia & "INNER JOIN cortevtacta cta ON cta.codemp=cco.codemp AND cta.sernegocio=cco.sernegocio AND cta.nronegocio=cco.nronegocio AND cta.tpocnc=cco.tpocnc AND cta.orden=cco.orden AND cta.codcta=cco.codcta "
  sSentencia = sSentencia & "INNER JOIN cortevta vta ON vta.codemp=cta.codemp AND vta.sernegocio=cta.sernegocio AND vta.nronegocio=cta.nronegocio "
  sSentencia = sSentencia & "WHERE cco.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND cco.sernegocio='" & oRecordset!sernegocio & "' "
  sSentencia = sSentencia & "AND cco.nronegocio='" & oRecordset!nronegocio & "' "
  sSentencia = sSentencia & "AND cco.porimpcco<>0.00 "
  sSentencia = sSentencia & "ORDER BY cco.tpocnc, cco.orden, cco.codcta"
  uocnnMain.Execute sSentencia, nNumRegistros
  
  uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  
  Exit Sub
ErrGrabar:
   gpErrores
   
   uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.

End Sub
Private Function ValidoPresupuesto(ByVal nTipoCambio As Double) As Boolean
  Dim sMensaje As String
  Dim nImporteVta As Double, nImportePre As Double
  Dim porstPspVta As ADODB.Recordset
  
  Set porstPspVta = New ADODB.Recordset
  With porstPspVta
    .ActiveConnection = uocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  
  ValidoPresupuesto = True
  With porstPspVta
    .Source = "SELECT psp.codcta, psp.codcco, cta.tpomon, "
    .Source = .Source & "ROUND(AVG(psp.impmn_" & gsMesAct & "), 2) AS imporpre_mn, "
    .Source = .Source & "ROUND(AVG(psp.impme_" & gsMesAct & "), 2) AS imporpre_me, "
    .Source = .Source & "ROUND(SUM(((CASE vcta.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim "
    .Source = .Source & "ELSE vta.imptot END)*(vcta.porimpcta/100))*(CASE WHEN vta.tpomon='" & TPOMON_EXT & "' THEN " & nTipoCambio & " ELSE 1 END)), 2) AS imporvta_mn, "
    .Source = .Source & "ROUND(SUM(((CASE vcta.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim "
    .Source = .Source & "ELSE vta.imptot END)*(vcta.porimpcta/100))/(CASE WHEN vta.tpomon='" & TPOMON_NAC & "' THEN " & nTipoCambio & " ELSE 1 END)), 2) AS imporvta_me "
    .Source = .Source & "FROM copsp psp "
    .Source = .Source & "INNER JOIN cocta cta ON cta.codemp=psp.codemp AND cta.pdoano=psp.pdoano AND cta.codcta=psp.codcta AND cta.indcco='" & INDCCO_INA & "' "
    .Source = .Source & "INNER JOIN cortevtacta vcta ON vcta.codemp=psp.codemp AND vcta.codcta=psp.codcta "
    .Source = .Source & "INNER JOIN cortevta vta ON vta.codemp=vcta.codemp AND vta.sernegocio=vcta.sernegocio AND vta.nronegocio=vcta.nronegocio AND vta.indestado='" & ESTCTA_ACT & "' "
    .Source = .Source & "WHERE psp.codemp ='" & gsCodEmp & "' "
    .Source = .Source & "AND psp.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND IFNULL(psp.codcco, '')='' "
    .Source = .Source & "GROUP BY psp.codcta, psp.codcco, cta.tpomon "
    .Source = .Source & "HAVING (CASE WHEN cta.tpomon='" & TPOMON_NAC & "' THEN imporvta_mn>imporpre_mn ELSE imporvta_me>imporpre_me END) "
    .Source = .Source & "UNION ALL "
    .Source = .Source & "SELECT psp.codcta, psp.codcco, cta.tpomon, "
    .Source = .Source & "ROUND(AVG(psp.impmn_" & gsMesAct & "), 2) AS imporpre_mn, "
    .Source = .Source & "ROUND(AVG(psp.impme_" & gsMesAct & "), 2) AS imporpre_me, "
    .Source = .Source & "ROUND(SUM(((CASE vcco.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim "
    .Source = .Source & "ELSE vta.imptot END)*((vcta.porimpcta * vcco.porimpcco)/10000))*(CASE WHEN vta.tpomon='" & TPOMON_EXT & "' THEN " & nTipoCambio & " ELSE 1 END)), 2) AS imporvta_mn, "
    .Source = .Source & "ROUND(SUM(((CASE vcco.tpocnc WHEN '1' THEN vta.impogr WHEN '2' THEN vta.impexp WHEN '3' THEN vta.impexo WHEN '4' THEN vta.impigv WHEN '5' THEN vta.impisc WHEN '6' THEN vta.impoim "
    .Source = .Source & "ELSE vta.imptot END)*((vcta.porimpcta * vcco.porimpcco)/10000))/(CASE WHEN vta.tpomon='" & TPOMON_NAC & "' THEN " & nTipoCambio & " ELSE 1 END)), 2) AS imporvta_me "
    .Source = .Source & "FROM copsp psp "
    .Source = .Source & "INNER JOIN cocta cta ON cta.codemp=psp.codemp AND cta.pdoano=psp.pdoano AND cta.codcta=psp.codcta AND cta.indcco='" & INDCCO_ACT & "' "
    .Source = .Source & "INNER JOIN cortevtacco vcco ON vcco.codemp=psp.codemp AND vcco.codcta=psp.codcta AND vcco.codcco=psp.codcco "
    .Source = .Source & "INNER JOIN cortevtacta vcta ON vcta.codemp=vcco.codemp AND vcta.sernegocio=vcco.sernegocio AND vcta.nronegocio=vcco.nronegocio AND vcta.tpocnc=vcco.tpocnc AND vcta.orden=vcco.orden AND vcta.codcta=vcco.codcta "
    .Source = .Source & "INNER JOIN cortevta vta ON vta.codemp=vcta.codemp AND vta.sernegocio=vcta.sernegocio AND vta.nronegocio=vcta.nronegocio AND vta.indestado='" & ESTCTA_ACT & "' "
    .Source = .Source & "WHERE psp.codemp ='" & gsCodEmp & "' "
    .Source = .Source & "AND psp.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND IFNULL(psp.codcco, '')<>'' "
    .Source = .Source & "GROUP BY psp.codcta, psp.codcco, cta.tpomon "
    .Source = .Source & "HAVING (CASE WHEN cta.tpomon='" & TPOMON_NAC & "' THEN imporvta_mn>imporpre_mn ELSE imporvta_me>imporpre_me END) "
    .Source = .Source & "ORDER BY codcta, codcco"
    .Open
  End With
  
  If porstPspVta.RecordCount > 0 Then
    MsgBox Choose(gsIdioma, "Importe de Ventas es Mayor al Importe del Presupuesto", "Sales Amount is Greater than the Amount of the Budget"), vbCritical
    ValidoPresupuesto = False
  End If
  If Not ValidoPresupuesto Then GoTo ErrorVerifica
  
ErrorVerifica:
  Set porstPspVta = Nothing

End Function
Private Function VerificaCtaCCo(ByVal oRecordset As ADODB.Recordset) As Boolean
  Dim nContador As Integer, nIndCco As Byte
  Dim sRegistro As String
  Dim nImporteRteVta As Double
  Dim nImporteCta_mn As Double, nImporteCta_me As Double
  Dim nImporteCCo_mn As Double
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
    nImporteRteVta = CDec(oRecordset(sRegistro))
    ' Verifico los importes de las cuentas
    If CDec(oRecordset(sRegistro)) <> 0 Then
      With porstCprCta
        .Source = "SELECT vta.orden, vta.codcta, vta.porimpcta, cta.indcco "
        .Source = .Source & "FROM CoRteVtaCta vta "
        .Source = .Source & "INNER JOIN cocta cta ON vta.codemp=cta.codemp AND vta.pdoano=cta.pdoano AND vta.codcta=cta.codcta "
        .Source = .Source & "WHERE vta.codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND vta.sernegocio='" & oRecordset!sernegocio & "' "
        .Source = .Source & "AND vta.nronegocio='" & oRecordset!nronegocio & "' "
        .Source = .Source & "AND vta.tpocnc='" & nContador & "' "
        .Source = .Source & "ORDER BY orden"
        .Open
      End With
      ' Valido los centro de costos
      If porstCprCta.RecordCount > 0 Then
        nImporteCta_mn = 0
        While Not porstCprCta.EOF
          nImporteCta_mn = nImporteCta_mn + CDec(porstCprCta!porimpcta)
          nIndCco = porstCprCta!indcco
          nImporteCCo_mn = 0
          If nIndCco = INDCCO_ACT Then
            With porstCprCco
              .Source = "SELECT vta.codcta, ROUND(SUM(vta.porimpcco), 2) AS porimpcco "
              .Source = .Source & "FROM CoRteVtaCCo vta "
              .Source = .Source & "INNER JOIN cocco cco ON vta.codemp=cco.codemp AND vta.pdoano=cco.pdoano AND vta.codcco=cco.codcco "
              .Source = .Source & "WHERE vta.codemp='" & gsCodEmp & "' "
              .Source = .Source & "AND vta.sernegocio='" & oRecordset!sernegocio & "' "
              .Source = .Source & "AND vta.nronegocio='" & oRecordset!nronegocio & "' "
              .Source = .Source & "AND vta.tpocnc='" & nContador & "' "
              .Source = .Source & "AND vta.orden='" & porstCprCta!orden & "' "
              .Source = .Source & "AND vta.codcta='" & porstCprCta!CodCta & "' "
              .Source = .Source & "GROUP BY vta.codcta"
              .Open
            End With
            ' Valido los centro de costos
            If porstCprCco.RecordCount > 0 Then
              nImporteCCo_mn = CDec(porstCprCco!porimpcco)
            End If
            porstCprCco.Close
            VerificaCtaCCo = (CDec(porstCprCta!porimpcta) = (porstCprCta!porimpcta * (nImporteCCo_mn / 100)))
            If Not VerificaCtaCCo Then GoTo ErrorVerifica
          End If
          porstCprCta.MoveNext
        Wend
      End If
      porstCprCta.Close
    End If
    
    ' Verifico información de rubro
    nImporteCta_me = Round(nImporteCta_mn / 100, 2)
    nImporteCta_mn = Round(nImporteRteVta * nImporteCta_me, 2)
    VerificaCtaCCo = (nImporteRteVta = nImporteCta_mn)
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

