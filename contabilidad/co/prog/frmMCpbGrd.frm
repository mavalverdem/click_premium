VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMCpbGrd 
   Caption         =   "[Entidad - Tipo Asiento]"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   1260
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8475
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8475
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8475
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
         Left            =   3600
         Picture         =   "frmMCpbGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmMCpbGrd.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   720
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
         Left            =   2880
         Picture         =   "frmMCpbGrd.frx":0414
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmMCpbGrd.frx":0516
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmMCpbGrd.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   720
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
         Height          =   560
         Left            =   7750
         Picture         =   "frmMCpbGrd.frx":071A
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
         Left            =   4320
         TabIndex        =   3
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   200
            Width           =   2415
         End
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
         Picture         =   "frmMCpbGrd.frx":0864
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8475
      _ExtentX        =   14949
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
End
Attribute VB_Name = "frmMCpbGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'sirve para darle el mes 00
'por defecto todo hira ahi
Public rcMesAct


Public uocnnMain As ADODB.Connection
Public uorstMain_0 As ADODB.Recordset, _
       uorstMain_1 As ADODB.Recordset, _
       uorstUltiItem As ADODB.Recordset
Public usConnStrgSele_0 As String, _
       usConnStrgOrde_0 As String, _
       usConnStrgSele_1 As String, _
       usConnStrgWher_1 As String, _
       usConnStrgOrde_1 As String
'       usCOnnStrgWher
Private pnColumnaOrd As Integer

'[Propio del formulario.
Public uorstCODro As ADODB.Recordset, _
       uorstTGTCb As ADODB.Recordset, _
       uorstCoCta As ADODB.Recordset, _
       uorstCoCCo As ADODB.Recordset, _
       uorstTGAux As ADODB.Recordset, _
       uorstTGTDc As ADODB.Recordset, _
       uorstcomacpbdet As ADODB.Recordset, _
       uorstCOTCbMes As ADODB.Recordset, _
       uorstcomacpbdetRP As ADODB.Recordset
Public uorstCOFjo As ADODB.Recordset, _
       uorstCOFjoDet As ADODB.Recordset

       
'       uorstTGArt As ADODB.Recordset, _
'       uorstTGSvc As ADODB.Recordset
']


Private Sub cmdGenera_Click()
   With frmPAsFmtTran
      .Show vbModal
   End With
   'Cambiar Formulario de Datos.

End Sub

'''Private Sub cmdProcesar_Click()
'''   With frmPAsFmtTran
'''      .Show vbModal
'''   End With
'''   'Cambiar Formulario de Datos.
'''
'''End Sub

Private Sub Form_Load()
'rcMesAct = "00"
'2014-03-31 error de validacion fecha cance prov
'rcMesAct = "01"
rcMesAct = "12"

 '[Recordsets                          'Cambiar.
   usConnStrgSele_0 = "SELECT CodDro, NroCpb, FehCpb, "
   usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "GloCpb, ", "GloCpbx, ")
   usConnStrgSele_0 = usConnStrgSele_0 & "(CASE TpoGnr WHEN " & TPOGNR_DRO & " THEN '" & TPOGNR_DRO_TXT & "' WHEN " & TPOGNR_CPR & " THEN '" & TPOGNR_CPR_TXT & "' WHEN " & TPOGNR_VTA & " THEN '" & TPOGNR_VTA_TXT & "' WHEN " & TPOGNR_HPR & " THEN '" & TPOGNR_HPR_TXT & "' WHEN " & TPOGNR_DST & " THEN '" & TPOGNR_DST_TXT & "' WHEN " & TPOGNR_DCA & " THEN '" & TPOGNR_DCA_TXT & "' WHEN " & TPOGNR_APE & " THEN '" & TPOGNR_APE_TXT & "' WHEN " & TPOGNR_CIE & " THEN '" & TPOGNR_CIE_TXT & "' WHEN " & TPOGNR_DRP & " THEN '" & TPOGNR_DRP_TXT & "' ELSE '" & TPOGNR_BAN_TXT & "' END) AS ccTpoGnr, "
   usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "GloCpbx, ", "GloCpb, ")
   usConnStrgSele_0 = usConnStrgSele_0 & "TpoGnr, MesPvs, IndNCu, IndAnu, UsrCre, FyHCre, UsrMdf, FyHMdf, codemp, pdoano, "
   usConnStrgSele_0 = usConnStrgSele_0 & IIf(ps_Plataforma = pSrvMySql, "Concat(CodDro, NroCpb)", "(CodDro+NroCpb)") & " AS cLlave "
   usConnStrgSele_0 = usConnStrgSele_0 & "FROM comacpbcab "
   usConnStrgSele_0 = usConnStrgSele_0 & "WHERE codemp='" & gsCodEmp & "' "
   usConnStrgSele_0 = usConnStrgSele_0 & "AND pdoano='" & gsAnoAct & "' "
   usConnStrgSele_0 = usConnStrgSele_0 & "AND MesPvs='" & rcMesAct & "' "
   usConnStrgOrde_0 = "ORDER BY CodDro, NroCpb"
   
   usConnStrgSele_1 = "SELECT comacpbdet.NroIte, comacpbdet.CodCta, comacpbdet.CodCCo, comacpbdet.CodAux, TGTDc.AbvTDc , comacpbdet.SerDoc, comacpbdet.NroDoc, "
   usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "comacpbdet.GloIte, ", "comacpbdet.GloItex, ")
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE comacpbdet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN comacpbdet.ImpMN ELSE 0 END) AS cImpMN_Deb, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE comacpbdet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN 0 ELSE comacpbdet.ImpMN END) AS cImpMN_Hab, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE comacpbdet.TpoGnr WHEN " & TPOGNR_DRO & " THEN '" & TPOGNR_DRO_TXT & "' WHEN " & TPOGNR_CPR & " THEN '" & TPOGNR_CPR_TXT & "' WHEN " & TPOGNR_VTA & " THEN '" & TPOGNR_VTA_TXT & "' WHEN " & TPOGNR_HPR & " THEN '" & TPOGNR_HPR_TXT & "' WHEN " & TPOGNR_DST & " THEN '" & TPOGNR_DST_TXT & "' WHEN " & TPOGNR_DCA & " THEN '" & TPOGNR_DCA_TXT & "' WHEN " & TPOGNR_APE & " THEN '" & TPOGNR_APE_TXT & "' WHEN " & TPOGNR_CIE & " THEN '" & TPOGNR_CIE_TXT & "' WHEN " & TPOGNR_DRP & " THEN '" & TPOGNR_DRP_TXT & "' ELSE '" & TPOGNR_BAN_TXT & "' END) AS ccTpoGnr, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE comacpbdet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN comacpbdet.ImpME ELSE 0 END) AS cImpME_Deb, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE comacpbdet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN 0 ELSE comacpbdet.ImpME END) AS cImpME_Hab, "
   usConnStrgSele_1 = usConnStrgSele_1 & "comacpbdet.BlqIte, comacpbdet.TpoMon, comacpbdet.ImpTcb, comacpbdet.TpoTCb, comacpbdet.TpoGnr, comacpbdet.CodTDc, "
   usConnStrgSele_1 = usConnStrgSele_1 & "comacpbdet.RefDoc, comacpbdet.CodDro, comacpbdet.NroCpb, comacpbdet.TpoCtb, comacpbdet.TpoPvs, comacpbdet.MesPvs, "
   usConnStrgSele_1 = usConnStrgSele_1 & "comacpbdet.FehOpe, comacpbdet.FeEDoc, comacpbdet.FeVDoc, comacpbdet.FeRDoc, comacpbdet.ImpMN, comacpbdet.ImpME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "comacpbdet.IndFjo_Det, comacpbdet.IndGnr_RP, comacpbdet.UsrCre, comacpbdet.FyHCre, comacpbdet.UsrMdf, comacpbdet.FyHMdf, "
   usConnStrgSele_1 = usConnStrgSele_1 & "comacpbdet.codemp, comacpbdet.pdoano, comacpbdet.pdocpr, "
   usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "comacpbdet.GloItex, ", "comacpbdet.GloIte, ")
   usConnStrgSele_1 = usConnStrgSele_1 & IIf(ps_Plataforma = pSrvMySql, "CONCAT(comacpbdet.CodDro, comacpbdet.NroCpb, comacpbdet.NroIte)", "(comacpbdet.CodDro+comacpbdet.NroCpb+RTrim(comacpbdet.NroIte))") & " AS cLlave "
   usConnStrgSele_1 = usConnStrgSele_1 & "FROM (comacpbdet "
   usConnStrgSele_1 = usConnStrgSele_1 & "LEFT JOIN TGTDc AS TGTDc ON comacpbdet.codemp=TGTDc.codemp AND comacpbdet.CodTDc=TGTDc.CodTDc) "
   usConnStrgWher_1 = "WHERE comacpbdet.codemp='" & gsCodEmp & "' "
   usConnStrgWher_1 = usConnStrgWher_1 & "AND comacpbdet.pdoano='" & gsAnoAct & "' "
   usConnStrgWher_1 = usConnStrgWher_1 & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(comacpbdet.CodDro, comacpbdet.NroCpb)", "(comacpbdet.CodDro+comacpbdet.NroCpb)") & "=' ' "
   usConnStrgOrde_1 = "ORDER BY comacpbdet.NroIte, comacpbdet.BlqIte"
   
   Set uocnnMain = New ADODB.Connection
   Set uorstMain_0 = New ADODB.Recordset
   Set uorstMain_1 = New ADODB.Recordset
   Set uorstUltiItem = New ADODB.Recordset
   Set uorstCODro = New ADODB.Recordset
   Set uorstTGTCb = New ADODB.Recordset
   Set uorstCoCta = New ADODB.Recordset
   Set uorstCoCCo = New ADODB.Recordset
   Set uorstTGAux = New ADODB.Recordset
   Set uorstTGTDc = New ADODB.Recordset
   Set uorstcomacpbdet = New ADODB.Recordset
   Set uorstCOTCbMes = New ADODB.Recordset
   Set uorstcomacpbdetRP = New ADODB.Recordset
   Set uorstCOFjo = New ADODB.Recordset
   Set uorstCOFjoDet = New ADODB.Recordset
   With uocnnMain
      .CursorLocation = adUseClient
'      .ConnectionString = CONNSTRG & gsRutBDS & gsNomBDS
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With uorstMain_0
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_0 & usConnStrgOrde_0
'     .CursorLocation = adUseClient   'Es el Default.
        
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open (usConnStrgSele_0 & usConnStrgOrde_0)
'      .Properties("Unique Table").Value = "VTFacCab"
      .Properties("Unique Table").Value = "comacpbcab"
   End With
   With uorstMain_1
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_1 & usConnStrgWher_1 & usConnStrgOrde_1
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "comacpbdet"
   End With
   With uorstUltiItem
      .ActiveConnection = uocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
'      .Open
   End With
   With uorstCODro
     .ActiveConnection = uocnnMain
     .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro, codemp, Cpb" & rcMesAct & ", "
     .Source = .Source & "codemp, pdoano "
     .Source = .Source & "FROM codro "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
''     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstTGTCb
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.FehTCb, a.ImpTCb_Cpr, a.ImpTCb_Vta "
     .Source = .Source & "FROM TGTCb a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCoCta
    .ActiveConnection = uocnnMain
    .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta, "
    .Source = .Source & "TpoTCb, TpoAnl, IndAjd, IndCCo, IndDoc, IndFjo, CodCCo_Def, "
    .Source = .Source & "CodCta_AjD_Deb, CodCta_AjD_Hab, CodCCo_AjD_Deb, CodCCo_AjD_Hab "
    .Source = .Source & "FROM COCta "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND TpoCta=" & TPOCTA_TRA & " "
    .Source = .Source & "AND EstCta='" & ESTCTA_ACT & "'"
    '     .CursorLocation = adUseClient   'Es el Default.
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
   With uorstTGAux
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.CodAux, a.RazAux "
     .Source = .Source & "FROM TGAux a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND a.EstAux='" & ESTAUX_ACT & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
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
   With uorstcomacpbdet
      .ActiveConnection = uocnnMain
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
   End With
   With uorstCOTCbMes
      .ActiveConnection = uocnnMain
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
   End With
   With uorstcomacpbdetRP
     .ActiveConnection = uocnnMain
     .Source = "SELECT MesPvs, CodDro, NroCpb, NroIte, CodAux, CodCta, "
     .Source = .Source & "CodTDc, SerDoc, NroDoc, ImpMN, ImpME, "
     .Source = .Source & "CodTDc_RtcPcp, SerDoc_RtcPcp, NroDoc_RtcPcp, "
     .Source = .Source & "feEDoc_RtcPcp, ImpMN_RtcPcp, ImpME_RtcPcp, IndRtcPcp, "
     .Source = .Source & "UsrCre, FyHCre, UsrMdf, FyHMdf, codemp, pdoano, "
     .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(MesPvs, CodDro, NroCpb, NroIte)", "(MesPvs+CodDro+NroCpb+RTrim(NroIte))") & " AS cLlave "
     .Source = .Source & "FROM comacpbdetRP "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND MesPvs='" & rcMesAct & "' "
     .Source = .Source & "ORDER BY NroIte, CodDro, NroCpb"
'     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
     .Properties("Unique Table").Value = "comacpbdetRP"
   End With
   With uorstCOFjo
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.CodFjo, " & Choose(gsIdioma, "a.DetFjo", "a.DetFjox") & " AS DetFjo "
     .Source = .Source & "FROM COFjo a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
     .Source = .Source & "AND a.pdoano='" & gsAnoAct & "'"
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodFjo)>2"
''     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstCOFjoDet
      .ActiveConnection = uocnnMain
      .Source = "SELECT MesPvs, CodDro, NroCpb, NroIte, NroOrd, CodCta, "
      .Source = .Source & "TpoCtb, ImpMN, ImpME, UsrCre, FyHCre, UsrMdf, FyHMdf, codemp, pdoano "
      .Source = .Source & "FROM comacpbdetFjo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "comacpbdetFjo"
   End With
 ']
   
  '[ Elimino y creo tabla temporal de detalle de flujo
  If ps_Plataforma = pSrvMySql Then
    uocnnMain.Execute "DROP TABLE IF EXISTS tmpcomacpbdetFjo"
    uocnnMain.Execute "CREATE TEMPORARY TABLE tmpcomacpbdetFjo SELECT * FROM comacpbdetFjo WHERE CodFjo='tmpflujo'"
  ElseIf ps_Plataforma = pSrvSql Then
    ' Activo detector de errores
    On Error Resume Next
    uocnnMain.Execute "DROP TABLE " & ps_Prefijo & "tmpcomacpbdetFjo"
    If Not (Err.Number = -2147217865 Or Err.Number = 0) Then
      MsgBox Err.Description, vbInformation
    End If
    On Error GoTo 0
    uocnnMain.Execute "SELECT * INTO #tmpcomacpbdetFjo FROM comacpbdetFjo WHERE CodFjo='tmpflujo'"
  End If
   ']
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
  ']
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain_0
End Sub

Private Sub Form_Activate()
   'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
   zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
   ppDatosGrid
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
'   uorstTGSvc.Close
'   uorstTGArt.Close
'   uorstTGCli.Close
''   uorstVTNroDoc.Close
''   uorstUltiItem.Close
'   uorstMain_1.Close
   uorstMain_0.Close
   uocnnMain.Close
'   Set uorstTGSvc = Nothing
'   Set uorstTGArt = Nothing
'   Set uorstTGCli = Nothing
'   Set uorstVTNroDoc = Nothing
'   Set uorstUltiItem = Nothing
'   Set uorstMain_1 = Nothing
   Set uorstCOTCbMes = Nothing
   Set uorstMain_0 = Nothing
   Set uorstCOFjoDet = Nothing
   Set uocnnMain = Nothing
   
End Sub

Public Sub cmdNuevo_Click()
   'Verificación de Mes Cerrado.
   If gbCieCpb Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   '[ No pertence al Formulario - Agregado por Angel
   With uorstMain_1
      .Close
      .Source = usConnStrgSele_1 & " Where comacpbdet.CodDro='    ' " & usConnStrgOrde_1
      .Open
      .Properties("Unique Table").Value = "comacpbdet"
   End With
'   gpTUg_Nuevo Me, frmTFacCab          'Cambiar Formulario de Datos.
   gpTUg_Nuevo Me, frmMCpbCab          'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err

   'Verificación de existencia de ítemes.
   If uorstMain_0.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If

   With frmMCpbCab                     'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
      .txtLlave(0).Enabled = False
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
   Dim dsCriterio As String
   
   'Verificación de Mes Cerrado.
   If gbCieCpb Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   'Verificación de existencia de ítemes.
   If uorstMain_0.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

   If frmMCpbGrd.uorstMain_0!tpognr <> TPOGNR_DRO Then
      MsgBox Choose(gsIdioma, "No se Puede Eliminar este Comprobante", "This Voucher can not be eliminated"), vbInformation
      Exit Sub
   End If

   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      uocnnMain.BeginTrans
      uorstMain_0.Delete
      uocnnMain.CommitTrans
   End If
   dgrMain.SetFocus

   Exit Sub
Err:
   gpErrores
   
   uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
'   gpTUg_Refrescar Me
   frmMCpbGrd.uorstMain_0.Requery
   frmMCpbGrd.ppDatosGrid
   dgrMain.SetFocus
End Sub

Public Sub cmdImprimir_Click(Index As Integer)
 '[Datos del formulario de impresión.  'Cambiar.
   frmLCpbma.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
   frmLCpbma.Show vbModal
 ']
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   usConnStrgOrde_0 = "ORDER BY "
'   Select Case pnColumnaOrd            'Cambiar.
'   Case 1, 2, 3
'      usConnStrgOrde_0 = usConnStrgOrde_0 & "1, 2, 3"
'   Case Else
      usConnStrgOrde_0 = usConnStrgOrde_0 & pnColumnaOrd + 1
'   End Select
   With uorstMain_0
      .Close
      .Source = usConnStrgSele_0 & usConnStrgOrde_0
      .Open
   End With
   Set dgrMain.DataSource = uorstMain_0
   ppDatosGrid

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub dgrMain_KeyUp(KeyCode As Integer, Shift As Integer)
   If uorstMain_0.RecordCount = 0 Then Exit Sub

   Select Case KeyCode
   Case vbKeyHome
      uorstMain_0.MoveFirst
   Case vbKeyEnd
      uorstMain_0.MoveLast
   End Select
End Sub

Private Sub txtBuscar_Change()
   On Error GoTo Err
   
   Dim dsCriterio As String
   Dim dvRegistroActual As Variant
            
   With uorstMain_0
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
      uorstMain_0.Bookmark = dvRegistroActual
   Else
      gpErrores
   End If
End Sub

Function pfNumItemCpb(ByVal s_Ano As String, ByVal s_Mes As String, ByVal s_Diario As String, s_Comprobante As String) As Integer
  
  ' s_Ano             Año donde  se genera
  ' s_Mes             Mes donde  se genera
  ' s_Diario          Codigo de diario para generar numero
  ' s_Comprobante     Numero de comprobante para generar item
    
  Dim porstRetorno As ADODB.Recordset
  Dim s_Sentencia As String
  
  s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroIte), 0) AS nNumMaxItem "
  s_Sentencia = s_Sentencia & "FROM comacpbdet "
  s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND pdoano='" & s_Ano & "' "
  s_Sentencia = s_Sentencia & "AND MesPvs='" & s_Mes & "' "
  s_Sentencia = s_Sentencia & "AND CodDro='" & s_Diario & "' "
  s_Sentencia = s_Sentencia & "AND NroCpb='" & s_Comprobante & "'"
  Set porstRetorno = New ADODB.Recordset
  With porstRetorno
    .ActiveConnection = frmMCpbGrd.uocnnMain
'    .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = s_Sentencia
    .Open
  End With
  pfNumItemCpb = CInt(porstRetorno!nNumMaxItem) + 1
  porstRetorno.Close
  Set porstRetorno = Nothing

End Function

Public Sub ppDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = Choose(gsIdioma, "Diario", "Journal")
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("CodDro").DefinedSize + 2)
        Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "NºComp.", "NºVoucher")
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("NroCpb").DefinedSize + 2)
         Case 2
'            .Item(dnNum).Caption = "Cliente"
'            .Item(dnNum).Width = 3000
            .Item(dnNum).Caption = Choose(gsIdioma, "Fecha", "Date")
            .Item(dnNum).Width = 100 * (7 + 4)
         Case 3
'            .Item(dnNum).Caption = "Observación"
'            .Item(dnNum).Width = 2200
            .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("GloCpb").DefinedSize - 16)
         Case 4
'            .Item(dnNum).Caption = "Anulada"
'            .Item(dnNum).Width = 850
            .Item(dnNum).Caption = Choose(gsIdioma, "Tipo", "Type")
'            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("TpoGnr").DefinedSize + 5)
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("TpoGnr").DefinedSize + 8)
            .Item(dnNum).Alignment = dbgCenter
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'[Código propio del formulario.
'Function ffTipoCta(cTpoCtb As String, cTipo As String) As Double
'  If cTpoCtb = "D" Then
'     If cTipo = "N" Then
'        ppTipoCta = IIf(IsNull(!ImpMN), 0, !ImpMN)
'     Else
'        ppTipoCta = IIf(IsNull(!ImpME), 0, !ImpME)
'     End If
'  Else
'     If cTipo = "E" Then
'        ppTipoCta = IIf(IsNull(!ImpMN), 0, !ImpMN)
'     Else
'        ppTipoCta = IIf(IsNull(!ImpME), 0, !ImpME)
'     End If
'  End If
'End Function
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



