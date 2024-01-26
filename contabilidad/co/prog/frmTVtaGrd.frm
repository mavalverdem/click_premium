VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTVtaGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   6390
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
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
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8475
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   8475
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
         Picture         =   "frmTVtaGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmTVtaGrd.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
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
            TabIndex        =   5
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
         Left            =   7750
         Picture         =   "frmTVtaGrd.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmTVtaGrd.frx":0396
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
         Picture         =   "frmTVtaGrd.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   2880
         Picture         =   "frmTVtaGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTVtaGrd"
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
Public uorstTGTCb As ADODB.Recordset
Public uorstCOCta As ADODB.Recordset
Public uorstCOCCo As ADODB.Recordset
Public uorstCODro As ADODB.Recordset
Public uorstCOVtaDocCta As ADODB.Recordset
Public uorstCOVtaDocCCo As ADODB.Recordset
Public uorstCOCpbCab As ADODB.Recordset
Public uorstCOCpbDet As ADODB.Recordset
Public uorstTemporal As ADODB.Recordset
Private porstCancel As ADODB.Recordset
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
']
']

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
   psConnStrgSele_Grd = "SELECT COVtaDoc.CodDro, COVtaDoc.NroCpb, c.AbvTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc, COVtaDoc.CodAux, b.RazAux, " _
                  & "  COVtaDoc.FeEDoc, COVtaDoc.TpoMon," _
                  & "  If(COVtaDoc.TpoMon='" & TPOMON_NAC & "',COVtaDoc.ImpTot_MN,COVtaDoc.ImpTot_ME) as cImpTot, " _
                  & "  If(COVtaDoc.IndGen,'x',' ') as cIndGen," _
                  & "  b.CodAux, c.CodTDc, Concat(COVtaDoc.CodTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc) as cLlave " _
                  & "FROM (COVtaDoc" _
                  & "  LEFT JOIN TGAux b ON COVtaDoc.CodAux = b.CodAux)" _
                  & "  LEFT JOIN TGTDc c ON COVtaDoc.CodTDc = c.CodTDc " _
                  & "WHERE COVtaDoc.MesPvs='" & gsMesAct & "' "
'                  & "  COVtaDoc.CodTDc, COVtaDoc.FehOpe, COVtaDoc.SerDoc_Fin, COVtaDoc.NroDoc_Fin," _
'                  & "  COVtaDoc.FeVDoc, COVtaDoc.ImpTCb, COVtaDoc.PctIGV, COVtaDoc.PctISC," _
'                  & "  COVtaDoc.RefDoc, COVtaDoc.GloDoc," _
'                  & "  COVtaDoc.MesPvs," _
'                  & "  COVtaDoc.ImpOGr_MN, COVtaDoc.ImpExp_MN, COVtaDoc.ImpExo_MN," _
'                  & "  COVtaDoc.ImpIGV_MN, COVtaDoc.ImpISC_MN, COVtaDoc.ImpOIm_MN, COVtaDoc.ImpTot_MN," _
'                  & "  COVtaDoc.ImpOGr_ME, COVtaDoc.ImpExp_ME, COVtaDoc.ImpExo_ME," _
'                  & "  COVtaDoc.ImpIGV_ME, COVtaDoc.ImpISC_ME, COVtaDoc.ImpOIm_ME, COVtaDoc.ImpTot_ME," _
'                  & "  COVtaDoc.IndCta_OGr, COVtaDoc.IndCta_Exp, COVtaDoc.IndCta_Exo," _
'                  & "  COVtaDoc.IndCta_IGV, COVtaDoc.IndCta_ISC, COVtaDoc.IndCta_OIm, COVtaDoc.IndCta_Tot," _
'                  & "  COVtaDoc.IndPreGen, COVtaDoc.IndGen, COVtaDoc.IndAnu," _
'                  & "  Concat(COVtaDoc.CodTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc) as cLlave," _
'                  & "  COVtaDoc.UsrCre, COVtaDoc.FyHCre, COVtaDoc.UsrMdf, COVtaDoc.FyHMdf "
   psConnStrgSele = "SELECT COVtaDoc.CodDro, COVtaDoc.NroCpb, COVtaDoc.SerDoc, COVtaDoc.NroDoc, COVtaDoc.CodAux, " _
                  & "  COVtaDoc.FeEDoc, COVtaDoc.TpoMon," _
                  & "  If(COVtaDoc.TpoMon='" & TPOMON_NAC & "',COVtaDoc.ImpTot_MN,COVtaDoc.ImpTot_ME) as cImpTot, " _
                  & "  COVtaDoc.CodTDc, COVtaDoc.FehOpe, COVtaDoc.SerDoc_Fin, COVtaDoc.NroDoc_Fin," _
                  & "  COVtaDoc.FeVDoc, COVtaDoc.ImpTCb, COVtaDoc.PctIGV, COVtaDoc.PctISC," _
                  & "  COVtaDoc.RefDoc, COVtaDoc.GloDoc," _
                  & "  COVtaDoc.MesPvs," _
                  & "  COVtaDoc.ImpOGr_MN, COVtaDoc.ImpExp_MN, COVtaDoc.ImpExo_MN," _
                  & "  COVtaDoc.ImpIGV_MN, COVtaDoc.ImpISC_MN, COVtaDoc.ImpOIm_MN, COVtaDoc.ImpTot_MN," _
                  & "  COVtaDoc.ImpOGr_ME, COVtaDoc.ImpExp_ME, COVtaDoc.ImpExo_ME," _
                  & "  COVtaDoc.ImpIGV_ME, COVtaDoc.ImpISC_ME, COVtaDoc.ImpOIm_ME, COVtaDoc.ImpTot_ME," _
                  & "  COVtaDoc.IndCta_OGr, COVtaDoc.IndCta_Exp, COVtaDoc.IndCta_Exo," _
                  & "  COVtaDoc.IndCta_IGV, COVtaDoc.IndCta_ISC, COVtaDoc.IndCta_OIm, COVtaDoc.IndCta_Tot," _
                  & "  COVtaDoc.IndPreGen, COVtaDoc.IndGen, COVtaDoc.IndAnu," _
                  & "  Concat(COVtaDoc.CodTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc) as cLlave," _
                  & "  COVtaDoc.UsrCre, COVtaDoc.FyHCre, COVtaDoc.UsrMdf, COVtaDoc.FyHMdf " _
                  & "FROM COVtaDoc " _
                  & "WHERE COVtaDoc.MesPvs='" & gsMesAct & "' "
   psConnStrgOrde = "ORDER BY COVtaDoc.CodTDc, COVtaDoc.SerDoc, COVtaDoc.NroDoc"
''   usConnStrgSele_COVtaDocCta = "SELECT COVtaDocCta.CodCta, b.DetCta, COVtaDocCta.ImpCta_MN, COVtaDocCta.ImpCta_ME," _
''                              & "  COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc," _
''                              & "  COVtaDocCta.TpoCnc," _
''                              & "  Concat(COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc, COVtaDocCta.TpoCnc) as cLlave," _
''                              & "  Concat(COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc, COVtaDocCta.TpoCnc, COVtaDocCta.CodCta) AS cLlave2," _
''                              & "  COVtaDocCta.UsrCre, COVtaDocCta.FyHCre, COVtaDocCta.UsrMdf, COVtaDocCta.FyHMdf " _
''                              & "FROM COVtaDocCta" _
''                              & "  LEFT JOIN COCta b ON COVtaDocCta.CodCta=b.CodCta "
   usConnStrgSele_COVtaDocCta = "SELECT COVtaDocCta.CodCta, COVtaDocCta.ImpCta_MN, COVtaDocCta.ImpCta_ME," _
                              & "  COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc," _
                              & "  COVtaDocCta.TpoCnc," _
                              & "  Concat(COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc, COVtaDocCta.TpoCnc) as cLlave," _
                              & "  Concat(COVtaDocCta.CodTDc, COVtaDocCta.SerDoc, COVtaDocCta.NroDoc, COVtaDocCta.TpoCnc, COVtaDocCta.CodCta) AS cLlave2," _
                              & "  COVtaDocCta.UsrCre, COVtaDocCta.FyHCre, COVtaDocCta.UsrMdf, COVtaDocCta.FyHMdf " _
                              & "FROM COVtaDocCta "
   usConnStrgWher_COVtaDocCta = ""
   usConnStrgOrde_COVtaDocCta = "ORDER BY 1, 2" ' DESC"
'   usConnStrgSele_COVtaDocCCo = "SELECT COVtaDocCCo.CodCCo, b.DetCCo, COVtaDocCCo.ImpCCo_MN, COVtaDocCCo.ImpCCo_ME, " _
'                              & "  COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc," _
'                              & "  COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta," _
'                              & "  Concat(COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta) as cLlave," _
'                              & "  Concat(COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc, COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta) as cLlave1," _
'                              & "  Concat(COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc, COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta, COVtaDocCCo.CodCCo) as cLlave2," _
'                              & "  COVtaDocCCo.UsrCre, COVtaDocCCo.FyHCre, COVtaDocCCo.UsrMdf, COVtaDocCCo.FyHMdf " _
'                              & "FROM COVtaDocCCo" _
'                              & "  LEFT JOIN COCCo b ON COVtaDocCCo.CodCCo=b.CodCCo "
   usConnStrgSele_COVtaDocCCo = "SELECT COVtaDocCCo.CodCCo, COVtaDocCCo.ImpCCo_MN, COVtaDocCCo.ImpCCo_ME," _
                              & "  COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta," _
                              & "  COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc," _
                              & "  Concat(COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta) as cLlave," _
                              & "  Concat(COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc, COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta) as cLlave1," _
                              & "  Concat(COVtaDocCCo.CodTDc, COVtaDocCCo.SerDoc, COVtaDocCCo.NroDoc, COVtaDocCCo.TpoCnc, COVtaDocCCo.CodCta, COVtaDocCCo.CodCCo) as cLlave2," _
                              & "  COVtaDocCCo.UsrCre, COVtaDocCCo.FyHCre, COVtaDocCCo.UsrMdf, COVtaDocCCo.FyHMdf " _
                              & "FROM COVtaDocCCo "
   usConnStrgWher_COVtaDocCCo = ""
   usConnStrgOrde_COVtaDocCCo = "ORDER BY 4, 5, 1"
   usConnStrgSele_COCpbDet = "SELECT COCpbDet.CodCta, COCpbDet.CodAux, COCpbDet.CodCCo, COCpbDet.GloIte," _
                           & "  If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "',COCpbDet.ImpMN,0) as cImpMN_Deb," _
                           & "  If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "',0,COCpbDet.ImpMN) as cImpMN_Hab," _
                           & "  If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "',COCpbDet.ImpME,0) as cImpME_Deb," _
                           & "  If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "',0,COCpbDet.ImpME) as cImpME_Hab," _
                           & "  If(COCpbDet.TpoGnr=" & TPOGNR_DST & ",'*','') as cTpoGnr, " _
                           & "  COCpbDet.MesPvs," _
                           & "  COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte, COCpbDet.FehOpe," _
                           & "  COCpbDet.CodTDc, COCpbDet.SerDoc, COCpbDet.NroDoc, COCpbDet.FeEDoc," _
                           & "  COCpbDet.FeVDoc, COCpbDet.FeRDoc, COCpbDet.RefDoc, COCpbDet.TpoMon," _
                           & "  COCpbDet.ImpTCb, COCpbDet.ImpMN, COCpbDet.ImpME, COCpbDet.TpoCtb," _
                           & "  COCpbDet.TpoGnr, " _
                           & "  Concat(COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte) as cLlave," _
                           & "  COCpbDet.UsrCre, COCpbDet.FyHCre " _
                           & "FROM COCpbDet "
   usConnStrgWher_COCpbDet = "WHERE COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='' AND COCpbDet.NroCpb='' "
   usConnStrgOrde_COCpbDet = "ORDER BY COCpbDet.NroIte"

   Set uocnnMain = New ADODB.Connection
   Set uocnnNoGrabable = New ADODB.Connection
   Set uorstMain = New ADODB.Recordset
   Set uorstMain_Grd = New ADODB.Recordset
   Set uorstTGAux = New ADODB.Recordset
   Set uorstTGTDc = New ADODB.Recordset
   Set uorstTGTCb = New ADODB.Recordset
   Set uorstCOCta = New ADODB.Recordset
   Set uorstCOCCo = New ADODB.Recordset
   Set uorstCODro = New ADODB.Recordset
   Set uorstCOVtaDocCta = New ADODB.Recordset
   Set uorstCOVtaDocCCo = New ADODB.Recordset
   Set uorstCOCpbCab = New ADODB.Recordset
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
      .Source = "SELECT a.CodTDc, a.DetTDc, a.SgnTDc " _
              & "FROM TGTDc a"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstTGTCb
      .ActiveConnection = uocnnMain
      .Source = "SELECT a.FehTCb, a.ImpTCb_Cpr, a.ImpTCb_Vta " _
              & "FROM TGTCb a"
'              & "WHERE Month(a.FehTCb)=" & Val(gsMesAct) & " AND Year(a.FehTCb)=" & Val(gsAnoAct)
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstCOCta
      .ActiveConnection = frmTVtaGrd.uocnnMain
      .Source = "SELECT a.CodCta, a.DetCta, a.TpoTCb, a.IndDoc, a.IndCCo " _
              & "FROM COCta a " _
              & "WHERE a.TpoCta=" & TPOCTA_TRA & " AND a.EstCta='" & ESTCTA_ACT & "'"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCOCCo
      .ActiveConnection = uocnnMain
      .Source = "SELECT a.CodCCo, a.DetCCo " _
              & "FROM COCCo a " _
              & "WHERE a.EstCCo='" & ESTCCO_ACT & "' AND Length(CodCCo) > 2"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCODro
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodDro, DetDro, Cpb" & gsMesAct & " " _
              & "FROM CODro " _
              & "WHERE Length(CodDro)=4"
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
      .Source = "SELECT CodDro, NroCpb, FehCpb, GloCpb, TpoGnr, IndNCu, MesPvs," _
              & "  Concat(CodDro, NroCpb) as cLlave," _
              & "  UsrCre, FyHCre " _
              & "FROM COCpbCab " _
              & "WHERE MesPvs='" & gsMesAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
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
      .Source = "SELECT a.CodAux, a.RazAux " _
              & "FROM TGAux a " _
              & "WHERE a.IndCli=1 AND a.EstAux='" & ESTAUX_ACT & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
']
   
   dgrMain.MarqueeStyle = dbgHighlightRow
   Set dgrMain.DataSource = uorstMain_Grd
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
   uorstCOCta.Close
   uorstCOCCo.Close
   uorstCODro.Close
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
   Set uorstCOCta = Nothing
   Set uorstCOCCo = Nothing
   Set uorstCODro = Nothing
   Set uorstCOVtaDocCta = Nothing
   Set uorstCOVtaDocCCo = Nothing
   Set uorstCOCpbCab = Nothing
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
   uorstMain.Find "cLlave='" & uorstMain_Grd!CodTDc & uorstMain_Grd!SerDoc & uorstMain_Grd!NroDoc & "'"
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
   
   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & "-" & Trim(dgrMain.Columns(2)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      With porstCancel
         .Source = "SELECT MesPvs, CodAux, CodTDc, SerDoc, NroDoc, TpoPvs " _
                 & "FROM COCpbDet " _
                 & "WHERE MesPvs='" & gsMesAct & "' AND CodAux='" & uorstMain_Grd!CodAux & "' AND CodTDc='" & uorstMain_Grd!CodTDc & "' AND SerDoc='" & uorstMain_Grd!SerDoc & "' AND NroDoc='" & uorstMain_Grd!NroDoc & "' And TpoPvs='" & TPOPVS_CAN & "'"
         .Open
         If porstCancel.RecordCount = 0 Then
            uorstMain.MoveFirst
            uorstMain.Find "cLlave = '" & uorstMain_Grd!CodTDc & uorstMain_Grd!SerDoc & uorstMain_Grd!NroDoc & "'"

            uocnnMain.BeginTrans       'INICIA TRANSACCION.
            uocnnMain.Execute "DELETE FROM COCpbCab WHERE MesPvs='" & gsMesAct & "' AND CodDro='" & Trim(dgrMain.Columns(0)) & "' And NroCpb='" & Trim(dgrMain.Columns(1)) & "' And TpoGnr='" & TPOGNR_VTA & "'"
            uorstMain.Properties("Unique Table").Value = "COVtaDoc"
            uorstMain.Delete
            uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.

           'Busca siguiente ítem.
            With uorstMain_Grd
               .MoveNext
               If .EOF Then .MoveLast
               dsLlaveSiguiente = !CodTDc & !SerDoc & !NroDoc
               .Requery
               If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
            End With
         Else
            MsgBox "Debe eliminar antes las Cancelaciones.", vbExclamation
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

Public Sub cmdImprimir_Click()
 '[Datos del formulario de impresión.  'Cambiar.
'   frmLCta.Caption = "Listado de " & Me.Caption
'   frmLCta.Show vbModal
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
            .Item(dnNum).Caption = "Diario"
            .Item(dnNum).Width = 500
         Case 1
            .Item(dnNum).Caption = "NºComp."
            .Item(dnNum).Width = 700
         Case 2
            .Item(dnNum).Caption = "TDc"
            .Item(dnNum).Width = 500
         Case 3
            .Item(dnNum).Caption = "Ser"
            .Item(dnNum).Width = 500
         Case 4
            .Item(dnNum).Caption = "Número"
            .Item(dnNum).Width = 1000
         Case 5
            .Item(dnNum).Caption = "Auxiliar"
            .Item(dnNum).Width = 1100
         Case 6
            .Item(dnNum).Caption = "Razón Social"
            .Item(dnNum).Width = 950
         Case 7
            .Item(dnNum).Caption = "F.Emisión"
            .Item(dnNum).Width = 1000
         Case 8
            .Item(dnNum).Caption = "M"
            .Item(dnNum).Width = 250
         Case 9
            .Item(dnNum).Caption = "Total"
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
   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property

