VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTHPrGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   1530
   ClientTop       =   1995
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
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
         Picture         =   "frmTHPrGrd.frx":0000
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
         Picture         =   "frmTHPrGrd.frx":0102
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
         Picture         =   "frmTHPrGrd.frx":024C
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
         Picture         =   "frmTHPrGrd.frx":0396
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
         Picture         =   "frmTHPrGrd.frx":0498
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
         Picture         =   "frmTHPrGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTHPrGrd"
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
Public uorstCOHPrDocCta As ADODB.Recordset
Public uorstCOHPrDocCCo As ADODB.Recordset
Public uorstCOCpbCab As ADODB.Recordset
Public uorstCOCpbDet As ADODB.Recordset
Public uorstTemporal As ADODB.Recordset
Private porstCancel As ADODB.Recordset
Public usConnStrgSele_COHPrDocCta As String, _
       usConnStrgWher_COHPrDocCta As String, _
       usConnStrgOrde_COHPrDocCta As String
Public usConnStrgSele_COHPrDocCCo As String, _
       usConnStrgWher_COHPrDocCCo As String, _
       usConnStrgOrde_COHPrDocCCo As String
Public usConnStrgSele_COCpbDet As String, _
       usConnStrgWher_COCpbDet As String, _
       usConnStrgOrde_COCpbDet As String

Public ubGrabaMas As Byte  '0:Nuevo documento 1:Cuenta grabado por cmdMas 2:Cuenta grabada directa.
'[Repetir en frmTHPr y frmTHPrMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
']

Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
   psConnStrgSele_Grd = "SELECT COHPrDoc.CodDro, COHPrDoc.NroCpb, COHPrDoc.CodAux, b.RazAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc, " _
                  & "  COHPrDoc.FeEDoc, COHPrDoc.TpoMon," _
                  & "  If(COHPrDoc.TpoMon='" & TPOMON_NAC & "',COHPrDoc.ImpBru_MN,COHPrDoc.ImpBru_ME) as cImpBru, " _
                  & "  If(COHPrDoc.IndGen,'x',' ') as cIndGen," _
                  & "  b.CodAux, Concat(COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc) as cLlave " _
                  & "FROM COHPrDoc" _
                  & "  LEFT JOIN TGAux b ON COHPrDoc.CodAux = b.CodAux " _
                  & "WHERE COHPrDoc.MesPvs='" & gsMesAct & "' "
   psConnStrgSele = "SELECT COHPrDoc.CodDro, COHPrDoc.NroCpb, COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc, " _
                  & "  COHPrDoc.FeEDoc, COHPrDoc.TpoMon," _
                  & "  If(COHPrDoc.TpoMon='" & TPOMON_NAC & "',COHPrDoc.ImpBru_MN,COHPrDoc.ImpBru_ME) as cImpBru, " _
                  & "  COHPrDoc.FehOpe, COHPrDoc.ImpTCb, COHPrDoc.PctIR4, COHPrDoc.PctIES," _
                  & "  COHPrDoc.RefDoc, COHPrDoc.GloDoc," _
                  & "  COHPrDoc.MesPvs," _
                  & "  COHPrDoc.ImpBru_MN, COHPrDoc.ImpIR4_MN, COHPrDoc.ImpIES_MN," _
                  & "  COHPrDoc.ImpORt_MN, COHPrDoc.ImpNet_MN," _
                  & "  COHPrDoc.ImpBru_ME, COHPrDoc.ImpIR4_ME, COHPrDoc.ImpIES_ME," _
                  & "  COHPrDoc.ImpORt_ME, COHPrDoc.ImpNet_ME," _
                  & "  COHPrDoc.IndAfeIR4, COHPrDoc.IndAfeIES, COHPrDoc.IndAfeORt," _
                  & "  COHPrDoc.IndCta_Bru, COHPrDoc.IndCta_IR4, COHPrDoc.IndCta_IES," _
                  & "  COHPrDoc.IndCta_ORt, COHPrDoc.IndCta_Net," _
                  & "  COHPrDoc.IndPreGen, COHPrDoc.IndGen, COHPrDoc.IndAnu," _
                  & "  Concat(COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc) as cLlave," _
                  & "  COHPrDoc.UsrCre, COHPrDoc.FyHCre, COHPrDoc.UsrMdf, COHPrDoc.FyHMdf " _
                  & "FROM COHPrDoc " _
                  & "WHERE COHPrDoc.MesPvs='" & gsMesAct & "' "
   psConnStrgOrde = "ORDER BY COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc"
''   usConnStrgSele_COHPrDocCta = "SELECT COHPrDocCta.CodCta, b.DetCta, COHPrDocCta.ImpCta_MN, COHPrDocCta.ImpCta_ME," _
''                              & "  COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc," _
''                              & "  COHPrDocCta.TpoCnc," _
''                              & "  Concat(COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc, COHPrDocCta.TpoCnc) as cLlave," _
''                              & "  Concat(COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc, COHPrDocCta.TpoCnc, COHPrDocCta.CodCta) as cLlave2," _
''                              & "  COHPrDocCta.UsrCre, COHPrDocCta.FyHCre, COHPrDocCta.UsrMdf, COHPrDocCta.FyHMdf " _
''                              & "FROM COHPrDocCta" _
''                              & "  LEFT JOIN COCta b ON COHPrDocCta.CodCta=b.CodCta "
   usConnStrgSele_COHPrDocCta = "SELECT COHPrDocCta.CodCta, COHPrDocCta.ImpCta_MN, COHPrDocCta.ImpCta_ME," _
                              & "  COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc," _
                              & "  COHPrDocCta.TpoCnc," _
                              & "  Concat(COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc, COHPrDocCta.TpoCnc) as cLlave," _
                              & "  Concat(COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc, COHPrDocCta.TpoCnc, COHPrDocCta.CodCta) as cLlave2," _
                              & "  COHPrDocCta.UsrCre, COHPrDocCta.FyHCre, COHPrDocCta.UsrMdf, COHPrDocCta.FyHMdf " _
                              & "FROM COHPrDocCta "
   usConnStrgWher_COHPrDocCta = ""
   usConnStrgOrde_COHPrDocCta = "ORDER BY 1, 2" ' DESC"
'   usConnStrgSele_COHPrDocCCo = "SELECT COHPrDocCCo.CodCCo, b.DetCCo, COHPrDocCCo.ImpCCo_MN, COHPrDocCCo.ImpCCo_ME, " _
'                              & "  COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc," _
'                              & "  COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta," _
'                              & "  Concat(COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta) as cLlave," _
'                              & "  Concat(COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc, COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta) as cLlave1," _
'                              & "  Concat(COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc, COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta, COHPrDocCCo.CodCCo) as cLlave2," _
'                              & "  COHPrDocCCo.UsrCre, COHPrDocCCo.FyHCre, COHPrDocCCo.UsrMdf, COHPrDocCCo.FyHMdf " _
'                              & "FROM COHPrDocCCo" _
'                              & "  LEFT JOIN COCCo b ON COHPrDocCCo.CodCCo=b.CodCCo "
   usConnStrgSele_COHPrDocCCo = "SELECT COHPrDocCCo.CodCCo, COHPrDocCCo.ImpCCo_MN, COHPrDocCCo.ImpCCo_ME, " _
                              & "  COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta," _
                              & "  COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc," _
                              & "  Concat(COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta) as cLlave," _
                              & "  Concat(COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc, COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta) as cLlave1," _
                              & "  Concat(COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc, COHPrDocCCo.TpoCnc, COHPrDocCCo.CodCta, COHPrDocCCo.CodCCo) as cLlave2," _
                              & "  COHPrDocCCo.UsrCre, COHPrDocCCo.FyHCre, COHPrDocCCo.UsrMdf, COHPrDocCCo.FyHMdf " _
                              & "FROM COHPrDocCCo "
   usConnStrgWher_COHPrDocCCo = ""
   usConnStrgOrde_COHPrDocCCo = "ORDER BY 4, 5, 1"
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
   Set uorstCOHPrDocCta = New ADODB.Recordset
   Set uorstCOHPrDocCCo = New ADODB.Recordset
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
      .Properties("Unique Table").Value = "COHPrDoc"
   End With
   With uorstMain
      .ActiveConnection = uocnnMain
      .Source = psConnStrgSele & psConnStrgOrde
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic 'adLockReadOnly
      .Open
      .Properties("Unique Table").Value = "COHPrDoc"
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
      .ActiveConnection = frmTHPrGrd.uocnnMain
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
              & "WHERE a.EstCCo='" & ESTCCO_ACT & "' AND Length(a.CodCCo) > 2"
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
   With uorstCOHPrDocCta
      .ActiveConnection = uocnnMain
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
   End With
   With uorstCOHPrDocCCo
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
              & "WHERE a.IndPrv=1 AND a.EstAux='" & ESTAUX_ACT & "'"
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
'[ARREGLAR. Genera demora al salir de la opci�n.
   If uorstCOHPrDocCta.State = adStateOpen Then uorstCOHPrDocCta.Close
   If uorstCOHPrDocCCo.State = adStateOpen Then uorstCOHPrDocCCo.Close
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
   Set uorstCOHPrDocCta = Nothing
   Set uorstCOHPrDocCCo = Nothing
   Set uorstCOCpbCab = Nothing
   Set uorstCOCpbDet = Nothing
   Set uorstMain_Grd = Nothing
   Set uorstMain = Nothing
   Set uocnnMain = Nothing
End Sub

Private Sub cmdNuevo_Click()
 '[Propio del formulario.
   'Verificaci�n de Mes Cerrado.
   If gbCieHpr Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   ubGrabaMas = INDMASCTA_INI
   uocnnMain.BeginTrans
 ']
   gpTUg_Nuevo Me, frmTHPr             'Cambiar Formulario de Datos.
   cmdRefrescar_Click
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err

   'Verificaci�n de existencia de �temes.
   If uorstMain_Grd.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If

 '[Propio del formulario.
   ubGrabaMas = INDMASCTA_CTA
 ']

 '[B�squeda del �tem.
uorstMain.Requery
   uorstMain.MoveFirst
   uorstMain.Find "cLlave='" & uorstMain_Grd!CodAux & uorstMain_Grd!SerDoc & uorstMain_Grd!NroDoc & "'"
 ']

   With frmTHPr                        'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitaci�n de Llaves.       'Cambiar.
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
   
   'Verificaci�n de Mes Cerrado.
   If gbCieHpr Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   'Verificaci�n de existencia de �temes.
   If uorstMain_Grd.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If
   
   'Mensaje de verificaci�n            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & "-" & Trim(dgrMain.Columns(2)) & "-" & Trim(dgrMain.Columns(3)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      With porstCancel
         .Source = "SELECT MesPvs, CodAux, CodTDc, SerDoc, NroDoc, TpoPvs " _
                 & "FROM COCpbDet " _
                 & "WHERE MesPvs='" & gsMesAct & "' AND CodAux='" & uorstMain_Grd!CodAux & "' AND CodTDc='" & CODTDC_HPR & "' AND SerDoc='" & uorstMain_Grd!SerDoc & "' AND NroDoc='" & uorstMain_Grd!NroDoc & "' And TpoPvs='" & TPOPVS_CAN & "'"
         .Open
         If porstCancel.RecordCount = 0 Then
            uorstMain.MoveFirst
            uorstMain.Find "cLlave = '" & uorstMain_Grd!CodAux & uorstMain_Grd!SerDoc & uorstMain_Grd!NroDoc & "'"

            uocnnMain.BeginTrans       'INICIA TRANSACCION.
            uocnnMain.Execute "DELETE FROM COCpbCab WHERE MesPvs='" & gsMesAct & "' And CodDro='" & Trim(dgrMain.Columns(0)) & "' And NroCpb='" & Trim(dgrMain.Columns(1)) & "' And TpoGnr='" & TPOGNR_HPR & "'"
            uorstMain.Properties("Unique Table").Value = "COHPrDoc"
            uorstMain.Delete
            uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.

           'Busca siguiente �tem.
            With uorstMain_Grd
               .MoveNext
               If .EOF Then .MoveLast
               dsLlaveSiguiente = !CodAux & !SerDoc & !NroDoc
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
            formularios que lo usan (formularios de registro �nico).
''   gpTUg_Refrescar Me
   uorstMain_Grd.Requery
   upDatosGrid
   
   dgrMain.SetFocus
']ARREGLAR.
End Sub

Public Sub cmdImprimir_Click()
 '[Datos del formulario de impresi�n.  'Cambiar.
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
      .Properties("Unique Table").Value = "COHPrDoc"
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
   
'[ARREGLAR: B�squeda con distintos tipos de columna.
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
            .Item(dnNum).Caption = "N�Comp."
            .Item(dnNum).Width = 700
         Case 2
            .Item(dnNum).Caption = "Auxiliar"
            .Item(dnNum).Width = 1100
         Case 3
            .Item(dnNum).Caption = "Raz�n Social"
            .Item(dnNum).Width = 1450
         Case 4
            .Item(dnNum).Caption = "Ser"
            .Item(dnNum).Width = 500
         Case 5
            .Item(dnNum).Caption = "N�mero"
            .Item(dnNum).Width = 1000
         Case 6
            .Item(dnNum).Caption = "F.Emisi�n"
            .Item(dnNum).Width = 1000
         Case 7
            .Item(dnNum).Caption = "M"
            .Item(dnNum).Width = 250
         Case 8
            .Item(dnNum).Caption = "Importe Bruto"
            .Item(dnNum).Width = 1200
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
            .Item(dnNum).Alignment = dbgRight
         Case 9
            .Item(dnNum).Caption = "G"
            .Item(dnNum).Width = 230
            .Item(dnNum).Alignment = dbgCenter
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

'[C�digo propio del formulario.

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



