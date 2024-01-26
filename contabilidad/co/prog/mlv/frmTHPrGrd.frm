VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTHPrGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   1530
   ClientTop       =   1995
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9195
      _ExtentX        =   16219
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
      ScaleWidth      =   9195
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   9195
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
         Picture         =   "frmTHPrGrd.frx":0000
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
         Picture         =   "frmTHPrGrd.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmTHPrGrd.frx":0414
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   4370
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
         Left            =   8475
         Picture         =   "frmTHPrGrd.frx":055E
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
         Picture         =   "frmTHPrGrd.frx":06A8
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
         Picture         =   "frmTHPrGrd.frx":07AA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
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
         Left            =   2880
         Picture         =   "frmTHPrGrd.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
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
Public uorstCodOnpAfp As ADODB.Recordset '2014-08-01 RR.HH afecto afp/onp
Public uorstTGAux As ADODB.Recordset
Public uorstTGTDc As ADODB.Recordset
Public uorstTGTCb As ADODB.Recordset
Public uorstCoCta As ADODB.Recordset
Public uorstCoCCo As ADODB.Recordset
Public uorstCODro As ADODB.Recordset
Public uorstCoAsiTipo As ADODB.Recordset
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

Private Sub cmdGenera_Click()
  Dim s_Sentencia As String
  
  'Verificación de Mes Cerrado.
  If gbCieCpr Then
    MsgBox TEXT_9016, vbCritical
    Exit Sub
  End If
  ' Genero información
  With porstCancel
    .Source = "SELECT hpr.CodDro, hpr.NroCpb, hpr.CodAux, hpr.SerDoc, hpr.NroDoc, "
    .Source = .Source & "hpr.FeEDoc, hpr.TpoMon, hpr.pdocpr, "
    .Source = .Source & "hpr.FehOpe, hpr.ImpTCb, hpr.PctIR4, hpr.PctIES,"
    .Source = .Source & "hpr.RefDoc, hpr.GloDoc, hpr.GloDocx, "
    .Source = .Source & "hpr.MesPvs, hpr.codasi, "
    .Source = .Source & "hpr.ImpBru_MN, hpr.ImpIR4_MN, hpr.ImpIES_MN, "
    .Source = .Source & "hpr.ImpORt_MN, hpr.ImpNet_MN, "
    .Source = .Source & "hpr.ImpBru_ME, hpr.ImpIR4_ME, hpr.ImpIES_ME, "
    .Source = .Source & "hpr.ImpORt_ME, hpr.ImpNet_ME, "
    .Source = .Source & "hpr.IndAfeIR4, hpr.IndAfeIES, hpr.IndAfeORt, "
    .Source = .Source & "hpr.IndCta_Bru, hpr.IndCta_IR4, hpr.IndCta_IES, "
    .Source = .Source & "hpr.IndCta_ORt, hpr.IndCta_Net, "
    .Source = .Source & "hpr.IndPreGen, hpr.IndGen, hpr.IndAnu "
    .Source = .Source & "FROM CoHPrDoc hpr "
    .Source = .Source & "LEFT JOIN CoCpbCab cab ON hpr.codemp=cab.codemp AND hpr.pdoano=cab.pdoano AND hpr.MesPvs=cab.MesPvs AND hpr.CodDro=cab.CodDro AND hpr.NroCpb=cab.NroCpb "
    .Source = .Source & "WHERE hpr.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND hpr.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND hpr.MesPvs='" & gsMesAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(hpr.IndGen", "ISNULL(hpr.IndGen") & ", '0')='0' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(cab.CodDro, cab.NroCpb)", "ISNULL((cab.CodDro+cab.NroCpb)") & ", '')='' "
    .Source = .Source & "ORDER BY hpr.CodDro, hpr.NroDoc"
    .Open
  End With
  
  'Valido las Cuentas esten Correctas(llenas para todas los valores)
  If porstCancel.RecordCount > 0 Then
    While Not porstCancel.EOF
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

End Sub
Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
  psConnStrgSele_Grd = "SELECT COHPrDoc.CodDro, COHPrDoc.NroCpb, COHPrDoc.CodAux, b.RazAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "COHPrDoc.FeEDoc, COHPrDoc.TpoMon, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE COHPrDoc.TpoMon WHEN '" & TPOMON_NAC & "' THEN COHPrDoc.ImpBru_MN ELSE COHPrDoc.ImpBru_ME END) as cImpBru, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE COHPrDoc.IndGen WHEN -1 THEN 'x' ELSE ' ' END) as cIndGen, b.CodAux, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "Concat(COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc)", "(COHPrDoc.CodAux+COHPrDoc.SerDoc+COHPrDoc.NroDoc)") & " AS cLlave "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "FROM COHPrDoc "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGAux b ON COHPrDoc.codemp=b.codemp AND COHPrDoc.CodAux=b.CodAux "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "WHERE COHPrDoc.codemp='" & gsCodEmp & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND COHPrDoc.pdoano='" & gsAnoAct & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND COHPrDoc.MesPvs='" & gsMesAct & "' "
  
  psConnStrgSele = "SELECT COHPrDoc.CodDro, COHPrDoc.NroCpb, COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.FeEDoc, COHPrDoc.TpoMon, "
  psConnStrgSele = psConnStrgSele & "(CASE COHPrDoc.TpoMon WHEN '" & TPOMON_NAC & "' THEN COHPrDoc.ImpBru_MN ELSE COHPrDoc.ImpBru_ME END) AS cImpBru, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.FehOpe, COHPrDoc.ImpTCb, COHPrDoc.PctIR4, COHPrDoc.PctIES,"
  psConnStrgSele = psConnStrgSele & "COHPrDoc.RefDoc, COHPrDoc.GloDoc, COHPrDoc.GloDocx, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.MesPvs, COHPrDoc.codasi, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.ImpBru_MN, COHPrDoc.ImpIR4_MN, COHPrDoc.ImpIES_MN, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.ImpORt_MN, COHPrDoc.ImpNet_MN, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.ImpBru_ME, COHPrDoc.ImpIR4_ME, COHPrDoc.ImpIES_ME, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.ImpORt_ME, COHPrDoc.ImpNet_ME, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.IndAfeIR4, COHPrDoc.IndAfeIES, COHPrDoc.IndAfeORt, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.IndCta_Bru, COHPrDoc.IndCta_IR4, COHPrDoc.IndCta_IES, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.IndCta_ORt, COHPrDoc.IndCta_Net, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.IndPreGen, COHPrDoc.IndGen, COHPrDoc.IndAnu, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.UsrCre, COHPrDoc.FyHCre, COHPrDoc.UsrMdf, COHPrDoc.FyHMdf, "
  psConnStrgSele = psConnStrgSele & "COHPrDoc.codemp, COHPrDoc.pdoano, CoHprDoc.pdocpr, CoHprDoc.codcon, "
  psConnStrgSele = psConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "Concat(COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc)", "(COHPrDoc.CodAux+COHPrDoc.SerDoc+COHPrDoc.NroDoc)") & " AS cLlave "
  psConnStrgSele = psConnStrgSele & "FROM COHPrDoc "
  psConnStrgSele = psConnStrgSele & "WHERE COHPrDoc.codemp='" & gsCodEmp & "' "
  psConnStrgSele = psConnStrgSele & "AND COHPrDoc.pdoano='" & gsAnoAct & "' "
  psConnStrgSele = psConnStrgSele & "AND COHPrDoc.MesPvs='" & gsMesAct & "' "
  psConnStrgOrde = "ORDER BY COHPrDoc.CodAux, COHPrDoc.SerDoc, COHPrDoc.NroDoc"
  
  usConnStrgSele_COHPrDocCta = "SELECT COHPrDocCta.CodCta, COHPrDocCta.ImpCta_MN, COHPrDocCta.ImpCta_ME, "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & Choose(gsIdioma, "COHPrDocCta.GloDet, ", "COHPrDocCta.GloDetx, ") & "COHPrDocCta.CodRuc, "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & "COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc, "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & "COHPrDocCta.TpoCnc, COHPrDocCta.Orden, "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & IIf(ps_Plataforma = pSrvMySql, "Concat(COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc, COHPrDocCta.TpoCnc, COHPrDocCta.Orden)", "(COHPrDocCta.CodAux+COHPrDocCta.SerDoc+COHPrDocCta.NroDoc+RTrim(COHPrDocCta.TpoCnc)+COHPrDocCta.Orden)") & " AS cLlave, "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & IIf(ps_Plataforma = pSrvMySql, "Concat(COHPrDocCta.CodAux, COHPrDocCta.SerDoc, COHPrDocCta.NroDoc, COHPrDocCta.TpoCnc, COHPrDocCta.Orden, COHPrDocCta.CodCta)", "(COHPrDocCta.CodAux+COHPrDocCta.SerDoc+COHPrDocCta.NroDoc+RTrim(COHPrDocCta.TpoCnc)+COHPrDocCta.Orden+COHPrDocCta.CodCta)") & " AS cLlave2, "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & "COHPrDocCta.UsrCre, COHPrDocCta.FyHCre, COHPrDocCta.UsrMdf, COHPrDocCta.FyHMdf, "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & Choose(gsIdioma, "COHPrDocCta.GloDetx, ", "COHPrDocCta.GloDet, ")
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & "COHPrDocCta.codemp, COHPrDocCta.pdoano "
  usConnStrgSele_COHPrDocCta = usConnStrgSele_COHPrDocCta & "FROM COHPrDocCta "
  usConnStrgWher_COHPrDocCta = ""
  usConnStrgOrde_COHPrDocCta = "ORDER BY 9, 10, 1" ' DESC"
  
  usConnStrgSele_COHPrDocCCo = "SELECT COHPrDocCCo.CodCCo, COHPrDocCCo.ImpCCo_MN, COHPrDocCCo.ImpCCo_ME, "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & "COHPrDocCCo.TpoCnc, COHPrDocCCo.Orden, COHPrDocCCo.CodCta, "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & "COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc, "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & IIf(ps_Plataforma = pSrvMySql, "Concat(COHPrDocCCo.TpoCnc, COHPrDocCCo.Orden, COHPrDocCCo.CodCta)", "(RTrim(COHPrDocCCo.TpoCnc)+COHPrDocCCo.Orden+COHPrDocCCo.CodCta)") & " AS cLlave, "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & IIf(ps_Plataforma = pSrvMySql, "Concat(COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc, COHPrDocCCo.TpoCnc, COHPrDocCCo.Orden, COHPrDocCCo.CodCta)", "(COHPrDocCCo.CodAux+COHPrDocCCo.SerDoc+COHPrDocCCo.NroDoc+RTrim(COHPrDocCCo.TpoCnc)+COHPrDocCCo.Orden+COHPrDocCCo.CodCta)") & " AS cLlave1, "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & IIf(ps_Plataforma = pSrvMySql, "Concat(COHPrDocCCo.CodAux, COHPrDocCCo.SerDoc, COHPrDocCCo.NroDoc, COHPrDocCCo.TpoCnc, COHPrDocCCo.Orden, COHPrDocCCo.CodCta, COHPrDocCCo.CodCCo)", "(COHPrDocCCo.CodAux+COHPrDocCCo.SerDoc+COHPrDocCCo.NroDoc+RTrim(COHPrDocCCo.TpoCnc)+COHPrDocCCo.Orden+COHPrDocCCo.CodCta+COHPrDocCCo.CodCCo)") & " AS cLlave2, "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & "COHPrDocCCo.UsrCre, COHPrDocCCo.FyHCre, COHPrDocCCo.UsrMdf, COHPrDocCCo.FyHMdf, "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & "COHPrDocCCo.codemp, COHPrDocCCo.pdoano "
  usConnStrgSele_COHPrDocCCo = usConnStrgSele_COHPrDocCCo & "FROM COHPrDocCCo "
  usConnStrgWher_COHPrDocCCo = ""
  usConnStrgOrde_COHPrDocCCo = "ORDER BY 4, 5, 6, 1"
  
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
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.TpoGnr, CoCpbDet.pdocpr, CoCpbDet.codcon, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & IIf(ps_Plataforma = pSrvMySql, "Concat(COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte)", "(COCpbDet.CodDro+COCpbDet.NroCpb+COCpbDet.NroIte)") & " AS cLlave, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & Choose(gsIdioma, "COCpbDet.GloItex, ", "COCpbDet.GloIte, ")
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.UsrCre, COCpbDet.FyHCre, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.codemp, COCpbDet.pdoano "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "FROM COCpbDet "
  usConnStrgWher_COCpbDet = "WHERE COCpbDet.codemp='" & gsCodEmp & "' COCpbDet.pdoano='" & gsAnoAct & "' "
  usConnStrgWher_COCpbDet = usConnStrgWher_COCpbDet & "AND COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='' AND COCpbDet.NroCpb='' "
  usConnStrgOrde_COCpbDet = "ORDER BY COCpbDet.NroIte"
  
  Set uocnnMain = New ADODB.Connection
  Set uocnnNoGrabable = New ADODB.Connection
  Set uorstMain = New ADODB.Recordset
  Set uorstMain_Grd = New ADODB.Recordset
  Set uorstCodOnpAfp = New ADODB.Recordset '2014-08-01 RR.HH afecto afp/onp
  Set uorstTGAux = New ADODB.Recordset
  Set uorstTGTDc = New ADODB.Recordset
  Set uorstTGTCb = New ADODB.Recordset
  Set uorstCoCta = New ADODB.Recordset
  Set uorstCoCCo = New ADODB.Recordset
  Set uorstCODro = New ADODB.Recordset
  Set uorstCoAsiTipo = New ADODB.Recordset
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
     .ActiveConnection = frmTHPrGrd.uocnnMain
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
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
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
    .Source = .Source & "AND a.TpoAsi='" & TPOGNR_HPR & "'"
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
'ini 2014-08-01 RR.HH afecto afp/onp
  With uorstCodOnpAfp
     .ActiveConnection = uocnnNoGrabable
     .Source = "SELECT "
     .Source = .Source & "    a.codaux, a.codafp,a.flagcomision,b.factor1,b.factor2,"
     .Source = .Source & "    b.factor3 , b.factor4, b.topeseg,a.fecnacimiento "
     .Source = .Source & "FROM codonpafp a "
     .Source = .Source & "LEFT JOIN Coentidadpen b ON a.Codemp=b.Codemp AND a.codafp=b.codafp "
     .Source = .Source & "WHERE a.CodEmp='" & gsCodEmp & "'"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
'fin 2014-08-01 RR.HH afecto afp/onp
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
  
  dgrMain.MarqueeStyle = dbgHighlightRow
  Set dgrMain.DataSource = uorstMain_Grd
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  cmdVerificar.Caption = Choose(gsIdioma, "&Verificar", "&Check")
  cmdGenera.Caption = Choose(gsIdioma, "&Generar", "&Generate")
  CaptionBotones Me, False, False, True, True, True, True, False, False, False, False, False, False, True, aLabel
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
   uorstCodOnpAfp.Close '2014-08-01 RR.HH afecto afp/onp
   uorstTGAux.Close
   uorstTGTDc.Close
   uorstTGTCb.Close
   uorstCoCta.Close
   uorstCoCCo.Close
   uorstCODro.Close
   uorstCoAsiTipo.Close
'[ARREGLAR. Genera demora al salir de la opción.
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
   Set uorstCodOnpAfp = Nothing '2014-08-01 RR.HH afecto afp/onp
   Set uorstTGAux = Nothing
   Set uorstTGTDc = Nothing
   Set uorstTGTCb = Nothing
   Set uorstCoCta = Nothing
   Set uorstCoCCo = Nothing
   Set uorstCODro = Nothing
   Set uorstCoAsiTipo = Nothing
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
   'Verificación de Mes Cerrado.
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
   uorstMain.Find "cLlave='" & uorstMain_Grd!codaux & uorstMain_Grd!serdoc & uorstMain_Grd!nrodoc & "'"
 ']

   With frmTHPr                        'Cambiar Formulario de Datos.
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
   If gbCieHpr Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   'Verificación de existencia de ítemes.
   If uorstMain_Grd.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If
   
   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & "-" & Trim(dgrMain.Columns(2)) & "-" & Trim(dgrMain.Columns(3)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      With porstCancel
        .Source = "SELECT MesPvs, CodAux, CodTDc, SerDoc, NroDoc, TpoPvs "
        .Source = .Source & "FROM COCpbDet "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND MesPvs='" & gsMesAct & "' AND CodAux='" & uorstMain_Grd!codaux & "' "
        .Source = .Source & "AND CodTDc='" & CODTDC_HPR & "' AND SerDoc='" & uorstMain_Grd!serdoc & "'"
        .Source = .Source & "AND NroDoc='" & uorstMain_Grd!nrodoc & "' AND TpoPvs<>'" & TPOPVS_CAN & "'"
         .Open
         If porstCancel.RecordCount = 0 Then
            uorstMain.MoveFirst
            uorstMain.Find "cLlave = '" & uorstMain_Grd!codaux & uorstMain_Grd!serdoc & uorstMain_Grd!nrodoc & "'"

            uocnnMain.BeginTrans       'INICIA TRANSACCION.
            uocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND MesPvs='" & gsMesAct & "' And CodDro='" & Trim(dgrMain.Columns(0)) & "' And NroCpb='" & Trim(dgrMain.Columns(1)) & "' And TpoGnr='" & TPOGNR_HPR & "'"
            uorstMain.Properties("Unique Table").Value = "COHPrDoc"
            uorstMain.Delete
            uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.

           'Busca siguiente ítem.
            With uorstMain_Grd
               .MoveNext
               If .EOF Then .MoveLast
               dsLlaveSiguiente = !codaux & !serdoc & !nrodoc
               .Requery
               If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
            End With
         Else
            MsgBox Choose(gsIdioma, "Debe eliminar antes las Cancelaciones.", " The Cancelations must be eliminated before."), vbExclamation
         End If
        'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
        fEstMayUpd
        'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
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
  Dim s_Sentencia As String
  Dim porstMRp As New ADODB.Recordset
 
  s_Sentencia = "SELECT '" & CODTDC_HPR & "' AS CodTDc, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc, '-',NroDoc)", "(SerDoc+'-'+NroDoc)") & " AS cDocumento, FehOpe, FeEdoc, "
  s_Sentencia = s_Sentencia & Choose(gsIdioma, "GloDoc", "GloDocx") & " AS GloDoc, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cMoneda, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpBru_MN ELSE ImpBru_ME END) AS cImpBas, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpIR4_MN ELSE ImpIR4_ME END) AS cImpRenta, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpIES_MN ELSE ImpIES_ME END) AS cImpuesto, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpNet_MN ELSE ImpNet_ME END) AS cImpNeto, "
  s_Sentencia = s_Sentencia & "a.CodDro, a.NroCpb "
  s_Sentencia = s_Sentencia & "FROM CoHprDoc AS a "
  s_Sentencia = s_Sentencia & "LEFT JOIN CoCpbCab AS b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb "
  s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND a.pdoano='" & gsAnoAct & "' "
  s_Sentencia = s_Sentencia & "AND a.MesPvs='" & gsMesAct & "' "
  s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(b.CodDro, b.NroCpb)", "ISNULL((b.CodDro+b.NroCpb)") & ", '')='' "
  s_Sentencia = s_Sentencia & "ORDER BY a.CodDro, a.NroCpb, SerDoc, NroDoc"
  With porstMRp
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = s_Sentencia
    .Open
  End With
  
  gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "DOCUMENTOS DE HONORARIOS NO CONTABILIZADOS", "NOT COUNTED DOCUMENTS OF FEES"), Date, True, False, porstMRp
  With frmMain.rptMain
    '[Datos y parámetros del reporte.  'Cambiar.
    .ReportFileName = gsRutRpt & "rptLHprCpb.rpt"
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
            .Item(dnNum).Caption = Choose(gsIdioma, "NºComp.", "NºVoucher")
            .Item(dnNum).Width = 700
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 1100
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "Razón Social", "Firm Name")
            .Item(dnNum).Width = 2150
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "Ser", "Ser")
            .Item(dnNum).Width = 500
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Número", "Number")
            .Item(dnNum).Width = 1000
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "F.Emisión", "Issue Date")
            .Item(dnNum).Width = 1000
         Case 7
            .Item(dnNum).Caption = Choose(gsIdioma, "Mon", "Cur")
            .Item(dnNum).Width = 250
         Case 8
            .Item(dnNum).Caption = Choose(gsIdioma, "Importe Bruto", "Gross Amount")
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

'[Código propio del formulario.
Private Sub ppGeneraCpbCab(ByVal oRecordset As ADODB.Recordset)
  On Error GoTo ErrGrabar
  Dim nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nNumRegistros As Long
  Dim sSentencia As String, sComprobante As String
  Dim sCodAux As String, sTpoCtb As String
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
  sSentencia = sSentencia & "'" & TPOGNR_HPR & "', "
  sSentencia = sSentencia & "'" & INDNCU_FAL & "', "
  sSentencia = sSentencia & "'" & INDANU_FAL & "', "
  sSentencia = sSentencia & "'" & gsAbvUsr & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
  sSentencia = sSentencia & "Null, Null)"
  uocnnMain.Execute sSentencia, nNumRegistros
  
  ' Información detalle cuentas
  With porstCprCta
    .Source = "SELECT hpr.tpocnc, hpr.orden, hpr.codcta, cco.codcco, hpr.glodet, hpr.glodetx, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "IFNULL(hpr.impcta_mn", "ISNULL(hpr.impcta_mn") & ", 0) AS impcta_mn, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "IFNULL(hpr.impcta_me", "ISNULL(hpr.impcta_me") & ", 0) AS impcta_me, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "IFNULL(cco.impcco_mn", "ISNULL(cco.impcco_mn") & ", 0) AS impcco_mn, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "IFNULL(cco.impcco_me", "ISNULL(cco.impcco_me") & ", 0) AS impcco_me, "
    .Source = .Source & "hpr.codruc, cta.indcco, cta.inddoc, cta.inddoc, cta.tpotcb, tdc.sgntdc "
    .Source = .Source & "FROM cohprdoccta hpr "
    .Source = .Source & "INNER JOIN cocta cta ON hpr.codemp=cta.codemp AND hpr.pdoano=cta.pdoano AND hpr.codcta=cta.codcta "
    .Source = .Source & "LEFT JOIN tgtdc tdc ON hpr.codemp=tdc.codemp AND tdc.codtdc='" & CODTDC_HPR & "' "
    .Source = .Source & "LEFT JOIN cohprdoccco cco ON hpr.codemp=cco.codemp AND hpr.pdoano=cco.pdoano AND hpr.codaux=cco.codaux "
    .Source = .Source & "AND hpr.serdoc=cco.serdoc AND hpr.nrodoc=cco.nrodoc AND hpr.tpocnc=cco.tpocnc AND hpr.orden=cco.orden AND hpr.codcta=cco.codcta "
    .Source = .Source & "WHERE hpr.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND hpr.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND hpr.codaux='" & oRecordset!codaux & "' "
    .Source = .Source & "AND hpr.serdoc='" & oRecordset!serdoc & "' "
    .Source = .Source & "AND hpr.nrodoc='" & oRecordset!nrodoc & "' "
    .Source = .Source & "ORDER BY hpr.tpocnc, hpr.orden"
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
      If (nImporte_me > 0) And (nImporte_mn > 0) Then
        sTpoCtb = IIf(porstCprCta!tpocnc = TPOCNC_TOT_HPR, IIf(porstCprCta!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(porstCprCta!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        sTpoCtb = IIf(porstCprCta!tpocnc = TPOCNC_TOT_HPR, IIf(porstCprCta!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(porstCprCta!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
      nRegistro = nRegistro + 1
      ' Grabación de cabecera de comprobante .....................
      sSentencia = "INSERT INTO CoCpbDet(codemp, pdoano, coddro, nrocpb, nroite, mespvs, blqite, codtdc, fehope, codcta, codcco, codaux, serdoc, nrodoc, feedoc, fevdoc, "
      sSentencia = sSentencia & "FeRDoc, RefDoc, GloIte, GloItex, TpoCtb, TpoPvs, TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, tpognr, pdocpr, UsrCre, FyHCre, UsrMdf, FyHMdf) "
      sSentencia = sSentencia & "VALUES("
      sSentencia = sSentencia & "'" & gsCodEmp & "', "
      sSentencia = sSentencia & "'" & gsAnoAct & "', "
      sSentencia = sSentencia & "'" & oRecordset!coddro & "', "
      sSentencia = sSentencia & "'" & sComprobante & "', "
      sSentencia = sSentencia & "'" & nRegistro & "', "
      sSentencia = sSentencia & "'" & gsMesAct & "', "
      sSentencia = sSentencia & "'" & nRegistro & "', "
      sSentencia = sSentencia & "'" & CODTDC_HPR & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!fehope, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & "'" & porstCprCta!CodCta & "', "
      sSentencia = sSentencia & IIf(IsNull(porstCprCta!codcco), "Null", "'" & porstCprCta!codcco & "'") & ", "
      sSentencia = sSentencia & IIf(sCodAux = "", "Null", "'" & sCodAux & "'") & ", "
      sSentencia = sSentencia & "'" & oRecordset!serdoc & "', "
      sSentencia = sSentencia & "'" & oRecordset!nrodoc & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(IsNull(oRecordset!RefDoc), "Null", "'" & oRecordset!RefDoc & "'") & ", "
      sSentencia = sSentencia & IIf(IsNull(porstCprCta!glodet), "Null", "'" & porstCprCta!glodet & "'") & ", "
      sSentencia = sSentencia & IIf(IsNull(porstCprCta!glodetx), "Null", "'" & porstCprCta!glodetx & "'") & ", "
      sSentencia = sSentencia & "'" & sTpoCtb & "', "
      sSentencia = sSentencia & "'" & TPOPVS_PVS & "', "
      sSentencia = sSentencia & "'" & oRecordset!tpomon & "', "
      sSentencia = sSentencia & "'" & porstCprCta!TpoTcb & "', "
      sSentencia = sSentencia & CDec(oRecordset!ImpTCb) & ", "
      sSentencia = sSentencia & nImporte_mn & ", "
      sSentencia = sSentencia & nImporte_me & ", "
      sSentencia = sSentencia & "'" & TPOGNR_HPR & "', "
      sSentencia = sSentencia & IIf(IsNull(oRecordset!pdocpr), "Null", "'" & oRecordset!pdocpr & "'") & ", "
      sSentencia = sSentencia & "'" & gsAbvUsr & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
      sSentencia = sSentencia & "Null, Null)"
      uocnnMain.Execute sSentencia, nNumRegistros
      porstCprCta.MoveNext
    Wend
  End If
  porstCprCta.Close
  'Si no está marcado para generar, marca el documento como no generado.
  sSentencia = "UPDATE CoHprDoc SET indpregen=" & INDPREGEN_ACT & ", indgen=-1 "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND codaux='" & oRecordset!codaux & "' "
  sSentencia = sSentencia & "AND serdoc='" & oRecordset!serdoc & "' "
  sSentencia = sSentencia & "AND nrodoc='" & oRecordset!nrodoc & "'"
  uocnnMain.Execute sSentencia, nNumRegistros
  uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  
  Exit Sub
ErrGrabar:
   gpErrores
   
   uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.

End Sub
Private Function VerificaCtaCCo(ByVal oRecordset As ADODB.Recordset) As Boolean
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
  For nContador = 1 To 5
    sRegistro = Choose(nContador, "impbru", "impir4", "impies", "import", "impnet")
    sIndicado = "indcta_" & Right(sRegistro, 3)
    nImporteCpr_mn = CDec(oRecordset(sRegistro & "_mn"))
    nImporteCpr_me = CDec(oRecordset(sRegistro & "_me"))
    nImporteCta_mn = 0
    nImporteCta_me = 0
    ' Verifico los importes de las cuentas
    If oRecordset(sIndicado) <> 0 Then
      With porstCprCta
        .Source = "SELECT hpr.orden, hpr.codcta, hpr.impcta_mn, hpr.impcta_me, cta.indcco "
        .Source = .Source & "FROM cohprdoccta hpr "
        .Source = .Source & "INNER JOIN cocta cta ON hpr.codemp=cta.codemp AND hpr.pdoano=cta.pdoano AND hpr.codcta=cta.codcta "
        .Source = .Source & "WHERE hpr.codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND hpr.pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND hpr.codaux='" & oRecordset!codaux & "' "
        .Source = .Source & "AND hpr.serdoc='" & oRecordset!serdoc & "' "
        .Source = .Source & "AND hpr.nrodoc='" & oRecordset!nrodoc & "' "
        .Source = .Source & "AND hpr.tpocnc='" & nContador & "' "
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
              .Source = "SELECT hpr.codcta, ROUND(SUM(hpr.impcco_mn), 2) AS impcco_mn, ROUND(SUM(hpr.impcco_me), 2) AS impcco_me "
              .Source = .Source & "FROM cohprdoccco hpr "
              .Source = .Source & "INNER JOIN cocco cco ON hpr.codemp=cco.codemp AND hpr.pdoano=cco.pdoano AND hpr.codcco=cco.codcco "
              .Source = .Source & "WHERE hpr.codemp='" & gsCodEmp & "' "
              .Source = .Source & "AND hpr.pdoano='" & gsAnoAct & "' "
              .Source = .Source & "AND hpr.codaux='" & oRecordset!codaux & "' "
              .Source = .Source & "AND hpr.serdoc='" & oRecordset!serdoc & "' "
              .Source = .Source & "AND hpr.nrodoc='" & oRecordset!nrodoc & "' "
              .Source = .Source & "AND hpr.tpocnc='" & nContador & "' "
              .Source = .Source & "AND hpr.orden='" & porstCprCta!orden & "' "
              .Source = .Source & "AND hpr.codcta='" & porstCprCta!CodCta & "' "
              .Source = .Source & "GROUP BY hpr.codcta "
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
    If Not VerificaCtaCCo Then GoTo ErrorVerifica
    VerificaCtaCCo = (nImporteCpr_me = nImporteCta_me)
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
   cmdGenera.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property
