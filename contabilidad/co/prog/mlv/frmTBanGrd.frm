VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTBanGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   5580
   ClientLeft      =   1635
   ClientTop       =   1605
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   5580
   ScaleWidth      =   9360
   Visible         =   0   'False
   Begin VB.CommandButton cmd_anple 
      Caption         =   "Sin Ple"
      Height          =   495
      Left            =   7560
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame framepagos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   3240
      TabIndex        =   11
      Top             =   3360
      Width           =   6060
      Begin VB.CommandButton cmdreporte 
         Height          =   495
         Left            =   5160
         Picture         =   "frmTBanGrd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdabajo 
         Height          =   255
         Left            =   5640
         Picture         =   "frmTBanGrd.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdarriba 
         Height          =   255
         Left            =   5640
         Picture         =   "frmTBanGrd.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdprocesar 
         Height          =   495
         Left            =   4680
         Picture         =   "frmTBanGrd.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdvalidar 
         Height          =   495
         Left            =   4200
         Picture         =   "frmTBanGrd.frx":0968
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdfiltrar 
         Height          =   495
         Left            =   3720
         Picture         =   "frmTBanGrd.frx":0AB2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdtodos 
         Caption         =   "Sin Filtros"
         Height          =   220
         Left            =   3720
         MaskColor       =   &H8000000F&
         TabIndex        =   18
         Top             =   130
         Width           =   2295
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   ".."
         Height          =   285
         Index           =   0
         Left            =   2400
         Picture         =   "frmTBanGrd.frx":0BFC
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox cmbmoneda 
         Height          =   315
         Left            =   2760
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtDato 
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   525
      End
      Begin VB.Label lbl 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   855
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
         Height          =   285
         Index           =   0
         Left            =   645
         TabIndex        =   15
         Top             =   480
         Width           =   1740
      End
      Begin VB.Label lbl 
         Caption         =   "Banco"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   9360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9360
      Begin VB.CommandButton cmdpagos 
         Appearance      =   0  'Flat
         Caption         =   "&Pagos"
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
         Picture         =   "frmTBanGrd.frx":0DA6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   850
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
         Picture         =   "frmTBanGrd.frx":0EF0
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
         Picture         =   "frmTBanGrd.frx":0FF2
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
         Picture         =   "frmTBanGrd.frx":10F4
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
         Picture         =   "frmTBanGrd.frx":11F6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdSalir 
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
         Left            =   8640
         Picture         =   "frmTBanGrd.frx":12F8
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
         Left            =   4560
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
         Picture         =   "frmTBanGrd.frx":1442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid dgrMain 
      Align           =   1  'Align Top
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   6800
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
Attribute VB_Name = "frmTBanGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2016-02-02.07  correccion ple

Option Explicit

Public uocnnMain As ADODB.Connection
Public uorstMain_0 As ADODB.Recordset, _
       uorstMain_0Fil As ADODB.Recordset, _
       uorstMain_1 As ADODB.Recordset, _
       uorstMain_1Fil As ADODB.Recordset, _
       uorstUltiItem As ADODB.Recordset, _
       uorstMain_Grd As ADODB.Recordset, _
       uorstMain_GrdFil As ADODB.Recordset

Public usConnStrgSele_Grd As String, _
       usConnStrgSele_0 As String, _
       usConnStrgOrde_0 As String, _
       usConnStrgSele_1 As String, _
       usConnStrgWher_1 As String, _
       usConnStrgOrde_1 As String
'      usCOnnStrgWher

Private pnColumnaOrd As Integer

'[Propio del formulario.
Public uorstCODro As ADODB.Recordset, _
       uorstTGTCb As ADODB.Recordset, _
       uorstCoCta As ADODB.Recordset, _
       uorstCoCCo As ADODB.Recordset, _
       uorstTGAux As ADODB.Recordset, _
       uorstTGTDc As ADODB.Recordset, _
       uorstCOBanDet As ADODB.Recordset, _
       uorstCOTCbMes As ADODB.Recordset, _
       uorstCOFjo As ADODB.Recordset, _
       uorstCoBco As ADODB.Recordset, _
       uorstmedio As ADODB.Recordset
']

Public uorstCodMon As ADODB.Recordset '2016-02-02.07  correccion ple

Public valor As String
Public formatob As Integer
Public xctactemn As String
Public xctacteme As String
Public textofiltrado As Boolean
Public reporte As ADODB.Recordset
Public sumatoria As Double
Public sumatoriacta As Double
Public cualsuma As Integer

Private Sub cmbmoneda_Click()
    cmdvalidar.Visible = False
    cmdprocesar.Visible = False
End Sub

Private Sub cmd_anple_Click()
frmTBanGrd_anple.Show
End Sub

Private Sub cmdabajo_Click()
    framepagos.Top = 3360
End Sub

Private Sub cmdarriba_Click()
    framepagos.Top = 620
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      'txtDato(Index).SetFocus
   End Select
   
   ayudaban = True
   ppAyuBus Index
   Activar = True
   
End Sub

Private Sub cmdpagos_Click()
textofiltrado = True
If framepagos.Visible = False Then
    framepagos.Top = 3360
    framepagos.Left = 3720
    framepagos.Visible = True
Else
    framepagos.Visible = False
End If
cmdprocesar.Visible = False
cmdvalidar.Visible = False
End Sub

Private Sub cmdFiltrar_Click()

  If textofiltrado = True Then
    If txtDato(0).Text = "" Then MsgBox Choose(gsIdioma, "Ingresar Banco?", "Enter Bank?"), vbCritical: Exit Sub
    If cmbmoneda.Text = "" Then MsgBox Choose(gsIdioma, "Ingresar Moneda? ", "Enter Money? "), vbCritical: Exit Sub
  End If
  proceso = True

  Set uorstMain_GrdFil = New ADODB.Recordset
  Set uorstMain_0Fil = New ADODB.Recordset
  Set uorstMain_1Fil = New ADODB.Recordset
  Dim sCase As String
  
  sCase = "(CASE cobancab.tpodoc WHEN " & TPODOC_DPS_IND & " THEN 'DPS-' WHEN " & TPODOC_GRO_IND & " THEN 'GRO-' "
  sCase = sCase & "WHEN " & TPODOC_TRA_IND & " THEN 'TRA-' WHEN " & TPODOC_ORD_IND & " THEN 'ORD-' WHEN " & TPODOC_DEB_IND & " THEN 'DEB-' "
  sCase = sCase & "WHEN " & TPODOC_CRE_IND & " THEN 'CRE-' WHEN " & TPODOC_CHQ_IND & " THEN 'CHQ-' WHEN " & TPODOC_OTR_IND & " THEN 'OTR-' "
  sCase = sCase & "WHEN " & TPODOC_EFE_IND & " THEN 'EFE-' WHEN " & TPODOC_PEX_IND & " THEN 'PEX-' WHEN " & TPODOC_LTR_IND & " THEN 'LTR-' WHEN " & TPODOC_CGE_IND & " THEN 'CGE-' END)"
  
  sCase = "bnmediopago.abvmed,'-'"
 
  usConnStrgSele_Grd = "SELECT cobancab.coddro, cobancab.nroban, cobancab.fehban, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "Concat(" & sCase & ", cobancab.docban)", "(" & sCase & "+cobancab.docban)") & " AS cDocuBan, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "tgaux.razaux, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & Choose(gsIdioma, "cobancab.globan, ", "cobancab.globanx, ")
  usConnStrgSele_Grd = usConnStrgSele_Grd & "(CASE cobancab.tpoban WHEN " & TPOBAN_ING & " THEN '" & TPOBAN_ING_TXT & "' WHEN " & TPOBAN_EGR & " THEN '" & TPOBAN_EGR_TXT & "' END) AS cTpoBan, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "(CASE cobancab.gencpb WHEN '1' THEN 'x' WHEN '0' THEN ' ' END) AS cGenCpb,(CASE cobancab.genprc WHEN '1' THEN 'P' WHEN '0' THEN ' ' END) AS cGenprc, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "cobancab.codcta, cobancab.codaux, cobancab.codcco, cobancab.codfjo, cobancab.tpodoc, cobancab.docban, cobancab.portador, cobancab.tpomon, cobancab.tpotcb, cobancab.imptcb, cobancab.impmn,impme, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & Choose(gsIdioma, "cobancab.globanx, ", "cobancab.globan, ")
  usConnStrgSele_Grd = usConnStrgSele_Grd & "cobancab.codbco, cobancab.gencpb, cobancab.genprc , cobancab.tpoban, cobancab.codemp, cobancab.pdoano, cobancab.mespvs, cobancab.usrcre, cobancab.fyhcre, cobancab.usrmdf, cobancab.fyhmdf, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "Concat(cobancab.coddro, cobancab.nroban)", "(cobancab.coddro+cobancab.nroban)") & " AS cLlave "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "FROM cobancab "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "LEFT JOIN tgaux ON cobancab.codemp=tgaux.codemp AND cobancab.codaux=tgaux.codaux "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "LEFT JOIN bnmediopago ON cobancab.codemp=bnmediopago.codemp AND cobancab.tpodoc=bnmediopago.codmed "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "WHERE cobancab.codemp='" & gsCodEmp & "' "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.pdoano='" & gsAnoAct & "' "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.mespvs='" & gsMesAct & "' "
  If textofiltrado = True Then
  usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.codbco='" & txtDato(0) & "' "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and bnmediopago.indmod=2 and cobancab.tpoban=1 and cobancab.genprc=0 "
  End If
  
  '*******************************************************************
  
  usConnStrgSele_0 = "SELECT cobancab.coddro, cobancab.nroban, cobancab.fehban, "
  usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "cobancab.globan, ", "cobancab.globanx, ")
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.codcta, cobancab.codaux, cobancab.codcco, cobancab.codfjo, cobancab.tpodoc, cobancab.docban, "
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.portador, cobancab.tpomon, cobancab.tpotcb, cobancab.imptcb, cobancab.impmn,impme, "
  usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "cobancab.globanx, ", "cobancab.globan, ")
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.codbco, cobancab.gencpb, cobancab.genprc,cobancab.tpoban, cobancab.codemp, cobancab.pdoano, cobancab.mespvs, "
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.usrcre, cobancab.fyhcre, cobancab.usrmdf, cobancab.fyhmdf, "
  usConnStrgSele_0 = usConnStrgSele_0 & IIf(ps_Plataforma = pSrvMySql, "Concat(cobancab.coddro, cobancab.nroban)", "(cobancab.coddro+cobancab.nroban)") & " AS cLlave "
  usConnStrgSele_0 = usConnStrgSele_0 & "FROM cobancab "
  usConnStrgSele_0 = usConnStrgSele_0 & "LEFT JOIN bnmediopago ON cobancab.codemp=bnmediopago.codemp AND cobancab.tpodoc=bnmediopago.codmed "
  usConnStrgSele_0 = usConnStrgSele_0 & "WHERE cobancab.codemp='" & gsCodEmp & "' "
  usConnStrgSele_0 = usConnStrgSele_0 & "AND cobancab.pdoano='" & gsAnoAct & "' "
  usConnStrgSele_0 = usConnStrgSele_0 & "AND cobancab.mespvs='" & gsMesAct & "' "
  If textofiltrado = True Then
    usConnStrgSele_0 = usConnStrgSele_0 & "AND cobancab.codbco='" & txtDato(0) & "' "
    usConnStrgSele_0 = usConnStrgSele_0 & "AND cobancab.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and bnmediopago.indmod=2 and cobancab.tpoban=1 "
  End If
  usConnStrgOrde_0 = "ORDER BY cobancab.coddro, cobancab.nroban"
  
  '*******************************************************************
  
  usConnStrgSele_1 = "SELECT cobandet.nroitem, cobandet.codcta, cobandet.codcco, cobandet.codaux, tgtdc.abvtdc, cobandet.serdoc, cobandet.nrodoc, "
  usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "cobandet.gloite, ", "cobandet.gloitex, ")
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_EGR & "' THEN cobandet.impmn ELSE 0 END) AS cImpMN_Deb, "
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_ING & "' THEN cobandet.impmn ELSE 0 END) AS cImpMN_Hab, "
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_EGR & "' THEN cobandet.impme ELSE 0 END) AS cImpME_Deb, "
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_ING & "' THEN cobandet.impme ELSE 0 END) AS cImpME_Hab, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.codtdc, cobandet.tpoban, cobandet.tpomon, cobandet.tpotcb, cobandet.imptcb, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.refdoc, cobandet.pvsdoc, cobandet.coddro, cobandet.nroban, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.impmn, cobandet.impme, "
  usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "cobandet.gloitex, ", "cobandet.gloite, ")
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.codemp, cobandet.pdoano, cobandet.mespvs, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.usrcre, cobandet.fyhcre, cobandet.usrmdf, cobandet.fyhmdf, "
  usConnStrgSele_1 = usConnStrgSele_1 & IIf(ps_Plataforma = pSrvMySql, "Concat(cobandet.coddro, cobandet.nroban, cobandet.nroitem)", "(cobandet.coddro+cobandet.nroban+RTrim(cobandet.nroitem))") & " AS cLlave,cobandet.codbco,cobandet.tpocta "
  usConnStrgSele_1 = usConnStrgSele_1 & "FROM (cobandet "
  usConnStrgSele_1 = usConnStrgSele_1 & "LEFT JOIN TGTDc AS TGTDc ON cobandet.codemp=TGTDc.codemp AND cobandet.codtdc=TGTDc.codtdc) "
  usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
  usConnStrgWher_1 = usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
  usConnStrgWher_1 = usConnStrgWher_1 & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cobandet.coddro, cobandet.nroban)", "(cobandet.coddro+cobandet.nroban)") & "=' ' "
  usConnStrgOrde_1 = "ORDER BY cobandet.nroitem"
  
  '*******************************************************************
  
  Set dgrMain.DataSource = Nothing
  
  With uorstMain_GrdFil
     .ActiveConnection = uocnnMain
     .Source = usConnStrgSele_Grd & usConnStrgOrde_0
  '  .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic 'adLockReadOnly
     .Open
     .Properties("Unique Table").Value = "cobancab"
  End With
  
  With uorstMain_0Fil
    .ActiveConnection = uocnnMain
    .Source = usConnStrgSele_0 & usConnStrgOrde_0
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "cobancab"
  End With
  
  With uorstMain_1Fil
    .ActiveConnection = uocnnMain
    .Source = usConnStrgSele_1 & usConnStrgWher_1 & usConnStrgOrde_1
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "cobandet"
  End With
  
  Set dgrMain.DataSource = uorstMain_GrdFil
  ppDatosGrid
      
  cmdprocesar.Visible = True
  cmdvalidar.Visible = True
  textofiltrado = True
  cmdtodos.Enabled = True
  
End Sub

Private Sub cmdtodos_Click()
textofiltrado = False
cmdFiltrar_Click
framepagos.Visible = False
End Sub

Private Sub cmdvalidar_Click()

    Dim i, j As Integer
    Dim k As Integer
    Dim row As Variant
    Dim sLinea As String

    Dim filtrocab As String
    k = 1
    filtrocab = " concat(det.coddro,det.nroban) in ("
   
    If dgrMain.SelBookmarks.Count = 0 Then
        MsgBox Choose(gsIdioma, "¿No Existe Ningun Comprobante Seleccionado", "There is Not Proof Selected? "), vbInformation, Choose(gsIdioma, "Bancos", "Banks")
        Exit Sub
    Else
        For Each row In dgrMain.SelBookmarks
            i = row
            If k = 1 Then
                filtrocab = filtrocab & dgrMain.Columns(0).CellValue(i) & "" & dgrMain.Columns(1).CellValue(i)
                Else
                filtrocab = filtrocab & "," & dgrMain.Columns(0).CellValue(i) & "" & dgrMain.Columns(1).CellValue(i)
            End If
            k = k + 1
        Next row
        filtrocab = filtrocab & ")"
    End If

  Set reporte = New ADODB.Recordset

   With reporte
      .ActiveConnection = uocnnMain
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
     
   usConnStrgSele_Grd = usConnStrgSele_Grd & "WHERE cobancab.codemp='" & gsCodEmp & "' "
   usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.pdoano='" & gsAnoAct & "' "
   usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.mespvs='" & gsMesAct & "' "
   
   With reporte
   If .State = adStateOpen Then .Close
        .Source = " select det.coddro,cab.nroban,det.codaux,a.razaux," & Choose(gsIdioma, "b.detbco", "b.detbcox") & " AS detbco,cta.tpomon,cta.nroctacte from cobandet det"
        .Source = .Source & " left join cobancab cab on det.codemp=cab.codemp and det.pdoano=cab.pdoano and det.mespvs=cab.mespvs and det.coddro=cab.coddro and det.nroban=cab.nroban"
        .Source = .Source & " left join tgaux a on det.codaux=a.codaux and det.codemp=a.codemp"
        .Source = .Source & " left join coctaban cta on det.codaux=cta.codaux and det.codemp=cta.codemp and cta.codbco='" & txtDato(0).Text & "' and cta.tpomon='" & Right(cmbmoneda.Text, 1) & "' "
        .Source = .Source & " left join cobco b on cta.codbco=b.codbco and cta.codemp=b.codemp"
        .Source = .Source & " where det.codemp='" & gsCodEmp & "' and det.pdoano='" & gsAnoAct & "' and det.mespvs='" & gsMesAct & "' and  " & filtrocab & "  and cta.nroctacte is null "
        .Source = .Source & " group by det.codaux "
        .Open
   End With
      
    gpEncabezadoRpt frmMain.rptMain, "Listado de Cuentas Corrientes por Auxiliar", Date, True, False, reporte
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLCuentas.rpt"
      '.MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
    End With

End Sub

Private Sub Form_Load()
  Dim sCase As String
  
  cmbmoneda.AddItem "MN"
  cmbmoneda.AddItem "ME"
  proceso = False
  textofiltrado = False

  cmdpagos.Caption = Choose(gsIdioma, " Pagos ", " Payments ")
  lbl(0).Caption = Choose(gsIdioma, " Bancos ", " Banks ")
  lbl(1).Caption = Choose(gsIdioma, " Moneda ", " Money ")
  cmdtodos.Caption = Choose(gsIdioma, " Sin Filtros ", " Without Filters ")
  
  sCase = "(CASE cobancab.tpodoc WHEN " & TPODOC_DPS_IND & " THEN 'DPS-' WHEN " & TPODOC_GRO_IND & " THEN 'GRO-' "
  sCase = sCase & "WHEN " & TPODOC_TRA_IND & " THEN 'TRA-' WHEN " & TPODOC_ORD_IND & " THEN 'ORD-' WHEN " & TPODOC_DEB_IND & " THEN 'DEB-' "
  sCase = sCase & "WHEN " & TPODOC_CRE_IND & " THEN 'CRE-' WHEN " & TPODOC_CHQ_IND & " THEN 'CHQ-' WHEN " & TPODOC_OTR_IND & " THEN 'OTR-' "
  sCase = sCase & "WHEN " & TPODOC_EFE_IND & " THEN 'EFE-' WHEN " & TPODOC_PEX_IND & " THEN 'PEX-' WHEN " & TPODOC_LTR_IND & " THEN 'LTR-' WHEN " & TPODOC_CGE_IND & " THEN 'CGE-' END)"
  
  sCase = "bnmediopago.abvmed,'-'"
  
  '[Recordsets                          'Cambiar.
  usConnStrgSele_Grd = "SELECT cobancab.coddro, cobancab.nroban, cobancab.fehban, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "Concat(" & sCase & ", cobancab.docban)", "(" & sCase & "+cobancab.docban)") & " AS cDocuBan, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "tgaux.razaux, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & Choose(gsIdioma, "cobancab.globan, ", "cobancab.globanx, ")
  usConnStrgSele_Grd = usConnStrgSele_Grd & "(CASE cobancab.tpoban WHEN " & TPOBAN_ING & " THEN '" & TPOBAN_ING_TXT & "' WHEN " & TPOBAN_EGR & " THEN '" & TPOBAN_EGR_TXT & "' END) AS cTpoBan, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "(CASE cobancab.gencpb WHEN '1' THEN 'x' WHEN '0' THEN ' ' END) AS cGenCpb,(CASE cobancab.genprc WHEN '1' THEN 'P' WHEN '0' THEN ' ' END) AS cGenprc, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "cobancab.codcta, cobancab.codaux, cobancab.codcco, cobancab.codfjo, cobancab.tpodoc, cobancab.docban, cobancab.portador, cobancab.tpomon, cobancab.tpotcb, cobancab.imptcb, cobancab.impmn,impme, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & Choose(gsIdioma, "cobancab.globanx, ", "cobancab.globan, ")
  usConnStrgSele_Grd = usConnStrgSele_Grd & "cobancab.codbco, cobancab.gencpb,cobancab.genprc, cobancab.tpoban, cobancab.codemp, cobancab.pdoano, cobancab.mespvs, cobancab.usrcre, cobancab.fyhcre, cobancab.usrmdf, cobancab.fyhmdf, "
  usConnStrgSele_Grd = usConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "Concat(cobancab.coddro, cobancab.nroban)", "(cobancab.coddro+cobancab.nroban)") & " AS cLlave "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "FROM cobancab "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "LEFT JOIN tgaux ON cobancab.codemp=tgaux.codemp AND cobancab.codaux=tgaux.codaux "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "LEFT JOIN bnmediopago ON cobancab.codemp=bnmediopago.codemp AND cobancab.tpodoc=bnmediopago.codmed "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "WHERE cobancab.codemp='" & gsCodEmp & "' "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.pdoano='" & gsAnoAct & "' "
  usConnStrgSele_Grd = usConnStrgSele_Grd & "AND cobancab.mespvs='" & gsMesAct & "' "
  
  usConnStrgSele_0 = "SELECT cobancab.coddro, cobancab.nroban, cobancab.fehban, "
  usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "cobancab.globan, ", "cobancab.globanx, ")
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.codcta, cobancab.codaux, cobancab.codcco, cobancab.codfjo, cobancab.tpodoc, cobancab.docban, "
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.portador, cobancab.tpomon, cobancab.tpotcb, cobancab.imptcb, cobancab.impmn,impme, "
  usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "cobancab.globanx, ", "cobancab.globan, ")
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.codbco, cobancab.gencpb, cobancab.genprc,cobancab.tpoban, cobancab.codemp, cobancab.pdoano, cobancab.mespvs, "
  usConnStrgSele_0 = usConnStrgSele_0 & "cobancab.usrcre, cobancab.fyhcre, cobancab.usrmdf, cobancab.fyhmdf, "
  usConnStrgSele_0 = usConnStrgSele_0 & IIf(ps_Plataforma = pSrvMySql, "Concat(cobancab.coddro, cobancab.nroban)", "(cobancab.coddro+cobancab.nroban)") & " AS cLlave "
  usConnStrgSele_0 = usConnStrgSele_0 & "FROM cobancab "
  usConnStrgSele_0 = usConnStrgSele_0 & "WHERE cobancab.codemp='" & gsCodEmp & "' "
  usConnStrgSele_0 = usConnStrgSele_0 & "AND cobancab.pdoano='" & gsAnoAct & "' "
  usConnStrgSele_0 = usConnStrgSele_0 & "AND cobancab.mespvs='" & gsMesAct & "' "
  usConnStrgOrde_0 = "ORDER BY cobancab.coddro, cobancab.nroban"
  
  usConnStrgSele_1 = "SELECT cobandet.nroitem, cobandet.codcta, cobandet.codcco, cobandet.codaux, tgtdc.abvtdc, cobandet.serdoc, cobandet.nrodoc, "
  usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "cobandet.gloite, ", "cobandet.gloitex, ")
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_EGR & "' THEN cobandet.impmn ELSE 0 END) AS cImpMN_Deb, "
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_ING & "' THEN cobandet.impmn ELSE 0 END) AS cImpMN_Hab, "
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_EGR & "' THEN cobandet.impme ELSE 0 END) AS cImpME_Deb, "
  usConnStrgSele_1 = usConnStrgSele_1 & "(CASE cobandet.tpoban WHEN '" & TPOBAN_ING & "' THEN cobandet.impme ELSE 0 END) AS cImpME_Hab, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.codtdc, cobandet.tpoban, cobandet.tpomon, cobandet.tpotcb, cobandet.imptcb, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.refdoc, cobandet.pvsdoc, cobandet.coddro, cobandet.nroban, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.impmn, cobandet.impme, "
  usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "cobandet.gloitex, ", "cobandet.gloite, ")
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.codemp, cobandet.pdoano, cobandet.mespvs, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.usrcre, cobandet.fyhcre, cobandet.usrmdf, cobandet.fyhmdf, "
  usConnStrgSele_1 = usConnStrgSele_1 & "cobandet.codmon, " '2016-02-02.07  correccion ple
  usConnStrgSele_1 = usConnStrgSele_1 & IIf(ps_Plataforma = pSrvMySql, "Concat(cobandet.coddro, cobandet.nroban, cobandet.nroitem)", "(cobandet.coddro+cobandet.nroban+RTrim(cobandet.nroitem))") & " AS cLlave,cobandet.codbco,cobandet.tpocta "
  usConnStrgSele_1 = usConnStrgSele_1 & "FROM (cobandet "
  usConnStrgSele_1 = usConnStrgSele_1 & "LEFT JOIN TGTDc AS TGTDc ON cobandet.codemp=TGTDc.codemp AND cobandet.codtdc=TGTDc.codtdc) "
  usConnStrgWher_1 = "WHERE cobandet.codemp='" & gsCodEmp & "' "
  usConnStrgWher_1 = usConnStrgWher_1 & "AND cobandet.pdoano='" & gsAnoAct & "' "
  usConnStrgWher_1 = usConnStrgWher_1 & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cobandet.coddro, cobandet.nroban)", "(cobandet.coddro+cobandet.nroban)") & "=' ' "
  usConnStrgOrde_1 = "ORDER BY cobandet.nroitem "
   
  Set uocnnMain = New ADODB.Connection
  Set uorstMain_Grd = New ADODB.Recordset
  Set uorstMain_0 = New ADODB.Recordset
  Set uorstMain_1 = New ADODB.Recordset
  Set uorstUltiItem = New ADODB.Recordset
  Set uorstCODro = New ADODB.Recordset
  Set uorstTGTCb = New ADODB.Recordset
  Set uorstCoCta = New ADODB.Recordset
  Set uorstCoCCo = New ADODB.Recordset
  Set uorstTGAux = New ADODB.Recordset
  Set uorstTGTDc = New ADODB.Recordset
  Set uorstCOBanDet = New ADODB.Recordset
  Set uorstCOTCbMes = New ADODB.Recordset
  Set uorstCOFjo = New ADODB.Recordset
  Set uorstCoBco = New ADODB.Recordset
  Set uorstmedio = New ADODB.Recordset
   
  Set uorstCodMon = New ADODB.Recordset '2016-02-02.07  correccion ple
  
  With uocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  
  With uorstMain_Grd
     .ActiveConnection = uocnnMain
     .Source = usConnStrgSele_Grd & usConnStrgOrde_0
     '.CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic 'adLockReadOnly
     .Open
     .Properties("Unique Table").Value = "cobancab"
  End With
  
  With uorstMain_0
    .ActiveConnection = uocnnMain
    .Source = usConnStrgSele_0 & usConnStrgOrde_0
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "cobancab"
  End With
  
  With uorstMain_1
    .ActiveConnection = uocnnMain
    .Source = usConnStrgSele_1 & usConnStrgWher_1 & usConnStrgOrde_1
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
    .Properties("Unique Table").Value = "cobandet"
  End With
  
  With uorstUltiItem
    .ActiveConnection = uocnnMain
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
  End With
  
  With uorstCODro
    .ActiveConnection = uocnnMain
    .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro, codemp, Cpb" & gsMesAct & ", "
    .Source = .Source & "codemp, pdoano "
    .Source = .Source & "FROM CODro "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  
  With uorstTGTCb
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.FehTCb, a.ImpTCb_Cpr, a.ImpTCb_Vta "
    .Source = .Source & "FROM TGTCb a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  
  With uorstCoCta
    .ActiveConnection = uocnnMain
    .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta, "
    .Source = .Source & "tpomon, TpoTCb, TpoAnl, IndAjd, IndCCo, IndDoc, IndFjo, CodCCo_Def, "
    .Source = .Source & "CodCta_AjD_Deb, CodCta_AjD_Hab, CodCCo_AjD_Deb, CodCCo_AjD_Hab, codbco "
    .Source = .Source & "FROM COCta "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND TpoCta=" & TPOCTA_TRA & " "
    .Source = .Source & "AND EstCta='" & ESTCTA_ACT & "'"
    '.CursorLocation = adUseClient   'Es el Default.
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
    '.CursorLocation = adUseClient   'Es el Default.
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
  
  With uorstCOBanDet
    .ActiveConnection = uocnnMain
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
  End With
  
  With uorstCOTCbMes
    .ActiveConnection = uocnnMain
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
  End With
  
  With uorstCOFjo
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.CodFjo, " & Choose(gsIdioma, "a.DetFjo", "a.DetFjox") & " AS DetFjo "
    .Source = .Source & "FROM COFjo a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodFjo)>2"
    ''     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  
  With uorstCoBco
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.codbco, " & Choose(gsIdioma, "a.detbco", "a.detbcox") & " AS detbco "
    .Source = .Source & "FROM cobco a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
    ''     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  
  With uorstmedio
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.codmed, abvmed,desmed,indmod "
    .Source = .Source & "FROM bnmediopago a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
    ''     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  ']
'ini 2016-02-02.07  correccion ple
  With uorstCodMon
     .ActiveConnection = CONNSTRG & gsNomBDC
     .Source = gf_tb_sunat(CODSUNAT_004)
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
'fin 2016-02-02.07  correccion ple
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, True, False, False, False, False, True, aLabel
  ']
  dgrMain.MarqueeStyle = dbgHighlightRow
  Set dgrMain.DataSource = uorstMain_Grd
  
  cmdprocesar.Visible = False
  cmdvalidar.Visible = False
  Activar = False
End Sub

Private Sub Form_Activate()
   
   'Orden: Nuevo, Eliminar, Vista Previa, Imprimir.
   
   zaOpciones = Array(gbPms01, gbPms03, gbPms04, gbPms05)
   ppDatosGrid
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(0).Caption
   
   If Activar = False Then
        framepagos.Visible = False
   Else
        framepagos.Visible = True
   End If
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   gpTUg_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   uorstCodMon.Close '2016-02-02.07  correccion ple

   uorstMain_Grd.Close
   uorstMain_0.Close
   uocnnMain.Close
   
   Set uorstCodMon = Nothing '2016-02-02.07  correccion ple
   
   Set uorstCOTCbMes = Nothing
   Set uorstMain_Grd = Nothing
   Set uorstMain_0 = Nothing
   Set uocnnMain = Nothing
End Sub

Public Sub cmdNuevo_Click()
  framepagos.Visible = False
  'Verificación de Mes Cerrado.
  If gbCieCpb Then MsgBox TEXT_9016, vbCritical: Exit Sub
   
  '[ No pertenece al Formulario - Agregado por Angel
  With uorstMain_1
    .Close
    .Source = usConnStrgSele_1 & " WHERE cobandet.coddro='    ' " & usConnStrgOrde_1
    .Open
    .Properties("Unique Table").Value = "cobandet"
  End With
  gpTUg_Nuevo Me, frmTBanCab          'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
  framepagos.Visible = False
  On Error GoTo Err
  
  If proceso = False Then
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then
    MsgBox TEXT_8001, vbCritical
    Exit Sub
  End If

  '[Búsqueda del ítem.
  uorstMain_0.Requery
  uorstMain_0.MoveFirst
  uorstMain_0.Find "cLlave='" & uorstMain_Grd!coddro & uorstMain_Grd!nroban & "'"
  ']
  Else
  'Verificación de existencia de ítemes.
  If uorstMain_GrdFil.RecordCount = 0 Then
    MsgBox TEXT_8001, vbCritical
    Exit Sub
  End If

  '[Búsqueda del ítem.
  uorstMain_0Fil.Requery
  uorstMain_0Fil.MoveFirst
  uorstMain_0Fil.Find "cLlave='" & uorstMain_GrdFil!coddro & uorstMain_GrdFil!nroban & "'"
  ']
  
  End If
  
  With frmTBanCab                     'Cambiar Formulario de Datos.
    .zbNuevo = False
    .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
    .txtLlave(0).Enabled = False
    .cboTpoBan.Enabled = False
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
  framepagos.Visible = False
  On Error GoTo Err
  
  Dim dsLlaveSiguiente As String
  
  'Verificación de Mes Cerrado.
  If gbCieCpb Then MsgBox TEXT_9016, vbCritical: Exit Sub
  'Verificación de existencia de ítemes.
  If uorstMain_Grd.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub
'ini 2016-05-27/28 nivel=asisten no elimin datos
   If gsNvlUsr = NVLUSR_ASIS Then
      MsgBox TEXT_9026, vbCritical
      Exit Sub
   End If
'fin 2016-05-27/28 nivel=asisten no elimin datos
  'Mensaje de verificación            'Cambiar.
  If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
    uorstMain_0.MoveFirst
    uorstMain_0.Find "cLlave = '" & uorstMain_Grd!coddro & uorstMain_Grd!nroban & "'"
    
    uocnnMain.BeginTrans
    ' Elimino el comprobante de diario
    uocnnMain.Execute "DELETE FROM cocpbcab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND mespvs='" & gsMesAct & "' AND coddro='" & Trim(dgrMain.Columns(0)) & "' AND nrocpb='" & Trim(dgrMain.Columns(1)) & "' AND tpognr='" & TPOGNR_BAN & "'"
    uorstMain_0.Properties("Unique Table").Value = "cobancab"
    uorstMain_0.Delete
    uocnnMain.CommitTrans
    
    'Busca siguiente ítem.
    With uorstMain_Grd
      .MoveNext
      If .EOF Then .MoveLast
      dsLlaveSiguiente = !coddro & !nroban
      .Requery
      If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
    End With
    ppDatosGrid
    ' actualizo recordset principal
    uorstMain_0.Requery
    If uorstMain_0.RecordCount > 0 Then uorstMain_0.Find "cLlave = '" & dsLlaveSiguiente & "'"
    
    'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
    fEstMayUpd
    'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
    
  End If
  dgrMain.SetFocus
  
  Exit Sub
Err:
  gpErrores
  
  uocnnMain.RollbackTrans             'RESTAURA TRANSACCION.
End Sub

Public Sub cmdRefrescar_Click()
  framepagos.Visible = False
  If proceso = False Then
    frmTBanGrd.uorstMain_Grd.Requery
  Else
   frmTBanGrd.uorstMain_GrdFil.Requery
  End If
  frmTBanGrd.ppDatosGrid
  dgrMain.SetFocus
End Sub

Public Sub cmdImprimir_Click(Index As Integer)
  framepagos.Visible = False
  '[Datos del formulario de impresión.  'Cambiar.
  frmLCpb.Caption = Choose(gsIdioma, "Listado de ", "Listing of ") & Me.Caption
  frmLCpb.Show vbModal
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
  usConnStrgOrde_0 = usConnStrgOrde_0 & pnColumnaOrd + 1
  With uorstMain_Grd
    .Close
    .Source = usConnStrgSele_Grd & usConnStrgOrde_0
    .Open
  End With
  Set dgrMain.DataSource = uorstMain_Grd
  ppDatosGrid
  
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
    '          Case vbDate
    '          dsCriterio = dgrMain.Columns(pnColumnaOrd).DataField & " = " & txtBuscar
    End Select
    .Find dsCriterio, , , 1
    If .EOF = True Then .Bookmark = dvRegistroActual
  End With
  Exit Sub
Err:
  If Err.Number = 3001 Then   'Se produce al llegar a EOF de adcMain.
    uorstMain_Grd.Bookmark = dvRegistroActual
  Else
    gpErrores
  End If
End Sub

Function pfNumItemBan(ByVal s_Ano As String, ByVal s_Mes As String, ByVal s_Diario As String, s_Comprobante As String) As Integer
  ' s_Ano             Año donde  se genera
  ' s_Mes             Mes donde  se genera
  ' s_Diario          Codigo de diario para generar numero
  ' s_Comprobante     Numero de comprobante para generar item
    
  Dim porstRetorno As ADODB.Recordset
  Dim s_Sentencia As String
  
  s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(nroitem), 0) AS nNumMaxItem "
  s_Sentencia = s_Sentencia & "FROM CoBanDet "
  s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND pdoano='" & s_Ano & "' "
  s_Sentencia = s_Sentencia & "AND MesPvs='" & s_Mes & "' "
  s_Sentencia = s_Sentencia & "AND CodDro='" & s_Diario & "' "
  s_Sentencia = s_Sentencia & "AND nroban='" & s_Comprobante & "'"
  Set porstRetorno = New ADODB.Recordset
  With porstRetorno
    .ActiveConnection = frmTBanGrd.uocnnMain
    '.CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = s_Sentencia
    .Open
  End With
  pfNumItemBan = CInt(porstRetorno!nNumMaxItem) + 1
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
        .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("coddro").DefinedSize + 1)
       Case 1
        .Item(dnNum).Caption = Choose(gsIdioma, "NºComp.", "NºVoucher")
        .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("nroban").DefinedSize + 1)
       Case 2
        .Item(dnNum).Caption = Choose(gsIdioma, "Fecha", "Date")
        .Item(dnNum).Width = 100 * (7 + 3)
       Case 3
        .Item(dnNum).Caption = Choose(gsIdioma, "Documento", "Document")
        .Item(dnNum).Width = 100 * (13)
       Case 4
        .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
        .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("razaux").DefinedSize - 42.8)
       Case 5
        .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
        .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("globan").DefinedSize - 42.5)
       Case 6
        .Item(dnNum).Caption = Choose(gsIdioma, "Tipo", "Type")
        .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("tpoban").DefinedSize + 4)
        .Item(dnNum).Alignment = dbgCenter
       Case 7
        .Item(dnNum).Caption = Choose(gsIdioma, "Gen", "Gen")
        .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("cgencpb").DefinedSize + 3)
        .Item(dnNum).Alignment = dbgCenter
       Case 8
        .Item(dnNum).Caption = Choose(gsIdioma, "Prc", "Prc")
        .Item(dnNum).Width = 100 * (uorstMain_Grd.Fields("cgenprc").DefinedSize + 3)
        .Item(dnNum).Alignment = dbgCenter
       Case Else
        .Item(dnNum).Visible = False
      End Select
    Next
  End With
End Sub
Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0                        'Cambiar (añadir índices).
    modAyuBus.Bco_CodBan "", txtDato(tnIndex).Text, 0, 0, Me.Top + framepagos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + framepagos.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frm0AyuBusBan.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frm0AyuBusBan.uvDato2
    formatob = frm0AyuBusBan.uvDato4
    xctactemn = frm0AyuBusBan.uvDato5
    xctacteme = frm0AyuBusBan.uvDato6
  End Select
End Sub
'[Código propio del formulario.
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
']

Private Sub cmdprocesar_Click()

    ' FORMATOS
    ' 1 BANCO DE CREDITO RENTING
    ' 2 BANCO CONTINENTAL
    ' 3 BANCO SCOTIA BANK
    ' 5 BANCO CITIBANK
    ' 6 BANCO DE CREDITO SODEXHO

    Dim i, j As Integer
    Dim k As Integer
    Dim row As Variant
    Dim sLinea As String
    Dim filtro As String
    Dim filtrocab As String
    
    If formatob <> 1 And formatob <> 2 And formatob <> 3 And formatob <> 5 And formatob <> 6 Then MsgBox Choose(gsIdioma, "No Existe formato de Exportacion para este Banco?", "There is no Export format for this Bank?"), vbCritical: Exit Sub
    
    k = 1
    filtro = " concat(cobandet.coddro,cobandet.nroban) in ("
    filtrocab = " concat(cobancab.coddro,cobancab.nroban) in ("
   
    If txtDato(0).Text = "" Then MsgBox Choose(gsIdioma, "Ingresar Banco?", "Enter Bank?"), vbCritical: Exit Sub
    If cmbmoneda.Text = "" Then MsgBox Choose(gsIdioma, "Ingresar Moneda? ", "Enter Money? "), vbCritical: Exit Sub
    
   
    If dgrMain.SelBookmarks.Count = 0 Then
        MsgBox "¿No Existe Ningun Comprobante Seleccionado", vbInformation, "Bancos"
        Exit Sub
    Else
        For Each row In dgrMain.SelBookmarks
            i = row
            If k = 1 Then
                filtro = filtro & dgrMain.Columns(0).CellValue(i) & "" & dgrMain.Columns(1).CellValue(i)
                filtrocab = filtrocab & dgrMain.Columns(0).CellValue(i) & "" & dgrMain.Columns(1).CellValue(i)
                Else
                filtro = filtro & "," & dgrMain.Columns(0).CellValue(i) & "" & dgrMain.Columns(1).CellValue(i)
                filtrocab = filtrocab & "," & dgrMain.Columns(0).CellValue(i) & "" & dgrMain.Columns(1).CellValue(i)
            End If
            sLinea = sLinea & IIf(sLinea <> Empty, ", ", Empty) & dgrMain.Columns(0).CellValue(i) & "-" & dgrMain.Columns(1).CellValue(i)
            k = k + 1
        Next row
        
        filtro = filtro & ")"
        filtrocab = filtrocab & ")"
        
        sLinea = "Ha Seleccionado los Comprobantes # " & sLinea
        MsgBox sLinea, vbInformation, row & " Comprobantes Seleccionados"
        
    End If
    
    Dim Rst As ADODB.Recordset
    Dim RstU As ADODB.Recordset
    Dim RstCuales As ADODB.Recordset
    Dim RstSumas As ADODB.Recordset
    
    Dim R As Boolean
    Dim sql As String
    
    Dim tipdoc As Integer
        
    Set Rst = New ADODB.Recordset
    Set RstU = New ADODB.Recordset
    Set RstCuales = New ADODB.Recordset
    Set RstSumas = New ADODB.Recordset
         
    If formatob = 1 Then
         
    sql = "select ' ' as espacio1,' ' as espacio2,case cobancab.tpodoc when 1 then 2 else '0' end,case cobancab.tpodoc when 1 then 'C' else 'G' end,coctaban.nroctacte as cta,razaux,case cobancab.tpomon when 'N' then 'S/' else 'US' end,case cobancab.tpomon when 'N' then cobandet.impmn else cobandet.impme end,'RUC' as DI,cobandet.codaux,case cobandet.codtdc when '07' then 'N' else 'F' end as Doc,concat(cobandet.serdoc,'',right(cobandet.nrodoc,6)),'1' as TA,cobandet.gloite,'0' as flag1,'0' as flag2,'0' as flag3,'' as direccion,'' as distrito,'' as provincia,'' as departamento,'' as contacto from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp "
    sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('003') and cobandet.tpoban=1 and " & filtro
         
    ElseIf formatob = 2 Then
    
    'sql = "select '002L' as espacio1,' ' as espacio2,case cobancab.tpodoc when 1 then 2 else '0' end,case cobancab.tpodoc when 1 then 'C' else 'G' end,coctaban.nroctacte as cta,razaux,case cobancab.tpomon when 'N' then 'S/' else 'US' end,case cobancab.tpomon when 'N' then cobandet.impmn else cobandet.impme end,'RUC' as DI,cobandet.codaux,case cobandet.codtdc when '07' then 'N' else 'F' end as Doc,concat(cobandet.serdoc,'',right(cobandet.nrodoc,6)),'1' as TA,cobandet.gloite,'0' as flag1,'0' as flag2,'0' as flag3,'' as direccion,'' as distrito,'' as provincia,'' as departamento,'' as contacto from cobandet "
    sql = "select ' ' as espacio,'002R' as espacio1,left(cobandet.codaux,11),'P' as espacio2,coctaban.nroctacte as cta,razaux,case cobancab.tpomon when 'N' then cobandet.impmn else cobandet.impme end,'F' as espacio3,concat(cobandet.serdoc,'',right(cobandet.nrodoc,6)),'N' as TA,cobandet.gloite,'' as flag1,'' as flag2,'' as contacto from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp "
    sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('003') and cobandet.tpoban=1 and " & filtro
    
    ElseIf formatob = 3 Then
        
    sql = "select cobandet.codaux,razaux,concat(serdoc,'-',right(nrodoc,9)),date_format(cobancab.fehban,'%Y%m%d'),case cobancab.tpomon when 'N' then cobandet.impmn else cobandet.impme end,"
    sql = sql & " if(cobandet.codbco=cobancab.codbco,case coctaban.tpocta when '" & TPOCTA_COR & "' then 2 else 3 end,4),"
    sql = sql & " if(cobandet.codbco=cobancab.codbco,left(coctaban.nroctacte,3),''),"
    sql = sql & " if(cobandet.codbco=cobancab.codbco,right(coctaban.nroctacte,7),''),"
    sql = sql & " '0' as Flag,"
    sql = sql & " if(isnull(tgaux.email),'',tgaux.email) as email,"
    sql = sql & " if(cobandet.codbco=cobancab.codbco,'',coctaban.nrocci) as codint"
    sql = sql & " from cobandet"
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano"
    sql = sql & " and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban"
    sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp"
    sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco"
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cobandet.pvsdoc=0 and cobancab.tpodoc in ('003') and cobandet.tpoban=1 and " & filtro
    
    ElseIf formatob = 5 Then
         
    'sql = "select "
    'sql = sql & " cobandet.coddro,cobandet.nroban,cobandet.codcta,cobandet.codaux,cobandet.serdoc,cobandet.nrodoc,cobandet.gloite,"
    'sql = sql & " cobandet.refdoc,cobandet.tpomon,concat(repeat(' ',17-length(cast(cobandet.impmn as char))),cast(cobandet.impmn as char)),concat(repeat(' ',17-length(cast(cobandet.impme as char))),cast(cobandet.impme as char)),cobandet.codbco,cobandet.codtdc,cobancab.tpodoc,"
    'sql = sql & " tgaux.razaux , tgaux.diraux, tgaux.email, coctaban.nroctacte, coctaban.nrocci, coctaban.tpocta"
    'sql = sql & " from cobandet"
    'sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    'sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp "
    'sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco "
    'sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    'sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    'sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    'sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    'sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    'sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('001','007') and cobandet.tpoban=1 and " & filtro
    'sql = sql & " order by cobandet.coddro,cobandet.nroban "
    
    sql = "select "
    sql = sql & " cobandet.coddro,cobandet.nroban,cobandet.codcta,cobandet.codaux,cobandet.serdoc,cobandet.nrodoc,cobandet.gloite,"
    sql = sql & " cobandet.refdoc,cobandet.tpomon,concat(repeat(' ',17-length(cast(cobandet.impmn as char))),cast(cobandet.impmn as char)),concat(repeat(' ',17-length(cast(cobandet.impme as char))),cast(cobandet.impme as char)),cobandet.codbco,cobandet.codtdc,cobancab.tpodoc,"
    sql = sql & " tgaux.razaux , tgaux.diraux, tgaux.email, coctaban.nroctacte, coctaban.nrocci, coctaban.tpocta"
    sql = sql & " from cobandet"
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp "
    sql = sql & " left join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "'"
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and cocta.inddoc=1 and cobancab.tpodoc in ('001','007') and cobandet.tpoban=1 and " & filtro
    sql = sql & " order by cobandet.coddro,cobandet.nroban "
    
    ElseIf formatob = 6 Then
    
    sql = "select ' ' as espacio1,' ' as espacio2,case cobancab.tpodoc when 1 then 2 else '0' end,case cobancab.tpodoc when 1 then 'C' else 'G' end,coctaban.nroctacte as cta,razaux,case cobancab.tpomon when 'N' then 'S/' else 'US' end,case cobancab.tpomon when 'N' then IF(left(cobandet.codcta,1)='1',-1*cobandet.impmn,cobandet.impmn) else IF(left(cobandet.codcta,1)='1',-1*cobandet.impme,cobandet.impme) end,'RUC' as DI,cobandet.codaux,case cobandet.codtdc when '07' then 'N' else 'F' end as Doc,concat(cobandet.serdoc,'',right(cobandet.nrodoc,6)),'1' as TA,cobandet.gloite,'0' as flag1,'0' as flag2,'0' as flag3,'' as direccion,'' as distrito,'' as provincia,'' as departamento,'' as contacto,coctaban.tpocta as tipocta from cobandet  "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp "
    sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('001')  and " & filtro & " order by 6,8 desc"

    End If
    
    Rst.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If Rst.RecordCount = 0 Then
        MsgBox " No Existen datos para Generar el Archivo ", vbInformation
        Exit Sub
    End If
    
    CommonDialog1.DialogTitle = "SaveAs"
    CommonDialog1.FileName = gsRUCEmp & ".txt"
    CommonDialog1.InitDir = "C:\"
    CommonDialog1.Filter = "*.txt"
    CommonDialog1.DefaultExt = "*.txt"
    CommonDialog1.Filter = "TextFiles(*.txt)|*.txt"
    CommonDialog1.ShowSave
    
    If formatob = 1 Then
    
    sql = "select " & IIf(cmbmoneda.Text = "MN", "cobandet.impmn", "cobandet.impme") & ",nroctacte,cobancab.tpodoc as doc,cobandet.codaux as aux from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('003') and cobandet.tpoban=1 and " & filtro
    
    RstSumas.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If xctactemn = "" Or Len(xctactemn) < 4 Then xctactemn = "0000"
    If xctacteme = "" Or Len(xctacteme) < 4 Then xctacteme = "0000"
    
        xctactemn = Replace(xctactemn, "-", "")
        xctacteme = Replace(xctacteme, "-", "")
       
    sumatoria = 0
    sumatoriacta = 0 + IIf(cmbmoneda.Text = "MN", CDbl(Right(xctactemn, Len(xctactemn) - 3)), CDbl(Right(xctacteme, Len(xctacteme) - 3)))
    
        If RstSumas.RecordCount = 0 Then
            Exit Sub
        Else
            RstSumas.MoveFirst
            ' recorre todo el recordset
            For cualsuma = 0 To RstSumas.RecordCount - 1
                sumatoria = sumatoria + RstSumas.Fields(0)
                tipdoc = RstSumas.Fields(2).Value
                
                If tipdoc = 1 Then
                    sumatoriacta = sumatoriacta + CDbl(Right(Replace(RstSumas.Fields(1), "-", ""), Len(Replace(RstSumas.Fields(1), "-", "")) - 3))
                ElseIf tipdoc = 12 Then
                    sumatoriacta = sumatoriacta + CDbl(RstSumas.Fields(3))
                Else
                    sumatoriacta = sumatoriacta + 0
                End If
                RstSumas.MoveNext
            Next
        End If
        
        cualsuma = RstSumas.RecordCount
        
    End If
    
    If formatob = 2 Then
    
    sql = "select " & IIf(cmbmoneda.Text = "MN", "cobandet.impmn", "cobandet.impme") & ",nroctacte,cobancab.tpodoc as doc,cobandet.codaux as aux from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('003') and cobandet.tpoban=1 and " & filtro
    
    RstSumas.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If xctactemn = "" Or Len(xctactemn) < 4 Then xctactemn = "0000"
    If xctacteme = "" Or Len(xctacteme) < 4 Then xctacteme = "0000"
    
        xctactemn = Replace(xctactemn, "-", "")
        xctacteme = Replace(xctacteme, "-", "")
    
    sumatoria = 0
    
        If RstSumas.RecordCount = 0 Then
            Exit Sub
        Else
            RstSumas.MoveFirst
            ' recorre todo el recordset
            For cualsuma = 0 To RstSumas.RecordCount - 1
                sumatoria = sumatoria + RstSumas.Fields(0)
                RstSumas.MoveNext
            Next
        End If
        
        cualsuma = RstSumas.RecordCount
        
    End If
    
    If formatob = 5 Then
    
    'sql = "select " & IIf(cmbmoneda.Text = "MN", "cobandet.impmn", "cobandet.impme") & ",nroctacte,cobancab.tpodoc as doc,cobandet.codaux as aux from cobandet "
    'sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    'sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    'sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco "
    'sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    'sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    'sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    'sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    'sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    'sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('001') and cobandet.tpoban=1 and " & filtro
    
    
    sql = "select " & IIf(cmbmoneda.Text = "MN", "cobandet.impmn", "cobandet.impme") & ",nroctacte,cobancab.tpodoc as doc,cobandet.codaux as aux from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    sql = sql & " left join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "'"
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and cocta.inddoc=1 and cobancab.tpodoc in ('001','007') and cobandet.tpoban=1 and " & filtro
    
    
    RstSumas.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
       
    sumatoria = 0
    
        If RstSumas.RecordCount = 0 Then
            Exit Sub
        Else
            RstSumas.MoveFirst
            ' recorre todo el recordset
            For cualsuma = 0 To RstSumas.RecordCount - 1
                sumatoria = sumatoria + RstSumas.Fields(0)
                RstSumas.MoveNext
            Next
        End If
        
        
    End If
    
    
    If formatob = 6 Then
    
    sql = "select " & IIf(cmbmoneda.Text = "MN", "cobandet.impmn", "cobandet.impme") & ",nroctacte,cobancab.tpodoc as doc,cobandet.codaux as aux,cobandet.codcta as cdocta from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('001')  and " & filtro
    
    RstSumas.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If xctactemn = "" Or Len(xctactemn) < 4 Then xctactemn = "0000"
    If xctacteme = "" Or Len(xctacteme) < 4 Then xctacteme = "0000"
    
        xctactemn = Replace(xctactemn, "-", "")
        xctacteme = Replace(xctacteme, "-", "")
       
    sumatoria = 0
    sumatoriacta = 0 + IIf(cmbmoneda.Text = "MN", CDbl(Right(xctactemn, Len(xctactemn) - 3)), CDbl(Right(xctacteme, Len(xctacteme) - 3)))
    
        If RstSumas.RecordCount = 0 Then
            Exit Sub
        Else
            RstSumas.MoveFirst
            ' recorre todo el recordset
            For cualsuma = 0 To RstSumas.RecordCount - 1
                sumatoria = sumatoria + IIf(Left(RstSumas.Fields(4), 1) = 1, RstSumas.Fields(0) * (-1), RstSumas.Fields(0))
                tipdoc = RstSumas.Fields(2).Value
                
                If tipdoc = 1 Then
                    sumatoriacta = sumatoriacta + CDbl(Right(Replace(RstSumas.Fields(1), "-", ""), Len(Replace(RstSumas.Fields(1), "-", "")) - 3))
                ElseIf tipdoc = 12 Then
                    sumatoriacta = sumatoriacta + CDbl(RstSumas.Fields(3))
                Else
                    sumatoriacta = sumatoriacta + 0
                End If
                
                RstSumas.MoveNext
            Next
        End If
        
        cualsuma = RstSumas.RecordCount
        
    End If
    
    
    If formatob = 1 Then
        R = Recordset_a_Csv1(Rst, CommonDialog1.FileName)
    ElseIf formatob = 2 Then
        R = Recordset_a_Csv2(Rst, CommonDialog1.FileName)
    ElseIf formatob = 3 Then
        R = Recordset_a_Csv3(Rst, CommonDialog1.FileName)
    ElseIf formatob = 5 Then
        R = Recordset_a_Csv5(Rst, CommonDialog1.FileName, filtro)
    ElseIf formatob = 6 Then
        R = Recordset_a_Csv6(filtro, filtrocab, CommonDialog1.FileName)
    End If
             
             
             
    MsgBox " Se generó el archivo " & gsRUCEmp & ".txt, con " & Rst.RecordCount & " Registros", vbInformation
    
    If formatob <> 5 Then
    
    If formatob = 6 Then
    
'    sql = "select cobancab.coddro,cobancab.nroban from cobandet "
'    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
'    sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
'    sql = sql & " inner join coctaban on cobancab.codaux=coctaban.codaux and cobancab.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
'    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
'    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
'    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
'    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
'    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cobandet.pvsdoc=0 and cobancab.tpodoc in ('001')  and " & filtro
'
    filtro = Replace(filtro, "cobandet", "cobancab")
    
    sql = "select coddro,nroban from cobancab "
    sql = sql & " where codemp='" & gsCodEmp & "' "
    sql = sql & " and pdoano='" & gsAnoAct & "' "
    sql = sql & " and mespvs='" & gsMesAct & "' "
    sql = sql & " and codbco='" & txtDato(0) & "' "
    sql = sql & " and tpodoc in ('001') and  " & filtro
    
    
    Else
    
    sql = "select cobancab.coddro,cobancab.nroban from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    sql = sql & " inner join coctaban on cobancab.codaux=coctaban.codaux and cobancab.codemp=coctaban.codemp and cobancab.codbco=coctaban.codbco "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cobandet.pvsdoc=0 and cobancab.tpodoc in ('003') and cobandet.tpoban=1 and " & filtro
    
    
        
    End If
    
    Else
    
    'sql = "select cobancab.coddro,cobancab.nroban from cobandet "
    'sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    'sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    'sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco "
    'sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    'sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    'sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    'sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    'sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cobandet.pvsdoc=0 and cobancab.tpodoc in ('001') and cobandet.tpoban=1 and " & filtro
 
    sql = "select cobancab.coddro,cobancab.nroban from cobandet "
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobancab.codaux=tgaux.codaux and cobancab.codemp=tgaux.codemp "
    sql = sql & " left join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "'"
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and cobandet.pvsdoc=0 and cobancab.tpodoc in ('001','007') and cobandet.tpoban=1 and " & filtro
 
    
    End If
    
    RstCuales.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    Dim cual As Integer
    
    If RstCuales.RecordCount = 0 Then
        Exit Sub
    Else
        filtrocab = " concat(cobancab.coddro,cobancab.nroban) in ("
        RstCuales.MoveFirst
        ' recorre todo el recordset
        For cual = 0 To RstCuales.RecordCount - 1
            If cual = 0 Then
               filtrocab = filtrocab & RstCuales.Fields(0) & "" & RstCuales.Fields(1)
               Else
               filtrocab = filtrocab & "," & RstCuales.Fields(0) & "" & RstCuales.Fields(1)
            End If
            RstCuales.MoveNext
        Next
        
    End If
    filtrocab = filtrocab & ")"
    
    sql = "update cobancab set cobancab.genprc=1 "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' and " & filtrocab
    
    RstU.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If Not Rst.State = adStateOpen Then
        Rst.Close
    End If
    If Not Rst Is Nothing Then
        Set Rst = Nothing
    End If
        
End Sub

Function Recordset_a_Csv1(rs As Recordset, path As String) As Boolean
    On Error GoTo Err_function
    Dim columna
    Dim fila As Integer
    Dim desde As Integer
    Dim cadenav As String
    Dim tamanno As Integer
    Dim ceros As String
    Dim i As Integer
    Dim sumatoriatexto As String
    Dim sumatoriactatexto As String
    Dim cualsumatexto As String
    Dim cadena As String
    
    sumatoriatexto = Replace(Format(sumatoria, "0.00"), ".", "")
    
    For i = Len(Trim(Str(sumatoriatexto))) To 14
        cadena = "0" & cadena
    Next
    
    sumatoriatexto = cadena & sumatoriatexto
    cadena = ""
    
    For i = Len(Trim(Str(sumatoriacta))) To 14
        cadena = "0" & cadena
    Next
    
    sumatoriactatexto = cadena & sumatoriacta
    
    For i = Len(cualsuma) To 6
        cualsumatexto = "0" & cualsumatexto
    Next
    
    cualsumatexto = cualsumatexto & cualsuma
    
    ' Crea el archivo
    Open path For Output As #1
    
    If cmbmoneda.Text = "MN" Then
        If xctactemn = "" Then MsgBox Choose(gsIdioma, "No existe Cuenta Corriente Moneda Nacional?", "There is no National Currency Current Account?"), vbCritical: 'Exit Sub
    Else
        If xctacteme = "" Then MsgBox Choose(gsIdioma, "No existe Cuenta Corriente Moneda Extranjera?", "There Current Account Foreign Currency?"), vbCritical: 'Exit Sub
    End If
    
    xctactemn = Replace(xctactemn, "-", "")
    xctacteme = Replace(xctacteme, "-", "")
    
    xctactemn = Left(xctactemn, 3) & IIf(cmbmoneda.Text = "MN", 0, 1) & Right(xctactemn, Len(xctactemn) - 3)
    xctacteme = Left(xctacteme, 3) & IIf(cmbmoneda.Text = "MN", 0, 1) & Right(xctacteme, Len(xctacteme) - 3)
    
    Print #1, "#1PC" & IIf(cmbmoneda.Text = "MN", xctactemn, xctacteme) & IIf(cmbmoneda.Text = "MN", "      S/", "      US") & sumatoriatexto & Format(Date, "dd") & Format(Date, "mm") & Format(Date, "yyyy") & "     REFERENCIA CASO" & sumatoriactatexto & cualsumatexto & "9" & Space(15) & "1";
    Print #1, "" & Chr(13) & Chr(10);
    
    ' Se mueve al primer registro
    rs.MoveFirst
    ' Recorre todo el Recordset
    For fila = 0 To rs.RecordCount - 1
           
           'Nombre del Campo
            cadenav = ""
            tamanno = 1
            valor = Trim(rs.Fields(0))
        
            If Len(valor) = tamanno Then
            ElseIf Len(valor) > tamanno Then
                    valor = Left(valor, tamanno)
            ElseIf Len(valor) < tamanno Then
                For desde = Len(valor) To tamanno - 1
                    cadenav = cadenav & ""
                Next
            End If
        
            valor = valor & cadenav
        
            Print #1, valor;
        
        ' Recorre Todos los Campos
        For columna = 1 To rs.Fields.Count - 1
            ' Imprime la Fila Actual en el fichero
            valor = IIf(rs.Fields(columna) = Null, "", Trim(Replace(rs.Fields(columna), "-", "")))
                        
            cadenav = ""
            
           Select Case columna
            Case 1
                tamanno = 1
            Case 2
                tamanno = 1
            Case 3
                tamanno = 1
            Case 4
                tamanno = 20
            Case 5
                tamanno = 40
            Case 6
                tamanno = 2
            Case 7
                tamanno = 15
            Case 8
                tamanno = 3
            Case 9
                tamanno = 12
            Case 10
                tamanno = 1
            Case 11
                tamanno = 10
            Case 12
                tamanno = 1
            Case 13
                tamanno = 40
            Case 14
                tamanno = 1
            Case 15
                tamanno = 1
            Case 16
                tamanno = 1
            Case 17
                tamanno = 40
            Case 18
                tamanno = 20
            Case 19
                tamanno = 20
            Case 20
                tamanno = 20
            Case 21
                tamanno = 40
           End Select
            
           If columna = 4 Then
                If Len(valor) = 13 Then
                    valor = Left(valor, 3) & IIf(cmbmoneda.Text = "MN", 0, 1) & Right(valor, Len(valor) - 3)
                ElseIf Len(valor) = 14 Then
                    valor = Left(valor, 3) & Right(valor, Len(valor) - 3)
                Else
                    valor = ""
                End If
           End If
           
           If Len(valor) = tamanno Then
           ElseIf Len(valor) > tamanno Then
           valor = Left(valor, tamanno)
           ElseIf Len(valor) < tamanno Then
                'If Columna <> 4 Then
                    For desde = Len(valor) To tamanno - 1
                        cadenav = cadenav & " "
                    Next
                'End If
           End If
           
           If columna = 7 Then
               cadenav = ""
               valor = FormatNumber(valor, 2)
               valor = Replace(valor, ".", "")
               valor = Replace(valor, ",", "")
               For i = 1 To 15 - Len(valor)
                    ceros = ceros & "0"
               Next
               valor = ceros & valor
               ceros = ""
           End If
            
           valor = valor & cadenav
                            
           Print #1, "" & valor;
        Next
            ' escribe una línea en blanco
        Print ""
            ' salto de carro
        Print #1, "" & Chr(13) & Chr(10);
            ' mueve el recordset al siguiente registro
        rs.MoveNext
        Next
    ' cierra el archivo
    Close #1
    Exit Function
Err_function:
    MsgBox Err.Description, vbCritical
    Close
End Function

Function Recordset_a_Csv2(rs As Recordset, path As String) As Boolean

On Error GoTo Err_function
    Dim columna
    Dim fila As Integer
    Dim desde As Integer
    Dim cadenav As String
    Dim tamanno As Integer
    Dim ceros As String
    Dim i As Integer
    Dim sumatoriatexto As String
    Dim sumatoriactatexto As String
    Dim cualsumatexto As String
    Dim cadena As String
    
    sumatoriatexto = Replace(Format(sumatoria, "0.00"), ".", "")
    
    For i = Len(Trim(Str(sumatoriatexto))) To 14
        cadena = "0" & cadena
    Next
    
    sumatoriatexto = cadena & sumatoriatexto
    cadena = ""
    
    For i = Len(cualsuma) To 5
        cualsumatexto = "0" & cualsumatexto
    Next
    
    cualsumatexto = cualsumatexto & cualsuma
    
    ' Crea el archivo
    Open path For Output As #1
    
    If cmbmoneda.Text = "MN" Then
        If xctactemn = "" Then MsgBox Choose(gsIdioma, "No existe Cuenta Corriente Moneda Nacional?", "There is no National Currency Current Account?"), vbCritical: 'Exit Sub
    Else
        If xctacteme = "" Then MsgBox Choose(gsIdioma, "No existe Cuenta Corriente Moneda Extranjera?", "There Current Account Foreign Currency?"), vbCritical: 'Exit Sub
    End If
    
    xctactemn = Replace(xctactemn, "-", "")
    xctacteme = Replace(xctacteme, "-", "")
    
    'xctactemn = Left(xctactemn, 3) & IIf(cmbmoneda.Text = "MN", 0, 1) & Right(xctactemn, Len(xctactemn) - 3)
    'xctacteme = Left(xctacteme, 3) & IIf(cmbmoneda.Text = "MN", 0, 1) & Right(xctacteme, Len(xctacteme) - 3)
    
    xctactemn = Left(xctactemn, 3) & Right(xctactemn, Len(xctactemn) - 3)
    xctacteme = Left(xctacteme, 3) & Right(xctacteme, Len(xctacteme) - 3)
    
    
    Print #1, "750" & IIf(cmbmoneda.Text = "MN", xctactemn, xctacteme) & IIf(cmbmoneda.Text = "MN", "PEN", "USD") & sumatoriatexto & "A" & Space(9) & "PROVEEDORES" & Space(14) & cualsumatexto & "N000000000000000000" & Space(50);
    Print #1, "" & Chr(13) & Chr(10);
    
    ' Se mueve al primer registro
    rs.MoveFirst
    ' Recorre todo el Recordset
    For fila = 0 To rs.RecordCount - 1
           
           'Nombre del Campo
            cadenav = ""
            tamanno = 1
            valor = Trim(rs.Fields(0))
        
            If Len(valor) = tamanno Then
            ElseIf Len(valor) > tamanno Then
                    valor = Left(valor, tamanno)
            ElseIf Len(valor) < tamanno Then
                For desde = Len(valor) To tamanno - 1
                    cadenav = cadenav & ""
                Next
            End If
        
            valor = valor & cadenav
        
            Print #1, valor;
        
        ' Recorre Todos los Campos
        For columna = 1 To rs.Fields.Count - 1
            ' Imprime la Fila Actual en el fichero
            valor = IIf(rs.Fields(columna) = Null, "", Trim(Replace(rs.Fields(columna), "-", "")))
                        
            cadenav = ""
            
           Select Case columna
            Case 1
                tamanno = 4
            Case 2
                tamanno = 12
            Case 3
                tamanno = 1
            Case 4
                tamanno = 20
            Case 5
                tamanno = 40
            Case 6
                tamanno = 15
            Case 7
                tamanno = 1
            Case 8
                tamanno = 12
            Case 9
                tamanno = 1
            Case 10
                tamanno = 40
            Case 11
                tamanno = 1
            Case 12
                tamanno = 50
            Case 13
                tamanno = 80
            Case Else
                tamanno = 0
            End Select
            
           'If Columna = 4 Then
           '     If Len(valor) = 13 Then
           '         valor = Left(valor, 3) & IIf(cmbmoneda.Text = "MN", 0, 1) & Right(valor, Len(valor) - 3)
           '     ElseIf Len(valor) = 14 Then
           '         valor = Left(valor, 3) & Right(valor, Len(valor) - 3)
           '     Else
           '         valor = ""
           '     End If
           'End If
           
           If Len(valor) = tamanno Then
           ElseIf Len(valor) > tamanno Then
           valor = Left(valor, tamanno)
           ElseIf Len(valor) < tamanno Then
                'If Columna <> 4 Then
                    For desde = Len(valor) To tamanno - 1
                        cadenav = cadenav & " "
                    Next
                'End If
           End If
           
           If columna = 6 Then
               cadenav = ""
               valor = FormatNumber(valor, 2)
               valor = Replace(valor, ".", "")
               valor = Replace(valor, ",", "")
               For i = 1 To 15 - Len(valor)
                    ceros = ceros & "0"
               Next
               valor = ceros & valor
               ceros = ""
           End If
            
           valor = valor & cadenav
                                                  
           Print #1, "" & valor;
        Next
            ' escribe una línea en blanco
        Print ""
            ' salto de carro
        Print #1, "" & Chr(13) & Chr(10);
            ' mueve el recordset al siguiente registro
        rs.MoveNext
        Next
    ' cierra el archivo
    Close #1
    Exit Function
Err_function:
    MsgBox Err.Description, vbCritical
    Close


End Function

Function Recordset_a_Csv3(rs As Recordset, path As String) As Boolean
    On Error GoTo Err_function
    Dim columna
    Dim fila As Integer
    Dim desde As Integer
    Dim cadenav As String
    Dim tamanno As Integer
    Dim ceros As String
    Dim i As Integer
    
    ' Crea el archivo
    Open path For Output As #1
    ' Se mueve al primer registro
    rs.MoveFirst
    ' recorre todo el recordset
    For fila = 0 To rs.RecordCount - 1
           
           ' nombre del campo
            cadenav = ""
            tamanno = 11
            valor = Trim(rs.Fields(0))
        
            If Len(valor) = tamanno Then
            ElseIf Len(valor) > tamanno Then
                    valor = Left(valor, tamanno)
            ElseIf Len(valor) < tamanno Then
                For desde = Len(valor) To tamanno - 1
                    cadenav = cadenav & ""
                Next
            End If
        
            valor = valor & cadenav
        
            Print #1, valor;
        
        ' recorre todos los campos
        For columna = 1 To rs.Fields.Count - 1
                     
            cadenav = ""
            valor = ""
            
            ' imprime la fila actual en el fichero
            valor = IIf(rs.Fields(columna) = Null, "", Trim(rs.Fields(columna)))
             
            Select Case columna
            Case 1
                tamanno = 60
            Case 2
                tamanno = 14
            Case 3
                tamanno = 8
            Case 4
                tamanno = 11
            Case 5
                tamanno = 1
            Case 6
                tamanno = 3
            Case 7
                tamanno = 7
            Case 8
                tamanno = 1
            Case 9
                tamanno = 30
            Case 10
                tamanno = 20
           End Select
            
            If Len(valor) = tamanno Then
            ElseIf Len(valor) > tamanno Then
            valor = Left(valor, tamanno)
            ElseIf Len(valor) < tamanno Then
                If columna <> 4 Then
                    For desde = Len(valor) To tamanno - 1
                        cadenav = cadenav & " "
                    Next
                End If
            End If
            
            If columna = 4 Then
               valor = FormatNumber(valor, 2)
               valor = Replace(valor, ".", "")
               valor = Replace(valor, ",", "")
               For i = 1 To 11 - Len(valor)
                    ceros = ceros & "0"
               Next
               valor = ceros & valor
               ceros = ""
            End If
            
            valor = valor & cadenav
                        
            Print #1, "" & valor;
        Next
            ' escribe una línea en blanco
        Print ""
            ' salto de carro
        Print #1, "" & Chr(13) & Chr(10);
            ' mueve el recordset al siguiente registro
        rs.MoveNext
        Next
    ' cierra el archivo
    Close #1
    Exit Function
Err_function:
    MsgBox Err.Description, vbCritical
    Close
End Function

Function Recordset_a_Csv5(rs As Recordset, path As String, filtro As String) As Boolean
    On Error GoTo Err_function
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim sql As String
    Dim RstCiti As ADODB.Recordset
    Dim conta As Integer
    Dim cuantos As Integer
    Dim xcont As Integer
    

    cuantos = 0
     
    Set RstCiti = New ADODB.Recordset

    'sql = "select "
    'sql = sql & " cobandet.coddro,cobandet.nroban,cobancab.tpodoc,cobandet.codaux,concat(repeat(' ',17-length(cast(sum(cobandet.impmn) as char))),cast(sum(cobandet.impmn) as char)),concat(repeat(' ',17-length(cast(sum(cobandet.impme) as char))),cast(sum(cobandet.impme) as char)),if(cobandet.codbco='07','072','071'),if(cobandet.tpomon='N','PEN','USD'),concat(left(cobancab.globan,35),repeat(' ',35-length(cobancab.globan))),if(cobandet.tpocta='A','02','01'),concat(left(tgaux.razaux,80),repeat(' ',80-length(tgaux.razaux))),"
    'sql = sql & " concat(left(ifnull(tgaux.diraux,''),35),repeat(' ',35-length(ifnull(tgaux.diraux,'')))),concat('Lima',repeat(' ',11)),if(cobandet.codbco='07',space(3),concat('0',cobandet.codbco)),if(cobandet.codbco='07',space(8),'00000000'),if(cobandet.codbco='07',space(35),coctaban.nrocci),if(cobandet.codbco='07',space(2),if(cobandet.tpocta='A','02','01')),if(cobandet.codbco='07',coctaban.nroctacte,space(10)),if(cobandet.codbco='07',if(cobandet.tpocta='A','02','01'),space(2)),if(cobandet.codbco='07','001',space(3)),concat(left(ifnull(tgaux.email,''),50),repeat(' ',50-length(ifnull(tgaux.email,'')))) "
    'sql = sql & " from cobandet"
    'sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    'sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp "
    'sql = sql & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco "
    'sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    'sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    'sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    'sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    'sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    'sql = sql & " and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "' and cocta.inddoc=1 and cobancab.tpodoc in ('001') and cobandet.tpoban=1 and " & filtro
    'sql = sql & " group by cobandet.coddro,cobandet.nroban,cobancab.tpodoc,cobandet.codaux order by 1,2"
    
    
    sql = "select "
    sql = sql & " cobandet.coddro,cobandet.nroban,cobancab.tpodoc,concat(cobandet.codaux,repeat(' ',20-length(cobandet.codaux))),concat(repeat(' ',17-length(cast(sum(cobandet.impmn) as char))),cast(sum(cobandet.impmn) as char)),concat(repeat(' ',17-length(cast(sum(cobandet.impme) as char))),cast(sum(cobandet.impme) as char)),if(cobancab.tpodoc='007','073',if(cobandet.codbco='07','072','071')),if(cobandet.tpomon='N','PEN','USD'),concat(left(cobancab.globan,35),repeat(' ',35-length(cobancab.globan))),if(cobandet.tpocta='A','02','01'),concat(left(tgaux.razaux,80),repeat(' ',80-length(tgaux.razaux))),"
    sql = sql & " concat(left(ifnull(tgaux.diraux,''),35),repeat(' ',35-length(ifnull(tgaux.diraux,'')))),concat('Lima',repeat(' ',11)),if(cobancab.tpodoc='007',space(3),if(cobandet.codbco='07',space(3),concat('0',cobandet.codbco))),if(cobancab.tpodoc='007',space(8),if(cobandet.codbco='07',space(8),'00000000')),if(cobancab.tpodoc='007',space(35),if(cobandet.codbco='07',space(35),coctaban.nrocci)),if(cobancab.tpodoc='007',space(2),if(cobandet.codbco='07',space(2),if(cobandet.tpocta='A','02','01'))),if(cobancab.tpodoc='007',space(10),if(cobandet.codbco='07',coctaban.nroctacte,space(10))),if(cobancab.tpodoc='007',space(2),if(cobandet.codbco='07',if(cobandet.tpocta='A','02','01'),space(2))),if(cobancab.tpodoc='007',space(3),if(cobandet.codbco='07','001',space(3))),concat(left(ifnull(tgaux.email,''),50),repeat(' ',50-length(ifnull(tgaux.email,'')))),left(cobancab.docban,3) "
    sql = sql & " from cobandet"
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban "
    sql = sql & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp "
    sql = sql & " left join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp and cobandet.codbco=coctaban.codbco and coctaban.tpomon='" & IIf(cmbmoneda.Text = "MN", "N", "E") & "'"
    sql = sql & " inner join cocta on cobandet.codcta=cocta.codcta and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and cocta.inddoc=1 and cobancab.tpodoc in ('001','007') and cobandet.tpoban=1 and " & filtro
    sql = sql & " group by cobandet.coddro,cobandet.nroban,cobancab.tpodoc,cobandet.codaux order by 1,2"
    
    RstCiti.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    ' Crea el archivo
    Open path For Output As #1
    
    RstCiti.MoveFirst
    xcont = RstCiti.RecordCount
    For i = 0 To RstCiti.RecordCount - 1
    
    If RstCiti.Fields(6) = "072" Then
    'cuenta citibank
    Print #1, "PAY604" & IIf(cmbmoneda.Text = "MN", Format(xctactemn, "0000000000"), Format(xctacteme, "0000000000")) & Format(Date, "yy") & Format(Date, "mm") & Format(Date, "dd") & RstCiti.Fields(6) & Format(RstCiti.Fields(0) & RstCiti.Fields(1), "000000000000000") & Format(i + 1, "00000000") & RstCiti.Fields(3) & RstCiti.Fields(7) & RstCiti.Fields(3) & Format(IIf(RstCiti.Fields(7) = "PEN", RstCiti.Fields(4) * 100, RstCiti.Fields(5) * 100), "000000000000000") & Space(6) & RstCiti.Fields(8) & Space(105) & "22" & RstCiti.Fields(9) & RstCiti.Fields(10) & RstCiti.Fields(11) & Space(35) & RstCiti.Fields(12) & Space(30) & RstCiti.Fields(13) & RstCiti.Fields(14) & RstCiti.Fields(15) & RstCiti.Fields(16) & Space(122) & RstCiti.Fields(17) & RstCiti.Fields(18) & RstCiti.Fields(19) & Space(55) & RstCiti.Fields(20) & Format(IIf(RstCiti.Fields(7) = "PEN", RstCiti.Fields(4) * 100, RstCiti.Fields(5) * 100), "000000000000000") & "2" & Space(267)
    End If
    If RstCiti.Fields(6) = "071" Then
    'cuenta otros bancos
    Print #1, "PAY604" & IIf(cmbmoneda.Text = "MN", Format(xctactemn, "0000000000"), Format(xctacteme, "0000000000")) & Format(Date, "yy") & Format(Date, "mm") & Format(Date, "dd") & RstCiti.Fields(6) & Format(RstCiti.Fields(0) & RstCiti.Fields(1), "000000000000000") & Format(i + 1, "00000000") & RstCiti.Fields(3) & RstCiti.Fields(7) & RstCiti.Fields(3) & Format(IIf(RstCiti.Fields(7) = "PEN", RstCiti.Fields(4) * 100, RstCiti.Fields(5) * 100), "000000000000000") & Space(6) & RstCiti.Fields(8) & Space(105) & "22" & RstCiti.Fields(9) & RstCiti.Fields(10) & RstCiti.Fields(11) & Space(35) & RstCiti.Fields(12) & Space(30) & RstCiti.Fields(13) & RstCiti.Fields(14) & RstCiti.Fields(15) & Space(15) & RstCiti.Fields(16) & Space(122) & RstCiti.Fields(17) & RstCiti.Fields(18) & "099" & Space(55) & RstCiti.Fields(20) & Format(IIf(RstCiti.Fields(7) = "PEN", RstCiti.Fields(4) * 100, RstCiti.Fields(5) * 100), "000000000000000") & "2" & Space(267)
    End If
    If RstCiti.Fields(6) = "073" Then
    'cuenta otros bancos
    Print #1, "PAY604" & IIf(cmbmoneda.Text = "MN", Format(xctactemn, "0000000000"), Format(xctacteme, "0000000000")) & Format(Date, "yy") & Format(Date, "mm") & Format(Date, "dd") & RstCiti.Fields(6) & Format(RstCiti.Fields(0) & RstCiti.Fields(1), "000000000000000") & Format(i + 1, "00000000") & RstCiti.Fields(3) & RstCiti.Fields(7) & RstCiti.Fields(3) & Format(IIf(RstCiti.Fields(7) = "PEN", RstCiti.Fields(4) * 100, RstCiti.Fields(5) * 100), "000000000000000") & Space(6) & RstCiti.Fields(8) & Space(105) & "00" & RstCiti.Fields(9) & RstCiti.Fields(10) & RstCiti.Fields(11) & Space(35) & RstCiti.Fields(12) & Space(30) & RstCiti.Fields(13) & RstCiti.Fields(14) & RstCiti.Fields(15) & Space(15) & RstCiti.Fields(16) & Space(107) & RstCiti.Fields(17) & RstCiti.Fields(18) & RstCiti.Fields(21) & Space(55) & RstCiti.Fields(20) & Format(IIf(RstCiti.Fields(7) = "PEN", RstCiti.Fields(4) * 100, RstCiti.Fields(5) * 100), "000000000000000") & "2" & Space(267)
    End If
    
    cuantos = cuantos + 1

        conta = 1
        
        rs.MoveFirst
        For j = 0 To rs.RecordCount - 1
        
            If rs.Fields(0) & rs.Fields(1) = RstCiti.Fields(0) & RstCiti.Fields(1) Then
            
                If conta = 1 Then
                
                    Print #1, "VOI604" & IIf(cmbmoneda.Text = "MN", Format(xctactemn, "0000000000"), Format(xctacteme, "0000000000")) & Format(RstCiti.Fields(0) & RstCiti.Fields(1), "000000000000000") & Format(i + 1, "00000000") & Format(1, "0000") & "  No. Documento        Monto       Descuento            Total              00000" & Space(127)
                    cuantos = cuantos + 1
                               
                    Print #1, "VOI604" & IIf(cmbmoneda.Text = "MN", Format(xctactemn, "0000000000"), Format(xctacteme, "0000000000")) & Format(RstCiti.Fields(0) & RstCiti.Fields(1), "000000000000000") & Format(i + 1, "00000000") & Format(2, "0000") & Format(rs.Fields(4) & rs.Fields(5), "0000000000000000") & rs.Fields(9) & "             0.00" & rs.Fields(9) & Space(140)
                    cuantos = cuantos + 1
                    
                    conta = 2
                    
                Else
               
                    Print #1, "VOI604" & IIf(cmbmoneda.Text = "MN", Format(xctactemn, "0000000000"), Format(xctacteme, "0000000000")) & Format(RstCiti.Fields(0) & RstCiti.Fields(1), "000000000000000") & Format(i + 1, "00000000") & Format(conta + 1, "0000") & Format(rs.Fields(4) & rs.Fields(5), "0000000000000000") & rs.Fields(9) & "             0.00" & rs.Fields(9) & Space(140)
                    cuantos = cuantos + 1
                    conta = conta + 1
                    
                End If
            
            End If
            
        rs.MoveNext
        Next
                 
                    Print #1, "VOI604" & IIf(cmbmoneda.Text = "MN", Format(xctactemn, "0000000000"), Format(xctacteme, "0000000000")) & Format(RstCiti.Fields(0) & RstCiti.Fields(1), "000000000000000") & Format(i + 1, "00000000") & Format(conta + 1, "0000") & "        TOTALES:" & RstCiti.Fields(4) & "             0.00" & RstCiti.Fields(4) & Space(140)
                    cuantos = cuantos + 1
        
    RstCiti.MoveNext
    Next
    
    sumatoria = sumatoria * 100
    
    Print #1, "TRL" & Format(xcont, "000000000000000") & Format(sumatoria, "000000000000000") & "000000000000000" & Format(cuantos, "000000000000000") & Space(37)
    
    Close #1
    
    Exit Function
Err_function:
    MsgBox Err.Description, vbCritical
    Close
End Function

Private Sub txtDato_Change(Index As Integer)
    cmdvalidar.Visible = False
    cmdprocesar.Visible = False
End Sub

Private Sub cmdreporte_Click()

Dim filtro As String
Dim RstImpresion As ADODB.Recordset
Dim RstConsulta As ADODB.Recordset
Dim contador As Integer

Dim xBanco As String
Dim xCuentaCorriente As String
Dim xCuentaContable As String
Dim xComprobante As String
Dim xGlosa As String
Dim xFecha As String

Set RstImpresion = New ADODB.Recordset
Set RstConsulta = New ADODB.Recordset

Dim sql As String

If txtDato(0).Text = "" Then MsgBox Choose(gsIdioma, "Ingresar Banco?", "Enter Bank?"), vbCritical: Exit Sub
If cmbmoneda.Text = "" Then MsgBox Choose(gsIdioma, "Ingresar Moneda? ", "Enter Money? "), vbCritical: Exit Sub
    
If dgrMain.SelBookmarks.Count = 0 Then
        MsgBox "¿No Existe Ningun Comprobante Seleccionado", vbInformation, "Bancos"
        Exit Sub
Else
        filtro = dgrMain.Columns(0) & "" & dgrMain.Columns(1)
End If

    sql = "select concat(cobancab.codbco,'-',cobco.detbco),concat(cobancab.codcta,'-',cocta.detcta),concat(coddro,'-',nroban),"
    sql = sql & " case cobancab.tpomon when 'N' then ctactemn else ctacteme end,globan,fehban "
    sql = sql & " from cobancab"
    sql = sql & " inner join cobco on cobancab.codemp=cobco.codemp and cobancab.codbco=cobco.codbco"
    sql = sql & " inner join cocta on cobancab.codemp=cocta.codemp and cobancab.codcta=cocta.codcta and cobancab.pdoano=cocta.pdoano"
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "'  and cobancab.pdoano='" & gsAnoAct & "'  and cobancab.mespvs='" & gsMesAct & "'  and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and concat(coddro,nroban) in (" & filtro & ")"

    RstConsulta.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If RstConsulta.RecordCount = 0 Then
    Else
            RstConsulta.MoveFirst
            For contador = 0 To RstConsulta.RecordCount - 1
                xBanco = RstConsulta.Fields(0)
                xCuentaContable = RstConsulta.Fields(1)
                xComprobante = RstConsulta.Fields(2)
                xCuentaCorriente = RstConsulta.Fields(3)
                xGlosa = RstConsulta.Fields(4)
                xFecha = RstConsulta.Fields(5)
                RstConsulta.MoveNext
            Next
    End If
    
    RstConsulta.Close
    

    With RstImpresion
      .ActiveConnection = uocnnMain
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
    End With
     
    With RstImpresion
    If .State = adStateOpen Then .Close
        .Source = "select 'RUC' as RUC,cobandet.codaux,razaux,diraux,coctaban.tpocta,coctaban.nroctacte,concat(cobandet.serdoc,'-',cobandet.nrodoc),"
        .Source = .Source & " case cobancab.tpomon when 'N' then 'S/' else 'US' end,"
        .Source = .Source & " sum(case cobancab.tpomon when 'N' then IF(left(cobandet.codcta,1)='1',-1*cobandet.impmn,cobandet.impmn)"
        .Source = .Source & " else IF(left(cobandet.codcta,1)='1',-1*cobandet.impme,cobandet.impme) end)"
        .Source = .Source & " from cobandet inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano"
        .Source = .Source & " and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban"
        .Source = .Source & " inner join tgaux on cobandet.codaux=tgaux.codaux and cobandet.codemp=tgaux.codemp"
        .Source = .Source & " inner join coctaban on cobandet.codaux=coctaban.codaux and cobandet.codemp=coctaban.codemp"
        .Source = .Source & " and cobancab.codbco=coctaban.codbco  inner join cocta on cobandet.codcta=cocta.codcta"
        .Source = .Source & " and cobandet.pdoano=cocta.pdoano and cobandet.codemp=cocta.codemp"
        .Source = .Source & " where cobancab.codemp='" & gsCodEmp & "'  and cobancab.pdoano='" & gsAnoAct & "'  and cobancab.mespvs='" & gsMesAct & "'"
        .Source = .Source & " and cobancab.codbco='" & txtDato(0) & "'  and coctaban.tpomon='N' and cocta.inddoc=1 and cobancab.tpodoc in ('001')"
        .Source = .Source & " and concat(cobandet.coddro,cobandet.nroban) in (" & filtro & ") group by cobandet.codaux"
        .Open
    End With
      
gpEncabezadoRpt frmMain.rptMain, "TRANSFERENCIA DE PROVEEDORES I/O AFILIADOS", Date, True, False, RstImpresion
With frmMain.rptMain
      .ReportFileName = gsRutRpt & "Transferencia.rpt"
      '[Parámetros adicionales.
      .ParameterFields(1) = "Banco;" & "Banco :" & xBanco & ";true"
      .ParameterFields(2) = "CuentaCorriente;" & "Cuenta Corriente :" & xCuentaCorriente & ";true"
      .ParameterFields(3) = "CuentaContable;" & "Cuenta Contable :" & xCuentaContable & ";true"
      .ParameterFields(4) = "Comprobante;" & "Comprobante :" & xComprobante & ";true"
      .ParameterFields(5) = "Glosa;" & "Glosa :" & xGlosa & ";true"
      .ParameterFields(6) = "Fecha;" & "Fecha :" & xFecha & ";true"
      ']
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
End With
     
End Sub

Function Recordset_a_Csv6(filtro As String, filtrocab As String, path As String) As Boolean
    On Error GoTo Err_function
    Dim sql As String
    Dim RstConsulta As ADODB.Recordset
    Dim RstDetalle As ADODB.Recordset
    Dim contador As Integer
    Dim xcontador As Integer
    Dim cantidaddeabonos As Integer
    Dim referencia As String
    Dim tipomoneda   As String
    Dim cuentacargo As String
    Dim importetotal As Double
    Dim checksum As Double
    Dim sumaproveedor As Double
    Dim x As Integer
    Dim cadena As String
    Set RstConsulta = New ADODB.Recordset
    Set RstDetalle = New ADODB.Recordset
    
    importetotal = 0
    checksum = 0
    sumaproveedor = 0
    cantidaddeabonos = 0
    
    ' Crea el Archivo
    Open path For Output As #1
    
    sql = "Select cobandet.codaux from cobandet"
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano"
    sql = sql & " and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban"
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and " & filtro
    sql = sql & " group by cobandet.codaux "
    
    RstConsulta.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If RstConsulta.RecordCount = 0 Then
            Exit Function
    Else
            RstConsulta.MoveFirst
            For contador = 0 To RstConsulta.RecordCount - 1
                contador = contador + 1
                RstConsulta.MoveNext
            Next
    End If
    cantidaddeabonos = contador - 1
    
    RstConsulta.Close
    
    sql = "select case cobancab.tpomon when 'N' then '0001' else '1001' end ,concat(left(globan,40),repeat(' ',40-length(left(globan,40)))),"
    sql = sql & " concat(case cobancab.tpomon when 'N' then replace(ctactemn,'-','') else replace(ctacteme,'-','') end,"
    sql = sql & " repeat(' ', 20-length(case cobancab.tpomon when 'N' then replace(ctactemn,'-','') else replace(ctacteme,'-','') end)))"
    sql = sql & " from cobancab inner join cobco on cobancab.codemp=cobco.codemp and cobancab.codbco=cobco.codbco "
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' "
    sql = sql & " and cobancab.pdoano='" & gsAnoAct & "' "
    sql = sql & " and cobancab.mespvs='" & gsMesAct & "' "
    sql = sql & " and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and " & filtrocab
    
    RstConsulta.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If RstConsulta.RecordCount = 0 Then
            Exit Function
    Else
            RstConsulta.MoveFirst
            For contador = 0 To RstConsulta.RecordCount - 1
                tipomoneda = RstConsulta.Fields(0)
                referencia = RstConsulta.Fields(1)
                cuentacargo = RstConsulta.Fields(2)
                RstConsulta.MoveNext
            Next
    End If
    
    RstConsulta.Close
    
    sql = "select sum(case cobandet.tpomon when 'N' then cobandet.impmn * IF(sgntdc=1,1,-1) else cobandet.impme * IF(sgntdc=1,1,-1) end)*(-1) from cobandet"
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano"
    sql = sql & " and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban"
    sql = sql & " inner join tgtdc on cobandet.codemp=tgtdc.codemp and cobandet.codtdc=tgtdc.codtdc"
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "'  and cobancab.pdoano='" & gsAnoAct & "'  and cobancab.mespvs='" & gsMesAct & "'  and cobancab.codbco='" & txtDato(0) & "'"
    sql = sql & " and " & filtro
    
    RstConsulta.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If RstConsulta.RecordCount = 0 Then
            Exit Function
    Else
            RstConsulta.MoveFirst
            For contador = 0 To RstConsulta.RecordCount - 1
                importetotal = importetotal + RstConsulta.Fields(0)
                RstConsulta.MoveNext
            Next
    End If
    
    
    RstConsulta.Close
    
    For x = 1 To 17 - Len(Round(importetotal, 2))
        cadena = cadena & "0"
    Next
    
    sql = " select replace(RIGHT(nroctacte,length(nroctacte)-3),'-','') from cobandet"
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano"
    sql = sql & " and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban"
    sql = sql & " inner join coctaban on cobandet.codemp=coctaban.codemp and  cobandet.codaux=coctaban.codaux and cobandet.codbco=coctaban.codbco and cobandet.tpomon=coctaban.tpomon"
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "'  and cobancab.pdoano='" & gsAnoAct & "'  and cobancab.mespvs='" & gsMesAct & "'  and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and " & filtro & " group by cobandet.codaux"
    
    RstConsulta.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If RstConsulta.RecordCount = 0 Then
            Exit Function
    Else
            RstConsulta.MoveFirst
            For contador = 0 To RstConsulta.RecordCount - 1
                checksum = checksum + RstConsulta.Fields(0)
                RstConsulta.MoveNext
            Next
    End If
    
    Print #1, "1" & Format(cantidaddeabonos, "000000") & Space(8) & "C" & tipomoneda & cuentacargo & cadena & Format(importetotal, "###0.00") & referencia & "S" & Format(checksum, "000000000000000")
    
    RstConsulta.Close
        
    sql = " select"
    sql = sql & " coctaban.tpocta,concat(replace(nroctacte,'-',''),repeat(' ',20-length(replace(nroctacte,'-','')))),"
    sql = sql & " CONCAT(cobandet.codaux,repeat(' ',12-length(cobandet.codaux))),"
    sql = sql & " CONCAT(left(tgaux.razaux,75),repeat(' ',75-length(left(tgaux.razaux,75)))),"
    sql = sql & " case cobancab.tpomon when 'N' then '0001' else '1001' end,"
    sql = sql & " sum(case cobandet.tpomon when 'N' then cobandet.impmn * IF(sgntdc=1,1,-1) else cobandet.impme * IF(sgntdc=1,1,-1) end) *(-1),cobandet.codaux "
    sql = sql & " From cobandet"
    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano"
    sql = sql & " and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban"
    sql = sql & " inner join coctaban on cobandet.codemp=coctaban.codemp and  cobandet.codaux=coctaban.codaux and cobandet.codbco=coctaban.codbco and cobandet.tpomon=coctaban.tpomon"
    sql = sql & " inner join tgaux on cobandet.codemp=tgaux.codemp and cobandet.codaux=tgaux.codaux"
    sql = sql & " inner join tgtdc on cobandet.codemp=tgtdc.codemp and cobandet.codtdc=tgtdc.codtdc"
    sql = sql & " where cobancab.codemp='" & gsCodEmp & "'  and cobancab.pdoano='" & gsAnoAct & "'  and cobancab.mespvs='" & gsMesAct & "'  and cobancab.codbco='" & txtDato(0) & "' "
    sql = sql & " and " & filtro & " group by cobandet.codaux "
       
    RstConsulta.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
    If RstConsulta.RecordCount = 0 Then
            Exit Function
    Else
            RstConsulta.MoveFirst
            For contador = 0 To RstConsulta.RecordCount - 1
            
                    sumaproveedor = RstConsulta.Fields(5)
                    
                    cadena = ""
                    
                    For x = 1 To 17 - Len(Format(sumaproveedor, "###0.00"))
                        cadena = cadena & "0"
                    Next
            
                    Print #1, "2" & RstConsulta.Fields(0) & RstConsulta.Fields(1) & "1" & "6" & RstConsulta.Fields(2) & Space(3) & RstConsulta.Fields(3) & Space(40) & Space(20) & RstConsulta.Fields(4) & cadena & Format(sumaproveedor, "###0.00") & "N"
                    
                    sql = " select case LEFT(tgtdc.abvtdc,1) when 'F' then 'E' else 'D' end,concat('0',serdoc,nrodoc),case cobandet.tpomon when 'N' then cobandet.impmn else cobandet.impme end from cobandet"
                    sql = sql & " inner join cobancab on cobandet.codemp=cobancab.codemp and cobandet.pdoano=cobancab.pdoano"
                    sql = sql & " and cobandet.mespvs=cobancab.mespvs and cobandet.coddro=cobancab.coddro and cobandet.nroban=cobancab.nroban"
                    sql = sql & " inner join tgtdc on cobandet.codemp=tgtdc.codemp and cobandet.codtdc=tgtdc.codtdc"
                    sql = sql & " where cobancab.codemp='" & gsCodEmp & "' and cobancab.pdoano='" & gsAnoAct & "' and cobancab.mespvs='" & gsMesAct & "'  and cobancab.codbco='" & txtDato(0) & "' "
                    sql = sql & " and cobandet.codaux='" & RstConsulta.Fields(6) & "'"
                    sql = sql & " and " & filtro & " order by 1 asc"
                    
                    RstDetalle.Open sql, uocnnMain, adOpenStatic, adLockOptimistic
    
                    If RstDetalle.RecordCount = 0 Then
                        Exit Function
                    Else
                        RstDetalle.MoveFirst
                        For xcontador = 0 To RstDetalle.RecordCount - 1
                        
                        cadena = ""
                    
                        For x = 1 To 17 - Len(Format(RstDetalle.Fields(2), "###0.00"))
                            cadena = cadena & "0"
                        Next
                        
                        Print #1, "3" & RstDetalle.Fields(0) & RstDetalle.Fields(1) & cadena & Format(RstDetalle.Fields(2), "###0.00")
                    
                        RstDetalle.MoveNext
                        Next
                    End If
                    
                    RstDetalle.Close
        
                RstConsulta.MoveNext
            Next
    End If
        
    Close #1
    
    Exit Function
Err_function:
    MsgBox Err.Description, vbCritical
    Close
End Function


    
