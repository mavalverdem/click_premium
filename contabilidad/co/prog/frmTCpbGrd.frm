VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTCpbGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   345
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8475
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
         Picture         =   "frmTCpbGrd.frx":0000
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
         Picture         =   "frmTCpbGrd.frx":014A
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
         Picture         =   "frmTCpbGrd.frx":0294
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
         Picture         =   "frmTCpbGrd.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "frmTCpbGrd.frx":0498
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
         Picture         =   "frmTCpbGrd.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTCpbGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
       uorstCOCta As ADODB.Recordset, _
       uorstCOCCo As ADODB.Recordset, _
       uorstTGAux As ADODB.Recordset, _
       uorstTGTDc As ADODB.Recordset, _
       uorstCOCpbDet As ADODB.Recordset, _
       uorstCOTCbMes As ADODB.Recordset, _
       uorstCOCpbDetRP As ADODB.Recordset
       
'       uorstTGArt As ADODB.Recordset, _
'       uorstTGSvc As ADODB.Recordset
']


Private Sub Form_Load()
 '[Recordsets                          'Cambiar.
'   usConnStrgSele_0 = "SELECT a.NroDoc, a.FEmDoc, b.DetCli, a.ObsDoc, Iif(a.IndAnu=0,'','Anulada') as cIndAnu," _
                    & "  a.FmaPgo, a.CodCli, a.DetCli_Doc, a.RUCCli_Doc, a.DirCli_Doc," _
                    & "  a.CodDtt_Doc, a.PctIGV, a.TotVVt, a.TotIGV, a.TotPVt," _
                    & "  a.IndAnu," _
                    & "  a.UsrCre, a.FyHCre, a.UsrMdf, a.FyHMdf " _
                    & "FROM VTFacCab a" _
                    & "  LEFT JOIN TGCli b ON a.CodCli=b.CodCli "
'   usConnStrgOrde_0 = "ORDER BY 1"
   usConnStrgSele_0 = "SELECT CodDro, NroCpb, FehCpb, GloCpb," _
                    & " If(TpoGnr=" & TPOGNR_DRO & ",'" & TPOGNR_DRO_TXT & "', If(TpoGnr=" & TPOGNR_CPR & ",'" & TPOGNR_CPR_TXT & "', If(TpoGnr=" & TPOGNR_VTA & ",'" & TPOGNR_VTA_TXT & "', If(TpoGnr=" & TPOGNR_HPR & ",'" & TPOGNR_HPR_TXT & "', If(TpoGnr=" & TPOGNR_DST & ",'" & TPOGNR_DST_TXT & "', If(TpoGnr=" & TPOGNR_DCA & ",'" & TPOGNR_DCA_TXT & "', If(TpoGnr=" & TPOGNR_APE & ",'" & TPOGNR_APE_TXT & "', '" & TPOGNR_CIE_TXT & "'))))))) as ccTpoGnr," _
                    & " TpoGnr, MesPvs, IndNCu, IndAnu, UsrCre, FyHCre, UsrMdf, FyHMdf, Concat(CodDro, NroCpb) as cLlave" _
                    & " FROM COCpbCab" _
                    & " WHERE MesPvs=" & gsMesAct & " "
'                    & "Where FehCpb>=" & gsDiaIni & "  And FehCpb<=" & gsDiaFin
'                    & "WHERE Month(FehCpb)=Val(" & gsMesAct & ") And Year(FehCpb)=Val('" & gsAnoAct & "') "
   usConnStrgOrde_0 = "ORDER BY CodDro, NroCpb"
   usConnStrgSele_1 = "SELECT COCpbDet.NroIte, COCpbDet.CodCta, COCpbDet.CodCCo, COCpbDet.CodAux, TGTDc.AbvTDc , COCpbDet.SerDoc, COCpbDet.NroDoc, COCpbDet.GloIte, " _
                    & " If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "', COCpbDet.ImpMN, 0) as cImpMN_Deb," _
                    & " If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "', 0, COCpbDet.ImpMN) as cImpMN_Hab," _
                    & " If(COCpbDet.TpoGnr=" & TPOGNR_DRO & ",'" & TPOGNR_DRO_TXT & "', If(COCpbDet.TpoGnr=" & TPOGNR_CPR & ", '" & TPOGNR_CPR_TXT & "', If(COCpbDet.TpoGnr=" & TPOGNR_VTA & ", '" & TPOGNR_VTA_TXT & "', If(COCpbDet.TpoGnr=" & TPOGNR_HPR & ",'" & TPOGNR_HPR_TXT & "', If(COCpbDet.TpoGnr=" & TPOGNR_DST & ",'" & TPOGNR_DST_TXT & "', If(COCpbDet.TpoGnr=" & TPOGNR_DCA & ",'" & TPOGNR_DCA_TXT & "', If(COCpbDet.TpoGnr=" & TPOGNR_APE & ",'" & TPOGNR_APE_TXT & "', '" & TPOGNR_CIE_TXT & "'))))))) as ccTpoGnr," _
                    & " If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "', COCpbDet.ImpME, 0) as cImpME_Deb," _
                    & " If(COCpbDet.TpoCtb='" & TPOCTB_DEB & "', 0, COCpbDet.ImpME) as cImpME_Hab," _
                    & " COCpbDet.BlqIte, COCpbDet.TpoMon, COCpbDet.ImpTcb, COCpbDet.TpoTCb, COCpbDet.TpoGnr, COCpbDet.CodTDc," _
                    & " COCpbDet.RefDoc, COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.TpoCtb, COCpbDet.TpoPvs, COCpbDet.MesPvs," _
                    & " COCpbDet.FehOpe, COCpbDet.FeEDoc, COCpbDet.FeVDoc, COCpbDet.FeRDoc, COCpbDet.ImpMN, COCpbDet.ImpME," _
                    & " Concat(COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte) as cLlave," _
                    & " COCpbDet.UsrCre, COCpbDet.FyHCre, COCpbDet.UsrMdf, COCpbDet.FyHMdf " _
                    & " FROM COCpbDet Left Join TGTDc as TGTDc On COCpbDet.CodTDc = TGTDc.CodTDc "
   usConnStrgWher_1 = " WHERE Concat(COCpbDet.CodDro, COCpbDet.NroCpb)=' '"
   usConnStrgOrde_1 = " ORDER BY COCpbDet.NroIte"
   
''SELECT CodCta, CodCCo, CodAux, TGTDc.AbvTDc, SerDoc, NroDoc, GloIte,If(TpoCtb='D',ImpMN,0) as cImpMN_Deb, If(TpoCtb='D',0,ImpMN) as cImpMN_Hab,
''   If(TpoGnr=1,'Diario', If(TpoGnr=1,'Compra', If(TpoGnr=2,'Venta', If(TpoGnr=3,'Hon.Prf.', If(TpoGnr=4,'Destino', If(TpoGnr=5,'Dif.T/C', If(TpoGnr=6,'Apertura', 'Cierre'))))))) as ccTpoGnr,
''   If(TpoCtb='D',ImpME,0) as cImpME_Deb, If(TpoCtb='D',0,ImpME) as cImpME_Hab, TpoMon, ImpTcb, TpoTCb, TpoGnr, RefDoc, CodDro, NroCpb, NroIte, TpoCtb, TpoPvs,
''   FehOpe, FeEDoc, FeVDoc, FeRDoc, ImpMN, ImpME, Concat(CodDro, NroCpb, NroIte) as cLlave, CoCpbDet.UsrCre, CoCpbDet.FyHCre, CoCpbDet.UsrMdf, CoCpbDet.FyHMdf, CoCpbDet.CodTDc
''  FROM CoCpbDet as CoCpbDet, TGTDc as TGTDc
''    WHERE Concat(CodDro, NroCpb)='0201000010 ' And CoCpbDet.CodTDc = TGTDc.CodTDc
''    Order by NroIte
    
   ' [ para uorstUltiItem
'   usCOnnStrgWher = "WHERE CodDro & NroCpb='" & frmTCpbCab.uorstMain!CodDro & frmTCpbCab.uorstMain!NroCpb & "' "
   ' ]
   
   Set uocnnMain = New ADODB.Connection
   Set uorstMain_0 = New ADODB.Recordset
   Set uorstMain_1 = New ADODB.Recordset
   Set uorstUltiItem = New ADODB.Recordset
   Set uorstCODro = New ADODB.Recordset
   Set uorstTGTCb = New ADODB.Recordset
   Set uorstCOCta = New ADODB.Recordset
   Set uorstCOCCo = New ADODB.Recordset
   Set uorstTGAux = New ADODB.Recordset
   Set uorstTGTDc = New ADODB.Recordset
   Set uorstCOCpbDet = New ADODB.Recordset
   Set uorstCOTCbMes = New ADODB.Recordset
   Set uorstCOCpbDetRP = New ADODB.Recordset
'   Set uorstTGArt = New ADODB.Recordset
'   Set uorstTGSvc = New ADODB.Recordset
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
      .Properties("Unique Table").Value = "COCpbCab"
   End With
   With uorstMain_1
      .ActiveConnection = uocnnMain
      .Source = usConnStrgSele_1 & usConnStrgWher_1 & usConnStrgOrde_1
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "COCpbDet"
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
'      .Source = "SELECT CodDro, DetDro, Cpb" & gsMesAct & "," _
'              & "  CodDro as cLlave " _
'              & "FROM CODro"
      .Source = "SELECT CodDro, DetDro, Cpb" & gsMesAct & " " _
              & "FROM CODro Where Length(CodDro) > 2"
''     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With uorstTGTCb
      .ActiveConnection = uocnnMain
      .Source = "SELECT FehTCb, ImpTCb_Cpr, ImpTCb_Vta " _
              & "FROM TGTCb"
'              & "WHERE Month(FehTCb)=" & Val(gsMesAct) & " AND Year(FehTCb)=" & Val(gsAnoAct)
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCOCta
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodCta, DetCta, TpoTCb, TpoAnl, IndAjd, IndCCo, IndDoc, CodCta_AjD_Deb, CodCta_AjD_Hab " _
              & "FROM COCta " _
              & "WHERE TpoCta = " & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCOCCo
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodCCo, DetCCo " _
              & "FROM COCCo " _
              & "WHERE EstCCo='" & ESTCCO_ACT & "' AND Length(CodCCo) > 2"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstTGAux
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodAux, RazAux " _
              & "FROM TGAux " _
              & "WHERE EstAux='" & ESTAUX_ACT & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstTGTDc
      .ActiveConnection = uocnnMain
      .Source = "SELECT CodTDc, DetTDc " _
              & "FROM TGTDc"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With uorstCOCpbDet
      .ActiveConnection = uocnnMain
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
   End With
   With uorstCOTCbMes
      .ActiveConnection = uocnnMain
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
   End With
   With uorstCOCpbDetRP
      .ActiveConnection = uocnnMain
      .Source = "SELECT MesPvs, CodDro, NroCpb, NroIte, CodAux, CodCta, CodTDc, SerDoc, NroDoc, " _
              & "  CodTDc_RtcPcp, SerDoc_RtcPcp, NroDoc_RtcPcp, " _
              & "  Concat(MesPvs, CodDro, NroCpb, NroIte) as cLlave, " _
              & "  UsrCre, FyHCre, UsrMdf, FyHMdf " _
              & "FROM COCpbDetRP "
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "COCpbDetRP"
   End With
'   With uorstTGArt
'      .ActiveConnection = uocnnMain
'      .Source = "SELECT CodArt, DetArt " _
'              & "FROM TGArt"
''     .CursorLocation = adUseClient   'Es el Default.
'      .CursorType = adOpenDynamic
'      .LockType = adLockReadOnly
'      .Open
'   End With
'   With uorstTGSvc
'      .ActiveConnection = uocnnMain
'      .Source = "SELECT CodSvc, DetSvc " _
'              & "FROM TGSvc"
''     .CursorLocation = adUseClient   'Es el Default.
'      .CursorType = adOpenDynamic
'      .LockType = adLockReadOnly
'      .Open
'   End With
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
      .Source = usConnStrgSele_1 & " Where COCpbDet.CodDro='    ' " & usConnStrgOrde_1
      .Open
      .Properties("Unique Table").Value = "COCpbDet"
   End With
'   gpTUg_Nuevo Me, frmTFacCab          'Cambiar Formulario de Datos.
   gpTUg_Nuevo Me, frmTCpbCab          'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
   On Error GoTo Err

   'Verificación de existencia de ítemes.
   If uorstMain_0.RecordCount = 0 Then
      MsgBox TEXT_8001, vbCritical
      Exit Sub
   End If

   With frmTCpbCab                     'Cambiar Formulario de Datos.
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
      MsgBox "No hay datos creados.", vbCritical
      Exit Sub
   End If

   If frmTCpbGrd.uorstMain_0!TpoGnr <> TPOGNR_DRO Then
      MsgBox "No se Puede Eliminar este Asiento", vbInformation
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
   frmTCpbGrd.uorstMain_0.Requery
   frmTCpbGrd.ppDatosGrid
   dgrMain.SetFocus
End Sub

Public Sub cmdImprimir_Click()
 '[Datos del formulario de impresión.  'Cambiar.
   frmLCpb.Caption = "Listado de " & Me.Caption
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

Public Sub ppDatosGrid()               'Cambiar Datos Grid.
   Dim dnNum As Integer
         
   With dgrMain.Columns
      For dnNum = 0 To .Count - 1
         Select Case dnNum
         Case 0
            .Item(dnNum).Caption = "Diario"
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("CodDro").DefinedSize + 2)
        Case 1
            .Item(dnNum).Caption = "NºComp."
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("NroCpb").DefinedSize + 2)
         Case 2
'            .Item(dnNum).Caption = "Cliente"
'            .Item(dnNum).Width = 3000
            .Item(dnNum).Caption = "Fecha"
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("FehCpb").DefinedSize + 4)
         Case 3
'            .Item(dnNum).Caption = "Observación"
'            .Item(dnNum).Width = 2200
            .Item(dnNum).Caption = "Glosa"
            .Item(dnNum).Width = 100 * (uorstMain_0.Fields("GloCpb").DefinedSize - 16)
         Case 4
'            .Item(dnNum).Caption = "Anulada"
'            .Item(dnNum).Width = 850
            .Item(dnNum).Caption = "Tipo"
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
   cmdImprimir.Enabled = IIf(taOpciones(2) Or taOpciones(3), True, False)
End Property

