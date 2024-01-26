VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRRegCpr 
   Caption         =   "[título]"
   ClientHeight    =   3075
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDiario 
      Caption         =   "Totaliza Diario"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3195
      TabIndex        =   5
      Top             =   765
      Width           =   1335
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Diario"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   6
      Top             =   930
      Width           =   4530
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4080
         Picture         =   "frmRRegCpr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
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
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   255
         Width           =   780
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
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   255
         Width           =   3240
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   12
      Top             =   1680
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   45
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6885
         Picture         =   "frmRRegCpr.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   255
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
         TabIndex        =   3
         Top             =   255
         Width           =   1260
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
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   4
         Top             =   255
         Width           =   5520
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmRRegCpr.frx":0354
      Left            =   6180
      List            =   "frmRRegCpr.frx":0356
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   900
      Width           =   1125
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7335
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2460
      Width           =   7335
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
         Height          =   570
         Index           =   0
         Left            =   0
         Picture         =   "frmRRegCpr.frx":0358
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Height          =   570
         Index           =   1
         Left            =   1245
         Picture         =   "frmRRegCpr.frx":088A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1125
      End
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
         Height          =   570
         Left            =   2355
         TabIndex        =   0
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
         Height          =   570
         Left            =   4920
         Picture         =   "frmRRegCpr.frx":098C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
      Begin MSComctlLib.Toolbar toolbar 
         Height          =   600
         Left            =   3600
         TabIndex        =   20
         Top             =   0
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1058
         ButtonWidth     =   1323
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exportar"
               Object.ToolTipText     =   "Exportar Registro de Documentos a Excel"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A1"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A2"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A3"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   1080
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegCpr.frx":0AD6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegCpr.frx":0C30
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegCpr.frx":0D8A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegCpr.frx":114C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegCpr.frx":1816
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   0
      Left            =   5325
      TabIndex        =   9
      Top             =   945
      Width           =   765
   End
End
Attribute VB_Name = "frmRRegCpr"
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
Private porstTGAux As ADODB.Recordset
Private porstCodro As ADODB.Recordset


']
Private Sub chkDiario_Click()
  fraRangos.Enabled = (chkDiario.Value = vbUnchecked)
  txtDato(1).Text = IIf(chkDiario.Value = vbChecked, "", txtDato(1).Text)
  lblDatoDeta(1).Caption = IIf(chkDiario.Value = vbChecked, "", lblDatoDeta(1).Caption)
End Sub

Private Sub pExporta(TpoRpt As Integer)
'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err
 
    Dim oProgress As New frmzProgressBar
    oProgress.Show
    oProgress.pgbProgreso.Value = 0: oProgress.pgbProgreso.Min = 0
    oProgress.pgbProgreso.Max = 7
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Min
    oProgress.Caption = "Procesando Compras"
 

    Dim pocnnTmp As ADODB.Connection '2014-04-14 Query timeout expired
    Set pocnnTmp = New ADODB.Connection '2014-04-14 Query timeout expired
    With pocnnTmp
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
         Dim cCadReporte  As String
         Dim sTabla As String
         sTabla = "xlsCprCta"
         pocnnTmp.Execute fDropTable2(sTabla, 1)
         
        cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
        cCadReporte = cCadReporte & "SELECT d.codemp,d.pdoano,c.MesPvs,d.codaux,d.codtdc, "
        cCadReporte = cCadReporte & "    d.serdoc , d.nrodoc, d.tpocnc, d.codcta "
        cCadReporte = cCadReporte & "FROM cocprdoccta d "
        cCadReporte = cCadReporte & "LEFT JOIN COCprDoc c "
        cCadReporte = cCadReporte & "    ON d.codemp=c.codemp AND d.pdoano=c.pdoano "
        cCadReporte = cCadReporte & "    AND d.codaux=c.codaux AND d.codtdc=c.codtdc "
        cCadReporte = cCadReporte & "    AND d.serdoc=c.serdoc AND d.nrodoc=c.nrodoc "
        'cCadReporte = cCadReporte & "WHERE  d.codemp='" & gsCodEmp & "' AND     concat(d.pdoano,c.MesPvs) = '" & gsAnoAct & gsMesAct & "' "
        cCadReporte = cCadReporte & "WHERE  d.codemp='" & gsCodEmp & "' "
       ' cCadReporte = cCadReporte & " AND     concat(d.pdoano,c.MesPvs) = '" & gsAnoAct & gsMesAct & "' "
       If TpoRpt = 1 Then
            cCadReporte = cCadReporte & " AND   concat(d.pdoano,c.MesPvs) = '" & gsAnoAct & gsMesAct & "' "
        ElseIf TpoRpt = 2 Then
            cCadReporte = cCadReporte & " AND   concat(d.pdoano,c.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
             cCadReporte = cCadReporte & "    concat(d.pdoano,c.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        Else
             cCadReporte = cCadReporte & " AND   concat(d.pdoano,c.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        End If
        
        cCadReporte = cCadReporte & "      AND d.tpocnc='11' "
        
        pocnnTmp.Execute cCadReporte
      
         sTabla = "xlsCprCab"
        'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & sTabla & " ", cCadReporte)
         pocnnTmp.Execute fDropTable2(sTabla, 1)

        cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
        cCadReporte = cCadReporte & "SELECT"
        cCadReporte = cCadReporte & "    CONCAT(a.pdoano,a.mespvs,'00') as CPERIODO,"
        cCadReporte = cCadReporte & "    concat(a.CodDro,a.NroCpb) as CNUMREGOPE,"
        cCadReporte = cCadReporte & "    date_format(a.FeEDoc,'%d/%m/%Y')as CFECCOM,"
        cCadReporte = cCadReporte & "    date_format(a.FevDOC,'%d/%m/%Y')as CFECVENPAG,"
        cCadReporte = cCadReporte & "    b.CodTDc AS CTIPDOCCOM,"
        '#IF(b.CodTDc<>'50',a.serdoc,mid(a.serdoc,2,3)) AS CNUMSER,
        cCadReporte = cCadReporte & "    IFNULL(case b.codtdc when '50' then a.codaduana when '52' then a.codaduana when '53' then a.codaduana else a.serdoc end, '-') AS CNUMSER,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.annodua,''),a.annodua,'0') as CEMIDUADSI,"
        '#a.NroDoc AS CNUMDCODFV,
        cCadReporte = cCadReporte & "    IFNULL(CASE b.codtdc when '50' then a.nrodua when '52' then a.nrodua when '53' then a.nrodua else a.nrodoc END, '') AS CNUMDCODFV,"
        cCadReporte = cCadReporte & "    '0' AS COSDCREFIS, MID(c.tpodci,2,1) AS CTIPDIDPRO,c.codaux AS CNUMDIDPRO,"
        
        cCadReporte = cCadReporte & "    replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        cCadReporte = cCadReporte & "    ifnull(MID(c.RazAux,1,60)  ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
        cCadReporte = cCadReporte & "    as CNOMRSOPRO,"
        
        '#IF((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1))<>0.00,(a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),'0.00') AS CBASIMPGRA
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS CBASIMPGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS CIGVGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CBASIMPGNG,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIGVGRANGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CBASIMPSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_ONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CIGVSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpExo_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CIMPTOTNGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpISC_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CISC,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOIm_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS COTRTRICGO,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpTot_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIMPTOTCOM,"
        cCadReporte = cCadReporte & "    format(a.imptcb,3) * 1  AS CTIPCAM,"
        cCadReporte = cCadReporte & "    IF(ifnull(codtdc_ref,''),date_format(feedoc_ref,'%d/%m/%Y'),'01/01/0001') as CFECCOMMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.codtdc_ref,''),a.codtdc_ref,'00')as CTIPCOMMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.serdoc_ref,''),a.serdoc_ref,'-') as CNUMSERMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.nrodoc_ref,''),a.nrodoc_ref,'-') as CNUMCOMMOD,"
        
        cCadReporte = cCadReporte & "    CASE WHEN a.codtdc_ref='91' THEN Concat(a.serdoc_ref, '-', a.nrodoc_ref) ELSE '-' END as CCOMNODOMI,"
        
        cCadReporte = cCadReporte & "    IF(ifnull(a.NroCDt,''),date_format(a.FehCDt,'%d/%m/%Y'),'01/01/0001') as CEMIDEPDET,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.NroCDt,''),a.NroCDt,'0')    as CNUMDEPDET,"
        
        cCadReporte = cCadReporte & "    ifnull(a.INDRETEN,'')   AS CCOMPGRET,"
        
        cCadReporte = cCadReporte & "    if(MONTH(a.FeEDoc)=a.mespvs,'1','6') as CESTOPE,"
        
        cCadReporte = cCadReporte & "    '0.00' AS CVALFACIMP ,"
        
        cCadReporte = cCadReporte & "    '' AS CINTDIAMAY,"
        cCadReporte = cCadReporte & "    '' AS CINTKARDEX,"
        cCadReporte = cCadReporte & "    '' AS CINTREG, "
        cCadReporte = cCadReporte & "    tsadetrac "
        cCadReporte = cCadReporte & "    ,'' xCol1 " '2015-05-14
        cCadReporte = cCadReporte & "    ,'' xCol2 " '2015-05-14
        cCadReporte = cCadReporte & "    ,a.tpomon " '2015-05-14
        
        cCadReporte = cCadReporte & "    ,replace(format((a.ImpTot_ME * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIMPTOTMEX "  '2015-07-13 adici vta/cpr me
        
        cCadReporte = cCadReporte & "    ,GloDoc " '2015-06-04 adicion glodoc
    
        cCadReporte = cCadReporte & "    ,a.codaux, a.codtdc, a.serdoc, a.NroDoc " '2015-07-15 adicion pgo segun diario

        cCadReporte = cCadReporte & "    ,ifnull(refdoc,'') refdoc " '2015-12-17 adicion ref
        cCadReporte = cCadReporte & "    ,e.codcta codcta " '2016-07-14 correcion error duplica doc
      
        cCadReporte = cCadReporte & "FROM ((((COCprDoc a "
        cCadReporte = cCadReporte & "LEFT JOIN TGTDc b on a.codemp=b.codemp and b.CodTDc = a.CodTDc) "
        cCadReporte = cCadReporte & "LEFT JOIN TGAux c on a.codemp=c.codemp and c.CodAux = a.CodAux) "
        cCadReporte = cCadReporte & "LEFT JOIN CODro d ON a.codemp=d.codemp and a.pdoano=d.pdoano and a.CodDro=d.CodDro) "
'ini 2016-07-14 duplica documento
        cCadReporte = cCadReporte & "LEFT JOIN xlsCprCta e "
        cCadReporte = cCadReporte & "ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.MesPvs=e.MesPvs "
        cCadReporte = cCadReporte & "AND a.codaux=e.codaux AND a.codtdc=e.codtdc AND a.serdoc=e.serdoc AND a.nrodoc=e.nrodoc) "
'fin 2016-07-14 duplica documento
        
        'cCadReporte = cCadReporte & "WHERE a.codemp='001' and a.pdoano='2012' and a.MesPvs >= '01' AND  a.MesPvs <= '04' "
        cCadReporte = cCadReporte & "WHERE "
        cCadReporte = cCadReporte & "    a.codemp='" & gsCodEmp & "' AND "
        
        'ini 2015-03-24
'        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
'        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & Left(cmbEjercicio.Text, 2) & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        If TpoRpt = 1 Then
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' "
'''        Else
        ElseIf TpoRpt = 2 Then
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
             cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        Else
'            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
             cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        End If
        'fin 2015-03-24
        
'ini 2015-01-09 adiciona ruc
      If Trim(txtDato(0).Text) <> "" Then
          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato(0).Text) & "' "
        End If
'fin 2015-01-09 adiciona ruc
        
        cCadReporte = cCadReporte & "ORDER BY a.pdoano, a.mespvs ,a.CodDro, a.NroCpb ASC "

    
    pocnnTmp.Execute cCadReporte
   oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
   DoEvents
   
'ini 2015-07-15 adicion pgo segun diario
    sTabla = "tmp_xls_pdte"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
'saldos de documento segun reporte historico
    cCadReporte = cCadReporte & "SELECT "
    cCadReporte = cCadReporte & "    a.pdoano AS cAno, a.MesPvs, a.CodCta, a.CodAux,a.CodTDc,"
    cCadReporte = cCadReporte & "    a.SerDoc, a.NroDoc, a.CodDro, a.NroCpb, Null AS codcco,"
    'cCadReporte = cCadReporte & "    Null AS detcco, CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocum, a.FehOpe, a.FeEDoc, a.FeVDoc,"
    cCadReporte = cCadReporte & "    Null AS detcco,"
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocum, "
    cCadReporte = cCadReporte & "a.FehOpe, a.FeEDoc, a.FeVDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, b.RazAux, "
    
    'cCadReporte = cCadReporte & "    a.RefDoc, a.GloIte AS GloIte, b.RazAux, (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
    cCadReporte = cCadReporte & "(CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) AS cDebeMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) AS cHaberMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END) AS cDebeME, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) AS cHaberME "
    
'    cCadReporte = cCadReporte & "    (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END) AS cDebeMN, (CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END) AS cHaberMN,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END) AS cDebeME, (CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END) AS cHaberME"
    cCadReporte = cCadReporte & "    ,a.TpoPvs"
    cCadReporte = cCadReporte & "    ,CONCAT(year(a.FehOpe),'-',LPAD(month(a.FehOpe),2,'0'),'-',LPAD(day(a.FehOpe),2,'0'),'-',a.CodDro,'-',a.NroCpb,'-',a.Nroite) AS x_clave "
    cCadReporte = cCadReporte & "FROM ((((COCpbDet a "
    cCadReporte = cCadReporte & "    LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    cCadReporte = cCadReporte & "    LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    cCadReporte = cCadReporte & "    LEFT JOIN Cocta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.Codcta=d.Codcta) "
    cCadReporte = cCadReporte & "    LEFT JOIN CoCCo e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.codcco=e.codcco) "
    cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
    cCadReporte = cCadReporte & "    AND a.pdoano='" & gsAnoAct & "' "
    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta, 2)>='01' AND LEFT(a.codcta, 2)<='FF' "
    cCadReporte = cCadReporte & "    AND (a.ImpMN<> 0.00 OR a.ImpME<> 0.00) "
    'cCadReporte = cCadReporte & "    AND a.Mespvs <='03' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND a.Mespvs <='" & gsMesAct & "' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND IFNULL(a.NroDoc, '') <>'' AND d.inddoc='1' "
    'cCadReporte = cCadReporte & "    # AND a.CodAux='10097267265'"
    cCadReporte = cCadReporte & "    AND a.TpoPvs='" & TPOPVS_CAN & "' " 'TPOPVS_CAN
    'cCadReporte = cCadReporte & "    AND a.TpoPvs='C' " 'TPOPVS_CAN
'2016-03-14 erro cuenta 422,428 se mezclan
    cCadReporte = cCadReporte & "    AND LEFT(a.codcta,3) = '421' "
    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta,3) <> '422' " '2015-09-03 erro cuenta anticipo duplicado
'2016-03-14 erro cuenta 422,428 se mezclan
    
'ini 2015-01-09 adiciona ruc
      If Trim(txtDato(0).Text) <> "" Then
          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato(0).Text) & "' "
        End If
'fin 2015-01-09 adiciona ruc
    
    cCadReporte = cCadReporte & "ORDER BY a.codcta, a.codaux, a.codtdc, a.serdoc, a.NroDoc, a.TpoPvs, a.MesPvs, a.FehOpe "
    
    
    pocnnTmp.Execute cCadReporte
   oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
   DoEvents

'*********************
    sTabla = "tmp_xls_pdte2"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
cCadReporte = cCadReporte & "SELECT "
cCadReporte = cCadReporte & "    codcta,codaux,cdocum,min(x_clave) x_clave "
cCadReporte = cCadReporte & "From tmp_xls_pdte "
cCadReporte = cCadReporte & "GROUP BY codcta, codaux,cdocum # x_clave "
cCadReporte = cCadReporte & "ORDER BY codcta, codaux,cdocum #x_clave "

    pocnnTmp.Execute cCadReporte
   oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
   DoEvents

'*********************
    sTabla = "tmp_xls_pdte3"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")

cCadReporte = cCadReporte & "SELECT "
cCadReporte = cCadReporte & "* "
cCadReporte = cCadReporte & "From tmp_xls_pdte "
cCadReporte = cCadReporte & "Where x_clave "
cCadReporte = cCadReporte & "    IN (select x_clave from tmp_xls_pdte2) "
    pocnnTmp.Execute cCadReporte
   oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
   DoEvents
    
'*********************
    
'fin 2015-07-15 adicion pgo segun diario
    
    
'ini exporta datos a excel

    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
    With porstTmp
       .ActiveConnection = pocnnTmp
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
       
'       .Source = "SELECT * FROM " & ps_Prefijo & sTabla

'ini 2015-07-15 adicion pgo segun diario
.Source = "SELECT "
.Source = .Source & "    CPERIODO,CNUMREGOPE,CFECCOM,CFECVENPAG,CTIPDOCCOM,"
.Source = .Source & "    CNUMSER,CEMIDUADSI,CNUMDCODFV,COSDCREFIS,CTIPDIDPRO,"
.Source = .Source & "    CNUMDIDPRO,CNOMRSOPRO,CBASIMPGRA,CIGVGRA,CBASIMPGNG,"
.Source = .Source & "    CIGVGRANGV,CBASIMPSCF,CIGVSCF,CIMPTOTNGV,CISC,"
.Source = .Source & "    COTRTRICGO,CIMPTOTCOM,CTIPCAM,CFECCOMMOD,CTIPCOMMOD,"
.Source = .Source & "    CNUMSERMOD,CNUMCOMMOD,CCOMNODOMI,CEMIDEPDET,CNUMDEPDET,"
.Source = .Source & "    CCOMPGRET,CESTOPE,CVALFACIMP,CINTDIAMAY,CINTKARDEX,"
.Source = .Source & "    CINTREG,tsadetrac,xCol1,xCol2,tpomon,"
.Source = .Source & "    CIMPTOTMEX,GloDoc,"
.Source = .Source & "    a.refdoc," '2015-12-17 adicion ref
'#   a.*,,b.FehOpe
.Source = .Source & "    b.FehOpe,"
.Source = .Source & "    IFNULL(b.cDebeMN,0)-IFNULL(b.cHaberMN,0) PgoMN,"
.Source = .Source & "    IFNULL(b.cDebeME,0)-IFNULL(b.cHaberME,0) PgoME "
.Source = .Source & "FROM xlsCprCab a "
.Source = .Source & "LEFT JOIN tmp_xls_pdte3 b "
'2016-07-14 duplica doc .Source = .Source & "    ON a.CodAux=b.CodAux and a.CTIPDOCCOM=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc "
.Source = .Source & "    ON a.codcta=b.codcta and a.CodAux=b.CodAux and a.CTIPDOCCOM=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc "

'fin 2015-07-15 adicion pgo segun diario
       
       .Open
   oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
   DoEvents
       
    End With

    Dim xArchPeriodo As String
    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
 
    'Set oSheet = oWBook.Worksheets(1)
 

    '*Set oExcel = New Excel.Application
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

    Set oWBook = oExcel.Workbooks.Add
    '*Set oWBook = oExcel.Workbooks.Open(dlbDirectorio(0).path & xArchPeriodo, , True) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    '*Set oSheet = oWBook.Worksheets("Clientes")
     Set oSheet = oWBook.Worksheets(1)
    '*oExcel.Visible = True

    With oSheet
    
'        oSheet.Select
'        Columns("M:V").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"

        oSheet.Select
        
        .Cells(1, 1).Value = "Registro de Compras"
        
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        
        .Cells(nRowI, 1).Value = "Registro de Compras"
        nRowI = nRowI + 2
        Dim x1 As Integer
        .Cells(nRowI, 1).Value = "Periodo"
        .Cells(nRowI, 2).Value = "Nº Reg."
        .Cells(nRowI, 3).Value = "F.Cmpra"
        .Cells(nRowI, 4).Value = "F. Pago"
        .Cells(nRowI, 5).Value = "T.Doc"
        .Cells(nRowI, 6).Value = "Serie"
        .Cells(nRowI, 7).Value = "CemiDuadsi"
        .Cells(nRowI, 8).Value = "Nº Doc."
        .Cells(nRowI, 9).Value = "COSDCREFIS"
        .Cells(nRowI, 10).Value = "T.Prv"
        .Cells(nRowI, 11).Value = "RUC"
        .Cells(nRowI, 12).Value = "R.Social"
        .Cells(nRowI, 13).Value = "B. Gravada"
        .Cells(nRowI, 14).Value = "IGV Grab"
        .Cells(nRowI, 15).Value = "B. G/N Gr"
        .Cells(nRowI, 16).Value = "IGV G/N Gr"
        .Cells(nRowI, 17).Value = "B. Sin CF"
        .Cells(nRowI, 18).Value = "Igv S CF"
        .Cells(nRowI, 19).Value = "CIMPTOTNGV"
        .Cells(nRowI, 20).Value = "CISSC"
        .Cells(nRowI, 21).Value = "COTRTRICGO"
        .Cells(nRowI, 22).Value = "CIMPTOTCOM"
        .Cells(nRowI, 23).Value = "CTIPCAM"
        .Cells(nRowI, 24).Value = "CFECCOMMOD"
        .Cells(nRowI, 25).Value = "CTIPCOMMOD"
        .Cells(nRowI, 26).Value = "CNUMSERMOD"
        .Cells(nRowI, 27).Value = "CNUMCOMMOD"
        .Cells(nRowI, 28).Value = "CCOMNODOMI"
        .Cells(nRowI, 29).Value = "CEMIDEPDET"
        .Cells(nRowI, 30).Value = "CNUMDEPDET"
        .Cells(nRowI, 31).Value = "CCOMPGRET"
        .Cells(nRowI, 32).Value = "CESTOPE"
        .Cells(nRowI, 33).Value = "CVALFACIMP"
        .Cells(nRowI, 34).Value = "CINTDIAMAY"
        .Cells(nRowI, 35).Value = "CINTKARDEX"
        .Cells(nRowI, 36).Value = "CINTREG"
        .Cells(nRowI, 37).Value = "tsadetrac"
        .Cells(nRowI, 38).Value = "DetaDetrac"
        .Cells(nRowI, 39).Value = "PorcDetra"
        .Cells(nRowI, 40).Value = "TpoMon"
        .Cells(nRowI, 41).Value = "Total ME"
        .Cells(nRowI, 42).Value = "Glosa"
        .Cells(nRowI, 43).Value = "Refer."
        .Cells(nRowI, 43 + 1).Value = "F.Pago"
        .Cells(nRowI, 44 + 1).Value = "1er Pgo MN"
        .Cells(nRowI, 45 + 1).Value = "1er Pgo ME"
     
        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
    oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
   DoEvents
       
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        .Columns.AutoFit ' ajusta el ancho de las columnas
        'Sheets(oSheet).Select
        
        'hay sale error definido por la aplicacion o el objeto 1004, cuando aplico estos comandos Select y NumberFormat
'        oSheet.Select
'solo sale error en esta        Columns("M:M").Select
'        Selection.NumberFormat = "#,##0.00"

'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("O:O").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("P:P").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("Q:Q").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("R:R").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("S:S").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("T:T").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("U:U").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("V:V").Select
'        Selection.NumberFormat = "#,##0.00"
        
        'crear tabla temporal
        'Dim xpocnnMain As ADODB.Connection
        'Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next
'        oSheet.Select
'        Cells(1, 1).Select

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 2).Value)
'        Next


 'ini 2015-07-02 adic tabla detrac
'*********************************
        Dim uorstcodetrac As ADODB.Recordset
        Set uorstcodetrac = New ADODB.Recordset
        Set uorstcodetrac = fRstDetrac(pocnnMain, uorstcodetrac)
       xrow1 = nRowI
        Dim nContador As Integer
        Dim s_Contenido As String
        Dim n_Detraccion As Double
        Dim s_detalle As String
        Do While Len(Trim(.Cells(xrow1, 2).Value)) <> 0
            s_Contenido = Left(.Cells(xrow1, 37).Value, 5)
            With uorstcodetrac
                If .RecordCount > 0 Then .MoveFirst
                    .Find "coddetrac='" & s_Contenido & "'"
                    If Not .EOF Then
                        oSheet.Cells(xrow1, 38).Value = !coddetrac
                        '2015-07-08 cambio de decima a % oSheet.Cells(xrow1, 39).Value = !pctdetrac * 100
                        oSheet.Cells(xrow1, 39).Value = !pctdetrac
                   End If
            End With
            xrow1 = xrow1 + 1
        Loop
        
        uorstcodetrac.Close
        Set uorstcodetrac = Nothing
        
    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing


'fin exporta datos a excel
  pocnnMain.Execute fDropTable("xlsCprCta", 1)
  pocnnMain.Execute fDropTable("xlsCprCab", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte2", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte3", 1)
  oProgress.pgbProgreso.Value = oProgress.pgbProgreso.Value + 1
   DoEvents
  
  Unload oProgress          ' Unload progress bar window

'fDropTable
   porstTmp.Close
   pocnnTmp.Close
   Set porstTmp = Nothing
   Set pocnnTmp = Nothing
   
  Exit Sub
Err:
  Unload oProgress          ' Unload progress bar window
    MsgBox (TEXT_6001)
  If pocnnTmp.State = adStateOpen Then
    porstTmp.Close
    pocnnTmp.Close
    Set porstTmp = Nothing
    Set pocnnTmp = Nothing
  End If
End Sub
Private Sub pExporta_2016_07_14(TpoRpt As Integer)
'TpoRpt=1 Del mes
'TpoRpt=2 Al mes
 On Error GoTo Err
 
'''ini 2016-07-13 correcion registro de compras
'    Dim xprt_resumen As Integer
'    xprt_resumen = 0
'    'If MsgBox(TEXT_1021 & " Desea Registro de Ventas [Si] o Detalle de Pagos [No]? ", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
'    'If MsgBox(" Desea Registro de Ventas [Si] o Detalle de Pagos [No]? ", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
'    If MsgBox(" Desea Registro de Ventas [Si] o Detalle de Pagos [No]? ", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
'        xprt_resumen = 0
'    Else
'        xprt_resumen = 1
'    End If
'''fin 2016-07-13 correcion registro de compras

    Dim pocnnTmp As ADODB.Connection '2014-04-14 Query timeout expired
    Set pocnnTmp = New ADODB.Connection '2014-04-14 Query timeout expired
    With pocnnTmp
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
    
    Dim cCadReporte  As String
    Dim sTabla As String
    sTabla = "xlsCprCab"
   'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & sTabla & " ", cCadReporte)
    pocnnTmp.Execute fDropTable2(sTabla, 1)

        cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
        cCadReporte = cCadReporte & "SELECT"
        cCadReporte = cCadReporte & "    CONCAT(a.pdoano,a.mespvs,'00') as CPERIODO,"
        cCadReporte = cCadReporte & "    concat(a.CodDro,a.NroCpb) as CNUMREGOPE,"
        cCadReporte = cCadReporte & "    date_format(a.FeEDoc,'%d/%m/%Y')as CFECCOM,"
        cCadReporte = cCadReporte & "    date_format(a.FevDOC,'%d/%m/%Y')as CFECVENPAG,"
        cCadReporte = cCadReporte & "    b.CodTDc AS CTIPDOCCOM,"
        '#IF(b.CodTDc<>'50',a.serdoc,mid(a.serdoc,2,3)) AS CNUMSER,
        cCadReporte = cCadReporte & "    IFNULL(case b.codtdc when '50' then a.codaduana when '52' then a.codaduana when '53' then a.codaduana else a.serdoc end, '-') AS CNUMSER,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.annodua,''),a.annodua,'0') as CEMIDUADSI,"
        '#a.NroDoc AS CNUMDCODFV,
        cCadReporte = cCadReporte & "    IFNULL(CASE b.codtdc when '50' then a.nrodua when '52' then a.nrodua when '53' then a.nrodua else a.nrodoc END, '') AS CNUMDCODFV,"
        cCadReporte = cCadReporte & "    '0' AS COSDCREFIS, MID(c.tpodci,2,1) AS CTIPDIDPRO,c.codaux AS CNUMDIDPRO,"
        
        cCadReporte = cCadReporte & "    replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        cCadReporte = cCadReporte & "    ifnull(MID(c.RazAux,1,60)  ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
        cCadReporte = cCadReporte & "    as CNOMRSOPRO,"
        
        '#IF((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1))<>0.00,(a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),'0.00') AS CBASIMPGRA
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS CBASIMPGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS CIGVGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CBASIMPGNG,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIGVGRANGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CBASIMPSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_ONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CIGVSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpExo_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CIMPTOTNGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpISC_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS CISC,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOIm_MN * IF(b.SgnTDc = 0, -1,1)),2),',','')  * 1 AS COTRTRICGO,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpTot_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIMPTOTCOM,"
        cCadReporte = cCadReporte & "    format(a.imptcb,3) * 1  AS CTIPCAM,"
        cCadReporte = cCadReporte & "    IF(ifnull(codtdc_ref,''),date_format(feedoc_ref,'%d/%m/%Y'),'01/01/0001') as CFECCOMMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.codtdc_ref,''),a.codtdc_ref,'00')as CTIPCOMMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.serdoc_ref,''),a.serdoc_ref,'-') as CNUMSERMOD,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.nrodoc_ref,''),a.nrodoc_ref,'-') as CNUMCOMMOD,"
        
        cCadReporte = cCadReporte & "    CASE WHEN a.codtdc_ref='91' THEN Concat(a.serdoc_ref, '-', a.nrodoc_ref) ELSE '-' END as CCOMNODOMI,"
        
        cCadReporte = cCadReporte & "    IF(ifnull(a.NroCDt,''),date_format(a.FehCDt,'%d/%m/%Y'),'01/01/0001') as CEMIDEPDET,"
        cCadReporte = cCadReporte & "    IF(ifnull(a.NroCDt,''),a.NroCDt,'0')    as CNUMDEPDET,"
        
        cCadReporte = cCadReporte & "    ifnull(a.INDRETEN,'')   AS CCOMPGRET,"
        
        cCadReporte = cCadReporte & "    if(MONTH(a.FeEDoc)=mespvs,'1','6') as CESTOPE,"
        
        cCadReporte = cCadReporte & "    '0.00' AS CVALFACIMP ,"
        
        cCadReporte = cCadReporte & "    '' AS CINTDIAMAY,"
        cCadReporte = cCadReporte & "    '' AS CINTKARDEX,"
        cCadReporte = cCadReporte & "    '' AS CINTREG, "
        cCadReporte = cCadReporte & "    tsadetrac "
        cCadReporte = cCadReporte & "    ,'' xCol1 " '2015-05-14
        cCadReporte = cCadReporte & "    ,'' xCol2 " '2015-05-14
        cCadReporte = cCadReporte & "    ,a.tpomon " '2015-05-14
        
        cCadReporte = cCadReporte & "    ,replace(format((a.ImpTot_ME * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1  AS CIMPTOTMEX "  '2015-07-13 adici vta/cpr me
        
        cCadReporte = cCadReporte & "    ,GloDoc " '2015-06-04 adicion glodoc
    
        cCadReporte = cCadReporte & "    ,a.codaux, a.codtdc, a.serdoc, a.NroDoc " '2015-07-15 adicion pgo segun diario

        cCadReporte = cCadReporte & "    ,ifnull(refdoc,'') refdoc " '2015-12-17 adicion ref
      
        cCadReporte = cCadReporte & "FROM (((COCprDoc a "
        cCadReporte = cCadReporte & "LEFT JOIN TGTDc b on a.codemp=b.codemp and b.CodTDc = a.CodTDc) "
        cCadReporte = cCadReporte & "LEFT JOIN TGAux c on a.codemp=c.codemp and c.CodAux = a.CodAux) "
        cCadReporte = cCadReporte & "LEFT JOIN CODro d ON a.codemp=d.codemp and a.pdoano=d.pdoano and a.CodDro=d.CodDro) "
        'cCadReporte = cCadReporte & "WHERE a.codemp='001' and a.pdoano='2012' and a.MesPvs >= '01' AND  a.MesPvs <= '04' "
        cCadReporte = cCadReporte & "WHERE "
        cCadReporte = cCadReporte & "    a.codemp='" & gsCodEmp & "' AND "
        
        'ini 2015-03-24
'        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
'        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & Left(cmbEjercicio.Text, 2) & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        If TpoRpt = 1 Then
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' "
'''        Else
        ElseIf TpoRpt = 2 Then
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
             cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        Else
'            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            '2015-05-25 teo dijo sacar YEAR(a.FeEDoc)  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
             cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' "
        End If
        'fin 2015-03-24
        'cCadReporte = cCadReporte & "    a.pdoano='" & sPdoAnoFin & "' AND "
        '2015-03-23  cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & gsMesAct & "' AND "
        '2015-03-20 cambio de periodo cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & "201502" & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        '2015-03-23 cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        
'ini 2015-01-09 adiciona ruc
      If Trim(txtDato(0).Text) <> "" Then
          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato(0).Text) & "' "
        End If
'fin 2015-01-09 adiciona ruc
        
        cCadReporte = cCadReporte & "ORDER BY a.pdoano, a.mespvs ,a.CodDro, a.NroCpb ASC "

    
    pocnnTmp.Execute cCadReporte
    
'ini 2015-07-15 adicion pgo segun diario
    sTabla = "tmp_xls_pdte"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
'saldos de documento segun reporte historico
    cCadReporte = cCadReporte & "SELECT "
    cCadReporte = cCadReporte & "    a.pdoano AS cAno, a.MesPvs, a.CodCta, a.CodAux,a.CodTDc,"
    cCadReporte = cCadReporte & "    a.SerDoc, a.NroDoc, a.CodDro, a.NroCpb, Null AS codcco,"
    'cCadReporte = cCadReporte & "    Null AS detcco, CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocum, a.FehOpe, a.FeEDoc, a.FeVDoc,"
    cCadReporte = cCadReporte & "    Null AS detcco,"
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(c.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocum, "
    cCadReporte = cCadReporte & "a.FehOpe, a.FeEDoc, a.FeVDoc, a.RefDoc, " & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, b.RazAux, "
    
    'cCadReporte = cCadReporte & "    a.RefDoc, a.GloIte AS GloIte, b.RazAux, (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
    cCadReporte = cCadReporte & "(CASE a.TpoMon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS cSigno, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END) AS cDebeMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END) AS cHaberMN, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END) AS cDebeME, "
    cCadReporte = cCadReporte & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END) AS cHaberME "
    
'    cCadReporte = cCadReporte & "    (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END) AS cDebeMN, (CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END) AS cHaberMN,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END) AS cDebeME, (CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END) AS cHaberME"
    cCadReporte = cCadReporte & "    ,a.TpoPvs"
    cCadReporte = cCadReporte & "    ,CONCAT(year(a.FehOpe),'-',LPAD(month(a.FehOpe),2,'0'),'-',LPAD(day(a.FehOpe),2,'0'),'-',a.CodDro,'-',a.NroCpb,'-',a.Nroite) AS x_clave "
    cCadReporte = cCadReporte & "FROM ((((COCpbDet a "
    cCadReporte = cCadReporte & "    LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
    cCadReporte = cCadReporte & "    LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
    cCadReporte = cCadReporte & "    LEFT JOIN Cocta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.Codcta=d.Codcta) "
    cCadReporte = cCadReporte & "    LEFT JOIN CoCCo e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.codcco=e.codcco) "
'    cCadReporte = cCadReporte & "WHERE a.codemp='010' "
'    cCadReporte = cCadReporte & "    AND a.pdoano='2014' "
    cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
    cCadReporte = cCadReporte & "    AND a.pdoano='" & gsAnoAct & "' "
    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta, 2)>='01' AND LEFT(a.codcta, 2)<='FF' "
    cCadReporte = cCadReporte & "    AND (a.ImpMN<> 0.00 OR a.ImpME<> 0.00) "
    'cCadReporte = cCadReporte & "    AND a.Mespvs <='03' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND a.Mespvs <='" & gsMesAct & "' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
    cCadReporte = cCadReporte & "    AND IFNULL(a.NroDoc, '') <>'' AND d.inddoc='1' "
    'cCadReporte = cCadReporte & "    # AND a.CodAux='10097267265'"
    cCadReporte = cCadReporte & "    AND a.TpoPvs='" & TPOPVS_CAN & "' " 'TPOPVS_CAN
    'cCadReporte = cCadReporte & "    AND a.TpoPvs='C' " 'TPOPVS_CAN
'2016-03-14 erro cuenta 422,428 se mezclan
    cCadReporte = cCadReporte & "    AND LEFT(a.codcta,3) = '421' "
    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta,3) <> '422' " '2015-09-03 erro cuenta anticipo duplicado
'2016-03-14 erro cuenta 422,428 se mezclan
    
'ini 2015-01-09 adiciona ruc
      If Trim(txtDato(0).Text) <> "" Then
          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato(0).Text) & "' "
        End If
'fin 2015-01-09 adiciona ruc
    
    cCadReporte = cCadReporte & "ORDER BY a.codcta, a.codaux, a.codtdc, a.serdoc, a.NroDoc, a.TpoPvs, a.MesPvs, a.FehOpe "


'    cCadReporte = cCadReporte & "SELECT "
'    cCadReporte = cCadReporte & "    a.pdoano AS cAno, a.MesPvs, a.CodCta, a.CodAux,a.CodTDc,"
'    cCadReporte = cCadReporte & "    a.SerDoc, a.NroDoc, a.CodDro, a.NroCpb, Null AS codcco,"
'    cCadReporte = cCadReporte & "    Null AS detcco, CONCAT(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocum, a.FehOpe, a.FeEDoc, a.FeVDoc,"
'    cCadReporte = cCadReporte & "    a.RefDoc, a.GloIte AS GloIte, b.RazAux, (CASE a.TpoMon WHEN 'N' THEN 'S/.' ELSE 'US$' END) AS cSigno,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpMN ELSE 0 END) AS cDebeMN, (CASE a.TpoCtb WHEN 'H' THEN a.ImpMN ELSE 0 END) AS cHaberMN,"
'    cCadReporte = cCadReporte & "    (CASE a.TpoCtb WHEN 'D' THEN a.ImpME ELSE 0 END) AS cDebeME, (CASE a.TpoCtb WHEN 'H' THEN a.ImpME ELSE 0 END) AS cHaberME"
'    cCadReporte = cCadReporte & "    ,a.TpoPvs"
'    cCadReporte = cCadReporte & "    ,CONCAT(year(a.FehOpe),'-',LPAD(month(a.FehOpe),2,'0'),'-',LPAD(day(a.FehOpe),2,'0'),'-',a.CodDro,'-',a.NroCpb,'-',a.Nroite) AS x_clave "
'    cCadReporte = cCadReporte & "FROM ((((COCpbDet a "
'    cCadReporte = cCadReporte & "    LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
'    cCadReporte = cCadReporte & "    LEFT JOIN TGTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
'    cCadReporte = cCadReporte & "    LEFT JOIN Cocta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.Codcta=d.Codcta) "
'    cCadReporte = cCadReporte & "    LEFT JOIN CoCCo e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.codcco=e.codcco) "
''    cCadReporte = cCadReporte & "WHERE a.codemp='010' "
''    cCadReporte = cCadReporte & "    AND a.pdoano='2014' "
'    cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
'    cCadReporte = cCadReporte & "    AND a.pdoano='" & gsAnoAct & "' "
'    'cCadReporte = cCadReporte & "    AND LEFT(a.codcta, 2)>='01' AND LEFT(a.codcta, 2)<='FF' "
'    cCadReporte = cCadReporte & "    AND (a.ImpMN<> 0.00 OR a.ImpME<> 0.00) "
'    'cCadReporte = cCadReporte & "    AND a.Mespvs <='03' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
'    cCadReporte = cCadReporte & "    AND a.Mespvs <='" & gsMesAct & "' AND IFNULL(a.CodAux, '') <>'' AND IFNULL(a.CodTDc, '') <>'' AND IFNULL(a.SerDoc, '') <>'' "
'    cCadReporte = cCadReporte & "    AND IFNULL(a.NroDoc, '') <>'' AND d.inddoc='1' "
'    'cCadReporte = cCadReporte & "    # AND a.CodAux='10097267265'"
'    cCadReporte = cCadReporte & "    AND a.TpoPvs='" & TPOPVS_CAN & "' " 'TPOPVS_CAN
'    'cCadReporte = cCadReporte & "    AND a.TpoPvs='C' " 'TPOPVS_CAN
'    cCadReporte = cCadReporte & "ORDER BY a.codcta, a.codaux, a.codtdc, a.serdoc, a.NroDoc, a.TpoPvs, a.MesPvs, a.FehOpe "
    
    
    pocnnTmp.Execute cCadReporte

'*********************
    sTabla = "tmp_xls_pdte2"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
cCadReporte = cCadReporte & "SELECT "
cCadReporte = cCadReporte & "    codcta,codaux,cdocum,min(x_clave) x_clave "
cCadReporte = cCadReporte & "From tmp_xls_pdte "
cCadReporte = cCadReporte & "GROUP BY codcta, codaux,cdocum # x_clave "
cCadReporte = cCadReporte & "ORDER BY codcta, codaux,cdocum #x_clave "

    pocnnTmp.Execute cCadReporte

'*********************
    sTabla = "tmp_xls_pdte3"
    pocnnTmp.Execute fDropTable2(sTabla, 1)
    cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")

cCadReporte = cCadReporte & "SELECT "
cCadReporte = cCadReporte & "* "
cCadReporte = cCadReporte & "From tmp_xls_pdte "
cCadReporte = cCadReporte & "Where x_clave "
cCadReporte = cCadReporte & "    IN (select x_clave from tmp_xls_pdte2) "
    pocnnTmp.Execute cCadReporte
    
'*********************
    
'fin 2015-07-15 adicion pgo segun diario
    
    
'ini exporta datos a excel

    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
    With porstTmp
       .ActiveConnection = pocnnTmp
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
       
'       .Source = "SELECT * FROM " & ps_Prefijo & sTabla

'ini 2015-07-15 adicion pgo segun diario
.Source = "SELECT "
.Source = .Source & "    CPERIODO,CNUMREGOPE,CFECCOM,CFECVENPAG,CTIPDOCCOM,"
.Source = .Source & "    CNUMSER,CEMIDUADSI,CNUMDCODFV,COSDCREFIS,CTIPDIDPRO,"
.Source = .Source & "    CNUMDIDPRO,CNOMRSOPRO,CBASIMPGRA,CIGVGRA,CBASIMPGNG,"
.Source = .Source & "    CIGVGRANGV,CBASIMPSCF,CIGVSCF,CIMPTOTNGV,CISC,"
.Source = .Source & "    COTRTRICGO,CIMPTOTCOM,CTIPCAM,CFECCOMMOD,CTIPCOMMOD,"
.Source = .Source & "    CNUMSERMOD,CNUMCOMMOD,CCOMNODOMI,CEMIDEPDET,CNUMDEPDET,"
.Source = .Source & "    CCOMPGRET,CESTOPE,CVALFACIMP,CINTDIAMAY,CINTKARDEX,"
.Source = .Source & "    CINTREG,tsadetrac,xCol1,xCol2,tpomon,"
.Source = .Source & "    CIMPTOTMEX,GloDoc,"
.Source = .Source & "    a.refdoc," '2015-12-17 adicion ref
'#   a.*,,b.FehOpe
.Source = .Source & "    b.FehOpe,"
.Source = .Source & "    IFNULL(b.cDebeMN,0)-IFNULL(b.cHaberMN,0) PgoMN,"
.Source = .Source & "    IFNULL(b.cDebeME,0)-IFNULL(b.cHaberME,0) PgoME "
.Source = .Source & "FROM xlsCprCab a "
.Source = .Source & "LEFT JOIN tmp_xls_pdte3 b "
.Source = .Source & "    ON a.CodAux=b.CodAux and a.CTIPDOCCOM=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc "

'fin 2015-07-15 adicion pgo segun diario
       
       .Open
    End With

    Dim xArchPeriodo As String
    xArchPeriodo = "plan 2011 txtpg.xlsx"

    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
 
    'Set oSheet = oWBook.Worksheets(1)
 

    '*Set oExcel = New Excel.Application
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

    Set oWBook = oExcel.Workbooks.Add
    '*Set oWBook = oExcel.Workbooks.Open(dlbDirectorio(0).path & xArchPeriodo, , True) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    '*Set oSheet = oWBook.Worksheets("Clientes")
     Set oSheet = oWBook.Worksheets(1)
    '*oExcel.Visible = True

    With oSheet
    
'        oSheet.Select
'        Columns("M:V").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"

        oSheet.Select
        
        .Cells(1, 1).Value = "Registro de Compras"
        
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        
        .Cells(nRowI, 1).Value = "Registro de Compras"
        nRowI = nRowI + 2
        Dim x1 As Integer
        .Cells(nRowI, 1).Value = "Periodo"
        .Cells(nRowI, 2).Value = "Nº Reg."
        .Cells(nRowI, 3).Value = "F.Cmpra"
        .Cells(nRowI, 4).Value = "F. Pago"
        .Cells(nRowI, 5).Value = "T.Doc"
        .Cells(nRowI, 6).Value = "Serie"
        .Cells(nRowI, 7).Value = "CemiDuadsi"
        .Cells(nRowI, 8).Value = "Nº Doc."
        .Cells(nRowI, 9).Value = "COSDCREFIS"
        .Cells(nRowI, 10).Value = "T.Prv"
        .Cells(nRowI, 11).Value = "RUC"
        .Cells(nRowI, 12).Value = "R.Social"
        .Cells(nRowI, 13).Value = "B. Gravada"
        .Cells(nRowI, 14).Value = "IGV Grab"
        .Cells(nRowI, 15).Value = "B. G/N Gr"
        .Cells(nRowI, 16).Value = "IGV G/N Gr"
        .Cells(nRowI, 17).Value = "B. Sin CF"
        .Cells(nRowI, 18).Value = "Igv S CF"
        .Cells(nRowI, 19).Value = "CIMPTOTNGV"
        .Cells(nRowI, 20).Value = "CISSC"
        .Cells(nRowI, 21).Value = "COTRTRICGO"
        .Cells(nRowI, 22).Value = "CIMPTOTCOM"
        .Cells(nRowI, 23).Value = "CTIPCAM"
        .Cells(nRowI, 24).Value = "CFECCOMMOD"
        .Cells(nRowI, 25).Value = "CTIPCOMMOD"
        .Cells(nRowI, 26).Value = "CNUMSERMOD"
        .Cells(nRowI, 27).Value = "CNUMCOMMOD"
        .Cells(nRowI, 28).Value = "CCOMNODOMI"
        .Cells(nRowI, 29).Value = "CEMIDEPDET"
        .Cells(nRowI, 30).Value = "CNUMDEPDET"
        .Cells(nRowI, 31).Value = "CCOMPGRET"
        .Cells(nRowI, 32).Value = "CESTOPE"
        .Cells(nRowI, 33).Value = "CVALFACIMP"
        .Cells(nRowI, 34).Value = "CINTDIAMAY"
        .Cells(nRowI, 35).Value = "CINTKARDEX"
        .Cells(nRowI, 36).Value = "CINTREG"
        .Cells(nRowI, 37).Value = "tsadetrac"
        .Cells(nRowI, 38).Value = "DetaDetrac"
        .Cells(nRowI, 39).Value = "PorcDetra"
        .Cells(nRowI, 40).Value = "TpoMon"
        .Cells(nRowI, 41).Value = "Total ME"
        .Cells(nRowI, 42).Value = "Glosa"
        .Cells(nRowI, 43).Value = "Refer."
        .Cells(nRowI, 43 + 1).Value = "F.Pago"
        .Cells(nRowI, 44 + 1).Value = "1er Pgo MN"
        .Cells(nRowI, 45 + 1).Value = "1er Pgo ME"
     
        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        .Columns.AutoFit ' ajusta el ancho de las columnas
        'Sheets(oSheet).Select
        
        'hay sale error definido por la aplicacion o el objeto 1004, cuando aplico estos comandos Select y NumberFormat
'        oSheet.Select
'solo sale error en esta        Columns("M:M").Select
'        Selection.NumberFormat = "#,##0.00"

'        Columns("N:N").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("O:O").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("P:P").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("Q:Q").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("R:R").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("S:S").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("T:T").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("U:U").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("V:V").Select
'        Selection.NumberFormat = "#,##0.00"
        
        'crear tabla temporal
        'Dim xpocnnMain As ADODB.Connection
        'Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next
'        oSheet.Select
'        Cells(1, 1).Select

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 2).Value)
'        Next


 'ini 2015-07-02 adic tabla detrac
'*********************************
        Dim uorstcodetrac As ADODB.Recordset
        Set uorstcodetrac = New ADODB.Recordset
        Set uorstcodetrac = fRstDetrac(pocnnMain, uorstcodetrac)
'        With uorstCoDetrac
'           .ActiveConnection = pocnnMain
'           .Source = "SELECT coddetrac, " & Choose(gsIdioma, "detdetrac", "detdetracx") & " AS DetDetrac,tsadetrac ,  "
'           .Source = .Source & "codemp "
'           .Source = .Source & "FROM codetrac  "
'           .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
'           .Source = .Source & "AND estdetrac ='" & ESTDETRAC_ACT & "' "
'           '.Source = .Source & "AND pdoano='" & gsAnoAct & "' "
'           '.Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
'           .CursorType = adOpenDynamic
'           .LockType = adLockOptimistic
'           .Open
'        End With
        
        
       xrow1 = nRowI
        Dim nContador As Integer
        Dim s_Contenido As String
        Dim n_Detraccion As Double
        Dim s_detalle As String
        Do While Len(Trim(.Cells(xrow1, 2).Value)) <> 0
            s_Contenido = Left(.Cells(xrow1, 37).Value, 5)
            With uorstcodetrac
                If .RecordCount > 0 Then .MoveFirst
                    .Find "coddetrac='" & s_Contenido & "'"
                    If Not .EOF Then
                        oSheet.Cells(xrow1, 38).Value = !coddetrac
                        '2015-07-08 cambio de decima a % oSheet.Cells(xrow1, 39).Value = !pctdetrac * 100
                        oSheet.Cells(xrow1, 39).Value = !pctdetrac
                   End If
            End With
            xrow1 = xrow1 + 1
        Loop
        
        uorstcodetrac.Close
        Set uorstcodetrac = Nothing


'*********************************
''       xrow1 = nRowI
''        Dim nContador As Integer
''        Dim s_Contenido As String
''        Dim n_Detraccion As Double
''        Dim s_detalle As String
''        Do While Len(Trim(.Cells(xrow1, 2).Value)) <> 0
''            'MsgBox (.Cells(xrow1, 37).Value)
''            's_Contenido = Left(.Cells(xrow1, 37).Value, 3)
''            s_Contenido = Left(.Cells(xrow1, 37).Value, 5)
''            'ini 2014-04-05 reclasificacion de cod detraccion
''            For nContador = 1 To UBound(aDtraccDet, 1)
''            'If Left(aDtraccDet(nContador), 3) = s_Contenido Then
''            If Left(aDtraccDet(nContador), 5) = s_Contenido Then
''                n_Detraccion = aDtraccPor(nContador)
''                s_detalle = aDtraccDet(nContador)
''                s_detalle = Mid(s_detalle, 7)
''                .Cells(xrow1, 38).Value = s_detalle
''                .Cells(xrow1, 39).Value = n_Detraccion * 100
''               Exit For
''            End If
''            Next nContador
''            xrow1 = xrow1 + 1
''        Loop
'fin 2015-07-02 adic tabla detrac
        
    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing


'fin exporta datos a excel
  pocnnMain.Execute fDropTable("xlsCprCab", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte2", 1)
  pocnnMain.Execute fDropTable("tmp_xls_pdte3", 1)
'  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS xlsCprCab", s_Sentencia)
'  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp_xls_pdte", s_Sentencia)
'  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp_xls_pdte2", s_Sentencia)
'  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmp_xls_pdte3", s_Sentencia)

'fDropTable
   porstTmp.Close
   pocnnTmp.Close
   Set porstTmp = Nothing
   Set pocnnTmp = Nothing

  Exit Sub
Err:
    MsgBox (TEXT_6001)
  If pocnnTmp.State = adStateOpen Then
    porstTmp.Close
    pocnnTmp.Close
    Set porstTmp = Nothing
    Set pocnnTmp = Nothing
  End If
End Sub

Private Sub cmdExporta_Click_2015_03_20(Index As Integer)
frmRRegCprExpo.Show vbModal
End Sub




Private Sub Form_Load()
   On Error GoTo Err
   
'ini sql8 2015-03-23
toolbar.Buttons(1).ButtonMenus(1).Text = "Del Mes"
toolbar.Buttons(1).ButtonMenus(2).Text = "Al Mes"
'fin sql8 2015-03-23
toolbar.Buttons(1).ButtonMenus(3).Text = "Historico" '2015-09-03 opc historico
 
   Dim dnContador As Integer
 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstTGAux = New ADODB.Recordset
   Set porstCodro = New ADODB.Recordset
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
   With porstTGAux
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodAux, RazAux "
      .Source = .Source & "FROM TGAux "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND IndPrv=" & INDAUX_PRV_ACT & " "
      .Source = .Source & "ORDER BY CodAux"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With porstCodro
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
    .Source = .Source & "FROM CODro "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "ORDER BY CodDro"
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
   End With
 ']
 '[Parámetros.                         'Cambiar.
   With cboTpoMon
      .AddItem TPOMON_NAC_TXT_1, 0
      .AddItem TPOMON_EXT_TXT_1, 1
   End With
  txtDato.Item(0).DataField = "CodAux"
  txtDato.Item(0).MaxLength = porstTGAux.Fields(txtDato.Item(0).DataField).DefinedSize
  txtDato.Item(1).DataField = "CodDro"
  txtDato.Item(1).MaxLength = porstCodro.Fields(txtDato.Item(1).DataField).DefinedSize
 ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Currency :")
  Next nElemento
  fraAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  chkDiario.Caption = Choose(gsIdioma, "Totaliza Diario", "Journal Totalizes")
  fraRangos.Caption = Choose(gsIdioma, "Diario", "Journal")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
'   With porstTgAux
'      .MoveLast
'      'txtDato(1).Text = !CodAux
'      .MoveFirst
'      txtDato(0).Text = !CodAux
'   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
  'Características de impresión.
   chkImpFecha.Value = vbChecked
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
   porstTGAux.Close
   pocnnMain.Close
   Set porstTGAux = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub
Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim dnContador As Byte
  
  ppHabilitacion False
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.FeEDoc, a.FehOpe, a.CodDro, a.NroCpb, b.AbvTDc, "
    .Source = .Source & "a.SerDoc, a.NroDoc, a.RefDoc, c.RUCAux, c.RazAux, a.NroCDt, a.FehCDt, "
    '.Source = .Source & "a.SerDoc, a.NroDoc, concat(a.mespvs,'-',a.RefDoc), c.RUCAux, c.RazAux, a.NroCDt, a.FehCDt, "
    '[ARREGLAR. Poder configurar el signo en Tipo de Documento. ImpIGV_OGr_MN
    If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
      .Source = .Source & "(a.ImpOGr_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOgr, "
      .Source = .Source & "(a.ImpOGN_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOGN, "
      .Source = .Source & "(a.ImpONG_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpONG, "
      .Source = .Source & "(a.ImpExo_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpExo, "
      .Source = .Source & "(a.ImpIGV_OGr_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGr, "
      .Source = .Source & "(a.ImpIGV_OGN_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGN, "
      .Source = .Source & "(a.ImpIGV_ONG_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVONG, "
      .Source = .Source & "(a.ImpISC_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpISC, "
      .Source = .Source & "(a.ImpOIm_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOIm, "
      .Source = .Source & "(a.ImpTot_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot, "
      .Source = .Source & "(a.ImpTot_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot_OM, "
    Else
      .Source = .Source & "(a.ImpOGr_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOgr, "
      .Source = .Source & "(a.ImpOGN_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOGN, "
      .Source = .Source & "(a.ImpONG_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpONG, "
      .Source = .Source & "(a.ImpExo_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpExo, "
      .Source = .Source & "(a.ImpIGV_OGr_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGr, "
      .Source = .Source & "(a.ImpIGV_OGN_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVOGN, "
      .Source = .Source & "(a.ImpIGV_ONG_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpIGVONG, "
      .Source = .Source & "(a.ImpISC_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpISC, "
      .Source = .Source & "(a.ImpOIm_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpOIm, "
      .Source = .Source & "(a.ImpTot_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot, "
      .Source = .Source & "(a.ImpTot_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS cImpTot_OM, "
    End If
    ']ARREGLAR.
    .Source = .Source & "b.CodTDc, d.DetDro, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "a.CodDro", IIf(Trim(txtDato(1).Text) <> "", "a.CodDro", "'drxx'")) & " AS grupo, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "'1'", IIf(Trim(txtDato(1).Text) <> "", "'2'", "'0'")) & " AS resumen "
    .Source = .Source & "FROM (((COCprDoc a "
    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    .Source = .Source & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
    .Source = .Source & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.MesPvs='" & gsMesAct & "' "
    '.Source = .Source & "AND a.MesPvs in ('01','02','03','04','05','06','07','08','09','10','11','12') "
    If Trim(txtDato(0).Text) <> "" Then
      .Source = .Source & "AND a.CodAux='" & Trim(txtDato(0).Text) & "' "
    End If
    If Trim(txtDato(1).Text) <> "" Then
      .Source = .Source & "AND Left(a.CodDro, " & Len(Trim(txtDato(1).Text)) & ")='" & Trim(txtDato(1).Text) & "' "
    End If
    .Source = .Source & "ORDER BY grupo, a.CodDro, a.NroCpb ASC"
    '.Source = .Source & "ORDER BY a.mespvs,grupo, a.CodDro, a.NroCpb ASC"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '       '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRRegCpr.rpt"
      .Formulas(7) = "pSigMon='" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, gsTpoMon_Sgn_ME, gsTpoMon_Sgn_MN) & "'"
      
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRRegCpr.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      
      '[Parámetros adicionales.
      If porstMRp.RecordCount > 0 Then
        porstMRp.MoveLast
        .Parameters("pPagePrinter") = porstMRp!coddro & porstMRp!NroCpb
        porstMRp.MoveFirst
      End If
      ']
      
      If Index = 0 Then
        .PreviewReport
      Else
        '[ARREGLAR: Revisar el uso de los tres primeros parámetros de Print.
        .Print 1, 0, 0, unCopias
        ']ARREGLAR.
      End If
      .UnLoadReport
    End With
    Set MRViewer = Nothing
  End If
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

'ini 2015-03-24
Private Sub toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  'no pinto datos Seleccion.Text = ButtonMenu.Text
  Select Case ButtonMenu.Key
'   Case "A1": pExporta_2016_07_14 1
'   Case "A2": pExporta_2016_07_14 2
   Case "A1": pExporta 1
   Case "A2": pExporta 2
'   Case "A" & Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
'    pnOpcion = Right(ButtonMenu.Key, Len(ButtonMenu.Key) - 1)
   Case "A3": pExporta 3 '2015-09-03 opc historico
  End Select

End Sub
'fin 2015-03-24

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
'End Sub

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
   Select Case Index    'Busca el dato en su tabla principal.
   Case 0                              'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
   Case 1
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
    If Len(txtDato(Index)) <> 4 Then txtDato(Index).SetFocus: Exit Sub
  End Select
End Sub
Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.Aux_Det "IndPrv=" & INDAUX_PRV_ACT & " ", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    Case 1
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
      With porstTGAux
         .MoveFirst
         .Find "CodAux='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & !razAux
         End If
      End With
    Case 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstCodro
         .MoveFirst
         .Find "Coddro='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
         End If
      End With
   End Select
End Function
Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
   optTipoImpresion(0).Enabled = tbHabilitar
   optTipoImpresion(1).Enabled = tbHabilitar
   cmdImprimir(0).Enabled = tbHabilitar
   cmdImprimir(1).Enabled = tbHabilitar
   cmdConfig.Enabled = tbHabilitar
   cmdSalir.Enabled = tbHabilitar

  'Controles del formulario.
'   cboTpoMon.Enabled = tbHabilitar
'   dtpFecha.Enabled = tbHabilitar
'   optTipo(0).Enabled = tbHabilitar
'   optTipo(1).Enabled = tbHabilitar
'   With txtDato
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With cmdDatoAyud
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
'   With lblDatoDeta
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

