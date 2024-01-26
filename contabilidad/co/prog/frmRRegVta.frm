VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRRegVta 
   Caption         =   "[título]"
   ClientHeight    =   3195
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDiario 
      Caption         =   "Totaliza Diario"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3150
      TabIndex        =   7
      Top             =   855
      Width           =   1335
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Diario"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   1020
      Width           =   4530
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   4080
         Picture         =   "frmRRegVta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   255
         Width           =   3240
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   1515
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5040
      TabIndex        =   14
      Top             =   1845
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   15
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   975
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraAuxiliar 
      Caption         =   "Cliente"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   0
      TabIndex        =   4
      Top             =   135
      Width           =   7290
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
         TabIndex        =   5
         Top             =   255
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6885
         Picture         =   "frmRRegVta.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   255
         Width           =   255
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
         TabIndex        =   6
         Top             =   255
         Width           =   5520
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmRRegVta.frx":0354
      Left            =   6150
      List            =   "frmRRegVta.frx":0356
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1050
      Width           =   1125
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7350
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2595
      Width           =   7350
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
         Height          =   570
         Left            =   4800
         Picture         =   "frmRRegVta.frx":0358
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
         Height          =   570
         Index           =   0
         Left            =   0
         Picture         =   "frmRRegVta.frx":04A2
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
         Height          =   570
         Index           =   1
         Left            =   1245
         Picture         =   "frmRRegVta.frx":09D4
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
                  Picture         =   "frmRRegVta.frx":0AD6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegVta.frx":0C30
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegVta.frx":0D8A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegVta.frx":114C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRRegVta.frx":1816
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
      TabIndex        =   11
      Top             =   1095
      Width           =   765
   End
End
Attribute VB_Name = "frmRRegVta"
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
Private porstCodro As ADODB.Recordset
Private porstTGAux As ADODB.Recordset
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

    Dim pocnnTmp As ADODB.Connection '2014-04-14 Query timeout expired
    Set pocnnTmp = New ADODB.Connection '2014-04-14 Query timeout expired
    With pocnnTmp
       .CursorLocation = adUseClient
       .ConnectionString = CONNSTRG & gsNomBDS
       .Open
    End With
    
    Dim cCadReporte  As String
    Dim sTabla As String
    sTabla = "xlsVtaCab"
   'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & sTabla & " ", cCadReporte)
    pocnnTmp.Execute fDropTable2(sTabla, 1)

        cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS " & sTabla & " ", "")
    cCadReporte = cCadReporte & "SELECT"
    cCadReporte = cCadReporte & "    concat(a.pdoano,a.mespvs,'00') AS VPERIODO,"
    cCadReporte = cCadReporte & "    concat(a.CodDro,a.NroCpb) as VNUMREGOPE,"
    cCadReporte = cCadReporte & "    date_format(a.feedoc,'%d/%m/%Y')as VFECCOM,"
    cCadReporte = cCadReporte & "    date_format(a.FevDOC,'%d/%m/%Y')as VFECVENPAG,"
    cCadReporte = cCadReporte & "    b.CodTDc as VTIPDOCCOM, a.SerDoc AS VNUMSER, a. NroDoc AS VNUMDOCCOI,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.NroDoc_Fin,''),a.NroDoc_Fin,'0') AS VNUMDOCCOF,"
    cCadReporte = cCadReporte & "    MID(c.tpodci,2,1) AS VTIPDIDCLI,"
    cCadReporte = cCadReporte & "    c.Codaux AS VNUMDIDCLI,"
    cCadReporte = cCadReporte & "    replace(replace(replace(replace(replace(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
    cCadReporte = cCadReporte & "    ifnull(MID(c.RazAux,1,60) ,''), '?', ' '), '*', ' '),'%',' '),'&',' '),'!',' '),'" & Chr(34) & "',' '),',',' '),'|',' '),'+',' '),')',' '),'$',' '),'~',' '),'ø',' '),'¥',' '),'¤', ' '),'°',' '),'º',' ')"
    cCadReporte = cCadReporte & "    as VAPENOMRSO,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpExp_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VVALFACEXP,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VBASIMPGRA,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpExo_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTEXO,"
    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTINA,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpISC_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VISC,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIGVIPM,"
    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VBASIMIVAP,"
    cCadReporte = cCadReporte & "    replace(format((0.00        * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIVAP,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpOIm_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VOTRTRICGO,"
    cCadReporte = cCadReporte & "    replace(format((a.ImpTot_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTCOM,"
    cCadReporte = cCadReporte & "    format(a.imptcb,3) * 1 AS VTIPCAM,"
    cCadReporte = cCadReporte & "    IF(ifnull(codtdc_ref,''),date_format(feedoc_ref,'%d/%m/%Y'),'01/01/0001') as VFECCOMMOD,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.codtdc_ref,''),a.codtdc_ref,'00') as VTIPCCOMOD,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.serdoc_ref,''),a.serdoc_ref,'-')  as VNUMSERMOD,"
    cCadReporte = cCadReporte & "    IF(ifnull(a.nrodoc_ref,''),a.nrodoc_ref,'-')  as VNUMCOMMOD,"
    cCadReporte = cCadReporte & "    IF(a.ImpTot_MN <>0.00,'1','2') as VESTOPE,"
    cCadReporte = cCadReporte & "    '' AS VINTDIAMAY,"
    cCadReporte = cCadReporte & "    '' AS VINTKARDEX,"
    cCadReporte = cCadReporte & "    '' AS VINTREG "
    cCadReporte = cCadReporte & "    ,a.TpoMon " '2015-05-14
    
    cCadReporte = cCadReporte & "    ,replace(format((a.ImpTot_ME * IF(b.SgnTDc = 0, -1,1)),2),',','') * 1 AS VIMPTOTMEX " '2015-07-13 adici vta/cpr me
    
    cCadReporte = cCadReporte & "    ,GloDoc " '2015-06-04 adicion glodoc
    cCadReporte = cCadReporte & "    ,ifnull(refdoc,'') refdoc " '2015-12-17 adicion ref
    cCadReporte = cCadReporte & "    ,date_format(a.feedoc,'%d/%m/%Y')as fehcdt,nrocdt,tsadetrac,pctdetrac " '2015-07-03 adicion campo detracc vta
           
    cCadReporte = cCadReporte & "FROM (((COVtaDoc a "
    cCadReporte = cCadReporte & "LEFT JOIN TGTDc b ON  a.codemp=b.codemp and a.CodTDc=b.CodTDc) "
    cCadReporte = cCadReporte & "LEFT JOIN TGAux c ON  a.codemp=c.codemp  and a.CodAux=c.CodAux) "
    cCadReporte = cCadReporte & "LEFT JOIN CODro d ON  a.codemp=d.codemp  and a.pdoano=d.pdoano and a.CodDro=d.CodDro) "
    'cCadReporte = cCadReporte & "WHERE a.codemp='001' and a.pdoano='2012' and a.Mespvs >='01'  and a.Mespvs <='05' AND IFNULL(a.CodAux, '')<>'' AND IFNULL(a.CodDro, '')<>'' "
    cCadReporte = cCadReporte & "WHERE "
    cCadReporte = cCadReporte & "   a.codemp='" & gsCodEmp & "' and "
        'ini 2015-03-24
'    cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) >='" & gsAnoAct & "01" & "'  and "
'    cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) <='" & gsAnoAct & Left(cmbEjercicio.Text, 2) & "' AND "
        If TpoRpt = 1 Then
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) = '" & gsAnoAct & gsMesAct & "' AND "
        'Else '2015-09-03 opc historico
        ElseIf TpoRpt = 2 Then
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND  "
'ini 2015-09-03 opc historico
        Else
            'cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & "01" & "' AND "
            cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND  "
'fin 2015-09-03 opc historico
        End If

        'fin 2015-03-24
    'cCadReporte = cCadReporte & "   a.pdoano='" & sPdoAnoFin & "' and "
    '2015-03-23 cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) >='" & gsAnoAct & gsMesAct & "'  and "
    '*cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) <='" & "201412" & "' AND "
    '2015-03-23cCadReporte = cCadReporte & "   concat(a.pdoano,a.Mespvs) <='" & gsAnoAct & gsMesAct & "' AND "
    cCadReporte = cCadReporte & "   IFNULL(a.CodAux, '')<>'' AND  "
    cCadReporte = cCadReporte & "   IFNULL(a.CodDro, '')<>'' "
'ini 2015-01-09 adiciona ruc
      If Trim(txtDato(0).Text) <> "" Then
          cCadReporte = cCadReporte & "AND a.codaux='" & Trim(txtDato(0).Text) & "' "
      End If
'fin 2015-01-09 adiciona ruc
    
    
    cCadReporte = cCadReporte & "ORDER BY a.mespvs ,a.CodTDc, a.SerDoc, a.NroDoc  ASC "

    
    pocnnTmp.Execute cCadReporte
    
'ini exporta datos a excel

    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
    With porstTmp
       .ActiveConnection = pocnnTmp
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
       .Source = "SELECT * FROM " & ps_Prefijo & sTabla
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
        oSheet.Select
        
        '.Cells(1, 1).Value = "Registro de Ventas"
        
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        
        .Cells(nRowI, 1).Value = "Registro de Ventas"
        nRowI = nRowI + 2
        Dim x1 As Integer
        .Cells(nRowI, 1).Value = "Periodo"
        .Cells(nRowI, 2).Value = "Nº Reg."
        .Cells(nRowI, 3).Value = "F.Vta"
        .Cells(nRowI, 4).Value = "F. Pago"
        .Cells(nRowI, 5).Value = "T.Doc"
        .Cells(nRowI, 6).Value = "Serie"
        .Cells(nRowI, 7).Value = "VNUMDOCCCOI"
        .Cells(nRowI, 8).Value = "Nº Doc."
        .Cells(nRowI, 9).Value = "Tpo.Cli"
        .Cells(nRowI, 10).Value = "RUC"
        .Cells(nRowI, 11).Value = "R.Social"
        .Cells(nRowI, 12).Value = "VVALFACEXP"
        .Cells(nRowI, 13).Value = "VBASIMPGRA"
        .Cells(nRowI, 14).Value = "VIMPTOTEXO"
        .Cells(nRowI, 15).Value = "VIMPTOTINA"
        .Cells(nRowI, 16).Value = "VISC"
        .Cells(nRowI, 17).Value = "VIGVIPM"
        .Cells(nRowI, 18).Value = "VBASIMIVAP"
        .Cells(nRowI, 19).Value = "VIVAP"
        .Cells(nRowI, 20).Value = "VOTRTRICGO"
        .Cells(nRowI, 21).Value = "CIMPTOTCOM"
        .Cells(nRowI, 22).Value = "VTIPCAM"
        .Cells(nRowI, 23).Value = "VFECCOMMOD"
        .Cells(nRowI, 24).Value = "VTIPCCOMOD"
        .Cells(nRowI, 25).Value = "VNUMSERMOD"
        .Cells(nRowI, 26).Value = "VNUMCOMMOD"
        .Cells(nRowI, 27).Value = "VESTOPE"
        .Cells(nRowI, 28).Value = "VINTDIAMAY"
        .Cells(nRowI, 29).Value = "VINTKARDEX"
        .Cells(nRowI, 30).Value = "VINTREG"
        .Cells(nRowI, 31).Value = "TpoMon"
        .Cells(nRowI, 32).Value = "Total ME"
        .Cells(nRowI, 33).Value = "Glosa" '2015-06-04 adicion glodoc
        .Cells(nRowI, 34).Value = "Refer." '2015-12-17 adicion ref
'ini 2015-07-03 adicion campo detracc vta
        .Cells(nRowI, 34 + 1).Value = "F.Detrac"
        .Cells(nRowI, 35 + 1).Value = "Doc.Detrac"
        .Cells(nRowI, 36 + 1).Value = "Tsa.Detrac"
        .Cells(nRowI, 37 + 1).Value = "% Detrac" '
'fin 2015-07-03 adicion campo detracc vta
        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        
        'hay sale error definido por la aplicacion o el objeto 1004, cuando aplico estos comandos Select y NumberFormat
'        oSheet.Select
'        Columns("L:L").Select
'        Selection.NumberFormat = "#,##0.00"
'        Columns("M:M").Select
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
'        Selection.NumberFormat = "#,##0.000"
        
        'crear tabla temporal
        'Dim xpocnnMain As ADODB.Connection
        'Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next
'        oSheet.Select
'        Cells(1, 1).Select

    End With
    'oExcel.Visible = True
    oExcel.Quit
    Set oExcel = Nothing


'fin exporta datos a excel

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
      .Source = .Source & "AND IndCli=" & INDAUX_CLI_ACT & " "
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
    '     .CursorLocation = adUseClient   'Es el Default.
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
  fraAuxiliar.Caption = Choose(gsIdioma, "Cliente", "Customer")
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
   porstCodro.Close
   pocnnMain.Close
   Set porstTGAux = Nothing
   Set porstCodro = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0       ', 1
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   Case 1       ', 1
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim dnContador As Byte
  
  ppHabilitacion False
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT a.FeEDoc, a.FehOpe, a.CodDro, a.NroCpb, "
    .Source = .Source & "b.AbvTDc, a.SerDoc, a. NroDoc, a.SerDoc_Fin, "
    .Source = .Source & "a.NroDoc_Fin , a.RefDoc, c.RucAux, c.RazAux, "
    '.Source = .Source & "a.NroDoc_Fin , concat(a.mespvs,'-',a.RefDoc), c.RucAux, c.RazAux, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.CodDro, '-', a.NroCpb)", "(a.CodDro+'-'+a.NroCpb)") & " AS cx1, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc, '-', a.NroDoc)", "(a.SerDoc+'-'+a.NroDoc)") & " AS cx2, "
    .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.SerDoc_Fin, '-', a.NroDoc_Fin)", "(a.SerDoc_Fin+'-'+a.NroDoc_Fin)") & " AS cx3, "
    '[ARREGLAR. Poder configurar el signo en Tipo de Documento.
    If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
      .Source = .Source & "(a.ImpOGr_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmOgr, "
      .Source = .Source & "(a.ImpExp_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmExp, "
      .Source = .Source & "(a.ImpExo_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmExo, "
      .Source = .Source & "(a.ImpIGV_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmIgv, "
      .Source = .Source & "(a.ImpISC_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmISC, "
      .Source = .Source & "(a.ImpOIm_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmOlm, "
      .Source = .Source & "(a.ImpTot_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmTot, "
      .Source = .Source & "(a.ImpTot_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmpTot, "
    Else
      .Source = .Source & "(ImpOGr_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmOgr, "
      .Source = .Source & "(a.ImpExp_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmExp, "
      .Source = .Source & "(a.ImpExo_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmExo, "
      .Source = .Source & "(a.ImpIGV_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmIgv, "
      .Source = .Source & "(a.ImpISC_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmISC, "
      .Source = .Source & "(a.ImpOIm_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmOlm, "
      .Source = .Source & "(a.ImpTot_ME * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmTot, "
      .Source = .Source & "(a.ImpTot_MN * (CASE b.SgnTDc WHEN " & SGNTDC_NEG & " THEN -1 ELSE 1 END)) AS clmpTot, "
    End If
']ARREGLAR.
    .Source = .Source & "b.CodTDc, d.DetDro, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "a.CodDro", IIf(Trim(txtDato(1).Text) <> "", "a.CodDro", "'drxx'")) & " AS grupo, "
    .Source = .Source & IIf(chkDiario.Value = vbChecked, "'1'", IIf(Trim(txtDato(1).Text) <> "", "'2'", "'0'")) & " AS resumen "
    .Source = .Source & "FROM (((COVtaDoc a "
    .Source = .Source & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    .Source = .Source & "LEFT JOIN TGAux c ON a.codemp=c.codemp AND a.CodAux=c.CodAux) "
    .Source = .Source & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.Mespvs ='" & gsMesAct & "' "
    '.Source = .Source & "AND a.Mespvs in ('01','02','03','04','05','06','07','08','09','10','11','12') "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
    If Trim(txtDato(0).Text) <> "" Then
      .Source = .Source & "AND a.CodAux = '" & Trim(txtDato(0).Text) & "' "
    End If
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodDro, '')<>'' "
    If Trim(txtDato(1).Text) <> "" Then
      .Source = .Source & "AND Left(a.CodDro, " & Len(Trim(txtDato(1).Text)) & ")='" & Trim(txtDato(1).Text) & "' "
    End If
    .Source = .Source & "ORDER BY grupo, a.CodTDc, a.SerDoc, a.NroDoc  ASC"
    '.Source = .Source & "ORDER BY a.mespvs,grupo, a.CodTDc, a.SerDoc, a.NroDoc  ASC"
    .Open
  End With

  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '       '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptRRegVta.rpt"
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
      .LoadReport gsRutRpt & "rptRRegVta.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value)
      '[Parámetros adicionales.
      .Parameters("pPagePrinter") = ""
      If porstMRp.RecordCount > 0 Then
        porstMRp.MoveLast
        .Parameters("pPagePrinter") = porstMRp!codtdc & porstMRp!serdoc & porstMRp!nrodoc
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
   '**************************beto
   If txtDato(0) = "" Then
       lblDatoDeta(0).Caption = ""
   End If
    
   If txtDato(1) = "" Then
       lblDatoDeta(1).Caption = ""
   End If
   
   
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   
  Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1    ', 1                           'Cambiar (añadir índices).
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
   Case 0                           'Cambiar (añadir índices).
      modAyuBus.Aux_Det "IndCli=" & INDAUX_CLI_ACT & " ", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAuxiliar.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAuxiliar.Left + txtDato(tnIndex).Left
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

