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
   ScaleHeight     =   6390
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdimportes 
      Caption         =   "Importes"
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
      Left            =   3585
      Picture         =   "frmTCpbGrd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   720
   End
   Begin VB.Frame framereporte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5085
      TabIndex        =   10
      Top             =   4365
      Width           =   3135
      Begin VB.TextBox txtimporte 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   14
         Top             =   400
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   2680
         Picture         =   "frmTCpbGrd.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   400
         Width           =   375
      End
      Begin VB.ComboBox cmbmoneda 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   400
         Width           =   735
      End
      Begin VB.ComboBox cmbDH 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   400
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Importe"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   165
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Debe/Haber"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   15
         Top             =   180
         Width           =   975
      End
   End
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
      Begin VB.CommandButton cmdreportes 
         Caption         =   "Reporte"
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
         Picture         =   "frmTCpbGrd.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "frmTCpbGrd.frx":03DE
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
         Left            =   5040
         TabIndex        =   0
         Top             =   15
         Width           =   2685
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
         Picture         =   "frmTCpbGrd.frx":0528
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
         Picture         =   "frmTCpbGrd.frx":0672
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
         Picture         =   "frmTCpbGrd.frx":0774
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
         Index           =   1
         Left            =   2880
         Picture         =   "frmTCpbGrd.frx":0876
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
         Picture         =   "frmTCpbGrd.frx":0978
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
       uorstCoCta As ADODB.Recordset, _
       uorstCoCCo As ADODB.Recordset, _
       uorstTGAux As ADODB.Recordset, _
       uorstTGTDc As ADODB.Recordset, _
       uorstCOCpbDet As ADODB.Recordset, _
       uorstCOTCbMes As ADODB.Recordset, _
       uorstCoCpbDetRP As ADODB.Recordset
Public uorstCOFjo As ADODB.Recordset, _
       uorstCOFjoDet As ADODB.Recordset, _
       uorstmedio As ADODB.Recordset
       
Private rconex As ADODB.Connection
Private rrecord As ADODB.Recordset


       
'       uorstTGArt As ADODB.Recordset, _
'       uorstTGSvc As ADODB.Recordset
']

Public reporte As ADODB.Recordset
Dim dbconex As ADODB.Connection
Dim sql As String

Private Sub cmd_Click()

If txtImporte.Text = "" Then Beep: MsgBox "Debe Ingresar Importe", vbExclamation: txtImporte.SetFocus: Exit Sub
If cmbmoneda.Text = "" Then Beep: MsgBox "Debe Ingresar Moneda", vbExclamation: cmbmoneda.SetFocus: Exit Sub
If cmbDH.Text = "" Then Beep: MsgBox "Debe Ingresar Debe o Haber", vbExclamation: cmbDH.SetFocus: Exit Sub
  
With rrecord
If .State = adStateOpen Then .Close
    If cmbmoneda = "MN" Then
        .Source = "select CONCAT(cocta.codcta,' ',cocta.detcta),coddro,nrocpb,nroite,fehope,nrodoc,codaux,refdoc,gloite,tpoctb,impmn from cocpbdet inner join cocta on cocpbdet.codemp=cocta.codemp and cocpbdet.pdoano=cocta.pdoano and cocpbdet.codcta=cocta.codcta where cocpbdet.codemp='" & gsCodEmp & "' and cocpbdet.pdoano='" & gsAnoAct & "' and cocpbdet.mespvs='" & gsMesAct & "' and cocpbdet.impmn='" & txtImporte.Text & "' and cocpbdet.tpoctb='" & Left(cmbDH, 1) & "'"
    Else
        .Source = "select CONCAT(cocta.codcta,' ',cocta.detcta),coddro,nrocpb,nroite,fehope,nrodoc,codaux,refdoc,gloite,tpoctb,impme from cocpbdet inner join cocta on cocpbdet.codemp=cocta.codemp and cocpbdet.pdoano=cocta.pdoano and cocpbdet.codcta=cocta.codcta where cocpbdet.codemp='" & gsCodEmp & "' and cocpbdet.pdoano='" & gsAnoAct & "' and cocpbdet.mespvs='" & gsMesAct & "' and cocpbdet.impme='" & txtImporte.Text & "' and cocpbdet.tpoctb='" & Left(cmbDH, 1) & "'"
    End If
    .Open
End With
 
gpEncabezadoRpt frmMain.rptMain, "Importes en " & IIf(cmbmoneda.Text = "MN", "Moneda Nacional", "Moneda Extranjera") & " - " & cmbDH.Text, Date, True, False, rrecord

With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptimportes.rpt"
      .WindowShowExportBtn = True
      .MarginLeft = 240
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
End With
  

End Sub

Private Sub cmdimportes_Click()
  If framereporte.Visible = False Then
      framereporte.Top = 620
      framereporte.Left = 3600
      framereporte.Visible = True
  Else
      framereporte.Visible = False
  End If
End Sub

Private Sub cmdreportes_Click()

    Dim Rst As ADODB.Recordset
    Dim RstConsulta As ADODB.Recordset
    Dim RstReporte As ADODB.Recordset
    
    Dim fila As Integer
    Dim filaconsulta As Integer
    Dim quediario As String
    Dim pasa As Boolean
    
    Set Rst = New ADODB.Recordset
    Set RstConsulta = New ADODB.Recordset
    
    dbconex.Execute "DROP TABLE IF EXISTS tmpdiario "
    sql = "CREATE TABLE IF NOT EXISTS tmpdiario "
    sql = sql & "(codaux varchar(11) NOT NULL, "
    sql = sql & "razaux varchar(60) NOT NULL, "
    sql = sql & "codtdc char(2) NOT NULL, "
    sql = sql & "serdoc char(4) NOT NULL, "
    sql = sql & "nrodoc varchar(10) NOT NULL, "
    sql = sql & "tpognr char(2) NOT NULL, "
    sql = sql & "nroite char(4) NOT NULL,  "
    sql = sql & "codcta varchar(16) NOT NULL, "
    sql = sql & "detcta varchar(60) NOT NULL, "
    sql = sql & "codcco varchar(5) NOT NULL, "
    sql = sql & "detcco varchar(40) NOT NULL, "
    sql = sql & "impmn decimal(12,2) DEFAULT '0', "
    sql = sql & "impme decimal(12,2) DEFAULT '0', "
    sql = sql & "coddro varchar(4) NOT NULL, "
    sql = sql & "nrocpb varchar(6) NOT NULL) "
        
    dbconex.Execute sql
   
    sql = " select a.codaux,a.codtdc,a.serdoc,a.nrodoc "
    sql = sql & " from cocpbdet a "
    sql = sql & " inner join cocpbcab b on a.codemp=b.codemp and a.pdoano=b.pdoano and a.mespvs=b.mespvs and a.coddro=b.coddro and a.nrocpb=b.nrocpb "
    sql = sql & " inner join tgaux aux on a.codemp=aux.codemp and a.codaux=aux.codaux "
    sql = sql & " inner join cocta cta on a.codemp=cta.codemp and a.pdoano=cta.pdoano and a.codcta=cta.codcta "
    sql = sql & " inner join cocco cost on a.codemp=cost.codemp and a.pdoano=cost.pdoano and a.codcco=cost.codcco "
    sql = sql & " where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' and month(b.fehcpb)='" & gsMesAct & "'"
    sql = sql & " group by a.codaux,a.codtdc,a.serdoc,a.nrodoc "
    
    Rst.Open sql, dbconex, adOpenStatic, adLockPessimistic
    
    If Rst.RecordCount = 0 Then
        Exit Sub
    End If
    
    Rst.MoveFirst
    For fila = 0 To Rst.RecordCount - 1
    
        sql = " select a.codcco "
        sql = sql & " from cocpbdet a "
        sql = sql & " inner join cocpbcab b on a.codemp=b.codemp and a.pdoano=b.pdoano and a.mespvs=b.mespvs and a.coddro=b.coddro and a.nrocpb=b.nrocpb "
        sql = sql & " inner join tgaux aux on a.codemp=aux.codemp and a.codaux=aux.codaux "
        sql = sql & " inner join cocta cta on a.codemp=cta.codemp and a.pdoano=cta.pdoano and a.codcta=cta.codcta "
        sql = sql & " inner join cocco cost on a.codemp=cost.codemp and a.pdoano=cost.pdoano and a.codcco=cost.codcco "
        sql = sql & " where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' and month(b.fehcpb)='" & gsMesAct & "'"
        sql = sql & " and a.codaux='" & Rst.Fields(0) & "' and a.codtdc='" & Rst.Fields(1) & "' and a.serdoc='" & Rst.Fields(2) & "' and a.nrodoc='" & Rst.Fields(3) & "'"
            
        RstConsulta.Open sql, dbconex, adOpenStatic, adLockPessimistic
        If RstConsulta.RecordCount > 0 Then
        
        RstConsulta.MoveFirst
            For filaconsulta = 0 To RstConsulta.RecordCount - 1
            
            If filaconsulta = 0 Then
                quediario = Left(RstConsulta.Fields(0), 1)
            Else
            
                If Left(RstConsulta(0), 1) <> quediario Then
                    pasa = True
                End If
            
            End If
            RstConsulta.MoveNext
            Next
            
        If pasa = True Then
        
                sql = "INSERT INTO tmpdiario(codaux,razaux,codtdc,serdoc,nrodoc,tpognr,nroite,codcta,detcta,codcco,detcco,impmn,impme,coddro,nrocpb )"
                sql = sql & " select a.codaux,aux.razaux,a.codtdc,a.serdoc,a.nrodoc,a.tpognr,a.nroite,a.codcta,cta.detcta,a.codcco,cost.detcco,a.impmn,a.impme,b.coddro,b.nrocpb  "
                sql = sql & " from cocpbdet a "
                sql = sql & " inner join cocpbcab b on a.codemp=b.codemp and a.pdoano=b.pdoano and a.mespvs=b.mespvs and a.coddro=b.coddro and a.nrocpb=b.nrocpb "
                sql = sql & " inner join tgaux aux on a.codemp=aux.codemp and a.codaux=aux.codaux "
                sql = sql & " inner join cocta cta on a.codemp=cta.codemp and a.pdoano=cta.pdoano and a.codcta=cta.codcta "
                sql = sql & " inner join cocco cost on a.codemp=cost.codemp and a.pdoano=cost.pdoano and a.codcco=cost.codcco "
                sql = sql & " where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' and month(b.fehcpb)='" & gsMesAct & "'"
                sql = sql & " and a.codaux='" & Rst.Fields(0) & "' and a.codtdc='" & Rst.Fields(1) & "' and a.serdoc='" & Rst.Fields(2) & "' and a.nrodoc='" & Rst.Fields(3) & "'"
            
                dbconex.Execute sql
                
        End If
                    
        quediario = ""
        pasa = False
        
        End If
    
        
        RstConsulta.Close
        Rst.MoveNext
    Next
    
    Rst.Close
    
   Set reporte = New ADODB.Recordset

   With reporte
      .ActiveConnection = dbconex
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
   End With
     
   With reporte
   If .State = adStateOpen Then .Close
        .Source = " select codaux,razaux,codtdc,serdoc,nrodoc,tpognr,nroite,codcta,detcta,codcco,detcco,format(impmn,2),format(impme,2),coddro,nrocpb "
        .Source = .Source & " from tmpdiario "
        .Open
   End With
      
   gpEncabezadoRpt frmMain.rptMain, "Listado Ref. Diario", Date, True, False, reporte
   
   With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLRcostos.rpt"
      '.MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
   End With

   dbconex.Execute "DROP TABLE IF EXISTS tmpdiario "
   
End Sub

Private Sub Form_Load()

   Set dbconex = New ADODB.Connection
   dbconex.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";connection="
   dbconex.CursorLocation = adUseClient
   dbconex.Open

   cmbDH.AddItem "Debe"
   cmbDH.AddItem "Haber"
   cmbmoneda.AddItem "MN"
   cmbmoneda.AddItem "ME"

 '[Recordsets                          'Cambiar.
   usConnStrgSele_0 = "SELECT CodDro, NroCpb, FehCpb, "
   usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "GloCpb, ", "GloCpbx, ")
   usConnStrgSele_0 = usConnStrgSele_0 & "(CASE TpoGnr WHEN " & TPOGNR_DRO & " THEN '" & TPOGNR_DRO_TXT & "' WHEN " & TPOGNR_CPR & " THEN '" & TPOGNR_CPR_TXT & "' WHEN " & TPOGNR_VTA & " THEN '" & TPOGNR_VTA_TXT & "' WHEN " & TPOGNR_HPR & " THEN '" & TPOGNR_HPR_TXT & "' WHEN " & TPOGNR_DST & " THEN '" & TPOGNR_DST_TXT & "' WHEN " & TPOGNR_DCA & " THEN '" & TPOGNR_DCA_TXT & "' WHEN " & TPOGNR_APE & " THEN '" & TPOGNR_APE_TXT & "' WHEN " & TPOGNR_CIE & " THEN '" & TPOGNR_CIE_TXT & "' WHEN " & TPOGNR_DRP & " THEN '" & TPOGNR_DRP_TXT & "' ELSE '" & TPOGNR_BAN_TXT & "' END) AS ccTpoGnr, "
   usConnStrgSele_0 = usConnStrgSele_0 & Choose(gsIdioma, "GloCpbx, ", "GloCpb, ")
   usConnStrgSele_0 = usConnStrgSele_0 & "TpoGnr, MesPvs, IndNCu, IndAnu, UsrCre, FyHCre, UsrMdf, FyHMdf, codemp, pdoano, "
   usConnStrgSele_0 = usConnStrgSele_0 & IIf(ps_Plataforma = pSrvMySql, "Concat(CodDro, NroCpb)", "(CodDro+NroCpb)") & " AS cLlave "
   usConnStrgSele_0 = usConnStrgSele_0 & "FROM COCpbCab "
   usConnStrgSele_0 = usConnStrgSele_0 & "WHERE codemp='" & gsCodEmp & "' "
   usConnStrgSele_0 = usConnStrgSele_0 & "AND pdoano='" & gsAnoAct & "' "
   usConnStrgSele_0 = usConnStrgSele_0 & "AND MesPvs='" & gsMesAct & "' "
   usConnStrgOrde_0 = "ORDER BY CodDro, NroCpb"
   
   usConnStrgSele_1 = "SELECT COCpbDet.NroIte, COCpbDet.CodCta, COCpbDet.CodCCo, COCpbDet.CodAux, TGTDc.AbvTDc , COCpbDet.SerDoc, COCpbDet.NroDoc, "
   usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "COCpbDet.GloIte, ", "COCpbDet.GloItex, ")
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN COCpbDet.ImpMN ELSE 0 END) AS cImpMN_Deb, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN 0 ELSE COCpbDet.ImpMN END) AS cImpMN_Hab, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE COCpbDet.TpoGnr WHEN " & TPOGNR_DRO & " THEN '" & TPOGNR_DRO_TXT & "' WHEN " & TPOGNR_CPR & " THEN '" & TPOGNR_CPR_TXT & "' WHEN " & TPOGNR_VTA & " THEN '" & TPOGNR_VTA_TXT & "' WHEN " & TPOGNR_HPR & " THEN '" & TPOGNR_HPR_TXT & "' WHEN " & TPOGNR_DST & " THEN '" & TPOGNR_DST_TXT & "' WHEN " & TPOGNR_DCA & " THEN '" & TPOGNR_DCA_TXT & "' WHEN " & TPOGNR_APE & " THEN '" & TPOGNR_APE_TXT & "' WHEN " & TPOGNR_CIE & " THEN '" & TPOGNR_CIE_TXT & "' WHEN " & TPOGNR_DRP & " THEN '" & TPOGNR_DRP_TXT & "' ELSE '" & TPOGNR_BAN_TXT & "' END) AS ccTpoGnr, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN COCpbDet.ImpME ELSE 0 END) AS cImpME_Deb, "
   usConnStrgSele_1 = usConnStrgSele_1 & "(CASE COCpbDet.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN 0 ELSE COCpbDet.ImpME END) AS cImpME_Hab, "
   usConnStrgSele_1 = usConnStrgSele_1 & "COCpbDet.BlqIte, COCpbDet.TpoMon, COCpbDet.ImpTcb, COCpbDet.TpoTCb, COCpbDet.TpoGnr, COCpbDet.CodTDc, "
   usConnStrgSele_1 = usConnStrgSele_1 & "COCpbDet.RefDoc, COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.TpoCtb, COCpbDet.TpoPvs, COCpbDet.MesPvs, "
   usConnStrgSele_1 = usConnStrgSele_1 & "COCpbDet.FehOpe, COCpbDet.FeEDoc, COCpbDet.FeVDoc, COCpbDet.FeRDoc, COCpbDet.ImpMN, COCpbDet.ImpME, "
   usConnStrgSele_1 = usConnStrgSele_1 & "COCpbDet.IndFjo_Det, COCpbDet.IndGnr_RP, COCpbDet.UsrCre, COCpbDet.FyHCre, COCpbDet.UsrMdf, COCpbDet.FyHMdf, "
   usConnStrgSele_1 = usConnStrgSele_1 & "COCpbDet.codemp, COCpbDet.pdoano, COCpbDet.pdocpr, COCpbDet.codcon, "
   usConnStrgSele_1 = usConnStrgSele_1 & Choose(gsIdioma, "COCpbDet.GloItex, ", "COCpbDet.GloIte, ")
   usConnStrgSele_1 = usConnStrgSele_1 & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte)", "(COCpbDet.CodDro+COCpbDet.NroCpb+RTrim(COCpbDet.NroIte))") & " AS cLlave, COCpbDet.tpodoc "
   usConnStrgSele_1 = usConnStrgSele_1 & "FROM (COCpbDet "
   usConnStrgSele_1 = usConnStrgSele_1 & "LEFT JOIN TGTDc AS TGTDc ON COCpbDet.codemp=TGTDc.codemp AND COCpbDet.CodTDc=TGTDc.CodTDc) "
   usConnStrgWher_1 = "WHERE COCpbDet.codemp='" & gsCodEmp & "' "
   usConnStrgWher_1 = usConnStrgWher_1 & "AND COCpbDet.pdoano='" & gsAnoAct & "' AND COCpbDet.MesPvs='" & gsMesAct & "' "
   usConnStrgWher_1 = usConnStrgWher_1 & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COCpbDet.CodDro, COCpbDet.NroCpb)", "(COCpbDet.CodDro+COCpbDet.NroCpb)") & "=' ' "
   usConnStrgOrde_1 = "ORDER BY COCpbDet.NroIte, COCpbDet.BlqIte"
   
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
   Set uorstCOCpbDet = New ADODB.Recordset
   Set uorstCOTCbMes = New ADODB.Recordset
   Set uorstCoCpbDetRP = New ADODB.Recordset
   Set uorstCOFjo = New ADODB.Recordset
   Set uorstCOFjoDet = New ADODB.Recordset
   Set uorstmedio = New ADODB.Recordset
   
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
     .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro, codemp, Cpb" & gsMesAct & ", "
     .Source = .Source & "codemp, pdoano "
     .Source = .Source & "FROM CODro "
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
    .Source = .Source & ",tpomon " '2015-08-20 correccion tipo mon cta
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
   With uorstCoCpbDetRP
     .ActiveConnection = uocnnMain
     .Source = "SELECT MesPvs, CodDro, NroCpb, NroIte, CodAux, CodCta, "
     .Source = .Source & "CodTDc, SerDoc, NroDoc, ImpMN, ImpME, "
     .Source = .Source & "CodTDc_RtcPcp, SerDoc_RtcPcp, NroDoc_RtcPcp, "
     .Source = .Source & "feEDoc_RtcPcp, ImpMN_RtcPcp, ImpME_RtcPcp, IndRtcPcp, "
     .Source = .Source & "UsrCre, FyHCre, UsrMdf, FyHMdf, codemp, pdoano, "
     .Source = .Source & IIf(ps_Plataforma = pSrvMySql, "CONVERT(CONCAT(MesPvs, CodDro, NroCpb, NroIte), char(14))", "(MesPvs+CodDro+NroCpb+RTrim(NroIte))") & " AS cLlave "
     .Source = .Source & "FROM CoCpbDetRP "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND MesPvs='" & gsMesAct & "' "
     .Source = .Source & "ORDER BY CodDro, NroCpb, NroIte"
'     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
     .Properties("Unique Table").Value = "CoCpbDetRP"
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
      .Source = .Source & "FROM CoCpbDetFjo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "CoCpbDetFjo"
   End With
   
   With uorstmedio
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.codmed, abvmed,desmed "
    .Source = .Source & "FROM bnmediopago a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "'"
    ''     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
   
 ']
 
  Set rconex = New ADODB.Connection
  Set rrecord = New ADODB.Recordset
 
  With rconex
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
  End With
  
  With rrecord
      .ActiveConnection = rconex
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
  End With
  
 
   
  '[ Elimino y creo tabla temporal de detalle de flujo
  If ps_Plataforma = pSrvMySql Then
    uocnnMain.Execute "DROP TABLE IF EXISTS tmpCoCpbDetFjo"
    uocnnMain.Execute "CREATE TEMPORARY TABLE tmpCoCpbDetFjo SELECT * FROM CoCpbDetFjo WHERE CodFjo='tmpflujo'"
  ElseIf ps_Plataforma = pSrvSql Then
    ' Activo detector de errores
    On Error Resume Next
    uocnnMain.Execute "DROP TABLE " & ps_Prefijo & "tmpCoCpbDetFjo"
    If Not (Err.Number = -2147217865 Or Err.Number = 0) Then
      MsgBox Err.Description, vbInformation
    End If
    On Error GoTo 0
    uocnnMain.Execute "SELECT * INTO #tmpCoCpbDetFjo FROM CoCpbDetFjo WHERE CodFjo='tmpflujo'"
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
   framereporte.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasGrid(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   gpTUg_Resize Me
   
End Sub

Private Sub cmbmoneda_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbDH_KeyPress(KeyAscii As Integer)
KeyAscii = 0
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
   If gbCieCpb Then MsgBox TEXT_9016, vbCritical: Exit Sub
  '[ No pertence al Formulario - Agregado por Angel
  With uorstMain_1
    .Close
    .Source = usConnStrgSele_1 & " WHERE COCpbDet.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND COCpbDet.pdoano='" & gsAnoAct & "' AND COCpbDet.mespvs='" & gsMesAct & "' "
    .Source = .Source & "AND COCpbDet.CodDro='    ' " & usConnStrgOrde_1
    .Open
    .Properties("Unique Table").Value = "COCpbDet"
  End With
  gpTUg_Nuevo Me, frmTCpbCab          'Cambiar Formulario de Datos.
End Sub

Public Sub cmdRevisar_click()
  On Error GoTo Err
  
  'Verificación de existencia de ítemes.
  If uorstMain_0.RecordCount = 0 Then MsgBox TEXT_8001, vbCritical: Exit Sub

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
   If gbCieCpb Then MsgBox TEXT_9016, vbCritical: Exit Sub
   
   'Verificación de existencia de ítemes.
   If uorstMain_0.RecordCount = 0 Then
      MsgBox Choose(gsIdioma, "No hay datos creados.", "There are not created data"), vbCritical
      Exit Sub
   End If

   If frmTCpbGrd.uorstMain_0!tpognr <> TPOGNR_DRO Then
      MsgBox Choose(gsIdioma, "No se Puede Eliminar este Comprobante", "This Voucher can not be eliminated"), vbInformation
      Exit Sub
   End If

   'Mensaje de verificación            'Cambiar.
   If MsgBox(TEXT_1021 & " " & Trim(dgrMain.Columns(0)) & " (" & Trim(dgrMain.Columns(1)) & ")?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
      uocnnMain.BeginTrans
      uorstMain_0.Delete
      uocnnMain.CommitTrans
      
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
  frmTCpbGrd.uorstMain_0.Requery
  frmTCpbGrd.ppDatosGrid
  dgrMain.SetFocus
End Sub

Public Sub cmdImprimir_Click(Index As Integer)
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
  s_Sentencia = s_Sentencia & "FROM CoCpbDet "
  s_Sentencia = s_Sentencia & "WHERE codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND pdoano='" & s_Ano & "' "
  s_Sentencia = s_Sentencia & "AND MesPvs='" & s_Mes & "' "
  s_Sentencia = s_Sentencia & "AND CodDro='" & s_Diario & "' "
  s_Sentencia = s_Sentencia & "AND NroCpb='" & s_Comprobante & "'"
  Set porstRetorno = New ADODB.Recordset
  With porstRetorno
    .ActiveConnection = frmTCpbGrd.uocnnMain
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

Private Sub txtimporte_Change()
  If Not IsNumeric(txtImporte.Text) Then
     txtImporte.Text = ""
  End If
End Sub
