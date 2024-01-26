VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTCprGrd 
   Caption         =   "[Entidad]"
   ClientHeight    =   6390
   ClientLeft      =   3075
   ClientTop       =   2160
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   Picture         =   "frmTCprGrd.frx":0000
   ScaleHeight     =   6390
   ScaleWidth      =   9195
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   9195
      Begin VB.CommandButton cmdreporte 
         Appearance      =   0  'Flat
         Caption         =   "&Reporte"
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
         Picture         =   "frmTCprGrd.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   850
      End
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
         Picture         =   "frmTCprGrd.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmTCprGrd.frx":0AC6
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
         Picture         =   "frmTCprGrd.frx":0BC8
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   5205
         TabIndex        =   0
         Top             =   0
         Width           =   2655
         Begin VB.TextBox txtBuscar 
            Height          =   285
            Left            =   120
            TabIndex        =   6
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
         Picture         =   "frmTCprGrd.frx":0D12
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
         Picture         =   "frmTCprGrd.frx":0E5C
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
         Picture         =   "frmTCprGrd.frx":0F5E
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
         Picture         =   "frmTCprGrd.frx":1060
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTCprGrd"
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
Public uorstCoCta As ADODB.Recordset
Public uorstCoCCo As ADODB.Recordset
Public uorstCoCCox As ADODB.Recordset
Public uorstCoCCoy As ADODB.Recordset
Public uorstCODro As ADODB.Recordset
Public uorstCoAsiTipo As ADODB.Recordset
Public uorstCOCprDocCta As ADODB.Recordset
Public uorstCOCprDocCCo As ADODB.Recordset
Public uorstCOCpbCab As ADODB.Recordset
Public uorstCOCpbDet As ADODB.Recordset
Public uorstTemporal As ADODB.Recordset
Private porstCancel As ADODB.Recordset

Public uorstcodetrac As ADODB.Recordset '2015-07-02 adic tabla detrac


Public usConnStrgSele_COCprDocCta As String, _
       usConnStrgWher_COCprDocCta As String, _
       usConnStrgOrde_COCprDocCta As String
Public usConnStrgSele_COCprDocCCo As String, _
       usConnStrgWher_COCprDocCCo As String, _
       usConnStrgOrde_COCprDocCCo As String
Public usConnStrgSele_COCpbDet As String, _
       usConnStrgWher_COCpbDet As String, _
       usConnStrgOrde_COCpbDet As String

Public ubGrabaMas As Byte  '0:Nuevo documento 1:Cuenta grabado por cmdMas 2:Cuenta grabada directa.

Public reporte As ADODB.Recordset

'[Repetir en frmTCpr y frmTCprMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']

Dim dbconex As ADODB.Connection
Dim sql As String

Private Sub cmdGenera_Click()
  Dim s_Sentencia As String
  
  'Verificación de Mes Cerrado.
  If gbCieCpr Then
    MsgBox TEXT_9016, vbCritical
    Exit Sub
  End If
  ' Genero información
  With porstCancel
    .Source = "SELECT cpr.CodDro, cpr.NroCpb, cpr.CodAux, cpr.SerDoc, cpr.NroDoc, "
    .Source = .Source & "cpr.FeEDoc, cpr.TpoMon, cpr.pdocpr, "
    .Source = .Source & "cpr.CodTDc, cpr.FehOpe, cpr.FeVDoc, "
    .Source = .Source & "cpr.FeRDoc, cpr.ImpTCb, cpr.PctIGV, cpr.PctISC, "
    .Source = .Source & "cpr.RefDoc, cpr.GloDoc, cpr.GloDocx, "
    .Source = .Source & "cpr.MesPvs, cpr.codasi, "
    .Source = .Source & "cpr.NroCDt, cpr.FehCDt, "
    .Source = .Source & "cpr.ImpOGr_MN, cpr.ImpOGN_MN, cpr.ImpONG_MN, cpr.ImpExo_MN, "
    .Source = .Source & "cpr.ImpIGV_MN, cpr.ImpISC_MN, cpr.ImpOIm_MN, cpr.ImpTot_MN, "
    .Source = .Source & "cpr.ImpOGr_ME, cpr.ImpOGN_ME, cpr.ImpONG_ME, cpr.ImpExo_ME, "
    .Source = .Source & "cpr.ImpIGV_ME, cpr.ImpISC_ME, cpr.ImpOIm_ME, cpr.ImpTot_ME, "
    .Source = .Source & "cpr.IndCta_OGr, cpr.IndCta_OGN, cpr.IndCta_ONG, cpr.IndCta_Exo, "
    .Source = .Source & "cpr.IndCta_IGV, cpr.IndCta_ISC, cpr.IndCta_OIm, cpr.IndCta_Tot, "
    .Source = .Source & "cpr.IndCDt, cpr.IndPreGen, cpr.IndGen, cpr.IndAnu, "
    .Source = .Source & "cpr.ImpIGV_OGr_MN, cpr.ImpIGV_OGN_MN, cpr.ImpIGV_ONG_MN, "
    .Source = .Source & "cpr.ImpIGV_OGr_ME, cpr.ImpIGV_OGN_ME, cpr.ImpIGV_ONG_ME, "
    .Source = .Source & "cpr.codemp, cpr.pdoano "
    .Source = .Source & "FROM COCprDoc cpr "
    .Source = .Source & "LEFT JOIN CoCpbCab cab ON cpr.codemp=cab.codemp AND cpr.pdoano=cab.pdoano AND cpr.MesPvs=cab.MesPvs AND cpr.CodDro=cab.CodDro AND cpr.NroCpb=cab.NroCpb "
    .Source = .Source & "WHERE cpr.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND cpr.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND cpr.MesPvs='" & gsMesAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(cpr.IndGen", "ISNULL(cpr.IndGen") & ", '0')='0' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(cab.CodDro, cab.NroCpb)", "ISNULL((cab.CodDro+cab.NroCpb)") & ", '')='' "
    .Source = .Source & "ORDER BY cpr.CodDro, cpr.NroDoc"
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

Private Sub cmdreporte_Click()

    Dim Rst As ADODB.Recordset
    Dim RstConsulta As ADODB.Recordset
    Dim RstReporte As ADODB.Recordset
    
    Dim fila As Integer
    Dim filaconsulta As Integer
    Dim quecostos As String
    Dim pasa As Boolean
    
    Set Rst = New ADODB.Recordset
    Set RstConsulta = New ADODB.Recordset
    
    dbconex.Execute "DROP TABLE IF EXISTS tmpcostos "
    sql = "CREATE TABLE IF NOT EXISTS tmpcostos "
    sql = sql & "(codaux varchar(11) NOT NULL, "
    sql = sql & "razaux varchar(60) NOT NULL, "
    sql = sql & "codtdc char(2) NOT NULL, "
    sql = sql & "serdoc char(4) NOT NULL, "
    sql = sql & "nrodoc varchar(10) NOT NULL, "
    sql = sql & "tpocnc char(2) NOT NULL, "
    sql = sql & "orden char(2) NOT NULL,  "
    sql = sql & "codcta varchar(8) NOT NULL, "
    sql = sql & "detcta varchar(60) NOT NULL, "
    sql = sql & "codcco varchar(5) NOT NULL, "
    sql = sql & "detcco varchar(40) NOT NULL, "
    sql = sql & "impcco_mn decimal(12,2) DEFAULT '0', "
    sql = sql & "impcco_me decimal(12,2) DEFAULT '0', "
    sql = sql & "coddro varchar(4) NOT NULL, "
    sql = sql & "nrocpb varchar(6) NOT NULL) "
    
    
    dbconex.Execute sql
   
    sql = " select a.codaux,a.codtdc,a.serdoc,a.nrodoc "
    sql = sql & " from cocprdoccco a "
    sql = sql & " inner join cocprdoc b on a.codemp=b.codemp and a.pdoano=b.pdoano and a.codaux=b.codaux and a.codtdc=b.codtdc and a.serdoc=b.serdoc and a.nrodoc=b.nrodoc "
    sql = sql & " inner join tgaux aux on a.codemp=aux.codemp and a.codaux=aux.codaux "
    sql = sql & " inner join cocta cta on a.codemp=cta.codemp and a.pdoano=cta.pdoano and a.codcta=cta.codcta "
    sql = sql & " inner join cocco cost on a.codemp=cost.codemp and a.pdoano=cost.pdoano and a.codcco=cost.codcco "
    sql = sql & " where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' and month(b.fehope)='" & gsMesAct & "'"
    sql = sql & " group by a.codaux,a.codtdc,a.serdoc,a.nrodoc "
    
    Rst.Open sql, dbconex, adOpenStatic, adLockPessimistic
    
    If Rst.RecordCount = 0 Then
        Exit Sub
    End If
    
    Rst.MoveFirst
    For fila = 0 To Rst.RecordCount - 1
    
        sql = " select a.codcco "
        sql = sql & " from cocprdoccco a "
        sql = sql & " inner join cocprdoc b on a.codemp=b.codemp and a.pdoano=b.pdoano and a.codaux=b.codaux and a.codtdc=b.codtdc and a.serdoc=b.serdoc and a.nrodoc=b.nrodoc "
        sql = sql & " inner join tgaux aux on a.codemp=aux.codemp and a.codaux=aux.codaux "
        sql = sql & " inner join cocta cta on a.codemp=cta.codemp and a.pdoano=cta.pdoano and a.codcta=cta.codcta "
        sql = sql & " inner join cocco cost on a.codemp=cost.codemp and a.pdoano=cost.pdoano and a.codcco=cost.codcco "
        sql = sql & " where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' and month(b.fehope)='" & gsMesAct & "'"
        sql = sql & " and a.codaux='" & Rst.Fields(0) & "' and a.codtdc='" & Rst.Fields(1) & "' and a.serdoc='" & Rst.Fields(2) & "' and a.nrodoc='" & Rst.Fields(3) & "'"
            
        RstConsulta.Open sql, dbconex, adOpenStatic, adLockPessimistic
    
        RstConsulta.MoveFirst
            For filaconsulta = 0 To RstConsulta.RecordCount - 1
            
            If filaconsulta = 0 Then
                quecostos = Left(RstConsulta.Fields(0), 1)
            Else
            
                If Left(RstConsulta(0), 1) <> quecostos Then
                    pasa = True
                End If
            
            End If
            RstConsulta.MoveNext
            Next
            
        If pasa = True Then
        
                sql = "INSERT INTO tmpcostos(codaux,razaux,codtdc,serdoc,nrodoc,tpocnc,orden,codcta,detcta,codcco,detcco,impcco_mn,impcco_me,coddro,nrocpb )"
                sql = sql & " select a.codaux,aux.razaux,a.codtdc,a.serdoc,a.nrodoc,a.tpocnc,a.orden,a.codcta,cta.detcta,a.codcco,cost.detcco,a.impcco_mn,a.impcco_me,b.coddro,b.nrocpb  "
                sql = sql & " from cocprdoccco a "
                sql = sql & " inner join cocprdoc b on a.codemp=b.codemp and a.pdoano=b.pdoano and a.codaux=b.codaux and a.codtdc=b.codtdc and a.serdoc=b.serdoc and a.nrodoc=b.nrodoc "
                sql = sql & " inner join tgaux aux on a.codemp=aux.codemp and a.codaux=aux.codaux "
                sql = sql & " inner join cocta cta on a.codemp=cta.codemp and a.pdoano=cta.pdoano and a.codcta=cta.codcta "
                sql = sql & " inner join cocco cost on a.codemp=cost.codemp and a.pdoano=cost.pdoano and a.codcco=cost.codcco "
                sql = sql & " where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' and month(b.fehope)='" & gsMesAct & "'"
                sql = sql & " and a.codaux='" & Rst.Fields(0) & "' and a.codtdc='" & Rst.Fields(1) & "' and a.serdoc='" & Rst.Fields(2) & "' and a.nrodoc='" & Rst.Fields(3) & "'"
            
                dbconex.Execute sql
                
        End If
                    
        quecostos = ""
        pasa = False
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
        .Source = " select codaux,razaux,codtdc,serdoc,nrodoc,tpocnc,orden,codcta,detcta,codcco,detcco,format(impcco_mn,2),format(impcco_me,2),coddro,nrocpb "
        .Source = .Source & " from tmpcostos "
        .Open
   End With
      
   gpEncabezadoRpt frmMain.rptMain, "Listado Ref. Centro de Costos", Date, True, False, reporte
   
   With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & "rptLRcostos.rpt"
      '.MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
   End With

   dbconex.Execute "DROP TABLE IF EXISTS tmpcostos "
   
End Sub

Private Sub cmdVerificar_Click()
  '[Datos del formulario de impresión.  'Cambiar.
  Dim s_Sentencia As String
  Dim porstMRp As New ADODB.Recordset
 
  s_Sentencia = "SELECT CodTDc, " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(SerDoc, '-',NroDoc)", "(SerDoc+'-'+NroDoc)") & " AS cDocumento, FehOpe, FeEdoc, "
  s_Sentencia = s_Sentencia & Choose(gsIdioma, "GloDoc", "GloDocx") & " AS GloDoc, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN 'S/.' ELSE 'US$' END) AS cMoneda, "
  s_Sentencia = s_Sentencia & "ROUND((CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN (ImpOGr_MN+ImpOGn_MN+ImpONG_MN) ELSE (ImpOGr_ME+ImpOGn_ME+ImpONG_ME) END), 2) AS cImpBas, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpExo_MN ELSE ImpExo_ME END) AS cExonerado, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpIGV_MN ELSE ImpIGV_ME END) AS cImpIGV, "
  s_Sentencia = s_Sentencia & "(CASE TpoMon WHEN '" & TPOMON_NAC & "' THEN ImpTot_MN ELSE ImpTot_ME END) AS cImpTotal, "
  s_Sentencia = s_Sentencia & "a.CodDro, a.NroCpb "
  s_Sentencia = s_Sentencia & "FROM CoCprDoc AS a "
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

  gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "DOCUMENTOS DE COMPRAS NO CONTABILIZADOS", "NOT COUNTED DOCUMENTS OF PURCHASES"), Date, True, False, porstMRp
  With frmMain.rptMain
    '[Datos y parámetros del reporte.  'Cambiar.
    .ReportFileName = gsRutRpt & "rptLCprCpb.rpt"
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

Private Sub Form_Load()

    Set dbconex = New ADODB.Connection
    dbconex.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";connection="
    dbconex.CursorLocation = adUseClient
    dbconex.Open

 '[Recordsets                          'Cambiar.
  psConnStrgSele_Grd = "SELECT COCprDoc.CodDro, COCprDoc.NroCpb, COCprDoc.CodAux, b.RazAux, c.AbvTDc, COCprDoc.SerDoc, COCprDoc.NroDoc, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "COCprDoc.FeEDoc, COCprDoc.TpoMon, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE COCprDoc.TpoMon WHEN '" & TPOMON_NAC & "' THEN COCprDoc.ImpTot_MN ELSE COCprDoc.ImpTot_ME END) as cImpTot, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "(CASE COCprDoc.IndGen WHEN -1 THEN 'x' ELSE ' ' END) as cIndGen, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "b.CodAux, c.CodTDc, "
  psConnStrgSele_Grd = psConnStrgSele_Grd & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COCprDoc.CodAux, COCprDoc.CodTDc, COCprDoc.SerDoc, COCprDoc.NroDoc)", "(COCprDoc.CodAux+COCprDoc.CodTDc+COCprDoc.SerDoc+COCprDoc.NroDoc)") & " AS cLlave "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "FROM (COCprDoc "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGAux b ON COCprDoc.codemp = b.codemp AND COCprDoc.CodAux = b.CodAux) "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "LEFT JOIN TGTDc c ON COCprDoc.codemp = c.codemp AND COCprDoc.CodTDc = c.CodTDc "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "WHERE COCprDoc.codemp='" & gsCodEmp & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND COCprDoc.pdoano='" & gsAnoAct & "' "
  psConnStrgSele_Grd = psConnStrgSele_Grd & "AND COCprDoc.MesPvs='" & gsMesAct & "' "
   
  psConnStrgSele = "SELECT COCprDoc.CodDro, COCprDoc.NroCpb, COCprDoc.CodAux, COCprDoc.SerDoc, COCprDoc.NroDoc, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.FeEDoc, COCprDoc.TpoMon, "
  psConnStrgSele = psConnStrgSele & "(CASE COCprDoc.TpoMon WHEN '" & TPOMON_NAC & "' THEN COCprDoc.ImpTot_MN ELSE COCprDoc.ImpTot_ME END) AS cImpTot, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.CodTDc, COCprDoc.FehOpe, COCprDoc.FeEDoc, COCprDoc.FeVDoc, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.FeRDoc, COCprDoc.ImpTCb, COCprDoc.PctIGV, COCprDoc.PctISC, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.RefDoc, COCprDoc.GloDoc, COCprDoc.GloDocx, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.MesPvs, COCprDoc.codasi, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.NroCDt, COCprDoc.FehCDt, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.ImpOGr_MN, COCprDoc.ImpOGN_MN, COCprDoc.ImpONG_MN, COCprDoc.ImpExo_MN, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.ImpIGV_MN, COCprDoc.ImpISC_MN, COCprDoc.ImpOIm_MN, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.impoi1_mn, COCprDoc.impoi2_mn, COCprDoc.impoi3_mn, COCprDoc.ImpTot_MN, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.ImpOGr_ME, COCprDoc.ImpOGN_ME, COCprDoc.ImpONG_ME, COCprDoc.ImpExo_ME, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.ImpIGV_ME, COCprDoc.ImpISC_ME, COCprDoc.ImpOIm_ME, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.impoi1_me, COCprDoc.impoi2_me, COCprDoc.impoi3_me, COCprDoc.ImpTot_Me, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.IndCta_OGr, COCprDoc.IndCta_OGN, COCprDoc.IndCta_ONG, COCprDoc.IndCta_Exo, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.IndCta_IGV, COCprDoc.IndCta_ISC, COCprDoc.IndCta_OIm, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.indcta_oi1, COCprDoc.indcta_oi2, COCprDoc.indcta_oi3, COCprDoc.indcta_tot, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.IndCDt, COCprDoc.IndPreGen, COCprDoc.IndGen, COCprDoc.IndAnu, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.ImpIGV_OGr_MN, COCprDoc.ImpIGV_OGN_MN, COCprDoc.ImpIGV_ONG_MN, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.ImpIGV_OGr_ME, COCprDoc.ImpIGV_OGN_ME, COCprDoc.ImpIGV_ONG_ME, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.UsrCre, COCprDoc.FyHCre, COCprDoc.UsrMdf, COCprDoc.FyHMdf, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.codemp, COCprDoc.pdoano, COCprDoc.pdocpr, COCprDoc.codcon, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.tpoimpuesto, COCprDoc.categoriadoc, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.indcprext, COCprDoc.codaduana, COCprDoc.annodua, COCprDoc.nrodua, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.indreten, COCprDoc.tsadetrac, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.pctdetrac, " '2015-07-02 adic tabla detrac
  psConnStrgSele = psConnStrgSele & "COCprDoc.codtdc_ref, COCprDoc.serdoc_ref, COCprDoc.nrodoc_ref, COCprDoc.feedoc_ref, "
  psConnStrgSele = psConnStrgSele & "COCprDoc.impbasref_mn, COCprDoc.impigvref_mn, COCprDoc.impbasref_me, COCprDoc.impigvref_me, "
  psConnStrgSele = psConnStrgSele & IIf(ps_Plataforma = pSrvMySql, "CONCAT(COCprDoc.CodAux, COCprDoc.CodTDc, COCprDoc.SerDoc, COCprDoc.NroDoc)", "(COCprDoc.CodAux+COCprDoc.CodTDc+COCprDoc.SerDoc+COCprDoc.NroDoc)") & " AS cLlave "
  psConnStrgSele = psConnStrgSele & "FROM COCprDoc "
  psConnStrgSele = psConnStrgSele & "WHERE COCprDoc.codemp='" & gsCodEmp & "' "
  psConnStrgSele = psConnStrgSele & "AND COCprDoc.pdoano='" & gsAnoAct & "' "
  psConnStrgSele = psConnStrgSele & "AND COCprDoc.MesPvs='" & gsMesAct & "' "
  
  psConnStrgOrde = "ORDER BY COCprDoc.CodAux, COCprDoc.CodTDc, COCprDoc.SerDoc, COCprDoc.NroDoc"
  
  usConnStrgSele_COCprDocCta = "SELECT COCprDocCta.CodCta, COCprDocCta.ImpCta_MN, COCprDocCta.ImpCta_ME, "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & Choose(gsIdioma, "COCprDocCta.GloDet, ", "COCprDocCta.GloDetx, ") & "COCprDocCta.CodRuc, "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & "COCprDocCta.CodAux, COCprDocCta.CodTDc, COCprDocCta.SerDoc, COCprDocCta.NroDoc, "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & "COCprDocCta.TpoCnc, COCprDocCta.Orden, "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & IIf(ps_Plataforma = pSrvMySql, "Concat(COCprDocCta.CodAux, COCprDocCta.CodTDc, COCprDocCta.SerDoc, COCprDocCta.NroDoc, COCprDocCta.TpoCnc, COCprDocCta.Orden)", "(COCprDocCta.CodAux+COCprDocCta.CodTDc+COCprDocCta.SerDoc+COCprDocCta.NroDoc+RTrim(COCprDocCta.TpoCnc)+COCprDocCta.Orden)") & " AS cLlave, "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & IIf(ps_Plataforma = pSrvMySql, "Concat(COCprDocCta.CodAux, COCprDocCta.CodTDc, COCprDocCta.SerDoc, COCprDocCta.NroDoc, COCprDocCta.TpoCnc, COCprDocCta.Orden, COCprDocCta.CodCta)", "(COCprDocCta.CodAux+COCprDocCta.CodTDc+COCprDocCta.SerDoc+COCprDocCta.NroDoc+RTrim(COCprDocCta.TpoCnc)+COCprDocCta.Orden+COCprDocCta.CodCta)") & " AS cLlave2, "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & Choose(gsIdioma, "COCprDocCta.GloDetx, ", "COCprDocCta.GloDet, ")
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & "COCprDocCta.UsrCre, COCprDocCta.FyHCre, COCprDocCta.UsrMdf, COCprDocCta.FyHMdf, "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & "COCprDocCta.codemp, COCprDocCta.pdoano "
  usConnStrgSele_COCprDocCta = usConnStrgSele_COCprDocCta & "FROM COCprDocCta "
  usConnStrgWher_COCprDocCta = ""
  usConnStrgOrde_COCprDocCta = "ORDER BY 10, 11, 1" ' DESC"
  
  usConnStrgSele_COCprDocCCo = "SELECT COCprDocCCo.CodCCo, COCprDocCCo.ImpCCo_MN, COCprDocCCo.ImpCCo_ME, "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & "COCprDocCCo.TpoCnc, COCprDocCCo.CodCta, COCprDocCCo.Orden, "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & "COCprDocCCo.CodAux, COCprDocCCo.CodTDc, COCprDocCCo.SerDoc, COCprDocCCo.NroDoc, "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & IIf(ps_Plataforma = pSrvMySql, "Concat(COCprDocCCo.TpoCnc, COCprDocCCo.Orden, COCprDocCCo.CodCta)", "(RTrim(COCprDocCCo.TpoCnc)+COCprDocCCo.Orden+COCprDocCCo.CodCta)") & " AS cLlave, "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & IIf(ps_Plataforma = pSrvMySql, "Concat(COCprDocCCo.CodAux, COCprDocCCo.CodTDc, COCprDocCCo.SerDoc, COCprDocCCo.NroDoc, COCprDocCCo.TpoCnc, COCprDocCCo.Orden, COCprDocCCo.CodCta)", "(COCprDocCCo.CodAux+COCprDocCCo.CodTDc+COCprDocCCo.SerDoc+COCprDocCCo.NroDoc+RTrim(COCprDocCCo.TpoCnc)+COCprDocCCo.Orden+COCprDocCCo.CodCta)") & " AS cLlave1, "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & IIf(ps_Plataforma = pSrvMySql, "Concat(COCprDocCCo.CodAux, COCprDocCCo.CodTDc, COCprDocCCo.SerDoc, COCprDocCCo.NroDoc, COCprDocCCo.TpoCnc, COCprDocCCo.Orden, COCprDocCCo.CodCta, COCprDocCCo.CodCCo)", "(COCprDocCCo.CodAux+COCprDocCCo.CodTDc+COCprDocCCo.SerDoc+COCprDocCCo.NroDoc+RTrim(COCprDocCCo.TpoCnc)+COCprDocCCo.Orden+COCprDocCCo.CodCta+COCprDocCCo.CodCCo)") & " AS cLlave2, "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & "COCprDocCCo.UsrCre, COCprDocCCo.FyHCre, COCprDocCCo.UsrMdf, COCprDocCCo.FyHMdf, "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & "COCprDocCCo.codemp, COCprDocCCo.pdoano "
  usConnStrgSele_COCprDocCCo = usConnStrgSele_COCprDocCCo & "FROM COCprDocCCo "
  usConnStrgWher_COCprDocCCo = ""
  usConnStrgOrde_COCprDocCCo = "ORDER BY 4, 6, 5, 1"
  
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
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & "COCpbDet.TpoGnr, COCpbDet.pdocpr, COCpbDet.codcon, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & IIf(ps_Plataforma = pSrvMySql, "Concat(COCpbDet.CodDro, COCpbDet.NroCpb, COCpbDet.NroIte)", "(COCpbDet.CodDro+COCpbDet.NroCpb+COCpbDet.NroIte)") & " AS cLlave, "
  usConnStrgSele_COCpbDet = usConnStrgSele_COCpbDet & Choose(gsIdioma, " COCpbDet.GloItex, ", " COCpbDet.GloIte, ")
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
  Set uorstTGAux = New ADODB.Recordset
  Set uorstTGTDc = New ADODB.Recordset
  Set uorstTGTCb = New ADODB.Recordset
  Set uorstCoCta = New ADODB.Recordset
  Set uorstCoCCo = New ADODB.Recordset
  Set uorstCoCCox = New ADODB.Recordset
  Set uorstCoCCoy = New ADODB.Recordset
  Set uorstCODro = New ADODB.Recordset
  Set uorstCoAsiTipo = New ADODB.Recordset
  Set uorstCOCprDocCta = New ADODB.Recordset
  Set uorstCOCprDocCCo = New ADODB.Recordset
  Set uorstCOCpbCab = New ADODB.Recordset
  Set uorstCOCpbDet = New ADODB.Recordset
  Set porstCancel = New ADODB.Recordset
  
  Set uorstcodetrac = New ADODB.Recordset '2015-07-02 adic tabla detrac
  
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
     .Properties("Unique Table").Value = "COCprDoc"
  End With
  With uorstMain
     .ActiveConnection = uocnnMain
     .Source = psConnStrgSele & psConnStrgOrde
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic 'adLockReadOnly
     .Open
     .Properties("Unique Table").Value = "COCprDoc"
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
     .ActiveConnection = frmTCprGrd.uocnnMain
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
  With uorstCoCCox
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.CodCCo, " & Choose(gsIdioma, "a.DetCCo", "a.DetCCox") & " AS DetCCo "
     .Source = .Source & "FROM COCCo a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND a.EstCCo='" & ESTCCO_ACT & "' AND a.indpdocpr=1 "
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodCCo)>2"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  With uorstCoCCoy
     .ActiveConnection = uocnnMain
     .Source = "SELECT a.CodCCo, " & Choose(gsIdioma, "a.DetCCo", "a.DetCCox") & " AS DetCCo "
     .Source = .Source & "FROM COCCo a "
     .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND a.EstCCo='" & ESTCCO_ACT & "' AND a.indpdocpr=0 "
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(a.CodCCo)>2"
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockReadOnly
     .Open
  End With
  With uorstCODro
     .ActiveConnection = uocnnMain
     .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro, Cpb" & gsMesAct & ", "
     .Source = .Source & "codemp, pdoano "
     .Source = .Source & "FROM CODro "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
     .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
  With uorstCoAsiTipo
    .ActiveConnection = uocnnMain
    .Source = "SELECT a.CodAsi, " & Choose(gsIdioma, "a.DetAsi", "a.DetAsix") & " AS DetAsi, a.TpoAsi "
    .Source = .Source & "FROM CoAsiTipo a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND a.TpoAsi='" & TPOGNR_CPR & "'"
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
  With uorstCOCprDocCta
     .ActiveConnection = uocnnMain
  '     .CursorLocation = adUseClient   'Es el Default.
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
  End With
  With uorstCOCprDocCCo
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
'ini 2015-07-02 adic tabla detrac
  With uorstcodetrac
     .ActiveConnection = uocnnMain
     .Source = "SELECT coddetrac, " & Choose(gsIdioma, "detdetrac", "detdetracx") & " AS DetDetrac,pctdetrac ,  "
     .Source = .Source & "codemp "
     .Source = .Source & "FROM codetrac  "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND estdetrac ='" & ESTDETRAC_ACT & "' "
     '.Source = .Source & "AND pdoano='" & gsAnoAct & "' "
     '.Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
     .CursorType = adOpenDynamic
     .LockType = adLockOptimistic
     .Open
  End With
  
'fin 2015-07-02 adic tabla detrac
  
  
  ']
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  CaptionBotones Me, False, False, True, True, True, True, False, False, False, False, False, False, True, aLabel
  cmdVerificar.Caption = Choose(gsIdioma, "&Verificar", "&Check")
  cmdGenera.Caption = Choose(gsIdioma, "&Generar", "&Generate")
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
   uorstCoCta.Close
   uorstCoCCo.Close
   uorstCODro.Close
   uorstCoAsiTipo.Close
   
   uorstcodetrac.Close '2015-07-02 adic tabla detrac
   
'[ARREGLAR. Genera demora al salir de la opción.
   If uorstCOCprDocCta.State = adStateOpen Then uorstCOCprDocCta.Close
   If uorstCOCprDocCCo.State = adStateOpen Then uorstCOCprDocCCo.Close
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
   Set uorstCoCta = Nothing
   Set uorstCoCCo = Nothing
   Set uorstCODro = Nothing
   Set uorstCoAsiTipo = Nothing
   Set uorstCOCprDocCta = Nothing
   Set uorstCOCprDocCCo = Nothing
   Set uorstCOCpbCab = Nothing
   Set uorstCOCpbDet = Nothing
   Set uorstMain_Grd = Nothing
   Set uorstMain = Nothing
   
    Set uorstcodetrac = Nothing '2015-07-02 adic tabla detrac
  
   Set uocnnMain = Nothing
   
   
End Sub

Private Sub cmdNuevo_Click()
 '[Propio del formulario.
   'Verificación de Mes Cerrado.
   If gbCieCpr Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   ubGrabaMas = INDMASCTA_INI
   uocnnMain.BeginTrans
 ']
   gpTUg_Nuevo Me, frmTCpr             'Cambiar Formulario de Datos.
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
   uorstMain.Find "cLlave='" & uorstMain_Grd!codaux & uorstMain_Grd!codtdc & uorstMain_Grd!serdoc & uorstMain_Grd!nrodoc & "'"
 ']

   With frmTCpr                        'Cambiar Formulario de Datos.
      .zbNuevo = False
      .upDatosDesconectados 1
    '[Deshabilitación de Llaves.       'Cambiar.
      .txtLlave(0).Enabled = False
      .txtLlave(1).Enabled = False
      .txtLlave(2).Enabled = False
      .txtLlave(3).Enabled = False
      .cmdLlaveAyud(0).Enabled = False
      .cmdLlaveAyud(1).Enabled = False
      .lblLlaveDeta(0).Enabled = False
      .lblLlaveDeta(1).Enabled = False
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
   If gbCieCpr Then
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
        .Source = "SELECT MesPvs, CodAux, CodTDc, SerDoc, NroDoc, TpoPvs "
        .Source = .Source & "FROM COCpbDet "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND MesPvs='" & gsMesAct & "' AND CodAux='" & uorstMain_Grd!codaux & "' "
        .Source = .Source & "AND CodTDc='" & uorstMain_Grd!codtdc & "' AND SerDoc='" & uorstMain_Grd!serdoc & "'"
        .Source = .Source & "AND NroDoc='" & uorstMain_Grd!nrodoc & "' AND TpoPvs<>'" & TPOPVS_CAN & "'"
        .Open
         If porstCancel.RecordCount = 0 Then
            uorstMain.MoveFirst
            uorstMain.Find "cLlave = '" & uorstMain_Grd!codaux & uorstMain_Grd!codtdc & uorstMain_Grd!serdoc & uorstMain_Grd!nrodoc & "'"

            uocnnMain.BeginTrans       'INICIA TRANSACCION.
            uocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND MesPvs='" & gsMesAct & "' And CodDro='" & Trim(dgrMain.Columns(0)) & "' And NroCpb='" & Trim(dgrMain.Columns(1)) & "' And TpoGnr='" & TPOGNR_CPR & "'"
            uorstMain.Properties("Unique Table").Value = "COCprDoc"
            uorstMain.Delete
            uocnnMain.CommitTrans      'CONFIRMA TRANSACCION.

           'Busca siguiente ítem.
            With uorstMain_Grd
               .MoveNext
               If .EOF Then .MoveLast
               dsLlaveSiguiente = !codaux & !codtdc & !serdoc & !nrodoc
               .Requery
               If .RecordCount > 0 Then .Find "cLlave = '" & dsLlaveSiguiente & "'"
            End With
            
            'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
            fEstMayUpd
            'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
            
         Else
            MsgBox Choose(gsIdioma, "Debe eliminar antes las Cancelaciones.", " The Cancelations must be eliminated before."), vbExclamation
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
Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub dgrMain_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo Err
   
'[ARREGLAR. No acepta ordenar por columna de tablas secundarias en el recordset.
   If ColIndex = 3 Or ColIndex = 4 Then Exit Sub
']ARREGLAR.

   pnColumnaOrd = ColIndex
   fraBuscar.Caption = TEXT_BUSCA & dgrMain.Columns(pnColumnaOrd).Caption
   txtBuscar = ""

   psConnStrgOrde = "ORDER BY "
   Select Case pnColumnaOrd            'Cambiar.
'   Case 1
'      psConnStrgOrde = psConnStrgOrde & "2, 3, 4, 5"
   Case Else
      psConnStrgOrde = psConnStrgOrde & pnColumnaOrd + 1
   End Select
   With uorstMain_Grd
      .Close
      .Properties("Unique Table").Value = "COCprDoc"
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
            .Item(dnNum).Caption = Choose(gsIdioma, "NºComp.", "NºVouchers")
            .Item(dnNum).Width = 700
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 1100
         Case 3
            .Item(dnNum).Caption = Choose(gsIdioma, "Razón Social", "Firm Name")
            .Item(dnNum).Width = 1650
         Case 4
            .Item(dnNum).Caption = Choose(gsIdioma, "TDc", "TDc") ' Type of Document
            .Item(dnNum).Width = 500
         Case 5
            .Item(dnNum).Caption = Choose(gsIdioma, "Ser", "Ser")
            .Item(dnNum).Width = 500
         Case 6
            .Item(dnNum).Caption = Choose(gsIdioma, "Número", "Number")
            .Item(dnNum).Width = 1000
         Case 7
            .Item(dnNum).Caption = Choose(gsIdioma, "F.Emisión", "Issue Date")
            .Item(dnNum).Width = 1000
         Case 8
            .Item(dnNum).Caption = Choose(gsIdioma, "M", "C")     '  Currency
            .Item(dnNum).Width = 250
         Case 9
            .Item(dnNum).Caption = Choose(gsIdioma, "Total", "Total")
            .Item(dnNum).Width = 1200
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
            .Item(dnNum).Alignment = dbgRight
         Case 10
            .Item(dnNum).Caption = Choose(gsIdioma, "G", "G")
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
'  On Error GoTo ErrGrabar
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
  sSentencia = sSentencia & "'" & TPOGNR_CPR & "', "
  sSentencia = sSentencia & "'" & INDNCU_FAL & "', "
  sSentencia = sSentencia & "'" & INDANU_FAL & "', "
  sSentencia = sSentencia & "'" & gsAbvUsr & "', "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
  sSentencia = sSentencia & "Null, Null)"
  uocnnMain.Execute sSentencia, nNumRegistros
  
  ' Información detalle cuentas
  With porstCprCta
    .Source = "SELECT cpr.tpocnc, cpr.orden, cpr.codcta, cco.codcco, cpr.glodet, cpr.glodetx, cpr.impcta_mn, cpr.impcta_me, cco.impcco_mn, cco.impcco_me, cpr.codruc, "
    .Source = .Source & "cta.indcco, cta.inddoc, cta.inddoc, cta.tpotcb, tdc.sgntdc "
    .Source = .Source & "FROM cocprdoccta cpr "
    .Source = .Source & "INNER JOIN cocta cta ON cpr.codemp=cta.codemp AND cpr.pdoano=cta.pdoano AND cpr.codcta=cta.codcta "
    .Source = .Source & "INNER JOIN tgtdc tdc ON cpr.codemp=tdc.codemp AND cpr.codtdc=tdc.codtdc "
    .Source = .Source & "LEFT JOIN cocprdoccco cco ON cpr.codemp=cco.codemp AND cpr.pdoano=cco.pdoano AND cpr.codaux=cco.codaux AND cpr.codtdc=cco.codtdc "
    .Source = .Source & "AND cpr.serdoc=cco.serdoc AND cpr.nrodoc=cco.nrodoc AND cpr.tpocnc=cco.tpocnc AND cpr.orden=cco.orden AND cpr.codcta=cco.codcta "
    .Source = .Source & "WHERE cpr.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND cpr.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND cpr.codaux='" & oRecordset!codaux & "' "
    .Source = .Source & "AND cpr.codtdc='" & oRecordset!codtdc & "' "
    .Source = .Source & "AND cpr.serdoc='" & oRecordset!serdoc & "' "
    .Source = .Source & "AND cpr.nrodoc='" & oRecordset!nrodoc & "' "
    .Source = .Source & "ORDER BY cpr.tpocnc, cpr.orden"
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
        sTpoCtb = IIf(porstCprCta!tpocnc = TPOCNC_TOT_CPR, IIf(porstCprCta!SgnTDc = SGNTDC_POS, TPOCTB_HAB, TPOCTB_DEB), IIf(porstCprCta!SgnTDc = SGNTDC_NEG, TPOCTB_HAB, TPOCTB_DEB))
      Else
        sTpoCtb = IIf(porstCprCta!tpocnc = TPOCNC_TOT_CPR, IIf(porstCprCta!SgnTDc = SGNTDC_NEG, TPOCTB_HAB, TPOCTB_DEB), IIf(porstCprCta!SgnTDc = SGNTDC_POS, TPOCTB_HAB, TPOCTB_DEB))
      End If
      nRegistro = nRegistro + 1
      ' Grabación de cabecera de comprobante
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
      sSentencia = sSentencia & "'" & oRecordset!codtdc & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!fehope, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & "'" & porstCprCta!CodCta & "', "
      sSentencia = sSentencia & IIf(IsNull(porstCprCta!codcco), "Null", "'" & porstCprCta!codcco & "'") & ", "
      sSentencia = sSentencia & IIf(sCodAux = "", "Null", "'" & sCodAux & "'") & ", "
      sSentencia = sSentencia & "'" & oRecordset!serdoc & "', "
      sSentencia = sSentencia & "'" & oRecordset!nrodoc & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!fevdoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(oRecordset!ferdoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
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
      sSentencia = sSentencia & "'" & TPOGNR_CPR & "', "
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
  sSentencia = "UPDATE CoCprDoc SET indpregen=" & INDPREGEN_ACT & ", indgen=-1 "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND codaux='" & oRecordset!codaux & "' "
  sSentencia = sSentencia & "AND codtdc='" & oRecordset!codtdc & "' "
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
  For nContador = 1 To 8
    sRegistro = Choose(nContador, "impogr", "impogn", "impong", "impexo", "impigv", "impisc", "impoim", "imptot")
    sIndicado = "indcta_" & Right(sRegistro, 3)
    nImporteCpr_mn = CDec(oRecordset(sRegistro & "_mn"))
    nImporteCpr_me = CDec(oRecordset(sRegistro & "_me"))
    nImporteCta_mn = 0
    nImporteCta_me = 0
    
    If nContador = 8 Then nContador = 11
    
    ' Verifico los importes de las cuentas
    If oRecordset(sIndicado) <> 0 Then
      
      With porstCprCta
        .Source = "SELECT cpr.orden, cpr.codcta, cpr.impcta_mn, cpr.impcta_me, cta.indcco "
        .Source = .Source & "FROM cocprdoccta cpr "
        .Source = .Source & "INNER JOIN cocta cta ON cpr.codemp=cta.codemp AND cpr.pdoano=cta.pdoano AND cpr.codcta=cta.codcta "
        .Source = .Source & "WHERE cpr.codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND cpr.pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND cpr.codaux='" & oRecordset!codaux & "' "
        .Source = .Source & "AND cpr.codtdc='" & oRecordset!codtdc & "' "
        .Source = .Source & "AND cpr.serdoc='" & oRecordset!serdoc & "' "
        .Source = .Source & "AND cpr.nrodoc='" & oRecordset!nrodoc & "' "
        .Source = .Source & "AND cpr.tpocnc='" & nContador & "' "
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
              .Source = "SELECT cpr.codcta, ROUND(SUM(cpr.impcco_mn), 2) AS impcco_mn, ROUND(SUM(cpr.impcco_me), 2) AS impcco_me "
              .Source = .Source & "FROM cocprdoccco cpr "
              .Source = .Source & "INNER JOIN cocco cco ON cpr.codemp=cco.codemp AND cpr.pdoano=cco.pdoano AND cpr.codcco=cco.codcco "
              .Source = .Source & "WHERE cpr.codemp='" & gsCodEmp & "' "
              .Source = .Source & "AND cpr.pdoano='" & gsAnoAct & "' "
              .Source = .Source & "AND cpr.codaux='" & oRecordset!codaux & "' "
              .Source = .Source & "AND cpr.codtdc='" & oRecordset!codtdc & "' "
              .Source = .Source & "AND cpr.serdoc='" & oRecordset!serdoc & "' "
              .Source = .Source & "AND cpr.nrodoc='" & oRecordset!nrodoc & "' "
              .Source = .Source & "AND cpr.tpocnc='" & nContador & "' "
              .Source = .Source & "AND cpr.orden='" & porstCprCta!orden & "' "
              .Source = .Source & "AND cpr.codcta='" & porstCprCta!CodCta & "' "
              .Source = .Source & "GROUP BY cpr.codcta "
              .Open
            End With
            ' Valido los centro de costos
            If porstCprCco.RecordCount > 0 Then
              nImporteCCo_mn = CDec(porstCprCco!impcco_mn)
              nImporteCCo_me = CDec(porstCprCco!impcco_me)
            End If
            porstCprCco.Close
            ' Verifico los importes de centro de costo
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
    
    'MsgBox Str(nContador) & " " & Str(nImporteCpr_mn) & " " & Str(nImporteCta_mn)
    
    
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

