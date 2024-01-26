VERSION 5.00
Begin VB.Form frmRRegCprExpo 
   Caption         =   "Exportar Datos"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
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
      Height          =   495
      Left            =   1920
      Picture         =   "frmRRegCprExpo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdExporta 
      Caption         =   "&Excel"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "frmRRegCprExpo.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdProcExcel 
      Caption         =   "Procesar &Excel"
      Height          =   400
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame frmUbicacion 
      Caption         =   " Carpeta "
      Height          =   2355
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.DirListBox dlbDirectorio 
         Height          =   1440
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   690
         Width           =   2235
      End
      Begin VB.DriveListBox drvUnidad 
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   400
         Width           =   2235
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Directorio :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   200
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmRRegCprExpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public pocnnMain As ADODB.Connection


Private Sub cmdExporta_Click(Index As Integer)
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
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CBASIMPGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGr_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CIGVGRA,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CBASIMPGNG,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_OGN_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CIGVGRANGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CBASIMPSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpIGV_ONG_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CIGVSCF,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpExo_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CIMPTOTNGV,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpISC_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CISC,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpOIm_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS COTRTRICGO,"
        cCadReporte = cCadReporte & "    replace(format((a.ImpTot_MN * IF(b.SgnTDc = 0, -1,1)),2),',','') AS CIMPTOTCOM,"
        cCadReporte = cCadReporte & "    format(a.imptcb,3) AS CTIPCAM,"
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
        cCadReporte = cCadReporte & "    '' AS CINTREG "
        
        
        cCadReporte = cCadReporte & "FROM (((COCprDoc a "
        cCadReporte = cCadReporte & "LEFT JOIN TGTDc b on a.codemp=b.codemp and b.CodTDc = a.CodTDc) "
        cCadReporte = cCadReporte & "LEFT JOIN TGAux c on a.codemp=c.codemp and c.CodAux = a.CodAux) "
        cCadReporte = cCadReporte & "LEFT JOIN CODro d ON a.codemp=d.codemp and a.pdoano=d.pdoano and a.CodDro=d.CodDro) "
        'cCadReporte = cCadReporte & "WHERE a.codemp='001' and a.pdoano='2012' and a.MesPvs >= '01' AND  a.MesPvs <= '04' "
        cCadReporte = cCadReporte & "WHERE "
        cCadReporte = cCadReporte & "    a.codemp='" & gsCodEmp & "' AND "
        'cCadReporte = cCadReporte & "    a.pdoano='" & sPdoAnoFin & "' AND "
        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) >= '" & gsAnoAct & gsMesAct & "' AND "
        '2015-03-20 cambio de periodo cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & "201502" & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        cCadReporte = cCadReporte & "    concat(a.pdoano,a.MesPvs) <= '" & gsAnoAct & gsMesAct & "' AND YEAR(a.FeEDoc)='" & gsAnoAct & "' "
        cCadReporte = cCadReporte & "ORDER BY mespvs ,a.CodDro, a.NroCpb ASC "

    
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
        Dim nRowI As Long, nColI As Long
        Dim nRecord As Long, nFields As Long
        Dim xrow1 As Long
        nRowI = 1: nColI = 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        
        'crear tabla temporal
        'Dim xpocnnMain As ADODB.Connection
        'Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")

'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next

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

Private Sub cmdProcExcel_Click()


'    Dim xArchPeriodo As String
'    xArchPeriodo = "plan 2011 txtpg.xls"
'
'
'    Dim oExcel As Excel.Application
'    Dim oWBook As Excel.Workbook
'    Dim oSheet As Excel.Worksheet
'
'    Set oExcel = New Excel.Application
'    'Set oWBook = oExcel.Workbooks.Add
'    'Set oSheet = oWBook.Worksheets(1)
'    Set oWBook = oExcel.Workbooks.Open(dlbDirectorio(0).path & "\" & xArchPeriodo, , True) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
'    Set oSheet = oWBook.Worksheets("Clientes")
'    oExcel.Visible = True
'    'oWBook.Worksheets("Clientes").Select
'    With oSheet
'        oSheet.Select
'        Dim nRowI As Long, nColI As Long
'        Dim nRecord As Long, nFields As Long
'        Dim xrow1 As Long
'        nRowI = 1: nColI = 1
'        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
'        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
'        nRowI = nRowI + 1 'limite inicial real
'        nRecord = (nRowI + nRecord)
'        If nRecord = 0 Then nRecord = nRowI
'        'crear tabla temporal
'        'Dim xpocnnMain As ADODB.Connection
'        Set pocnnMain = fOpenTmp(pocnnMain, "ex2aux")
'
'        For xrow1 = nRowI To nRecord
'            MsgBox (.Cells(xrow1, 1).Value)
'        Next
'
'    End With
'    'oExcel.Visible = True
'    oExcel.Quit
'    Set oExcel = Nothing

End Sub

Private Sub cmdSalir_Click()
        Unload Me

End Sub

Private Sub dlbDirectorio_Change(Index As Integer)
'  flbArchivo.path = dlbDirectorio(0).path
'  flbArchivo.Refresh
End Sub

'Private Sub drvUnidad_Change(Index As Integer)
'  dlbDirectorio(Index).path = drvUnidad(Index).Drive
'  dlbDirectorio(Index).Refresh
'End Sub


Private Sub Form_Load()
' On Error GoTo Err

'  Dim s_Conexion As String, sSentencia As String
'  s_Conexion = CONNSTRG & gsNomBDS
'  Set pocnnMain = New ADODB.Connection
'  With pocnnMain
'    If .State = adStateOpen Then .Close
'    .ConnectionTimeout = 15
'    .CursorLocation = adUseClient
'    .ConnectionString = s_Conexion
'    .Open
'  End With
'  pocnnMain.BeginTrans               'INICIA TRANSACCION.
'
'  ppImporta_Tablas
'  ppTransfir_Tablas
'
'  pocnnMain.CommitTrans                           ' CONFIRMA TRANSACCION.
  
'  cmdAceptar.Enabled = True
'  cmdSalir.Enabled = True
'  cmdSalir.SetFocus
'  pocnnMain.Close
'  Set pocnnMain = Nothing
'  Exit Sub
'Err:
'  If pocnnMain.State = adStateOpen Then
''    pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
'    pocnnMain.Close
'    Set pocnnMain = Nothing
'  End If
'
''  cmdSalir.Enabled = True
''  cmdSalir.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
'   uorstMain.Close
'   uocnnMain.Close
'   Set uorstMain = Nothing
'   Set uocnnMain = Nothing
End Sub
