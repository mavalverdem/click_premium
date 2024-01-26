VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPPDTRet 
   Caption         =   "[título]"
   ClientHeight    =   4170
   ClientLeft      =   2250
   ClientTop       =   2385
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5745
   Begin TabDlg.SSTab tabProceso 
      Height          =   2835
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5001
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   494
      TabCaption(0)   =   "Configuración de Parámetros"
      TabPicture(0)   =   "frmPPDTRet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmRegistro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmUbicacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboTpoMon"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmTipoEmpresa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame frmTipoEmpresa 
         Caption         =   " Tipo de Empresa "
         ForeColor       =   &H00000080&
         Height          =   900
         Left            =   150
         TabIndex        =   14
         Top             =   350
         Width           =   2500
         Begin VB.OptionButton optTipoEmpresa 
            Caption         =   "No Agente"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   16
            Top             =   510
            Width           =   2200
         End
         Begin VB.OptionButton optTipoEmpresa 
            Caption         =   "Agente"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   250
            Value           =   -1  'True
            Width           =   2200
         End
      End
      Begin VB.ComboBox cboTpoMon 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2355
         Width           =   1260
      End
      Begin VB.Frame frmUbicacion 
         Caption         =   " Carpeta "
         ForeColor       =   &H00000080&
         Height          =   2325
         Left            =   2880
         TabIndex        =   3
         Top             =   350
         Width           =   2535
         Begin VB.DriveListBox drvUnidad 
            Height          =   315
            Left            =   150
            TabIndex        =   11
            Top             =   400
            Width           =   2235
         End
         Begin VB.DirListBox dlbDirectorio 
            Height          =   1440
            Left            =   150
            TabIndex        =   4
            Top             =   690
            Width           =   2235
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   200
            Width           =   765
         End
      End
      Begin VB.Frame frmRegistro 
         Caption         =   " Registros "
         ForeColor       =   &H00000080&
         Height          =   900
         Left            =   150
         TabIndex        =   0
         Top             =   1305
         Width           =   2500
         Begin VB.CheckBox chkInformacion 
            Caption         =   "&Percepciones"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   2
            Top             =   510
            Value           =   1  'Checked
            Width           =   2200
         End
         Begin VB.CheckBox chkInformacion 
            Caption         =   "&Retenciones"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   250
            Value           =   1  'Checked
            Width           =   2200
         End
      End
      Begin VB.Label lblTexto 
         Caption         =   "Moneda"
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   1
         Left            =   555
         TabIndex        =   13
         Top             =   2415
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   400
      Left            =   1380
      TabIndex        =   9
      Top             =   3675
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   400
      Left            =   3060
      TabIndex        =   8
      Top             =   3675
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProgreso 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      Caption         =   "Procesando Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1785
   End
End
Attribute VB_Name = "frmPPDTRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    On Error GoTo Err
    
    If MsgBox("¿ Estás Seguro de Generar archivo de información ? ", vbQuestion + vbYesNo) = vbYes Then
      cmdAceptar.Enabled = False
      cmdSalir.Enabled = False
      pgbProgreso(0).Value = 0: pgbProgreso(0).Min = 0
      ' Genero los archivos de información
      ppGenArchivo
      
      MsgBox TEXT_8008, vbInformation
      cmdAceptar.Enabled = True
      cmdSalir.Enabled = True
      cmdSalir.SetFocus
    End If
    Exit Sub
Err:
    MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
    cmdSalir.Enabled = True
    cmdSalir.SetFocus

End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Function pfOutApoRet(s_Expresion As String) As String

s_Expresion = Trim$(s_Expresion)
If s_Expresion <> "" Then
    ' saco los enters de la cadena de caracteres
    While InStr(s_Expresion, Chr(13)) <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(13)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(13)) + 1)
    Wend
    ' saco los retornos de la cadena de caracteres
    While InStr(s_Expresion, Chr(10)) <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, Chr(10)) - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, Chr(10)) + 1)
    Wend
    ' saco los apostrofes de la cadena de caracteres
    While InStr(s_Expresion, "'") <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "'") - 1) & "´" & Mid$(s_Expresion, InStr(s_Expresion, "'") + 1)
    Wend
    ' saco los rayas de la cadena de caracteres
    While InStr(s_Expresion, "|") <> 0
        s_Expresion = Left$(s_Expresion, InStr(s_Expresion, "|") - 1) & " " & Mid$(s_Expresion, InStr(s_Expresion, "|") + 1)
    Wend
End If
pfOutApoRet = Trim$(s_Expresion)

End Function

Private Sub ppGenArchivo()
  
  Dim sSentencia As String, sLinea As String
  Dim nContador As Integer, nArchivo As Integer
  Dim nRegistro As Double, nNumRegistros As Double, nImporte As Double
  Dim sArchivo As String, nSecuencia As Integer
  Dim sCaracter  As String

  Dim pocnnMain As ADODB.Connection
  Dim porstTmp As ADODB.Recordset

  ' Seteo y activo la coneccion
  Set pocnnMain = New ADODB.Connection
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  
  ' Seteo el recordset temporal
  Set porstTmp = New ADODB.Recordset
  sCaracter = "|"
  ' Importo las tablas de acuerdo a la selección
  For nContador = 0 To chkInformacion.Count - 1
    ' Verifico que se haya seleccionado
    If chkInformacion(nContador).Value = vbChecked Then
      ' Obtengo el archivo de texto libre
      nArchivo = FreeFile
      ' Tabla temporal de comprobantes de retencion
      sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 15)='#tmpRetenPerce_') DROP TABLE #tmpRetenPerce"
      pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRetenPerce", sSentencia)
      
      sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpRetenPerce ", "")
      sSentencia = sSentencia & "SELECT DISTINCT a.*, b.FeEDoc, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & Choose(nContador + 1, TPOCTB_HAB, TPOCTB_DEB) & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0) & "_RtcPcp ELSE 0 END), 0) AS cImpRetDeb, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & Choose(nContador + 1, TPOCTB_DEB, TPOCTB_HAB) & "' THEN a.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0) & "_RtcPcp ELSE 0 END), 0) AS cImpRetHab, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN b.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0) & " ELSE 0 END), 0) AS cImpCanDeb, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN b.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0) & " ELSE 0 END), 0) AS cImpCanHab "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpRetenPerce ")
      sSentencia = sSentencia & "FROM (CoCPbDetRP a "
      sSentencia = sSentencia & "LEFT JOIN CoCPbDet b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.MesPvs=b.MesPvs AND a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb AND a.NroIte=b.NroIte) "
      sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND a.MesPvs='" & gsMesAct & "' "
      sSentencia = sSentencia & "AND a.CodTDc_RtcPcp='" & Choose(nContador + 1, gsCodTDc_Rtc, gsCodTDc_Pcp) & "' "
      pocnnMain.Execute sSentencia, nNumRegistros
      
      ' Tabla temporal de total de comprobantes de retencion
      pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpImporterp", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#tmpImporterp_') DROP TABLE " & ps_Prefijo & "tmpImporterp")
      sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpImporterp ", "")
      sSentencia = sSentencia & "SELECT DISTINCT MesPvs, CodAux, CodTDc_RtcPcp, SerDoc_RtcPcp, NroDoc_RtcPcp, MIN(FeEDoc_RtcPcp) AS FeEDoc_RtcPcp, "
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cImpRetDeb, 0)), 2) AS cTotRetDeb, "
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cImpRetHab, 0)), 2) AS cTotRetHab, "
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cImpCanDeb, 0)), 2) AS cTotCanDeb, "
      sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cImpCanHab, 0)), 2) AS cTotCanHab "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpImporterp ")
      sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpRetenPerce "
      sSentencia = sSentencia & "GROUP BY MesPvs, CodAux, CodTDc_RtcPcp, SerDoc_RtcPcp, NroDoc_RtcPcp "
      sSentencia = sSentencia & "ORDER BY MesPvs, CodAux, CodTDc_RtcPcp, SerDoc_RtcPcp, NroDoc_RtcPcp"
      pocnnMain.Execute sSentencia, nNumRegistros
      
      ' Tabla temporal de comprobantes de retencion con todos datos
      sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 15)='#tmpcocpbdetrp_') DROP TABLE #tmpcocpbdetrp"
      pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpcocpbdetrp", sSentencia)
      sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpcocpbdetrp ", "")
      sSentencia = sSentencia & "SELECT DISTINCT a.*, cTotRetDeb, cTotRetHab, cTotCanDeb, cTotCanHab "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpcocpbdetrp ")
      sSentencia = sSentencia & "FROM (" & ps_Prefijo & "tmpRetenPerce a "
      sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmpImporteRP b ON a.MesPvs=b.MesPvs AND a.CodAux=b.CodAux AND a.CodTDc_RtcPcp=b.CodTDc_RtcPcp AND a.SerDoc_RtcPcp=b.SerDoc_RtcPcp AND a.NroDoc_RtcPcp=b.NroDoc_RtcPcp) "
      sSentencia = sSentencia & "ORDER BY a.MesPvs, a.CodAux, a.CodTDc_RtcPcp, a.SerDoc_RtcPcp, a.NroDoc_RtcPcp"
      pocnnMain.Execute sSentencia, nNumRegistros
      
      ' Generacion de la tabla de seleccion
      sSentencia = "SELECT DISTINCT c.RUCAux, c.TpoPer, c.RazAux, d.ApePatAux, d.ApeMatAux, d.NomAux, "
      sSentencia = sSentencia & "a.SerDoc_RtcPcp, a.NroDoc_RtcPcp, a.FeEDoc_RtcPcp AS cFecOpeRet, "
      sSentencia = sSentencia & "a.FeEDoc AS cFecEmiRet, a.cImpCanDeb, a.cImpCanHab, a.cImpRetDeb, a.cImpRetHab, "
      sSentencia = sSentencia & "a.CodTDc, a.SerDoc, a.NroDoc, b.FehOpe AS cFecOpeDoc, b.FeEDoc AS cFecEmiDoc, "
      sSentencia = sSentencia & "a.cTotCanDeb, a.cTotCanHab, a.cTotRetDeb, a.cTotRetHab, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN b.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0) & " ELSE 0 END), 0) AS cDebeDoc, "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE b.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN b.Imp" & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0) & " ELSE 0 END), 0) AS cHaberDoc "
      sSentencia = sSentencia & "FROM (((" & ps_Prefijo & "tmpCoCPbDetRP a "
      sSentencia = sSentencia & "LEFT JOIN CoCPbDet b ON b.codemp='" & gsCodEmp & "' AND b.pdoano='" & gsAnoAct & "' AND a.CodCta=b.CodCta AND a.CodAux=b.CodAux AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc) "
      sSentencia = sSentencia & "LEFT JOIN TgAux c ON c.codemp='" & gsCodEmp & "' AND a.CodAux=c.CodAux) "
      sSentencia = sSentencia & "LEFT JOIN TgAuxNat d ON c.codemp=d.codemp AND c.CodAux=d.CodAux) "
      sSentencia = sSentencia & "WHERE b.TpoPvs='" & TPOPVS_PVS & "' "
      sSentencia = sSentencia & "ORDER BY c.RUCAux, cFecOpeRet, cFecEmiRet, a.SerDoc_RtcPcp, a.NroDoc_RtcPcp"
      
      ' Abro el recordset temporal
      With porstTmp
        If .State = adStateOpen Then .Close
        .ActiveConnection = pocnnMain
        .Source = sSentencia
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
        .Open
      End With
      If Not (porstTmp.BOF And porstTmp.EOF) Then
        ' Barro todo el recordset y lo grabo en el archivo
        lblProgreso(0).Caption = Choose(gsIdioma, "Exportando Archivo:", "Exporting File:") & Mid(Trim(chkInformacion(nContador).Caption), 2)
        nNumRegistros = porstTmp.RecordCount
        pgbProgreso(0).Max = nNumRegistros
        pgbProgreso(0).Value = pgbProgreso(0).Min
        nRegistro = 0
        ' Nombre del archivo de texto a generar
        sArchivo = dlbDirectorio.path & "\" & IIf(optTipoEmpresa(1).Value, "0621", Choose(nContador + 1, "0626", "0633")) & gsRUCEmp & gsAnoAct & gsMesAct & IIf(optTipoEmpresa(0).Value, "", Choose(nContador + 1, "r", "p")) & ".txt"
        ' Elimino archivo de texto si existe
        If Dir$(sArchivo, vbNormal) <> "" Then Kill sArchivo
        If Dir$(sArchivo, vbNormal) = "" Then
          Open sArchivo For Output Access Write Lock Read Write As #nArchivo
          porstTmp.MoveFirst
          While Not porstTmp.EOF
            nRegistro = nRegistro + 1
            ' Diseño y grabro la linea en el archivo
            sLinea = ""
            sLinea = sLinea & porstTmp!rucaux & sCaracter
            If optTipoEmpresa(0).Value Then
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!TpoPer = TPOPER_JUR, porstTmp!razaux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!TpoPer = TPOPER_NAT, porstTmp!ApePatAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!TpoPer = TPOPER_NAT, porstTmp!ApeMatAux, "")) & sCaracter
                sLinea = sLinea & pfOutApoRet(IIf(porstTmp!TpoPer = TPOPER_NAT, porstTmp!NomAux, "")) & sCaracter
            End If
            sLinea = sLinea & porstTmp!SerDoc_RtcPcp & sCaracter
            sLinea = sLinea & Right(porstTmp!NroDoc_RtcPcp, 8) & sCaracter
            sLinea = sLinea & Format(porstTmp!cFecOpeRet, "dd/mm/yyyy") & sCaracter
            nImporte = Abs(Round(CDec(((porstTmp!cTotCanDeb + IIf(nContador = 1, porstTmp!cTotRetHab, 0)) - (porstTmp!cTotCanHab + IIf(nContador = 1, porstTmp!cTotRetDeb, 0)))), 2))
            '[ Si no es agente de retencion/percepcion
            If optTipoEmpresa(1).Value Then
              nImporte = Abs(Round(CDec((porstTmp!cTotRetDeb - porstTmp!cTotRetHab)), 2))
            End If
            sLinea = sLinea & Format(nImporte, "############0.00") & sCaracter
            sLinea = sLinea & porstTmp!CodTDc & sCaracter
            sLinea = sLinea & porstTmp!SerDoc & sCaracter
            sLinea = sLinea & Right(porstTmp!NroDoc, 8) & sCaracter
            sLinea = sLinea & Format(porstTmp!cFecEmiDoc, "dd/mm/yyyy") & sCaracter
            nImporte = Abs(Round(CDec(porstTmp!cDebeDoc - porstTmp!cHaberDoc), 2))
            sLinea = sLinea & Format(nImporte, "############0.00") & sCaracter
            Print #nArchivo, sLinea
            pgbProgreso(0).Value = nRegistro
            DoEvents
            porstTmp.MoveNext
          Wend
          Close #nArchivo
        End If
      End If
      porstTmp.Close
    End If
  Next nContador
  ' Cierro y saco de memoria los objetos
  Set porstTmp = Nothing
  pocnnMain.Close
  Set pocnnMain = Nothing

End Sub

Private Sub drvUnidad_Change()

dlbDirectorio.path = drvUnidad.Drive
dlbDirectorio.Refresh

End Sub

Private Sub Form_Activate()
   
  drvUnidad.Drive = gsRutSis
  dlbDirectorio.path = gsRutSis
  
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  cmdSalir.SetFocus

End Sub

Private Sub Form_Load()
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Directorio :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Directory :", "Currency :")
  Next nElemento
  tabProceso.TabCaption(0) = Choose(gsIdioma, "Configuración de Parámetros", "Configuration of Parameters")
  frmTipoEmpresa.Caption = Choose(gsIdioma, " Tipo de Empresa ", " Type of Company ")
  optTipoEmpresa(0).Caption = Choose(gsIdioma, "Agente", "Agent")
  optTipoEmpresa(1).Caption = Choose(gsIdioma, "No Agente", "Non Agent")
  frmUbicacion.Caption = Choose(gsIdioma, " Carpeta ", " Location ")
  frmRegistro.Caption = Choose(gsIdioma, " Registros ", " Registers ")
  chkInformacion(0).Caption = Choose(gsIdioma, "&Retenciones", "&Withholding")
  chkInformacion(1).Caption = Choose(gsIdioma, "&Percepciones", "&Perceptions")
  lblProgreso(0).Caption = Choose(gsIdioma, "Procesando Archivo:", "Processing File:")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
End Sub
