VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRFluCjaRea 
   Caption         =   "[título]"
   ClientHeight    =   2865
   ClientLeft      =   3750
   ClientTop       =   2445
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6180
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cmbMoneda 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3600
      TabIndex        =   13
      Text            =   "Moneda"
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cmbPeriodo 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   675
      TabIndex        =   11
      Text            =   "Periodo"
      Top             =   270
      Width           =   2000
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2640
      TabIndex        =   8
      Top             =   1020
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   9
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   990
      Width           =   2175
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   5
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   6180
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2205
      Width           =   6180
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
         Left            =   5040
         Picture         =   "frmRFluCjaRea.frx":0000
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
         Picture         =   "frmRFluCjaRea.frx":014A
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
         Picture         =   "frmRFluCjaRea.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
      Begin MSComctlLib.Toolbar toolbar 
         Height          =   600
         Left            =   3720
         TabIndex        =   16
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
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A1"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "A2"
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
                  Picture         =   "frmRFluCjaRea.frx":077E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRFluCjaRea.frx":08D8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRFluCjaRea.frx":0A32
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRFluCjaRea.frx":0DF4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRFluCjaRea.frx":14BE
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label lblTexto 
      Caption         =   "Periodo"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   30
      TabIndex        =   14
      Top             =   285
      Width           =   600
   End
   Begin VB.Label lblTexto 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   1
      Left            =   2805
      TabIndex        =   12
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "frmRFluCjaRea"
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

Public pocnnMain As ADODB.Connection

Private Sub Form_Activate()
   'Orden: Vista Previa, Imprimir, Exportar.
   zaOpciones = Array(gbPms04, gbPms05, gbPms06)
End Sub

Private Sub Form_Load()

'ini 2015-04-23 exporte excel
toolbar.Buttons(1).ButtonMenus(1).Text = "Del Mes"
toolbar.Buttons(1).ButtonMenus(2).Text = "Al Mes"
'fin 2015-04-23 exporte excel

  Dim dnContador As Integer
'[Parametros
   For dnContador = 0 To 13
    If gsIdioma = NvlUsr_Sup Then
      cmbPeriodo.AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre") & " " & gsAnoAct
    Else
      cmbPeriodo.AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing") & " " & gsAnoAct
    End If
   Next dnContador
   cmbPeriodo.ListIndex = Val(gsMesAct)
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Periodo :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Period :", "Currency :")
  Next nElemento
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
 
   With cmbmoneda
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
   End With
   cmbmoneda.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   
  ' Características de impresión.
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

End Sub

Private Sub Form_Resize()
   On Error Resume Next
  
   picOpciones.Width = Me.Width - 120
   cmdSalir.Left = picOpciones.Width - 1135
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    
  Dim sSentencia As String, sMoneda As String
  Dim porstMRp As ADODB.Recordset

  'On Error GoTo Err
    
  ppHabilitacion False
  
  ' Seteo y activo la coneccion
  Set pocnnMain = New ADODB.Connection
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
  sMoneda = IIf(cmbmoneda.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  ' Genero la tabla temporal saldos anteriores
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpSaldos", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 10)='#tmpSaldos') DROP TABLE #tmpSaldos")
  
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpSaldos ", "")
  sSentencia = sSentencia & "SELECT DISTINCT "
  sSentencia = sSentencia & IIf(OptTipo(0).Value, "'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "'", "(MesPvs +1 )") & " AS MesPvs, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mSalDebe, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mSalHaber "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmpSaldos ", "")
  sSentencia = sSentencia & "FROM ((CoCpbDetFjo a "
  sSentencia = sSentencia & "LEFT JOIN CoFjo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodFjo=b.CodFjo) "
  sSentencia = sSentencia & "LEFT JOIN CoCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs<'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "' AND c.IndFjo='" & INDFJO_ACT & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodFjo, '')<>'' "
  sSentencia = sSentencia & IIf(OptTipo(1).Value, "GROUP BY MesPvs ORDER BY MesPvs", "")
  pocnnMain.Execute sSentencia
  
  ' Creo la sentencia de seleccion del reporte
  sSentencia = "SELECT DISTINCT a.MesPvs, b.TpoFjo, (CASE b.TpoFjo WHEN '" & TPOFJO_ING & "' THEN '" & UCase(TPOFJO_ING_TXT) & "' ELSE '" & UCase(TPOFJO_EGR_TXT) & "' END) AS mDesTitulo, "
  sSentencia = sSentencia & "LEFT(a.CodFjo, 2) As mCodNivel, c.DetFjo AS mDetNivel, a.CodFjo, " & Choose(gsIdioma, "b.DetFjo", "b.DetFjox") & " AS DetFjo, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mDebe, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2) AS mHaber, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AVG(e.mSalDebe), 0) AS mSalDebe, "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AVG(e.mSalHaber), 0) AS mSalHaber "
  sSentencia = sSentencia & "FROM ((((CoCpbDetFjo a "
  sSentencia = sSentencia & "LEFT JOIN CoFjo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodFjo=b.CodFjo) "
  sSentencia = sSentencia & "LEFT JOIN CoFjo c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND LEFT(a.CodFjo, 2)=c.CodFjo) "
  sSentencia = sSentencia & "LEFT JOIN CoCta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCta=d.CodCta) "
  sSentencia = sSentencia & "LEFT JOIN " & ps_Prefijo & "tmpSaldos e ON a.MesPvs=e.MesPvs) "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND a.MesPvs" & IIf(OptTipo(0).Value, "=", "<=") & "'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodFjo, '')<>'' "
  sSentencia = sSentencia & "AND d.IndFjo='" & INDFJO_ACT & "' "
  sSentencia = sSentencia & "GROUP BY a.MesPvs, b.TpoFjo, a.CodFjo, c.DetFjo, b.DetFjo "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (mDebe<>0.00 OR mHaber<>0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ISNULL(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2)<>0.00 "
    sSentencia = sSentencia & "OR ROUND(ISNULL(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END), 0), 2)<>0.00) "
  End If
  sSentencia = sSentencia & "ORDER BY a.MesPvs, b.TpoFjo, a.CodFjo"
  
  If OptTipo(1).Value Then
    ' Genero tabla temporal de flujo anual
    ppGene_FlujoAnual sSentencia
    sSentencia = "SELECT * "
    sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpFlujoAnual "
    sSentencia = sSentencia & "ORDER BY TpoFjo, CodFjo"
  End If
  ' Seteo y activo el recordset
  Set porstMRp = New ADODB.Recordset
  With porstMRp
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
'    .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = sSentencia
    .Open
  End With
    
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cmbmoneda.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRFluCjaReM.rpt", "rptRFluCjaReA.rpt")
      '         .WindowShowGroupTree = True
      .WindowState = crptMaximized
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  End If
  
  porstMRp.Close
  Set porstMRp = Nothing
  pocnnMain.Close
  Set pocnnMain = Nothing
  ppHabilitacion True
  Exit Sub
Err:
  pocnnMain.Close
  Set pocnnMain = Nothing
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

']

Private Sub ppGene_FlujoAnual(sSentencia As String)
  
  Dim porstReg As ADODB.Recordset
  Dim porstTmp As ADODB.Recordset
  Dim nDebeIni As Double, nHaberIni As Double
  Dim nSaldoDeb As Double, nSaldoHab As Double, nImporte As Double
  Dim sSaldoMes As String, sCommandText As String
  Dim sConversion As String

  sConversion = "CONVERT(" & IIf(ps_Plataforma = pSrvMySql, "0, decimal(18, 2)", "decimal(18, 2), 0") & ")"
  ' Creo la tabla temporal del reporte
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpFlujoAnual", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#tmpFlujoAnual') DROP TABLE #tmpFlujoAnual")
  
  sCommandText = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpFlujoAnual ", "")
  sCommandText = sCommandText & "SELECT DISTINCT b.TpoFjo, (CASE b.TpoFjo WHEN '" & TPOFJO_ING & "' THEN '" & UCase(TPOFJO_ING_TXT) & "' ELSE '" & UCase(TPOFJO_EGR_TXT) & "' END) AS mDesTitulo, "
  sCommandText = sCommandText & "LEFT(a.CodFjo, 2) As mCodNivel, c.DetFjo AS mDetNivel, a.CodFjo, " & Choose(gsIdioma, "b.DetFjo", "b.DetFjox") & " AS DetFjo, "
  sCommandText = sCommandText & sConversion & " AS mDeb00, " & sConversion & " AS mHab00, " & sConversion & " AS mDeb01, " & sConversion & " AS mHab01, "
  sCommandText = sCommandText & sConversion & " AS mDeb02, " & sConversion & " AS mHab02, " & sConversion & " AS mDeb03, " & sConversion & " AS mHab03, "
  sCommandText = sCommandText & sConversion & " AS mDeb04, " & sConversion & " AS mHab04, " & sConversion & " AS mDeb05, " & sConversion & " AS mHab05, "
  sCommandText = sCommandText & sConversion & " AS mDeb06, " & sConversion & " AS mHab06, " & sConversion & " AS mDeb07, " & sConversion & " AS mHab07, "
  sCommandText = sCommandText & sConversion & " AS mDeb08, " & sConversion & " AS mHab08, " & sConversion & " AS mDeb09, " & sConversion & " AS mHab09, "
  sCommandText = sCommandText & sConversion & " AS mDeb10, " & sConversion & " AS mHab10, " & sConversion & " AS mDeb11, " & sConversion & " AS mHab11, "
  sCommandText = sCommandText & sConversion & " AS mDeb12, " & sConversion & " AS mHab12, " & sConversion & " AS mDeb13, " & sConversion & " AS mHab13, "
  sCommandText = sCommandText & sConversion & " AS mDebTo, " & sConversion & " AS mHabTo, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb00, " & sConversion & " AS mSalHab00, " & sConversion & " AS mSalDeb01, " & sConversion & " AS mSalHab01, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb02, " & sConversion & " AS mSalHab02, " & sConversion & " AS mSalDeb03, " & sConversion & " AS mSalHab03, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb04, " & sConversion & " AS mSalHab04, " & sConversion & " AS mSalDeb05, " & sConversion & " AS mSalHab05, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb06, " & sConversion & " AS mSalHab06, " & sConversion & " AS mSalDeb07, " & sConversion & " AS mSalHab07, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb08, " & sConversion & " AS mSalHab08, " & sConversion & " AS mSalDeb09, " & sConversion & " AS mSalHab09, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb10, " & sConversion & " AS mSalHab10, " & sConversion & " AS mSalDeb11, " & sConversion & " AS mSalHab11, "
  sCommandText = sCommandText & sConversion & " AS mSalDeb12, " & sConversion & " AS mSalHab12, " & sConversion & " AS mSalDeb13, " & sConversion & " AS mSalHab13, "
  sCommandText = sCommandText & sConversion & " AS mSalDebTo, " & sConversion & " AS mSalHabTo "
  sCommandText = sCommandText & IIf(ps_Plataforma = pSrvSql, "INTO #tmpFlujoAnual ", "")
  sCommandText = sCommandText & "FROM (((CoCpbDetFjo a "
  sCommandText = sCommandText & "LEFT JOIN CoFjo b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodFjo=b.CodFjo) "
  sCommandText = sCommandText & "LEFT JOIN CoFjo c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND LEFT(a.CodFjo, 2)=c.CodFjo) "
  sCommandText = sCommandText & "LEFT JOIN CoCta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCta=d.CodCta) "
  sCommandText = sCommandText & "WHERE a.codemp='" & gsCodEmp & "' "
  sCommandText = sCommandText & "AND a.pdoano='" & gsAnoAct & "' "
  sCommandText = sCommandText & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodFjo, '')<>'' "
  sCommandText = sCommandText & "AND a.MesPvs" & IIf(OptTipo(0).Value, "=", "<=") & "'" & gfCeros(cmbPeriodo.ListIndex, 2, 0, "0") & "' "
  sCommandText = sCommandText & "AND d.IndFjo='" & INDFJO_ACT & "' "
  sCommandText = sCommandText & "GROUP BY b.TpoFjo, a.CodFjo, b.DetFjo, c.DetFjo "
  sCommandText = sCommandText & "ORDER BY b.TpoFjo, a.CodFjo"
  pocnnMain.Execute sCommandText
  
  ' Seteo el recordset temporal para la grabacion
  Set porstTmp = New ADODB.Recordset
  With porstTmp
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open
  End With
  If Not (porstTmp.BOF And porstTmp.EOF) Then
    porstTmp.MoveFirst
    sSaldoMes = "": nSaldoDeb = 0: nSaldoHab = 0
    nDebeIni = CDec(IIf(IsNull(porstTmp!mSalDebe), 0, porstTmp!mSalDebe))
    nHaberIni = CDec(IIf(IsNull(porstTmp!mSalHaber), 0, porstTmp!mSalHaber))
    Set porstReg = New ADODB.Recordset
    With porstReg
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = "SELECT * "
      .Source = .Source & "FROM " & ps_Prefijo & "tmpFlujoAnual "
      .Source = .Source & "ORDER BY CodFjo"
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
      
    Do While Not porstTmp.EOF
      nImporte = 0
      If sSaldoMes <> porstTmp!mespvs Then
        sSaldoMes = porstTmp!mespvs
        nImporte = CDec(IIf(IsNull(porstTmp!mSalDebe), 0, porstTmp!mSalDebe))
        nSaldoDeb = CDec(nSaldoDeb) + nImporte
        nImporte = CDec(IIf(IsNull(porstTmp!mSalHaber), 0, porstTmp!mSalHaber))
        nSaldoHab = CDec(nSaldoHab) + nImporte
        ' Actualizo el saldo inicial del mes
        sCommandText = "UPDATE " & ps_Prefijo & "tmpFlujoAnual SET "
        sCommandText = sCommandText & "mSalDeb" & porstTmp!mespvs & "=" & CDbl(nSaldoDeb) & ", "
        sCommandText = sCommandText & "mSalHab" & porstTmp!mespvs & "=" & CDbl(nSaldoHab) & " "
        pocnnMain.Execute sCommandText
        porstReg.Requery
      End If
      ' Modifico los saldos de los flujos
      With porstReg
        If .RecordCount <> 0 Then .MoveFirst
        .Find "CodFjo = '" & porstTmp!CodFjo & "'"
        !CodFjo = porstTmp!CodFjo
        nImporte = IIf(IsNull(.Fields("mDeb" & porstTmp!mespvs)), 0, .Fields("mDeb" & porstTmp!mespvs))
        .Fields("mDeb" & porstTmp!mespvs) = CDec(nImporte) + porstTmp!mDebe
        nImporte = IIf(IsNull(.Fields("mHab" & porstTmp!mespvs)), 0, .Fields("mHab" & porstTmp!mespvs))
        .Fields("mHab" & porstTmp!mespvs) = CDec(nImporte) + porstTmp!mHaber
        !mDebTo = CDec(!mDebTo) + porstTmp!mDebe
        !mHabTo = CDec(!mHabTo) + porstTmp!mHaber
        !mSalDebTo = nDebeIni
        !mSalHabTo = nHaberIni
        .UpdateBatch
        .Requery
      End With
      porstTmp.MoveNext
    Loop
    porstReg.Close
    Set porstReg = Nothing
  End If
  porstTmp.Close
  Set porstTmp = Nothing

End Sub

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   
   MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
   optTipoImpresion(0).Enabled = tbHabilitar
   optTipoImpresion(1).Enabled = tbHabilitar
   cmdImprimir(0).Enabled = tbHabilitar
   cmdImprimir(1).Enabled = tbHabilitar
   cmdConfig.Enabled = tbHabilitar
   cmdSalir.Enabled = tbHabilitar

End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

'ini 2015-04-23 exporte excel
Private Sub toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  'no pinto datos Seleccion.Text = ButtonMenu.Text
  Select Case ButtonMenu.Key
   Case "A1": pExporta 1
   Case "A2": pExporta 2
  End Select
End Sub
'fin 2015-04-23 exporte excel
'ini 2015-04-23 exporte excel
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
    sTabla = "xlsCpbCabUsu"
   'pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS " & sTabla & " ", cCadReporte)
    pocnnTmp.Execute fDropTable2(sTabla, 1)
'ini 2015-04-23 exporte excel
Dim sMoneda As String
sMoneda = IIf(cmbmoneda.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
cCadReporte = cCadReporte & "select a.codemp,a.pdoano,a.mespvs,a.coddro,f.detdro,a.nrocpb,a.nroite,a.codfjo,e.detfjo,a.codcta,a.tpoctb,"
cCadReporte = cCadReporte & "IF(b.nroite=a.nroite,(case a.tpoctb when 'D' then a.imp" & sMoneda & " else 0.00 END),0) as debeMN,"
cCadReporte = cCadReporte & "IF(b.nroite=a.nroite,(case a.tpoctb when 'H' then a.imp" & sMoneda & " else 0.00 END),0) as haberMN,"
cCadReporte = cCadReporte & "b.codemp,b.pdoano,b.mespvs,b.coddro,b.nrocpb,b.nroite,"
cCadReporte = cCadReporte & "b.fehope,b.feedoc,b.codcta,d.detcta,b.codaux,c.razaux,b.codtdc,b.serdoc,b.nrodoc,b.refdoc,b.gloite,b.tpopvs,b.tpomon,"
cCadReporte = cCadReporte & "(case b.tpoctb when 'D' then b.imp" & sMoneda & " else 0.00 END) AS DebeMN, (case b.tpoctb when 'H' then b.imp" & sMoneda & " else 0.00 END) AS HaberMN "
cCadReporte = cCadReporte & "from (((((cocpbdetfjo a "
cCadReporte = cCadReporte & "LEFT JOIN cocpbdet b ON a.codemp=b.codemp and a.pdoano=b.pdoano and a.mespvs=b.mespvs and a.coddro=b.coddro and a.nrocpb=b.nrocpb) "
cCadReporte = cCadReporte & "LEFT JOIN TGAux c ON b.codemp=c.codemp and b.CodAux=c.CodAux) "
cCadReporte = cCadReporte & "LEFT JOIN COCta d ON b.codemp=d.codemp and b.pdoano=d.pdoano and b.CodCta=d.CodCta) "
cCadReporte = cCadReporte & "LEFT JOIN cofjo e ON a.codemp=e.codemp and a.pdoano=e.pdoano and a.Codfjo=e.Codfjo) "
cCadReporte = cCadReporte & "LEFT JOIN codro f ON a.codemp=f.codemp and a.pdoano=f.pdoano and a.coddro=f.coddro) "
'2015-04-23 exporte excelc CadReporte = cCadReporte & "where a.codemp='" & gsCodEmp & "' and a.pdoano='" & sPdoAnoFin & "' and a.mespvs<='" & sMesPvsFin & "' "
cCadReporte = cCadReporte & "where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' "
'cCadReporte = cCadReporte & " and a.mespvs<='" & sMesPvsFin & "' "
If TpoRpt = 1 Then
cCadReporte = cCadReporte & " and a.mespvs='" & gsMesAct & "' "
Else
cCadReporte = cCadReporte & " and a.mespvs<='" & gsMesAct & "' "
End If
'************************************************************ sMesPvsFin
cCadReporte = cCadReporte & "UNION "
'**********************************************************
cCadReporte = cCadReporte & "select a.codemp,a.pdoano,a.mespvs,a.coddro,f.detdro,a.nrocpb,a.nroite,b.codfjo,'' as detfjo,a.codcta,a.tpoctb,"
cCadReporte = cCadReporte & "(case a.tpoctb when 'D' then a.imp" & sMoneda & " else 0.00 END) AS DebeMN, (case a.tpoctb when 'H' then a.imp" & sMoneda & " else 0.00 END) AS HaberMN,"
cCadReporte = cCadReporte & "a.codemp,a.pdoano,a.mespvs,a.coddro,a.nrocpb,a.nroite,a.fehope,a.feedoc,a.codcta,d.detcta,a.codaux,'' as razaux,a.codtdc,a.serdoc,a.nrodoc,a.refdoc,a.gloite,a.tpopvs,a.tpomon,"
cCadReporte = cCadReporte & "(case a.tpoctb when 'D' then a.imp" & sMoneda & " else 0.00 END) AS DebeMN, (case a.tpoctb when 'H' then a.imp" & sMoneda & " else 0.00 END) AS HaberMN "
cCadReporte = cCadReporte & "from (((cocpbdet a "
cCadReporte = cCadReporte & "LEFT JOIN cocpbdetfjo b ON a.codemp=b.codemp and a.pdoano=b.pdoano and a.mespvs=b.mespvs and a.coddro=b.coddro and a.nrocpb=b.nrocpb) "
cCadReporte = cCadReporte & "LEFT JOIN COCta d ON a.codemp=d.codemp and a.pdoano=d.pdoano and a.CodCta=d.CodCta) "
cCadReporte = cCadReporte & "LEFT JOIN codro f ON a.codemp=f.codemp and a.pdoano=f.pdoano and a.coddro=f.coddro) "
cCadReporte = cCadReporte & "where a.codemp='" & gsCodEmp & "' and a.pdoano='" & gsAnoAct & "' "
'cCadReporte = cCadReporte & " and a.mespvs<='" & gsMesAct & "'"
If TpoRpt = 1 Then
cCadReporte = cCadReporte & " and a.mespvs='" & gsMesAct & "' "
Else
cCadReporte = cCadReporte & " and a.mespvs<='" & gsMesAct & "' "
End If
cCadReporte = cCadReporte & " and  mid(a.codcta,1,2)='10' and a.imp" & sMoneda & "<>0.00 and  a.indfjo_det<>9 and ifnull(b.codfjo,'')=''; "

'2015-04-23 exporte excelcCadReporte = cCadReporte & "where a.codemp='" & gsCodEmp & "' and a.pdoano='" & sPdoAnoFin & "' and a.mespvs<='" & sMesPvsFin & "' and  mid(a.codcta,1,2)='10' and a.imp" & xMon & "<>0.00 and  a.indfjo_det<>9 and ifnull(b.codfjo,'')=''; "
'fin 2015-04-23 exporte excel

    
'''    pocnnTmp.Execute cCadReporte
'fin 2015-04-23 exporte excel
    
'ini exporta datos a excel

    Dim porstTmp As ADODB.Recordset
    Set porstTmp = New ADODB.Recordset
    With porstTmp
       .ActiveConnection = pocnnTmp
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenForwardOnly
       .LockType = adLockReadOnly
       '2015-04-23 exporte excel .Source = "SELECT * FROM " & ps_Prefijo & sTabla
       .Source = cCadReporte
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
        
        .Cells(nRowI, 1).Value = "Informacion del Flujo Caja"
        nRowI = nRowI + 2
        Dim x1 As Integer
'ini 2015-04-23 exporte excel
Dim n As Integer
'total 33 colum +1
n = 0
n = n + 1: .Cells(nRowI, n).Value = "codemp"
n = n + 1: .Cells(nRowI, n).Value = "pdoano"
n = n + 1: .Cells(nRowI, n).Value = "mespvs"
n = n + 1: .Cells(nRowI, n).Value = "coddro"
n = n + 1: .Cells(nRowI, n).Value = "detdro"
n = n + 1: .Cells(nRowI, n).Value = "nrocpb"
n = n + 1: .Cells(nRowI, n).Value = "nroite"
n = n + 1: .Cells(nRowI, n).Value = "codfjo"
n = n + 1: .Cells(nRowI, n).Value = "detfjo"
n = n + 1: .Cells(nRowI, n).Value = "codcta"
n = n + 1: .Cells(nRowI, n).Value = "tpoctb"
n = n + 1: .Cells(nRowI, n).Value = "debeMN"
n = n + 1: .Cells(nRowI, n).Value = "haberMN"
n = n + 1: .Cells(nRowI, n).Value = "codemp"
n = n + 1: .Cells(nRowI, n).Value = "pdoano"
n = n + 1: .Cells(nRowI, n).Value = "mespvs"
n = n + 1: .Cells(nRowI, n).Value = "coddro"
n = n + 1: .Cells(nRowI, n).Value = "nrocpb"
n = n + 1: .Cells(nRowI, n).Value = "nroite"
n = n + 1: .Cells(nRowI, n).Value = "fehope"
n = n + 1: .Cells(nRowI, n).Value = "feedoc"
n = n + 1: .Cells(nRowI, n).Value = "codcta"
n = n + 1: .Cells(nRowI, n).Value = "detcta"
n = n + 1: .Cells(nRowI, n).Value = "codaux"
n = n + 1: .Cells(nRowI, n).Value = "razaux"
n = n + 1: .Cells(nRowI, n).Value = "codtdc"
n = n + 1: .Cells(nRowI, n).Value = "serdoc"
n = n + 1: .Cells(nRowI, n).Value = "nrodoc"
n = n + 1: .Cells(nRowI, n).Value = "refdoc"
n = n + 1: .Cells(nRowI, n).Value = "gloite"
n = n + 1: .Cells(nRowI, n).Value = "tpopvs"
n = n + 1: .Cells(nRowI, n).Value = "tpomon"
n = n + 1: .Cells(nRowI, n).Value = "DebeMN"
n = n + 1: .Cells(nRowI, n).Value = "HaberMN"
        
'''        .Cells(nRowI, 1).Value = "Periodo"
'''        .Cells(nRowI, 2).Value = "Nº Reg."
'''        .Cells(nRowI, 3).Value = "F.Cmpra"
'''        .Cells(nRowI, 4).Value = "F. Pago"
'''        .Cells(nRowI, 5).Value = "T.Doc"
'''        .Cells(nRowI, 6).Value = "Serie"
'''        .Cells(nRowI, 7).Value = "CemiDuadsi"
'''        .Cells(nRowI, 8).Value = "Nº Doc."
'''        .Cells(nRowI, 9).Value = "COSDCREFIS"
'''        .Cells(nRowI, 10).Value = "T.Prv"
'''        .Cells(nRowI, 11).Value = "RUC"
'''        .Cells(nRowI, 12).Value = "R.Social"
'''        .Cells(nRowI, 13).Value = "B. Gravada"
'''        .Cells(nRowI, 14).Value = "IGV Grab"
'''        .Cells(nRowI, 15).Value = "B. G/N Gr"
'''        .Cells(nRowI, 16).Value = "IGV G/N Gr"
'''        .Cells(nRowI, 17).Value = "B. Sin CF"
'''        .Cells(nRowI, 18).Value = "Igv S CF"
'''        .Cells(nRowI, 19).Value = "CIMPTOTNGV"
'''        .Cells(nRowI, 20).Value = "CISSC"
'''        .Cells(nRowI, 21).Value = "COTRTRICGO"
'''        .Cells(nRowI, 22).Value = "CIMPTOTCOM"
'''        .Cells(nRowI, 23).Value = "CTIPCAM"
'''        .Cells(nRowI, 24).Value = "CFECCOMMOD"
'''        .Cells(nRowI, 25).Value = "CTIPCOMMOD"
'''        .Cells(nRowI, 26).Value = "CNUMSERMOD"
'''        .Cells(nRowI, 27).Value = "CNUMCOMMOD"
'''        .Cells(nRowI, 28).Value = "CCOMNODOMI"
'''        .Cells(nRowI, 29).Value = "CEMIDEPDET"
'''        .Cells(nRowI, 30).Value = "CNUMDEPDET"
'''        .Cells(nRowI, 31).Value = "CCOMPGRET"
'''        .Cells(nRowI, 32).Value = "CESTOPE"
'''        .Cells(nRowI, 33).Value = "CVALFACIMP"
'''        .Cells(nRowI, 34).Value = "CINTDIAMAY"
'''        .Cells(nRowI, 35).Value = "CINTKARDEX"
'''        .Cells(nRowI, 36).Value = "CINTREG"
'''        .Cells(nRowI, 37).Value = "tsadetrac"
'''        .Cells(nRowI, 38).Value = "DetaDetrac"
'''        .Cells(nRowI, 39).Value = "PorcDetra"
'fin 2015-04-23 exporte excel
       
        'nRowI = nRowI + 1
        nRecord = .Cells(nRowI, nColI).CurrentRegion.Rows.Count
        nFields = .Cells(nRowI, nColI).CurrentRegion.Columns.Count
        nRowI = nRowI + 1 'limite inicial real
        nRecord = (nRowI + nRecord)
        If nRecord = 0 Then nRecord = nRowI
        
        .Range(.Cells(nRowI, 1), .Cells(.Rows.Count, nFields)).ClearContents
        
        .Cells(nRowI, nColI).CopyFromRecordset porstTmp
        .Columns.AutoFit ' ajusta el ancho de las columnas
        '.Cells(nRecord + 1, 32).Value = "=SUM(AF4:AF775)"
'        Range("AF89").Select
'        Range("AF89").FormulaR1C1 = "=@SUMA(AF4:AF88)"
        '=SUMA(AF4:AF88)
        'Sheets(oSheet).Select
        'Columns("N:N").Select
     '"=Sum(A1:A10)"
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
'fin 2015-04-23 exporte excel

