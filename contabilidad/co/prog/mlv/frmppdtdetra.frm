VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPPDTDetra 
   Caption         =   "[título]"
   ClientHeight    =   3090
   ClientLeft      =   2640
   ClientTop       =   3960
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFormaPago 
      Height          =   315
      ItemData        =   "frmppdtdetra.frx":0000
      Left            =   1170
      List            =   "frmppdtdetra.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Width           =   2070
   End
   Begin VB.TextBox txtSecuencia 
      Alignment       =   1  'Right Justify
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
      Left            =   1170
      MaxLength       =   20
      TabIndex        =   4
      Top             =   825
      Width           =   990
   End
   Begin ComctlLib.ProgressBar pgbEtapa1 
      Height          =   285
      Left            =   210
      TabIndex        =   10
      Top             =   2010
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame fraRangos 
      Caption         =   " Banco "
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4635
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4260
         Picture         =   "frmppdtdetra.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   225
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
         TabIndex        =   1
         Top             =   225
         Width           =   450
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
         Left            =   570
         TabIndex        =   2
         Top             =   225
         Width           =   3705
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmppdtdetra.frx":01AE
      Left            =   3585
      List            =   "frmppdtdetra.frx":01B0
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   825
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog cdlDirectorio 
      Left            =   4005
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   855
      TabIndex        =   12
      Top             =   2490
      Width           =   1150
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   2610
      TabIndex        =   11
      Top             =   2490
      Width           =   1150
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Forma Pago :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   165
      TabIndex        =   7
      Top             =   1245
      Width           =   945
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Secuencia :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   165
      TabIndex        =   3
      Top             =   870
      Width           =   855
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   2730
      TabIndex        =   5
      Top             =   870
      Width           =   660
   End
   Begin VB.Label lblTexto 
      Caption         =   "Procesando Detracciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   255
      TabIndex        =   9
      Top             =   1740
      Width           =   2355
   End
End
Attribute VB_Name = "frmPPDTDetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2015-01-14
'he revisado este frm 100% en fuentes de mysql y lo he pasado a sql.
'sin errores. luego he copiado el frm en  OCont_sql
'frm compatible con fuentes de OCont_sql

Option Explicit

Private udFecha As Date
Private unCopias As Integer
Private unMargenIzquierdo As Integer
Private usDEstino As String
Private usOrientacionRpt As String
Private usOrientacionOri As String

Private pocnnMain As ADODB.Connection
Private porstCoBco As ADODB.Recordset

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0
    txtDato(Index).SetFocus
  End Select
  ppAyuBus Index
End Sub

Private Sub Form_Activate()
  cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
  Dim s_Archivo  As String, s_Expresion As String
  Dim s_Sql As String
  Dim porstValidacion As ADODB.Recordset
  
  ' Verifico que existan registros y seleccionados
  If txtDato(0).Text = "" Then Beep: MsgBox "Debe Ingresar Entidad Bancaria de transferencia", vbExclamation: txtDato(0).SetFocus: Exit Sub
  If CInt(txtSecuencia.Text) = 0 Then Beep: MsgBox "Secuencia de transferencia no valido", vbExclamation: txtSecuencia.SetFocus: Exit Sub
  ' Inicializo variables
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  pgbEtapa1.Value = 0: pgbEtapa1.Min = 0
  pgbEtapa1.Value = pgbEtapa1.Min
  
  Set porstValidacion = New ADODB.Recordset
  ' Valida archivo
  s_Sql = "SELECT DISTINCT 1 AS opcion, 'Transferencia Cuenta Detracción' AS desopcion, '2' AS caso, "
  s_Sql = s_Sql & IIf(ps_Plataforma = pSrvMySql, "CONCAT('Cuenta Detracción Vacio :', IFNULL(cta.nroctacte, ''), '/', cpr.codaux, ' / ', cpr.serdoc, '-', cpr.nrodoc)", "('Cuenta Detracción Vacio :' + ISNULL(cta.nroctacte, '')+ '/'+cpr.codaux+' / '+cpr.serdoc+'-'+cpr.nrodoc)") & " AS descripcion, '123' AS registro "
'ini 2015-01-14 conver a sql
  'ini 2014-04-05 error cuando aplico sql
  '-ORDER BY items must appear in the select list if SELECT DISTINCT is specified.
  If ps_Plataforma = pSrvMySql Then
  Else
  s_Sql = s_Sql & ",cpr.codaux, cpr.serdoc, cpr.nrodoc "
  End If
  'fin 2014-04-05 error cuando aplico sql
'fin 2015-01-14 conver a sql
  s_Sql = s_Sql & "FROM cocprdoc cpr "
  s_Sql = s_Sql & "INNER JOIN tgaux aux ON aux.codemp=cpr.codemp AND aux.codaux=cpr.codaux "
  s_Sql = s_Sql & "LEFT JOIN coctaban cta ON cta.codemp=cpr.codemp AND cta.codaux=cpr.codaux AND cta.codbco='" & txtDato(0).Text & "' AND cta.tpocta='" & TPOCTA_COR & "' AND cta.tpomon='" & TPOMON_NAC & "' "
  s_Sql = s_Sql & "WHERE cpr.codemp='" & gsCodEmp & "' "
  s_Sql = s_Sql & "AND cpr.pdoano='" & gsAnoAct & "' "
  s_Sql = s_Sql & "AND cpr.mespvs<='" & gsMesAct & "' "
  s_Sql = s_Sql & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cpr.tsadetrac, '0')<>'" & INDCDT_INA & "' "
  s_Sql = s_Sql & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cpr.nrocdt, '')='' "
  s_Sql = s_Sql & "AND cpr.indcdt='" & INDCDT_ACT & "' "
  s_Sql = s_Sql & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cta.nroctacte, '')='' "
  s_Sql = s_Sql & "ORDER BY cpr.codaux, cpr.serdoc, cpr.nrodoc"
  With porstValidacion
    .ActiveConnection = pocnnMain
    .Source = s_Sql
    .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open
  End With
  
  If (Not (porstValidacion.BOF And porstValidacion.EOF) Or porstValidacion.RecordCount > 1) Then
    MsgBox Choose(gsIdioma, "Validación de Información tiene Errores", "Validation of information has errors") & Chr$(13) & Choose(gsIdioma, "Presione Aceptar para Imprimir Reporte de Validación y pueda corregir sus Errores", "You Press Accept to print report of validation and can correct errors"), vbCritical
    ' Listado de Errores
    gpEncabezadoRpt frmMain.rptMain, Choose(gsIdioma, "Errores o Alertas de la Validación de Información", "Erros or Alerts of the Validation of Information"), Date, True, False, porstValidacion
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptLInfVal.rpt"
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .Action = 1
    End With
    porstValidacion.Close
    GoTo Finalizar
  End If
  porstValidacion.Close
  
  s_Expresion = Right(gsAnoAct, 2) & gfPadL(txtSecuencia.Text, 4, "0")
  s_Archivo = "D" & gsRUCEmp & s_Expresion & ".txt"
  
  On Error GoTo CancelaDialogo
  cdlDirectorio.DialogTitle = "Grabar Archivo Como"
  cdlDirectorio.CancelError = True
  cdlDirectorio.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
  cdlDirectorio.FileName = s_Archivo
  cdlDirectorio.DefaultExt = ".txt"
  cdlDirectorio.Filter = "Archivos de texto(*.txt)|*.txt|Todos los archivos(*.*)|*.*"
  cdlDirectorio.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then
    MsgBox error(Err.Number)
    Exit Sub
  End If
  On Error GoTo 0
  
  ChDir App.path
  If MsgBox("¿ Estás Seguro de Generar Archivo de Detracciones? ", vbQuestion + vbYesNo) = vbYes Then
    s_Archivo = cdlDirectorio.FileName
    ppExportaDetraccion s_Archivo, s_Expresion
    MsgBox TEXT_8008, vbInformation
  End If
  ChDrive Left$(App.path, 1)
  ChDir App.path
Finalizar:
  Set porstValidacion = Nothing
  ' Incializa controles
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub ppExportaDetraccion(ByVal sArchivo As String, ByVal sTransaccion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim s_Contenido As String, s_Detraccion As String
  Dim dsTexto As String, s_Caracter As String, s_Sentencia As String
  Dim nRegistro As Long, nRegistros As Long, n_Longitud As Long
  Dim n_Importe As Double, n_Detraccion As Double
  Dim nSumatoriaTotal As Double
  Dim porstRecordset As ADODB.Recordset
 
  Set porstRecordset = New ADODB.Recordset
  ' Información de archivo
  s_Sentencia = "SELECT cpr.codaux, cpr.serdoc, cpr.nrodoc, cpr.feedoc, cpr.tsadetrac, aux.razaux, aux.rucaux, cta.nroctacte, "
  s_Sentencia = s_Sentencia & "cpr.imptot_" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT) & " AS nImporte "
'ini 2015-01-13 error sum de bruto
  s_Sentencia = s_Sentencia & ",cpr.CodTdc "
'fin 2015-01-13 error sum de bruto
  s_Sentencia = s_Sentencia & "FROM cocprdoc cpr "
  s_Sentencia = s_Sentencia & "INNER JOIN tgaux aux ON aux.codemp=cpr.codemp AND aux.codaux=cpr.codaux "
  s_Sentencia = s_Sentencia & "LEFT JOIN coctaban cta ON cta.codemp=cpr.codemp AND cta.codaux=cpr.codaux AND cta.codbco='" & txtDato(0).Text & "' AND cta.tpocta='" & TPOCTA_COR & "' AND cta.tpomon='" & TPOMON_NAC & "' "
  s_Sentencia = s_Sentencia & "WHERE cpr.codemp='" & gsCodEmp & "' "
  s_Sentencia = s_Sentencia & "AND cpr.pdoano='" & gsAnoAct & "' "
  s_Sentencia = s_Sentencia & "AND cpr.mespvs<='" & gsMesAct & "' "
'ini 2015-01-14 conver a sql
  's_Sentencia = s_Sentencia & "AND IFNULL(cpr.tsadetrac, '0')<>'" & INDCDT_INA & "' "
  's_Sentencia = s_Sentencia & "AND IFNULL(cpr.nrocdt, '')='' "
  '2014-04-04 error conver sql s_Sentencia = s_Sentencia & "AND IFNULL(cpr.tsadetrac, '0')<>'" & INDCDT_INA & "' "
  '2014-04-04 error conver sql s_Sentencia = s_Sentencia & "AND IFNULL(cpr.nrocdt, '')='' "
  s_Sentencia = s_Sentencia & "AND " & fIsNull() & "cpr.tsadetrac, '0')<>'" & INDCDT_INA & "' "
  s_Sentencia = s_Sentencia & "AND " & fIsNull() & "cpr.nrocdt, '')='' "
'fin 2015-01-14 conver a sql
  
  s_Sentencia = s_Sentencia & "AND cpr.indcdt='" & INDCDT_ACT & "' "
  s_Sentencia = s_Sentencia & "ORDER BY cpr.codaux, cpr.serdoc, cpr.nrodoc"
  With porstRecordset
    .ActiveConnection = pocnnMain
    .Source = s_Sentencia
    .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Open
  End With
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    pgbEtapa1.Max = porstRecordset.RecordCount
    pgbEtapa1.Value = pgbEtapa1.Min
    ' Creo objeto de archivo
    Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
    Set potxtFileExp = pofsoFileExp.CreateTextFile(sArchivo, True)
    s_Caracter = " "
    nSumatoriaTotal = 0: nRegistros = 0
    
  Dim nContador As Integer ' 2014-04-05 reclasificacion de cod detraccion
    
    While Not porstRecordset.EOF
      s_Contenido = Left(porstRecordset!tsadetrac, 3)
'ini 2015-07-02 adic tabla detrac
         n_Detraccion = fTsaDetrac(pocnnMain, s_Contenido)
     
'      'ini 2014-04-05 reclasificacion de cod detraccion
'      For nContador = 1 To UBound(aDtraccDet, 1)
'        If Left(aDtraccDet(nContador), 3) = s_Contenido Then
'            n_Detraccion = aDtraccPor(nContador)
'            Exit For
'        End If
'      Next nContador
'      'fin 2014-04-05 reclasificacion de cod detraccion
'fin 2015-07-02 adic tabla detrac
      
      nSumatoriaTotal = nSumatoriaTotal + Round(CDec(porstRecordset!nImporte) * n_Detraccion, 0)
      nRegistros = nRegistros + 1
      porstRecordset.MoveNext
    Wend
    porstRecordset.MoveFirst
    
    ' Registro incial de archivo
    dsTexto = ""
    ' 1: Indicador de maestra - constante
    s_Contenido = "*": n_Longitud = 1
    dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
    ' 2: ruc empresa
    s_Contenido = gsRUCEmp: n_Longitud = 11
    dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
    ' 3: razon social empresa
    s_Contenido = gsRazEmp: n_Longitud = 35
    s_Contenido = Left(gsRazEmp, n_Longitud)
    dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
    ' 4: numero de lote
    s_Contenido = sTransaccion: n_Longitud = 6
    dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
    ' 5: importe total
    s_Contenido = Format(CDbl(nSumatoriaTotal), "############0.00") * 100: n_Longitud = 15
    
    dsTexto = dsTexto & gfPadL(s_Contenido, n_Longitud, "0")
    potxtFileExp.WriteLine dsTexto
    
    ' Detalle de archivo
    nRegistro = 0
  'Dim nContador As Integer '2015-01-13 error sum de bruto
    
    While Not porstRecordset.EOF
      dsTexto = ""
'ini 2015-01-14 corrige datos
      '0: ruc proveedor
      dsTexto = dsTexto & "6"
'fin 2015-01-14 corrige datos
     
      ' 1: ruc proveedor
      s_Contenido = porstRecordset!rucaux: n_Longitud = 11
      dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
'ini 2015-01-14 corrige datos
      ' 1A: razon social proveedor
      s_Contenido = porstRecordset!razAux: n_Longitud = 35
      dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
'fin 2015-01-14 corrige datos
      ' 2: numero de proforma - constante
      s_Contenido = "": n_Longitud = 9
      If cboFormaPago.ListIndex = 1 Then ' banco nacion
        ' 2: periodo tributario
        s_Contenido = Format(porstRecordset!feedoc, "yyyymm")
      End If
      dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, "0")
      ' 3: bien o servicio
      s_Contenido = porstRecordset!tsadetrac: n_Longitud = 3
      s_Contenido = Left(s_Contenido, n_Longitud)
      dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
      ' 4: cuenta bancaria proveedor
      s_Contenido = porstRecordset!nroctacte: n_Longitud = 11
      dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
      ' 5: importe depósito
      s_Contenido = Left(porstRecordset!tsadetrac, 3)
      
'ini 2015-07-02 adic tabla detrac
         n_Detraccion = fTsaDetrac(pocnnMain, s_Contenido)
      
''ini 2015-01-13 error sum de bruto
'      For nContador = 1 To UBound(aDtraccDet, 1)
'        If Left(aDtraccDet(nContador), 3) = s_Contenido Then
'            n_Detraccion = aDtraccPor(nContador)
'            Exit For
'        End If
'      Next nContador
''fin 2015-01-13 error sum de bruto
'fin 2015-07-02 adic tabla detrac
     n_Importe = Round(CDec(porstRecordset!nImporte) * n_Detraccion, 0)
      
      s_Contenido = Format(CDbl(n_Importe), "############0.00") * 100: n_Longitud = 15
      dsTexto = dsTexto & gfPadL(s_Contenido, n_Longitud, "0")
      ' 6: tipo de operación
      s_Contenido = porstRecordset!tsadetrac: n_Longitud = 2
      s_Contenido = Right(s_Contenido, n_Longitud)
      dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
      If cboFormaPago.ListIndex = 0 Then ' internet
        ' 7: periodo tributario
        s_Contenido = Format(porstRecordset!feedoc, "yyyymm"): n_Longitud = 6
        dsTexto = dsTexto & gfPadR(s_Contenido, n_Longitud, s_Caracter)
      End If
'ini 2015-01-13 error sum de bruto
        ' 8: tipo documento
      dsTexto = dsTexto & porstRecordset!codtdc
        ' 9: serie documento
      dsTexto = dsTexto & porstRecordset!serdoc
        ' 10: numero documento
      dsTexto = dsTexto & Right(porstRecordset!nrodoc, 8)
'fin 2015-01-13 error sum de bruto
      potxtFileExp.WriteLine dsTexto
      ' Incremento correlativo
      nRegistro = nRegistro + 1
      pgbEtapa1.Value = nRegistro
      porstRecordset.MoveNext
    Wend
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  End If
  porstRecordset.Close
  ' Actualizo lasecuencia de envio
  s_Sentencia = "UPDATE cocfg SET numera_dtr=" & CInt(txtSecuencia.Text)
  pocnnMain.Execute s_Sentencia

End Sub

Private Sub Form_Load()
   Dim sSentencia As String
   Dim nSecuencia As Integer
   
  Set pocnnMain = New ADODB.Connection
  Set porstCoBco = New ADODB.Recordset
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
   With porstCoBco
    .ActiveConnection = pocnnMain
    .Source = "SELECT a.codbco, "
    .Source = .Source & Choose(gsIdioma, "a.detbco", "a.detbcox") & " AS detbco "
    .Source = .Source & "FROM cobco a "
    .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
  '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open
   End With
  
  ' [Parámetros.                         'Cambiar.
  txtDato(0).MaxLength = porstCoBco!codbco.DefinedSize
  ' Obtengo la secuencia
  '2015-01-14 conver a sql sSentencia = "SELECT IFNULL(numera_dtr, 0) AS numera_dtr "
  '2014-04-05 error sql sSentencia = "SELECT IFNULL(numera_dtr, 0) AS numera_dtr "
  sSentencia = "SELECT " & fIsNull() & "numera_dtr, 0) AS numera_dtr "
  sSentencia = sSentencia & "FROM cocfg "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "'"
  nSecuencia = CInt(gfRetornaValor(CONNSTRG & gsNomBDS, sSentencia)) + 1
  txtSecuencia.MaxLength = 4
  txtSecuencia.Text = nSecuencia
  
  With cboTpoMon
    .AddItem TPOMON_NAC_TXT_1, 0
    .AddItem TPOMON_EXT_TXT_1, 1
  End With
  cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
  
  With cboFormaPago
    .AddItem "Internet", 0
    .AddItem "Banco de la Nación", 1
  End With
  cboFormaPago.ListIndex = 0
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(4, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Secuencia : ", "Moneda : ", "Forma Pago :", "Procesando Detracciones")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Sequence : ", "Currency : ", "Method Payment :", "Processing Detractions")
  Next nElemento
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  fraRangos.Caption = Choose(gsIdioma, " Banco ", " Bank ")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
  
  'Características de impresión.
  udFecha = Date                      'Fecha en el encabezado.
  unCopias = 1                        'Cantidad de Copias.
  unMargenIzquierdo = 240             'Margen izquierdo.
  usDEstino = PRN_DEST_GRAF           'PRN_DEST_GRAF:ica _
                                       PRN_DEST_MATR:icial.
  usOrientacionRpt = PRN_ORIE_VERT    'PRN_ORIE_VERT:ical _
                                       PRN_ORIE_HORI:zontal.

End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0
    modAyuBus.Bco_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
    With porstCoBco
      .MoveFirst
      .Find "codbco='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detbco), "", !detbco)
      End If
    End With
  End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
  porstCoBco.Close
  pocnnMain.Close
  Set porstCoBco = Nothing
  Set pocnnMain = Nothing
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
  txtDato(Index).SelStart = 0
  txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub
Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus Index
End Sub
Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index    'Busca el dato en su tabla principal.
   Case 0                          'Cambiar (añadir índices).
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
End Sub

Private Sub txtSecuencia_GotFocus()
  txtSecuencia.SelStart = 0
  txtSecuencia.SelLength = txtSecuencia.MaxLength
End Sub
Private Sub txtSecuencia_KeyPress(KeyAscii As Integer)
  If Len(Trim(txtSecuencia)) + 1 = txtSecuencia.MaxLength Then SendKeys "{TAB}"
End Sub
Private Sub txtSecuencia_Validate(Cancel As Boolean)
  txtSecuencia.Text = Interaction.IIf(Not IsNumeric(txtSecuencia.Text), 0, txtSecuencia.Text)
  txtSecuencia.Text = Math.Abs(Conversion.Val(txtSecuencia.Text))
  txtSecuencia.Text = Strings.FormatNumber(Conversion.CInt(txtSecuencia.Text), 0)
End Sub
