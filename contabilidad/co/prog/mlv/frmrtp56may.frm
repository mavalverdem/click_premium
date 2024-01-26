VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRTp56May 
   Caption         =   "[título]"
   ClientHeight    =   4125
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCabecera 
      Caption         =   "Imprime Cabecera"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   2280
      Width           =   1800
   End
   Begin VB.CheckBox chkFolio 
      Caption         =   "Folio Inicial"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   2280
      Width           =   1800
   End
   Begin VB.CheckBox chkRango 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1260
      TabIndex        =   22
      Top             =   2640
      Width           =   180
   End
   Begin VB.Frame fraRngPeriodo 
      Caption         =   " Rango Saldos "
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   15
      TabIndex        =   21
      Top             =   2640
      Width           =   3870
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   300
         Width           =   1410
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   23
         Top             =   345
         Width           =   705
      End
   End
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame fraMeses 
      Caption         =   " Rango de Meses "
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   15
      TabIndex        =   11
      Top             =   1440
      Width           =   4245
      Begin VB.ComboBox cmbMeses 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   280
         Width           =   1410
      End
      Begin VB.ComboBox cmbMeses 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   2670
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   280
         Width           =   1410
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Fin  : "
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   2235
         TabIndex        =   14
         Top             =   345
         Width           =   345
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Inicio : "
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   345
         Width           =   555
      End
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   26
      Top             =   2820
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   27
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1005
         TabIndex        =   28
         Top             =   315
         Width           =   1035
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmrtp56may.frx":0000
      Left            =   5640
      List            =   "frmrtp56may.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1515
      Width           =   1350
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   15
      TabIndex        =   5
      Top             =   75
      Width           =   6975
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6615
         Picture         =   "frmrtp56may.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   495
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
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6615
         Picture         =   "frmrtp56may.frx":01AE
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   855
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
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   585
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
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   5550
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
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   5550
      End
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7005
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3585
      Width           =   7005
      Begin VB.CommandButton cmdExporta 
         Caption         =   "&Genera Archivo"
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
         Left            =   3600
         Picture         =   "frmrtp56may.frx":0358
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Height          =   495
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
         Height          =   495
         Left            =   5745
         Picture         =   "frmrtp56may.frx":045A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Height          =   495
         Index           =   0
         Left            =   0
         Picture         =   "frmrtp56may.frx":05A4
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
         Height          =   495
         Index           =   1
         Left            =   1245
         Picture         =   "frmrtp56may.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin MSComDlg.CommonDialog cdlUbicacion 
      Left            =   3945
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTexto 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   4860
      TabIndex        =   16
      Top             =   1560
      Width           =   675
   End
End
Attribute VB_Name = "frmRTp56May"
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
Private porstCOCta As ADODB.Recordset
']

Private Sub chkRango_Click()
  fraRngPeriodo.Enabled = (chkRango.Value = vbChecked)
  cmbMeses(0).ListIndex = IIf((chkRango.Value = vbChecked), 1, cmbMeses(0).ListIndex)
  cmbMeses(0).Enabled = (chkRango.Value = vbUnchecked)
End Sub

Private Sub cmdExporta_Click()
  Dim s_MesIni As String, s_MesFin As String
  Dim s_SalAno As String, s_SalMes As String
  Dim sArchivo As String, sCadena As String
  Dim sCaracter As String, sMoneda As String, sRegistro As String
  Dim nImporte As Double
  
  s_SalAno = gsAnoAct
  s_MesIni = Format(cmbMeses(0).ListIndex, "00")
  s_MesFin = Format(cmbMeses(1).ListIndex, "00")
  If Not (s_MesFin >= s_MesIni) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial; Verificar", "End Month must be equal or more than opening; Verify"), vbExclamation: cmbMeses(1).SetFocus: Exit Sub
  
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_SalAno = Right(cmbPeriodo(0), 4)
    s_SalMes = Format(cmbPeriodo(1).ListIndex, "00")
    If (s_SalAno = gsAnoAct) And Not (s_SalMes <= s_MesIni) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbMeses(0).SetFocus: Exit Sub
    s_MesIni = s_SalMes
  End If
  ppHabilitacion False
   
  '[ Inicializo variables y nombre de archivo
  sArchivo = gsRUCEmp & gsAnoAct & gsMesAct & ".ema"
  sCaracter = ";"
  cdlUbicacion.FileName = sArchivo
  cdlUbicacion.ShowSave
  Open sArchivo For Output As #1
  
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  ' Recupero la informacion
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT det.codcta, " & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, "
    .Source = .Source & "det.codcco, " & Choose(gsIdioma, "cco.detcco", "cco.detccox") & " AS detcco, cta.tpomon, "
    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.imp" & sMoneda & " ELSE 0 END), 2) AS nDebe, "
    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.imp" & sMoneda & " ELSE 0 END), 2) AS nHaber "
    .Source = .Source & "FROM cocpbdet det "
    .Source = .Source & "LEFT JOIN cocta cta ON det.codemp=cta.codemp AND det.pdoano=cta.pdoano AND det.codcta=cta.codcta "
    .Source = .Source & "LEFT JOIN cocco cco ON det.codemp=cco.codemp AND det.pdoano=cco.pdoano AND det.codcco=cco.codcco "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & ">='" & s_SalAno & s_MesIni & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & "<='" & gsAnoAct & s_MesFin & "' "
    .Source = .Source & "AND det.mespvs NOT IN ('00', '13') "
    .Source = .Source & "AND det.codcta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "GROUP BY det.codcta, det.codcco "
    .Source = .Source & "ORDER BY det.codcta"
    .Open
  End With
  
  ' Verifico si existe registros
  If porstMRp.RecordCount > 0 Then
    porstMRp.MoveFirst
    Do While Not porstMRp.EOF
      sMoneda = IIf(porstMRp!tpomon = TPOMON_NAC, "PEN", "USD")
      nImporte = CDec(porstMRp!nDebe - porstMRp!nHaber)
      ' Genero la cadena si importe es diferente de cero
      If nImporte <> 0 Then
        ' Inicializo la cadena
        sCadena = ""
        sRegistro = Trim(IIf(IsNull(porstMRp!codcta), "", porstMRp!codcta))
        sCadena = sCadena & sRegistro & sCaracter
        sRegistro = Trim(IIf(IsNull(porstMRp!codcco), "", porstMRp!codcco))
        sCadena = sCadena & sRegistro & sCaracter
        sCadena = sCadena & sCaracter
        sCadena = sCadena & sCaracter
        sCadena = sCadena & sCaracter
        sCadena = sCadena & sMoneda & sCaracter
        sRegistro = Left(Trim(IIf(IsNull(porstMRp!codcta), "", porstMRp!codcta)), 1)
        sRegistro = IIf(sRegistro < 6, "", "-")
        sCadena = sCadena & sRegistro & sCaracter
        sCadena = sCadena & Format(Abs(nImporte), "#0.00")
        Print #1, sCadena
      End If
      porstMRp.MoveNext
    Loop
  End If
  Close #1
  porstMRp.Close
  ppHabilitacion True

End Sub

Private Sub Form_Load()
'   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   
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
   With porstCOCta
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCta, " & Choose(gsIdioma, "DetCta", "DetCtax") & " AS DetCta "
      .Source = .Source & "FROM COCta "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "ORDER BY CodCta"
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
   
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuentas :", "Inicio :", "Fin :", "Moneda :", "Inicio :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Accounts :", "Beginning :", "End :", "Currency :", "Beginning :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraMeses.Caption = Choose(gsIdioma, "Rango de Meses", "Range of Months")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraRngPeriodo.Caption = Choose(gsIdioma, "Rango Saldos", "Range Balances")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !codcta
      .MoveFirst
      txtDato(0).Text = !codcta
   End With
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
   ' Configuro los controles de año y mes
    For dnContador = (Val(gsAnoAct) - 10) To Val(gsAnoAct): cmbPeriodo(0).AddItem Choose(gsIdioma, "Año ", "Year ") & dnContador: Next dnContador
    cmbPeriodo(0).ListIndex = 9
    
    For dnContador = 0 To 13
      If gsIdioma = NvlUsr_Sup Then
        cmbMeses(0).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
        cmbMeses(1).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
        cmbPeriodo(1).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
      Else
        cmbMeses(0).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
        cmbMeses(1).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
        cmbPeriodo(1).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
      End If
    Next dnContador
    cmbMeses(0).ListIndex = Val(gsMesAct)
    cmbMeses(1).ListIndex = Val(gsMesAct)
    cmbPeriodo(1).ListIndex = 0
    fraRngPeriodo.Enabled = False
 
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
   porstCOCta.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim nContador As Integer, sMoneda As String
  Dim s_MesIni As String, s_MesFin As String
  Dim s_SalAno As String, s_SalMes As String
  Dim s_Sentencia As String, s_Sql As String
  Dim l_CreateTB As Boolean, n_Index As Integer
  Dim s_Catalogo As String
  Dim sSalAntDeb As String, sSalAntHab As String
  Dim sRegistro As String
    
  s_MesIni = Format(cmbMeses(0).ListIndex, "00")
  s_MesFin = Format(cmbMeses(1).ListIndex, "00")
  If Not (s_MesFin >= s_MesIni) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial; Verificar", "End Month must be equal or more than opening; Verify"), vbExclamation: cmbMeses(1).SetFocus: Exit Sub
  
  ' Valido el rango de periodos
  If chkRango.Value = vbChecked Then
    s_SalAno = Right(cmbPeriodo(0), 4)
    s_SalMes = Format(cmbPeriodo(1).ListIndex, "00")
    If (s_SalAno = gsAnoAct) And Not (s_SalMes <= s_MesIni) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbMeses(0).SetFocus: Exit Sub
  End If
  ppHabilitacion False
   
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  ' Obtengo suma de saldos anteriores
  If chkRango.Value = vbChecked Then
    ' Genero la tabla temporal de saldos
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRngSaldos", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpRngSaldos') DROP TABLE #tmpRngSaldos")
    
    s_Sentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpRngSaldos ", "")
    s_Sentencia = s_Sentencia & "SELECT * "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmpRngSaldos ", "")
    s_Sentencia = s_Sentencia & "FROM CoCtaAcu WHERE CodCta='tmp'"
    pocnnMain.Execute s_Sentencia
    For nContador = Val(s_SalAno) To (Val(gsAnoAct) - 1)
      s_Catalogo = Format(nContador, "0000")
      sSalAntDeb = "": sSalAntHab = ""
      s_SalMes = IIf(nContador = Val(s_SalAno), s_SalMes, "01")
      For n_Index = Val(s_SalMes) To 12
        sSalAntDeb = sSalAntDeb & "AcuD" & Format(Trim(n_Index), "00") & "_" & sMoneda & IIf(n_Index = 12, "", ", ")
        sSalAntHab = sSalAntHab & "AcuH" & Format(Trim(n_Index), "00") & "_" & sMoneda & IIf(n_Index = 12, "", ", ")
      Next n_Index
      s_Sentencia = "INSERT INTO " & ps_Prefijo & "tmpRngSaldos (CodCta, " & sSalAntDeb & ", " & sSalAntHab & ") "
      s_Sentencia = s_Sentencia & "SELECT a.CodCta, " & sSalAntDeb & ", " & sSalAntHab & " "
      s_Sentencia = s_Sentencia & "FROM (CoCtaAcu a "
      s_Sentencia = s_Sentencia & "LEFT JOIN CoCta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
      s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
      s_Sentencia = s_Sentencia & "AND a.pdoano='" & s_Catalogo & "' "
      s_Sentencia = s_Sentencia & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
      s_Sentencia = s_Sentencia & "ORDER BY a.CodCta"
      pocnnMain.Execute s_Sentencia
    Next nContador
    ' Genero tabla temporal con saldo finales
    sSalAntDeb = "": sSalAntHab = ""
    For n_Index = 0 To 13
      s_Sql = "AcuD" & Format(Trim(n_Index), "00") & "_" & sMoneda
      sSalAntDeb = sSalAntDeb & "ROUND(SUM(" & s_Sql & "), 2) AS " & s_Sql & IIf(n_Index = 13, "", ", ")
      s_Sql = "AcuH" & Format(Trim(n_Index), "00") & "_" & sMoneda
      sSalAntHab = sSalAntHab & "ROUND(SUM(" & s_Sql & "), 2) AS " & s_Sql & IIf(n_Index = 13, "", ", ")
    Next n_Index
    
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpSaldosIni", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpSaldosIni') DROP TABLE #tmpSaldosIni")
    s_Sentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpSaldosIni ", "")
    s_Sentencia = s_Sentencia & "SELECT '" & gsCodEmp & "' AS CodEmp, '" & gsAnoAct & "' AS pdoano, Codcta, " & sSalAntDeb & ", " & sSalAntHab & " "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvSql, "INTO #tmpSaldosIni ", "")
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "tmpRngSaldos "
    s_Sentencia = s_Sentencia & "GROUP BY CodCta "
    s_Sentencia = s_Sentencia & "ORDER BY CodCta"
    pocnnMain.Execute s_Sentencia
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRngSaldos", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpRngSaldos') DROP TABLE #tmpRngSaldos")
  End If
   
   ' Cadena de saldo anterior
   With porstMRp
      If .State = adStateOpen Then .Close
      s_Catalogo = IIf(ps_Plataforma = pSrvMySql, "tmpSaldosIni", "#tmpSaldosIni")
      s_Catalogo = IIf(chkRango.Value = vbChecked, s_Catalogo, "coctaacu")
      s_Sentencia = "SELECT a.MesPvs AS MesPvs, a.CodCta AS CodCta, a.CodDro AS CodDro, a.NroCpb AS NroCpb, a.NroIte AS NroIte, a.FehOpe, "
      s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc)", "(e.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
      s_Sentencia = s_Sentencia & "a.CodAux, b.RazAux, a.RefDoc, "
      s_Sentencia = s_Sentencia & Choose(gsIdioma, "a.GloIte", "a.GloItex") & " AS GloIte, "
      s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
      s_Sentencia = s_Sentencia & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
      s_Sentencia = s_Sentencia & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta , " & Choose(gsIdioma, "d.DetDro", "d.DetDrox") & " AS DetDro, e.AbvTDc, "
      If s_MesIni <> "00" Then
        sSalAntDeb = "ROUND(("
        sSalAntHab = "ROUND(("
        s_SalMes = IIf(chkRango.Value = vbChecked, "13", s_MesIni)
        For nContador = 0 To (Val(s_SalMes) - 1)
          sSalAntDeb = sSalAntDeb & "acu.AcuD" & Format(nContador, "00") & "_" & sMoneda & IIf(nContador = (Val(s_SalMes) - 1), ")", "+")
          sSalAntHab = sSalAntHab & "acu.AcuH" & Format(nContador, "00") & "_" & sMoneda & IIf(nContador = (Val(s_SalMes) - 1), ")", "+")
        Next nContador
        sSalAntDeb = sSalAntDeb & ", 2)"
        sSalAntHab = sSalAntHab & ", 2)"
        s_Sentencia = s_Sentencia & sSalAntDeb & " AS cAntCtaDeb, "
        s_Sentencia = s_Sentencia & sSalAntHab & " AS cAntCtaHab "
      Else
        s_Sentencia = s_Sentencia & "0 AS cAntCtaDeb, 0 AS cAntCtaHab "
      End If
      s_Sentencia = s_Sentencia & "FROM ((((COCpbDet a "
      s_Sentencia = s_Sentencia & "LEFT JOIN TGAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
      s_Sentencia = s_Sentencia & "LEFT JOIN COCta c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND a.CodCta=c.CodCta) "
      s_Sentencia = s_Sentencia & "LEFT JOIN CODro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
      s_Sentencia = s_Sentencia & "LEFT JOIN TGTDc e ON a.codemp=e.codemp AND a.CodTDc=e.CodTDc) "
      s_Sentencia = s_Sentencia & "LEFT JOIN " & s_Catalogo & " acu ON a.codemp=acu.codemp AND a.pdoano=acu.pdoano AND a.CodCta=acu.CodCta "
      s_Sentencia = s_Sentencia & "WHERE a.codemp='" & gsCodEmp & "' "
      s_Sentencia = s_Sentencia & "AND a.pdoano='" & gsAnoAct & "' "
      s_Sentencia = s_Sentencia & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
      s_Sentencia = s_Sentencia & "AND a.MesPvs>='" & s_MesIni & "' AND a.MesPvs<='" & s_MesFin & "' "
      If s_MesIni <> "00" Then
        s_Sentencia = s_Sentencia & "UNION "
        s_Sentencia = s_Sentencia & "SELECT '00' AS MesPvs, c.CodCta AS CodCta, '', '', '', NULL, '', '', '', '', '', 0, 0, "
        s_Sentencia = s_Sentencia & Choose(gsIdioma, "c.DetCta", "c.DetCtax") & " AS DetCta , '', '', "
        s_Sentencia = s_Sentencia & sSalAntDeb & " AS cAntCtaDeb, "
        s_Sentencia = s_Sentencia & sSalAntHab & " AS cAntCtaHab "
        s_Sentencia = s_Sentencia & "FROM (COCta c "
        s_Sentencia = s_Sentencia & "LEFT JOIN COCpbDet a ON c.codemp=a.codemp AND c.pdoano=a.pdoano AND c.CodCta=a.CodCta) "
        s_Sentencia = s_Sentencia & "LEFT JOIN " & s_Catalogo & " acu ON c.codemp=acu.codemp AND c.pdoano=acu.pdoano AND c.CodCta=acu.CodCta "
        s_Sentencia = s_Sentencia & "WHERE c.codemp='" & gsCodEmp & "' "
        s_Sentencia = s_Sentencia & "AND c.pdoano='" & gsAnoAct & "' "
        s_Sentencia = s_Sentencia & "AND c.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
        s_Sentencia = s_Sentencia & "AND c.TpoCta='" & TPOCTA_TRA & "' "
        If ps_Plataforma = pSrvMySql Then
          s_Sentencia = s_Sentencia & "HAVING ROUND(cAntCtaDeb-cAntCtaHab, 2)<>0.00 "
        Else
          s_Sentencia = s_Sentencia & "AND ROUND(" & sSalAntDeb & "-" & sSalAntHab & ", 2)<>0.00 "
        End If
      End If
      s_Sentencia = s_Sentencia & "ORDER BY CodCta, MesPvs, CodDro, NroCpb, NroIte"
      .Source = s_Sentencia
      .Open
   End With

  sRegistro = IIf(chkRango.Value = vbChecked, cmbPeriodo(1).Text & " " & Right(cmbPeriodo(0).Text, 4) & " - ", "")
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      .ReportFileName = gsRutRpt & "rptr56mayaux.rpt"
      '[ Formulas adicionales del reporte
      .Formulas(5) = "mPeriodo='" & sRegistro & " " & gfMesLet("01" & gsMesAct & gsAnoAct, 0, "", 1, " ", 1) & "'"
      sRegistro = FormatNumber(Left(Trim(gsRUCEmp), 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(6) = "mRucEmpresa='" & sRegistro & Mid(Trim(gsRUCEmp), 9) & "'"
      .Formulas(7) = "mDireccion='" & gsDirEmp & "'"
      .Formulas(8) = "mDistrito='" & gsLocEmp & "'"
      .Formulas(9) = "mActividad='" & gsGirEmp & "'"
      .Formulas(10) = "mRepresentante='" & gsRepEmp & "'"
      sRegistro = FormatNumber(Left(gsRepDNIEmp, 8), 0)
      sRegistro = Replace(sRegistro, ",", ".")
      .Formulas(11) = "mDniRepresentante='" & sRegistro & Mid(gsRepDNIEmp, 9) & "'"
      sRegistro = IIf(chkCabecera.Value = vbChecked, "S", "N")
      .ParameterFields(1) = "Cabecera;" & sRegistro & ";true"
      sRegistro = IIf(chkFolio.Value = vbChecked, "S", "N")
      .ParameterFields(2) = "FolioInicial;" & sRegistro & ";true"
      
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .WindowState = crptMaximized
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    
    With MRViewer
      .DataRecordSet = porstMRp
      .LoadReport gsRutRpt & "rptRMayAux.mrp"
      
      gpEncabezadoMRp MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value
      '[Parámetros adicionales.
      '         .Parameters("pTipoFecha") = IIf(optFecha(0).Value, "Emisión", "Cancelac.")
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
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpSaldosIni", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#tmpSaldosIni') DROP TABLE #tmpSaldosIni")

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
'FALTA VALIDAR LOS DATOS NUMERICOS
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
'   Select Case Index    'Completa con ceros a la izquierda.
'   Case 0, 1                           'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0, 1
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstCOCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
         End If
      End With
   End Select
End Function

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  MousePointer = IIf(tbHabilitar, vbDefault, vbHourglass)
  optTipoImpresion(0).Enabled = tbHabilitar
  optTipoImpresion(1).Enabled = tbHabilitar
  cmdImprimir(0).Enabled = tbHabilitar
  cmdImprimir(1).Enabled = tbHabilitar
  cmdConfig.Enabled = tbHabilitar
  cmdSalir.Enabled = tbHabilitar
  cmdExporta.Enabled = tbHabilitar
End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

