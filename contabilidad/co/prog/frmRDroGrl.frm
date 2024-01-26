VERSION 5.00
Begin VB.Form frmRDroGrl 
   Caption         =   "[título]"
   ClientHeight    =   4005
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2940
      TabIndex        =   19
      Top             =   2760
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   20
         Top             =   315
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.Frame fraNivel 
      Caption         =   "Nivel de Cuentas"
      ForeColor       =   &H80000002&
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   1305
      Width           =   2430
      Begin VB.OptionButton option1 
         Caption         =   "3 Dígitos"
         ForeColor       =   &H80000001&
         Height          =   240
         Index           =   1
         Left            =   1215
         TabIndex        =   7
         Top             =   300
         Width           =   960
      End
      Begin VB.OptionButton option1 
         Caption         =   "2 Dígitos"
         ForeColor       =   &H80000001&
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   1365
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
      ScaleWidth      =   5115
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3465
      Width           =   5115
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
         Picture         =   "frmRDroGrl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmRDroGrl.frx":0102
         Style           =   1  'Graphical
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
         Height          =   495
         Left            =   3720
         Picture         =   "frmRDroGrl.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkNuevaPagina 
      Alignment       =   1  'Right Justify
      Caption         =   "Nueva página por cada Diario"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   2535
      TabIndex        =   8
      Top             =   1440
      Width           =   2535
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmRDroGrl.frx":077E
      Left            =   3975
      List            =   "frmRDroGrl.frx":0780
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1875
      Width           =   1140
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5115
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   4755
         Picture         =   "frmRDroGrl.frx":0782
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4755
         Picture         =   "frmRDroGrl.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   4
         Top             =   480
         Width           =   570
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
         TabIndex        =   5
         Top             =   840
         Width           =   570
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Diarios"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   480
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
         Left            =   675
         TabIndex        =   14
         Top             =   495
         Width           =   4095
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
         Left            =   675
         TabIndex        =   13
         Top             =   840
         Width           =   4095
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
      Height          =   240
      Index           =   1
      Left            =   3300
      TabIndex        =   11
      Top             =   1920
      Width           =   645
   End
End
Attribute VB_Name = "frmRDroGrl"
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
']


Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
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

   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodDro"
         .Item(dnContador).MaxLength = porstCodro.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Diarios:", "Moneda:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Journals:", "Currency:")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraNivel.Caption = Choose(gsIdioma, "Nivel de Cuentas", "Account Level")
  option1(0).Caption = Choose(gsIdioma, "2 Dígitos", "2 Digits")
  option1(1).Caption = Choose(gsIdioma, "3 Dígitos", "3 Digits")
  chkNuevaPagina.Caption = Choose(gsIdioma, "Nueva página por cada Diario", "New page for each Journal")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCodro
      .MoveLast
      txtDato(1).Text = !coddro
      .MoveFirst
      txtDato(0).Text = !coddro
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   chkNuevaPagina.Value = 0
   option1(0).Value = True
   
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
   porstCodro.Close
   pocnnMain.Close
   Set porstCodro = Nothing
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
   Dim sMoneda As String

  ppHabilitacion False
   
  sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT, TPOMON_EXT_TXT)
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = "SELECT LEFT(a.CodDro, 2) As Diario, a.CodCta, " & Choose(gsIdioma, "b.DetCta", "b.DetCtax") & " AS DetCta, a.CodDro, " & Choose(gsIdioma, "c.DetDro", "c.DetDrox") & " AS DetDro, "
    .Source = .Source & Choose(gsIdioma, "d.DetDro", "d.DetDrox") & " AS Otro, LEFT(a.CodCta ,2) AS CuentaReg, "
    .Source = .Source & IIf(option1(1).Value, "LEFT(a.CodCta ,3) AS CuentaReg1, " & Choose(gsIdioma, "e.DetDro", "e.DetDrox") & " AS Otra, ", " ")
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cDebe, "
    .Source = .Source & "(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.Imp" & sMoneda & " ELSE 0 END) AS cHaber, "
    .Source = .Source & Choose(gsIdioma, "b2.DetCta", "b2.DetCtax") & " AS cDetalle2, "
    .Source = .Source & IIf(option1(1).Value, Choose(gsIdioma, "b3.DetCta", "b3.DetCtax"), "'x'") & " AS cDetalle3 "
    If usDEstino = PRN_DEST_GRAF Then
      .Source = .Source & "FROM COCpbDet a "
      .Source = .Source & "LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta "
      .Source = .Source & "LEFT JOIN CoDro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND LEFT(a.CodDro, 2)=RTrim(c.CodDro) "
      .Source = .Source & "LEFT JOIN CoDro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro "
      .Source = .Source & IIf(option1(1).Value, "LEFT JOIN CoDro e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.CodDro=e.CodDro ", " ")
      .Source = .Source & "LEFT JOIN cocta b2 ON a.codemp=b2.codemp AND a.pdoano=b2.pdoano AND LEFT(a.CodCta, 2)=b2.CodCta "
      .Source = .Source & IIf(option1(1).Value, "LEFT JOIN cocta b3 ON a.codemp=b3.codemp AND a.pdoano=b3.pdoano AND LEFT(a.CodCta, 3)=b3.CodCta ", " ")
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND a.Mespvs ='" & gsMesAct & "' "
      .Source = .Source & "AND a.CodDro BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
      .Source = .Source & "ORDER BY Diario, a.CodDro, a.CodCta ASC"
    Else
      pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRDroGrlCtaR2", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 17)='#trptRDroGrlCtaR2') DROP TABLE #trptRDroGrlCtaR2")
      cmdImprimir(Index).Tag = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptRDroGrlCtaR2 ", "")
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "SELECT LEFT(a.CodDro, 2) As Diario, LEFT(a.CodCta, 2) AS CuentaReg, "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "ROUND(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 2) AS cDebeTCtaR, "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "ROUND(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 2) AS cHaberTCtaR "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & IIf(ps_Plataforma = pSrvSql, "INTO #trptRDroGrlCtaR2 ", "")
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "FROM COCpbDet a "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "LEFT JOIN CoDro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND LEFT(a.CodDro, 2)=RTrim(c.CodDro) "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "LEFT JOIN CoDro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "WHERE a.codemp='" & gsCodEmp & "' "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "AND a.pdoano='" & gsAnoAct & "' "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "AND a.Mespvs ='" & gsMesAct & "' "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "AND a.CodDro BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
      cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "GROUP BY LEFT(a.CodDro, 2), LEFT(a.CodCta, 2)"
      pocnnMain.Execute cmdImprimir(Index).Tag
      If option1(1).Value = True Then
        pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRDroGrlCtaR3", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 17)='#trptRDroGrlCtaR3') DROP TABLE #trptRDroGrlCtaR3")
        cmdImprimir(Index).Tag = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS trptRDroGrlCtaR3 ", "")
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "SELECT LEFT(a.CodDro, 2) AS Diario, LEFT(a.CodCta, 3) AS CuentaReg, "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "ROUND(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END)), 2) AS cDebeTCtaR3, "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "ROUND(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END)), 2) AS cHaberTCtaR3 "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & IIf(ps_Plataforma = pSrvSql, "INTO #trptRDroGrlCtaR3 ", "")
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "FROM (((COCpbDet a "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta) "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "LEFT JOIN CoDro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND LEFT(a.CodDro, 2)=RTRim(c.CodDro)) "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "LEFT JOIN CoDro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro) "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "WHERE a.codemp='" & gsCodEmp & "' "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "AND a.pdoano='" & gsAnoAct & "' "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "AND a.Mespvs='" & gsMesAct & "' "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "AND a.CodDro BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
        cmdImprimir(Index).Tag = cmdImprimir(Index).Tag & "GROUP BY LEFT(a. CodDro,2), LEFT(a.CodCta ,3)"
        pocnnMain.Execute cmdImprimir(Index).Tag
      End If
      .Source = .Source & IIf(option1(1).Value, ", f.cDebeTCtaR, f.cHaberTCtaR,g.cDebeTCtaR3, g.cHaberTCtaR3", ",e.cDebeTCtaR, e.cHaberTCtaR") & " FROM COCpbDet a "
      .Source = .Source & "LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND a.CodCta=b.CodCta "
      .Source = .Source & "LEFT JOIN CoDro c ON a.codemp=c.codemp AND a.pdoano=c.pdoano AND LEFT(a.CodDro, 2)=RTrim(c.CodDro) "
      .Source = .Source & "LEFT JOIN CoDro d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodDro=d.CodDro "
      .Source = .Source & "LEFT JOIN cocta b2 ON a.codemp=b2.codemp AND a.pdoano=b2.pdoano AND LEFT(a.CodCta, 2)=b2.CodCta "
      If option1(1).Value Then
        .Source = .Source & "LEFT JOIN cocta b3 ON a.codemp=b3.codemp AND a.pdoano=b3.pdoano AND LEFT(a.CodCta, 3)=b3.CodCta "
        .Source = .Source & "LEFT JOIN CoDro e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.CodDro=e.CodDro "
        .Source = .Source & "LEFT JOIN " & ps_Prefijo & "trptRDroGrlCtaR2 f ON LEFT(a.CodDro, 2)=f.Diario AND LEFT(a.CodCta, 2)=f.CuentaReg "
        .Source = .Source & "LEFT JOIN " & ps_Prefijo & "trptRDroGrlCtaR3 g ON LEFT(a.CodDro, 2)=g.Diario AND LEFT(a.CodCta, 3)=g.CuentaReg "
      Else
        .Source = .Source & "LEFT JOIN " & ps_Prefijo & "trptRDroGrlCtaR2 e ON LEFT(a.CodDro, 2)=e.Diario AND LEFT(a.CodCta, 2)=e.CuentaReg "
      End If
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND a.Mespvs ='" & gsMesAct & "' "
      .Source = .Source & "AND a.CodDro BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
      .Source = .Source & "ORDER BY Diario, a.CodDro, a.CodCta ASC"
    End If
    .Open
  End With
   
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(option1(1).Value, "rptRDroGrl3.rpt", "rptRDroGrl.rpt")
      '         .WindowShowGroupTree = True
      
      'Fórmular propias.
      .Formulas(7) = "sNuevaPagina='" & IIf(chkNuevaPagina.Value, "S", "N") & "'"
      ']
      .WindowState = crptMaximized
      .MarginLeft = unMargenIzquierdo
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      If option1(0).Value = True Then
        .LoadReport gsRutRpt & "rptRDroGrl" & IIf(chkNuevaPagina.Value, "s", "") & ".mrp"
      End If
      If option1(1).Value = True Then
        .LoadReport gsRutRpt & "rptRDroGrl3" & IIf(chkNuevaPagina.Value, "s", "") & ".mrp"
      End If
      
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
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRDroGrlCtaR2", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 17)='#trptRDroGrlCtaR2') DROP TABLE #trptRDroGrlCtaR2")
    pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptRDroGrlCtaR3", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 17)='#trptRDroGrlCtaR3') DROP TABLE #trptRDroGrlCtaR3")
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
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
      With porstCodro
         .MoveFirst
         .Find "CoddRO='" & txtDato(tnIndex).Text & "'"
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

