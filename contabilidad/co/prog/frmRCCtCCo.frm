VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRCCtCCo 
   Caption         =   "[título]"
   ClientHeight    =   5760
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   34
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   5100
      TabIndex        =   33
      Top             =   4320
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
         Left            =   1005
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Pendientes al"
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   0
      TabIndex        =   32
      Top             =   3360
      Width           =   2175
      Begin MSComCtl2.DTPicker TxtFecha 
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Format          =   51052545
         CurrentDate     =   37974
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   2175
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   11
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rangos"
      ForeColor       =   &H80000002&
      Height          =   3210
      Left            =   0
      TabIndex        =   18
      Top             =   45
      Width           =   7290
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   6920
         Picture         =   "frmRCCtCCo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2715
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
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   2700
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   4
         Left            =   4560
         Picture         =   "frmRCCtCCo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   4560
         Picture         =   "frmRCCtCCo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   540
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
         Index           =   4
         Left            =   135
         TabIndex        =   5
         Top             =   900
         Width           =   630
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
         Index           =   3
         Left            =   135
         TabIndex        =   4
         Top             =   540
         Width           =   630
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6600
         Picture         =   "frmRCCtCCo.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1980
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6600
         Picture         =   "frmRCCtCCo.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1620
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
         Left            =   135
         TabIndex        =   7
         Top             =   1965
         Width           =   945
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
         Left            =   135
         TabIndex        =   6
         Top             =   1620
         Width           =   945
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Auxiliar"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   31
         Top             =   2475
         Width           =   495
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
         Index           =   2
         Left            =   1395
         TabIndex        =   30
         Top             =   2700
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costos"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   28
         Top             =   270
         Width           =   1215
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
         Index           =   4
         Left            =   765
         TabIndex        =   25
         Top             =   900
         Width           =   3795
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
         Index           =   3
         Left            =   765
         TabIndex        =   24
         Top             =   540
         Width           =   3795
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
         TabIndex        =   23
         Top             =   1965
         Width           =   5520
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
         TabIndex        =   22
         Top             =   1620
         Width           =   5520
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   1350
         Width           =   585
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   6195
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3420
      Width           =   1080
   End
   Begin VB.PictureBox picOpciones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7290
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5220
      Width           =   7290
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
         Left            =   3720
         Picture         =   "frmRCCtCCo.frx":0852
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
         Height          =   495
         Index           =   0
         Left            =   0
         Picture         =   "frmRCCtCCo.frx":099C
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
         Left            =   1260
         Picture         =   "frmRCCtCCo.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
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
      Index           =   3
      Left            =   5340
      TabIndex        =   16
      Top             =   3465
      Width           =   765
   End
End
Attribute VB_Name = "frmRCCtCCo"
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
Private porstTGAux As ADODB.Recordset
Private porstCoCCo As ADODB.Recordset

']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
    Set pocnnMain = New ADODB.Connection
    Set porstMRp = New ADODB.Recordset
    Set porstCOCta = New ADODB.Recordset   'Cuentas
    Set porstTGAux = New ADODB.Recordset   'Auxiliar
    Set porstCoCCo = New ADODB.Recordset   'Centro de Costos
   
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
 
    With porstTGAux
       .ActiveConnection = pocnnMain
       .Source = "SELECT CodAux, RazAux "
       .Source = .Source & "FROM TgAux "
       .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
       .Source = .Source & "ORDER BY CodAux"
    '     .CursorLocation = adUseClient   'Es el Default.
       .CursorType = adOpenDynamic
       .LockType = adLockReadOnly
       .Open
    End With
    
    With porstCoCCo
       .ActiveConnection = pocnnMain
       .Source = "SELECT CodCCo, " & Choose(gsIdioma, "DetCCo", "DetCCox") & " AS DetCCo "
       .Source = .Source & "FROM CoCCo "
       .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
       .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
       .Source = .Source & "ORDER BY CodCCo"
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
      For dnContador = 2 To 2
         .Item(dnContador).DataField = "CodAux"
         .Item(dnContador).MaxLength = porstTGAux.Fields(.Item(dnContador).DataField).DefinedSize
      Next
      For dnContador = 3 To 4
         .Item(dnContador).DataField = "CodCCo"
         .Item(dnContador).MaxLength = porstCoCCo.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   
   End With
 ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(4, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Centro de Costos :", "Cuentas :", "Auxiliar :", "Moneda :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Cost Center :", "Accounts :", "Auxiliary :", "Currency :")
  Next nElemento
  fraRangos.Caption = Choose(gsIdioma, "Rango", "Range")
  fraFecha.Caption = Choose(gsIdioma, "Pendientes al ", "Pending to ")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  chkImpFecha.Caption = Choose(gsIdioma, "Imprime Fecha", "Print Date")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !CodCta
      .MoveFirst
      txtDato(0).Text = !CodCta
   End With
  
   With porstTGAux
      '.MoveLast
      'txtDato(1).Text = !CodCta
      'Beto
      '.MoveFirst
      'txtDato(2).Text = !CodAux
   End With
   With porstCoCCo
      .MoveLast
      txtDato(4).Text = !CodCCo
      .MoveFirst
      txtDato(3).Text = !CodCCo
   End With
  
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
   If txtDato(2).Text <> "" Then ppAyuDet 2
  
   If txtDato(3).Text <> "" Then ppAyuDet 3
   If txtDato(4).Text <> "" Then ppAyuDet 4
  
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   OptTipo(0).Value = True
   TxtFecha.MinDate = "01/" & gsMesAct & "/" & gsAnoAct
   TxtFecha.MaxDate = gfUltDia(TxtFecha.MinDate)
   TxtFecha.Value = TxtFecha.MaxDate

   
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
   
   porstCOCta.Close
   porstTGAux.Close
   porstCoCCo.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstTGAux = Nothing
   Set porstCoCCo = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtDato(Index).SetFocus
   Case 2     ', 3
      txtDato(Index).SetFocus
'      mskDato(Index).SetFocus
   Case 3, 4
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim Fecha As Variant
  Dim cCadReporte As String

  ppHabilitacion False
  cCadReporte = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmpRptCCoAtg_') DROP TABLE #tmpRptCCoAtg"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpRptCCoAtg", cCadReporte)
  
  cCadReporte = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE tmpRptCCoAtg ", "")
  cCadReporte = cCadReporte & "SELECT a.CodCCo, " & Choose(gsIdioma, "e.DetCCo", "e.DetCCox") & " AS DetCCo, a.CodAux, a.CodCta, c.AbvTDc, a.CodTDc, a.SerDoc, a.NroDoc, b.RazAux, " & Choose(gsIdioma, "d.DetCta", "d.DetCtax") & " AS DetCta, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN (CASE WHEN '" & cboTpoMon.ListIndex & "'=0 THEN a.ImpMN ELSE a.ImpME END) ELSE 0 END)), 0), 2) AS Debe, "
  cCadReporte = cCadReporte & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN (CASE WHEN '" & cboTpoMon.ListIndex & "'=0 THEN a.ImpMN ELSE a.ImpME END) ELSE 0 END)), 0), 2) AS Haber "
  cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpRptCCoAtg ")
  cCadReporte = cCadReporte & "FROM ((((COCpbDet a "
  cCadReporte = cCadReporte & "LEFT JOIN TgAux b ON a.codemp=b.codemp AND a.CodAux=b.CodAux) "
  cCadReporte = cCadReporte & "LEFT JOIN TgTDc c ON a.codemp=c.codemp AND a.CodTDc=c.CodTDc) "
  cCadReporte = cCadReporte & "LEFT JOIN CoCta d ON a.codemp=d.codemp AND a.pdoano=d.pdoano AND a.CodCta=d.CodCta) "
  cCadReporte = cCadReporte & "LEFT JOIN CoCCo e ON a.codemp=e.codemp AND a.pdoano=e.pdoano AND a.CodCCo=e.CodCCo) "
  cCadReporte = cCadReporte & "WHERE a.codemp='" & gsCodEmp & "' "
  cCadReporte = cCadReporte & "AND a.pdoano='" & gsAnoAct & "' "
  cCadReporte = cCadReporte & "AND a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
  cCadReporte = cCadReporte & "AND a.CodCCo BETWEEN '" & txtDato(3).Text & "' AND '" & txtDato(4).Text & "' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodTDc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.SerDoc, '') <>'' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.NroDoc, '') <>'' AND d.inddoc='1' "
  If Trim(txtDato(2).Text) <> "" Then
    cCadReporte = cCadReporte & "AND a.CodAux='" & txtDato(2).Text & "' "
  End If
  cCadReporte = cCadReporte & "GROUP BY a.CodCCo, a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, e.detcco, d.detcta, b.razaux, c.abvtdc "
  If ps_Plataforma = pSrvMySql Then
    cCadReporte = cCadReporte & "HAVING (ROUND(Debe - Haber, 2) <> 0.00) "
  Else
    cCadReporte = cCadReporte & "HAVING (ROUND(ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN (CASE WHEN '" & cboTpoMon.ListIndex & "'=0 THEN a.ImpMN ELSE a.ImpME END) ELSE 0 END)), 0), 2) - "
    cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN (CASE WHEN '" & cboTpoMon.ListIndex & "'=0 THEN a.ImpMN ELSE a.ImpME END) ELSE 0 END)), 0), 2), 2) <> 0.00) "
  End If
  cCadReporte = cCadReporte & "ORDER BY a.CodCCo, a.CodAux, a.CodCta, a.CodTDc, a.SerDoc, a.NroDoc"
  pocnnMain.Execute cCadReporte
  
  Fecha = Format(TxtFecha.Value, "yyyy-mm-dd")
  
  If OptTipo(1).Value = True Then
    cCadReporte = "SELECT MIN(b.FeEDoc) AS FeEDoc, MIN(b.FeVDoc) AS FeVDoc, a.CodAux, a.CodCta, Space(2) AS CodTDc, a.RazAux, a.DetCta, a.CodCCo, a.DetCCo, Space(18) AS cDocume, "
    If ps_Plataforma = pSrvMySql Then
      cCadReporte = cCadReporte & "ROUND(IFNULL(SUM((CASE WHEN DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=30 THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo30, "
      cCadReporte = cCadReporte & "ROUND(IFNULL(SUM((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>30 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=60) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo60, "
      cCadReporte = cCadReporte & "ROUND(IFNULL(SUM((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>60 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=90) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo90, "
      cCadReporte = cCadReporte & "ROUND(IFNULL(SUM((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>90 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=120) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo120, "
      cCadReporte = cCadReporte & "ROUND(IFNULL(SUM((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>120 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=360) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo360, "
      cCadReporte = cCadReporte & "ROUND(IFNULL(SUM((CASE WHEN DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>360 THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldoMas "
    ElseIf ps_Plataforma = pSrvSql Then
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE WHEN DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=30 THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo30, "
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>30 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=60) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo60, "
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>60 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=90) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo90, "
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>90 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=120) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo120, "
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>120 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=360) THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldo360, "
      cCadReporte = cCadReporte & "ROUND(ISNULL(SUM((CASE WHEN DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>360 THEN (a.Debe - a.Haber) ELSE 0 END)), 0), 2) AS cSaldoMas "
    End If
  Else
    cCadReporte = "SELECT b.FeEDoc, b.FeVDoc, a.CodAux, a.CodCta, a.CodTDc, a.RazAux, a.DetCta, a.CodCCo, a.DetCCo, "
    cCadReporte = cCadReporte & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.AbvTDc, '-', a.SerDoc, '-', a.NroDoc)", "(a.AbvTDc+'-'+a.SerDoc+'-'+a.NroDoc)") & " AS cDocume, "
    If ps_Plataforma = pSrvMySql Then
      cCadReporte = cCadReporte & "ROUND((CASE WHEN DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=30 THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo30, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>30 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=60) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo60, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>60 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=90) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo90, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>90 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=120) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo120, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>120 AND DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))<=360) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo360, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN DATEDIFF(DATE_FORMAT('" & Fecha & "', '%Y-%m-%d'), DATE_FORMAT(b.FeVDoc, '%Y-%m-%d'))>360 THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldoMas "
    ElseIf ps_Plataforma = pSrvSql Then
      cCadReporte = cCadReporte & "ROUND((CASE WHEN DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=30 THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo30, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>30 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=60) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo60, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>60 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=90) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo90, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>90 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=120) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo120, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN (DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>120 AND DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))<=360) THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldo360, "
      cCadReporte = cCadReporte & "ROUND((CASE WHEN DATEDIFF(dd, CONVERT(smalldatetime, b.FeVDoc, 120), CONVERT(smalldatetime, '" & Fecha & "', 120))>360 THEN (a.Debe - a.Haber) ELSE 0 END), 2) AS cSaldoMas "
    End If
  End If
  cCadReporte = cCadReporte & "FROM (" & ps_Prefijo & "tmpRptCCoAtg a "
  cCadReporte = cCadReporte & "LEFT JOIN COCpbDet b ON a.CodCCo=b.CodCCo AND a.CodAux=b.CodAux AND a.CodCta=b.CodCta AND a.CodTDc=b.CodTDc AND a.SerDoc=b.SerDoc AND a.NroDoc=b.NroDoc) "
  cCadReporte = cCadReporte & "WHERE b.codemp='" & gsCodEmp & "' "
  cCadReporte = cCadReporte & "AND b.pdoano='" & gsAnoAct & "' "
  cCadReporte = cCadReporte & "AND b.TpoPvs='" & TPOPVS_PVS & "' "
  cCadReporte = cCadReporte & "AND " & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(b.FeVDoc, '%Y-%m-%d')<='" & Fecha & "'", "CONVERT(smalldatetime, b.FeVDoc, 103)<=CONVERT(smalldatetime, '" & Format(TxtFecha.Value, "dd/mm/yyyy") & "', 103)") & " "
  If OptTipo(1).Value = True Then
    cCadReporte = cCadReporte & "GROUP BY a.CodCCo, a.CodCta, a.CodAux, a.DetCCo, a.DetCta, a.RazAux "
    If ps_Plataforma = pSrvMySql Then
      cCadReporte = cCadReporte & "HAVING (ROUND(cSaldo30 + cSaldo60 + cSaldo90 + cSaldo120 + cSaldoMas, 2) <> 0.00) "
    Else
      cCadReporte = cCadReporte & "HAVING (ROUND(ISNULL(SUM(a.Debe - a.Haber), 0), 2) <> 0.00) "
    End If
  End If
  cCadReporte = cCadReporte & "ORDER BY a.CodCCo, a.CodCta, a.CodAux" & IIf(OptTipo(0).Value, ", a.CodTDc, a.SerDoc, a.NroDoc", "")
  
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = cCadReporte
    .Open
  End With
    
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " (" & IIf(OptTipo(0).Value = True, Choose(gsIdioma, "Detalle", "Detail") & " / " & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0), Choose(gsIdioma, "Resumen", "Summary") & " / " & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0)) & ")", udFecha, True, chkImpFecha.Value, porstMRp
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptRCCtCCoDet.rpt", "rptRCCtCCoRes.rpt")
      .WindowShowExportBtn = IIf(paOpciones(2), True, False)
      .MarginLeft = unMargenIzquierdo
      .WindowState = crptMaximized
      .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
      .Action = 1
    End With
  Else
    Set MRViewer = New MRViewerObject
    With MRViewer
      .DataRecordSet = porstMRp
      If OptTipo(0).Value = True Then
        .LoadReport gsRutRpt & "rptRCCtCCoDet.mrp"
      Else
        .LoadReport gsRutRpt & "rptRCCtCCoRes.mrp"
      End If
      Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(OptTipo(0).Value = True, Choose(gsIdioma, "Detalle", "Detail") & " / " & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0), Choose(gsIdioma, "Resumen", "Summary") & " / " & IIf(cboTpoMon.ListIndex = 0, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0)) & ")", udFecha, True, chkImpFecha.Value)
      
      '[Parámetros adicionales.
      .Parameters("pPeriodoAdc") = "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
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
   If txtDato(2) = "" Then
       lblDatoDeta(2).Caption = ""
       'lblDatoDeta(0).Caption = "Todos..."
   End If

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
   Case 2    ', 1                           'Cambiar (añadir índices). - Auxiliar
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   Case 3, 4                           'Cambiar (añadir índices).
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
   Case 2         ', 1                           'Cambiar (añadir índices).
      modAyuBus.Aux_Det "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 3, 4                           'Cambiar (añadir índices).
      modAyuBus.CCo_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
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
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetCta), "", !DetCta)
         End If
      End With
   
   Case 2   ', 1  - Auxiliar
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
            lblDatoDeta(tnIndex).Caption = " " & !RazAux
         End If
      End With
   Case 3, 4
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With porstCoCCo
         .MoveFirst
         .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetCCo), "", !DetCCo)
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

Private Sub TxtFecha_LostFocus()

    If Month(TxtFecha.Value) <> gsMesAct Or Year(TxtFecha.Value) <> gsAnoAct Then
        MsgBox "Fecha Fuera de Periodo Actual.", vbInformation, Me.Caption
        TxtFecha.SetFocus
    End If

End Sub


