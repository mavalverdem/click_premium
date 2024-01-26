VERSION 5.00
Begin VB.Form frmLCon 
   Caption         =   "[título]"
   ClientHeight    =   4755
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRango 
      Caption         =   "Rango de Auxiliares "
      ForeColor       =   &H00800000&
      Height          =   1020
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   7035
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   585
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   1
         Left            =   6660
         Picture         =   "frmLCon.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1260
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   0
         Left            =   6660
         Picture         =   "frmLCon.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   270
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
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   4
         Top             =   585
         Width           =   5265
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
         Left            =   1380
         TabIndex        =   2
         Top             =   270
         Width           =   5265
      End
   End
   Begin VB.Frame fraRango 
      Caption         =   " Rango de Periodos "
      ForeColor       =   &H00800000&
      Height          =   1095
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1095
      Width           =   4215
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   2310
         TabIndex        =   11
         Text            =   "Mes Final"
         Top             =   645
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   2310
         TabIndex        =   8
         Text            =   "Mes Inicio"
         Top             =   300
         Width           =   1710
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   855
         TabIndex        =   10
         Text            =   "Año Final"
         Top             =   645
         Width           =   1245
      End
      Begin VB.ComboBox cmbPeriodo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   7
         Text            =   "Año Inicio"
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   345
         Width           =   765
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Fin :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   690
         Width           =   765
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   3450
      Width           =   2295
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   19
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4860
      TabIndex        =   20
      Top             =   3450
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1035
         TabIndex        =   22
         Top             =   315
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.Frame fraRango 
      Caption         =   "Rango de Servicios "
      ForeColor       =   &H00800000&
      Height          =   1020
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   2325
      Width           =   7035
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   2
         Left            =   6660
         Picture         =   "frmLCon.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   270
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   1600
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   3
         Left            =   6660
         Picture         =   "frmLCon.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   585
         Width           =   1600
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
         Index           =   2
         Left            =   1725
         TabIndex        =   14
         Top             =   270
         Width           =   4920
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
         Index           =   3
         Left            =   1725
         TabIndex        =   16
         Top             =   585
         Width           =   4920
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
      ScaleWidth      =   7080
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4215
      Width           =   7080
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
         TabIndex        =   25
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
         Picture         =   "frmLCon.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Picture         =   "frmLCon.frx":07F2
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Picture         =   "frmLCon.frx":0D24
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmLCon"
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
Private porstConSer As ADODB.Recordset
Private porstTGAux As ADODB.Recordset
Private plRecupera As Boolean
']

Private Sub cmbPeriodo_Click(Index As Integer)
  ppRecuperaInformacion 1, Index
End Sub

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstConSer = New ADODB.Recordset
   Set porstTGAux = New ADODB.Recordset
   
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
    .Source = "SELECT codaux, razaux "
    .Source = .Source & "FROM tgaux "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND indcli=1 "
    .Source = .Source & "ORDER BY codaux "
  '  .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
   
  With porstConSer
    .ActiveConnection = pocnnMain
    .Source = "SELECT codcon, " & Choose(gsIdioma, "detcon", "deconx") & " AS detcon "
    .Source = .Source & "FROM coconser "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND mespvs='" & gsMesAct & "' "
    .Source = .Source & "ORDER BY codcon"
'     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
 ']

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(2, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Inicio :", "Fin :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Beginning :", "End :")
  Next nElemento
  fraRango(0).Caption = Choose(gsIdioma, "Rango de Auxiliares", "Range of Auxiliarys")
  fraRango(1).Caption = Choose(gsIdioma, "Rango de Periodos", "Range of Periods")
  fraRango(2).Caption = Choose(gsIdioma, "Rango de Servicios", "Range of Services")
  fraTipo.Caption = Choose(gsIdioma, "Tipo", "Type")
  OptTipo(0).Caption = Choose(gsIdioma, "Detalle", "Detail")
  OptTipo(1).Caption = Choose(gsIdioma, "Resumen", "Summary")
  fraTipoImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  optTipoImpresion(0).Caption = Choose(gsIdioma, "Matricial", "Dot Matrix")
  optTipoImpresion(1).Caption = Choose(gsIdioma, "Gráfica", "Graphic")
  CaptionBotones Me, False, False, False, False, False, False, True, True, True, False, False, False, True, aLabel
   
  'Límites de rangos.
  plRecupera = False
  txtDato(0).MaxLength = porstTGAux.Fields("codaux").DefinedSize
  txtDato(1).MaxLength = porstTGAux.Fields("codaux").DefinedSize
  With porstTGAux
    .MoveLast
    txtDato(1).Text = !codaux
    .MoveFirst
    txtDato(0).Text = !codaux
  End With
  txtDato(0).Tag = txtDato(0).Text
  txtDato(1).Tag = txtDato(1).Text
  'Busca detalle de códigos            '(habilitar/deshabilitar).
  If txtDato(0).Text <> "" Then ppAyuDet 0
  If txtDato(1).Text <> "" Then ppAyuDet 1
   
 '[Datos predeterminados.              'Cambiar.
  ' Configuro los controles de año y mes
  For dnContador = (Val(gsAnoAct) - 9) To Val(gsAnoAct)
    cmbPeriodo(0).AddItem Choose(gsIdioma, "Año ", "Year ") & dnContador
    cmbPeriodo(1).AddItem Choose(gsIdioma, "Año ", "Year ") & dnContador
  Next dnContador
  cmbPeriodo(0).ListIndex = 9
  cmbPeriodo(1).ListIndex = 9
  
  For dnContador = 0 To 13
    If gsIdioma = NvlUsr_Sup Then
      cmbPeriodo(2).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
      cmbPeriodo(3).AddItem Choose(dnContador + 1, "Apertura", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre", "Cierre")
    Else
      cmbPeriodo(2).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
      cmbPeriodo(3).AddItem Choose(dnContador + 1, "Opening", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Closing")
    End If
  Next dnContador
  cmbPeriodo(2).ListIndex = Val(gsMesAct)
  cmbPeriodo(3).ListIndex = Val(gsMesAct)
   
  txtDato(2).MaxLength = porstConSer.Fields("codcon").DefinedSize
  txtDato(3).MaxLength = porstConSer.Fields("codcon").DefinedSize
  With porstConSer
    .MoveLast
    txtDato(3).Text = !codcon
    .MoveFirst
    txtDato(2).Text = !codcon
  End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
  If txtDato(2).Text <> "" Then ppAyuDet 2
  If txtDato(3).Text <> "" Then ppAyuDet 3
  
  'Otros.
  OptTipo(0).Value = True
  plRecupera = True
  
  'Características de impresión.
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
   porstConSer.Close
   pocnnMain.Close
   Set porstConSer = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2, 3
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
  Dim sSentencia As String, s_Sentencia As String
  Dim s_AnoIni As String, s_AnoFin As String
  Dim s_MesIni As String, s_MesFin As String
  
  s_AnoIni = Right(cmbPeriodo(0).Text, 4)
  s_AnoFin = Right(cmbPeriodo(1).Text, 4)
  s_MesIni = Format(cmbPeriodo(2).ListIndex, "00")
  s_MesFin = Format(cmbPeriodo(3).ListIndex, "00")
  If Not (s_AnoFin >= s_AnoIni) Then MsgBox Choose(gsIdioma, "Ejercicio Final debe ser mayor o igual que Inicial; Verificar", "End Fiscal year must be equal or more than opening; Verify"), vbExclamation: cmbPeriodo(1).SetFocus: Exit Sub
  If (s_AnoFin = s_AnoIni) And Not (s_MesIni <= s_MesFin) Then MsgBox Choose(gsIdioma, "Mes Final debe ser mayor o igual que Inicial de Saldos", "End month must be equal or more than opening balance"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  If (s_AnoFin = gsAnoAct) And Not (s_MesFin <= gsMesAct) Then MsgBox Choose(gsIdioma, "Mes Final debe ser menor o igual que Mes Activo", "End month must be smaller or just as Active Month"), vbExclamation: cmbPeriodo(3).SetFocus: Exit Sub
  
  ppHabilitacion False
    
  ' Elimino la tabla temporal de saldos
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmprptlstcon_') DROP TABLE #tmprptlstcon"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmprptlstcon", sSentencia)
  
  ' Genero temporal ejecutado
  sSentencia = "CREATE TABLE tmprptlstcon "
  If OptTipo(0).Value Then
    ' Información detalle comprobante
    sSentencia = sSentencia & "SELECT det.codemp, det.codcon, " & Choose(gsIdioma, "ser.detcon", "ser.detconx") & " AS detcon, ser.codaux, ser.fehcon, "
    sSentencia = sSentencia & "(CASE ser.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS ctpomon, "
    sSentencia = sSentencia & "(CASE ser.tpomon WHEN '" & TPOMON_NAC & "' THEN ser.impmn ELSE ser.impme END) AS nImporte, "
    sSentencia = sSentencia & "det.pdoano, det.mespvs, det.coddro, det.nrocpb, tdc.abvtdc, det.serdoc, det.nrodoc, det.feedoc, det.refdoc, "
    sSentencia = sSentencia & Choose(gsIdioma, "det.gloite", "det.gloitex") & " AS gloite, "
    sSentencia = sSentencia & "ROUND(IFNULL(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN (CASE ser.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impmn ELSE det.impme END) ELSE 0 END, 0), 2) AS impdeb, "
    sSentencia = sSentencia & "ROUND(IFNULL(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN (CASE ser.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impmn ELSE det.impme END) ELSE 0 END, 0), 2) AS imphab, "
    sSentencia = sSentencia & "det.codcta, (CASE det.tpognr WHEN '" & TPOGNR_VTA & "' THEN '1' ELSE '2' END) AS sOrden "
    sSentencia = sSentencia & "FROM ((cocpbdet det "
    sSentencia = sSentencia & "INNER JOIN coconser ser ON ser.codemp=det.codemp AND ser.codcon=det.codcon AND ser.codaux BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' ) "
    sSentencia = sSentencia & "LEFT JOIN tgtdc tdc ON tdc.codemp=det.codemp AND tdc.codtdc=det.codtdc) "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & ">='" & s_AnoIni & s_MesIni & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & "<='" & s_AnoFin & s_MesFin & "' "
    sSentencia = sSentencia & "AND det.codcon BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
    sSentencia = sSentencia & "ORDER BY codcon, pdoano, mespvs, coddro, nrocpb"
    ' Información reporte final
    s_Sentencia = "SELECT sal.codcon, sal.detcon, sal.codaux, sal.fehcon, sal.ctpomon, sal.nimporte, aux.razaux, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(sal.pdoano, '-', sal.mespvs)", "(sal.pdoano+'-'+sal.mespvs)") & " AS speriodo, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(sal.coddro, '-', sal.nrocpb)", "(sal.coddro+'-'+sal.nrocpb)") & " AS snrocpb, "
    s_Sentencia = s_Sentencia & IIf(ps_Plataforma = pSrvMySql, "CONCAT(sal.abvtdc, '-', sal.serdoc, '-', sal.nrodoc)", "(sal.abvtdc+'-'+sal.serdoc+'-'+sal.nrodoc)") & " AS sdocumento, "
    s_Sentencia = s_Sentencia & "sal.feedoc, sal.refdoc, sal.gloite, "
    s_Sentencia = s_Sentencia & "ROUND(IFNULL(CASE sal.sorden WHEN '1' THEN sal.imphab-sal.impdeb ELSE 0 END, 0), 2) AS impfactura, "
    s_Sentencia = s_Sentencia & "ROUND(IFNULL(CASE sal.sorden WHEN '2' THEN sal.impdeb-sal.imphab ELSE 0 END, 0), 2) AS impconsumo, "
    s_Sentencia = s_Sentencia & "sal.codcta, sal.sorden "
    s_Sentencia = s_Sentencia & "FROM " & ps_Prefijo & "tmprptlstcon sal "
    s_Sentencia = s_Sentencia & "LEFT JOIN tgaux aux ON aux.codemp=sal.codemp AND aux.codaux=sal.codaux "
    s_Sentencia = s_Sentencia & "UNION ALL "
    s_Sentencia = s_Sentencia & "SELECT ser.codcon, "
    s_Sentencia = s_Sentencia & Choose(gsIdioma, "ser.detcon", "ser.detconx") & " AS detcon, ser.codaux, ser.fehcon, "
    s_Sentencia = s_Sentencia & "(CASE ser.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS ctpomon, "
    s_Sentencia = s_Sentencia & "(CASE ser.tpomon WHEN '" & TPOMON_NAC & "' THEN ser.impmn ELSE ser.impme END) AS nimporte, "
    s_Sentencia = s_Sentencia & "aux.razaux, "
    s_Sentencia = s_Sentencia & "Null AS speriodo, Null AS snrocpb, Null AS sdocumento, Null AS feedoc, Null AS refdoc, Null AS gloite, "
    s_Sentencia = s_Sentencia & "0.00 AS impfactura, 0.00 AS impconsumo, Null AS codcta, '0' AS sorden "
    s_Sentencia = s_Sentencia & "FROM coconser ser "
    s_Sentencia = s_Sentencia & "LEFT JOIN tgaux aux ON ser.codemp=aux.codemp AND ser.codaux=aux.codaux "
    s_Sentencia = s_Sentencia & "WHERE ser.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(ser.pdoano, ser.mespvs)", "(ser.pdoano+ser.mespvs)") & ">='" & s_AnoIni & s_MesIni & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(ser.pdoano, ser.mespvs)", "(ser.pdoano+ser.mespvs)") & "<='" & s_AnoFin & s_MesFin & "' "
    s_Sentencia = s_Sentencia & "AND ser.codaux BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    s_Sentencia = s_Sentencia & "AND ser.codcon BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
    s_Sentencia = s_Sentencia & "AND NOT EXISTS (SELECT * FROM " & ps_Prefijo & "tmprptlstcon sal WHERE sal.codemp=ser.codemp AND sal.codaux=ser.codaux AND sal.codcon=ser.codcon) "
    s_Sentencia = s_Sentencia & "ORDER BY codcon, sorden, speriodo, snrocpb"
  Else
    ' Información detalle comprobante
    sSentencia = sSentencia & "SELECT det.codemp, det.pdocpr, det.codcta, "
    sSentencia = sSentencia & "ROUND(SUM(IFNULL(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN (CASE pdo.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impmn ELSE det.impme END) ELSE 0 END, 0)), 2) AS impdeb, "
    sSentencia = sSentencia & "ROUND(SUM(IFNULL(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN (CASE pdo.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impmn ELSE det.impme END) ELSE 0 END, 0)), 2) AS imphab "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmprptlstpdo ")
    sSentencia = sSentencia & "FROM cocpbdet det "
    sSentencia = sSentencia & "INNER JOIN copdocprcta cta ON cta.codemp=det.codemp AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cta.coddpe, cta.pdocpr)", "(cta.coddpe+cta.pdocpr)") & "=det.pdocpr AND cta.codcta=det.codcta AND cta.codcco=det.codcco "
    sSentencia = sSentencia & "INNER JOIN copdocpr pdo ON pdo.codemp=cta.codemp AND pdo.pdoano=cta.pdoano AND pdo.mespvs=cta.mespvs AND pdo.coddpe=cta.coddpe AND pdo.pdocpr=cta.pdocpr "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & ">='" & s_AnoIni & s_MesIni & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.pdoano, det.mespvs)", "(det.pdoano+det.mespvs)") & "<='" & s_AnoFin & s_MesFin & "' "
    sSentencia = sSentencia & "AND det.pdocpr BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
    sSentencia = sSentencia & "GROUP BY det.pdocpr, det.codcta "
    sSentencia = sSentencia & "ORDER BY pdocpr, codcta"
    
    ' Información reporte final
    s_Sentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(pdo.coddpe,pdo.pdocpr)", "(pdo.coddpe+pdo.pdocpr)") & " AS pdocpr, "
    s_Sentencia = s_Sentencia & Choose(gsIdioma, "pdo.detpdo", "pdo.detpdox") & " AS detpdo, pdo.codaux, pdo.fehpdo, cta.codcta, cta.codcco, "
    s_Sentencia = s_Sentencia & "(CASE pdo.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS ctpomon, "
    s_Sentencia = s_Sentencia & "(CASE pdo.tpomon WHEN '" & TPOMON_NAC & "' THEN pdo.impmn ELSE pdo.impme END) AS nimportepdo, "
    s_Sentencia = s_Sentencia & "(CASE pdo.tpomon WHEN '" & TPOMON_NAC & "' THEN cta.impcta_mn ELSE cta.impcta_me END) AS nimportecta, "
    s_Sentencia = s_Sentencia & "ROUND(IFNULL(sal.impdeb-sal.imphab, 0), 2) AS impconsumocta, "
    s_Sentencia = s_Sentencia & "pdo.impdife, aux.razaux "
    s_Sentencia = s_Sentencia & "FROM copdocprcta cta "
    s_Sentencia = s_Sentencia & "INNER JOIN copdocpr pdo ON pdo.codemp=cta.codemp AND pdo.pdoano=cta.pdoano AND pdo.mespvs=cta.mespvs AND pdo.coddpe=cta.coddpe AND pdo.pdocpr=cta.pdocpr "
    s_Sentencia = s_Sentencia & "LEFT JOIN " & ps_Prefijo & "tmprptlstpdo sal ON sal.codemp=cta.codemp AND sal.pdocpr=" & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cta.coddpe, cta.pdocpr)", "(cta.coddpe+cta.pdocpr)") & " AND sal.codcta=cta.codcta "
    s_Sentencia = s_Sentencia & "LEFT JOIN tgaux aux ON pdo.codemp=aux.codemp AND pdo.codaux=aux.codaux "
    s_Sentencia = s_Sentencia & "WHERE cta.codemp='" & gsCodEmp & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cta.pdoano, cta.mespvs)", "(cta.pdoano+cta.mespvs)") & ">='" & s_AnoIni & s_MesIni & "' "
    s_Sentencia = s_Sentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(cta.pdoano, cta.mespvs)", "(cta.pdoano+cta.mespvs)") & "<='" & s_AnoFin & s_MesFin & "' "
    s_Sentencia = s_Sentencia & "AND concat(cta.coddpe,cta.pdocpr) BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "' "
    s_Sentencia = s_Sentencia & "ORDER BY pdocpr, codcta"
  End If
  pocnnMain.Execute sSentencia
  
  With porstMRp
    If .State = adStateOpen Then .Close
    .Source = s_Sentencia
    .Open
   End With
   
  usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
  If usDEstino = PRN_DEST_GRAF Then
    gpEncabezadoRpt frmMain.rptMain, Me.Caption & " - " & IIf(OptTipo(0).Value, "Detalle", "Resumen"), udFecha, True, False, porstMRp
    
    With frmMain.rptMain
      '[Datos y parámetros del reporte.  'Cambiar.
      .ReportFileName = gsRutRpt & IIf(OptTipo(0).Value, "rptlcondet.rpt", "rptlcon.rpt")
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
      .LoadReport gsRutRpt & "rptLPdo.mrp"
      
      Call gpEncabezadoMRp(MRViewer, Me.Caption, udFecha, True)
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
  ' Elimino la tabla temporal de saldos
  sSentencia = "IF EXISTS (SELECT * FROM tempdb.information_schema.tables WHERE LEFT(table_name, 14)='#tmprptlstcon_') DROP TABLE #tmprptlstcon"
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmprptlstcon", sSentencia)
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

Private Sub txtDato_LostFocus(Index As Integer)
  If Index <= 1 Then ppRecuperaInformacion 0, Index
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2, 3                         'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Dim sWhere As String
  Select Case tnIndex
   Case 0, 1                          'Cambiar (añadir índices).
    modAyuBus.Aux_Det "IndCli=1", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRango(0).Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRango(0).Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 2, 3                           'Cambiar (añadir índices).
    sWhere = IIf(ps_Plataforma = pSrvMySql, "CONCAT(pdoano, mespvs)", "(pdoano+mespvs)") & ">='" & Right(cmbPeriodo(0).Text, 4) & Format(cmbPeriodo(2).ListIndex, "00") & "' "
    sWhere = sWhere & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(pdoano, mespvs)", "(pdoano+mespvs)") & "<='" & Right(cmbPeriodo(1).Text, 4) & Format(cmbPeriodo(3).ListIndex, "00") & "' "
    sWhere = sWhere & "AND codaux BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "'"
    modAyuBus.Con_Rpt sWhere, txtDato(tnIndex).Text, 0, 0, Me.Top + fraRango(2).Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRango(2).Left + txtDato(tnIndex).Left
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
    With porstTGAux
      If .RecordCount > 0 Then .MoveFirst
      .Find "codaux='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!RazAux), "", !RazAux)
      End If
    End With
   Case 2, 3
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstConSer
      .MoveFirst
      .Find "codcon ='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcon), "", !detcon)
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

End Sub
Private Sub ppRecuperaInformacion(ByVal nSecuencia As Integer, nIndex As Integer)
  
  If nSecuencia = 0 And (txtDato(0).Tag = txtDato(0).Text And txtDato(1).Tag = txtDato(1).Text) Then Exit Sub
  If nSecuencia = 1 And (cmbPeriodo(0).Tag = cmbPeriodo(0).Text And cmbPeriodo(1).Tag = cmbPeriodo(1).Text And cmbPeriodo(2).Tag = cmbPeriodo(2).Text And cmbPeriodo(3).Tag = cmbPeriodo(3).Text) Then Exit Sub
  If nSecuencia = 0 Then
    txtDato(nIndex).Tag = txtDato(nIndex).Text
  Else
    cmbPeriodo(nIndex).Tag = cmbPeriodo(nIndex).Text
  End If
  If Not plRecupera Then Exit Sub
  
  ' Información de pedidos
  With porstConSer
    If .State = adStateOpen Then .Close
    .Source = "SELECT codcon, " & Choose(gsIdioma, "detcon", "detconx") & " AS detcon "
    .Source = .Source & "FROM coconser "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(pdoano, mespvs)", "(pdoano+mespvs)") & ">='" & Right(cmbPeriodo(0).Text, 4) & Format(cmbPeriodo(2).ListIndex, "00") & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(pdoano, mespvs)", "(pdoano+mespvs)") & "<='" & Right(cmbPeriodo(1).Text, 4) & Format(cmbPeriodo(3).ListIndex, "00") & "' "
    .Source = .Source & "AND codaux BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' "
    .Source = .Source & "ORDER BY codcon"
    .Open
    .MoveLast
    txtDato(3).Text = !codcon
    .MoveFirst
    txtDato(2).Text = !codcon
  End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
  If txtDato(2).Text <> "" Then ppAyuDet 2
  If txtDato(3).Text <> "" Then ppAyuDet 3

End Sub

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   paOpciones = taOpciones
   cmdImprimir(0).Enabled = taOpciones(0)
   cmdImprimir(1).Enabled = taOpciones(1)
End Property

