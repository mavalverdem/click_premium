VERSION 5.00
Begin VB.Form frmRMayAux 
   Caption         =   "[título]"
   ClientHeight    =   3330
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Meses"
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   0
      TabIndex        =   20
      Top             =   1440
      Width           =   2655
      Begin VB.TextBox txtDato 
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
         Index           =   3
         Left            =   1860
         TabIndex        =   7
         Top             =   280
         Width           =   375
      End
      Begin VB.TextBox txtDato 
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
         Index           =   2
         Left            =   660
         TabIndex        =   6
         Top             =   280
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "a"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   22
         Top             =   345
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   345
         Width           =   210
      End
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   19
      Top             =   2040
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmRMayAux.frx":0000
      Left            =   5640
      List            =   "frmRMayAux.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1515
      Width           =   1350
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H80000002&
      Height          =   1275
      Left            =   0
      TabIndex        =   13
      Top             =   75
      Width           =   6990
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6615
         Picture         =   "frmRMayAux.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6615
         Picture         =   "frmRMayAux.frx":01AE
         Style           =   1  'Graphical
         TabIndex        =   14
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
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2790
      Width           =   7005
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
         Picture         =   "frmRMayAux.frx":0358
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
         Picture         =   "frmRMayAux.frx":04A2
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
         Picture         =   "frmRMayAux.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label Label18 
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
      Left            =   4965
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmRMayAux"
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

Private Sub Form_Load()
   On Error GoTo Err
  
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
      .Source = "SELECT CodCta, DetCta " _
              & "FROM COCta " _
              & "ORDER BY CodCta"
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
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !CodCta
      .MoveFirst
      txtDato(0).Text = !CodCta
   End With
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
    txtDato(2).Text = gsMesAct
    txtDato(3).Text = gsMesAct
    txtDato(2).MaxLength = 2
    txtDato(3).MaxLength = 2
  
  'Otros.
   cboTpoMon.ListIndex = IIf(gsTpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
   
   
'   With TxtFecha
'      For dnContador = 0 To 1
'        .Item(dnContador).MinDate = "01/01/" & gsAnoAct
'        .Item(dnContador).MaxDate = "31/12/" & gsAnoAct
'      Next
'   End With
'   TxtFecha(0).Value = "01/" & gsMesAct & "/" & gsAnoAct
'   TxtFecha(1).Value = gfUltDia(TxtFecha(0).Value)
   
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
   Dim dnContador As Byte, sMoneda As String
    
   ppHabilitacion False
   
   sMoneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0)
   With porstMRp
      If .State = adStateOpen Then .Close
      .Source = "SELECT a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.NroIte, a.FehOpe,"
      .Source = .Source & "  CONCAT(e.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) AS cDocume,"
      .Source = .Source & "  a.CodAux, b.RazAux, a.RefDoc, a.GloIte,"
      .Source = .Source & " IF(a.TpoCtb='" & TPOCTB_DEB & "', Imp" & sMoneda & ", 0) AS cDebe,"
      .Source = .Source & " IF(a.TpoCtb='" & TPOCTB_HAB & "', Imp" & sMoneda & ", 0) AS cHaber,"
      .Source = .Source & "  c.DetCta , d.DetDro, e.AbvTDc,"
      If txtDato(2) <> "00" Then
      .Source = .Source _
              & "  (" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ") AS cAntCtaDeb," _
              & "  (" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab"
      Else
      .Source = .Source _
              & "  0 AS cAntCtaDeb, 0 AS cAntCtaHab"
      End If
      .Source = .Source & " FROM ((((COCpbDet a"
      .Source = .Source & "  LEFT JOIN TGAux b ON a.CodAux=b.CodAux)"
      .Source = .Source & "  LEFT JOIN COCta c ON a.CodCta=c.CodCta)"
      .Source = .Source & "  LEFT JOIN CODro d ON a.CodDro=d.CodDro)"
      .Source = .Source & "  LEFT JOIN TGTDc e ON a.CodTDc=e.CodTDc)"
      .Source = .Source & "  LEFT JOIN COCtaAcu ON a.CodCta=COCtaAcu.CodCta"
      .Source = .Source & " WHERE a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "'"
      .Source = .Source & " AND a.MesPvs>='" & txtDato(2).Text & "' AND a.MesPvs<='" & txtDato(3).Text & "'"
      If txtDato(2) <> "00" Then
        .Source = .Source & "UNION"
        .Source = .Source & " SELECT '00', c.CodCta, '', '', '', NULL, '', '', '', '', '', 0, 0,"
        .Source = .Source & "  c.DetCta , '', '',"
        .Source = .Source & " ROUND(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 1, 2)) & ", 2) AS cAntCtaDeb,"
        .Source = .Source & " ROUND(" & gsAcuAnt(IIf(cboTpoMon.ListIndex = 0, 3, 4)) & ") AS cAntCtaHab"
        .Source = .Source & " FROM (COCta c"
        .Source = .Source & "  LEFT JOIN COCpbDet a ON c.CodCta=a.CodCta)"
        .Source = .Source & "  LEFT JOIN COCtaAcu ON c.CodCta=COCtaAcu.CodCta"
        .Source = .Source & " WHERE c.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "'"
        .Source = .Source & " AND a.MesPvs>='" & txtDato(2).Text & "' AND a.MesPvs<='" & txtDato(3).Text & "'"
        .Source = .Source & " AND c.TpoCta='" & TPOCTA_TRA & "'"
        .Source = .Source & " HAVING ROUND(cAntCtaDeb-cAntCtaHab, 2)<>0.00"
      End If
      .Source = .Source & " ORDER BY a.MesPvs, a.CodCta, a.CodDro, a.NroCpb, a.NroIte"
      .Open
    '15/12/2003 Angel
   End With

   usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
   If usDEstino = PRN_DEST_GRAF Then
      'Genero la tabla temporal de impresion
      pocnnMain.Execute "DROP TABLE IF EXISTS trptRMayAux"
      cmdImprimir(Index).Tag = "CREATE TABLE IF NOT EXISTS trptRMayAux " & porstMRp.Source
      pocnnMain.Execute cmdImprimir(Index).Tag

      Call gpEncabezadoRpt(frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True)
        
        If porstMRp.RecordCount = 0 Then
           pocnnMain.Execute "INSERT INTO trptRMayAux (CODAUX) VALUES ('')"
        End If
      With frmMain.rptMain
         '[Datos y parámetros del reporte.  'Cambiar.
        '    .WindowShowGroupTree = True
          'Fórmular propias.
        '         .Formulas(6) = "fSaldoInicial=" & gsIniMesCnt
         ']
         .ReportFileName = gsRutRpt & "rptRMayAux.rpt"
         .Connect = "Provider=MySqlProv;Extended Properties=" & CONNSTRG & gsNomBDS
         .WindowShowExportBtn = IIf(paOpciones(2), True, False)
         .WindowState = crptMaximized
         .WindowShowRefreshBtn = False
         .MarginLeft = unMargenIzquierdo
         .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
         .Action = 1
      End With
        ' Elimino la tabla temporal de impresion
      pocnnMain.Execute "DROP TABLE IF EXISTS trptRMayAux"
   Else
      Set MRViewer = New MRViewerObject

      With MRViewer
         .DataRecordSet = porstMRp
         .LoadReport gsRutRpt & "rptRMayAux.mrp"

         Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & ")", udFecha, True)
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
            lblDatoDeta(tnIndex).Caption = " " & !DetCta
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

