VERSION 5.00
Begin VB.Form frmRMayAuxCCo 
   Caption         =   "[título]"
   ClientHeight    =   4635
   ClientLeft      =   1620
   ClientTop       =   1515
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6990
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkImpFecha 
      Caption         =   "Imprime Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4800
      TabIndex        =   22
      Top             =   3240
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   23
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      Left            =   5745
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2475
      Width           =   1260
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Rango"
      ForeColor       =   &H00800000&
      Height          =   2265
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   6990
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
         TabIndex        =   7
         Top             =   1845
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
         Index           =   2
         Left            =   135
         TabIndex        =   6
         Top             =   1485
         Width           =   630
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   4500
         Picture         =   "frmRMayAuxCCo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1500
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   4500
         Picture         =   "frmRMayAuxCCo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1860
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   6585
         Picture         =   "frmRMayAuxCCo.frx":0354
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
         Left            =   135
         TabIndex        =   4
         Top             =   495
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
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   945
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   6585
         Picture         =   "frmRMayAuxCCo.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   855
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
         Height          =   315
         Index           =   3
         Left            =   780
         TabIndex        =   21
         Top             =   1845
         Width           =   3720
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
         Left            =   780
         TabIndex        =   20
         Top             =   1500
         Width           =   3720
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
         Left            =   1065
         TabIndex        =   17
         Top             =   495
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
         Index           =   1
         Left            =   1065
         TabIndex        =   16
         Top             =   855
         Width           =   5520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centros de Costo"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   1260
         Width           =   1215
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
      ScaleWidth      =   6990
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4095
      Width           =   6990
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
         Picture         =   "frmRMayAuxCCo.frx":06A8
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
         Picture         =   "frmRMayAuxCCo.frx":07F2
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
         Picture         =   "frmRMayAuxCCo.frx":0D24
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Moneda"
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   5085
      TabIndex        =   12
      Top             =   2520
      Width           =   600
   End
End
Attribute VB_Name = "frmRMayAuxCCo"
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
Private porstCoCCo As ADODB.Recordset

']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   Set porstCoCCo = New ADODB.Recordset
   
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
              & "FROM CoCta"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   With porstCoCCo
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodCCo, DetCCo " _
              & "FROM CoCCo"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With

 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "CodCta"
         .Item(dnContador).MaxLength = porstCOCta.Fields(.Item(dnContador).DataField).DefinedSize
      Next
      For dnContador = 2 To 3
         .Item(dnContador).DataField = "CodCCo"
         .Item(dnContador).MaxLength = porstCoCCo.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
    With cboTpoMon
        .AddItem TPOMON_NAC_TXT_1, 0
        .AddItem TPOMON_EXT_TXT_1, 1
    End With
    cboTpoMon.ListIndex = TPOMON_NAC_IND

 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCta
      .MoveLast
      txtDato(1).Text = !CodCta
      .MoveFirst
      txtDato(0).Text = !CodCta
   End With
   With porstCoCCo
      .MoveLast
      txtDato(3).Text = !CodCCo
      .MoveFirst
      txtDato(2).Text = !CodCCo
   End With
   
  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
   If txtDato(2).Text <> "" Then ppAyuDet 2
   If txtDato(3).Text <> "" Then ppAyuDet 3
  
  'Otros.
   
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
   porstCoCCo.Close
   porstCOCta.Close
   pocnnMain.Close
   Set porstCOCta = Nothing
   Set porstCoCCo = Nothing
   Set porstMRp = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2, 3
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    Dim dnContador      As Byte
    Dim CadCrystal      As String
    Dim s_Moneda As String
    Dim s_SaldoDeb As String, s_SaldoHab As String
       
   ppHabilitacion False
    
    s_SaldoDeb = "ROUND(": s_SaldoHab = "ROUND("
    s_Moneda = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_0, TPOMON_EXT_TXT_0)
    For dnContador = 0 To (Val(gsMesAct) - 1)
      s_SaldoDeb = s_SaldoDeb & "AcuD" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = (Val(gsMesAct) - 1), "", "+")
      s_SaldoHab = s_SaldoHab & "AcuH" & gfCeros(Trim(dnContador), 2, 0, "0") & "_" & s_Moneda & IIf(dnContador = (Val(gsMesAct) - 1), "", "+")
    Next dnContador
    s_SaldoDeb = s_SaldoDeb & ", 2) AS cSalDeb,"
    s_SaldoHab = s_SaldoHab & ", 2) AS cSalHab"
    
    With porstMRp
       If .State = adStateOpen Then .Close
        .Source = "SELECT a.CodDro, a.NroCpb, a.FehOpe,"
        .Source = .Source & " CONCAT(c.AbvTDc, '-', a.SerDoc, '-', a.NroDoc) as cDocume,"
        .Source = .Source & " a.RefDoc, a.GloIte, a.CodCCo, b.DetCCo, a.CodCta, d.DetCta,"
        .Source = .Source & " IF(a.TpoCtb='" & TPOCTB_DEB & "', Imp" & s_Moneda & ", 0) as cDebe,"
        .Source = .Source & " IF(a.TpoCtb='" & TPOCTB_HAB & "', Imp" & s_Moneda & ", 0) as cHaber,"
        .Source = .Source & " " & s_SaldoDeb & " " & s_SaldoHab
        .Source = .Source & " FROM ((((COCpbDet a"
        .Source = .Source & " LEFT JOIN CoCCo b ON a.CodCCo=b.CodCCo)"
        .Source = .Source & " LEFT JOIN CoCCoAcu e ON a.CodCta=e.CodCta AND a.CodCCo=e.CodCCo)"
        .Source = .Source & " LEFT JOIN TGTDc c ON a.CodTDc=c.CodTDc)"
        .Source = .Source & " LEFT JOIN COCta d ON a.CodCta=d.CodCta)"
        .Source = .Source & " WHERE a.CodCta BETWEEN '" & txtDato(0).Text & "' AND '" & txtDato(1).Text & "' AND a.Mespvs ='" & gsMesAct & "'"
        .Source = .Source & " AND a.CodCCo BETWEEN '" & txtDato(2).Text & "' AND '" & txtDato(3).Text & "'"
        .Source = .Source & " ORDER BY a.CodCCo, a.CodCta, a.CodDro, a.NroCpb, a.NroIte "
       .Open
    End With
   
   usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
   If usDEstino = PRN_DEST_GRAF Then
        Call gpEncabezadoRpt(frmMain.rptMain, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
        pocnnMain.Execute "DROP TABLE IF EXISTS trptRMayAuxCCo"
        CadCrystal = "CREATE TABLE IF NOT EXISTS trptRMayAuxCCo " & porstMRp.Source
        pocnnMain.Execute CadCrystal

        With frmMain.rptMain
            '[Datos y parámetros del reporte.  'Cambiar.
            .ReportFileName = gsRutRpt & "rptRMayAuxCCo.rpt"
            .WindowShowExportBtn = IIf(paOpciones(2), True, False)
            .MarginLeft = unMargenIzquierdo
            .WindowState = crptMaximized
            .Connect = "Provider=MySqlProv;Extended Properties=" & CONNSTRG & gsNomBDS
            .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
            .Action = 1
      End With
      pocnnMain.Execute "DROP TABLE IF EXISTS trptRMayAuxCCo"
   Else
      Set MRViewer = New MRViewerObject

      With MRViewer
         .DataRecordSet = porstMRp
         .LoadReport gsRutRpt & "rptRMayAuxCCo.mrp"

         Call gpEncabezadoMRp(MRViewer, Me.Caption & " (" & IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC_TXT_1, TPOMON_EXT_TXT_1) & " )", udFecha, True, chkImpFecha.Value)
        '[Parámetros adicionales.
'         .Parameters("pPeriodoAdc") = IIf(optFecha(0).Value, "Emisión", "Cancelac.")
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
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   'Select Case Index    'Completa con ceros a la izquierda.
   'Case 0, 1                           'Cambiar (añadir índices).
   '   If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
   '      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
   '   End If
   'End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1, 2, 3                         'Cambiar (añadir índices).
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
     Case 2, 3                           'Cambiar (añadir índices).
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
            lblDatoDeta(tnIndex).Caption = " " & !DetCta
         End If
      End With
    Case 2, 3
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
            lblDatoDeta(tnIndex).Caption = " " & !DetCCo
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

