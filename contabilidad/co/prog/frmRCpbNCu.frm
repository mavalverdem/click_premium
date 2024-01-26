VERSION 5.00
Begin VB.Form frmRCpbNCu 
   Caption         =   "[título]"
   ClientHeight    =   2295
   ClientLeft      =   1620
   ClientTop       =   345
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTipoImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2640
      TabIndex        =   17
      Top             =   1020
      Width           =   2175
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Gráfica"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton optTipoImpresion 
         Caption         =   "Matricial"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   990
      Width           =   2175
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   1035
         TabIndex        =   7
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle"
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Frame fraRangos 
      Caption         =   "Mes"
      ForeColor       =   &H80000002&
      Height          =   780
      Left            =   0
      TabIndex        =   11
      Top             =   90
      Width           =   3060
      Begin VB.ComboBox CmbMes 
         Height          =   315
         ItemData        =   "frmRCpbNCu.frx":0000
         Left            =   1035
         List            =   "frmRCpbNCu.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   315
         Width           =   1905
      End
      Begin VB.CheckBox chkMes 
         Caption         =   "Todos"
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4125
         Picture         =   "frmRCpbNCu.frx":00A8
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1170
         Visible         =   0   'False
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
         Left            =   165
         TabIndex        =   8
         Top             =   1155
         Width           =   315
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   4125
         Picture         =   "frmRCpbNCu.frx":0252
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1530
         Visible         =   0   'False
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
         Left            =   165
         TabIndex        =   9
         Top             =   1515
         Visible         =   0   'False
         Width           =   315
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
         Left            =   465
         TabIndex        =   15
         Top             =   1155
         Width           =   3675
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
         Left            =   465
         TabIndex        =   14
         Top             =   1515
         Visible         =   0   'False
         Width           =   3675
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
      ScaleWidth      =   4845
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1755
      Width           =   4845
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
         Picture         =   "frmRCpbNCu.frx":03FC
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
         Picture         =   "frmRCpbNCu.frx":0546
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
         Picture         =   "frmRCpbNCu.frx":0A78
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRCpbNCu"
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
Private porstCOCpbDet As ADODB.Recordset
Private porstCrystal As ADODB.Recordset
']

Private Sub Form_Load()
   On Error GoTo Err
  
   Dim dnContador As Integer

 '[Recordsets.                         'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstMRp = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   
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
   With porstCOCpbDet
      .ActiveConnection = pocnnMain
      .Source = "SELECT MesPvs " _
              & "FROM CocpbDet"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
 ']

 '[Parámetros.                         'Cambiar.
   With txtDato
      For dnContador = 0 To 1
         .Item(dnContador).DataField = "MesPvs"
         .Item(dnContador).MaxLength = porstCOCpbDet.Fields(.Item(dnContador).DataField).DefinedSize
      Next
   End With
 ']
   
 '[Datos predeterminados.              'Cambiar.
  'Límites de rangos.
   With porstCOCpbDet
      .MoveLast
      txtDato(1).Text = !MesPvs
      .MoveFirst
      txtDato(0).Text = !MesPvs
   End With

  'Busca detalle de códigos            '(habilitar/deshabilitar).
   If txtDato(0).Text <> "" Then ppAyuDet 0
   If txtDato(1).Text <> "" Then ppAyuDet 1
  
  'Otros.
  CmbMes.ListIndex = gsMesAct
   
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
   porstCOCpbDet.Close
   pocnnMain.Close
   Set porstCOCpbDet = Nothing
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
   Dim dnContador As Byte
   Dim CadCrystal As String

   ppHabilitacion False
    
    With porstMRp
       If .State = adStateOpen Then .Close
       If OptTipo(0).Value = True Then
            .Source = "SELECT a.CodCta, a.CodDro, a.NroCpb, a.GloIte , a.CodAux, b.MesPvs, d.RazAux, a.BlqIte, a.FehOpe,"
       Else
            .Source = "SELECT a.CodCta, a.CodDro, a.NroCpb, b.GloCpb , a.CodAux, b.MesPvs, d.RazAux, a.BlqIte, b.FehCpb,"
       End If
       .Source = .Source _
               & "  Concat(c.AbvTDc,'-',a.SerDoc,'-',a.NroDoc) as cNroDoc," _
               & "  IF(a.TpoCtb='D',ImpMN,0) as clmpDeb," _
               & "  IF(a.TpoCtb='H',ImpMN,0) as clmpHab," _
               & "  IF(a.TpoCtb='D',ImpME,0) as clmpDebME," _
               & "  IF(a.TpoCtb='H',ImpME,0) as clmpHabME," _
               & "  Concat(a.CodDro,'-', a.NroCpb) as cDroCpb " _
               & "FROM (((CocPbDet a" _
               & "  LEFT JOIN CoCpbCab b ON a.CodDro=b.CodDro AND a.NroCpb=b.NroCpb" _
               & "  LEFT JOIN TgTDc c ON a.CodTDc=c.CodTDc " _
               & "  LEFT JOIN TgAux d ON a.CodAux=d.CodAux )))" _
               & "WHERE b.IndNCu=1 "
        If chkMes.Value = False Then
           .Source = .Source & " AND b.MesPvs = '" + Format(CmbMes.ListIndex, "00") + "' "
        Else
           .Source = .Source & " AND b.MesPvs <= '" + Format(CmbMes.ListIndex, "00") + "' "
        End If
'[Raúl 090104.
'       .Source = .Source & " ORDER BY b.MesPvs, a.CodDro, a.NroCpb, a.CodCta"
       .Source = .Source & " ORDER BY b.MesPvs, a.CodDro, a.NroCpb, a.NroIte"
']Raúl.
       .Open
    End With
   
   usDEstino = IIf(optTipoImpresion(0).Value, PRN_DEST_MATR, PRN_DEST_GRAF)
   If usDEstino = PRN_DEST_GRAF Then
      Call gpEncabezadoRpt(frmMain.rptMain, Me.Caption & IIf(OptTipo(0).Value = True, " (Detallado)", " (Resumen)"), udFecha, True)
        
      Prepara_Crystal
        
      With frmMain.rptMain
         '[Datos y parámetros del reporte.  'Cambiar.
         If OptTipo(0).Value = True Then
            .ReportFileName = gsRutRpt & "rptRCpbNCu.rpt"
         Else
            .ReportFileName = gsRutRpt & "rptRCpbNCuRes.rpt"
         End If
         .SelectionFormula = "{cotmprpt.UsrCre}='" & gsCodUsr & "' AND {cotmprpt.NomRpt}='rptRCpbNCu'"
         .WindowShowExportBtn = IIf(paOpciones(2), True, False)
         .MarginLeft = unMargenIzquierdo
         .WindowState = crptMaximized
         .Connect = "Provider=MySqlProv;Extended Properties=" & CONNSTRG & gsNomBDS
         .Destination = IIf(crptToPrinter = Index, crptToPrinter, crptToWindow)
         .Action = 1
      End With
   Else
      Set MRViewer = New MRViewerObject
      With MRViewer
        .DataRecordSet = porstMRp
        If OptTipo(0).Value = True Then
            .LoadReport gsRutRpt & "rptRCpbNCu.mrp"
        Else
            .LoadReport gsRutRpt & "rptRCpbNCuRes.mrp"
        End If
     
        Call gpEncabezadoMRp(MRViewer, Me.Caption & IIf(OptTipo(0).Value = True, " (Detallado)", " (Resumen)"), udFecha, True)
     
        '[Parámetros adicionales.
        If chkMes.Value = False Then
            .Parameters("pPeriodoAdc") = Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
        Else
            .Parameters("pPeriodoAdc") = "A " & Format(CDate(gsMesAct & " " & gsAnoAct), "mmmm") & " " & gsAnoAct
        End If
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
   Select Case Index    'Completa con ceros a la izquierda.
   Case 0, 1                           'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Busca el dato en su tabla principal.
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0, 1                           'Cambiar (añadir índices).
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + fraRangos.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraRangos.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   
   Select Case tnIndex                 'Cambiar.
    Case 0, 1
      If Val(txtDato(0)) > 12 Or Val(txtDato(0)) = 0 Then
        txtDato(0).Text = "99"
        lblDatoDeta(0).Caption = "Todos los Meses"
      Else
        lblDatoDeta(0).Caption = Format(CVDate("01" & "/" & (txtDato(tnIndex) & "/" & Year(udFecha))), "mmmm")
      End If
'      If txtDato(tnIndex).Text = "" Then
'         lblDatoDeta(tnIndex).Caption = ""
'         Exit Function
'      End If
'      With porstCocpbDet
'         .MoveFirst
'         .Find "MesPvs='" & txtDato(tnIndex).Text & "'"
'         If .EOF Then
'            MsgBox TEXT_8006, vbExclamation
'            ppAyuDet = True
'         Else
'            lblDatoDeta(tnIndex).Caption = " " & !DetTDc
'         End If
'      End With
   End Select
End Function

'[Propio del formulario.

Private Sub ChkMes_Click()
    If chkMes.Value = 0 Then CmbMes.Enabled = True
    If chkMes.Value = 1 Then CmbMes.Enabled = False
End Sub

Private Sub Prepara_Crystal()
    Dim cad     As String
    Dim cad1    As String
    Dim c       As Byte
    
    'jp
    Set porstCrystal = New ADODB.Recordset
    
    With porstCrystal
      .ActiveConnection = pocnnMain
      .CursorType = adOpenForwardOnly
      .LockType = adLockOptimistic
      .Source = "SELECT * FROM COTmpRpt WHERE UsrCre='" & gsCodUsr & "' AND NomRpt = 'rptRCpbNCu'"
    End With

    pocnnMain.Execute "DELETE FROM COTmpRpt WHERE UsrCre='" & gsCodUsr & "' AND NomRpt = 'rptRCpbNCu'"
    If porstCrystal.State = adStateOpen Then porstCrystal.Close
    porstCrystal.Open
    If porstMRp.RecordCount > 0 Then
        porstMRp.MoveFirst
        Do While Not porstMRp.EOF
            pocnnMain.BeginTrans   '[ INICIA TRANSACCION ]
            porstCrystal.AddNew
            porstCrystal.Fields!UsrCre = gsCodUsr
            porstCrystal.Fields!NomRpt = "rptRCpbNCu"
            porstCrystal.Fields!CodCta = porstMRp.Fields!CodCta
            porstCrystal.Fields!CodDro = porstMRp.Fields!CodDro
            porstCrystal.Fields!NroCpb = porstMRp.Fields!NroCpb
            If OptTipo(0).Value = True Then
                porstCrystal.Fields!GloIte = porstMRp.Fields!GloIte
                porstCrystal.Fields!FehOpe = porstMRp.Fields!FehOpe
            Else
                porstCrystal.Fields!GloIte = porstMRp.Fields!GloCpb
                porstCrystal.Fields!FehOpe = porstMRp.Fields!FehCpb
            End If
            porstCrystal.Fields!CodAux = porstMRp.Fields!CodAux
            porstCrystal.Fields!MesPvs = porstMRp.Fields!MesPvs
            porstCrystal.Fields!RazAux = porstMRp.Fields!RazAux
            porstCrystal.Fields!BlqIte = porstMRp.Fields!BlqIte
            porstCrystal.Fields!cDocum = porstMRp.Fields!cNroDoc
            porstCrystal.Fields!cDroCpb = porstMRp.Fields!cDroCpb
            porstCrystal.Fields!NumCol1 = porstMRp.Fields!clmpDeb
            porstCrystal.Fields!NumCol2 = porstMRp.Fields!clmpHab
            porstCrystal.Fields!NumCol3 = porstMRp.Fields!clmpDebME
            porstCrystal.Fields!NumCol4 = porstMRp.Fields!clmpHabME
            porstCrystal.Update
            pocnnMain.CommitTrans  '[ CONFIRMA TRANSACCION ]
            porstMRp.MoveNext
        Loop
    End If
    porstCrystal.Close
    
    Set porstCrystal = Nothing
    'jp

End Sub

']

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

