VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPPDTHPrMes 
   Caption         =   "[título]"
   ClientHeight    =   3390
   ClientLeft      =   2640
   ClientTop       =   3960
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CmnDlgUbica 
      Left            =   225
      Top             =   1395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   893
      TabIndex        =   2
      Top             =   2535
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   2573
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbEtapa1 
      Height          =   345
      Left            =   225
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label LblProces 
      Caption         =   "Procesando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   405
      Width           =   1635
   End
End
Attribute VB_Name = "frmPPDTHPrMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCOHPrDoc As ADODB.Recordset
Public pbNuevo As Boolean
Public pcNroCpb As String

Private Sub Form_Activate()
  LblProces.Visible = False
  cmdSalir.SetFocus
End Sub

Private Sub cmdAceptar_Click()
  On Error GoTo Err
  
  Dim dnContador As Integer
  
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  LblProces.Visible = True
  pgbEtapa1.Value = 0

  Set pocnnMain = New ADODB.Connection
  Set porstCOHPrDoc = New ADODB.Recordset
  
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
   
  With porstCOHPrDoc
    .ActiveConnection = pocnnMain
    .Source = "SELECT det.codaux, aux.tpodci, aux.rucaux, nat.numdci, det.codtdc, det.serdoc, det.nrodoc, "
    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(det.impmn, 0)*-1 ELSE IFNULL(det.impmn, 0) END), 2) AS impmn, "
    .Source = .Source & "MIN(hpr.feedoc) AS feedoc, MAX(det.fehope) AS fehope, hpr.indafeir4, hpr.impir4_mn, hpr.impir4_me, import_mn, import_me, hpr.tpomon, "
    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_mn, 0)*-1 ELSE IFNULL(hpr.impbru_mn, 0) END), 2) AS impmnh, "
    .Source = .Source & "ROUND(SUM(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN IFNULL(hpr.impbru_me, 0)*-1 ELSE IFNULL(hpr.impbru_me, 0) END), 2) AS impmeh, "
    .Source = .Source & "ROUND(AVG(det.imptcb), 4) AS imptcb "
    .Source = .Source & "FROM (((cocpbdet det "
    .Source = .Source & "INNER JOIN cohprdoc hpr ON det.codemp=hpr.codemp AND det.codaux=hpr.codaux AND det.serdoc=hpr.serdoc AND det.nrodoc=hpr.nrodoc) "
    .Source = .Source & "LEFT JOIN tgaux aux ON det.codemp=aux.codemp AND det.codaux=aux.codaux) "
    .Source = .Source & "LEFT JOIN tgauxnat nat ON aux.codemp=nat.codemp AND aux.codaux=nat.codaux) "
    .Source = .Source & "WHERE det.codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND det.pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND det.mespvs='" & gsMesAct & "' "
    .Source = .Source & "AND det.codtdc='" & CODTDC_HPR & "' "
    .Source = .Source & "AND det.tpopvs='" & TPOPVS_CAN & "' "
    .Source = .Source & "GROUP BY det.codaux, det.serdoc, det.nrodoc "
    '     .CursorLocation = adUseClient   'Es el Default.
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  
  'Etapa1 : Generando Texto segun lectura de Tabla.
  dnContador = 0
  pgbEtapa1.Min = 0
  pgbEtapa1.Value = pgbEtapa1.Min
  ppEtapa_01
  
  porstCOHPrDoc.Close
  pocnnMain.Close
  Set porstCOHPrDoc = Nothing
  Set pocnnMain = Nothing
  
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub ppEtapa_01()   ' Generacion de Texto en File
  Dim dnContador As Integer
  Dim dsTexto, dsFile As String
  Dim sCaracter As String
  Dim sImporte As String
  
  On Error GoTo CancelaDialogo
  
  dnContador = 0
  pgbEtapa1.Min = 0
  dsFile = "0601" & gsAnoAct & gsMesAct & gsRUCEmp & ".4ta"
  CmnDlgUbica.FileName = dsFile
  CmnDlgUbica.CancelError = True
  CmnDlgUbica.ShowSave
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then MsgBox error(Err.Number): Exit Sub
  On Error GoTo 0
  
  Open dsFile For Output As #1
  sCaracter = "|"
  Do
    With porstCOHPrDoc
      If .RecordCount = 0 Then Exit Do
      .MoveFirst
      pgbEtapa1.Max = .RecordCount
      pgbEtapa1.Value = pgbEtapa1.Min
      Do
        dsTexto = ""
        'dsTexto = Trim(IIf(IsNull(!codtdi), "", !codtdi)) & sCaracter
        
        dsTexto = Mid(!TpoDci, 2, 1) & sCaracter
        
        dsTexto = dsTexto & IIf(!TpoDci = "06", !rucaux, !numdci) & sCaracter
        
        'dsTexto = dsTexto & Trim(IIf(IsNull(!numdci), "", !numdci)) & sCaracter
        dsTexto = dsTexto & "R" & sCaracter
        dsTexto = dsTexto & Trim(IIf(IsNull(!serdoc), "", IIf(Left(!serdoc, 1) = "E", !serdoc, Right(!serdoc, 3)))) & sCaracter
        dsTexto = dsTexto & Trim(IIf(IsNull(!nrodoc), "", Mid(!nrodoc, 3, 8))) & sCaracter
        sImporte = Format(!ImpMNH, "############.00")
        If !tpomon = TPOMON_EXT Then
          sImporte = Format(Round(CDec(!ImpMEH) * CDec(!ImpTCb), 2), "############.00")
        End If
    '    sImporte = Replace(sImporte, ".", "")
        dsTexto = dsTexto & Trim(sImporte) & sCaracter
        dsTexto = dsTexto & Format(!feedoc, "dd/mm/yyyy") & sCaracter
        dsTexto = dsTexto & Format(!fehope, "dd/mm/yyyy") & sCaracter
                
        dsTexto = dsTexto & IIf(!ImpIR4_MN + !ImpIR4_ME <> 0, "1", "0") & sCaracter & "3" & sCaracter & sCaracter
        'dsTexto = dsTexto & Trim(!IndAfeIR4) & sCaracter
        Print #1, dsTexto
        dnContador = dnContador + 1
        pgbEtapa1.Value = dnContador
        .MoveNext
      Loop Until .EOF
    End With
    Exit Do
  Loop
  Close #1
  MsgBox TEXT_8008, vbInformation

End Sub

Private Sub Form_Load()
  
  '[ Cargo los mensajes de botones
  ReDim aLabel(0, 0)
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  LblProces.Caption = Choose(gsIdioma, "Procesando", "Processing")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']

End Sub
