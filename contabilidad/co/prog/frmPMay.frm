VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPMay 
   Caption         =   "[título]"
   ClientHeight    =   4560
   ClientLeft      =   2460
   ClientTop       =   2235
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Estado Proceso"
      Height          =   495
      Left            =   1560
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox chkmayorizar 
      Caption         =   "Mayorizar Hasta el Periodo"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "Mayoriza &Auxiliares"
      Height          =   200
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin VB.TextBox TxtDato 
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Text            =   "2101"
      Top             =   1170
      Width           =   465
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   330
      Index           =   0
      Left            =   4335
      Picture         =   "frmPMay.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1170
      Width           =   255
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "Mayoriza &Centro de Costos"
      Height          =   200
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   2325
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "&Mayoriza Cuentas"
      Height          =   200
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin VB.CheckBox chkProceso 
      Caption         =   "&Generación Asientos Destinos"
      Height          =   200
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Value           =   1  'Checked
      Width           =   4000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1950
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2580
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   3255
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ingrese Diario"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   900
      Width           =   1275
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
      Height          =   330
      Index           =   0
      Left            =   690
      TabIndex        =   10
      Top             =   1170
      Width           =   3660
   End
End
Attribute VB_Name = "frmPMay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCoTCbMes As ADODB.Recordset
Public porstCodro As ADODB.Recordset
Dim sqlCuenta As String
Dim lVisualizar As Boolean



'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
Private Sub Command1_Click()
  
  With frmPMayEst
    '.ConfiguraPrn 0, Me
    .Show vbModal
    '.ConfiguraPrn 1, Me
  End With

End Sub
'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion

Private Sub Form_Load()


  chkmayorizar.Caption = "Mayorizar Hasta el Periodo: " & gsAnoAct & gsMesAct
  
  'Abrir Tablas.
  Set pocnnMain = New ADODB.Connection
  Set porstCoTCbMes = New ADODB.Recordset
  Set porstCodro = New ADODB.Recordset
  
  With pocnnMain
    .CursorLocation = adUseClient
    .ConnectionString = CONNSTRG & gsNomBDS
    .Open
  End With
   
  With porstCodro
    .ActiveConnection = pocnnMain
    .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
    .Source = .Source & "FROM CODro "
    .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
    .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
    .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4"
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
  End With
   
  With porstCoTCbMes
    .ActiveConnection = pocnnMain
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
  End With
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(1, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Ingrese Diario")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Enter Journal")
  Next nElemento
  chkProceso(0).Caption = Choose(gsIdioma, "&Generación Asientos Destinos", "&Generate Destination Entries")
  chkProceso(1).Caption = Choose(gsIdioma, "&Mayoriza Cuentas", "Centralization of &Accounts")
  chkProceso(2).Caption = Choose(gsIdioma, "Mayoriza &Centro de Costos", "Centralization of &Cost Center")
  chkProceso(3).Caption = Choose(gsIdioma, "Mayoriza &Auxiliares", "Centralization of Au&xiliaries")
  cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']
'ini 2015-05-18 validacion frm
    If gsMesAct = "00" Or gsMesAct = "13" Then
    chkProceso(0).Value = False
    End If
'fin 2015-05-18 validacion frm

End Sub

Private Sub Form_Activate()
   chkProceso(0).Enabled = True
   chkProceso(1).Enabled = True
   chkProceso(2).Enabled = True
   chkProceso(3).Enabled = True
   cmdSalir.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo error
   porstCoTCbMes.Close
   porstCodro.Close
   pocnnMain.Close
   Set porstCoTCbMes = Nothing
   Set porstCodro = Nothing
   Set pocnnMain = Nothing
error:
End Sub

Private Sub cmdAceptar_Click()
  Dim nContador As Integer, nMesIni As Integer
  Dim consulta As ADODB.Recordset
  
  On Error GoTo ErrorMayoriza
  
  ' Validacion de mes cerrado
  If (gbCieCpb And chkProceso(0).Value) Then MsgBox TEXT_9016, vbCritical: cmdSalir.SetFocus: Exit Sub
  If (chkProceso(0).Value And txtDato(0).Text = "") Then MsgBox TEXT_6002, vbCritical: txtDato(0).SetFocus: Exit Sub
  
    'ini 2015-06-24 control flag mayoriza
'ini 2015-07-27 error mayoriza varios meses
    'If chkProceso(0).Value Then
    If chkProceso(0).Value And chkmayorizar.Value Then
'fin 2015-07-27 error mayoriza varios meses
        If gcCierre(gsAnoAct, gsMesAct) = 1 Then Exit Sub
    End If
    'fin 2015-06-24 control flag mayoriza

  
  Set consulta = New ADODB.Recordset
  cmdAceptar.Enabled = False
  cmdSalir.Enabled = False
  pgbProceso(0).Value = 0: pgbProceso(0).Min = 0
  pgbProceso(1).Value = 0: pgbProceso(1).Min = 0
  pgbProceso(2).Value = 0: pgbProceso(2).Min = 0
  pgbProceso(3).Value = 0: pgbProceso(3).Min = 0
  nMesIni = IIf(chkmayorizar.Value = Checked, 0, CInt(gsMesAct))
  
  pocnnMain.BeginTrans                'INICIA TRANSACCION.
  
  For nContador = nMesIni To CInt(gsMesAct)
    ' Tipo de cambio
    With porstCoTCbMes
      If .State = adStateOpen Then .Close
      .Source = "SELECT ImpTCb_Cpr, ImpTCb_Vta "
      .Source = .Source & "FROM cotcbmes "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND mespvs='" & Format(nContador, "00") & "'"
      .Open
    End With
    
    'Paso 1: Elimino los asientos de destino del mes y regenero las cuentas de destino.
    pocnnMain.Execute "DELETE FROM CoCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr='" & TPOGNR_DST & "' AND MesPvs='" & Format(nContador, "00") & "'"
    If gnProDestino = NvlUsr_Sup Then
      pocnnMain.Execute "DELETE FROM CoCpbDet WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr='" & TPOGNR_DST & "' AND MesPvs='" & Format(nContador, "00") & "'"
    End If
    If (chkProceso(0).Value And gnProDestino = NvlUsr_Adm) Then ppGene_CuentaDst_Gen Format(nContador, "00"), 0
    If (chkProceso(0).Value And gnProDestino = NvlUsr_Sup) Then ppGene_CuentaDst_Cpr Format(nContador, "00"), 0
    'Paso 2: Reemplazo con cero los campos de acumulado mensual y mayorizo las cuentas.
    If chkProceso(1).Value Then ppMayoriza_Cuenta Format(nContador, "00"), 1
    'Paso 3: Reemplazo con cero los campos de acumulado mensual y mayorizo los centro de costo.
    If chkProceso(2).Value Then ppMayoriza_CenCosto Format(nContador, "00"), 2
    'Paso 4: Reemplazo con cero los campos de acumulado mensual y mayorizo los auxiliares.
    If chkProceso(3).Value Then ppMayoriza_Auxiliar Format(nContador, "00"), 3
    
'ini 2015-09-25 se necesita que reconozca el mes diferente al actual
    fEstMayUpd -1, Format(nContador, "00")
'fin 2015-09-25 se necesita que reconozca el mes diferente al actual
    
  Next nContador
  
  pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
   
  MsgBox TEXT_8008, vbInformation
  cmdAceptar.Enabled = True
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  
'ini 2015-09-25 se necesita que reconozca el mes diferente al actual
'''ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
''fEstMayUpd -1
'''fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
'fin 2015-09-25 se necesita que reconozca el mes diferente al actual
  
  
  Exit Sub

ErrorMayoriza:
   
  If lVisualizar = True Then
    With consulta
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sqlCuenta
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
    End With
    
    If consulta.RecordCount > 0 Then
      consulta.MoveFirst
      Do While Not consulta.EOF
        MsgBox " Error, Cuenta Contable: " & consulta!mespvs & "- " & consulta!codcta
        consulta.MoveNext
      Loop
    End If
    
    consulta.Close
    Set consulta = Nothing
  End If
   
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description
  cmdSalir.Enabled = True
  cmdSalir.SetFocus
  
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      ppAyuBus AYUDAT, Index
      txtDato(Index).SetFocus
   End Select
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYUDAT Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4 ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
   End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYUDAT Then
      Select Case tnIndex              'Cambiar.
      Case 0
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With porstCodro
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodDro='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
            End If
         End With
      End Select
   End If
End Function

Private Sub ppGene_CuentaDst_Gen(ByVal sPeriodo As String, ByVal nIndex As Integer)
  Static porstCOCpbCab As ADODB.Recordset
  Static porstCOCpbDet As ADODB.Recordset
  Static porstCOCpbDst As ADODB.Recordset
  Static porstUltCoCpb  As ADODB.Recordset
   
  Static sNroComprobante As String
  Static sSentencia As String, sGrabacion As String
  Static nNroItem As Integer, nRegistros As Integer
  Static nItems As Integer, nLimite As Integer
   
  Set porstCOCpbCab = New ADODB.Recordset
  Set porstCOCpbDet = New ADODB.Recordset
  Set porstCOCpbDst = New ADODB.Recordset
  Set porstUltCoCpb = New ADODB.Recordset
   
  'Elimina el asiento de cuentas de destino.
  pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DST) & " AND MesPvs='" & sPeriodo & "'"
  pocnnMain.Execute "DELETE FROM COCpbDet WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DST) & " AND MesPvs='" & sPeriodo & "'"
  pgbProceso(nIndex).Min = 0
  'Genero la sentencia de selección de acumulados de destinos.
  sSentencia = "SELECT b.CodCta_Dst_Deb AS cCuenta, "
  sSentencia = sSentencia & "b.CodCCo_Dst_Deb AS cCenCosto, '" & TPOCTB_DEB & "' AS cTpoCtb, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2), 2) AS nImporteMN, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2) , 2) AS nImporteME "
  sSentencia = sSentencia & "FROM COCpbDet a, CoCta b "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND b.codemp=a.codemp "
  sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
  sSentencia = sSentencia & "AND b.CodCta=a.CodCta "
  sSentencia = sSentencia & "AND b.TpoCta =" & TPOCTA_TRA & " "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Dst_Deb, '') <> '' "
  sSentencia = sSentencia & "AND MesPvs='" & sPeriodo & "' "
  sSentencia = sSentencia & "GROUP BY b.CodCta_Dst_Deb, b.CodCCo_Dst_Deb "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (nImporteMN <> 0.00 OR nImporteME <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) - "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2), 2) <> 0.00 "
    sSentencia = sSentencia & "OR ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) - "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2) , 2) <> 0.00) "
  End If
  sSentencia = sSentencia & "UNION "
  sSentencia = sSentencia & "SELECT b.CodCta_Dst_Hab AS cCuenta, "
  sSentencia = sSentencia & "b.CodCCo_Dst_Hab AS cCenCosto, '" & TPOCTB_HAB & "' AS cTpoCtb, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2), 2) AS nImporteMN, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2) , 2) AS nImporteME "
  sSentencia = sSentencia & "FROM COCpbDet a, COCta b "
  sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND b.codemp=a.codemp "
  sSentencia = sSentencia & "AND b.pdoano=a.pdoano "
  sSentencia = sSentencia & "AND b.CodCta=a.CodCta "
  sSentencia = sSentencia & "AND b.TpoCta =" & TPOCTA_TRA & " "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Dst_Hab, '') <> '' "
  sSentencia = sSentencia & "AND MesPvs='" & sPeriodo & "' "
  sSentencia = sSentencia & "GROUP BY b.CodCta_Dst_Hab, b.CodCCo_Dst_Hab "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (nImporteMN <> 0.00 OR nImporteME <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) - "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2), 2) <> 0.00 "
    sSentencia = sSentencia & "OR ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) - "
    sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2) , 2) <> 0.00) "
  End If
  'Abro el recordset de seleccion de destinos.
  With porstCOCpbDst
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  If porstCOCpbDst.RecordCount > 0 Then
    porstCOCpbDst.MoveFirst
    pgbProceso(nIndex).Max = porstCOCpbDst.RecordCount
    pgbProceso(nIndex).Value = pgbProceso(nIndex).Min
    'Genero la sentencia de selección cabecera de comprobantes.
    sGrabacion = "SELECT codemp, pdoano, CodDro, NroCpb, FehCpb, GloCpb, glocpbx, MesPvs, "
    sGrabacion = sGrabacion & "TpoGnr, IndNCu, IndAnu, "
    sGrabacion = sGrabacion & "UsrCre, FyHCre "
    sGrabacion = sGrabacion & "FROM COCpbCab "
    sGrabacion = sGrabacion & "WHERE codemp='" & gsCodEmp & "' "
    sGrabacion = sGrabacion & "AND pdoano='" & gsAnoAct & "' "
    sGrabacion = sGrabacion & "AND CodDro=''"
    'Abro el recordset de grabación de la cabecera de comprobante.
    With porstCOCpbCab
      .ActiveConnection = pocnnMain
      .Source = sGrabacion
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    'Obtengo el número e inserto la cabecera del comprobante.
    With porstUltCoCpb
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), 0) AS cUltNroCpb "
      .Source = .Source & "FROM COCpbCab "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND MesPvs='" & sPeriodo & "' "
      .Source = .Source & "AND CodDro='" & txtDato(0).Text & "'"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
      sNroComprobante = !cUltNroCpb
      .Close
    End With
    sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
    With porstCOCpbCab
      .AddNew
      !codemp = gsCodEmp
      !pdoano = gsAnoAct
      !mespvs = sPeriodo
      !coddro = txtDato(0).Text
      !NroCpb = sNroComprobante
      !FehCpb = gfUltDia("01/" & sPeriodo & "/" & gsAnoAct)
      !glocpb = "Asiento de Cuentas de Destino"
      !glocpbx = "Destination of Accounts Entries"
      !tpognr = TPOGNR_DST
      !IndNCu = INDNCU_FAL
      !IndAnu = INDANU_FAL
      !UsrCre = gsAbvUsr
      !FyHCre = Now
      .Update
    End With
    nNroItem = 0
     
    'Genero la sentencia de selección detalles de comprobantes.
    sGrabacion = "SELECT codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, "
    sGrabacion = sGrabacion & "SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, gloitex, TpoCtb, TpoPvs, "
    sGrabacion = sGrabacion & "TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, "
    sGrabacion = sGrabacion & "UsrCre, FyHCre "
'ini 2016-06-24 correcion codmon asto destino y dif cam
    sGrabacion = sGrabacion & ",CodMon "
'fin 2016-06-24 correcion codmon asto destino y dif cam
    sGrabacion = sGrabacion & "FROM COCpbDet "
    sGrabacion = sGrabacion & "WHERE codemp='" & gsCodEmp & "' "
    sGrabacion = sGrabacion & "AND pdoano='" & gsAnoAct & "' "
    sGrabacion = sGrabacion & "AND CodDro=''"
    'Abro el recordset de grabación de la cabecera de comprobante.
    With porstCOCpbDet
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sGrabacion
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
     
     nRegistros = nNroItem
     Do While Not porstCOCpbDst.EOF
        nLimite = IIf((porstCOCpbDst!nImporteMN >= 0 And porstCOCpbDst!nImporteME >= 0) Or (porstCOCpbDst!nImporteMN <= 0 And porstCOCpbDst!nImporteME <= 0), 1, 2)
        nRegistros = nRegistros + 1
        For nItems = 1 To nLimite
           nNroItem = nNroItem + 1
           ' Asigno la ubicacion del importe D/H
           sSentencia = porstCOCpbDst!cTpoCtb
           If nLimite = 1 Then
               sSentencia = IIf((porstCOCpbDst!nImporteMN < 0 Or porstCOCpbDst!nImporteME < 0), IIf(porstCOCpbDst!cTpoCtb = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB), sSentencia)
           Else
               sSentencia = IIf(Choose(nItems, (porstCOCpbDst!nImporteMN >= 0), (porstCOCpbDst!nImporteME >= 0)), sSentencia, IIf(porstCOCpbDst!cTpoCtb = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB))
           End If
           ' Adiciono el detalle del comprobante.
           With porstCOCpbDet
              .AddNew
              !codemp = gsCodEmp
              !pdoano = gsAnoAct
              !mespvs = sPeriodo
              !coddro = txtDato(0).Text
              !NroCpb = sNroComprobante
              !NroIte = nNroItem
              !blqite = nNroItem
              !codcta = porstCOCpbDst!cCuenta
              !fehope = gfUltDia("01/" & sPeriodo & "/" & gsAnoAct)
              !feedoc = gfUltDia("01/" & sPeriodo & "/" & gsAnoAct)
              !fevdoc = gfUltDia("01/" & sPeriodo & "/" & gsAnoAct)
              !ferdoc = gfUltDia("01/" & sPeriodo & "/" & gsAnoAct)
              !codtdc = Null
              !codcco = porstCOCpbDst!cCenCosto
              !codaux = Null
              !serdoc = Null
              !nrodoc = Null
              !GloIte = "Asiento de Cuenta de Destino"
              !GloItex = "Destination of Account Entries"
              'If (IIf(nLimite = 1, (porstCOCpbDst!nImporteMN >= 0 Or porstCOCpbDst!nImporteME >= 0), Choose(nItems, (porstCOCpbDst!nImporteMN >= 0), (porstCOCpbDst!nImporteME >= 0))), porstCOCpbDst!cTpoCtb, IIf(porstCOCpbDst!cTpoCtb = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB))
              !TpoCtb = sSentencia
              !tpomon = TPOMON_NAC
'ini 2016-06-24 correcion codmon asto destino y dif cam
              !codmon = IIf(!tpomon = TPOMON_NAC, CODMON_NAC, CODMON_EXT)
'fin 2016-06-24 correcion codmon asto destino y dif cam
              !TpoTcb = TPOTCB_VTA
              !ImpTCb = IIf(IsNull(porstCoTCbMes!ImpTCb_Vta), 1, porstCoTCbMes!ImpTCb_Vta)
              !ImpMN = Abs(IIf(nLimite = 1, porstCOCpbDst!nImporteMN, Choose(nItems, porstCOCpbDst!nImporteMN, 0)))
              !ImpME = Abs(IIf(nLimite = 1, porstCOCpbDst!nImporteME, Choose(nItems, 0, porstCOCpbDst!nImporteME)))
              !TpoPvs = TPOPVS_OTR
              !tpognr = TPOGNR_DST
              !UsrCre = gsAbvUsr
              !FyHCre = Now
           End With
           pgbProceso(nIndex).Value = nRegistros
        Next nItems
        porstCOCpbDst.MoveNext
     Loop
     porstCOCpbDet.UpdateBatch
     'Cierro y saco de memoria los recordset.
     porstCOCpbDet.Close
     porstCOCpbCab.Close
     Set porstCOCpbDet = Nothing
     Set porstCOCpbCab = Nothing
     Set porstUltCoCpb = Nothing
  End If
  'Cierro y saco de memoria los recordset.
  porstCOCpbDst.Close
  Set porstCOCpbDst = Nothing
End Sub

Private Sub ppGene_CuentaDst_Cpr(ByVal sPeriodo As String, ByVal nIndex As Integer)
  Dim sDiario As String, sNroComprobante As String
  Static sSentencia As String, sGrabacion As String, sExpresion As String
  Static nNroItem As Long, nRegistros As Long
  Static nItems As Integer, nLimite As Integer
   
  Dim porstCOCpbDst As ADODB.Recordset
  Dim porstUltItem  As ADODB.Recordset
   
  Set porstCOCpbDst = New ADODB.Recordset
  Set porstUltItem = New ADODB.Recordset
   
  'Elimina el asiento de cuentas de destino.
  pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DST) & " AND MesPvs='" & sPeriodo & "'"
  pocnnMain.Execute "DELETE FROM COCpbDet WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DST) & " AND MesPvs='" & sPeriodo & "'"
  pgbProceso(nIndex).Min = 0
  pocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpdestino", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 11)='#tmpdestino') DROP TABLE #tmpdestino")
  'Genero la sentencia de selección de acumulados de destinos.
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdestino ", "")
  sSentencia = sSentencia & "SELECT "
  sSentencia = sSentencia & "det.codemp, det.pdoano, det.coddro, det.nrocpb, det.nroite + 1 AS nroite, det.blqite, det.codtdc, det.fehope, "
  sSentencia = sSentencia & "cta.codcta_dst_deb AS codcta, det.codcco, det.codaux, det.serdoc, det.nrodoc, det.feedoc, det.fevdoc, det.ferdoc, "
  sSentencia = sSentencia & "det.refdoc, det.pdocpr, det.gloite, det.gloitex, '" & TPOCTB_DEB & "' AS tpoctb, det.tpopvs, det.tpomon, det.tpotcb, "
  sSentencia = sSentencia & "det.imptcb, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END, 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END, 0), 2), 2) AS impmn, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END, 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END, 0), 2), 2) AS impme, "
  sSentencia = sSentencia & "det.tpognr, det.indfjo_det, det.indgnr_rp, det.tpodoc, det.codcon, det.codmon, cta.codcco_dst_deb AS codcco_dst "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvSql, "INTO " & ps_Prefijo & "tmpdestino ", "")
  sSentencia = sSentencia & "FROM cocpbdet det, cocpbcab cab, cocta cta "
  sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND det.mespvs='" & sPeriodo & "' "
  sSentencia = sSentencia & "AND cab.codemp=det.codemp AND cab.pdoano=det.pdoano AND cab.mespvs=det.mespvs AND cab.coddro=det.coddro AND cab.nrocpb=det.nrocpb "
  sSentencia = sSentencia & "AND cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
  sSentencia = sSentencia & "AND cta.tpocta=" & TPOCTA_TRA & " AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cta.codcta_dst_deb, '') <> '' "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (impmn <> 0.00 OR impme <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END, 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END, 0), 2), 2) <> 0.00 "
    sSentencia = sSentencia & "ROUND(ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END, 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END, 0), 2), 2) <> 0.00) "
  End If
  sSentencia = sSentencia & "UNION "
  sSentencia = sSentencia & "SELECT "
  sSentencia = sSentencia & "det.codemp, det.pdoano, det.coddro, det.nrocpb, nroite + 2 AS nroite, det.blqite, det.codtdc, det.fehope, "
  sSentencia = sSentencia & "cta.codcta_dst_hab AS codcta, det.codcco, det.codaux, det.serdoc, det.nrodoc, det.feedoc, det.fevdoc, det.ferdoc, "
  sSentencia = sSentencia & "det.refdoc, det.pdocpr, det.gloite, det.gloitex, '" & TPOCTB_HAB & "' AS tpoctb, det.tpopvs, det.tpomon, det.tpotcb, "
  sSentencia = sSentencia & "det.imptcb, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END, 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END, 0), 2), 2) AS impmn, "
  sSentencia = sSentencia & "ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END, 0), 2) - "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END, 0), 2), 2) AS impme, "
  sSentencia = sSentencia & "det.tpognr, det.indfjo_det, det.indgnr_rp, det.tpodoc, det.codcon, det.codmon, cta.codcco_dst_hab AS codcco_dst "
  sSentencia = sSentencia & "FROM cocpbdet det, cocpbcab cab, cocta cta "
  sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND det.mespvs='" & sPeriodo & "' "
  sSentencia = sSentencia & "AND cab.codemp=det.codemp AND cab.pdoano=det.pdoano AND cab.mespvs=det.mespvs AND cab.coddro=det.coddro AND cab.nrocpb=det.nrocpb "
  sSentencia = sSentencia & "AND cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
  sSentencia = sSentencia & "AND cta.tpocta=" & TPOCTA_TRA & " AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(cta.codcta_dst_hab, '') <> '' "
  If ps_Plataforma = pSrvMySql Then
    sSentencia = sSentencia & "HAVING (impmn <> 0.00 OR impme <> 0.00) "
  ElseIf ps_Plataforma = pSrvSql Then
    sSentencia = sSentencia & "HAVING (ROUND(ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impmn ELSE 0 END, 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impmn ELSE 0 END, 0), 2), 2) <> 0.00 "
    sSentencia = sSentencia & "ROUND(ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_DEB & "' THEN det.impme ELSE 0 END, 0), 2) - "
    sSentencia = sSentencia & "ROUND(ISNULL(CASE det.tpoctb WHEN '" & TPOCTB_HAB & "' THEN det.impme ELSE 0 END, 0), 2), 2) <> 0.00) "
  End If
  sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite"
  pocnnMain.Execute sSentencia, nRegistros
  
  'Abro el recordset de seleccion de destinos.
  sSentencia = "SELECT tmp.*, cta.inddoc, cta.indcco, cta.codcco_def "
  sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpdestino tmp, cocta cta "
  sSentencia = sSentencia & "WHERE tmp.codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND tmp.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND cta.codemp=tmp.codemp AND cta.pdoano=tmp.pdoano AND cta.codcta=tmp.codcta "
  sSentencia = sSentencia & "AND cta.tpocta=" & TPOCTA_TRA & " "
  sSentencia = sSentencia & "HAVING (impmn <> 0.00 OR impme <> 0.00) "
  sSentencia = sSentencia & "ORDER BY coddro, nrocpb, nroite"
  With porstCOCpbDst
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  If porstCOCpbDst.RecordCount > 0 Then
    porstCOCpbDst.MoveFirst
    pgbProceso(nIndex).Max = porstCOCpbDst.RecordCount
    pgbProceso(nIndex).Value = pgbProceso(nIndex).Min

    sDiario = ""
    sNroComprobante = ""
    nRegistros = 0
    While Not porstCOCpbDst.EOF
      If Not (sDiario = porstCOCpbDst!coddro And sNroComprobante = porstCOCpbDst!NroCpb) Then
        sDiario = porstCOCpbDst!coddro
        sNroComprobante = porstCOCpbDst!NroCpb
        ' Secuencia final
        With porstUltItem
          If .State = adStateOpen Then .Close
          .ActiveConnection = pocnnMain
          .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroIte), 0) AS nMaxItem "
          .Source = .Source & "FROM cocpbdet "
          .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
          .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
          .Source = .Source & "AND mespvs='" & sPeriodo & "' "
          .Source = .Source & "AND coddro='" & sDiario & "' "
          .Source = .Source & "AND nrocpb='" & sNroComprobante & "' "
          .CursorType = adOpenDynamic
          .LockType = adLockReadOnly
          .Open
          nNroItem = !nMaxItem
          .Close
        End With
      End If
      nLimite = IIf((CDec(porstCOCpbDst!ImpMN) >= 0 And CDec(porstCOCpbDst!ImpME) >= 0) Or (CDec(porstCOCpbDst!ImpMN) <= 0 And CDec(porstCOCpbDst!ImpME) <= 0), 1, 2)
      For nItems = 1 To nLimite
        nNroItem = nNroItem + 1
        ' Adiciono el detalle del comprobante.
        sGrabacion = "INSERT INTO cocpbdet VALUES ("
        sGrabacion = sGrabacion & "'" & gsCodEmp & "', "
        sGrabacion = sGrabacion & "'" & gsAnoAct & "', "
        sGrabacion = sGrabacion & "'" & sPeriodo & "', "
        sGrabacion = sGrabacion & "'" & sDiario & "', "
        sGrabacion = sGrabacion & "'" & sNroComprobante & "', "
        sGrabacion = sGrabacion & nNroItem & ", "
        sGrabacion = sGrabacion & nNroItem & ", "
        ' Tipo de documento
        sExpresion = IIf((IsNull(porstCOCpbDst!codtdc) Or porstCOCpbDst!IndDoc = INDDOC_INA), "", porstCOCpbDst!codtdc)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        sGrabacion = sGrabacion & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(porstCOCpbDst!fehope, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
        sGrabacion = sGrabacion & "'" & porstCOCpbDst!codcta & "', "
        ' Centro de costo
        sExpresion = IIf((IsNull(porstCOCpbDst!codcco) Or porstCOCpbDst!indcco = INDCCO_INA), "", porstCOCpbDst!codcco)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        ' Auxiliar
        sExpresion = IIf((IsNull(porstCOCpbDst!codaux) Or porstCOCpbDst!IndDoc = INDDOC_INA), "", porstCOCpbDst!codaux)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        ' serie
        sExpresion = IIf((IsNull(porstCOCpbDst!serdoc) Or porstCOCpbDst!IndDoc = INDDOC_INA), "", porstCOCpbDst!serdoc)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        ' documento
        sExpresion = IIf((IsNull(porstCOCpbDst!nrodoc) Or porstCOCpbDst!IndDoc = INDDOC_INA), "", porstCOCpbDst!nrodoc)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        sGrabacion = sGrabacion & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(porstCOCpbDst!feedoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
        sGrabacion = sGrabacion & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(porstCOCpbDst!fevdoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
        sGrabacion = sGrabacion & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(smalldatetime, ") & "'" & Format(porstCOCpbDst!ferdoc, "yyyy-mm-dd") & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d'", "120") & "), "
        ' referencia
        sExpresion = IIf(IsNull(porstCOCpbDst!RefDoc), "", porstCOCpbDst!RefDoc)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        ' pedido
        sExpresion = IIf(IsNull(porstCOCpbDst!pdocpr), "", porstCOCpbDst!pdocpr)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        ' glosa
        sExpresion = IIf(IsNull(porstCOCpbDst!GloIte), "", porstCOCpbDst!GloIte)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        ' glosa traduccion
        sExpresion = IIf(IsNull(porstCOCpbDst!GloItex), "", porstCOCpbDst!GloItex)
        sExpresion = IIf(sExpresion = "", "Null", "'" & sExpresion & "'")
        sGrabacion = sGrabacion & sExpresion & ", "
        ' partida doble
        sExpresion = porstCOCpbDst!TpoCtb
        If nLimite = 1 Then
          sExpresion = IIf((CDec(porstCOCpbDst!ImpMN) < 0 Or CDec(porstCOCpbDst!ImpME) < 0), IIf(sExpresion = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB), sExpresion)
        Else
          sExpresion = IIf(Choose(nItems, (CDec(porstCOCpbDst!ImpMN) >= 0), (CDec(porstCOCpbDst!ImpME) >= 0)), sExpresion, IIf(sExpresion = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB))
        End If
        sGrabacion = sGrabacion & "'" & sExpresion & "', "
        sGrabacion = sGrabacion & "'" & TPOPVS_OTR & "', "
        sGrabacion = sGrabacion & "'" & porstCOCpbDst!tpomon & "', "
        sGrabacion = sGrabacion & "'" & porstCOCpbDst!TpoTcb & "', "
        sGrabacion = sGrabacion & "'" & CDec(porstCOCpbDst!ImpTCb) & "', "
        sGrabacion = sGrabacion & Abs(CDec(IIf(nLimite = 1, porstCOCpbDst!ImpMN, Choose(nItems, porstCOCpbDst!ImpMN, 0)))) & ", "
        sGrabacion = sGrabacion & Abs(CDec(IIf(nLimite = 1, porstCOCpbDst!ImpME, Choose(nItems, 0, porstCOCpbDst!ImpME)))) & ", "
        sGrabacion = sGrabacion & TPOGNR_DST & ", "
        sGrabacion = sGrabacion & INDFJO_INA & ", "
        sGrabacion = sGrabacion & INDFJO_INA & ", "
        sGrabacion = sGrabacion & IIf(IsNull(porstCOCpbDst!tpodoc), "Null", "'" & porstCOCpbDst!tpodoc & "'") & ", "
        ' Gerente defaul vacio
        sExpresion = "Null"
        sGrabacion = sGrabacion & sExpresion & ", "
        sGrabacion = sGrabacion & IIf(IsNull(porstCOCpbDst!codcon), "Null", "'" & porstCOCpbDst!codcon & "'") & ", "
        sGrabacion = sGrabacion & IIf(IsNull(porstCOCpbDst!codmon), "Null", "'" & porstCOCpbDst!codmon & "'") & ", "
        sGrabacion = sGrabacion & "'" & gsAbvUsr & "', "
        sGrabacion = sGrabacion & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
        sGrabacion = sGrabacion & "Null, Null, "
        ' rectifica, adicionar, deducible default
        sGrabacion = sGrabacion & "Null, Null, 0)"
        pocnnMain.Execute sGrabacion
      Next nItems
      nRegistros = nRegistros + 1
      pgbProceso(nIndex).Value = nRegistros
      porstCOCpbDst.MoveNext
    Wend
    'Cierro y saco de memoria los recordset.
    Set porstUltItem = Nothing
  End If
  'Cierro y saco de memoria los recordset.
  porstCOCpbDst.Close
  Set porstCOCpbDst = Nothing
End Sub

Private Sub ppMayoriza_Auxiliar(ByVal sPeriodo As String, ByVal nIndex As Integer)
  Dim porstCOCpbDet As ADODB.Recordset
  Dim sSentencia As String, sGrabacion As String
  Dim nLenCuenta As Integer, nProgreso As Integer
  Dim nImporte As Double
  Dim sCadWhere As String, nNivel As Integer
   
  Set porstCOCpbDet = New ADODB.Recordset
   
  ' Inicializo los saldos de cuenta y auxiliar
  pocnnMain.Execute "UPDATE COAuxAcu SET AcuD" & sPeriodo & "_MN=0, AcuH" & sPeriodo & "_MN=0, AcuD" & sPeriodo & "_ME=0, AcuH" & sPeriodo & "_ME=0 WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "'"
  
  ' Inserto las cuentas no existentes
  For nLenCuenta = 1 To Len(gsNivCta)
    nNivel = Mid(gsNivCta, nLenCuenta, 1)
    sSentencia = "INSERT INTO CoAuxAcu (codemp, pdoano, CodCta, CodAux, UsrCre, FyHCre) "
    sSentencia = sSentencia & "SELECT DISTINCT det.codemp, det.pdoano, LEFT(det.CodCta, " & nNivel & "), det.CodAux, '" & gsAbvUsr & "', "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & " "
    sSentencia = sSentencia & "FROM CoCpbDet det "
    sSentencia = sSentencia & "LEFT JOIN CoAuxAcu sal ON det.codemp=sal.codemp AND det.pdoano=sal.pdoano AND LEFT(det.CodCta, " & nNivel & ")=sal.CodCta AND det.CodAux=sal.CodAux "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND det.MesPvs='" & sPeriodo & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.CodCta, '')<>'' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.CodAux, '')<>'' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(sal.CodCta, sal.CodAux), '')", "ISNULL((sal.CodCta+sal.CodAux), '')") & "='' "
    sSentencia = sSentencia & "ORDER BY LEFT(det.CodCta, " & nNivel & "), det.CodAux"
    pocnnMain.Execute sSentencia
  Next nLenCuenta
  pgbProceso(nIndex).Value = 1
   
  ' Genero la sentencia de seleccion de acumulados de cuentas y auxiliares
  sSentencia = "SELECT CodCta, CodAux, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END), 0), 2) AS nDebeMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END), 0), 2) AS nHaberMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END), 0), 2) AS nDebeME, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END), 0), 2) AS nHaberME "
  sSentencia = sSentencia & "FROM COCpbDet "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND MesPvs='" & sPeriodo & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodCta, '') <> '' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodAux, '') <> '' "
  sSentencia = sSentencia & "GROUP BY CodCta, CodAux "
  sSentencia = sSentencia & "ORDER BY CodCta, CodAux"
  ' Abro el recordset de seleccion de las cuenta y auxiliar
  With porstCOCpbDet
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  If porstCOCpbDet.RecordCount > 0 Then
    porstCOCpbDet.MoveFirst
    pgbProceso(nIndex).Max = porstCOCpbDet.RecordCount
    pgbProceso(nIndex).Value = pgbProceso(nIndex).Min
    nProgreso = 0
    Do While Not porstCOCpbDet.EOF
      nImporte = 0
      sCadWhere = ""
      For nLenCuenta = 1 To Len(gsNivCta)
        nNivel = Mid(gsNivCta, nLenCuenta, 1)
        sCadWhere = sCadWhere & "'" & Left(porstCOCpbDet!codcta, nNivel) & "'"
        sCadWhere = sCadWhere & IIf(nLenCuenta = Len(gsNivCta), "", ", ")
      Next nLenCuenta
      ' Actualizo los saldos existentes
      sSentencia = "UPDATE CoAuxAcu SET "
      sSentencia = sSentencia & "AcuD" & sPeriodo & "_MN=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuD" & sPeriodo & "_MN, 0)+" & CDec(porstCOCpbDet!nDebeMN) & ", "
      sSentencia = sSentencia & "AcuH" & sPeriodo & "_MN=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuH" & sPeriodo & "_MN, 0)+" & CDec(porstCOCpbDet!nHaberMN) & ", "
      sSentencia = sSentencia & "AcuD" & sPeriodo & "_ME=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuD" & sPeriodo & "_ME, 0)+" & CDec(porstCOCpbDet!nDebeME) & ", "
      sSentencia = sSentencia & "AcuH" & sPeriodo & "_ME=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuH" & sPeriodo & "_ME, 0)+" & CDec(porstCOCpbDet!nHaberME) & ", "
      sSentencia = sSentencia & "UsrMdf= '" & gsAbvUsr & "', "
      sSentencia = sSentencia & "FyHMdf=" & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & " "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND CodCta IN (" & sCadWhere & ") "
      sSentencia = sSentencia & "AND CodAux='" & porstCOCpbDet!codaux & "'"
      pocnnMain.Execute sSentencia
      nProgreso = nProgreso + 1
      pgbProceso(nIndex).Value = nProgreso
      porstCOCpbDet.MoveNext
    Loop
  End If
  ' Cierro y saco de memoria los recordset
  porstCOCpbDet.Close
  Set porstCOCpbDet = Nothing
End Sub

Private Sub ppMayoriza_CenCosto(ByVal sPeriodo As String, ByVal nIndex As Integer)
  Dim porstCOCpbDet As ADODB.Recordset
  Dim sSentencia As String, sGrabacion As String
  Dim nLenCuenta As Integer, nLenCosto As Integer, nProgreso As Integer
  Dim nNivel As Integer, nNivelCosto As Integer
  Dim sCadWhere As String, sCadWhereCosto As String
  Dim nImporte As Double
   
  Set porstCOCpbDet = New ADODB.Recordset
   
  ' Inicializo los saldos de cuenta y centro de costo
  pocnnMain.Execute "UPDATE COCCoAcu SET AcuD" & sPeriodo & "_MN=0, AcuH" & sPeriodo & "_MN=0, AcuD" & sPeriodo & "_ME=0, AcuH" & sPeriodo & "_ME=0 WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "'"
  
  ' Elimino los registros en cero
  sSentencia = "DELETE FROM CoCCoAcu "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND ((AcuD00_MN + AcuD01_MN + AcuD02_MN + AcuD03_MN + AcuD04_MN + AcuD05_MN + AcuD06_MN + "
  sSentencia = sSentencia & "AcuD07_MN + AcuD08_MN + AcuD09_MN + AcuD10_MN + AcuD11_MN + AcuD12_MN + AcuD13_MN) = 0.00 "
  sSentencia = sSentencia & "AND (AcuH00_MN + AcuH01_MN + AcuH02_MN + AcuH03_MN + AcuH04_MN + AcuH05_MN + AcuH06_MN + "
  sSentencia = sSentencia & "AcuH07_MN + AcuH08_MN + AcuH09_MN + AcuH10_MN + AcuH11_MN + AcuH12_MN + AcuH13_MN) = 0.00 "
  sSentencia = sSentencia & "AND (AcuD00_ME + AcuD01_ME + AcuD02_ME + AcuD03_ME + AcuD04_ME + AcuD05_ME + AcuD06_ME + "
  sSentencia = sSentencia & "AcuD07_ME + AcuD08_ME + AcuD09_ME + AcuD10_ME + AcuD11_ME + AcuD12_ME + AcuD13_ME) = 0.00 "
  sSentencia = sSentencia & "AND (AcuH00_ME + AcuH01_ME + AcuH02_ME + AcuH03_ME + AcuH04_ME + AcuH05_ME + AcuH06_ME + "
  sSentencia = sSentencia & "AcuH07_ME + AcuH08_ME + AcuH09_ME + AcuH10_ME + AcuH11_ME + AcuH12_ME + AcuH13_ME) = 0.00)"
  pocnnMain.Execute sSentencia
  
  ' Inserto las cuentas no existentes
  For nLenCuenta = 1 To Len(gsNivCta)
    nNivel = Mid(gsNivCta, nLenCuenta, 1)
    ' Inserto los centro de costos no existentes
    For nLenCosto = 1 To Len(gsNivCCo)
      nNivelCosto = Mid(gsNivCCo, nLenCosto, 1)
      sSentencia = "INSERT INTO CoCCoAcu (codemp, pdoano, CodCta, CodCCo, UsrCre, FyHCre) "
      sSentencia = sSentencia & "SELECT DISTINCT det.codemp, det.pdoano, LEFT(det.CodCta, " & nNivel & "), LEFT(det.CodCCo, " & nNivelCosto & "), '" & gsAbvUsr & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & " "
      sSentencia = sSentencia & "FROM CoCpbDet det "
      sSentencia = sSentencia & "LEFT JOIN CoCCoAcu sal ON det.codemp=sal.codemp AND det.pdoano=sal.pdoano AND LEFT(det.CodCta, " & nNivel & ")=sal.CodCta AND LEFT(det.CodCCo, " & nNivelCosto & ")=sal.CodCCo "
      sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND det.MesPvs='" & sPeriodo & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.CodCta, '')<>'' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(det.CodCCo, '')<>'' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL(CONCAT(sal.CodCta, sal.CodCCo), '')", "ISNULL((sal.CodCta+sal.CodCCo), '')") & "='' "
      sSentencia = sSentencia & "ORDER BY LEFT(det.CodCta, " & nNivel & "), LEFT(det.CodCCo, " & nNivelCosto & ")"
      pocnnMain.Execute sSentencia
    Next nLenCosto
  Next nLenCuenta
  pgbProceso(nIndex).Value = 1
  ' Genero la sentencia de seleccion de acumulados de cuenta y centro de costo
  sSentencia = "SELECT CodCta, CodCco, "
  sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END), 0)), 2) AS nDebeMN, "
  sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END), 0)), 2) AS nHaberMN, "
  sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END), 0)), 2) AS nDebeME, "
  sSentencia = sSentencia & "ROUND(SUM(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "((CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END), 0)), 2) AS nHaberME "
  sSentencia = sSentencia & "FROM COCpbDet "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND MesPvs='" & sPeriodo & "' "
  sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(CodCco, '') <> '' "
  sSentencia = sSentencia & "GROUP BY CodCta, CodCco "
  sSentencia = sSentencia & "ORDER BY CodCta, CodCco"
  ' Abro el recordset de seleccion de las cuenta y centro de costo
  With porstCOCpbDet
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  If porstCOCpbDet.RecordCount > 0 Then
    porstCOCpbDet.MoveFirst
    pgbProceso(nIndex).Max = porstCOCpbDet.RecordCount
    pgbProceso(nIndex).Value = pgbProceso(nIndex).Min
    nProgreso = 0
    Do While Not porstCOCpbDet.EOF
      nImporte = 0
      sCadWhere = "": sCadWhereCosto = ""
      ' Obtengo las cuentas a actualizar
      For nLenCuenta = 1 To Len(gsNivCta)
        nNivel = Mid(gsNivCta, nLenCuenta, 1)
        sCadWhere = sCadWhere & "'" & Left(porstCOCpbDet!codcta, nNivel) & "'"
        sCadWhere = sCadWhere & IIf(nLenCuenta = Len(gsNivCta), "", ", ")
      Next nLenCuenta
      ' Obtengo las centro de costos a actualizar
      For nLenCosto = 1 To Len(gsNivCCo)
        nNivelCosto = Mid(gsNivCCo, nLenCosto, 1)
        sCadWhereCosto = sCadWhereCosto & "'" & Left(porstCOCpbDet!codcco, nNivelCosto) & "'"
        sCadWhereCosto = sCadWhereCosto & IIf(nLenCosto = Len(gsNivCCo), "", ", ")
      Next nLenCosto
      
      ' Actualizo los saldos existentes
      sSentencia = "UPDATE CoCCoAcu SET "
      sSentencia = sSentencia & "AcuD" & sPeriodo & "_MN=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuD" & sPeriodo & "_MN, 0)+" & CDec(porstCOCpbDet!nDebeMN) & ", "
      sSentencia = sSentencia & "AcuH" & sPeriodo & "_MN=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuH" & sPeriodo & "_MN, 0)+" & CDec(porstCOCpbDet!nHaberMN) & ", "
      sSentencia = sSentencia & "AcuD" & sPeriodo & "_ME=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuD" & sPeriodo & "_ME, 0)+" & CDec(porstCOCpbDet!nDebeME) & ", "
      sSentencia = sSentencia & "AcuH" & sPeriodo & "_ME=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuH" & sPeriodo & "_ME, 0)+" & CDec(porstCOCpbDet!nHaberME) & ", "
      sSentencia = sSentencia & "UsrMdf= '" & gsAbvUsr & "', "
      sSentencia = sSentencia & "FyHMdf=" & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & " "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND CodCta IN (" & sCadWhere & ") "
      sSentencia = sSentencia & "AND CodCCo IN (" & sCadWhereCosto & ")"
      pocnnMain.Execute sSentencia
      nProgreso = nProgreso + 1
      pgbProceso(nIndex).Value = nProgreso
      porstCOCpbDet.MoveNext
    Loop
  End If
  ' Cierro y saco de memoria el recordset de detalle
  porstCOCpbDet.Close
  Set porstCOCpbDet = Nothing
  
End Sub

Private Sub ppMayoriza_Cuenta(ByVal sPeriodo As String, ByVal nIndex As Integer)
  Static porstCOCpbDet As ADODB.Recordset
  
  Dim sSentencia As String, sGrabacion As String
  Dim nLenCuenta As Integer, nProgreso As Integer
  Dim nImporte As Double
  Dim sCadWhere As String, nNivel As Integer
   
  Set porstCOCpbDet = New ADODB.Recordset
   
  ' Inicializo los saldos de las cuentas
  pocnnMain.Execute "UPDATE COCtaAcu SET AcuD" & sPeriodo & "_MN=0, AcuH" & sPeriodo & "_MN=0, AcuD" & sPeriodo & "_ME=0, AcuH" & sPeriodo & "_ME=0 WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "'"
  ' Inserto las cuentas no existentes
  For nLenCuenta = 1 To Len(gsNivCta)
    nNivel = Mid(gsNivCta, nLenCuenta, 1)
    
    sSentencia = "INSERT INTO CoCtaAcu (codemp, pdoano, CodCta, UsrCre, FyHCre) "
    sSentencia = sSentencia & "SELECT DISTINCT det.codemp, det.pdoano, LEFT(det.CodCta, " & nNivel & "), '" & gsAbvUsr & "', "
    sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & " "
    sSentencia = sSentencia & "FROM CoCpbDet det "
    sSentencia = sSentencia & "LEFT JOIN CoCtaAcu sal ON det.codemp=sal.codemp AND det.pdoano=sal.pdoano AND LEFT(det.CodCta, " & nNivel & ")=sal.CodCta "
    sSentencia = sSentencia & "WHERE det.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND det.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND det.MesPvs='" & sPeriodo & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(sal.CodCta, '')='' "
    sSentencia = sSentencia & "ORDER BY LEFT(det.CodCta, " & nNivel & ")"
    
    sqlCuenta = ""
    For nProgreso = nLenCuenta To Len(gsNivCta)
      sqlCuenta = sqlCuenta & "SELECT DISTINCT det.codemp, det.pdoano, det.mespvs, LEFT(det.CodCta, " & Mid(gsNivCta, nProgreso, 1) & ") AS codcta "
      sqlCuenta = sqlCuenta & "FROM cocpbdet det "
      sqlCuenta = sqlCuenta & "WHERE det.codemp='" & gsCodEmp & "' "
      sqlCuenta = sqlCuenta & "AND det.pdoano='" & gsAnoAct & "' "
      sqlCuenta = sqlCuenta & "AND det.mespvs='" & sPeriodo & "' "
      sqlCuenta = sqlCuenta & "AND NOT EXISTS (SELECT * FROM cocta cta WHERE cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=LEFT(det.CodCta, " & Mid(gsNivCta, nProgreso, 1) & ")) "
      sqlCuenta = sqlCuenta & IIf(nProgreso = Len(gsNivCta), "", "UNION ")
    Next nProgreso
    sqlCuenta = sqlCuenta & "ORDER BY codcta"
    lVisualizar = True
    pocnnMain.Execute sSentencia
  Next nLenCuenta
  
  lVisualizar = False
  
  pgbProceso(nIndex).Value = 1
   
  ' Genero la sentencia de seleccion de acumulados de las cuentas
  sSentencia = "SELECT CodCta, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpMN ELSE 0 END), 0), 2) AS nDebeMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpMN ELSE 0 END), 0), 2) AS nHaberMN, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_DEB & "' THEN ImpME ELSE 0 END), 0), 2) AS nDebeME, "
  sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE TpoCtb WHEN '" & TPOCTB_HAB & "' THEN ImpME ELSE 0 END), 0), 2) AS nHaberME "
  sSentencia = sSentencia & "FROM COCpbDet "
  sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND MesPvs='" & sPeriodo & "' "
  sSentencia = sSentencia & "GROUP BY CodCta "
  sSentencia = sSentencia & "ORDER BY CodCta"
  ' Abro el recordset de seleccion de las cuentas
  With porstCOCpbDet
    If .State = adStateOpen Then .Close
    .ActiveConnection = pocnnMain
    .Source = sSentencia
    .CursorType = adOpenDynamic
    .LockType = adLockReadOnly
    .Open
  End With
  If porstCOCpbDet.RecordCount > 0 Then
    porstCOCpbDet.MoveFirst
    pgbProceso(nIndex).Max = porstCOCpbDet.RecordCount
    pgbProceso(nIndex).Value = pgbProceso(nIndex).Min
    nProgreso = 0
    Do While Not porstCOCpbDet.EOF
      nImporte = 0
      sCadWhere = ""
      For nLenCuenta = 1 To Len(gsNivCta)
        nNivel = Mid(gsNivCta, nLenCuenta, 1)
        sCadWhere = sCadWhere & "'" & Left(porstCOCpbDet!codcta, nNivel) & "'"
        sCadWhere = sCadWhere & IIf(nLenCuenta = Len(gsNivCta), "", ", ")
      Next nLenCuenta
      ' Actualizo los saldos existentes
      sSentencia = "UPDATE CoCtaAcu SET "
      sSentencia = sSentencia & "AcuD" & sPeriodo & "_MN=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuD" & sPeriodo & "_MN, 0)+" & CDec(porstCOCpbDet!nDebeMN) & ", "
      sSentencia = sSentencia & "AcuH" & sPeriodo & "_MN=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuH" & sPeriodo & "_MN, 0)+" & CDec(porstCOCpbDet!nHaberMN) & ", "
      sSentencia = sSentencia & "AcuD" & sPeriodo & "_ME=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuD" & sPeriodo & "_ME, 0)+" & CDec(porstCOCpbDet!nDebeME) & ", "
      sSentencia = sSentencia & "AcuH" & sPeriodo & "_ME=" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(AcuH" & sPeriodo & "_ME, 0)+" & CDec(porstCOCpbDet!nHaberME) & ", "
      sSentencia = sSentencia & "UsrMdf= '" & gsAbvUsr & "', "
      sSentencia = sSentencia & "FyHMdf=" & IIf(ps_Plataforma = pSrvMySql, "", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "'" & IIf(ps_Plataforma = pSrvMySql, "", ", 120)") & " "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND CodCta IN (" & sCadWhere & ")"
      pocnnMain.Execute sSentencia
      nProgreso = nProgreso + 1
      pgbProceso(nIndex).Value = nProgreso
      porstCOCpbDet.MoveNext
    Loop
  End If
  ' Cierro y saco de memoria los recordset
  porstCOCpbDet.Close
  Set porstCOCpbDet = Nothing

End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   txtDato(Index).SelStart = 0
   txtDato(Index).SelLength = txtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYUDAT, Index
   End If
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Cancel = ppAyuDet(AYUDAT, Index)
   If Cancel Then Exit Sub
   cmdDatoAyud(0).Enabled = True
   cmdAceptar.Enabled = True
   Exit Sub
End Sub

