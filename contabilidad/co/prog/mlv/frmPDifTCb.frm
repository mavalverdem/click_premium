VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPDifTCb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[título]"
   ClientHeight    =   4725
   ClientLeft      =   3135
   ClientTop       =   1590
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4650
   Begin VB.CheckBox chkprocesar 
      Caption         =   "Procesar Hasta el Periodo"
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   0
      Width           =   3255
   End
   Begin VB.TextBox txtDato 
      Height          =   300
      Index           =   1
      Left            =   90
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2580
      Width           =   465
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   4155
      Picture         =   "frmPDifTCb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2580
      Width           =   255
   End
   Begin VB.TextBox txtDato 
      Height          =   300
      Index           =   2
      Left            =   90
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3165
      Width           =   465
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   2
      Left            =   4155
      Picture         =   "frmPDifTCb.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3165
      Width           =   255
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   4185
      Picture         =   "frmPDifTCb.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   255
   End
   Begin VB.TextBox txtDato 
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Text            =   "9201"
      Top             =   450
      Width           =   465
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   345
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   1170
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   450
      Left            =   2040
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   450
      Left            =   3360
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   345
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   1890
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   345
      Index           =   2
      Left            =   105
      TabIndex        =   16
      Top             =   3825
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
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
      Height          =   300
      Index           =   1
      Left            =   510
      TabIndex        =   12
      Top             =   2580
      Width           =   3660
   End
   Begin VB.Label lblTexto 
      Caption         =   "Flujo de Caja Ganancia :"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   3
      Left            =   90
      TabIndex        =   19
      Top             =   2325
      Width           =   2800
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
      Height          =   300
      Index           =   2
      Left            =   510
      TabIndex        =   13
      Top             =   3165
      Width           =   3660
   End
   Begin VB.Label lblTexto 
      Caption         =   "Flujo de Caja Perdida :"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   2910
      Width           =   2800
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ajuste por Cuenta + Auxiliar"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   5
      Left            =   105
      TabIndex        =   17
      Top             =   3585
      Width           =   2800
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ajuste por Cuenta"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   1620
      Width           =   2800
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ajuste por Documento"
      ForeColor       =   &H80000002&
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   900
      Width           =   2800
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
      Height          =   300
      Index           =   0
      Left            =   540
      TabIndex        =   11
      Top             =   450
      Width           =   3660
   End
   Begin VB.Label lblTexto 
      Caption         =   "Ingrese Diario"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   180
      Width           =   2800
   End
End
Attribute VB_Name = "frmPDifTCb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCoTCbMes As ADODB.Recordset
Public porstCodro As ADODB.Recordset
Public porstCoFjo As ADODB.Recordset

Dim cnn As ADODB.Connection
Private Sub Form_Load()

'ini 2015-05-18 validacion frm
    If gsMesAct = "00" Or gsMesAct = "13" Then
        MsgBox TEXT_9018, vbCritical
        'Me.Hide
'        End
        'Return
'        Me.Visible = False
        pExitForm = 1 'permite salir del Form_Load sin que salga error de conflicto de proc al entrar 2 veces
        'Unload Me
        Exit Sub
    End If
'fin 2015-05-18 validacion frm

  pgbProceso(0).Value = 0
  pgbProceso(1).Value = 0
  
  chkprocesar.Caption = "Procesar Hasta el Periodo: " & gsAnoAct & gsMesAct
 
  Set cnn = New ADODB.Connection
  If ps_Puerto = "" Then
     cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";connection="
  Else
     cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & gsNomBDS & ";Port=" & ps_Puerto & ";connection="
  End If
  cnn.CursorLocation = adUseClient
  cnn.Open
  
  
  'Abrir Tablas.
   
  Set pocnnMain = New ADODB.Connection
  Set porstCoTCbMes = New ADODB.Recordset
  Set porstCodro = New ADODB.Recordset
  Set porstCoFjo = New ADODB.Recordset

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
   
  With porstCoFjo
      .ActiveConnection = pocnnMain
      .Source = "SELECT CodFjo, DetFjo "
      .Source = .Source & "FROM COFjo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodFjo)=4 "
      .Source = .Source & "ORDER BY CodFjo"
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
  End With
   
  With porstCoTCbMes
      .ActiveConnection = pocnnMain
      .Source = "SELECT ImpTCb_Cpr, ImpTCb_Vta "
      .Source = .Source & "FROM COTCbMes "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND MesPvs='" & gsMesAct & "'"
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
      .Open
        'ini 2015-05-18 validacion frm
        If .EOF() Then
            MsgBox TEXT_9019, vbCritical
            pExitForm = 1 'permite salir del Form_Load sin que salga error de conflicto de proc al entrar 2 veces
            Exit Sub
        End If
        'fin 2015-05-18 validacion frm
   
  End With
   
  '[ Cargo los mensajes de botones
   Dim nElemento As Integer
   ReDim aLabel(6, 2)
   For nElemento = 0 To UBound(aLabel, 1) - 1
     aLabel(nElemento, 0) = Choose(nElemento + 1, "Ingrese Diario", "Ajuste por Documento", "Ajuste por Cuenta", "Flujo de Caja Ganancia", "Flujo de Caja Perdida", "Ajuste por Cuenta + Auxiliar")
     aLabel(nElemento, 1) = Choose(nElemento + 1, "Enter Journal", "Adjustment for Document", "Adjustment for Account", "Cash Flow Profit", "Cash Flow Loss", "Adjustment for Account + Auxiliary")
   Next nElemento
   cmdAceptar.Caption = Choose(gsIdioma, "&Procesar", "&Process")
   CaptionBotones Me, False, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']
  
  'Modificado para que solicite el tipo de cambio diferente de o  TC
  ' If porstCoTCbMes.RecordCount = 0 Then
   
   If porstCoTCbMes.RecordCount = 0 Or (porstCoTCbMes!ImpTCb_Cpr = 0#) Or (porstCoTCbMes!ImpTCb_Vta = 0#) Then
        MsgBox Choose(gsIdioma, "Para Procesar la Diferencia de Cambio del Cierre de Mes, debe Ingresar T/Cambio del Fin de Mes", "It will Process the Closing Difference of Exchange, if you must Enter R/Exchange of End Month"), vbCritical
        Exit Sub
   End If
   
End Sub

Private Sub Form_Activate()
   cmdAceptar.Enabled = False
   cmdSalir.Enabled = True
   cmdSalir.SetFocus
End Sub



Private Sub Form_Unload(Cancel As Integer)
'error cuando salen sin usar estos rs y conex
'ini 2015-05-18 validacion frm
'''   porstCoTCbMes.Close
'''   porstCodro.Close
'''   porstCoFjo.Close
'''   pocnnMain.Close
'''
'''   Set porstCoTCbMes = Nothing
'''   Set porstCodro = Nothing
'''   Set porstCoFjo = Nothing
'''   Set pocnnMain = Nothing
'fin 2015-05-18 validacion frm
fRstClose porstCoTCbMes
fRstClose porstCodro
fRstClose porstCoFjo
fCnnClose pocnnMain

End Sub

Private Sub cmdAceptar_Click()

  Dim i As Integer
  Dim dnContador As Integer
  On Error GoTo Err
   
   'Verificación de Mes Cerrado.
  If gbCieCpb Then
    MsgBox TEXT_9016, vbCritical
    Exit Sub
  End If
   
 '[Propio del formulario.
  For dnContador = 0 To txtDato.Count - 1
    If txtDato.Item(dnContador).Text = "" Then
      If dnContador = 0 Then
        MsgBox TEXT_6002, vbCritical
        txtDato(dnContador).SetFocus
        Exit Sub
      End If
    End If
  Next dnContador
   
   If gnIndMNE <> INDMNE_ACT Then
      MsgBox Choose(gsIdioma, "La Empresa trabaja sólo con una Moneda", "The Company works only one Currency"), vbExclamation
      Exit Sub
   End If
   
  If chkprocesar.Value = Checked Then
    If gsMesAct = "00" Then gsMesAct = "01"
    If gsMesAct = "13" Then gsMesAct = "12"
    
    'ini 2015-06-24 control flag mayoriza
    If gcCierre(gsAnoAct, gsMesAct) = 1 Then Exit Sub
    'fin 2015-06-24 control flag mayoriza
   
    For i = 1 To Int(gsMesAct)
      gsMesAct = Format(i, "00")
      porstCoTCbMes.Close
      With porstCoTCbMes
        .ActiveConnection = cnn
        .Source = "SELECT ImpTCb_Cpr, ImpTCb_Vta "
        .Source = .Source & "FROM COTCbMes "
        .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
        .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
        .Source = .Source & "AND MesPvs='" & gsMesAct & "'"
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
      End With
    
      If porstCoTCbMes.RecordCount = 0 Or (porstCoTCbMes!ImpTCb_Cpr = 0#) Or (porstCoTCbMes!ImpTCb_Vta = 0#) Then
        MsgBox Choose(gsIdioma, "Para Procesar la Diferencia de Cambio del Cierre de Mes, debe Ingresar T/Cambio del Fin de Mes", "It will Process the Closing Difference of Exchange, if you must Enter R/Exchange of End Month"), vbCritical
        Exit Sub
      End If
      pgbProceso(0).Value = 0: pgbProceso(0).Min = 0
      pgbProceso(1).Value = 0: pgbProceso(1).Min = 0
      pgbProceso(2).Value = 0: pgbProceso(2).Min = 0
      
      pocnnMain.BeginTrans                'INICIA TRANSACCION.
      
      'Paso 1 : Elimino los comprobantes de ajuste del mes
      pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DCA) & " AND MesPvs='" & gsMesAct & "'"
      'Paso 2 : Generacion de Ajustes por Documento
      ppAjuste_Documento
      'Paso 3 : Generacion de Ajustes por Cuenta
      ppAjuste_SaldoCuenta
      'Paso 4 : Generacion de Ajustes por Cuenta
      ppAjuste_Auxiliar
      
      pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
      MsgBox TEXT_8008 & " Periodo: " & gsAnoAct & gsMesAct, vbInformation
    Next
  Else
    pgbProceso(0).Value = 0: pgbProceso(0).Min = 0
    pgbProceso(1).Value = 0: pgbProceso(1).Min = 0
    pgbProceso(2).Value = 0: pgbProceso(2).Min = 0
       
    pocnnMain.BeginTrans                'INICIA TRANSACCION.
    
    'Paso 1 : Elimino los comprobantes de ajuste del mes
    pocnnMain.Execute "DELETE FROM COCpbCab WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND TpoGnr=" & Str(TPOGNR_DCA) & " AND MesPvs='" & gsMesAct & "'"
    'Paso 2 : Generacion de Ajustes por Documento
    ppAjuste_Documento
    'Paso 3 : Generacion de Ajustes por Cuenta
    ppAjuste_SaldoCuenta
    'Paso 4 : Generacion de Ajustes por Cuenta
    ppAjuste_Auxiliar
    
    pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
    MsgBox TEXT_8008 & " Periodo: " & gsAnoAct & gsMesAct, vbInformation
   End If
   
  Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1, 2
      ppAyuBus AYUDAT, Index
      txtDato(Index).SetFocus
   End Select
End Sub

Private Sub cmdSalir_Click()
   Unload Me
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
'   On Error GoTo Err
   Dim dbSalir As Boolean
   Dim dvRegistro As Variant
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Cancel = ppAyuDet(AYUDAT, Index)
   If Cancel Then Exit Sub
   cmdAceptar.Enabled = True
   
   Exit Sub
'Err:
'   gpErrores
End Sub
Private Sub ppAjuste_SaldoCuenta()
   Dim porstCOCpbCab As ADODB.Recordset
   Dim porstCOCpbDet As ADODB.Recordset
   Dim porstCoCpbDetFjo As ADODB.Recordset
   Dim porstCoCpbAjD As ADODB.Recordset
   Dim porstUltCoCpb  As ADODB.Recordset
   
   Dim sNroComprobante As String, sTpoTcb As String, sCodFjo As String
   Dim nNroItem As Long, nContador As Integer
   
   Dim sCenCosto As String, sCodCCo_Ajd As String
   Dim sTpoCtb_Ajd As String, sTpoMon_Ajd As String
   Dim aCodCta_Ajd(), sCodCta_Ajd As String
   Dim nImpTCb_Ajd As Double, nImporte As Double
   Dim nImpor_Ajd As Double
   Dim nImpMN_Ajd As Double, nImpME_Ajd As Double
   Dim nImpMN_Sal As Double, nImpME_Sal As Double
   
   Set porstCOCpbCab = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   Set porstCoCpbDetFjo = New ADODB.Recordset
   Set porstCoCpbAjD = New ADODB.Recordset
   Set porstUltCoCpb = New ADODB.Recordset
   
   pgbProceso(1).Min = 0
   ' Abro el recordset de seleccion de destinos
   With porstCoCpbAjD
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      'Genero la sentencia de seleccion saldos de cuentas pendientes
      .Source = "SELECT a.CodCta, b.TpoMon, b.TpoTcb, b.NatCta, b.IndFjo, "
      .Source = .Source & "b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab, "
      .Source = .Source & "b.CodCCo_Def, b.CodCCo_Ajd_Deb, b.CodCCo_Ajd_Hab, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) AS nImpMN_Deb, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2) AS nImpMN_Hab, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) AS nImpME_Deb, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2) AS nImpME_Hab "
      .Source = .Source & "FROM COCpbDet a, CoCta b "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND b.codemp=a.codemp "
      .Source = .Source & "AND b.pdoano=a.pdoano "
      .Source = .Source & "AND b.CodCta=a.CodCta "
      .Source = .Source & "AND a.MesPvs<='" & gsMesAct & "' "
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & "AND b.IndAjd='" & INDAJD_ACT & "' "
      .Source = .Source & "AND b.TpoAnl='" & TPOANL_CTA & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodCta, '')<>'' "
      .Source = .Source & "AND (" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Ajd_Deb, '')<>'' OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Ajd_Hab, '')<>'') "
      .Source = .Source & "GROUP BY a.CodCta, b.TpoMon, b.TpoTcb, b.NatCta, b.IndFjo, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab, b.CodCCo_Def, b.CodCCo_Ajd_Deb, b.CodCCo_Ajd_Hab "
      If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "HAVING (ROUND(nImpMN_Deb - nImpMN_Hab, 2) <> 0.00) OR (ROUND(nImpME_Deb - nImpME_Hab, 2) <> 0.00) "
      ElseIf ps_Plataforma = pSrvSql Then
        .Source = .Source & "HAVING (ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) - "
        .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2), 2)<> 0.00) OR "
        .Source = .Source & "(ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) - "
        .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2), 2) <> 0.00) "
      End If
      .Source = .Source & "ORDER BY b.CodCta_Ajd_Deb, b.CodCCo_Ajd_Deb, a.CodCta"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
      .Close
      .Open
   End With
   If Not (porstCoCpbAjD.EOF Or porstCoCpbAjD.BOF) Then
      porstCoCpbAjD.MoveFirst
      pgbProceso(1).Max = porstCoCpbAjD.RecordCount
      pgbProceso(1).Value = pgbProceso(1).Min
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbCab
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion cabecera de comprobantes
         .Source = "SELECT codemp, pdoano, CodDro, NroCpb, FehCpb, GloCpb, GloCpbx, MesPvs, "
         .Source = .Source & "TpoGnr, IndNCu, IndAnu, "
         .Source = .Source & "UsrCre, FyHCre "
         .Source = .Source & "FROM COCpbCab "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ' Obtengo el numero e inserto la cabecera del comprobante
      With porstUltCoCpb
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), 0) AS cUltNroCpb "
         .Source = .Source & "FROM COCpbCab "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND MesPvs='" & gsMesAct & "' "
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
         !mespvs = gsMesAct
         !coddro = txtDato(0).Text
         !NroCpb = sNroComprobante
         !FehCpb = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
         !glocpb = "Ajustes por Diferencia de Cambio Cuenta Cierre de Mes " & gsMesAct
         !glocpbx = "Adjustment by Difference of Exchange Account Closing " & gsMesAct
         !tpognr = TPOGNR_DCA
         !IndNCu = INDNCU_FAL
         !IndAnu = INDANU_FAL
         !UsrCre = gsAbvUsr
         !FyHCre = Now
         .Update
      End With
      nNroItem = 0
      
      ' Abro el recordset de grabacion del detalle de comprobante
      With porstCOCpbDet
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion detalles de comprobantes
         .Source = "SELECT codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, "
         .Source = .Source & "SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, gloite, gloitex, TpoCtb, TpoPvs, IndFjo_Det, "
         .Source = .Source & "TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, "
         .Source = .Source & "UsrCre, FyHCre "
'ini 2016-06-24 correcion codmon asto destino y dif cam
         .Source = .Source & ",Codmon "
'fin 2016-06-24 correcion codmon asto destino y dif cam
         .Source = .Source & "FROM COCpbDet "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
     ' Abro el recordset de grabacion detalle de flujo de caja
      With porstCoCpbDetFjo
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion detalles de comprobantes
         .Source = "SELECT codemp, pdoano, MesPvs, CodDro, NroCpb, NroIte, NroOrd, CodFjo, "
         .Source = .Source & "CodCta, TpoCtb, ImpMN, ImpME, UsrCre, FyHCre "
         .Source = .Source & "FROM CoCpbDetFjo "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      
      ReDim aCodCta_Ajd(7, 0)
      Do While Not porstCoCpbAjD.EOF
         ' Calculos para determinar el ajuste
         nNroItem = nNroItem + 1
         nImpMN_Sal = gfRedond(porstCoCpbAjD!nImpMN_Deb - porstCoCpbAjD!nImpMN_Hab, 2)
         nImpME_Sal = gfRedond(porstCoCpbAjD!nImpME_Deb - porstCoCpbAjD!nImpME_Hab, 2)
         nImpor_Ajd = 0
         nImpME_Ajd = 0
         nImpMN_Ajd = 0
         sCodFjo = ""
         sTpoTcb = porstCoCpbAjD!TpoTcb
         If porstCoCpbAjD!NatCta = NATCTA_DEU Then
            sTpoTcb = IIf((IIf(porstCoCpbAjD!tpomon = TPOMON_NAC, nImpMN_Sal, nImpME_Sal) > 0), sTpoTcb, IIf(sTpoTcb = TPOTCB_CPR, TPOTCB_VTA, TPOTCB_CPR))
         Else
            sTpoTcb = IIf((IIf(porstCoCpbAjD!tpomon = TPOMON_NAC, nImpMN_Sal, nImpME_Sal) < 0), sTpoTcb, IIf(sTpoTcb = TPOTCB_CPR, TPOTCB_VTA, TPOTCB_CPR))
         End If
         nImpTCb_Ajd = IIf(sTpoTcb = TPOTCB_VTA, porstCoTCbMes!ImpTCb_Vta, porstCoTCbMes!ImpTCb_Cpr)
         
         
         If nImpTCb_Ajd > 0 And (nImpMN_Sal <> 0 Or nImpME_Sal <> 0) Then
            If porstCoCpbAjD!tpomon = TPOMON_EXT Then
               If nImpME_Sal > 0 Then
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpMN_Deb - (porstCoCpbAjD!nImpMN_Hab + (Abs(nImpME_Sal) * nImpTCb_Ajd))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpMN_Hab - (porstCoCpbAjD!nImpMN_Deb + (Abs(nImpME_Sal) * nImpTCb_Ajd))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(nImpME_Sal * nImpTCb_Ajd, 2)
               If nImporte <> nImpMN_Sal Then
                  nImpor_Ajd = gfRedond(nImporte - nImpMN_Sal, 2)
                  nImpMN_Ajd = Abs(nImpor_Ajd)
                  nImpME_Ajd = gfRedond(IIf(nImpTCb_Ajd = 0, 0, nImpMN_Ajd / nImpTCb_Ajd), 2)
               End If
            Else
               If nImpMN_Sal > 0 Then
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpME_Deb - (porstCoCpbAjD!nImpME_Hab + IIf(nImpTCb_Ajd = 0, 0, Abs(nImpMN_Sal) / nImpTCb_Ajd))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpME_Hab - (porstCoCpbAjD!nImpME_Deb + IIf(nImpTCb_Ajd = 0, 0, Abs(nImpMN_Sal) / nImpTCb_Ajd))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(IIf(nImpTCb_Ajd = 0, 0, nImpMN_Sal / nImpTCb_Ajd), 2)
               If nImporte <> nImpME_Sal Then
                  nImpor_Ajd = gfRedond(nImporte - nImpME_Sal, 2)
                  nImpME_Ajd = Abs(nImpor_Ajd)
                  nImpMN_Ajd = gfRedond(nImpME_Ajd * nImpTCb_Ajd, 2)
               End If
            End If
            If gfRedond(nImpor_Ajd, 2) <> 0 Then
               ' Adiciono el detalle del comprobante cuenta de documentos
               sTpoMon_Ajd = IIf(porstCoCpbAjD!tpomon = TPOMON_EXT, TPOMON_NAC, TPOMON_EXT)
               sCenCosto = IIf(IsNull(porstCoCpbAjD!codcco_def), "", porstCoCpbAjD!codcco_def)
               sCodFjo = IIf(porstCoCpbAjD!IndFjo = INDFJO_ACT, IIf(sTpoCtb_Ajd = TPOCTB_DEB, txtDato(1), txtDato(2)), "")
               ppInsDetalle_Cpb porstCOCpbDet, porstCoCpbDetFjo, sNroComprobante, nNroItem, sCodFjo, porstCoCpbAjD!codcta, sCenCosto, "", "", "", "", sTpoCtb_Ajd, sTpoMon_Ajd, sTpoTcb, IIf(IsNull(nImpTCb_Ajd), 1, nImpTCb_Ajd), nImpMN_Ajd, nImpME_Ajd
               
               sCodCta_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, porstCoCpbAjD!CodCta_AjD_Deb, porstCoCpbAjD!CodCta_AjD_Hab)
               sCodCCo_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, IIf(IsNull(porstCoCpbAjD!CodCCo_AjD_Deb), "", porstCoCpbAjD!CodCCo_AjD_Deb), IIf(IsNull(porstCoCpbAjD!CodCCo_AjD_Hab), "", porstCoCpbAjD!CodCCo_AjD_Hab))
               sTpoCtb_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
               sCodCCo_Ajd = IIf(IsNull(sCodCCo_Ajd), "", sCodCCo_Ajd)
               For nContador = 1 To UBound(aCodCta_Ajd, 2)
                  ' Verifico los datos de la cuenta de ajuste
                  If aCodCta_Ajd(1, nContador) = sCodCta_Ajd And aCodCta_Ajd(2, nContador) = sTpoCtb_Ajd And aCodCta_Ajd(3, nContador) = sTpoMon_Ajd Then
                     Exit For
                  End If
               Next nContador
               If nContador > UBound(aCodCta_Ajd, 2) Then
                  ReDim Preserve aCodCta_Ajd(7, UBound(aCodCta_Ajd, 2) + 1)
               End If
               aCodCta_Ajd(1, nContador) = sCodCta_Ajd
               aCodCta_Ajd(2, nContador) = sCodCCo_Ajd
               aCodCta_Ajd(3, nContador) = sTpoCtb_Ajd
               aCodCta_Ajd(4, nContador) = sTpoMon_Ajd
               aCodCta_Ajd(5, nContador) = IIf(IsNull(nImpTCb_Ajd), 1, nImpTCb_Ajd)
               aCodCta_Ajd(6, nContador) = gfRedond(aCodCta_Ajd(6, nContador) + nImpMN_Ajd, 2)
               aCodCta_Ajd(7, nContador) = gfRedond(aCodCta_Ajd(7, nContador) + nImpME_Ajd, 2)
            End If
         End If
         pgbProceso(1).Value = nNroItem
         porstCoCpbAjD.MoveNext
      Loop
      ' Adiciono el detalle del comprobante cuenta (perdidad o ganacia)
      For nContador = 1 To UBound(aCodCta_Ajd, 2)
         nNroItem = nNroItem + 1
         sCodCta_Ajd = aCodCta_Ajd(1, nContador)
         sCodCCo_Ajd = aCodCta_Ajd(2, nContador)
         sTpoCtb_Ajd = aCodCta_Ajd(3, nContador)
         sTpoMon_Ajd = aCodCta_Ajd(4, nContador)
         nImpTCb_Ajd = aCodCta_Ajd(5, nContador)
         nImpMN_Ajd = aCodCta_Ajd(6, nContador)
         nImpME_Ajd = aCodCta_Ajd(7, nContador)
         ppInsDetalle_Cpb porstCOCpbDet, porstCoCpbDetFjo, sNroComprobante, nNroItem, "", sCodCta_Ajd, sCodCCo_Ajd, "", "", "", "", sTpoCtb_Ajd, sTpoMon_Ajd, "V", nImpTCb_Ajd, nImpMN_Ajd, nImpME_Ajd
      Next nContador
      porstCOCpbDet.UpdateBatch
      porstCoCpbDetFjo.UpdateBatch
      ' Cierro y saco de memoria los recordset
      porstCoCpbDetFjo.Close
      porstCOCpbDet.Close
      porstCOCpbCab.Close
      Set porstCoCpbDetFjo = Nothing
      Set porstCOCpbDet = Nothing
      Set porstCOCpbCab = Nothing
      Set porstUltCoCpb = Nothing
   End If
   ' Cierro y saco de memoria los recordset
   porstCoCpbAjD.Close
   Set porstCoCpbAjD = Nothing
End Sub
Private Sub ppAjuste_Documento()
   Dim porstCOCpbCab As ADODB.Recordset
   Dim porstCOCpbDet As ADODB.Recordset
   Dim porstCoCpbAjD As ADODB.Recordset
   Dim porstUltCoCpb  As ADODB.Recordset
   
   Dim sNroComprobante As String
   Dim nNroItem As Long, nContador As Integer
   
   Dim sCenCosto As String, sCodCCo_Ajd As String
   Dim sTpoCtb_Ajd As String, sTpoMon_Ajd As String
   Dim aCodCta_Ajd(), sCodCta_Ajd As String
   Dim nImpTCb_Ajd As Double, nImporte As Double
   Dim nImpor_Ajd As Double
   Dim nImpMN_Ajd As Double, nImpME_Ajd As Double
   Dim nImpMN_Sal As Double, nImpME_Sal As Double
   
   Set porstCOCpbCab = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   Set porstCoCpbAjD = New ADODB.Recordset
   Set porstUltCoCpb = New ADODB.Recordset
   
   pgbProceso(0).Min = 0
   ' Abro el recordset de seleccion de destinos
   With porstCoCpbAjD
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      'Genero la sentencia de seleccion documentos pendientes
      .Source = "SELECT a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, b.TpoMon, "
      .Source = .Source & "b.TpoTcb, b.IndFjo, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab, "
      .Source = .Source & "b.CodCCo_Def, b.CodCCo_Ajd_Deb, b.CodCCo_Ajd_Hab, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) AS nImpMN_Deb, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2) AS nImpMN_Hab, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) AS nImpME_Deb, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2) AS nImpME_Hab "
      .Source = .Source & "FROM COCpbDet a, CoCta b "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND b.codemp=a.codemp "
      .Source = .Source & "AND b.pdoano=a.pdoano "
      .Source = .Source & "AND b.CodCta=a.CodCta "
      .Source = .Source & "AND a.MesPvs<='" & gsMesAct & "' "
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & "AND b.IndAjd='" & INDAJD_ACT & "' "
      .Source = .Source & "AND b.IndDoc='" & INDDOC_ACT & "' "
      .Source = .Source & "AND b.TpoAnl='" & TPOANL_DOC & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodTDc, '')<>'' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.SerDoc, '')<>'' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.NroDoc, '')<>'' "
      .Source = .Source & "AND (" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Ajd_Deb, '')<>'' OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Ajd_Hab, '')<>'') "
      .Source = .Source & "GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, b.TpoMon, b.TpoTcb, b.IndFjo, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab, b.CodCCo_Def, b.CodCCo_Ajd_Deb, b.CodCCo_Ajd_Hab "
      If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "HAVING (ROUND(nImpMN_Deb - nImpMN_Hab, 2) <> 0.00) OR (ROUND(nImpME_Deb - nImpME_Hab, 2) <> 0.00) "
      ElseIf ps_Plataforma = pSrvSql Then
        .Source = .Source & "HAVING (ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) - "
        .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2), 2)<> 0.00) OR "
        .Source = .Source & "(ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) - "
        .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2), 2) <> 0.00) "
      End If
      .Source = .Source & "ORDER BY b.CodCta_Ajd_Deb, b.CodCCo_Ajd_Deb, a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc "
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   If porstCoCpbAjD.RecordCount > 0 Then
      porstCoCpbAjD.MoveFirst
      pgbProceso(0).Max = porstCoCpbAjD.RecordCount
      pgbProceso(0).Value = pgbProceso(0).Min
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbCab
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion cabecera de comprobantes
         .Source = "SELECT codemp, pdoano, MesPvs, CodDro, NroCpb, FehCpb, glocpb, glocpbx, "
         .Source = .Source & "TpoGnr, IndNCu, IndAnu, "
         .Source = .Source & "UsrCre, FyHCre "
         .Source = .Source & "FROM COCpbCab "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ' Obtengo el numero e inserto la cabecera del comprobante
      With porstUltCoCpb
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), 0) AS cUltNroCpb "
         .Source = .Source & "FROM COCpbCab "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND MesPvs='" & gsMesAct & "' "
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
         !mespvs = gsMesAct
         !coddro = txtDato(0).Text
         !NroCpb = sNroComprobante
         !FehCpb = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
         !glocpb = "Ajustes por Diferencia de Cambio Documento Cierre de Mes " & gsMesAct
         !glocpbx = "Adjustment by Difference of Exchange Document Closing " & gsMesAct
         !tpognr = TPOGNR_DCA
         !IndNCu = INDNCU_FAL
         !IndAnu = INDANU_FAL
         !UsrCre = gsAbvUsr
         !FyHCre = Now
         .Update
      End With
      nNroItem = 0
      
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbDet
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion detalles de comprobantes
         .Source = "SELECT codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, "
         .Source = .Source & "SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, gloitex, TpoCtb, TpoPvs, IndFjo_Det, "
         .Source = .Source & "TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, "
         .Source = .Source & "UsrCre, FyHCre "
'ini 2016-06-24 correcion codmon asto destino y dif cam
         .Source = .Source & ", Codmon "
'fin 2016-06-24 correcion codmon asto destino y dif cam
         .Source = .Source & "FROM COCpbDet "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ReDim aCodCta_Ajd(7, 0)
      Do While Not porstCoCpbAjD.EOF
         ' Calculos para determinar el ajuste
         nNroItem = nNroItem + 1
         nImpTCb_Ajd = IIf(porstCoCpbAjD!TpoTcb = TPOTCB_VTA, porstCoTCbMes!ImpTCb_Vta, porstCoTCbMes!ImpTCb_Cpr)
         nImpMN_Sal = gfRedond(porstCoCpbAjD!nImpMN_Deb - porstCoCpbAjD!nImpMN_Hab, 2)
         nImpME_Sal = gfRedond(porstCoCpbAjD!nImpME_Deb - porstCoCpbAjD!nImpME_Hab, 2)
         nImpor_Ajd = 0
         nImpME_Ajd = 0
         nImpMN_Ajd = 0
         If nImpTCb_Ajd > 0 And (nImpMN_Sal <> 0 Or nImpME_Sal <> 0) Then
            If porstCoCpbAjD!tpomon = TPOMON_EXT Then
               If nImpME_Sal > 0 Then
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpMN_Deb - (porstCoCpbAjD!nImpMN_Hab + (Abs(nImpME_Sal) * nImpTCb_Ajd))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpMN_Hab - (porstCoCpbAjD!nImpMN_Deb + (Abs(nImpME_Sal) * nImpTCb_Ajd))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(nImpME_Sal * nImpTCb_Ajd, 2)
               If nImporte <> nImpMN_Sal Then
                  nImpor_Ajd = gfRedond(nImporte - nImpMN_Sal, 2)
                  nImpMN_Ajd = Abs(nImpor_Ajd)
                  nImpME_Ajd = gfRedond(IIf(nImpTCb_Ajd = 0, 0, nImpMN_Ajd / nImpTCb_Ajd), 2)
               End If
            Else
               If nImpMN_Sal > 0 Then
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpME_Deb - (porstCoCpbAjD!nImpME_Hab + IIf(nImpTCb_Ajd = 0, 0, Abs(nImpMN_Sal) / nImpTCb_Ajd))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpME_Hab - (porstCoCpbAjD!nImpME_Deb + IIf(nImpTCb_Ajd = 0, 0, Abs(nImpMN_Sal) / nImpTCb_Ajd))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(IIf(nImpTCb_Ajd = 0, 0, nImpMN_Sal / nImpTCb_Ajd), 2)
               If nImporte <> nImpME_Sal Then
                  nImpor_Ajd = gfRedond(nImporte - nImpME_Sal, 2)
                  nImpME_Ajd = Abs(nImpor_Ajd)
                  nImpMN_Ajd = gfRedond(nImpME_Ajd * nImpTCb_Ajd, 2)
               End If
            End If
            If nImpor_Ajd <> 0 Then
               ' Adiciono el detalle del comprobante cuenta de documentos
               sTpoMon_Ajd = IIf(porstCoCpbAjD!tpomon = TPOMON_EXT, TPOMON_NAC, TPOMON_EXT)
               sCenCosto = IIf(IsNull(porstCoCpbAjD!codcco_def), "", porstCoCpbAjD!codcco_def)
               ppInsDetalle_Cpb porstCOCpbDet, porstCOCpbDet, sNroComprobante, nNroItem, "", porstCoCpbAjD!codcta, sCenCosto, porstCoCpbAjD!codtdc, porstCoCpbAjD!codaux, porstCoCpbAjD!serdoc, porstCoCpbAjD!nrodoc, sTpoCtb_Ajd, sTpoMon_Ajd, porstCoCpbAjD!TpoTcb, IIf(IsNull(nImpTCb_Ajd), 1, nImpTCb_Ajd), nImpMN_Ajd, nImpME_Ajd
               
               sCodCta_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, porstCoCpbAjD!CodCta_AjD_Deb, porstCoCpbAjD!CodCta_AjD_Hab)
               sCodCCo_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, IIf(IsNull(porstCoCpbAjD!CodCCo_AjD_Deb), "", porstCoCpbAjD!CodCCo_AjD_Deb), IIf(IsNull(porstCoCpbAjD!CodCCo_AjD_Hab), "", porstCoCpbAjD!CodCCo_AjD_Hab))
               sTpoCtb_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
               sCodCCo_Ajd = IIf(IsNull(sCodCCo_Ajd), "", sCodCCo_Ajd)
               For nContador = 1 To UBound(aCodCta_Ajd, 2)
                  ' Verifico los datos de la cuenta de ajuste
                  If aCodCta_Ajd(1, nContador) = sCodCta_Ajd And aCodCta_Ajd(2, nContador) = sCodCCo_Ajd And aCodCta_Ajd(3, nContador) = sTpoCtb_Ajd And aCodCta_Ajd(4, nContador) = sTpoMon_Ajd Then
                     Exit For
                  End If
               Next nContador
               If nContador > UBound(aCodCta_Ajd, 2) Then
                  ReDim Preserve aCodCta_Ajd(7, UBound(aCodCta_Ajd, 2) + 1)
               End If
               aCodCta_Ajd(1, nContador) = sCodCta_Ajd
               aCodCta_Ajd(2, nContador) = sCodCCo_Ajd
               aCodCta_Ajd(3, nContador) = sTpoCtb_Ajd
               aCodCta_Ajd(4, nContador) = sTpoMon_Ajd
               aCodCta_Ajd(5, nContador) = IIf(IsNull(nImpTCb_Ajd), 1, nImpTCb_Ajd)
               aCodCta_Ajd(6, nContador) = gfRedond(aCodCta_Ajd(6, nContador) + nImpMN_Ajd, 2)
               aCodCta_Ajd(7, nContador) = gfRedond(aCodCta_Ajd(7, nContador) + nImpME_Ajd, 2)
            End If
         End If
         pgbProceso(0).Value = nNroItem
         porstCoCpbAjD.MoveNext
      Loop
      ' Adiciono el detalle del comprobante cuenta de documentos(perdidad o ganacia)
      For nContador = 1 To UBound(aCodCta_Ajd, 2)
         nNroItem = nNroItem + 1
         sCodCta_Ajd = aCodCta_Ajd(1, nContador)
         sCodCCo_Ajd = aCodCta_Ajd(2, nContador)
         sTpoCtb_Ajd = aCodCta_Ajd(3, nContador)
         sTpoMon_Ajd = aCodCta_Ajd(4, nContador)
         nImpTCb_Ajd = aCodCta_Ajd(5, nContador)
         nImpMN_Ajd = aCodCta_Ajd(6, nContador)
         nImpME_Ajd = aCodCta_Ajd(7, nContador)
         ppInsDetalle_Cpb porstCOCpbDet, porstCOCpbDet, sNroComprobante, nNroItem, "", sCodCta_Ajd, sCodCCo_Ajd, "", "", "", "", sTpoCtb_Ajd, sTpoMon_Ajd, "V", nImpTCb_Ajd, nImpMN_Ajd, nImpME_Ajd
      Next nContador
      porstCOCpbDet.UpdateBatch
      ' Cierro y saco de memoria los recordset
      porstCOCpbDet.Close
      porstCOCpbCab.Close
      Set porstCOCpbDet = Nothing
      Set porstCOCpbCab = Nothing
      Set porstUltCoCpb = Nothing
   End If
   ' Cierro y saco de memoria los recordset
   porstCoCpbAjD.Close
   Set porstCoCpbAjD = Nothing
End Sub
Private Sub ppAjuste_Auxiliar()
   Static porstCOCpbCab As ADODB.Recordset
   Static porstCOCpbDet As ADODB.Recordset
   Static porstCoCpbAjD As ADODB.Recordset
   Static porstUltCoCpb  As ADODB.Recordset
   
   Dim sNroComprobante As String
   Dim nNroItem As Long, nContador As Integer
   
   Dim sCenCosto As String, sCodCCo_Ajd As String
   Dim sTpoCtb_Ajd As String, sTpoMon_Ajd As String
   Dim aCodCta_Ajd(), sCodCta_Ajd As String
   Dim nImpTCb_Ajd As Double, nImporte As Double
   Dim nImpor_Ajd As Double
   Dim nImpMN_Ajd As Double, nImpME_Ajd As Double
   Dim nImpMN_Sal As Double, nImpME_Sal As Double
   
   Set porstCOCpbCab = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   Set porstCoCpbAjD = New ADODB.Recordset
   Set porstUltCoCpb = New ADODB.Recordset
   
   pgbProceso(2).Min = 0
   ' Abro el recordset de seleccion de destinos
   With porstCoCpbAjD
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      'Genero la sentencia de seleccion documentos pendientes
      .Source = "SELECT a.CodCta, a.CodAux, b.TpoMon, "
      .Source = .Source & "b.TpoTcb, b.IndFjo, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab, "
      .Source = .Source & "b.CodCCo_Def, b.CodCCo_Ajd_Deb, b.CodCCo_Ajd_Hab, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) AS nImpMN_Deb, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2) AS nImpMN_Hab, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) AS nImpME_Deb, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2) AS nImpME_Hab "
      .Source = .Source & "FROM COCpbDet a, CoCta b "
      .Source = .Source & "WHERE a.codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND a.pdoano='" & gsAnoAct & "' "
      .Source = .Source & "AND b.codemp=a.codemp "
      .Source = .Source & "AND b.pdoano=a.pdoano "
      .Source = .Source & "AND b.CodCta=a.CodCta "
      .Source = .Source & "AND a.MesPvs<='" & gsMesAct & "' "
      .Source = .Source & "AND b.TpoCta='" & TPOCTA_TRA & "' "
      .Source = .Source & "AND b.IndAjd='" & INDAJD_ACT & "' "
      .Source = .Source & "AND b.IndDoc='" & INDAUX_ACT & "' "
      .Source = .Source & "AND b.TpoAnl='" & TPOANL_AUX & "' "
      .Source = .Source & "AND " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(a.CodAux, '')<>'' "
      .Source = .Source & "AND (" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Ajd_Deb, '')<>'' OR " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(b.CodCta_Ajd_Hab, '')<>'') "
      .Source = .Source & "GROUP BY a.CodCta, a.CodAux, b.TpoMon, b.TpoTcb, b.IndFjo, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab, b.CodCCo_Def, b.CodCCo_Ajd_Deb, b.CodCCo_Ajd_Hab "
      If ps_Plataforma = pSrvMySql Then
        .Source = .Source & "HAVING (ROUND(nImpMN_Deb - nImpMN_Hab, 2) <> 0.00) OR (ROUND(nImpME_Deb - nImpME_Hab, 2) <> 0.00) "
      ElseIf ps_Plataforma = pSrvSql Then
        .Source = .Source & "HAVING (ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpMN ELSE 0 END), 0), 2) - "
        .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpMN ELSE 0 END), 0), 2), 2)<> 0.00) OR "
        .Source = .Source & "(ROUND(ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_DEB & "' THEN a.ImpME ELSE 0 END), 0), 2) - "
        .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(CASE a.TpoCtb WHEN '" & TPOCTB_HAB & "' THEN a.ImpME ELSE 0 END), 0), 2), 2) <> 0.00) "
      End If
      .Source = .Source & "ORDER BY b.CodCta_Ajd_Deb, b.CodCCo_Ajd_Deb, a.CodCta, a.CodAux"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
      .Close
      .Open
   End With
   If porstCoCpbAjD.RecordCount > 0 Then
      porstCoCpbAjD.MoveFirst
      pgbProceso(2).Max = porstCoCpbAjD.RecordCount
      pgbProceso(2).Value = pgbProceso(2).Min
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbCab
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion cabecera de comprobantes
         .Source = "SELECT codemp, pdoano, CodDro, NroCpb, FehCpb, GloCpb, glocpbx, MesPvs, "
         .Source = .Source & "TpoGnr, IndNCu, IndAnu, "
         .Source = .Source & "UsrCre, FyHCre "
         .Source = .Source & "FROM COCpbCab "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ' Obtengo el numero e inserto la cabecera del comprobante
      With porstUltCoCpb
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroCpb), 0) AS cUltNroCpb "
         .Source = .Source & "FROM COCpbCab "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND MesPvs='" & gsMesAct & "' "
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
         !mespvs = gsMesAct
         !coddro = txtDato(0).Text
         !NroCpb = sNroComprobante
         !FehCpb = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
         !glocpb = "Ajustes por Diferencia de Cambio Auxiliar Cierre de Mes " & gsMesAct
         !glocpbx = "Adjustment by Difference of Exchange Auxiliary Closing " & gsMesAct
         !tpognr = TPOGNR_DCA
         !IndNCu = INDNCU_FAL
         !IndAnu = INDANU_FAL
         !UsrCre = gsAbvUsr
         !FyHCre = Now
         .Update
      End With
      nNroItem = 0
      
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbDet
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion detalles de comprobantes
         .Source = "SELECT codemp, pdoano, CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, "
         .Source = .Source & "SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, gloitex, TpoCtb, TpoPvs, IndFjo_Det, "
         .Source = .Source & "TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, "
         .Source = .Source & "UsrCre, FyHCre "
'ini 2016-06-24 correcion codmon asto destino y dif cam
         .Source = .Source & ",Codmon "
'fin 2016-06-24 correcion codmon asto destino y dif cam
         .Source = .Source & "FROM COCpbDet "
         .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
         .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
         .Source = .Source & "AND CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ReDim aCodCta_Ajd(7, 0)
      Do While Not porstCoCpbAjD.EOF
         ' Calculos para determinar el ajuste
         nNroItem = nNroItem + 1
         nImpTCb_Ajd = IIf(porstCoCpbAjD!TpoTcb = TPOTCB_VTA, porstCoTCbMes!ImpTCb_Vta, porstCoTCbMes!ImpTCb_Cpr)
         nImpMN_Sal = gfRedond(porstCoCpbAjD!nImpMN_Deb - porstCoCpbAjD!nImpMN_Hab, 2)
         nImpME_Sal = gfRedond(porstCoCpbAjD!nImpME_Deb - porstCoCpbAjD!nImpME_Hab, 2)
         nImpor_Ajd = 0
         nImpME_Ajd = 0
         nImpMN_Ajd = 0
         If nImpTCb_Ajd > 0 And (nImpMN_Sal <> 0 Or nImpME_Sal <> 0) Then
            If porstCoCpbAjD!tpomon = TPOMON_EXT Then
               If nImpME_Sal > 0 Then
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpMN_Deb - (porstCoCpbAjD!nImpMN_Hab + (Abs(nImpME_Sal) * nImpTCb_Ajd))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpMN_Hab - (porstCoCpbAjD!nImpMN_Deb + (Abs(nImpME_Sal) * nImpTCb_Ajd))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(nImpME_Sal * nImpTCb_Ajd, 2)
               If nImporte <> nImpMN_Sal Then
                  nImpor_Ajd = gfRedond(nImporte - nImpMN_Sal, 2)
                  nImpMN_Ajd = Abs(nImpor_Ajd)
                  nImpME_Ajd = gfRedond(IIf(nImpTCb_Ajd = 0, 0, nImpMN_Ajd / nImpTCb_Ajd), 2)
               End If
            Else
               If nImpMN_Sal > 0 Then
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpME_Deb - (porstCoCpbAjD!nImpME_Hab + IIf(nImpTCb_Ajd = 0, 0, Abs(nImpMN_Sal) / nImpTCb_Ajd))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_Ajd = IIf((porstCoCpbAjD!nImpME_Hab - (porstCoCpbAjD!nImpME_Deb + IIf(nImpTCb_Ajd = 0, 0, Abs(nImpMN_Sal) / nImpTCb_Ajd))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(IIf(nImpTCb_Ajd = 0, 0, nImpMN_Sal / nImpTCb_Ajd), 2)
               If nImporte <> nImpME_Sal Then
                  nImpor_Ajd = gfRedond(nImporte - nImpME_Sal, 2)
                  nImpME_Ajd = Abs(nImpor_Ajd)
                  nImpMN_Ajd = gfRedond(nImpME_Ajd * nImpTCb_Ajd, 2)
               End If
            End If
            
            If nImpor_Ajd <> 0 Then
               ' Adiciono el detalle del comprobante cuenta de documentos
               sTpoMon_Ajd = IIf(porstCoCpbAjD!tpomon = TPOMON_EXT, TPOMON_NAC, TPOMON_EXT)
               sCenCosto = IIf(IsNull(porstCoCpbAjD!codcco_def), "", porstCoCpbAjD!codcco_def)
               ppInsDetalle_Cpb porstCOCpbDet, porstCOCpbDet, sNroComprobante, nNroItem, "", porstCoCpbAjD!codcta, sCenCosto, "", porstCoCpbAjD!codaux, "", "", sTpoCtb_Ajd, sTpoMon_Ajd, porstCoCpbAjD!TpoTcb, IIf(IsNull(nImpTCb_Ajd), 1, nImpTCb_Ajd), nImpMN_Ajd, nImpME_Ajd
               
               sCodCta_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, porstCoCpbAjD!CodCta_AjD_Deb, porstCoCpbAjD!CodCta_AjD_Hab)
               sCodCCo_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, IIf(IsNull(porstCoCpbAjD!CodCCo_AjD_Deb), "", porstCoCpbAjD!CodCCo_AjD_Deb), IIf(IsNull(porstCoCpbAjD!CodCCo_AjD_Hab), "", porstCoCpbAjD!CodCCo_AjD_Hab))
               sTpoCtb_Ajd = IIf(sTpoCtb_Ajd = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
               sCodCCo_Ajd = IIf(IsNull(sCodCCo_Ajd), "", sCodCCo_Ajd)
               For nContador = 1 To UBound(aCodCta_Ajd, 2)
                  ' Verifico los datos de la cuenta de ajuste
                  If aCodCta_Ajd(1, nContador) = sCodCta_Ajd And aCodCta_Ajd(2, nContador) = sCodCCo_Ajd And aCodCta_Ajd(3, nContador) = sTpoCtb_Ajd And aCodCta_Ajd(4, nContador) = sTpoMon_Ajd Then
                     Exit For
                  End If
               Next nContador
               If nContador > UBound(aCodCta_Ajd, 2) Then
                  ReDim Preserve aCodCta_Ajd(7, UBound(aCodCta_Ajd, 2) + 1)
               End If
               aCodCta_Ajd(1, nContador) = sCodCta_Ajd
               aCodCta_Ajd(2, nContador) = sCodCCo_Ajd
               aCodCta_Ajd(3, nContador) = sTpoCtb_Ajd
               aCodCta_Ajd(4, nContador) = sTpoMon_Ajd
               aCodCta_Ajd(5, nContador) = IIf(IsNull(nImpTCb_Ajd), 1, nImpTCb_Ajd)
               aCodCta_Ajd(6, nContador) = gfRedond(aCodCta_Ajd(6, nContador) + nImpMN_Ajd, 2)
               aCodCta_Ajd(7, nContador) = gfRedond(aCodCta_Ajd(7, nContador) + nImpME_Ajd, 2)
            End If
         End If
         pgbProceso(2).Value = nNroItem
         porstCoCpbAjD.MoveNext
      Loop
      ' Adiciono el detalle del comprobante cuenta de documentos(perdidad o ganacia)
      For nContador = 1 To UBound(aCodCta_Ajd, 2)
         nNroItem = nNroItem + 1
         sCodCta_Ajd = aCodCta_Ajd(1, nContador)
         sCodCCo_Ajd = aCodCta_Ajd(2, nContador)
         sTpoCtb_Ajd = aCodCta_Ajd(3, nContador)
         sTpoMon_Ajd = aCodCta_Ajd(4, nContador)
         nImpTCb_Ajd = aCodCta_Ajd(5, nContador)
         nImpMN_Ajd = aCodCta_Ajd(6, nContador)
         nImpME_Ajd = aCodCta_Ajd(7, nContador)
         ppInsDetalle_Cpb porstCOCpbDet, porstCOCpbDet, sNroComprobante, nNroItem, "", sCodCta_Ajd, sCodCCo_Ajd, "", "", "", "", sTpoCtb_Ajd, sTpoMon_Ajd, "V", nImpTCb_Ajd, nImpMN_Ajd, nImpME_Ajd
      Next nContador
      porstCOCpbDet.UpdateBatch
      ' Cierro y saco de memoria los recordset
      porstCOCpbDet.Close
      porstCOCpbCab.Close
      Set porstCOCpbDet = Nothing
      Set porstCOCpbCab = Nothing
      Set porstUltCoCpb = Nothing
   End If
   ' Cierro y saco de memoria los recordset
   porstCoCpbAjD.Close
   Set porstCoCpbAjD = Nothing

End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYUDAT Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4 ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case 1, 2                          'Cambiar (añadir índices).
         modAyuBus.Fjo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodFjo)=4 ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
'               cmdAceptar.Enabled = True
'               cmdSalir.Enabled = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
            End If
         End With
      Case 1, 2
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With porstCoFjo
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodFjo='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & !DetFjo
            End If
         End With
      End Select
   End If
End Function

Private Sub ppInsDetalle_Cpb(porstCOCpbDet As ADODB.Recordset, porstCoCpbDetFjo As ADODB.Recordset, cNroCpb As String, nNroIte As Long, sCodFlujo As String, cCodCta As String, cCodCCo As String, cCodTDc As String, cCodAux As String, cSerDoc As String, cNroDoc As String, cTpoCtb As String, cTpoMon As String, cTpoTcb As String, nImpTCb As Double, nImpMN As Double, nImpME As Double)
  
  Static INDMASFJO_INI As Byte, INDMASFJO_MAS As Byte
  
  INDMASFJO_INI = 0: INDMASFJO_MAS = 1

  ' Adiciono el detalle del comprobante
  With porstCOCpbDet
    .AddNew
    !codemp = gsCodEmp
    !pdoano = gsAnoAct
    !mespvs = gsMesAct
    !coddro = txtDato(0).Text
    !NroCpb = cNroCpb
    !NroIte = nNroIte
    !blqite = nNroIte
    !codcta = cCodCta
    !fehope = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
    !feedoc = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
    !fevdoc = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
    !ferdoc = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
    !codtdc = IIf(cCodTDc = "", Null, cCodTDc)
    !codcco = IIf(cCodCCo = "", Null, cCodCCo)
    !codaux = IIf(cCodAux = "", Null, cCodAux)
    !serdoc = IIf(cSerDoc = "", Null, cSerDoc)
    !nrodoc = IIf(cNroDoc = "", Null, cNroDoc)
    !GloIte = "Ajuste Diferencia de Cambio Cierre Mes"
    !GloItex = "Adjustment by Difference of Exchange Closing"
    !TpoCtb = cTpoCtb
    !tpomon = cTpoMon
'ini 2016-06-24 correcion codmon asto destino y dif cam
    !codmon = IIf(cTpoMon = TPOMON_NAC, CODMON_NAC, CODMON_EXT)
'fin 2016-06-24 correcion codmon asto destino y dif cam
    !TpoTcb = cTpoTcb
    !ImpTCb = nImpTCb
    !ImpMN = IIf(porstCOCpbDet!tpomon = TPOMON_NAC, nImpMN, 0)
    !ImpME = IIf(porstCOCpbDet!tpomon = TPOMON_EXT, nImpME, 0)
    !TpoPvs = TPOPVS_OTR
    !tpognr = TPOGNR_DCA
    !indfjo_det = IIf(sCodFlujo = "", INDMASFJO_INI, INDMASFJO_MAS)
    !UsrCre = gsAbvUsr
    !FyHCre = Now
  End With
  ' Verifico si tiene flujo de caja
  If sCodFlujo <> "" Then
    ' adiciono detalle de flujo de caja
    With porstCoCpbDetFjo
      .AddNew
      !codemp = gsCodEmp
      !pdoano = gsAnoAct
      !mespvs = gsMesAct
      !coddro = txtDato(0).Text
      !NroCpb = cNroCpb
      !NroIte = nNroIte
      !NroOrd = 1
      !CodFjo = sCodFlujo
      !codcta = cCodCta
      !TpoCtb = cTpoCtb
      !ImpMN = IIf(cTpoMon = TPOMON_NAC, nImpMN, 0)
      !ImpME = IIf(cTpoMon = TPOMON_EXT, nImpME, 0)
      !UsrCre = gsAbvUsr
      !FyHCre = Now
    End With
  End If

End Sub
