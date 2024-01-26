VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPDifTCb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[título]"
   ClientHeight    =   3630
   ClientLeft      =   2925
   ClientTop       =   2700
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   375
      Index           =   0
      Left            =   4185
      Picture         =   "frmPDifTCb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   450
      Width           =   255
   End
   Begin VB.TextBox TxtDato 
      Height          =   330
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
      TabIndex        =   4
      Top             =   1260
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   893
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   2573
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProceso 
      Height          =   345
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   1980
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
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Ajuste por Cuenta + Auxiliar"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   2595
   End
   Begin VB.Label Label3 
      Caption         =   "Ajuste por Cuenta"
      ForeColor       =   &H80000002&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   1710
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Ajuste por Documento"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   990
      Width           =   1635
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
      Height          =   360
      Index           =   0
      Left            =   540
      TabIndex        =   7
      Top             =   450
      Width           =   3660
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese Diario"
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmPDifTCb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCOTCbMes As ADODB.Recordset
Public porstCodro As ADODB.Recordset

Private Sub Form_Load()
   pgbProceso(0).Value = 0
   pgbProceso(1).Value = 0
  
  'Abrir Tablas.
   
   Set pocnnMain = New ADODB.Connection
   Set porstCOTCbMes = New ADODB.Recordset
   Set porstCodro = New ADODB.Recordset
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstCodro
      .ActiveConnection = pocnnMain
      .Source = "Select CodDro, DetDro From CODro Where Length(CodDro)=4"
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   With porstCOTCbMes
      .ActiveConnection = pocnnMain
      .Source = "SELECT ImpTCb_Cpr, ImpTCb_Vta FROM COTCbMes " _
              & "WHERE MesPvs='" & gsMesAct & "'"
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
      .Open
   End With
   
   If porstCOTCbMes.RecordCount = 0 Then
      MsgBox "Para Procesar la Diferencia de Cambio del Cierre de Mes, debe Ingresar T/Cambio del Fin de Mes", vbCritical
      Exit Sub
   End If
   
End Sub

Private Sub Form_Activate()
   cmdAceptar.Enabled = False
   cmdSalir.Enabled = True
   cmdSalir.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   porstCOTCbMes.Close
   porstCodro.Close
   pocnnMain.Close
   Set porstCOTCbMes = Nothing
   Set porstCodro = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdAceptar_Click()
   On Error GoTo Err
   'Verificación de Mes Cerrado.
   If gbCieCpb Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
 '[Propio del formulario.
   If txtDato(0).Text = "" Then
      MsgBox TEXT_6002, vbCritical
      txtDato(0).SetFocus
      Exit Sub
   End If
   
   If gnIndMNE <> INDMNE_ACT Then
      MsgBox "La Empresa trabaja sólo con una Moneda", vbExclamation
      Exit Sub
   End If
   pgbProceso(0).Value = 0: pgbProceso(0).Min = 0
   pgbProceso(1).Value = 0: pgbProceso(1).Min = 0
   pgbProceso(2).Value = 0: pgbProceso(2).Min = 0
   
   pocnnMain.BeginTrans                'INICIA TRANSACCION.
  
  'Paso 1 : Elimino los comprobantes de ajuste del mes
   pocnnMain.Execute "DELETE FROM COCpbCab WHERE TpoGnr=" & Str(TPOGNR_DCA) & " And MesPvs=" & gsMesAct
  'Paso 2 : Generacion de Ajustes por Documento
   ppAjuste_Documento
  'Paso 3 : Generacion de Ajustes por Cuenta
   ppAjuste_SaldoCuenta
  'Paso 4 : Generacion de Ajustes por Cuenta
   ppajuste_Auxiliar

   pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
   
   MsgBox TEXT_8008, vbInformation
  
   Exit Sub
Err:
  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      ppAyuBus AYUDAT, Index
      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
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
   cmdDatoAyud(0).Enabled = True
   cmdAceptar.Enabled = True
   Exit Sub
'Err:
'   gpErrores
End Sub
Private Sub ppAjuste_SaldoCuenta()
   Static porstCOCpbCab As ADODB.Recordset
   Static porstCOCpbDet As ADODB.Recordset
   Static porstCOCpbAjD As ADODB.Recordset
   Static porstUltCoCpb  As ADODB.Recordset
   
   Static sNroComprobante As String, sTpoTcb As String
   Static nNroItem As Integer, nContador As Integer
   
   Static sTpoCtb_AjD As String, sTpoMon_AjD As String
   Static aCodCta_AjD(), sCodCta_AjD As String
   Static nImpTCb_AjD As Double, nImporte As Double
   Static nImpor_AjD As Double
   Static nImpMN_AjD As Double, nImpME_AjD As Double
   Static nImpMN_Sal As Double, nImpME_Sal As Double
   
   Set porstCOCpbCab = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   Set porstCOCpbAjD = New ADODB.Recordset
   Set porstUltCoCpb = New ADODB.Recordset
   
   pgbProceso(1).Min = 0
   ' Abro el recordset de seleccion de destinos
   With porstCOCpbAjD
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      'Genero la sentencia de seleccion saldos de cuentas pendientes
      .Source = "SELECT a.CodCta, b.TpoMon, b.TpoTcb, b.NatCta, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='D', a.ImpMN, 0)), 0), 2) AS nImpMN_Deb," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='H', a.ImpMN, 0)), 0), 2) AS nImpMN_Hab," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='D', a.ImpME, 0)), 0), 2) AS nImpME_Deb," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='H', a.ImpME, 0)), 0), 2) AS nImpME_Hab " _
              & "FROM COCpbDet a, CoCta b " _
              & "WHERE a.CodCta=b.CodCta" _
              & "  AND a.MesPvs<='" & gsMesAct & "'" _
              & "  AND b.TpoCta='" & TPOCTA_TRA & "'" _
              & "  AND b.IndAjd='" & INDAJD_ACT & "'" _
              & "  AND b.TpoAnl='" & TPOANL_CTA & "' " _
              & "  AND (IFNULL(b.CodCta_Ajd_Deb, '')<>'' OR IFNULL(b.CodCta_Ajd_Hab, '')<>'') " _
              & "GROUP BY a.CodCta " _
              & "HAVING (ROUND(nImpMN_Deb - nImpMN_Hab, 2) <> 0.00) OR (ROUND(nImpME_Deb - nImpME_Hab, 2) <> 0.00) " _
              & "ORDER BY b.CodCta_Ajd_Deb, a.CodCta"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   If porstCOCpbAjD.RecordCount > 0 Then
      porstCOCpbAjD.MoveFirst
      pgbProceso(1).Max = porstCOCpbAjD.RecordCount
      pgbProceso(1).Value = pgbProceso(1).Min
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbCab
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion cabecera de comprobantes
         .Source = "SELECT CodDro, NroCpb, FehCpb, GloCpb, MesPvs," _
                 & "  TpoGnr, IndNCu, IndAnu," _
                 & "  UsrCre, FyHCre " _
                 & "FROM COCpbCab " _
                 & "WHERE CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ' Obtengo el numero e inserto la cabecera del comprobante
      With porstUltCoCpb
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = "SELECT IFNULL(MAX(NroCpb), 0) AS cUltNroCpb FROM COCpbCab WHERE CodDro='" & txtDato(0).Text & "' AND MesPvs='" & gsMesAct & "'"
         .CursorType = adOpenDynamic
         .LockType = adLockReadOnly
         .Open
         sNroComprobante = !cUltNroCpb
         .Close
      End With
      sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
      With porstCOCpbCab
         .AddNew
         !CodDro = txtDato(0).Text
         !NroCpb = sNroComprobante
         !FehCpb = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
         !GloCpb = "Ajustes por Diferencia de Cambio Cuenta Cierre de Mes " & gsMesAct
         !MesPvs = gsMesAct
         !TpoGnr = TPOGNR_DCA
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
         .Source = "SELECT CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux," _
                 & "  SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, TpoCtb, TpoPvs," _
                 & "  TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr," _
                 & "  UsrCre, FyHCre " _
                 & "FROM COCpbDet " _
                 & "WHERE CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ReDim aCodCta_AjD(6, 0)
      Do While Not porstCOCpbAjD.EOF
         ' Calculos para determinar el ajuste
         nNroItem = nNroItem + 1
         nImpMN_Sal = gfRedond(porstCOCpbAjD!nImpMN_Deb - porstCOCpbAjD!nImpMN_Hab, 2)
         nImpME_Sal = gfRedond(porstCOCpbAjD!nImpME_Deb - porstCOCpbAjD!nImpME_Hab, 2)
         nImpor_AjD = 0
         nImpME_AjD = 0
         nImpMN_AjD = 0
         sTpoTcb = porstCOCpbAjD!TpoTcb
         If porstCOCpbAjD!NatCta = NATCTA_DEU Then
            sTpoTcb = IIf((IIf(porstCOCpbAjD!TpoMon = TPOMON_NAC, nImpMN_Sal, nImpME_Sal) > 0), sTpoTcb, IIf(sTpoTcb = TPOTCB_CPR, TPOTCB_VTA, TPOTCB_CPR))
         Else
            sTpoTcb = IIf((IIf(porstCOCpbAjD!TpoMon = TPOMON_NAC, nImpMN_Sal, nImpME_Sal) < 0), sTpoTcb, IIf(sTpoTcb = TPOTCB_CPR, TPOTCB_VTA, TPOTCB_CPR))
         End If
         nImpTCb_AjD = IIf(sTpoTcb = TPOTCB_VTA, porstCOTCbMes!ImpTCb_Vta, porstCOTCbMes!ImpTCb_Cpr)
         If nImpTCb_AjD > 0 And (nImpMN_Sal <> 0 Or nImpME_Sal <> 0) Then
            If porstCOCpbAjD!TpoMon = TPOMON_EXT Then
               If nImpMN_Sal > 0 Then
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpMN_Deb - (porstCOCpbAjD!nImpMN_Hab + (Abs(nImpME_Sal) * nImpTCb_AjD))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpMN_Hab - (porstCOCpbAjD!nImpMN_Deb + (Abs(nImpME_Sal) * nImpTCb_AjD))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(nImpME_Sal * nImpTCb_AjD, 2)
               If nImporte <> nImpMN_Sal Then
                  nImpor_AjD = gfRedond(nImporte - nImpMN_Sal, 2)
                  nImpMN_AjD = Abs(nImpor_AjD)
                  nImpME_AjD = gfRedond(IIf(nImpTCb_AjD = 0, 0, nImpMN_AjD / nImpTCb_AjD), 2)
               End If
            Else
               If nImpME_Sal > 0 Then
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpME_Deb - (porstCOCpbAjD!nImpME_Hab + IIf(nImpTCb_AjD = 0, 0, Abs(nImpMN_Sal) / nImpTCb_AjD))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpME_Hab - (porstCOCpbAjD!nImpME_Deb + IIf(nImpTCb_AjD = 0, 0, Abs(nImpMN_Sal) / nImpTCb_AjD))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(IIf(nImpTCb_AjD = 0, 0, nImpMN_Sal / nImpTCb_AjD), 2)
               If nImporte <> nImpME_Sal Then
                  nImpor_AjD = gfRedond(nImporte - nImpME_Sal, 2)
                  nImpME_AjD = Abs(nImpor_AjD)
                  nImpMN_AjD = gfRedond(nImpME_AjD * nImpTCb_AjD, 2)
               End If
            End If
            If gfRedond(nImpor_AjD, 2) <> 0 Then
               ' Adiciono el detalle del comprobante cuenta de documentos
               sTpoMon_AjD = IIf(porstCOCpbAjD!TpoMon = TPOMON_EXT, TPOMON_NAC, TPOMON_EXT)
               ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, porstCOCpbAjD!CodCta, "", "", "", "", sTpoCtb_AjD, sTpoMon_AjD, sTpoTcb, IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD), nImpMN_AjD, nImpME_AjD
               
               sCodCta_AjD = IIf(sTpoCtb_AjD = TPOCTB_DEB, porstCOCpbAjD!CodCta_AjD_Deb, porstCOCpbAjD!CodCta_AjD_Hab)
               sTpoCtb_AjD = IIf(sTpoCtb_AjD = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
               For nContador = 1 To UBound(aCodCta_AjD, 2)
                  ' Verifico los datos de la cuenta de ajuste
                  If aCodCta_AjD(1, nContador) = sCodCta_AjD And aCodCta_AjD(2, nContador) = sTpoCtb_AjD And aCodCta_AjD(3, nContador) = sTpoMon_AjD Then
                     Exit For
                  End If
               Next nContador
               If nContador > UBound(aCodCta_AjD, 2) Then
                  ReDim Preserve aCodCta_AjD(6, UBound(aCodCta_AjD, 2) + 1)
               End If
               aCodCta_AjD(1, nContador) = sCodCta_AjD
               aCodCta_AjD(2, nContador) = sTpoCtb_AjD
               aCodCta_AjD(3, nContador) = sTpoMon_AjD
               aCodCta_AjD(4, nContador) = IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD)
               aCodCta_AjD(5, nContador) = gfRedond(aCodCta_AjD(5, nContador) + nImpMN_AjD, 2)
               aCodCta_AjD(6, nContador) = gfRedond(aCodCta_AjD(6, nContador) + nImpME_AjD, 2)
            End If
         End If
         pgbProceso(1).Value = nNroItem
         porstCOCpbAjD.MoveNext
      Loop
      ' Adiciono el detalle del comprobante cuenta de documentos(perdidad o ganacia)
      For nContador = 1 To UBound(aCodCta_AjD, 2)
         nNroItem = nNroItem + 1
         sCodCta_AjD = aCodCta_AjD(1, nContador)
         sTpoCtb_AjD = aCodCta_AjD(2, nContador)
         sTpoMon_AjD = aCodCta_AjD(3, nContador)
         nImpTCb_AjD = aCodCta_AjD(4, nContador)
         nImpMN_AjD = aCodCta_AjD(5, nContador)
         nImpME_AjD = aCodCta_AjD(6, nContador)
         ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, sCodCta_AjD, "", "", "", "", sTpoCtb_AjD, sTpoMon_AjD, "V", nImpTCb_AjD, nImpMN_AjD, nImpME_AjD
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
   porstCOCpbAjD.Close
   Set porstCOCpbAjD = Nothing
End Sub
Private Sub ppAjuste_Documento()
   Static porstCOCpbCab As ADODB.Recordset
   Static porstCOCpbDet As ADODB.Recordset
   Static porstCOCpbAjD As ADODB.Recordset
   Static porstUltCoCpb  As ADODB.Recordset
   
   Static sNroComprobante As String
   Static nNroItem As Integer, nContador As Integer
   
   Static sTpoCtb_AjD As String, sTpoMon_AjD As String
   Static aCodCta_AjD(), sCodCta_AjD As String
   Static nImpTCb_AjD As Double, nImporte As Double
   Static nImpor_AjD As Double
   Static nImpMN_AjD As Double, nImpME_AjD As Double
   Static nImpMN_Sal As Double, nImpME_Sal As Double
   
   Set porstCOCpbCab = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   Set porstCOCpbAjD = New ADODB.Recordset
   Set porstUltCoCpb = New ADODB.Recordset
   
   pgbProceso(0).Min = 0
   ' Abro el recordset de seleccion de destinos
   With porstCOCpbAjD
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      'Genero la sentencia de seleccion documentos pendientes
      .Source = "SELECT a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc, b.TpoMon," _
              & "  b.TpoTcb, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='D', a.ImpMN, 0)), 0), 2) AS nImpMN_Deb," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='H', a.ImpMN, 0)), 0), 2) AS nImpMN_Hab," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='D', a.ImpME, 0)), 0), 2) AS nImpME_Deb," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='H', a.ImpME, 0)), 0), 2) AS nImpME_Hab " _
              & "FROM COCpbDet a, CoCta b " _
              & "WHERE a.CodCta=b.CodCta" _
              & "  AND a.MesPvs<='" & gsMesAct & "'" _
              & "  AND b.TpoCta='" & TPOCTA_TRA & "'" _
              & "  AND b.IndAjd='" & INDAJD_ACT & "'" _
              & "  AND b.IndDoc='" & INDDOC_ACT & "'" _
              & "  AND b.TpoAnl='" & TPOANL_DOC & "'" _
              & "  AND IFNULL(a.CodAux, '')<>''" _
              & "  AND IFNULL(a.CodTDc, '')<>''" _
              & "  AND IFNULL(a.SerDoc, '')<>''" _
              & "  AND IFNULL(a.NroDoc, '')<>'' " _
              & "  AND (IFNULL(b.CodCta_Ajd_Deb, '')<>'' OR IFNULL(b.CodCta_Ajd_Hab, '')<>'') " _
              & "GROUP BY a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc " _
              & "HAVING (ROUND(nImpMN_Deb - nImpMN_Hab, 2) <> 0.00) OR (ROUND(nImpME_Deb - nImpME_Hab, 2) <> 0.00) " _
              & "ORDER BY b.CodCta_Ajd_Deb, a.CodCta, a.CodAux, a.CodTDc, a.SerDoc, a.NroDoc"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   If porstCOCpbAjD.RecordCount > 0 Then
      porstCOCpbAjD.MoveFirst
      pgbProceso(0).Max = porstCOCpbAjD.RecordCount
      pgbProceso(0).Value = pgbProceso(0).Min
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbCab
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion cabecera de comprobantes
         .Source = "SELECT CodDro, NroCpb, FehCpb, GloCpb, MesPvs, " _
                 & "  TpoGnr, IndNCu, IndAnu," _
                 & "  UsrCre, FyHCre " _
                 & "FROM COCpbCab " _
                 & "WHERE CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ' Obtengo el numero e inserto la cabecera del comprobante
      With porstUltCoCpb
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = "SELECT IFNULL(MAX(NroCpb), 0) AS cUltNroCpb FROM COCpbCab WHERE CodDro='" & txtDato(0).Text & "' AND MesPvs='" & gsMesAct & "'"
         .CursorType = adOpenDynamic
         .LockType = adLockReadOnly
         .Open
         sNroComprobante = !cUltNroCpb
         .Close
      End With
      sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
      With porstCOCpbCab
         .AddNew
         !CodDro = txtDato(0).Text
         !NroCpb = sNroComprobante
         !FehCpb = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
         !GloCpb = "Ajustes por Diferencia de Cambio Documento Cierre de Mes " & gsMesAct
         !MesPvs = gsMesAct
         !TpoGnr = TPOGNR_DCA
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
         .Source = "SELECT CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux," _
                 & "  SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, TpoCtb, TpoPvs," _
                 & "  TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr," _
                 & "  UsrCre, FyHCre " _
                 & "FROM COCpbDet " _
                 & "WHERE CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ReDim aCodCta_AjD(6, 0)
      Do While Not porstCOCpbAjD.EOF
         ' Calculos para determinar el ajuste
         nNroItem = nNroItem + 1
         nImpTCb_AjD = IIf(porstCOCpbAjD!TpoTcb = TPOTCB_VTA, porstCOTCbMes!ImpTCb_Vta, porstCOTCbMes!ImpTCb_Cpr)
         nImpMN_Sal = gfRedond(porstCOCpbAjD!nImpMN_Deb - porstCOCpbAjD!nImpMN_Hab, 2)
         nImpME_Sal = gfRedond(porstCOCpbAjD!nImpME_Deb - porstCOCpbAjD!nImpME_Hab, 2)
         nImpor_AjD = 0
         nImpME_AjD = 0
         nImpMN_AjD = 0
         If nImpTCb_AjD > 0 And (nImpMN_Sal <> 0 Or nImpME_Sal <> 0) Then
            If porstCOCpbAjD!TpoMon = TPOMON_EXT Then
               If nImpMN_Sal > 0 Then
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpMN_Deb - (porstCOCpbAjD!nImpMN_Hab + (Abs(nImpME_Sal) * nImpTCb_AjD))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpMN_Hab - (porstCOCpbAjD!nImpMN_Deb + (Abs(nImpME_Sal) * nImpTCb_AjD))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(nImpME_Sal * nImpTCb_AjD, 2)
               If nImporte <> nImpMN_Sal Then
                  nImpor_AjD = gfRedond(nImporte - nImpMN_Sal, 2)
                  nImpMN_AjD = Abs(nImpor_AjD)
                  nImpME_AjD = gfRedond(IIf(nImpTCb_AjD = 0, 0, nImpMN_AjD / nImpTCb_AjD), 2)
               End If
            Else
               If nImpME_Sal > 0 Then
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpME_Deb - (porstCOCpbAjD!nImpME_Hab + IIf(nImpTCb_AjD = 0, 0, Abs(nImpMN_Sal) / nImpTCb_AjD))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpME_Hab - (porstCOCpbAjD!nImpME_Deb + IIf(nImpTCb_AjD = 0, 0, Abs(nImpMN_Sal) / nImpTCb_AjD))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(IIf(nImpTCb_AjD = 0, 0, nImpMN_Sal / nImpTCb_AjD), 2)
               If nImporte <> nImpME_Sal Then
                  nImpor_AjD = gfRedond(nImporte - nImpME_Sal, 2)
                  nImpME_AjD = Abs(nImpor_AjD)
                  nImpMN_AjD = gfRedond(nImpME_AjD * nImpTCb_AjD, 2)
               End If
            End If
            If nImpor_AjD <> 0 Then
               ' Adiciono el detalle del comprobante cuenta de documentos
               sTpoMon_AjD = IIf(porstCOCpbAjD!TpoMon = TPOMON_EXT, TPOMON_NAC, TPOMON_EXT)
               ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, porstCOCpbAjD!CodCta, porstCOCpbAjD!CodTDc, porstCOCpbAjD!CodAux, porstCOCpbAjD!SerDoc, porstCOCpbAjD!NroDoc, sTpoCtb_AjD, sTpoMon_AjD, porstCOCpbAjD!TpoTcb, IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD), nImpMN_AjD, nImpME_AjD
               
               sCodCta_AjD = IIf(sTpoCtb_AjD = TPOCTB_DEB, porstCOCpbAjD!CodCta_AjD_Deb, porstCOCpbAjD!CodCta_AjD_Hab)
               sTpoCtb_AjD = IIf(sTpoCtb_AjD = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
               For nContador = 1 To UBound(aCodCta_AjD, 2)
                  ' Verifico los datos de la cuenta de ajuste
                  If aCodCta_AjD(1, nContador) = sCodCta_AjD And aCodCta_AjD(2, nContador) = sTpoCtb_AjD And aCodCta_AjD(3, nContador) = sTpoMon_AjD Then
                     Exit For
                  End If
               Next nContador
               If nContador > UBound(aCodCta_AjD, 2) Then
                  ReDim Preserve aCodCta_AjD(6, UBound(aCodCta_AjD, 2) + 1)
               End If
               aCodCta_AjD(1, nContador) = sCodCta_AjD
               aCodCta_AjD(2, nContador) = sTpoCtb_AjD
               aCodCta_AjD(3, nContador) = sTpoMon_AjD
               aCodCta_AjD(4, nContador) = IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD)
               aCodCta_AjD(5, nContador) = gfRedond(aCodCta_AjD(5, nContador) + nImpMN_AjD, 2)
               aCodCta_AjD(6, nContador) = gfRedond(aCodCta_AjD(6, nContador) + nImpME_AjD, 2)
            End If
         End If
         pgbProceso(0).Value = nNroItem
         porstCOCpbAjD.MoveNext
      Loop
      ' Adiciono el detalle del comprobante cuenta de documentos(perdidad o ganacia)
      For nContador = 1 To UBound(aCodCta_AjD, 2)
         nNroItem = nNroItem + 1
         sCodCta_AjD = aCodCta_AjD(1, nContador)
         sTpoCtb_AjD = aCodCta_AjD(2, nContador)
         sTpoMon_AjD = aCodCta_AjD(3, nContador)
         nImpTCb_AjD = aCodCta_AjD(4, nContador)
         nImpMN_AjD = aCodCta_AjD(5, nContador)
         nImpME_AjD = aCodCta_AjD(6, nContador)
         ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, sCodCta_AjD, "", "", "", "", sTpoCtb_AjD, sTpoMon_AjD, "V", nImpTCb_AjD, nImpMN_AjD, nImpME_AjD
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
   porstCOCpbAjD.Close
   Set porstCOCpbAjD = Nothing
End Sub
Private Sub ppajuste_Auxiliar()
   Static porstCOCpbCab As ADODB.Recordset
   Static porstCOCpbDet As ADODB.Recordset
   Static porstCOCpbAjD As ADODB.Recordset
   Static porstUltCoCpb  As ADODB.Recordset
   
   Static sNroComprobante As String
   Static nNroItem As Integer, nContador As Integer
   
   Static sTpoCtb_AjD As String, sTpoMon_AjD As String
   Static aCodCta_AjD(), sCodCta_AjD As String
   Static nImpTCb_AjD As Double, nImporte As Double
   Static nImpor_AjD As Double
   Static nImpMN_AjD As Double, nImpME_AjD As Double
   Static nImpMN_Sal As Double, nImpME_Sal As Double
   
   Set porstCOCpbCab = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   Set porstCOCpbAjD = New ADODB.Recordset
   Set porstUltCoCpb = New ADODB.Recordset
   
   pgbProceso(2).Min = 0
   ' Abro el recordset de seleccion de destinos
   With porstCOCpbAjD
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      'Genero la sentencia de seleccion documentos pendientes
      .Source = "SELECT a.CodCta, a.CodAux, b.TpoMon," _
              & "  b.TpoTcb, b.CodCta_Ajd_Deb, b.CodCta_Ajd_Hab," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='D', a.ImpMN, 0)), 0), 2) AS nImpMN_Deb," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='H', a.ImpMN, 0)), 0), 2) AS nImpMN_Hab," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='D', a.ImpME, 0)), 0), 2) AS nImpME_Deb," _
              & "  ROUND(IFNULL(SUM(IF(a.TpoCtb='H', a.ImpME, 0)), 0), 2) AS nImpME_Hab " _
              & "FROM COCpbDet a, CoCta b " _
              & "WHERE a.CodCta=b.CodCta" _
              & "  AND a.MesPvs<='" & gsMesAct & "'" _
              & "  AND b.TpoCta='" & TPOCTA_TRA & "'" _
              & "  AND b.IndAjd='" & INDAJD_ACT & "'" _
              & "  AND b.IndDoc='" & INDAUX_ACT & "'" _
              & "  AND b.TpoAnl='" & TPOANL_AUX & "'" _
              & "  AND IFNULL(a.CodAux, '')<>''" _
              & "  AND (IFNULL(b.CodCta_Ajd_Deb, '')<>'' OR IFNULL(b.CodCta_Ajd_Hab, '')<>'') " _
              & "GROUP BY a.CodCta, a.CodAux " _
              & "HAVING (ROUND(nImpMN_Deb - nImpMN_Hab, 2) <> 0.00) OR (ROUND(nImpME_Deb - nImpME_Hab, 2) <> 0.00) " _
              & "ORDER BY b.CodCta_Ajd_Deb, a.CodCta, a.CodAux"
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   If porstCOCpbAjD.RecordCount > 0 Then
      porstCOCpbAjD.MoveFirst
      pgbProceso(2).Max = porstCOCpbAjD.RecordCount
      pgbProceso(2).Value = pgbProceso(2).Min
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbCab
         .ActiveConnection = pocnnMain
         'Genero la sentencia de seleccion cabecera de comprobantes
         .Source = "SELECT CodDro, NroCpb, FehCpb, GloCpb, MesPvs, " _
                 & "  TpoGnr, IndNCu, IndAnu," _
                 & "  UsrCre, FyHCre " _
                 & "FROM COCpbCab " _
                 & "WHERE CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ' Obtengo el numero e inserto la cabecera del comprobante
      With porstUltCoCpb
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = "SELECT IFNULL(MAX(NroCpb), 0) AS cUltNroCpb FROM COCpbCab WHERE CodDro='" & txtDato(0).Text & "' AND MesPvs='" & gsMesAct & "'"
         .CursorType = adOpenDynamic
         .LockType = adLockReadOnly
         .Open
         sNroComprobante = !cUltNroCpb
         .Close
      End With
      sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
      With porstCOCpbCab
         .AddNew
         !CodDro = txtDato(0).Text
         !NroCpb = sNroComprobante
         !FehCpb = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
         !GloCpb = "Ajustes por Diferencia de Cambio Documento Cierre de Mes " & gsMesAct
         !MesPvs = gsMesAct
         !TpoGnr = TPOGNR_DCA
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
         .Source = "SELECT CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux," _
                 & "  SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, TpoCtb, TpoPvs," _
                 & "  TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr," _
                 & "  UsrCre, FyHCre " _
                 & "FROM COCpbDet " _
                 & "WHERE CodDro=''"
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ReDim aCodCta_AjD(6, 0)
      Do While Not porstCOCpbAjD.EOF
         ' Calculos para determinar el ajuste
         nNroItem = nNroItem + 1
         nImpTCb_AjD = IIf(porstCOCpbAjD!TpoTcb = TPOTCB_VTA, porstCOTCbMes!ImpTCb_Vta, porstCOTCbMes!ImpTCb_Cpr)
         nImpMN_Sal = gfRedond(porstCOCpbAjD!nImpMN_Deb - porstCOCpbAjD!nImpMN_Hab, 2)
         nImpME_Sal = gfRedond(porstCOCpbAjD!nImpME_Deb - porstCOCpbAjD!nImpME_Hab, 2)
         nImpor_AjD = 0
         nImpME_AjD = 0
         nImpMN_AjD = 0
         If nImpTCb_AjD > 0 And (nImpMN_Sal <> 0 Or nImpME_Sal <> 0) Then
            If porstCOCpbAjD!TpoMon = TPOMON_EXT Then
               If nImpMN_Sal > 0 Then
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpMN_Deb - (porstCOCpbAjD!nImpMN_Hab + (Abs(nImpME_Sal) * nImpTCb_AjD))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpMN_Hab - (porstCOCpbAjD!nImpMN_Deb + (Abs(nImpME_Sal) * nImpTCb_AjD))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(nImpME_Sal * nImpTCb_AjD, 2)
               If nImporte <> nImpMN_Sal Then
                  nImpor_AjD = gfRedond(nImporte - nImpMN_Sal, 2)
                  nImpMN_AjD = Abs(nImpor_AjD)
                  nImpME_AjD = gfRedond(IIf(nImpTCb_AjD = 0, 0, nImpMN_AjD / nImpTCb_AjD), 2)
               End If
            Else
               If nImpME_Sal > 0 Then
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpME_Deb - (porstCOCpbAjD!nImpME_Hab + IIf(nImpTCb_AjD = 0, 0, Abs(nImpMN_Sal) / nImpTCb_AjD))) < 0, TPOCTB_DEB, TPOCTB_HAB)
               Else
                  sTpoCtb_AjD = IIf((porstCOCpbAjD!nImpME_Hab - (porstCOCpbAjD!nImpME_Deb + IIf(nImpTCb_AjD = 0, 0, Abs(nImpMN_Sal) / nImpTCb_AjD))) < 0, TPOCTB_HAB, TPOCTB_DEB)
               End If
               nImporte = gfRedond(IIf(nImpTCb_AjD = 0, 0, nImpMN_Sal / nImpTCb_AjD), 2)
               If nImporte <> nImpME_Sal Then
                  nImpor_AjD = gfRedond(nImporte - nImpME_Sal, 2)
                  nImpME_AjD = Abs(nImpor_AjD)
                  nImpMN_AjD = gfRedond(nImpME_AjD * nImpTCb_AjD, 2)
               End If
            End If
            If nImpor_AjD <> 0 Then
               ' Adiciono el detalle del comprobante cuenta de documentos
               sTpoMon_AjD = IIf(porstCOCpbAjD!TpoMon = TPOMON_EXT, TPOMON_NAC, TPOMON_EXT)
               ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, porstCOCpbAjD!CodCta, "", porstCOCpbAjD!CodAux, "", "", sTpoCtb_AjD, sTpoMon_AjD, porstCOCpbAjD!TpoTcb, IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD), nImpMN_AjD, nImpME_AjD
               'ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, porstCOCpbAjD!CodCta, "", "", "", "", sTpoCtb_AjD, sTpoMon_AjD, sTpoTcb, IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD), nImpMN_AjD, nImpME_AjDppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, porstCOCpbAjD!CodCta, "", "", "", "", sTpoCtb_AjD, sTpoMon_AjD, sTpoTcb, IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD), nImpMN_AjD, nImpME_AjD
               
               sCodCta_AjD = IIf(sTpoCtb_AjD = TPOCTB_DEB, porstCOCpbAjD!CodCta_AjD_Deb, porstCOCpbAjD!CodCta_AjD_Hab)
               sTpoCtb_AjD = IIf(sTpoCtb_AjD = TPOCTB_DEB, TPOCTB_HAB, TPOCTB_DEB)
               For nContador = 1 To UBound(aCodCta_AjD, 2)
                  ' Verifico los datos de la cuenta de ajuste
                  If aCodCta_AjD(1, nContador) = sCodCta_AjD And aCodCta_AjD(2, nContador) = sTpoCtb_AjD And aCodCta_AjD(3, nContador) = sTpoMon_AjD Then
                     Exit For
                  End If
               Next nContador
               If nContador > UBound(aCodCta_AjD, 2) Then
                  ReDim Preserve aCodCta_AjD(6, UBound(aCodCta_AjD, 2) + 1)
               End If
               aCodCta_AjD(1, nContador) = sCodCta_AjD
               aCodCta_AjD(2, nContador) = sTpoCtb_AjD
               aCodCta_AjD(3, nContador) = sTpoMon_AjD
               aCodCta_AjD(4, nContador) = IIf(IsNull(nImpTCb_AjD), 1, nImpTCb_AjD)
               aCodCta_AjD(5, nContador) = gfRedond(aCodCta_AjD(5, nContador) + nImpMN_AjD, 2)
               aCodCta_AjD(6, nContador) = gfRedond(aCodCta_AjD(6, nContador) + nImpME_AjD, 2)
            End If
         End If
         pgbProceso(2).Value = nNroItem
         porstCOCpbAjD.MoveNext
      Loop
      ' Adiciono el detalle del comprobante cuenta de documentos(perdidad o ganacia)
      For nContador = 1 To UBound(aCodCta_AjD, 2)
         nNroItem = nNroItem + 1
         sCodCta_AjD = aCodCta_AjD(1, nContador)
         sTpoCtb_AjD = aCodCta_AjD(2, nContador)
         sTpoMon_AjD = aCodCta_AjD(3, nContador)
         nImpTCb_AjD = aCodCta_AjD(4, nContador)
         nImpMN_AjD = aCodCta_AjD(5, nContador)
         nImpME_AjD = aCodCta_AjD(6, nContador)
         ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, sCodCta_AjD, "", "", "", "", sTpoCtb_AjD, sTpoMon_AjD, "V", nImpTCb_AjD, nImpMN_AjD, nImpME_AjD
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
   porstCOCpbAjD.Close
   Set porstCOCpbAjD = Nothing

End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYUDAT Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod " Length(CodDro)=4 ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'         cmdAceptar.Enabled = True
'         cmdSalir.Enabled = True
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
               lblDatoDeta(tnIndex).Caption = " " & !DetDro
            End If
         End With
      End Select
   End If
End Function

Private Sub ppInsDetalle_Cpb(porstCOCpbDet As ADODB.Recordset, cNroCpb As String, nNroIte As Integer, cCodCta As String, cCodTDc As String, cCodAux As String, cSerDoc As String, cNroDoc As String, cTpoCtb As String, cTpoMon As String, cTpoTcb As String, nImpTCb As Double, nImpMN As Double, nImpME As Double)
   ' Adiciono el detalle del comprobante
   With porstCOCpbDet
      .AddNew
      !CodDro = txtDato(0).Text
      !NroCpb = cNroCpb
      !NroIte = nNroIte
      !MesPvs = gsMesAct
      !BlqIte = nNroIte
      !CodCta = cCodCta
      !FehOpe = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
      !FeEDoc = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
      !FeVDoc = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
      !FeRDoc = gfUltDia("01/" & gsMesAct & "/" & gsAnoAct)
      !CodTDc = IIf(cCodTDc = "", Null, cCodTDc)
      !CodCCo = IIf(Mid(cCodCta, 1, 1) = "6", CODCCO_AJD, IIf(Mid(cCodCta, 1, 1) = "7", CODCCO_AJD, IIf(Mid(cCodCta, 1, 1) = "9", CODCCO_AJD, Null)))
      !CodAux = IIf(cCodAux = "", Null, cCodAux)
      !SerDoc = IIf(cSerDoc = "", Null, cSerDoc)
      !NroDoc = IIf(cNroDoc = "", Null, cNroDoc)
      !GloIte = "Ajuste Diferencia de Cambio Cierre Mes"
      !TpoCtb = cTpoCtb
      !TpoMon = cTpoMon
      !TpoTcb = cTpoTcb
      !ImpTCb = nImpTCb
      !ImpMN = IIf(porstCOCpbDet!TpoMon = TPOMON_NAC, nImpMN, 0)
      !ImpME = IIf(porstCOCpbDet!TpoMon = TPOMON_EXT, nImpME, 0)
      !TpoPvs = TPOPVS_OTR
      !TpoGnr = TPOGNR_DCA
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With
End Sub
