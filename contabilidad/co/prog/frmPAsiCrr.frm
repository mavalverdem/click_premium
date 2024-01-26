VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPAsiCrr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[título]"
   ClientHeight    =   2910
   ClientLeft      =   2925
   ClientTop       =   2700
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   375
      Index           =   0
      Left            =   4185
      Picture         =   "frmPAsiCrr.frx":0000
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
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Default         =   -1  'True
      Height          =   495
      Left            =   2573
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Asientos de Cierre"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   90
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   450
      Width           =   3660
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese Diario"
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmPAsiCrr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pocnnMain As ADODB.Connection
Public porstCodro As ADODB.Recordset
Public porstCOCtaAcu As ADODB.Recordset
Public porstCOCta As ADODB.Recordset
Public porstCOICM As ADODB.Recordset
Dim sSentencia As String

Private Sub Form_Load()
   pgbProceso(0).Value = 0
  
  'Abrir Tablas.
   
   Set pocnnMain = New ADODB.Connection
   Set porstCodro = New ADODB.Recordset
   Set porstCOCtaAcu = New ADODB.Recordset
   Set porstCOCta = New ADODB.Recordset
   Set porstCOICM = New ADODB.Recordset
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
   
   sSentencia = "SELECT CodCta, "
   sSentencia = sSentencia & gsAcuAnt(1) & " as cAcuD_MN, " & gsAcuAnt(3) & " as cAcuH_MN, "
   sSentencia = sSentencia & "COCtaAcu.AcuD00_MN, COCtaAcu.AcuD01_MN, COCtaAcu.AcuD02_MN, COCtaAcu.AcuD03_MN, COCtaAcu.AcuD04_MN, COCtaAcu.AcuD05_MN, COCtaAcu.AcuD06_MN, COCtaAcu.AcuD07_MN, COCtaAcu.AcuD08_MN, COCtaAcu.AcuD09_MN, COCtaAcu.AcuD10_MN, COCtaAcu.AcuD11_MN, COCtaAcu.AcuD12_MN, COCtaAcu.AcuD13_MN, "
   sSentencia = sSentencia & "COCtaAcu.AcuH00_MN, COCtaAcu.AcuH01_MN, COCtaAcu.AcuH02_MN, COCtaAcu.AcuH03_MN, COCtaAcu.AcuH04_MN, COCtaAcu.AcuH05_MN, COCtaAcu.AcuH06_MN, COCtaAcu.AcuH07_MN, COCtaAcu.AcuH08_MN, COCtaAcu.AcuH09_MN, COCtaAcu.AcuH10_MN, COCtaAcu.AcuH11_MN, COCtaAcu.AcuH12_MN, COCtaAcu.AcuH13_MN, UsrCre, FyHCre, UsrMdf, FyHMdf "
   sSentencia = sSentencia & "FROM COCtaAcu "
   sSentencia = sSentencia & "ORDER BY CodCta"
   With porstCOCtaAcu
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sSentencia
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
   
   sSentencia = "SELECT CodCta, IndMoe "
   sSentencia = sSentencia & "FROM COCta "
   sSentencia = sSentencia & "ORDER BY CodCta"
   With porstCOCta
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sSentencia
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   
   sSentencia = "SELECT MesICM, ImpInd "
   sSentencia = sSentencia & "FROM CoICM "
   sSentencia = sSentencia & "ORDER BY MesICM"
   With porstCOICM
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sSentencia
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
End Sub

Private Sub Form_Activate()
'   If gsMesAct <> 13 Then
'      Unload Me
'   Else
      cmdAceptar.Enabled = False
      cmdSalir.Enabled = True
      cmdSalir.SetFocus
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   porstCodro.Close
   pocnnMain.Close
   Set porstCodro = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdAceptar_Click()
'   On Error GoTo Err
   
 '[Propio del formulario.
   If TxtDato(0).Text = "" Then
      MsgBox TEXT_6002, vbCritical
      TxtDato(0).SetFocus
      Exit Sub
   End If
   
   If gnIndMNE <> INDMNE_ACT Then
      MsgBox "La Empresa trabaja solo con una Moneda", vbExclamation
      Exit Sub
   End If
   pgbProceso(0).Value = 0: pgbProceso(0).Min = 0
   
   pocnnMain.BeginTrans                'INICIA TRANSACCION.
   '[Paso 1 : Elimino los comprobantes de ajuste del mes
     pocnnMain.Execute "DELETE FROM COCpbCab WHERE TpoGnr=" & Str(TPOGNR_DCA) & " And MesPvs=" & gsMesAct
   '[Paso 2 : Generacion de Asientos de Cierre
     ppAjuste_Cierre
   pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
   
   MsgBox TEXT_8008, vbInformation
  
   Exit Sub
'Err:
'  pocnnMain.RollbackTrans              'RESTAURA TRANSACCION.
'  MsgBox TEXT_6001 & " " & Err.Number & " : " & Err.Description

End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      ppAyuBus AYUDAT, Index
      TxtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
   End Select
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   TxtDato(Index).SelStart = 0
   TxtDato(Index).SelLength = TxtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYUDAT, Index
   End If
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(TxtDato(Index))) + 1 = TxtDato(Index).MaxLength Then
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

Private Sub ppAjuste_Cierre()
   Static porstCOCpbCab As ADODB.Recordset
   Static porstCOCpbDet As ADODB.Recordset
   Static porstCOCie As ADODB.Recordset
   Static porstCOCieCta As ADODB.Recordset
   Static porstUltCoCpb As ADODB.Recordset
   Static porstCOHojTra1 As ADODB.Recordset
   Static porstCOHojTra2 As ADODB.Recordset
   Static porstCOHojTra3 As ADODB.Recordset
   
   Static sNroComprobante As String, sTpoTcb As String
   Static sGrabacion As String
   Static nNroItem As Integer, nContador As Integer
   Static nImpAD As Double, nImpAH As Double
   Static nImpTD As Double, nImpTH As Double
   
   Static sTpoCtb_AjD As String, sTpoMon_AjD As String
   Static aCodCta_AjD(), sCodCta_AjD As String
   Static nImpTCb_AjD As Double, nImporte As Double
   Static nImpor_AjD As Double
   Static nImpMN_AjD As Double, nImpME_AjD As Double
   Static nImpMN_Sal As Double, nImpME_Sal As Double
   Static bIndCta As Boolean
   
   Set porstCOCpbCab = New ADODB.Recordset
   Set porstCOCpbDet = New ADODB.Recordset
   Set porstCOCie = New ADODB.Recordset
   Set porstCOCieCta = New ADODB.Recordset
   Set porstUltCoCpb = New ADODB.Recordset
   Set porstCOHojTra1 = New ADODB.Recordset
   Set porstCOHojTra2 = New ADODB.Recordset
   Set porstCOHojTra3 = New ADODB.Recordset
   
   pgbProceso(0).Min = 0
   sSentencia = "SELECT sum(ImpSalI+ImpAdq-ImpVtRr) as sSalD, sum(ImpSalIA+ImpAdqA-ImpVtRrA) as sSalH, "
   sSentencia = sSentencia & "sum(ImpSalIH+ImpDepH-ImpVtRrH) as sDepD, sum(ImpSalIDA+ImpDepA-ImpVtRrDA) as sDepH "
   sSentencia = sSentencia & "FROM COHojTra1 "
   With porstCOHojTra1
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sSentencia
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   
   sSentencia = "SELECT sum(ImpValH) as sValD, sum(ImpValA) sValH  "
   sSentencia = sSentencia & "FROM COHojTra2 "
   With porstCOHojTra2
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sSentencia
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   
   sSentencia = "SELECT sum(ImpSalIH+ImpAmtH+ImpAmtAdH) as sAmtD, sum(ImpSalIA+ImpAmtA+ImpAmtAdA) as sAmtH "
   sSentencia = sSentencia & "FROM COHojTra3 "
   With porstCOHojTra3
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sSentencia
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   
   sSentencia = "SELECT NroCie, DetCie "
   sSentencia = sSentencia & "FROM COcie "
   sSentencia = sSentencia & "ORDER BY NroCie"
   With porstCOCie
      If .State = adStateOpen Then .Close
      .ActiveConnection = pocnnMain
      .Source = sSentencia
      .CursorType = adOpenDynamic
      .LockType = adLockReadOnly
      .Open
   End With
   
   If porstCOCie.RecordCount > 0 Then
      sSentencia = "SELECT Count(*) as ConReg "
      sSentencia = sSentencia & "FROM COCieCta "
      With porstCOCieCta
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = sSentencia
         .CursorType = adOpenDynamic
         .LockType = adLockReadOnly
         .Open
      End With
      pgbProceso(0).Max = porstCOCieCta!ConReg
      pgbProceso(0).Value = pgbProceso(0).Min
      pocnnMain.Execute "DELETE FROM COCpbCab WHERE TpoGnr='" & TPOGNR_CIE & "' AND CodDro='" & TxtDato(0).Text & "'"
      porstCOCie.MoveFirst
      ' Genero la sentencia de seleccion cabecera de comprobantes
      sGrabacion = "SELECT CodDro, NroCpb, FehCpb, GloCpb, MesPvs, "
      sGrabacion = sGrabacion & "TpoGnr, IndNCu, IndAnu, UsrCre, FyHCre "
      sGrabacion = sGrabacion & "FROM COCpbCab "
      sGrabacion = sGrabacion & "WHERE CodDro=''"
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbCab
         .ActiveConnection = pocnnMain
         .Source = sGrabacion
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      ' Genero la sentencia de seleccion detalles de comprobantes
      sGrabacion = "SELECT CodDro, NroCpb, NroIte, MesPvs, BlqIte, CodTDc, FehOpe, CodCta, CodCCo, CodAux, "
      sGrabacion = sGrabacion & "SerDoc, NroDoc, FeEDoc, FeVDoc, FeRDoc, RefDoc, GloIte, TpoCtb, TpoPvs, "
      sGrabacion = sGrabacion & "TpoMon, TpoTCb, ImpTCb, ImpMN, ImpME, TpoGnr, UsrCre, FyHCre "
      sGrabacion = sGrabacion & "FROM COCpbDet "
      sGrabacion = sGrabacion & "WHERE CodDro=''"
      ' Abro el recordset de grabacion de la cabecera de comprobante
      With porstCOCpbDet
         If .State = adStateOpen Then .Close
         .ActiveConnection = pocnnMain
         .Source = sGrabacion
         .CursorType = adOpenDynamic
         .LockType = adLockOptimistic
         .Open
      End With
      nNroItem = 1
      ReDim aCodCta_AjD(3, 0)
      Do While Not porstCOCie.EOF
         ' Obtengo el numero e inserto la cabecera del comprobante
         With porstUltCoCpb
            If .State = adStateOpen Then .Close
            .ActiveConnection = pocnnMain
            .Source = "SELECT IFNULL(MAX(NroCpb), 0) AS cUltNroCpb FROM COCpbCab WHERE CodDro='" & TxtDato(0).Text & "'" ' AND MesPvs='" & gsMesAct & "'"
            .CursorType = adOpenDynamic
            .LockType = adLockReadOnly
            .Open
            sNroComprobante = !cUltNroCpb
            .Close
         End With
         sNroComprobante = gfCeros(sNroComprobante, 6, 1, "0")
         With porstCOCpbCab
            .AddNew
            !CodDro = TxtDato(0).Text
            !NroCpb = sNroComprobante
            !FehCpb = gfUltDia("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct)
            !GloCpb = porstCOCie!DetCie
            !MesPvs = gsMesAct
            !TpoGnr = TPOGNR_CIE
            !IndNCu = INDNCU_FAL
            !IndAnu = INDANU_FAL
            !UsrCre = gsAbvUsr
            !FyHCre = Now
            .Update
         End With
         sSentencia = "SELECT NroIte, CodCta, IndHTr, TpoHTr, TpoHT1, FmlCie, TpoCtb, IndCCt, TpoCtbI, ImpMNI, IndAMo "
         sSentencia = sSentencia & "FROM COCieCta "
         sSentencia = sSentencia & "WHERE NroCie='" & porstCOCie!NroCie & "'"
         sSentencia = sSentencia & "ORDER BY NroCie, NroIte"
         With porstCOCieCta
            If .State = adStateOpen Then .Close
            .ActiveConnection = pocnnMain
            .Source = sSentencia
            .CursorType = adOpenDynamic
            .LockType = adLockReadOnly
            .Open
         End With
         nImpTD = 0
         nImpTH = 0
         porstCOCieCta.MoveFirst
         Do While Not porstCOCieCta.EOF
            ' Calculos para determinar el ajuste
            bIndCta = False
            If porstCOCieCta!IndHTr = 1 Then
               If porstCOCieCta!TpoHTr = 0 Then
                  With porstCOHojTra1
                     .MoveFirst
                     If Not .EOF Then
                        If porstCOCieCta!TpoHT1 = 0 Then
                           nImporte = CDec(!sSalD) - CDec(!sSalH)
                        Else
                           nImporte = CDec(!sDepD) - CDec(!sDepH)
                        End If
                     End If
                  End With
               ElseIf porstCOCieCta!TpoHTr = 1 Then
                  With porstCOHojTra2
                     .MoveFirst
                     If Not .EOF Then
                        nImporte = CDec(!sValD) - CDec(!sValH)
                     End If
                  End With
               ElseIf porstCOCieCta!TpoHTr = 2 Then
                  With porstCOHojTra3
                     .MoveFirst
                     If Not .EOF Then
                        nImporte = CDec(!sAmtD) - CDec(!sAmtH)
                     End If
                  End With
               End If
            Else
               If IsNull(porstCOCieCta!FmlCie) Or Trim(porstCOCieCta!FmlCie) = "" Then
                  If porstCOCieCta!ImpMNI > 0 Then
                     nImporte = CDec(porstCOCieCta!ImpMNI)
                     sTpoCtb_AjD = IIf(IsNull(porstCOCieCta!TpoCtbI), TPOCTB_DEB, porstCOCieCta!TpoCtbI)
                     bIndCta = True
                  Else
                     'INDICADOR DE AJUSTE EN MANTEN. CUENTA CIERRE IndAMo
                     If porstCOCieCta!IndAMo = 1 Then
                        ppAjuste_CorrMon porstCOCieCta!CodCta, nImpAD, nImpAH
                        nImporte = CDec(nImpAD) - CDec(nImpAH)
                     Else
                        With porstCOCta
                           .MoveFirst
                           .Find "CodCta='" & porstCOCieCta!CodCta & "'"
                           If Not .EOF Then
                              If !IndMoe Then
                                 ppAjuste_CorrMon porstCOCieCta!CodCta, nImpAD, nImpAH
                                 nImporte = CDec(nImpAD) - CDec(nImpAH)
                              Else
                                 porstCOCtaAcu.MoveFirst
                                 porstCOCtaAcu.Find "CodCta='" & porstCOCieCta!CodCta & "'"
                                 If Not porstCOCtaAcu.EOF Then
                                    nImporte = CDec(IIf(IsNull(porstCOCtaAcu!cAcuD_MN), 0, porstCOCtaAcu!cAcuD_MN)) - CDec(IIf(IsNull(porstCOCtaAcu!cAcuH_MN), 0, porstCOCtaAcu!cAcuH_MN))
                                 End If
                              End If
                           End If
                        End With
                     End If
                  End If
               Else
                  'INDICADOR DE AJUSTE EN MANTEN. CUENTA CIERRE IndAMo
                  If porstCOCieCta!IndAMo = 1 Then
                     ResuelveFormula porstCOCieCta!FmlCie, nImpAD, nImpAH, True
                  Else
                     ResuelveFormula porstCOCieCta!FmlCie, nImpAD, nImpAH, False
                  End If
                  nImporte = CDec(nImpAD) - CDec(nImpAH)
               End If
            End If
            nImpMN_AjD = Abs(CDec(nImporte))
            'Tipo contable lo toma del indicador TpoCtb del Manten. de Cierre Ctas.
            If nImporte > 0 And Not bIndCta Then
               'sTpoCtb_AjD = TPOCTB_HAB
               sTpoCtb_AjD = IIf(porstCOCieCta!TpoCtb = 0, TPOCTB_HAB, TPOCTB_DEB)
            Else
               'sTpoCtb_AjD = TPOCTB_DEB
               sTpoCtb_AjD = IIf(porstCOCieCta!TpoCtb = 0, TPOCTB_DEB, TPOCTB_HAB)
            End If
            If sTpoCtb_AjD = TPOCTB_HAB Then
               nImpTH = CDec(nImpTH) + CDec(nImpMN_AjD)
            Else
               nImpTD = CDec(nImpTD) + CDec(nImpMN_AjD)
            End If
            sTpoMon_AjD = TPOMON_NAC
            sTpoTcb = TPOTCB_VTA
            nImpME_AjD = 0
            nImpTCb_AjD = 1
            ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem, porstCOCieCta!CodCta, "", "", "", "", _
            sTpoCtb_AjD, sTpoMon_AjD, sTpoTcb, nImpTCb_AjD, CDec(nImpMN_AjD), nImpME_AjD, porstCOCie!DetCie
            For nContador = 1 To UBound(aCodCta_AjD, 2)
               ' Verifico los datos de la cuenta de ajuste
               If aCodCta_AjD(1, nContador) = porstCOCieCta!CodCta And aCodCta_AjD(2, nContador) = sTpoCtb_AjD Then
                  Exit For
               End If
            Next nContador
            If nContador > UBound(aCodCta_AjD, 2) Then
               ReDim Preserve aCodCta_AjD(3, UBound(aCodCta_AjD, 2) + 1)
            End If
            aCodCta_AjD(1, nContador) = porstCOCieCta!CodCta
            aCodCta_AjD(2, nContador) = sTpoCtb_AjD
            aCodCta_AjD(3, nContador) = gfRedond(CDec(aCodCta_AjD(3, nContador)) + CDec(nImpMN_AjD), 2)
            nNroItem = nNroItem + 1
            porstCOCieCta.MoveNext
            If Not porstCOCieCta.EOF Then
               If porstCOCieCta!IndCCt = 1 And nImpTD <> nImpTH Then
                  If nImpTD > nImpTH Then
                     nImpMN_AjD = CDec(nImpTD) - CDec(nImpTH)
                     sTpoCtb_AjD = TPOCTB_HAB
                  Else
                     nImpMN_AjD = CDec(nImpTH) - CDec(nImpTD)
                     sTpoCtb_AjD = TPOCTB_DEB
                  End If
                  ppInsDetalle_Cpb porstCOCpbDet, sNroComprobante, nNroItem + 1, porstCOCieCta!CodCta, "", "", "", "", _
                     sTpoCtb_AjD, sTpoMon_AjD, sTpoTcb, nImpTCb_AjD, CDec(nImpMN_AjD), nImpME_AjD, porstCOCie!DetCie
                  porstCOCpbCab!IndNCu = INDNCU_VER
                  porstCOCpbCab.Update
                  For nContador = 1 To UBound(aCodCta_AjD, 2)
                     ' Verifico los datos de la cuenta de ajuste
                     If aCodCta_AjD(1, nContador) = porstCOCieCta!CodCta And aCodCta_AjD(2, nContador) = sTpoCtb_AjD Then
                        Exit For
                     End If
                  Next nContador
                  If nContador > UBound(aCodCta_AjD, 2) Then
                     ReDim Preserve aCodCta_AjD(3, UBound(aCodCta_AjD, 2) + 1)
                  End If
                  aCodCta_AjD(1, nContador) = porstCOCieCta!CodCta
                  aCodCta_AjD(2, nContador) = sTpoCtb_AjD
                  aCodCta_AjD(3, nContador) = gfRedond(CDec(aCodCta_AjD(3, nContador)) + CDec(nImpMN_AjD), 2)
                  nNroItem = nNroItem + 1
                  Exit Do
               End If
            End If
            pgbProceso(0).Value = IIf(nNroItem > pgbProceso(0).Max, pgbProceso(0).Max, nNroItem)
         Loop
         ' Adiciono el detalle del comprobante cuenta de documentos(perdidad o ganacia)
         porstCOCpbDet.UpdateBatch
         porstCOCie.MoveNext
      Loop
      
      For nContador = 1 To UBound(aCodCta_AjD, 2)
         With porstCOCtaAcu
            .MoveFirst
            .Find "CodCta='" & aCodCta_AjD(1, nContador) & "'"
            If .EOF Then
               .AddNew
               .Fields("UsrCre") = gsAbvUsr
               .Fields("FyHCre") = Now
            Else
               .Fields("UsrMdf") = gsAbvUsr
               .Fields("FyHMdf") = Now
            End If
            .Fields("CodCta") = aCodCta_AjD(1, nContador)
            .Fields("AcuD" & gsMesAct & "_MN") = IIf(aCodCta_AjD(2, nContador) = TPOCTB_DEB, CDec(aCodCta_AjD(3, nContador)), 0)
            .Fields("AcuH" & gsMesAct & "_MN") = IIf(aCodCta_AjD(2, nContador) = TPOCTB_HAB, CDec(aCodCta_AjD(3, nContador)), 0)
            .Update
         End With
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
   porstCOCieCta.Close
   Set porstCOCieCta = Nothing
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYUDAT Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod " Length(CodDro)=4 ", TxtDato(tnIndex).Text, 0, 0, Me.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + TxtDato(tnIndex).Left
         TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
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
         If TxtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With porstCodro
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodDro='" & TxtDato(tnIndex).Text & "'"
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

Private Sub ppInsDetalle_Cpb(porstCOCpbDet As ADODB.Recordset, cNroCpb As String, nNroIte As Integer, cCodCta As String, cCodTDc As String, cCodAux As String, cSerDoc As String, cNroDoc As String, cTpoCtb As String, cTpoMon As String, cTpoTcb As String, nImpTCb As Double, nImpMN As Double, nImpME As Double, cGlosa As String)
   ' Adiciono el detalle del comprobante
   With porstCOCpbDet
      .AddNew
      !CodDro = TxtDato(0).Text
      !NroCpb = cNroCpb
      !NroIte = nNroIte
      !MesPvs = gsMesAct
      !BlqIte = nNroIte
      !CodCta = cCodCta
      !FehOpe = gfUltDia("01/" & gfMesAct(gsMesAct) & "/" & gsAnoAct)
      !FeEDoc = Null
      !FeVDoc = Null
      !FeRDoc = Null
      !CodTDc = IIf(cCodTDc = "", Null, cCodTDc)
      !CodCCo = Null
      !CodAux = IIf(cCodAux = "", Null, cCodAux)
      !SerDoc = IIf(cSerDoc = "", Null, cSerDoc)
      !NroDoc = IIf(cNroDoc = "", Null, cNroDoc)
      !GloIte = cGlosa
      !TpoCtb = cTpoCtb
      !TpoMon = cTpoMon
      !TpoTcb = cTpoTcb
      !ImpTCb = nImpTCb
      !ImpMN = IIf(porstCOCpbDet!TpoMon = TPOMON_NAC, CDec(nImpMN), 0)
      !ImpME = IIf(porstCOCpbDet!TpoMon = TPOMON_EXT, CDec(nImpME), 0)
      !TpoPvs = TPOPVS_PVS
      !TpoGnr = TPOGNR_CIE
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With
End Sub

Private Sub ppAjuste_CorrMon(cCodCta As String, ImpAD As Double, ImpAH As Double)
Dim dnImpIndB As Double, dnFactAj As Double, dnImpDif As Double
Dim dnSum As Double, dnSumAj As Double, dnfil As Long
   porstCOICM.MoveFirst
   porstCOICM.Find "MesICM= '" & gfMesAct(gsMesAct) & "'"  'con que mes si el cierre es al 13
   dnImpIndB = CDec(porstCOICM!ImpInd)
   dnSum = 0: dnSumAj = 0
   With porstCOCtaAcu
       If .RecordCount > 0 And porstCOCta.RecordCount > 0 Then
          .MoveFirst
          .Find "CodCta='" & cCodCta & "'"
         If Not .EOF Then
            For dnfil = 1 To 12
               If dnfil <= CInt(gfMesAct(gsMesAct)) Then
                   porstCOICM.MoveFirst
                   porstCOICM.Find "MesICM='" & Format(dnfil, "00") & "'"
                   dnFactAj = gfRedond(CDec(dnImpIndB) / CDec(IIf(porstCOICM!ImpInd = 0, 1, porstCOICM!ImpInd)), 3)
                   dnImpDif = CDec(porstCOCtaAcu.Fields("AcuD" & Format(dnfil, "00") & "_MN")) - CDec(porstCOCtaAcu.Fields("AcuH" & Format(dnfil, "00") & "_MN"))
                   dnSum = CDec(dnSum) + Format(CDec(dnImpDif), FORMATO_NUM_2)
                   dnSumAj = CDec(dnSumAj) + Format(CDec(dnImpDif) * CDec(dnFactAj), FORMATO_NUM_2)
                End If
            Next dnfil
         End If
         ImpAD = gfRedond(CDec(dnSum), 2)
         ImpAH = gfRedond(CDec(dnSumAj), 2)
      End If
   End With
End Sub

Private Sub ResuelveFormula(ByVal s_Cadena As String, ImpAD As Double, ImpAH As Double, bIndAMo As Boolean)
   Static sVariable As String, sCaso As String, sSigno As String
   Static nInicio As Integer, nFinal As Integer, nLen As Integer, nContador As Integer
   Static nImpDebe As Double, nImpHaber As Double
   Static nSImpDebe As Double, nSImpHaber As Double

   nInicio = 1: nFinal = 1: nLen = 0: nContador = 0
   nImpDebe = 0: nImpHaber = 0: nSImpDebe = 0: nSImpHaber = 0: 'nImpSaldo(0) = 0: nImpSaldo(1) = 0
   sCaso = Left(s_Cadena, 1)

   Do While nContador <= Len(s_Cadena)
     Select Case sCaso
       Case "["         ' Cuenta
         nInicio = (InStr(nInicio, s_Cadena, "[", vbTextCompare)) + 1
         nFinal = InStr(nInicio, s_Cadena, "]", vbTextCompare)
         nLen = (nFinal - nInicio)
         sVariable = Mid$(s_Cadena, nInicio, nLen)
         sSigno = Left(sVariable, 1)
         sVariable = IIf(IsNumeric(sSigno), sVariable, Mid(sVariable, 2))
         nImpDebe = 0: nImpHaber = 0
         With porstCOCta
            .MoveFirst
            .Find "CodCta='" & sVariable & "'"
            If Not .EOF Then
               If !IndMoe = 1 Then
                  ppAjuste_CorrMon sVariable, nImpDebe, nImpHaber
               Else
                  'INDICADOR DE AJUSTE EN MANTEN. CUENTA CIERRE IndAMo
                  If bIndAMo Then
                     ppAjuste_CorrMon sVariable, nImpDebe, nImpHaber
                  Else
                     If porstCOCtaAcu.RecordCount > 0 Then porstCOCtaAcu.MoveFirst
                     porstCOCtaAcu.Find "CodCta='" & sVariable & "'"
                     If Not porstCOCtaAcu.EOF Then
                        nImpDebe = CDec(porstCOCtaAcu!cAcuD_MN)
                        nImpHaber = CDec(porstCOCtaAcu!cAcuH_MN)
                     End If
                  End If
               End If
            End If
         End With
         sCaso = Mid(s_Cadena, nFinal + 1, 1)
         nContador = nFinal
       Case "+"         ' Signo Positivo
         nSImpDebe = CDec(nSImpDebe) + CDec(nImpDebe)
         nSImpHaber = CDec(nSImpHaber) + CDec(nImpHaber)
         sCaso = Mid(s_Cadena, nFinal + 2, 1)
         nContador = nFinal + 1
       Case "-"         ' Signo Negativo
         nSImpDebe = CDec(nSImpDebe) - CDec(nImpDebe)
         nSImpHaber = CDec(nSImpHaber) - CDec(nImpHaber)
         sCaso = Mid(s_Cadena, nFinal + 2, 1)
         nContador = nFinal + 1
       Case Else        ' Otro Caso
         nSImpDebe = CDec(nSImpDebe) + CDec(nImpDebe)
         nSImpHaber = CDec(nSImpHaber) + CDec(nImpHaber)
         nContador = nContador + 1
     End Select
    Loop
    ImpAD = CDec(nSImpDebe)
    ImpAH = CDec(nSImpHaber)
End Sub

