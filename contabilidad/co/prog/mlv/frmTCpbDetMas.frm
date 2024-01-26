VERSION 5.00
Begin VB.Form frmTCpbDetMas 
   Caption         =   "[Entidad]"
   ClientHeight    =   1935
   ClientLeft      =   3405
   ClientTop       =   3285
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   7275
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "##,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
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
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "##,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
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
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1980
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3480
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
         Height          =   560
         Left            =   2690
         Picture         =   "frmTCpbDetMas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdDeshacer 
         Caption         =   "&Deshacer"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   1950
         Picture         =   "frmTCpbDetMas.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   1220
         Picture         =   "frmTCpbDetMas.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdCorregir 
         Caption         =   "&Corregir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   480
         Picture         =   "frmTCpbDetMas.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdAvanzar 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         Picture         =   "frmTCpbDetMas.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   345
         Width           =   360
      End
      Begin VB.CommandButton cmdRetroceder 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         Picture         =   "frmTCpbDetMas.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   6840
      Picture         =   "frmTCpbDetMas.frx":07EC
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
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
      Left            =   660
      TabIndex        =   0
      Top             =   135
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   15
      X2              =   7200
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
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
      Left            =   60
      TabIndex        =   15
      Top             =   780
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "M.N."
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
      Left            =   840
      TabIndex        =   14
      Top             =   780
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "M.E."
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
      Left            =   3240
      TabIndex        =   13
      Top             =   780
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F. Caja:"
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
      Left            =   60
      TabIndex        =   11
      Top             =   180
      Width           =   540
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
      Left            =   1305
      TabIndex        =   10
      Top             =   135
      Width           =   5535
   End
End
Attribute VB_Name = "frmTCpbDetMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmTCpbDetMasGrd                     'Cambiar Formulario de Grid.
    '[Datos                            'Cambiar.
      txtDato(0).MaxLength = .uorstMain!CodFjo.DefinedSize
      txtDato(1).MaxLength = 14
      txtDato(2).MaxLength = 14
      txtDato(1).TabIndex = Choose(frmTCpbDet.cboTpoMon.ListIndex + 1, 1, 2)
      txtDato(2).TabIndex = Choose(frmTCpbDet.cboTpoMon.ListIndex + 1, 2, 1)    ']
   End With
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
      cmdCorregir.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
   
End Sub

Private Sub Form_Activate()
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub cmdCorregir_Click()
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   cmdSalir.Enabled = True
   upHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato((frmTCpbDet.cboTpoMon.ListIndex + 1)).SetFocus
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmTCpbDetMasGrd                     'Cambiar Formulario de Grid.
      frmTCpbGrd.uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstMain.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
            !FyHMdf = Now
         End If
         .Update
      End With
      frmTCpbGrd.uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      frmTCpbDetMasGrd.uorstMain_Grd.Requery
      frmTCpbDetMasGrd.ppDatosGrid
   
      If pbNuevo Then
         upDatosPredeterminados
       '[Dato con el foco al añadir.   'Cambiar.
         txtDato(0).SetFocus
       ']
      Else
         cmdRetroceder.Enabled = True
         cmdAvanzar.Enabled = True
         cmdCorregir.Enabled = True
         cmdGrabar.Enabled = False
         cmdDeshacer.Enabled = False
         upHabilitacion False
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
  
   frmTCpbGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Private Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtDato(Index).SetFocus
   Case 4
      txtDato(Index).SetFocus
      
'   Case 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
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

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
   If Index = 1 Then
     If frmTCpbDet.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
       txtDato(2).Text = Format(gfRedond(CDec(txtDato(1).Text) / CDec(frmTCpbDet.txtDato(8).Text), 2), FORMATO_NUM_1)
     ElseIf CDec(txtDato(2).Text) = 0 Then
       txtDato(2).Text = Format(gfRedond(CDec(txtDato(1).Text) / CDec(frmTCpbDet.txtDato(8).Text), 2), FORMATO_NUM_1)
     End If
   ElseIf Index = 2 Then
     If frmTCpbDet.cboTpoMon.ListIndex = TPOMON_EXT_IND Then
       txtDato(1).Text = Format(gfRedond(CDec(txtDato(2).Text) * CDec(frmTCpbDet.txtDato(8).Text), 2), FORMATO_NUM_1)
     ElseIf CDec(txtDato(1).Text) = 0 Then
       txtDato(1).Text = Format(gfRedond(CDec(txtDato(2).Text) * CDec(frmTCpbDet.txtDato(8).Text), 2), FORMATO_NUM_1)
     End If
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
   Dim dvRegistroActual As Variant

  'Asigna 0 a campos numéricos si están vacíos. y da formato
   Select Case Index
   Case 1, 2                           'Cambiar (añadir índices).
      If txtDato(Index).Text = "" Or Not IsNumeric(txtDato(Index).Text) Then
         txtDato(Index).Text = 0
      End If
      txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
   End Select

  'Busca el dato en su tabla principal.
   Select Case Index
   Case 0                              'Cambiar (añadir índices).
    If Len(Trim(txtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      '[
      With frmTCpbDetMasGrd.uorstMain_Grd
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            .MoveFirst
            .Find "CodFjo='" & txtDato(0).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistroActual
         End If
      End With
      cmdGrabar.Enabled = True
      upHabilitacion True
    Else
      lblDatoDeta(Index).Caption = ""
      cmdGrabar.Enabled = False
      upHabilitacion False
    End If
    cmdDatoAyud(0).Enabled = True
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.Fjo_Cod IIf(ps_Plataforma = pSrvMySql, "Length ", "Len") & "(CodFjo)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmTCpbDetMasGrd.uorstCOFjo
         .MoveFirst
         .Find "CodFjo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmTCpbDetMasGrd.uorstCOFjo!DetFjo
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
  Dim uorstTmp As ADODB.Recordset
  Dim nNroOrden As Integer, pn_NroItem As Integer
  Dim sSentencia As String
   
  On Error GoTo Err

  With frmTCpbDetMasGrd.uorstMain    'Cambiar RecordSet.
    If tnFase = 0 Then
      ' Llaves.
      If pbNuevo Then
        Set uorstTmp = New ADODB.Recordset
        pn_NroItem = frmTCpbDet.pnNroIte
        nNroOrden = 0
        ' Si numero de item de detalle es cero
        If pn_NroItem = 0 Then
          With uorstTmp
            .ActiveConnection = frmTCpbGrd.uocnnMain
            .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroIte), 0) AS cUltIte "
            .Source = .Source & "FROM COCpbDet "
            .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
            .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
            .Source = .Source & "AND MesPvs='" & gsMesAct & "' "
            .Source = .Source & "AND coddro='" & frmTCpbCab.txtLlave(0).Text & "' "
            .Source = .Source & "AND NroCpb='" & frmTCpbCab.txtLlave(1).Text & "'"
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open
            pn_NroItem = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
          .Close
          End With
        End If
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !MesPvs = gsMesAct
        !CodDro = frmTCpbCab.txtLlave(0).Text
        !NroCpb = frmTCpbCab.txtLlave(1).Text
        !NroIte = pn_NroItem
        With uorstTmp
          .ActiveConnection = frmTCpbGrd.uocnnMain
          .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroOrd), 0) AS cUltOrd "
          .Source = .Source & "FROM " & ps_Prefijo & "TmpCoCpbDetFjo "
          .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
          .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
          .Source = .Source & "AND MesPvs='" & gsMesAct & "' "
          .Source = .Source & "AND CodDro='" & frmTCpbCab.txtLlave(0).Text & "' "
          .Source = .Source & "AND NroCpb='" & frmTCpbCab.txtLlave(1).Text & "' "
          .Source = .Source & "AND NroIte=" & pn_NroItem
        ' .CursorLocation = adUseClient   'Es el Default.
          .CursorType = adOpenForwardOnly
          .LockType = adLockReadOnly
          .Open
          nNroOrden = IIf(IsNull(!cUltOrd), 1, !cUltOrd + 1)
          .Close
        End With
        !NroOrd = nNroOrden
        !CodFjo = IIf(txtDato(0).Text = "", Null, txtDato(0))
      End If
      ' Datos.
      !CodCta = IIf(frmTCpbDet.txtDato(0).Text = "", Null, frmTCpbDet.txtDato(0).Text)
      !TpoCtb = IIf(frmTCpbDet.txtImporte(0).Text = 0 And frmTCpbDet.txtImporte(2).Text = 0, TPOCTB_HAB, TPOCTB_DEB)
      !ImpMN = CDec(txtDato(1).Text)
      !ImpME = CDec(txtDato(2).Text)
    Else
      ' Datos.
      txtDato(0).Text = IIf(IsNull(!CodFjo), "", !CodFjo)
      txtDato(1).Text = Format(!ImpMN, FORMATO_NUM_1)
      txtDato(2).Text = Format(!ImpME, FORMATO_NUM_1)
      '[ Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      ppAyuDet 0
    End If
  End With
      
  Exit Sub
Err:
   gpErrores
   
   Resume
      
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
  Static uorstTmp As ADODB.Recordset
  
  ' Datos
  txtDato(0).Text = ""
  txtDato(1).Text = Format(0, FORMATO_NUM_1)
  txtDato(2).Text = Format(0, FORMATO_NUM_1)
   
  '[ Obtengo los importes restantes
  If pbNuevo Then
    txtDato(1).Text = Format(CDec(frmTCpbDet.txtImporte(0).Text) + CDec(frmTCpbDet.txtImporte(1).Text), FORMATO_NUM_1)
    txtDato(2).Text = Format(CDec(frmTCpbDet.txtImporte(2).Text) + CDec(frmTCpbDet.txtImporte(3).Text), FORMATO_NUM_1)
    Set uorstTmp = New ADODB.Recordset
    With uorstTmp
      .ActiveConnection = frmTCpbGrd.uocnnMain
      .Source = "SELECT ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(ImpMN), 0), 2) AS ImpTotMN, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(ImpME), 0), 2) AS ImpTotME "
      .Source = .Source & "FROM " & ps_Prefijo & "TmpCoCpbDetFjo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
      .Open
      txtDato(1).Text = Format(CDec(txtDato(1).Text) - !ImpTotMN, FORMATO_NUM_1)
      txtDato(2).Text = Format(CDec(txtDato(2).Text) - !ImpTotME, FORMATO_NUM_1)
      .Close
    End With
  End If
  ']
  ' Ayudas.
  lblDatoDeta(0).Caption = ""
  Set uorstTmp = Nothing
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
     For dnContador = 0 To .Count - 1
       If dnContador = 0 Then
         .Item(dnContador).Enabled = pbNuevo
       Else
         .Item(dnContador).Enabled = tbHabilitar
       End If
     Next
   End With

  'Ayudas.
   cmdDatoAyud(0).Enabled = pbNuevo
   lblDatoDeta(0).Enabled = pbNuevo
End Sub

'[Propio del formulario.

']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
End Property

