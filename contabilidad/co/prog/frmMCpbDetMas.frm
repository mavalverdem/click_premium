VERSION 5.00
Begin VB.Form frmMCpbDetMas 
   Caption         =   "[Entidad Tipo Asiento]"
   ClientHeight    =   1935
   ClientLeft      =   3405
   ClientTop       =   3285
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleMode       =   0  'User
   ScaleWidth      =   11308.89
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
      Left            =   645
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   6825
      Picture         =   "frmMCpbDetMas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1965
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1185
      Width           =   3480
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
         Picture         =   "frmMCpbDetMas.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   360
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
         Picture         =   "frmMCpbDetMas.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   345
         Width           =   360
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
         Picture         =   "frmMCpbDetMas.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmMCpbDetMas.frx":0648
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmMCpbDetMas.frx":074A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   720
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
         Height          =   560
         Left            =   2690
         Picture         =   "frmMCpbDetMas.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
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
      Left            =   3585
      TabIndex        =   1
      Top             =   705
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
      Index           =   1
      Left            =   1185
      TabIndex        =   0
      Top             =   705
      Width           =   1815
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
      Left            =   1290
      TabIndex        =   15
      Top             =   120
      Width           =   5535
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
      Left            =   45
      TabIndex        =   14
      Top             =   165
      Width           =   540
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
      Left            =   3225
      TabIndex        =   13
      Top             =   765
      Width           =   300
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
      Left            =   825
      TabIndex        =   12
      Top             =   765
      Width           =   315
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
      Left            =   45
      TabIndex        =   11
      Top             =   765
      Width           =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   0
      X2              =   11168.99
      Y1              =   555
      Y2              =   555
   End
End
Attribute VB_Name = "frmMCpbDetMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean
Private rcMesAct As String
'rcMesAct
Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmMCpbDetMasGrd                     'Cambiar Formulario de Grid.
    '[Datos                            'Cambiar.
      TxtDato(0).MaxLength = .uorstMain!CodFjo.DefinedSize
      TxtDato(1).MaxLength = 14
      TxtDato(2).MaxLength = 14
      TxtDato(1).TabIndex = Choose(frmMCpbDet.cboTpoMon.ListIndex + 1, 1, 2)
      TxtDato(2).TabIndex = Choose(frmMCpbDet.cboTpoMon.ListIndex + 1, 2, 1)    ']
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
   TxtDato((frmMCpbDet.cboTpoMon.ListIndex + 1)).SetFocus
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmMCpbDetMasGrd                     'Cambiar Formulario de Grid.
      frmMCpbGrd.uocnnMain.BeginTrans            'INICIA TRANSACCION.
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
      frmMCpbGrd.uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      frmMCpbDetMasGrd.uorstMain_Grd.Requery
      frmMCpbDetMasGrd.ppDatosGrid
   
      If pbNuevo Then
         upDatosPredeterminados
       '[Dato con el foco al añadir.   'Cambiar.
         TxtDato(0).SetFocus
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
  
   frmMCpbGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
      TxtDato(Index).SetFocus
   Case 4
      TxtDato(Index).SetFocus
      
'   Case 3
'      mskDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   TxtDato(Index).SelStart = 0
   TxtDato(Index).SelLength = TxtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(TxtDato(Index))) + 1 = TxtDato(Index).MaxLength Then
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
     If frmMCpbDet.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
       TxtDato(2).Text = Format(gfRedond(CDec(TxtDato(1).Text) / CDec(frmMCpbDet.TxtDato(8).Text), 2), FORMATO_NUM_1)
     ElseIf CDec(TxtDato(2).Text) = 0 Then
       TxtDato(2).Text = Format(gfRedond(CDec(TxtDato(1).Text) / CDec(frmMCpbDet.TxtDato(8).Text), 2), FORMATO_NUM_1)
     End If
   ElseIf Index = 2 Then
     If frmMCpbDet.cboTpoMon.ListIndex = TPOMON_EXT_IND Then
       TxtDato(1).Text = Format(gfRedond(CDec(TxtDato(2).Text) * CDec(frmMCpbDet.TxtDato(8).Text), 2), FORMATO_NUM_1)
     ElseIf CDec(TxtDato(1).Text) = 0 Then
       TxtDato(1).Text = Format(gfRedond(CDec(TxtDato(2).Text) * CDec(frmMCpbDet.TxtDato(8).Text), 2), FORMATO_NUM_1)
     End If
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
   Dim dvRegistroActual As Variant

  'Asigna 0 a campos numéricos si están vacíos. y da formato
   Select Case Index
   Case 1, 2                           'Cambiar (añadir índices).
      If TxtDato(Index).Text = "" Or Not IsNumeric(TxtDato(Index).Text) Then
         TxtDato(Index).Text = 0
      End If
      TxtDato(Index).Text = Format(TxtDato(Index).Text, FORMATO_NUM_1)
   End Select

  'Busca el dato en su tabla principal.
   Select Case Index
   Case 0                              'Cambiar (añadir índices).
    If Len(Trim(TxtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      '[
      With frmMCpbDetMasGrd.uorstMain_Grd
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            .MoveFirst
            .Find "CodFjo='" & TxtDato(0).Text & "'"
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
      modAyuBus.Fjo_Cod IIf(ps_Plataforma = pSrvMySql, "Length ", "Len") & "(CodFjo)=4", TxtDato(tnIndex).Text, 0, 0, Me.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + TxtDato(tnIndex).Left
      TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0
      If TxtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMCpbDetMasGrd.uorstCOFjo
         .MoveFirst
         .Find "CodFjo='" & TxtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & frmMCpbDetMasGrd.uorstCOFjo!DetFjo
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

  With frmMCpbDetMasGrd.uorstMain    'Cambiar RecordSet.
    If tnFase = 0 Then
      ' Llaves.
      If pbNuevo Then
        Set uorstTmp = New ADODB.Recordset
        pn_NroItem = frmMCpbDet.pnNroIte
        nNroOrden = 0
        ' Si numero de item de detalle es cero
        If pn_NroItem = 0 Then
          With uorstTmp
            .ActiveConnection = frmMCpbGrd.uocnnMain
            .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroIte), 0) AS cUltIte "
            .Source = .Source & "FROM comacpbdet "
            .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
            .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
            .Source = .Source & "AND MesPvs='" & rcMesAct & "' "
            .Source = .Source & "AND coddro='" & frmMCpbCab.txtLlave(0).Text & "' "
            .Source = .Source & "AND NroCpb='" & frmMCpbCab.txtLlave(1).Text & "'"
            .CursorType = adOpenForwardOnly
            .LockType = adLockReadOnly
            .Open
            pn_NroItem = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
          .Close
          End With
        End If
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !mespvs = rcMesAct
        !coddro = frmMCpbCab.txtLlave(0).Text
        !NroCpb = frmMCpbCab.txtLlave(1).Text
        !NroIte = pn_NroItem
        With uorstTmp
          .ActiveConnection = frmMCpbGrd.uocnnMain
          .Source = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(NroOrd), 0) AS cUltOrd "
          .Source = .Source & "FROM " & ps_Prefijo & "TmpcomacpbdetFjo "
          .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
          .Source = .Source & "AND pdoano='" & gsAnoAct & "' "
          .Source = .Source & "AND MesPvs='" & rcMesAct & "' "
          .Source = .Source & "AND CodDro='" & frmMCpbCab.txtLlave(0).Text & "' "
          .Source = .Source & "AND NroCpb='" & frmMCpbCab.txtLlave(1).Text & "' "
          .Source = .Source & "AND NroIte=" & pn_NroItem
        ' .CursorLocation = adUseClient   'Es el Default.
          .CursorType = adOpenForwardOnly
          .LockType = adLockReadOnly
          .Open
          nNroOrden = IIf(IsNull(!cUltOrd), 1, !cUltOrd + 1)
          .Close
        End With
        !NroOrd = nNroOrden
        !CodFjo = IIf(TxtDato(0).Text = "", Null, TxtDato(0))
      End If
      ' Datos.
      !codcta = IIf(frmMCpbDet.TxtDato(0).Text = "", Null, frmMCpbDet.TxtDato(0).Text)
      !TpoCtb = IIf(frmMCpbDet.txtImporte(0).Text = 0 And frmMCpbDet.txtImporte(2).Text = 0, TPOCTB_HAB, TPOCTB_DEB)
      !ImpMN = CDec(TxtDato(1).Text)
      !ImpME = CDec(TxtDato(2).Text)
    Else
      ' Datos.
      TxtDato(0).Text = IIf(IsNull(!CodFjo), "", !CodFjo)
      TxtDato(1).Text = Format(!ImpMN, FORMATO_NUM_1)
      TxtDato(2).Text = Format(!ImpME, FORMATO_NUM_1)
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
  TxtDato(0).Text = ""
  TxtDato(1).Text = Format(0, FORMATO_NUM_1)
  TxtDato(2).Text = Format(0, FORMATO_NUM_1)
   
  '[ Obtengo los importes restantes
  If pbNuevo Then
    TxtDato(1).Text = Format(CDec(frmMCpbDet.txtImporte(0).Text) + CDec(frmMCpbDet.txtImporte(1).Text), FORMATO_NUM_1)
    TxtDato(2).Text = Format(CDec(frmMCpbDet.txtImporte(2).Text) + CDec(frmMCpbDet.txtImporte(3).Text), FORMATO_NUM_1)
    Set uorstTmp = New ADODB.Recordset
    With uorstTmp
      .ActiveConnection = frmMCpbGrd.uocnnMain
      .Source = "SELECT ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(ImpMN), 0), 2) AS ImpTotMN, "
      .Source = .Source & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(ImpME), 0), 2) AS ImpTotME "
      .Source = .Source & "FROM " & ps_Prefijo & "TmpcomacpbdetFjo "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'      .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
      .Open
      TxtDato(1).Text = Format(CDec(TxtDato(1).Text) - !ImpTotMN, FORMATO_NUM_1)
      TxtDato(2).Text = Format(CDec(TxtDato(2).Text) - !ImpTotME, FORMATO_NUM_1)
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
   With TxtDato
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


