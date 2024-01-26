VERSION 5.00
Begin VB.Form frmTPdoMas 
   Caption         =   "[Entidad]"
   ClientHeight    =   2070
   ClientLeft      =   5160
   ClientTop       =   5355
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   7455
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   2
      Left            =   1200
      TabIndex        =   8
      Top             =   870
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   3
      Left            =   3360
      TabIndex        =   10
      Top             =   870
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   4
      Left            =   5505
      TabIndex        =   12
      Top             =   870
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   1
      Left            =   990
      TabIndex        =   4
      Top             =   510
      Width           =   615
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   1
      Left            =   7065
      Picture         =   "frmTPdoMas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   510
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   135
      Width           =   975
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   7065
      Picture         =   "frmTPdoMas.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   135
      Width           =   280
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1980
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1350
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
         Picture         =   "frmTPdoMas.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmTPdoMas.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmTPdoMas.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmTPdoMas.frx":06A2
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmTPdoMas.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmTPdoMas.frx":0996
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Diferencial"
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
      Index           =   5
      Left            =   5160
      TabIndex        =   11
      Top             =   900
      Width           =   300
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "C.Costo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   525
      Width           =   975
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
      Height          =   285
      Index           =   1
      Left            =   1590
      TabIndex        =   5
      Top             =   510
      Width           =   5475
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
      Height          =   285
      Index           =   0
      Left            =   1950
      TabIndex        =   2
      Top             =   135
      Width           =   5115
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   150
      Width           =   960
   End
   Begin VB.Label lblTexto 
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
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   900
      Width           =   570
   End
   Begin VB.Label lblTexto 
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
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   900
      Width           =   300
   End
   Begin VB.Label lblTexto 
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
      Index           =   4
      Left            =   3000
      TabIndex        =   9
      Top             =   900
      Width           =   300
   End
End
Attribute VB_Name = "frmTPdoMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
Private pnCta_IndCCo As Integer
Private pcCodCCo_Def As String


Private Sub Form_Load()
  pbValidada = False
  Me.KeyPreview = True
  
  With frmTPdoGrd                     'Cambiar Formulario de Grid.
    '[Datos                            'Cambiar.
    txtDato(0).MaxLength = .uorstCoDPeCta!codcta.DefinedSize
    txtDato(1).MaxLength = .uorstCoDPeCta!codcco.DefinedSize
    txtDato(2).MaxLength = 14
    txtDato(3).MaxLength = 14
    txtDato(4).MaxLength = 14
    txtDato(2).TabIndex = Choose(frmTPdo.cboTpoMon.ListIndex + 1, 8, 9)
    txtDato(3).TabIndex = Choose(frmTPdo.cboTpoMon.ListIndex + 1, 9, 8)    ']
    ']
  End With
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  cmdAvanzar.Enabled = (Not pbNuevo)
  cmdRetroceder.Enabled = (Not pbNuevo)
  cmdCorregir.Enabled = (Not pbNuevo)
  upHabilitacion pbNuevo
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(6, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuenta :", "C.Costo :", "Importe :", "MN", "ME", "DIF")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Account :", "C.Center :", "Amount :", "NC", "FC", "DIF")
  Next nElemento
  cmdGrabar.Caption = Choose(gsIdioma, "&Aceptar", "&Accept")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, False, True, True, aLabel
  ']
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
  upHabilitacion True
  
  '[Dato con el foco al corregir.       'Cambiar.
  txtDato((frmTPdo.cboTpoMon.ListIndex + 2)).SetFocus
  ']
End Sub

Private Sub cmdGrabar_Click()
  On Error GoTo Err
  
  If Len(Trim(txtDato(0).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If pnCta_IndCCo = INDCCO_ACT And Len(Trim(txtDato(1).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(1).SetFocus: Exit Sub
  With frmTPdoGrd                     'Cambiar Formulario de Grid.
    frmTPdoGrd.uocnnMain.BeginTrans            'INICIA TRANSACCION.
    If pbNuevo Then
      .uorstCoDPeCta.AddNew
    End If
    upDatosDesconectados 0
    With .uorstCoDPeCta
      If pbNuevo Then
        !UsrCre = gsAbvUsr
        !FyHCre = Now
      Else
        !UsrMdf = gsAbvUsr
        !FyHMdf = Now
      End If
      .Update
    End With
    frmTPdoGrd.uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
    
    If pbNuevo Then
      .uorstCoDPeCta.Requery
      frmTPdoMasGrd.ppDatosGrid
      '[Búsqueda de llave actual.     'Cambiar.
      .uorstCoDPeCta.Find "codcta='" & txtDato(0).Text & "'"
      ']
      cmdGrabar.Enabled = False
      upHabilitacion True
      
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
  frmTPdoGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
   Case 1
    txtDato(Index).SetFocus
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
  If KeyCode = vbKeyF2 Then ppAyuBus Index
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
  Select Case Index
   Case 0
    txtDato(Index + 1).Enabled = (txtDato(Index).Text <> "" And pnCta_IndCCo = INDCCO_ACT)
    cmdDatoAyud(Index + 1).Enabled = (txtDato(0).Text <> "" And pnCta_IndCCo = INDCCO_ACT)
   Case 2, 3
    ' Convierto importe en cero
    If CDec(txtDato(Index).Text) = 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      If Index = 2 And frmTPdo.cboTpoMon.ListIndex = TPOMON_EXT_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index + 1).Text) * CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 3 And frmTPdo.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index - 1).Text) / CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    ElseIf CDec(txtDato(Index).Text) <> 0 Then
      If Index = 2 And frmTPdo.cboTpoMon.ListIndex = TPOMON_NAC_IND And (txtDato(3).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index + 1).Text = Format(Round(CDec(txtDato(Index).Text) / CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 3 And frmTPdo.cboTpoMon.ListIndex = TPOMON_EXT_IND And (txtDato(2).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index - 1).Text = Format(Round(CDec(txtDato(Index).Text) * CDec(frmTPdo.txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    End If
  End Select
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  Dim dvRegistroActual As Variant

  'Completa con ceros a la izquierda.
  Select Case Index
   Case 0
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
    If lblDatoDeta(Index).Caption <> "" Then
      ' Cuenta repetida
      With frmTPdoGrd.uorstCoDPeCta
        If Not (.BOF Or .EOF) And .RecordCount > 0 Then
          dvRegistroActual = .Bookmark
          .MoveFirst
          .Find "codcta='" & txtDato(Index).Text & "'"
          If Not .EOF Then
            MsgBox TEXT_8007, vbExclamation
            If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
            Cancel = True
            Exit Sub
          End If
          .Bookmark = dvRegistroActual
        End If
      End With
      pnCta_IndCCo = frmTPdoGrd.uorstCOCta!IndCCo
      pcCodCCo_Def = IIf(IsNull(frmTPdoGrd.uorstCOCta!codcco_def), "", frmTPdoGrd.uorstCOCta!codcco_def)
      ' Actualizo los datos adicionales
      txtDato(1).Text = IIf(txtDato(1).Text = "", IIf(frmTPdo.psCodCCo_Pdo = "", pcCodCCo_Def, frmTPdo.psCodCCo_Pdo), txtDato(1).Text)
      txtDato(1).Text = IIf(pnCta_IndCCo = INDCCO_ACT, txtDato(1).Text, "")
      lblDatoDeta(1).Caption = IIf(pnCta_IndCCo = INDCCO_ACT, lblDatoDeta(1).Caption, "")
      ' Habilito controles
      txtDato(1).Enabled = (pnCta_IndCCo = INDCCO_ACT)
      cmdDatoAyud(1).Enabled = (pnCta_IndCCo = INDCCO_ACT)
      cmdGrabar.Enabled = True
      upHabilitacion True
    Else
      cmdGrabar.Enabled = False
      upHabilitacion False
    End If
   Case 1
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
   Case 2, 3, 4
    txtDato(Index).Text = Format(CDec(IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)), FORMATO_NUM_1)
  End Select
  Exit Sub

Err:
  gpErrores
  
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0
    modAyuBus.Cta_Cod "tpocta=" & TPOCTA_TRA & " AND estcta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1
    modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(codcco)=5 AND estcco='" & ESTCCO_ACT & "' AND indpdocpr='" & INDCCO_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
    With frmTPdoGrd.uorstCOCta
      If .RecordCount > 0 Then .MoveFirst
      .Find "codcta='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
      End If
    End With
   Case 1
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With frmTPdoGrd.uorstCoCCo
      If .RecordCount > 0 Then .MoveFirst
      .Find "codcco='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcco), "", !detcco)
      End If
    End With
  End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  On Error GoTo Err

  With frmTPdoGrd.uorstCoDPeCta    'Cambiar RecordSet.
    If tnFase = 0 Then
      ' Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !mespvs = gsMesAct
        !coddpe = frmTPdo.txtLlave(0).Text
        !pdocpr = frmTPdo.txtLlave(1).Text
        !codcta = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      End If
      ' Datos.
      !codcco = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
      !impcta_mn = CDec(txtDato(2).Text)
      !impcta_me = CDec(txtDato(3).Text)
      !impctadif = CDec(txtDato(4).Text)
    Else
      ' Llaves.
      txtDato(0).Text = IIf(IsNull(!codcta), "", !codcta)
      ' Datos.
      txtDato(1).Text = IIf(IsNull(!codcco), "", !codcco)
      txtDato(2).Text = Format(!impcta_mn, FORMATO_NUM_1)
      txtDato(3).Text = Format(!impcta_me, FORMATO_NUM_1)
      txtDato(4).Text = Format(!impctadif, FORMATO_NUM_1)
      
      txtDato(2).Tag = Format(txtDato(2).Text, FORMATO_NUM_1)
      txtDato(3).Tag = Format(txtDato(3).Text, FORMATO_NUM_1)
      txtDato(4).Tag = Format(txtDato(4).Text, FORMATO_NUM_1)
      '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      ppAyuDet 0
      ppAyuDet 1
    End If
  End With
      
  Exit Sub
Err:
   gpErrores
   
   Resume
      
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
  Dim sSentencia As String
  
  txtDato(0).Text = ""
  txtDato(1).Text = ""
  txtDato(2).Text = Format(0, FORMATO_NUM_1)
  txtDato(3).Text = Format(0, FORMATO_NUM_1)
  txtDato(4).Text = Format(0, FORMATO_NUM_1)
  '[ Obtengo los importes restantes
  If pbNuevo Then
    txtDato(2).Text = Format(CDec(frmTPdo.txtDato(4).Text), FORMATO_NUM_1)
    txtDato(3).Text = Format(CDec(frmTPdo.txtDato(5).Text), FORMATO_NUM_1)
    txtDato(4).Text = Format(CDec(frmTPdo.txtDato(6).Text), FORMATO_NUM_1)
    With frmTPdoGrd
      sSentencia = "SELECT ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impcta_mn), 0), 2) AS ImporteMN, "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impcta_me), 0), 2) AS ImporteME, "
      sSentencia = sSentencia & "ROUND(" & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(SUM(impctadif), 0), 2) AS ImporteDF "
      sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpcopdocprcta "
      Set .porstCancel = .uocnnMain.Execute(sSentencia)
      txtDato(2).Text = Format(CDec(txtDato(2).Text) - .porstCancel!ImporteMN, FORMATO_NUM_1)
      txtDato(3).Text = Format(CDec(txtDato(3).Text) - .porstCancel!ImporteME, FORMATO_NUM_1)
      txtDato(4).Text = Format(CDec(txtDato(4).Text) - .porstCancel!ImporteDF, FORMATO_NUM_1)
      .porstCancel.Close
    End With
  End If
  txtDato(2).Tag = Format(txtDato(2).Text, FORMATO_NUM_2)
  txtDato(3).Tag = Format(txtDato(3).Text, FORMATO_NUM_1)
  txtDato(4).Tag = Format(txtDato(4).Text, FORMATO_NUM_1)
  ']
  ' Ayudas.
  lblDatoDeta(0).Caption = ""
  lblDatoDeta(1).Caption = ""
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
