VERSION 5.00
Begin VB.Form frmTCprMasCCo 
   Caption         =   "[Entidad]"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3480
      TabIndex        =   2
      Top             =   480
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
      Left            =   1140
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   922
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1020
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
         Picture         =   "frmTCprMasCCo.frx":0000
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
         Picture         =   "frmTCprMasCCo.frx":014A
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
         Picture         =   "frmTCprMasCCo.frx":024C
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
         Picture         =   "frmTCprMasCCo.frx":034E
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
         Picture         =   "frmTCprMasCCo.frx":0498
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   338
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
         Picture         =   "frmTCprMasCCo.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   5040
      Picture         =   "frmTCprMasCCo.frx":07EC
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   615
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
      Index           =   1
      Left            =   60
      TabIndex        =   15
      Top             =   540
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
      Index           =   2
      Left            =   780
      TabIndex        =   14
      Top             =   540
      Width           =   315
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
      Index           =   3
      Left            =   3120
      TabIndex        =   13
      Top             =   540
      Width           =   300
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
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   3735
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "frmTCprMasCCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
']

Private Sub Form_Load()
   pbValidada = False

   Me.KeyPreview = True
   
   With frmTCprGrd                     'Cambiar Formulario de Grid.
    '[Datos                            'Cambiar.
      txtDato(0).MaxLength = .uorstCOCprDocCCo!codcco.DefinedSize
      'txtDato(1).MaxLength = .uorstCOCprDocCCo!ImpCCo_MN.DefinedSize
      'txtDato(2).MaxLength = .uorstCOCprDocCCo!ImpCCo_ME.DefinedSize
      txtDato(1).MaxLength = 14
      txtDato(2).MaxLength = 14
      
    ']
   End With
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   cmdRetroceder.Enabled = (Not pbNuevo)
   cmdCorregir.Enabled = (Not pbNuevo)
   cmdAvanzar.Enabled = (Not pbNuevo)
   upHabilitacion pbNuevo
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(4, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "C.Costo :", "Importe :", "MN", "ME")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "C.Center :", "Amount :", "NC", "FC")
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

Private Sub Form_KeyPress(KeyAscii As Integer)
''   Call gpTeclasData2(KeyAscii)
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
   txtDato(1).SetFocus
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmTCprGrd                     'Cambiar Formulario de Grid.
'      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstCOCprDocCCo.AddNew
      End If
      upDatosDesconectados 0
      With .uorstCOCprDocCCo
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
            !FyHMdf = Now
         End If
         .Update
      End With
'      .uorstCCCfg.Update
'      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.

      If pbNuevo Then
         .uorstCOCprDocCCo.Requery
         .upDatosGrid
''       '[B�squeda de llave actual.     'Cambiar.
''         .uorstCOCprDocCta.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
''       ']
          cmdGrabar.Enabled = False
          cmdDeshacer.Enabled = False
          cmdAvanzar.Enabled = False
          cmdRetroceder.Enabled = False
          cmdCorregir.Enabled = False
          upHabilitacion True
   
         upDatosPredeterminados
       '[Dato con el foco al a�adir.   'Cambiar.
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
  
'   frmTCprGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Private Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. A�adir �ndices.
   Case 0
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
'[ARREGLAR: Retrocede si Shift est� presionado.
   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.
 
 '[Convierte a may�sculas.
'   If Index = 0 Then                   'Cambiar (a�adir �ndices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
   If Index > 0 Then
      With frmTCpr
         If .chkMonedaActiva.Value = vbChecked Then
            If Index = 1 Then
               If .cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                  txtDato(2).Text = Format(gfRedond(CDec(txtDato(1).Text) / CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               ElseIf CDec(txtDato(2).Text) = 0 Then
                  txtDato(2).Text = Format(gfRedond(CDec(txtDato(1).Text) / CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               End If
            End If
            If Index = 2 Then
               If .cboTpoMon.ListIndex = TPOMON_EXT_IND Then
                  txtDato(1).Text = Format(gfRedond(CDec(txtDato(2).Text) * CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               ElseIf CDec(txtDato(1).Text) = 0 Then
                  txtDato(1).Text = Format(gfRedond(CDec(txtDato(2).Text) * CDec(.txtDato(4).Text), 2), FORMATO_NUM_1)
               End If
            End If
         End If
      End With
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
Dim dnContador As Byte
Dim dvRegistroActual As Variant

   On Error GoTo Err

  'Completa con ceros a la izquierda.
   Select Case Index
   Case 0                              'Cambiar (a�adir �ndices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Asigna 0 a campos num�ricos si est�n vac�os.
   Case 1, 2                             'Cambiar (a�adir �ndices).
      If txtDato.Item(Index).Text = "" Then
         txtDato.Item(Index).Text = 0
      End If
   End Select

  'Asigna 0 a campos num�ricos si est�n vac�os.
''   Select Case Index
''   Case 1, 2                           'Cambiar (a�adir �ndices).
''      Cancel = ppAyuDet(Index)
''      If Cancel Then Exit Sub
''   End Select

  'Da formato.
   Select Case Index
   Case 1, 2
      txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
   End Select

  'Busca el dato en su tabla principal.
   Select Case Index
   Case 0                              'Cambiar (a�adir �ndices).
    If Len(Trim(txtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      With frmTCprGrd.uorstCOCprDocCCo
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            .MoveFirst
            .Find "cLlave2='" & frmTCpr.txtLlave(0).Text & frmTCpr.txtLlave(1).Text & frmTCpr.txtLlave(2).Text & frmTCpr.txtLlave(3).Text & frmTCprMasGrd.unIndice & frmTCprGrd.uorstCOCprDocCta!orden & frmTCprMasGrd.dgrMain.Columns(0).Text & txtDato(0).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistroActual
         End If
      End With
'      If pbNuevo Then
'         For dnContador = 1 To txtDato.Count - 1
'            txtDato(dnContador).Enabled = True
'         Next
'         txtDato(1).SetFocus
'      End If
      upHabilitacion True
      cmdGrabar.Enabled = True
    Else
      upHabilitacion False
      cmdGrabar.Enabled = False
    End If
    cmdDatoAyud(0).Enabled = True
 
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Dim s_PedidoCco  As String
  
  s_PedidoCco = "AND indpdocpr='" & IIf(frmTCpr.txtDato(45).Text = "", INDCCO_INA, INDCCO_ACT) & "' "
  Select Case tnIndex
  Case 0                              'Cambiar (a�adir �ndices).
    modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 " & s_PedidoCco, txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
    txtDato(tnIndex).Text = frmOAyuBus.uvDato1
    lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
  End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
  Dim s_PedidoCco As String
  
  Select Case tnIndex                 'Cambiar.
   Case 0
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With frmTCprGrd.uorstCoCCo
      .MoveFirst
      .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetCCo), "", !DetCCo)
      End If
    End With
  End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmTCprGrd.uorstCOCprDocCCo    'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !codemp = gsCodEmp
            !pdoano = gsAnoAct
            !CodAux = frmTCpr.txtLlave(0).Text
            !codtdc = frmTCpr.txtLlave(1).Text
            !SerDoc = frmTCpr.txtLlave(2).Text
            !NroDoc = frmTCpr.txtLlave(3).Text
            !tpocnc = frmTCprMasGrd.unIndice
            !orden = frmTCprGrd.uorstCOCprDocCta!orden
            !codcta = frmTCprMasGrd.dgrMain.Columns(0).Text
         End If

        'Datos.
         !codcco = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
         !ImpCCo_MN = CDec(txtDato(1).Text)
         !ImpCCo_ME = CDec(txtDato(2).Text)
      Else
        'Datos.
         txtDato(0).Text = IIf(IsNull(!codcco), "", !codcco)
         txtDato(1).Text = Format(!ImpCCo_MN, FORMATO_NUM_1)
         txtDato(2).Text = Format(!ImpCCo_ME, FORMATO_NUM_1)
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
   
   Resume
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
   Dim dnContador As Integer

  'Datos.
   txtDato(0).Text = ""
   txtDato(1).Text = Format(0, FORMATO_NUM_1)
   txtDato(2).Text = Format(0, FORMATO_NUM_1)

'///Angel 22/12/2003
'///Envio de valor desde la pantalla de inicio de registro
   If pbNuevo Then
      If frmTCpr.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
         txtDato(1).Text = Format(frmTCprGrd.uorstCOCprDocCta!ImpCta_MN, FORMATO_NUM_1)
      Else
         txtDato(2).Text = Format(frmTCprGrd.uorstCOCprDocCta!ImpCta_ME, FORMATO_NUM_1)
      End If
      With frmTCprGrd.uorstCOCprDocCCo
         If Not .EOF And .RecordCount > 0 Then
            .MoveFirst
            Do
               If frmTCpr.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                  txtDato(1).Text = Format(CDec(txtDato(1).Text) - !ImpCCo_MN, FORMATO_NUM_1)
               Else
                  txtDato(2).Text = Format(CDec(txtDato(2).Text) - !ImpCCo_ME, FORMATO_NUM_1)
               End If
               .MoveNext
            Loop Until .EOF
            .MoveFirst
         End If
      End With
   End If
'///
   
  'Ayudas.
   For dnContador = 0 To 0
      lblDatoDeta(dnContador).Caption = ""
   Next
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
'   cboTpoMon.Enabled = tbHabilitar
'   chkMonedaActiva.Enabled = tbHabilitar
'   chkDesactivar.Enabled = tbHabilitar
'   dtpDato(3).Enabled = tbHabilitar
'   With mskDato
'      For dnContador = 0 To .Count - 1
'         .Item(dnContador).Enabled = tbHabilitar
'      Next
'   End With
   With txtDato
      For dnContador = 0 To .Count - 1
'/// Angel 12/12/2003
'/// Se agrego por el motivo de no permitir el cambio de un C.Costo al momento de corregir
         If dnContador = 0 Then
            .Item(dnContador).Enabled = pbNuevo
         Else
            .Item(dnContador).Enabled = tbHabilitar
         End If
'         .Item(dncontador).Enabled = tbHabilitar
'///
      Next
   End With

  'Ayudas.
   cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar
End Sub

'[Propio del formulario.

']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo

  'Habilita o no los botones.
   cmdDeshacer.Enabled = IIf(pbNuevo = True, False, True)
End Property
