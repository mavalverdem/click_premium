VERSION 5.00
Begin VB.Form frmTVtaMasCCo 
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
      Index           =   1
      Left            =   1140
      TabIndex        =   1
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
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   922
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1020
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
         Picture         =   "frmTVtaMasCCo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmTVtaMasCCo.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
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
         Picture         =   "frmTVtaMasCCo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmTVtaMasCCo.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmTVtaMasCCo.frx":05A0
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmTVtaMasCCo.frx":06A2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   5040
      Picture         =   "frmTVtaMasCCo.frx":07EC
      Style           =   1  'Graphical
      TabIndex        =   10
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
      Left            =   3120
      TabIndex        =   15
      Top             =   540
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
      Left            =   780
      TabIndex        =   14
      Top             =   540
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
      Left            =   1320
      TabIndex        =   12
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label6 
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
      Left            =   60
      TabIndex        =   11
      Top             =   180
      Width           =   615
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
      TabIndex        =   9
      Top             =   540
      Width           =   570
   End
End
Attribute VB_Name = "frmTVtaMasCCo"
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
   
   With frmTVtaGrd                     'Cambiar Formulario de Grid.
    '[Datos                            'Cambiar.
      txtDato(0).MaxLength = .uorstCOVtaDocCCo!CodCCo.DefinedSize
      'txtDato(1).MaxLength = .uorstCOVtaDocCCo!ImpCCo_MN.DefinedSize
      'txtDato(2).MaxLength = .uorstCOVtaDocCCo!ImpCCo_ME.DefinedSize
      txtDato(1).MaxLength = 14
      txtDato(2).MaxLength = 14
    ']
   End With
   cmdGrabar.Enabled = pbNuevo
   cmdDeshacer.Enabled = pbNuevo
   cmdRetroceder.Enabled = (Not pbNuevo)
   cmdCorregir.Enabled = (Not pbNuevo)
   cmdAvanzar.Enabled = (Not pbNuevo)
   upHabilitacion (pbNuevo)
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
'///Angel 12/12/2003
'/// Se agrego habilitacion del boton salir
   cmdSalir.Enabled = True
'///
   upHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(1).SetFocus
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmTVtaGrd                     'Cambiar Formulario de Grid.
'      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstCOVtaDocCCo.AddNew
      End If
      upDatosDesconectados 0
      With .uorstCOVtaDocCCo
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
'            !FyHMdf = Now
         End If
         .Update
      End With
'      .uorstCCCfg.Update
'      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstCOVtaDocCCo.Requery
         .upDatosGrid
''       '[Búsqueda de llave actual.     'Cambiar.
''         .uorstCOVtaDocCta.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
''       ']
'''         cmdGrabar.Enabled = False
'''         upHabilitacion False
   
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
  
'   frmTVtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
 
 '[Convierte a mayúsculas.
'   If Index = 0 Then                   'Cambiar (añadir índices).
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
      With frmTVta
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
   On Error GoTo Err
   Dim dnContador As Integer
   Dim dvRegistroActual As Variant
'Completa con ceros a la izquierda.
   Select Case Index
   Case 0                              'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Select Case Index    'Asigna 0 a campos numéricos si están vacíos.
   Case 1, 2                             'Cambiar (añadir índices).
      If txtDato.Item(Index).Text = "" Then
         txtDato.Item(Index).Text = 0
      End If
   End Select

  'Asigna 0 a campos numéricos si están vacíos.
''   Select Case Index
''   Case 1, 2                           'Cambiar (añadir índices).
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
   Case 0                              'Cambiar (añadir índices).
    If Len(Trim(txtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      With frmTVtaGrd.uorstCOVtaDocCCo
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            .MoveFirst
             .Find "cLlave2='" & frmTVta.txtLlave(0).Text & frmTVta.txtLlave(1).Text & frmTVta.txtLlave(2).Text & frmTVtaMasGrd.unIndice & frmTVtaMasGrd.dgrMain.Columns(0).Text & txtDato(0).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistroActual <> -1 Then .Bookmark = dvRegistroActual
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistroActual
         End If
      End With
      
      upHabilitacion True
      cmdGrabar.Enabled = True
    Else
      upHabilitacion False
      cmdGrabar.Enabled = False
    End If
    cmdDatoAyud(0).Enabled = True
      
'///Angel 12/12/2003
'/// Agrega habilitacion de objetos textos de importes
'      If pbNuevo Then
'         For dnContador = 1 To txtDato.Count - 1
'            txtDato(dnContador).Enabled = True
'         Next
'         txtDato(1).SetFocus
'      End If
'///
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.CCo_Cod "Length(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
      With frmTVtaGrd.uorstCOCCo
         .MoveFirst
         .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & !DetCCo
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmTVtaGrd.uorstCOVtaDocCCo    'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !CodTDc = frmTVta.txtLlave(0).Text
            !SerDoc = frmTVta.txtLlave(1).Text
            !NroDoc = frmTVta.txtLlave(2).Text
            !TpoCnc = frmTVtaMasGrd.unIndice
            !CodCta = frmTVtaMasGrd.dgrMain.Columns(0).Text
         End If

        'Datos.
'         !TpoMon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
'         !IndAjD = IIf(chkIndAjD.Value = vbChecked, INDAJD_ACT, INDAJD_INA)
'         !CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         !FehOpe = dtpDato(3).Value
'         !Tf1Cta = mskDato(0).Text
'         !CodMon = optTpoMon(1).Value
         !CodCCo = IIf(txtDato(0).Text = "", Null, txtDato(0))
         !ImpCCo_MN = txtDato(1).Text
         !ImpCCo_ME = txtDato(2).Text
      Else
        'Datos.
'         cboTpoMon.ListIndex = IIf(!TpoMon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
'         chkIndCCo.Value = IIf(!IndCCo = INDCCO_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(!CodSoc), "", !CodSoc)
'         dtpDato(3).Value = !FehOpe
'         optTpoMon(1).Value = uorstMain!CodMon
'         mskDato(0).Text = IIf(IsNull(.uorstMain!Tf1Cta), "", .uorstMain!Tf1Cta)
         txtDato(0).Text = IIf(IsNull(!CodCCo), "", !CodCCo)
'/// Angel 12/12/2003
'/// Se cambio campos ImpCta por ImpCCo para Centros de Costos
         txtDato(1).Text = Format(!ImpCCo_MN, FORMATO_NUM_1)
         txtDato(2).Text = Format(!ImpCCo_ME, FORMATO_NUM_1)
'         txtDato(1).Text = Format(!ImpCta_MN, FORMATO_NUM_1)
'         txtDato(2).Text = Format(!ImpCta_ME, FORMATO_NUM_1)
'///
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
'   cboTpoMon.ListIndex = TPOMON_NAC_IND
'   chkEstado.Value = vbChecked
'   dtpDato(3).Value = Date
'   optTpoMon(1).Value = True
   txtDato(0).Text = ""
   txtDato(1).Text = Format(0, FORMATO_NUM_1)
   txtDato(2).Text = Format(0, FORMATO_NUM_1)
'///Angel 22/12/2003
'///Envio de valor desde la pantalla de inicio de registro
   If pbNuevo Then
      If frmTVta.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
         txtDato(1).Text = Format(frmTVtaGrd.uorstCOVtaDocCta!ImpCta_MN, FORMATO_NUM_1)
      Else
         txtDato(2).Text = Format(frmTVtaGrd.uorstCOVtaDocCta!ImpCta_ME, FORMATO_NUM_1)
      End If
      With frmTVtaGrd.uorstCOVtaDocCCo
         If Not .EOF And .RecordCount > 0 Then
            .MoveFirst
            Do
               If frmTVta.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
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
'/// Angel 12/12/2003
'/// Se agrego por el motivo de no permitir el cambio de un C.Costo al momento de corregir
      For dnContador = 0 To .Count - 1
         If dnContador = 0 Then
            .Item(dnContador).Enabled = pbNuevo
         Else
            .Item(dnContador).Enabled = tbHabilitar
         End If
'         .Item(dncontador).Enabled = tbHabilitar
      Next
'///
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


