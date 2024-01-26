VERSION 5.00
Begin VB.Form frmTVtaMasCta 
   Caption         =   "[Entidad]"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   7440
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1980
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   13
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
         Picture         =   "frmTVtaMasCta.frx":0000
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
         Picture         =   "frmTVtaMasCta.frx":014A
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
         Picture         =   "frmTVtaMasCta.frx":024C
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
         Picture         =   "frmTVtaMasCta.frx":034E
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
         Picture         =   "frmTVtaMasCta.frx":0498
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
         Picture         =   "frmTVtaMasCta.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   360
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
      Left            =   3600
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
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7140
      Picture         =   "frmTVtaMasCta.frx":07EC
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
      Top             =   120
      Width           =   975
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
      Left            =   840
      TabIndex        =   14
      Top             =   540
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
      Left            =   60
      TabIndex        =   12
      Top             =   540
      Width           =   570
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   60
      TabIndex        =   11
      Top             =   180
      Width           =   555
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
      Left            =   1620
      TabIndex        =   10
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmTVtaMasCta"
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
    '[Llaves                           'Cambiar
      TxtDato(0).MaxLength = .uorstCOVtaDocCta!CodCta.DefinedSize
    ']
    '[Datos                            'Cambiar.
      'txtDato(1).MaxLength = .uorstCOVtaDocCta!ImpCta_MN.DefinedSize
      'txtDato(2).MaxLength = .uorstCOVtaDocCta!ImpCta_ME.DefinedSize
      TxtDato(1).MaxLength = 14
      TxtDato(2).MaxLength = 14
      
   End With
'//Raul 12/12/2003
'// ¿?
'   cmdRetroceder.Enabled = (Not pbNuevo)
'   cmdCorregir.Enabled = (Not pbNuevo)
'   cmdAvanzar.Enabled = (Not pbNuevo)
'//
'/// Angel 12/12/2003
'// Habilitar los botones cuando es Nuevo
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
'   cmdSalir.Enabled = pbNuevo
   cmdCorregir.Enabled = (Not pbNuevo)
   cmdRetroceder.Enabled = (Not pbNuevo)
   cmdAvanzar.Enabled = (Not pbNuevo)
'//   cmdGrabar.Enabled = False
'//   cmdDeshacer.Enabled = False
'///
   upHabilitacion pbNuevo
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
   TxtDato(1).SetFocus
 ']
End Sub

Private Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmTVtaGrd                     'Cambiar Formulario de Grid.
'      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
      If pbNuevo Then
         .uorstCOVtaDocCta.AddNew
      End If
      upDatosDesconectados 0
      With .uorstCOVtaDocCta
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
         .uorstCOVtaDocCta.Requery
         .upDatosGrid
''       '[Búsqueda de llave actual.     'Cambiar.
''         .uorstCOVtaDocCta.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
''       ']
'''         cmdGrabar.Enabled = False
'''         upHabilitacion False
   
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
   If Index = 0 Then
   
   ElseIf Index > 0 Then
      With frmTVta
         If .chkMonedaActiva.Value = vbChecked Then
            If Index = 1 Then
               If .cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                  TxtDato(2).Text = Format(gfRedond(CDec(TxtDato(1).Text) / CDec(.TxtDato(4).Text), 2), FORMATO_NUM_1)
               ElseIf CDec(TxtDato(2).Text) = 0 Then
                  TxtDato(2).Text = Format(gfRedond(CDec(TxtDato(1).Text) / CDec(.TxtDato(4).Text), 2), FORMATO_NUM_1)
               End If
            End If
            If Index = 2 Then
               If .cboTpoMon.ListIndex = TPOMON_EXT_IND Then
                  TxtDato(1).Text = Format(gfRedond(CDec(TxtDato(2).Text) * CDec(.TxtDato(4).Text), 2), FORMATO_NUM_1)
               ElseIf CDec(TxtDato(1).Text) = 0 Then
                  TxtDato(1).Text = Format(gfRedond(CDec(TxtDato(2).Text) * CDec(.TxtDato(4).Text), 2), FORMATO_NUM_1)
               End If
            End If
         End If
      End With
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistroActual As Variant
  
  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 1, 21 To 28                    'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

  'Asigna 0 a campos numéricos si están vacíos.
   Select Case Index
   Case 1, 2                           'Cambiar (añadir índices).
      If TxtDato(Index).Text = "" Then
         TxtDato(Index).Text = 0
      End If
   End Select

  'Da formato.
   Select Case Index
   Case 1, 2
      TxtDato(Index).Text = Format(TxtDato(Index).Text, FORMATO_NUM_1)
   End Select

  'Busca el dato en su tabla principal.
   Select Case Index
   Case 0                              'Cambiar (añadir índices).
    If Len(Trim(TxtDato(Index).Text)) <> 0 Then
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
      '[
      With frmTVtaGrd.uorstCOVtaDocCta
         If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            .MoveFirst
'///Angel 12/12/2003
'/// Solo se agrego el la propiedad Text en TxtDato(0) al final de la instruccion de busqueda
             .Find "cLlave2='" & frmTVta.txtLlave(0).Text & frmTVta.txtLlave(1).Text & frmTVta.txtLlave(2).Text & frmTVtaMasGrd.unIndice & TxtDato(0).Text & "'"
'///
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
      upHabilitacion False
      cmdGrabar.Enabled = False
    End If
    cmdDatoAyud(0).Enabled = True
      ']
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", TxtDato(tnIndex).Text, 0, 0, Me.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + TxtDato(tnIndex).Left
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
      With frmTVtaGrd.uorstCOCta
         .MoveFirst
         .Find "CodCta='" & TxtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & !DetCta
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmTVtaGrd.uorstCOVtaDocCta    'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !CodTDc = frmTVta.txtLlave(0).Text
            !SerDoc = frmTVta.txtLlave(1).Text
            !NroDoc = frmTVta.txtLlave(2).Text
            !TpoCnc = frmTVtaMasGrd.unIndice
         End If

        'Datos.
'         !TpoMon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
'         !IndAjD = IIf(chkIndAjD.Value = vbChecked, INDAJD_ACT, INDAJD_INA)
'         !CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         !FehOpe = dtpDato(3).Value
'         !Tf1Cta = mskDato(0).Text
'         !CodMon = optTpoMon(1).Value
         !CodCta = IIf(TxtDato(0).Text = "", Null, TxtDato(0))
         !ImpCta_MN = TxtDato(1).Text
         !ImpCta_ME = TxtDato(2).Text
      Else
        'Datos.
'         cboTpoMon.ListIndex = IIf(!TpoMon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
'         chkIndCCo.Value = IIf(!IndCCo = INDCCO_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(!CodSoc), "", !CodSoc)
'         dtpDato(3).Value = !FehOpe
'         optTpoMon(1).Value = uorstMain!CodMon
'         mskDato(0).Text = IIf(IsNull(.uorstMain!Tf1Cta), "", .uorstMain!Tf1Cta)
         TxtDato(0).Text = IIf(IsNull(!CodCta), "", !CodCta)
         TxtDato(1).Text = Format(!ImpCta_MN, FORMATO_NUM_1)
         TxtDato(2).Text = Format(!ImpCta_ME, FORMATO_NUM_1)
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
   TxtDato(0).Text = ""
   TxtDato(1).Text = Format(0, FORMATO_NUM_1)
   TxtDato(2).Text = Format(0, FORMATO_NUM_1)
'///Angel 22/12/2003
'///Envio de valor desde la pantalla de inicio de registro
   If pbNuevo Then
      If frmTVta.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
         TxtDato(1).Text = Format(frmTVta.TxtDato(frmTVtaMasGrd.unIndice + 4).Text, FORMATO_NUM_1)
      Else
         TxtDato(2).Text = Format(frmTVta.TxtDato(frmTVtaMasGrd.unIndice + 11).Text, FORMATO_NUM_1)
      End If
      With frmTVtaGrd.uorstCOVtaDocCta
         If Not .EOF And .RecordCount > 0 Then
            .MoveFirst
            Do
               If frmTVta.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                  TxtDato(1).Text = Format(CDec(TxtDato(1).Text) - !ImpCta_MN, FORMATO_NUM_1)
               Else
                  TxtDato(2).Text = Format(CDec(TxtDato(2).Text) - !ImpCta_ME, FORMATO_NUM_1)
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
   With TxtDato
'/// Angel 12/12/2003
'/// Se agrego condicion para no permitir la correccion del dato Cuenta, solo para los importes
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

End Property
