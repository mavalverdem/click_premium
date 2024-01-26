VERSION 5.00
Begin VB.Form frmMPsp 
   Caption         =   "[Entidad]"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   1320
      Width           =   3495
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
         Index           =   2
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox CboTpoGru 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden:"
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
         Left            =   2400
         TabIndex        =   19
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1020
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   3420
      TabIndex        =   2
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   305
      Index           =   0
      Left            =   7200
      Picture         =   "frmMPsp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2010
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2520
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
         Picture         =   "frmMPsp.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmMPsp.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   338
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
         Picture         =   "frmMPsp.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
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
         Picture         =   "frmMPsp.frx":0648
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
         Picture         =   "frmMPsp.frx":074A
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmMPsp.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.TextBox txtLlave 
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
      Width           =   950
   End
   Begin VB.Label lblLlaveDeta 
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
      Left            =   1680
      TabIndex        =   16
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Importe M.E.:"
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
      Left            =   2460
      TabIndex        =   14
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Importe M.N.:"
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
      Top             =   840
      Width           =   930
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
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   7440
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMPsp"
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
   
   With frmMPspGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!CodCta.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
      'txtDato(0).MaxLength = .uorstMain.Fields("ImpMN_" & gsMesAct).DefinedSize
      'txtDato(1).MaxLength = .uorstMain.Fields("ImpME_" & gsMesAct).DefinedSize
      'corregido solo se podia ingresar 8 caracteres contando las comas y decimales
      txtDato(0).Text = .uorstMain.Fields("ImpMN_" & gsMesAct).DefinedSize
      txtDato(1).Text = .uorstMain.Fields("ImpME_" & gsMesAct).DefinedSize
      
      
      txtDato(2).MaxLength = .uorstMain.Fields("OrdRep").DefinedSize - 1
      With CboTpoGru
         .AddItem TPOGRU1_TXT_1, TPOGRU1_IND
         .AddItem TPOGRU2_TXT_1, TPOGRU2_IND
         .AddItem TPOGRU3_TXT_1, TPOGRU3_IND
      End With
    ']
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
End Sub

Private Sub Form_Activate()
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
   If txtLlave(0).Text <> "" Then ppAyuDet 0
 ']
''   If pbNuevo Then
''      With frmMPspGrd.porstUltOrdRep
''         .Open
''         txtDato(2).Text = !OrdRep
''         .Close
''      End With
''   End If
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMPspGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMPspGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMPspGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err

   With frmMPspGrd                     'Cambiar Formulario de Grid.
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
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
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain.Find "CodCta='" & txtLlave(0).Text & "'"
       ']
         cmdGrabar.Enabled = False
         upHabilitacion False
   
         upDatosPredeterminados
       '[Llave con el foco al añadir.  'Cambiar.
         txtLlave(0).SetFocus
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
  
   frmMPspGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtLlave(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
'   Select Case Index                   'Cambiar. Añadir índices.
'   Case 0, 1
'      txtDato(Index).SetFocus
'   Case 2, 3
'      mskDato(Index).SetFocus
'   End Select
'   ppAyuBus Index
End Sub

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
'End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
'         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
'      End If
'   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
   Case 0
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMPspGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodCta='" & txtLlave(0).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistro <> -1 Then .Bookmark = dvRegistro
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistro
         End If
      End With
      
      cmdGrabar.Enabled = True
      upHabilitacion True
      pbValidada = True
      txtDato(0).SetFocus
   Else
      cmdGrabar.Enabled = False
      upHabilitacion False
      pbValidada = False
   End If
      
   Exit Sub
Err:
   gpErrores
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
   If Index = 1 Then                   'Cambiar (añadir índices).
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
 ']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
'   On Error GoTo Err

  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 12                             'Cambiar (añadir índices).
'      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
'         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
'      End If
'   End Select

  'Asigna 0 a campos numéricos si están vacíos.
   Select Case Index
   Case 0, 1                             'Cambiar (añadir índices).
      If Not IsNumeric(txtDato(Index).Text) Then
         txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      Else
         txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
      End If
   End Select

  'Busca el dato en su tabla principal.
'   Select Case Index
'   Case 12                             'Cambiar (añadir índices).
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
      
'   Exit Sub
'Err:
'   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "TpoCta=1", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
      lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
   Select Case tnIndex                 'Cambiar.
   Case 0
      If txtLlave(tnIndex).Text = "" Then
         lblLlaveDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMPspGrd.porstCOCta
         .MoveFirst
         .Find "CodCta='" & txtLlave(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblLlaveDeta(tnIndex).Caption = " " & !DetCta
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMPspGrd
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain!CodCta = txtLlave(0).Text
         End If

        'Datos.
'         uorstMain!EstTDc = IIf(chkEstado.Value = vbChecked, ESTTDC_ACT, ESTTDC_INA)
'         uorstMain!CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         uorstMain!FehOpe = dtpFecha.Value
'         uorstMain!CodMon = optMoneda(1).Value
         .uorstMain!OrdRep = IIf(CboTpoGru.ListIndex = TPOGRU1_IND, TPOGRU1_TXT_0, IIf(CboTpoGru.ListIndex = TPOGRU2_IND, TPOGRU2_TXT_0, TPOGRU3_TXT_0)) & gfCeros(txtDato(2).Text, 2, 0, "0")
         .uorstMain.Fields("ImpMN_" & gsMesAct) = CDec(txtDato(0).Text)
         .uorstMain.Fields("ImpME_" & gsMesAct) = CDec(txtDato(1).Text)
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain!CodCta
      
        'Datos.
'         chkEstado.Value = IIf(uorstMain!EstTDc = ESTTDc_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(uorstMain!CodSoc), "", uorstMain!CodSoc)
'         dtpFecha.Value = uorstMain!FehOpe
'         optMoneda(1).Value = uorstMain!CodMon
         txtDato(0).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_" & gsMesAct)), 0, .uorstMain.Fields("ImpMN_" & gsMesAct)), FORMATO_NUM_1)
         txtDato(1).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_" & gsMesAct)), 0, .uorstMain.Fields("ImpME_" & gsMesAct)), FORMATO_NUM_1)
         CboTpoGru.ListIndex = IIf(Left(.uorstMain!OrdRep, 1) = TPOGRU1_TXT_0, TPOGRU1_IND, IIf(Left(.uorstMain!OrdRep, 1) = TPOGRU2_TXT_0, TPOGRU2_IND, TPOGRU3_IND))
         txtDato(2).Text = Right(.uorstMain!OrdRep, 2)
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
   
   Resume
End Sub

Public Sub upDatosPredeterminados()    'Cambiar.
   Dim dnContador As Integer

  'Llaves.
   txtLlave(0).Text = ""

  'Datos.
'   chkEstado.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 2
         .Item(dnContador).Text = Format(0, FORMATO_NUM_1)
      Next
      .Item(dnContador).Text = ""
   End With
   CboTpoGru.ListIndex = TPOGRU1_IND

  'Ayudas.
   lblLlaveDeta(0).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer
  'Llaves
   With txtLlave
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
      Next
   End With
  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   CboTpoGru.Enabled = tbHabilitar

  'Ayudas.
'   cmdDatoAyud(0).Enabled = tbHabilitar
   lblLlaveDeta(0).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
   cmdLlaveAyud(0).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
End Sub

'[Código propio del formulario.

']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
   
   'Orden: Corregir.
   zaOpciones = Array(gbPms02)
End Property

Public Property Get zaOpciones() As Variant
End Property
Public Property Let zaOpciones(ByVal taOpciones As Variant)
   cmdCorregir.Enabled = IIf(pbNuevo, False, taOpciones(0))
End Property

