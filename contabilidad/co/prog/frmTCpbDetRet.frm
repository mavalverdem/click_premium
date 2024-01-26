VERSION 5.00
Begin VB.Form frmTCpbDetRet 
   Caption         =   "[Entidad]"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboRtcPcp 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   877
      ScaleHeight     =   690
      ScaleWidth      =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   780
      Width           =   3000
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
         Left            =   0
         Picture         =   "frmTCpbDetRet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdAceptar 
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
         Left            =   735
         Picture         =   "frmTCpbDetRet.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   1470
         Picture         =   "frmTCpbDetRet.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   2205
         Picture         =   "frmTCpbDetRet.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   6
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
      Index           =   1
      Left            =   3750
      TabIndex        =   2
      Top             =   240
      Width           =   1155
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
      Left            =   3315
      TabIndex        =   1
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Tipo :"
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
      TabIndex        =   9
      Top             =   300
      Width           =   390
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "N° Documento :"
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
      Left            =   2100
      TabIndex        =   7
      Top             =   300
      Width           =   1110
   End
End
Attribute VB_Name = "frmTCpbDetRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean

'[Propio del formulario.
']

Private Sub Form_Load()
   Me.KeyPreview = True
   
   With cboRtcPcp
      .AddItem "Percepción", 0
      .AddItem "Retención", 1
   End With
   txtDato(0).MaxLength = frmTCpbGrd.uorstCOCpbDetRP!SerDoc_RtcPcp.DefinedSize
   txtDato(1).MaxLength = frmTCpbGrd.uorstCOCpbDetRP!NroDoc_RtcPcp.DefinedSize
   
   cmdAceptar.Enabled = pbNuevo
   cmdDeshacer.Enabled = pbNuevo
   cmdCorregir.Enabled = Not pbNuevo
   
   ppHabilitacion pbNuevo
End Sub

Private Sub Form_Activate()
   With frmTCpbGrd.uorstCOCpbDetRP    'Cambiar RecordSet.
      If .RecordCount > 0 Then .MoveFirst
     .Find "cLlave='" & gsMesAct & frmTCpbCab.txtLlave(0).Text & frmTCpbCab.txtLlave(1).Text & IIf(pbNuevo, frmTCpbDet.pnNroIte, frmTCpbGrd.uorstMain_1!NroIte) & "'"
      If Not .EOF Then
         ppDatosDesconectados 1
      Else
         ppDatosPredeterminados
      End If
   End With
   
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

Private Sub Form_Unload(Cancel As Integer)
'   frmTCpbGrd.uorstCOCpbDetRP.Requery
End Sub

Private Sub cmdCorregir_Click()
   cmdCorregir.Enabled = False
   cmdAceptar.Enabled = True
   cmdDeshacer.Enabled = True
   cmdSalir.Enabled = True
   ppHabilitacion True
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
End Sub

Private Sub cmdAceptar_Click()
  'Validacion de Datos.
   If Len(Trim(txtDato(0).Text)) = 0 Or Len(Trim(txtDato(1).Text)) = 0 Then
      MsgBox TEXT_6002, vbExclamation
      If Len(Trim(txtDato(0).Text)) = 0 Then
         txtDato(0).SetFocus
      Else
         txtDato(1).SetFocus
      End If
      frmTCpbDet.pbHayRtcPcp = False
      Exit Sub
   End If
   
   frmTCpbDet.pbHayRtcPcp = True
   
   cmdCorregir.Enabled = True
   cmdAceptar.Enabled = False
   cmdDeshacer.Enabled = False
   ppHabilitacion False
   pbNuevo = False
End Sub

Private Sub cmdDeshacer_Click()
   If pbNuevo Then ppDatosPredeterminados Else ppDatosDesconectados 1

   cmdCorregir.Enabled = True
   cmdAceptar.Enabled = False
   cmdDeshacer.Enabled = False
   ppHabilitacion False
End Sub

Private Sub cmdSalir_Click()
   Me.Hide
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
'   If Index = 0 Then
'   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
   Dim dvRegistroActual As Variant

  'Completa con ceros a la izquierda.
   Select Case Index
   Case 0, 1                    'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

  'Asigna 0 a campos numéricos si están vacíos.

  'Busca el dato en su tabla principal.
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
'   Select Case tnIndex
'   Case 0                              'Cambiar (añadir índices).
'      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
'      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
'   Select Case tnIndex                 'Cambiar.
'   Case 0
'      If txtDato(tnIndex).Text = "" Then
'         lblDatoDeta(tnIndex).Caption = ""
'         Exit Function
'      End If
'      With frmTCprGrd.uorstCOCta
'         .MoveFirst
'         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
'         If .EOF Then
'            MsgBox TEXT_8006, vbExclamation
'            ppAyuDet = True
'         Else
'            lblDatoDeta(tnIndex).Caption = " " & !DetCta
'         End If
'      End With
'   End Select
End Function

Private Sub ppDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmTCpbGrd
      If tnFase = 0 Then
        'Datos.
'         .uorstMain!EstDro = IIf(chkEstado.Value = vbChecked, ESTDRO_ACT, ESTDRO_INA)
'         .uorstMain!CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         .uorstMain!FehOpe = dtpFecha.Value
'         .uorstMain!CodMon = optMoneda(1).Value
'         .uorstCOCpbDetRP!DetDro = txtDato(0).Text
      Else
        'Datos.
         cboRtcPcp.ListIndex = IIf(.uorstCOCpbDetRP!CodTDc = gsCodTDc_Rtc, 1, 0)
'         chkEstado.Value = IIf(.uorstMain!EstDro = ESTDRO_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(.uorstMain!CodSoc), "", .uorstMain!CodSoc)
'         dtpFecha.Value = .uorstMain!FehOpe
'         optMoneda(1).Value = .uorstMain!CodMon
         txtDato(0).Text = .uorstCOCpbDetRP!SerDoc_RtcPcp
         txtDato(1).Text = .uorstCOCpbDetRP!NroDoc_RtcPcp
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
   
   Resume
End Sub

Private Sub ppDatosPredeterminados()    'Cambiar.
   Dim dnContador As Integer

  'Datos.
   cboRtcPcp.ListIndex = 1
'   chkEstado.Value = vbChecked
'   dtpDato(3).Value = Date
'   optTpoMon(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With
   
  'Ayudas.
End Sub

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   cboRtcPcp.Enabled = tbHabilitar
'   chkMonedaActiva.Enabled = tbHabilitar
'   chkDesactivar.Enabled = tbHabilitar
'   dtpDato(3).Enabled = tbHabilitar
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With

  'Ayudas.
End Sub

'[Propio del formulario.

']

Public Property Get zbNuevo() As Boolean
   zbNuevo = pbNuevo
End Property
Public Property Let zbNuevo(ByVal tbNuevo As Boolean)
   pbNuevo = tbNuevo
End Property



