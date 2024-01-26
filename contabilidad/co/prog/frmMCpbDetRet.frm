VERSION 5.00
Begin VB.Form frmMCpbDetRet 
   Caption         =   "[Entidad Tipo Asiento]"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3255
      TabIndex        =   7
      Top             =   240
      Width           =   435
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
      Left            =   3690
      TabIndex        =   6
      Top             =   240
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   817
      ScaleHeight     =   690
      ScaleWidth      =   3000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1140
      Width           =   3000
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
         Picture         =   "frmMCpbDetRet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   1470
         Picture         =   "frmMCpbDetRet.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.ComboBox cboRtcPcp 
      Height          =   315
      Left            =   420
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1395
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "##.00"
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
      Left            =   1140
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "###,###,###.00"
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
      Index           =   3
      Left            =   3270
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblTexto 
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
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   300
      Width           =   1110
   End
   Begin VB.Label lblTexto 
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
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   300
      Width           =   390
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Base Im"
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
      Left            =   0
      TabIndex        =   10
      Top             =   780
      Width           =   570
   End
   Begin VB.Label lblTexto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " M.N."
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
      Left            =   705
      TabIndex        =   9
      Top             =   780
      Width           =   360
   End
   Begin VB.Label lblTexto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " M.E."
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
      Left            =   2850
      TabIndex        =   8
      Top             =   780
      Width           =   330
   End
End
Attribute VB_Name = "frmMCpbDetRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboRtcPcp_LostFocus()
  '[ Habilita los importes de acuerdo(agente)
  If cboRtcPcp.ListIndex = 1 Then
    txtDato(2).Enabled = (gsIndPcp = "N")
    txtDato(3).Enabled = (gsIndPcp = "N")
  ElseIf cboRtcPcp.ListIndex = 2 Then
    txtDato(2).Enabled = (gsIndRtc = "N")
    txtDato(3).Enabled = (gsIndRtc = "N")
  Else
    txtDato(2).Enabled = False: txtDato(3).Enabled = False
  End If
  ']
End Sub

'[Propio del formulario.
']

Private Sub Form_Load()
   
   Me.KeyPreview = True
   Me.Caption = Choose(gsIdioma, "Documento de Retención o Percepción", "Document of Winhholding or Perception")
   
   With cboRtcPcp
      .AddItem Choose(gsIdioma, "Ninguno", "Neither"), 0
      .AddItem Choose(gsIdioma, "Percepción", "Perception"), 1
      .AddItem Choose(gsIdioma, "Retención", "Withholding"), 2
   End With
   txtDato(0).MaxLength = frmMCpbDet.txtDato(4).MaxLength
   txtDato(1).MaxLength = frmMCpbDet.txtDato(5).MaxLength
   txtDato(2).MaxLength = 14
   txtDato(3).MaxLength = 14
   '[ Datos
   cboRtcPcp.ListIndex = IIf(frmMCpbDet.psTpoDocRP = gsCodTDc_Pcp, 1, IIf(frmMCpbDet.psTpoDocRP = gsCodTDc_Rtc, 2, 0))
   txtDato(0).Text = frmMCpbDet.psSerDocRP
   txtDato(1).Text = frmMCpbDet.psNroDocRP
   txtDato(2).Text = Format(IIf(IsNull(frmMCpbDet.pnImpDcMNRP), 0, frmMCpbDet.pnImpDcMNRP), FORMATO_NUM_1)
   txtDato(3).Text = Format(IIf(IsNull(frmMCpbDet.pnImpDcMERP), 0, frmMCpbDet.pnImpDcMERP), FORMATO_NUM_1)
   txtDato(2).Tag = Format(txtDato(2).Text, FORMATO_NUM_1)
   txtDato(3).Tag = Format(txtDato(3).Text, FORMATO_NUM_1)
   txtDato(2).Enabled = False: txtDato(3).Enabled = False
   ']
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(5, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Tipo:", "Nº Documento:", "Base Imp", "M.N.:", "M.E.:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Type:", "Nº Documet:", "Base Amou", "N.C.:", "F.C.:")
  Next nElemento
  CaptionBotones Me, True, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
 ']
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub cmdAceptar_Click()
  'Validacion de Datos.
   If (cboRtcPcp.ListIndex = 0 And (Len(Trim(txtDato(0).Text)) <> 0 Or Len(Trim(txtDato(1).Text)) <> 0)) Then Beep: MsgBox Choose(gsIdioma, "Seleccione Tipo de Documento", "You Select Type Document"), vbExclamation: cboRtcPcp.SetFocus: frmMCpbDet.pbHayRtcPcp = False: Exit Sub
   If (cboRtcPcp.ListIndex = 1 And gsIndPcp = "N" And CDec(txtDato(2).Text) = 0) Then Beep: MsgBox Choose(gsIdioma, "Ingrese Importe Base de Impuesto", "You enter Base amount of Tax"), vbExclamation: txtDato(2).SetFocus: frmMCpbDet.pbHayRtcPcp = False: Exit Sub
   If (cboRtcPcp.ListIndex = 2 And gsIndRtc = "N" And CDec(txtDato(2).Text) = 0) Then Beep: MsgBox Choose(gsIdioma, "Ingrese Importe Base de Impuesto", "You enter Base amount of Tax"), vbExclamation: txtDato(2).SetFocus: frmMCpbDet.pbHayRtcPcp = False: Exit Sub
   If (cboRtcPcp.ListIndex <> 0 And (Len(Trim(txtDato(0).Text)) = 0 Or Len(Trim(txtDato(1).Text)) = 0)) Then
      MsgBox TEXT_6002, vbExclamation
      If Len(Trim(txtDato(0).Text)) = 0 Then
         txtDato(0).SetFocus
      Else
         txtDato(1).SetFocus
      End If
      frmMCpbDet.pbHayRtcPcp = False
      Exit Sub
   End If
   frmMCpbDet.psTpoDocRP = Choose(cboRtcPcp.ListIndex + 1, "", gsCodTDc_Pcp, gsCodTDc_Rtc)
   frmMCpbDet.psSerDocRP = txtDato(0).Text
   frmMCpbDet.psNroDocRP = txtDato(1).Text
   frmMCpbDet.pnImpDcMNRP = Format(txtDato(2).Text, FORMATO_NUM_1)
   frmMCpbDet.pnImpDcMERP = Format(txtDato(3).Text, FORMATO_NUM_1)
   frmMCpbDet.pbHayRtcPcp = (cboRtcPcp.ListIndex <> 0)
   Unload Me
   
End Sub

Private Sub cmdSalir_Click()
   Unload Me
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

Private Sub txtDato_LostFocus(Index As Integer)

  Select Case Index
   Case 2, 3
    If Val(txtDato(Index).Text) = 0 Then
        txtDato(Index).Text = Format(0, FORMATO_NUM_1)
    End If
    If Index = 2 Then
      If CDec(txtDato(Index).Text) <> 0 Then
        If frmMCpbDet.cboTpoMon.ListIndex = TPOMON_NAC_IND And (CDec(txtDato(3).Text) = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
          txtDato(3).Text = Format(gfRedond(CDec(txtDato(Index).Text) / CDec(frmMCpbDet.txtDato(8).Text), 2), FORMATO_NUM_1)
        End If
      End If
    Else
      If CDec(txtDato(Index).Text) <> 0 Then
        If frmMCpbDet.cboTpoMon.ListIndex = TPOMON_EXT_IND And (CDec(txtDato(2).Text) = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
          txtDato(2).Text = Format(gfRedond(CDec(txtDato(Index).Text) * CDec(frmMCpbDet.txtDato(8).Text), 2), FORMATO_NUM_1)
        End If
      End If
    End If
  End Select

End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
  
  'Completa con ceros a la izquierda.
   Select Case Index
   Case 0, 1                    'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

   Exit Sub
Err:
   gpErrores
End Sub

'[Propio del formulario.

']

