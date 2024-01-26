VERSION 5.00
Begin VB.Form frmMCieCta 
   Appearance      =   0  'Flat
   Caption         =   "[Entidad]"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIndAMo 
      Caption         =   "&Con Ajuste Cta. Mon."
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   3240
      Width           =   1875
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Importe manual"
      ForeColor       =   &H80000002&
      Height          =   765
      Index           =   3
      Left            =   40
      TabIndex        =   27
      Top             =   2760
      Width           =   4515
      Begin VB.TextBox txtImporte 
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
         Index           =   1
         Left            =   2760
         TabIndex        =   8
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txtImporte 
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
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Haber"
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
         Left            =   2280
         TabIndex        =   29
         Top             =   345
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Debe"
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
         Left            =   165
         TabIndex        =   28
         Top             =   345
         Width           =   375
      End
   End
   Begin VB.CheckBox chkIndCCt 
      Caption         =   "&Cuenta Centralizacion"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo Contable"
      ForeColor       =   &H80000002&
      Height          =   525
      Index           =   2
      Left            =   4680
      TabIndex        =   24
      Top             =   840
      Width           =   2835
      Begin VB.OptionButton optTpoCtb 
         Caption         =   "Haber"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   26
         Top             =   200
         Width           =   855
      End
      Begin VB.OptionButton optTpoCtb 
         Caption         =   "Debe"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   200
         Width           =   855
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Fórmula"
      ForeColor       =   &H80000002&
      Height          =   1240
      Index           =   1
      Left            =   4665
      TabIndex        =   23
      Top             =   1440
      Width           =   2835
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
         Height          =   810
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   320
         Width           =   2595
      End
   End
   Begin VB.CheckBox chkIndHTr 
      Alignment       =   1  'Right Justify
      Caption         =   "&Hoja de trabajo"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   53
      TabIndex        =   1
      Top             =   1005
      Width           =   1515
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo de hojas de trabajo"
      ForeColor       =   &H80000002&
      Height          =   1240
      Index           =   0
      Left            =   40
      TabIndex        =   21
      Top             =   1440
      Width           =   4515
      Begin VB.ComboBox cboTpoHT1 
         Height          =   315
         ItemData        =   "frmMCieCta.frx":0000
         Left            =   3060
         List            =   "frmMCieCta.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   285
         Width           =   1335
      End
      Begin VB.OptionButton optTpoHTr 
         Caption         =   "Hoja de trabajo N° 3"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   1935
      End
      Begin VB.OptionButton optTpoHTr 
         Caption         =   "Hoja de trabajo N° 2"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optTpoHTr 
         Caption         =   "Hoja de trabajo N° 1"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo importe :"
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
         Left            =   2040
         TabIndex        =   22
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2010
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3720
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
         Picture         =   "frmMCieCta.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmMCieCta.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmMCieCta.frx":0250
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmMCieCta.frx":0352
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmMCieCta.frx":049C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmMCieCta.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   360
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
      Left            =   713
      TabIndex        =   0
      Top             =   240
      Width           =   950
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   305
      Index           =   0
      Left            =   7193
      Picture         =   "frmMCieCta.frx":07F0
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   240
      Width           =   255
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
      TabIndex        =   19
      Top             =   300
      Width           =   555
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
      TabIndex        =   18
      Top             =   240
      Width           =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   7480
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmMCieCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean
Private pbInd As Boolean


Private Sub chkIndHTr_Click()
   If pbInd Then
      If chkIndHTr.Value = vbChecked Then
         optTpoHTr(0).Enabled = True
         optTpoHTr(1).Enabled = True
         optTpoHTr(2).Enabled = True
         If optTpoHTr(0).Value = True Then cboTpoHT1.Enabled = True
         txtDato(0).Enabled = False
         txtImporte(0).Enabled = False
         txtImporte(1).Enabled = False
         chkIndCCt.Enabled = False
         chkIndAMo.Enabled = False
      Else
         optTpoHTr(0).Enabled = False
         optTpoHTr(1).Enabled = False
         optTpoHTr(2).Enabled = False
         cboTpoHT1.Enabled = False
         txtDato(0).Enabled = True
         txtImporte(0).Enabled = True
         txtImporte(1).Enabled = True
         chkIndCCt.Enabled = True
         chkIndAMo.Enabled = True
      End If
   End If
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtLlave(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

'[Propio del formulario.
'Private porstTGTPv As ADODB.Recordset
']

Private Sub Form_Load()
   pbValidada = False
   pbInd = False

   Me.KeyPreview = True
   
   With frmMCieGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain_1!CodCta.DefinedSize
    ']
   
    '[Datos.                           'Cambiar.
      With cboTpoHT1
         .AddItem TPOHT1_SAL_TXT, 0
         .AddItem TPOHT1_DEP_TXT, 1
      End With
      txtImporte(0).MaxLength = 14
      txtImporte(1).MaxLength = 14
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
   If txtLlave(0).Text <> "" Then ppAyuDet 0
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not frmMCieGrd.uorstMain_1.EOF Then
      If frmMCieGrd.uorstMain_1.EditMode <> adEditNone Then frmMCieGrd.uorstMain_1.CancelUpdate 'Cambiar Formulario de Grid.
   End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMCieGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMCieGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   pbInd = True
   upHabilitacion (True)
 '[Dato con el foco al corregir.       'Cambiar.
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err
   
   If Not ResuelveFormula(Trim(txtDato(0).Text), chkIndAMo.Value) Then
      With frmMCieGrd                     'Cambiar Formulario de Grid.
         .uocnnMain.BeginTrans            'INICIA TRANSACCION.
         If pbNuevo Then
            .uorstMain_1.AddNew
         End If
   
         upDatosDesconectados 0
   
         With .uorstMain_1
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
            .uorstMain_1.Requery
            .upDatosGrid 1
          '[Búsqueda de llave actual.     'Cambiar.
            .uorstMain_1.Find "CodCta='" & txtLlave(0).Text & "'"
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
   End If
      
   Exit Sub
Err:
   gpErrores
  
   frmMCieGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   pbInd = False
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub optTpoHTr_Click(Index As Integer)
   If optTpoHTr(0).Value = True And pbInd Then
      cboTpoHT1.Enabled = True
   Else
      cboTpoHT1.Enabled = False
   End If
End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
   txtImporte.Item(Index).SelStart = 0
   txtImporte.Item(Index).SelLength = txtImporte.Item(Index).MaxLength
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
   If Val(txtImporte(Index).Text) = 0 Then
      txtImporte(Index).Text = Format(0, FORMATO_NUM_1)
   End If
   
   Select Case Index
   Case 0
      If CDec(txtImporte(Index).Text) <> 0 Then
         txtImporte(1).Text = Format(0, FORMATO_NUM_1)
      End If
   Case 1
      If CDec(txtImporte(Index).Text) <> 0 Then
         txtImporte(0).Text = Format(0, FORMATO_NUM_1)
      End If
   End Select
   
   If CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0 And Trim(txtDato(0).Text) = "" Then
      chkIndCCt.Enabled = True
      chkIndAMo.Enabled = True
   Else
      chkIndAMo.Enabled = False
      chkIndCCt.Enabled = False
      txtDato(0).Text = ""
   End If
   
   txtImporte(Index).Text = Format(CDec(txtImporte(Index).Text), FORMATO_NUM_1)
End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
 '[Convierte a mayúsculas.
'   If Index = 1 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then cmdLlaveAyud(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
   Select Case Index
   Case 0
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
'      With frmMCieGrd.uorstMain_1
'         If Not (.BOF And .EOF) Then
'            dvRegistro = .Bookmark
'            .MoveFirst
'            .Find "CodCta='" & txtLlave(0).Text & "'"
'            If Not .EOF Then
'               MsgBox TEXT_8007, vbExclamation
'               If dvRegistro <> -1 Then .Bookmark = dvRegistro
'               Cancel = True
'               Exit Sub
'            End If
'            .Bookmark = dvRegistro
'         End If
'      End With
      
      cmdGrabar.Enabled = True
      upHabilitacion True
      pbValidada = True
      pbInd = True
   Else
      cmdGrabar.Enabled = False
      upHabilitacion False
      pbValidada = False
      pbInd = False
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
'   If Index = 1 Or Index = 2 Then      'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
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
   Select Case Index
   Case 0                             'Cambiar (añadir índices).
      If Trim(txtDato(0).Text) = "" And CDec(txtImporte(0).Text) = 0 And CDec(txtImporte(1).Text) = 0 Then
         chkIndCCt.Enabled = True
      Else
         txtImporte(0).Text = Format(0, FORMATO_NUM_1)
         txtImporte(1).Text = Format(0, FORMATO_NUM_1)
         chkIndCCt.Enabled = False
         chkIndAMo.Enabled = True
      End If
   End Select

  'Asigna 0 a campos numéricos si están vacíos.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
'      If Not IsNumeric(txtDato(Index).Text) Then
'         txtDato(Index).Text = 0
'      End If
'   End Select

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
      modAyuBus.Cta_Cod "TpoCta=1", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
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
      With frmMCieGrd.uorstCOCta
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

   With frmMCieGrd.uorstMain_1                     'Cambiar Formulario de Grid.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !NroCie = frmMCieGrd.uorstMain_0!NroCie
            With frmMCieGrd.uorstUltiItem
               .Source = "SELECT IFNULL(MAX(NroIte), 0) AS cUltIte " _
                       & "FROM COCieCta " _
                       & "WHERE NroCie='" & frmMCieGrd.uorstMain_0!NroCie & "'"
               .Open
               frmMCieGrd.uorstMain_1!NroIte = IIf(IsNull(!cUltIte), 1, !cUltIte + 1)
               .Close
            End With
         End If

        'Datos.
            !CodCta = Trim(txtLlave(0).Text)
            !IndHTr = IIf(chkIndHTr.Value = vbChecked, INDHTR_ACT, INDHTR_INA)
            If chkIndHTr.Value = vbChecked Then
               txtImporte(0).Text = Format(0, FORMATO_NUM_1)
               txtImporte(1).Text = Format(0, FORMATO_NUM_1)
               txtDato(0).Text = ""
            End If
            !IndCCt = IIf(chkIndCCt.Value = vbChecked, INDCCT_ACT, INDCCT_INA)
            !IndAMo = IIf(chkIndAMo.Value = vbChecked, INDAMO_ACT, INDAMO_INA)
            If chkIndHTr.Value = vbUnchecked Then
               optTpoHTr(0).Value = True
               cboTpoHT1.ListIndex = 0
            End If
            If Not optTpoHTr(0).Value Then
               cboTpoHT1.ListIndex = 0
            End If
            !TpoHTr = Switch(optTpoHTr(0).Value, TPOHTR_HT1, optTpoHTr(1).Value, TPOHTR_HT2, optTpoHTr(2).Value, TPOHTR_HT3)
            !TpoHT1 = Choose(cboTpoHT1.ListIndex + 1, TPOHT1_SAL, TPOHT1_DEP)
            !FmlCie = txtDato(0).Text
            !TpoCtb = Switch(optTpoCtb(0).Value, 0, optTpoCtb(1).Value, 1)
            !ImpMNI = CDec(IIf(txtImporte(0).Text <> 0, txtImporte(0).Text, txtImporte(1).Text))
            !TpoCtbI = IIf(txtImporte(0).Text = 0, 1, 0)
      Else
        'Llaves.
      
        'Datos.
         txtLlave(0).Text = !CodCta
         chkIndHTr.Value = IIf(!IndHTr = INDHTR_ACT, vbChecked, vbUnchecked)
         chkIndCCt.Value = IIf(!IndCCt = INDCCT_ACT, vbChecked, vbUnchecked)
         chkIndAMo.Value = IIf(!IndAMo = INDAMO_ACT, vbChecked, vbUnchecked)
         optTpoHTr(!TpoHTr).Value = True
         optTpoCtb(!TpoCtb).Value = True
         Select Case !TpoHT1
         Case TPOHT1_SAL
            cboTpoHT1.ListIndex = 0
         Case TPOHT1_DEP
            cboTpoHT1.ListIndex = 1
         End Select
         txtDato(0).Text = IIf(IsNull(!FmlCie), "", !FmlCie)
         If !TpoCtbI = 0 Then
            txtImporte(0).Text = Format(IIf(IsNull(!ImpMNI), 0, !ImpMNI), FORMATO_NUM_1)
            txtImporte(1).Text = Format(0, FORMATO_NUM_1)
         Else
            txtImporte(0).Text = Format(0, FORMATO_NUM_1)
            txtImporte(1).Text = Format(IIf(IsNull(!ImpMNI), 0, !ImpMNI), FORMATO_NUM_1)
         End If
         If IsNull(!FmlCie) Or Trim(!FmlCie) = "" Or !ImpMNI > 0 Then
            chkIndCCt.Enabled = False
         End If
         If !ImpMNI > 0 Or chkIndHTr.Value = vbChecked Then
            chkIndAMo.Enabled = False
         End If
        'Busca detalle de códigos.
         'cmdCorregir.Enabled = False
'         ppAyuDet 0
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
   chkIndHTr.Value = vbUnchecked
   chkIndCCt.Value = vbUnchecked
   chkIndAMo.Value = vbUnchecked
''   dcoSocio.BoundText = ""
''   dtpFecha.Value = Date
   optTpoHTr(0).Value = True
   optTpoCtb(0).Value = True
   cboTpoHT1.ListIndex = 0
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With
   With txtImporte
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = Format(0, FORMATO_NUM_1)
      Next
   End With

  'Ayudas.
'   lblDatoDeta(2).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtLlave
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
      Next
   End With
   With txtImporte
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   chkIndHTr.Enabled = tbHabilitar
   With optTpoHTr
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = IIf(pbInd And chkIndHTr.Value = vbChecked, tbHabilitar, False)
      Next
   End With
   With optTpoCtb
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   cboTpoHT1.Enabled = IIf(pbInd And chkIndHTr.Value = vbChecked And optTpoHTr(0).Value = True, tbHabilitar, False)
   txtDato(0).Enabled = IIf(pbInd And chkIndHTr.Value = vbUnchecked, tbHabilitar, False)
   txtImporte(0).Enabled = IIf(pbInd And chkIndHTr.Value = vbUnchecked, tbHabilitar, False)
   txtImporte(1).Enabled = IIf(pbInd And chkIndHTr.Value = vbUnchecked, tbHabilitar, False)
   chkIndCCt.Enabled = IIf(pbInd And chkIndHTr.Value = vbUnchecked, tbHabilitar, False)
   chkIndAMo.Enabled = IIf(pbInd And chkIndHTr.Value = vbUnchecked, tbHabilitar, False)
  'Ayudas.
  'Ayudas.
'   cmdDatoAyud(0).Enabled = tbHabilitar
   lblLlaveDeta(0).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
   cmdLlaveAyud(0).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
End Sub

'[Código propio del formulario.

Private Function ResuelveFormula(ByVal s_Cadena As String, nIndAMo As Integer) As Boolean
   Static sVariable As String, sCaso As String, sSigno As String
   Static nInicio As Integer, nFinal As Integer, nLen As Integer, nContador As Integer

   nInicio = 1: nFinal = 1: nLen = 0: nContador = 0
   sCaso = Left(s_Cadena, 1)

   If nIndAMo = 1 Then
      Do While nContador <= Len(s_Cadena)
        Select Case sCaso
          Case "["         ' Cuenta
            nInicio = (InStr(nInicio, s_Cadena, "[", vbTextCompare)) + 1
            nFinal = InStr(nInicio, s_Cadena, "]", vbTextCompare)
            nLen = (nFinal - nInicio)
            sVariable = Mid$(s_Cadena, nInicio, nLen)
            sSigno = Left(sVariable, 1)
            sVariable = IIf(IsNumeric(sSigno), sVariable, Mid(sVariable, 2))
            With frmMCieGrd.uorstCOCta
               .MoveFirst
               .Find "CodCta='" & sVariable & "'"
               If Not .EOF Then
                  If !IndMoe = 0 Then
                     If MsgBox("La Cuenta [" & sVariable & "], no es Cuenta Monetaria desea grabar.", vbExclamation + vbYesNo) = vbYes Then
                        ResuelveFormula = False
                     Else
                        ResuelveFormula = True
                        Exit Do
                     End If
                  End If
               End If
            End With
            sCaso = Mid(s_Cadena, nFinal + 1, 1)
            nContador = nFinal
          Case "+"         ' Signo Positivo
            sCaso = Mid(s_Cadena, nFinal + 2, 1)
            nContador = nFinal + 1
          Case "-"         ' Signo Negativo
            sCaso = Mid(s_Cadena, nFinal + 2, 1)
            nContador = nFinal + 1
          Case Else        ' Otro Caso
            nContador = nContador + 1
        End Select
      Loop
   End If
End Function

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


