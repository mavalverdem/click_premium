VERSION 5.00
Begin VB.Form frmMOnp 
   Caption         =   "[Entidad]"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frarangos 
      Height          =   4095
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   7335
      Begin VB.CheckBox chkEstAfp 
         Caption         =   "Activo"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   3720
         Width           =   795
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
         Index           =   6
         Left            =   1440
         TabIndex        =   13
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   2880
         Width           =   1410
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
         Index           =   5
         Left            =   1440
         TabIndex        =   12
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   2520
         Width           =   1410
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
         Index           =   4
         Left            =   1440
         TabIndex        =   11
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   2160
         Width           =   1410
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
         Index           =   3
         Left            =   1440
         TabIndex        =   10
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1800
         Width           =   1410
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
         Index           =   2
         Left            =   1440
         TabIndex        =   9
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1440
         Width           =   1410
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
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   915
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
         Left            =   1440
         TabIndex        =   7
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   660
         Width           =   4290
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
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   1035
         Width           =   4290
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   7
         Left            =   6960
         Picture         =   "frmMOnp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3240
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
         Index           =   7
         Left            =   1440
         TabIndex        =   14
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tope Segu.:"
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
         Index           =   7
         Left            =   360
         TabIndex        =   27
         Top             =   2940
         Width           =   870
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Factor 4     :"
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
         Index           =   6
         Left            =   360
         TabIndex        =   26
         Top             =   2580
         Width           =   870
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Factor 3     :"
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
         Left            =   360
         TabIndex        =   25
         Top             =   2220
         Width           =   870
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Factor 2     :"
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
         Left            =   360
         TabIndex        =   24
         Top             =   1860
         Width           =   870
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Factor 1     :"
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
         Left            =   360
         TabIndex        =   23
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Enti Pensión:"
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
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
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
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cód.SUNAT:"
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
         Left            =   360
         TabIndex        =   20
         Top             =   1095
         Width           =   900
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
         Index           =   7
         Left            =   2400
         TabIndex        =   19
         Top             =   3240
         Width           =   4515
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta : "
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
         Index           =   8
         Left            =   360
         TabIndex        =   18
         Top             =   3240
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1935
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4440
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
         Picture         =   "frmMOnp.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmMOnp.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Picture         =   "frmMOnp.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmMOnp.frx":0648
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmMOnp.frx":074A
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
         Picture         =   "frmMOnp.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMOnp"
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
   
   With frmMOnpGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!CodAfp.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
      txtDato(gsIdioma - 1).MaxLength = .uorstMain!Desafp.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain!CodSunat.DefinedSize
    ']
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(7, 2)
  'Contribution
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Enti.Pensión:", "Descripción:", "Cód.SUNAT:", _
    "Aporte:", "Flujo :", "Mixta :", "Seguro :", "Tope Segu.:", "Cuenta:")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Pension.Enti:", "Description:", "Cod.SUNAT:", _
    "Contri.:", "Flow :", "Mixed :", "You sure:", "Top.Insuran.:", "Account:")
  Next nElemento
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']
   chkEstAfp.Caption = Choose(gsIdioma, "Activo", "Active")
   Me.Caption = Choose(gsIdioma, "Registrar ONP/AFP", "register ONP/AFP")
   
End Sub

Private Sub Form_Activate()
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
'   If txtDato(0).Text <> "" Then ppAyuDet 0
 ']

   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMOnpGrd.uorstMain.BOF And frmMOnpGrd.uorstMain.EOF) Then
   frmMOnpGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMOnpGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMOnpGrd.uorstMain, Me 'Cambiar Formulario de Grid.
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
   Dim dvFeCre, dvFeMdf As Variant
   On Error GoTo Err
   With frmMOnpGrd                     'Cambiar Formulario de Grid.
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
         .uorstMain.Find "CodAfp='" & txtLlave(0).Text & "'"
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
  
   frmMOnpGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      'txtDato(2).SetFocus
   End Select
   ppAyuBus Index
End Sub




Private Sub txtDato_LostFocus(Index As Integer)
   Select Case Index
   Case 2, 3, 4, 5, 6
       If CDec(txtDato(Index).Text) <> 0 Then
       txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
       End If
   End Select
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

Private Sub txtllave_GotFocus(Index As Integer)
   txtLlave(Index).SelStart = 0
   txtLlave(Index).SelLength = txtLlave(Index).MaxLength
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
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
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMOnpGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodAfp='" & txtLlave(0).Text & "'"
            If Not .EOF Then
               MsgBox TEXT_8007, vbExclamation
               If dvRegistro <> -1 Then .Bookmark = dvRegistro
               Cancel = True
               Exit Sub
            End If
            .Bookmark = dvRegistro
         End If
      End With
      
'[REVISAR.
'   If Index = 0 Then
'      If Len(txtLlave(0).Text) = 1 Or Len(txtLlave(0).Text) = 3 Then
'         MsgBox Choose(gsIdioma, "El diario debe ser de 2 o 4 caracteres.", "The journal must be  2 or 4 characters."), vbExclamation
'         Cancel = True
'         Exit Sub
'      End If
'      If Len(Trim(txtLlave(0).Text)) = 4 Then
'         With frmMOnpGrd.uorstCoEntidadPen
'            .Requery
'            .Find "CodAfp='" & Mid(txtLlave(0).Text, 1, 2) & "'"
'            If .EOF Then
'               MsgBox Choose(gsIdioma, "El diario ", "The journal ") & Mid(txtLlave(0).Text, 1, 2) & Choose(gsIdioma, " no existe.", " no exist."), vbCritical
'               Cancel = True
'               Exit Sub
'            End If
'         End With
'      End If
'   End If
']

      cmdGrabar.Enabled = True
      upHabilitacion True
      pbValidada = True
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
   If KeyAscii <> 8 Then
      If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
         SendKeys "{TAB}"
      End If
   End If
']ARREGLAR.
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
 Select Case Index    'Busca el dato en su tabla principal.
   Case 7                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 7                          'Cambiar (añadir índices).
'      modAyuBus.Lib_Cod "", txtDato(2).Text, 0, 0, Me.Top + frarangos.Top + txtDato(2).Top + txtDato(2).Height, Me.Left + frarangos.Left + txtDato(2).Left
'      txtDato(2).Text = frmOAyuBus.uvDato1
'      lblDatoDeta(7).Caption = " " & frmOAyuBus.uvDato2
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      
   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
 Select Case tnIndex                 'Cambiar.
   Case 7
      If txtDato(tnIndex).Text = "" Then
         lblDatoDeta(tnIndex).Caption = ""
         Exit Function
      End If
      With frmMOnpGrd.uorstCoCta
         .MoveFirst
         .Find "CodCta='" & txtDato(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMOnpGrd.uorstMain
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !codemp = gsCodEmp
            '.uorstMain!pdoano = gsAnoAct
            !CodAfp = txtLlave(0).Text
         End If

        'Datos.
         !Desafp = txtDato(0).Text
         !CodSunat = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
         !Factor1 = IIf(txtDato(2).Text = "", Null, CDec(txtDato(2).Text))
         !Factor2 = IIf(txtDato(3).Text = "", Null, CDec(txtDato(3).Text))
         !Factor3 = IIf(txtDato(4).Text = "", Null, CDec(txtDato(4).Text))
         !Factor4 = IIf(txtDato(5).Text = "", Null, CDec(txtDato(5).Text))
         !topeseg = IIf(txtDato(6).Text = "", Null, CDec(txtDato(6).Text))
         !CodCta = txtDato(7).Text
         !Estadoafp = IIf(chkEstAfp.Value = vbChecked, ESTCTA_ACT, ESTCTA_INA)
   Else
        'Llaves.
         txtLlave(0).Text = !CodAfp
      
        'Datos.
         txtDato(0).Text = IIf(IsNull(!Desafp), "", !Desafp)
         txtDato(1).Text = IIf(IsNull(!CodSunat), "", !CodSunat)
         txtDato(2).Text = IIf(IsNull(!Factor1), "", Format(!Factor1, FORMATO_NUM_1))
         txtDato(3).Text = IIf(IsNull(!Factor2), "", Format(!Factor2, FORMATO_NUM_1))
         txtDato(4).Text = IIf(IsNull(!Factor3), "", Format(!Factor3, FORMATO_NUM_1))
         txtDato(5).Text = IIf(IsNull(!Factor4), "", Format(!Factor4, FORMATO_NUM_1))
         txtDato(6).Text = IIf(IsNull(!topeseg), "", Format(!topeseg, FORMATO_NUM_1))
         txtDato(7).Text = IIf(IsNull(!CodCta), "", !CodCta)
         chkEstAfp.Value = IIf(!Estadoafp = ESTCTA_ACT, vbChecked, vbUnchecked)
         ppAyuDet 7
         
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
   chkEstAfp.Value = vbChecked

  'Datos.
'   chkEstado.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With

  'Ayudas.
   lblDatoDeta(7).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With

  'Ayudas.
   cmdDatoAyud(7).Enabled = tbHabilitar
   lblDatoDeta(7).Enabled = tbHabilitar
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


