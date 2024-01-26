VERSION 5.00
Begin VB.Form frmTCprMasIgv 
   Caption         =   "[Entidad]"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   5925
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
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   960
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
      Left            =   1665
      TabIndex        =   4
      Top             =   960
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
      Index           =   4
      Left            =   3960
      TabIndex        =   3
      Top             =   600
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
      Left            =   1665
      TabIndex        =   2
      Top             =   600
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
      Index           =   0
      Left            =   1665
      TabIndex        =   0
      Top             =   240
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
      Index           =   3
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2340
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1500
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
         Picture         =   "frmTCprMasIgv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   1815
         Picture         =   "frmTCprMasIgv.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   720
      End
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
      Index           =   8
      Left            =   3600
      TabIndex        =   17
      Top             =   1020
      Width           =   300
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
      Index           =   7
      Left            =   1365
      TabIndex        =   16
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label lblTexto 
      Caption         =   "Op. No Grav.:"
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
      Left            =   60
      TabIndex        =   15
      Top             =   1020
      Width           =   1290
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
      Index           =   5
      Left            =   3600
      TabIndex        =   14
      Top             =   660
      Width           =   300
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
      Index           =   4
      Left            =   1365
      TabIndex        =   13
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblTexto 
      Caption         =   "Op. Gr./No Gr.:"
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
      Left            =   60
      TabIndex        =   12
      Top             =   660
      Width           =   1290
   End
   Begin VB.Label lblTexto 
      Caption         =   "Op. Gravada:"
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
      TabIndex        =   11
      Top             =   300
      Width           =   1290
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
      Index           =   1
      Left            =   1365
      TabIndex        =   10
      Top             =   300
      Width           =   255
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
      Index           =   2
      Left            =   3600
      TabIndex        =   9
      Top             =   300
      Width           =   300
   End
End
Attribute VB_Name = "frmTCprMasIgv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
  Dim nContador  As Integer
  Dim nImporteMN As Double, nImporteME As Double
  '[Datos                            'Cambiar.
  With frmTCpr                     'Cambiar Formulario de Grid.
    For nContador = 0 To 2
     .txtDato(53 + nContador).Text = Format(txtDato(nContador).Text, FORMATO_NUM_1)
     .txtDato(56 + nContador).Text = Format(txtDato(3 + nContador).Text, FORMATO_NUM_1)
     nImporteMN = gfRedond(nImporteMN + CDec(txtDato(nContador).Text), 2)
     nImporteME = gfRedond(nImporteME + CDec(txtDato(3 + nContador).Text), 2)
    Next nContador
    .txtDato(9).Text = Format(nImporteMN, FORMATO_NUM_1)
    .txtDato(20).Text = Format(nImporteME, FORMATO_NUM_1)
  End With
  ']
  Unload Me

End Sub

'[Propio del formulario.
']
Private Sub Form_Activate()
   Dim nContador As Integer
   If frmTCpr.cboTpoMon.ListIndex = TPOMON_EXT_IND Then
     For nContador = 0 To 5
      txtDato(nContador).TabIndex = Choose(nContador + 1, 1, 3, 5, 0, 2, 4)
     Next nContador
     txtDato(3).SetFocus
   End If
End Sub


Private Sub Form_Load()
   
   Dim nContador As Integer
   Me.KeyPreview = True
   '[Datos                            'Cambiar.
   For nContador = 0 To 5
    txtDato(nContador).MaxLength = frmTCpr.txtDato(53 + nContador).MaxLength
    txtDato(nContador).Text = frmTCpr.txtDato(53 + nContador).Text
   Next nContador
   ']
   
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(9, 2)
  
  Me.Caption = Choose(gsIdioma, "Distribución del IGV Compras", "Distribution of the GST Purchases")
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Op. Gravada :", "MN", "ME", "Op. Gr./No Gr. :", "MN", "ME", "Op. No Grav. :", "MN", "ME")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Op. with Taxes :", "NC", "FC", "Op. with/without Taxes :", "NC", "FC", "Op.Without Taxes :", "NC", "FC")
  Next nElemento
  CaptionBotones Me, True, False, False, False, False, False, False, False, False, False, False, False, True, aLabel
  ']
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub txtDato_GotFocus(Index As Integer)
   txtDato.Item(Index).SelStart = 0
   txtDato.Item(Index).SelLength = txtDato.Item(Index).MaxLength
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

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
   
   If Val(txtDato(Index).Text) = 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
   End If
   
   Select Case Index
   Case 0, 1, 2
    If frmTCpr.chkMonedaActiva.Value = vbChecked And frmTCpr.cboTpoMon.ListIndex = TPOMON_NAC_IND Then
       txtDato(Index + 3).Text = Format(Round(CDec(txtDato(Index).Text) / CDec(frmTCpr.txtDato(4).Text), 2), FORMATO_NUM_1)
    End If
   Case 3, 4, 5
    If frmTCpr.chkMonedaActiva.Value = vbChecked And frmTCpr.cboTpoMon.ListIndex = TPOMON_EXT_IND Then
       txtDato(Index - 3).Text = Format(Round(CDec(txtDato(Index).Text) * CDec(frmTCpr.txtDato(4).Text), 2), FORMATO_NUM_1)
    End If
   End Select
   
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  'Asigna 0 a campos numéricos si están vacíos.
  If txtDato(Index).Text = "" Or Not IsNumeric(txtDato(Index).Text) Then
     txtDato(Index).Text = 0
  End If
  txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)

End Sub

