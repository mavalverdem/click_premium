VERSION 5.00
Begin VB.Form frmMEFiLin 
   Caption         =   "[Entidad]"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   6555
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
      Index           =   1
      Left            =   1020
      TabIndex        =   6
      Top             =   975
      Width           =   5355
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
      Index           =   1
      Left            =   1665
      TabIndex        =   2
      Top             =   90
      Width           =   435
   End
   Begin VB.ComboBox cboEstilo 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      ItemData        =   "frmMEFiLin.frx":0000
      Left            =   1020
      List            =   "frmMEFiLin.frx":0002
      TabIndex        =   8
      Top             =   1350
      Width           =   1500
   End
   Begin VB.CheckBox chkEfecto 
      Caption         =   "Font Subrayado"
      Height          =   195
      Left            =   3165
      TabIndex        =   9
      Top             =   1395
      Width           =   1935
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Formato CONASEV"
      ForeColor       =   &H80000002&
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   25
      Top             =   4470
      Width           =   2715
      Begin VB.OptionButton optIndLat 
         Caption         =   "Izquierda"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optIndLat 
         Caption         =   "Derecha"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
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
      Left            =   600
      TabIndex        =   29
      Top             =   5160
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
      Left            =   1020
      TabIndex        =   4
      Top             =   600
      Width           =   5355
   End
   Begin VB.CheckBox chkBsePct 
      Caption         =   "Base para el cálculo de porcentajes"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   5220
      Width           =   5205
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
      ForeColor       =   &H80000002&
      Height          =   1740
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1740
      Width           =   6255
      Begin VB.ComboBox cboEstilo 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         ItemData        =   "frmMEFiLin.frx":0004
         Left            =   4590
         List            =   "frmMEFiLin.frx":0006
         TabIndex        =   17
         Top             =   210
         Width           =   1500
      End
      Begin VB.Frame fraLinea 
         Caption         =   " Bordes "
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   2400
         TabIndex        =   18
         Top             =   1035
         Width           =   3735
         Begin VB.ComboBox cboBorde 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            ItemData        =   "frmMEFiLin.frx":0008
            Left            =   2640
            List            =   "frmMEFiLin.frx":000A
            TabIndex        =   22
            Top             =   200
            Width           =   1000
         End
         Begin VB.ComboBox cboBorde 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            ItemData        =   "frmMEFiLin.frx":000C
            Left            =   840
            List            =   "frmMEFiLin.frx":000E
            TabIndex        =   20
            Top             =   200
            Width           =   1000
         End
         Begin VB.Label lblTexto 
            Caption         =   "Inferior :"
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   6
            Left            =   1920
            TabIndex        =   21
            Top             =   250
            Width           =   735
         End
         Begin VB.Label lblTexto 
            Caption         =   "Superior :"
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.OptionButton optTpoLin 
         Caption         =   "Mascara"
         ForeColor       =   &H8000000D&
         Height          =   200
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1410
         Width           =   2775
      End
      Begin VB.OptionButton optTpoLin 
         Caption         =   "Título"
         ForeColor       =   &H8000000D&
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton optTpoLin 
         Caption         =   "Subtotal (separa con una línea simple arriba y una simple abajo)"
         ForeColor       =   &H8000000D&
         Height          =   200
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   570
         Width           =   5775
      End
      Begin VB.OptionButton optTpoLin 
         Caption         =   "Total (separa con una línea simple arriba y  una línea doble abajo)"
         ForeColor       =   &H8000000D&
         Height          =   200
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   5175
      End
      Begin VB.OptionButton optTpoLin 
         Caption         =   "Sólo realiza operaciones"
         ForeColor       =   &H8000000D&
         Height          =   200
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1110
         Width           =   2775
      End
      Begin VB.Label lblTexto 
         Alignment       =   1  'Right Justify
         Caption         =   "Estilo Font :"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   4
         Left            =   3630
         TabIndex        =   16
         Top             =   255
         Width           =   900
      End
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
      Height          =   795
      Index           =   2
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   3570
      Width           =   5655
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
      Left            =   1020
      TabIndex        =   1
      Top             =   90
      Width           =   435
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1462
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5550
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
         Picture         =   "frmMEFiLin.frx":0010
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Picture         =   "frmMEFiLin.frx":015A
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "frmMEFiLin.frx":025C
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "frmMEFiLin.frx":035E
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "frmMEFiLin.frx":04A8
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "frmMEFiLin.frx":0652
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Traducción:"
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
      TabIndex        =   5
      Top             =   1035
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1545
      X2              =   1590
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Label lblTexto 
      Alignment       =   1  'Right Justify
      Caption         =   "Estilo Font :"
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   1395
      Width           =   900
   End
   Begin VB.Label lblTexto 
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
      Index           =   8
      Left            =   60
      TabIndex        =   28
      Top             =   5220
      Width           =   495
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
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fórmula:"
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
      Left            =   60
      TabIndex        =   23
      Top             =   3465
      Width           =   615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Línea:"
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
      TabIndex        =   0
      Top             =   150
      Width           =   435
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   6360
      Y1              =   495
      Y2              =   495
   End
End
Attribute VB_Name = "frmMEFiLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean

'[Propio del formulario.
'Private porstTGTPv As ADODB.Recordset
']

Private Sub Form_Load()
   Dim n_Index As Integer
   
   pbValidada = False

   Me.KeyPreview = True
   With frmMEFiGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = (.uorstMain_1!NroLin.DefinedSize - 1)
      txtLlave(1).MaxLength = 1
    ']
   
    '[Datos.                           'Cambiar.
      txtDato(gsIdioma - 1).MaxLength = .uorstMain_1!DetLin.DefinedSize
      txtDato(2 - gsIdioma).MaxLength = .uorstMain_1!DetLinx.DefinedSize
      txtDato(2).MaxLength = .uorstMain_1!FmlLin.DefinedSize
      txtDato(3).MaxLength = .uorstMain_1!grppct.DefinedSize
    ']
   End With
   For n_Index = 0 To 2
    If gsIdioma = NvlUsr_Sup Then
      cboBorde(0).AddItem Choose(n_Index + 1, "Ninguno", "Simple", "Doble")
      cboBorde(1).AddItem Choose(n_Index + 1, "Ninguno", "Simple", "Doble")
    Else
      cboBorde(0).AddItem Choose(n_Index + 1, "Neither", "Single", "Double")
      cboBorde(1).AddItem Choose(n_Index + 1, "Neither", "Single", "Double")
    End If
   Next n_Index
   cboBorde(0).ListIndex = 0
   cboBorde(1).ListIndex = 0
   
   For n_Index = 0 To 3
    If gsIdioma = NvlUsr_Sup Then
     cboEstilo(0).AddItem Choose(n_Index + 1, "Normal", "Cursiva", "Negrita", "Negrita Cursiva")
     cboEstilo(1).AddItem Choose(n_Index + 1, "Normal", "Cursiva", "Negrita", "Negrita Cursiva")
    Else
     cboEstilo(0).AddItem Choose(n_Index + 1, "Normal", "Italic", "Bold", "Bold Italic")
     cboEstilo(1).AddItem Choose(n_Index + 1, "Normal", "Italic", "Bold", "Bold Italic")
    End If
   Next n_Index
   cboEstilo(0).ListIndex = 0
   cboEstilo(1).ListIndex = 0
   chkEfecto.Value = Unchecked
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(9, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Línea :", "Descripción :", "Traducción :", "Estilo Font :", "Estilo Font :", "Superior :", "Inferior :", "Fórmula :", "Grupo :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Line :", "Description:", "Translation:", "Font Style :", "Font Style :", "Top :", "Bottom :", "Formula :", "Group :")
  Next nElemento
  chkEfecto.Caption = Choose(gsIdioma, "Font Subrayado", "Underlined Font")
  fraTipo(0).Caption = Choose(gsIdioma, " Tipo ", " Type ")
  optTpoLin(0).Caption = Choose(gsIdioma, "Título", "Title")
  optTpoLin(1).Caption = Choose(gsIdioma, "Subtotal (separa con una línea simple arriba y una simple abajo)", "Subtotal(Separate with one line simple up and one line simple down)")
  optTpoLin(2).Caption = Choose(gsIdioma, "Total (separa con una línea simple arriba y  una línea doble abajo)", "Total(Separate with one line simple up and one line double down)")
  optTpoLin(3).Caption = Choose(gsIdioma, "Sólo realiza operaciones", "Only It makes Operations")
  optTpoLin(4).Caption = Choose(gsIdioma, "Mascara", "Mask")
  fraLinea.Caption = Choose(gsIdioma, " Bordes ", " Edges ")
  fraTipo(1).Caption = Choose(gsIdioma, "Formato CONASEV", "Format CONASEV")
  optIndLat(0).Caption = Choose(gsIdioma, "Derecha", "Right")
  optIndLat(1).Caption = Choose(gsIdioma, "Izquierda", "Left")
  chkBsePct.Caption = Choose(gsIdioma, "Base para el cálculo de porcentajes", "Base for the calculate of percentages")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
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

Private Sub Form_Unload(Cancel As Integer)
  If Not (frmMEFiGrd.uorstMain_1.BOF And frmMEFiGrd.uorstMain_1.EOF) Then
   frmMEFiGrd.uorstMain_1.CancelUpdate 'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMEFiGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMEFiGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
   If frmMEFiGrd.uorstMain_0!IndCnv Then
      optIndLat.Item(0).Enabled = True
      optIndLat.Item(1).Enabled = True
   End If
 
 '[Dato con el foco al corregir.       'Cambiar.
   txtDato(0).SetFocus
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err
   
   With frmMEFiGrd                     'Cambiar Formulario de Grid.
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
         .uorstMain_1.Find "NroLin='" & txtLlave(0).Text & txtLlave(1).Text & "'"
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
  
   frmMEFiGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
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

Private Sub optTpoLin_Click(Index As Integer)
   fraLinea.Enabled = (Index = 3 And optTpoLin(3).Enabled)
   If Index <> 3 Then
    cboBorde(0).ListIndex = 0
    cboBorde(1).ListIndex = 0
   End If
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

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
 '[Convierte a mayúsculas.
'   If Index = 1 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
   If pbValidada Then txtDato(0).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
   Case 0
      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength - 1 Then
         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
      End If
   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      Cancel = ppAyuDet(Index)
'      If Cancel Then Exit Sub
'   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 And Index = 1 Then
      With frmMEFiGrd.uorstMain_1        'Cambiar Formulario de Grid.
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "NroLin='" & txtLlave(0).Text & txtLlave(1).Text & "'"
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
      If frmMEFiGrd.uorstMain_0!IndCnv Then
         optIndLat.Item(0).Enabled = True
         optIndLat.Item(1).Enabled = True
      End If
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
   If Index = 2 Or Index = 3 Then      'Cambiar (añadir índices).
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
'   Select Case tnIndex
'   Case 2                              'Cambiar (añadir índices).
'      modAyuBus.Dtt_Cod txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
'      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
'      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'   End Select
End Sub

Private Function ppAyuDet(tnIndex As Integer)
'   Select Case tnIndex                 'Cambiar.
'   Case 2
'      If txtDato(tnIndex).Text = "" Then
'         lblDatoDeta(tnIndex).Caption = ""
'         Exit Function
'      End If
'      With porstTGDtt
'         .MoveFirst
'         .Find "CodDtt='" & txtDato(tnIndex).Text & "'"
'         If .EOF Then
'            MsgBox TEXT_8006, vbExclamation
'            ppAyuDet = True
'         Else
'            lblDatoDeta(tnIndex).Caption = " " & !DetDtt
'         End If
'      End With
'   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   With frmMEFiGrd                     'Cambiar Formulario de Grid.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain_1!codemp = frmMEFiGrd.uorstMain_0!codemp
            .uorstMain_1!pdoano = frmMEFiGrd.uorstMain_0!pdoano
            .uorstMain_1!CodEfi = frmMEFiGrd.uorstMain_0!CodEfi
            .uorstMain_1!NroLin = txtLlave(0).Text & Trim(txtLlave(1).Text)
         End If

        'Datos.
         .uorstMain_1!BsePct = IIf(chkBsePct.Value = vbChecked, BSEPCT_ACT, BSEPCT_INA)
'         uorstMain_1!CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
'         uorstMain_1!FehOpe = dtpFecha.Value
         '[arturo'
         'se cambio los indices de radio button van de 0 a 3
         .uorstMain_1!TpoLin = Switch(optTpoLin(0).Value, TPOLIN_TIT, optTpoLin(1).Value, TPOLIN_STO, optTpoLin(2).Value, TPOLIN_TOT, optTpoLin(3).Value, TPOLIN_OPE, optTpoLin(4).Value, TPOLIN_MAS)
         .uorstMain_1!IndLat = Switch(optIndLat(0).Value, INDLAT_DER, optIndLat(1).Value, INDLAT_IZQ)
         .uorstMain_1!DetLin = txtDato(gsIdioma - 1).Text
         .uorstMain_1!DetLinx = txtDato(2 - gsIdioma).Text
         .uorstMain_1!FmlLin = txtDato(2).Text
         .uorstMain_1!grppct = txtDato(3).Text
         .uorstMain_1!IndBdeSup = cboBorde(0).ListIndex
         .uorstMain_1!IndBdeInf = cboBorde(1).ListIndex
         
         .uorstMain_1!IndFonDet = cboEstilo(0).ListIndex
         .uorstMain_1!IndFonDet_Syd = IIf(chkEfecto.Value = Checked, 1, 0)
         .uorstMain_1!IndFonImp = cboEstilo(1).ListIndex
      Else
        'Llaves.
         txtLlave(0).Text = Left(.uorstMain_1!NroLin, 3)
         txtLlave(1).Text = Mid(.uorstMain_1!NroLin, 4, 1)
      
        'Datos.
         chkBsePct.Value = IIf(.uorstMain_1!BsePct = BSEPCT_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(uorstMain_1!CodSoc), "", uorstMain_1!CodSoc)
'         dtpFecha.Value = uorstMain_1!FehOpe
'         optMoneda(1).Value = uorstMain_1!CodMon
         optTpoLin(.uorstMain_1!TpoLin).Value = True
         optIndLat(.uorstMain_1!IndLat).Value = True
         txtDato(gsIdioma - 1).Text = IIf(IsNull(.uorstMain_1!DetLin), "", .uorstMain_1!DetLin)
         txtDato(2 - gsIdioma).Text = IIf(IsNull(.uorstMain_1!DetLinx), "", .uorstMain_1!DetLinx)
         txtDato(2).Text = IIf(IsNull(.uorstMain_1!FmlLin), "", .uorstMain_1!FmlLin)
         txtDato(3).Text = IIf(IsNull(.uorstMain_1!grppct), "", .uorstMain_1!grppct)
         
         cboBorde(0).ListIndex = .uorstMain_1!IndBdeSup
         cboBorde(1).ListIndex = .uorstMain_1!IndBdeInf
        
         cboEstilo(0).ListIndex = .uorstMain_1!IndFonDet
         chkEfecto.Value = IIf(.uorstMain_1!IndFonDet_Syd = 1, Checked, Unchecked)
         cboEstilo(1).ListIndex = .uorstMain_1!IndFonImp
        'Busca detalle de códigos.
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
   txtLlave(1).Text = ""

  'Datos.
   chkBsePct.Value = vbUnchecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
   optTpoLin(1).Value = True
   optIndLat(0).Value = True
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With

  'Ayudas.
'   lblDatoDeta(2).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   chkBsePct.Enabled = tbHabilitar
   With optTpoLin
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   With optIndLat
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = False
      Next
   End With
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   fraLinea.Enabled = (tbHabilitar And optTpoLin(3).Value)
   
   cboEstilo(0).Enabled = tbHabilitar
   cboEstilo(1).Enabled = tbHabilitar
   chkEfecto.Enabled = tbHabilitar

   
  'Ayudas.
'   cmdDatoAyud(0).Enabled = tbHabilitar
'   lblDatoDeta(0).Enabled = tbHabilitar
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


