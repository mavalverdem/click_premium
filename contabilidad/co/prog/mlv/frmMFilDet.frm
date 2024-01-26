VERSION 5.00
Begin VB.Form frmMFilDet 
   Caption         =   "[Entidad]"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox checkpro 
      Caption         =   "Prorrateo"
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
      Height          =   255
      Left            =   6480
      TabIndex        =   44
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   1
      Left            =   7920
      Picture         =   "frmMFilDet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   960
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   1
      Left            =   1020
      TabIndex        =   38
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtDato 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   1020
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   7920
      Picture         =   "frmMFilDet.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   280
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
      Left            =   1020
      TabIndex        =   9
      Top             =   1800
      Width           =   6795
   End
   Begin VB.Frame fraColumna 
      Caption         =   " Columnas "
      ForeColor       =   &H00000080&
      Height          =   1710
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   8145
      Begin VB.CheckBox checkflag 
         Caption         =   "INCENTIVE"
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
         Height          =   255
         Index           =   2
         Left            =   6480
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox checkflag 
         Caption         =   "GIFT"
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
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   42
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox checkflag 
         Caption         =   "FOOD"
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
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   41
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         ItemData        =   "frmMFilDet.frx":0354
         Left            =   6480
         List            =   "frmMFilDet.frx":0356
         TabIndex        =   35
         Top             =   240
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         ItemData        =   "frmMFilDet.frx":0358
         Left            =   3720
         List            =   "frmMFilDet.frx":035A
         TabIndex        =   26
         Top             =   1290
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         ItemData        =   "frmMFilDet.frx":035C
         Left            =   1080
         List            =   "frmMFilDet.frx":035E
         TabIndex        =   18
         Top             =   1290
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         ItemData        =   "frmMFilDet.frx":0360
         Left            =   3720
         List            =   "frmMFilDet.frx":0362
         TabIndex        =   24
         Top             =   930
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         ItemData        =   "frmMFilDet.frx":0364
         Left            =   1080
         List            =   "frmMFilDet.frx":0366
         TabIndex        =   16
         Top             =   930
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         ItemData        =   "frmMFilDet.frx":0368
         Left            =   3720
         List            =   "frmMFilDet.frx":036A
         TabIndex        =   22
         Top             =   570
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         ItemData        =   "frmMFilDet.frx":036C
         Left            =   1080
         List            =   "frmMFilDet.frx":036E
         TabIndex        =   14
         Top             =   570
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         ItemData        =   "frmMFilDet.frx":0370
         Left            =   3720
         List            =   "frmMFilDet.frx":0372
         TabIndex        =   20
         Top             =   210
         Width           =   1530
      End
      Begin VB.ComboBox cboColumna 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmMFilDet.frx":0374
         Left            =   1080
         List            =   "frmMFilDet.frx":0376
         TabIndex        =   12
         Top             =   210
         Width           =   1530
      End
      Begin VB.Label lblTexto 
         Caption         =   "Flag :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   12
         Left            =   5400
         TabIndex        =   36
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 8 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   11
         Left            =   2760
         TabIndex        =   25
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 7 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 6 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   10
         Left            =   2760
         TabIndex        =   23
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 5 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   975
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 4 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   9
         Left            =   2760
         TabIndex        =   21
         Top             =   630
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 3 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   615
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 1 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblTexto 
         Caption         =   "Columna 2 :"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   8
         Left            =   2760
         TabIndex        =   19
         Top             =   270
         Width           =   885
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
      Index           =   2
      Left            =   1020
      TabIndex        =   7
      Top             =   1320
      Width           =   6795
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
      Left            =   1440
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3840
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
         Picture         =   "frmMFilDet.frx":0378
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "frmMFilDet.frx":04C2
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Left            =   1200
         Picture         =   "frmMFilDet.frx":05C4
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Picture         =   "frmMFilDet.frx":06C6
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Picture         =   "frmMFilDet.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "frmMFilDet.frx":09BA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   60
         Width           =   360
      End
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
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   39
      Top             =   960
      Width           =   5790
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Centro de Costos:"
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
      Height          =   420
      Index           =   13
      Left            =   60
      TabIndex        =   37
      Top             =   920
      Width           =   825
      WordWrap        =   -1  'True
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
      Height          =   285
      Index           =   0
      Left            =   1980
      TabIndex        =   5
      Top             =   600
      Width           =   5790
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
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1545
      X2              =   1590
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Label lblTexto 
      Caption         =   "Traducción:"
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   8
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label lblTexto 
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
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   555
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
      X2              =   8160
      Y1              =   495
      Y2              =   495
   End
End
Attribute VB_Name = "frmMFilDet"
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
   With frmMFilGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = (.uorstMain_1!NroLin.DefinedSize - 1)
      txtLlave(1).MaxLength = 1
    ']
   
    '[Datos.                           'Cambiar.
      txtDato(0).MaxLength = .uorstMain_1!codcta.DefinedSize
      txtDato(2).MaxLength = .uorstMain_1!DetLin.DefinedSize
      txtDato(3).MaxLength = .uorstMain_1!DetLinx.DefinedSize
    ']
   End With
   
   For n_Index = 0 To 8
     
     cboColumna(0).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     cboColumna(1).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     cboColumna(2).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     cboColumna(3).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     cboColumna(4).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     cboColumna(5).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     cboColumna(6).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     cboColumna(7).AddItem Choose(n_Index + 1, IIf(gsIdioma = NvlUsr_Sup, "Ninguno", "Neither"), "CLO", "FUN_CCO", "LINK_TOBV", "SVC_FOO", "SVC_GIF", "SVC_EXT", "VFT_FLO", IIf(gsIdioma = NvlUsr_Sup, "Saldo", "Balance"))
     
   Next n_Index
   
   cboColumna(8).AddItem ""
   cboColumna(8).AddItem "Pedidos"
   cboColumna(8).AddItem "Vales"
   cboColumna(8).AddItem "Lectura"
   
   cboColumna(0).ListIndex = 0: cboColumna(1).ListIndex = 0
   cboColumna(2).ListIndex = 0: cboColumna(3).ListIndex = 0
   cboColumna(4).ListIndex = 0: cboColumna(5).ListIndex = 0
   cboColumna(6).ListIndex = 0: cboColumna(7).ListIndex = 0
   
  If pbNuevo Then
    cmdRetroceder.Enabled = False
    cmdAvanzar.Enabled = False
  End If
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(12, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Línea :", "Cuenta :", "Descripción :", "Traducción :", "Columna 3 :", "Columna 4 :", "Columna 5 :", "Columna 6 :", "Columna 7 :", "Columna 8 :", "Columna 9 :", "Columna 10 :", "Flag :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Line :", "Account :", "Description:", "Translation:", "Column 3 :", "Column 4 :", "Column 5 :", "Column 6 :", "Column 7 :", "Column 8 :", "Column 9 :", "Column 10 :", "Flag :")
  Next nElemento
  fraColumna.Caption = Choose(gsIdioma, " Columnas ", " Columns ")
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
  If Not (frmMFilGrd.uorstMain_1.BOF And frmMFilGrd.uorstMain_1.EOF) Then
   frmMFilGrd.uorstMain_1.CancelUpdate 'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMFilGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMFilGrd.uorstMain_1, Me 'Cambiar Formulario de Grid.
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
   
   With frmMFilGrd                     'Cambiar Formulario de Grid.
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
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain_1.Requery
         .upDatosGrid 1
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain_1.Find "nrolin='" & txtLlave(0).Text & txtLlave(1).Text & "'"
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
  
   frmMFilGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
     txtDato(Index).SetFocus
   Case 1
     txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
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
   Select Case Index                   'Cambiar (añadir índices).
   Case 0
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
 
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 And Index = 1 Then
      With frmMFilGrd.uorstMain_1        'Cambiar Formulario de Grid.
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
   If Index = 1 Or Index = 2 Then      'Cambiar (añadir índices).
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
 ']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus Index
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err

  'Busca el dato en su tabla principal.
   Select Case Index
    Case 0                             'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
    Case 1                             'Cambiar (añadir índices).
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tnIndex As Integer)
  Select Case tnIndex
   Case 0                              'Cambiar (añadir índices).
        modAyuBus.Cta_Cod "estcta='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
        txtDato(tnIndex).Text = frmOAyuBus.uvDato1
        lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1
        modAyuBus.CCo_Cod "estcco='" & ESTCTA_ACT & "'", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
    With frmMFilGrd.uorstCOCta
      If .RecordCount > 0 Then .MoveFirst
      .Find "codcta='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
      End If
    End With
  Case 1
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With frmMFilGrd.uorstCoCCo
      If .RecordCount > 0 Then .MoveFirst
      .Find "codcco='" & txtDato(tnIndex).Text & "'"
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
  
  With frmMFilGrd                     'Cambiar Formulario de Grid.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        .uorstMain_1!codemp = frmMFilGrd.uorstMain_0!codemp
        .uorstMain_1!pdoano = frmMFilGrd.uorstMain_0!pdoano
        .uorstMain_1!codfil = frmMFilGrd.uorstMain_0!codfil
        .uorstMain_1!NroLin = txtLlave(0).Text & Trim(txtLlave(1).Text)
      End If
      
      'Datos.
      .uorstMain_1!codcta = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      .uorstMain_1!codcco = IIf(txtDato(1).Text = "", "", txtDato(1).Text)
      .uorstMain_1!DetLin = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
      .uorstMain_1!DetLinx = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
      .uorstMain_1!coldet1 = cboColumna(0).ListIndex
      .uorstMain_1!coldet2 = cboColumna(1).ListIndex
      .uorstMain_1!coldet3 = cboColumna(2).ListIndex
      .uorstMain_1!coldet4 = cboColumna(3).ListIndex
      .uorstMain_1!coldet5 = cboColumna(4).ListIndex
      .uorstMain_1!coldet6 = cboColumna(5).ListIndex
      .uorstMain_1!coldet7 = cboColumna(6).ListIndex
      .uorstMain_1!coldet8 = cboColumna(7).ListIndex
      .uorstMain_1!colflag = IIf(cboColumna(8).ListIndex = -1, 0, cboColumna(8).ListIndex)
      .uorstMain_1!flagpro = IIf(checkpro.Value = Checked, 1, 0)
      .uorstMain_1!flagfood = IIf(checkflag(0).Value = Checked, 1, 0)
      .uorstMain_1!flaggift = IIf(checkflag(1).Value = Checked, 1, 0)
      .uorstMain_1!flagincentive = IIf(checkflag(2).Value = Checked, 1, 0)
      
    Else
      'Llaves.
      txtLlave(0).Text = Left(.uorstMain_1!NroLin, 3)
      txtLlave(1).Text = Mid(.uorstMain_1!NroLin, 4, 1)
      
      'Datos.
      txtDato(0).Text = IIf(IsNull(.uorstMain_1!codcta), "", .uorstMain_1!codcta)
      txtDato(1).Text = IIf(IsNull(.uorstMain_1!codcco), "", .uorstMain_1!codcco)
      txtDato(2).Text = IIf(IsNull(.uorstMain_1!DetLin), "", .uorstMain_1!DetLin)
      txtDato(3).Text = IIf(IsNull(.uorstMain_1!DetLinx), "", .uorstMain_1!DetLinx)
      cboColumna(0).ListIndex = .uorstMain_1!coldet1
      cboColumna(1).ListIndex = .uorstMain_1!coldet2
      cboColumna(2).ListIndex = .uorstMain_1!coldet3
      cboColumna(3).ListIndex = .uorstMain_1!coldet4
      cboColumna(4).ListIndex = .uorstMain_1!coldet5
      cboColumna(5).ListIndex = .uorstMain_1!coldet6
      cboColumna(6).ListIndex = .uorstMain_1!coldet7
      cboColumna(7).ListIndex = .uorstMain_1!coldet8
      cboColumna(8).ListIndex = .uorstMain_1!colflag
      checkpro.Value = IIf(.uorstMain_1!flagpro = 1, Checked, Unchecked)
      checkflag(0).Value = IIf(.uorstMain_1!flagfood = 1, Checked, Unchecked)
      checkflag(1).Value = IIf(.uorstMain_1!flaggift = 1, Checked, Unchecked)
      checkflag(2).Value = IIf(.uorstMain_1!flagincentive = 1, Checked, Unchecked)
      
      'Busca detalle de códigos.
       ppAyuDet 0
       ppAyuDet 1
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
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Text = ""
    Next
  End With
  For dnContador = 0 To 7
    cboColumna(dnContador).ListIndex = 0
  Next dnContador
  'Ayudas.
  lblDatoDeta(0).Caption = ""
  lblDatoDeta(1).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  fraColumna.Enabled = tbHabilitar
   
  'Ayudas.
   cmdDatoAyud(0).Enabled = tbHabilitar
   cmdDatoAyud(1).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar
   lblDatoDeta(1).Enabled = tbHabilitar
   
   checkpro.Enabled = tbHabilitar
   checkflag(0).Enabled = tbHabilitar
   checkflag(1).Enabled = tbHabilitar
   checkflag(2).Enabled = tbHabilitar
   
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
