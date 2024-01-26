VERSION 5.00
Begin VB.Form frmMAux 
   Caption         =   "[Entidad]"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOnpAfp 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      Picture         =   "frmMAux.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmbcuentas 
      Cancel          =   -1  'True
      Enabled         =   0   'False
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
      Left            =   6720
      Picture         =   "frmMAux.frx":03BC
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1410
      Width           =   720
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
      Index           =   8
      Left            =   1095
      TabIndex        =   6
      Top             =   1455
      Width           =   2460
   End
   Begin VB.Frame fraTpoper 
      Caption         =   " Tipo de Persona "
      ForeColor       =   &H00800000&
      Height          =   570
      Left            =   60
      TabIndex        =   10
      Top             =   2175
      Width           =   7410
      Begin VB.OptionButton opbTpoper 
         Caption         =   "Juridica"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   250
         Width           =   1460
      End
      Begin VB.OptionButton opbTpoper 
         Caption         =   "Natural"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   1
         Left            =   1770
         TabIndex        =   12
         Top             =   250
         Width           =   1460
      End
      Begin VB.OptionButton opbTpoper 
         Caption         =   "No Domiciliado"
         ForeColor       =   &H00C00000&
         Height          =   200
         Index           =   2
         Left            =   3465
         TabIndex        =   13
         Top             =   250
         Width           =   1460
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
      Index           =   7
      Left            =   765
      TabIndex        =   26
      Top             =   4470
      Width           =   2910
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
      Left            =   1110
      TabIndex        =   5
      Top             =   1095
      Width           =   6390
   End
   Begin VB.ComboBox cmbDocIdentidad 
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1845
      Width           =   2745
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   60
      TabIndex        =   31
      Top             =   4860
      Width           =   975
      Begin VB.CheckBox chkEstAux 
         Caption         =   "Activo"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Frame fraCuadro 
      Caption         =   "Tipo"
      ForeColor       =   &H80000002&
      Height          =   735
      Index           =   1
      Left            =   3840
      TabIndex        =   27
      Top             =   4275
      Width           =   3615
      Begin VB.CheckBox chkIndCli 
         Caption         =   "Cliente"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox chkIndPrv 
         Caption         =   "Proveedor"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   1155
      End
      Begin VB.CheckBox chkIndOtr 
         Caption         =   "Otro"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraCuadro 
      Caption         =   " Persona Natural "
      ForeColor       =   &H00800000&
      Height          =   1470
      Index           =   0
      Left            =   60
      TabIndex        =   14
      Top             =   2745
      Width           =   7410
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
         Left            =   5145
         TabIndex        =   24
         Top             =   1005
         Width           =   1250
      End
      Begin VB.ComboBox cmbDocIdentidad 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1020
         Width           =   2190
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
         Left            =   1380
         TabIndex        =   16
         Top             =   285
         Width           =   2190
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
         Left            =   1380
         TabIndex        =   20
         Top             =   660
         Width           =   2190
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
         Left            =   5145
         TabIndex        =   18
         Top             =   300
         Width           =   2190
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documen. :"
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
         Left            =   135
         TabIndex        =   21
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Docuem. :"
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
         Left            =   3885
         TabIndex        =   23
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Apell.Paterno:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Apell.Materno:"
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
         Left            =   3885
         TabIndex        =   17
         Top             =   360
         Width           =   1035
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   1250
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2010
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5115
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
         Picture         =   "frmMAux.frx":06C6
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Picture         =   "frmMAux.frx":0870
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "frmMAux.frx":0A1A
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "frmMAux.frx":0B64
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "frmMAux.frx":0C66
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   60
         Width           =   720
      End
      Begin VB.CommandButton cmdSalir 
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
         Picture         =   "frmMAux.frx":0D68
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   60
         Width           =   720
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
      Index           =   0
      Left            =   1110
      TabIndex        =   3
      Top             =   750
      Width           =   6390
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
      TabIndex        =   1
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail:"
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
      Index           =   10
      Left            =   75
      TabIndex        =   40
      Top             =   1530
      Width           =   840
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Rubro :"
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
      Index           =   9
      Left            =   105
      TabIndex        =   25
      Top             =   4530
      Width           =   525
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
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
      Left            =   75
      TabIndex        =   4
      Top             =   1140
      Width           =   720
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "RUC:"
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
      TabIndex        =   7
      Top             =   1860
      Width           =   360
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Razón Social:"
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
      TabIndex        =   2
      Top             =   780
      Width           =   990
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Auxiliar:"
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
      Top             =   180
      Width           =   585
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
Attribute VB_Name = "frmMAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public psConnStrgSel2, psConnStrgCon2, psConnStrgOrd2
Private pbNuevo As Boolean
Private pbValidada As Boolean
Private pvVacio As Boolean
Private validador As Boolean

Private Sub cmbcuentas_Click()
  CtaAuxiliar = txtLlave(0).Text
  DesAuxiliar = txtDato(0).Text
  frmMCbaGrd.Show vbModal
End Sub

Private Sub cmdOnpAfp_Click()
  frmMAuxOnpAfp.Show vbModal
End Sub

Private Sub Form_Load()
   validador = False
   pbValidada = False
   Dim n_Contador As Integer

   Me.KeyPreview = True
   
   With frmMAuxGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!codaux.DefinedSize
    ']
    '[Datos                            'Cambiar.
      txtDato(0).MaxLength = .uorstMain!razAux.DefinedSize
      txtDato(1).MaxLength = .uorstMain!DirAux.DefinedSize
      txtDato(2).MaxLength = .uorstMain!rucaux.DefinedSize
      txtDato(3).MaxLength = CInt(.uorstMain!razAux.DefinedSize / 3)
      txtDato(4).MaxLength = CInt(.uorstMain!razAux.DefinedSize / 3)
      txtDato(5).MaxLength = CInt(.uorstMain!razAux.DefinedSize / 3)
      txtDato(6).MaxLength = .uorstMain!rucaux.DefinedSize
      txtDato(7).MaxLength = .uorstMain!rubro.DefinedSize
      txtDato(8).MaxLength = .uorstMain!email.DefinedSize
    ']
   End With
   
  ' Configuro tipos documentos de identidad
  For n_Contador = 0 To 5
    If gsIdioma = NvlUsr_Sup Then
      cmbDocIdentidad(0).AddItem Choose(n_Contador + 1, "Otros Tipos de Documentos", "Libreta Electoral o DNI", "Carnet de Extranjería", "RUC", "Pasaporte", "Cédula Diplomática de Identidad")
    Else
      cmbDocIdentidad(0).AddItem Choose(n_Contador + 1, "Other Types of Documents", "Libreta Electoral o DNI", "Card of Extranjeria", "RUT", "Passport", "Diplomatic Identity Certificate")
    End If
    cmbDocIdentidad(0).ItemData(n_Contador) = Choose(n_Contador + 1, 0, 1, 4, 6, 7, 8)
  Next n_Contador
  cmbDocIdentidad(0).ListIndex = 0
   
  For n_Contador = 0 To 4
    If gsIdioma = NvlUsr_Sup Then
      cmbDocIdentidad(1).AddItem Choose(n_Contador + 1, "Ninguno", "Documento Nacional de Identidad", "Carnet de Extranjería", "Pasaporte", "Partida de Nacimiento")
    Else
      cmbDocIdentidad(1).AddItem Choose(n_Contador + 1, "Neither", "National document of Identity", "Card of Extranjeria", "Passport", "Game of Birth")
    End If
    cmbDocIdentidad(1).ItemData(n_Contador) = Choose(n_Contador + 1, 0, 1, 4, 7, 11)
  Next n_Contador
  cmbDocIdentidad(1).ListIndex = 0
   
  If pbNuevo Then
    cmdRetroceder.Enabled = False
    cmdAvanzar.Enabled = False
  End If
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(11, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Auxiliar:", "Razón Social:", "Dirección:", "RUC:", "Apellido Paterno:", "Apellido Materno:", "Nombres:", "Tipo Documen. :", "Nro. Documen. :", "Rubro :", "E-mail")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Auxiliary:", "Firm Name:", "Address:", "RUT:", "Pat. Last Name:", "Mat. Last Name:", "Names:", "Type Documen.:", "Num. Documen. :", "Heading :", "E-mail")
  Next nElemento
  fraTpoper.Caption = Choose(gsIdioma, " Tipo de Persona ", " Type of Person ")
  opbTpoper(0).Caption = Choose(gsIdioma, "Jurídica", "Jurídica")
  opbTpoper(1).Caption = Choose(gsIdioma, "Natural", "Natural")
  opbTpoper(2).Caption = Choose(gsIdioma, "No Domiciliado", "No Domiciliado")
  fraCuadro(0).Caption = Choose(gsIdioma, "Persona Natural", "Natural Person")
  fraCuadro(1).Caption = Choose(gsIdioma, "Tipo", "Type")
  chkIndCli.Caption = Choose(gsIdioma, "Cliente", "Customer")
  chkIndPrv.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  chkIndOtr.Caption = Choose(gsIdioma, "Otro", "Other")
  chkEstAux.Caption = Choose(gsIdioma, "&Activo", "&Active")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']

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
  If Not (frmMAuxGrd.uorstMain.BOF And frmMAuxGrd.uorstMain.EOF) Then
   frmMAuxGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMAuxGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMAuxGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Public Sub cmdCorregir_Click()
  Dim nOpbIndex As Integer
   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
   nOpbIndex = IIf(opbTpoper(0).Value, 0, IIf(opbTpoper(1).Value, 1, 2))
   opbTpoper_Click nOpbIndex
 
 '[Dato con el foco al corregir.       'Cambiar.
   If txtDato(0).Enabled Then
      txtDato(0).SetFocus
   Else
      txtDato(1).SetFocus
   End If
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err

'  [Valida datos de Persona Natural y RUC antes de Grabar
   If Len(Trim(txtDato(2).Text)) < 11 Then
      MsgBox Choose(gsIdioma, "El RUC no puede ser menor de 11 digitos", "The RUT can not be less than 11 digits"), vbExclamation
      txtDato(2).SetFocus
      Exit Sub
   End If
   If opbTpoper(1).Value Then
      If Len(Trim(txtDato(3))) = 0 And Len(Trim(txtDato(4))) = 0 And Len(Trim(txtDato(5))) = 0 And opbTpoper(1).Value Then
         MsgBox Choose(gsIdioma, "Necesita Registrar Datos de Persona Natural", "You need to register data of Natural Person"), vbExclamation
         txtDato(3).SetFocus: Exit Sub
      End If
      If cmbDocIdentidad(1).ListIndex <= 0 Then
        MsgBox Choose(gsIdioma, "Necesita Registrar Datos de Persona Natural", "You need to register data of Natural Person"), vbExclamation
        cmbDocIdentidad(1).SetFocus: Exit Sub
      End If
      If Trim(txtDato(6).Text) = "" Then
        MsgBox Choose(gsIdioma, "Necesita Registrar Datos de Persona Natural", "You need to register data of Natural Person"), vbExclamation
        txtDato(6).SetFocus: Exit Sub
      End If
   Else
      If Not opbTpoper(1).Value And Len(Trim(txtDato(0).Text)) = 0 Then
         MsgBox Choose(gsIdioma, "No es Persona Natural, Debe ingresar la Razon Social", "It not is Natural Person, you must enter Firm Name"), vbExclamation
         txtDato(0).SetFocus
         Exit Sub
      End If
   End If
   ']
   '[Variables para Datos de Persona Natural'
   psConnStrgSel2 = "SELECT codaux, nomaux, apepataux, apemataux, codtdi, numdci, "
   psConnStrgSel2 = psConnStrgSel2 & "codemp, UsrCre, FyHCre, UsrMdf, FyHMdf "
   psConnStrgSel2 = psConnStrgSel2 & "FROM TgAuxNat "
   psConnStrgSel2 = psConnStrgSel2 & "WHERE codemp='" & gsCodEmp & "' "
   psConnStrgCon2 = "AND CodAux="
   psConnStrgOrd2 = " ORDER BY 1"
   pvVacio = False
   ']
   
   With frmMAuxGrd                     'Cambiar Formulario de Grid.
      
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
      '[Actualiza Datos de Persona Natural'
      If opbTpoper(1).Value Then
         With .uorstMai2
            If pbNuevo Or pvVacio Then
               !UsrCre = gsAbvUsr
               !FyHCre = Now
            Else
               !UsrMdf = gsAbvUsr
               !FyHMdf = Now
            End If
            .Update
         End With
      End If
      ']
      
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
   
      If pbNuevo Then
         .uorstMain.Requery
         .ppDatosGrid
       '[Búsqueda de llave actual.     'Cambiar.
         .uorstMain.Find "CodAux='" & txtLlave(0).Text & "'"
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
  
   frmMAuxGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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
'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub

Private Sub txtLlave_Change(Index As Integer)
If txtLlave(0).Text <> "" Then
    cmbcuentas.Enabled = True
    cmdOnpAfp.Enabled = True '2014-08-04 adicion frmoppafp
End If
End Sub

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
   'Cambiar.
   If pbValidada Then txtDato(0).SetFocus
'   If pbValidada Then
'      txtDato(1).Text = txtLlave(0).Text
'      txtDato(0).SetFocus
'   End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
'   On Error GoTo Err

   Dim dvRegistro As Variant
   Dim nOpbIndex As Integer
   
   VerificadorRUC
   
   'If validador = False Then Exit Sub
   'txtLlave(0).Enabled = False
   
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
      With frmMAuxGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodAux='" & txtLlave(0).Text & "'"
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
      nOpbIndex = IIf(opbTpoper(0).Value, 0, IIf(opbTpoper(1).Value, 1, 2))
      opbTpoper_Click nOpbIndex
      pbValidada = True
   Else
      cmdGrabar.Enabled = False
      upHabilitacion False
      pbValidada = False
   End If
'Err:
'   gpErrores
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

'[Angel
'[ARREGLAR: Se está pintando lo digitado antes del actual.
   If Index = 5 Or Index = 3 Or Index = 4 Then
      txtDato(0).Text = txtDato(3).Text & " " & txtDato(4).Text & "," & txtDato(5).Text
   End If
']ARREGLAR.
']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

  'Completa con ceros a la izquierda.
'   Select Case Index
'   Case 2                              'Cambiar (añadir índices).
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
  
  '[ Para Validar Datos de Auxiliares
  Select Case Index
  Case 2
     If Len(Trim(txtDato(2).Text)) < 11 Then
        MsgBox Choose(gsIdioma, "El RUC no puede ser menor de 11 digitos", "The RUT can not be less than 11 digits"), vbExclamation
        Cancel = True
        txtDato(2).SetFocus
     End If
  Case 3 To 5
     If Len(Trim(txtDato(3))) = 0 And Len(Trim(txtDato(4))) = 0 And Len(Trim(txtDato(5))) = 0 And opbTpoper(1).Value Then
        MsgBox Choose(gsIdioma, "Necesita Registrar Datos de Persona Natural", "You need to register data of Natural Person"), vbExclamation
        Cancel = True
        txtDato(3).SetFocus
     End If
  End Select
  ']
  Exit Sub
Err:
   gpErrores
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
  Dim n_Index As Integer
   
  On Error GoTo Err

  With frmMAuxGrd
    '[Datos No pertenecen al formulario  -Angel 26/11/2003
    '[Datos de Persona Natural'
    With .uorstMai2
      '[ARREGLAR. No estaba al 31/12/2003. Puesto por Raúl. Sin esto no permite mdfi ningún Auxiliar.
      psConnStrgSel2 = "SELECT CodAux, NomAux, ApePatAux, ApeMatAux, codtdi, numdci, "
      psConnStrgSel2 = psConnStrgSel2 & "codemp, UsrCre, FyHCre, UsrMdf, FyHMdf "
      psConnStrgSel2 = psConnStrgSel2 & "FROM TgAuxNat "
      psConnStrgSel2 = psConnStrgSel2 & "WHERE codemp='" & gsCodEmp & "' "
      psConnStrgCon2 = "AND CodAux="
      psConnStrgOrd2 = " ORDER BY 1"
      ']ARREGLAR.
      .Close
         
      '[ARREGLAR. Línea cambiada por If. (RAUL 09/01/2004).
      If tnFase = 0 Then
        .Source = psConnStrgSel2 & psConnStrgCon2 & "'" & txtLlave(0).Text & "'" & psConnStrgOrd2
      Else
        .Source = psConnStrgSel2 & psConnStrgCon2 & "'" & frmMAuxGrd.uorstMain!codaux & "'" & psConnStrgOrd2
      End If
      ']ARREGLAR.
      .Open
    End With
    
    '[Datos de Persona Natural'
    If tnFase = 1 Then
      opbTpoper(1).Value = (.uorstMain!TpoPer = TPOPER_NAT)
    End If
    
    If opbTpoper(1).Value And (pbNuevo Or .uorstMai2.RecordCount = 0) Then
      .uorstMai2.AddNew
      pvVacio = True
    End If
    ']
      
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        .uorstMain!codemp = gsCodEmp
        .uorstMain!codaux = txtLlave(0).Text
      End If
      ' Reemplazo los caracteres
      txtDato(0).Text = gfSacaEntRetApos(txtDato(0).Text)
      txtDato(3).Text = gfSacaEntRetApos(txtDato(3).Text)
      txtDato(4).Text = gfSacaEntRetApos(txtDato(4).Text)
      txtDato(5).Text = gfSacaEntRetApos(txtDato(5).Text)
      
      'Datos.
      n_Index = cmbDocIdentidad(0).ListIndex
      .uorstMain!DirAux = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
      .uorstMain!rucaux = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
      .uorstMain!TpoDci = IIf(cmbDocIdentidad(0).ItemData(n_Index) = 8, "0A", Format(cmbDocIdentidad(0).ItemData(n_Index), "00"))
      .uorstMain!EstAux = IIf(chkEstAux.Value = vbChecked, ESTAUX_ACT, ESTAUX_INA)
      .uorstMain!IndCli = IIf(chkIndCli.Value = vbChecked, INDAUX_CLI_ACT, INDAUX_CLI_INA)
      .uorstMain!IndPrv = IIf(chkIndPrv.Value = vbChecked, INDAUX_PRV_ACT, INDAUX_PRV_INA)
      .uorstMain!IndOtr = IIf(chkIndOtr.Value = vbChecked, INDAUX_OTR_ACT, INDAUX_OTR_INA)
      .uorstMain!TpoPer = IIf(opbTpoper(0).Value, TPOPER_JUR, IIf(opbTpoper(1).Value, TPOPER_NAT, TPOPER_DOM))
      .uorstMain!rubro = IIf(txtDato(7).Text = "", Null, txtDato(7).Text)
      .uorstMain!email = IIf(txtDato(8).Text = "", Null, txtDato(8).Text)
      
      If opbTpoper(1).Value Then
        .uorstMain!razAux = Trim(txtDato(3).Text) & " " & Trim(txtDato(4).Text) & "," & Trim(txtDato(5).Text)
        .uorstMai2!codemp = gsCodEmp
        .uorstMai2!codaux = txtLlave(0).Text
        .uorstMai2!ApePatAux = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
        .uorstMai2!ApeMatAux = IIf(txtDato(4).Text = "", Null, txtDato(4).Text)
        .uorstMai2!NomAux = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
        n_Index = cmbDocIdentidad(1).ListIndex
        '13/04/2009
        .uorstMai2!codtdi = Format(cmbDocIdentidad(1).ItemData(n_Index), "00")
        .uorstMai2!numdci = IIf(txtDato(6).Text = "", Null, txtDato(6).Text)
      Else
        .uorstMain!razAux = txtDato(0).Text
      End If
    Else
      'Llaves.
      txtLlave(0).Text = .uorstMain!codaux
      
      'Datos.
      txtDato(1).Text = IIf(IsNull(.uorstMain!DirAux), "", .uorstMain!DirAux)
      txtDato(2).Text = IIf(IsNull(.uorstMain!rucaux), "", .uorstMain!rucaux)
      
      n_Index = IIf(Trim(.uorstMain!TpoDci) = "0A", 5, IIf(Val(.uorstMain!TpoDci) <= 1, Val(.uorstMain!TpoDci), IIf(Val(.uorstMain!TpoDci) = 4, 2, Val(.uorstMain!TpoDci) - 3)))
      cmbDocIdentidad(0).ListIndex = n_Index
      chkEstAux.Value = IIf(.uorstMain!EstAux = ESTAUX_ACT, vbChecked, vbUnchecked)
      chkIndCli.Value = IIf(.uorstMain!IndCli = INDAUX_CLI_ACT, vbChecked, vbUnchecked)
      chkIndPrv.Value = IIf(.uorstMain!IndPrv = INDAUX_PRV_ACT, vbChecked, vbUnchecked)
      chkIndOtr.Value = IIf(.uorstMain!IndOtr = INDAUX_OTR_ACT, vbChecked, vbUnchecked)
      opbTpoper(0).Value = (.uorstMain!TpoPer = TPOPER_JUR)
      opbTpoper(1).Value = (.uorstMain!TpoPer = TPOPER_NAT)
      opbTpoper(2).Value = (.uorstMain!TpoPer = TPOPER_DOM)
      txtDato(7).Text = IIf(IsNull(.uorstMain!rubro), "", .uorstMain!rubro)
      txtDato(8).Text = IIf(IsNull(.uorstMain!email), "", .uorstMain!email)
    
      If opbTpoper(1).Value Then
        txtDato(3).Text = IIf(IsNull(.uorstMai2!ApePatAux), "", .uorstMai2!ApePatAux)
        txtDato(4).Text = IIf(IsNull(.uorstMai2!ApeMatAux), "", .uorstMai2!ApeMatAux)
        txtDato(5).Text = IIf(IsNull(.uorstMai2!NomAux), "", .uorstMai2!NomAux)
        txtDato(0).Text = txtDato(3).Text & " " & txtDato(4).Text & "," & txtDato(5).Text
        n_Index = Val(IIf(IsNull(.uorstMai2!codtdi), 0, .uorstMai2!codtdi))
        n_Index = IIf(n_Index = 4, 2, IIf(n_Index = 7, 3, IIf(n_Index = 11, 4, n_Index)))
        cmbDocIdentidad(1).ListIndex = n_Index
        txtDato(6).Text = IIf(IsNull(.uorstMai2!numdci), "", .uorstMai2!numdci)
      Else
        txtDato(0).Text = IIf(IsNull(.uorstMain!razAux), "", .uorstMain!razAux)
      End If
    End If
  End With
      
  Exit Sub
Err:
  gpErrores
 '  Resume

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
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Text = ""
      Next
   End With
   chkEstAux.Value = vbChecked
   opbTpoper(0).Value = True
   chkIndCli = vbUnchecked
   chkIndOtr = vbUnchecked
   chkIndPrv = vbUnchecked

  'Ayudas.
'   lblDatoDeta(2).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
   End With
   
   chkIndCli.Enabled = tbHabilitar
   chkIndPrv.Enabled = tbHabilitar
   chkIndOtr.Enabled = tbHabilitar
   opbTpoper(0).Enabled = tbHabilitar
   opbTpoper(1).Enabled = tbHabilitar
   opbTpoper(2).Enabled = tbHabilitar
   chkEstAux.Enabled = tbHabilitar
   cmbDocIdentidad(0).Enabled = tbHabilitar
   cmbDocIdentidad(1).Enabled = tbHabilitar

End Sub

'[Código propio del formulario.
Private Sub opbTpoper_Click(Index As Integer)
  If Index <> 1 Then
    fraCuadro(0).Enabled = False
    txtDato(3).Text = "": txtDato(4).Text = "": txtDato(5).Text = ""
    cmbDocIdentidad(1).ListIndex = 0: txtDato(6).Text = ""
  End If
  If (Index = 1 And opbTpoper(Index).Enabled) Then fraCuadro(0).Enabled = True
  txtDato(0).Enabled = (Index <> 1 And opbTpoper(Index).Enabled)
End Sub
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
Function VerificadorRUC()
Dim lnSuma As Long
Dim lnResiduo As Long
Dim lnResta As Long
Dim lnDigitoVerificador As Long

lnSuma = 0
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 1, 1)) * 5
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 2, 1)) * 4
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 3, 1)) * 3
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 4, 1)) * 2
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 5, 1)) * 7
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 6, 1)) * 6
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 7, 1)) * 5
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 8, 1)) * 4
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 9, 1)) * 3
lnSuma = lnSuma + Val(Mid(txtLlave(0).Text, 10, 1)) * 2
lnResiduo = lnSuma Mod 11
lnResta = 11 - lnResiduo

If lnResta = 10 Then
    lnDigitoVerificador = 0
ElseIf lnResta = 11 Then
    lnDigitoVerificador = 1
Else
    lnDigitoVerificador = lnResta
End If

If lnDigitoVerificador = Val(Right(txtLlave(0).Text, 1)) Then
    MsgBox "Ruc Correcto "  'lnDigitoVerificador
    validador = True
Else
    MsgBox "Ruc Incorrecto "  'lnDigitoVerificador
    validador = False
End If


End Function

