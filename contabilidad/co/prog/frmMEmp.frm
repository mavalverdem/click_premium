VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmMEmp 
   Appearance      =   0  'Flat
   Caption         =   "[Entidad]"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEstEmp 
      Caption         =   "Activo"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4920
      TabIndex        =   40
      Top             =   1800
      Width           =   1995
   End
   Begin VB.CheckBox chkBuenContri 
      Caption         =   "Buen Contribuyente"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2640
      TabIndex        =   39
      Top             =   1800
      Width           =   1995
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
      Left            =   1080
      TabIndex        =   11
      Top             =   2160
      Width           =   6430
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
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Width           =   1275
   End
   Begin TabDlg.SSTab tabRegistro 
      Height          =   3405
      Left            =   285
      TabIndex        =   12
      Top             =   2625
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6006
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Reprentante Legal"
      TabPicture(0)   =   "frmMEmp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCuadro(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contador"
      TabPicture(1)   =   "frmMEmp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCuadro(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Logo"
      TabPicture(2)   =   "frmMEmp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmCuadro(6)"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraCuadro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1665
         Index           =   1
         Left            =   -74895
         TabIndex        =   22
         Top             =   420
         Width           =   7110
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
            Index           =   12
            Left            =   3780
            TabIndex        =   30
            Top             =   1230
            Width           =   3000
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
            Index           =   10
            Left            =   3780
            TabIndex        =   26
            Top             =   555
            Width           =   3000
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
            Index           =   11
            Left            =   345
            TabIndex        =   28
            Top             =   1230
            Width           =   3000
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
            Index           =   9
            Left            =   345
            TabIndex        =   24
            Top             =   555
            Width           =   3000
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Documento de Identidad :"
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
            Index           =   13
            Left            =   3780
            TabIndex        =   29
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Materno :"
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
            Index           =   11
            Left            =   3780
            TabIndex        =   25
            Top             =   270
            Width           =   1290
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Paterno :"
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
            Left            =   345
            TabIndex        =   23
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nombres :"
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
            Index           =   12
            Left            =   345
            TabIndex        =   27
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame fraCuadro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1665
         Index           =   0
         Left            =   105
         TabIndex        =   13
         Top             =   420
         Width           =   7110
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
            Left            =   345
            TabIndex        =   15
            Top             =   555
            Width           =   3000
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
            Left            =   345
            TabIndex        =   19
            Top             =   1230
            Width           =   3000
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
            Left            =   3780
            TabIndex        =   17
            Top             =   555
            Width           =   3000
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
            Left            =   3780
            TabIndex        =   21
            Top             =   1230
            Width           =   3000
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nombres :"
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
            Left            =   345
            TabIndex        =   18
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Paterno :"
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
            Left            =   345
            TabIndex        =   14
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Apellido Materno :"
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
            Left            =   3780
            TabIndex        =   16
            Top             =   270
            Width           =   1290
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Documento de Identidad :"
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
            Left            =   3780
            TabIndex        =   20
            Top             =   960
            Width           =   1815
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2925
         Index           =   6
         Left            =   -73320
         TabIndex        =   38
         Top             =   360
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   5159
         _StockProps     =   14
         Caption         =   " Logo  "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   2520
            Index           =   0
            Left            =   210
            Shape           =   4  'Rounded Rectangle
            Top             =   285
            Width           =   3420
         End
         Begin VB.Image imgLogo 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Left            =   225
            ToolTipText     =   "Haga doble click para logo empresa"
            Top             =   600
            Width           =   3375
         End
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
      TabIndex        =   7
      Top             =   1440
      Width           =   6430
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
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   6430
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2175
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6150
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
         Picture         =   "frmMEmp.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "frmMEmp.frx":01FE
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "frmMEmp.frx":03A8
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "frmMEmp.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Picture         =   "frmMEmp.frx":05F4
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "frmMEmp.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   36
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
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   6430
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
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Actvidad :"
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
      Left            =   60
      TabIndex        =   10
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "RUC :"
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
      Left            =   60
      TabIndex        =   8
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Localidad :"
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
      TabIndex        =   6
      Top             =   1500
      Width           =   780
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Dirección :"
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
      TabIndex        =   4
      Top             =   1140
      Width           =   765
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
      Caption         =   "Empresa:"
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
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   7820
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMEmp"
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
   
   With frmMEmpGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!codemp.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
      txtDato(0).MaxLength = .uorstMain!RazEmp.DefinedSize
      txtDato(1).MaxLength = .uorstMain!direccion.DefinedSize
      txtDato(2).MaxLength = .uorstMain!localidademp.DefinedSize
      txtDato(3).MaxLength = .uorstMain!RUCEmp.DefinedSize
      txtDato(4).MaxLength = .uorstMain!actividademp.DefinedSize
      txtDato(5).MaxLength = .uorstMain!repapepaterno.DefinedSize
      txtDato(6).MaxLength = .uorstMain!repapematerno.DefinedSize
      txtDato(7).MaxLength = .uorstMain!repnombre.DefinedSize
      txtDato(8).MaxLength = .uorstMain!repdocumento.DefinedSize
      txtDato(9).MaxLength = .uorstMain!conapepaterno.DefinedSize
      txtDato(10).MaxLength = .uorstMain!conapematerno.DefinedSize
      txtDato(11).MaxLength = .uorstMain!connombre.DefinedSize
      txtDato(12).MaxLength = .uorstMain!condocumento.DefinedSize
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
  ReDim aLabel(14, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Empresa:", "Razón Social:", "Dirección :", "Localidad :", "R.U.C.:", "Actividad :", "Apellido Paterno :", "Apellido Materno :", "Nombres :", "Documento de Identidad :", "Apellido Paterno :", "Apellido Materno :", "Nombres :", "Documento de Identidad :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Company:", "Firm Name:", "Address:", "Locality :", "R.U.T.:", "Activity :", "Pat. Last Name:", "Mat. Last Name:", "Names:", "Identity Card :", "Pat. Last Name:", "Mat. Last Name:", "Names:", "Identity Card :")
  Next nElemento
  chkBuenContri.Caption = Choose(gsIdioma, "&Buen Contribuyente", "&Good Taxpayer") '2015-08-27 ctr obligac sunat
  chkEstEmp.Caption = Choose(gsIdioma, "&Activo", "&Active") '2016-06-02 adicion campo EstAct en Empresa
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
   frmMEmpGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMEmpGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMEmpGrd.uorstMain, Me 'Cambiar Formulario de Grid.
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

   With frmMEmpGrd                     'Cambiar Formulario de Grid.
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
         .uorstMain.Find "CodEmp='" & txtLlave(0).Text & "'"
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
      ' Actualizo si es la misma empresa
      If gsCodEmp = frmMEmpGrd.uorstMain!codemp Then
        gsRazEmp = IIf(IsNull(frmMEmpGrd.uorstMain!RazEmp), "", frmMEmpGrd.uorstMain!RazEmp)
        gsRUCEmp = IIf(IsNull(frmMEmpGrd.uorstMain!RUCEmp), "", frmMEmpGrd.uorstMain!RUCEmp)
        gsDirEmp = IIf(IsNull(frmMEmpGrd.uorstMain!direccion), "", frmMEmpGrd.uorstMain!direccion)
        gsLocEmp = IIf(IsNull(frmMEmpGrd.uorstMain!localidademp), "", frmMEmpGrd.uorstMain!localidademp)
        gsGirEmp = IIf(IsNull(frmMEmpGrd.uorstMain!actividademp), "", frmMEmpGrd.uorstMain!actividademp)
        gsRepEmp = IIf(IsNull(frmMEmpGrd.uorstMain!repnombre), "", frmMEmpGrd.uorstMain!repnombre & ", ") & IIf(IsNull(frmMEmpGrd.uorstMain!repapepaterno), "", frmMEmpGrd.uorstMain!repapepaterno & " ") & IIf(IsNull(frmMEmpGrd.uorstMain!repapematerno), "", frmMEmpGrd.uorstMain!repapematerno)
        gsRepDNIEmp = IIf(IsNull(frmMEmpGrd.uorstMain!repdocumento), "", frmMEmpGrd.uorstMain!repdocumento)
        gsConEmp = IIf(IsNull(frmMEmpGrd.uorstMain!connombre), "", frmMEmpGrd.uorstMain!connombre & ", ") & IIf(IsNull(frmMEmpGrd.uorstMain!conapepaterno), "", frmMEmpGrd.uorstMain!conapepaterno & " ") & IIf(IsNull(frmMEmpGrd.uorstMain!conapematerno), "", frmMEmpGrd.uorstMain!conapematerno)
        gsConDNIEmp = IIf(IsNull(frmMEmpGrd.uorstMain!condocumento), "", frmMEmpGrd.uorstMain!condocumento)
        
        gsBuenContriEmp = IIf(IsNull(frmMEmpGrd.uorstMain!BuenContri), "", frmMEmpGrd.uorstMain!BuenContri) '2015-08-27 ctr obligac sunat
        '2016-06-02 adicion campo EstAct en Empresa
        'aqui no va EstEmp, por que no es variable global
        frmMain.lblVar(0) = gsRazEmp
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
  
   frmMEmpGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
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


Private Sub imgLogo_DblClick()
  
  On Error GoTo CancelaDialogo
  frmMain.cdlDialogo.DialogTitle = "Seleccionar Imagen"
  frmMain.cdlDialogo.CancelError = True
  frmMain.cdlDialogo.Flags = cdlOFNHideReadOnly
  frmMain.cdlDialogo.DefaultExt = ".bmp"
  frmMain.cdlDialogo.Filter = "Imagen BMP (*.bmp)|*.bmp|Imagen JPEG(*.jpg)|*.jpg|Imagen GIF (*.gif)|*.gif|Todos los archivos(*.*)|*.*"
  frmMain.cdlDialogo.FilterIndex = 1
  frmMain.cdlDialogo.ShowOpen
  imgLogo.Picture = LoadPicture(frmMain.cdlDialogo.FileName)
  imgLogo.Tag = frmMain.cdlDialogo.FileName
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then
    MsgBox error(Err.Number)
    Exit Sub
  End If
  On Error GoTo 0

End Sub
Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'ini 2015-01-07 adiciono imagen empresa
Dim s_Estado_Ina As Integer
s_Estado_Ina = 0
'fin 2015-01-07 adiciono imagen empresa

  ' Elimino la fotografia
  If Button = vbRightButton And Shift = s_Estado_Ina Then
    If MsgBox("Desea Eliminar Logo de la Empresa", vbQuestion + vbYesNo) = vbYes Then
      imgLogo.Picture = LoadPicture("")
      imgLogo.Tag = ""
    End If
  End If
End Sub


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
   Select Case Index                   'Cambiar (añadir índices).
   Case 0
      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
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
   If Len(Trim(txtLlave(Index).Text)) <> 0 Then
      With frmMEmpGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "CodEmp='" & txtLlave(0).Text & "'"
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
'   If Index = 1 Then                   'Cambiar (añadir índices).
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

   With frmMEmpGrd
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain!codemp = txtLlave(0).Text
         End If

        'Datos.
         .uorstMain!RazEmp = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
         .uorstMain!direccion = IIf(txtDato(1).Text = "", Null, txtDato(1).Text)
         .uorstMain!localidademp = IIf(txtDato(2).Text = "", Null, txtDato(2).Text)
         .uorstMain!RUCEmp = IIf(txtDato(3).Text = "", Null, txtDato(3).Text)
         .uorstMain!actividademp = IIf(txtDato(4).Text = "", Null, txtDato(4).Text)
         .uorstMain!repapepaterno = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
         .uorstMain!repapematerno = IIf(txtDato(6).Text = "", Null, txtDato(6).Text)
         .uorstMain!repnombre = IIf(txtDato(7).Text = "", Null, txtDato(7).Text)
         .uorstMain!repdocumento = IIf(txtDato(8).Text = "", Null, txtDato(8).Text)
         .uorstMain!conapepaterno = IIf(txtDato(9).Text = "", Null, txtDato(9).Text)
         .uorstMain!conapematerno = IIf(txtDato(10).Text = "", Null, txtDato(10).Text)
         .uorstMain!connombre = IIf(txtDato(11).Text = "", Null, txtDato(11).Text)
         .uorstMain!condocumento = IIf(txtDato(12).Text = "", Null, txtDato(12).Text)
         .uorstMain!BuenContri = IIf(chkBuenContri.Value = vbChecked, ESTDBUEN_CONTRI_ACT, ESTDBUEN_CONTRI_INA) '2015-08-27 ctr obligac sunat
         .uorstMain!EstEmp = IIf(chkEstEmp.Value = vbChecked, ESTEMPR_ACT, ESTEMPR_INA) '2016-06-02 adicion campo EstAct en Empresa
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain!codemp
      
        'Datos.
         txtDato(0).Text = IIf(IsNull(.uorstMain!RazEmp), "", .uorstMain!RazEmp)
         txtDato(1).Text = IIf(IsNull(.uorstMain!direccion), "", .uorstMain!direccion)
         txtDato(2).Text = IIf(IsNull(.uorstMain!localidademp), "", .uorstMain!localidademp)
         txtDato(3).Text = IIf(IsNull(.uorstMain!RUCEmp), "", .uorstMain!RUCEmp)
         txtDato(4).Text = IIf(IsNull(.uorstMain!actividademp), "", .uorstMain!actividademp)
         txtDato(5).Text = IIf(IsNull(.uorstMain!repapepaterno), "", .uorstMain!repapepaterno)
         txtDato(6).Text = IIf(IsNull(.uorstMain!repapematerno), "", .uorstMain!repapematerno)
         txtDato(7).Text = IIf(IsNull(.uorstMain!repnombre), "", .uorstMain!repnombre)
         txtDato(8).Text = IIf(IsNull(.uorstMain!repdocumento), "", .uorstMain!repdocumento)
         txtDato(9).Text = IIf(IsNull(.uorstMain!conapepaterno), "", .uorstMain!conapepaterno)
         txtDato(10).Text = IIf(IsNull(.uorstMain!conapematerno), "", .uorstMain!conapematerno)
         txtDato(11).Text = IIf(IsNull(.uorstMain!connombre), "", .uorstMain!connombre)
         txtDato(12).Text = IIf(IsNull(.uorstMain!condocumento), "", .uorstMain!condocumento)
         
         chkBuenContri.Value = IIf(.uorstMain!BuenContri = ESTDBUEN_CONTRI_ACT, vbChecked, vbUnchecked) '2015-08-27 ctr obligac sunat
         chkEstEmp.Value = IIf(.uorstMain!EstEmp = ESTEMPR_ACT, vbChecked, vbUnchecked) '2016-06-02 adicion campo EstAct en Empresa
'ini 2015-01-07 adiciono imagen empresa
'''ReadImagen .uorstMain, imgLogo, "logoemp"
'fin 2015-01-07 adiciono imagen empresa

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
   '2016-06-02 adicion campo EstAct en Empresa chkBuenContri.Value = vbChecked
   chkBuenContri.Value = vbUnchecked
   chkEstEmp.Value = vbChecked '2016-06-02 adicion campo EstAct en Empresa
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
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
   chkBuenContri.Enabled = tbHabilitar
   chkEstEmp.Enabled = tbHabilitar '2016-06-02 adicion campo EstAct en Empresa
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With

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
