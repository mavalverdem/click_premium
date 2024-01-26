VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTPdo 
   Caption         =   "[Título]"
   ClientHeight    =   6480
   ClientLeft      =   900
   ClientTop       =   1650
   ClientWidth     =   8385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   8385
   Begin VB.ComboBox cboCalcularIGV 
      Height          =   315
      ItemData        =   "frmTPdo.frx":0000
      Left            =   6600
      List            =   "frmTPdo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Frame fraCuadro 
      Height          =   1020
      Index           =   1
      Left            =   60
      TabIndex        =   34
      Top             =   4575
      Width           =   8220
      Begin VB.CommandButton cmdProducto 
         Caption         =   "Producto"
         Height          =   315
         Left            =   6975
         TabIndex        =   51
         Top             =   360
         Width           =   1110
      End
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   10
         Left            =   1110
         TabIndex        =   39
         Top             =   600
         Width           =   1440
      End
      Begin VB.CommandButton cmdMasProducto 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Picture         =   "frmTPdo.frx":0004
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   9
         Left            =   1110
         TabIndex        =   36
         Top             =   255
         Width           =   680
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   9
         Left            =   6615
         Picture         =   "frmTPdo.frx":0106
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   255
         Width           =   255
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
         Index           =   10
         Left            =   2535
         TabIndex        =   40
         Top             =   600
         Width           =   4095
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
         Index           =   9
         Left            =   1770
         TabIndex        =   37
         Top             =   255
         Width           =   4875
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cen. Costo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   35
         Top             =   255
         Width           =   890
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   38
         Top             =   615
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdFormato 
      Cancel          =   -1  'True
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5805
      Picture         =   "frmTPdo.frx":02B0
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5745
      Width           =   720
   End
   Begin VB.TextBox txtLlave 
      Height          =   280
      Index           =   2
      Left            =   810
      TabIndex        =   6
      Top             =   465
      Width           =   1920
   End
   Begin VB.PictureBox picextension 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5760
      Picture         =   "frmTPdo.frx":0886
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   57
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkExtension 
      Alignment       =   1  'Right Justify
      Caption         =   "Extensión"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1050
   End
   Begin VB.Frame fraCuadro 
      Height          =   1050
      Index           =   0
      Left            =   60
      TabIndex        =   26
      Top             =   3450
      Width           =   8220
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   8
         Left            =   7875
         Picture         =   "frmTPdo.frx":09D0
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   8
         Left            =   4425
         TabIndex        =   32
         Top             =   540
         Width           =   680
      End
      Begin VB.CommandButton cmdMas 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         Picture         =   "frmTPdo.frx":0B7A
         TabIndex        =   27
         Top             =   540
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   7
         Left            =   4080
         Picture         =   "frmTPdo.frx":0C7C
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   7
         Left            =   375
         TabIndex        =   29
         Top             =   540
         Width           =   980
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   12
         Left            =   375
         TabIndex        =   28
         Top             =   210
         Width           =   1800
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   13
         Left            =   4425
         TabIndex        =   31
         Top             =   210
         Width           =   1440
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
         Index           =   8
         Left            =   5085
         TabIndex        =   33
         Top             =   540
         Width           =   2835
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
         Height          =   280
         Index           =   7
         Left            =   1335
         TabIndex        =   30
         Top             =   540
         Width           =   2745
      End
   End
   Begin VB.TextBox txtLlave 
      ForeColor       =   &H80000012&
      Height          =   280
      Index           =   0
      Left            =   810
      TabIndex        =   1
      Top             =   120
      Width           =   520
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   5010
      Picture         =   "frmTPdo.frx":0E26
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   120
      Width           =   280
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   6
      Left            =   4800
      TabIndex        =   25
      Top             =   3135
      Width           =   1690
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   2355
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5685
      Width           =   3480
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
         Picture         =   "frmTPdo.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   47
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
         Picture         =   "frmTPdo.frx":111A
         Style           =   1  'Graphical
         TabIndex        =   46
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
         Picture         =   "frmTPdo.frx":121C
         Style           =   1  'Graphical
         TabIndex        =   45
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
         Picture         =   "frmTPdo.frx":131E
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Picture         =   "frmTPdo.frx":1468
         Style           =   1  'Graphical
         TabIndex        =   43
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
         Picture         =   "frmTPdo.frx":1612
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   5
      Left            =   2685
      TabIndex        =   23
      Top             =   3135
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   4
      Left            =   600
      TabIndex        =   21
      Top             =   3135
      Width           =   1690
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   3
      Left            =   2880
      TabIndex        =   19
      Top             =   2475
      Width           =   735
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   2
      Left            =   1080
      TabIndex        =   15
      Top             =   2100
      Width           =   7050
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   1740
      Width           =   7050
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   990
      Width           =   1280
   End
   Begin VB.TextBox txtLlave 
      Height          =   280
      Index           =   1
      Left            =   7035
      TabIndex        =   5
      Top             =   360
      Width           =   1140
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "Proveedor"
      Height          =   315
      Left            =   7035
      TabIndex        =   50
      Top             =   1380
      Width           =   1110
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   7880
      Picture         =   "frmTPdo.frx":17BC
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   990
      Width           =   255
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTPdo.frx":1966
      Left            =   1080
      List            =   "frmTPdo.frx":1968
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2475
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Left            =   1080
      TabIndex        =   11
      Top             =   1350
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   63635457
      CurrentDate     =   37102
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Interno :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   58
      Top             =   510
      Width           =   585
   End
   Begin VB.Shape shpCuadro 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   675
      Left            =   5640
      Top             =   75
      Width           =   2670
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Proyecto :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   165
      Width           =   735
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
      Height          =   285
      Index           =   0
      Left            =   1335
      TabIndex        =   2
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe Diferencial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   11
      Left            =   4830
      TabIndex        =   24
      Top             =   2895
      Width           =   1335
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe ME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   10
      Left            =   2715
      TabIndex        =   22
      Top             =   2895
      Width           =   780
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe MN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   9
      Left            =   615
      TabIndex        =   20
      Top             =   2895
      Width           =   795
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Traducción :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   6
      Left            =   60
      TabIndex        =   14
      Top             =   2145
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   4
      Left            =   60
      TabIndex        =   10
      Top             =   1395
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8300
      Y1              =   840
      Y2              =   840
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
      Left            =   2340
      TabIndex        =   9
      Top             =   990
      Width           =   5535
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Glosa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   5
      Left            =   60
      TabIndex        =   12
      Top             =   1785
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Moneda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   7
      Left            =   60
      TabIndex        =   16
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "T.Cambio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   8
      Left            =   2040
      TabIndex        =   18
      Top             =   2505
      Width           =   705
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nro Pedido :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   7065
      TabIndex        =   4
      Top             =   90
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   1035
      Width           =   900
   End
End
Attribute VB_Name = "frmTPdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbCorregir As Boolean
Private pbValidada As Boolean
Private pbFecha As Boolean

Private pnCta_IndCCo As Integer
Private pnCta_TpoTcb As String
Private pcCodCCo_Def As String
Public psCodCCo_Pdo As String

Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
Private Sub cmdFormato_Click()
  Dim sSQL As String
  Dim sImporteLetras As String, sSignoMoneda As String
  Dim nImporteTotal As Double, nImporteIgv As Double
  Dim nFormato As Integer, nRegistro As Integer, nContador As Integer
  Dim nDiferencia As Integer, nLen As Integer
  Dim porstRegistro As New ADODB.Recordset
  ' xx = IIf(chkCalcularIGV.Value = Val(CODPDO_IGV), "I.G.V. (" & Str(gnPctIGV) & "%) ", _
    "Retención (" & Str(gnPctIR4) & "%)")

  ' Inicializo las variables de impresion
  'nImporteIgv = Round((CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) * gnPctIGV) / 100, 2)
  'nImporteTotal = Round(CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) + nImporteIgv, 2)
  'ini 2014-05-22 pdo c/igv
  ' 2014-07-18 error igv if chkCalcularIGV.Value = Val(CODPDO_IGV) Then
  If cboCalcularIGV.ListIndex = Val(CODPDO_IGV) Or cboCalcularIGV.ListIndex = Val(CODPDO_IGVG) Then
    nImporteIgv = Round((CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) * gnPctIGV) / 100, 2)
    nImporteTotal = Round(CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) + nImporteIgv, 2)
  Else
    nImporteIgv = 0
    'ini 2014-07-02 limite calcul reten S/.1500 o conversion de dolares > 1500
    Dim xLimite As Double
    If cboTpoMon.ListIndex + 1 = 1 Then
        'xLimite = CDec(TxtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text)
        xLimite = CDec(txtDato(4).Text)
   Else
        'xLimite = CDec(TxtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text)
        xLimite = CDec(txtDato(5).Text)
        'xLimite = fDiv0(xLimite, CDec(TxtDato(3).Text))
        xLimite = xLimite * CDec(txtDato(3).Text)
    End If
    'xLimite
    'fin 2014-07-02 limite calcul reten S/.1500 o conversion de dolares > 1500
    
    'retencion 4ta categoria
    'If CDec(TxtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) > Choose(cboTpoMon.ListIndex + 1, 1500, 3000) Then
    If xLimite > 1500 Then
       nImporteIgv = Round((CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) * gnPctIR4) / 100, 2)
    End If
    nImporteTotal = Round(CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) - nImporteIgv, 2)
  End If
  'fin 2014-05-22 pdo c/igv
  '2014-06-02 se debe restar si es <> igv nImporteTotal = Round(CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 4, 5)).Text) + nImporteIgv, 2)
  sImporteLetras = "SON : " & gfNumLet(nImporteTotal, Choose(cboTpoMon.ListIndex + 1, "N", "E"))
  sSignoMoneda = gfEnmasc(IIf(IsNull(frmTPdoGrd.uorstMain!UsrMdf), frmTPdoGrd.uorstMain!UsrCre, frmTPdoGrd.uorstMain!UsrMdf))
  
  sSQL = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.coddpe, det.pdocpr)", "(det.coddpe+det.pdocpr)") & " AS documento, det.codprod AS codcta, "
  sSQL = sSQL & "det.coddpe, det.pdocpr, cpr.fehpdo AS emision, cpr.nrointerno, "
  sSQL = sSQL & "cpr.codaux, aux.razaux, aux.diraux, aux.rucaux, aux.email, cpr.tpomon, cfg.pctigv, "
  sSQL = sSQL & "prod." & Choose(gsIdioma, "detprod", "detprodx") & " AS glocta, "
  sSQL = sSQL & "cpr." & Choose(gsIdioma, "detpdo", "detpdox") & " AS detpdo, prod.unimed, det.cantiprod, "
  
  sSQL = sSQL & "(CASE cpr.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impouni_mn ELSE det.impouni_me END) AS impunitario, "
  sSQL = sSQL & "(CASE cpr.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impprod_mn ELSE det.impprod_me END) AS impbase, "
  
  'sSQL = sSQL & "20 AS impunitario, "
  'sSQL = sSQL & "30 AS impbase, "
  
  sSQL = sSQL & nImporteIgv & " AS impigv, " & nImporteTotal & " AS imptotal, "
  sSQL = sSQL & "'" & sImporteLetras & "' AS importeletra, "
  sSQL = sSQL & "(CASE cpr.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS signomon, "
  'sSQL = sSQL & "'x' AS signomon, "
  sSQL = sSQL & "dpe." & Choose(gsIdioma, "detdpe", "detdpex") & " AS detdpe, "
  sSQL = sSQL & "'" & sSignoMoneda & "' AS selaborado "
  sSQL = sSQL & "FROM copdocpr cpr "
  sSQL = sSQL & "INNER JOIN copdocprprod det ON cpr.codemp=det.codemp AND cpr.pdoano=det.pdoano AND cpr.mespvs=det.mespvs AND cpr.coddpe=det.coddpe AND cpr.pdocpr=det.pdocpr "
  sSQL = sSQL & "INNER JOIN tgaux aux ON cpr.codemp=aux.codemp AND cpr.codaux=aux.codaux "
  sSQL = sSQL & "INNER JOIN tgcfg cfg ON cpr.codemp=cfg.codemp AND cpr.pdoano=cfg.pdoano "
  sSQL = sSQL & "INNER JOIN codpe dpe ON dpe.codemp=cpr.codemp AND dpe.coddpe=cpr.coddpe "
  sSQL = sSQL & "LEFT JOIN cocprprod prod ON prod.codemp=det.codemp AND prod.pdoano=det.pdoano AND prod.codprod=det.codprod "
  sSQL = sSQL & "WHERE cpr.codemp='" & gsCodEmp & "' "
  sSQL = sSQL & "AND cpr.pdoano='" & gsAnoAct & "' "
  sSQL = sSQL & "AND cpr.mespvs='" & gsMesAct & "' "
  sSQL = sSQL & "AND cpr.coddpe='" & txtllave(0).Text & "' "
  sSQL = sSQL & "AND cpr.pdocpr='" & txtllave(1).Text & "' "
  sSQL = sSQL & "ORDER BY det.codprod, det.codcco"
  With porstRegistro
    If .State = adStateOpen Then .Close
    .ActiveConnection = frmTPdoGrd.uocnnMain
    .CursorLocation = adUseClient
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = sSQL
    .Open
  End With
  ' Verifico tiene detalle producto
  If porstRegistro.RecordCount = 0 Then
    sSQL = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.coddpe, det.pdocpr)", "(det.coddpe+det.pdocpr)") & " AS documento, det.codcta, "
    sSQL = sSQL & "det.coddpe, det.pdocpr, cpr.fehpdo AS emision, cpr.nrointerno, "
    sSQL = sSQL & "cpr.codaux, aux.razaux, aux.diraux, aux.rucaux, aux.email, cpr.tpomon, cfg.pctigv, "
    sSQL = sSQL & "cta." & Choose(gsIdioma, "detcta", "detctax") & " AS glocta, "
    sSQL = sSQL & "cpr." & Choose(gsIdioma, "detpdo", "detpdox") & " AS detpdo, Null AS unimed, 0 AS cantiprod,0.00 AS impunitario, "
    sSQL = sSQL & "(CASE cpr.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impcta_mn ELSE det.impcta_me END) AS impbase, "
    sSQL = sSQL & nImporteIgv & " AS impigv, " & nImporteTotal & " AS imptotal, "
    sSQL = sSQL & "'" & sImporteLetras & "' AS importeletra, "
    sSQL = sSQL & "(CASE cpr.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS signomon, "
    sSQL = sSQL & "dpe." & Choose(gsIdioma, "detdpe", "detdpex") & " AS detdpe, "
    sSQL = sSQL & "'" & sSignoMoneda & "' AS selaborado "
    sSQL = sSQL & "FROM copdocpr cpr "
    sSQL = sSQL & "LEFT JOIN copdocprcta det ON cpr.codemp=det.codemp AND cpr.pdoano=det.pdoano AND cpr.mespvs=det.mespvs AND cpr.coddpe=det.coddpe AND cpr.pdocpr=det.pdocpr "
    sSQL = sSQL & "INNER JOIN tgaux aux ON cpr.codemp=aux.codemp AND cpr.codaux=aux.codaux "
    sSQL = sSQL & "INNER JOIN tgcfg cfg ON cpr.codemp=cfg.codemp AND cpr.pdoano=cfg.pdoano "
    sSQL = sSQL & "INNER JOIN codpe dpe ON dpe.codemp=cpr.codemp AND dpe.coddpe=cpr.coddpe "
    sSQL = sSQL & "LEFT JOIN cocta cta ON cta.codemp=det.codemp AND cta.pdoano=det.pdoano AND cta.codcta=det.codcta "
    sSQL = sSQL & "WHERE cpr.codemp='" & gsCodEmp & "' "
    sSQL = sSQL & "AND cpr.pdoano='" & gsAnoAct & "' "
    sSQL = sSQL & "AND cpr.mespvs='" & gsMesAct & "' "
    sSQL = sSQL & "AND cpr.coddpe='" & txtllave(0).Text & "' "
    sSQL = sSQL & "AND cpr.pdocpr='" & txtllave(1).Text & "' "
    sSQL = sSQL & "ORDER BY det.codcta, det.codcco"
    With porstRegistro
      If .State = adStateOpen Then .Close
      .ActiveConnection = frmTPdoGrd.uocnnMain
      .CursorLocation = adUseClient
      .CursorType = adOpenForwardOnly
      .LockType = adLockReadOnly
      .Source = sSQL
      .Open
    End With
  End If
  
  ' Verifico si se puede imprimir
  If porstRegistro.RecordCount = 0 Then MsgBox Choose(gsIdioma, "El documento no tiene detalle de impresión", "The document does not have impression detail"), vbCritical: Exit Sub
  nRegistro = CInt(porstRegistro.RecordCount)
  
  ' Genero la tabla temporal de reporte
  If ps_Plataforma = pSrvMySql Then
    frmTPdoGrd.uocnnMain.Execute "DROP TABLE IF EXISTS trptdoccompra"
    sSQL = "CREATE TEMPORARY TABLE IF NOT EXISTS trptdoccompra (documento varchar(8) Not Null, "
    sSQL = sSQL & "secuencia smallint(2) Default '0', codcta varchar(20) Null,"
    sSQL = sSQL & "coddpe char(4) Not Null, pdocpr varchar(8) Not Null, "
    sSQL = sSQL & "emision date Null, referencia varchar(15) Null, "
    sSQL = sSQL & "detdpe varchar(40) Null, "
    sSQL = sSQL & "codaux varchar(11) Null, razaux varchar(80) Null, "
    sSQL = sSQL & "diraux varchar(80) Null, rucaux varchar(11) Null, "
    sSQL = sSQL & "email varchar(40) Null, "
    sSQL = sSQL & "tpomon char(1) Null, signomon char(3) Null, "
    sSQL = sSQL & "pctigv decimal(4,2) Default '0', "
    sSQL = sSQL & "glocta varchar(60) Null, "
    '2014-05-27 detped a 100c sSQL = sSQL & "detpdo varchar(50) Null, "
    sSQL = sSQL & "detpdo varchar(100) Null, "
    sSQL = sSQL & "unimed char(3) Null, "
    sSQL = sSQL & "cantiprod decimal(7,2) Default '0', impunitario decimal(12,2) Default '0', "
    sSQL = sSQL & "impbase decimal(12,2) Default '0', impigv decimal(12,2) Default '0', "
    sSQL = sSQL & "imptotal decimal(12,2) Default '0', importeletra varchar(250) Null, "
    sSQL = sSQL & "forimp smallint(1) Default '0', selaborado varchar(10) Null, "
    sSQL = sSQL & "PRIMARY KEY (documento, secuencia))"
  ElseIf ps_Plataforma = pSrvSql Then
    frmTPdoGrd.uocnnMain.Execute "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#trptdoccompra') DROP TABLE #trptdoccompra"
    sSQL = "CREATE TABLE #trptdoccompra (documento varchar(8) Not Null, "
    sSQL = sSQL & "secuencia smallint default '0', codcta varchar(20) Null,"
    sSQL = sSQL & "coddpe char(4) Not Null, pdocpr varchar(8) Not Null, "
    sSQL = sSQL & "emision smalldatetime Null, referencia varchar(15) Null, "
    sSQL = sSQL & "detdpe varchar(40) Null, "
    sSQL = sSQL & "codaux varchar(11) Null, razaux varchar(80) Null, "
    sSQL = sSQL & "diraux varchar(80) Null, rucaux varchar(11) Null, "
    sSQL = sSQL & "email varchar(40) Null, "
    sSQL = sSQL & "tpomon char(1) Null, signomon char(3) Null, "
    sSQL = sSQL & "pctigv decimal(4,2) Default '0', "
    sSQL = sSQL & "glocta varchar(60) Null, "
    '2014-05-27 detped a 100c sSQL = sSQL & "detpdo varchar(50) NULL, "
    sSQL = sSQL & "detpdo varchar(100) NULL, "
    sSQL = sSQL & "unimed char(3) Null, "
    sSQL = sSQL & "cantiprod decimal(7,2) Default '0', impunitario decimal(12,2) Default '0', "
    sSQL = sSQL & "impbase decimal(12,2) Default '0', impigv decimal(12,2) Default '0', "
    sSQL = sSQL & "imptotal decimal(12,2) Default '0', importeletra varchar(250) Null, "
    sSQL = sSQL & "forimp smallint Default '0', selaborado varchar(10) Null, "
    sSQL = sSQL & "PRIMARY KEY (documento, secuencia))"
  End If
  frmTPdoGrd.uocnnMain.Execute sSQL
  
  nRegistro = 0: nContador = 0
  ' Genero la informació de impresión
  While Not porstRegistro.EOF
    'nDiferencia = ppNumeroLinea(IIf(IsNull(porstRegistro!glodet0), "", porstRegistro!glodet0) & IIf(IsNull(porstRegistro!glodet1), "", porstRegistro!glodet1))
    nDiferencia = ppNumeroLinea(IIf(IsNull(porstRegistro!detpdo), "", porstRegistro!detpdo))
    nContador = nContador + nDiferencia
    nRegistro = nRegistro + 1
    sSQL = "INSERT INTO " & ps_Prefijo & "trptdoccompra "
    sSQL = sSQL & "(documento, secuencia, coddpe, pdocpr, codcta, emision, referencia, detdpe, codaux, razaux, diraux, rucaux, email, tpomon, "
    sSQL = sSQL & "signomon, pctigv, glocta, detpdo, unimed, cantiprod, impunitario, impbase, impigv, imptotal, importeletra, forimp, selaborado) "
    sSQL = sSQL & "VALUES ('" & porstRegistro!documento & "', "
    sSQL = sSQL & nRegistro & ", "
    sSQL = sSQL & "'" & porstRegistro!coddpe & "', "
    sSQL = sSQL & "'" & porstRegistro!pdocpr & "', "
    sSQL = sSQL & "'" & porstRegistro!CodCta & "', "
    If ps_Plataforma = pSrvMySql Then
      sSQL = sSQL & "DATE_FORMAT('" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
    Else
      sSQL = sSQL & "CONVERT(smalldatetime, '" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', 120), "
    End If
    sSQL = sSQL & IIf(IsNull(porstRegistro!nrointerno), "Null", "'" & porstRegistro!nrointerno & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!detdpe), "Null", "'" & porstRegistro!detdpe & "'") & ", "
    sSQL = sSQL & "'" & porstRegistro!codaux & "', "
    sSQL = sSQL & "'" & porstRegistro!razAux & "', "
    sSQL = sSQL & IIf(IsNull(porstRegistro!DirAux), "Null", "'" & porstRegistro!DirAux & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!rucaux), "Null", "'" & porstRegistro!rucaux & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!email), "Null", "'" & porstRegistro!email & "'") & ", "
    sSQL = sSQL & "'" & porstRegistro!tpomon & "', "
    sSQL = sSQL & "'" & porstRegistro!signomon & "', "
    sSQL = sSQL & CDec(porstRegistro!PctIGV) & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glocta), "Null", "'" & porstRegistro!glocta & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!detpdo), "Null", "'" & porstRegistro!detpdo & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!unimed), "Null", "'" & porstRegistro!unimed & "'") & ", "
    sSQL = sSQL & CDec(porstRegistro!cantiprod) & ", "
    sSQL = sSQL & CDec(porstRegistro!impunitario) & ", "
    sSQL = sSQL & CDec(porstRegistro!impbase) & ", "
    sSQL = sSQL & CDec(porstRegistro!impigv) & ", "
    sSQL = sSQL & CDec(porstRegistro!imptotal) & ", "
    sSQL = sSQL & "'" & sImporteLetras & "', "
    sSQL = sSQL & "'" & nFormato & "', "
    sSQL = sSQL & IIf(IsNull(porstRegistro!selaborado), "Null", "'" & porstRegistro!selaborado & "'") & ")"
    frmTPdoGrd.uocnnMain.Execute sSQL
    porstRegistro.MoveNext
  Wend
  porstRegistro.MovePrevious
  
  ' Inserto los detalles adicionales
  nRegistro = nContador + 1
  For nContador = nRegistro To 7
    sSQL = "INSERT INTO " & ps_Prefijo & "trptdoccompra "
    sSQL = sSQL & "(documento, secuencia, coddpe, pdocpr, codcta, emision, referencia, detdpe, codaux, razaux, diraux, rucaux, email, tpomon, "
    sSQL = sSQL & "signomon, pctigv, glocta, detpdo, unimed, cantiprod, impunitario, impbase, impigv, imptotal, importeletra, forimp, selaborado) "
    sSQL = sSQL & "VALUES ('" & porstRegistro!documento & "', "
    sSQL = sSQL & nContador & ", "
    sSQL = sSQL & "'" & porstRegistro!coddpe & "', "
    sSQL = sSQL & "'" & porstRegistro!pdocpr & "', "
    sSQL = sSQL & "Null, "
    If ps_Plataforma = pSrvMySql Then
      sSQL = sSQL & "DATE_FORMAT('" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
    Else
      sSQL = sSQL & "CONVERT(smalldatetime, '" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', 120), "
    End If
    sSQL = sSQL & IIf(IsNull(porstRegistro!nrointerno), "Null", "'" & porstRegistro!nrointerno & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!detdpe), "Null", "'" & porstRegistro!detdpe & "'") & ", "
    sSQL = sSQL & "'" & porstRegistro!codaux & "', "
    sSQL = sSQL & "'" & porstRegistro!razAux & "', "
    sSQL = sSQL & IIf(IsNull(porstRegistro!DirAux), "Null", "'" & porstRegistro!DirAux & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!rucaux), "Null", "'" & porstRegistro!rucaux & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!email), "Null", "'" & porstRegistro!email & "'") & ", "
    sSQL = sSQL & "'" & porstRegistro!tpomon & "', "
    sSQL = sSQL & "'" & porstRegistro!signomon & "', "
    sSQL = sSQL & CDec(porstRegistro!PctIGV) & ", "
    sSQL = sSQL & "Null, "
    sSQL = sSQL & IIf(IsNull(porstRegistro!detpdo), "Null", "'" & porstRegistro!detpdo & "'") & ", "
    sSQL = sSQL & "Null, 0, 0, 0, "
    sSQL = sSQL & CDec(porstRegistro!impigv) & ", "
    sSQL = sSQL & CDec(porstRegistro!imptotal) & ", "
    sSQL = sSQL & "'" & sImporteLetras & "', "
    sSQL = sSQL & "'" & nFormato & "', "
    sSQL = sSQL & IIf(IsNull(porstRegistro!selaborado), "Null", "'" & porstRegistro!selaborado & "'") & ")"
    frmTPdoGrd.uocnnMain.Execute sSQL
  Next nContador
  
  ' Obtengo los registrso de impresion
  With porstRegistro
    If .State = adStateOpen Then .Close
    .ActiveConnection = frmTPdoGrd.uocnnMain
    .CursorLocation = adUseClient
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = "SELECT * FROM " & ps_Prefijo & "trptdoccompra ORDER BY documento, secuencia"
    .Open
  End With
  ' Realizo la impresion
  gpEncabezadoRpt frmMain.rptMain, Me.Caption, Date, True, False, porstRegistro
  With frmMain.rptMain
    '[Datos y parámetros del reporte
'ini 2014-05-22 pdo c/igv
    Dim xx As String
    
'2014-07-18 error igv xx = IIf(chkCalcularIGV.Value = Val(CODPDO_IGV), "I.G.V. (" & Str(gnPctIGV) & "%) ",
    xx = IIf(cboCalcularIGV.ListIndex = Val(CODPDO_IGV) Or cboCalcularIGV.ListIndex = Val(CODPDO_IGVG), "I.G.V. (" & Str(gnPctIGV) & "%) ", _
    "Retención (" & Str(gnPctIR4) & "%)")
    .Formulas(9) = "sTpoCalculo='" & xx & "'"
'fin 2014-05-22 pdo c/igv
    .ReportFileName = gsRutRpt & "rptdoccompra.rpt"
    .WindowState = crptMaximized
    .MarginLeft = 240
    .Destination = crptToWindow
    .Action = 1
  End With
  porstRegistro.Close
  Set porstRegistro = Nothing
  frmTPdoGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptdoccompra", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#trptdoccompra') DROP TABLE #trptdoccompra")

End Sub

Private Sub cmdMas_Click()
  frmTPdoMasGrd.Show vbModal
'  ppAbreCtaCCo
End Sub

Private Sub cmdMasProducto_Click()
  If txtDato(9).Text = "" Then MsgBox TEXT_6002, vbExclamation: txtDato(9).SetFocus: Exit Sub
  frmTPdoMasProdGrd.Show vbModal
End Sub
Private Sub cmdProducto_Click()
  frmMProdGrd.Show vbModal
  frmTPdoGrd.uorstCoCprProd.Requery
End Sub
']
Private Sub Form_Load()
  pbValidada = False
  pbFecha = True
  Me.KeyPreview = True
  
  With frmTPdoGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
    txtllave(0).MaxLength = .uorstMain!coddpe.DefinedSize
    txtllave(1).MaxLength = .uorstMain!pdocpr.DefinedSize
    txtllave(2).MaxLength = .uorstMain!nrointerno.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
    End With
    
    '2014-05-22 chkCalcularIGV.Value = 1
    'ini 2014-07-18 error igv
    'chkCalcularIGV.Caption = Choose(gsIdioma, "Calcular I.G.&V.", "Calculate GST")
    With cboCalcularIGV
      .AddItem CODPDO_HPR_TXT, CODPDO_HPR
      .AddItem CODPDO_IGV_TXT, CODPDO_IGV
      .AddItem CODPDO_IGVG_TXT, CODPDO_IGVG
    End With
    'fin 2014-07-18 error igv

    txtDato(0).MaxLength = .uorstMain!codaux.DefinedSize
    txtDato(gsIdioma).MaxLength = .uorstMain!detpdo.DefinedSize
    txtDato(3 - gsIdioma).MaxLength = .uorstMain!detpdox.DefinedSize
    txtDato(3).MaxLength = 7
    txtDato(4).MaxLength = 14
    txtDato(5).MaxLength = 14
    txtDato(6).MaxLength = 14
    txtDato(7).MaxLength = 8
    txtDato(8).MaxLength = 5
    txtDato(9).MaxLength = 5
    txtDato(10).MaxLength = 20
    txtllave(1).Enabled = False
    chkExtension.Enabled = False
  End With
   
  If pbNuevo Then
    cmdRetroceder.Enabled = False
    cmdAvanzar.Enabled = False
  End If
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  cmdFormato.Enabled = Not pbNuevo
  upHabilitacion False

  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(16, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Proyecto :", "Nro Pedido :", "Interno :", "Proveedor :", "Fecha :", "Glosa :", "Traducción :", "Moneda :", "T.Cambio:", "Importe MN :", "Importe ME :", "Importe Diferencial :", "Cuenta Contable :", "Centro de Costo :", "Cen. Costo :", "Producto :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Project :", "Nro Order :", "Internal :", "Supplier :", "Date :", "Gloss :", "Translation :", "Currency :", "R.Exchange :", "Amount NC :", "Amount FC :", "Amount Differential :", "Accountable Account :", "Cost Center :", "Cost Center :", "Product :")
  Next nElemento
  cmdAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  chkExtension.Caption = Choose(gsIdioma, "Extensión", "Extention")
  cmdProducto.Caption = Choose(gsIdioma, "Producto", "Product")
  cmdFormato.Caption = Choose(gsIdioma, "&Imprimir", "&Print")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']

  '[Propio del formulario.
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  ']
End Sub

Private Sub Form_Activate()
  
 '[Busca detalle de códigos.           'Cambiar (habilitar/deshabilitar).
  If txtDato(7).Text <> "" Then
    ppAyuDet AYUDAT, 7
    pnCta_IndCCo = frmTPdoGrd.uorstCoCta!indcco
    pnCta_TpoTcb = frmTPdoGrd.uorstCoCta!TpoTcb
    pnCta_TpoTcb = TPOTCB_VTA
    pcCodCCo_Def = IIf(IsNull(frmTPdoGrd.uorstCoCta!codcco_def), "", frmTPdoGrd.uorstCoCta!codcco_def)
    ' Actualiza los datos de centro de costo
    txtDato(8).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(7).Enabled)
    cmdDatoAyud(8).Enabled = (pnCta_IndCCo = INDCCO_ACT And txtDato(8).Enabled)
  End If
  If txtDato(0).Text <> "" Then ppAyuDet AYUDAT, 0
  If txtDato(7).Text <> "" Then ppAyuDet AYUDAT, 7
  If txtDato(8).Text <> "" Then ppAyuDet AYUDAT, 8
  If txtDato(9).Text <> "" Then ppAyuDet AYUDAT, 8
  If txtDato(10).Text <> "" Then ppAyuDet AYUDAT, 8
 ']
  
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If

 '[Propio del formulario.
  If Not pbNuevo Then
    dtpDato.Tag = dtpDato.Value
  End If
  txtDato(3).Tag = txtDato(3).Text
  If txtllave(0).Text <> "" Then Call ppAyuDet(AYULLA, 0)
 ']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not frmTPdoGrd.uorstMain.EOF Then
    If frmTPdoGrd.uorstMain.EditMode <> adEditNone Then frmTPdoGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmTPdoGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTPdoGrd.uorstMain_Grd.MoveFirst
   frmTPdoGrd.uorstMain_Grd.Find "cLlave='" & txtllave(0).Text & txtllave(1).Text & txtDato(0).Text & "'"
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmTPdoGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTPdoGrd.uorstMain_Grd.MoveFirst
   frmTPdoGrd.uorstMain_Grd.Find "cLlave='" & txtllave(0).Text & txtllave(1).Text & txtDato(0).Text & "'"
End Sub

Public Sub cmdCorregir_Click()
  'Verificación de Mes Cerrado.
  If gbCieCpr Then
    MsgBox TEXT_9016, vbCritical
    Exit Sub
  End If
  
  pbCorregir = True
  
  cmdRetroceder.Enabled = False
  cmdAvanzar.Enabled = False
  cmdCorregir.Enabled = False
  cmdFormato.Enabled = False
  cmdGrabar.Enabled = True
  cmdDeshacer.Enabled = True
  upHabilitacion True
  txtDato(0).Enabled = False
  cmdDatoAyud(0).Enabled = False
  ' Caracteristicas de cuentas
  txtDato(7).Enabled = (cmdMas.Tag <> INDMASCTA_MAS)
  cmdDatoAyud(7).Enabled = (cmdMas.Tag <> INDMASCTA_MAS)
  txtDato(8).Enabled = (txtDato(7).Text <> "" And cmdMas.Tag <> INDMASCTA_MAS And pnCta_IndCCo = INDCCO_ACT)
  cmdDatoAyud(8).Enabled = (txtDato(7).Text <> "" And cmdMas.Tag <> INDMASCTA_MAS And pnCta_IndCCo = INDCCO_ACT)
  cmdMas.Enabled = (txtDato(7).Text = "" Or cmdMas.Tag <> INDMASCTA_CTA)
  
  '[Dato con el foco al corregir.       'Cambiar.
  txtllave(2).SetFocus
  ']
  txtDato(4).Enabled = False
  txtDato(5).Enabled = False
  
End Sub

Public Sub cmdGrabar_Click()
  Dim sSentencia As String
  On Error GoTo Err
  
  If Len(Trim(txtDato(0).Text)) = 0 Then MsgBox TEXT_6002, vbExclamation: txtDato(0).SetFocus: Exit Sub
  If cboTpoMon.ListIndex = TPOMON_NAC_IND And CDec(txtDato(4).Text) = 0 Then
    MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Nacional.", "You Must enter the amount in National Currency."), vbInformation
    txtDato(4).SetFocus
    Exit Sub
  ElseIf cboTpoMon.ListIndex = TPOMON_EXT_IND And CDec(txtDato(5).Text) = 0 Then
    MsgBox Choose(gsIdioma, "Debe ingresar el importe en Moneda Extranjera.", "You Must enter the amount in Foreign Currency."), vbInformation
    txtDato(5).SetFocus
    Exit Sub
  End If
   
  If CDec(txtDato(6).Text) = 0 Then
    MsgBox Choose(gsIdioma, "Debe ingresar el importe Diferencial.", "You Must enter the amount Differential."), vbInformation
    txtDato(6).SetFocus
    Exit Sub
  End If
  If cmdMas.Tag <> INDMASCTA_MAS And Len(Trim(txtDato(7).Text)) = 0 Then MsgBox Choose(gsIdioma, "Debe ingresar cuenta contable.", "You Must enter the accountable account."), vbExclamation: txtDato(7).SetFocus: Exit Sub
  If cmdMas.Tag = INDMASCTA_CTA And Len(Trim(txtDato(7).Text)) <> 0 And pnCta_IndCCo = INDCCO_ACT And Len(Trim(txtDato(8).Text)) = 0 Then MsgBox Choose(gsIdioma, "Debe ingresar centro de costos.", "You Must enter the cost center."), vbExclamation: txtDato(8).SetFocus: Exit Sub
  If cmdMas.Tag = INDMASCTA_MAS And Len(Trim(txtDato(7).Text)) = 0 Then MsgBox Choose(gsIdioma, "Debe ingresar cuenta contable.", "You Must enter the accountable account."), vbExclamation: cmdMas.SetFocus: Exit Sub
  ' valida presupuesto
  If Not pfValidoPresupuesto(txtllave(0).Text, txtllave(1).Text) Then Exit Sub
  
  With frmTPdoGrd                     'Cambiar Formulario de Grid.
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
    ' Elimino e inserto las productos
    .uocnnMain.Execute "DELETE FROM copdocprprod " & .usConnStrgWher_CoPdoCprProd
    sSentencia = "INSERT INTO copdocprprod SELECT * FROM " & ps_Prefijo & "tmpcopdocprprod ORDER BY codprod"
    .uocnnMain.Execute sSentencia
    ' Elimino e inserto las cuentas contables
    .uocnnMain.Execute "DELETE FROM copdocprcta " & .usConnStrgWher_CoPdoCprCta
    sSentencia = "INSERT INTO copdocprcta SELECT * FROM " & ps_Prefijo & "tmpcopdocprcta ORDER BY codcta, codcco"
    If cmdMas.Tag = INDMASCTA_CTA Then
      sSentencia = "INSERT INTO copdocprcta(codemp, pdoano, mespvs, coddpe, pdocpr, codcta, codcco, impcta_mn, impcta_me, impctadif, usrcre, fyhcre, usrmdf, fyhmdf) "
      sSentencia = sSentencia & "VALUES("
      sSentencia = sSentencia & "'" & gsCodEmp & "', "
      sSentencia = sSentencia & "'" & gsAnoAct & "', "
      sSentencia = sSentencia & "'" & gsMesAct & "', "
      sSentencia = sSentencia & "'" & txtllave(0).Text & "', "
      sSentencia = sSentencia & "'" & txtllave(1).Text & "', "
      sSentencia = sSentencia & IIf(txtDato(7).Text = "", "Null", "'" & txtDato(7).Text & "'") & ", "
      sSentencia = sSentencia & IIf(txtDato(8).Text = "", "Null", "'" & txtDato(8).Text & "'") & ", "
      sSentencia = sSentencia & CDec(txtDato(4).Text) & ", "
      sSentencia = sSentencia & CDec(txtDato(5).Text) & ", "
      sSentencia = sSentencia & CDec(txtDato(6).Text) & ", "
      sSentencia = sSentencia & "'" & gsAbvUsr & "', "
      sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "), "
      If pbNuevo Then
        sSentencia = sSentencia & "Null, Null)"
      Else
        sSentencia = sSentencia & "'" & gsAbvUsr & "', "
        sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & ")) "
      End If
    End If
    .uocnnMain.Execute sSentencia
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
    
    ' Refresco la grilla verificar
    .uorstMain_Grd.Requery
    .upDatosGrid
    '[Búsqueda de llave actual.     'Cambiar.
    .uorstMain_Grd.Find "cLlave='" & txtllave(0).Text & txtllave(1).Text & txtDato(0).Text & "'"
    ']
    If pbNuevo Then
      pbValidada = False
      cmdGrabar.Enabled = False
      upHabilitacion False
      txtllave(1).Enabled = False
      '[ No Pertenece al Formulario
      .uorstMain.Requery
      
      upDatosPredeterminados
      txtllave(0).Enabled = True
      '[Llave con el foco al añadir.  'Cambiar.
      txtllave(0).SetFocus
      ']
    Else
      cmdRetroceder.Enabled = True
      cmdAvanzar.Enabled = True
      cmdCorregir.Enabled = True
      cmdGrabar.Enabled = False
      cmdDeshacer.Enabled = False
      cmdFormato.Enabled = True
      upHabilitacion False
    End If
   End With
      
   Exit Sub
Err:
   gpErrores
  
   frmTPdoGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
   cmdFormato.Enabled = True
   gpTUe_Deshacer Me
End Sub

Public Sub cmdSalir_Click()
  If pbNuevo Or pbCorregir Then pbCorregir = False
  Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 7, 8
    If (pnCta_IndCCo = INDCCO_ACT And Index = 8) Or Index <> 8 Then
      txtDato(Index).SetFocus
    End If
   Case 9
    txtDato(Index).SetFocus
  End Select
  If (pnCta_IndCCo = INDCCO_ACT And Index = 8) Or Index <> 8 Then ppAyuBus AYUDAT, Index
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
  
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0
    txtllave(Index).SetFocus
  End Select
  ppAyuBus AYULLA, Index

End Sub

Private Sub picextension_Click()
  Dim respuesta As Long
  Dim sql As String
  
  If pbNuevo Then
    If chkExtension.Value = 0 Then Exit Sub
    If InStr(txtllave(1), "-") = 0 Then
    Else
      respuesta = MsgBox("Se Copiara datos del Pedido " & txtllave(0) & Left(txtllave(1), InStr(txtllave(1), "-") - 1) & "  al Pedido # " & txtllave(0) & txtllave(1) & " con Fecha " & dtpDato & " es correcto?", vbYesNo)
      If respuesta = vbYes Then
        With frmTPdoGrd
          sql = "INSERT INTO copdocpr (codemp,pdoano,mespvs,coddpe,pdocpr,indext,codaux,detpdo,detpdox,fehpdo,indcta,tpomon,imptcb,impmn,impme,impdife,usrcre,fyhcre,usrmdf,fyhmdf) "
          sql = sql & " SELECT codemp,'" & gsAnoAct & "' as pdoano,'" & gsMesAct & "' as mespvs,coddpe,'" & txtllave(1) & "' as pdocpr,indext,codaux,detpdo,detpdox,'" & Format(dtpDato, s_FmtFeHoMysql_0) & "',indcta,tpomon,imptcb,impmn,impme,impdife,'" & gsAbvUsr & "' as usrcre,'" & Format(Now, s_FmtFeHoMysql_0) & "' as fyhcre,null as usrmdf,null as fyhmdf "
          sql = sql & " FROM copdocpr "
          sql = sql & " WHERE codemp='" & gsCodEmp & "' and coddpe='" & txtllave(0) & "' and (pdocpr='" & Mid(Left(txtllave(1), InStr(txtllave(1), "-") - 1), 2, 10) & "' or pdocpr='" & Left(txtllave(1), InStr(txtllave(1), "-") - 1) & "')"
          On Error GoTo error
          Set .porstCancel = .uocnnMain.Execute(sql)
        End With
        With frmTPdoGrd
          sql = "INSERT INTO copdocprcta (codemp,pdoano,mespvs,coddpe,pdocpr,codcta,codcco,impcta_mn,impcta_me,impctadif,usrcre,fyhcre,usrmdf,fyhmdf) "
          sql = sql & " SELECT codemp,'" & gsAnoAct & "' as pdoano,'" & gsMesAct & "' as mespvs,coddpe,'" & txtllave(1) & "' as pdocpr,codcta,codcco,impcta_mn,impcta_me,impctadif,'" & gsAbvUsr & "' as usrcre,'" & Format(Now, s_FmtFeHoMysql_0) & "' as fyhcre,null as usrmdf,null as fyhmdf "
          sql = sql & " FROM copdocprcta "
          sql = sql & " WHERE codemp='" & gsCodEmp & "' and coddpe='" & txtllave(0) & "' and (pdocpr='" & Mid(Left(txtllave(1), InStr(txtllave(1), "-") - 1), 2, 10) & "' or pdocpr='" & Left(txtllave(1), InStr(txtllave(1), "-") - 1) & "')"
          On Error GoTo error
          Set .porstCancel = .uocnnMain.Execute(sql)
        End With
        frmTPdoGrd.uorstMain_Grd.Requery
        frmTPdoGrd.upDatosGrid
        Unload Me
      Else
        Exit Sub
      End If
    End If
  End If
error:

End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
  txtllave(Index).SelStart = 0
  txtllave(Index).SelLength = txtllave(Index).MaxLength
End Sub

Private Sub txtLlave_KeyPress(Index As Integer, KeyAscii As Integer)
'''[ARREGLAR: Retrocede si Shift está presionado.
''   If Len(Trim(txtLlave(Index))) + 1 = txtLlave(Index).MaxLength Then
''      SendKeys "{TAB}"
''   End If
''']ARREGLAR.
 
 '[Convierte a mayúsculas.
'   If Index = 0 Then                   'Cambiar (añadir índices).
'      KeyAscii = Asc(UCase(Chr(KeyAscii)))
'   End If
 ']
End Sub
Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus AYULLA, Index
End Sub
Private Sub txtLlave_LostFocus(Index As Integer)
  If (pbValidada And Index <> 2) Then
    If Len(txtllave(0)) <> 4 Then
      txtllave(0).Enabled = True
      txtllave(0).SetFocus    'Cambiar.
      Exit Sub
    End If
    txtllave(0).Enabled = False
    cmdLlaveAyud(0).Enabled = False
    If txtllave(2).Enabled Then
      txtllave(2).SetFocus
    ElseIf dtpDato.Enabled Then
      dtpDato.SetFocus
    End If
  End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  Dim dvRegistro As Variant
  Dim sSentencia As String
  
  If Index = 2 Then Exit Sub
  '[Valida la llave.                    'Cambiar.
  Select Case Index
   Case 0
    Cancel = ppAyuDet(AYULLA, Index)
    If Cancel Then Exit Sub
    If Len(txtllave(0)) <> 4 Then txtllave(0).SetFocus: Exit Sub
    psCodCCo_Pdo = IIf(IsNull(frmTPdoGrd.uorstCoDPe!codcco), "", frmTPdoGrd.uorstCoDPe!codcco)
    If pbNuevo Then
      With frmTPdoGrd
        sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(CAST(pdocpr As decimal)), '0000') AS cNumMaxPdo "
        sSentencia = sSentencia & "FROM copdocpr "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND coddpe='" & txtllave(0).Text & "' "
        sSentencia = sSentencia & "AND indext='" & INDMASCTA_INI & "'"
        Set .porstCancel = .uocnnMain.Execute(sSentencia)
        txtllave(1).Text = gfCeros(.porstCancel!cNumMaxPdo, 4, 1, "0")
        .porstCancel.Close
      End With
    End If
   Case 1
    If Len(Trim(txtllave(Index).Text)) <> 0 And Len(Trim(txtllave(Index).Text)) < (txtllave(Index).MaxLength / 2) Then
       txtllave(Index).Text = gfCeros(txtllave(Index).Text, (txtllave(Index).MaxLength / 2), 0, "0")
    End If
    ' Genero correlativo de extensión
    If (pbNuevo And chkExtension.Value = vbChecked) Then
      With frmTPdoGrd
        sSentencia = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "IFNULL", "ISNULL") & "(MAX(pdocpr), '00000000') AS cNumMaxPdo "
        sSentencia = sSentencia & "FROM copdocpr "
        sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
        sSentencia = sSentencia & "AND coddpe='" & txtllave(0).Text & "' "
        sSentencia = sSentencia & "AND LEFT(pdocpr, " & Len(txtllave(1).Text) & ")='" & txtllave(1).Text & "'"
        Set .porstCancel = .uocnnMain.Execute(sSentencia)
        txtllave(1).Text = txtllave(1).Text & "-" & IIf(Len(.porstCancel!cNumMaxPdo) > 4, Val(Mid(.porstCancel!cNumMaxPdo, 6)) + 1, "1")
        .porstCancel.Close
      End With
    End If
  End Select
  ']
  
  ' Valido la llave
  If Len(Trim(txtllave(0).Text)) <> 0 And Len(Trim(txtllave(1).Text)) <> 0 Then
    With frmTPdoGrd                  'Cambiar Formulario de Grid.
      sSentencia = "SELECT mespvs FROM copdocpr "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND coddpe='" & txtllave(0).Text & "' "
      sSentencia = sSentencia & "AND pdocpr='" & txtllave(1).Text & "'"
      Set .porstCancel = .uocnnMain.Execute(sSentencia)
      If .porstCancel.RecordCount > 0 Then
        MsgBox TEXT_8007 & Chr(13) & Choose(gsIdioma, "(mes ", "(month ") & gfMesLet("01" & .porstCancel!mespvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
        Cancel = True
        Exit Sub
      End If
      .porstCancel.Close
    End With
    
    With frmTPdoGrd.uorstMain
      If Not (.BOF And .EOF) Then
        dvRegistro = .Bookmark
        .MoveFirst
        .Find "cLlave1='" & txtllave(0).Text & txtllave(1).Text & "'"
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
    chkExtension.Enabled = True
    txtllave(1).Enabled = True
    txtllave(2).Enabled = True
    pbValidada = True
    ppAbreCuentaPedido
    ppAbreProductoPedido
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
  If Index >= 4 And Index <= 6 Then
    If Val(txtDato(3).Text) = 0 Then
      txtDato(3).Text = Format(0, FORMATO_NUM_2)
      txtDato(3).SetFocus
      MsgBox TEXT_9015, vbExclamation
      Exit Sub
    End If
  End If
  txtDato(Index).SelStart = 0
  txtDato(Index).SelLength = txtDato(Index).MaxLength + IIf(Index >= 3 And Index <= 6, 1, 0)
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
  '[ARREGLAR: Retrocede si Shift está presionado.
  If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus AYUDAT, Index
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
   
  Select Case Index
   Case 3
    If Val(txtDato(Index).Text) > 0 Then
      txtDato(Index).Text = Format(Val(txtDato(Index).Text), FORMATO_NUM_2)
    End If
   Case 4, 5
    If CDec(txtDato(3).Text) <= 0 Then
      MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
      txtDato(3).SetFocus
      Exit Sub
    End If
    ' Convierto importe en cero
    If CDec(txtDato(Index).Text) = 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      If Index = 4 And cboTpoMon.ListIndex = TPOMON_EXT_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index + 1).Text) * CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 5 And cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDato(Index).Text = Format(Round(CDec(txtDato(Index - 1).Text) / CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    ElseIf CDec(txtDato(Index).Text) <> 0 Then
      If Index = 4 And cboTpoMon.ListIndex = TPOMON_NAC_IND And (txtDato(Index - 1).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index + 1).Text = Format(Round(CDec(txtDato(Index).Text) / CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      ElseIf Index = 5 And cboTpoMon.ListIndex = TPOMON_EXT_IND And (txtDato(Index - 1).Text = 0 Or CDec(txtDato(Index).Text) <> CDec(txtDato(Index).Tag)) Then
        txtDato(Index - 1).Text = Format(Round(CDec(txtDato(Index).Text) * CDec(txtDato(3).Text), 2), FORMATO_NUM_1)
      End If
    End If
   Case 7
    If (txtDato(Index).Text = "" And txtDato(Index).Tag <> txtDato(Index).Text) Then
      ' Inicializo y elimino las cuentas
      lblDatoDeta(Index).Caption = ""
      txtDato(Index + 1).Text = ""
      lblDatoDeta(Index + 1).Caption = ""
      frmTPdoGrd.uocnnMain.Execute "DELETE FROM " & ps_Prefijo & "tmpcopdocprcta"
      frmTPdoGrd.uorstCoDPeCta.Requery
      cmdMas.Tag = INDMASCTA_INI
    End If
    txtDato(Index).Tag = txtDato(Index).Text
    cmdMas.Tag = IIf(txtDato(Index).Text <> "" And cmdMas.Tag = INDMASCTA_INI, INDMASCTA_CTA, cmdMas.Tag)
    txtDato(Index).Enabled = (cmdMas.Tag <> INDMASCTA_MAS)
    cmdDatoAyud(Index).Enabled = (cmdMas.Tag <> INDMASCTA_MAS)
    txtDato(Index + 1).Enabled = (txtDato(7).Text <> "" And cmdMas.Tag <> INDMASCTA_MAS And pnCta_IndCCo = INDCCO_ACT)
    cmdDatoAyud(Index + 1).Enabled = (txtDato(7).Text <> "" And cmdMas.Tag <> INDMASCTA_MAS And pnCta_IndCCo = INDCCO_ACT)
    cmdMas.Enabled = (txtDato(Index).Text = "" Or cmdMas.Tag <> INDMASCTA_CTA)
   Case 10
    If (txtDato(Index).Text = "" And txtDato(Index).Tag <> txtDato(Index).Text) Then
      ' Inicializo y elimino las cuentas
      lblDatoDeta(Index).Caption = ""
      txtDato(Index + 1).Text = ""
      lblDatoDeta(Index + 1).Caption = ""
      frmTPdoGrd.uocnnMain.Execute "DELETE FROM " & ps_Prefijo & "tmpcopdocprcta"
      frmTPdoGrd.uorstCoDPeCta.Requery
      cmdMas.Tag = INDMASCTA_INI
    End If
    txtDato(Index).Tag = txtDato(Index).Text
    cmdMas.Tag = IIf(txtDato(Index).Text <> "" And cmdMas.Tag = INDMASCTA_INI, INDMASCTA_CTA, cmdMas.Tag)
    txtDato(Index).Enabled = (cmdMas.Tag <> INDMASCTA_MAS)
    cmdDatoAyud(Index).Enabled = (cmdMas.Tag <> INDMASCTA_MAS)
    txtDato(Index + 1).Enabled = (txtDato(7).Text <> "" And cmdMas.Tag <> INDMASCTA_MAS And pnCta_IndCCo = INDCCO_ACT)
    cmdDatoAyud(Index + 1).Enabled = (txtDato(7).Text <> "" And cmdMas.Tag <> INDMASCTA_MAS And pnCta_IndCCo = INDCCO_ACT)
    cmdMas.Enabled = (txtDato(Index).Text = "" Or cmdMas.Tag <> INDMASCTA_CTA)
   End Select

End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  
  'Completa con ceros a la izquierda.
  Select Case Index
   Case 0, 8, 9
    Cancel = ppAyuDet(AYUDAT, Index)
    If Cancel Then Exit Sub
   Case 7
    Cancel = ppAyuDet(AYUDAT, Index)
    If Cancel Then Exit Sub
    
    If lblDatoDeta(Index).Caption <> "" Then
      pnCta_TpoTcb = frmTPdoGrd.uorstCoCta!TpoTcb
      pnCta_TpoTcb = TPOTCB_VTA
      pnCta_IndCCo = frmTPdoGrd.uorstCoCta!indcco
      pcCodCCo_Def = IIf(IsNull(frmTPdoGrd.uorstCoCta!codcco_def), "", frmTPdoGrd.uorstCoCta!codcco_def)
      
      ' Actualizo los datos adicionales
      txtDato(8).Text = IIf(txtDato(8).Text = "", IIf(psCodCCo_Pdo = "", pcCodCCo_Def, psCodCCo_Pdo), txtDato(8).Text)
      txtDato(8).Text = IIf(pnCta_IndCCo = INDCCO_ACT, txtDato(8).Text, "")
      lblDatoDeta(8).Caption = IIf(pnCta_IndCCo = INDCCO_ACT, lblDatoDeta(8).Caption, "")
      txtDato(8).Enabled = (pnCta_IndCCo = INDCCO_ACT)
      cmdDatoAyud(8).Enabled = (pnCta_IndCCo = INDCCO_ACT)
    End If
    Case 3
      txtDato(Index).Text = Format(CDec(IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)), FORMATO_NUM_2)
    Case 4, 5, 6
      txtDato(Index).Text = Format(CDec(IIf(Not IsNumeric(txtDato(Index).Text), 0, txtDato(Index).Text)), FORMATO_NUM_1)
   End Select
   
   Exit Sub
Err:
   gpErrores
End Sub
Private Function pfValidoPresupuesto(ByVal sProyecto As String, ByVal sPedido As String) As Boolean
  Dim sMensaje As String, sSentencia As String
  Dim nImporteCpr As Double, nImportePre As Double
  Dim porstPspCpr As ADODB.Recordset
  
  Set porstPspCpr = New ADODB.Recordset
  With porstPspCpr
    .ActiveConnection = frmTPdoGrd.uocnnMain
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
  End With
  
  pfValidoPresupuesto = True
  ' informacion de validacion
  sSentencia = "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#tmpvalidarpre') DROP TABLE #tmpvalidarpre"
  frmTPdoGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpvalidarpre", sSentencia)
    
  sSentencia = IIf(ps_Plataforma = pSrvMySql, "CREATE TEMPORARY TABLE IF NOT EXISTS tmpvalidarpre ", "")
  sSentencia = sSentencia & "SELECT pdo.codcta, pdo.codcco, cta.tpomon, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impmn_" & gsMesAct & "), 2) AS imporpre_mn, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impme_" & gsMesAct & "), 2) AS imporpre_me, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_mn), 2) AS impopdo_mn, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_me), 2) AS impopdo_me "
  sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "", "INTO #tmpvalidarpre ")
  sSentencia = sSentencia & "FROM copdocprcta pdo "
  sSentencia = sSentencia & "INNER JOIN copsp psp ON psp.codemp=pdo.codemp AND psp.pdoano=pdo.pdoano AND psp.codcta=pdo.codcta AND psp.codcco=pdo.codcco "
  sSentencia = sSentencia & "INNER JOIN cocta cta ON cta.codemp=pdo.codemp AND cta.pdoano=pdo.pdoano AND cta.codcta=pdo.codcta AND cta.indcco='" & INDCCO_INA & "' "
  sSentencia = sSentencia & "WHERE pdo.codemp ='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdo.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND pdo.mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND pdo.coddpe<>'" & sProyecto & "' "
  sSentencia = sSentencia & "AND pdo.pdocpr<>'" & sPedido & "' "
  sSentencia = sSentencia & "AND IFNULL(pdo.codcco, '')='' "
  sSentencia = sSentencia & "GROUP BY pdo.codcta, pdo.codcco, cta.tpomon "
  sSentencia = sSentencia & "UNION ALL "
  sSentencia = sSentencia & "SELECT pdo.codcta, pdo.codcco, cta.tpomon, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impmn_" & gsMesAct & "), 2) AS imporpre_mn, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impme_" & gsMesAct & "), 2) AS imporpre_me, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_mn), 2) AS impopdo_mn, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_me), 2) AS impopdo_me "
  sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpcopdocprcta pdo "
  sSentencia = sSentencia & "INNER JOIN copsp psp ON psp.codemp=pdo.codemp AND psp.pdoano=pdo.pdoano AND psp.codcta=pdo.codcta AND psp.codcco=pdo.codcco "
  sSentencia = sSentencia & "INNER JOIN cocta cta ON cta.codemp=pdo.codemp AND cta.pdoano=pdo.pdoano AND cta.codcta=pdo.codcta AND cta.indcco='" & INDCCO_INA & "' "
  sSentencia = sSentencia & "WHERE pdo.codemp ='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdo.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND pdo.mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND pdo.coddpe='" & sProyecto & "' "
  sSentencia = sSentencia & "AND pdo.pdocpr='" & sPedido & "' "
  sSentencia = sSentencia & "AND IFNULL(pdo.codcco, '')='' "
  sSentencia = sSentencia & "GROUP BY pdo.codcta, pdo.codcco, cta.tpomon "
  sSentencia = sSentencia & "ORDER BY codcta, codcco"
  frmTPdoGrd.uocnnMain.Execute sSentencia
  
  sSentencia = "INSERT INTO " & ps_Prefijo & "tmpvalidarpre "
  sSentencia = sSentencia & "SELECT pdo.codcta, pdo.codcco, cta.tpomon, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impmn_" & gsMesAct & "), 2) AS imporpre_mn, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impme_" & gsMesAct & "), 2) AS imporpre_me, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_mn), 2) AS impopdo_mn, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_me), 2) AS impopdo_me "
  sSentencia = sSentencia & "FROM copdocprcta pdo "
  sSentencia = sSentencia & "INNER JOIN copsp psp ON psp.codemp=pdo.codemp AND psp.pdoano=pdo.pdoano AND psp.codcta=pdo.codcta AND psp.codcco=pdo.codcco "
  sSentencia = sSentencia & "INNER JOIN cocta cta ON cta.codemp=pdo.codemp AND cta.pdoano=pdo.pdoano AND cta.codcta=pdo.codcta AND cta.indcco='" & INDCCO_ACT & "' "
  sSentencia = sSentencia & "WHERE pdo.codemp ='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdo.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND pdo.mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND pdo.coddpe<>'" & sProyecto & "' "
  sSentencia = sSentencia & "AND pdo.pdocpr<>'" & sPedido & "' "
  sSentencia = sSentencia & "AND IFNULL(pdo.codcco, '')<>'' "
  sSentencia = sSentencia & "GROUP BY pdo.codcta, pdo.codcco, cta.tpomon "
  sSentencia = sSentencia & "UNION ALL "
  sSentencia = sSentencia & "SELECT pdo.codcta, pdo.codcco, cta.tpomon, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impmn_" & gsMesAct & "), 2) AS imporpre_mn, "
  sSentencia = sSentencia & "ROUND(AVG(psp.impme_" & gsMesAct & "), 2) AS imporpre_me, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_mn), 2) AS impopdo_mn, "
  sSentencia = sSentencia & "ROUND(SUM(pdo.impcta_me), 2) AS impopdo_me "
  sSentencia = sSentencia & "FROM " & ps_Prefijo & "tmpcopdocprcta pdo "
  sSentencia = sSentencia & "INNER JOIN copsp psp ON psp.codemp=pdo.codemp AND psp.pdoano=pdo.pdoano AND psp.codcta=pdo.codcta AND psp.codcco=pdo.codcco "
  sSentencia = sSentencia & "INNER JOIN cocta cta ON cta.codemp=pdo.codemp AND cta.pdoano=pdo.pdoano AND cta.codcta=pdo.codcta AND cta.indcco='" & INDCCO_ACT & "' "
  sSentencia = sSentencia & "WHERE pdo.codemp ='" & gsCodEmp & "' "
  sSentencia = sSentencia & "AND pdo.pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & "AND pdo.mespvs='" & gsMesAct & "' "
  sSentencia = sSentencia & "AND pdo.coddpe='" & sProyecto & "' "
  sSentencia = sSentencia & "AND pdo.pdocpr='" & sPedido & "' "
  sSentencia = sSentencia & "AND IFNULL(pdo.codcco, '')<>'' "
  sSentencia = sSentencia & "GROUP BY pdo.codcta, pdo.codcco, cta.tpomon "
  sSentencia = sSentencia & "ORDER BY codcta, codcco"
  frmTPdoGrd.uocnnMain.Execute sSentencia
  
  ' seleciono informacion
  With porstPspCpr
    .Source = "SELECT pdo.codcta, pdo.codcco, pdo.tpomon, "
    .Source = .Source & "ROUND(AVG(pdo.imporpre_mn), 2) AS imporpre_mn, "
    .Source = .Source & "ROUND(AVG(pdo.imporpre_me), 2) AS imporpre_me, "
    .Source = .Source & "ROUND(SUM(pdo.impopdo_mn), 2) AS impopdo_mn, "
    .Source = .Source & "ROUND(SUM(pdo.impopdo_me), 2) AS impopdo_me "
    .Source = .Source & "FROM " & ps_Prefijo & "tmpvalidarpre pdo "
    .Source = .Source & "GROUP BY pdo.codcta, pdo.codcco, pdo.tpomon "
    .Source = .Source & "HAVING (CASE WHEN pdo.tpomon='" & TPOMON_NAC & "' THEN impopdo_mn>imporpre_mn ELSE impopdo_me>imporpre_me END) "
    .Source = .Source & "ORDER BY codcta, codcco"
    .Open
  End With
  
  If porstPspCpr.RecordCount > 0 Then
    MsgBox Choose(gsIdioma, "Importe de Pedidos es Mayor al Importe del Presupuesto", "Orders Amount is Greater than the Amount of the Budget"), vbCritical
    pfValidoPresupuesto = False
  End If
  If Not pfValidoPresupuesto Then GoTo ErrorVerifica
  
ErrorVerifica:
  Set porstPspCpr = Nothing
  ' informacion de validacion
  sSentencia = "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 14)='#tmpvalidarpre') DROP TABLE #tmpvalidarpre"
  frmTPdoGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS tmpvalidarpre", sSentencia)

End Function
Private Sub ppAbreCuentaPedido()
  frmTPdoGrd.usConnStrgWher_CoPdoCprCta = "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
  frmTPdoGrd.usConnStrgWher_CoPdoCprCta = frmTPdoGrd.usConnStrgWher_CoPdoCprCta & "AND mespvs='" & gsMesAct & "' "
  frmTPdoGrd.usConnStrgWher_CoPdoCprCta = frmTPdoGrd.usConnStrgWher_CoPdoCprCta & "AND coddpe='" & txtllave(0).Text & "' "
  frmTPdoGrd.usConnStrgWher_CoPdoCprCta = frmTPdoGrd.usConnStrgWher_CoPdoCprCta & "AND pdocpr='" & txtllave(1).Text & "' "
  frmTPdoGrd.uocnnMain.Execute "DELETE FROM " & ps_Prefijo & "tmpcopdocprcta"
  frmTPdoGrd.uocnnMain.Execute "INSERT INTO " & ps_Prefijo & "tmpcopdocprcta SELECT * FROM copdocprcta " & frmTPdoGrd.usConnStrgWher_CoPdoCprCta
  ' Información de objeto
  With frmTPdoGrd.uorstCoDPeCta
    If .State = adStateOpen Then .Close
    .Source = frmTPdoGrd.usConnStrgSele_CoPdoCprCta & frmTPdoGrd.usConnStrgWher_CoPdoCprCta & frmTPdoGrd.usConnStrgOrde_CoPdoCprCta
    .Open
    .Properties("Unique Table").Value = ps_Prefijo & "tmpcopdocprcta"
  End With
End Sub
Private Sub ppAbreProductoPedido()
  frmTPdoGrd.usConnStrgWher_CoPdoCprProd = "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
  frmTPdoGrd.usConnStrgWher_CoPdoCprProd = frmTPdoGrd.usConnStrgWher_CoPdoCprProd & "AND mespvs='" & gsMesAct & "' "
  frmTPdoGrd.usConnStrgWher_CoPdoCprProd = frmTPdoGrd.usConnStrgWher_CoPdoCprProd & "AND coddpe='" & txtllave(0).Text & "' "
  frmTPdoGrd.usConnStrgWher_CoPdoCprProd = frmTPdoGrd.usConnStrgWher_CoPdoCprProd & "AND pdocpr='" & txtllave(1).Text & "' "
  frmTPdoGrd.uocnnMain.Execute "DELETE FROM " & ps_Prefijo & "tmpcopdocprprod"
  frmTPdoGrd.uocnnMain.Execute "INSERT INTO " & ps_Prefijo & "tmpcopdocprprod SELECT * FROM copdocprprod " & frmTPdoGrd.usConnStrgWher_CoPdoCprProd
  ' Información de objeto
  With frmTPdoGrd.uorstCoPdoCprProd
    If .State = adStateOpen Then .Close
    .Source = frmTPdoGrd.usConnStrgSele_CoPdoCprProd & frmTPdoGrd.usConnStrgWher_CoPdoCprProd & frmTPdoGrd.usConnStrgOrde_CoPdoCprProd
    .Open
    .Properties("Unique Table").Value = ps_Prefijo & "tmpcopdocprprod"
  End With
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYULLA Then
    Select Case tnIndex
     Case 0                           'Cambiar (añadir índices).
      modAyuBus.DPe_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(coddpe)=4", txtllave(tnIndex).Text, 0, 0, Me.Top + txtllave(tnIndex).Top + txtllave(tnIndex).Height, Me.Left + txtllave(tnIndex).Left
      txtllave(tnIndex).Text = frmOAyuBus.uvDato1
      lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
  Else
    Select Case tnIndex
     Case 0                           'Cambiar (añadir índices).
      modAyuBus.Aux_Det "IndPrv=1", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 7
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 8, 9
      modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "' AND indpdocpr='" & INDCCO_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
  End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYULLA Then
    Select Case tnIndex                 'Cambiar.
     Case 0
      If txtllave(tnIndex).Text = "" Then
        lblLlaveDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTPdoGrd.uorstCoDPe
        If .RecordCount > 0 Then .MoveFirst
        .Find "coddpe='" & txtllave(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblLlaveDeta(tnIndex).Caption = " " & IIf(IsNull(!detdpe), "", !detdpe)
        End If
      End With
    End Select
  Else
    Select Case tnIndex                 'Cambiar.
     Case 0
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTPdoGrd.uorstTGAux
        If .RecordCount > 0 Then .MoveFirst
        .Find "codaux='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!razAux), "", !razAux)
        End If
      End With
     Case 7
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTPdoGrd.uorstCoCta
        If .RecordCount > 0 Then .MoveFirst
        .Find "codcta='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
        End If
      End With
     Case 8, 9
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTPdoGrd.uorstCoCCo
        If .RecordCount > 0 Then .MoveFirst
        .Find "codcco='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detcco), "", !detcco)
        End If
      End With
     Case 10
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTPdoGrd.uorstCoCprProd
        If .RecordCount > 0 Then .MoveFirst
        .Find "codprod='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detprod), "", !detprod)
        End If
      End With
    End Select
  End If
  
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
  'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  
  On Error GoTo Err

  '[Propio del formulario.
  Dim dnContador As Byte
  ']
  With frmTPdoGrd.uorstMain           'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !mespvs = gsMesAct
        !coddpe = txtllave(0).Text
        !pdocpr = txtllave(1).Text
        !indext = chkExtension.Value
      End If

      'Datos.
      !nrointerno = txtllave(2).Text
      !codaux = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      !fehpdo = dtpDato.Value
      !detpdo = IIf(txtDato(gsIdioma).Text = "", Null, txtDato(gsIdioma).Text)
      !detpdox = IIf(txtDato(3 - gsIdioma).Text = "", Null, txtDato(3 - gsIdioma).Text)
      !tpomon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
      '2014-07-18 !tpoigv = IIf(chkCalcularIGV.Value = Val(CODPDO_IGV), CODPDO_IGV, CODPDO_HPR) '2014-05-22 pdo c/igv
      !tpoigv = Choose(cboCalcularIGV.ListIndex + 1, CODPDO_HPR, CODPDO_IGV, CODPDO_IGVG)
      !ImpTCb = CDec(txtDato(3).Text)
      !ImpMN = CDec(txtDato(4).Text)
      !ImpME = CDec(txtDato(5).Text)
      !impdife = CDec(txtDato(6).Text)
      !indcta = cmdMas.Tag
    Else
      'Llaves.
      txtllave(0).Text = !coddpe
      txtllave(1).Text = !pdocpr
      txtllave(2).Text = IIf(IsNull(!nrointerno), "", !nrointerno)
      chkExtension.Value = !indext
      
      'Datos.
      txtDato(0).Text = IIf(IsNull(!codaux), "", !codaux)
      dtpDato.Value = !fehpdo
      txtDato(gsIdioma).Text = IIf(IsNull(!detpdo), "", !detpdo)
      txtDato(3 - gsIdioma).Text = IIf(IsNull(!detpdox), "", !detpdox)
      cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      '2014-07-18 chkCalcularIGV.Value = IIf(!tpoigv = CODPDO_IGV, Val(CODPDO_IGV), Val(CODPDO_HPR)) '2014-05-22 pdo c/igv
      cboCalcularIGV.ListIndex = !tpoigv
      txtDato(3).Text = Format(!ImpTCb, FORMATO_NUM_2)
      txtDato(4).Text = Format(!ImpMN, FORMATO_NUM_1)
      txtDato(5).Text = Format(!ImpME, FORMATO_NUM_1)
      txtDato(6).Text = Format(!impdife, FORMATO_NUM_1)
      
      ' Actualizo la información de cuentas
      cmdMas.Tag = !indcta
      ppAbreCuentaPedido
      txtDato(7).Text = IIf(IsNull(frmTPdoGrd.uorstCoDPeCta!CodCta), "", frmTPdoGrd.uorstCoDPeCta!CodCta)
      txtDato(8).Text = IIf(IsNull(frmTPdoGrd.uorstCoDPeCta!codcco), "", frmTPdoGrd.uorstCoDPeCta!codcco)
      txtDato(7).Tag = txtDato(7).Text
      ' Actualizo la información de productos
      cmdMasProducto.Tag = !indcta
      txtDato(9).Text = ""
      txtDato(10).Text = ""
      ppAbreProductoPedido
      If frmTPdoGrd.uorstCoPdoCprProd.RecordCount Then
        txtDato(9).Text = IIf(IsNull(frmTPdoGrd.uorstCoPdoCprProd!codcco), "", frmTPdoGrd.uorstCoPdoCprProd!codcco)
        txtDato(10).Text = IIf(IsNull(frmTPdoGrd.uorstCoPdoCprProd!codprod), "", frmTPdoGrd.uorstCoPdoCprProd!codprod)
      End If
      
      txtDato(4).Tag = Format(txtDato(6).Text, FORMATO_NUM_1)
      txtDato(5).Tag = Format(txtDato(7).Text, FORMATO_NUM_1)
      '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      ppAyuDet AYUDAT, 0
      psCodCCo_Pdo = IIf(IsNull(frmTPdoGrd.uorstCoDPe!codcco), "", frmTPdoGrd.uorstCoDPe!codcco)
      ppAyuDet AYUDAT, 7
      ppAyuDet AYUDAT, 8
      ppAyuDet AYUDAT, 9
      ppAyuDet AYUDAT, 10
      ']
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
  txtllave(0).Text = ""
  txtllave(1).Text = ""
  txtllave(2).Text = ""
  chkExtension.Value = vbUnchecked
  
  'Datos.
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  
'2014-07-18 chkCalcularIGV.Value = 1  '2014-05-22
  cboCalcularIGV.ListIndex = CODPDO_IGV
  
  dtpDato.Value = Date
  pnCta_TpoTcb = TPOTCB_VTA
  For dnContador = 0 To 10
    txtDato(dnContador).Text = ""
  Next
  txtDato(3).Text = Format(0, FORMATO_NUM_2)
  txtDato(4).Text = Format(0, FORMATO_NUM_1)
  txtDato(5).Text = Format(0, FORMATO_NUM_1)
  txtDato(6).Text = Format(0, FORMATO_NUM_1)
  
  txtDato(3).Tag = Format(0, FORMATO_NUM_2)
  txtDato(4).Tag = Format(0, FORMATO_NUM_1)
  txtDato(5).Tag = Format(0, FORMATO_NUM_1)

  'Ayudas.
  lblLlaveDeta(0).Caption = ""
  lblDatoDeta(0).Caption = ""
  lblDatoDeta(7).Caption = ""
  lblDatoDeta(8).Caption = ""
  lblDatoDeta(9).Caption = ""
  lblDatoDeta(10).Caption = ""
  
  ' Inicializo detalle de cuenta
  txtDato(7).Tag = ""
  cmdMas.Tag = INDMASCTA_INI
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Byte
  
  'Datos.

'2014-07-18 error igv  chkCalcularIGV.Enabled = tbHabilitar '2014-05-22
cboCalcularIGV.Enabled = tbHabilitar '2014-05-22

  cboTpoMon.Enabled = tbHabilitar
  dtpDato.Enabled = tbHabilitar
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  txtllave(2).Enabled = tbHabilitar
  txtDato(10).Enabled = False
  'Ayudas.
  cmdLlaveAyud(0).Enabled = (pbNuevo)
  cmdMas.Enabled = tbHabilitar
  cmdMasProducto.Enabled = tbHabilitar
  cmdDatoAyud(0).Enabled = tbHabilitar
  cmdDatoAyud(7).Enabled = tbHabilitar
  cmdDatoAyud(8).Enabled = tbHabilitar
  cmdDatoAyud(9).Enabled = tbHabilitar
End Sub

Private Sub cmdAuxiliar_Click()
  frmMAuxGrd.Show vbModal
  frmTPdoGrd.uorstTGAux.Requery
End Sub

Private Sub dtpDato_Validate(Cancel As Boolean)
      
  If Not (Month(dtpDato.Value) >= Val(gsMesAct) And Year(dtpDato.Value) >= Val(gsAnoAct)) Then
    MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
    dtpDato.SetFocus
    Cancel = True
    Exit Sub
  End If
  dtpDato.Tag = 0
  If (dtpDato.Tag <> dtpDato.Value) Then
    dtpDato.Tag = dtpDato.Value
    With frmTPdoGrd.uorstTGTCb
      If .RecordCount <> 0 Then
        .MoveFirst
        .Find "(FehTCb) = '" & Format(dtpDato.Value, "yyyy/mm/dd") & "'"
        ' [Adicional Agregado por Angel
        If .EOF Then
          MsgBox TEXT_9015, vbExclamation
          txtDato(3).Text = Format(0, FORMATO_NUM_2)
          txtDato(3).SetFocus
          Cancel = True
          Exit Sub
        Else
          txtDato(3).Text = Format(IIf(pnCta_TpoTcb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta), FORMATO_NUM_2)
        End If
        ']
      Else
         txtDato(3).Text = Format(0, FORMATO_NUM_2)
      End If
    End With
  End If

End Sub

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
