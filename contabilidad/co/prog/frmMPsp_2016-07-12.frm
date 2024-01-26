VERSION 5.00
Begin VB.Form frmMPsp 
   Caption         =   "[Entidad]"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7635
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox checkmeses 
      Caption         =   "Incluir en el calculo los Meses posteriores"
      Height          =   615
      Left            =   5640
      TabIndex        =   59
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Calcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   5640
      TabIndex        =   58
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   305
      Index           =   1
      Left            =   7200
      Picture         =   "frmMPsp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
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
      Left            =   720
      TabIndex        =   55
      Top             =   600
      Width           =   950
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   23
      Left            =   3420
      TabIndex        =   24
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   22
      Left            =   1800
      TabIndex        =   23
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   21
      Left            =   3420
      TabIndex        =   22
      Top             =   4740
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   20
      Left            =   1800
      TabIndex        =   21
      Top             =   4740
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   19
      Left            =   3420
      TabIndex        =   20
      Top             =   4380
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   18
      Left            =   1800
      TabIndex        =   19
      Top             =   4380
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   17
      Left            =   3420
      TabIndex        =   18
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   16
      Left            =   1800
      TabIndex        =   17
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   15
      Left            =   3420
      TabIndex        =   16
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   14
      Left            =   1800
      TabIndex        =   15
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Index           =   13
      Left            =   3420
      TabIndex        =   14
      Top             =   3420
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   13
      Top             =   3420
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   3420
      TabIndex        =   12
      Top             =   3060
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   11
      Top             =   3060
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   3420
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   3420
      TabIndex        =   8
      Top             =   2460
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   7
      Top             =   2460
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   3420
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   3420
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame fraImpresion 
      Caption         =   "Impresión"
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   240
      TabIndex        =   39
      Top             =   5400
      Width           =   3495
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
         Index           =   24
         Left            =   2760
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox CboTpoGru 
         Height          =   315
         Left            =   720
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Orden:"
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
         Index           =   16
         Left            =   2160
         TabIndex        =   41
         Top             =   420
         Width           =   495
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
         Index           =   15
         Left            =   120
         TabIndex        =   40
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1500
      Width           =   1575
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
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
      Left            =   3420
      TabIndex        =   2
      Top             =   1500
      Width           =   1575
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   305
      Index           =   0
      Left            =   7200
      Picture         =   "frmMPsp.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1920
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6240
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
         Picture         =   "frmMPsp.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Picture         =   "frmMPsp.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Picture         =   "frmMPsp.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "frmMPsp.frx":07F2
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Picture         =   "frmMPsp.frx":08F4
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Picture         =   "frmMPsp.frx":09F6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   60
         Width           =   720
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   950
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
      Index           =   1
      Left            =   1680
      TabIndex        =   56
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "C.Costos:"
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
      Index           =   17
      Left            =   60
      TabIndex        =   54
      Top             =   600
      Width           =   705
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Diciembre:"
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
      Left            =   240
      TabIndex        =   53
      Top             =   5100
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Noviembre:"
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
      Left            =   240
      TabIndex        =   52
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Octubre.:"
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
      Left            =   240
      TabIndex        =   51
      Top             =   4440
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Setiembre:"
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
      Left            =   240
      TabIndex        =   50
      Top             =   4140
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Agosto...:"
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
      Left            =   240
      TabIndex        =   49
      Top             =   3780
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Julio.......:"
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
      Left            =   240
      TabIndex        =   48
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Junio......:"
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
      Left            =   240
      TabIndex        =   47
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Mayo......:"
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
      Left            =   240
      TabIndex        =   46
      Top             =   2820
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Abril.......:"
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
      Left            =   240
      TabIndex        =   45
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Marzo.....:"
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
      Left            =   240
      TabIndex        =   44
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Febrero...:"
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
      Left            =   240
      TabIndex        =   43
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Enero.......:"
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
      Left            =   240
      TabIndex        =   42
      Top             =   1560
      Width           =   900
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
      TabIndex        =   38
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe M.E.:"
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
      Index           =   14
      Left            =   3600
      TabIndex        =   36
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Importe M.N.:"
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
      Left            =   1920
      TabIndex        =   34
      Top             =   1200
      Width           =   930
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
      Index           =   0
      Left            =   60
      TabIndex        =   33
      Top             =   180
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   7440
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmMPsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbValidada As Boolean
Private ultimo As Integer
Private Sub Calcular_Click()

Dim meses As Boolean
Dim opcion As String
Dim i As Integer

If checkmeses.Value = Checked Then
    meses = True
Else
    meses = False
End If
Select Case ultimo
Case 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22
    opcion = "IZQ"
    For i = ultimo To 22 Step 2
        txtDato(i).Text = txtDato(ultimo).Text
        If meses = False Then Exit For
    Next
Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23
    opcion = "DER"
    For i = ultimo To 23 Step 2
        txtDato(i).Text = txtDato(ultimo).Text
        If meses = False Then Exit For
    Next
End Select
End Sub

Private Sub Calcular_Click_2016_01_17()
Dim i As Integer
Dim k As Integer
Dim x As Integer
Dim Rst As ADODB.Recordset
Dim sql As String
Dim cambio As Double
Dim opcion As String
Dim meses As Boolean

If txtLlave(0) = "" Or txtLlave(1) = "" Then Exit Sub

'Select Case MsgBox("Procesar los Meses Siguientes", vbYesNoCancel + vbQuestion, "Contabilidad")
'Case vbYes
'    meses = True
'Case vbNo
'    meses = False
'Case vbCancel
'    Exit Sub
'End Select

If checkmeses.Value = Checked Then
    meses = True
Else
    meses = False
End If

Select Case ultimo
Case 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22
    opcion = "IZQ"
    For i = ultimo To 22 Step 2
        txtDato(i).Text = txtDato(ultimo).Text
        If meses = False Then Exit For
    Next
Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23
    opcion = "DER"
    For i = ultimo To 23 Step 2
        txtDato(i).Text = txtDato(ultimo).Text
        If meses = False Then Exit For
    Next
End Select

Set Rst = New ADODB.Recordset

sql = "select MesPvs, ImpTCb_Cpr, pdoano FROM CoTCbMes where codemp='" & gsCodEmp & "' and pdoano='" & gsAnoAct & "' and MesPvs='" & gsMesAct & "'"
Rst.Open sql, frmMPspGrd.uocnnMain, adOpenStatic, adLockOptimistic
    
If Rst.RecordCount = 0 Then
    cambio = 0
Else
    Rst.MoveFirst
    For x = 0 To Rst.RecordCount - 1
        cambio = Rst.Fields(1)
        Rst.MoveNext
    Next
End If
Rst.Close

Select Case ultimo
Case 0, 1
    x = 1
Case 2, 3
    x = 2
Case 4, 5
    x = 3
Case 6, 7
    x = 4
Case 8, 9
    x = 5
Case 10, 11
    x = 6
Case 12, 13
    x = 7
Case 14, 15
    x = 8
Case 16, 17
    x = 9
Case 18, 19
    x = 10
Case 20, 21
    x = 11
Case 22, 23
    x = 12
End Select

Select Case opcion
Case "IZQ"
    For i = x To 12
        Select Case i
        Case 1
            k = 1
        Case 2
            k = 3
        Case 3
            k = 5
        Case 4
            k = 7
        Case 5
            k = 9
        Case 6
            k = 11
        Case 7
            k = 13
        Case 8
            k = 15
        Case 9
            k = 17
        Case 10
            k = 19
        Case 11
            k = 21
        Case 12
            k = 23
        End Select
        If cambio = 0 Then
            txtDato(k).Text = Format(0, FORMATO_NUM_1)
        Else
            txtDato(k).Text = Format(txtDato(k - 1).Text / cambio, FORMATO_NUM_1)
        End If
        If meses = False Then Exit For
    Next
Case "DER"
    For i = x To 12
        Select Case i
        Case 1
            k = 0
        Case 2
            k = 2
        Case 3
            k = 4
        Case 4
            k = 6
        Case 5
            k = 8
        Case 6
            k = 10
        Case 7
            k = 12
        Case 8
            k = 14
        Case 9
            k = 16
        Case 10
            k = 18
        Case 11
            k = 20
        Case 12
            k = 22
        End Select
        If cambio = 0 Then
            txtDato(k).Text = Format(0, FORMATO_NUM_1)
        Else
            txtDato(k).Text = Format(txtDato(k + 1).Text * cambio, FORMATO_NUM_1)
        End If
        If meses = False Then Exit For
    Next
End Select
End Sub

Private Sub Form_Load()
   pbValidada = False
   
   'Calcular.Enabled = False

   Me.KeyPreview = True
   
   With frmMPspGrd                     'Cambiar Formulario de Grid.
    '[Llaves                           'Cambiar
      txtLlave(0).MaxLength = .uorstMain!CodCta.DefinedSize
      txtLlave(1).MaxLength = .uorstMain!codcco.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
       
      'txtDato(0).MaxLength = .uorstMain.Fields("ImpMN_" & gsMesAct).DefinedSize
      'txtDato(1).MaxLength = .uorstMain.Fields("ImpME_" & gsMesAct).DefinedSize
      'corregido solo se podia ingresar 8 caracteres contando las comas y decimales
      txtDato(0).Text = .uorstMain.Fields("ImpMN_" & "01").DefinedSize
      txtDato(1).Text = .uorstMain.Fields("ImpME_" & "01").DefinedSize
      txtDato(2).Text = .uorstMain.Fields("ImpMN_" & "02").DefinedSize
      txtDato(3).Text = .uorstMain.Fields("ImpME_" & "02").DefinedSize
      txtDato(4).Text = .uorstMain.Fields("ImpMN_" & "03").DefinedSize
      txtDato(5).Text = .uorstMain.Fields("ImpME_" & "03").DefinedSize
      txtDato(6).Text = .uorstMain.Fields("ImpMN_" & "04").DefinedSize
      txtDato(7).Text = .uorstMain.Fields("ImpME_" & "04").DefinedSize
      txtDato(8).Text = .uorstMain.Fields("ImpMN_" & "05").DefinedSize
      txtDato(9).Text = .uorstMain.Fields("ImpME_" & "05").DefinedSize
      txtDato(10).Text = .uorstMain.Fields("ImpMN_" & "06").DefinedSize
      txtDato(11).Text = .uorstMain.Fields("ImpME_" & "06").DefinedSize
      txtDato(12).Text = .uorstMain.Fields("ImpMN_" & "07").DefinedSize
      txtDato(13).Text = .uorstMain.Fields("ImpME_" & "07").DefinedSize
      txtDato(14).Text = .uorstMain.Fields("ImpMN_" & "08").DefinedSize
      txtDato(15).Text = .uorstMain.Fields("ImpME_" & "08").DefinedSize
      txtDato(16).Text = .uorstMain.Fields("ImpMN_" & "09").DefinedSize
      txtDato(17).Text = .uorstMain.Fields("ImpME_" & "09").DefinedSize
      txtDato(18).Text = .uorstMain.Fields("ImpMN_" & "10").DefinedSize
      txtDato(19).Text = .uorstMain.Fields("ImpME_" & "10").DefinedSize
      txtDato(20).Text = .uorstMain.Fields("ImpMN_" & "11").DefinedSize
      txtDato(21).Text = .uorstMain.Fields("ImpME_" & "11").DefinedSize
      txtDato(22).Text = .uorstMain.Fields("ImpMN_" & "12").DefinedSize
      txtDato(23).Text = .uorstMain.Fields("ImpME_" & "12").DefinedSize
      
      txtDato(24).MaxLength = .uorstMain.Fields("OrdRep").DefinedSize - 1
      With CboTpoGru
         .AddItem TPOGRU1_TXT_1, TPOGRU1_IND
         .AddItem TPOGRU2_TXT_1, TPOGRU2_IND
         .AddItem TPOGRU3_TXT_1, TPOGRU3_IND
      End With
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
  ReDim aLabel(17, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Cuenta :", "Enero :", "Febrero :", "Marzo :", "Abril :", "Mayo :", "Junio :", "Julio :", "Agosto :", "Setiembre :", "Octubre :", "Noviembre :", "Diciembre :", "Importe M.N.:", "Importe M.E.:", "Grupo :", "Orden :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Account :", "January :", "February :", "March :", "April :", "May :", "June :", "July :", "August :", "September :", "October :", "November :", "December :", "Amount N.C.:", "Amount F.C.:", "Group :", "Order :")
  Next nElemento
  fraImpresion.Caption = Choose(gsIdioma, "Impresión", "Printing")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']

End Sub

Private Sub Form_Activate()
 '[Busca detalle de códigos            'Cambiar (habilitar/deshabilitar).
   If txtLlave(0).Text <> "" Then ppAyuDet 0
   If txtLlave(1).Text <> "" Then ppAyuDet 1
 ']
''   If pbNuevo Then
''      With frmMPspGrd.porstUltOrdRep
''         .Open
''         txtDato(2).Text = !OrdRep
''         .Close
''      End With
''   End If
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not frmMPspGrd.uorstMain.EOF Then
     If frmMPspGrd.uorstMain.EditMode <> adEditNone Then frmMPspGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmMPspGrd.uorstMain, Me 'Cambiar Formulario de Grid.
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmMPspGrd.uorstMain, Me 'Cambiar Formulario de Grid.
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
 
 'Calcular.Enabled = True
 
End Sub

Public Sub cmdGrabar_Click()
  

   With frmMPspGrd                     'Cambiar Formulario de Grid.
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
         .uorstMain.Find "CodCta='" & txtLlave(0).Text & "'"
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
   frmMPspGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.

End Sub

Public Sub cmdDeshacer_Click()
   gpTUe_Deshacer Me
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      'txtLlave(Index).SetFocus
   End Select
   ppAyuBus Index
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
   'If pbValidada Then txtDato(0).SetFocus
End Sub

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus Index
   End If
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
   'If lblLlaveDeta(0) = "" Or lblLlaveDeta(1) = "" Then Exit Sub
   
   If lblLlaveDeta(0) = "" Then Exit Sub
   
  'Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
'   Select Case Index                   'Cambiar (añadir índices).
'   Case 0
'      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
'         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
'      End If
'   End Select
   
  'Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index                   'Cambiar (añadir índices).
   Case 0, 1
      Cancel = ppAyuDet(Index)
      If Cancel Then Exit Sub
   End Select
 
  
  'Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 Then 'And Len(Trim(txtLlave(1).Text)) <> 0 Then
      With frmMPspGrd.uorstMain
         If Not (.BOF And .EOF) Then
            dvRegistro = .Bookmark
            .MoveFirst
            .Find "llave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
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
      txtDato(0).SetFocus
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
   
   ultimo = Index
  
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(txtDato(Index))) + 1 = txtDato(Index).MaxLength Then
      SendKeys "{TAB}"
   End If
']ARREGLAR.

 '[Convierte a mayúsculas.
   If Index = 1 Then                   'Cambiar (añadir índices).
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
   Select Case Index
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23       'Cambiar (añadir índices).
      If Not IsNumeric(txtDato(Index).Text) Then
         txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      Else
         txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
      End If
   End Select

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
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA, "", 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
      lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   Case 1
      modAyuBus.CCo_Cod "length(codcco)=2 ", "", 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
      With frmMPspGrd.porstCOCta
         .MoveFirst
         .Find "CodCta='" & txtLlave(tnIndex).Text & "'"
         If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
         Else
            lblLlaveDeta(tnIndex).Caption = " " & IIf(IsNull(!detcta), "", !detcta)
         End If
      End With
   End Select
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   'On Error GoTo Err

   With frmMPspGrd
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            .uorstMain!codemp = gsCodEmp
            .uorstMain!pdoano = gsAnoAct
            .uorstMain!CodCta = txtLlave(0).Text
            .uorstMain!codcco = txtLlave(1).Text
         End If

        'Datos.
        .uorstMain!OrdRep = IIf(CboTpoGru.ListIndex = TPOGRU1_IND, TPOGRU1_TXT_0, IIf(CboTpoGru.ListIndex = TPOGRU2_IND, TPOGRU2_TXT_0, TPOGRU3_TXT_0)) & gfCeros(txtDato(24).Text, 3, 0, "0")
        .uorstMain.Fields("ImpMN_" & "01") = CDec(txtDato(0).Text)
        .uorstMain.Fields("ImpME_" & "01") = CDec(txtDato(1).Text)
        .uorstMain.Fields("ImpMN_" & "02") = CDec(txtDato(2).Text)
        .uorstMain.Fields("ImpME_" & "02") = CDec(txtDato(3).Text)
        .uorstMain.Fields("ImpMN_" & "03") = CDec(txtDato(4).Text)
        .uorstMain.Fields("ImpME_" & "03") = CDec(txtDato(5).Text)
        .uorstMain.Fields("ImpMN_" & "04") = CDec(txtDato(6).Text)
        .uorstMain.Fields("ImpME_" & "04") = CDec(txtDato(7).Text)
        .uorstMain.Fields("ImpMN_" & "05") = CDec(txtDato(8).Text)
        .uorstMain.Fields("ImpME_" & "05") = CDec(txtDato(9).Text)
        .uorstMain.Fields("ImpMN_" & "06") = CDec(txtDato(10).Text)
        .uorstMain.Fields("ImpME_" & "06") = CDec(txtDato(11).Text)
        .uorstMain.Fields("ImpMN_" & "07") = CDec(txtDato(12).Text)
        .uorstMain.Fields("ImpME_" & "07") = CDec(txtDato(13).Text)
        .uorstMain.Fields("ImpMN_" & "08") = CDec(txtDato(14).Text)
        .uorstMain.Fields("ImpME_" & "08") = CDec(txtDato(15).Text)
        .uorstMain.Fields("ImpMN_" & "09") = CDec(txtDato(16).Text)
        .uorstMain.Fields("ImpME_" & "09") = CDec(txtDato(17).Text)
        .uorstMain.Fields("ImpMN_" & "10") = CDec(txtDato(18).Text)
        .uorstMain.Fields("ImpME_" & "10") = CDec(txtDato(19).Text)
        .uorstMain.Fields("ImpMN_" & "11") = CDec(txtDato(20).Text)
        .uorstMain.Fields("ImpME_" & "11") = CDec(txtDato(21).Text)
        .uorstMain.Fields("ImpMN_" & "12") = CDec(txtDato(22).Text)
        .uorstMain.Fields("ImpME_" & "12") = CDec(txtDato(23).Text)
      Else
        'Llaves.
         txtLlave(0).Text = .uorstMain!CodCta
         txtLlave(1).Text = .uorstMain!codcco
       
        'Datos.
'         chkEstado.Value = IIf(uorstMain!EstTDc = ESTTDc_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(uorstMain!CodSoc), "", uorstMain!CodSoc)
'         dtpFecha.Value = uorstMain!FehOpe
'         optMoneda(1).Value = uorstMain!CodMon
         'txtDato(0).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_" & gsMesAct)), 0, .uorstMain.Fields("ImpMN_" & gsMesAct)), FORMATO_NUM_1)
         'txtDato(1).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_" & gsMesAct)), 0, .uorstMain.Fields("ImpME_" & gsMesAct)), FORMATO_NUM_1)
         
        txtDato(0).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_01")), 0, .uorstMain.Fields("ImpMN_01")), FORMATO_NUM_1)
        txtDato(1).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_01")), 0, .uorstMain.Fields("ImpME_01")), FORMATO_NUM_1)
        txtDato(2).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_02")), 0, .uorstMain.Fields("ImpMN_02")), FORMATO_NUM_1)
        txtDato(3).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_02")), 0, .uorstMain.Fields("ImpME_02")), FORMATO_NUM_1)
        txtDato(4).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_03")), 0, .uorstMain.Fields("ImpMN_03")), FORMATO_NUM_1)
        txtDato(5).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_03")), 0, .uorstMain.Fields("ImpME_03")), FORMATO_NUM_1)
        txtDato(6).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_04")), 0, .uorstMain.Fields("ImpMN_04")), FORMATO_NUM_1)
        txtDato(7).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_04")), 0, .uorstMain.Fields("ImpME_04")), FORMATO_NUM_1)
        txtDato(8).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_05")), 0, .uorstMain.Fields("ImpMN_05")), FORMATO_NUM_1)
        txtDato(9).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_05")), 0, .uorstMain.Fields("ImpME_05")), FORMATO_NUM_1)
        txtDato(10).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_06")), 0, .uorstMain.Fields("ImpMN_06")), FORMATO_NUM_1)
        txtDato(11).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_06")), 0, .uorstMain.Fields("ImpME_06")), FORMATO_NUM_1)
        txtDato(12).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_07")), 0, .uorstMain.Fields("ImpMN_07")), FORMATO_NUM_1)
        txtDato(13).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_07")), 0, .uorstMain.Fields("ImpME_07")), FORMATO_NUM_1)
        txtDato(14).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_08")), 0, .uorstMain.Fields("ImpMN_08")), FORMATO_NUM_1)
        txtDato(15).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_08")), 0, .uorstMain.Fields("ImpME_08")), FORMATO_NUM_1)
        txtDato(16).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_09")), 0, .uorstMain.Fields("ImpMN_09")), FORMATO_NUM_1)
        txtDato(17).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_09")), 0, .uorstMain.Fields("ImpME_09")), FORMATO_NUM_1)
        txtDato(18).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_10")), 0, .uorstMain.Fields("ImpMN_10")), FORMATO_NUM_1)
        txtDato(19).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_10")), 0, .uorstMain.Fields("ImpME_10")), FORMATO_NUM_1)
        txtDato(20).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_11")), 0, .uorstMain.Fields("ImpMN_11")), FORMATO_NUM_1)
        txtDato(21).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_11")), 0, .uorstMain.Fields("ImpME_11")), FORMATO_NUM_1)
        txtDato(22).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpMN_12")), 0, .uorstMain.Fields("ImpMN_12")), FORMATO_NUM_1)
        txtDato(23).Text = Format(IIf(IsNull(.uorstMain.Fields("ImpME_12")), 0, .uorstMain.Fields("ImpME_12")), FORMATO_NUM_1)
        CboTpoGru.ListIndex = IIf(Left(.uorstMain!OrdRep, 1) = TPOGRU1_TXT_0, TPOGRU1_IND, IIf(Left(.uorstMain!OrdRep, 1) = TPOGRU2_TXT_0, TPOGRU2_IND, TPOGRU3_IND))
        txtDato(24).Text = Right(.uorstMain!OrdRep, 3)
        
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
'   chkEstado.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
   With txtDato
      For dnContador = 0 To .Count - 2
         .Item(dnContador).Text = Format(0, FORMATO_NUM_1)
      Next
      .Item(dnContador).Text = ""
   End With
   CboTpoGru.ListIndex = TPOGRU1_IND

  'Ayudas.
   lblLlaveDeta(0).Caption = ""
   lblLlaveDeta(1).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer
  'Llaves
   With txtLlave
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
      Next
   End With
  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   CboTpoGru.Enabled = tbHabilitar

  'Ayudas.
'   cmdDatoAyud(0).Enabled = tbHabilitar
   lblLlaveDeta(0).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
   lblLlaveDeta(1).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
   cmdLlaveAyud(0).Enabled = IIf(pbNuevo, Not tbHabilitar, False)
   
   
   
   
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

