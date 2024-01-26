VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTRteVta 
   Caption         =   "[Título]"
   ClientHeight    =   6900
   ClientLeft      =   840
   ClientTop       =   1245
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   9480
   Begin VB.CheckBox chkIndEstado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contrato Activo"
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
      Height          =   255
      Left            =   210
      TabIndex        =   97
      Top             =   3375
      Width           =   1500
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
      Left            =   7905
      TabIndex        =   26
      Top             =   2655
      Width           =   560
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
      Left            =   1110
      TabIndex        =   28
      Top             =   3015
      Width           =   560
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   7
      Left            =   6990
      Picture         =   "frmTRteVta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   3015
      Width           =   255
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   5
      Left            =   6990
      Picture         =   "frmTRteVta.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   2655
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
      Index           =   5
      Left            =   1110
      TabIndex        =   23
      Top             =   2655
      Width           =   315
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
      Height          =   280
      Index           =   3
      Left            =   4050
      TabIndex        =   16
      Top             =   1650
      Width           =   4515
   End
   Begin VB.ComboBox cboRetencion 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2010
      Width           =   2295
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
      Height          =   510
      Index           =   4
      Left            =   4050
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   2010
      Width           =   4530
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
      Height          =   500
      Left            =   8730
      Picture         =   "frmTRteVta.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   2670
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "Cli&ente"
      Height          =   315
      Left            =   8625
      TabIndex        =   94
      Top             =   3210
      Width           =   825
   End
   Begin VB.CheckBox chkCalcularISC 
      Caption         =   "Calcular I.&S.C."
      ForeColor       =   &H00800000&
      Height          =   200
      Left            =   1740
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   8640
      ScaleHeight     =   2535
      ScaleWidth      =   885
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   120
      Width           =   885
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
         Left            =   98
         Picture         =   "frmTRteVta.frx":092A
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1935
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
         Left            =   98
         Picture         =   "frmTRteVta.frx":0A74
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1380
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
         Left            =   98
         Picture         =   "frmTRteVta.frx":0B76
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   825
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
         Left            =   98
         Picture         =   "frmTRteVta.frx":0C78
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   270
         Width           =   720
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
         Left            =   98
         Picture         =   "frmTRteVta.frx":0DC2
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   0
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
         Left            =   458
         Picture         =   "frmTRteVta.frx":0F6C
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   8295
      Picture         =   "frmTRteVta.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   990
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
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   990
      Width           =   1275
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTRteVta.frx":12C0
      Left            =   1080
      List            =   "frmTRteVta.frx":12C2
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1650
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   1
      Left            =   3330
      TabIndex        =   6
      Top             =   615
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   113901569
      CurrentDate     =   37102
   End
   Begin VB.CheckBox chkCalcularIGV 
      Caption         =   "Calcular I.G.&V."
      ForeColor       =   &H00800000&
      Height          =   200
      Left            =   60
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1485
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
      Height          =   285
      Index           =   2
      Left            =   4050
      TabIndex        =   15
      Top             =   1320
      Width           =   4515
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
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   1320
      Width           =   2310
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
      Left            =   2145
      TabIndex        =   2
      Top             =   120
      Width           =   1245
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
      Left            =   1275
      TabIndex        =   1
      Top             =   120
      Width           =   570
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   8
      Top             =   615
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   113901569
      CurrentDate     =   37102
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   2775
      Left            =   0
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   4095
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   441
      ForeColor       =   8388608
      TabCaption(0)   =   "I&mportes"
      TabPicture(0)   =   "frmTRteVta.frx":12C4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(20)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTexto(13)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTexto(12)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTexto(14)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTexto(16)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTexto(15)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTexto(19)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblTexto(18)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTexto(17)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDatoDeta(15)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDatoDeta(16)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDatoDeta(17)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDatoDeta(18)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDatoDeta(19)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDatoDeta(20)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDatoDeta(21)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDatoDeta(22)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblDatoDeta(23)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblDatoDeta(24)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblDatoDeta(25)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblDatoDeta(26)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblDatoDeta(27)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblDatoDeta(28)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDato(8)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtDato(15)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtDato(16)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtDato(17)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtDato(18)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtDato(19)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtDato(20)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtDato(21)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtDato(22)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtDato(23)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtDato(24)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtDato(25)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtDato(9)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtDato(10)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtDato(11)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtDato(12)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtDato(13)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtDato(14)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdMas(1)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdMas(2)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdMas(3)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdMas(4)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdMas(5)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cmdMas(6)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cmdMas(7)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtDato(26)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtDato(27)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtDato(28)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "C&uentas"
      TabPicture(1)   =   "frmTRteVta.frx":12E0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgrDetalle"
      Tab(1).ControlCount=   1
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
         Index           =   28
         Left            =   6840
         TabIndex        =   81
         Top             =   2310
         Width           =   675
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
         Index           =   27
         Left            =   6840
         TabIndex        =   74
         Top             =   2025
         Width           =   675
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
         Index           =   26
         Left            =   6840
         TabIndex        =   67
         Top             =   1740
         Width           =   675
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
         Height          =   280
         Index           =   7
         Left            =   3000
         Picture         =   "frmTRteVta.frx":12FC
         TabIndex        =   78
         Top             =   2310
         Width           =   255
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
         Height          =   280
         Index           =   6
         Left            =   3000
         Picture         =   "frmTRteVta.frx":13FE
         TabIndex        =   71
         Top             =   2025
         Width           =   255
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
         Height          =   280
         Index           =   5
         Left            =   3000
         Picture         =   "frmTRteVta.frx":1500
         TabIndex        =   64
         Top             =   1740
         Width           =   255
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
         Height          =   280
         Index           =   4
         Left            =   3000
         Picture         =   "frmTRteVta.frx":1602
         TabIndex        =   57
         Top             =   1455
         Width           =   255
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
         Height          =   280
         Index           =   3
         Left            =   3000
         Picture         =   "frmTRteVta.frx":1704
         TabIndex        =   50
         Top             =   1170
         Width           =   255
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
         Height          =   280
         Index           =   2
         Left            =   3000
         Picture         =   "frmTRteVta.frx":1806
         TabIndex        =   43
         Top             =   885
         Width           =   255
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
         Height          =   280
         Index           =   1
         Left            =   3000
         Picture         =   "frmTRteVta.frx":1908
         TabIndex        =   36
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,###,###,##0.00"
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
         Index           =   14
         Left            =   1320
         TabIndex        =   77
         Top             =   2310
         Width           =   1695
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,###,###,##0.00"
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
         Index           =   13
         Left            =   1320
         TabIndex        =   70
         Top             =   2025
         Width           =   1695
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,###,###,##0.00"
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
         Index           =   12
         Left            =   1320
         TabIndex        =   63
         Top             =   1740
         Width           =   1695
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,###,###,##0.00"
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
         Index           =   11
         Left            =   1320
         TabIndex        =   56
         Top             =   1455
         Width           =   1695
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,###,###,##0.00"
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
         Index           =   10
         Left            =   1320
         TabIndex        =   49
         Top             =   1170
         Width           =   1695
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Index           =   9
         Left            =   1320
         TabIndex        =   42
         Top             =   885
         Width           =   1695
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
         Index           =   25
         Left            =   6840
         TabIndex        =   60
         Top             =   1455
         Width           =   675
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
         Index           =   24
         Left            =   6840
         TabIndex        =   53
         Top             =   1170
         Width           =   675
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
         Index           =   23
         Left            =   6840
         TabIndex        =   46
         Top             =   885
         Width           =   675
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
         Index           =   22
         Left            =   6840
         TabIndex        =   39
         Top             =   600
         Width           =   675
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
         Index           =   21
         Left            =   3300
         TabIndex        =   79
         Top             =   2310
         Width           =   975
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
         Index           =   20
         Left            =   3300
         TabIndex        =   72
         Top             =   2025
         Width           =   975
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
         Index           =   19
         Left            =   3300
         TabIndex        =   65
         Top             =   1740
         Width           =   975
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
         Index           =   18
         Left            =   3300
         TabIndex        =   58
         Top             =   1455
         Width           =   975
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
         Index           =   17
         Left            =   3300
         TabIndex        =   51
         Top             =   1170
         Width           =   975
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
         Index           =   16
         Left            =   3300
         TabIndex        =   44
         Top             =   885
         Width           =   975
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
         Index           =   15
         Left            =   3300
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   2325
         Left            =   -74880
         TabIndex        =   84
         Top             =   345
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4101
         _Version        =   393216
         ForeColor       =   -2147483630
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
         Index           =   8
         Left            =   1320
         TabIndex        =   35
         Top             =   600
         Width           =   1695
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
         Index           =   28
         Left            =   7500
         TabIndex        =   82
         Top             =   2310
         Width           =   1845
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
         Index           =   27
         Left            =   7500
         TabIndex        =   75
         Top             =   2025
         Width           =   1845
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
         Index           =   26
         Left            =   7500
         TabIndex        =   68
         Top             =   1740
         Width           =   1845
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
         Index           =   25
         Left            =   7500
         TabIndex        =   61
         Top             =   1455
         Width           =   1845
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
         Index           =   24
         Left            =   7500
         TabIndex        =   54
         Top             =   1170
         Width           =   1845
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
         Index           =   23
         Left            =   7500
         TabIndex        =   47
         Top             =   885
         Width           =   1845
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
         Index           =   22
         Left            =   7500
         TabIndex        =   40
         Top             =   600
         Width           =   1845
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
         Index           =   21
         Left            =   4260
         TabIndex        =   80
         Top             =   2310
         Width           =   2565
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
         Index           =   20
         Left            =   4260
         TabIndex        =   73
         Top             =   2025
         Width           =   2565
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
         Index           =   19
         Left            =   4260
         TabIndex        =   66
         Top             =   1740
         Width           =   2565
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
         Index           =   18
         Left            =   4260
         TabIndex        =   59
         Top             =   1455
         Width           =   2565
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
         Index           =   17
         Left            =   4260
         TabIndex        =   52
         Top             =   1170
         Width           =   2565
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
         Index           =   16
         Left            =   4260
         TabIndex        =   45
         Top             =   885
         Width           =   2565
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
         Index           =   15
         Left            =   4260
         TabIndex        =   38
         Top             =   600
         Width           =   2565
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V.:"
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
         Index           =   17
         Left            =   90
         TabIndex        =   55
         Top             =   1485
         Width           =   450
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "I.S.C.:"
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
         Index           =   18
         Left            =   90
         TabIndex        =   62
         Top             =   1770
         Width           =   420
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Otros:"
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
         Index           =   19
         Left            =   90
         TabIndex        =   69
         Top             =   2010
         Width           =   450
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Exportacion.:"
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
         Left            =   90
         TabIndex        =   41
         Top             =   915
         Width           =   945
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Exoneradas:"
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
         Index           =   16
         Left            =   90
         TabIndex        =   48
         Top             =   1185
         Width           =   915
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   14
         Left            =   90
         TabIndex        =   34
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
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
         Left            =   3420
         TabIndex        =   32
         Top             =   345
         Width           =   3285
      End
      Begin VB.Label lblTexto 
         Alignment       =   2  'Center
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
         Left            =   6960
         TabIndex        =   33
         Top             =   345
         Width           =   2265
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL:"
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
         Index           =   20
         Left            =   90
         TabIndex        =   76
         Top             =   2385
         Width           =   555
      End
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   615
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      _Version        =   393216
      Format          =   113901569
      CurrentDate     =   37102
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serie :"
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
      Left            =   7350
      TabIndex        =   25
      Top             =   2685
      Width           =   465
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   1905
      X2              =   2055
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diario :"
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
      Left            =   210
      TabIndex        =   27
      Top             =   3045
      Width           =   495
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
      Height          =   300
      Index           =   7
      Left            =   1665
      TabIndex        =   29
      Top             =   3015
      Width           =   5355
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
      Height          =   300
      Index           =   5
      Left            =   1425
      TabIndex        =   24
      Top             =   2655
      Width           =   5565
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Doc. :"
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
      Left            =   210
      TabIndex        =   22
      Top             =   2685
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   1155
      Left            =   60
      Top             =   2565
      Width           =   8475
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Operaci.:"
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
      Left            =   60
      TabIndex        =   19
      Top             =   2010
      Width           =   1005
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Operación:"
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
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   60
      X2              =   8500
      Y1              =   495
      Y2              =   495
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
      TabIndex        =   11
      Top             =   990
      Width           =   5970
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
      Index           =   6
      Left            =   3540
      TabIndex        =   14
      Top             =   1350
      Width           =   465
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Referencia:"
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
      Top             =   1350
      Width           =   840
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
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Vencimiento:"
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
      Left            =   4860
      TabIndex        =   7
      Top             =   660
      Width           =   1065
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Emisión:"
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
      Left            =   2610
      TabIndex        =   5
      Top             =   660
      Width           =   720
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   1110
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      TabIndex        =   9
      Top             =   1020
      Width           =   525
   End
End
Attribute VB_Name = "frmTRteVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbNuevo As Boolean
Private pbCorregir As Boolean
Private pbValidada As Boolean
Private pbFecha As Boolean

'[Propio del formulario.
Public unVerMonNac As Byte
Private Const MINIMOINDICEIMPORTE As Byte = 8, _
              MINIMOINDICEMAS As Byte = 1, _
              MINIMOINDICECUENTA As Byte = 15, _
              MINIMOINDICECCOSTO As Byte = 22, _
              CANTIDADIMPORTES As Byte = 7
'[Repetir en frmTVtaMasGrd.
Private Const DIFERENCIAMASIMPORTE As Byte = 7, _
              DIFERENCIAMASCUENTA As Byte = 14, _
              DIFERENCIAMASCCOSTO As Byte = 21
Private Const CUENTASCONCCOSTO As Byte = 7
']

'[Repetir en frmTrteVtaGrd y fmrTVtaMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
Private Const ps_OrdenCta As String = "01"
Private Sub cboRetencion_Click()
  If cboRetencion.Tag <> cboRetencion.ListIndex Then
    txtDato(4).Text = gsGloDoc_Rtc(cboRetencion.ListIndex)
  End If
  cboRetencion.Tag = cboRetencion.ListIndex
End Sub

Private Sub cmdFormato_Click()
  Dim sSQL As String
  Dim sImporteLetras As String, sSignoMoneda As String
  Dim nImporteTotal As Double, nImporteIgv As Double
  Dim nFormato As Integer, nRegistro As Integer, nContador As Integer
  Dim nDiferencia As Integer, nLen As Integer
  Dim porstRegistro As New ADODB.Recordset
  
  ' Inicializo las variables de impresion
  nImporteIgv = CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 9, 15)).Text)
  nImporteTotal = CDec(txtDato(Choose(cboTpoMon.ListIndex + 1, 11, 18)).Text)
  sImporteLetras = "SON : " & gfNumLet(nImporteTotal, Choose(cboTpoMon.ListIndex + 1, "N", "E"))
  
  sSQL = "SELECT " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(det.codtdc, det.serdoc, det.nrodoc)", "(det.codtdc+det.serdoc+det.nrodoc)") & " AS documento, det.orden AS secuencia, det.codtdc, "
  sSQL = sSQL & "det.serdoc, det.nrodoc, vta.feedoc AS emision, vta.fevdoc AS modifica, "
  sSQL = sSQL & "vta.codaux, aux.razaux, aux.diraux, aux.rucaux, vta.tpomon, vta.pctigv, vta.refdoc, "
  sSQL = sSQL & "det.glodet0, det.glodet1, vta.glodoc_rtc AS glortc, "
  sSQL = sSQL & "(CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN det.impcta_mn ELSE det.impcta_me END) AS impbase, "
  sSQL = sSQL & "(CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN vta.impigv_mn ELSE vta.impigv_me END) AS impigv, "
  sSQL = sSQL & "(CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN vta.imptot_mn ELSE vta.imptot_me END) AS imptotal, "
  sSQL = sSQL & "doc.dettdc, doc.forimp, '" & sImporteLetras & "' AS importeletra, "
  sSQL = sSQL & "(CASE vta.tpomon WHEN '" & TPOMON_NAC & "' THEN '" & gsTpoMon_Sgn_MN & "' ELSE '" & gsTpoMon_Sgn_ME & "' END) AS signomon "
  sSQL = sSQL & "FROM covtadoc vta "
  sSQL = sSQL & "LEFT JOIN CORteVtaCta det ON vta.codemp=det.codemp AND vta.pdoano=det.pdoano AND vta.codtdc=det.codtdc AND vta.serdoc=det.serdoc AND vta.nrodoc=det.nrodoc "
  sSQL = sSQL & "INNER JOIN tgaux aux ON vta.codemp=aux.codemp AND vta.codaux=aux.codaux "
  sSQL = sSQL & "INNER JOIN tgtdc doc ON vta.codemp=doc.codemp AND vta.codtdc=doc.codtdc "
  sSQL = sSQL & "WHERE vta.codemp='" & gsCodEmp & "' "
  sSQL = sSQL & "AND vta.pdoano='" & gsAnoAct & "' "
  sSQL = sSQL & "AND vta.mespvs='" & gsMesAct & "' "
  sSQL = sSQL & "AND vta.codtdc='" & txtLlave(0).Text & "' "
  sSQL = sSQL & "AND vta.serdoc='" & txtLlave(1).Text & "' "
  sSQL = sSQL & "AND vta.nrodoc='" & txtLlave(2).Text & "' "
  sSQL = sSQL & "AND det.tpocnc<= 3 "
  sSQL = sSQL & "ORDER BY det.tpocnc, det.orden"
  With porstRegistro
    If .State = adStateOpen Then .Close
    .ActiveConnection = frmTRteVtaGrd.uocnnMain
    .CursorLocation = adUseClient
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = sSQL
    .Open
  End With
  ' Verifico si se puede imprimir
  If porstRegistro.RecordCount = 0 Then MsgBox Choose(gsIdioma, "El documento no tiene detalle de impresión", "The document does not have impression detail"), vbCritical: Exit Sub
  nRegistro = CInt(porstRegistro.RecordCount)
  nFormato = CInt(porstRegistro!forimp)
  If (nFormato <= 0 Or nFormato >= 99) Then MsgBox Choose(gsIdioma, "No existe formato de impresión", "Not exist. format trint"), vbCritical: Exit Sub
  
  ' Genero la tabla temporal de reporte
  If ps_Plataforma = pSrvMySql Then
    frmTRteVtaGrd.uocnnMain.Execute "DROP TABLE IF EXISTS trptdocventa"
    sSQL = "CREATE TEMPORARY TABLE IF NOT EXISTS trptdocventa (documento varchar(16) NOT NULL, "
    sSQL = sSQL & "secuencia smallint(1) DEFAULT '0', codtdc char(2) NOT NULL, "
    sSQL = sSQL & "serdoc char(4) NOT NULL, nrodoc varchar(10) NOT NULL, "
    sSQL = sSQL & "emision date NULL, modifica date NULL, "
    sSQL = sSQL & "codaux varchar(11) NULL, razaux varchar(60) NULL, "
    sSQL = sSQL & "diraux varchar(80) NULL, rucaux varchar(11) NULL, "
    sSQL = sSQL & "tpomon char(1) NULL, signomon char(3) NULL, "
    sSQL = sSQL & "pctigv decimal(4,2) DEFAULT '0', refdoc varchar(20) NULL, "
    sSQL = sSQL & "glodet0 varchar(250) NULL, glodet1 varchar(250) NULL, "
    sSQL = sSQL & "glortc varchar(250) NULL, "
    sSQL = sSQL & "impbase decimal(12,2) DEFAULT '0', impigv decimal(12,2) DEFAULT '0', "
    sSQL = sSQL & "imptotal decimal(12,2) DEFAULT '0', dettdc varchar(40) NULL, "
    sSQL = sSQL & "forimp smallint(1) DEFAULT '0',  importeletra varchar(250) NULL, "
    sSQL = sSQL & "PRIMARY KEY (documento, secuencia))"
  ElseIf ps_Plataforma = pSrvSql Then
    frmTRteVtaGrd.uocnnMain.Execute "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptdocventa') DROP TABLE #trptdocventa"
    sSQL = "CREATE TABLE #trptdocventa (documento varchar(16) NOT NULL, "
    sSQL = sSQL & "secuencia smallint DEFAULT '0', codtdc char(2) NOT NULL, "
    sSQL = sSQL & "serdoc char(4) NOT NULL, nrodoc varchar(10) NOT NULL, "
    sSQL = sSQL & "emision smalldatetime NULL, modifica smalldatetime NULL, "
    sSQL = sSQL & "codaux varchar(11) NULL, razaux varchar(60) NULL, "
    sSQL = sSQL & "diraux varchar(80) NULL, rucaux varchar(11) NULL, "
    sSQL = sSQL & "tpomon char(1) NULL, signomon char(3) NULL, "
    sSQL = sSQL & "pctigv decimal(4,2) DEFAULT '0', refdoc varchar(20) NULL, "
    sSQL = sSQL & "glodet0 varchar(250) NULL, glodet1 varchar(250) NULL, "
    sSQL = sSQL & "glortc varchar(250) NULL, "
    sSQL = sSQL & "impbase decimal(12,2) DEFAULT '0', impigv decimal(12,2) DEFAULT '0', "
    sSQL = sSQL & "imptotal decimal(12,2) DEFAULT '0', dettdc varchar(40) NULL, "
    sSQL = sSQL & "forimp smallint DEFAULT '0',  importeletra varchar(250) NULL, "
    sSQL = sSQL & "PRIMARY KEY (documento, secuencia))"
  End If
  frmTRteVtaGrd.uocnnMain.Execute sSQL
  
  nRegistro = 0: nContador = 0
  ' Genero la informació de impresión
  While Not porstRegistro.EOF
    nDiferencia = ppNumeroLinea(IIf(IsNull(porstRegistro!glodet0), "", porstRegistro!glodet0) & IIf(IsNull(porstRegistro!glodet1), "", porstRegistro!glodet1))
    nContador = nContador + nDiferencia
    nRegistro = nRegistro + 1
    sSQL = "INSERT INTO " & ps_Prefijo & "trptdocventa "
    sSQL = sSQL & "(documento, secuencia, codtdc, serdoc, nrodoc, emision, modifica, codaux, razaux, "
    sSQL = sSQL & "diraux, rucaux, tpomon, signomon, pctigv, refdoc, glodet0, glodet1, glortc, "
    sSQL = sSQL & "impbase, impigv, imptotal, dettdc , forimp, importeletra) "
    sSQL = sSQL & "VALUES ('" & porstRegistro!documento & "', "
    sSQL = sSQL & nRegistro & ", "
    sSQL = sSQL & "'" & porstRegistro!CodTDc & "', "
    sSQL = sSQL & "'" & porstRegistro!serdoc & "', "
    sSQL = sSQL & "'" & porstRegistro!NroDoc & "', "
    If ps_Plataforma = pSrvMySql Then
      sSQL = sSQL & "DATE_FORMAT('" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
      sSQL = sSQL & "DATE_FORMAT('" & Format(porstRegistro!modifica, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
    Else
      sSQL = sSQL & "CONVERT(smalldatetime, '" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', 120), "
      sSQL = sSQL & "CONVERT(smalldatetime, '" & Format(porstRegistro!modifica, "yyyy-mm-dd") & "', 120), "
    End If
    sSQL = sSQL & "'" & porstRegistro!codaux & "', "
    sSQL = sSQL & "'" & porstRegistro!razAux & "', "
    sSQL = sSQL & "'" & porstRegistro!DirAux & "', "
    sSQL = sSQL & "'" & porstRegistro!rucaux & "', "
    sSQL = sSQL & "'" & porstRegistro!tpomon & "', "
    sSQL = sSQL & "'" & porstRegistro!signomon & "', "
    sSQL = sSQL & CDec(porstRegistro!PctIGV) & ", "
    sSQL = sSQL & "'" & porstRegistro!refdoc & "', "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glodet0), "Null", "'" & porstRegistro!glodet0 & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glodet1), "Null", "'" & porstRegistro!glodet1 & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glortc), "Null", "'" & porstRegistro!glortc & "'") & ", "
    sSQL = sSQL & CDec(porstRegistro!impbase) & ", "
    sSQL = sSQL & CDec(porstRegistro!impigv) & ", "
    sSQL = sSQL & CDec(porstRegistro!imptotal) & ", "
    sSQL = sSQL & "'" & porstRegistro!dettdc & "', "
    ' sSQL = sSQL & "'" & porstRegistro!forimp & "', "
    sSQL = sSQL & "'" & sImporteLetras & "')"
    frmTRteVtaGrd.uocnnMain.Execute sSQL
    porstRegistro.MoveNext
  Wend
  porstRegistro.MovePrevious
  
  ' Inserto los detalles adicionales
  nRegistro = nContador + 1
  For nContador = nRegistro To 7
    sSQL = "INSERT INTO " & ps_Prefijo & "trptdocventa "
    sSQL = sSQL & "(documento, secuencia, codtdc, serdoc, nrodoc, emision, modifica, codaux, razaux, "
    sSQL = sSQL & "diraux, rucaux, tpomon, signomon, pctigv, refdoc, glodet0, glodet1, glortc, "
    sSQL = sSQL & "impbase, impigv, imptotal, dettdc , forimp, importeletra) "
    sSQL = sSQL & "VALUES ('" & porstRegistro!documento & "', "
    sSQL = sSQL & nContador & ", "
    sSQL = sSQL & "'" & porstRegistro!CodTDc & "', "
    sSQL = sSQL & "'" & porstRegistro!serdoc & "', "
    sSQL = sSQL & "'" & porstRegistro!NroDoc & "', "
    If ps_Plataforma = pSrvMySql Then
      sSQL = sSQL & "DATE_FORMAT('" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
      sSQL = sSQL & "DATE_FORMAT('" & Format(porstRegistro!modifica, "yyyy-mm-dd") & "', '%Y-%m-%d'), "
    Else
      sSQL = sSQL & "CONVERT(smalldatetime, '" & Format(porstRegistro!emision, "yyyy-mm-dd") & "', 120), "
      sSQL = sSQL & "CONVERT(smalldatetime, '" & Format(porstRegistro!modifica, "yyyy-mm-dd") & "', 120), "
    End If
    sSQL = sSQL & "'" & porstRegistro!codaux & "', "
    sSQL = sSQL & "'" & porstRegistro!razAux & "', "
    sSQL = sSQL & "'" & porstRegistro!DirAux & "', "
    sSQL = sSQL & "'" & porstRegistro!rucaux & "', "
    sSQL = sSQL & "'" & porstRegistro!tpomon & "', "
    sSQL = sSQL & "'" & porstRegistro!signomon & "', "
    sSQL = sSQL & CDec(porstRegistro!PctIGV) & ", "
    sSQL = sSQL & "'" & porstRegistro!refdoc & "', "
    sSQL = sSQL & "Null" & ", "
    sSQL = sSQL & "Null" & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glortc), "Null", "'" & porstRegistro!glortc & "'") & ", "
    sSQL = sSQL & "0" & ", "
    sSQL = sSQL & CDec(porstRegistro!impigv) & ", "
    sSQL = sSQL & CDec(porstRegistro!imptotal) & ", "
    sSQL = sSQL & "'" & porstRegistro!dettdc & "', "
'    sSQL = sSQL & "'" & porstRegistro!forimp & "', "
    sSQL = sSQL & "'" & sImporteLetras & "')"
    frmTRteVtaGrd.uocnnMain.Execute sSQL
  Next nContador
  
  ' Obtengo los registrso de impresion
  With porstRegistro
    If .State = adStateOpen Then .Close
    .ActiveConnection = frmTRteVtaGrd.uocnnMain
    .CursorLocation = adUseClient
    .CursorType = adOpenForwardOnly
    .LockType = adLockReadOnly
    .Source = "SELECT * FROM " & ps_Prefijo & "trptdocventa ORDER BY documento, secuencia"
    .Open
  End With
  ' Realizo la impresion
  gpEncabezadoRpt frmMain.rptMain, Me.Caption, Date, True, False, porstRegistro
  With frmMain.rptMain
    '[Datos y parámetros del reporte
    .ReportFileName = gsRutRpt & "rptdocvcenta" & nFormato & ".rpt"
    .WindowState = crptMaximized
    .MarginLeft = 240
    .Destination = crptToWindow
    .Action = 1
  End With
  porstRegistro.Close
  Set porstRegistro = Nothing
  frmTRteVtaGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptdocventa", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptdocventa') DROP TABLE #trptdocventa")

End Sub

']
Private Sub Form_Load()
  pbValidada = False
  pbFecha = True
  Me.KeyPreview = True
   
  With frmTRteVtaGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
    txtLlave(0).MaxLength = .uorstMain!sernegocio.DefinedSize
    txtLlave(1).MaxLength = .uorstMain!nronegocio.DefinedSize
    ']
    
    '[Datos                            'Cambiar.
    With cboTpoMon
      .AddItem TPOMON_NAC_TXT_2, TPOMON_NAC_IND
      .AddItem TPOMON_EXT_TXT_2, TPOMON_EXT_IND
    End With
    With cboRetencion
      .AddItem Choose(gsIdioma, "Ninguna", "None"), TPOGRU1_IND
      .AddItem Choose(gsIdioma, "Sujeta a Detracción", "Holds to Deduction"), TPOGRU2_IND
      .AddItem Choose(gsIdioma, "No Sujeta a Detracción", "Not Holds to Deduction"), TPOGRU3_IND
    End With
    
    txtDato(0).MaxLength = .uorstMain!codaux.DefinedSize
    txtDato(1).MaxLength = .uorstMain!refdoc.DefinedSize
    txtDato(Choose(gsIdioma, 2, 3)).MaxLength = .uorstMain!GloDoc.DefinedSize
    txtDato(Choose(gsIdioma, 3, 2)).MaxLength = .uorstMain!glodocx.DefinedSize
    txtDato(4).MaxLength = .uorstMain!glodoc_rtc.DefinedSize
    txtDato(5).MaxLength = .uorstMain!CodTDc.DefinedSize
    txtDato(6).MaxLength = .uorstMain!serdoc.DefinedSize
    txtDato(7).MaxLength = .uorstMain!coddro.DefinedSize
    txtDato(8).MaxLength = 14
    txtDato(9).MaxLength = 14
    txtDato(10).MaxLength = 14
    txtDato(11).MaxLength = 14
    txtDato(12).MaxLength = 14
    txtDato(13).MaxLength = 14
    txtDato(14).MaxLength = 14
    txtDato(15).MaxLength = 8
    txtDato(16).MaxLength = 8
    txtDato(17).MaxLength = 8
    txtDato(18).MaxLength = 8
    txtDato(19).MaxLength = 8
    txtDato(20).MaxLength = 8
    txtDato(21).MaxLength = 8
    txtDato(22).MaxLength = 5
    txtDato(23).MaxLength = 5
    txtDato(24).MaxLength = 5
    txtDato(25).MaxLength = 5
    txtDato(26).MaxLength = 5
    txtDato(27).MaxLength = 5
    txtDato(28).MaxLength = 5
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
  ReDim aLabel(21, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Nº Documento :", "F.Operación:", "F.Emisión:", "F.Vencimiento:", "Cliente :", "Referencia:", "Glosa:", "Moneda:", "Tipo Operaci.:", "Tipo Doc.:", "Serie:", "Diario:", "Cuenta Contable", "Centro de Costo", "Op. Gravada :", "Exportación :", "Exonerado :", "IGV :", "ISC :", "Otros :", "Total :", "Tipo Tasa :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Nº Document :", "Operti.Date:", "IssueDate:", "Due Date:", "Client :", "Reference :", "Gloss:", "Currency:", "Type Doc.:", "Serie:", "Journal:", "Accountable Account", "Cost Center", "Op. with Taxes :", "Export :", "Discharged :", "GST :", "SCT :", "Others :", "Total :", "Rate Type :")
  Next nElemento
  chkCalcularIGV.Caption = Choose(gsIdioma, "Calcular I.G.&V.", "Calculate G.&S.T.")
  chkCalcularISC.Caption = Choose(gsIdioma, "Calcular I.S.&C.", "Calculate S.&C.T.")
  chkIndEstado.Caption = Choose(gsIdioma, "Contrato Activo", "Active Contract")
  cmdAuxiliar.Caption = Choose(gsIdioma, "Cliente", "Client")
  cmdFormato.Caption = Choose(gsIdioma, "&Imprimir", "&Print")
  sstMain.TabCaption(0) = Choose(gsIdioma, "I&mportes", "A&mounts")
  sstMain.TabCaption(1) = Choose(gsIdioma, "C&uentas", "Acco&unts")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']

'[Propio del formulario.
   dgrDetalle.MarqueeStyle = dbgHighlightRow
   Set dgrDetalle.DataSource = frmTRteVtaGrd.uorstCOCpbDet
   
   cboTpoMon.ListIndex = TPOMON_NAC_IND
   sstMain.Tab = 0
']
End Sub

Private Sub Form_Activate()
  If Not pbNuevo And cmdCorregir.Enabled Then
    cmdCorregir.SetFocus
  End If
  
  '[Propio del formulario.
  If Not pbNuevo Then
    dtpDato(0).Tag = dtpDato(0).Value
  End If
  ']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not frmTRteVtaGrd.uorstMain.EOF Then
    If frmTRteVtaGrd.uorstMain.EditMode <> adEditNone Then frmTRteVtaGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
  End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmTRteVtaGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTRteVtaGrd.uorstMain_Grd.MoveFirst
   frmTRteVtaGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmTRteVtaGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTRteVtaGrd.uorstMain_Grd.MoveFirst
   frmTRteVtaGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
End Sub

Public Sub cmdCorregir_Click()
  'Verificación de Mes Cerrado.
  If gbCieVta Then MsgBox TEXT_9016, vbCritical: Exit Sub
  
  pbCorregir = True
  frmTRteVtaGrd.uocnnMain.BeginTrans     'Cambiar Formulario de Grid. 'INICIA TRANSACCION.
  
  cmdRetroceder.Enabled = False
  cmdAvanzar.Enabled = False
  cmdCorregir.Enabled = False
  cmdFormato.Enabled = False
  cmdGrabar.Enabled = True
  cmdDeshacer.Enabled = True
  upHabilitacion True
  
  ' Dato con el foco al corregir
  dtpDato(0).SetFocus
  ' Para no cambiar fechas
  pbFecha = False
  
End Sub

Public Sub cmdGrabar_Click()
'   On Error GoTo Err

  '[Propio del formulario.
  Dim dnSumaMN As Double, dnSumaME As Double
  
  If txtDato(0).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(0).SetFocus: Exit Sub
  If txtDato(5).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(5).SetFocus: Exit Sub
  If txtDato(6).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(6).SetFocus: Exit Sub
  If txtDato(7).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(7).SetFocus: Exit Sub
   
  With frmTRteVtaGrd.uorstMain
    dnSumaMN = CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text) + CDec(txtDato(11).Text) + CDec(txtDato(12).Text) + CDec(txtDato(13).Text)
    If dnSumaMN <> CDec(txtDato(MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1).Text) Then
      If (cboTpoMon.ListIndex = TPOMON_EXT_IND And cmdMas(CANTIDADIMPORTES).Tag <> INDMASCTA_MAS And (Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1).Text)) <= 0.02)) Then
        If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1).Text)))) & "." & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
        txtDato(MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1).Text = Format(dnSumaMN, FORMATO_NUM_1)
      Else
        If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(11).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
      End If
    End If
  End With

   ' Valido las Cuentas esten Correctas(llenas para todas los valores)
  If Not ValidaCtasCCo Then
    If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
  End If
 ']

  With frmTRteVtaGrd                     'Cambiar Formulario de Grid.
    If pbNuevo And frmTRteVtaGrd.ubGrabaMas = 0 Then
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
    '[Propio del formulario.
'    ppGenera
    ']
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
    
    '[Actualiza grid..
    .uorstMain_Grd.Requery
    .upDatosGrid
    .uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & "'"
    ']
    pbCorregir = False
  
    If pbNuevo Then
      pbValidada = False
      cmdGrabar.Enabled = False
      upHabilitacion False
      frmTRteVtaGrd.ubGrabaMas = INDMASCTA_INI
    
      upDatosPredeterminados
      pbFecha = True
      '[Llave habilitar  'Cambiar.
      txtLlave(0).Enabled = True
      txtLlave(1).Enabled = True
      ']
      '[Llave con el foco al añadir.  'Cambiar.
      txtLlave(0).SetFocus
      ']
      .uocnnMain.BeginTrans            'INICIA TRANSACCION.
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
  
  frmTRteVtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
  '[Propio del formulario.
  frmTRteVtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
  pbCorregir = False
  ']
  cmdFormato.Enabled = True
  gpTUe_Deshacer Me
End Sub

Public Sub cmdSalir_Click()
  If pbNuevo Or pbCorregir Then
    pbCorregir = False
    frmTRteVtaGrd.uocnnMain.RollbackTrans 'RESTAURA TRANSACCION.
  End If
  Unload Me
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 5, 7, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
    txtDato(Index).SetFocus
  End Select
  ppAyuBus AYUDAT, Index
End Sub

Private Sub txtLlave_GotFocus(Index As Integer)
  txtLlave(Index).SelStart = 0
  txtLlave(Index).SelLength = txtLlave(Index).MaxLength
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

Private Sub txtLlave_LostFocus(Index As Integer)
  If pbValidada Then                  'Cambiar.
    txtLlave(0).Enabled = False
    txtLlave(1).Enabled = False
    If dtpDato(0).Enabled Then
      dtpDato(0).SetFocus
    End If
  End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
  Dim dvRegistro As Variant
  
  On Error GoTo Err
   
  '[Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
  Select Case Index
   Case 0, 1                        'Cambiar (añadir índices).
    If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
      txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
    End If
  End Select
  ']
 
  '[Valida la llave.                    'Cambiar.
  If (Len(Trim(txtLlave(0).Text)) <> 0 And Len(Trim(txtLlave(1).Text)) <> 0) Then
    With frmTRteVtaGrd                  'Cambiar Formulario de Grid.
      Set .uorstTemporal = .uocnnMain.Execute("SELECT pdoano FROM CoRteVta WHERE codemp='" & gsCodEmp & "' AND sernegocio='" & txtLlave(0).Text & "' AND nronegocio='" & txtLlave(1).Text & "'")
      If .uorstTemporal.RecordCount > 0 Then
        MsgBox TEXT_8007 & Chr(13), vbExclamation
        Cancel = True
        Exit Sub
      End If
      .uorstTemporal.Close
    End With
    '[Propio del formulario.
    If frmTRteVtaGrd.ubGrabaMas = 0 Then
      frmTRteVtaGrd.ubGrabaMas = 1
      With frmTRteVtaGrd
        If pbNuevo Then
          .uorstMain.AddNew
          .uorstMain!UsrCre = gsAbvUsr
          .uorstMain!FyHCre = Now
        End If
        upDatosDesconectados 0
        .uorstMain.Update
      End With
    End If
    ']
    cmdGrabar.Enabled = True
    upHabilitacion True
    pbValidada = True
  Else
    cmdGrabar.Enabled = False
    upHabilitacion False
    pbValidada = False
  End If
  ']

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
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    ppAyuBus AYUDAT, Index
  End If
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
''   Dim doColumna As Field
  Select Case Index
   Case MINIMOINDICEIMPORTE To MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1
    If CDec(txtDato(Index).Text) = 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
    End If
    If Index = MINIMOINDICEIMPORTE Then
      If chkCalcularIGV.Value Then txtDato(Index + 3).Text = Format(CDec(txtDato(Index).Text) * CDec(gnPctIGV) / 100, FORMATO_NUM_1)
      If chkCalcularISC.Value Then txtDato(Index + 4).Text = Format(CDec(txtDato(Index).Text) * CDec(gnPctISC) / 100, FORMATO_NUM_1)
    End If
      
    'Cálculo del total.
    If (Index = 14 And txtDato(Index).Text = 0) Then
      txtDato(Index).Text = Format(CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text) + CDec(txtDato(11).Text) + CDec(txtDato(12).Text) + CDec(txtDato(13).Text), FORMATO_NUM_1)
    End If
   Case MINIMOINDICECUENTA To MINIMOINDICECCOSTO - 1 'Cambiar (añadir índices).
    If txtDato(Index).Text = "" Then
      txtDato(Index + CUENTASCONCCOSTO).Text = ""
      lblDatoDeta(Index).Caption = ""
      lblDatoDeta(Index + CUENTASCONCCOSTO).Caption = ""
      txtDato(Index + CUENTASCONCCOSTO).Enabled = False
      lblDatoDeta(Index + CUENTASCONCCOSTO).Enabled = False
      cmdMas(Index - MINIMOINDICECUENTA + 1).Enabled = True
      If (Not pbNuevo) And cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_CTA Then
        If frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount > 0 Then
          ppAbreCtaCCo
          If frmTRteVtaGrd.uorstCoRteVtaCta.State = adStateOpen Then
            frmTRteVtaGrd.uorstCoRteVtaCta.MoveFirst
            Do
              If frmTRteVtaGrd.uorstCoRteVtaCta!sernegocio = txtLlave(0).Text And _
               frmTRteVtaGrd.uorstCoRteVtaCta!nronegocio = txtLlave(1).Text And _
               Trim(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc) = Trim(Str(Index - MINIMOINDICECUENTA + 1)) Then
                frmTRteVtaGrd.uorstCoRteVtaCta.Delete
              End If
              frmTRteVtaGrd.uorstCoRteVtaCta.MoveNext
            Loop Until frmTRteVtaGrd.uorstCoRteVtaCta.EOF
            frmTRteVtaGrd.uorstCoRteVtaCta.Requery
          End If
        End If
      End If
      cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI
    ElseIf cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI Then
      cmdMas(Index - MINIMOINDICECUENTA + 1).Enabled = False
    End If
  End Select
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  On Error GoTo Err
  
  'Completa con ceros a la izquierda.
  Select Case Index
   Case 6, MINIMOINDICECCOSTO To (MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1)
    If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
      txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
    End If
  End Select
  
  'Asigna 0 a campos numéricos si están vacíos.
   Select Case Index
    Case MINIMOINDICEIMPORTE To (MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1)
    If Not IsNumeric(txtDato(Index).Text) Then
      txtDato(Index).Text = 0
    End If
   End Select

  Select Case Index
   Case MINIMOINDICEIMPORTE To (MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1)
    txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
  End Select

  'Busca el dato en su tabla principal.
  Select Case Index
   Case 0, 5, 7, MINIMOINDICECUENTA To (MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1)
    Cancel = ppAyuDet(AYUDAT, Index)
    If Cancel Then Exit Sub
    If Index >= MINIMOINDICECUENTA And Index < MINIMOINDICECCOSTO Then
      If frmTRteVtaGrd.uorstCOCta.RecordCount > 0 Then
        If Not frmTRteVtaGrd.uorstCOCta.EOF Then
          If frmTRteVtaGrd.uorstCOCta!indcco = INDCCO_ACT Then
            ' Inicializo el centro de costos
            txtDato(Index + CUENTASCONCCOSTO).Tag = txtDato(Index + CUENTASCONCCOSTO).Text
            txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index).Tag <> txtDato(Index).Text, "", txtDato(Index + CUENTASCONCCOSTO).Text)
            If Not IsNull(frmTRteVtaGrd.uorstCOCta!codcco_def) Then
              txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index + CUENTASCONCCOSTO).Text = "", frmTRteVtaGrd.uorstCOCta!codcco_def, txtDato(Index + CUENTASCONCCOSTO).Text)
            Else
              txtDato(Index + CUENTASCONCCOSTO).Text = txtDato(Index + CUENTASCONCCOSTO).Tag
            End If
            txtDato(Index + CUENTASCONCCOSTO).Enabled = True
            cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = True
          Else
            txtDato(Index + CUENTASCONCCOSTO).Text = ""
            lblDatoDeta(Index + CUENTASCONCCOSTO).Caption = ""
            txtDato(Index + CUENTASCONCCOSTO).Enabled = False
            cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = False
          End If
        End If
        txtDato(Index).Tag = txtDato(Index).Text
      End If
    End If
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYUDAT Then
    Select Case tnIndex
     Case 0
      modAyuBus.Aux_Det "IndCli=1", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 5
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 7
      modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
    End Select
  End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
  If tsTipo = AYUDAT Then
    Select Case tnIndex                 'Cambiar.
     Case 0
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTRteVtaGrd.uorstTGAux
        If .RecordCount > 0 Then .MoveFirst
        .Find "codaux='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & !razAux
        End If
      End With
     Case 5
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTRteVtaGrd.uorstTGTDc
        If .RecordCount > 0 Then .MoveFirst
        .Find "codtdc='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!dettdc), "", !dettdc)
        End If
      End With
     Case 7
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTRteVtaGrd.uorstCODro
        If .RecordCount > 0 Then .MoveFirst
        .Find "coddro='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!DetDro), "", !DetDro)
        End If
      End With
     Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTRteVtaGrd.uorstCOCta
        If .RecordCount > 0 Then .MoveFirst
        .Find "CodCta='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & !detcta
        End If
      End With
     Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTRteVtaGrd.uorstCoCCo
        If .RecordCount > 0 Then .MoveFirst
        .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & !detcco
        End If
      End With
    End Select
  End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
  'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  Dim dnContador As Integer
  
  On Error GoTo Err

  With frmTRteVtaGrd.uorstMain           'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !sernegocio = txtLlave(0).Text
        !nronegocio = txtLlave(1).Text
        !PctIGV = CDec(gnPctIGV)
        !PctISC = CDec(gnPctISC)
      End If

      'Datos.
      !fehope = dtpDato(0).Value
      !feedoc = dtpDato(1).Value
      !fevdoc = dtpDato(2).Value
      !codaux = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      !refdoc = txtDato(1).Text
      !GloDoc = IIf(txtDato(Choose(gsIdioma, 2, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 2, 3)).Text)
      !glodocx = IIf(txtDato(Choose(gsIdioma, 3, 2)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 2)).Text)
      !tpomon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
      !TpoGlo_Rtc = cboRetencion.ListIndex
      !glodoc_rtc = IIf(txtDato(4).Text = "" Or cboRetencion.ListIndex = TPOGRU1_IND, Null, txtDato(4).Text)
      !CodTDc = IIf(txtDato(5).Text = "", Null, txtDato(5).Text)
      !serdoc = IIf(txtDato(6).Text = "", Null, txtDato(6).Text)
      !coddro = IIf(txtDato(7).Text = "", Null, txtDato(7).Text)
      !indestado = IIf(chkIndEstado.Value = vbChecked, ESTCCO_ACT, ESTCCO_INA)
      !impogr = CDec(txtDato(8).Text)
      !impexp = CDec(txtDato(9).Text)
      !impexo = CDec(txtDato(10).Text)
      !impigv = CDec(txtDato(11).Text)
      !impisc = CDec(txtDato(12).Text)
      !impoim = CDec(txtDato(13).Text)
      !imptot = CDec(txtDato(14).Text)
      
      '[Actualización de datos por Cuentas y Centros de Costo registrados directamente.
      ppAbreCtaCCo
      For dnContador = MINIMOINDICEMAS To (MINIMOINDICEMAS + CANTIDADIMPORTES - 1)
        If cmdMas(dnContador).Tag = INDMASCTA_CTA And Val(txtDato(dnContador + DIFERENCIAMASCUENTA).Text) <> 1 Then
          With frmTRteVtaGrd.uorstCoRteVtaCta
            .MoveFirst
            .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & dnContador & ps_OrdenCta & "'"
            If Not .EOF Then
              .Delete
              .Update
              .Requery
              frmTRteVtaGrd.uorstCoRteVtaCCo.Requery
              frmTRteVta.cmdMas(dnContador).Tag = INDMASCTA_INI
            End If
          End With
        End If
            
        If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
          cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Then
          With frmTRteVtaGrd.uorstCoRteVtaCta
            If .RecordCount <> 0 Then .MoveFirst
              .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & dnContador & ps_OrdenCta & "'"
              If .EOF Then
                .AddNew
                !codemp = gsCodEmp
                !pdoano = gsAnoAct
                !sernegocio = txtLlave(0).Text
                !nronegocio = txtLlave(1).Text
                !tpocnc = dnContador
                !orden = ps_OrdenCta
                !UsrCre = gsAbvUsr
                !FyHCre = Now
              Else
                !UsrMdf = gsAbvUsr
                !FyHMdf = Now
              End If
              !codcta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
              !glodet0 = txtDato(2).Text
              !porimpcta = CDec(100)
              .Update
          End With
          If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Or _
             cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Then
            With frmTRteVtaGrd.uorstCoRteVtaCCo
              If .RecordCount <> 0 Then .MoveFirst
              .Find "cLlave1='" & txtLlave(0).Text & txtLlave(1).Text & dnContador & ps_OrdenCta & txtDato(dnContador + DIFERENCIAMASCUENTA).Text & "'"
              If .EOF Then
                .AddNew
                !codemp = gsCodEmp
                !pdoano = gsAnoAct
                !sernegocio = txtLlave(0).Text
                !nronegocio = txtLlave(1).Text
                !tpocnc = dnContador
                !orden = ps_OrdenCta
                !codcta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
                !UsrCre = gsAbvUsr
                !FyHCre = Now
              Else
                !UsrMdf = gsAbvUsr
                !FyHMdf = Now
              End If
              !codcco = txtDato(dnContador + DIFERENCIAMASCCOSTO).Text
              !porimpcco = CDec(100)
              .Update
            End With
          End If
          frmTRteVta.cmdMas(dnContador).Tag = INDMASCTA_CTA
        End If
      Next
      ']
    Else
      'Llaves.
      txtLlave(0).Text = !sernegocio
      txtLlave(1).Text = !nronegocio
      
      'Datos.
      dtpDato(0).Value = !fehope
      dtpDato(1).Value = !feedoc
      dtpDato(2).Value = !fevdoc
      txtDato(0).Text = IIf(IsNull(!codaux), "", !codaux)
      txtDato(1).Text = IIf(IsNull(!refdoc), "", !refdoc)
      txtDato(Choose(gsIdioma, 2, 3)).Text = IIf(IsNull(!GloDoc), "", !GloDoc)
      txtDato(Choose(gsIdioma, 3, 2)).Text = IIf(IsNull(!glodocx), "", !glodocx)
      cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      
      txtDato(5).Text = IIf(IsNull(!CodTDc), "", !CodTDc)
      txtDato(6).Text = IIf(IsNull(!serdoc), "", !serdoc)
      txtDato(7).Text = IIf(IsNull(!coddro), "", !coddro)
      chkIndEstado.Value = IIf(!indestado = ESTCCO_ACT, vbChecked, vbUnchecked)
      txtDato(8).Text = Format(!impogr, FORMATO_NUM_1)
      txtDato(9).Text = Format(!impexp, FORMATO_NUM_1)
      txtDato(10).Text = Format(!impexo, FORMATO_NUM_1)
      txtDato(11).Text = Format(!impigv, FORMATO_NUM_1)
      txtDato(12).Text = Format(!impisc, FORMATO_NUM_1)
      txtDato(13).Text = Format(!impoim, FORMATO_NUM_1)
      txtDato(14).Text = Format(!imptot, FORMATO_NUM_1)
      For dnContador = MINIMOINDICECUENTA To (MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1)
        txtDato(dnContador).Text = ""
        txtDato(dnContador).Tag = ""
      Next dnContador
      cboRetencion.Tag = !TpoGlo_Rtc
      cboRetencion.ListIndex = !TpoGlo_Rtc
      txtDato(4).Text = IIf(IsNull(!glodoc_rtc), "", !glodoc_rtc)

      '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      ppAyuDet AYUDAT, 0
      ppAyuDet AYUDAT, 5
      ppAyuDet AYUDAT, 7
      For dnContador = MINIMOINDICECUENTA To (MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1)
        ppAyuDet AYUDAT, dnContador
      Next dnContador
      ']
      
      '[Propio del formulario.
      For dnContador = MINIMOINDICEMAS To CANTIDADIMPORTES
        cmdMas(dnContador).Tag = INDMASCTA_MAS
      Next dnContador
        
      ' MA Obtengo las cuenta y centro de costo
      ppAbreCtaCCo
      With frmTRteVtaGrd.uorstCoRteVtaCta
        dnContador = 0
        While Not .EOF
          ' Cuenta por importe
          If dnContador <> CByte(!tpocnc) Then
            dnContador = CByte(!tpocnc)
            txtDato(dnContador + DIFERENCIAMASCUENTA).Text = !codcta
            txtDato(dnContador + DIFERENCIAMASCUENTA).Tag = !codcta
            ppAyuDet AYUDAT, (dnContador + DIFERENCIAMASCUENTA)
            With frmTRteVtaGrd.uorstCoRteVtaCCo
              If .RecordCount > 0 Then
                .MoveFirst
                .Find "cLlave = " & dnContador & ps_OrdenCta & frmTRteVtaGrd.uorstCoRteVtaCta!codcta
                If Not .EOF Then
                  txtDato(dnContador + DIFERENCIAMASCCOSTO).Text = !codcco
                  ppAyuDet AYUDAT, (dnContador + DIFERENCIAMASCCOSTO)
                End If
              End If
            End With
          End If
          .MoveNext
        Wend
      End With
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
  txtLlave(0).Text = ""
  txtLlave(1).Text = ""

  'Datos.
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  dtpDato(0).Value = Date
  dtpDato(1).Value = Date
  dtpDato(2).Value = Date
  For dnContador = 0 To 7
    txtDato(dnContador).Text = ""
  Next dnContador
  For dnContador = MINIMOINDICEIMPORTE To (MINIMOINDICEIMPORTE + CANTIDADIMPORTES - 1)
    txtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
  Next
  For dnContador = MINIMOINDICECUENTA To (MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1)
    txtDato(dnContador).Text = ""
    txtDato(dnContador).Tag = ""
  Next
  cboRetencion.Tag = gsTpoGlo_Rtc
  cboRetencion.ListIndex = gsTpoGlo_Rtc
  txtDato(4).Text = gsGloDoc_Rtc(gsTpoGlo_Rtc)
  chkIndEstado.Value = vbChecked
   
  '[Propio del formulario.
  For dnContador = MINIMOINDICEMAS To (MINIMOINDICEMAS + CANTIDADIMPORTES - 1)
    cmdMas(dnContador).Tag = INDMASCTA_MAS
  Next
  ']
  'Ayudas.
  lblDatoDeta(0).Caption = ""
  For dnContador = MINIMOINDICECUENTA To (MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1)
    lblDatoDeta(dnContador).Caption = ""
  Next
  lblDatoDeta(5).Caption = ""
  lblDatoDeta(7).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

  'Datos.
  cboTpoMon.Enabled = tbHabilitar
  chkCalcularIGV.Enabled = tbHabilitar
  chkCalcularISC.Enabled = tbHabilitar
  chkIndEstado.Enabled = tbHabilitar
  dtpDato(0).Enabled = tbHabilitar
  dtpDato(1).Enabled = tbHabilitar
  dtpDato(2).Enabled = tbHabilitar
  cboRetencion.Enabled = tbHabilitar
   With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next dnContador
   End With
   
  'Ayudas.
  cmdDatoAyud(0).Enabled = tbHabilitar
  lblDatoDeta(0).Enabled = tbHabilitar
  cmdDatoAyud(5).Enabled = tbHabilitar
  lblDatoDeta(5).Enabled = tbHabilitar
  cmdDatoAyud(7).Enabled = tbHabilitar
  lblDatoDeta(7).Enabled = tbHabilitar
  
  If tbHabilitar Then
    For dnContador = MINIMOINDICEMAS To CANTIDADIMPORTES
      cmdMas(dnContador).Enabled = Not (cmdMas(dnContador).Tag = INDMASCTA_CTA)
      Call upHabilitaCuenta((Not cmdMas(dnContador).Tag = INDMASCTA_MAS), dnContador)
    Next dnContador
  Else
    For dnContador = MINIMOINDICEMAS To CANTIDADIMPORTES
      cmdMas(dnContador).Enabled = False
      Call upHabilitaCuenta(False, dnContador)
      Call upHabilitaCCosto(False, dnContador)
    Next dnContador
  End If
  ']
End Sub

Private Sub cmdAuxiliar_Click()
  frmMAuxGrd.Show vbModal
  frmTRteVtaGrd.uorstTGAux.Requery
End Sub

Private Sub cmdMas_Click(Index As Integer) 'Cambiar Formulario de Grid.
  frmTRteVtaMasGrd.unIndice = Index
  frmTRteVtaMasGrd.Show vbModal
  ppAbreCtaCCo
  ppAyuDet AYUDAT, Index + DIFERENCIAMASCUENTA
  If Index <= CUENTASCONCCOSTO Then
    ppAyuDet AYUDAT, Index + DIFERENCIAMASCCOSTO
  End If
End Sub

Private Sub dtpDato_Validate(Index As Integer, Cancel As Boolean)
Dim dnContador As Byte
   If Index = 0 Then
      If Month(dtpDato(0).Value) > Val(gsMesAct) And Year(dtpDato(0).Value) >= Val(gsAnoAct) Then
         MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
         dtpDato(Index).SetFocus
         Cancel = True
         Exit Sub
      End If
      dtpDato(0).Tag = 0
      If (dtpDato(0).Tag <> dtpDato(0).Value) Then
         dtpDato(0).Tag = dtpDato(0).Value
      End If
      If pbFecha Then
         With dtpDato
            For dnContador = 0 To .Count - 2
               .Item(dnContador).Value = dtpDato(Index).Value
            Next
         End With
         pbFecha = False
      End If
   End If
   If Index = 3 Then
      If Month(dtpDato(3).Value) <> Val(gsMesAct) Or Year(dtpDato(3).Value) <> Val(gsAnoAct) Then
         MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
         dtpDato(Index).SetFocus
         Cancel = True
         Exit Sub
      End If
      If txtDato(4).Text = 0 Then
         If pbFecha Then
            With dtpDato
            For dnContador = 0 To .Count - 2
               .Item(dnContador).Value = dtpDato(Index).Value
            Next
            End With
            pbFecha = False
         End If
      Else
         If pbFecha Then
            With dtpDato
            For dnContador = 1 To .Count - 2
               .Item(dnContador).Value = dtpDato(Index).Value
            Next
            End With
            pbFecha = False
         End If
      End If
   
   End If
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
  If (PreviousTab = 0 And sstMain.Tab = 1) Then
    ppDatosWhere
  End If
  dgrDetalle.SetFocus
End Sub

Private Sub ppAbreCtaCCo()
   With frmTRteVtaGrd.uorstCoRteVtaCCo
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = "WHERE cortevtacco.codemp='" & frmTRteVtaGrd.uorstMain!codemp & "' "
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND cortevtacco.sernegocio='" & frmTRteVtaGrd.uorstMain!sernegocio & "' "
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & "AND cortevtacco.nronegocio='" & frmTRteVtaGrd.uorstMain!nronegocio & "' "
    If .State = adStateOpen Then .Close
    .Source = frmTRteVtaGrd.usConnStrgSele_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCCo
    .Open
    .Properties("Unique Table").Value = "CoRteVtaCCo"
   End With
   With frmTRteVtaGrd.uorstCoRteVtaCta
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta = "WHERE CORteVtaCta.codemp='" & frmTRteVtaGrd.uorstMain!codemp & "' "
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta & "AND CORteVtaCta.sernegocio='" & frmTRteVtaGrd.uorstMain!sernegocio & "' "
    frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta = frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta & "AND CORteVtaCta.nronegocio='" & frmTRteVtaGrd.uorstMain!nronegocio & "' "
    If .State = adStateOpen Then .Close
    .Source = frmTRteVtaGrd.usConnStrgSele_CoRteVtaCta & frmTRteVtaGrd.usConnStrgWher_CoRteVtaCta & frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCta
    .Open
    .Properties("Unique Table").Value = "CoRteVtaCta"
   End With
End Sub

Private Sub ppGenera()
'   Dim dnContador As Integer
'   Dim dnNumeroItem As Integer
'   Dim dbProcesaCuenta As Boolean
'
'   Dim sSentencia As String
'   Dim siexiste As Boolean
'   Dim masdedos As Boolean
'   Dim cuenta As Integer
'
'    siexiste = False
'  masdedos = False
'
'  'Aqui Esta
'  sSentencia = "SELECT coddro, nrocpb"
'  sSentencia = sSentencia & " FROM cocpbcab  "
'  sSentencia = sSentencia & " WHERE codemp='" & gsCodEmp & "' "
'  sSentencia = sSentencia & " AND pdoano='" & gsAnoAct & "' "
'  sSentencia = sSentencia & " AND mespvs='" & gsMesAct & "'"
'  sSentencia = sSentencia & " AND coddro='" & txtDato(0).Text & "'"
'  sSentencia = sSentencia & " AND nrocpb='" & txtDato(1).Text & "'"
'  Set frmTRteVtaGrd.uorstTemporal = frmTRteVtaGrd.uocnnMain.Execute(sSentencia)
'  If Not (frmTRteVtaGrd.uorstTemporal.BOF Or frmTRteVtaGrd.uorstTemporal.EOF) And frmTRteVtaGrd.uorstTemporal.RecordCount > 0 Then
'    While Not frmTRteVtaGrd.uorstTemporal.EOF
'      siexiste = True
'      frmTRteVtaGrd.uorstTemporal.MoveNext
'    Wend
'  Else
'  End If
'  frmTRteVtaGrd.uorstTemporal.Close
'
'  cuenta = 0
'  sSentencia = "SELECT coddro,nrocpb "
'  sSentencia = sSentencia & " FROM covtadoc  "
'  sSentencia = sSentencia & " WHERE codemp='" & gsCodEmp & "' "
'  sSentencia = sSentencia & " AND pdoano='" & gsAnoAct & "' "
'  sSentencia = sSentencia & " AND mespvs='" & gsMesAct & "'"
'  sSentencia = sSentencia & " AND coddro='" & txtDato(0).Text & "'"
'  sSentencia = sSentencia & " AND nrocpb='" & txtDato(1).Text & "'"
'  Set frmTRteVtaGrd.uorstTemporal = frmTRteVtaGrd.uocnnMain.Execute(sSentencia)
'  If Not (frmTRteVtaGrd.uorstTemporal.BOF Or frmTRteVtaGrd.uorstTemporal.EOF) And frmTRteVtaGrd.uorstTemporal.RecordCount > 0 Then
'    While Not frmTRteVtaGrd.uorstTemporal.EOF
'      cuenta = cuenta + 1
'      frmTRteVtaGrd.uorstTemporal.MoveNext
'    Wend
'  Else
'  End If
'  frmTRteVtaGrd.uorstTemporal.Close
'
'  If cuenta >= 2 Then masdedos = True
'
'
'   ' MA 26-08-2011 / Tipo documento clave
'   frmTRteVtaGrd.uorstTGTDc.MoveFirst
'   frmTRteVtaGrd.uorstTGTDc.Find "codtdc='" & txtLlave(0).Text & "'"
'
'   With frmTRteVtaGrd.uorstCOCpbCab
'     'Captura del Siguiente Número.
'      If txtDato(1).Text = "" Then
'        txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
'        txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
'        txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
'        txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
'        txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
'        frmTRteVtaGrd.uocnnMain.Execute txtDato(1).Tag
'        ' Actualizo numero de comprobante tabla de detalle
'        frmTRteVtaGrd.uorstMain!NroCpb = txtDato(1).Text
'        frmTRteVtaGrd.uorstMain.Update
'       Else
'       If masdedos = True Then
'        txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
'        txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
'        txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
'        txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
'        txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
'        frmTRteVtaGrd.uocnnMain.Execute txtDato(1).Tag
'        ' Actualizo numero de comprobante tabla de detalle
'        frmTRteVtaGrd.uorstMain!NroCpb = txtDato(1).Text
'        frmTRteVtaGrd.uorstMain.Update
'        End If
'    End If
'
'      ppDatosWhere
'
'     'Si no hay cuentas, marca el documento como no generado.
'      If frmTRteVtaGrd.uorstCoRteVtaCta.RecordCount = 0 Then
'         frmTRteVtaGrd.uorstMain!indgen = False
'         frmTRteVtaGrd.uorstMain.Update
'         Exit Sub
'      End If
'
'     'Crea encabezado de Comprobante.
'      .AddNew
'      !codemp = gsCodEmp
'      !pdoano = gsAnoAct
'      !mespvs = gsMesAct
'      !coddro = txtDato(0).Text
'      !NroCpb = txtDato(1).Text
'      !FehCpb = dtpDato(3).Value
'      !tpognr = TPOGNR_VTA
'      !IndNCu = INDNCU_FAL
'      !glocpb = IIf(txtDato(Choose(gsIdioma, 3, 37)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 37)).Text)
'      !glocpbx = IIf(txtDato(Choose(gsIdioma, 37, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 37, 3)).Text)
'      !UsrCre = gsAbvUsr
'      !FyHCre = Now
'   End With
''[ Teo, Miguel Angel Refresco los recordset de cuentas y centros de costos
'  frmTRteVtaGrd.uorstCoRteVtaCta.Requery
'  frmTRteVtaGrd.uorstCoRteVtaCCo.Requery
'']
'   With frmTRteVtaGrd.uorstCoRteVtaCta
'     'Crea ítemes de Comprobante.
'      .MoveFirst
'      Do
'         dbProcesaCuenta = True
'
'        'Itemes con Centro de Costo.
'         If !tpocnc <= CUENTASCONCCOSTO Then
'            With frmTRteVtaGrd.uorstCoRteVtaCCo
'               If .RecordCount <> 0 Then
'                  .MoveFirst
'                  .Find "cLlave = " & Trim(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc) & frmTRteVtaGrd.uorstCoRteVtaCta!orden & frmTRteVtaGrd.uorstCoRteVtaCta!codcta
'                  If Not .EOF Then
'                     Do
'                        dnNumeroItem = dnNumeroItem + 1
'                        ppGenera1 True, dnNumeroItem, IIf(CInt(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc) >= 4, "", txtDato(40).Text)
'                        .MoveNext
'                        If .EOF Then Exit Do
'                        If !cLlave <> Trim(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc) & frmTRteVtaGrd.uorstCoRteVtaCta!orden & frmTRteVtaGrd.uorstCoRteVtaCta!codcta Then Exit Do
'                     Loop
'                     dbProcesaCuenta = False
'                  End If
'               End If
'            End With
'         End If
'
'        'Itemes sin Centro de Costo.
'         If dbProcesaCuenta Then
'            dnNumeroItem = dnNumeroItem + 1
'            ppGenera1 False, dnNumeroItem, IIf(CInt(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc) >= 4, "", txtDato(40).Text)
'         End If
'         .MoveNext
'      Loop Until .EOF
'   End With
'
'   frmTRteVtaGrd.uorstMain!indgen = True
'   txtDato(0).Enabled = False
'   txtDato(1).Enabled = False
'   cmdDatoAyud(0).Enabled = False
'   lblDatoDeta(0).Enabled = False
'
'   frmTRteVtaGrd.uorstCOCpbCab.Update
'   frmTRteVtaGrd.uorstCOCpbDet.UpdateBatch
'   frmTRteVtaGrd.uorstMain.Update
End Sub

Private Sub ppGenera1(tbCCosto As Boolean, tnNumeroItem As Integer, ByVal sContrato As String)
   
  With frmTRteVtaGrd.uorstCOCpbDet
    .AddNew
    !codemp = gsCodEmp
    !pdoano = gsAnoAct
    !coddro = txtDato(0).Text
    !NroCpb = txtDato(1).Text
    !NroIte = tnNumeroItem
    !mespvs = gsMesAct
    !codcta = frmTRteVtaGrd.uorstCoRteVtaCta!codcta
    !fehope = dtpDato(3).Value
    frmTRteVtaGrd.uorstCOCta.MoveFirst
    frmTRteVtaGrd.uorstCOCta.Find "CodCta='" & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & "'"
    If frmTRteVtaGrd.uorstCOCta!indcco = INDCCO_ACT Then If tbCCosto Then !codcco = frmTRteVtaGrd.uorstCoRteVtaCCo!codcco
    If frmTRteVtaGrd.uorstCOCta!IndDoc = INDDOC_ACT Then
      !codaux = txtDato(33).Text
    Else
      If Len(Trim(frmTRteVtaGrd.uorstCoRteVtaCta!codruc)) > 0 Then
        !codaux = frmTRteVtaGrd.uorstCoRteVtaCta!codruc
      Else
        !codaux = txtDato(33).Text
      End If
    End If
    !CodTDc = txtLlave(0).Text
    !serdoc = txtLlave(1).Text
    !NroDoc = txtLlave(2).Text
    !feedoc = dtpDato(0).Value
    !fevdoc = dtpDato(1).Value
    !ferdoc = dtpDato(0).Value
    !refdoc = txtDato(2).Text
    !GloIte = Left(Trim(frmTRteVtaGrd.uorstCoRteVtaCta!glodet0), 60)
    !GloItex = Left(Trim(frmTRteVtaGrd.uorstCoRteVtaCta!glodet0x), 60)
    !codcon = IIf(sContrato = "", Null, sContrato)
    If tbCCosto Then
      If (frmTRteVtaGrd.uorstCoRteVtaCCo!impcco_me > 0) And (frmTRteVtaGrd.uorstCoRteVtaCCo!impcco_mn > 0) Then
        !TpoCtb = IIf(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        !TpoCtb = IIf(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
    Else
      If (frmTRteVtaGrd.uorstCoRteVtaCta!impcta_me > 0) And (frmTRteVtaGrd.uorstCoRteVtaCta!impcta_mn > 0) Then
        !TpoCtb = IIf(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        !TpoCtb = IIf(frmTRteVtaGrd.uorstCoRteVtaCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
    End If
    !tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
    !ImpTCb = CDec(txtDato(4).Text)
    If tbCCosto Then
      !ImpMN = CDec(Abs(frmTRteVtaGrd.uorstCoRteVtaCCo!impcco_mn))
      !ImpME = CDec(Abs(frmTRteVtaGrd.uorstCoRteVtaCCo!impcco_me))
    Else
      !ImpMN = CDec(Abs(frmTRteVtaGrd.uorstCoRteVtaCta!impcta_mn))
      !ImpME = CDec(Abs(frmTRteVtaGrd.uorstCoRteVtaCta!impcta_me))
    End If
    'modificado tc
'    !TpoPvs = IIf(frmTRteVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG And cboCategoria.ListIndex >= CategoriaDocumento.RetencionIva, TPOPVS_PVS, TPOPVS_PVS)
    !tpognr = TPOGNR_VTA
    !UsrCre = gsAbvUsr
    !FyHCre = Now
  End With

End Sub

Public Sub upHabilitaCuenta(tbHabilita As Boolean, tnIndice As Byte)
  txtDato(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
  lblDatoDeta(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
  If tnIndice <= CUENTASCONCCOSTO Then
    Call upHabilitaCCosto(tbHabilita, tnIndice)
  End If
End Sub

Public Sub upHabilitaCCosto(tbHabilita As Boolean, tnIndice As Byte)
  If Not tbHabilita Or txtDato(tnIndice + DIFERENCIAMASCUENTA).Text = "" Then
    txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
    lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
  Else
    frmTRteVtaGrd.uorstCOCta.MoveFirst
    frmTRteVtaGrd.uorstCOCta.Find "CodCta='" & txtDato(tnIndice + DIFERENCIAMASCUENTA).Text & "'"
    txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTRteVtaGrd.uorstCOCta!indcco = INDCCO_ACT)
    lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTRteVtaGrd.uorstCOCta!indcco = INDCCO_ACT)
  End If
End Sub

Private Sub ppDatosWhere()             'Cambiar.
  With frmTRteVtaGrd
    .usConnStrgWher_COCpbDet = "WHERE cta.codemp='" & gsCodEmp & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND cta.sernegocio='" & txtLlave(0).Text & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND cta.nronegocio='" & txtLlave(1).Text & "' "
    With .uorstCOCpbDet
      .Close
      .Source = frmTRteVtaGrd.usConnStrgSele_COCpbDet & frmTRteVtaGrd.usConnStrgWher_COCpbDet & frmTRteVtaGrd.usConnStrgOrde_COCpbDet
      .Open
    End With
    Set dgrDetalle.DataSource = .uorstCOCpbDet
  End With
  ppDatosGrid
End Sub

Public Sub ppDatosGrid()               'Cambiar Datos Grid.
  Dim dnNum As Integer
  
  With dgrDetalle.Columns
    For dnNum = 0 To .Count - 1
      Select Case dnNum
       Case 0
        .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
        .Item(dnNum).Width = 1000
       Case 1
        .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
        .Item(dnNum).Width = 1150
       Case 2
        .Item(dnNum).Caption = Choose(gsIdioma, "C.Cto.", "C.Center")
        .Item(dnNum).Width = 800
       Case 3
        .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
        .Item(dnNum).Width = 2500
       Case 4
        .Item(dnNum).Caption = Choose(gsIdioma, "Mon", "Cur")
        .Item(dnNum).Width = 400
        .Item(dnNum).Alignment = dbgCenter
       Case 5
        .Item(dnNum).Caption = Choose(gsIdioma, "Debe", "Debit")
        .Item(dnNum).Width = 1400
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 6
        .Item(dnNum).Caption = Choose(gsIdioma, "Haber", "Credit")
        .Item(dnNum).Width = 1400
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case Else
        .Item(dnNum).Visible = False
      End Select
    Next dnNum
  End With
End Sub

Private Function ValidaCtasCCo() As Boolean
   Dim dnContador, dnIndCCo As Byte
   Dim dvRegistroActual As Variant
   Dim dnTotalCuentaMN, dnTotalCuentaME, dnTotalImporteMN, dnTotalImporteME As Double
    
  ValidaCtasCCo = True
  
  For dnContador = INDMASCTA_INI To (CANTIDADIMPORTES - 1)
    If (CDec(txtDato(MINIMOINDICEIMPORTE + dnContador).Text) <> 0) And _
      (cmdMas(dnContador + 1).Tag = INDMASCTA_MAS And Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
      ValidaCtasCCo = Not (txtDato(MINIMOINDICECUENTA + dnContador).Text = "")
      If Not ValidaCtasCCo Then Exit Function
      
      If frmTRteVtaGrd.ubGrabaMas = INDMASCTA_MAS Then
        With frmTRteVtaGrd.uorstCoRteVtaCta
          If Not (.BOF Or .EOF) And .RecordCount > 0 Then
            dvRegistroActual = .Bookmark
            dnTotalCuentaMN = 0
            dnTotalCuentaME = 0
            .MoveFirst
            Do
              dnIndCCo = 0
              If Trim(!tpocnc) = Trim(Str(dnContador + 1)) Then
                dnTotalCuentaMN = dnTotalCuentaMN + !porimpcta
                With frmTRteVtaGrd.uorstCOCta
                  .MoveFirst
                  .Find "CodCta='" & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & "'"
                  If Not .EOF Then
                    dnIndCCo = frmTRteVtaGrd.uorstCOCta!indcco
                  End If
                End With
              End If
              If dnIndCCo = INDCCO_ACT Then
                With frmTRteVtaGrd.uorstCoRteVtaCCo
                  If .State = adStateOpen Then .Close
                  frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo = "WHERE cortevtacco.sernegocio='" & frmTRteVtaGrd.uorstMain!sernegocio & "' And cortevtacco.nronegocio='" & frmTRteVtaGrd.uorstMain!nronegocio & "' And cortevtacco.TpoCnc='" & Trim(Str(dnContador + 1)) & "' And cortevtacco.codcta='" & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & "' "
                  .Source = frmTRteVtaGrd.usConnStrgSele_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgWher_CoRteVtaCCo & frmTRteVtaGrd.usConnStrgOrde_CoRteVtaCCo
                  .Open
                  If .RecordCount = 0 Then
                    MsgBox Choose(gsIdioma, "Cuenta ", "Account ") & frmTRteVtaGrd.uorstCoRteVtaCta!codcta & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
                    ValidaCtasCCo = False
                    .Close
                    Exit Function
                  End If
                  .Close
                End With
              End If
              .MoveNext
            Loop Until .EOF
            .Bookmark = dvRegistroActual
          End If
        End With
        dnTotalImporteMN = gfRedond(CDec(txtDato(MINIMOINDICEIMPORTE + dnContador).Text), 2)
        dnTotalCuentaME = Round(dnTotalCuentaMN / 100, 2)
        dnTotalCuentaMN = Round(dnTotalImporteMN * dnTotalCuentaME, 2)
        If Not (CDec(dnTotalCuentaMN) = CDec(dnTotalImporteMN)) Then
          ValidaCtasCCo = False
          Exit Function
        End If
      End If
    ElseIf Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0 Then
      dnIndCCo = 0
      With frmTRteVtaGrd.uorstCOCta
        .MoveFirst
        .Find "CodCta='" & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & "'"
        If Not .EOF Then
          dnIndCCo = frmTRteVtaGrd.uorstCOCta!indcco
        End If
      End With
      If dnIndCCo = INDCCO_ACT And txtDato(dnContador + 1 + DIFERENCIAMASCUENTA + CUENTASCONCCOSTO).Text = "" Then
        MsgBox Choose(gsIdioma, "Cuenta ", "Account ") & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
        ValidaCtasCCo = False
        Exit Function
      End If
    ElseIf Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) = 0 And ((CDec(txtDato(MINIMOINDICEIMPORTE + dnContador).Text) <> 0)) Then
      ValidaCtasCCo = False
    End If
  Next dnContador
    
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
