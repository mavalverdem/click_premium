VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTVta 
   Caption         =   "[Título]"
   ClientHeight    =   6390
   ClientLeft      =   8310
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "Cli&ente"
      Height          =   375
      Left            =   8250
      TabIndex        =   124
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   119
      Top             =   2520
      Width           =   7155
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
         Left            =   6315
         TabIndex        =   14
         Top             =   180
         Width           =   735
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
         Left            =   600
         TabIndex        =   13
         Top             =   180
         Width           =   555
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4860
         Picture         =   "frmTVta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Diario:"
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
         Left            =   120
         TabIndex        =   123
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Comprobante:"
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
         Left            =   5265
         TabIndex        =   122
         Top             =   240
         Width           =   1005
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
         Index           =   0
         Left            =   1140
         TabIndex        =   121
         Top             =   180
         Width           =   3735
      End
   End
   Begin VB.CheckBox chkIndPreGen 
      Caption         =   "Cuentas &Registradas"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   6960
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3540
      Width           =   1815
   End
   Begin VB.CheckBox chkCalcularISC 
      Caption         =   "Calcular I.&S.C."
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   1500
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3180
      Width           =   1335
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
      Index           =   34
      Left            =   6840
      TabIndex        =   3
      Top             =   600
      Width           =   435
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
      Index           =   35
      Left            =   7260
      TabIndex        =   4
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   8640
      ScaleHeight     =   2610
      ScaleWidth      =   885
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   120
      Width           =   885
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
         Left            =   98
         Picture         =   "frmTVta.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1990
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
         Picture         =   "frmTVta.frx":02F4
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1440
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
         Picture         =   "frmTVta.frx":03F6
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   880
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
         Picture         =   "frmTVta.frx":04F8
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   325
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
         Picture         =   "frmTVta.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   55
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
         Left            =   458
         Picture         =   "frmTVta.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#0.000"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   2220
      Width           =   735
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
      TabIndex        =   0
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   5820
      Picture         =   "frmTVta.frx":0996
      Style           =   1  'Graphical
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   135
      Width           =   255
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   33
      Left            =   7860
      Picture         =   "frmTVta.frx":0B40
      Style           =   1  'Graphical
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1515
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
      Index           =   33
      Left            =   1080
      TabIndex        =   8
      Top             =   1500
      Width           =   1275
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTVta.frx":0CEA
      Left            =   1080
      List            =   "frmTVta.frx":0CEC
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2220
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   6
      Top             =   1140
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   37102
   End
   Begin VB.CheckBox chkCalcularIGV 
      Caption         =   "Calcular I.G.&V."
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   60
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3180
      Width           =   1335
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
      Left            =   3900
      TabIndex        =   10
      Top             =   1860
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
      Height          =   315
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   1860
      Width           =   2235
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
      Index           =   2
      Left            =   7260
      TabIndex        =   2
      Top             =   120
      Width           =   1155
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
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   1
      Left            =   5640
      TabIndex        =   7
      Top             =   1140
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   37102
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   2835
      Left            =   0
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3540
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   5001
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   441
      ForeColor       =   8388608
      TabCaption(0)   =   "I&mportes"
      TabPicture(0)   =   "frmTVta.frx":0CEE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label24"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDatoDeta(19)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDatoDeta(20)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDatoDeta(21)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDatoDeta(22)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDatoDeta(23)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDatoDeta(24)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDatoDeta(25)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDatoDeta(26)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblDatoDeta(27)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblDatoDeta(28)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblDatoDeta(29)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblDatoDeta(30)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblDatoDeta(31)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblDatoDeta(32)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDato(12)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtDato(13)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtDato(14)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtDato(15)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtDato(16)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtDato(17)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtDato(18)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtDato(5)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdDatoAyud(19)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtDato(19)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtDato(20)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdDatoAyud(20)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdDatoAyud(21)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtDato(21)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmdDatoAyud(22)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtDato(22)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdDatoAyud(23)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtDato(23)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdDatoAyud(24)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtDato(24)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdDatoAyud(25)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtDato(25)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdDatoAyud(26)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtDato(26)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cmdDatoAyud(27)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtDato(27)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cmdDatoAyud(28)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtDato(28)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cmdDatoAyud(29)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtDato(29)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtDato(6)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtDato(7)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtDato(8)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtDato(9)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtDato(10)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtDato(11)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "cmdMas(1)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmdMas(2)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmdMas(3)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "cmdMas(4)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cmdMas(5)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cmdMas(6)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "cmdMas(7)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "chkMonedaActiva"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "chkDesactivar"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtDato(30)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "cmdDatoAyud(30)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtDato(31)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cmdDatoAyud(31)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txtDato(32)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "cmdDatoAyud(32)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).ControlCount=   74
      TabCaption(1)   =   "C&uentas"
      TabPicture(1)   =   "frmTVta.frx":0D0A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dgrDetalle"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   32
         Left            =   9060
         Picture         =   "frmTVta.frx":0D26
         Style           =   1  'Graphical
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   2430
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
         Index           =   32
         Left            =   6840
         TabIndex        =   54
         Top             =   2400
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   31
         Left            =   9060
         Picture         =   "frmTVta.frx":0ED0
         Style           =   1  'Graphical
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   2130
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
         Index           =   31
         Left            =   6840
         TabIndex        =   49
         Top             =   2100
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   30
         Left            =   9060
         Picture         =   "frmTVta.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   1830
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
         Index           =   30
         Left            =   6840
         TabIndex        =   44
         Top             =   1800
         Width           =   675
      End
      Begin VB.CheckBox chkDesactivar 
         Caption         =   "Des&activar Cuentas"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   4980
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox chkMonedaActiva 
         Caption         =   "M&oneda activa"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   330
         Width           =   1635
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   -67080
         ScaleHeight     =   270
         ScaleWidth      =   1575
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   0
         Width           =   1575
         Begin VB.CommandButton cmdGenerar 
            Caption         =   "&Generar"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   0
            Visible         =   0   'False
            Width           =   700
         End
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "&Borrar"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   780
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   0
            Visible         =   0   'False
            Width           =   700
         End
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
         Index           =   7
         Left            =   3000
         Picture         =   "frmTVta.frx":1224
         TabIndex        =   52
         Top             =   2425
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
         Height          =   285
         Index           =   6
         Left            =   3000
         Picture         =   "frmTVta.frx":1326
         TabIndex        =   47
         Top             =   2125
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
         Height          =   285
         Index           =   5
         Left            =   3000
         Picture         =   "frmTVta.frx":1428
         TabIndex        =   42
         Top             =   1825
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
         Height          =   285
         Index           =   4
         Left            =   3000
         Picture         =   "frmTVta.frx":152A
         TabIndex        =   37
         Top             =   1525
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
         Height          =   285
         Index           =   3
         Left            =   3000
         Picture         =   "frmTVta.frx":162C
         TabIndex        =   32
         Top             =   1225
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
         Height          =   285
         Index           =   2
         Left            =   3000
         Picture         =   "frmTVta.frx":172E
         TabIndex        =   27
         Top             =   925
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
         Height          =   285
         Index           =   1
         Left            =   3000
         Picture         =   "frmTVta.frx":1830
         TabIndex        =   22
         Top             =   625
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
         Index           =   11
         Left            =   1320
         TabIndex        =   50
         Top             =   2400
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
         TabIndex        =   45
         Top             =   2100
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
         Index           =   9
         Left            =   1320
         TabIndex        =   40
         Top             =   1800
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
         Index           =   8
         Left            =   1320
         TabIndex        =   35
         Top             =   1500
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
         Index           =   7
         Left            =   1320
         TabIndex        =   30
         Top             =   1200
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
         Index           =   6
         Left            =   1320
         TabIndex        =   25
         Top             =   900
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
         Index           =   29
         Left            =   6840
         TabIndex        =   39
         Top             =   1500
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   29
         Left            =   9060
         Picture         =   "frmTVta.frx":1932
         Style           =   1  'Graphical
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   1525
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
         Index           =   28
         Left            =   6840
         TabIndex        =   34
         Top             =   1200
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   28
         Left            =   9060
         Picture         =   "frmTVta.frx":1ADC
         Style           =   1  'Graphical
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   1225
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
         Index           =   27
         Left            =   6840
         TabIndex        =   29
         Top             =   900
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   27
         Left            =   9060
         Picture         =   "frmTVta.frx":1C86
         Style           =   1  'Graphical
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   925
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
         Index           =   26
         Left            =   6840
         TabIndex        =   24
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   26
         Left            =   9060
         Picture         =   "frmTVta.frx":1E30
         Style           =   1  'Graphical
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   625
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
         Index           =   25
         Left            =   3300
         TabIndex        =   53
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   25
         Left            =   6540
         Picture         =   "frmTVta.frx":1FDA
         Style           =   1  'Graphical
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   2425
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
         Index           =   24
         Left            =   3300
         TabIndex        =   48
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   24
         Left            =   6540
         Picture         =   "frmTVta.frx":2184
         Style           =   1  'Graphical
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2125
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
         Index           =   23
         Left            =   3300
         TabIndex        =   43
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   23
         Left            =   6540
         Picture         =   "frmTVta.frx":232E
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1825
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
         Index           =   22
         Left            =   3300
         TabIndex        =   38
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   22
         Left            =   6540
         Picture         =   "frmTVta.frx":24D8
         Style           =   1  'Graphical
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1525
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
         Index           =   21
         Left            =   3300
         TabIndex        =   33
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   21
         Left            =   6540
         Picture         =   "frmTVta.frx":2682
         Style           =   1  'Graphical
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1225
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   20
         Left            =   6540
         Picture         =   "frmTVta.frx":282C
         Style           =   1  'Graphical
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   925
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
         Index           =   20
         Left            =   3300
         TabIndex        =   28
         Top             =   900
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
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   19
         Left            =   6540
         Picture         =   "frmTVta.frx":29D6
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   625
         Width           =   255
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   80
         Top             =   420
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5741
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
         Index           =   5
         Left            =   1320
         TabIndex        =   20
         Top             =   600
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
         Index           =   18
         Left            =   1320
         TabIndex        =   51
         Top             =   2400
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
         Index           =   17
         Left            =   1320
         TabIndex        =   46
         Top             =   2100
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
         Index           =   16
         Left            =   1320
         TabIndex        =   41
         Top             =   1800
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
         Index           =   15
         Left            =   1320
         TabIndex        =   36
         Top             =   1500
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
         Index           =   14
         Left            =   1320
         TabIndex        =   31
         Top             =   1200
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
         Index           =   13
         Left            =   1320
         TabIndex        =   26
         Top             =   900
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
         Index           =   12
         Left            =   1320
         TabIndex        =   21
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
         Height          =   315
         Index           =   32
         Left            =   7500
         TabIndex        =   117
         Top             =   2400
         Width           =   1575
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
         Index           =   31
         Left            =   7500
         TabIndex        =   115
         Top             =   2100
         Width           =   1575
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
         Index           =   30
         Left            =   7500
         TabIndex        =   113
         Top             =   1800
         Width           =   1575
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
         Index           =   29
         Left            =   7500
         TabIndex        =   106
         Top             =   1500
         Width           =   1575
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
         Index           =   28
         Left            =   7500
         TabIndex        =   104
         Top             =   1200
         Width           =   1575
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
         Index           =   27
         Left            =   7500
         TabIndex        =   102
         Top             =   900
         Width           =   1575
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
         Index           =   26
         Left            =   7500
         TabIndex        =   100
         Top             =   600
         Width           =   1575
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
         Index           =   25
         Left            =   4260
         TabIndex        =   98
         Top             =   2400
         Width           =   2295
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
         Index           =   24
         Left            =   4260
         TabIndex        =   96
         Top             =   2100
         Width           =   2295
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
         Index           =   23
         Left            =   4260
         TabIndex        =   94
         Top             =   1800
         Width           =   2295
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
         Index           =   22
         Left            =   4260
         TabIndex        =   92
         Top             =   1500
         Width           =   2295
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
         Index           =   21
         Left            =   4260
         TabIndex        =   90
         Top             =   1200
         Width           =   2295
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
         Index           =   20
         Left            =   4260
         TabIndex        =   88
         Top             =   900
         Width           =   2295
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
         Index           =   19
         Left            =   4260
         TabIndex        =   86
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   195
         TabIndex        =   79
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   195
         TabIndex        =   78
         Top             =   1860
         Width           =   420
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   195
         TabIndex        =   77
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   180
         TabIndex        =   76
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   180
         TabIndex        =   75
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label21 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   180
         TabIndex        =   74
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label23 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   3420
         TabIndex        =   73
         Top             =   345
         Width           =   3285
      End
      Begin VB.Label Label24 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   6960
         TabIndex        =   72
         Top             =   345
         Width           =   2265
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   180
         TabIndex        =   71
         Top             =   2460
         Width           =   555
      End
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   3
      Left            =   1080
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   37102
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rango Final:"
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
      Left            =   5890
      TabIndex        =   118
      Top             =   660
      Width           =   885
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8400
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   60
      TabIndex        =   107
      Top             =   1200
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8400
      Y1              =   540
      Y2              =   540
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
      Left            =   1200
      TabIndex        =   84
      Top             =   120
      Width           =   4635
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
      Index           =   33
      Left            =   2340
      TabIndex        =   82
      Top             =   1500
      Width           =   5535
   End
   Begin VB.Label Label20 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   3420
      TabIndex        =   70
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label Label19 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   60
      TabIndex        =   69
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   60
      TabIndex        =   68
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label16 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   4560
      TabIndex        =   67
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label15 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   1920
      TabIndex        =   66
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label14 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   2460
      TabIndex        =   65
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "NºDoc.:"
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
      Left            =   6240
      TabIndex        =   64
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Doc.:"
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
      TabIndex        =   63
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   60
      TabIndex        =   62
      Top             =   1560
      Width           =   525
   End
End
Attribute VB_Name = "frmTVta"
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
Private Const MINIMOINDICEIMPORTEMN As Byte = 5, _
              MINIMOINDICEIMPORTEME As Byte = 12, _
              MINIMOINDICEMAS As Byte = 1, _
              MINIMOINDICECUENTA As Byte = 19, _
              MINIMOINDICECCOSTO As Byte = 26, _
              CANTIDADIMPORTES As Byte = 7
'[Repetir en frmTVtaMasGrd.
Private Const DIFERENCIAMASIMPORTE As Byte = 4, _
              DIFERENCIAMASCUENTA As Byte = 18, _
              DIFERENCIAMASCCOSTO As Byte = 25
Private Const CUENTASCONCCOSTO As Byte = 7
']

'[Repetir en frmTVtaGrd y fmrTVtaMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2

']
']
Private Sub Form_Load()
   pbValidada = False
   pbFecha = True
   Me.KeyPreview = True
   
   With frmTVtaGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
      txtLlave(0).MaxLength = .uorstMain!CodTDc.DefinedSize
      txtLlave(1).MaxLength = .uorstMain!SerDoc.DefinedSize
      txtLlave(2).MaxLength = .uorstMain!NroDoc.DefinedSize
    ']
   
    '[Datos                            'Cambiar.
      With cboTpoMon
         .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
         .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
      End With
'      mskDato(0).MaxLength = .uorstMain!Tf1Cta.DefinedSize + 1
      txtDato(0).MaxLength = .uorstMain!CodDro.DefinedSize
      txtDato(1).MaxLength = .uorstMain!NroCpb.DefinedSize
      txtDato(2).MaxLength = .uorstMain!RefDoc.DefinedSize
      txtDato(3).MaxLength = .uorstMain!GloDoc.DefinedSize
      txtDato(4).MaxLength = .uorstMain!ImpTCb.DefinedSize
      txtDato(5).MaxLength = 14
      txtDato(6).MaxLength = 14
      txtDato(7).MaxLength = 14
      txtDato(8).MaxLength = 14
      txtDato(9).MaxLength = 14
      txtDato(10).MaxLength = 14
      txtDato(11).MaxLength = 14
      txtDato(12).MaxLength = 14
      txtDato(13).MaxLength = 14
      txtDato(14).MaxLength = 14
      txtDato(15).MaxLength = 14
      txtDato(16).MaxLength = 14
      txtDato(17).MaxLength = 14
      txtDato(18).MaxLength = 14
      txtDato(19).MaxLength = 8
      txtDato(20).MaxLength = 8
      txtDato(21).MaxLength = 8
      txtDato(22).MaxLength = 8
      txtDato(23).MaxLength = 8
      txtDato(24).MaxLength = 8
      txtDato(25).MaxLength = 8
      txtDato(26).MaxLength = 5
      txtDato(27).MaxLength = 5
      txtDato(28).MaxLength = 5
      txtDato(29).MaxLength = 5
      txtDato(30).MaxLength = 5
      txtDato(31).MaxLength = 5
      txtDato(32).MaxLength = 5
      txtDato(33).MaxLength = .uorstMain!CodAux.DefinedSize
      txtDato(34).MaxLength = .uorstMain!SerDoc_Fin.DefinedSize
      txtDato(35).MaxLength = .uorstMain!NroDoc_Fin.DefinedSize
   End With
   
   If pbNuevo Then
      cmdRetroceder.Enabled = False
      cmdAvanzar.Enabled = False
   End If
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   upHabilitacion False

'[Propio del formulario.
   dgrDetalle.MarqueeStyle = dbgHighlightRow
   Set dgrDetalle.DataSource = frmTVtaGrd.uorstCOCpbDet
   
   cboTpoMon.ListIndex = TPOMON_NAC_IND
   chkMonedaActiva.Value = vbChecked
   sstMain.Tab = 0
']
End Sub

Private Sub Form_Activate()
''   ppDatosGrid
 
   If Not pbNuevo And cmdCorregir.Enabled Then
      cmdCorregir.SetFocus
   End If

 '[Propio del formulario.
   If Not pbNuevo Then
      dtpDato(3).Tag = dtpDato(3).Value
   End If
   txtDato(4).Tag = txtDato(4).Text
 ']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not frmTVtaGrd.uorstMain.EOF Then
      If frmTVtaGrd.uorstMain.EditMode <> adEditNone Then frmTVtaGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
   End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmTVtaGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTVtaGrd.uorstMain_Grd.MoveFirst
   frmTVtaGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmTVtaGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTVtaGrd.uorstMain_Grd.MoveFirst
   frmTVtaGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
End Sub

Public Sub cmdCorregir_Click()
   'Verificación de Mes Cerrado.
   If gbCieVta Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   pbCorregir = True
   frmTVtaGrd.uocnnMain.BeginTrans     'Cambiar Formulario de Grid. 'INICIA TRANSACCION.

   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
   txtDato(0).Enabled = (chkIndPreGen.Value = 0)
   lblDatoDeta(0).Enabled = (chkIndPreGen.Value = 0)
   cmdDatoAyud(0).Enabled = (chkIndPreGen.Value = 0)

 '[Dato con el foco al corregir.       'Cambiar.
   dtpDato(3).SetFocus
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err

 '[Propio del formulario.
   Dim dnSumaMN As Double, _
       dnSumaME As Double
   
   If txtDato(33).Text = "" Then
      MsgBox TEXT_6002, vbCritical
      txtDato(33).SetFocus
      Exit Sub
   End If
   
   If txtDato(0).Text = "" Then
      MsgBox TEXT_6002, vbCritical
      txtDato(0).SetFocus
      Exit Sub
   End If
   
   With frmTVtaGrd.uorstMain
      dnSumaMN = CDec(txtDato(5).Text) + CDec(txtDato(6).Text) + CDec(txtDato(7).Text) + CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text)
      dnSumaME = CDec(txtDato(12).Text) + CDec(txtDato(13).Text) + CDec(txtDato(14).Text) + CDec(txtDato(15).Text) + CDec(txtDato(16).Text) + CDec(txtDato(17).Text)
'      If gfRedond(CDec(txtDato(5).Text) + CDec(txtDato(6).Text) + CDec(txtDato(7).Text) + CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text), 2) <> CDec(txtDato(11).Text) Then
      If dnSumaMN <> CDec(txtDato(11).Text) Then
'         If MsgBox(TEXT_9011, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
         If MsgBox(TEXT_9011 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(11).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
            Exit Sub
         End If
'      ElseIf gfRedond(CDec(TxtDato(12).Text) + CDec(TxtDato(13).Text) + CDec(TxtDato(14).Text) + CDec(TxtDato(15).Text) + CDec(TxtDato(16).Text) + CDec(TxtDato(17).Text), 2) <> CDec(TxtDato(18).Text) Then
      ElseIf dnSumaME <> CDec(txtDato(18).Text) Then
'         If MsgBox(TEXT_9012, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
         If MsgBox(TEXT_9012 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(dnSumaME - CDec(txtDato(18).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
            Exit Sub
         End If
      End If
   End With

   ' Valido las Cuentas esten Correctas(llenas para todas los valores)
   If chkIndPreGen.Value = vbChecked Then
      chkIndPreGen.Value = IIf(ValidaCtasCCo, 1, 0)
      If Not (chkIndPreGen.Value = vbChecked) Then
         If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
            Exit Sub
         End If
      End If
   End If
 ']

   With frmTVtaGrd                     'Cambiar Formulario de Grid.
      If pbNuevo And frmTVtaGrd.ubGrabaMas = 0 Then
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
      ppGenera
    ']
'      .uorstCCCfg.Update
      .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      
    '[Actualiza grid..
      .uorstMain_Grd.Requery
      .upDatosGrid
      .uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
    ']

      pbCorregir = False
   
      If pbNuevo Then
'         .uorstMain.Requery
'         .upDatosGrid
'       '[Búsqueda de llave actual.     'Cambiar.
'         .uorstMain.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
'       ']
         pbValidada = False
         cmdGrabar.Enabled = False
         upHabilitacion False
         frmTVtaGrd.ubGrabaMas = INDMASCTA_INI
   
         upDatosPredeterminados
         pbFecha = True
       '[Llave habilitar  'Cambiar.
         txtLlave(0).Enabled = True
         txtLlave(1).Enabled = True
         txtLlave(2).Enabled = True
         lblLlaveDeta(0).Enabled = True
         cmdLlaveAyud(0).Enabled = True
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
         upHabilitacion False
      End If
   End With
      
   Exit Sub
Err:
   gpErrores
  
   frmTVtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
 '[Propio del formulario.
   frmTVtaGrd.uorstCOCpbCab.CancelUpdate
   frmTVtaGrd.uorstCOCpbDet.CancelBatch
   frmTVtaGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
   pbCorregir = False
 ']
   
   gpTUe_Deshacer Me
End Sub

Public Sub cmdSalir_Click()
   If pbNuevo Or pbCorregir Then
      pbCorregir = False
      frmTVtaGrd.uocnnMain.RollbackTrans 'RESTAURA TRANSACCION.
   End If
   Unload Me
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0
      txtLlave(Index).SetFocus
'   Case 1
'      mskLlave(Index).SetFocus
   End Select
   ppAyuBus AYULLA, Index
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1, 33
      txtDato(Index).SetFocus
'   Case 0, 1
'      mskDato(Index).SetFocus
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

Private Sub txtLlave_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYULLA, Index
   End If
End Sub

Private Sub txtLlave_LostFocus(Index As Integer)
'   If pbValidada Then dtpDato(3).SetFocus 'Cambiar.
   If pbValidada Then                  'Cambiar.
      txtLlave(0).Enabled = False
      txtLlave(1).Enabled = False
      txtLlave(2).Enabled = False
      lblLlaveDeta(0).Enabled = False
      cmdLlaveAyud(0).Enabled = False
      
      If txtDato(34).Enabled Then
         txtDato(34).SetFocus
      ElseIf dtpDato(3).Enabled Then
         dtpDato(3).SetFocus
      End If
   End If
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
 '[Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index
   Case 0, 1, 2                        'Cambiar (añadir índices).
      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
      End If
   End Select
 ']
   
 '[Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index
   Case 0                              'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYULLA, Index)
      If Cancel Then Exit Sub
   End Select
 ']
 
 '[Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 And Len(Trim(txtLlave(1).Text)) <> 0 And Len(Trim(txtLlave(2).Text)) <> 0 Then
      With frmTVtaGrd                  'Cambiar Formulario de Grid.
         Set .uorstTemporal = .uocnnMain.Execute("SELECT MesPvs FROM COVtaDoc WHERE CodTDc ='" & txtLlave(0).Text & "' AND SerDoc='" & txtLlave(1).Text & "' AND NroDoc='" & txtLlave(2).Text & "'")
         If .uorstTemporal.RecordCount > 0 Then
            MsgBox TEXT_8007 & Chr(13) & "(mes " & gfMesLet("01" & .uorstTemporal!MesPvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
            Cancel = True
            Exit Sub
         End If
         .uorstTemporal.Close
      End With
    '[Propio del formulario.
      If frmTVtaGrd.ubGrabaMas = 0 Then
         frmTVtaGrd.ubGrabaMas = 1
         With frmTVtaGrd
            If pbNuevo Then
               .uorstMain.AddNew
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

'Private Sub mskDato_GotFocus(Index As Integer)
'   mskDato(Index).SelStart = 0
'   mskDato(Index).SelLength = mskDato(Index).MaxLength
'End Sub
'
'Private Sub mskDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF2 Then
'      ppAyuBus Index
'   End If
'End Sub

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
   If KeyCode = vbKeyF2 Then
      ppAyuBus AYUDAT, Index
   End If
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
''   Dim doColumna As Field
''
   Select Case Index
''   Case 4
''      If txtDato(Index).Tag <> frmTVtaGrd.uorstCOCta!TpoTCb Or CDec(txtDato(Index).Text) = 0 Then
''         txtDato(Index).Tag = frmTVtaGrd.uorstCOCta!TpoTCb
''         With frmTVtaGrd.uorstTGTCb
''            If .RecordCount <> 0 Then .MoveFirst
''            .Find "FehTCb = '" & dtpDato(3).Value & "'"
'''            uorstMain!ImpTCb = IIf(frmTVtaGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
''         End With
''      End If
   Case MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
'///Angel 12/12/2003
'/// Agregado para validar el T/C
      If CDec(txtDato(4).Text) <= 0 Then
         MsgBox "No se ha ingresado Tipo de Cambio para Esta Fecha", vbCritical
         txtDato(4).SetFocus
         Exit Sub
      End If
'///

      If CDec(txtDato(Index).Text) = 0 Then
         txtDato(Index).Text = Format(0, FORMATO_NUM_1)
'///Angel 12/12/2003
'///Validacion de T/C para importes en cero cuando hay valor en la otra moneda
         If Index >= MINIMOINDICEIMPORTEMN And Index < MINIMOINDICEIMPORTEME And cboTpoMon.ListIndex = TPOMON_EXT_IND Then
            txtDato(Index).Text = Format(CDec(txtDato(Index + CANTIDADIMPORTES).Text) * CDec(txtDato(4).Text), FORMATO_NUM_1)
         ElseIf Index >= MINIMOINDICEIMPORTEME And Index < (MINIMOINDICEIMPORTEME + CANTIDADIMPORTES) And cboTpoMon.ListIndex = TPOMON_NAC_IND Then
            txtDato(Index).Text = Format(CDec(txtDato(Index - CANTIDADIMPORTES).Text) / CDec(txtDato(4).Text), FORMATO_NUM_1)
         End If
'///
      End If
      If Index = 5 Or Index = 12 Then
         If chkCalcularIGV Then txtDato(Index + 3).Text = Format(CDec(txtDato(Index).Text) * CDec(gnPctIGV) / 100, FORMATO_NUM_1)
         If chkCalcularISC Then txtDato(Index + 4).Text = Format(CDec(txtDato(Index).Text) * CDec(gnPctISC) / 100, FORMATO_NUM_1)
         If chkMonedaActiva.Value = vbChecked Then
            If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
               If CDec(txtDato(Index + 3).Text) > 0 Then txtDato((Index + CANTIDADIMPORTES) + 3).Text = Format(gfRedond(CDec(txtDato(Index + 3).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
               If CDec(txtDato(Index + 4).Text) > 0 Then txtDato((Index + CANTIDADIMPORTES) + 4).Text = Format(gfRedond(CDec(txtDato(Index + 4).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
            Else
               If CDec(txtDato(Index + 3).Text) > 0 Then txtDato((Index - CANTIDADIMPORTES) + 3).Text = Format(gfRedond(CDec(txtDato(Index + 3).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
               If CDec(txtDato(Index + 4).Text) > 0 Then txtDato((Index - CANTIDADIMPORTES) + 4).Text = Format(gfRedond(CDec(txtDato(Index + 4).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
            End If
         End If
      End If
      
     'Cálculo del total.
      If (Index = 11 And txtDato(Index).Text = 0) Or (Index = 18 And txtDato(Index).Text = 0) Then
         If Index = 11 Then
            txtDato(11).Text = Format(CDec(txtDato(5).Text) + CDec(txtDato(6).Text) + CDec(txtDato(7).Text) + CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text), FORMATO_NUM_1)
         Else
            txtDato(18).Text = Format(CDec(txtDato(12).Text) + CDec(txtDato(13).Text) + CDec(txtDato(14).Text) + CDec(txtDato(15).Text) + CDec(txtDato(16).Text) + CDec(txtDato(17).Text), FORMATO_NUM_1)
         End If
      End If
   
      ' Miguel Angel 25/01/2004 Convierte el monto si es la moneda funcional
      ' Quito esto al momento de comvertir If CDec(txtDato(Index + CANTIDADIMPORTES).Text) = 0 Then
      If chkMonedaActiva.Value = vbChecked Then
         If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
            txtDato(Index + CANTIDADIMPORTES).Text = Format(gfRedond(CDec(txtDato(Index).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
         Else
            txtDato(Index - CANTIDADIMPORTES).Text = Format(gfRedond(CDec(txtDato(Index).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
         End If
      End If
   Case MINIMOINDICECUENTA To MINIMOINDICECCOSTO - 1 'Cambiar (añadir índices).
      If txtDato(Index).Text = "" Then
         txtDato(Index + CUENTASCONCCOSTO).Text = ""
         lblDatoDeta(Index).Caption = ""
         lblDatoDeta(Index + CUENTASCONCCOSTO).Caption = ""
         txtDato(Index + CUENTASCONCCOSTO).Enabled = False
         lblDatoDeta(Index + CUENTASCONCCOSTO).Enabled = False
         cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = False
         cmdMas(Index - MINIMOINDICECUENTA + 1).Enabled = True
         If (Not pbNuevo) And cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_CTA Then
            If frmTVtaGrd.uorstCOVtaDocCta.RecordCount > 0 Then
               ppAbreCtaCCo
               If frmTVtaGrd.uorstCOVtaDocCta.State = adStateOpen Then
                  frmTVtaGrd.uorstCOVtaDocCta.MoveFirst
                  Do
                     If frmTVtaGrd.uorstCOVtaDocCta!CodTDc = txtLlave(0).Text And _
                       frmTVtaGrd.uorstCOVtaDocCta!SerDoc = txtLlave(1).Text And _
                       frmTVtaGrd.uorstCOVtaDocCta!NroDoc = txtLlave(2).Text And _
                       frmTVtaGrd.uorstCOVtaDocCta!TpoCnc = Trim(Str(Index - MINIMOINDICECUENTA + 1)) Then
'                        frmTVtaGrd.uorstCOVtaDocCCo.MoveFirst
'                        Do
'                           If frmTVtaGrd.uorstCOVtaDocCCo!CodAux = txtLlave(0).Text And _
                             frmTVtaGrd.uorstCOVtaDocCCo!CodTDc = txtLlave(1).Text And _
                             frmTVtaGrd.uorstCOVtaDocCCo!SerDoc = txtLlave(2).Text And _
                             frmTVtaGrd.uorstCOVtaDocCCo!NroDoc = txtLlave(3).Text And _
                             frmTVtaGrd.uorstCOVtaDocCCo!TpoCnc = Trim(Str(Index - MINIMOINDICECUENTA + 1)) And _
                             frmTVtaGrd.uorstCOVtaDocCCo!CodCta = frmTVtaGrd.uorstCOVtaDocCta!CodCta Then
'                              frmTVtaGrd.uorstCOVtaDocCCo.Delete
'                           End If
'                           frmTVtaGrd.uorstCOVtaDocCCo.MoveNext
'                        Loop Until frmTVtaGrd.uorstCOVtaDocCCo.EOF
'                        frmTVtaGrd.uorstCOVtaDocCCo.Requery
                        frmTVtaGrd.uorstCOVtaDocCta.Delete
                     End If
                     frmTVtaGrd.uorstCOVtaDocCta.MoveNext
                  Loop Until frmTVtaGrd.uorstCOVtaDocCta.EOF
                  frmTVtaGrd.uorstCOVtaDocCta.Requery
               End If
            End If
         End If
         cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI
      ElseIf cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI Then
            cmdMas(Index - MINIMOINDICECUENTA + 1).Enabled = False
'            txtDato(Index + CUENTASCONCCOSTO).Enabled = True
'            cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = True
      End If
   End Select
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

  'Completa con ceros a la izquierda.
   Select Case Index
   Case MINIMOINDICECCOSTO To CANTIDADIMPORTES - 1, 34, 35 'Cambiar (añadir índices).
      If Len(Trim(txtDato(Index).Text)) <> 0 And Len(Trim(txtDato(Index).Text)) <> txtDato(Index).MaxLength Then
         txtDato(Index) = gfCeros(txtDato(Index).Text, txtDato(Index).MaxLength, 0, "0")
      End If
   End Select

  'Asigna 0 a campos numéricos si están vacíos.
   Select Case Index
   Case 4, MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      If Not IsNumeric(txtDato(Index).Text) Then
         txtDato(Index).Text = 0
      End If
   End Select

  'Da formato.
   Select Case Index
   Case 4
      txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_2)
   Case MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      txtDato(Index).Text = Format(txtDato(Index).Text, FORMATO_NUM_1)
   End Select

  'Busca el dato en su tabla principal.
   Select Case Index
   Case 0, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYUDAT, Index)
      If Cancel Then Exit Sub
'      If lblDatoDeta(Index).Caption <> "" Then
      If Index >= MINIMOINDICECUENTA And Index < MINIMOINDICECCOSTO Then
'         If frmTVtaGrd.uorstCOCta.RecordCount > 0 And txtDato(Index + CUENTASCONCCOSTO).Text <> "" Then
         If frmTVtaGrd.uorstCOCta.RecordCount > 0 Then
            If Not frmTVtaGrd.uorstCOCta.EOF Then
               If frmTVtaGrd.uorstCOCta!IndCCo = INDCCO_ACT Then
                  txtDato(Index + CUENTASCONCCOSTO).Enabled = True
                  cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = True
               Else
                  txtDato(Index + CUENTASCONCCOSTO).Text = ""
                  lblDatoDeta(Index + CUENTASCONCCOSTO).Caption = ""
                  txtDato(Index + CUENTASCONCCOSTO).Enabled = False
                  cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = False
               End If
            End If
        End If
      End If
'[Propio del formulario (¿Angel?).
'      If frmTVtaGrd.ubGrabaMas = 0 Then
'         frmTVtaGrd.ubGrabaMas = 1
'         With frmTVtaGrd
'            If pbNuevo Then
'               .uorstMain.AddNew
'            End If
'            upDatosDesconectados 0
'            .uorstMain.Update
'         End With
'      End If
']
'      End If
   Case 33                             'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYUDAT, Index)
      If Cancel Then Exit Sub
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.TDc_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
   Else
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod "Length(CodDro)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
         modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
         modAyuBus.CCo_Cod "Length(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case 33                          'Cambiar (añadir índices).
         modAyuBus.Aux_Det "IndCli=1", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
         txtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
   End If
End Sub

Private Function ppAyuDet(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex                 'Cambiar.
      Case 0
         If txtLlave(tnIndex).Text = "" Then
            lblLlaveDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTVtaGrd.uorstTGTDc
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodTDc='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !DetTDc
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
         With frmTVtaGrd.uorstCODro
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodDro='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & !DetDro
            End If
         End With
      Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTVtaGrd.uorstCOCta
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodCta='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
'[ARREGLAR. Encontrar forma de mostrar el caption de los Label así no haya espacios en blanco.
'               lblDatoDeta(tnIndex).Caption = " " & !DetCta
               lblDatoDeta(tnIndex).Caption = " " & Left(!DetCta, 18)
']ARREGLAR.
            End If
         End With
      Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTVtaGrd.uorstCOCCo
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
'[ARREGLAR. Encontrar forma de mostrar el caption de los Label así no haya espacios en blanco.
'               lblDatoDeta(tnIndex).Caption = " " & !DetCCo
               lblDatoDeta(tnIndex).Caption = " " & Left(!DetCCo, 12)
']ARREGLAR.
            End If
         End With
      Case 33
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTVtaGrd.uorstTGAux
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodAux='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & !RazAux
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
   With frmTVtaGrd.uorstMain           'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !CodTDc = txtLlave(0).Text
            !SerDoc = txtLlave(1).Text
            !NroDoc = txtLlave(2).Text
            !MesPvs = gsMesAct
            !PctIGV = CDec(gnPctIGV)
            !PctISC = CDec(gnPctISC)
         End If

        'Datos.
         !TpoMon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
         !IndPreGen = IIf(chkIndPreGen.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
'         !CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
         !FehOpe = dtpDato(3).Value
         !FeEDoc = dtpDato(0).Value
         !FeVDoc = dtpDato(1).Value
'         !Tf1Cta = mskDato(0).Text
'         !CodMon = optTpoMon(1).Value
         !CodDro = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
         !NroCpb = txtDato(1).Text
         !RefDoc = txtDato(2).Text
         !GloDoc = txtDato(3).Text
         !ImpTCb = txtDato(4).Text
         !ImpOGr_MN = txtDato(5).Text
         !ImpExp_MN = txtDato(6).Text
         !ImpExo_MN = txtDato(7).Text
         !ImpIGV_MN = txtDato(8).Text
         !ImpISC_MN = txtDato(9).Text
         !ImpOIm_MN = txtDato(10).Text
         !ImpTot_MN = txtDato(11).Text
         !ImpOGr_ME = txtDato(12).Text
         !ImpExp_ME = txtDato(13).Text
         !ImpExo_ME = txtDato(14).Text
         !ImpIGV_ME = txtDato(15).Text
         !ImpISC_ME = txtDato(16).Text
         !ImpOIm_ME = txtDato(17).Text
         !ImpTot_ME = txtDato(18).Text
         !CodAux = IIf(txtDato(33).Text = "", Null, txtDato(33).Text)
         !SerDoc_Fin = txtDato(34).Text
         !NroDoc_Fin = txtDato(35).Text
'[Propio del formulario.
         '.Update
']

       '[Actualización de datos por Cuentas y Centros de Costo registrados directamente.
         ppAbreCtaCCo
         For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
            If cmdMas(dnContador).Tag = INDMASCTA_CTA And Val(txtDato(dnContador + DIFERENCIAMASCUENTA).Text) <> 1 Then
               With frmTVtaGrd.uorstCOVtaDocCta
                  .MoveFirst
                  .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & "'"
                  If Not .EOF Then
                     .Delete
                     .Update
                     .Requery
                     frmTVtaGrd.uorstCOVtaDocCCo.Requery
                     Call upActualizaMas(dnContador, INDMASCTA_INI)
                  End If
               End With
            End If
            
            If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
               cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Then
               With frmTVtaGrd.uorstCOVtaDocCta
                  If .RecordCount <> 0 Then .MoveFirst
                  .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & "'"
                  If .EOF Then
                     .AddNew
                     !CodTDc = txtLlave(0).Text
                     !SerDoc = txtLlave(1).Text
                     !NroDoc = txtLlave(2).Text
                     !TpoCnc = dnContador
                     !UsrCre = gsAbvUsr
                     !FyHCre = Now
                  Else
                     !UsrMdf = gsAbvUsr
                     !FyHMdf = Now
                  End If
                  !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
                  !ImpCta_MN = txtDato(dnContador + DIFERENCIAMASIMPORTE).Text
                  !ImpCta_ME = txtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text
                  .Update
               End With
               If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Or _
                  cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Then
                  With frmTVtaGrd.uorstCOVtaDocCCo
                     If .RecordCount <> 0 Then .MoveFirst
                     .Find "cLlave1='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & txtDato(dnContador + DIFERENCIAMASCUENTA).Text & "'"
                     If .EOF Then
                        .AddNew
                        !CodTDc = txtLlave(0).Text
                        !SerDoc = txtLlave(1).Text
                        !NroDoc = txtLlave(2).Text
                        !TpoCnc = dnContador
                        !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
                        !UsrCre = gsAbvUsr
                        !FyHCre = Now
                     Else
                        !UsrMdf = gsAbvUsr
                        !FyHMdf = Now
                     End If
                     !CodCCo = txtDato(dnContador + DIFERENCIAMASCCOSTO).Text
                     !ImpCCo_MN = txtDato(dnContador + DIFERENCIAMASIMPORTE).Text
                     !ImpCCo_ME = txtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text
                     .Update
                  End With
               End If
               Call upActualizaMas(dnContador, INDMASCTA_CTA)
            End If
         Next
       ']
      Else
        'Llaves.
         txtLlave(0).Text = !CodTDc
         txtLlave(1).Text = !SerDoc
         txtLlave(2).Text = !NroDoc

        'Datos.
         cboTpoMon.ListIndex = IIf(!TpoMon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
         chkIndPreGen.Value = IIf(!IndPreGen = INDPREGEN_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(!CodSoc), "", !CodSoc)
         dtpDato(3).Value = !FehOpe
         dtpDato(0).Value = !FeEDoc
         dtpDato(1).Value = !FeVDoc
'         optTpoMon(1).Value = uorstMain!CodMon
'         mskDato(0).Text = IIf(IsNull(.uorstMain!Tf1Cta), "", .uorstMain!Tf1Cta)
         txtDato(0).Text = IIf(IsNull(!CodDro), "", !CodDro)
         txtDato(1).Text = IIf(IsNull(!NroCpb), "", !NroCpb)
         txtDato(2).Text = IIf(IsNull(!RefDoc), "", !RefDoc)
         txtDato(3).Text = IIf(IsNull(!GloDoc), "", !GloDoc)
         txtDato(4).Text = Format(!ImpTCb, FORMATO_NUM_2)
         txtDato(5).Text = Format(!ImpOGr_MN, FORMATO_NUM_1)
         txtDato(6).Text = Format(!ImpExp_MN, FORMATO_NUM_1)
         txtDato(7).Text = Format(!ImpExo_MN, FORMATO_NUM_1)
         txtDato(8).Text = Format(!ImpIGV_MN, FORMATO_NUM_1)
         txtDato(9).Text = Format(!ImpISC_MN, FORMATO_NUM_1)
         txtDato(10).Text = Format(!ImpOIm_MN, FORMATO_NUM_1)
         txtDato(11).Text = Format(!ImpTot_MN, FORMATO_NUM_1)
         txtDato(12).Text = Format(!ImpOGr_ME, FORMATO_NUM_1)
         txtDato(13).Text = Format(!ImpExp_ME, FORMATO_NUM_1)
         txtDato(14).Text = Format(!ImpExo_ME, FORMATO_NUM_1)
         txtDato(15).Text = Format(!ImpIGV_ME, FORMATO_NUM_1)
         txtDato(16).Text = Format(!ImpISC_ME, FORMATO_NUM_1)
         txtDato(17).Text = Format(!ImpOIm_ME, FORMATO_NUM_1)
         txtDato(18).Text = Format(!ImpTot_ME, FORMATO_NUM_1)
         For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
            txtDato(dnContador).Text = ""
         Next
         txtDato(33).Text = IIf(IsNull(!CodAux), "", !CodAux)
         txtDato(34).Text = IIf(IsNull(!SerDoc_Fin), "", !SerDoc_Fin)
         txtDato(35).Text = IIf(IsNull(!NroDoc_Fin), "", !NroDoc_Fin)
      
       '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
         ppAyuDet AYULLA, 0
      '   ppAyuDet AYULLA, 1
         ppAyuDet AYUDAT, 0
         ppAyuDet AYUDAT, 19
         ppAyuDet AYUDAT, 20
         ppAyuDet AYUDAT, 21
         ppAyuDet AYUDAT, 22
         ppAyuDet AYUDAT, 23
         ppAyuDet AYUDAT, 24
         ppAyuDet AYUDAT, 25
         ppAyuDet AYUDAT, 26
         ppAyuDet AYUDAT, 27
         ppAyuDet AYUDAT, 28
         ppAyuDet AYUDAT, 29
         ppAyuDet AYUDAT, 30
         ppAyuDet AYUDAT, 31
         ppAyuDet AYUDAT, 32
         ppAyuDet AYUDAT, 33
       ']
      
       '[Propio del formulario.
         For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
            cmdMas(dnContador).Tag = Choose(dnContador, !IndCta_OGr, !IndCta_Exp, !IndCta_Exo, !IndCta_IGV, !IndCta_ISC, !IndCta_OIm, !IndCta_Tot)
         Next

         ppAbreCtaCCo
         With frmTVtaGrd.uorstCOVtaDocCta
            If .RecordCount > 0 Then
               For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
                  If Val(txtDato(dnContador + DIFERENCIAMASIMPORTE).Text) <> 0 Then
                     .MoveFirst
                     .Find "TpoCnc = " & dnContador
                     If Not .EOF Then
                        txtDato(dnContador + DIFERENCIAMASCUENTA).Text = !CodCta
                        ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCUENTA
                        With frmTVtaGrd.uorstCOVtaDocCCo
                           If .RecordCount > 0 Then
                              .MoveFirst
                              .Find "cLlave = " & dnContador & frmTVtaGrd.uorstCOVtaDocCta!CodCta
                              If Not .EOF Then
                                 txtDato(dnContador + DIFERENCIAMASCCOSTO).Text = !CodCCo
                                 ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCCOSTO
                              End If
                           End If
                        End With
                     End If
                  End If
               Next
            End If
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
   txtLlave(2).Text = ""

  'Datos.
   cboTpoMon.ListIndex = TPOMON_NAC_IND
   chkIndPreGen.Value = vbUnchecked
   dtpDato(3).Value = Date
   dtpDato(0).Value = Date
   dtpDato(1).Value = Date
'   optTpoMon(1).Value = True
   For dnContador = 0 To 3
      txtDato(dnContador).Text = ""
   Next
   txtDato(4).Text = Format(0, FORMATO_NUM_2)
   For dnContador = MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
      txtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
   Next
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      txtDato(dnContador).Text = ""
   Next
   txtDato(33).Text = ""
   txtDato(34).Text = ""
   txtDato(35).Text = ""

 '[Propio del formulario.
   For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      cmdMas(dnContador).Tag = INDMASCTA_INI
   Next
 ']

  'Ayudas.
   lblLlaveDeta(0).Caption = ""
   lblDatoDeta(0).Caption = ""
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      lblDatoDeta(dnContador).Caption = ""
   Next
   lblDatoDeta(33).Caption = ""
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte

  'Datos.
   cboTpoMon.Enabled = tbHabilitar
   chkCalcularIGV.Enabled = tbHabilitar
   chkCalcularISC.Enabled = tbHabilitar
   chkDesactivar.Enabled = tbHabilitar
   chkIndPreGen.Enabled = tbHabilitar
   chkMonedaActiva.Enabled = tbHabilitar
   dtpDato(0).Enabled = tbHabilitar
   dtpDato(1).Enabled = tbHabilitar
   dtpDato(3).Enabled = tbHabilitar
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   txtDato(34).Enabled = (txtLlave(0).Text = CODTDC_BOL) And tbHabilitar
   txtDato(35).Enabled = (txtLlave(0).Text = CODTDC_BOL) And tbHabilitar
   
  'Ayudas.
   cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar
   cmdDatoAyud(33).Enabled = tbHabilitar
   lblDatoDeta(33).Enabled = tbHabilitar

'[Propio del formulario
   txtDato(1).Enabled = False 'Deshabilitación del Comprobante.

   If tbHabilitar Then
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
         cmdMas(dnContador).Enabled = Not (cmdMas(dnContador).Tag = INDMASCTA_CTA)
         Call upHabilitaCuenta((Not cmdMas(dnContador).Tag = INDMASCTA_MAS), dnContador)
      Next
   Else
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
         cmdMas(dnContador).Enabled = False
         Call upHabilitaCuenta(False, dnContador)
         Call upHabilitaCCosto(False, dnContador)
      Next
   End If
']
End Sub

'[Propio del formulario

'Private Sub cmdGenerar_Click()
'   If txtDato(0).Text = "" Or txtDato(1).Text = "" Then
'      MsgBox "Falta el Diario y/o el Nº de Comprobante", vbExclamation
'      Exit Sub
'   End If
'
'   If uorstMain!IndGen Then
'      cmdBorrar_Click
'   Else
'      ppDatosWhere
'   End If
'   ppGenera
'
' '[Necesario para que actualize las columnas temporales.
'   frmTVtaGrd.uorstCOCpbDet.Requery
'   DatosGrid
' ']
'End Sub

'Private Sub cmdBorrar_Click()
'   If Not frmTVtaGrd.uorstCOCpbCab.EOF Then
'      frmTVtaGrd.uorstCOCpbCab.Delete
'      uorstMain!IndGen = False
'      frmTVtaGrd.uorstCOCpbCab.Update
'      uorstMain.Update
'
'      txtDato(0).Enabled = True
'      txtDato(1).Enabled = True
'      cmdDatoAyud.Item(0).Enabled = True
'      lblDatoDeta.Item(0).Enabled = True
'
'      ppDatosWhere
'   End If
'End Sub

Private Sub cboTpoMon_Click()
   unVerMonNac = IIf(chkMonedaActiva.Value, cboTpoMon.ListIndex, IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_IND, TPOMON_NAC_IND))
   ppCambioTpoMon
End Sub

Private Sub chkDesactivar_Click()
   Dim dnContador As Integer
   For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      If cmdMas(dnContador).Tag = INDMASCTA_CTA Then
         cmdMas(dnContador).Enabled = cmdMas(dnContador).Enabled = True
      Else
         cmdMas(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
      End If
   Next
'   For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
'      cmdMas(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
'   Next
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      If (dnContador >= MINIMOINDICECUENTA And dnContador < MINIMOINDICECCOSTO) Then
         If cmdMas(dnContador - MINIMOINDICECUENTA + 1).Enabled Then
            txtDato(dnContador).Enabled = False
            lblDatoDeta(dnContador).Enabled = False
            cmdDatoAyud(dnContador).Enabled = False
         Else
            txtDato(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            lblDatoDeta(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            cmdDatoAyud(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
         End If
      ElseIf dnContador > MINIMOINDICECCOSTO Then
         If cmdMas(dnContador - MINIMOINDICECCOSTO + 1).Enabled Then
            txtDato(dnContador).Enabled = False
            lblDatoDeta(dnContador).Enabled = False
            cmdDatoAyud(dnContador).Enabled = False
         Else
            txtDato(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            lblDatoDeta(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            cmdDatoAyud(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
         End If
      End If
   Next
'   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
'      txtDato(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
'      lblDatoDeta(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
'      cmdDatoAyud(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
'   Next
End Sub

Private Sub chkMonedaActiva_Click()
   unVerMonNac = IIf(chkMonedaActiva.Value, cboTpoMon.ListIndex, IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_IND, TPOMON_NAC_IND))
   ppCambioTpoMon
End Sub

Private Sub cmdAuxiliar_Click()
    frmMAuxGrd.Show vbModal
    frmTVtaGrd.uorstTGAux.Requery
End Sub

Private Sub cmdMas_Click(Index As Integer) 'Cambiar Formulario de Grid.
'   If frmTVtaGrd.ubGrabaMas = 0 Then
'      frmTVtaGrd.ubGrabaMas = 1
'      With frmTVtaGrd
'         If pbNuevo Then
'            .uorstMain.AddNew
'         End If
'         upDatosDesconectados 0
'         .uorstMain.Update
'      End With
'   End If

   frmTVtaMasGrd.unIndice = Index
   frmTVtaMasGrd.Show vbModal
   ppAbreCtaCCo
   ppAyuDet AYUDAT, Index + DIFERENCIAMASCUENTA
   If Index <= CUENTASCONCCOSTO Then
      ppAyuDet AYUDAT, Index + DIFERENCIAMASCCOSTO
   End If
End Sub

Private Sub dtpDato_Validate(Index As Integer, Cancel As Boolean)
Dim dnContador As Byte
   If Index = 3 Then
      If Month(dtpDato(3).Value) > Val(gsMesAct) And Year(dtpDato(3).Value) >= Val(gsAnoAct) Then
         MsgBox "La fecha No Corresponde al Periodo de Operacion", vbCritical
         dtpDato(Index).SetFocus
         Cancel = True
         Exit Sub
      End If
      If dtpDato(3).Tag <> dtpDato(3).Value Then
         dtpDato(3).Tag = dtpDato(3).Value
         With frmTVtaGrd.uorstTGTCb
            If .RecordCount <> 0 Then .MoveFirst
            .Find "(FehTCb) = '" & Format(dtpDato(3).Value, "yyyy/mm/dd") & "'"
            If .EOF Then
               MsgBox "No se ha ingresado Tipo de Cambio para esta fecha.", vbCritical
               Cancel = True
               Exit Sub
            Else
'               uorstMain!ImpTCb = IIf(frmTVtaGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
''               frmTVtaGrd.uorstMain!ImpTCb = !ImpTCb_Vta
               txtDato(4).Text = Format(!ImpTCb_Vta, FORMATO_NUM_2)
            End If
         End With
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
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
   If PreviousTab = 0 And sstMain.Tab = 1 Then
      ppDatosWhere
   End If
   dgrDetalle.SetFocus
End Sub

Private Sub ppAbreCtaCCo()
   With frmTVtaGrd.uorstCOVtaDocCCo
      frmTVtaGrd.usConnStrgWher_COVtaDocCCo = "WHERE COVtaDocCCo.CodTDc='" & frmTVtaGrd.uorstMain!CodTDc & "' AND COVtaDocCCo.SerDoc='" & frmTVtaGrd.uorstMain!SerDoc & "' AND COVtaDocCCo.NroDoc='" & frmTVtaGrd.uorstMain!NroDoc & "' "
      If .State = adStateOpen Then .Close
      .Source = frmTVtaGrd.usConnStrgSele_COVtaDocCCo & frmTVtaGrd.usConnStrgWher_COVtaDocCCo & frmTVtaGrd.usConnStrgOrde_COVtaDocCCo
      .Open
      .Properties("Unique Table").Value = "COVtaDocCCo"
   End With
   With frmTVtaGrd.uorstCOVtaDocCta
      frmTVtaGrd.usConnStrgWher_COVtaDocCta = "WHERE COVtaDocCta.CodTDc='" & frmTVtaGrd.uorstMain!CodTDc & "' AND COVtaDocCta.SerDoc='" & frmTVtaGrd.uorstMain!SerDoc & "'  AND COVtaDocCta.NroDoc='" & frmTVtaGrd.uorstMain!NroDoc & "' "
      If .State = adStateOpen Then .Close
      .Source = frmTVtaGrd.usConnStrgSele_COVtaDocCta & frmTVtaGrd.usConnStrgWher_COVtaDocCta & frmTVtaGrd.usConnStrgOrde_COVtaDocCta
      .Open
      .Properties("Unique Table").Value = "COVtaDocCta"
   End With
End Sub

Private Sub ppCambioTpoMon()
   Dim dnContador As Integer
   
   chkMonedaActiva.Caption = IIf(unVerMonNac = TPOMON_NAC_IND, TPOMON_NAC_TXT_2, TPOMON_EXT_TXT_2)
   For dnContador = MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1
      txtDato(dnContador).Visible = (unVerMonNac = TPOMON_NAC_IND)
   Next
   For dnContador = MINIMOINDICEIMPORTEME To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
      txtDato(dnContador).Visible = Not (unVerMonNac = TPOMON_NAC_IND)
   Next
End Sub

Private Sub ppGenera()
   Dim dnContador As Integer
   Dim dnNumeroItem As Integer
   Dim dbProcesaCuenta As Boolean

   If txtDato(1).Text <> "" Then
      ppDatosWhere

      With frmTVtaGrd.uorstCOCpbCab
        'Si existe, elimina Comprobante existente.
         If .RecordCount > 0 Then
            .MoveFirst
            .Find "cLlave='" & txtDato(0).Text & txtDato(1).Text & "'"
            If Not .EOF Then .Delete
         End If
      End With
   End If

   With frmTVtaGrd.uorstCOCpbCab
     'Si no está marcado para generar, marca el documento como no generado.
      If chkIndPreGen.Value = vbUnchecked Then
         frmTVtaGrd.uorstMain!IndGen = False
         frmTVtaGrd.uorstMain.Update
         Exit Sub
      End If

     'Captura del Siguiente Número.
      If txtDato(1).Text = "" Then
         With frmTVtaGrd.uorstCODro
            If .RecordCount <> 0 Then .MoveFirst
            .Find "CodDro = '" & txtDato(0).Text & "'"
            If IsNull(.Fields(2).Value) Then .Fields(2).Value = gfCeros("", .Fields(2).DefinedSize, 0, "0")
            txtDato(1).Text = gfCeros(.Fields(2).Value, .Fields(2).DefinedSize, 1, "0")
            .Fields(2).Value = txtDato(1).Text
            .Update
         End With
         frmTVtaGrd.uorstMain!NroCpb = txtDato(1).Text
         frmTVtaGrd.uorstMain.Update
      End If
      
      ppDatosWhere
   
     'Si no hay cuentas, marca el documento como no generado.
      If frmTVtaGrd.uorstCOVtaDocCta.RecordCount = 0 Then
         frmTVtaGrd.uorstMain!IndGen = False
         frmTVtaGrd.uorstMain.Update
         Exit Sub
      End If

     'Crea encabezado de Comprobante.
      .AddNew
      !MesPvs = gsMesAct
      !CodDro = txtDato(0).Text
      !NroCpb = txtDato(1).Text
      !FehCpb = dtpDato(3).Value
      !TpoGnr = TPOGNR_VTA
      !IndNCu = INDNCU_FAL
      !GloCpb = txtDato(3).Text
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With

   With frmTVtaGrd.uorstCOVtaDocCta
     'Crea ítemes de Comprobante.
      .MoveFirst
      Do
         dbProcesaCuenta = True

        'Itemes con Centro de Costo.
         If !TpoCnc <= CUENTASCONCCOSTO Then
            With frmTVtaGrd.uorstCOVtaDocCCo
               If .RecordCount <> 0 Then
                  .MoveFirst
                  .Find "cLlave = " & Trim(frmTVtaGrd.uorstCOVtaDocCta!TpoCnc) & frmTVtaGrd.uorstCOVtaDocCta!CodCta
                  If Not .EOF Then
                     Do
                        dnNumeroItem = dnNumeroItem + 1
                        Call ppGenera1(True, dnNumeroItem)
                        .MoveNext
                        If .EOF Then Exit Do
                        If !cLlave <> Trim(frmTVtaGrd.uorstCOVtaDocCta!TpoCnc) & frmTVtaGrd.uorstCOVtaDocCta!CodCta Then Exit Do
                     Loop
                     dbProcesaCuenta = False
                  End If
               End If
            End With
         End If

        'Itemes sin Centro de Costo.
         If dbProcesaCuenta Then
            dnNumeroItem = dnNumeroItem + 1
            Call ppGenera1(False, dnNumeroItem)
         End If

         .MoveNext
      Loop Until .EOF
   End With

   frmTVtaGrd.uorstMain!IndGen = True
   txtDato(0).Enabled = False
   txtDato(1).Enabled = False
   cmdDatoAyud(0).Enabled = False
   lblDatoDeta(0).Enabled = False

   frmTVtaGrd.uorstCOCpbCab.Update
   frmTVtaGrd.uorstCOCpbDet.UpdateBatch
   frmTVtaGrd.uorstMain.Update
End Sub

Private Sub ppGenera1(tbCCosto As Boolean, tnNumeroItem As Integer)
   With frmTVtaGrd.uorstCOCpbDet
      .AddNew
      !CodDro = txtDato(0).Text
      !NroCpb = txtDato(1).Text
      !NroIte = tnNumeroItem
      !MesPvs = gsMesAct
      !CodCta = frmTVtaGrd.uorstCOVtaDocCta!CodCta
      !FehOpe = dtpDato(3).Value
      frmTVtaGrd.uorstCOCta.MoveFirst
      frmTVtaGrd.uorstCOCta.Find "CodCta='" & frmTVtaGrd.uorstCOVtaDocCta!CodCta & "'"
      If frmTVtaGrd.uorstCOCta!IndCCo = INDCCO_ACT Then If tbCCosto Then !CodCCo = frmTVtaGrd.uorstCOVtaDocCCo!CodCCo
      If frmTVtaGrd.uorstCOCta!IndDoc = INDDOC_ACT Then
         !CodTDc = txtLlave(0).Text
         !SerDoc = txtLlave(1).Text
         !NroDoc = txtLlave(2).Text
         !FeEDoc = dtpDato(0).Value
         !FeVDoc = dtpDato(1).Value
         !FeRDoc = dtpDato(0).Value
         !RefDoc = txtDato(2).Text
         !CodAux = txtDato(33).Text
      End If
      !GloIte = txtDato(3).Text
      !TpoCtb = IIf(frmTVtaGrd.uorstCOVtaDocCta!TpoCnc = TPOCNC_TOT_VTA, IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      !TpoMon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      !ImpTCb = txtDato(4).Text
      If tbCCosto Then
'         !ImpMN = frmTVtaGrd.uorstCOVtaDocCCo!ImpCCo * IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, txtDato(4).Text)
'         !ImpME = frmTVtaGrd.uorstCOVtaDocCCo!ImpCCo / IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, txtDato(4).Text, 1)
         !ImpMN = frmTVtaGrd.uorstCOVtaDocCCo!ImpCCo_MN
         !ImpME = frmTVtaGrd.uorstCOVtaDocCCo!ImpCCo_ME
      Else
'         !ImpMN = frmTVtaGrd.uorstCOVtaDocCta!ImpCta * IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, txtDato(4).Text)
'         !ImpME = frmTVtaGrd.uorstCOVtaDocCta!ImpCta / IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, txtDato(4).Text, 1)
         !ImpMN = frmTVtaGrd.uorstCOVtaDocCta!ImpCta_MN
         !ImpME = frmTVtaGrd.uorstCOVtaDocCta!ImpCta_ME
      End If
      !TpoGnr = TPOGNR_VTA
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With
End Sub

Public Sub upActualizaMas(pnIndice As Byte, pnValor As Byte)
   frmTVta.cmdMas(pnIndice).Tag = pnValor 'Necesaria la referencia por ser llamado externamente.
   With frmTVtaGrd.uorstMain
      Select Case pnIndice
      Case 1
         !IndCta_OGr = pnValor
      Case 2
         !IndCta_Exp = pnValor
      Case 3
         !IndCta_Exo = pnValor
      Case 4
         !IndCta_IGV = pnValor
      Case 5
         !IndCta_ISC = pnValor
      Case 6
         !IndCta_OIm = pnValor
      Case 7
         !IndCta_Tot = pnValor
      End Select
   End With
End Sub

Public Sub upHabilitaCuenta(tbHabilita As Boolean, tnIndice As Byte)
   txtDato(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
   lblDatoDeta(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
   cmdDatoAyud(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
   If tnIndice <= CUENTASCONCCOSTO Then
      Call upHabilitaCCosto(tbHabilita, tnIndice)
   End If
End Sub

Public Sub upHabilitaCCosto(tbHabilita As Boolean, tnIndice As Byte)
   
   If Not tbHabilita Or txtDato(tnIndice + DIFERENCIAMASCUENTA).Text = "" Then
      txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
      lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
      cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
   Else
      frmTVtaGrd.uorstCOCta.MoveFirst
      frmTVtaGrd.uorstCOCta.Find "CodCta='" & txtDato(tnIndice + DIFERENCIAMASCUENTA).Text & "'"
      txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTVtaGrd.uorstCOCta!IndCCo = INDCCO_ACT)
      lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTVtaGrd.uorstCOCta!IndCCo = INDCCO_ACT)
      cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTVtaGrd.uorstCOCta!IndCCo = INDCCO_ACT)
   End If
End Sub

Private Sub ppDatosWhere()             'Cambiar.
   With frmTVtaGrd
      .uorstCOCpbCab.Requery
   
      .usConnStrgWher_COCpbDet = "WHERE COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='" & txtDato(0).Text & "' AND COCpbDet.NroCpb='" & txtDato(1).Text & "' "
      With .uorstCOCpbDet
         .Close
         .Source = frmTVtaGrd.usConnStrgSele_COCpbDet & frmTVtaGrd.usConnStrgWher_COCpbDet & frmTVtaGrd.usConnStrgOrde_COCpbDet
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
            .Item(dnNum).Caption = "Cuenta"
            .Item(dnNum).Width = 900
         Case 1
            .Item(dnNum).Caption = "Auxiliar"
            .Item(dnNum).Width = 1150
         Case 2
            .Item(dnNum).Caption = "C.Cto."
            .Item(dnNum).Width = 600
         Case 3
            .Item(dnNum).Caption = "Glosa"
            .Item(dnNum).Width = 1500
         Case 4
            .Item(dnNum).Caption = "Debe " & TPOMON_NAC_TXT_0
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 5
            .Item(dnNum).Caption = "Haber " & TPOMON_NAC_TXT_0
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 6
            .Item(dnNum).Caption = "Debe " & TPOMON_EXT_TXT_0
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 7
            .Item(dnNum).Caption = "Haber " & TPOMON_EXT_TXT_0
            .Item(dnNum).Width = 1000
            .Item(dnNum).Alignment = dbgRight
            .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
         Case 8
            .Item(dnNum).Caption = "Ds"
            .Item(dnNum).Width = 330
            .Item(dnNum).Alignment = dbgCenter
         Case Else
            .Item(dnNum).Visible = False
         End Select
      Next
   End With
End Sub

Private Function ValidaCtasCCo() As Boolean
   Dim dnContador, dnIndCCo As Byte
   Dim dvRegistroActual As Variant
   Dim dnTotalCuentaMN, dnTotalCuentaME, dnTotalImporteMN, dnTotalImporteME As Double
    
   ValidaCtasCCo = True
   
   For dnContador = INDMASCTA_INI To CANTIDADIMPORTES - 1
''       If (CDec(txtDato(MINIMOINDICEIMPORTEMN + dncontador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dncontador).Text) <> 0) And _
''        (cmdMas(dncontador + 1).Tag = INDMASCTA_INI And Len(Trim(txtDato(dncontador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
''        cmdMas(dncontador + 1).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dncontador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
       If ((CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) And _
          (cmdMas(dnContador + 1).Tag = INDMASCTA_MAS And Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
          ValidaCtasCCo = Not (txtDato(MINIMOINDICECUENTA + dnContador).Text = "")
          If Not ValidaCtasCCo Then Exit Function
          
          If frmTVtaGrd.ubGrabaMas = 1 Then
             With frmTVtaGrd.uorstCOVtaDocCta
                If Not (.BOF Or .EOF) And .RecordCount > 0 Then
                   dvRegistroActual = .Bookmark
                   dnTotalCuentaMN = 0
                   dnTotalCuentaME = 0
                   .MoveFirst
                   Do
                      dnIndCCo = 0
                      If !TpoCnc = Trim(Str(dnContador + 1)) Then
                         dnTotalCuentaMN = dnTotalCuentaMN + !ImpCta_MN
                         dnTotalCuentaME = dnTotalCuentaME + !ImpCta_ME
                         With frmTVtaGrd.uorstCOCta
                            .MoveFirst
                            .Find "CodCta='" & frmTVtaGrd.uorstCOVtaDocCta!CodCta & "'"
                            If Not .EOF Then
                               dnIndCCo = frmTVtaGrd.uorstCOCta!IndCCo
                            End If
                         End With
                      End If
                      If dnIndCCo = INDCCO_ACT Then
                         With frmTVtaGrd.uorstCOVtaDocCCo
                            If .State = adStateOpen Then .Close
                            frmTVtaGrd.usConnStrgWher_COVtaDocCCo = "WHERE COVtaDocCCo.SerDoc='" & frmTVtaGrd.uorstMain!SerDoc & "' And COVtaDocCCo.NroDoc='" & frmTVtaGrd.uorstMain!NroDoc & "' And COVtaDocCCo.TpoCnc='" & Trim(Str(dnContador + 1)) & "' And COVtaDocCCo.CodCta='" & frmTVtaGrd.uorstCOVtaDocCta!CodCta & "' "
                            .Source = frmTVtaGrd.usConnStrgSele_COVtaDocCCo & frmTVtaGrd.usConnStrgWher_COVtaDocCCo & frmTVtaGrd.usConnStrgOrde_COVtaDocCCo
                            .Open
                            If .RecordCount = 0 Then
                               MsgBox "Cuenta " & frmTVtaGrd.uorstCOVtaDocCta!CodCta & " requiere C.Costo", vbInformation
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
'               dnTotalImporteMN = gfRedond(CDec(txtDato(5).Text) + CDec(txtDato(6).Text) + CDec(txtDato(7).Text) + CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text) + CDec(txtDato(11).Text), 2)
'               dnTotalImporteME = gfRedond(CDec(txtDato(12).Text) + CDec(txtDato(13).Text) + CDec(txtDato(14).Text) + CDec(txtDato(15).Text) + CDec(txtDato(16).Text) + CDec(txtDato(17).Text) + CDec(txtDato(18).Text), 2)
             dnTotalImporteMN = gfRedond(CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text), 2)
             dnTotalImporteME = gfRedond(CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text), 2)
             If Not (CDec(dnTotalCuentaMN) = CDec(dnTotalImporteMN)) Then
                ValidaCtasCCo = False
                Exit Function
             End If
             If Not (CDec(dnTotalCuentaME) = CDec(dnTotalImporteME)) Then
                ValidaCtasCCo = False
                Exit Function
             End If
          End If
       ElseIf Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0 Then
          dnIndCCo = 0
          With frmTVtaGrd.uorstCOCta
             .MoveFirst
             .Find "CodCta='" & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & "'"
             If Not .EOF Then
                dnIndCCo = frmTVtaGrd.uorstCOCta!IndCCo
             End If
          End With
          If dnIndCCo = INDCCO_ACT And txtDato(dnContador + 1 + DIFERENCIAMASCUENTA + CUENTASCONCCOSTO).Text = "" Then
             MsgBox "Cuenta " & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & " requiere C.Costo", vbInformation
'             MsgBox "Cuenta " & frmTVtaGrd.uorstCOVtaDocCta!CodCta & " requiere C.Costo", vbInformation
             ValidaCtasCCo = False
             Exit Function
          End If
       ElseIf Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) = 0 And ((CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) Then
         ValidaCtasCCo = False
       End If
    Next dnContador
    ' Valido que los detalles sean iguales los importes
''    With frmTVtaGrd.uorstCOVtaDocCta
''       If Not (.BOF Or .EOF) And .RecordCount > 0 Then
''          dvRegistroActual = .Bookmark
''          .MoveFirst
''          Do
''            dnTotalCuentaMN = dnTotalCuentaMN + !ImpCta_MN
''            dnTotalCuentaME = dnTotalCuentaME + !ImpCta_ME
''            .MoveNext
''          Loop Until .EOF
''          .Bookmark = dvRegistroActual
''       End If
''    End With
''    dnTotalImporteMN = gfRedond(CDec(txtDato(5).Text) + CDec(txtDato(6).Text) + CDec(txtDato(7).Text) + CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text) + CDec(txtDato(11).Text), 2)
''    dnTotalImporteME = gfRedond(CDec(txtDato(12).Text) + CDec(txtDato(13).Text) + CDec(txtDato(14).Text) + CDec(txtDato(15).Text) + CDec(txtDato(16).Text) + CDec(txtDato(17).Text) + CDec(txtDato(18).Text), 2)
''    If Not (CDec(dnTotalCuentaMN) = CDec(dnTotalImporteMN)) Then
''        ValidaCtasCCo = False
''        Exit Function
''    End If
''    If Not (CDec(dnTotalCuentaME) = CDec(dnTotalImporteME)) Then
''        ValidaCtasCCo = False
''        Exit Function
''    End If
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

