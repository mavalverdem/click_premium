VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTCpr 
   Caption         =   "[Título]"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Tic&ket"
      Height          =   495
      Left            =   7320
      TabIndex        =   147
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "&Proveedor"
      Height          =   375
      Left            =   8250
      TabIndex        =   140
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   1
      Left            =   3000
      TabIndex        =   137
      Top             =   2640
      Width           =   4995
      Begin VB.CheckBox chkIndCDt 
         Caption         =   "Detracción"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   218
         Width           =   1095
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
         Index           =   37
         Left            =   2040
         TabIndex        =   17
         Top             =   180
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDato 
         Height          =   315
         Index           =   4
         Left            =   3600
         TabIndex        =   18
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61538305
         CurrentDate     =   37102
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Form. Nº"
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
         Left            =   1380
         TabIndex        =   139
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         Left            =   3105
         TabIndex        =   138
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CheckBox chkIndPreGen 
      Caption         =   "Cuentas &Registradas"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   6960
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CheckBox chkCalcularIGV 
      Caption         =   "Calcular I.G.&V."
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   60
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CheckBox chkCalcularISC 
      Caption         =   "Calcular I.S.&C."
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   1500
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   132
      Top             =   2040
      Width           =   7155
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4860
         Picture         =   "frmTCpr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   195
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
         Left            =   600
         TabIndex        =   12
         Top             =   180
         Width           =   555
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
         Left            =   6315
         TabIndex        =   13
         Top             =   180
         Width           =   735
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
         TabIndex        =   136
         Top             =   180
         Width           =   3735
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
         TabIndex        =   135
         Top             =   240
         Width           =   1005
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
         TabIndex        =   134
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   8640
      ScaleHeight     =   2610
      ScaleWidth      =   885
      TabIndex        =   125
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
         Picture         =   "frmTCpr.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   68
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
         Picture         =   "frmTCpr.frx":02F4
         Style           =   1  'Graphical
         TabIndex        =   67
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
         Picture         =   "frmTCpr.frx":03F6
         Style           =   1  'Graphical
         TabIndex        =   66
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
         Picture         =   "frmTCpr.frx":04F8
         Style           =   1  'Graphical
         TabIndex        =   65
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
         Picture         =   "frmTCpr.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   63
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
         Picture         =   "frmTCpr.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   64
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
      Left            =   2580
      TabIndex        =   11
      Top             =   1740
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
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   5820
      Picture         =   "frmTCpr.frx":0996
      Style           =   1  'Graphical
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   495
      Width           =   255
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7680
      Picture         =   "frmTCpr.frx":0B40
      Style           =   1  'Graphical
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   135
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
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTCpr.frx":0CEA
      Left            =   1020
      List            =   "frmTCpr.frx":0CEC
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1740
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   3060
      TabIndex        =   5
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   37102
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
      Left            =   3840
      TabIndex        =   9
      Top             =   1380
      Width           =   4635
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
      TabIndex        =   8
      Top             =   1380
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
      Index           =   3
      Left            =   7260
      TabIndex        =   3
      Top             =   480
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
      Index           =   2
      Left            =   6840
      TabIndex        =   2
      Top             =   480
      Width           =   435
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   1
      Left            =   5100
      TabIndex        =   6
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   37102
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   2
      Left            =   7200
      TabIndex        =   7
      Top             =   1020
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   37102
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   3
      Left            =   1020
      TabIndex        =   4
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   37102
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   3135
      Left            =   0
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3240
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   441
      ForeColor       =   8388608
      TabCaption(0)   =   "I&mportes"
      TabPicture(0)   =   "frmTCpr.frx":0CEE
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
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDatoDeta(21)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDatoDeta(22)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDatoDeta(23)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDatoDeta(24)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDatoDeta(25)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDatoDeta(26)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDatoDeta(27)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDatoDeta(28)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDatoDeta(29)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblDatoDeta(30)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblDatoDeta(31)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblDatoDeta(32)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblDatoDeta(33)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblDatoDeta(34)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblDatoDeta(35)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblDatoDeta(36)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkDesactivar"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdDatoAyud(21)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtDato(21)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtDato(22)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdDatoAyud(22)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdDatoAyud(23)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtDato(23)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdDatoAyud(24)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtDato(24)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdDatoAyud(25)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtDato(25)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmdDatoAyud(26)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtDato(26)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdDatoAyud(27)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtDato(27)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdDatoAyud(28)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtDato(28)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdDatoAyud(29)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtDato(29)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdDatoAyud(30)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtDato(30)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cmdDatoAyud(31)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtDato(31)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cmdDatoAyud(32)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtDato(32)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cmdDatoAyud(33)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtDato(33)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "cmdDatoAyud(34)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtDato(34)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "cmdMas(1)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmdMas(2)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cmdMas(3)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cmdMas(4)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "cmdMas(5)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmdMas(6)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmdMas(7)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "cmdMas(8)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "chkMonedaActiva"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cmdDatoAyud(36)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "cmdDatoAyud(35)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txtDato(35)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txtDato(36)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtDato(13)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txtDato(14)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtDato(15)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txtDato(16)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txtDato(17)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txtDato(18)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txtDato(19)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txtDato(20)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "txtDato(5)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "txtDato(6)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txtDato(7)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "txtDato(8)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "txtDato(9)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "txtDato(10)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "txtDato(11)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txtDato(12)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "txtDato(38)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "txtDato(39)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "txtDato(40)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "txtDato(41)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txtDato(42)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "txtDato(43)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "cmdMasIGV"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).ControlCount=   91
      TabCaption(1)   =   "C&uentas"
      TabPicture(1)   =   "frmTCpr.frx":0D0A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dgrDetalle"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdMasIGV 
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
         Left            =   1000
         Picture         =   "frmTCpr.frx":0D26
         TabIndex        =   42
         Top             =   1825
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
         Index           =   43
         Left            =   0
         TabIndex        =   146
         Top             =   0
         Visible         =   0   'False
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
         Index           =   42
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
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
         Index           =   41
         Left            =   0
         TabIndex        =   144
         Top             =   0
         Visible         =   0   'False
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
         Index           =   40
         Left            =   0
         TabIndex        =   143
         Top             =   0
         Visible         =   0   'False
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
         Index           =   39
         Left            =   0
         TabIndex        =   142
         Top             =   0
         Visible         =   0   'False
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
         Index           =   38
         Left            =   0
         TabIndex        =   141
         Top             =   0
         Visible         =   0   'False
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
         TabIndex        =   58
         Top             =   2700
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
         TabIndex        =   53
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
         TabIndex        =   48
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
         TabIndex        =   43
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
         TabIndex        =   37
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
         TabIndex        =   32
         Top             =   1200
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
         Index           =   6
         Left            =   1320
         TabIndex        =   27
         Top             =   900
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
         Index           =   5
         Left            =   1320
         TabIndex        =   22
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
         Index           =   20
         Left            =   1320
         TabIndex        =   59
         Top             =   2700
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
         Index           =   19
         Left            =   1320
         TabIndex        =   54
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
         Index           =   18
         Left            =   1320
         TabIndex        =   49
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
         Index           =   17
         Left            =   1320
         TabIndex        =   44
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
         Index           =   16
         Left            =   1320
         TabIndex        =   38
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
         Index           =   15
         Left            =   1320
         TabIndex        =   33
         Top             =   1200
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
         TabIndex        =   28
         Top             =   900
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
         TabIndex        =   23
         Top             =   600
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
         Index           =   36
         Left            =   6840
         TabIndex        =   62
         Top             =   2700
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
         Index           =   35
         Left            =   6840
         TabIndex        =   57
         Top             =   2400
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   35
         Left            =   9060
         Picture         =   "frmTCpr.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   2430
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   36
         Left            =   9060
         Picture         =   "frmTCpr.frx":0FD2
         Style           =   1  'Graphical
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   2730
         Width           =   255
      End
      Begin VB.CheckBox chkMonedaActiva 
         Caption         =   "M&oneda activa"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   330
         Width           =   1635
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
         Index           =   8
         Left            =   3000
         Picture         =   "frmTCpr.frx":117C
         TabIndex        =   60
         Top             =   2725
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
         Index           =   7
         Left            =   3000
         Picture         =   "frmTCpr.frx":127E
         TabIndex        =   55
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
         Picture         =   "frmTCpr.frx":1380
         TabIndex        =   50
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
         Picture         =   "frmTCpr.frx":1482
         TabIndex        =   45
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
         Picture         =   "frmTCpr.frx":1584
         TabIndex        =   39
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
         Picture         =   "frmTCpr.frx":1686
         TabIndex        =   34
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
         Picture         =   "frmTCpr.frx":1788
         TabIndex        =   29
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
         Picture         =   "frmTCpr.frx":188A
         TabIndex        =   24
         Top             =   625
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   -70170
         ScaleHeight     =   270
         ScaleWidth      =   1575
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   75
         Visible         =   0   'False
         Width           =   1575
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
            TabIndex        =   124
            Top             =   0
            Visible         =   0   'False
            Width           =   700
         End
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
            TabIndex        =   123
            Top             =   0
            Visible         =   0   'False
            Width           =   700
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
         Index           =   34
         Left            =   6840
         TabIndex        =   52
         Top             =   2100
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   34
         Left            =   9060
         Picture         =   "frmTCpr.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   119
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
         Index           =   33
         Left            =   6840
         TabIndex        =   47
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   33
         Left            =   9060
         Picture         =   "frmTCpr.frx":1B36
         Style           =   1  'Graphical
         TabIndex        =   117
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
         Index           =   32
         Left            =   6840
         TabIndex        =   41
         Top             =   1500
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   32
         Left            =   9060
         Picture         =   "frmTCpr.frx":1CE0
         Style           =   1  'Graphical
         TabIndex        =   115
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
         Index           =   31
         Left            =   6840
         TabIndex        =   36
         Top             =   1200
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   31
         Left            =   9060
         Picture         =   "frmTCpr.frx":1E8A
         Style           =   1  'Graphical
         TabIndex        =   113
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
         Index           =   30
         Left            =   6840
         TabIndex        =   31
         Top             =   900
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   30
         Left            =   9060
         Picture         =   "frmTCpr.frx":2034
         Style           =   1  'Graphical
         TabIndex        =   111
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
         Index           =   29
         Left            =   6840
         TabIndex        =   26
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   29
         Left            =   9060
         Picture         =   "frmTCpr.frx":21DE
         Style           =   1  'Graphical
         TabIndex        =   109
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
         Index           =   28
         Left            =   3300
         TabIndex        =   61
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   28
         Left            =   6540
         Picture         =   "frmTCpr.frx":2388
         Style           =   1  'Graphical
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   2725
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
         Left            =   3300
         TabIndex        =   56
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   27
         Left            =   6540
         Picture         =   "frmTCpr.frx":2532
         Style           =   1  'Graphical
         TabIndex        =   105
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
         Index           =   26
         Left            =   3300
         TabIndex        =   51
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   26
         Left            =   6540
         Picture         =   "frmTCpr.frx":26DC
         Style           =   1  'Graphical
         TabIndex        =   103
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
         Index           =   25
         Left            =   3300
         TabIndex        =   46
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   25
         Left            =   6540
         Picture         =   "frmTCpr.frx":2886
         Style           =   1  'Graphical
         TabIndex        =   101
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
         Index           =   24
         Left            =   3300
         TabIndex        =   40
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   24
         Left            =   6540
         Picture         =   "frmTCpr.frx":2A30
         Style           =   1  'Graphical
         TabIndex        =   99
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
         Index           =   23
         Left            =   3300
         TabIndex        =   35
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   23
         Left            =   6540
         Picture         =   "frmTCpr.frx":2BDA
         Style           =   1  'Graphical
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1225
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   22
         Left            =   6540
         Picture         =   "frmTCpr.frx":2D84
         Style           =   1  'Graphical
         TabIndex        =   95
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
         Index           =   22
         Left            =   3300
         TabIndex        =   30
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
         Index           =   21
         Left            =   3300
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   21
         Left            =   6540
         Picture         =   "frmTCpr.frx":2F2E
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   625
         Width           =   255
      End
      Begin VB.CheckBox chkDesactivar 
         Caption         =   "Des&activar Cuentas"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   4980
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   88
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   180
         TabIndex        =   131
         Top             =   1260
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   180
         TabIndex        =   130
         Top             =   960
         Width           =   1080
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
         Index           =   36
         Left            =   7500
         TabIndex        =   128
         Top             =   2700
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
         Index           =   35
         Left            =   7500
         TabIndex        =   129
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
         Index           =   34
         Left            =   7500
         TabIndex        =   120
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
         Index           =   33
         Left            =   7500
         TabIndex        =   118
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
         Index           =   32
         Left            =   7500
         TabIndex        =   116
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
         Index           =   31
         Left            =   7500
         TabIndex        =   114
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
         Index           =   30
         Left            =   7500
         TabIndex        =   112
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
         Index           =   29
         Left            =   7500
         TabIndex        =   110
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
         Index           =   28
         Left            =   4260
         TabIndex        =   108
         Top             =   2700
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
         Index           =   27
         Left            =   4260
         TabIndex        =   106
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
         Index           =   26
         Left            =   4260
         TabIndex        =   104
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
         Index           =   25
         Left            =   4260
         TabIndex        =   102
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
         Index           =   24
         Left            =   4260
         TabIndex        =   100
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
         Index           =   23
         Left            =   4260
         TabIndex        =   98
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
         Index           =   22
         Left            =   4260
         TabIndex        =   96
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
         Index           =   21
         Left            =   4260
         TabIndex        =   94
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
         TabIndex        =   87
         Top             =   1860
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
         TabIndex        =   86
         Top             =   2160
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
         TabIndex        =   85
         Top             =   2460
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Exonerado.:"
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
         TabIndex        =   84
         Top             =   1560
         Width           =   870
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
         Top             =   2760
         Width           =   555
      End
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
      TabIndex        =   121
      Top             =   1080
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8400
      Y1              =   900
      Y2              =   900
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
      Left            =   1200
      TabIndex        =   92
      Top             =   480
      Width           =   4635
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
      Left            =   2160
      TabIndex        =   90
      Top             =   120
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
      Left            =   3360
      TabIndex        =   79
      Top             =   1440
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
      TabIndex        =   78
      Top             =   1440
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
      TabIndex        =   77
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "F.Recepc.:"
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
      Left            =   6420
      TabIndex        =   76
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "F.Vcmto.:"
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
      Left            =   4380
      TabIndex        =   75
      Top             =   1080
      Width           =   690
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
      Left            =   1860
      TabIndex        =   74
      Top             =   1800
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
      Left            =   2340
      TabIndex        =   73
      Top             =   1080
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
      TabIndex        =   72
      Top             =   540
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
      TabIndex        =   71
      Top             =   540
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor:"
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
      TabIndex        =   70
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmTCpr"
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
              MINIMOINDICEIMPORTEME As Byte = 13, _
              MINIMOINDICEMAS As Byte = 1, _
              MINIMOINDICECUENTA As Byte = 21, _
              MINIMOINDICECCOSTO As Byte = 29, _
              CANTIDADIMPORTES As Byte = 8
'[Repetir en frmTCprMasGrd.
Private Const DIFERENCIAMASIMPORTE As Byte = 4, _
              DIFERENCIAMASCUENTA As Byte = 20, _
              DIFERENCIAMASCCOSTO As Byte = 28
Private Const CUENTASCONCCOSTO As Byte = 8
']

'[Repetir en frmTCprGrd y frmTCprMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
']

Private Sub chkIndCDt_Click()
    If Not (chkIndCDt.Value = vbChecked) Then TxtDato(37).Text = ""
    TxtDato(37).Enabled = (chkIndCDt.Value)
    dtpDato(4).Enabled = (chkIndCDt.Value)
End Sub

Private Sub cmdMasIGV_Click()
   
   frmTCprMasIgv.Show vbModal

End Sub

Private Sub Command1_Click()
'txtDato(5).Text = 0
'txtDato(8).Text = 0
'txtDato(13).Text = 0
'txtDato(16).Text = 0
'
'If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
'    txtDato(5).Text = Format(CDec(txtDato(9).Text * 100) / 19, FORMATO_NUM_1)
'    txtDato(8).Text = Format(CDec(txtDato(12).Text - txtDato(5).Text - txtDato(9).Text), FORMATO_NUM_1)
'    txtDato(13).Text = Format(CDec(txtDato(5) / txtDato(4)), FORMATO_NUM_1)
'    txtDato(16).Text = Format(CDec(txtDato(8) / txtDato(4)), FORMATO_NUM_1)
'Else
'    txtDato(13).Text = Format(CDec(txtDato(17).Text * 100) / 19, FORMATO_NUM_1)
'    txtDato(16).Text = Format(CDec(txtDato(20).Text - txtDato(13).Text - txtDato(17).Text), FORMATO_NUM_1)
'    txtDato(5).Text = Format(CDec(txtDato(13) * txtDato(4)), FORMATO_NUM_1)
'    txtDato(8).Text = Format(CDec(txtDato(16) * txtDato(4)), FORMATO_NUM_1)
'End If

  'reverción del igv y calculo de la base exonerada
         Dim TOTMN, TOTME As Byte
         If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
            TOTMN = Val(TxtDato(5).Text) + Val(TxtDato(6).Text) + Val(TxtDato(7).Text) + Val(TxtDato(8).Text)
            'If TOTMN = 0 Then
            If TOTMN = 0 And Val(TxtDato(9).Text) > 0 And Val(TxtDato(12).Text) > 0 Then
                Dim XRESPMN As String
                XRESPMN = MsgBox("No ha registrado importes Grabados ni Exonerados deseas calcularlos", vbYesNo, "Consulta M.N")
                If XRESPMN = 6 Then
                    TxtDato(5).Text = 0
                    TxtDato(8).Text = 0
                    TxtDato(13).Text = 0
                    TxtDato(16).Text = 0
                    TxtDato(5).Text = Format(CDec(TxtDato(9).Text * 100) / CDec(gnPctIGV), FORMATO_NUM_1)
                    TxtDato(8).Text = Format(CDec(TxtDato(12).Text - TxtDato(5).Text - TxtDato(9).Text), FORMATO_NUM_1)
                    TxtDato(13).Text = Format(CDec(TxtDato(5) / TxtDato(4)), FORMATO_NUM_1)
                    TxtDato(16).Text = Format(CDec(TxtDato(8) / TxtDato(4)), FORMATO_NUM_1)
                End If
            End If
         Else
            TOTME = Val(TxtDato(13).Text) + Val(TxtDato(14).Text) + Val(TxtDato(15).Text) + Val(TxtDato(16).Text)
            'Print TOTME
            'If TOTME = 0 Then
             If TOTME = 0 And Val(TxtDato(17).Text) > 0 And Val(TxtDato(20).Text) > 0 Then
                Dim XRESPME As String
                XRESPME = MsgBox("No ha registrado importes Grabados ni Exonerados deseas calcularlos", vbYesNo, "Consulta M.E.")
                If XRESPME = 6 Then
                    TxtDato(5).Text = 0
                    TxtDato(8).Text = 0
                    TxtDato(13).Text = 0
                    TxtDato(16).Text = 0
                    TxtDato(13).Text = Format(CDec(TxtDato(17).Text * 100) / CDec(gnPctIGV), FORMATO_NUM_1)
                    TxtDato(16).Text = Format(CDec(TxtDato(20).Text - TxtDato(13).Text - TxtDato(17).Text), FORMATO_NUM_1)
                    TxtDato(5).Text = Format(CDec(TxtDato(13) * TxtDato(4)), FORMATO_NUM_1)
                    TxtDato(8).Text = Format(CDec(TxtDato(16) * TxtDato(4)), FORMATO_NUM_1)
                End If
            End If
         End If
        ' fin del p
End Sub

Private Sub Form_Load()
   pbValidada = False
   pbFecha = True
   Me.KeyPreview = True
   
   With frmTCprGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
      txtLlave(0).MaxLength = .uorstMain!CodAux.DefinedSize
      txtLlave(1).MaxLength = .uorstMain!CodTDc.DefinedSize
      txtLlave(2).MaxLength = .uorstMain!SerDoc.DefinedSize
      txtLlave(3).MaxLength = .uorstMain!NroDoc.DefinedSize
    ']
   
    '[Datos                            'Cambiar.
      With cboTpoMon
         .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
         .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
      End With
'      mskDato(0).MaxLength = .uorstMain!Tf1Cta.DefinedSize + 1
      TxtDato(0).MaxLength = .uorstMain!CodDro.DefinedSize
      TxtDato(1).MaxLength = .uorstMain!NroCpb.DefinedSize
      TxtDato(2).MaxLength = .uorstMain!RefDoc.DefinedSize
      TxtDato(3).MaxLength = .uorstMain!GloDoc.DefinedSize
      TxtDato(4).MaxLength = .uorstMain!ImpTCb.DefinedSize
      TxtDato(5).MaxLength = 16
      TxtDato(6).MaxLength = 16
      TxtDato(7).MaxLength = 16
      TxtDato(8).MaxLength = 16
      TxtDato(9).MaxLength = 16
      TxtDato(10).MaxLength = 16
      TxtDato(11).MaxLength = 16
      TxtDato(12).MaxLength = 16
      TxtDato(13).MaxLength = 16
      TxtDato(14).MaxLength = 16
      TxtDato(15).MaxLength = 16
      TxtDato(16).MaxLength = 16
      TxtDato(17).MaxLength = 16
      TxtDato(18).MaxLength = 16
      TxtDato(19).MaxLength = 16
      TxtDato(20).MaxLength = 16
      TxtDato(21).MaxLength = 8
      TxtDato(22).MaxLength = 8
      TxtDato(23).MaxLength = 8
      TxtDato(24).MaxLength = 8
      TxtDato(25).MaxLength = 8
      TxtDato(26).MaxLength = 8
      TxtDato(27).MaxLength = 8
      TxtDato(28).MaxLength = 8
      TxtDato(29).MaxLength = 5
      TxtDato(30).MaxLength = 5
      TxtDato(31).MaxLength = 5
      TxtDato(32).MaxLength = 5
      TxtDato(33).MaxLength = 5
      TxtDato(34).MaxLength = 5
      TxtDato(35).MaxLength = 5
      TxtDato(36).MaxLength = 5
      TxtDato(37).MaxLength = .uorstMain!NroCDt.DefinedSize
      TxtDato(38).MaxLength = 16
      TxtDato(39).MaxLength = 16
      TxtDato(40).MaxLength = 16
      TxtDato(41).MaxLength = 16
      TxtDato(42).MaxLength = 16
      TxtDato(43).MaxLength = 16
    ']
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
   Set dgrDetalle.DataSource = frmTCprGrd.uorstCOCpbDet
   
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
   TxtDato(4).Tag = TxtDato(4).Text
 ']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not frmTCprGrd.uorstMain.EOF Then
      If frmTCprGrd.uorstMain.EditMode <> adEditNone Then frmTCprGrd.uorstMain.CancelUpdate 'Cambiar Formulario de Grid.
   End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmTCprGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTCprGrd.uorstMain_Grd.MoveFirst
   frmTCprGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "'"
End Sub

Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmTCprGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTCprGrd.uorstMain_Grd.MoveFirst
   frmTCprGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "'"
End Sub

Public Sub cmdCorregir_Click()
   'Verificación de Mes Cerrado.
   If gbCieCpr Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
   pbCorregir = True
   frmTCprGrd.uocnnMain.BeginTrans     'Cambiar Formulario de Grid. 'INICIA TRANSACCION.

   cmdRetroceder.Enabled = False
   cmdAvanzar.Enabled = False
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   upHabilitacion (True)
   TxtDato(0).Enabled = (chkIndPreGen.Value = 0)
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
   
'  If txtDato(33).Text = "" Then
'     MsgBox TEXT_6002, vbCritical
'     txtDato(33).SetFocus
'     Exit Sub
'  End If
   
   With frmTCprGrd.uorstMain
      dnSumaMN = CDec(TxtDato(5).Text) + CDec(TxtDato(6).Text) + CDec(TxtDato(7).Text) + CDec(TxtDato(8).Text) + CDec(TxtDato(9).Text) + CDec(TxtDato(10).Text) + CDec(TxtDato(11).Text)
      dnSumaME = CDec(TxtDato(13).Text) + CDec(TxtDato(14).Text) + CDec(TxtDato(15).Text) + CDec(TxtDato(16).Text) + CDec(TxtDato(17).Text) + CDec(TxtDato(18).Text) + CDec(TxtDato(19).Text)
'      If gfRedond(CDec(TxtDato(5).Text) + CDec(TxtDato(6).Text) + CDec(TxtDato(7).Text) + CDec(TxtDato(8).Text) + CDec(TxtDato(9).Text) + CDec(TxtDato(10).Text) + CDec(TxtDato(11).Text), 2) <> CDec(TxtDato(12).Text) Then
      If dnSumaMN <> CDec(TxtDato(12).Text) Then
'         If MsgBox(TEXT_9011, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
         If MsgBox(TEXT_9011 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(dnSumaMN - CDec(TxtDato(12).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
            Exit Sub
         End If
'      ElseIf gfRedond(CDec(TxtDato(13).Text) + CDec(TxtDato(14).Text) + CDec(TxtDato(15).Text) + CDec(TxtDato(16).Text) + CDec(TxtDato(17).Text) + CDec(TxtDato(18).Text) + CDec(TxtDato(19).Text), 2) <> CDec(TxtDato(20).Text) Then
      ElseIf dnSumaME <> CDec(TxtDato(20).Text) Then
'         If MsgBox(TEXT_9012, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
         If MsgBox(TEXT_9012 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(dnSumaME - CDec(TxtDato(20).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
            Exit Sub
         End If
      End If
   End With
   'Valido las Cuentas esten Correctas(llenas para todas los valores)
   If chkIndPreGen.Value = vbChecked Then
      chkIndPreGen.Value = IIf(ValidaCtasCCo, 1, 0)
      If Not (chkIndPreGen.Value = vbChecked) Then
        If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
          Exit Sub
        End If
      End If
   End If
 ']

   With frmTCprGrd                     'Cambiar Formulario de Grid.
      If pbNuevo And frmTCprGrd.ubGrabaMas = 0 Then
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
      .uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "'"
    ']

      pbCorregir = False
   
      If pbNuevo Then
'         .uorstMain.Requery
'         .upDatosGrid
'       '[Búsqueda de llave actual.     'Cambiar.
'         .uorstMain.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "'"
'       ']
         pbValidada = False
         cmdGrabar.Enabled = False
         upHabilitacion False
         frmTCprGrd.ubGrabaMas = INDMASCTA_INI

         upDatosPredeterminados
         pbFecha = True
       '[Llave habilitar  'Cambiar.
         txtLlave(0).Enabled = True
         txtLlave(1).Enabled = True
         txtLlave(2).Enabled = True
         txtLlave(3).Enabled = True
         lblLlaveDeta(0).Enabled = True
         lblLlaveDeta(1).Enabled = True
         cmdLlaveAyud(0).Enabled = True
         cmdLlaveAyud(1).Enabled = True
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
  
   frmTCprGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
 '[Propio del formulario.
   frmTCprGrd.uorstCOCpbCab.CancelUpdate
   frmTCprGrd.uorstCOCpbDet.CancelBatch
   frmTCprGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
   pbCorregir = False
 ']
   
   gpTUe_Deshacer Me
End Sub

Public Sub cmdSalir_Click()
   If pbNuevo Or pbCorregir Then
      pbCorregir = False
      frmTCprGrd.uocnnMain.RollbackTrans 'RESTAURA TRANSACCION.
   End If

   Unload Me
End Sub

Private Sub cmdLlaveAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 1
      txtLlave(Index).SetFocus
'   Case 1
'      mskLlave(Index).SetFocus
   End Select
   ppAyuBus AYULLA, Index
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 0, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      TxtDato(Index).SetFocus
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
   If pbValidada Then                  'Cambiar.
'      dtpDato(3).SetFocus
      
      txtLlave(0).Enabled = False
      txtLlave(1).Enabled = False
      txtLlave(2).Enabled = False
      txtLlave(3).Enabled = False
      lblLlaveDeta(0).Enabled = False
      lblLlaveDeta(1).Enabled = False
      cmdLlaveAyud(0).Enabled = False
      cmdLlaveAyud(1).Enabled = False
   End If
   If pbValidada And dtpDato(3).Enabled Then dtpDato(3).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
'   On Error GoTo Err

   Dim dvRegistro As Variant
   
 '[Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index
   Case 1, 2, 3                        'Cambiar (añadir índices).
      If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
         txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
      End If
   End Select
 ']
   
 '[Busca el dato en su tabla principal.'Cambiar (habilitar/deshabilitar).
   Select Case Index
   Case 0, 1                           'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYULLA, Index)
      If Cancel Then Exit Sub
   End Select
 ']
 
 '[Valida la llave.                    'Cambiar.
   If Len(Trim(txtLlave(0).Text)) <> 0 And Len(Trim(txtLlave(1).Text)) <> 0 And Len(Trim(txtLlave(2).Text)) <> 0 And Len(Trim(txtLlave(3).Text)) <> 0 Then
      With frmTCprGrd                  'Cambiar Formulario de Grid.
         Set .uorstTemporal = .uocnnMain.Execute("SELECT MesPvs FROM COCprDoc WHERE CodAux='" & txtLlave(0).Text & "' AND CodTDc ='" & txtLlave(1).Text & "' AND SerDoc='" & txtLlave(2).Text & "' AND NroDoc='" & txtLlave(3).Text & "'")
         If .uorstTemporal.RecordCount > 0 Then
            MsgBox TEXT_8007 & Chr(13) & "(mes " & gfMesLet("01" & .uorstTemporal!MesPvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
            Cancel = True
            Exit Sub
         End If
         .uorstTemporal.Close
      End With
    '[Propio del formulario.
      If frmTCprGrd.ubGrabaMas = 0 Then
         frmTCprGrd.ubGrabaMas = 1
         With frmTCprGrd
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
   TxtDato(Index).SelStart = 0
   TxtDato(Index).SelLength = TxtDato(Index).MaxLength
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
'[ARREGLAR: Retrocede si Shift está presionado.
   If Len(Trim(TxtDato(Index))) + 1 = TxtDato(Index).MaxLength Then
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
''      If txtDato(Index).Tag <> frmTCprGrd.uorstCOCta!TpoTCb Or CDec(txtDato(Index).Text) = 0 Then
''         txtDato(Index).Tag = frmTCprGrd.uorstCOCta!TpoTCb
''         With frmTCprGrd.uorstTGTCb
''            If .RecordCount <> 0 Then .MoveFirst
''            .Find "FehTCb = '" & dtpDato(3).Value & "'"
'''            uorstMain!ImpTCb = IIf(frmTCprGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
''         End With
''      End If
   Case MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
      If CDec(TxtDato(4).Text) <= 0 Then
         MsgBox "No se ha ingresado Tipo de Cambio para Esta Fecha", vbCritical
         TxtDato(4).SetFocus
         Exit Sub
      End If
      If CDec(TxtDato(Index).Text) = 0 Then
         TxtDato(Index).Text = Format(0, FORMATO_NUM_1)
         If Index >= MINIMOINDICEIMPORTEMN And Index < MINIMOINDICEIMPORTEME And cboTpoMon.ListIndex = TPOMON_EXT_IND Then
            TxtDato(Index).Text = Format(CDec(TxtDato(Index + CANTIDADIMPORTES).Text) * CDec(TxtDato(4).Text), FORMATO_NUM_1)
         ElseIf Index >= MINIMOINDICEIMPORTEME And Index < (MINIMOINDICEIMPORTEME + CANTIDADIMPORTES) And cboTpoMon.ListIndex = TPOMON_NAC_IND Then
            TxtDato(Index).Text = Format(CDec(TxtDato(Index - CANTIDADIMPORTES).Text) / CDec(TxtDato(4).Text), FORMATO_NUM_1)
         End If
      End If
      If (Index >= 5 And Index <= 7) Then
         If chkCalcularIGV Then TxtDato(MINIMOINDICEIMPORTEMN + 4).Text = Format((CDec(TxtDato(5).Text) + CDec(TxtDato(6).Text) + CDec(TxtDato(7).Text)) * CDec(gnPctIGV) / 100, FORMATO_NUM_1)
         If chkCalcularISC Then TxtDato(MINIMOINDICEIMPORTEMN + 5).Text = Format((CDec(TxtDato(5).Text) + CDec(TxtDato(6).Text) + CDec(TxtDato(7).Text)) * CDec(gnPctISC) / 100, FORMATO_NUM_1)
         ' Calculo individual ma 31/01/2004
         If chkCalcularIGV Then TxtDato(33 + Index).Text = Format((CDec(TxtDato(Index).Text)) * CDec(gnPctIGV) / 100, FORMATO_NUM_1)
         If (chkMonedaActiva.Value = vbChecked) And (cboTpoMon.ListIndex = TPOMON_NAC_IND) Then
            If CDec(TxtDato(MINIMOINDICEIMPORTEMN + 4).Text) > 0 Then TxtDato(MINIMOINDICEIMPORTEME + 4).Text = Format(gfRedond(CDec(TxtDato(MINIMOINDICEIMPORTEMN + 4).Text) / CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
            If CDec(TxtDato(MINIMOINDICEIMPORTEMN + 5).Text) > 0 Then TxtDato(MINIMOINDICEIMPORTEME + 5).Text = Format(gfRedond(CDec(TxtDato(MINIMOINDICEIMPORTEMN + 5).Text) / CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
            ' Calculo individual ma 31/01/2004
            If CDec(TxtDato(33 + Index).Text) > 0 Then TxtDato(36 + Index).Text = Format(gfRedond(CDec(TxtDato(33 + Index).Text) / CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
         End If
      ElseIf (Index >= 13 And Index <= 15) Then
         If chkCalcularIGV Then TxtDato(MINIMOINDICEIMPORTEME + 4).Text = Format((CDec(TxtDato(13).Text) + CDec(TxtDato(14).Text) + CDec(TxtDato(15).Text)) * CDec(gnPctIGV) / 100, FORMATO_NUM_1)
         If chkCalcularISC Then TxtDato(MINIMOINDICEIMPORTEME + 5).Text = Format((CDec(TxtDato(13).Text) + CDec(TxtDato(14).Text) + CDec(TxtDato(15).Text)) * CDec(gnPctISC) / 100, FORMATO_NUM_1)
         ' Calculo individual ma 31/01/2004
         If chkCalcularIGV Then TxtDato(28 + Index).Text = Format((CDec(TxtDato(Index).Text)) * CDec(gnPctIGV) / 100, FORMATO_NUM_1)
         If (chkMonedaActiva.Value = vbChecked) And (cboTpoMon.ListIndex = TPOMON_EXT_IND) Then
            If CDec(TxtDato(MINIMOINDICEIMPORTEME + 4).Text) > 0 Then TxtDato(MINIMOINDICEIMPORTEMN + 4).Text = Format(gfRedond(CDec(TxtDato(MINIMOINDICEIMPORTEME + 4).Text) * CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
            If CDec(TxtDato(MINIMOINDICEIMPORTEME + 5).Text) > 0 Then TxtDato(MINIMOINDICEIMPORTEMN + 5).Text = Format(gfRedond(CDec(TxtDato(MINIMOINDICEIMPORTEME + 5).Text) * CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
            ' Calculo individual ma 31/01/2004
            If CDec(TxtDato(28 + Index).Text) > 0 Then TxtDato(25 + Index).Text = Format(gfRedond(CDec(TxtDato(28 + Index).Text) * CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
         End If
      End If
      
     'Cálculo del total.
      If (Index = 12 And TxtDato(Index).Text = 0) Or (Index = 20 And TxtDato(Index).Text = 0) Then
         If Index = 12 Then
            TxtDato(12).Text = Format(CDec(TxtDato(5).Text) + CDec(TxtDato(6).Text) + CDec(TxtDato(7).Text) + CDec(TxtDato(8).Text) + CDec(TxtDato(9).Text) + CDec(TxtDato(10).Text) + CDec(TxtDato(11).Text), FORMATO_NUM_1)
         Else
            TxtDato(20).Text = Format(CDec(TxtDato(13).Text) + CDec(TxtDato(14).Text) + CDec(TxtDato(15).Text) + CDec(TxtDato(16).Text) + CDec(TxtDato(17).Text) + CDec(TxtDato(18).Text) + CDec(TxtDato(19).Text), FORMATO_NUM_1)
         End If
      End If
   
      ' Miguel Angel 25/01/2004 Convierte el monto si es la moneda funcional
      ' Quito esto al momento de comvertir If CDec(txtDato(Index + CANTIDADIMPORTES).Text) = 0 Then
      If chkMonedaActiva.Value = vbChecked Then
         If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
            TxtDato(Index + CANTIDADIMPORTES).Text = Format(gfRedond(CDec(TxtDato(Index).Text) / CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
         Else
            TxtDato(Index - CANTIDADIMPORTES).Text = Format(gfRedond(CDec(TxtDato(Index).Text) * CDec(TxtDato(4).Text), 2), FORMATO_NUM_1)
         End If
      End If
'///Angel 18/12/2003
'///Se agrega para la eliminacion del dato del centro de costo digitado directamente
   Case MINIMOINDICECUENTA To MINIMOINDICECCOSTO - 1 'Cambiar (añadir índices).
      If TxtDato(Index).Text = "" Then
         TxtDato(Index + CUENTASCONCCOSTO).Text = ""
         lblDatoDeta(Index).Caption = ""
         lblDatoDeta(Index + CUENTASCONCCOSTO).Caption = ""
         TxtDato(Index + CUENTASCONCCOSTO).Enabled = False
         lblDatoDeta(Index + CUENTASCONCCOSTO).Enabled = False
         cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = False
         cmdMas(Index - MINIMOINDICECUENTA + 1).Enabled = True
         If (Not pbNuevo) And cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_CTA Then
            If frmTCprGrd.uorstCOCprDocCta.RecordCount > 0 Then
               ppAbreCtaCCo
               If frmTCprGrd.uorstCOCprDocCta.State = adStateOpen Then
                  frmTCprGrd.uorstCOCprDocCta.MoveFirst
                  Do
                     If frmTCprGrd.uorstCOCprDocCta!CodAux = txtLlave(0).Text And _
                       frmTCprGrd.uorstCOCprDocCta!CodTDc = txtLlave(1).Text And _
                       frmTCprGrd.uorstCOCprDocCta!SerDoc = txtLlave(2).Text And _
                       frmTCprGrd.uorstCOCprDocCta!NroDoc = txtLlave(3).Text And _
                       frmTCprGrd.uorstCOCprDocCta!TpoCnc = Trim(Str(Index - MINIMOINDICECUENTA + 1)) Then
'                        frmTCprGrd.uorstCOCprDocCCo.MoveFirst
'                        Do
'                           If frmTCprGrd.uorstCOCprDocCCo!CodAux = txtLlave(0).Text And _
                             frmTCprGrd.uorstCOCprDocCCo!CodTDc = txtLlave(1).Text And _
                             frmTCprGrd.uorstCOCprDocCCo!SerDoc = txtLlave(2).Text And _
                             frmTCprGrd.uorstCOCprDocCCo!NroDoc = txtLlave(3).Text And _
                             frmTCprGrd.uorstCOCprDocCCo!TpoCnc = Trim(Str(Index - MINIMOINDICECUENTA + 1)) And _
                             frmTCprGrd.uorstCOCprDocCCo!CodCta = frmTCprGrd.uorstCOCprDocCta!CodCta Then
'                              frmTCprGrd.uorstCOCprDocCCo.Delete
'                           End If
'                           frmTCprGrd.uorstCOCprDocCCo.MoveNext
'                        Loop Until frmTCprGrd.uorstCOCprDocCCo.EOF
'                        frmTCprGrd.uorstCOCprDocCCo.Requery
                        frmTCprGrd.uorstCOCprDocCta.Delete
                     End If
                     frmTCprGrd.uorstCOCprDocCta.MoveNext
                  Loop Until frmTCprGrd.uorstCOCprDocCta.EOF
                  frmTCprGrd.uorstCOCprDocCta.Requery
               End If
            End If
         End If
         cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI
      ElseIf cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI Then
            cmdMas(Index - MINIMOINDICECUENTA + 1).Enabled = INDMASCTA_INI
''/// Angel 22/12/2003
''/// Cambio True por False a las 2 siguientes lineas
'            txtDato(Index + CUENTASCONCCOSTO).Enabled = False
'            cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = False
''///
      End If
   End Select
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

  'Completa con ceros a la izquierda.
   Select Case Index
   Case 37, MINIMOINDICECCOSTO To CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      If Len(Trim(TxtDato(Index).Text)) <> 0 And Len(Trim(TxtDato(Index).Text)) <> TxtDato(Index).MaxLength Then
         TxtDato(Index) = gfCeros(TxtDato(Index).Text, TxtDato(Index).MaxLength, 0, "0")
      End If
   End Select

  'Asigna 0 a campos numéricos si están vacíos.
   Select Case Index
   Case 4, MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      If Not IsNumeric(TxtDato(Index).Text) Then
         TxtDato(Index).Text = 0
      End If
   End Select

  'Da formato.
   Select Case Index
   Case 4
      TxtDato(Index).Text = Format(TxtDato(Index).Text, FORMATO_NUM_2)
   Case MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      TxtDato(Index).Text = Format(TxtDato(Index).Text, FORMATO_NUM_1)
   End Select

  'Busca el dato en su tabla principal.
   Select Case Index
   Case 0, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYUDAT, Index)
      If Cancel Then Exit Sub
      If Index >= MINIMOINDICECUENTA And Index < MINIMOINDICECCOSTO Then
'         If frmTCprGrd.uorstCOCta.RecordCount > 0 And txtDato(Index + CUENTASCONCCOSTO).Text = "" Then
         If frmTCprGrd.uorstCOCta.RecordCount > 0 Then
            If Not frmTCprGrd.uorstCOCta.EOF Then
                If frmTCprGrd.uorstCOCta!IndCCo = INDCCO_ACT Then
                   TxtDato(Index + CUENTASCONCCOSTO).Enabled = True
                   cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = True
    '///Angel 22/12/2003
                Else
                   TxtDato(Index + CUENTASCONCCOSTO).Text = ""
                   lblDatoDeta(Index + CUENTASCONCCOSTO).Caption = ""
                   TxtDato(Index + CUENTASCONCCOSTO).Enabled = False
                   cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = False
    '///
                End If
            End If
         End If
      End If
'      If frmTCprGrd.ubGrabaMas = 0 Then
'         frmTCprGrd.ubGrabaMas = 1
'         With frmTCprGrd
'            If pbNuevo Then
'               .uorstMain.AddNew
'            End If
'            upDatosDesconectados 0
'            .uorstMain.Update
'         End With
'      End If
   End Select
      
   Exit Sub
Err:
   gpErrores
End Sub

Private Sub ppAyuBus(tsTipo As String, tnIndex As Integer)
   If tsTipo = AYULLA Then
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Aux_Det "IndPrv=1", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case 1                           'Cambiar (añadir índices).
         modAyuBus.TDc_Cod "", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
         txtLlave(tnIndex).Text = frmOAyuBus.uvDato1
         lblLlaveDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      End Select
   Else
      Select Case tnIndex
      Case 0                           'Cambiar (añadir índices).
         modAyuBus.Dro_Cod "Length(CodDro)=4", TxtDato(tnIndex).Text, 0, 0, Me.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + TxtDato(tnIndex).Left
         TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
         modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", TxtDato(tnIndex).Text, 0, 0, Me.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + TxtDato(tnIndex).Left
         TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
         lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
      Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
         modAyuBus.CCo_Cod "Length(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "' ", TxtDato(tnIndex).Text, 0, 0, Me.Top + TxtDato(tnIndex).Top + TxtDato(tnIndex).Height, Me.Left + TxtDato(tnIndex).Left
         TxtDato(tnIndex).Text = frmOAyuBus.uvDato1
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
         With frmTCprGrd.uorstTGAux
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodAux='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !RazAux
            End If
         End With
      Case 1
         If txtLlave(tnIndex).Text = "" Then
            lblLlaveDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTCprGrd.uorstTGTDc
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
         If TxtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTCprGrd.uorstCODro
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodDro='" & TxtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & !DetDro
            End If
         End With
      Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1
         If TxtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTCprGrd.uorstCOCta
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodCta='" & TxtDato(tnIndex).Text & "'"
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
         If TxtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         With frmTCprGrd.uorstCOCCo
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodCCo='" & TxtDato(tnIndex).Text & "'"
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
      End Select
   End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

 '[Propio del formulario.
   Dim dnContador As Byte
 ']

   With frmTCprGrd.uorstMain           'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !CodAux = txtLlave(0).Text
            !CodTDc = txtLlave(1).Text
            !SerDoc = txtLlave(2).Text
            !NroDoc = txtLlave(3).Text
            !MesPvs = gsMesAct
            !PctIGV = CDec(gnPctIGV)
            !PctISC = CDec(gnPctISC)
         End If

        'Datos.
         !TpoMon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
         !IndPreGen = IIf(chkIndPreGen.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
         !IndCDt = IIf(chkIndCDt.Value = vbChecked, INDCDT_ACT, INDCDT_INA)
'         !CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
         !FehOpe = dtpDato(3).Value
         !FeEDoc = dtpDato(0).Value
         !FeVDoc = dtpDato(1).Value
         !FeRDoc = dtpDato(2).Value
         !FehCDt = dtpDato(4).Value
'         !Tf1Cta = mskDato(0).Text
'         !CodMon = optTpoMon(1).Value
         !CodDro = IIf(TxtDato(0).Text = "", Null, TxtDato(0).Text)
         !NroCpb = TxtDato(1).Text
         !RefDoc = TxtDato(2).Text
         !GloDoc = TxtDato(3).Text
         !ImpTCb = TxtDato(4).Text
         !ImpOGr_MN = TxtDato(5).Text
         !ImpOGN_MN = TxtDato(6).Text
         !ImpONG_MN = TxtDato(7).Text
         !ImpExo_MN = TxtDato(8).Text
         !ImpIGV_MN = TxtDato(9).Text
         !ImpISC_MN = TxtDato(10).Text
         !ImpOIm_MN = TxtDato(11).Text
         !ImpTot_MN = TxtDato(12).Text
         !ImpOGr_ME = TxtDato(13).Text
         !ImpOGN_ME = TxtDato(14).Text
         !ImpONG_ME = TxtDato(15).Text
         !ImpExo_ME = TxtDato(16).Text
         !ImpIGV_ME = TxtDato(17).Text
         !ImpISC_ME = TxtDato(18).Text
         !ImpOIm_ME = TxtDato(19).Text
         !ImpTot_ME = TxtDato(20).Text
         !NroCDt = TxtDato(37).Text
         !ImpIGV_OGr_MN = TxtDato(38).Text
         !ImpIGV_OGN_MN = TxtDato(39).Text
         !ImpIGV_ONG_MN = TxtDato(40).Text
         !ImpIGV_OGr_ME = TxtDato(41).Text
         !ImpIGV_OGN_ME = TxtDato(42).Text
         !ImpIGV_ONG_ME = TxtDato(43).Text

       '[Actualización de datos por Cuentas y Centros de Costo registrados directamente.
         ppAbreCtaCCo
         For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
            If cmdMas(dnContador).Tag = INDMASCTA_CTA And Val(TxtDato(dnContador + DIFERENCIAMASCUENTA).Text) <> 1 Then
               With frmTCprGrd.uorstCOCprDocCta
                  .MoveFirst
                  .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & dnContador & "'"
                  If Not .EOF Then
                     .Delete
                     .Update
                     .Requery
                     frmTCprGrd.uorstCOCprDocCCo.Requery
                     Call upActualizaMas(dnContador, INDMASCTA_INI)
                  End If
               End With
            End If
            
            If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(TxtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
               cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(TxtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Then
               With frmTCprGrd.uorstCOCprDocCta
                  If .RecordCount <> 0 Then .MoveFirst
                  .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & dnContador & "'"
                  If .EOF Then
                     .AddNew
                     !CodAux = txtLlave(0).Text
                     !CodTDc = txtLlave(1).Text
                     !SerDoc = txtLlave(2).Text
                     !NroDoc = txtLlave(3).Text
                     !TpoCnc = dnContador
                     !UsrCre = gsAbvUsr
                     !FyHCre = Now
                  Else
                     !UsrMdf = gsAbvUsr
                     !FyHMdf = Now
                  End If
                  '[ 20/01/2004 Miguel Angel Capturo la cuenta anterior si es modificacion
                  TxtDato(dnContador + DIFERENCIAMASCUENTA).Tag = IIf(pbNuevo, TxtDato(dnContador + DIFERENCIAMASCUENTA).Text, !CodCta)
                  ']
                  !CodCta = TxtDato(dnContador + DIFERENCIAMASCUENTA).Text
                  !ImpCta_MN = TxtDato(dnContador + DIFERENCIAMASIMPORTE).Text
                  !ImpCta_ME = TxtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text
                  .Update
               End With
               If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(TxtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Or _
                  cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(TxtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Then
                  With frmTCprGrd.uorstCOCprDocCCo
                     If .RecordCount <> 0 Then .MoveFirst
                     .Find "cLlave1='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & dnContador & TxtDato(dnContador + DIFERENCIAMASCUENTA).Text & "'"
                     If .EOF Then
                        .AddNew
                        !CodAux = txtLlave(0).Text
                        !CodTDc = txtLlave(1).Text
                        !SerDoc = txtLlave(2).Text
                        !NroDoc = txtLlave(3).Text
                        !TpoCnc = dnContador
                        !CodCta = TxtDato(dnContador + DIFERENCIAMASCUENTA).Text
                        !UsrCre = gsAbvUsr
                        !FyHCre = Now
                     Else
                        !UsrMdf = gsAbvUsr
                        !FyHMdf = Now
                     End If
                     !CodCCo = TxtDato(dnContador + DIFERENCIAMASCCOSTO).Text
                     !ImpCCo_MN = TxtDato(dnContador + DIFERENCIAMASIMPORTE).Text
                     !ImpCCo_ME = TxtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text
                     .Update
                  End With
               End If
               Call upActualizaMas(dnContador, INDMASCTA_CTA)
            End If
         Next
       ']
      Else
        'Llaves.
         txtLlave(0).Text = !CodAux
         txtLlave(1).Text = !CodTDc
         txtLlave(2).Text = !SerDoc
         txtLlave(3).Text = !NroDoc

        'Datos.
         cboTpoMon.ListIndex = IIf(!TpoMon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
         chkIndCDt.Value = IIf(!IndCDt = INDCDT_ACT, vbChecked, vbUnchecked)
         chkIndPreGen.Value = IIf(!IndPreGen = INDPREGEN_ACT, vbChecked, vbUnchecked)
'         dcoSocio.BoundText = IIf(IsNull(!CodSoc), "", !CodSoc)
         dtpDato(3).Value = !FehOpe
         dtpDato(0).Value = !FeEDoc
         dtpDato(1).Value = !FeVDoc
         dtpDato(2).Value = !FeRDoc
         dtpDato(4).Value = !FehCDt
'         optTpoMon(1).Value = uorstMain!CodMon
'         mskDato(0).Text = IIf(IsNull(.uorstMain!Tf1Cta), "", .uorstMain!Tf1Cta)
         TxtDato(0).Text = IIf(IsNull(!CodDro), "", !CodDro)
         TxtDato(1).Text = IIf(IsNull(!NroCpb), "", !NroCpb)
         TxtDato(2).Text = IIf(IsNull(!RefDoc), "", !RefDoc)
         TxtDato(3).Text = IIf(IsNull(!GloDoc), "", !GloDoc)
         TxtDato(4).Text = Format(!ImpTCb, FORMATO_NUM_2)
         TxtDato(5).Text = Format(!ImpOGr_MN, FORMATO_NUM_1)
         TxtDato(6).Text = Format(!ImpOGN_MN, FORMATO_NUM_1)
         TxtDato(7).Text = Format(!ImpONG_MN, FORMATO_NUM_1)
         TxtDato(8).Text = Format(!ImpExo_MN, FORMATO_NUM_1)
         TxtDato(9).Text = Format(!ImpIGV_MN, FORMATO_NUM_1)
         TxtDato(10).Text = Format(!ImpISC_MN, FORMATO_NUM_1)
         TxtDato(11).Text = Format(!ImpOIm_MN, FORMATO_NUM_1)
         TxtDato(12).Text = Format(!ImpTot_MN, FORMATO_NUM_1)
         TxtDato(13).Text = Format(!ImpOGr_ME, FORMATO_NUM_1)
         TxtDato(14).Text = Format(!ImpOGN_ME, FORMATO_NUM_1)
         TxtDato(15).Text = Format(!ImpONG_ME, FORMATO_NUM_1)
         TxtDato(16).Text = Format(!ImpExo_ME, FORMATO_NUM_1)
         TxtDato(17).Text = Format(!ImpIGV_ME, FORMATO_NUM_1)
         TxtDato(18).Text = Format(!ImpISC_ME, FORMATO_NUM_1)
         TxtDato(19).Text = Format(!ImpOIm_ME, FORMATO_NUM_1)
         TxtDato(20).Text = Format(!ImpTot_ME, FORMATO_NUM_1)
         For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
            TxtDato(dnContador).Text = ""
         Next
         TxtDato(37).Text = IIf(IsNull(!NroCDt), "", !NroCDt)
         TxtDato(38).Text = Format(!ImpIGV_OGr_MN, FORMATO_NUM_1)
         TxtDato(39).Text = Format(!ImpIGV_OGN_MN, FORMATO_NUM_1)
         TxtDato(40).Text = Format(!ImpIGV_ONG_MN, FORMATO_NUM_1)
         TxtDato(41).Text = Format(!ImpIGV_OGr_ME, FORMATO_NUM_1)
         TxtDato(42).Text = Format(!ImpIGV_OGN_ME, FORMATO_NUM_1)
         TxtDato(43).Text = Format(!ImpIGV_ONG_ME, FORMATO_NUM_1)
      
       '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
         ppAyuDet AYULLA, 0
         ppAyuDet AYULLA, 1
      '   ppAyuDet AYULLA, 1
         ppAyuDet AYUDAT, 0
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
         ppAyuDet AYUDAT, 34
         ppAyuDet AYUDAT, 35
         ppAyuDet AYUDAT, 36
       ']
      
       '[Propio del formulario.
         For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
            cmdMas(dnContador).Tag = Choose(dnContador, !IndCta_OGr, !IndCta_OGN, !IndCta_ONG, !IndCta_Exo, !IndCta_IGV, !IndCta_ISC, !IndCta_OIm, !IndCta_Tot)
         Next

         ppAbreCtaCCo
         With frmTCprGrd.uorstCOCprDocCta
            If .RecordCount > 0 Then
               For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
                  If Val(TxtDato(dnContador + DIFERENCIAMASIMPORTE).Text) <> 0 Then
                     .MoveFirst
                     .Find "TpoCnc = " & dnContador
                     If Not .EOF Then
                        TxtDato(dnContador + DIFERENCIAMASCUENTA).Text = !CodCta
                        ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCUENTA
                        With frmTCprGrd.uorstCOCprDocCCo
                           If .RecordCount > 0 Then
                              .MoveFirst
                              .Find "cLlave = " & dnContador & frmTCprGrd.uorstCOCprDocCta!CodCta
                              If Not .EOF Then
                                 TxtDato(dnContador + DIFERENCIAMASCCOSTO).Text = !CodCCo
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
   txtLlave(3).Text = ""

  'Datos.
   cboTpoMon.ListIndex = TPOMON_NAC_IND
   chkIndCDt.Value = vbUnchecked
   chkIndPreGen.Value = vbUnchecked
   dtpDato(3).Value = Date
   dtpDato(0).Value = Date
   dtpDato(1).Value = Date
   dtpDato(2).Value = Date
   dtpDato(4).Value = Date
'   optTpoMon(1).Value = True
   For dnContador = 0 To 3
      TxtDato(dnContador).Text = ""
   Next
   TxtDato(4).Text = Format(0, FORMATO_NUM_2)
   For dnContador = MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
      TxtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
   Next
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      TxtDato(dnContador).Text = ""
   Next
   TxtDato(37).Text = ""
   ' Importes de distribucion de igv
   For dnContador = 38 To 43
      TxtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
   Next
   
 '[Propio del formulario.
   For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      cmdMas(dnContador).Tag = INDMASCTA_INI
   Next
 ']

  'Ayudas.
   lblLlaveDeta(0).Caption = ""
   lblLlaveDeta(1).Caption = ""
   lblDatoDeta(0).Caption = ""
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      lblDatoDeta(dnContador).Caption = ""
   Next
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Byte
'///Angel 17/12/2003
'/// Agregado para habilitar textos despues de grabar uno nuevo.
  'Llaves
   txtLlave(0).Enabled = pbNuevo
   txtLlave(1).Enabled = pbNuevo
   txtLlave(2).Enabled = pbNuevo
   txtLlave(3).Enabled = pbNuevo
'///

  'Datos.
   cboTpoMon.Enabled = tbHabilitar
   chkCalcularIGV.Enabled = tbHabilitar
   chkCalcularISC.Enabled = tbHabilitar
   chkDesactivar.Enabled = tbHabilitar
   chkIndCDt.Enabled = tbHabilitar
   chkIndPreGen.Enabled = tbHabilitar
   chkMonedaActiva.Enabled = tbHabilitar
   dtpDato(3).Enabled = tbHabilitar
   dtpDato(0).Enabled = tbHabilitar
   dtpDato(1).Enabled = tbHabilitar
   dtpDato(2).Enabled = tbHabilitar
   dtpDato(4).Enabled = (tbHabilitar And chkIndCDt.Value)
   With TxtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   TxtDato(37).Enabled = (tbHabilitar And chkIndCDt.Value)
   ' Inhabilito el total de IGV
   TxtDato(9).Enabled = False
   TxtDato(17).Enabled = False
   cmdMasIGV.Enabled = tbHabilitar
   
   For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      Call upHabilitaCuenta(False, dnContador)
   Next

  'Ayudas.
   cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar

'[Propio del formulario
   TxtDato(1).Enabled = False 'Deshabilitación del Comprobante.
   
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
'   frmTCprGrd.uorstCOCpbDet.Requery
'   DatosGrid
' ']
'End Sub

'Private Sub cmdBorrar_Click()
'   If Not frmTCprGrd.uorstCOCpbCab.EOF Then
'      frmTCprGrd.uorstCOCpbCab.Delete
'      uorstMain!IndGen = False
'      frmTCprGrd.uorstCOCpbCab.Update
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
'///Angel 18/12/2003
'///Se agrego condicion e instruccion ELSE, solo se dejo instruccion else
      If cmdMas(dnContador).Tag = INDMASCTA_CTA Then
         cmdMas(dnContador).Enabled = cmdMas(dnContador).Enabled = True
      Else
         cmdMas(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
      End If
   Next
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
'///Angel 18/12/2003
'///Se agrego condicion e instruccion ELSE, solo se dejo instruccion else
      If (dnContador >= MINIMOINDICECUENTA And dnContador < MINIMOINDICECCOSTO) Then
         If cmdMas(dnContador - MINIMOINDICECUENTA + 1).Enabled Then
            TxtDato(dnContador).Enabled = False
            lblDatoDeta(dnContador).Enabled = False
            cmdDatoAyud(dnContador).Enabled = False
         Else
            TxtDato(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            lblDatoDeta(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            cmdDatoAyud(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
         End If
      ElseIf dnContador > MINIMOINDICECCOSTO Then
         If cmdMas(dnContador - MINIMOINDICECCOSTO + 1).Enabled Then
            TxtDato(dnContador).Enabled = False
            lblDatoDeta(dnContador).Enabled = False
            cmdDatoAyud(dnContador).Enabled = False
         Else
            TxtDato(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            lblDatoDeta(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
            cmdDatoAyud(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
         End If
      End If
   Next
End Sub

Private Sub chkMonedaActiva_Click()
   unVerMonNac = IIf(chkMonedaActiva.Value, cboTpoMon.ListIndex, IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_IND, TPOMON_NAC_IND))
   ppCambioTpoMon
End Sub

Private Sub cmdAuxiliar_Click()
    frmMAuxGrd.Show vbModal
    frmTCprGrd.uorstTGAux.Requery
End Sub

Private Sub cmdMas_Click(Index As Integer) 'Cambiar Formulario de Grid.
'   If frmTCprGrd.ubGrabaMas = 0 Then
'      frmTCprGrd.ubGrabaMas = 1
'      With frmTCprGrd
'         If pbNuevo Then
'            .uorstMain.AddNew
'         End If
'         upDatosDesconectados 0
'         .uorstMain.Update
'      End With
'   End If

   frmTCprMasGrd.unIndice = Index
   frmTCprMasGrd.Show vbModal
   ppAbreCtaCCo
   ppAyuDet AYUDAT, Index + DIFERENCIAMASCUENTA
   If Index <= CUENTASCONCCOSTO Then
      ppAyuDet AYUDAT, Index + DIFERENCIAMASCCOSTO
   End If
End Sub

Private Sub dtpDato_Validate(Index As Integer, Cancel As Boolean)
   Dim dnContador As Byte
   If Index = 0 Then
'///Angel 22/12/2003
      If Month(dtpDato(0).Value) > Val(gsMesAct) And Year(dtpDato(0).Value) >= Val(gsAnoAct) Then
         MsgBox "La fecha No Corresponde al Periodo de Operacion", vbCritical
         dtpDato(Index).SetFocus
         Cancel = True
         Exit Sub
      End If
'///
      If dtpDato(0).Tag <> dtpDato(0).Value Then
         dtpDato(0).Tag = dtpDato(0).Value
         With frmTCprGrd.uorstTGTCb
            If .RecordCount <> 0 Then .MoveLast: .MoveFirst
            .Find "FehTCb = '" & dtpDato(0).Value & "'"
            If .EOF Then
               MsgBox "No se ha ingresado Tipo de Cambio para esta fecha.", vbCritical
               Cancel = True
'///Angel 22/12/2003
               Exit Sub
'///
            Else
'               uorstMain!ImpTCb = IIf(frmTCprGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
''               frmTCprGrd.uorstMain!ImpTCb = !ImpTCb_Vta
               TxtDato(4).Text = Format(!ImpTCb_Vta, FORMATO_NUM_2)
            End If
         End With
      End If
   End If
   If Index = 3 Then
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
   With frmTCprGrd.uorstCOCprDocCCo
      frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!CodAux & "' AND COCprDocCCo.CodTDc='" & frmTCprGrd.uorstMain!CodTDc & "' AND COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!SerDoc & "' AND COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!NroDoc & "' "
      If .State = adStateOpen Then .Close
      .Source = frmTCprGrd.usConnStrgSele_COCprDocCCo & frmTCprGrd.usConnStrgWher_COCprDocCCo & frmTCprGrd.usConnStrgOrde_COCprDocCCo
      .Open
      .Properties("Unique Table").Value = "COCprDocCCo"
   End With
   With frmTCprGrd.uorstCOCprDocCta
      frmTCprGrd.usConnStrgWher_COCprDocCta = " WHERE COCprDocCta.CodAux='" & frmTCprGrd.uorstMain!CodAux & "' AND COCprDocCta.CodTDc='" & frmTCprGrd.uorstMain!CodTDc & "' AND COCprDocCta.SerDoc='" & frmTCprGrd.uorstMain!SerDoc & "' AND COCprDocCta.NroDoc='" & frmTCprGrd.uorstMain!NroDoc & "' "
      If .State = adStateOpen Then .Close
      .Source = frmTCprGrd.usConnStrgSele_COCprDocCta & frmTCprGrd.usConnStrgWher_COCprDocCta & frmTCprGrd.usConnStrgOrde_COCprDocCta
      .Open
      .Properties("Unique Table").Value = "COCprDocCta"
   End With
End Sub

Private Sub ppCambioTpoMon()
   Dim dnContador As Integer
   
   chkMonedaActiva.Caption = IIf(unVerMonNac = TPOMON_NAC_IND, TPOMON_NAC_TXT_2, TPOMON_EXT_TXT_2)
   For dnContador = MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1
      TxtDato(dnContador).Visible = (unVerMonNac = TPOMON_NAC_IND)
   Next
   For dnContador = MINIMOINDICEIMPORTEME To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
      TxtDato(dnContador).Visible = Not (unVerMonNac = TPOMON_NAC_IND)
   Next
End Sub

Private Sub ppGenera()
   Dim dnContador As Integer
   Dim dnNumeroItem As Integer
   Dim dbProcesaCuenta As Boolean

   If TxtDato(1).Text <> "" Then
      ppDatosWhere

      With frmTCprGrd.uorstCOCpbCab
        'Si existe, elimina Comprobante existente.
         If .RecordCount > 0 Then
            .MoveFirst
            .Find "cLlave='" & TxtDato(0).Text & TxtDato(1).Text & "'"
            If Not .EOF Then .Delete
         End If
      End With
   End If

   With frmTCprGrd.uorstCOCpbCab
     'Si no está marcado para generar, marca el documento como no generado.
      If chkIndPreGen.Value = vbUnchecked Then
         frmTCprGrd.uorstMain!IndGen = False
         frmTCprGrd.uorstMain.Update
         Exit Sub
      End If

     'Captura del Siguiente Número.
      If TxtDato(1).Text = "" Then
         With frmTCprGrd.uorstCODro
            If .RecordCount <> 0 Then .MoveFirst
            .Find "CodDro = '" & TxtDato(0).Text & "'"
            If IsNull(.Fields(2).Value) Then .Fields(2).Value = gfCeros("", .Fields(2).DefinedSize, 0, "0")
            TxtDato(1).Text = gfCeros(.Fields(2).Value, .Fields(2).DefinedSize, 1, "0")
            .Fields(2).Value = TxtDato(1).Text
            .Update
         End With
         frmTCprGrd.uorstMain!NroCpb = TxtDato(1).Text
         frmTCprGrd.uorstMain.Update
      End If
      
      ppDatosWhere
   
     'Si no hay cuentas, marca el documento como no generado.
      If frmTCprGrd.uorstCOCprDocCta.RecordCount = 0 Then
         frmTCprGrd.uorstMain!IndGen = False
         frmTCprGrd.uorstMain.Update
         Exit Sub
      End If

     'Crea encabezado de Comprobante.
      .AddNew
      !MesPvs = gsMesAct
      !CodDro = TxtDato(0).Text
      !NroCpb = TxtDato(1).Text
      !FehCpb = dtpDato(3).Value
      !TpoGnr = TPOGNR_CPR
      !IndNCu = INDNCU_FAL
      !GloCpb = TxtDato(3).Text
      !UsrCre = gsAbvUsr
      !FyHCre = Now
'Angel 15/12/2003
      
   End With

   With frmTCprGrd.uorstCOCprDocCta
     'Crea ítemes de Comprobante.
      .MoveFirst
      Do
         dbProcesaCuenta = True

        'Itemes con Centro de Costo.
         If Val(!TpoCnc) <= CUENTASCONCCOSTO Then
            With frmTCprGrd.uorstCOCprDocCCo
               If .RecordCount <> 0 Then
                  .MoveFirst
                  .Find "cLlave = " & Trim(frmTCprGrd.uorstCOCprDocCta!TpoCnc) & frmTCprGrd.uorstCOCprDocCta!CodCta
                  If Not .EOF Then
                     Do
                        dnNumeroItem = dnNumeroItem + 1
                        Call ppGenera1(True, dnNumeroItem)
                        .MoveNext
                        If .EOF Then Exit Do
                        If !cLlave <> Trim(frmTCprGrd.uorstCOCprDocCta!TpoCnc) & frmTCprGrd.uorstCOCprDocCta!CodCta Then Exit Do
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

   frmTCprGrd.uorstMain!IndGen = True
   TxtDato(0).Enabled = False
   TxtDato(1).Enabled = False
   cmdDatoAyud(0).Enabled = False
   lblDatoDeta(0).Enabled = False

   frmTCprGrd.uorstCOCpbCab.Update
   frmTCprGrd.uorstCOCpbDet.UpdateBatch
   frmTCprGrd.uorstMain.Update
End Sub

Private Sub ppGenera1(tbCCosto As Boolean, tnNumeroItem As Integer)
   With frmTCprGrd.uorstCOCpbDet
      .AddNew
      !CodDro = TxtDato(0).Text
      !NroCpb = TxtDato(1).Text
      !NroIte = tnNumeroItem
      !MesPvs = gsMesAct
      !CodCta = frmTCprGrd.uorstCOCprDocCta!CodCta
      !FehOpe = dtpDato(3).Value
      ' Busco en el plan de cuentas
      frmTCprGrd.uorstCOCta.MoveFirst
      frmTCprGrd.uorstCOCta.Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "'"
      If frmTCprGrd.uorstCOCta!IndCCo = INDCCO_ACT Then If tbCCosto Then !CodCCo = frmTCprGrd.uorstCOCprDocCCo!CodCCo
      If frmTCprGrd.uorstCOCta!IndDoc = INDDOC_ACT Then
         !CodAux = txtLlave(0).Text
         !CodTDc = txtLlave(1).Text
         !SerDoc = txtLlave(2).Text
         !NroDoc = txtLlave(3).Text
         !FeEDoc = dtpDato(0).Value
         !FeVDoc = dtpDato(1).Value
         !FeRDoc = dtpDato(2).Value
         !RefDoc = TxtDato(2).Text
      End If
      !GloIte = TxtDato(3).Text
      !TpoCtb = IIf(frmTCprGrd.uorstCOCprDocCta!TpoCnc = TPOCNC_TOT_CPR, IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_HAB, TPOCTB_DEB), IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_HAB, TPOCTB_DEB))
      !TpoMon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      !ImpTCb = CDec(TxtDato(4).Text)
      If tbCCosto Then
'         !ImpMN = frmTCprGrd.uorstCOCprDocCCo!ImpCCo * IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, txtDato(4).Text)
'         !ImpME = frmTCprGrd.uorstCOCprDocCCo!ImpCCo / IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, txtDato(4).Text, 1)
         !ImpMN = frmTCprGrd.uorstCOCprDocCCo!ImpCCo_MN
         !ImpME = frmTCprGrd.uorstCOCprDocCCo!ImpCCo_ME
      Else
'         !ImpMN = frmTCprGrd.uorstCOCprDocCta!ImpCta * IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, txtDato(4).Text)
'         !ImpME = frmTCprGrd.uorstCOCprDocCta!ImpCta / IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, txtDato(4).Text, 1)
         !ImpMN = frmTCprGrd.uorstCOCprDocCta!ImpCta_MN
         !ImpME = frmTCprGrd.uorstCOCprDocCta!ImpCta_ME
      End If
      !TpoGnr = TPOGNR_CPR
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With
End Sub

Public Sub upActualizaMas(pnIndice As Byte, pnValor As Byte)
   frmTCpr.cmdMas(pnIndice).Tag = pnValor 'Necesaria la referencia por ser llamado externamente.
   With frmTCprGrd.uorstMain
      Select Case pnIndice
      Case 1
         !IndCta_OGr = pnValor
      Case 2
         !IndCta_OGN = pnValor
      Case 3
         !IndCta_ONG = pnValor
      Case 4
         !IndCta_Exo = pnValor
      Case 5
         !IndCta_IGV = pnValor
      Case 6
         !IndCta_ISC = pnValor
      Case 7
         !IndCta_OIm = pnValor
      Case 8
         !IndCta_Tot = pnValor
      End Select
   End With
End Sub

Public Sub upHabilitaCuenta(tbHabilita As Boolean, tnIndice As Byte)
   TxtDato(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
   lblDatoDeta(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
   cmdDatoAyud(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
   If tnIndice <= CUENTASCONCCOSTO Then
      Call upHabilitaCCosto(tbHabilita, tnIndice)
   End If
End Sub

Public Sub upHabilitaCCosto(tbHabilita As Boolean, tnIndice As Byte)
   If Not tbHabilita Or TxtDato(tnIndice + DIFERENCIAMASCUENTA).Text = "" Then
      TxtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
      lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
      cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
   Else
      frmTCprGrd.uorstCOCta.MoveFirst
      frmTCprGrd.uorstCOCta.Find "CodCta='" & TxtDato(tnIndice + DIFERENCIAMASCUENTA).Text & "'"
      TxtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTCprGrd.uorstCOCta!IndCCo = INDCCO_ACT)
      lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTCprGrd.uorstCOCta!IndCCo = INDCCO_ACT)
      cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTCprGrd.uorstCOCta!IndCCo = INDCCO_ACT)
   End If
End Sub

Private Sub ppDatosWhere()             'Cambiar.
   With frmTCprGrd
      .uorstCOCpbCab.Requery
   
      .usConnStrgWher_COCpbDet = "WHERE COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='" & TxtDato(0).Text & "' AND COCpbDet.NroCpb='" & TxtDato(1).Text & "' "
      With .uorstCOCpbDet
         .Close
         .Source = frmTCprGrd.usConnStrgSele_COCpbDet & frmTCprGrd.usConnStrgWher_COCpbDet & frmTCprGrd.usConnStrgOrde_COCpbDet
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
'       If (CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0) And _
        (cmdMas(dnContador + 1).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
        cmdMas(dnContador + 1).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
       If ((CDec(TxtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(TxtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) And _
        (cmdMas(dnContador + 1).Tag = INDMASCTA_MAS And Len(Trim(TxtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
          ValidaCtasCCo = Not (TxtDato(MINIMOINDICECUENTA + dnContador).Text = "")
          If Not ValidaCtasCCo Then Exit Function
          If frmTCprGrd.ubGrabaMas = 1 Then
             With frmTCprGrd.uorstCOCprDocCta
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
                         With frmTCprGrd.uorstCOCta
                            .MoveFirst
                            .Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "'"
                            If Not .EOF Then
                               dnIndCCo = frmTCprGrd.uorstCOCta!IndCCo
                            End If
                         End With
                      End If
                      If dnIndCCo = INDCCO_ACT Then
                         With frmTCprGrd.uorstCOCprDocCCo
                            If .State = adStateOpen Then .Close
                            frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!CodAux & "' And COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!SerDoc & "' And COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!NroDoc & "' And COCprDocCCo.TpoCnc='" & Trim(Str(dnContador + 1)) & "' And COCprDocCCo.CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "' "
                            .Source = frmTCprGrd.usConnStrgSele_COCprDocCCo & frmTCprGrd.usConnStrgWher_COCprDocCCo & frmTCprGrd.usConnStrgOrde_COCprDocCCo
                            .Open
                            If .RecordCount = 0 Then
                               MsgBox "Cuenta " & frmTCprGrd.uorstCOCprDocCta!CodCta & " requiere C.Costo", vbInformation
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
'              dnTotalImporteMN = gfRedond(CDec(txtDato(5).Text) + CDec(txtDato(6).Text) + CDec(txtDato(7).Text) + CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text) + CDec(txtDato(11).Text), 2)
'              dnTotalImporteME = gfRedond(CDec(txtDato(12).Text) + CDec(txtDato(13).Text) + CDec(txtDato(14).Text) + CDec(txtDato(15).Text) + CDec(txtDato(16).Text) + CDec(txtDato(17).Text) + CDec(txtDato(18).Text), 2)
             dnTotalImporteMN = gfRedond(CDec(TxtDato(MINIMOINDICEIMPORTEMN + dnContador).Text), 2)
             dnTotalImporteME = gfRedond(CDec(TxtDato(MINIMOINDICEIMPORTEME + dnContador).Text), 2)
             If Not (CDec(dnTotalCuentaMN) = CDec(dnTotalImporteMN)) Then
                ValidaCtasCCo = False
                Exit Function
             End If
             If Not (CDec(dnTotalCuentaME) = CDec(dnTotalImporteME)) Then
                ValidaCtasCCo = False
                Exit Function
             End If
          End If
       ElseIf Len(Trim(TxtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0 Then
          dnIndCCo = 0
          With frmTCprGrd.uorstCOCta
             .MoveFirst
             .Find "CodCta='" & TxtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & "'"
             If Not .EOF Then
                dnIndCCo = frmTCprGrd.uorstCOCta!IndCCo
             End If
          End With
          If dnIndCCo = INDCCO_ACT And TxtDato(dnContador + 1 + DIFERENCIAMASCUENTA + CUENTASCONCCOSTO).Text = "" Then
             MsgBox "Cuenta " & TxtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & " requiere C.Costo", vbInformation
'             MsgBox "Cuenta " & frmTCprGrd.uorstCOCprDocCta!CodCta & " requiere C.Costo", vbInformation
             ValidaCtasCCo = False
             Exit Function
          End If
       ElseIf Len(Trim(TxtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) = 0 And ((CDec(TxtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(TxtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) Then
         ValidaCtasCCo = False
       End If
    Next dnContador
    ' Valido que los detalles sean iguales los importes
''    With frmTCprGrd.uorstCOCprDocCta
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



