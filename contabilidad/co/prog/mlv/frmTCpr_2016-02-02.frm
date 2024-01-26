VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTCpr 
   Caption         =   "[Título]"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   9480
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
      Index           =   60
      Left            =   1185
      TabIndex        =   43
      Top             =   3570
      Width           =   2250
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   60
      Left            =   8040
      Picture         =   "frmTCpr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   3570
      Width           =   255
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "&Proveedor"
      Height          =   375
      Left            =   8490
      TabIndex        =   154
      Top             =   3015
      Width           =   975
   End
   Begin VB.ComboBox cboCategoria 
      Height          =   315
      ItemData        =   "frmTCpr.frx":01AA
      Left            =   7005
      List            =   "frmTCpr.frx":01AC
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   2520
      Width           =   1485
   End
   Begin VB.ComboBox cboImpuesto 
      Height          =   315
      ItemData        =   "frmTCpr.frx":01AE
      Left            =   7005
      List            =   "frmTCpr.frx":01B0
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   1965
      Width           =   1485
   End
   Begin VB.Frame fraAsiento 
      Height          =   540
      Left            =   0
      TabIndex        =   38
      Top             =   2910
      Width           =   6960
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   51
         Left            =   1560
         TabIndex        =   40
         Top             =   165
         Width           =   560
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   51
         Left            =   6600
         Picture         =   "frmTCpr.frx":01B2
         Style           =   1  'Graphical
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   165
         Width           =   255
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Asiento Tipo :"
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
         Index           =   29
         Left            =   105
         TabIndex        =   39
         Top             =   180
         Width           =   990
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
         Index           =   51
         Left            =   2100
         TabIndex        =   41
         Top             =   165
         Width           =   4500
      End
   End
   Begin VB.TextBox txtDato 
      Alignment       =   1  'Right Justify
      Height          =   280
      Index           =   4
      Left            =   2835
      TabIndex        =   25
      Top             =   1545
      Width           =   735
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   49
      Left            =   3840
      TabIndex        =   21
      Top             =   1545
      Width           =   4635
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   3
      Left            =   3840
      TabIndex        =   20
      Top             =   1215
      Width           =   4635
   End
   Begin VB.TextBox txtDato 
      Height          =   280
      Index           =   2
      Left            =   1020
      TabIndex        =   18
      Top             =   1215
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
      Height          =   280
      Index           =   3
      Left            =   7320
      TabIndex        =   8
      Top             =   420
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
      Height          =   315
      Index           =   2
      Left            =   6840
      TabIndex        =   7
      Top             =   420
      Width           =   525
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
      Height          =   280
      Index           =   1
      Left            =   900
      TabIndex        =   4
      Top             =   420
      Width           =   315
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
      Height          =   315
      Index           =   0
      Left            =   900
      TabIndex        =   1
      Top             =   90
      Width           =   1275
   End
   Begin VB.Frame fraPedido 
      Height          =   540
      Left            =   0
      TabIndex        =   26
      Top             =   1830
      Width           =   6945
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   50
         Left            =   6600
         Picture         =   "frmTCpr.frx":035C
         Style           =   1  'Graphical
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   165
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
         Height          =   315
         Index           =   50
         Left            =   1005
         TabIndex        =   28
         Top             =   165
         Width           =   1410
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
         Index           =   50
         Left            =   2415
         TabIndex        =   29
         Top             =   165
         Width           =   4170
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
         Index           =   28
         Left            =   105
         TabIndex        =   27
         Top             =   180
         Width           =   870
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tic&ket"
      Height          =   450
      Left            =   7125
      TabIndex        =   196
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdPedido 
      Caption         =   "P&edido"
      Height          =   375
      Left            =   8490
      TabIndex        =   155
      Top             =   3525
      Width           =   975
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   1
      Left            =   2730
      TabIndex        =   47
      Top             =   3990
      Width           =   5715
      Begin VB.CheckBox chkIndCDt 
         Caption         =   "Detracción"
         ForeColor       =   &H00800000&
         Height          =   200
         Left            =   105
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   218
         Width           =   1080
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
         Index           =   52
         Left            =   1920
         TabIndex        =   50
         Top             =   180
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpDato 
         Height          =   315
         Index           =   4
         Left            =   4320
         TabIndex        =   52
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65667073
         CurrentDate     =   37102
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Form. Nº:"
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
         Left            =   1200
         TabIndex        =   49
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblTexto 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   14
         Left            =   3840
         TabIndex        =   51
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CheckBox chkIndPreGen 
      Caption         =   "Cuentas &Registradas"
      ForeColor       =   &H00C00000&
      Height          =   200
      Left            =   6960
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   4590
      Width           =   1815
   End
   Begin VB.CheckBox chkCalcularIGV 
      Caption         =   "Calcular I.G.&V."
      ForeColor       =   &H00800000&
      Height          =   200
      Left            =   60
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4230
      Width           =   1365
   End
   Begin VB.CheckBox chkCalcularISC 
      Caption         =   "Calcular I.S.&C."
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1365
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   540
      Index           =   0
      Left            =   0
      TabIndex        =   30
      Top             =   2370
      Width           =   6960
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   1
         Left            =   6120
         TabIndex        =   35
         Top             =   165
         Width           =   735
      End
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   0
         Left            =   690
         TabIndex        =   32
         Top             =   165
         Width           =   555
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   0
         Left            =   4755
         Picture         =   "frmTCpr.frx":0506
         Style           =   1  'Graphical
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   165
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
         Index           =   0
         Left            =   1230
         TabIndex        =   33
         Top             =   165
         Width           =   3540
      End
      Begin VB.Label lblTexto 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   12
         Left            =   5085
         TabIndex        =   34
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label lblTexto 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   195
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   8640
      ScaleHeight     =   2610
      ScaleWidth      =   885
      TabIndex        =   191
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
         Picture         =   "frmTCpr.frx":06B0
         Style           =   1  'Graphical
         TabIndex        =   153
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
         Picture         =   "frmTCpr.frx":07FA
         Style           =   1  'Graphical
         TabIndex        =   152
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
         Picture         =   "frmTCpr.frx":08FC
         Style           =   1  'Graphical
         TabIndex        =   151
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
         Picture         =   "frmTCpr.frx":09FE
         Style           =   1  'Graphical
         TabIndex        =   150
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
         Picture         =   "frmTCpr.frx":0B48
         Style           =   1  'Graphical
         TabIndex        =   148
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
         Picture         =   "frmTCpr.frx":0CF2
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   1
      Left            =   5835
      Picture         =   "frmTCpr.frx":0E9C
      Style           =   1  'Graphical
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   420
      Width           =   255
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   7695
      Picture         =   "frmTCpr.frx":1046
      Style           =   1  'Graphical
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   90
      Width           =   255
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTCpr.frx":11F0
      Left            =   1020
      List            =   "frmTCpr.frx":11F2
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1545
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   285
      Index           =   0
      Left            =   5265
      TabIndex        =   14
      Top             =   870
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65667073
      CurrentDate     =   37102
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   285
      Index           =   1
      Left            =   7305
      TabIndex        =   16
      Top             =   870
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65667073
      CurrentDate     =   37102
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   285
      Index           =   2
      Left            =   3180
      TabIndex        =   12
      Top             =   870
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65667073
      CurrentDate     =   37102
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   10
      Top             =   870
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65667073
      CurrentDate     =   37102
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   4080
      Left            =   0
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   4590
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   7197
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   441
      ForeColor       =   8388608
      TabCaption(0)   =   "I&mportes"
      TabPicture(0)   =   "frmTCpr.frx":11F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(22)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTexto(27)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTexto(26)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTexto(15)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTexto(18)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTexto(21)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTexto(20)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblTexto(19)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDatoDeta(27)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDatoDeta(28)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDatoDeta(29)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDatoDeta(30)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDatoDeta(31)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDatoDeta(32)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDatoDeta(33)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDatoDeta(34)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDatoDeta(38)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblDatoDeta(39)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblDatoDeta(40)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblDatoDeta(41)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblDatoDeta(42)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblDatoDeta(43)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblDatoDeta(44)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblDatoDeta(45)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblTexto(16)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblTexto(17)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblDatoDeta(48)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblDatoDeta(47)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblDatoDeta(46)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblDatoDeta(37)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblDatoDeta(36)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblDatoDeta(35)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblTexto(23)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblTexto(24)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblTexto(25)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "chkDesactivar"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmdDatoAyud(27)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtDato(27)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtDato(28)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdDatoAyud(28)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdDatoAyud(29)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtDato(29)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdDatoAyud(30)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtDato(30)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdDatoAyud(31)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtDato(31)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cmdDatoAyud(32)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtDato(32)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "cmdDatoAyud(33)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtDato(33)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cmdDatoAyud(34)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtDato(34)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cmdDatoAyud(38)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtDato(38)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "cmdDatoAyud(39)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtDato(39)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmdDatoAyud(40)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtDato(40)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cmdDatoAyud(41)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtDato(41)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmdDatoAyud(42)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtDato(42)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "cmdDatoAyud(43)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txtDato(43)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cmdMas(1)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "cmdMas(2)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "cmdMas(3)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "cmdMas(4)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "cmdMas(5)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "cmdMas(6)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cmdMas(7)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cmdMas(8)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "chkMonedaActiva"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "cmdDatoAyud(45)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "cmdDatoAyud(44)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txtDato(44)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "txtDato(45)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "txtDato(16)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txtDato(17)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "txtDato(18)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "txtDato(19)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "txtDato(20)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "txtDato(21)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txtDato(22)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "txtDato(23)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "txtDato(5)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "txtDato(6)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "txtDato(7)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txtDato(8)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "txtDato(9)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "txtDato(10)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "txtDato(11)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "txtDato(12)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "txtDato(53)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "txtDato(54)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "txtDato(55)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "txtDato(56)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "cmdMasIGV"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "txtDato(26)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "txtDato(25)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "txtDato(24)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "txtDato(48)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "txtDato(47)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "cmdDatoAyud(47)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "cmdDatoAyud(48)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "cmdMas(11)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "cmdMas(10)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "cmdMas(9)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "txtDato(46)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "cmdDatoAyud(46)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "txtDato(37)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "cmdDatoAyud(37)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "txtDato(36)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "cmdDatoAyud(36)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "txtDato(35)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "cmdDatoAyud(35)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "txtDato(13)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "txtDato(14)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "txtDato(15)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "txtDato(57)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "txtDato(58)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).ControlCount=   121
      TabCaption(1)   =   "C&uentas"
      TabPicture(1)   =   "frmTCpr.frx":1210
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgrDetalle"
      Tab(1).Control(1)=   "Picture2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "pdb-dua-nc-nd-snd"
      TabPicture(2)   =   "frmTCpr.frx":122C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraReferencia"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraDetracion"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraExterior"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chkIndreten"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
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
         Index           =   58
         Left            =   0
         TabIndex        =   221
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
         Index           =   57
         Left            =   0
         TabIndex        =   220
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
         Index           =   15
         Left            =   1320
         TabIndex        =   141
         Top             =   3645
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
         TabIndex        =   133
         Top             =   3330
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
         TabIndex        =   125
         Top             =   3015
         Width           =   1695
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   35
         Left            =   6540
         Picture         =   "frmTCpr.frx":1248
         Style           =   1  'Graphical
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   3045
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
         Index           =   35
         Left            =   3300
         TabIndex        =   128
         Top             =   3015
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   36
         Left            =   6540
         Picture         =   "frmTCpr.frx":13F2
         Style           =   1  'Graphical
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   3360
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
         Index           =   36
         Left            =   3300
         TabIndex        =   136
         Top             =   3330
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   37
         Left            =   6540
         Picture         =   "frmTCpr.frx":159C
         Style           =   1  'Graphical
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   3675
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
         Index           =   37
         Left            =   3300
         TabIndex        =   144
         Top             =   3645
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   46
         Left            =   9060
         Picture         =   "frmTCpr.frx":1746
         Style           =   1  'Graphical
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   3045
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
         Index           =   46
         Left            =   6840
         TabIndex        =   130
         Top             =   3015
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
         Height          =   285
         Index           =   9
         Left            =   3000
         Picture         =   "frmTCpr.frx":18F0
         TabIndex        =   127
         Top             =   3045
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
         Index           =   10
         Left            =   3000
         Picture         =   "frmTCpr.frx":19F2
         TabIndex        =   135
         Top             =   3360
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
         Index           =   11
         Left            =   3000
         Picture         =   "frmTCpr.frx":1AF4
         TabIndex        =   143
         Top             =   3675
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   48
         Left            =   9060
         Picture         =   "frmTCpr.frx":1BF6
         Style           =   1  'Graphical
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   3675
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   47
         Left            =   9060
         Picture         =   "frmTCpr.frx":1DA0
         Style           =   1  'Graphical
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   3360
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
         Index           =   47
         Left            =   6840
         TabIndex        =   138
         Top             =   3330
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
         Index           =   48
         Left            =   6840
         TabIndex        =   146
         Top             =   3645
         Width           =   675
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
         Index           =   24
         Left            =   1320
         TabIndex        =   126
         Top             =   3015
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
         Index           =   25
         Left            =   1320
         TabIndex        =   134
         Top             =   3330
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
         Index           =   26
         Left            =   1320
         TabIndex        =   142
         Top             =   3645
         Width           =   1695
      End
      Begin VB.CheckBox chkIndreten 
         Caption         =   "Afecto a Retencion"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   -74910
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1860
      End
      Begin VB.Frame fraExterior 
         Caption         =   " Compra Externa "
         ForeColor       =   &H00C00000&
         Height          =   690
         Left            =   -74910
         TabIndex        =   197
         Top             =   495
         Width           =   3195
         Begin VB.CheckBox chkIndCprext 
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   1380
            TabIndex        =   198
            TabStop         =   0   'False
            Top             =   15
            Width           =   180
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   1
            Left            =   1800
            TabIndex        =   201
            Top             =   270
            Width           =   500
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   0
            Left            =   1365
            TabIndex        =   200
            Top             =   270
            Width           =   440
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   2
            Left            =   2280
            TabIndex        =   202
            Top             =   270
            Width           =   800
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Numero DUA :"
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
            Index           =   30
            Left            =   60
            TabIndex        =   199
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.Frame fraDetracion 
         Caption         =   " Detracion "
         ForeColor       =   &H00C00000&
         Height          =   690
         Left            =   -71640
         TabIndex        =   204
         Top             =   495
         Width           =   6030
         Begin VB.ComboBox cboDetraccion 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   270
            Width           =   5085
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tasa :"
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
            Index           =   31
            Left            =   60
            TabIndex        =   205
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame fraReferencia 
         Caption         =   " Documento de Referencia "
         ForeColor       =   &H00C00000&
         Height          =   1680
         Left            =   -71640
         TabIndex        =   207
         Top             =   1200
         Width           =   6030
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   4
            Left            =   1440
            TabIndex        =   211
            Top             =   615
            Width           =   1185
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   3
            Left            =   945
            TabIndex        =   210
            Top             =   615
            Width           =   500
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
            Index           =   59
            Left            =   945
            TabIndex        =   55
            Top             =   285
            Width           =   315
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   59
            Left            =   5625
            Picture         =   "frmTCpr.frx":1F4A
            Style           =   1  'Graphical
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   285
            Width           =   255
         End
         Begin VB.TextBox txtDetalle 
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
            Left            =   945
            TabIndex        =   218
            Top             =   1275
            Width           =   1695
         End
         Begin VB.TextBox txtDetalle 
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
            Left            =   945
            TabIndex        =   215
            Top             =   930
            Width           =   1695
         End
         Begin VB.TextBox txtDetalle 
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
            Left            =   945
            TabIndex        =   219
            Top             =   1275
            Width           =   1695
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   7
            Left            =   945
            TabIndex        =   216
            Top             =   930
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtpDetalle 
            Height          =   315
            Index           =   0
            Left            =   3555
            TabIndex        =   213
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65667073
            CurrentDate     =   37102
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   32
            Left            =   60
            TabIndex        =   208
            Top             =   300
            Width           =   765
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. :"
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
            Index           =   33
            Left            =   60
            TabIndex        =   209
            Top             =   645
            Width           =   645
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
            Index           =   59
            Left            =   1245
            TabIndex        =   54
            Top             =   285
            Width           =   4410
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "F. Emisión :"
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
            Index           =   34
            Left            =   2685
            TabIndex        =   212
            Top             =   645
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "I.G.V. :"
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
            Index           =   36
            Left            =   60
            TabIndex        =   217
            Top             =   1305
            Width           =   495
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Base Impo. :"
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
            Index           =   35
            Left            =   60
            TabIndex        =   214
            Top             =   960
            Width           =   885
         End
      End
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
         Picture         =   "frmTCpr.frx":20F4
         TabIndex        =   92
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
         Index           =   56
         Left            =   0
         TabIndex        =   195
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
         Index           =   55
         Left            =   0
         TabIndex        =   194
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
         Index           =   54
         Left            =   0
         TabIndex        =   193
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
         Index           =   53
         Left            =   0
         TabIndex        =   192
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
         TabIndex        =   117
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
         TabIndex        =   109
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
         TabIndex        =   101
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
         TabIndex        =   93
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
         TabIndex        =   84
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
         TabIndex        =   76
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
         TabIndex        =   68
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
         TabIndex        =   60
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
         Index           =   23
         Left            =   1320
         TabIndex        =   118
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
         Index           =   22
         Left            =   1320
         TabIndex        =   110
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
         Index           =   21
         Left            =   1320
         TabIndex        =   102
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
         Index           =   20
         Left            =   1320
         TabIndex        =   94
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
         Index           =   19
         Left            =   1320
         TabIndex        =   85
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
         Index           =   18
         Left            =   1320
         TabIndex        =   77
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
         Index           =   17
         Left            =   1320
         TabIndex        =   69
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
         Index           =   16
         Left            =   1320
         TabIndex        =   61
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
         Index           =   45
         Left            =   6840
         TabIndex        =   122
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
         Index           =   44
         Left            =   6840
         TabIndex        =   114
         Top             =   2400
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   44
         Left            =   9060
         Picture         =   "frmTCpr.frx":21F6
         Style           =   1  'Graphical
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   2430
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   45
         Left            =   9060
         Picture         =   "frmTCpr.frx":23A0
         Style           =   1  'Graphical
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   2730
         Width           =   255
      End
      Begin VB.CheckBox chkMonedaActiva 
         Caption         =   "M&oneda activa"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1320
         TabIndex        =   59
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
         Picture         =   "frmTCpr.frx":254A
         TabIndex        =   119
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
         Picture         =   "frmTCpr.frx":264C
         TabIndex        =   111
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
         Picture         =   "frmTCpr.frx":274E
         TabIndex        =   103
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
         Picture         =   "frmTCpr.frx":2850
         TabIndex        =   95
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
         Picture         =   "frmTCpr.frx":2952
         TabIndex        =   86
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
         Picture         =   "frmTCpr.frx":2A54
         TabIndex        =   78
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
         Picture         =   "frmTCpr.frx":2B56
         TabIndex        =   70
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
         Picture         =   "frmTCpr.frx":2C58
         TabIndex        =   62
         Top             =   625
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   -70170
         ScaleHeight     =   270
         ScaleWidth      =   1575
         TabIndex        =   186
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
            TabIndex        =   190
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
            TabIndex        =   189
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
         Index           =   43
         Left            =   6840
         TabIndex        =   106
         Top             =   2100
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   43
         Left            =   9060
         Picture         =   "frmTCpr.frx":2D5A
         Style           =   1  'Graphical
         TabIndex        =   182
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
         Index           =   42
         Left            =   6840
         TabIndex        =   98
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   42
         Left            =   9060
         Picture         =   "frmTCpr.frx":2F04
         Style           =   1  'Graphical
         TabIndex        =   181
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
         Index           =   41
         Left            =   6840
         TabIndex        =   89
         Top             =   1500
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   41
         Left            =   9060
         Picture         =   "frmTCpr.frx":30AE
         Style           =   1  'Graphical
         TabIndex        =   180
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
         Index           =   40
         Left            =   6840
         TabIndex        =   81
         Top             =   1200
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   40
         Left            =   9060
         Picture         =   "frmTCpr.frx":3258
         Style           =   1  'Graphical
         TabIndex        =   179
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
         Index           =   39
         Left            =   6840
         TabIndex        =   73
         Top             =   900
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   39
         Left            =   9060
         Picture         =   "frmTCpr.frx":3402
         Style           =   1  'Graphical
         TabIndex        =   178
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
         Index           =   38
         Left            =   6840
         TabIndex        =   65
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   38
         Left            =   9060
         Picture         =   "frmTCpr.frx":35AC
         Style           =   1  'Graphical
         TabIndex        =   177
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
         Index           =   34
         Left            =   3300
         TabIndex        =   120
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   34
         Left            =   6540
         Picture         =   "frmTCpr.frx":3756
         Style           =   1  'Graphical
         TabIndex        =   173
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
         Index           =   33
         Left            =   3300
         TabIndex        =   112
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   33
         Left            =   6540
         Picture         =   "frmTCpr.frx":3900
         Style           =   1  'Graphical
         TabIndex        =   172
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
         Index           =   32
         Left            =   3300
         TabIndex        =   104
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   32
         Left            =   6540
         Picture         =   "frmTCpr.frx":3AAA
         Style           =   1  'Graphical
         TabIndex        =   171
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
         Index           =   31
         Left            =   3300
         TabIndex        =   96
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   31
         Left            =   6540
         Picture         =   "frmTCpr.frx":3C54
         Style           =   1  'Graphical
         TabIndex        =   170
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
         Index           =   30
         Left            =   3300
         TabIndex        =   87
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   30
         Left            =   6540
         Picture         =   "frmTCpr.frx":3DFE
         Style           =   1  'Graphical
         TabIndex        =   169
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
         Index           =   29
         Left            =   3300
         TabIndex        =   79
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   29
         Left            =   6540
         Picture         =   "frmTCpr.frx":3FA8
         Style           =   1  'Graphical
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   1225
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   28
         Left            =   6540
         Picture         =   "frmTCpr.frx":4152
         Style           =   1  'Graphical
         TabIndex        =   165
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
         Index           =   28
         Left            =   3300
         TabIndex        =   71
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
         Index           =   27
         Left            =   3300
         TabIndex        =   63
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   27
         Left            =   6540
         Picture         =   "frmTCpr.frx":42FC
         Style           =   1  'Graphical
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   625
         Width           =   255
      End
      Begin VB.CheckBox chkDesactivar 
         Caption         =   "Des&activar Cuentas"
         ForeColor       =   &H00C00000&
         Height          =   200
         Left            =   4980
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   0
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   159
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
         Index           =   25
         Left            =   90
         TabIndex        =   139
         Top             =   3705
         Width           =   555
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Otros 3 :"
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
         Index           =   24
         Left            =   105
         TabIndex        =   131
         Top             =   3390
         Width           =   630
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Otros 2:"
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
         Index           =   23
         Left            =   105
         TabIndex        =   123
         Top             =   3075
         Width           =   585
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
         Left            =   4260
         TabIndex        =   129
         Top             =   3015
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
         Index           =   36
         Left            =   4260
         TabIndex        =   137
         Top             =   3330
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
         Index           =   37
         Left            =   4260
         TabIndex        =   145
         Top             =   3645
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
         Index           =   46
         Left            =   7500
         TabIndex        =   132
         Top             =   3015
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
         Index           =   47
         Left            =   7500
         TabIndex        =   140
         Top             =   3330
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
         Index           =   48
         Left            =   7500
         TabIndex        =   147
         Top             =   3645
         Width           =   1575
      End
      Begin VB.Label lblTexto 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   17
         Left            =   90
         TabIndex        =   74
         Top             =   1260
         Width           =   990
      End
      Begin VB.Label lblTexto 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   16
         Left            =   90
         TabIndex        =   66
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
         Index           =   45
         Left            =   7500
         TabIndex        =   124
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
         Index           =   44
         Left            =   7500
         TabIndex        =   116
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
         Index           =   43
         Left            =   7500
         TabIndex        =   108
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
         Index           =   42
         Left            =   7500
         TabIndex        =   100
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
         Index           =   41
         Left            =   7500
         TabIndex        =   91
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
         Index           =   40
         Left            =   7500
         TabIndex        =   83
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
         Index           =   39
         Left            =   7500
         TabIndex        =   75
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
         Index           =   38
         Left            =   7500
         TabIndex        =   67
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
         Index           =   34
         Left            =   4260
         TabIndex        =   121
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
         Index           =   33
         Left            =   4260
         TabIndex        =   113
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
         Index           =   32
         Left            =   4260
         TabIndex        =   105
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
         Index           =   31
         Left            =   4260
         TabIndex        =   97
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
         Index           =   30
         Left            =   4260
         TabIndex        =   88
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
         Index           =   29
         Left            =   4260
         TabIndex        =   80
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
         Index           =   28
         Left            =   4260
         TabIndex        =   72
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
         Index           =   27
         Left            =   4260
         TabIndex        =   64
         Top             =   600
         Width           =   2295
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
         Index           =   19
         Left            =   105
         TabIndex        =   90
         Top             =   1860
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
         Index           =   20
         Left            =   105
         TabIndex        =   99
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Otros :"
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
         Index           =   21
         Left            =   105
         TabIndex        =   107
         Top             =   2460
         Width           =   495
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Exonerado :"
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
         TabIndex        =   82
         Top             =   1560
         Width           =   870
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
         Index           =   15
         Left            =   90
         TabIndex        =   57
         Top             =   660
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
         Index           =   26
         Left            =   3420
         TabIndex        =   158
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
         Index           =   27
         Left            =   6960
         TabIndex        =   157
         Top             =   345
         Width           =   2265
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Otros 1 :"
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
         Index           =   22
         Left            =   90
         TabIndex        =   115
         Top             =   2760
         Width           =   630
      End
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   285
      Index           =   5
      Left            =   3195
      TabIndex        =   222
      Top             =   540
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65667073
      CurrentDate     =   37102
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato :"
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
      Index           =   37
      Left            =   165
      TabIndex        =   42
      Top             =   3600
      Width           =   705
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
      Index           =   60
      Left            =   3420
      TabIndex        =   44
      Top             =   3570
      Width           =   4620
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   480
      Left            =   15
      Top             =   3480
      Width           =   8415
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
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   900
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8500
      Y1              =   780
      Y2              =   780
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
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   420
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
      Height          =   280
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   90
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
      Index           =   8
      Left            =   3360
      TabIndex        =   19
      Top             =   1245
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
      Index           =   7
      Left            =   60
      TabIndex        =   17
      Top             =   1245
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
      Index           =   9
      Left            =   60
      TabIndex        =   22
      Top             =   1575
      Width           =   615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F. T.Cambio :"
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
      Left            =   2340
      TabIndex        =   11
      Top             =   900
      Width           =   930
   End
   Begin VB.Label lblTexto 
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   5
      Left            =   6585
      TabIndex        =   15
      Top             =   900
      Width           =   690
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
      Index           =   10
      Left            =   1860
      TabIndex        =   24
      Top             =   1575
      Width           =   705
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
      Index           =   4
      Left            =   4545
      TabIndex        =   13
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lblTexto 
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   6240
      TabIndex        =   6
      Top             =   465
      Width           =   555
   End
   Begin VB.Label lblTexto 
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   465
      Width           =   720
   End
   Begin VB.Label lblTexto 
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   135
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
              MINIMOINDICEIMPORTEME As Byte = 16, _
              MINIMOINDICEMAS As Byte = 1, _
              MINIMOINDICECUENTA As Byte = 27, _
              MINIMOINDICECCOSTO As Byte = 38, _
              CANTIDADIMPORTES As Byte = 11
'[Repetir en frmTCprMasGrd.
Private Const DIFERENCIAMASIMPORTE As Byte = 4, _
              DIFERENCIAMASCUENTA As Byte = 26, _
              DIFERENCIAMASCCOSTO As Byte = 37
Private Const CUENTASCONCCOSTO As Byte = 11
']

'[Repetir en frmTCprGrd y frmTCprMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
Private Const ps_OrdenCta As String = "01"
Private s_PedidoValiDsc As String
']
Private filtro As Boolean


Private Sub chkIndCDt_Click()
  If Not (chkIndCDt.Value = vbChecked) Then txtDato(52).Text = ""
  txtDato(52).Enabled = (chkIndCDt.Value)
  dtpDato(4).Enabled = (chkIndCDt.Value)
End Sub
Private Sub cmdMasIGV_Click()
   frmTCprMasIgv.Show vbModal
End Sub

Private Sub cmdPedido_Click()
  frmTPdoGrd.Show vbModal
End Sub

Private Sub Command1_Click()
  Dim TOTMN, TOTME As Byte
  
  ' Reverción del igv y calculo de la base exonerada
  If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
    TOTMN = Val(txtDato(MINIMOINDICEIMPORTEMN).Text) + Val(txtDato(MINIMOINDICEIMPORTEMN + 1).Text) + Val(txtDato(MINIMOINDICEIMPORTEMN + 2).Text) + Val(txtDato(MINIMOINDICEIMPORTEMN + 3).Text)
    If TOTMN = 0 And Val(txtDato(MINIMOINDICEIMPORTEMN + 4).Text) > 0 And Val(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text) > 0 Then
      Dim XRESPMN As String
      XRESPMN = MsgBox(Choose(gsIdioma, "No ha registrado importes " & TPOMON_NAC_TXT_2 & " Gravados o Exonerados; deseas calcularlos?", "You have not registered amount with tax or discharged. Do you want to calculate them?"), vbYesNo, Me.Caption)
      If XRESPMN = 6 Then
        txtDato(MINIMOINDICEIMPORTEMN).Text = 0
        txtDato(MINIMOINDICEIMPORTEMN + 3).Text = 0
        txtDato(MINIMOINDICEIMPORTEME).Text = 0
        txtDato(MINIMOINDICEIMPORTEME + 3).Text = 0
        txtDato(MINIMOINDICEIMPORTEMN).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEMN + 4).Text * 100) / CDec(gnPctIGV), FORMATO_NUM_1)
        txtDato(MINIMOINDICEIMPORTEMN + 3).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text - txtDato(MINIMOINDICEIMPORTEMN).Text - txtDato(MINIMOINDICEIMPORTEMN + 4).Text), FORMATO_NUM_1)
        txtDato(MINIMOINDICEIMPORTEME).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEMN) / txtDato(4)), FORMATO_NUM_1)
        txtDato(MINIMOINDICEIMPORTEME + 3).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEMN + 3) / txtDato(4)), FORMATO_NUM_1)
      End If
    End If
  Else
    TOTME = Val(txtDato(MINIMOINDICEIMPORTEME).Text) + Val(txtDato(MINIMOINDICEIMPORTEME + 1).Text) + Val(txtDato(MINIMOINDICEIMPORTEME + 2).Text) + Val(txtDato(MINIMOINDICEIMPORTEME + 3).Text)
    If TOTME = 0 And Val(txtDato(MINIMOINDICEIMPORTEME + 4).Text) > 0 And Val(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text) > 0 Then
      Dim XRESPME As String
      XRESPME = MsgBox(Choose(gsIdioma, "No ha registrado importes " & TPOMON_EXT_TXT_2 & " Gravados o Exonerados; deseas calcularlos?", "You have not registered amount with tax or discharged. Do you want to calculate them?"), vbYesNo, Me.Caption)
      If XRESPME = 6 Then
        txtDato(MINIMOINDICEIMPORTEMN).Text = 0
        txtDato(MINIMOINDICEIMPORTEMN + 3).Text = 0
        txtDato(MINIMOINDICEIMPORTEME).Text = 0
        txtDato(MINIMOINDICEIMPORTEME + 3).Text = 0
        txtDato(MINIMOINDICEIMPORTEME).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEME + 4).Text * 100) / CDec(gnPctIGV), FORMATO_NUM_1)
        txtDato(MINIMOINDICEIMPORTEME + 3).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text - txtDato(MINIMOINDICEIMPORTEME).Text - txtDato(MINIMOINDICEIMPORTEME + 4).Text), FORMATO_NUM_1)
        
        txtDato(MINIMOINDICEIMPORTEMN).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEME) * txtDato(4)), FORMATO_NUM_1)
        txtDato(MINIMOINDICEIMPORTEMN + 3).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEME + 3) * txtDato(4)), FORMATO_NUM_1)
      End If
    End If
  End If
  ' fin del p
End Sub

Private Sub Form_Load()
  Dim nContador As Integer
  
  filtro = True
  
   pbValidada = False
   pbFecha = True
   Me.KeyPreview = True
   
   With frmTCprGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
      txtLlave(0).MaxLength = .uorstMain!codaux.DefinedSize
      txtLlave(1).MaxLength = .uorstMain!codtdc.DefinedSize
      txtLlave(2).MaxLength = .uorstMain!serdoc.DefinedSize
      txtLlave(3).MaxLength = .uorstMain!nrodoc.DefinedSize
    ']
   
    '[Datos                            'Cambiar.
      With cboTpoMon
        .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
        .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
      End With
      
      With cboImpuesto
        .AddItem TEXT_Ninguno, TipoImpuesto.Ninguno
        .AddItem TEXT_ResponsableInscrito, TipoImpuesto.ResponsableInscrito
        .AddItem TEXT_ResponsableMonotributo, TipoImpuesto.ResponsableMonotributo
        .AddItem TEXT_Exepto, TipoImpuesto.Exepto
        .AddItem TEXT_NoAlcanzado, TipoImpuesto.NoAlcanzado
        .AddItem TEXT_ConsumidosFinal, TipoImpuesto.ConsumidosFinal
      End With
      With cboCategoria
        .AddItem TEXT_Ninguno, CategoriaDocumento.Ninguno
        .AddItem TEXT_ImpuestoDetallado, CategoriaDocumento.ImpuestoDetallado
        .AddItem TEXT_FacturaPublica, CategoriaDocumento.FacturaPublica
        .AddItem TEXT_FacturaContador, CategoriaDocumento.FacturaContador
        .AddItem TEXT_RetencionIva, CategoriaDocumento.RetencionIva
        .AddItem TEXT_RetencionIB, CategoriaDocumento.RetencionIB
        .AddItem TEXT_RetencionIG, CategoriaDocumento.RetenconIG
        .AddItem TEXT_RetencionSuss, CategoriaDocumento.RetencionSuss
        .AddItem TEXT_RetencionOtro, CategoriaDocumento.RetencionOtro
      End With
      
'      mskDato(0).MaxLength = .uorstMain!Tf1Cta.DefinedSize + 1
      txtDato(0).MaxLength = .uorstMain!coddro.DefinedSize
      txtDato(1).MaxLength = .uorstMain!NroCpb.DefinedSize
      txtDato(2).MaxLength = .uorstMain!refdoc.DefinedSize
      txtDato(Choose(gsIdioma, 3, 49)).MaxLength = .uorstMain!GloDoc.DefinedSize
      txtDato(Choose(gsIdioma, 49, 3)).MaxLength = .uorstMain!glodocx.DefinedSize
      txtDato(50).MaxLength = .uorstMain!pdocpr.DefinedSize
      txtDato(51).MaxLength = .uorstMain!codasi.DefinedSize
      txtDato(60).MaxLength = .uorstMain!codcon.DefinedSize
      txtDato(4).MaxLength = .uorstMain!ImpTCb.DefinedSize
      ' Importes
      For nContador = MINIMOINDICEIMPORTEMN To DIFERENCIAMASCUENTA
        txtDato(nContador).MaxLength = 16
      Next nContador
      For nContador = MINIMOINDICECUENTA To DIFERENCIAMASCCOSTO
        txtDato(nContador).MaxLength = 8
      Next nContador
      For nContador = MINIMOINDICECCOSTO To (MINIMOINDICECCOSTO + CUENTASCONCCOSTO - 1)
        txtDato(nContador).MaxLength = 5
      Next nContador
      txtDato(52).MaxLength = .uorstMain!NroCDt.DefinedSize
      ' importes igv
      txtDato(53).MaxLength = 16
      txtDato(54).MaxLength = 16
      txtDato(55).MaxLength = 16
      txtDato(56).MaxLength = 16
      txtDato(57).MaxLength = 16
      txtDato(58).MaxLength = 16
      ' Datos detalle pdb
      txtDetalle(0).MaxLength = .uorstMain!codaduana.DefinedSize
      txtDetalle(1).MaxLength = .uorstMain!annodua.DefinedSize
      txtDetalle(2).MaxLength = .uorstMain!nrodua.DefinedSize
      txtDato(59).MaxLength = .uorstMain!codtdc_ref.DefinedSize
      txtDetalle(3).MaxLength = .uorstMain!serdoc_ref.DefinedSize
      txtDetalle(4).MaxLength = .uorstMain!nrodoc_ref.DefinedSize
      txtDetalle(5).MaxLength = 14
      txtDetalle(6).MaxLength = 14
      txtDetalle(7).MaxLength = 14
      txtDetalle(8).MaxLength = 14
      
      cboDetraccion.AddItem Choose(gsIdioma, "Ninguna", "Neither"), 0
'ini 2015-07-02 adic tabla detrac
        With frmTCprGrd.uorstcodetrac
            If .RecordCount > 0 Then .MoveFirst
            If Not .EOF Then
                Do While Not .EOF
                      '2015-07-08 cambio de decima a % cboDetraccion.AddItem !coddetrac & " " & !detdetrac & " " & Trim(Str(!pctdetrac * 100)) & "%"
                     cboDetraccion.AddItem !coddetrac & " " & !detdetrac & " " & Trim(Str(!pctdetrac)) & "%"
                    .MoveNext
                Loop
            End If
        End With
     
''      For nContador = 1 To UBound(aDtraccDet, 1)
''        If aDtraccEst(nContador) = 1 Then
''        cboDetraccion.AddItem aDtraccDet(nContador)
''        End If
''      Next nContador
'fin 2015-07-02 adic tabla detrac

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
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(38, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Proveedor:", "Tipo Doc.:", "NºDoc.:", "F.Operación:", "F.Emisión:", "F.Vcmto.:", "F.T.Cambio :", "Referencia:", "Glosa:", "Moneda:", "T.Cambio:", "Diario:", "Comprobante:", "Form. Nº:", "Fecha:", "Op.Gravada:", "Op.Gr./No Gr.:", "Op. No Grav.:", "Exonerado :", "IGV :", "ISC :", "Otros :", "Otros 1-P.IVA :", "Otros 2-P.IB :", "Otros 3 :", "Total :", "Cuenta Contable", "Centro de Costo", "Nro Pedido :", "Asiento Tipo :", "Numero DUA :", "Tipo Tasa :", "Tipo Doc.:", "Nº Doc. :", "F. Emisión :", "Base Impo. :", "IGV :", "Ord.Servicio :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Supplier :", "Type Doc.:", "NºDoc.:", "Operti.Date:", "IssueDate:", "Due Date:", "Exchange Date:", "Reference :", "Gloss:", "Currency:", "R.Exchange:", "Journal:", "Voucher:", "N° Form:", "Date:", "Op. with Taxes:", "Op. with/without Taxes:", "Op.Without Taxes:", "Discharged :", "GST :", "SCT :", "Others :", "Others 1-P.IVA :", "Others 2-P.IB :", "Others 3 :", "Total :", "Accountable Account", "Cost Center", "Nro Order :", "Standar Recorded :", "DUA Number :", "Rate Type :", "Type Doc.:", "Nº Doc. :", "IssueDate :", "Base Amount:", "GST :", "Ord.Service :")
  Next nElemento
  chkCalcularIGV.Caption = Choose(gsIdioma, "Calcular I.G.&V.", "Calculate G.&S.T.")
  chkCalcularISC.Caption = Choose(gsIdioma, "Calcular I.S.&C.", "Calculate S.&C.T.")
  chkIndCDt.Caption = Choose(gsIdioma, "Detracción", "Deduction")
  chkDesactivar.Caption = Choose(gsIdioma, "Des&activar Cuentas", "Dis&able Accounts")
  chkIndPreGen.Caption = Choose(gsIdioma, "Cuentas &Registradas", "&Registered Accounts")
  fraExterior.Caption = Choose(gsIdioma, " Compra Externa ", " External Purchase ")
  chkIndreten.Caption = Choose(gsIdioma, "Afecto a Retención", "Afecto Withholding")
  fraDetracion.Caption = Choose(gsIdioma, " Detracción ", " Deduction ")
  fraReferencia.Caption = Choose(gsIdioma, " Documento de Referencia ", " Reference Document ")
  cmdAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  cmdPedido.Caption = Choose(gsIdioma, "Pedido", "Order")
  sstMain.TabCaption(0) = Choose(gsIdioma, "I&mportes", "A&mounts")
  sstMain.TabCaption(1) = Choose(gsIdioma, "C&uentas", "Acco&unts")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']
   dgrDetalle.MarqueeStyle = dbgHighlightRow
   Set dgrDetalle.DataSource = frmTCprGrd.uorstCOCpbDet
   
   cboTpoMon.ListIndex = TPOMON_NAC_IND
   chkMonedaActiva.Value = vbChecked
   sstMain.Tab = 0
']
End Sub

Private Sub Form_Activate()
'' ppDatosGrid
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
  If gbCieCpr Then MsgBox TEXT_9016, vbCritical: Exit Sub
  
  pbCorregir = True
  frmTCprGrd.uocnnMain.BeginTrans     'Cambiar Formulario de Grid. 'INICIA TRANSACCION.
  
  cmdRetroceder.Enabled = False
  cmdAvanzar.Enabled = False
  cmdCorregir.Enabled = False
  cmdGrabar.Enabled = True
  cmdDeshacer.Enabled = True
  upHabilitacion (True)
  txtDato(0).Enabled = (chkIndPreGen.Value = 0)
  lblDatoDeta(0).Enabled = (chkIndPreGen.Value = 0)
  cmdDatoAyud(0).Enabled = (chkIndPreGen.Value = 0)
  
  ' Dato con el foco al corregir
  dtpDato(3).SetFocus
  ' Para no cambiar fechas
  pbFecha = False

End Sub

Public Sub cmdGrabar_Click()

  On Error GoTo Err
  
  '[Propio del formulario.
  Dim dnSumaMN As Double, dnSumaME As Double
  Dim nIndices As Integer
  'ini 2015-05-21 valida detrac
  If chkIndCDt.Value = 1 Then
    If cboDetraccion.ListIndex = 0 Then MsgBox TEXT_6002, vbCritical: cboDetraccion.SetFocus: Exit Sub
'    If txtDato(59).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(59).SetFocus: Exit Sub
'    If txtDetalle(3).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDetalle(3).SetFocus: Exit Sub
'    If txtDetalle(4).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDetalle(4).SetFocus: Exit Sub
'    If txtDetalle(5).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDetalle(5).SetFocus: Exit Sub
'    If txtDetalle(6).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDetalle(6).SetFocus: Exit Sub
  End If
  'ini 2015-05-21 valida detrac
  
  ' validacion de compra exterior
  If txtLlave(1).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(0).SetFocus: Exit Sub
  ' validacion de compra exterior
  If ((txtLlave(1).Text = "50" Or txtLlave(1).Text = "52") And chkIndCprext.Value = Unchecked) Then MsgBox Choose(gsIdioma, "Seleccione Parametro Compra Externa", "Purchase Select External Parameter"), vbCritical: chkIndCprext.SetFocus: Exit Sub
  If ((txtLlave(1).Text = "50" Or txtLlave(1).Text = "52") And (txtDetalle(0).Text = "" Or txtDetalle(1).Text = "" Or txtDetalle(2).Text = "")) Then MsgBox Choose(gsIdioma, "Debe ingresar Documento Unico de Tra", "You Must enter the Document Series Reference"), vbCritical: txtDetalle(3).SetFocus: Exit Sub
  
  ' validacion de nota credito
  If ((txtLlave(1).Text = "07" Or txtLlave(1).Text = "08" Or txtLlave(1).Text = "91") And txtDato(59).Text = "") Then MsgBox Choose(gsIdioma, "Seleccione Tipo de Documento de Referencia", "Select Document Type Reference"), vbCritical: txtDato(59).SetFocus: Exit Sub
  If ((txtLlave(1).Text = "07" Or txtLlave(1).Text = "08" Or txtLlave(1).Text = "91") And txtDetalle(3).Text = "") Then MsgBox Choose(gsIdioma, "Debe ingresar Serie de Documento de Referencia", "You Must enter the Document Series Reference"), vbCritical: txtDetalle(3).SetFocus: Exit Sub
  If ((txtLlave(1).Text = "07" Or txtLlave(1).Text = "08" Or txtLlave(1).Text = "91") And txtDetalle(4).Text = "") Then MsgBox Choose(gsIdioma, "Debe ingresar Documento de Referencia", "You Must enter the Document Reference"), vbCritical: txtDetalle(4).SetFocus: Exit Sub
  
  ' validacion de informacion
  If txtDato(0).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(0).SetFocus: Exit Sub
  With frmTCprGrd.uorstMain
    For nIndices = 0 To CANTIDADIMPORTES - 2
      dnSumaMN = dnSumaMN + CDec(txtDato(MINIMOINDICEIMPORTEMN + nIndices).Text)
      dnSumaME = dnSumaME + CDec(txtDato(MINIMOINDICEIMPORTEME + nIndices).Text)
    Next nIndices
    If dnSumaMN <> CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text) Then
      If (cboTpoMon.ListIndex = TPOMON_EXT_IND And cmdMas(CANTIDADIMPORTES).Tag <> INDMASCTA_MAS And (Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text)) <= 0.02)) Then
        If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text)))) & "." & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
        txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text = Format(dnSumaMN, FORMATO_NUM_1)
      Else
        If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
      End If
    ElseIf dnSumaME <> CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text) Then
      If (cboTpoMon.ListIndex = TPOMON_NAC_IND And cmdMas(CANTIDADIMPORTES).Tag <> INDMASCTA_MAS And (Abs(dnSumaME - CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)) <= 0.02)) Then
        If MsgBox(TEXT_9012 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaME - CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)))) & "." & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
        txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text = Format(dnSumaME, FORMATO_NUM_1)
      Else
        If MsgBox(TEXT_9012 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaME - CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
      End If
    End If
  End With
  
  ' Valido los saldos del pedido
  If txtDato(50).Text <> "" Then
    If Not pfValidaPedido(txtDato(50).Text, "N") Then Exit Sub
  End If
  
  ' Genero las cuentas de acuerdo al asiento tipo
  If txtDato(51).Text <> "" And pbNuevo Then
    ppInsDelCtaCos txtDato(51).Text, INDCCO_INA
    ppInsDelCtaCos txtDato(51).Text, INDCCO_ACT
  End If
  
  'Valido las Cuentas esten Correctas(llenas para todas los valores)
  If chkIndPreGen.Value = vbChecked Then
  
    'ini 2014-07-10 inhabilita y activa cuentas registradas
'    chkIndPreGen.Value = IIf(ValidaCtasCCo, 1, 0)
'    If Not (chkIndPreGen.Value = vbChecked) Then
'      If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
'        Exit Sub
'      End If
'    End If
    '************
    If ValidaCtasCCo = False Then
      If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
        Exit Sub
      Else
        chkIndPreGen.Value = 0
      End If
    End If
    'fin 2014-07-10 inhabilita y activa cuentas registradas
    '************
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
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
  
    '[Actualiza grid..
    .uorstMain_Grd.Requery
    .upDatosGrid
    .uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "'"
    ']
  
    pbCorregir = False
    
    If pbNuevo Then
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
  
'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
fEstMayUpd
'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion

'If fEstMayUpd(frmTCprGrd.uocnnMain) = 0 Then
'End If
'fEstMayUpd (frmTCprGrd.uocnnMain)
'fEstMayUpd frmTCprGrd.uocnnMain
  
  
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
  End Select
  ppAyuBus AYULLA, Index
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
  Select Case Index                   'Cambiar. Añadir índices.
   Case 0, 50, 51, 59, 60, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
    txtDato(Index).SetFocus
  End Select
  Select Case Index
    Case 38, 39, 40, 41
        filtro = False
    Case Else
        filtro = True
  End Select
  ppAyuBus AYUDAT, Index
End Sub



Private Sub txtDetalle_GotFocus(Index As Integer)
  txtDetalle(Index).SelStart = 0
  txtDetalle(Index).SelLength = txtDetalle(Index).MaxLength
End Sub
Private Sub txtDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
  If Len(Trim(txtDetalle(Index))) + 1 = txtDetalle(Index).MaxLength Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtDetalle_LostFocus(Index As Integer)
  Select Case Index
   Case 5, 6, 7, 8
    If CDec(txtDato(4).Text) <= 0 Then
      MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
      txtDato(4).SetFocus
      Exit Sub
    End If
      
    If CDec(txtDetalle(Index).Text) = 0 Then
      txtDetalle(Index).Text = Format(0, FORMATO_NUM_1)
      If (Index = 5 Or Index = 6) And cboTpoMon.ListIndex = TPOMON_EXT_IND Then
        txtDetalle(Index).Text = Format(CDec(txtDetalle(IIf(Index = 5, 7, 8)).Text) * CDec(txtDato(4).Text), FORMATO_NUM_1)
      ElseIf (Index = 7 Or Index = 8) And cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDetalle(Index).Text = Format(CDec(txtDetalle(IIf(Index = 7, 5, 6)).Text) / CDec(txtDato(4).Text), FORMATO_NUM_1)
      End If
    End If
    If chkMonedaActiva.Value = vbChecked Then
      If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDetalle(IIf(Index = 5, 7, 8)).Text = Format(gfRedond(CDec(txtDetalle(Index).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
      Else
        txtDetalle(IIf(Index = 7, 5, 6)).Text = Format(gfRedond(CDec(txtDetalle(Index).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
      End If
    End If
  End Select
End Sub
Private Sub txtDetalle_Validate(Index As Integer, Cancel As Boolean)
  'Completa con ceros a la izquierda.
  Select Case Index
   Case 0, 1, 2, 3, 4
    If Len(Trim(txtDetalle(Index).Text)) <> 0 And Len(Trim(txtDetalle(Index).Text)) <> txtDetalle(Index).MaxLength Then
      txtDetalle(Index) = gfCeros(txtDetalle(Index).Text, txtDetalle(Index).MaxLength, 0, "0")
    End If
   Case 5, 6, 7, 8
    txtDetalle(Index).Text = IIf(Not IsNumeric(txtDetalle(Index).Text), 0, txtDetalle(Index).Text)
    txtDetalle(Index).Text = Format(txtDetalle(Index).Text, FORMATO_NUM_1)
  End Select
End Sub

Private Sub txtllave_GotFocus(Index As Integer)
'ini 2015-08-07 valor 1683 defec
      If txtLlave(1).Text = "10" And Index = 1 Then
        txtLlave(2).Text = "1683"
        txtLlave(3).SetFocus
      End If
'fin 2015-08-07 valor 1683 defec

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
'ini 2015-08-07 valor 1683 defec
      If txtLlave(1).Text = "10" And Index = 1 Then
        txtLlave(2).Text = "1683"
        txtLlave(3).SetFocus
      End If
'fin 2015-08-07 valor 1683 defec

  If pbValidada Then                  'Cambiar.
    txtLlave(0).Enabled = False
    txtLlave(1).Enabled = False
    txtLlave(2).Enabled = False
    txtLlave(3).Enabled = False
    lblLlaveDeta(0).Enabled = False
    lblLlaveDeta(1).Enabled = False
    cmdLlaveAyud(0).Enabled = False
    cmdLlaveAyud(1).Enabled = False
    'ini 2014-07-09 inhabilita y activa cuentas registradas
    chkIndPreGen.Value = 1 'activar chek
    chkIndPreGen.Enabled = False
    'fin 2014-07-09 inhabilita y activa cuentas registradas
  End If
  If pbValidada And dtpDato(3).Enabled Then dtpDato(3).SetFocus 'Cambiar.
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
  'On Error GoTo Err
  Dim dvRegistro As Variant
   
  '[Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
  Select Case Index
'2014-08-20 va hacia arriba
   Case 1                      'Cambiar (añadir índices).
     If Len(Trim(txtLlave(Index).Text)) <> 0 And Len(Trim(txtLlave(Index).Text)) <> txtLlave(Index).MaxLength Then
      txtLlave(Index) = gfCeros(txtLlave(Index).Text, txtLlave(Index).MaxLength, 0, "0")
    End If
  Case 2, 3                         'Cambiar (añadir índices).
   
   
      'ini 2014-07-10 validacion T.Doc=05
      'If txtLlave(1).Text = "05" And InStr(1, "1234", Trim(Str(Val(txtLlave(2).Text)))) = 0 Then
      '2014-07-16 corregido validacion de tipodoc=05
      If txtLlave(1).Text = "05" And Not (Val(txtLlave(2).Text) >= 1 And Val(txtLlave(2).Text) <= 4) Then
        'MsgBox ("Error, en la SERIE debe poner los siguientes digitos: 1=Boleto Manual, 2=Boleto Automatico, 3=Boleto Electronico, 4=Otros")
        MsgBox (TEXT_9017)
        txtLlave(Index).SelStart = 0
        txtLlave(Index).SelLength = txtLlave(Index).MaxLength
        Cancel = True
        Exit Sub
      End If
      'fin 2014-07-10 validacion T.Doc=05
'ini 2015-08-07 valor 1683 defec
''      If txtLlave(1).Text = "10" And Val(txtLlave(2).Text) <> 1683 Then
''        MsgBox (TEXT_8010)
''        txtLlave(Index).SelStart = 0
''        txtLlave(Index).SelLength = txtLlave(Index).MaxLength
''        Cancel = True
''        Exit Sub
''      End If
'fin 2015-08-07 valor 1683 defec
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
      Set .uorstTemporal = .uocnnMain.Execute("SELECT MesPvs FROM COCprDoc WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND CodAux='" & txtLlave(0).Text & "' AND CodTDc ='" & txtLlave(1).Text & "' AND SerDoc='" & txtLlave(2).Text & "' AND NroDoc='" & txtLlave(3).Text & "'")
      If .uorstTemporal.RecordCount > 0 Then
        MsgBox TEXT_8007 & Chr(13) & Choose(gsIdioma, "(mes ", "(month ") & gfMesLet("01" & .uorstTemporal!mespvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
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
  
  '[Convierte a mayúsculas.
  '   If Index = 1 Then                   'Cambiar (añadir índices).
  '      KeyAscii = Asc(UCase(Chr(KeyAscii)))
  '   End If
  ']
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then ppAyuBus AYUDAT, Index
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
  
  Select Case Index
   Case MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
    If CDec(txtDato(4).Text) <= 0 Then
      MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
      txtDato(4).SetFocus
      Exit Sub
    End If
    If CDec(txtDato(Index).Text) = 0 Then
      txtDato(Index).Text = Format(0, FORMATO_NUM_1)
      If Index >= MINIMOINDICEIMPORTEMN And Index < MINIMOINDICEIMPORTEME And cboTpoMon.ListIndex = TPOMON_EXT_IND Then
        txtDato(Index).Text = Format(CDec(txtDato(Index + CANTIDADIMPORTES).Text) * CDec(txtDato(4).Text), FORMATO_NUM_1)
      ElseIf Index >= MINIMOINDICEIMPORTEME And Index < (MINIMOINDICEIMPORTEME + CANTIDADIMPORTES) And cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDato(Index).Text = Format(CDec(txtDato(Index - CANTIDADIMPORTES).Text) / CDec(txtDato(4).Text), FORMATO_NUM_1)
      End If
    End If
    If (Index >= MINIMOINDICEIMPORTEMN And Index <= MINIMOINDICEIMPORTEMN + 2) Then
      If chkCalcularIGV Then txtDato(MINIMOINDICEIMPORTEMN + 4).Text = Format(Round(CDec(txtDato(MINIMOINDICEIMPORTEMN).Text) * CDec(gnPctIGV) / 100, 2) + Round(CDec(txtDato(MINIMOINDICEIMPORTEMN + 1).Text) * CDec(gnPctIGV1) / 100, 2) + Round(CDec(txtDato(MINIMOINDICEIMPORTEMN + 2).Text) * CDec(gnPctIGV2) / 100, 2), FORMATO_NUM_1)
      If chkCalcularISC Then txtDato(MINIMOINDICEIMPORTEMN + 5).Text = Format((CDec(txtDato(MINIMOINDICEIMPORTEMN).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 1).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 2).Text)) * CDec(gnPctISC) / 100, FORMATO_NUM_1)
      ' Calculo individual ma 31/01/2004
      If chkCalcularIGV Then txtDato(48 + Index).Text = Format((CDec(txtDato(Index).Text)) * CDec(IIf(Index = MINIMOINDICEIMPORTEMN, gnPctIGV, IIf(Index = MINIMOINDICEIMPORTEMN + 1, gnPctIGV1, gnPctIGV2))) / 100, FORMATO_NUM_1)
      If (chkMonedaActiva.Value = vbChecked) And (cboTpoMon.ListIndex = TPOMON_NAC_IND) Then
        If CDec(txtDato(MINIMOINDICEIMPORTEMN + 4).Text) > 0 Then txtDato(MINIMOINDICEIMPORTEME + 4).Text = Format(gfRedond(CDec(txtDato(MINIMOINDICEIMPORTEMN + 4).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
        If CDec(txtDato(MINIMOINDICEIMPORTEMN + 5).Text) > 0 Then txtDato(MINIMOINDICEIMPORTEME + 5).Text = Format(gfRedond(CDec(txtDato(MINIMOINDICEIMPORTEMN + 5).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
        ' Calculo individual ma 31/01/2004
        If CDec(txtDato(48 + Index).Text) > -0.01 Then txtDato(51 + Index).Text = Format(Round(CDec(txtDato(48 + Index).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
      End If
    ElseIf (Index >= MINIMOINDICEIMPORTEME And Index <= MINIMOINDICEIMPORTEME + 2) Then
      If chkCalcularIGV Then txtDato(MINIMOINDICEIMPORTEME + 4).Text = Format(Round(CDec(txtDato(MINIMOINDICEIMPORTEME).Text) * CDec(gnPctIGV) / 100, 2) + Round(CDec(txtDato(MINIMOINDICEIMPORTEME + 1).Text) * CDec(gnPctIGV1) / 100, 2) + Round(CDec(txtDato(MINIMOINDICEIMPORTEME + 2).Text) * CDec(gnPctIGV2) / 100, 2), FORMATO_NUM_1)
      If chkCalcularISC Then txtDato(MINIMOINDICEIMPORTEME + 5).Text = Format((CDec(txtDato(MINIMOINDICEIMPORTEME).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 1).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 2).Text)) * CDec(gnPctISC) / 100, FORMATO_NUM_1)
      ' Calculo individual ma 31/01/2004
      If chkCalcularIGV Then txtDato(40 + Index).Text = Format((CDec(txtDato(Index).Text)) * CDec(IIf(Index = MINIMOINDICEIMPORTEME, gnPctIGV, IIf(Index = MINIMOINDICEIMPORTEME + 1, gnPctIGV1, gnPctIGV2))) / 100, FORMATO_NUM_1)
      If (chkMonedaActiva.Value = vbChecked) And (cboTpoMon.ListIndex = TPOMON_EXT_IND) Then
        If CDec(txtDato(MINIMOINDICEIMPORTEME + 4).Text) > 0 Then txtDato(MINIMOINDICEIMPORTEMN + 4).Text = Format(gfRedond(CDec(txtDato(MINIMOINDICEIMPORTEME + 4).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
        If CDec(txtDato(MINIMOINDICEIMPORTEME + 5).Text) > 0 Then txtDato(MINIMOINDICEIMPORTEMN + 5).Text = Format(gfRedond(CDec(txtDato(MINIMOINDICEIMPORTEME + 5).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
        ' Calculo individual ma 31/01/2004
        If CDec(txtDato(40 + Index).Text) > -0.01 Then txtDato(37 + Index).Text = Format(Round(CDec(txtDato(40 + Index).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
      End If
    End If
  
    'Cálculo del total.
    If (Index = MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1 And txtDato(Index).Text = 0) Or (Index = MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1 And txtDato(Index).Text = 0) Then
      If Index = MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1 Then
        txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEMN).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 1).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 2).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 3).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 4).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 5).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 6).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 7).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 8).Text) + CDec(txtDato(MINIMOINDICEIMPORTEMN + 9).Text), FORMATO_NUM_1)
      Else
        txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text = Format(CDec(txtDato(MINIMOINDICEIMPORTEME).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 1).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 2).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 3).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 4).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 5).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 6).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 7).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 8).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + 9).Text), FORMATO_NUM_1)
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
    '///Angel 18/12/2003
    '///Se agrega para la eliminacion del dato del centro de costo digitado directamente
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
        If frmTCprGrd.uorstCOCprDocCta.RecordCount > 0 Then
          ppAbreCtaCCo
          If frmTCprGrd.uorstCOCprDocCta.State = adStateOpen Then
            frmTCprGrd.uorstCOCprDocCta.MoveFirst
            Do
              If frmTCprGrd.uorstCOCprDocCta!codaux = txtLlave(0).Text And _
              frmTCprGrd.uorstCOCprDocCta!codtdc = txtLlave(1).Text And _
              frmTCprGrd.uorstCOCprDocCta!serdoc = txtLlave(2).Text And _
              frmTCprGrd.uorstCOCprDocCta!nrodoc = txtLlave(3).Text And _
              Trim(frmTCprGrd.uorstCOCprDocCta!tpocnc) = Trim(Str(Index - MINIMOINDICECUENTA + 1)) Then
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
    End If
   Case 50
    ' Inicializo clon de pedido
    txtDato(Index).Tag = txtDato(Index).Text
   Case 51
    If txtDato(Index).Text <> "" And pbNuevo Then
      chkDesactivar.Value = vbChecked
    End If
  End Select
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err
 
    Select Case Index
    Case 38, 39, 40, 41
        filtro = False
    Case Else
        filtro = True
    End Select

 
  'Completa con ceros a la izquierda.
   Select Case Index
   Case 52, MINIMOINDICECCOSTO To CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
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
   Case 0, 50, 51, 59, 60, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYUDAT, Index)
      If Cancel Then Exit Sub
      If Index >= MINIMOINDICECUENTA And Index < MINIMOINDICECCOSTO Then
'         If frmTCprGrd.uorstCOCta.RecordCount > 0 And txtDato(Index + CUENTASCONCCOSTO).Text = "" Then
         If frmTCprGrd.uorstCoCta.RecordCount > 0 Then
            If Not frmTCprGrd.uorstCoCta.EOF Then
                If frmTCprGrd.uorstCoCta!indcco = INDCCO_ACT Then
                  ' Inicializo el centro de costos
                  txtDato(Index + CUENTASCONCCOSTO).Tag = txtDato(Index + CUENTASCONCCOSTO).Text
                  txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index).Tag <> txtDato(Index).Text, "", txtDato(Index + CUENTASCONCCOSTO).Text)
                  If Not IsNull(frmTCprGrd.uorstCoCta!codcco_def) Then
                    txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index + CUENTASCONCCOSTO).Text = "", frmTCprGrd.uorstCoCta!codcco_def, txtDato(Index + CUENTASCONCCOSTO).Text)
                  Else
                    txtDato(Index + CUENTASCONCCOSTO).Text = txtDato(Index + CUENTASCONCCOSTO).Tag
                  End If
                  txtDato(Index + CUENTASCONCCOSTO).Enabled = True
                  cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = True
    '///Angel 22/12/2003
                Else
                  txtDato(Index + CUENTASCONCCOSTO).Text = ""
                  lblDatoDeta(Index + CUENTASCONCCOSTO).Caption = ""
                  txtDato(Index + CUENTASCONCCOSTO).Enabled = False
                  cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = False
    '///
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
  Dim s_PedidoCco As String
  
  'If gnindpedido = 1 And txtDato(50) = "" And filtro = False Then
  s_PedidoCco = "AND indpdocpr=" & IIf(txtDato(50) = "", INDCCO_INA, INDCCO_ACT) & " "
  s_PedidoCco = ""
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
      modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "' " & s_PedidoCco, txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 50                           'Cambiar (añadir índices).
      ' Filtro de seleccion
      cmdDatoAyud(tnIndex).Tag = "a.codaux = '" & txtLlave(0).Text & "' "
      If ps_Plataforma = pSrvMySql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND a.fehpdo<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
      ElseIf ps_Plataforma = pSrvSql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND a.fehpdo<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
      End If
      modAyuBus.Pdo_Sal cmdDatoAyud(tnIndex).Tag, txtDato(tnIndex).Text, 0, 0, Me.Top + fraPedido.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraPedido.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 51
      modAyuBus.Asi_Cod "tpoasi='" & TPOGNR_CPR & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAsiento.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAsiento.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 59
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + sstMain.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, sstMain.Left + Me.Left + fraReferencia.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 60                             ' orden de servicio
      ' Filtro de seleccion
      If ps_Plataforma = pSrvMySql Then
        cmdDatoAyud(tnIndex).Tag = "a.fehcon<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
      ElseIf ps_Plataforma = pSrvSql Then
        cmdDatoAyud(tnIndex).Tag = "a.fehcon<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
      End If
      'cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND a.tpognr='" & INDANU_FAL & "' "
      modAyuBus.Con_Sal cmdDatoAyud(tnIndex).Tag, txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
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
         With frmTCprGrd.uorstTGAux
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodAux='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !razAux
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
               lblLlaveDeta(tnIndex).Caption = " " & IIf(IsNull(!dettdc), "", !dettdc)
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
         With frmTCprGrd.uorstCODro
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodDro='" & txtDato(tnIndex).Text & "'"
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
         With frmTCprGrd.uorstCoCta
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodCta='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
                'ini 2015-06-30 correccion tipo mon cta
                If tnIndex = 37 And !tpomon <> Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT) Then
                          MsgBox TEXT_9021, vbExclamation
                          ppAyuDet = True
                End If
                'fin 2015-06-30 correccion tipo mon cta
              lblDatoDeta(tnIndex).Caption = " " & Left(!detcta, 18)
            End If
         End With
      Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
         
         If gnIndPedido = 1 And txtDato(50) <> "" And filtro = False Then
            With frmTCprGrd.uorstCoCCox
              If .RecordCount > 0 Then .MoveFirst
              .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
              If .EOF Then
                MsgBox TEXT_8006, vbExclamation
                ppAyuDet = True
              Else
                lblDatoDeta(tnIndex).Caption = " " & Left(!detcco, 12)
              End If
            End With
         ElseIf gnIndPedido = 1 And txtDato(50) = "" And filtro = False Then
            With frmTCprGrd.uorstCoCCoy
              If .RecordCount > 0 Then .MoveFirst
              .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
              If .EOF Then
                MsgBox TEXT_8006, vbExclamation
                ppAyuDet = True
              Else
                lblDatoDeta(tnIndex).Caption = " " & Left(!detcco, 12)
              End If
            End With
         Else
            With frmTCprGrd.uorstCoCCo
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblDatoDeta(tnIndex).Caption = " " & Left(!detcco, 12)
            End If
            End With
         End If
       Case 50
        If txtDato(tnIndex).Text = "" Then
          lblDatoDeta(tnIndex).Caption = ""
          Exit Function
        End If
        ppAyuDet = Not pfValidaPedido(txtDato(tnIndex).Text, "S")
       Case 51
        If txtDato(tnIndex).Text = "" Then
          lblDatoDeta(tnIndex).Caption = ""
          Exit Function
        End If
        With frmTCprGrd.uorstCoAsiTipo
          If .RecordCount > 0 Then .MoveFirst
          .Find "codasi='" & txtDato(tnIndex).Text & "'"
          If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
          Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detasi), "", !detasi)
          End If
        End With
       Case 59
        If txtDato(tnIndex).Text = "" Then
          lblDatoDeta(tnIndex).Caption = ""
          Exit Function
        End If
        With frmTCprGrd.uorstTGTDc
          If .RecordCount > 0 Then .MoveFirst
          .Find "codtdc='" & txtDato(tnIndex).Text & "'"
          If .EOF Then
            MsgBox TEXT_8006, vbExclamation
            ppAyuDet = True
          Else
            lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!dettdc), "", !dettdc)
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
  Dim dnIndices As Integer
  ']

  With frmTCprGrd.uorstMain           'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !codaux = txtLlave(0).Text
        !codtdc = txtLlave(1).Text
        !serdoc = txtLlave(2).Text
        !nrodoc = txtLlave(3).Text
        !mespvs = gsMesAct
        !PctIGV = CDec(gnPctIGV)
        !PctISC = CDec(gnPctISC)
      End If

      'Datos.
      !tpomon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
      !indpregen = IIf(chkIndPreGen.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      !indcdt = IIf(chkIndCDt.Value = vbChecked, INDCDT_ACT, INDCDT_INA)
      !fehope = dtpDato(3).Value
      !feedoc = dtpDato(0).Value
      !fevdoc = dtpDato(1).Value
      !feedoc_ref = dtpDato(2).Value
      !ferdoc = dtpDato(5).Value
      
      !FehCDt = dtpDato(4).Value
      !coddro = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      !NroCpb = txtDato(1).Text
      !refdoc = txtDato(2).Text
      !GloDoc = IIf(txtDato(Choose(gsIdioma, 3, 49)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 49)).Text)
      !glodocx = IIf(txtDato(Choose(gsIdioma, 49, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 49, 3)).Text)
      !pdocpr = IIf(txtDato(50).Text = "", Null, txtDato(50).Text)
      !codcon = IIf(txtDato(60).Text = "", Null, txtDato(60).Text)
      !codasi = IIf(txtDato(51).Text = "", Null, txtDato(51).Text)
      !NroCDt = txtDato(52).Text
      !ImpTCb = CDec(txtDato(4).Text)
      ' Importes
      !impogr_mn = CDec(txtDato(5).Text)
      !ImpOGN_MN = CDec(txtDato(6).Text)
      !ImpONG_MN = CDec(txtDato(7).Text)
      !impexo_mn = CDec(txtDato(8).Text)
      !impigv_mn = CDec(txtDato(9).Text)
      !impisc_mn = CDec(txtDato(10).Text)
      !impoim_mn = CDec(txtDato(11).Text)
      !impoi1_mn = CDec(txtDato(12).Text)
      !impoi2_mn = CDec(txtDato(13).Text)
      !impoi3_mn = CDec(txtDato(14).Text)
      !imptot_mn = CDec(txtDato(15).Text)
      !ImpIGV_OGr_MN = CDec(txtDato(53).Text)
      !ImpIGV_OGN_MN = CDec(txtDato(54).Text)
      !ImpIGV_ONG_MN = CDec(txtDato(55).Text)
      !impogr_me = CDec(txtDato(16).Text)
      !ImpOGN_ME = CDec(txtDato(17).Text)
      !ImpONG_ME = CDec(txtDato(18).Text)
      !impexo_me = CDec(txtDato(19).Text)
      !impigv_me = CDec(txtDato(20).Text)
      !impisc_me = CDec(txtDato(21).Text)
      !impoim_me = CDec(txtDato(22).Text)
      !impoi1_me = CDec(txtDato(23).Text)
      !impoi2_me = CDec(txtDato(24).Text)
      !impoi3_me = CDec(txtDato(25).Text)
      !imptot_me = CDec(txtDato(26).Text)
      !ImpIGV_OGr_ME = CDec(txtDato(56).Text)
      !ImpIGV_OGN_ME = CDec(txtDato(57).Text)
      !ImpIGV_ONG_ME = CDec(txtDato(58).Text)
      
      ' Datos adicionales
      !tpoimpuesto = cboImpuesto.ListIndex
      !categoriadoc = cboCategoria.ListIndex
      
      ' Informacion pdb
      !indcprext = IIf(chkIndCprext.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      !codaduana = IIf(txtDetalle(0).Text = "", Null, txtDetalle(0).Text)
      !annodua = IIf(txtDetalle(1).Text = "", Null, txtDetalle(1).Text)
      !nrodua = IIf(txtDetalle(2).Text = "", Null, txtDetalle(2).Text)
      !indreten = IIf(chkIndreten.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      
'ini 2015-07-02 adic tabla detrac
      '!tsadetrac = IIf(cboDetraccion.ListIndex = 0, Null, Left(cboDetraccion.Text, 5))
         With frmTCprGrd.uorstcodetrac
            If .RecordCount > 0 Then .MoveFirst
            .Find "coddetrac='" & Left(cboDetraccion.Text, 5) & "'"
            If .EOF Then
               'MsgBox TEXT_8006, vbExclamation
               'ppAyuDet = True
            Else
               'lblLlaveDeta(tnIndex).Caption = " " & !razAux
               frmTCprGrd.uorstMain!tsadetrac = IIf(cboDetraccion.ListIndex = 0, Null, !coddetrac)
               frmTCprGrd.uorstMain!pctdetrac = IIf(cboDetraccion.ListIndex = 0, 0#, !pctdetrac)
           End If
         End With
'fin 2015-07-02 adic tabla detrac
      
      !codtdc_ref = IIf(txtDato(59).Text = "", Null, txtDato(59).Text)
      !serdoc_ref = IIf(txtDetalle(3).Text = "", Null, txtDetalle(3).Text)
      !nrodoc_ref = IIf(txtDetalle(4).Text = "", Null, txtDetalle(4).Text)
      dtpDetalle(0).Value = dtpDato(2).Value
      !feedoc_ref = dtpDetalle(0).Value
      !impbasref_mn = CDec(txtDetalle(5).Text)
      !impbasref_me = CDec(txtDetalle(7).Text)
      !impigvref_mn = CDec(txtDetalle(6).Text)
      !impigvref_me = CDec(txtDetalle(8).Text)
       
      '[Actualización de datos por Cuentas y Centros de Costo registrados directamente.
      ppAbreCtaCCo
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
        If cmdMas(dnContador).Tag = INDMASCTA_CTA And Val(txtDato(dnContador + DIFERENCIAMASCUENTA).Text) <> 1 Then
          With frmTCprGrd.uorstCOCprDocCta
            .MoveFirst
            .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & dnContador & ps_OrdenCta & "'"
            If Not .EOF Then
              .Delete
              .Update
              .Requery
              frmTCprGrd.uorstCOCprDocCCo.Requery
              Call upActualizaMas(dnContador, INDMASCTA_INI)
            End If
          End With
        End If
        
        If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
           cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Then
          With frmTCprGrd.uorstCOCprDocCta
            If .RecordCount <> 0 Then .MoveFirst
            .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & dnContador & ps_OrdenCta & "'"
            If .EOF Then
              .AddNew
              !codemp = gsCodEmp
              !pdoano = gsAnoAct
              !codaux = txtLlave(0).Text
              !codtdc = txtLlave(1).Text
              !serdoc = txtLlave(2).Text
              !nrodoc = txtLlave(3).Text
              !tpocnc = dnContador
              !orden = ps_OrdenCta
              !UsrCre = gsAbvUsr
              !FyHCre = Now
            Else
              !UsrMdf = gsAbvUsr
              !FyHMdf = Now
            End If
            '[ 20/01/2004 Miguel Angel Capturo la cuenta anterior si es modificacion
            txtDato(dnContador + DIFERENCIAMASCUENTA).Tag = IIf(pbNuevo, txtDato(dnContador + DIFERENCIAMASCUENTA).Text, !CodCta)
            ']
            !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
            !glodet = IIf(txtDato(Choose(gsIdioma, 3, 49)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 49)).Text)
            !glodetx = IIf(txtDato(Choose(gsIdioma, 49, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 49, 3)).Text)
            !impcta_mn = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE).Text)
            !impcta_me = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text)
            .Update
          End With
          If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Or _
             cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Then
            With frmTCprGrd.uorstCOCprDocCCo
              If .RecordCount <> 0 Then .MoveFirst
              .Find "cLlave1='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & dnContador & ps_OrdenCta & txtDato(dnContador + DIFERENCIAMASCUENTA).Text & "'"
              If .EOF Then
                .AddNew
                !codemp = gsCodEmp
                !pdoano = gsAnoAct
                !codaux = txtLlave(0).Text
                !codtdc = txtLlave(1).Text
                !serdoc = txtLlave(2).Text
                !nrodoc = txtLlave(3).Text
                !tpocnc = dnContador
                !orden = ps_OrdenCta
                !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
                !UsrCre = gsAbvUsr
                !FyHCre = Now
              Else
                !UsrMdf = gsAbvUsr
                !FyHMdf = Now
              End If
              !codcco = txtDato(dnContador + DIFERENCIAMASCCOSTO).Text
              !impcco_mn = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE).Text)
              !impcco_me = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text)
              .Update
            End With
          End If
          upActualizaMas dnContador, INDMASCTA_CTA
        End If
      Next
      ']
    Else
      'Llaves.
      txtLlave(0).Text = !codaux
      txtLlave(1).Text = !codtdc
      txtLlave(2).Text = !serdoc
      txtLlave(3).Text = !nrodoc

      'Datos.
      cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      chkIndCDt.Value = IIf(!indcdt = INDCDT_ACT, vbChecked, vbUnchecked)
      chkIndPreGen.Value = IIf(!indpregen = INDPREGEN_ACT, vbChecked, vbUnchecked)
      dtpDato(3).Value = !fehope
      dtpDato(0).Value = !feedoc
      dtpDato(1).Value = !fevdoc
      dtpDato(2).Value = !feedoc_ref
      dtpDato(5).Value = !ferdoc
      dtpDato(4).Value = !FehCDt
      txtDato(0).Text = IIf(IsNull(!coddro), "", !coddro)
      txtDato(1).Text = IIf(IsNull(!NroCpb), "", !NroCpb)
      txtDato(2).Text = IIf(IsNull(!refdoc), "", !refdoc)
      txtDato(Choose(gsIdioma, 3, 49)).Text = IIf(IsNull(!GloDoc), "", !GloDoc)
      txtDato(Choose(gsIdioma, 49, 3)).Text = IIf(IsNull(!glodocx), "", !glodocx)
      txtDato(50).Text = IIf(IsNull(!pdocpr), "", !pdocpr)
      txtDato(60).Text = IIf(IsNull(!codcon), "", !codcon)
      txtDato(51).Text = IIf(IsNull(!codasi), "", !codasi)
      txtDato(4).Text = Format(!ImpTCb, FORMATO_NUM_2)
      txtDato(52).Text = IIf(IsNull(!NroCDt), "", !NroCDt)
      ' Importes mn y me
      txtDato(5).Text = Format(!impogr_mn, FORMATO_NUM_1)
      txtDato(6).Text = Format(!ImpOGN_MN, FORMATO_NUM_1)
      txtDato(7).Text = Format(!ImpONG_MN, FORMATO_NUM_1)
      txtDato(8).Text = Format(!impexo_mn, FORMATO_NUM_1)
      txtDato(9).Text = Format(!impigv_mn, FORMATO_NUM_1)
      txtDato(10).Text = Format(!impisc_mn, FORMATO_NUM_1)
      txtDato(11).Text = Format(!impoim_mn, FORMATO_NUM_1)
      txtDato(12).Text = Format(!impoi1_mn, FORMATO_NUM_1)
      txtDato(13).Text = Format(!impoi2_mn, FORMATO_NUM_1)
      txtDato(14).Text = Format(!impoi3_mn, FORMATO_NUM_1)
      txtDato(15).Text = Format(!imptot_mn, FORMATO_NUM_1)
      txtDato(16).Text = Format(!impogr_me, FORMATO_NUM_1)
      txtDato(17).Text = Format(!ImpOGN_ME, FORMATO_NUM_1)
      txtDato(18).Text = Format(!ImpONG_ME, FORMATO_NUM_1)
      txtDato(19).Text = Format(!impexo_me, FORMATO_NUM_1)
      txtDato(20).Text = Format(!impigv_me, FORMATO_NUM_1)
      txtDato(21).Text = Format(!impisc_me, FORMATO_NUM_1)
      txtDato(22).Text = Format(!impoim_me, FORMATO_NUM_1)
      txtDato(23).Text = Format(!impoi1_me, FORMATO_NUM_1)
      txtDato(24).Text = Format(!impoi2_me, FORMATO_NUM_1)
      txtDato(25).Text = Format(!impoi3_me, FORMATO_NUM_1)
      txtDato(26).Text = Format(!imptot_me, FORMATO_NUM_1)
      txtDato(53).Text = Format(!ImpIGV_OGr_MN, FORMATO_NUM_1)
      txtDato(54).Text = Format(!ImpIGV_OGN_MN, FORMATO_NUM_1)
      txtDato(55).Text = Format(!ImpIGV_ONG_MN, FORMATO_NUM_1)
      txtDato(56).Text = Format(!ImpIGV_OGr_ME, FORMATO_NUM_1)
      txtDato(57).Text = Format(!ImpIGV_OGN_ME, FORMATO_NUM_1)
      txtDato(58).Text = Format(!ImpIGV_ONG_ME, FORMATO_NUM_1)
      For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
        txtDato(dnContador).Text = ""
        txtDato(dnContador).Tag = ""
      Next
      
      ' Datos adicionales
      cboImpuesto.ListIndex = !tpoimpuesto
      cboCategoria.ListIndex = !categoriadoc
      
      ' Informacion pdb
      chkIndCprext.Value = IIf(!indcprext = INDPREGEN_ACT, vbChecked, vbUnchecked)
      txtDetalle(0).Text = IIf(IsNull(!codaduana), "", !codaduana)
      txtDetalle(1).Text = IIf(IsNull(!annodua), "", !annodua)
      txtDetalle(2).Text = IIf(IsNull(!nrodua), "", !nrodua)
      chkIndreten.Value = IIf(!indreten = INDPREGEN_ACT, vbChecked, vbUnchecked)
      cboDetraccion.ListIndex = 0
      If Not IsNull(!tsadetrac) Then
        For dnContador = 1 To cboDetraccion.ListCount - 1
          If !tsadetrac = Left(cboDetraccion.List(dnContador), 5) Then
            cboDetraccion.ListIndex = dnContador
            Exit For
          End If
        Next dnContador
      End If
      txtDato(59).Text = IIf(IsNull(!codtdc_ref), "", !codtdc_ref)
      txtDetalle(3).Text = IIf(IsNull(!serdoc_ref), "", !serdoc_ref)
      txtDetalle(4).Text = IIf(IsNull(!nrodoc_ref), "", !nrodoc_ref)
      dtpDetalle(0).Value = IIf(IsNull(!feedoc_ref), !fehope, !feedoc_ref)
      txtDetalle(5).Text = Format(!impbasref_mn, FORMATO_NUM_1)
      txtDetalle(7).Text = Format(!impbasref_me, FORMATO_NUM_1)
      txtDetalle(6).Text = Format(!impigvref_mn, FORMATO_NUM_1)
      txtDetalle(8).Text = Format(!impigvref_me, FORMATO_NUM_1)
      ppAyuDet AYUDAT, 59
      
      '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      s_PedidoValiDsc = "S"
      ppAyuDet AYULLA, 0
      ppAyuDet AYULLA, 1
      ppAyuDet AYUDAT, 0
      For dnIndices = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
        ppAyuDet AYUDAT, dnIndices
      Next dnIndices
      ppAyuDet AYUDAT, 50
      ppAyuDet AYUDAT, 51
      ppAyuDet AYUDAT, 60
      ']
      
      '[Propio del formulario.
      ' Inicializo clon de pedido
      txtDato(50).Tag = txtDato(50).Text
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
        cmdMas(dnContador).Tag = Choose(dnContador, !indcta_ogr, !IndCta_OGN, !IndCta_ONG, !indcta_exo, !indcta_igv, !indcta_isc, !indcta_oim, !indcta_oi1, !indcta_oi2, !indcta_oi3, !indcta_tot)
      Next

      ppAbreCtaCCo
      With frmTCprGrd.uorstCOCprDocCta
        If .RecordCount > 0 Then
          For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
            If Val(txtDato(dnContador + DIFERENCIAMASIMPORTE).Text) <> 0 Then
              .MoveFirst
              .Find "TpoCnc = " & dnContador
              If Not .EOF Then
                txtDato(dnContador + DIFERENCIAMASCUENTA).Text = !CodCta
                txtDato(dnContador + DIFERENCIAMASCUENTA).Tag = !CodCta
                ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCUENTA
                With frmTCprGrd.uorstCOCprDocCCo
                  If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "cLlave = " & dnContador & frmTCprGrd.uorstCOCprDocCta!orden & frmTCprGrd.uorstCOCprDocCta!CodCta
                    If Not .EOF Then
                      txtDato(dnContador + DIFERENCIAMASCCOSTO).Text = !codcco
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
   dtpDato(5).Value = Date
'   optTpoMon(1).Value = True
   For dnContador = 0 To 3
      txtDato(dnContador).Text = ""
   Next
   txtDato(4).Text = Format(0, FORMATO_NUM_2)
   txtDato(49).Text = ""
   txtDato(50).Text = ""
   txtDato(51).Text = ""
   txtDato(60).Text = ""
   For dnContador = MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
      txtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
   Next
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      txtDato(dnContador).Text = ""
      txtDato(dnContador).Tag = ""
   Next
   txtDato(52).Text = ""
   ' Importes de distribucion de igv
   For dnContador = 53 To 58
      txtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
   Next
   ' Datos adicionales
   cboImpuesto.ListIndex = TipoImpuesto.Ninguno
   cboCategoria.ListIndex = CategoriaDocumento.Ninguno
   
   ' PDB
   chkIndCprext.Value = vbUnchecked
   txtDetalle(0).Text = ""
   txtDetalle(1).Text = ""
   txtDetalle(2).Text = ""
   chkIndreten.Value = vbUnchecked
   cboDetraccion.ListIndex = 0
   txtDato(59).Text = ""
   txtDetalle(3).Text = ""
   txtDetalle(4).Text = ""
   dtpDetalle(0).Value = Date
   txtDetalle(5).Text = Format(0, FORMATO_NUM_1)
   txtDetalle(7).Text = Format(0, FORMATO_NUM_1)
   txtDetalle(6).Text = Format(0, FORMATO_NUM_1)
   txtDetalle(8).Text = Format(0, FORMATO_NUM_1)
   lblDatoDeta(59).Caption = ""
   
  '[Propio del formulario.
  ' Inicializo clon de pedido
  s_PedidoValiDsc = "N"
  txtDato(50).Tag = txtDato(50).Text
  For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
    cmdMas(dnContador).Tag = INDMASCTA_INI
  Next
  ']

  'Ayudas.
  lblLlaveDeta(0).Caption = ""
  lblLlaveDeta(1).Caption = ""
  lblDatoDeta(0).Caption = ""
  lblDatoDeta(50).Caption = ""
  lblDatoDeta(51).Caption = ""
  lblDatoDeta(60).Caption = ""
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
  dtpDato(5).Enabled = tbHabilitar
  dtpDato(4).Enabled = (tbHabilitar And chkIndCDt.Value)
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next dnContador
  End With
  txtDato(52).Enabled = (tbHabilitar And chkIndCDt.Value)
  txtDato(51).Enabled = (tbHabilitar And pbNuevo)
  ' Inhabilito el total de IGV
  txtDato(9).Enabled = False
  txtDato(20).Enabled = False
  cmdMasIGV.Enabled = tbHabilitar
   
   For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      Call upHabilitaCuenta(False, dnContador)
   Next

  'Ayudas.
   cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar
   cmdDatoAyud(50).Enabled = tbHabilitar
   lblDatoDeta(50).Enabled = tbHabilitar
   cmdDatoAyud(51).Enabled = (tbHabilitar And pbNuevo)
   lblDatoDeta(51).Enabled = (tbHabilitar And pbNuevo)
   cmdDatoAyud(60).Enabled = tbHabilitar
   lblDatoDeta(60).Enabled = tbHabilitar

  ' Datos adicionales
  cboImpuesto.Enabled = tbHabilitar
  cboCategoria.Enabled = tbHabilitar
  
  ' PDB
   chkIndCprext.Enabled = tbHabilitar
   txtDetalle(0).Enabled = tbHabilitar
   txtDetalle(1).Enabled = tbHabilitar
   txtDetalle(2).Enabled = tbHabilitar
   chkIndreten.Enabled = tbHabilitar
   cboDetraccion.Enabled = tbHabilitar
   txtDato(59).Enabled = tbHabilitar
   txtDetalle(3).Enabled = tbHabilitar
   txtDetalle(4).Enabled = tbHabilitar
   dtpDetalle(0).Enabled = tbHabilitar
   txtDetalle(5).Enabled = tbHabilitar
   txtDetalle(6).Enabled = tbHabilitar
   txtDetalle(7).Enabled = tbHabilitar
   txtDetalle(8).Enabled = tbHabilitar
   cmdDatoAyud(59).Enabled = tbHabilitar
   lblDatoDeta(59).Enabled = tbHabilitar

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
Private Sub cboTpoMon_Click()
  unVerMonNac = IIf(chkMonedaActiva.Value, cboTpoMon.ListIndex, IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_IND, TPOMON_NAC_IND))
  ppCambioTpoMon
End Sub

Private Sub chkDesactivar_Click()
  Dim dnContador As Integer
  
  For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
    ' Inabilito cuentas por asiento
    If cmdMas(dnContador).Tag = INDMASCTA_CTA Then
      cmdMas(dnContador).Enabled = (cmdMas(dnContador).Enabled = True)
    Else
      cmdMas(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
    End If
  Next
  
  For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
    '///Angel 18/12/2003
    '///Se agrego condicion e instruccion ELSE, solo se dejo instruccion else
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
  txtDato(51).Text = IIf((pbNuevo And chkDesactivar.Value = vbUnchecked), "", txtDato(51).Text)
  lblDatoDeta(51).Caption = IIf((pbNuevo And chkDesactivar.Value = vbUnchecked), "", lblDatoDeta(51).Caption)
  
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
    If Month(dtpDato(0).Value) > Val(gsMesAct) And Year(dtpDato(0).Value) >= Val(gsAnoAct) Then
      MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
      dtpDato(Index).SetFocus
      Cancel = True
      Exit Sub
    End If
  End If
  
  If Index = 2 Then
    If Month(dtpDato(2).Value) > Val(gsMesAct) And Year(dtpDato(2).Value) >= Val(gsAnoAct) Then
      MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
      dtpDato(Index).SetFocus
      Cancel = True
      Exit Sub
    End If
    dtpDato(2).Tag = 0
    If (dtpDato(2).Tag <> dtpDato(2).Value) Then
      dtpDato(2).Tag = dtpDato(2).Value
      With frmTCprGrd.uorstTGTCb
        If .RecordCount <> 0 Then .MoveLast: .MoveFirst
        .Find "FehTCb = '" & dtpDato(2).Value & "'"
        If .EOF Then
          MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
          Cancel = True
          Exit Sub
        Else
          txtDato(4).Text = Format(!ImpTCb_Vta, FORMATO_NUM_2)
        End If
      End With
    End If
  End If
  
  If Index = 3 Then
    If Month(dtpDato(3).Value) <> Val(gsMesAct) Or Year(dtpDato(3).Value) <> Val(gsAnoAct) Then
      MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
      dtpDato(Index).SetFocus
      Cancel = True
      Exit Sub
    End If
  
    If pbFecha Then
      If txtDato(4).Text = 0 Then
        With dtpDato
          For dnContador = 0 To .Count - 3
            .Item(dnContador).Value = dtpDato(Index).Value
          Next
          dtpDato(5).Value = dtpDato(Index).Value
        End With
        pbFecha = False
      Else
        With dtpDato
          For dnContador = 1 To .Count - 3
            .Item(dnContador).Value = dtpDato(Index).Value
          Next
          dtpDato(5).Value = dtpDato(Index).Value
        End With
        pbFecha = False
      End If
    End If
  End If
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
  If PreviousTab = 0 And sstMain.Tab = 1 Then
    ppDatosWhere
  End If
  'dgrDetalle.SetFocus
End Sub

Private Sub ppAbreCtaCCo()
  With frmTCprGrd.uorstCOCprDocCCo
    frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.codemp='" & frmTCprGrd.uorstMain!codemp & "' AND COCprDocCCo.pdoano='" & frmTCprGrd.uorstMain!pdoano & "' "
    frmTCprGrd.usConnStrgWher_COCprDocCCo = frmTCprGrd.usConnStrgWher_COCprDocCCo & "AND COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!codaux & "' AND COCprDocCCo.CodTDc='" & frmTCprGrd.uorstMain!codtdc & "' "
    frmTCprGrd.usConnStrgWher_COCprDocCCo = frmTCprGrd.usConnStrgWher_COCprDocCCo & "AND COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!serdoc & "' AND COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!nrodoc & "' "
    If .State = adStateOpen Then .Close
    .Source = frmTCprGrd.usConnStrgSele_COCprDocCCo & frmTCprGrd.usConnStrgWher_COCprDocCCo & frmTCprGrd.usConnStrgOrde_COCprDocCCo
    .Open
    .Properties("Unique Table").Value = "COCprDocCCo"
  End With
  With frmTCprGrd.uorstCOCprDocCta
    frmTCprGrd.usConnStrgWher_COCprDocCta = "WHERE COCprDocCta.codemp='" & frmTCprGrd.uorstMain!codemp & "' AND COCprDocCta.pdoano='" & frmTCprGrd.uorstMain!pdoano & "' "
    frmTCprGrd.usConnStrgWher_COCprDocCta = frmTCprGrd.usConnStrgWher_COCprDocCta & "AND COCprDocCta.CodAux='" & frmTCprGrd.uorstMain!codaux & "' AND COCprDocCta.CodTDc='" & frmTCprGrd.uorstMain!codtdc & "' "
    frmTCprGrd.usConnStrgWher_COCprDocCta = frmTCprGrd.usConnStrgWher_COCprDocCta & "AND COCprDocCta.SerDoc='" & frmTCprGrd.uorstMain!serdoc & "' AND COCprDocCta.NroDoc='" & frmTCprGrd.uorstMain!nrodoc & "' "
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
    txtDato(dnContador).Visible = (unVerMonNac = TPOMON_NAC_IND)
  Next
  For dnContador = MINIMOINDICEIMPORTEME To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
    txtDato(dnContador).Visible = Not (unVerMonNac = TPOMON_NAC_IND)
  Next
  ' PDB FOB, documento referencia
  txtDetalle(5).Visible = (unVerMonNac = TPOMON_NAC_IND)
  txtDetalle(6).Visible = (unVerMonNac = TPOMON_NAC_IND)
  txtDetalle(7).Visible = Not (unVerMonNac = TPOMON_NAC_IND)
  txtDetalle(8).Visible = Not (unVerMonNac = TPOMON_NAC_IND)
End Sub

Private Sub ppGenera()
  Dim dnContador As Integer
  Dim dnNumeroItem As Integer
  Dim dbProcesaCuenta As Boolean
  Dim sSentencia As String
  Dim siexiste As Boolean
  Dim masdedos As Boolean
  Dim cuenta As Integer
  
  siexiste = False
  masdedos = False
  
  'Aqui Esta
  sSentencia = "SELECT coddro,nrocpb"
  sSentencia = sSentencia & " FROM cocpbcab  "
  sSentencia = sSentencia & " WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & " AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & " AND mespvs='" & gsMesAct & "'"
  sSentencia = sSentencia & " AND coddro='" & txtDato(0).Text & "'"
  sSentencia = sSentencia & " AND nrocpb='" & txtDato(1).Text & "'"
  Set frmTCprGrd.uorstTemporal = frmTCprGrd.uocnnMain.Execute(sSentencia)
  If Not (frmTCprGrd.uorstTemporal.BOF Or frmTCprGrd.uorstTemporal.EOF) And frmTCprGrd.uorstTemporal.RecordCount > 0 Then
    While Not frmTCprGrd.uorstTemporal.EOF
            siexiste = True
            frmTCprGrd.uorstTemporal.MoveNext
    Wend
  Else
  End If
  frmTCprGrd.uorstTemporal.Close
  
  cuenta = 0
  sSentencia = "SELECT coddro,nrocpb "
  sSentencia = sSentencia & " FROM cocprdoc  "
  sSentencia = sSentencia & " WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & " AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & " AND mespvs='" & gsMesAct & "'"
  sSentencia = sSentencia & " AND coddro='" & txtDato(0).Text & "'"
  sSentencia = sSentencia & " AND nrocpb='" & txtDato(1).Text & "'"
  Set frmTCprGrd.uorstTemporal = frmTCprGrd.uocnnMain.Execute(sSentencia)
  If Not (frmTCprGrd.uorstTemporal.BOF Or frmTCprGrd.uorstTemporal.EOF) And frmTCprGrd.uorstTemporal.RecordCount > 0 Then
    While Not frmTCprGrd.uorstTemporal.EOF
            cuenta = cuenta + 1
            frmTCprGrd.uorstTemporal.MoveNext
    Wend
  Else
  End If
  frmTCprGrd.uorstTemporal.Close
  
  If cuenta >= 2 Then masdedos = True
  
  If txtDato(1).Text <> "" Then
    ppDatosWhere
        If masdedos = False Then
            With frmTCprGrd.uorstCOCpbCab
                'Si existe, elimina Comprobante existente.
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "cLlave='" & txtDato(0).Text & txtDato(1).Text & "'"
                    If Not .EOF Then .Delete
                End If
            End With
        End If
  End If

  ' MA 26-08-2011 / Tipo documento clave
  frmTCprGrd.uorstTGTDc.MoveFirst
  frmTCprGrd.uorstTGTDc.Find "codtdc='" & txtLlave(1).Text & "'"

  With frmTCprGrd.uorstCOCpbCab
    'Si no está marcado para generar, marca el documento como no generado.
    If chkIndPreGen.Value = vbUnchecked Then
     
      frmTCprGrd.uorstMain!indgen = False
      frmTCprGrd.uorstMain.Update
      Exit Sub
    End If
    
    ' Captura del siguiente numero de comprobante
    If txtDato(1).Text = "" Then
      txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
      txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
      txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
      txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
      txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
      frmTCprGrd.uocnnMain.Execute txtDato(1).Tag
      ' Actualizo numero de comprobante tabla de detalle
      frmTCprGrd.uorstMain!NroCpb = txtDato(1).Text
      frmTCprGrd.uorstMain.Update
    Else
    If masdedos = True Then
      txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
      txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
      txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
      txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
      txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
      frmTCprGrd.uocnnMain.Execute txtDato(1).Tag
      ' Actualizo numero de comprobante tabla de detalle
      frmTCprGrd.uorstMain!NroCpb = txtDato(1).Text
      frmTCprGrd.uorstMain.Update
    End If
    End If
    
    ppDatosWhere
   
    'Si no hay cuentas, marca el documento como no generado.
    If frmTCprGrd.uorstCOCprDocCta.RecordCount = 0 Then
      frmTCprGrd.uorstMain!indgen = False
      frmTCprGrd.uorstMain.Update
      Exit Sub
    End If

    'Crea encabezado de Comprobante.
    .AddNew
    !codemp = gsCodEmp
    !pdoano = gsAnoAct
    !mespvs = gsMesAct
    !coddro = txtDato(0).Text
    !NroCpb = txtDato(1).Text
    !FehCpb = dtpDato(3).Value
    !tpognr = TPOGNR_CPR
    !IndNCu = INDNCU_FAL
    !glocpb = IIf(txtDato(Choose(gsIdioma, 3, 49)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 49)).Text)
    !glocpbx = IIf(txtDato(Choose(gsIdioma, 49, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 49, 3)).Text)
    !UsrCre = gsAbvUsr
    !FyHCre = Now
    'Angel 15/12/2003
  End With

  '[ Teo, Miguel Angel Refresco los recordset de cuentas y centros de costos
  frmTCprGrd.uorstCOCprDocCta.Requery
  frmTCprGrd.uorstCOCprDocCCo.Requery
  ']
  With frmTCprGrd.uorstCOCprDocCta
    'Crea ítemes de Comprobante.
    .MoveFirst
    Do
      dbProcesaCuenta = True
      'Itemes con Centro de Costo.
      If Val(!tpocnc) <= CUENTASCONCCOSTO Then
        With frmTCprGrd.uorstCOCprDocCCo
          If .RecordCount <> 0 Then
            .MoveFirst
            .Find "cLlave = " & Trim(frmTCprGrd.uorstCOCprDocCta!tpocnc) & frmTCprGrd.uorstCOCprDocCta!orden & frmTCprGrd.uorstCOCprDocCta!CodCta
            If Not .EOF Then
              Do
                dnNumeroItem = dnNumeroItem + 1
                ppGenera1 True, dnNumeroItem, IIf(CInt(frmTCprGrd.uorstCOCprDocCta!tpocnc) >= 5, "", txtDato(60).Text)
                .MoveNext
                If .EOF Then Exit Do
                If !cLlave <> Trim(frmTCprGrd.uorstCOCprDocCta!tpocnc) & frmTCprGrd.uorstCOCprDocCta!orden & frmTCprGrd.uorstCOCprDocCta!CodCta Then Exit Do
              Loop
              dbProcesaCuenta = False
            End If
          End If
        End With
      End If
      
      'Itemes sin Centro de Costo.
      If dbProcesaCuenta Then
        dnNumeroItem = dnNumeroItem + 1
        ppGenera1 False, dnNumeroItem, IIf(CInt(frmTCprGrd.uorstCOCprDocCta!tpocnc) >= 5, "", txtDato(60).Text)
      End If
      .MoveNext
    Loop Until .EOF
  End With

  frmTCprGrd.uorstMain!indgen = True
  txtDato(0).Enabled = False
  txtDato(1).Enabled = False
  cmdDatoAyud(0).Enabled = False
  lblDatoDeta(0).Enabled = False
  
  frmTCprGrd.uorstCOCpbCab.Update
  frmTCprGrd.uorstCOCpbDet.UpdateBatch
  frmTCprGrd.uorstMain.Update

End Sub

Private Sub ppGenera1(tbCCosto As Boolean, tnNumeroItem As Integer, ByVal sContrato As String)
  
  With frmTCprGrd.uorstCOCpbDet
    .AddNew
    !codemp = gsCodEmp
    !pdoano = gsAnoAct
    !coddro = txtDato(0).Text
    !NroCpb = txtDato(1).Text
    !NroIte = tnNumeroItem
    !mespvs = gsMesAct
    !CodCta = frmTCprGrd.uorstCOCprDocCta!CodCta
    !fehope = dtpDato(3).Value
    ' Busco en el plan de cuentas
    frmTCprGrd.uorstCoCta.MoveFirst
    frmTCprGrd.uorstCoCta.Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "'"
    If frmTCprGrd.uorstCoCta!indcco = INDCCO_ACT Then If tbCCosto Then !codcco = frmTCprGrd.uorstCOCprDocCCo!codcco
    If frmTCprGrd.uorstCoCta!IndDoc = INDDOC_ACT Then
      !codaux = txtLlave(0).Text
    Else
      If Len(Trim(frmTCprGrd.uorstCOCprDocCta!codruc)) > 0 Then
        !codaux = frmTCprGrd.uorstCOCprDocCta!codruc
      Else
        !codaux = txtLlave(0).Text
      End If
    End If
    !codtdc = txtLlave(1).Text
    !serdoc = txtLlave(2).Text
    !nrodoc = txtLlave(3).Text
    !feedoc = dtpDato(0).Value
    !fevdoc = dtpDato(1).Value
    !ferdoc = dtpDato(5).Value
    !refdoc = txtDato(2).Text
    !pdocpr = IIf(txtDato(50).Text = "", Null, txtDato(50).Text)
    !codcon = IIf(sContrato = "", Null, sContrato)
    !GloIte = frmTCprGrd.uorstCOCprDocCta!glodet
    !GloItex = frmTCprGrd.uorstCOCprDocCta!glodetx
    If tbCCosto Then
      If (frmTCprGrd.uorstCOCprDocCCo!impcco_me > 0) Or (frmTCprGrd.uorstCOCprDocCCo!impcco_mn > 0) Then
        !TpoCtb = IIf(frmTCprGrd.uorstCOCprDocCta!tpocnc = TPOCNC_TOT_CPR, IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_HAB, TPOCTB_DEB), IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_HAB, TPOCTB_DEB))
      Else
        !TpoCtb = IIf(frmTCprGrd.uorstCOCprDocCta!tpocnc = TPOCNC_TOT_CPR, IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_HAB, TPOCTB_DEB), IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_HAB, TPOCTB_DEB))
      End If
    Else
      If (frmTCprGrd.uorstCOCprDocCta!impcta_me > 0) Or (frmTCprGrd.uorstCOCprDocCta!impcta_mn > 0) Then
        !TpoCtb = IIf(frmTCprGrd.uorstCOCprDocCta!tpocnc = TPOCNC_TOT_CPR, IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_HAB, TPOCTB_DEB), IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_HAB, TPOCTB_DEB))
      Else
        !TpoCtb = IIf(frmTCprGrd.uorstCOCprDocCta!tpocnc = TPOCNC_TOT_CPR, IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_HAB, TPOCTB_DEB), IIf(frmTCprGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_HAB, TPOCTB_DEB))
      End If
    End If
    !tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
    !ImpTCb = CDec(txtDato(4).Text)
    If tbCCosto Then
      !ImpMN = CDec(Abs(frmTCprGrd.uorstCOCprDocCCo!impcco_mn))
      !ImpME = CDec(Abs(frmTCprGrd.uorstCOCprDocCCo!impcco_me))
    Else
      !ImpMN = CDec(Abs(frmTCprGrd.uorstCOCprDocCta!impcta_mn))
      !ImpME = CDec(Abs(frmTCprGrd.uorstCOCprDocCta!impcta_me))
    End If
    !tpognr = TPOGNR_CPR
    !UsrCre = gsAbvUsr
    !FyHCre = Now
  End With

End Sub

Private Sub ppInsDelCtaCos(ByVal s_Asiento As String, ByVal n_TipoTran As Integer)
  Dim sSentencia As String, sTpoCnc As String, sOrden As String
  Dim nOrden As Long
  Dim nImporteMN As Double, nImporteME As Double
  Dim nImpoCtaMN As Double, nImpoCtaME As Double
  Dim nImpoCCoMN As Double, nImpoCCoME As Double
  Dim nPorcentaje As Double
  
  ' Refresco las cuentas y centro de costos
  If n_TipoTran = INDCCO_ACT Then
    sSentencia = "SELECT " & Choose(cboTpoMon.ListIndex + 1, "a.codcta_mn", "a.codcta_me") & " AS codcta, "
    sSentencia = sSentencia & "a.tpocnc, a.orden, a.codcco, a.pordst, b.inddoc, b.indcco "
    sSentencia = sSentencia & "FROM coasidet a "
    sSentencia = sSentencia & "LEFT JOIN cocta b ON a.codemp=b.codemp AND a.pdoano=b.pdoano AND " & Choose(cboTpoMon.ListIndex + 1, "a.codcta_mn", "a.codcta_me") & "=b.codcta "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND a.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND a.codasi='" & s_Asiento & "'"
    sSentencia = sSentencia & "ORDER BY a.tpocnc, a.orden"
    Set frmTCprGrd.uorstTemporal = frmTCprGrd.uocnnMain.Execute(sSentencia)
    If Not (frmTCprGrd.uorstTemporal.BOF Or frmTCprGrd.uorstTemporal.EOF) And frmTCprGrd.uorstTemporal.RecordCount > 0 Then
      While Not frmTCprGrd.uorstTemporal.EOF
        ' Inicializo el orden de cuenta
        nOrden = IIf(sTpoCnc = frmTCprGrd.uorstTemporal!tpocnc, nOrden, 0)
        sTpoCnc = frmTCprGrd.uorstTemporal!tpocnc
        nImporteMN = CDec(txtDato(Val(sTpoCnc) + MINIMOINDICEIMPORTEMN - 1).Text)
        nImporteME = CDec(txtDato(Val(sTpoCnc) + MINIMOINDICEIMPORTEME - 1).Text)
        ' Inserto las cuentas por compra
        If nImporteMN <> 0 Or nImporteME <> 0 Then
          nOrden = nOrden + 1
          nImpoCtaMN = Round(nImporteMN * (CDec(frmTCprGrd.uorstTemporal!pordst) / 100), 2)
          nImpoCtaME = Round(nImporteME * (CDec(frmTCprGrd.uorstTemporal!pordst) / 100), 2)
          With frmTCprGrd.uorstCOCprDocCta    'Cambiar RecordSet.
            .AddNew
            'Llaves.
            !codemp = gsCodEmp
            !pdoano = gsAnoAct
            !codaux = txtLlave(0).Text
            !codtdc = txtLlave(1).Text
            !serdoc = txtLlave(2).Text
            !nrodoc = txtLlave(3).Text
            !tpocnc = sTpoCnc
            !orden = Format(nOrden, "00")
            'Datos.
            !CodCta = frmTCprGrd.uorstTemporal!CodCta
            !codruc = IIf(frmTCprGrd.uorstTemporal!IndDoc = INDDOC_ACT, txtLlave(0).Text, Null)
            !glodet = IIf(txtDato(Choose(gsIdioma, 3, 49)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 49)).Text)
            !glodetx = IIf(txtDato(Choose(gsIdioma, 49, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 49, 3)).Text)
            !impcta_mn = CDec(nImpoCtaMN)
            !impcta_me = CDec(nImpoCtaME)
            !UsrCre = gsAbvUsr
            !FyHCre = Now
            .Update
          End With
          cmdMas(Val(sTpoCnc)).Tag = INDMASCTA_MAS
          ' Cuenta incial actualiza texto de costos
          If nOrden = 1 Then
            txtDato(Val(sTpoCnc) + (MINIMOINDICECUENTA - 1)).Text = frmTCprGrd.uorstTemporal!CodCta
            cmdMas(Val(sTpoCnc)).Tag = INDMASCTA_INI
          End If
          upActualizaMas Val(sTpoCnc), cmdMas(Val(sTpoCnc)).Tag
          
          ' Inserto el centro de costo
          If ((Not IsNull(frmTCprGrd.uorstTemporal!codcco)) And frmTCprGrd.uorstTemporal!indcco = INDCCO_ACT) Then
            With frmTCprGrd.uorstCOCprDocCCo    'Cambiar RecordSet.
              .AddNew
              'Llaves.
              !codemp = gsCodEmp
              !pdoano = gsAnoAct
              !codaux = txtLlave(0).Text
              !codtdc = txtLlave(1).Text
              !serdoc = txtLlave(2).Text
              !nrodoc = txtLlave(3).Text
              !tpocnc = sTpoCnc
              !orden = Format(nOrden, "00")
              !CodCta = frmTCprGrd.uorstTemporal!CodCta
              'Datos.
              !codcco = frmTCprGrd.uorstTemporal!codcco
              !impcco_mn = CDec(nImpoCtaMN)
              !impcco_me = CDec(nImpoCtaME)
              !UsrCre = gsAbvUsr
              !FyHCre = Now
              .Update
            End With
            ' Cuenta incial actualiza texto de costos
            If nOrden = 1 Then
              txtDato(Val(sTpoCnc) + (MINIMOINDICECCOSTO - 1)).Text = frmTCprGrd.uorstTemporal!codcco
            End If
          End If
        End If
        frmTCprGrd.uorstTemporal.MoveNext
      Wend
    End If
    frmTCprGrd.uorstTemporal.Close
  ElseIf n_TipoTran = INDCCO_INA Then
    sSentencia = "DELETE FROM COCprDocCta "
    sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND codaux='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND codtdc='" & txtLlave(1).Text & "' "
    sSentencia = sSentencia & "AND serdoc='" & txtLlave(2).Text & "' "
    sSentencia = sSentencia & "AND nrodoc='" & txtLlave(3).Text & "'"
    frmTCprGrd.uocnnMain.Execute sSentencia
  End If
  ppAbreCtaCCo

End Sub

Public Sub upActualizaMas(pnIndice As Byte, pnValor As Byte)
   frmTCpr.cmdMas(pnIndice).Tag = pnValor 'Necesaria la referencia por ser llamado externamente.
   With frmTCprGrd.uorstMain
      Select Case pnIndice
      Case 1
         !indcta_ogr = pnValor
      Case 2
         !IndCta_OGN = pnValor
      Case 3
         !IndCta_ONG = pnValor
      Case 4
         !indcta_exo = pnValor
      Case 5
         !indcta_igv = pnValor
      Case 6
         !indcta_isc = pnValor
      Case 7
         !indcta_oim = pnValor
      Case 8
         !indcta_oi1 = pnValor
      Case 9
         !indcta_oi2 = pnValor
      Case 10
         !indcta_oi3 = pnValor
      Case 11
         !indcta_tot = pnValor
      End Select
   End With
End Sub

Public Sub upHabilitaCuenta(tbHabilita As Boolean, tnIndice As Byte)
  txtDato(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
  lblDatoDeta(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
  cmdDatoAyud(tnIndice + DIFERENCIAMASCUENTA).Enabled = tbHabilita
  If tnIndice <= CUENTASCONCCOSTO Then
    upHabilitaCCosto tbHabilita, tnIndice
  End If
End Sub

Public Sub upHabilitaCCosto(tbHabilita As Boolean, tnIndice As Byte)
  If Not tbHabilita Or txtDato(tnIndice + DIFERENCIAMASCUENTA).Text = "" Then
    txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
    lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
    cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = False
  Else
    frmTCprGrd.uorstCoCta.MoveFirst
    frmTCprGrd.uorstCoCta.Find "CodCta='" & txtDato(tnIndice + DIFERENCIAMASCUENTA).Text & "'"
    txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTCprGrd.uorstCoCta!indcco = INDCCO_ACT)
    lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTCprGrd.uorstCoCta!indcco = INDCCO_ACT)
    cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTCprGrd.uorstCoCta!indcco = INDCCO_ACT)
  End If
End Sub

Private Sub ppDatosWhere()             'Cambiar.
  With frmTCprGrd
    .uorstCOCpbCab.Requery
    .usConnStrgWher_COCpbDet = "WHERE COCpbDet.codemp='" & gsCodEmp & "' AND COCpbDet.pdoano='" & gsAnoAct & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='" & txtDato(0).Text & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND COCpbDet.NroCpb='" & txtDato(1).Text & "' "
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
        .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
        .Item(dnNum).Width = 900
       Case 1
        .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Axiliary")
        .Item(dnNum).Width = 1150
       Case 2
        .Item(dnNum).Caption = Choose(gsIdioma, "C.Cto.", "Cost Center")
        .Item(dnNum).Width = 600
       Case 3
        .Item(dnNum).Caption = Choose(gsIdioma, "Glosa", "Gloss")
        .Item(dnNum).Width = 1500
       Case 4
        .Item(dnNum).Caption = Choose(gsIdioma, "Debe ", "Debit ") & TPOMON_NAC_TXT_0
        .Item(dnNum).Width = 1000
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 5
        .Item(dnNum).Caption = Choose(gsIdioma, "Haber ", "Credit ") & TPOMON_NAC_TXT_0
        .Item(dnNum).Width = 1000
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 6
        .Item(dnNum).Caption = Choose(gsIdioma, "Debe ", "Debit ") & TPOMON_EXT_TXT_0
        .Item(dnNum).Width = 1000
        .Item(dnNum).Alignment = dbgRight
        .Item(dnNum).NumberFormat = FORMATO_NUM_1 & " "
       Case 7
        .Item(dnNum).Caption = Choose(gsIdioma, "Haber ", "Credit ") & TPOMON_EXT_TXT_0
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
    Next dnNum
  End With
End Sub

Private Function ValidaCtasCCo() As Boolean
  Dim dnContador, dnIndCCo As Byte
  Dim dvRegistroActual As Variant
  Dim dnTotalCuentaMN, dnTotalCuentaME, dnTotalImporteMN, dnTotalImporteME As Double
    
  ValidaCtasCCo = True
  For dnContador = INDMASCTA_INI To CANTIDADIMPORTES - 1
    If ((CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) And _
       (cmdMas(dnContador + 1).Tag = INDMASCTA_MAS And Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
      ValidaCtasCCo = Not (txtDato(MINIMOINDICECUENTA + dnContador).Text = "")
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
              If Trim(!tpocnc) = Trim(Str(dnContador + 1)) Then
                dnTotalCuentaMN = dnTotalCuentaMN + !impcta_mn
                dnTotalCuentaME = dnTotalCuentaME + !impcta_me
                With frmTCprGrd.uorstCoCta
                  .MoveFirst
                  .Find "CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "'"
                  If Not .EOF Then
                    dnIndCCo = frmTCprGrd.uorstCoCta!indcco
                  End If
                End With
              End If
              If dnIndCCo = INDCCO_ACT Then
                With frmTCprGrd.uorstCOCprDocCCo
                  If .State = adStateOpen Then .Close
                  frmTCprGrd.usConnStrgWher_COCprDocCCo = "WHERE COCprDocCCo.CodAux='" & frmTCprGrd.uorstMain!codaux & "' And COCprDocCCo.SerDoc='" & frmTCprGrd.uorstMain!serdoc & "' And COCprDocCCo.NroDoc='" & frmTCprGrd.uorstMain!nrodoc & "' And COCprDocCCo.TpoCnc='" & Trim(Str(dnContador + 1)) & "' And COCprDocCCo.CodCta='" & frmTCprGrd.uorstCOCprDocCta!CodCta & "' "
                  .Source = frmTCprGrd.usConnStrgSele_COCprDocCCo & frmTCprGrd.usConnStrgWher_COCprDocCCo & frmTCprGrd.usConnStrgOrde_COCprDocCCo
                  .Open
                  If .RecordCount = 0 Then
                    MsgBox Choose(gsIdioma, "Cuenta ", "Account ") & frmTCprGrd.uorstCOCprDocCta!CodCta & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
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
      With frmTCprGrd.uorstCoCta
        .MoveFirst
        .Find "CodCta='" & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & "'"
        If Not .EOF Then
          dnIndCCo = frmTCprGrd.uorstCoCta!indcco
        End If
      End With
      If dnIndCCo = INDCCO_ACT And txtDato(dnContador + 1 + DIFERENCIAMASCUENTA + CUENTASCONCCOSTO).Text = "" Then
        MsgBox Choose(gsIdioma, "Cuenta ", "Account") & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
        ValidaCtasCCo = False
        Exit Function
      End If
    ElseIf Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) = 0 And ((CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) Then
        ValidaCtasCCo = False
    End If
  Next dnContador
End Function

Private Function pfValidaPedido(ByVal sPedido As String, ByVal sModificar As String) As Boolean
  Dim pnImporteMN As Double, pnImporteME As Double
  Dim pnImporte As Double, pnImporDiferen As Double
  Dim sSentencia As String, sMoneda As String
  Dim sCuenta As String, sCenCosto As String, sDetalle As String
  Dim sDetaCuenta As String, sDetaCenCosto As String, sMensage As String
  Dim sRegistro As String, nOrden As Long
    
  pfValidaPedido = True
  sCuenta = "": sCenCosto = "": sDetalle = "": sMoneda = TPOMON_NAC
  sDetaCuenta = "": sDetaCenCosto = ""
  pnImporteMN = 0: pnImporteME = 0
  pnImporte = 0: pnImporDiferen = 0
  
  ' Descripcion del pedido
  If s_PedidoValiDsc = "S" Then
    pfValidaPedido = False
    With frmTCprGrd                  'Cambiar Formulario de Grid.
      sSentencia = "SELECT " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS detpdo "
      sSentencia = sSentencia & "FROM copdocpr a "
      sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & gsAnoAct & gsMesAct & "' "
      sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(a.coddpe, a.pdocpr)", "(a.coddpe+a.pdocpr)") & "='" & sPedido & "' "
      sSentencia = sSentencia & "AND a.codaux='" & txtLlave(0).Text & "' "
      If ps_Plataforma = pSrvMySql Then
        sSentencia = sSentencia & "AND a.fehpdo<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
      ElseIf ps_Plataforma = pSrvSql Then
        sSentencia = sSentencia & "AND a.fehpdo<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
      End If
      Set .uorstTemporal = .uocnnMain.Execute(sSentencia)
      If Not (.uorstTemporal.BOF Or .uorstTemporal.EOF) And .uorstTemporal.RecordCount > 0 Then
        sDetalle = IIf(IsNull(.uorstTemporal!detpdo), "", .uorstTemporal!detpdo)
        pfValidaPedido = True
      End If
      .uorstTemporal.Close
    End With
    lblDatoDeta(50).Caption = " " & sDetalle
    Exit Function
  End If
  ' Movimientos por compras
  With frmTCprGrd                  'Cambiar Formulario de Grid.
    sSentencia = "SELECT p.codaux, p.coddpe, p.pdocpr, "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "ROUND(SUM("
      sSentencia = sSentencia & "((IFNULL(a.impogr_mn, 0)+IFNULL(a.impogn_mn, 0)+IFNULL(a.impong_mn, 0)+IFNULL(a.impexo_mn, 0))*"
      sSentencia = sSentencia & "(CASE b.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+ "
      sSentencia = sSentencia & "(IFNULL(d.impbru_mn, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      sSentencia = sSentencia & "), 2) AS impcpr_mn, "
      sSentencia = sSentencia & "ROUND(SUM("
      sSentencia = sSentencia & "((IFNULL(a.impogr_me, 0)+IFNULL(a.impogn_me, 0)+IFNULL(a.impong_me, 0)+IFNULL(a.impexo_me, 0))*"
      sSentencia = sSentencia & "(CASE b.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+ "
      sSentencia = sSentencia & "(IFNULL(d.impbru_me, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      sSentencia = sSentencia & "), 2) AS impcpr_me "
    Else
      sSentencia = sSentencia & "ROUND(SUM("
      sSentencia = sSentencia & "((ISNULL(a.impogr_mn, 0)+ISNULL(a.impogn_mn, 0)+ISNULL(a.impong_mn, 0)+ISNULL(a.impexo_mn, 0))*"
      sSentencia = sSentencia & "(CASE b.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+ "
      sSentencia = sSentencia & "(ISNULL(d.impbru_mn, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      sSentencia = sSentencia & "), 2) AS impcpr_mn, "
      sSentencia = sSentencia & "ROUND(SUM("
      sSentencia = sSentencia & "((ISNULL(a.impogr_me, 0)+ISNULL(a.impogn_me, 0)+ISNULL(a.impong_me, 0)+ISNULL(a.impexo_me, 0))*"
      sSentencia = sSentencia & "(CASE b.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+ "
      sSentencia = sSentencia & "(ISNULL(d.impbru_me, 0)*(CASE e.SgnTDc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))"
      sSentencia = sSentencia & "), 2) AS impcpr_me "
    End If
    sSentencia = sSentencia & "FROM ((((copdocpr p "
    sSentencia = sSentencia & "LEFT JOIN cocprdoc a ON p.codemp=a.codemp AND p.codaux=a.codaux AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(p.coddpe, p.pdocpr)", "(p.coddpe+p.pdocpr)") & "=a.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND Concat(a.pdoano, a.codtdc, a.serdoc, a.nrodoc)<>'" & gsAnoAct & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "' "
      sSentencia = sSentencia & "AND a.feedoc<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "') "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND (a.pdoano+a.codtdc+a.serdoc+a.nrodoc)<>'" & gsAnoAct & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "' "
      sSentencia = sSentencia & "AND a.feedoc<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103)) "
    End If
    sSentencia = sSentencia & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    sSentencia = sSentencia & "LEFT JOIN cohprdoc d ON p.codemp=d.codemp AND p.codaux=d.codaux AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(p.coddpe, p.pdocpr)", "(p.coddpe+p.pdocpr)") & "=d.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND d.feedoc<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "') "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND d.feedoc<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103)) "
    End If
    sSentencia = sSentencia & "LEFT JOIN TGTDc e ON d.codemp=e.codemp AND e.CodTDc='" & CODTDC_HPR & "') "
    sSentencia = sSentencia & "WHERE p.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND p.codaux='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(p.coddpe, p.pdocpr)", "(p.coddpe+p.pdocpr)") & "='" & sPedido & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(p.pdoano, p.mespvs)", "(p.pdoano+p.mespvs)") & "<='" & gsAnoAct & gsMesAct & "' "
    sSentencia = sSentencia & "GROUP BY p.codemp, p.codaux, p.coddpe, p.pdocpr "
    ' Obtengo los importes de movimientos
    Set .uorstTemporal = .uocnnMain.Execute(sSentencia)
    If Not (.uorstTemporal.BOF Or .uorstTemporal.EOF) And .uorstTemporal.RecordCount > 0 Then
      pnImporteMN = CDec(.uorstTemporal!impcpr_mn)
      pnImporteME = CDec(.uorstTemporal!impcpr_me)
    End If
    .uorstTemporal.Close
  End With
  ' Saldo de pedido
  With frmTCprGrd                  'Cambiar Formulario de Grid.
    sSentencia = "SELECT a.coddpe, a.pdocpr, " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS detpdo, a.tpomon, "
    sSentencia = sSentencia & "a.impmn, a.impme, a.impdife "
    sSentencia = sSentencia & "FROM copdocpr a "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & gsAnoAct & gsMesAct & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(a.coddpe, a.pdocpr)", "(a.coddpe+a.pdocpr)") & "='" & sPedido & "' "
    sSentencia = sSentencia & "AND a.codaux='" & txtLlave(0).Text & "' "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND a.fehpdo<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND a.fehpdo<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
    End If
    sSentencia = sSentencia & "ORDER BY a.coddpe, a.pdocpr"
    Set .uorstTemporal = .uocnnMain.Execute(sSentencia)
    If Not (.uorstTemporal.BOF Or .uorstTemporal.EOF) And .uorstTemporal.RecordCount > 0 Then
      pnImporDiferen = CDec(.uorstTemporal!impdife) * -1
      pnImporteMN = CDec(.uorstTemporal!ImpMN) - pnImporteMN
      pnImporteME = CDec(.uorstTemporal!ImpME) - pnImporteME
      sMoneda = IIf(IsNull(.uorstTemporal!tpomon), "", .uorstTemporal!tpomon)
    End If
    .uorstTemporal.Close
  End With
  ' Asigno los importes de acuerdo a al moneda
  pnImporte = CDec(txtDato(IIf(sMoneda = TPOMON_NAC, MINIMOINDICEIMPORTEMN, MINIMOINDICEIMPORTEME)).Text)
  pnImporte = pnImporte + CDec(txtDato(IIf(sMoneda = TPOMON_NAC, MINIMOINDICEIMPORTEMN, MINIMOINDICEIMPORTEME) + 1).Text)
  pnImporte = pnImporte + CDec(txtDato(IIf(sMoneda = TPOMON_NAC, MINIMOINDICEIMPORTEMN, MINIMOINDICEIMPORTEME) + 2).Text)
  pnImporte = pnImporte + CDec(txtDato(IIf(sMoneda = TPOMON_NAC, MINIMOINDICEIMPORTEMN, MINIMOINDICEIMPORTEME) + 3).Text)
  pnImporte = Round(IIf(sMoneda = TPOMON_NAC, pnImporteMN, pnImporteME) - pnImporte, 2)
    
  ' Inicializo la descripcion pedido
  lblDatoDeta(50).Caption = " " & sDetalle
  If pnImporte < pnImporDiferen Then
    sMensage = TEXT_8006 & Choose(gsIdioma, " y/o excede importe de Pedido", " and/or exceeds amount of Order")
    MsgBox sMensage, vbExclamation
    pfValidaPedido = False
    Exit Function
  End If
  
  ' Modifico los datos generales
  If sModificar = "S" And txtDato(50).Tag <> sPedido Then
    ' Información de las cuentas
    sSentencia = "SELECT pdc.codcta, pdc.codcco, cta.inddoc, " & Choose(gsIdioma, "cta.detcta", "cta.detctax") & " AS detcta, "
    sSentencia = sSentencia & Choose(gsIdioma, "cco.detcco", "cco.detccox") & " AS detcco, ROUND(AVG(pdc.impctadif), 2) AS impctadif, "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "ROUND(pdc.impcta_mn-ROUND(SUM((IFNULL(cpc.impcta_mn, 0)*(CASE cpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      sSentencia = sSentencia & "(IFNULL(hpc.impcta_mn, 0)*(CASE hpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))), 2), 2) AS impctamn, "
      sSentencia = sSentencia & "ROUND(pdc.impcta_me-ROUND(SUM((IFNULL(cpc.impcta_me, 0)*(CASE cpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      sSentencia = sSentencia & "(IFNULL(hpc.impcta_me, 0)*(CASE hpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))), 2), 2) AS impctame "
    Else
      sSentencia = sSentencia & "ROUND(pdc.impcta_mn-ROUND(SUM((ISNULL(cpc.impcta_mn, 0)*(CASE cpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      sSentencia = sSentencia & "(ISNULL(hpc.impcta_mn, 0)*(CASE hpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))), 2), 2) AS impctamn, "
      sSentencia = sSentencia & "ROUND(pdc.impcta_me-ROUND(SUM((ISNULL(cpc.impcta_me, 0)*(CASE cpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))+"
      sSentencia = sSentencia & "(ISNULL(hpc.impcta_me, 0)*(CASE hpd.sgntdc WHEN '" & SGNTDC_NEG & "' THEN -1 ELSE 1 END))), 2), 2) AS impctame "
    End If
    sSentencia = sSentencia & "FROM copdocpr pdo "
    sSentencia = sSentencia & "LEFT JOIN copdocprcta pdc ON pdo.codemp=pdc.codemp AND pdo.pdoano=pdc.pdoano AND pdo.mespvs=pdc.mespvs AND pdo.coddpe=pdc.coddpe AND pdo.pdocpr=pdc.pdocpr "
    sSentencia = sSentencia & "LEFT JOIN cocprdoc cpr ON pdo.codemp=cpr.codemp AND pdo.codaux=cpr.codaux AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(pdo.coddpe, pdo.pdocpr)", "(pdo.coddpe+pdo.pdocpr)") & "=cpr.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND Concat(cpr.pdoano, cpr.codtdc, cpr.serdoc, cpr.nrodoc)<>'" & gsAnoAct & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "' "
      sSentencia = sSentencia & "AND cpr.feedoc<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND (cpr.pdoano+cpr.codtdc+cpr.serdoc+cpr.nrodoc)<>'" & gsAnoAct & txtLlave(1).Text & txtLlave(2).Text & txtLlave(3).Text & "' "
      sSentencia = sSentencia & "AND cpr.feedoc<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
    End If
    sSentencia = sSentencia & "LEFT JOIN cocprdoccta cpc ON cpr.codemp=cpc.codemp AND cpr.pdoano=cpc.pdoano AND cpr.codaux=cpc.codaux AND cpr.codtdc=cpc.codtdc AND cpr.serdoc=cpc.serdoc AND cpr.nrodoc=cpc.nrodoc AND pdc.codcta=cpc.codcta AND cpc.tpocnc<=4 "
    sSentencia = sSentencia & "LEFT JOIN tgtdc cpd ON cpr.codemp=cpd.codemp AND cpr.codtdc=cpd.codtdc "
    sSentencia = sSentencia & "LEFT JOIN cohprdoc hpr ON pdo.codemp=hpr.codemp AND pdo.codaux=hpr.codaux AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(pdo.coddpe, pdo.pdocpr)", "(pdo.coddpe+pdo.pdocpr)") & "=hpr.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND hpr.feedoc<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND hpr.feedoc<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
    End If
    sSentencia = sSentencia & "LEFT JOIN cohprdoccta hpc ON hpr.codemp=hpc.codemp AND hpr.pdoano=hpc.pdoano AND hpr.codaux=hpc.codaux AND hpr.serdoc=hpc.serdoc AND hpr.nrodoc=hpc.nrodoc AND pdc.codcta=hpc.codcta AND hpc.tpocnc=1 "
    sSentencia = sSentencia & "LEFT JOIN tgtdc hpd ON hpr.codemp=hpd.codemp AND hpd.codtdc='" & CODTDC_HPR & "' "
    sSentencia = sSentencia & "LEFT JOIN cocta cta ON pdc.codemp=cta.codemp AND pdc.pdoano=cta.pdoano AND pdc.codcta=cta.codcta "
    sSentencia = sSentencia & "LEFT JOIN cocco cco ON pdc.codemp=cco.codemp AND pdc.pdoano=cco.pdoano AND pdc.codcco=cco.codcco "
    sSentencia = sSentencia & "WHERE pdo.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND pdo.codaux='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(pdo.pdoano, pdo.mespvs)", "(pdo.pdoano+pdo.mespvs)") & "<='" & gsAnoAct & gsMesAct & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "Concat(pdo.coddpe, pdo.pdocpr)", "(pdo.coddpe+pdo.pdocpr)") & "='" & sPedido & "' "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND pdo.fehpdo<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND pdo.fehpdo<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
    End If
    sSentencia = sSentencia & "GROUP BY pdc.codcta, pdc.codcco, cta.inddoc, " & Choose(gsIdioma, "cta.detcta, cco.detcco", "cta.detctax, cco.detccox") & " "
    sSentencia = sSentencia & "ORDER BY pdc.codcta, pdc.codcco"
    With frmTCprGrd                  'Cambiar Formulario de Grid.
      Set .uorstTemporal = .uocnnMain.Execute(sSentencia)
      ' Elimino las cuentas y centro de costos
      sSentencia = "DELETE FROM cocprdoccta "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND codaux='" & txtLlave(0).Text & "' AND codtdc='" & txtLlave(1).Text & "' "
      sSentencia = sSentencia & "AND serdoc='" & txtLlave(2).Text & "' AND nrodoc='" & txtLlave(3).Text & "' "
      sSentencia = sSentencia & "AND tpocnc='1'"
      .uocnnMain.Execute sSentencia
      If Not (.uorstTemporal.BOF Or .uorstTemporal.EOF) And .uorstTemporal.RecordCount > 0 Then
        sCuenta = IIf(IsNull(.uorstTemporal!CodCta), "", .uorstTemporal!CodCta)
        sCenCosto = IIf(IsNull(.uorstTemporal!codcco), "", .uorstTemporal!codcco)
        sDetaCuenta = IIf(IsNull(.uorstTemporal!detcta), "", .uorstTemporal!detcta)
        sDetaCenCosto = IIf(IsNull(.uorstTemporal!detcco), "", .uorstTemporal!detcco)
        pnImporteMN = 0: pnImporteME = 0: nOrden = 0
        ' Inserto cuentas y centro de costos
        While Not .uorstTemporal.EOF
          pnImporDiferen = CDec(.uorstTemporal!impctadif) * -1
          pnImporte = CDec(IIf(sMoneda = TPOMON_NAC, .uorstTemporal!impctamn, .uorstTemporal!impctame))
          If Not (pnImporte < pnImporDiferen) Then
            sSentencia = "INSERT INTO cocprdoccta(codemp, pdoano, codaux, codtdc, serdoc, nrodoc, tpocnc, orden, codcta, glodet, glodetx, codruc, impcta_mn, impcta_me, usrcre, fyhcre) "
            sSentencia = sSentencia & " VALUES("
            sSentencia = sSentencia & "'" & gsCodEmp & "', "
            sSentencia = sSentencia & "'" & gsAnoAct & "', "
            sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
            sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
            sSentencia = sSentencia & "'" & txtLlave(2).Text & "', "
            sSentencia = sSentencia & "'" & txtLlave(3).Text & "', "
            nOrden = nOrden + 1
            sSentencia = sSentencia & "'1', '" & Format(nOrden, "00") & "', "
            sSentencia = sSentencia & "'" & .uorstTemporal!CodCta & "', "
            sSentencia = sSentencia & IIf(txtDato(Choose(gsIdioma, 3, 49)).Text = "", "Null", "'" & txtDato(Choose(gsIdioma, 3, 49)).Text & "'") & ", "
            sSentencia = sSentencia & IIf(txtDato(Choose(gsIdioma, 49, 3)).Text = "", "Null", "'" & txtDato(Choose(gsIdioma, 49, 3)).Text & "'") & ", "
            sSentencia = sSentencia & IIf(.uorstTemporal!IndDoc = INDDOC_ACT, "'" & txtLlave(0).Text & "'", "Null") & ", "
            sRegistro = Round(IIf(sMoneda = TPOMON_NAC, .uorstTemporal!impctamn, (.uorstTemporal!impctame) * CDec(txtDato(4).Text)), 2)
            pnImporteMN = pnImporteMN + CDec(sRegistro)
            sSentencia = sSentencia & CDec(sRegistro) & ", "
            sRegistro = Round(IIf(sMoneda = TPOMON_EXT, .uorstTemporal!impctame, (.uorstTemporal!impctamn) / CDec(txtDato(4).Text)), 2)
            pnImporteME = pnImporteME + CDec(sRegistro)
            sSentencia = sSentencia & CDec(sRegistro) & ", "
            sSentencia = sSentencia & "'" & gsAbvUsr & "', "
            sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "))"
            .uocnnMain.Execute sSentencia
            sRegistro = IIf(IsNull(.uorstTemporal!codcco), "", .uorstTemporal!codcco)
            If sRegistro <> "" Then
              sSentencia = "INSERT INTO cocprdoccco(codemp, pdoano, codaux, codtdc, serdoc, nrodoc, tpocnc, orden, codcta, codcco, impcco_mn, impcco_me, usrcre, fyhcre) "
              sSentencia = sSentencia & " VALUES("
              sSentencia = sSentencia & "'" & gsCodEmp & "', "
              sSentencia = sSentencia & "'" & gsAnoAct & "', "
              sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
              sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
              sSentencia = sSentencia & "'" & txtLlave(2).Text & "', "
              sSentencia = sSentencia & "'" & txtLlave(3).Text & "', "
              sSentencia = sSentencia & "'1', '" & Format(nOrden, "00") & "', "
              sSentencia = sSentencia & "'" & .uorstTemporal!CodCta & "', "
              sSentencia = sSentencia & "'" & .uorstTemporal!codcco & "', "
              sRegistro = Round(IIf(sMoneda = TPOMON_NAC, .uorstTemporal!impctamn, (.uorstTemporal!impctame) * CDec(txtDato(4).Text)), 2)
              sSentencia = sSentencia & CDec(sRegistro) & ", "
              sRegistro = Round(IIf(sMoneda = TPOMON_EXT, .uorstTemporal!impctame, (.uorstTemporal!impctamn) / CDec(txtDato(4).Text)), 2)
              sSentencia = sSentencia & CDec(sRegistro) & ", "
              sSentencia = sSentencia & "'" & gsAbvUsr & "', "
              sSentencia = sSentencia & IIf(ps_Plataforma = pSrvMySql, "DATE_FORMAT(", "CONVERT(datetime, ") & "'" & Format(Now, s_FmtFeHoMysql_0) & "', " & IIf(ps_Plataforma = pSrvMySql, "'%Y-%m-%d %T'", "120") & "))"
              .uocnnMain.Execute sSentencia
            End If
          End If
          .uorstTemporal.MoveNext
        Wend
        ppAbreCtaCCo
        ' Inicializo tipo de registro
        cmdMas(1).Tag = INDMASCTA_MAS
      End If
      frmTCprGrd.uorstTemporal.Close
    End With
    txtDato(MINIMOINDICEIMPORTEMN).Text = Format(pnImporteMN, FORMATO_NUM_1)
    txtDato(MINIMOINDICEIMPORTEME).Text = Format(pnImporteME, FORMATO_NUM_1)
    txtDato(MINIMOINDICECUENTA).Text = sCuenta
    txtDato(MINIMOINDICECCOSTO).Text = sCenCosto
    lblDatoDeta(MINIMOINDICECUENTA).Caption = " " & sDetaCuenta
    lblDatoDeta(MINIMOINDICECCOSTO).Caption = " " & sDetaCenCosto
  End If
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

