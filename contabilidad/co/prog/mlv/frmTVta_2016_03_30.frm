VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTVta 
   Caption         =   "[Título]"
   ClientHeight    =   8325
   ClientLeft      =   840
   ClientTop       =   1245
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   9525
   Begin VB.CheckBox chkDesactivar 
      Caption         =   "Des&activar Cuentas"
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   1800
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   1
      Left            =   3720
      TabIndex        =   179
      Top             =   4560
      Width           =   5715
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
         Left            =   1920
         TabIndex        =   181
         Top             =   180
         Width           =   1815
      End
      Begin VB.CheckBox chkIndCDt 
         Caption         =   "Detracción"
         ForeColor       =   &H00800000&
         Height          =   200
         Left            =   105
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   218
         Width           =   1080
      End
      Begin MSComCtl2.DTPicker dtpDato 
         Height          =   315
         Index           =   4
         Left            =   4320
         TabIndex        =   182
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65798145
         CurrentDate     =   37102
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
         Index           =   39
         Left            =   3840
         TabIndex        =   184
         Top             =   240
         Width           =   495
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
         Index           =   38
         Left            =   1200
         TabIndex        =   183
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   1
      Left            =   3720
      Picture         =   "frmTVta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
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
      Left            =   1050
      TabIndex        =   33
      Top             =   2730
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CheckBox chkIndVtaext 
      Caption         =   " Venta Externa"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   615
      Width           =   1725
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   40
      Left            =   8115
      Picture         =   "frmTVta.frx":01AA
      Style           =   1  'Graphical
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   3150
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
      Left            =   1230
      TabIndex        =   37
      Top             =   3150
      Width           =   2250
   End
   Begin VB.ComboBox cboImpuesto 
      Height          =   315
      ItemData        =   "frmTVta.frx":0354
      Left            =   7005
      List            =   "frmTVta.frx":0356
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   3705
      Width           =   1545
   End
   Begin VB.ComboBox cboCategoria 
      Height          =   315
      ItemData        =   "frmTVta.frx":0358
      Left            =   7005
      List            =   "frmTVta.frx":035A
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   4260
      Width           =   1545
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
      Index           =   37
      Left            =   4050
      TabIndex        =   25
      Top             =   2040
      Width           =   4515
   End
   Begin VB.Frame fraAsiento 
      Height          =   540
      Left            =   0
      TabIndex        =   46
      Top             =   4095
      Width           =   6960
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   38
         Left            =   1560
         TabIndex        =   48
         Top             =   165
         Width           =   560
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   38
         Left            =   6600
         Picture         =   "frmTVta.frx":035C
         Style           =   1  'Graphical
         TabIndex        =   126
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
         Index           =   23
         Left            =   105
         TabIndex        =   47
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
         Index           =   38
         Left            =   2100
         TabIndex        =   49
         Top             =   165
         Width           =   4500
      End
   End
   Begin VB.ComboBox cboRetencion 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2400
      Width           =   2925
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
      Index           =   36
      Left            =   4050
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   2400
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
      Picture         =   "frmTVta.frx":0506
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   2670
      Width           =   720
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "Cli&ente"
      Height          =   315
      Left            =   8625
      TabIndex        =   142
      Top             =   3210
      Width           =   825
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   525
      Index           =   0
      Left            =   0
      TabIndex        =   39
      Top             =   3570
      Width           =   6960
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
         Left            =   6150
         TabIndex        =   44
         Top             =   165
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
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   41
         Top             =   165
         Width           =   555
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   0
         Left            =   4725
         Picture         =   "frmTVta.frx":0ADC
         Style           =   1  'Graphical
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   165
         Width           =   255
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
         Left            =   90
         TabIndex        =   40
         Top             =   180
         Width           =   450
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
         Left            =   5100
         TabIndex        =   43
         Top             =   180
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
         Height          =   285
         Index           =   0
         Left            =   1170
         TabIndex        =   42
         Top             =   165
         Width           =   3540
      End
   End
   Begin VB.CheckBox chkIndPreGen 
      Caption         =   "Cuentas &Registradas"
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   1800
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CheckBox chkCalcularISC 
      Caption         =   "Calcular I.&S.C."
      ForeColor       =   &H00800000&
      Height          =   200
      Left            =   120
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4680
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
      Index           =   34
      Left            =   6885
      TabIndex        =   8
      Top             =   570
      Width           =   510
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
      Index           =   35
      Left            =   7380
      TabIndex        =   9
      Top             =   570
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   8640
      ScaleHeight     =   2535
      ScaleWidth      =   885
      TabIndex        =   138
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
         Picture         =   "frmTVta.frx":0C86
         Style           =   1  'Graphical
         TabIndex        =   116
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
         Picture         =   "frmTVta.frx":0DD0
         Style           =   1  'Graphical
         TabIndex        =   115
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
         Picture         =   "frmTVta.frx":0ED2
         Style           =   1  'Graphical
         TabIndex        =   114
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
         Picture         =   "frmTVta.frx":0FD4
         Style           =   1  'Graphical
         TabIndex        =   113
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
         Picture         =   "frmTVta.frx":111E
         Style           =   1  'Graphical
         TabIndex        =   111
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
         Picture         =   "frmTVta.frx":12C8
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   0
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
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   29
      Top             =   2040
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
      Height          =   285
      Index           =   0
      Left            =   900
      TabIndex        =   1
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   0
      Left            =   5820
      Picture         =   "frmTVta.frx":1472
      Style           =   1  'Graphical
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   280
      Index           =   33
      Left            =   7860
      Picture         =   "frmTVta.frx":161C
      Style           =   1  'Graphical
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   1380
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
      Index           =   33
      Left            =   1080
      TabIndex        =   19
      Top             =   1380
      Width           =   1275
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTVta.frx":17C6
      Left            =   1080
      List            =   "frmTVta.frx":17C8
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   2040
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   2
      Left            =   3345
      TabIndex        =   13
      Top             =   1005
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65798145
      CurrentDate     =   37102
   End
   Begin VB.CheckBox chkCalcularIGV 
      Caption         =   "Calcular I.G.&V."
      ForeColor       =   &H00800000&
      Height          =   200
      Left            =   120
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4920
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
      Index           =   3
      Left            =   4050
      TabIndex        =   24
      Top             =   1710
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
      Index           =   2
      Left            =   1080
      TabIndex        =   22
      Top             =   1710
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
      Height          =   285
      Index           =   2
      Left            =   7380
      TabIndex        =   5
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
      Height          =   285
      Index           =   1
      Left            =   6885
      TabIndex        =   4
      Top             =   120
      Width           =   510
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   5415
      TabIndex        =   15
      Top             =   1005
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65798145
      CurrentDate     =   37102
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   3375
      Left            =   0
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   5280
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   441
      ForeColor       =   8388608
      TabCaption(0)   =   "I&mportes"
      TabPicture(0)   =   "frmTVta.frx":17CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(19)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTexto(21)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTexto(20)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTexto(13)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTexto(15)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTexto(14)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTexto(18)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblTexto(17)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTexto(16)"
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
      Tab(0).Control(67)=   "txtDato(30)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "cmdDatoAyud(30)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txtDato(31)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cmdDatoAyud(31)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txtDato(32)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "cmdDatoAyud(32)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).ControlCount=   73
      TabCaption(1)   =   "C&uentas"
      TabPicture(1)   =   "frmTVta.frx":17E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgrDetalle"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PDB-NC-ND"
      TabPicture(2)   =   "frmTVta.frx":1802
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDetracion"
      Tab(2).Control(1)=   "fraReferencia"
      Tab(2).Control(2)=   "fraPercepcion"
      Tab(2).Control(3)=   "fraExterior"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Sunat"
      TabPicture(3)   =   "frmTVta.frx":181E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdDatoAyud(43)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtDato(43)"
      Tab(3).Control(2)=   "lblTexto(42)"
      Tab(3).Control(3)=   "lblDatoDeta(43)"
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   43
         Left            =   -67860
         Picture         =   "frmTVta.frx":183A
         Style           =   1  'Graphical
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   480
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
         Index           =   43
         Left            =   -73440
         TabIndex        =   189
         Top             =   480
         Width           =   1275
      End
      Begin VB.Frame fraDetracion 
         Caption         =   "Detacción"
         ForeColor       =   &H00C00000&
         Height          =   570
         Left            =   -71640
         TabIndex        =   185
         Top             =   2640
         Width           =   6030
         Begin VB.ComboBox cboDetraccion 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   150
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
            Index           =   40
            Left            =   60
            TabIndex        =   187
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.Frame fraReferencia 
         Caption         =   " Documento de Referencia "
         ForeColor       =   &H00C00000&
         Height          =   1680
         Left            =   -71670
         TabIndex        =   163
         Top             =   1020
         Width           =   6030
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
            Index           =   10
            Left            =   945
            TabIndex        =   173
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
            Index           =   11
            Left            =   945
            TabIndex        =   176
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
            Index           =   12
            Left            =   945
            TabIndex        =   174
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
            Index           =   13
            Left            =   945
            TabIndex        =   177
            Top             =   1275
            Width           =   1695
         End
         Begin VB.CommandButton cmdDatoAyud 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Height          =   280
            Index           =   39
            Left            =   5625
            Picture         =   "frmTVta.frx":19E4
            Style           =   1  'Graphical
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   285
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
            Index           =   39
            Left            =   945
            TabIndex        =   165
            Top             =   285
            Width           =   315
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   8
            Left            =   945
            TabIndex        =   168
            Top             =   615
            Width           =   500
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   9
            Left            =   1440
            TabIndex        =   169
            Top             =   615
            Width           =   1185
         End
         Begin MSComCtl2.DTPicker dtpDetalle 
            Height          =   315
            Index           =   2
            Left            =   3405
            TabIndex        =   171
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65798145
            CurrentDate     =   37102
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
            Index           =   33
            Left            =   60
            TabIndex        =   172
            Top             =   960
            Width           =   885
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
            Index           =   34
            Left            =   60
            TabIndex        =   175
            Top             =   1305
            Width           =   495
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
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   32
            Left            =   2685
            TabIndex        =   170
            Top             =   645
            Width           =   720
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
            Index           =   39
            Left            =   1245
            TabIndex        =   166
            Top             =   285
            Width           =   4410
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
            Index           =   31
            Left            =   60
            TabIndex        =   167
            Top             =   645
            Width           =   645
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
            Index           =   30
            Left            =   60
            TabIndex        =   164
            Top             =   300
            Width           =   765
         End
      End
      Begin VB.Frame fraPercepcion 
         Caption         =   " Percepcion "
         ForeColor       =   &H00C00000&
         Height          =   690
         Left            =   -71685
         TabIndex        =   156
         Top             =   330
         Width           =   6030
         Begin VB.ComboBox cboPercepcion 
            Height          =   315
            Left            =   945
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   270
            Width           =   2130
         End
         Begin VB.CheckBox chkIndpercep 
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   1065
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   15
            Width           =   180
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   6
            Left            =   4830
            TabIndex        =   162
            Top             =   270
            Width           =   1110
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   5
            Left            =   4335
            TabIndex        =   161
            Top             =   270
            Width           =   500
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
            Index           =   28
            Left            =   60
            TabIndex        =   158
            Top             =   300
            Width           =   795
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento :"
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
            Index           =   29
            Left            =   3165
            TabIndex        =   160
            Top             =   300
            Width           =   1125
         End
      End
      Begin VB.Frame fraExterior 
         Caption         =   " Venta Externa "
         ForeColor       =   &H00C00000&
         Height          =   2355
         Left            =   -74925
         TabIndex        =   144
         Top             =   330
         Width           =   3195
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
            Index           =   3
            Left            =   1365
            TabIndex        =   154
            Top             =   1365
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
            Index           =   4
            Left            =   1365
            TabIndex        =   155
            Top             =   1365
            Width           =   1695
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   2
            Left            =   2280
            TabIndex        =   148
            Top             =   270
            Width           =   800
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   0
            Left            =   1365
            TabIndex        =   146
            Top             =   270
            Width           =   440
         End
         Begin VB.TextBox txtDetalle 
            ForeColor       =   &H80000012&
            Height          =   280
            Index           =   1
            Left            =   1800
            TabIndex        =   147
            Top             =   270
            Width           =   500
         End
         Begin MSComCtl2.DTPicker dtpDetalle 
            Height          =   315
            Index           =   1
            Left            =   1365
            TabIndex        =   152
            Top             =   1005
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65798145
            CurrentDate     =   37102
         End
         Begin MSComCtl2.DTPicker dtpDetalle 
            Height          =   315
            Index           =   0
            Left            =   1365
            TabIndex        =   150
            Top             =   660
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65798145
            CurrentDate     =   37102
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Valor FOB :"
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
            Index           =   27
            Left            =   60
            TabIndex        =   153
            Top             =   1395
            Width           =   840
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "F. Regularizar :"
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
            Index           =   26
            Left            =   60
            TabIndex        =   151
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "F. Embarque :"
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
            Index           =   25
            Left            =   60
            TabIndex        =   149
            Top             =   705
            Width           =   990
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
            Index           =   24
            Left            =   60
            TabIndex        =   145
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   32
         Left            =   9060
         Picture         =   "frmTVta.frx":1B8E
         Style           =   1  'Graphical
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   2310
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
         TabIndex        =   109
         Top             =   2310
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   31
         Left            =   9060
         Picture         =   "frmTVta.frx":1D38
         Style           =   1  'Graphical
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   2025
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
         TabIndex        =   101
         Top             =   2025
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   30
         Left            =   9060
         Picture         =   "frmTVta.frx":1EE2
         Style           =   1  'Graphical
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   1740
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
         TabIndex        =   93
         Top             =   1740
         Width           =   675
      End
      Begin VB.CheckBox chkMonedaActiva 
         Caption         =   "M&oneda activa"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1320
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   345
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
         Height          =   280
         Index           =   7
         Left            =   3000
         Picture         =   "frmTVta.frx":208C
         TabIndex        =   106
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
         Picture         =   "frmTVta.frx":218E
         TabIndex        =   98
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
         Picture         =   "frmTVta.frx":2290
         TabIndex        =   90
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
         Picture         =   "frmTVta.frx":2392
         TabIndex        =   82
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
         Picture         =   "frmTVta.frx":2494
         TabIndex        =   74
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
         Picture         =   "frmTVta.frx":2596
         TabIndex        =   66
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
         Picture         =   "frmTVta.frx":2698
         TabIndex        =   58
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
         Index           =   11
         Left            =   1320
         TabIndex        =   104
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
         Index           =   10
         Left            =   1320
         TabIndex        =   96
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
         Index           =   9
         Left            =   1320
         TabIndex        =   88
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
         Index           =   8
         Left            =   1320
         TabIndex        =   80
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
         Index           =   7
         Left            =   1320
         TabIndex        =   72
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
         Index           =   6
         Left            =   1320
         TabIndex        =   64
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
         Index           =   29
         Left            =   6840
         TabIndex        =   85
         Top             =   1455
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   29
         Left            =   9060
         Picture         =   "frmTVta.frx":279A
         Style           =   1  'Graphical
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   1455
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
         TabIndex        =   77
         Top             =   1170
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   28
         Left            =   9060
         Picture         =   "frmTVta.frx":2944
         Style           =   1  'Graphical
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1170
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
         TabIndex        =   69
         Top             =   885
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   27
         Left            =   9060
         Picture         =   "frmTVta.frx":2AEE
         Style           =   1  'Graphical
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   885
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
         TabIndex        =   61
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   26
         Left            =   9060
         Picture         =   "frmTVta.frx":2C98
         Style           =   1  'Graphical
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   107
         Top             =   2310
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   25
         Left            =   6540
         Picture         =   "frmTVta.frx":2E42
         Style           =   1  'Graphical
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   2310
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
         TabIndex        =   99
         Top             =   2025
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   24
         Left            =   6540
         Picture         =   "frmTVta.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   2025
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
         TabIndex        =   91
         Top             =   1740
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   23
         Left            =   6540
         Picture         =   "frmTVta.frx":3196
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   1740
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
         TabIndex        =   83
         Top             =   1455
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   22
         Left            =   6540
         Picture         =   "frmTVta.frx":3340
         Style           =   1  'Graphical
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   1455
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
         TabIndex        =   75
         Top             =   1170
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   21
         Left            =   6540
         Picture         =   "frmTVta.frx":34EA
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   1170
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   20
         Left            =   6540
         Picture         =   "frmTVta.frx":3694
         Style           =   1  'Graphical
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   885
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
         TabIndex        =   67
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
         Index           =   19
         Left            =   3300
         TabIndex        =   59
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   19
         Left            =   6540
         Picture         =   "frmTVta.frx":383E
         Style           =   1  'Graphical
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   600
         Width           =   255
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   2325
         Left            =   -74880
         TabIndex        =   120
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
         Index           =   5
         Left            =   1320
         TabIndex        =   56
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
         TabIndex        =   105
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
         Index           =   17
         Left            =   1320
         TabIndex        =   97
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
         Index           =   16
         Left            =   1320
         TabIndex        =   89
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
         Index           =   15
         Left            =   1320
         TabIndex        =   81
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
         Index           =   14
         Left            =   1320
         TabIndex        =   73
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
         Index           =   13
         Left            =   1320
         TabIndex        =   65
         Top             =   885
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
         TabIndex        =   57
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Codigo de Moneda:"
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
         Index           =   42
         Left            =   -74880
         TabIndex        =   192
         Top             =   510
         Width           =   1380
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
         Index           =   43
         Left            =   -72180
         TabIndex        =   191
         Top             =   480
         Width           =   4335
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
         Index           =   32
         Left            =   7500
         TabIndex        =   110
         Top             =   2310
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
         Height          =   280
         Index           =   31
         Left            =   7500
         TabIndex        =   102
         Top             =   2025
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
         Height          =   280
         Index           =   30
         Left            =   7500
         TabIndex        =   94
         Top             =   1740
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
         Height          =   280
         Index           =   29
         Left            =   7500
         TabIndex        =   86
         Top             =   1455
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
         Height          =   280
         Index           =   28
         Left            =   7500
         TabIndex        =   78
         Top             =   1170
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
         Height          =   280
         Index           =   27
         Left            =   7500
         TabIndex        =   70
         Top             =   885
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
         Height          =   280
         Index           =   26
         Left            =   7500
         TabIndex        =   62
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
         Height          =   280
         Index           =   25
         Left            =   4260
         TabIndex        =   108
         Top             =   2310
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
         Height          =   280
         Index           =   24
         Left            =   4260
         TabIndex        =   100
         Top             =   2025
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
         Height          =   280
         Index           =   23
         Left            =   4260
         TabIndex        =   92
         Top             =   1740
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
         Height          =   280
         Index           =   22
         Left            =   4260
         TabIndex        =   84
         Top             =   1455
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
         Height          =   280
         Index           =   21
         Left            =   4260
         TabIndex        =   76
         Top             =   1170
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
         Height          =   280
         Index           =   20
         Left            =   4260
         TabIndex        =   68
         Top             =   885
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
         Height          =   280
         Index           =   19
         Left            =   4260
         TabIndex        =   60
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
         Index           =   16
         Left            =   90
         TabIndex        =   79
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
         Index           =   17
         Left            =   90
         TabIndex        =   87
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
         Index           =   18
         Left            =   90
         TabIndex        =   95
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
         Index           =   14
         Left            =   90
         TabIndex        =   63
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
         Index           =   15
         Left            =   90
         TabIndex        =   71
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
         Index           =   13
         Left            =   90
         TabIndex        =   55
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
         Index           =   20
         Left            =   3420
         TabIndex        =   119
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
         Index           =   21
         Left            =   6960
         TabIndex        =   118
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
         Index           =   19
         Left            =   90
         TabIndex        =   103
         Top             =   2385
         Width           =   555
      End
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      Top             =   1005
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65798145
      CurrentDate     =   37102
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   1
      Left            =   7455
      TabIndex        =   17
      Top             =   1005
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   65798145
      CurrentDate     =   37102
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
      Left            =   1800
      TabIndex        =   34
      Top             =   2730
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Ubigeo :"
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
      Left            =   60
      TabIndex        =   32
      Top             =   2730
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Vcmto. :"
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
      Left            =   6660
      TabIndex        =   16
      Top             =   1050
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
      Height          =   300
      Index           =   40
      Left            =   3480
      TabIndex        =   38
      Top             =   3150
      Width           =   4650
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ord.Servicio :"
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
      Index           =   35
      Left            =   210
      TabIndex        =   36
      Top             =   3180
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   480
      Left            =   60
      Top             =   3060
      Width           =   8475
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Operac :"
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
      Left            =   60
      TabIndex        =   30
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label lblTexto 
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   5940
      TabIndex        =   7
      Top             =   615
      Width           =   885
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8500
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Operación :"
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
      TabIndex        =   10
      Top             =   1050
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   60
      X2              =   8500
      Y1              =   495
      Y2              =   495
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
      Left            =   1200
      TabIndex        =   2
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
      Height          =   285
      Index           =   33
      Left            =   2340
      TabIndex        =   20
      Top             =   1440
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
      Left            =   3540
      TabIndex        =   23
      Top             =   1740
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
      TabIndex        =   21
      Top             =   1740
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
      TabIndex        =   26
      Top             =   2070
      Width           =   615
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "F.Emisión :"
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
      Left            =   4635
      TabIndex        =   14
      Top             =   1050
      Width           =   765
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
      Left            =   1920
      TabIndex        =   28
      Top             =   2070
      Width           =   705
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
      Index           =   36
      Left            =   2355
      TabIndex        =   12
      Top             =   1050
      Width           =   930
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
      Index           =   1
      Left            =   6285
      TabIndex        =   3
      Top             =   150
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
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   720
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
      Index           =   6
      Left            =   60
      TabIndex        =   18
      Top             =   1410
      Width           =   525
   End
End
Attribute VB_Name = "frmTVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2016-02-02.06  correccion ple
Option Explicit

'Public pglodet_len As Integer '2016-02-23 control tamaño de glosa en formato
'Public pglodet_len_max As Integer  '2016-02-23 tamaño maximo de glosa en formato

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
Private Const ps_OrdenCta As String = "01"

'ini 2015-07-07 detrac vtas
Private Sub chkIndCDt_Click()
  If Not (chkIndCDt.Value = vbChecked) Then txtDato(42).Text = ""
  txtDato(42).Enabled = (chkIndCDt.Value)
  dtpDato(4).Enabled = (chkIndCDt.Value)
End Sub
'fin 2015-07-07 detrac vtas


Private Sub cboRetencion_Click()
  If cboRetencion.Tag <> cboRetencion.ListIndex Then
    txtDato(36).Text = gsGloDoc_Rtc(cboRetencion.ListIndex)
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
  sSQL = sSQL & "LEFT JOIN covtadoccta det ON vta.codemp=det.codemp AND vta.pdoano=det.pdoano AND vta.codtdc=det.codtdc AND vta.serdoc=det.serdoc AND vta.nrodoc=det.nrodoc "
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
    .ActiveConnection = frmTVtaGrd.uocnnMain
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
    frmTVtaGrd.uocnnMain.Execute "DROP TABLE IF EXISTS trptdocventa"
    sSQL = "CREATE TEMPORARY TABLE IF NOT EXISTS trptdocventa (documento varchar(16) NOT NULL, "
    sSQL = sSQL & "secuencia smallint(1) DEFAULT '0', codtdc char(2) NOT NULL, "
    sSQL = sSQL & "serdoc char(4) NOT NULL, nrodoc varchar(10) NOT NULL, "
    sSQL = sSQL & "emision date NULL, modifica date NULL, "
    'sSQL = sSQL & "codaux varchar(11) NULL, razaux varchar(60) NULL, " '2015-08-12
    sSQL = sSQL & "codaux varchar(11) NULL, razaux varchar(80) NULL, "
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
    frmTVtaGrd.uocnnMain.Execute "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptdocventa') DROP TABLE #trptdocventa"
    sSQL = "CREATE TABLE #trptdocventa (documento varchar(16) NOT NULL, "
    sSQL = sSQL & "secuencia smallint DEFAULT '0', codtdc char(2) NOT NULL, "
    sSQL = sSQL & "serdoc char(4) NOT NULL, nrodoc varchar(10) NOT NULL, "
    sSQL = sSQL & "emision smalldatetime NULL, modifica smalldatetime NULL, "
    'sSQL = sSQL & "codaux varchar(11) NULL, razaux varchar(60) NULL, " '2105-08-12
    sSQL = sSQL & "codaux varchar(11) NULL, razaux varchar(80) NULL, "
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
  frmTVtaGrd.uocnnMain.Execute sSQL
  
  nRegistro = 0: nContador = 0
  ' Genero la informació de impresión
'  Dim xlen_det As Integer '2016-02-23 control tamaño de glosa en formato
  While Not porstRegistro.EOF
  
'ini 2016-02-23 control tamaño de glosa en formato
'    xlen_det = xlen_det + _
'    Len(Trim(IIf(IsNull(porstRegistro!glodet0), "", porstRegistro!glodet0))) + _
'    Len(Trim(IIf(IsNull(porstRegistro!glodet1), "", porstRegistro!glodet1)))
'fin 2016-02-23 control tamaño de glosa en formato
   
    nDiferencia = ppNumeroLinea(IIf(IsNull(porstRegistro!glodet0), "", porstRegistro!glodet0) & IIf(IsNull(porstRegistro!glodet1), "", porstRegistro!glodet1))
    nContador = nContador + nDiferencia
    nRegistro = nRegistro + 1
    sSQL = "INSERT INTO " & ps_Prefijo & "trptdocventa "
    sSQL = sSQL & "(documento, secuencia, codtdc, serdoc, nrodoc, emision, modifica, codaux, razaux, "
    sSQL = sSQL & "diraux, rucaux, tpomon, signomon, pctigv, refdoc, glodet0, glodet1, glortc, "
    sSQL = sSQL & "impbase, impigv, imptotal, dettdc , forimp, importeletra) "
    sSQL = sSQL & "VALUES ('" & porstRegistro!documento & "', "
    sSQL = sSQL & nRegistro & ", "
    sSQL = sSQL & "'" & porstRegistro!codtdc & "', "
    sSQL = sSQL & "'" & porstRegistro!serdoc & "', "
    sSQL = sSQL & "'" & porstRegistro!nrodoc & "', "
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
    sSQL = sSQL & "'" & porstRegistro!RefDoc & "', "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glodet0), "Null", "'" & porstRegistro!glodet0 & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glodet1), "Null", "'" & porstRegistro!glodet1 & "'") & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glortc), "Null", "'" & porstRegistro!glortc & "'") & ", "
    sSQL = sSQL & CDec(porstRegistro!impbase) & ", "
    sSQL = sSQL & CDec(porstRegistro!impigv) & ", "
    sSQL = sSQL & CDec(porstRegistro!imptotal) & ", "
    sSQL = sSQL & "'" & porstRegistro!dettdc & "', "
    sSQL = sSQL & "'" & cboImpuesto.ListIndex & "', "
    ' sSQL = sSQL & "'" & porstRegistro!forimp & "', "
    sSQL = sSQL & "'" & sImporteLetras & "')"
    frmTVtaGrd.uocnnMain.Execute sSQL
    porstRegistro.MoveNext
  Wend
  porstRegistro.MovePrevious
  
'ini 2016-02-23 control tamaño de glosa en formato
'    If xlen_det > frmTVta.pglodet_len_max Then
'        MsgBox ("No puede ingresar mas de " & Str(frmTVta.pglodet_len_max) & " caracteres")
'        porstRegistro.Close
'        Set porstRegistro = Nothing
'        frmTVtaGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptdocventa", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptdocventa') DROP TABLE #trptdocventa")
'       Exit Sub
'    End If
'fin 2016-02-23 control tamaño de glosa en formato
  
  
  ' Inserto los detalles adicionales
  nRegistro = nContador + 1
  For nContador = nRegistro To 7
    sSQL = "INSERT INTO " & ps_Prefijo & "trptdocventa "
    sSQL = sSQL & "(documento, secuencia, codtdc, serdoc, nrodoc, emision, modifica, codaux, razaux, "
    sSQL = sSQL & "diraux, rucaux, tpomon, signomon, pctigv, refdoc, glodet0, glodet1, glortc, "
    sSQL = sSQL & "impbase, impigv, imptotal, dettdc , forimp, importeletra) "
    sSQL = sSQL & "VALUES ('" & porstRegistro!documento & "', "
    sSQL = sSQL & nContador & ", "
    sSQL = sSQL & "'" & porstRegistro!codtdc & "', "
    sSQL = sSQL & "'" & porstRegistro!serdoc & "', "
    sSQL = sSQL & "'" & porstRegistro!nrodoc & "', "
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
    sSQL = sSQL & "'" & porstRegistro!RefDoc & "', "
    sSQL = sSQL & "Null" & ", "
    sSQL = sSQL & "Null" & ", "
    sSQL = sSQL & IIf(IsNull(porstRegistro!glortc), "Null", "'" & porstRegistro!glortc & "'") & ", "
    sSQL = sSQL & "0" & ", "
    sSQL = sSQL & CDec(porstRegistro!impigv) & ", "
    sSQL = sSQL & CDec(porstRegistro!imptotal) & ", "
    sSQL = sSQL & "'" & porstRegistro!dettdc & "', "
    sSQL = sSQL & "'" & cboImpuesto.ListIndex & "', "
'    sSQL = sSQL & "'" & porstRegistro!forimp & "', "
    sSQL = sSQL & "'" & sImporteLetras & "')"
    frmTVtaGrd.uocnnMain.Execute sSQL
  Next nContador
  
  ' Obtengo los registrso de impresion
  With porstRegistro
    If .State = adStateOpen Then .Close
    .ActiveConnection = frmTVtaGrd.uocnnMain
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
  frmTVtaGrd.uocnnMain.Execute IIf(ps_Plataforma = pSrvMySql, "DROP TABLE IF EXISTS trptdocventa", "IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE LEFT(name, 13)='#trptdocventa') DROP TABLE #trptdocventa")

End Sub

']
Private Sub Form_Load()

'   pglodet_len = 0 '2016-02-23 control tamaño de glosa en formato
'   'pglodet_len_max = 940 '2016-02-23 control tamaño de glosa en formato
'   pglodet_len_max = 300 '2016-02-23 control tamaño de glosa en formato
   
   pbValidada = False
   pbFecha = True
   Me.KeyPreview = True
   
   With frmTVtaGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
      txtLlave(0).MaxLength = .uorstMain!codtdc.DefinedSize
      txtLlave(1).MaxLength = .uorstMain!serdoc.DefinedSize
      txtLlave(2).MaxLength = .uorstMain!nrodoc.DefinedSize
    ']
   
    '[Datos                            'Cambiar.
      With cboTpoMon
        .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
        .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
      End With
      With cboRetencion
        .AddItem Choose(gsIdioma, "Ninguna", "None"), TPOGRU1_IND
        .AddItem Choose(gsIdioma, "Sujeta a Detracción", "Holds to Deduction"), TPOGRU2_IND
        .AddItem Choose(gsIdioma, "No Sujeta a Detracción", "Not Holds to Deduction"), TPOGRU3_IND
      End With
      
      With cboImpuesto
        .AddItem TEXT_Ninguno, TipoImpuesto.Ninguno
        .AddItem TEXT_ResponsableInscrito, TipoImpuesto.ResponsableInscrito
        .AddItem TEXT_ResponsableMonotributo, TipoImpuesto.ResponsableMonotributo
        .AddItem TEXT_Exepto, TipoImpuesto.Exepto
        .AddItem TEXT_NoAlcanzado, TipoImpuesto.NoAlcanzado
        .AddItem TEXT_ConsumidosFinal, TipoImpuesto.ConsumidosFinal
        .AddItem "Factura Semanal", TipoImpuesto.TipoSemanal
        .AddItem "Factura Quincenal", TipoImpuesto.TipoQuincenal
        .AddItem "Factura Mensual", TipoImpuesto.TipoMensual
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
      
      txtDato(40).MaxLength = .uorstMain!codcon.DefinedSize
      txtDato(0).MaxLength = .uorstMain!coddro.DefinedSize
      txtDato(1).MaxLength = .uorstMain!NroCpb.DefinedSize
      txtDato(2).MaxLength = .uorstMain!RefDoc.DefinedSize
      txtDato(Choose(gsIdioma, 3, 37)).MaxLength = .uorstMain!GloDoc.DefinedSize
      txtDato(Choose(gsIdioma, 37, 3)).MaxLength = .uorstMain!glodocx.DefinedSize
      txtDato(38).MaxLength = .uorstMain!codasi.DefinedSize
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
      txtDato(33).MaxLength = .uorstMain!codaux.DefinedSize
      txtDato(34).MaxLength = .uorstMain!SerDoc_Fin.DefinedSize
      txtDato(35).MaxLength = .uorstMain!nrodoc_fin.DefinedSize
      txtDato(36).MaxLength = .uorstMain!glodoc_rtc.DefinedSize
      ' Datos detalle pdb
      txtDetalle(0).MaxLength = .uorstMain!codaduana.DefinedSize
      txtDetalle(1).MaxLength = .uorstMain!annodua.DefinedSize
      txtDetalle(2).MaxLength = .uorstMain!nrodua.DefinedSize
      txtDetalle(3).MaxLength = 14
      txtDetalle(4).MaxLength = 14
      txtDetalle(5).MaxLength = .uorstMain!serpercep.DefinedSize
      txtDetalle(6).MaxLength = .uorstMain!nropercep.DefinedSize
      txtDato(39).MaxLength = .uorstMain!codtdc_ref.DefinedSize
      txtDetalle(8).MaxLength = .uorstMain!serdoc_ref.DefinedSize
      txtDetalle(9).MaxLength = .uorstMain!nrodoc_ref.DefinedSize
      txtDetalle(10).MaxLength = 14
      txtDetalle(11).MaxLength = 14
      txtDetalle(12).MaxLength = 14
      txtDetalle(13).MaxLength = 14
      With cboPercepcion
        .AddItem Choose(gsIdioma, "Ninguna", "Neither"), 0
        .AddItem Choose(gsIdioma, "Combustibles - 1%", "Fuels - 1%"), 1
        .AddItem Choose(gsIdioma, "Importaciones - 10%", "Imports - 10%"), 2
        .AddItem Choose(gsIdioma, "Importaciones - 5%", "Imports - 5%"), 3
        .AddItem Choose(gsIdioma, "Importaciones - 3.5%", "Imports - 3.5%"), 4
        .AddItem Choose(gsIdioma, "Ventas Internas - 10%", "Internal Sales - 10%"), 5
        .AddItem Choose(gsIdioma, "Ventas Internas - 2%", "Internal Sales - 2%"), 6
        .AddItem Choose(gsIdioma, "Ventas Internas - 0.5%", "Internal Sales - 0.5%"), 7
      End With
      
'ini 2015-07-08 adic tabla detrac
       cboDetraccion.AddItem Choose(gsIdioma, "Ninguna", "Neither"), 0
       With frmTVtaGrd.uorstcodetrac
            '2015-07-27 error eof detra sin reg.MoveFirst
            If .RecordCount > 0 Then .MoveFirst
            If Not .EOF Then
                '.MoveFirst
                Do While Not .EOF
                      '2015-07-08 cambio de decima a % cboDetraccion.AddItem !coddetrac & " " & !detdetrac & " " & Trim(Str(!pctdetrac * 100)) & "%"
                     cboDetraccion.AddItem !coddetrac & " " & !detdetrac & " " & Trim(Str(!pctdetrac)) & "%"
                    .MoveNext
                Loop
            End If
        End With
'fin 2015-07-08 adic tabla detrac

    '2016-02-02.09  correccion ple gpcbo_sunat_ins cboCodMon, frmTVtaGrd.uorstCodMon '2016-02-02.06  correccion ple
    txtDato(43).MaxLength = .uorstMain!codmon.DefinedSize
          
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
  ReDim aLabel(37, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Tipo Doc.:", "NºDoc.:", "Rango Final:", "F.Operación :", "F.Emisión:", "F.Vcmto. :", "Cliente :", "Referencia:", "Glosa:", "Moneda:", "T.Cambio:", "Diario:", "Comprobante:", "Op. Gravada :", "Exportación :", "Exonerado :", "IGV :", "ISC :", "Otros :", "Total :", "Cuenta Contable", "Centro de Costo", "Tipo Operaci.:", "Asiento Tipo :", "Numero DUA :", "F. Embarque :", "F. Regularizar :", "Valor F.O.B. :", "Tipo Tasa :", "Nº Documento : ", "Tipo Doc. :", "NºDoc. :", "F.Emisión :", "Base Impo. :", "I.G.V. :", "Ord.Servicio :", "F.T.Cambio :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Type Doc.:", "NºDoc.:", "Final Range:", "Operti.Date:", "IssueDate :", "Due Date:", "Client :", "Reference :", "Gloss:", "Currency:", "R.Exchange:", "Journal:", "Voucher:", "Op. with Taxes :", "Export :", "Discharged :", "GST :", "SCT :", "Others :", "Total :", "Accountable Account", "Cost Center", "Operti. Type:", "Standar Recorded :", "Number DUA :", "Boarding Date :", "Regularize Date :", "Value F.O.B. :", "Type Tax :", "Nº Document : ", "Type Doc. :", "NºDoc. :", "IssueDate :", "Tax Basis :", "G.S.T. :", "Ord.Service :", "Exchange Date:")
  Next nElemento
  chkCalcularIGV.Caption = Choose(gsIdioma, "Calcular I.G.&V.", "Calculate G.&S.T.")
  chkCalcularISC.Caption = Choose(gsIdioma, "Calcular I.S.&C.", "Calculate S.&C.T.")
  
  chkIndCDt.Caption = Choose(gsIdioma, "Detracción", "Deduction") '2015-07-07 detrac vtas
  
  chkDesactivar.Caption = Choose(gsIdioma, "Des&activar Cuentas", "Dis&able Accounts")
  chkIndPreGen.Caption = Choose(gsIdioma, "Cuentas &Registradas", "&Registered Accounts")
  
  chkIndVtaext.Caption = Choose(gsIdioma, " Venta &Externa ", " &External Sale ")
  fraExterior.Caption = Choose(gsIdioma, " Venta Externa ", " External Sale ")
  fraPercepcion.Caption = Choose(gsIdioma, " Percepción ", " Perception ")
  fraReferencia.Caption = Choose(gsIdioma, " Documento de Referencia ", " Reference Document ")
  cmdAuxiliar.Caption = Choose(gsIdioma, "Cliente", "Client")
  cmdFormato.Caption = Choose(gsIdioma, "&Imprimir", "&Print")
  sstMain.TabCaption(0) = Choose(gsIdioma, "I&mportes", "A&mounts")
  sstMain.TabCaption(1) = Choose(gsIdioma, "C&uentas", "Acco&unts")
  sstMain.TabCaption(2) = Choose(gsIdioma, "&PDB-NC-ND", "&PDB-NC-ND")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']

'[Propio del formulario.
   dgrDetalle.MarqueeStyle = dbgHighlightRow
   Set dgrDetalle.DataSource = frmTVtaGrd.uorstCOCpbDet
   
   cboTpoMon.ListIndex = TPOMON_NAC_IND
   chkMonedaActiva.Value = vbChecked
   sstMain.Tab = 0
']
End Sub

Private Sub Form_Activate()
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
  If gbCieVta Then MsgBox TEXT_9016, vbCritical: Exit Sub
  
  pbCorregir = True
  frmTVtaGrd.uocnnMain.BeginTrans     'Cambiar Formulario de Grid. 'INICIA TRANSACCION.
  
  cmdRetroceder.Enabled = False
  cmdAvanzar.Enabled = False
  cmdCorregir.Enabled = False
  cmdFormato.Enabled = False
  cmdGrabar.Enabled = True
  cmdDeshacer.Enabled = True
  upHabilitacion True
  txtDato(0).Enabled = (chkIndPreGen.Value = 0)
  lblDatoDeta(0).Enabled = (chkIndPreGen.Value = 0)
  cmdDatoAyud(0).Enabled = (chkIndPreGen.Value = 0)
  
  ' Dato con el foco al corregir
  dtpDato(3).SetFocus
  ' Para no cambiar fechas
  pbFecha = False
  
End Sub

Public Sub cmdGrabar_Click()
  '[Propio del formulario.
  Dim dnSumaMN As Double, dnSumaME As Double

'ini 2015-07-07 detrac vtas
  If chkIndCDt.Value = 1 Then
    If cboDetraccion.ListIndex = 0 Then MsgBox TEXT_6002, vbCritical: cboDetraccion.SetFocus: Exit Sub
  End If
'fin 2015-07-07 detrac vtas

'   On Error GoTo Err
   
  If txtDato(33).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(33).SetFocus: Exit Sub
  If txtDato(0).Text = "" Then MsgBox TEXT_6002, vbCritical: txtDato(0).SetFocus: Exit Sub
  ' validacion de venta exterior
  If (chkIndVtaext.Value = vbChecked And (txtDetalle(0).Text = "" Or txtDetalle(1).Text = "" Or txtDetalle(2).Text = "")) Then MsgBox Choose(gsIdioma, "Ingrese el Documento de Venta Externa", "Enter the Document External Sale"), vbCritical: txtDetalle(0).SetFocus: Exit Sub
  If (chkIndVtaext.Value = vbChecked And CDec(txtDetalle(3).Text) <= 0) Then MsgBox Choose(gsIdioma, "Ingrese Importe de Venta Externa", "Enter Amount of Sale External"), vbCritical: txtDetalle(3).SetFocus: Exit Sub
  ' validacion de nota credito
  If ((txtLlave(0).Text = "07" Or txtLlave(0).Text = "08") And txtDato(39).Text = "" And (CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)) > 0) Then MsgBox Choose(gsIdioma, "Seleccione Tipo de Documento de Referencia", "Select Document Type Reference"), vbCritical: txtDato(39).SetFocus: Exit Sub
  If ((txtLlave(0).Text = "07" Or txtLlave(0).Text = "08") And txtDetalle(8).Text = "" And (CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)) > 0) Then MsgBox Choose(gsIdioma, "Debe ingresar Serie de Documento de Referencia", "You Must enter the Document Series Reference"), vbCritical: txtDetalle(8).SetFocus: Exit Sub
  If ((txtLlave(0).Text = "07" Or txtLlave(0).Text = "08") And txtDetalle(9).Text = "" And (CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text) + CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)) > 0) Then MsgBox Choose(gsIdioma, "Debe ingresar Documento de Referencia", "You Must enter the Document Reference"), vbCritical: txtDetalle(9).SetFocus: Exit Sub

  With frmTVtaGrd.uorstMain
    dnSumaMN = CDec(txtDato(5).Text) + CDec(txtDato(6).Text) + CDec(txtDato(7).Text) + CDec(txtDato(8).Text) + CDec(txtDato(9).Text) + CDec(txtDato(10).Text)
    dnSumaME = CDec(txtDato(12).Text) + CDec(txtDato(13).Text) + CDec(txtDato(14).Text) + CDec(txtDato(15).Text) + CDec(txtDato(16).Text) + CDec(txtDato(17).Text)
    If dnSumaMN <> CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text) Then
      If (cboTpoMon.ListIndex = TPOMON_EXT_IND And cmdMas(CANTIDADIMPORTES).Tag <> INDMASCTA_MAS And (Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text)) <= 0.02)) Then
        If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text)))) & "." & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
          txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text = Format(dnSumaMN, FORMATO_NUM_1)
        Else
          If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(11).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
        End If
    ElseIf dnSumaME <> CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text) Then
      If (cboTpoMon.ListIndex = TPOMON_NAC_IND And cmdMas(CANTIDADIMPORTES).Tag <> INDMASCTA_MAS And (Abs(dnSumaME - CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)) <= 0.02)) Then
        If MsgBox(TEXT_9012 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaME - CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)))) & "." & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
        txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text = Format(dnSumaME, FORMATO_NUM_1)
      Else
        If MsgBox(TEXT_9012 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaME - CDec(txtDato(18).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
      End If
    End If
  End With

  ' Genero las cuentas de acuerdo al asiento tipo
  If txtDato(38).Text <> "" And pbNuevo Then
    ppInsDelCtaCos txtDato(38).Text, INDCCO_INA
    ppInsDelCtaCos txtDato(38).Text, INDCCO_ACT
  End If

  ' Valido las Cuentas esten Correctas(llenas para todas los valores)
  If chkIndPreGen.Value = vbChecked Then
    'ini 2014-07-10 inhabilita y activa cuentas registradas
'    chkIndPreGen.Value = IIf(ValidaCtasCCo, 1, 0)
'    If Not (chkIndPreGen.Value = vbChecked) Then
'      If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
'    End If
    '************
    If ValidaCtasCCo = False Then
      If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
        Exit Sub
      Else
        chkIndPreGen.Value = 0
      End If
    End If
    '************
    'fin 2014-07-10 inhabilita y activa cuentas registradas
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
'    .uorstCCCfg.Update
    .uocnnMain.CommitTrans           'CONFIRMA TRANSACCION.
      
    '[Actualiza grid..
    .uorstMain_Grd.Requery
    .upDatosGrid
    .uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
    ']
      pbCorregir = False

    If pbNuevo Then
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
      cmdFormato.Enabled = True
      upHabilitacion False
    End If
  End With
      
'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
fEstMayUpd
'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
      
      
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
   cmdFormato.Enabled = True
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
   Case 0, 38, 40, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1, 33
      txtDato(Index).SetFocus
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
   Case 3, 4, 10, 11, 12, 13
    If CDec(txtDato(4).Text) <= 0 Then
      MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
      txtDato(4).SetFocus
      Exit Sub
    End If
      
    If CDec(txtDetalle(Index).Text) = 0 Then
      txtDetalle(Index).Text = Format(0, FORMATO_NUM_1)
      If (Index = 3 Or Index = 10 Or Index = 11) And cboTpoMon.ListIndex = TPOMON_EXT_IND Then
        txtDetalle(Index).Text = Format(CDec(txtDetalle(IIf(Index = 3, 4, IIf(Index = 10, 12, 13))).Text) * CDec(txtDato(4).Text), FORMATO_NUM_1)
      ElseIf (Index = 4 Or Index = 12 Or Index = 13) And cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDetalle(Index).Text = Format(CDec(txtDetalle(IIf(Index = 4, 3, IIf(Index = 12, 10, 11))).Text) / CDec(txtDato(4).Text), FORMATO_NUM_1)
      End If
    End If
    If chkMonedaActiva.Value = vbChecked Then
      If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
        txtDetalle(IIf(Index = 3, 4, IIf(Index = 10, 12, 13))).Text = Format(gfRedond(CDec(txtDetalle(Index).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
      Else
        txtDetalle(IIf(Index = 4, 3, IIf(Index = 12, 10, 11))).Text = Format(gfRedond(CDec(txtDetalle(Index).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
      End If
    End If
  End Select

End Sub

Private Sub txtDetalle_Validate(Index As Integer, Cancel As Boolean)
  
  'Completa con ceros a la izquierda.
  Select Case Index
   Case 0, 1, 2, 5, 6, 8, 9
    If Len(Trim(txtDetalle(Index).Text)) <> 0 And Len(Trim(txtDetalle(Index).Text)) <> txtDetalle(Index).MaxLength Then
      txtDetalle(Index) = gfCeros(txtDetalle(Index).Text, txtDetalle(Index).MaxLength, 0, "0")
    End If
   Case 3, 4, 10, 11, 12, 13
    txtDetalle(Index).Text = IIf(Not IsNumeric(txtDetalle(Index).Text), 0, txtDetalle(Index).Text)
    txtDetalle(Index).Text = Format(txtDetalle(Index).Text, FORMATO_NUM_1)
  End Select

End Sub

Private Sub txtllave_GotFocus(Index As Integer)
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
    'ini 2014-07-09 inhabilita y activa cuentas registradas
    chkIndPreGen.Value = 1 'activar chek
    chkIndPreGen.Enabled = False
    'fin 2014-07-09 inhabilita y activa cuentas registradas
   End If
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
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
         Set .uorstTemporal = .uocnnMain.Execute("SELECT MesPvs FROM COVtaDoc WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND CodTDc ='" & txtLlave(0).Text & "' AND SerDoc='" & txtLlave(1).Text & "' AND NroDoc='" & txtLlave(2).Text & "'")
         If .uorstTemporal.RecordCount > 0 Then
            MsgBox TEXT_8007 & Chr(13) & Choose(gsIdioma, "(mes ", "(month ") & gfMesLet("01" & .uorstTemporal!mespvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
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
  Select Case Index
   Case MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
    '///Angel 12/12/2003
    '/// Agregado para validar el T/C
    If CDec(txtDato(4).Text) <= 0 Then
      MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
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
              If frmTVtaGrd.uorstCOVtaDocCta!codtdc = txtLlave(0).Text And _
               frmTVtaGrd.uorstCOVtaDocCta!serdoc = txtLlave(1).Text And _
               frmTVtaGrd.uorstCOVtaDocCta!nrodoc = txtLlave(2).Text And _
               Trim(frmTVtaGrd.uorstCOVtaDocCta!tpocnc) = Trim(Str(Index - MINIMOINDICECUENTA + 1)) Then
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
    End If
   Case 38
    If txtDato(Index).Text <> "" And pbNuevo Then
      chkDesactivar.Value = vbChecked
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
   Case 0, 38, 40, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
    Cancel = ppAyuDet(AYUDAT, Index)
    If Cancel Then Exit Sub
'    If lblDatoDeta(Index).Caption <> "" Then
    If Index >= MINIMOINDICECUENTA And Index < MINIMOINDICECCOSTO Then
'      If frmTVtaGrd.uorstCOCta.RecordCount > 0 And txtDato(Index + CUENTASCONCCOSTO).Text <> "" Then
      If frmTVtaGrd.uorstCoCta.RecordCount > 0 Then
        If Not frmTVtaGrd.uorstCoCta.EOF Then
          If frmTVtaGrd.uorstCoCta!indcco = INDCCO_ACT Then
            ' Inicializo el centro de costos
            txtDato(Index + CUENTASCONCCOSTO).Tag = txtDato(Index + CUENTASCONCCOSTO).Text
            txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index).Tag <> txtDato(Index).Text, "", txtDato(Index + CUENTASCONCCOSTO).Text)
            If Not IsNull(frmTVtaGrd.uorstCoCta!codcco_def) Then
              txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index + CUENTASCONCCOSTO).Text = "", frmTVtaGrd.uorstCoCta!codcco_def, txtDato(Index + CUENTASCONCCOSTO).Text)
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

   '2016-02-02.09  correccion ple  Case 33                             'Cambiar (añadir índices).
   Case 33, 43                            'Cambiar (añadir índices).
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
      modAyuBus.Dro_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodDro)=4", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case MINIMOINDICECUENTA To MINIMOINDICECUENTA + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "TpoCta=" & TPOCTA_TRA & " AND EstCta='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      modAyuBus.CCo_Cod IIf(ps_Plataforma = pSrvMySql, "Length", "Len") & "(CodCCo)=5 AND EstCCo='" & ESTCTA_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 33                          'Cambiar (añadir índices).
      modAyuBus.Aux_Det "IndCli=1", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 38
      modAyuBus.Asi_Cod "tpoasi='" & TPOGNR_VTA & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAsiento.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAsiento.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 39                           'Cambiar (añadir índices).
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 40                             ' orden de servicio
      ' Filtro de seleccion
      cmdDatoAyud(tnIndex).Tag = "a.codaux = '" & txtDato(33).Text & "' "
      If ps_Plataforma = pSrvMySql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND a.fehcon<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
      ElseIf ps_Plataforma = pSrvSql Then
        cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND a.fehcon<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
      End If
      cmdDatoAyud(tnIndex).Tag = cmdDatoAyud(tnIndex).Tag & "AND a.tpognr='" & INDANU_FAL & "' "
      modAyuBus.Con_Sal cmdDatoAyud(tnIndex).Tag, txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'ini 2016-02-02.09  correccion ple
     Case 43   'codmon Cambiar (añadir índices).
      Dim xxCodTabla As String
      xxCodTabla = IIf(tnIndex = 43, CODSUNAT_004, xxCodTabla)
      modAyuBus.Sunat_Cod " estsunat ='" & ESTSUNAT_ACT & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + txtDato(tnIndex).Left, xxCodTabla
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
'fin 2016-02-02.09  correccion ple
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
      With frmTVtaGrd.uorstCODro
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
      With frmTVtaGrd.uorstCoCta
        If .RecordCount > 0 Then .MoveFirst
        .Find "CodCta='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
 
            'ini 2015-06-30 correccion tipo mon cta
            If tnIndex = 25 And !tpomon <> Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT) Then
                      MsgBox TEXT_9021, vbExclamation
                      ppAyuDet = True
            End If
            'fin 2015-06-30 correccion tipo mon cta
       
          '[ARREGLAR. Encontrar forma de mostrar el caption de los Label así no haya espacios en blanco.
          lblDatoDeta(tnIndex).Caption = " " & Left(!detcta, 18)
          ']ARREGLAR.
        End If
      End With
     Case MINIMOINDICECCOSTO To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTVtaGrd.uorstCoCCo
        If .RecordCount > 0 Then .MoveFirst
        .Find "CodCCo='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          '[ARREGLAR. Encontrar forma de mostrar el caption de los Label así no haya espacios en blanco.
          lblDatoDeta(tnIndex).Caption = " " & Left(!detcco, 12)
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
          lblDatoDeta(tnIndex).Caption = " " & !razAux
        End If
      End With
     Case 38
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTVtaGrd.uorstCoAsiTipo
        If .RecordCount > 0 Then .MoveFirst
        .Find "codasi='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detasi), "", !detasi)
        End If
      End With
     Case 39
      If txtDato(tnIndex).Text = "" Then
        lblDatoDeta(tnIndex).Caption = ""
        Exit Function
      End If
      With frmTVtaGrd.uorstTGTDc
        If .RecordCount > 0 Then .MoveFirst
        .Find "CodTDc='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!dettdc), "", !dettdc)
        End If
      End With
       Case 43
         If txtDato(tnIndex).Text = "" Then
            lblDatoDeta(tnIndex).Caption = ""
            Exit Function
         End If
        Dim zxrst As Recordset
        Set zxrst = IIf(tnIndex = 43, frmTVtaGrd.uorstCodMon, zxrst)
         If gf_tb_sunat_seek(zxrst, txtDato(tnIndex).Text, lblDatoDeta(tnIndex)) Then
           MsgBox TEXT_8006, vbExclamation
           ppAyuDet = True
         Else
           'lblDatoDeta(tnIndex).Caption = " " & !razAux
         End If
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
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !codtdc = txtLlave(0).Text
        !serdoc = txtLlave(1).Text
        !nrodoc = txtLlave(2).Text
        !mespvs = gsMesAct
        !PctIGV = CDec(gnPctIGV)
        !PctISC = CDec(gnPctISC)
      End If

      'Datos.
      !tpomon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
      !indpregen = IIf(chkIndPreGen.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      
'ini 2015-07-08 adic tabla detrac
      '!tsadetrac = IIf(cboDetraccion.ListIndex = 0, Null, Left(cboDetraccion.Text, 5))
         With frmTVtaGrd.uorstcodetrac
            If .RecordCount > 0 Then .MoveFirst
            .Find "coddetrac='" & Left(cboDetraccion.Text, 5) & "'"
            If .EOF Then
               'MsgBox TEXT_8006, vbExclamation
               'ppAyuDet = True
            Else
               'lblLlaveDeta(tnIndex).Caption = " " & !razAux
               frmTVtaGrd.uorstMain!tsadetrac = IIf(cboDetraccion.ListIndex = 0, Null, !coddetrac)
               frmTVtaGrd.uorstMain!pctdetrac = IIf(cboDetraccion.ListIndex = 0, 0#, !pctdetrac)
           End If
         End With
'fin 2015-07-08 adic tabla detrac

       '2016-02-02.09  correccion ple  gpcbo_sunat_update cboCodMon, frmTVtaGrd.uorstCodMon, "CodMon", 3, frmTVtaGrd.uorstMain '2016-02-02.06  correccion ple
      '2016-02-02.09  correccion ple !CodMon = IIf(txtDato(63).Text = "", Null, txtDato(63).Text) '2016-02-02.08  correccion ple
      !codmon = IIf(txtDato(43).Text = "", Null, txtDato(43).Text) '2016-02-02.08  correccion ple

      
'ini 2015-07-07 detrac vtas
     !indcdt = IIf(chkIndCDt.Value = vbChecked, INDCDT_ACT, INDCDT_INA)
      !FehCDt = dtpDato(4).Value
      !NroCDt = txtDato(42).Text
'fin 2015-07-07 detrac vtas

      !fehope = dtpDato(3).Value
      !feedoc_ref = dtpDato(2).Value
      !feedoc = dtpDato(0).Value
      !fevdoc = dtpDato(1).Value
      !codcon = IIf(txtDato(40).Text = "", Null, txtDato(40).Text)
      !coddro = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      !NroCpb = txtDato(1).Text
      !RefDoc = txtDato(2).Text
      !GloDoc = IIf(txtDato(Choose(gsIdioma, 3, 37)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 37)).Text)
      !glodocx = IIf(txtDato(Choose(gsIdioma, 37, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 37, 3)).Text)
      !codasi = IIf(txtDato(38).Text = "", Null, txtDato(38).Text)
      !ImpTCb = CDec(txtDato(4).Text)
      !impogr_mn = CDec(txtDato(5).Text)
      !ImpExp_mn = CDec(txtDato(6).Text)
      !impexo_mn = CDec(txtDato(7).Text)
      !impigv_mn = CDec(txtDato(8).Text)
      !impisc_mn = CDec(txtDato(9).Text)
      !impoim_mn = CDec(txtDato(10).Text)
      !imptot_mn = CDec(txtDato(11).Text)
      !impogr_me = CDec(txtDato(12).Text)
      !impexp_me = CDec(txtDato(13).Text)
      !impexo_me = CDec(txtDato(14).Text)
      !impigv_me = CDec(txtDato(15).Text)
      !impisc_me = CDec(txtDato(16).Text)
      !impoim_me = CDec(txtDato(17).Text)
      !imptot_me = CDec(txtDato(18).Text)
      !codaux = IIf(txtDato(33).Text = "", Null, txtDato(33).Text)
      !SerDoc_Fin = txtDato(34).Text
      !nrodoc_fin = txtDato(35).Text
      !TpoGlo_Rtc = cboRetencion.ListIndex
      !glodoc_rtc = IIf(txtDato(36).Text = "" Or cboRetencion.ListIndex = TPOGRU1_IND, Null, txtDato(36).Text)
      ' Datos adicionales
      !tpoimpuesto = cboImpuesto.ListIndex
      !categoriadoc = cboCategoria.ListIndex
      
      ' Informacion pdb
      !indvtaext = IIf(chkIndVtaext.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      !codaduana = IIf(txtDetalle(0).Text = "", Null, txtDetalle(0).Text)
      !annodua = IIf(txtDetalle(1).Text = "", Null, txtDetalle(1).Text)
      !nrodua = IIf(txtDetalle(2).Text = "", Null, txtDetalle(2).Text)
      !feembarq = dtpDetalle(0).Value
      !feregula = dtpDetalle(1).Value
      !impfob_mn = CDec(txtDetalle(3).Text)
      !impfob_me = CDec(txtDetalle(4).Text)
      chkIndpercep.Value = IIf(!indpercep = INDPREGEN_ACT, vbChecked, vbUnchecked)
      !tsapercep = cboPercepcion.ListIndex
      !serpercep = IIf(txtDetalle(5).Text = "", Null, txtDetalle(5).Text)
      !nropercep = IIf(txtDetalle(6).Text = "", Null, txtDetalle(6).Text)
      !codtdc_ref = IIf(txtDato(39).Text = "", Null, txtDato(39).Text)
      !serdoc_ref = IIf(txtDetalle(8).Text = "", Null, txtDetalle(8).Text)
      !nrodoc_ref = IIf(txtDetalle(9).Text = "", Null, txtDetalle(9).Text)
      dtpDetalle(2).Value = dtpDato(2).Value
      !feedoc_ref = dtpDetalle(2).Value
      !impbasref_mn = CDec(txtDetalle(10).Text)
      !impbasref_me = CDec(txtDetalle(12).Text)
      !impigvref_mn = CDec(txtDetalle(11).Text)
      !impigvref_me = CDec(txtDetalle(13).Text)
      
      '[Actualización de datos por Cuentas y Centros de Costo registrados directamente.
      ppAbreCtaCCo
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
        If cmdMas(dnContador).Tag = INDMASCTA_CTA And Val(txtDato(dnContador + DIFERENCIAMASCUENTA).Text) <> 1 Then
          With frmTVtaGrd.uorstCOVtaDocCta
            .MoveFirst
            .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & ps_OrdenCta & "'"
            If Not .EOF Then
              .Delete
              .Update
              .Requery
              frmTVtaGrd.uorstCOVtaDocCCo.Requery
              upActualizaMas dnContador, INDMASCTA_INI
            End If
          End With
        End If
            
        If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
          cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Then
          With frmTVtaGrd.uorstCOVtaDocCta
            If .RecordCount <> 0 Then .MoveFirst
              .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & ps_OrdenCta & "'"
              If .EOF Then
                .AddNew
                !codemp = gsCodEmp
                !pdoano = gsAnoAct
                !codtdc = txtLlave(0).Text
                !serdoc = txtLlave(1).Text
                !nrodoc = txtLlave(2).Text
                !tpocnc = dnContador
                !orden = ps_OrdenCta
                !UsrCre = gsAbvUsr
                !FyHCre = Now
              Else
                !UsrMdf = gsAbvUsr
                !FyHMdf = Now
              End If
              !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
              !glodet0 = txtDato(3).Text
              !impcta_mn = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE).Text)
              !impcta_me = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text)
              .Update
          End With
          If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Or _
             cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Then
            With frmTVtaGrd.uorstCOVtaDocCCo
              If .RecordCount <> 0 Then .MoveFirst
              .Find "cLlave1='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & ps_OrdenCta & txtDato(dnContador + DIFERENCIAMASCUENTA).Text & "'"
              If .EOF Then
                .AddNew
                !codemp = gsCodEmp
                !pdoano = gsAnoAct
                !codtdc = txtLlave(0).Text
                !serdoc = txtLlave(1).Text
                !nrodoc = txtLlave(2).Text
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
      txtLlave(0).Text = !codtdc
      txtLlave(1).Text = !serdoc
      txtLlave(2).Text = !nrodoc
      
      'Datos.
      cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      chkIndPreGen.Value = IIf(!indpregen = INDPREGEN_ACT, vbChecked, vbUnchecked)
      
'ini 2015-07-08 detrac vtas
      cboDetraccion.ListIndex = 0
      If Not IsNull(!tsadetrac) Then
        For dnContador = 1 To cboDetraccion.ListCount - 1
          If !tsadetrac = Left(cboDetraccion.List(dnContador), 5) Then
            cboDetraccion.ListIndex = dnContador
            Exit For
          End If
        Next dnContador
      End If
'ini 2015-07-08 detrac vtas
      
    '2016-02-02.09  correccion ple gpcbo_sunat_index cboCodMon, "CodMon", 3, frmTVtaGrd.uorstMain '2016-02-02.06  correccion ple
    txtDato(43).Text = IIf(IsNull(!codmon), "", !codmon) ' 2016-02-02.08  correccion ple
        ppAyuDet AYUDAT, 43
   
'ini 2015-07-07 detrac vtas
      chkIndCDt.Value = IIf(!indcdt = INDCDT_ACT, vbChecked, vbUnchecked)
      If Not IsNull(!FehCDt) Then '2015-07-08 error null
      dtpDato(4).Value = !FehCDt
      End If
      txtDato(42).Text = IIf(IsNull(!NroCDt), "", !NroCDt)
'fin 2015-07-07 detrac vtas
    
      '         dcoSocio.BoundText = IIf(IsNull(!CodSoc), "", !CodSoc)
      dtpDato(3).Value = !fehope
      dtpDato(2).Value = !feedoc_ref
      dtpDato(0).Value = !feedoc
      dtpDato(1).Value = !fevdoc
      '         optTpoMon(1).Value = uorstMain!CodMon
      '         mskDato(0).Text = IIf(IsNull(.uorstMain!Tf1Cta), "", .uorstMain!Tf1Cta)
      txtDato(40).Text = IIf(IsNull(!codcon), "", !codcon)
      txtDato(0).Text = IIf(IsNull(!coddro), "", !coddro)
      txtDato(1).Text = IIf(IsNull(!NroCpb), "", !NroCpb)
      txtDato(2).Text = IIf(IsNull(!RefDoc), "", !RefDoc)
      txtDato(Choose(gsIdioma, 3, 37)).Text = IIf(IsNull(!GloDoc), "", !GloDoc)
      txtDato(Choose(gsIdioma, 37, 3)).Text = IIf(IsNull(!glodocx), "", !glodocx)
      txtDato(38).Text = IIf(IsNull(!codasi), "", !codasi)
      txtDato(4).Text = Format(!ImpTCb, FORMATO_NUM_2)
      txtDato(5).Text = Format(!impogr_mn, FORMATO_NUM_1)
      txtDato(6).Text = Format(!ImpExp_mn, FORMATO_NUM_1)
      txtDato(7).Text = Format(!impexo_mn, FORMATO_NUM_1)
      txtDato(8).Text = Format(!impigv_mn, FORMATO_NUM_1)
      txtDato(9).Text = Format(!impisc_mn, FORMATO_NUM_1)
      txtDato(10).Text = Format(!impoim_mn, FORMATO_NUM_1)
      txtDato(11).Text = Format(!imptot_mn, FORMATO_NUM_1)
      txtDato(12).Text = Format(!impogr_me, FORMATO_NUM_1)
      txtDato(13).Text = Format(!impexp_me, FORMATO_NUM_1)
      txtDato(14).Text = Format(!impexo_me, FORMATO_NUM_1)
      txtDato(15).Text = Format(!impigv_me, FORMATO_NUM_1)
      txtDato(16).Text = Format(!impisc_me, FORMATO_NUM_1)
      txtDato(17).Text = Format(!impoim_me, FORMATO_NUM_1)
      txtDato(18).Text = Format(!imptot_me, FORMATO_NUM_1)
      For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
        txtDato(dnContador).Text = ""
        txtDato(dnContador).Tag = ""
      Next
      txtDato(33).Text = IIf(IsNull(!codaux), "", !codaux)
      txtDato(34).Text = IIf(IsNull(!SerDoc_Fin), "", !SerDoc_Fin)
      txtDato(35).Text = IIf(IsNull(!nrodoc_fin), "", !nrodoc_fin)
        
      cboRetencion.Tag = !TpoGlo_Rtc
      cboRetencion.ListIndex = !TpoGlo_Rtc
      txtDato(36).Text = IIf(IsNull(!glodoc_rtc), "", !glodoc_rtc)
      
      ' Datos adicionales
      cboImpuesto.ListIndex = !tpoimpuesto
      cboCategoria.ListIndex = !categoriadoc
      
      ' Informacion pdb
      chkIndVtaext.Value = IIf(!indvtaext = INDPREGEN_ACT, vbChecked, vbUnchecked)
      txtDetalle(0).Text = IIf(IsNull(!codaduana), "", !codaduana)
      txtDetalle(1).Text = IIf(IsNull(!annodua), "", !annodua)
      txtDetalle(2).Text = IIf(IsNull(!nrodua), "", !nrodua)
      dtpDetalle(0).Value = IIf(IsNull(!feembarq), !fehope, !feembarq)
      dtpDetalle(1).Value = IIf(IsNull(!feregula), !fehope, !feregula)
      txtDetalle(3).Text = Format(!impfob_mn, FORMATO_NUM_1)
      txtDetalle(4).Text = Format(!impfob_me, FORMATO_NUM_1)
      chkIndpercep.Value = IIf(!indpercep = INDPREGEN_ACT, vbChecked, vbUnchecked)
      cboPercepcion.ListIndex = !tsapercep
      txtDetalle(5).Text = IIf(IsNull(!serpercep), "", !serpercep)
      txtDetalle(6).Text = IIf(IsNull(!nropercep), "", !nropercep)
      txtDato(39).Text = IIf(IsNull(!codtdc_ref), "", !codtdc_ref)
      txtDetalle(8).Text = IIf(IsNull(!serdoc_ref), "", !serdoc_ref)
      txtDetalle(9).Text = IIf(IsNull(!nrodoc_ref), "", !nrodoc_ref)
      dtpDetalle(2).Value = IIf(IsNull(!feedoc_ref), !fehope, !feedoc_ref)
      txtDetalle(10).Text = Format(!impbasref_mn, FORMATO_NUM_1)
      txtDetalle(12).Text = Format(!impbasref_me, FORMATO_NUM_1)
      txtDetalle(11).Text = Format(!impigvref_mn, FORMATO_NUM_1)
      txtDetalle(13).Text = Format(!impigvref_me, FORMATO_NUM_1)
      ppAyuDet AYUDAT, 39
      
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
      ppAyuDet AYUDAT, 38
      ppAyuDet AYUDAT, 40
      ']
      
      '[Propio del formulario.
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
        cmdMas(dnContador).Tag = Choose(dnContador, !indcta_ogr, !indcta_exp, !indcta_exo, !indcta_igv, !indcta_isc, !indcta_oim, !indcta_tot)
      Next
        
      ' MA Obtengo las cuenta y centro de costo
      ppAbreCtaCCo
      With frmTVtaGrd.uorstCOVtaDocCta
        dnContador = 0
        While Not .EOF
          ' Cuenta por importe
          If dnContador <> CByte(!tpocnc) Then
            dnContador = CByte(!tpocnc)
            txtDato(dnContador + DIFERENCIAMASCUENTA).Text = !CodCta
            txtDato(dnContador + DIFERENCIAMASCUENTA).Tag = !CodCta
            ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCUENTA
            With frmTVtaGrd.uorstCOVtaDocCCo
              If .RecordCount > 0 Then
                .MoveFirst
                .Find "cLlave = " & dnContador & ps_OrdenCta & frmTVtaGrd.uorstCOVtaDocCta!CodCta
                If Not .EOF Then
                  txtDato(dnContador + DIFERENCIAMASCCOSTO).Text = !codcco
                  ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCCOSTO
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
   txtLlave(2).Text = ""

  'Datos.
   cboTpoMon.ListIndex = TPOMON_NAC_IND
   chkIndPreGen.Value = vbUnchecked
'ini 2015-07-08 detrac vtas
   cboDetraccion.ListIndex = 0
'fin 2015-07-08 detrac vtas
    'cboCodMon.ListIndex = 0 '2016-02-02.06  correccion ple
    '2016-02-02.09  correccion ple  gpcbo_sunat_index2 cboCodMon, IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CODMON_NAC, CODMON_EXT), 3 '2016-02-02.06  correccion ple
    'txtDato(43).Text = "" 'fin 2016-02-02.03 correccion ple
     txtDato(43).Text = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CODMON_NAC, CODMON_EXT) '2016-02-02.06  correccion ple
         ppAyuDet AYUDAT, 43
    
'ini 2015-07-07 detrac vtas
   chkIndCDt.Value = vbUnchecked
   txtDato(42).Text = ""
   dtpDato(4).Value = Date
'fin 2015-07-07 detrac vtas
   dtpDato(2).Value = Date 'fin 2015-08-25 falta predetermindo
  
   dtpDato(3).Value = Date
   dtpDato(0).Value = Date
   dtpDato(1).Value = Date
'   optTpoMon(1).Value = True
   For dnContador = 0 To 3
      txtDato(dnContador).Text = ""
   Next
   txtDato(40).Text = ""
   txtDato(37).Text = ""
   txtDato(38).Text = ""
   txtDato(4).Text = Format(0, FORMATO_NUM_2)
   For dnContador = MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
      txtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
   Next
   For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
    txtDato(dnContador).Text = ""
    txtDato(dnContador).Tag = ""
   Next
   txtDato(33).Text = ""
   txtDato(34).Text = ""
   txtDato(35).Text = ""
   cboRetencion.Tag = gsTpoGlo_Rtc
   cboRetencion.ListIndex = gsTpoGlo_Rtc
   txtDato(36).Text = gsGloDoc_Rtc(gsTpoGlo_Rtc)
   ' Datos adicionales
   cboImpuesto.ListIndex = TipoImpuesto.Ninguno
   cboCategoria.ListIndex = CategoriaDocumento.Ninguno
   
   ' PDB
   chkIndVtaext.Value = vbUnchecked
   txtDetalle(0).Text = ""
   txtDetalle(1).Text = ""
   txtDetalle(2).Text = ""
   dtpDetalle(0).Value = Date
   dtpDetalle(1).Value = Date
   txtDetalle(3).Text = Format(0, FORMATO_NUM_1)
   txtDetalle(4).Text = Format(0, FORMATO_NUM_1)
   chkIndpercep.Value = vbUnchecked
   cboPercepcion.ListIndex = 0
   txtDetalle(5).Text = ""
   txtDetalle(6).Text = ""
   txtDato(39).Text = ""
   txtDetalle(8).Text = ""
   txtDetalle(9).Text = ""
   dtpDetalle(2).Value = Date
   txtDetalle(10).Text = Format(0, FORMATO_NUM_1)
   txtDetalle(11).Text = Format(0, FORMATO_NUM_1)
   txtDetalle(12).Text = Format(0, FORMATO_NUM_1)
   txtDetalle(13).Text = Format(0, FORMATO_NUM_1)
   lblDatoDeta(39).Caption = ""

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
   lblDatoDeta(38).Caption = ""
   lblDatoDeta(40).Caption = ""
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
   dtpDato(2).Enabled = tbHabilitar
   dtpDato(3).Enabled = tbHabilitar
   cboRetencion.Enabled = tbHabilitar
   With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
   End With
   
'ini 2015-07-08 detrac vtas
   cboDetraccion.Enabled = tbHabilitar
'fin 2015-07-08 detrac vtas

   '2016-02-02.09  correccion ple cboCodMon.Enabled = tbHabilitar '2016-02-02.06  correccion ple
    txtDato(43).Enabled = tbHabilitar 'fin 2016-02-02.03 correccion ple
   cmdDatoAyud(43).Enabled = tbHabilitar
   lblDatoDeta(43).Enabled = tbHabilitar
   
'ini 2015-07-07 detrac vtas
   chkIndCDt.Enabled = tbHabilitar
   dtpDato(4).Enabled = (tbHabilitar And chkIndCDt.Value)
   txtDato(42).Enabled = (tbHabilitar And chkIndCDt.Value)
'fin 2015-07-07 detrac vtas
   
   txtDato(34).Enabled = (txtLlave(0).Text = CODTDC_BOL Or txtLlave(0).Text = CODTDC_TIC) And tbHabilitar
   txtDato(35).Enabled = (txtLlave(0).Text = CODTDC_BOL Or txtLlave(0).Text = CODTDC_TIC) And tbHabilitar
   txtDato(38).Enabled = (tbHabilitar And pbNuevo)
   
  'Ayudas.
   cmdDatoAyud(40).Enabled = tbHabilitar
   lblDatoDeta(40).Enabled = tbHabilitar
   cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar
   cmdDatoAyud(33).Enabled = tbHabilitar
   lblDatoDeta(33).Enabled = tbHabilitar
   cmdDatoAyud(38).Enabled = (tbHabilitar And pbNuevo)
   lblDatoDeta(38).Enabled = (tbHabilitar And pbNuevo)
  
  ' Datos adicionales
  cboImpuesto.Enabled = tbHabilitar
  cboCategoria.Enabled = tbHabilitar
  
  ' PDB
   chkIndVtaext.Enabled = tbHabilitar
   txtDetalle(0).Enabled = tbHabilitar
   txtDetalle(1).Enabled = tbHabilitar
   txtDetalle(2).Enabled = tbHabilitar
   dtpDetalle(0).Enabled = tbHabilitar
   dtpDetalle(1).Enabled = tbHabilitar
   txtDetalle(3).Enabled = tbHabilitar
   txtDetalle(4).Enabled = tbHabilitar
   chkIndpercep.Enabled = tbHabilitar
   cboPercepcion.Enabled = tbHabilitar
   txtDetalle(5).Enabled = tbHabilitar
   txtDetalle(6).Enabled = tbHabilitar
   txtDato(39).Enabled = tbHabilitar
   txtDetalle(8).Enabled = tbHabilitar
   txtDetalle(9).Enabled = tbHabilitar
   dtpDetalle(2).Enabled = tbHabilitar
   txtDetalle(10).Enabled = tbHabilitar
   txtDetalle(11).Enabled = tbHabilitar
   txtDetalle(12).Enabled = tbHabilitar
   txtDetalle(13).Enabled = tbHabilitar
   cmdDatoAyud(39).Enabled = tbHabilitar
   lblDatoDeta(39).Enabled = tbHabilitar

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
'   If cboCodMon.ListIndex <> -1 Then
'    gpcbo_sunat_index2 cboCodMon, IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CODMON_NAC, CODMON_EXT), 3 '2016-02-02.06  correccion ple
'   End If

  ' If Trim(txtDato(43).Text) = "" Then
     txtDato(43).Text = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, CODMON_NAC, CODMON_EXT) '2016-02-02.06  correccion ple
         ppAyuDet AYUDAT, 43
  ' End If
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
   txtDato(38).Text = IIf((pbNuevo And chkDesactivar.Value = vbUnchecked), "", txtDato(38).Text)
   lblDatoDeta(38).Caption = IIf((pbNuevo And chkDesactivar.Value = vbUnchecked), "", lblDatoDeta(38).Caption)
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
  If Index = 0 Then
    If Month(dtpDato(Index).Value) <> Val(gsMesAct) Or Year(dtpDato(Index).Value) <> Val(gsAnoAct) Then
      MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
      dtpDato(Index).SetFocus
      Cancel = True
      Exit Sub
    End If
      
    If pbFecha Then
      With dtpDato
        For dnContador = 1 To .Count - 3
          .Item(dnContador).Value = dtpDato(Index).Value
        Next
      End With
      pbFecha = False
    End If
  End If
   
  If Index = 2 Then
    If Month(dtpDato(Index).Value) > Val(gsMesAct) And Year(dtpDato(Index).Value) >= Val(gsAnoAct) Then
      MsgBox Choose(gsIdioma, "La fecha No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
      dtpDato(Index).SetFocus
      Cancel = True
      Exit Sub
    End If
    dtpDato(Index).Tag = 0
    If (dtpDato(Index).Tag <> dtpDato(Index).Value) Then
      dtpDato(Index).Tag = dtpDato(Index).Value
      With frmTVtaGrd.uorstTGTCb
        If .RecordCount <> 0 Then .MoveFirst
        .Find "(FehTCb) = '" & Format(dtpDato(Index).Value, "yyyy/mm/dd") & "'"
        If .EOF Then
          MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
          Cancel = True
          Exit Sub
        Else
          txtDato(4).Text = Format(!ImpTCb_Vta, FORMATO_NUM_2)
        End If
      End With
    End If
    If pbFecha Then
      With dtpDato
        For dnContador = 0 To .Count - 3
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
          For dnContador = 0 To .Count - 3
            .Item(dnContador).Value = dtpDato(Index).Value
          Next
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
   dgrDetalle.SetFocus
End Sub

Private Sub ppAbreCtaCCo()
   With frmTVtaGrd.uorstCOVtaDocCCo
    frmTVtaGrd.usConnStrgWher_COVtaDocCCo = "WHERE COVtaDocCCo.codemp='" & frmTVtaGrd.uorstMain!codemp & "' AND COVtaDocCCo.pdoano='" & frmTVtaGrd.uorstMain!pdoano & "' "
    frmTVtaGrd.usConnStrgWher_COVtaDocCCo = frmTVtaGrd.usConnStrgWher_COVtaDocCCo & "AND COVtaDocCCo.CodTDc='" & frmTVtaGrd.uorstMain!codtdc & "' AND COVtaDocCCo.SerDoc='" & frmTVtaGrd.uorstMain!serdoc & "' "
    frmTVtaGrd.usConnStrgWher_COVtaDocCCo = frmTVtaGrd.usConnStrgWher_COVtaDocCCo & "AND COVtaDocCCo.NroDoc='" & frmTVtaGrd.uorstMain!nrodoc & "' "
    If .State = adStateOpen Then .Close
    .Source = frmTVtaGrd.usConnStrgSele_COVtaDocCCo & frmTVtaGrd.usConnStrgWher_COVtaDocCCo & frmTVtaGrd.usConnStrgOrde_COVtaDocCCo
    .Open
    .Properties("Unique Table").Value = "COVtaDocCCo"
   End With
   With frmTVtaGrd.uorstCOVtaDocCta
    frmTVtaGrd.usConnStrgWher_COVtaDocCta = "WHERE COVtaDocCta.codemp='" & frmTVtaGrd.uorstMain!codemp & "' AND COVtaDocCta.pdoano='" & frmTVtaGrd.uorstMain!pdoano & "' "
    frmTVtaGrd.usConnStrgWher_COVtaDocCta = frmTVtaGrd.usConnStrgWher_COVtaDocCta & "AND COVtaDocCta.CodTDc='" & frmTVtaGrd.uorstMain!codtdc & "' AND COVtaDocCta.SerDoc='" & frmTVtaGrd.uorstMain!serdoc & "' "
    frmTVtaGrd.usConnStrgWher_COVtaDocCta = frmTVtaGrd.usConnStrgWher_COVtaDocCta & "AND COVtaDocCta.NroDoc='" & frmTVtaGrd.uorstMain!nrodoc & "' "
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
   ' PDB FOB, documento referencia
   txtDetalle(3).Visible = (unVerMonNac = TPOMON_NAC_IND)
   txtDetalle(4).Visible = Not (unVerMonNac = TPOMON_NAC_IND)
   txtDetalle(10).Visible = (unVerMonNac = TPOMON_NAC_IND)
   txtDetalle(11).Visible = (unVerMonNac = TPOMON_NAC_IND)
   txtDetalle(12).Visible = Not (unVerMonNac = TPOMON_NAC_IND)
   txtDetalle(13).Visible = Not (unVerMonNac = TPOMON_NAC_IND)

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
  sSentencia = "SELECT coddro, nrocpb"
  sSentencia = sSentencia & " FROM cocpbcab  "
  sSentencia = sSentencia & " WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & " AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & " AND mespvs='" & gsMesAct & "'"
  sSentencia = sSentencia & " AND coddro='" & txtDato(0).Text & "'"
  sSentencia = sSentencia & " AND nrocpb='" & txtDato(1).Text & "'"
  Set frmTVtaGrd.uorstTemporal = frmTVtaGrd.uocnnMain.Execute(sSentencia)
  If Not (frmTVtaGrd.uorstTemporal.BOF Or frmTVtaGrd.uorstTemporal.EOF) And frmTVtaGrd.uorstTemporal.RecordCount > 0 Then
    While Not frmTVtaGrd.uorstTemporal.EOF
            siexiste = True
            frmTVtaGrd.uorstTemporal.MoveNext
    Wend
  Else
  End If
  frmTVtaGrd.uorstTemporal.Close
  
  cuenta = 0
  sSentencia = "SELECT coddro,nrocpb "
  sSentencia = sSentencia & " FROM covtadoc  "
  sSentencia = sSentencia & " WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & " AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & " AND mespvs='" & gsMesAct & "'"
  sSentencia = sSentencia & " AND coddro='" & txtDato(0).Text & "'"
  sSentencia = sSentencia & " AND nrocpb='" & txtDato(1).Text & "'"
  Set frmTVtaGrd.uorstTemporal = frmTVtaGrd.uocnnMain.Execute(sSentencia)
  If Not (frmTVtaGrd.uorstTemporal.BOF Or frmTVtaGrd.uorstTemporal.EOF) And frmTVtaGrd.uorstTemporal.RecordCount > 0 Then
    While Not frmTVtaGrd.uorstTemporal.EOF
            cuenta = cuenta + 1
            frmTVtaGrd.uorstTemporal.MoveNext
    Wend
  Else
  End If
  frmTVtaGrd.uorstTemporal.Close
   
  If cuenta >= 2 Then masdedos = True


   If txtDato(1).Text <> "" Then
      ppDatosWhere
      If masdedos = False Then
      With frmTVtaGrd.uorstCOCpbCab
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
   frmTVtaGrd.uorstTGTDc.MoveFirst
   frmTVtaGrd.uorstTGTDc.Find "codtdc='" & txtLlave(0).Text & "'"
  
   With frmTVtaGrd.uorstCOCpbCab
     'Si no está marcado para generar, marca el documento como no generado.
      If chkIndPreGen.Value = vbUnchecked Then
         frmTVtaGrd.uorstMain!indgen = False
         frmTVtaGrd.uorstMain.Update
         Exit Sub
      End If

     'Captura del Siguiente Número.
      If txtDato(1).Text = "" Then
        txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
        txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
        txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
        txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
        txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
        frmTVtaGrd.uocnnMain.Execute txtDato(1).Tag
        ' Actualizo numero de comprobante tabla de detalle
        frmTVtaGrd.uorstMain!NroCpb = txtDato(1).Text
        frmTVtaGrd.uorstMain.Update
       Else
       If masdedos = True Then
        txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
        txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
        txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
        txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
        txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
        frmTVtaGrd.uocnnMain.Execute txtDato(1).Tag
        ' Actualizo numero de comprobante tabla de detalle
        frmTVtaGrd.uorstMain!NroCpb = txtDato(1).Text
        frmTVtaGrd.uorstMain.Update
        End If
    End If
      
      ppDatosWhere
   
     'Si no hay cuentas, marca el documento como no generado.
      If frmTVtaGrd.uorstCOVtaDocCta.RecordCount = 0 Then
         frmTVtaGrd.uorstMain!indgen = False
         frmTVtaGrd.uorstMain.Update
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
      !tpognr = TPOGNR_VTA
      !IndNCu = INDNCU_FAL
      !glocpb = IIf(txtDato(Choose(gsIdioma, 3, 37)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 37)).Text)
      !glocpbx = IIf(txtDato(Choose(gsIdioma, 37, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 37, 3)).Text)
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With
'[ Teo, Miguel Angel Refresco los recordset de cuentas y centros de costos
  frmTVtaGrd.uorstCOVtaDocCta.Requery
  frmTVtaGrd.uorstCOVtaDocCCo.Requery
']
   With frmTVtaGrd.uorstCOVtaDocCta
     'Crea ítemes de Comprobante.
      .MoveFirst
      Do
         dbProcesaCuenta = True

        'Itemes con Centro de Costo.
         If !tpocnc <= CUENTASCONCCOSTO Then
            With frmTVtaGrd.uorstCOVtaDocCCo
               If .RecordCount <> 0 Then
                  .MoveFirst
                  .Find "cLlave = " & Trim(frmTVtaGrd.uorstCOVtaDocCta!tpocnc) & frmTVtaGrd.uorstCOVtaDocCta!orden & frmTVtaGrd.uorstCOVtaDocCta!CodCta
                  If Not .EOF Then
                     Do
                        dnNumeroItem = dnNumeroItem + 1
                        ppGenera1 True, dnNumeroItem, IIf(CInt(frmTVtaGrd.uorstCOVtaDocCta!tpocnc) >= 4, "", txtDato(40).Text)
                        .MoveNext
                        If .EOF Then Exit Do
                        If !cLlave <> Trim(frmTVtaGrd.uorstCOVtaDocCta!tpocnc) & frmTVtaGrd.uorstCOVtaDocCta!orden & frmTVtaGrd.uorstCOVtaDocCta!CodCta Then Exit Do
                     Loop
                     dbProcesaCuenta = False
                  End If
               End If
            End With
         End If

        'Itemes sin Centro de Costo.
         If dbProcesaCuenta Then
            dnNumeroItem = dnNumeroItem + 1
            ppGenera1 False, dnNumeroItem, IIf(CInt(frmTVtaGrd.uorstCOVtaDocCta!tpocnc) >= 4, "", txtDato(40).Text)
         End If
         .MoveNext
      Loop Until .EOF
   End With

   frmTVtaGrd.uorstMain!indgen = True
   txtDato(0).Enabled = False
   txtDato(1).Enabled = False
   cmdDatoAyud(0).Enabled = False
   lblDatoDeta(0).Enabled = False

   frmTVtaGrd.uorstCOCpbCab.Update
   frmTVtaGrd.uorstCOCpbDet.UpdateBatch
   frmTVtaGrd.uorstMain.Update
End Sub

Private Sub ppGenera1(tbCCosto As Boolean, tnNumeroItem As Integer, ByVal sContrato As String)
   
  With frmTVtaGrd.uorstCOCpbDet
    .AddNew
    !codemp = gsCodEmp
    !pdoano = gsAnoAct
    !coddro = txtDato(0).Text
    !NroCpb = txtDato(1).Text
    !NroIte = tnNumeroItem
    !mespvs = gsMesAct
    !CodCta = frmTVtaGrd.uorstCOVtaDocCta!CodCta
    !fehope = dtpDato(3).Value
    frmTVtaGrd.uorstCoCta.MoveFirst
    frmTVtaGrd.uorstCoCta.Find "CodCta='" & frmTVtaGrd.uorstCOVtaDocCta!CodCta & "'"
    If frmTVtaGrd.uorstCoCta!indcco = INDCCO_ACT Then If tbCCosto Then !codcco = frmTVtaGrd.uorstCOVtaDocCCo!codcco
    If frmTVtaGrd.uorstCoCta!IndDoc = INDDOC_ACT Then
      !codaux = txtDato(33).Text
    Else
      If Len(Trim(frmTVtaGrd.uorstCOVtaDocCta!codruc)) > 0 Then
        !codaux = frmTVtaGrd.uorstCOVtaDocCta!codruc
      Else
        !codaux = txtDato(33).Text
      End If
    End If
    !codtdc = txtLlave(0).Text
    !serdoc = txtLlave(1).Text
    !nrodoc = txtLlave(2).Text
    !feedoc = dtpDato(0).Value
    !fevdoc = dtpDato(1).Value
    !ferdoc = dtpDato(0).Value
    !RefDoc = txtDato(2).Text
    !GloIte = Left(Trim(frmTVtaGrd.uorstCOVtaDocCta!glodet0), 60)
    !GloItex = Left(Trim(frmTVtaGrd.uorstCOVtaDocCta!glodet0x), 60)
    !codcon = IIf(sContrato = "", Null, sContrato)
    If tbCCosto Then
      If (frmTVtaGrd.uorstCOVtaDocCCo!impcco_me > 0) And (frmTVtaGrd.uorstCOVtaDocCCo!impcco_mn > 0) Then
        !TpoCtb = IIf(frmTVtaGrd.uorstCOVtaDocCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        !TpoCtb = IIf(frmTVtaGrd.uorstCOVtaDocCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
    Else
      If (frmTVtaGrd.uorstCOVtaDocCta!impcta_me > 0) And (frmTVtaGrd.uorstCOVtaDocCta!impcta_mn > 0) Then
        !TpoCtb = IIf(frmTVtaGrd.uorstCOVtaDocCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        !TpoCtb = IIf(frmTVtaGrd.uorstCOVtaDocCta!tpocnc = TPOCNC_TOT_VTA, IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
    End If
    !tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
    !ImpTCb = CDec(txtDato(4).Text)
    If tbCCosto Then
      !ImpMN = CDec(Abs(frmTVtaGrd.uorstCOVtaDocCCo!impcco_mn))
      !ImpME = CDec(Abs(frmTVtaGrd.uorstCOVtaDocCCo!impcco_me))
    Else
      !ImpMN = CDec(Abs(frmTVtaGrd.uorstCOVtaDocCta!impcta_mn))
      !ImpME = CDec(Abs(frmTVtaGrd.uorstCOVtaDocCta!impcta_me))
    End If
    'modificado tc
    '!TpoPvs = IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG And cboCategoria.ListIndex >= CategoriaDocumento.RetencionIva, TPOPVS_CAN, TPOPVS_PVS)
    !TpoPvs = IIf(frmTVtaGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG And cboCategoria.ListIndex >= CategoriaDocumento.RetencionIva, TPOPVS_PVS, TPOPVS_PVS)
    !tpognr = TPOGNR_VTA
    '!codmon = ""
    '2016-02-02.09  correccion ple gpcbo_sunat_update cboCodMon, frmTVtaGrd.uorstCodMon, "CodMon", 3, frmTVtaGrd.uorstCOCpbDet '2016-02-02.08  correccion ple
     !codmon = IIf(txtDato(43).Text = "", Null, txtDato(43).Text)
   
    !UsrCre = gsAbvUsr
    !FyHCre = Now
  End With

End Sub

Private Sub ppInsDelCtaCos(ByVal s_Asiento As String, ByVal n_TipoTran As Integer)
  Dim sSentencia As String, sTpoCnc As String, sOrden As String
  Dim nOrden As Long
  Dim nImporteMN As Double, nImporteME As Double
  Dim nImpoCtaMN As Double, nImpoCtaME As Double
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
    Set frmTVtaGrd.uorstTemporal = frmTVtaGrd.uocnnMain.Execute(sSentencia)
    If Not (frmTVtaGrd.uorstTemporal.BOF Or frmTVtaGrd.uorstTemporal.EOF) And frmTVtaGrd.uorstTemporal.RecordCount > 0 Then
      While Not frmTVtaGrd.uorstTemporal.EOF
        ' Inicializo el orden de cuenta
        nOrden = IIf(sTpoCnc = frmTVtaGrd.uorstTemporal!tpocnc, nOrden, 0)
        sTpoCnc = frmTVtaGrd.uorstTemporal!tpocnc
        nImporteMN = CDec(txtDato(Val(sTpoCnc) + 4).Text)
        nImporteME = CDec(txtDato(Val(sTpoCnc) + 11).Text)
        ' Inserto las cuentas por compra
        If nImporteMN <> 0 Or nImporteME <> 0 Then
          nOrden = nOrden + 1
          nImpoCtaMN = Round(nImporteMN * (CDec(frmTVtaGrd.uorstTemporal!pordst) / 100), 2)
          nImpoCtaME = Round(nImporteME * (CDec(frmTVtaGrd.uorstTemporal!pordst) / 100), 2)
          With frmTVtaGrd.uorstCOVtaDocCta    'Cambiar RecordSet.
            .AddNew
            'Llaves.
            
            !codemp = gsCodEmp
            !pdoano = gsAnoAct
            !codtdc = txtLlave(0).Text
            !serdoc = txtLlave(1).Text
            !nrodoc = txtLlave(2).Text
            !tpocnc = sTpoCnc
            !orden = Format(nOrden, "00")
            'Datos.
            !CodCta = frmTVtaGrd.uorstTemporal!CodCta
            !codruc = IIf(frmTVtaGrd.uorstTemporal!IndDoc = INDDOC_ACT, txtDato(33).Text, Null)
            !glodet0 = IIf(txtDato(Choose(gsIdioma, 3, 37)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 37)).Text)
            !glodet0x = IIf(txtDato(Choose(gsIdioma, 37, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 37, 3)).Text)
            !impcta_mn = CDec(nImpoCtaMN)
            !impcta_me = CDec(nImpoCtaME)
            !UsrCre = gsAbvUsr
            !FyHCre = Now
            .Update
          End With
          cmdMas(Val(sTpoCnc)).Tag = INDMASCTA_MAS
          ' Cuenta incial actualiza texto de costos
          If nOrden = 1 Then
            txtDato(Val(sTpoCnc) + (MINIMOINDICECUENTA - 1)).Text = frmTVtaGrd.uorstTemporal!CodCta
            cmdMas(Val(sTpoCnc)).Tag = INDMASCTA_INI
          End If
          upActualizaMas Val(sTpoCnc), cmdMas(Val(sTpoCnc)).Tag
          
          ' Inserto el centro de costo
          If ((Not IsNull(frmTVtaGrd.uorstTemporal!codcco)) And frmTVtaGrd.uorstTemporal!indcco = INDCCO_ACT) Then
            With frmTVtaGrd.uorstCOVtaDocCCo    'Cambiar RecordSet.
              .AddNew
              'Llaves.
              !codemp = gsCodEmp
              !pdoano = gsAnoAct
              !codtdc = txtLlave(0).Text
              !serdoc = txtLlave(1).Text
              !nrodoc = txtLlave(2).Text
              !tpocnc = sTpoCnc
              !orden = Format(nOrden, "00")
              !CodCta = frmTVtaGrd.uorstTemporal!CodCta
              'Datos.
              !codcco = frmTVtaGrd.uorstTemporal!codcco
              !impcco_mn = CDec(nImpoCtaMN)
              !impcco_me = CDec(nImpoCtaME)
              !UsrCre = gsAbvUsr
              !FyHCre = Now
              .Update
            End With
            ' Cuenta incial actualiza texto de costos
            If nOrden = 1 Then
              txtDato(Val(sTpoCnc) + (MINIMOINDICECCOSTO - 1)).Text = frmTVtaGrd.uorstTemporal!codcco
            End If
          End If
        End If
        frmTVtaGrd.uorstTemporal.MoveNext
      Wend
    End If
    frmTVtaGrd.uorstTemporal.Close
  ElseIf n_TipoTran = INDCCO_INA Then
    sSentencia = "DELETE FROM CoVtaDocCta "
    sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND codtdc='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND serdoc='" & txtLlave(1).Text & "' "
    sSentencia = sSentencia & "AND nrodoc='" & txtLlave(2).Text & "'"
    frmTVtaGrd.uocnnMain.Execute sSentencia
  End If
  ppAbreCtaCCo

End Sub

Public Sub upActualizaMas(pnIndice As Byte, pnValor As Byte)
   frmTVta.cmdMas(pnIndice).Tag = pnValor 'Necesaria la referencia por ser llamado externamente.
   With frmTVtaGrd.uorstMain
      Select Case pnIndice
      Case 1
         !indcta_ogr = pnValor
      Case 2
         !indcta_exp = pnValor
      Case 3
         !indcta_exo = pnValor
      Case 4
         !indcta_igv = pnValor
      Case 5
         !indcta_isc = pnValor
      Case 6
         !indcta_oim = pnValor
      Case 7
         !indcta_tot = pnValor
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
      frmTVtaGrd.uorstCoCta.MoveFirst
      frmTVtaGrd.uorstCoCta.Find "CodCta='" & txtDato(tnIndice + DIFERENCIAMASCUENTA).Text & "'"
      txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTVtaGrd.uorstCoCta!indcco = INDCCO_ACT)
      lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTVtaGrd.uorstCoCta!indcco = INDCCO_ACT)
      cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTVtaGrd.uorstCoCta!indcco = INDCCO_ACT)
   End If
End Sub

Private Sub ppDatosWhere()             'Cambiar.
  With frmTVtaGrd
    .uorstCOCpbCab.Requery
    
    .usConnStrgWher_COCpbDet = "WHERE COCpbDet.codemp='" & gsCodEmp & "' AND COCpbDet.pdoano='" & gsAnoAct & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='" & txtDato(0).Text & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND COCpbDet.NroCpb='" & txtDato(1).Text & "' "
    '2016-02-02.08  correccion ple
    'del grid frmtvtagrid.proc=load viene con el selec de la venta,
    'pero aqui recien retoma valores del cocpbdet
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
            .Item(dnNum).Caption = Choose(gsIdioma, "Cuenta", "Account")
            .Item(dnNum).Width = 900
         Case 1
            .Item(dnNum).Caption = Choose(gsIdioma, "Auxiliar", "Auxiliary")
            .Item(dnNum).Width = 1150
         Case 2
            .Item(dnNum).Caption = Choose(gsIdioma, "C.Cto.", "C.Center")
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
                      If Trim(!tpocnc) = Trim(Str(dnContador + 1)) Then
                         dnTotalCuentaMN = dnTotalCuentaMN + !impcta_mn
                         dnTotalCuentaME = dnTotalCuentaME + !impcta_me
                         With frmTVtaGrd.uorstCoCta
                            .MoveFirst
                            .Find "CodCta='" & frmTVtaGrd.uorstCOVtaDocCta!CodCta & "'"
                            If Not .EOF Then
                               dnIndCCo = frmTVtaGrd.uorstCoCta!indcco
                            End If
                         End With
                      End If
                      If dnIndCCo = INDCCO_ACT Then
                         With frmTVtaGrd.uorstCOVtaDocCCo
                            If .State = adStateOpen Then .Close
                            frmTVtaGrd.usConnStrgWher_COVtaDocCCo = "WHERE COVtaDocCCo.SerDoc='" & frmTVtaGrd.uorstMain!serdoc & "' And COVtaDocCCo.NroDoc='" & frmTVtaGrd.uorstMain!nrodoc & "' And COVtaDocCCo.TpoCnc='" & Trim(Str(dnContador + 1)) & "' And COVtaDocCCo.CodCta='" & frmTVtaGrd.uorstCOVtaDocCta!CodCta & "' "
                            .Source = frmTVtaGrd.usConnStrgSele_COVtaDocCCo & frmTVtaGrd.usConnStrgWher_COVtaDocCCo & frmTVtaGrd.usConnStrgOrde_COVtaDocCCo
                            .Open
                            If .RecordCount = 0 Then
                               MsgBox Choose(gsIdioma, "Cuenta ", "Account ") & frmTVtaGrd.uorstCOVtaDocCta!CodCta & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
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
          With frmTVtaGrd.uorstCoCta
             .MoveFirst
             .Find "CodCta='" & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & "'"
             If Not .EOF Then
                dnIndCCo = frmTVtaGrd.uorstCoCta!indcco
             End If
          End With
          If dnIndCCo = INDCCO_ACT And txtDato(dnContador + 1 + DIFERENCIAMASCUENTA + CUENTASCONCCOSTO).Text = "" Then
             MsgBox Choose(gsIdioma, "Cuenta ", "Account ") & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
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


