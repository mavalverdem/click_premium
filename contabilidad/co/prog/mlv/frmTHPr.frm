VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTHPr 
   Caption         =   "[Título]"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7470
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
      Index           =   28
      Left            =   1140
      TabIndex        =   34
      Top             =   3930
      Width           =   1250
   End
   Begin VB.CommandButton cmdDatoAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   300
      Index           =   28
      Left            =   7995
      Picture         =   "frmTHPr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   3930
      Width           =   255
   End
   Begin VB.CommandButton cmdPedido 
      Caption         =   "P&edido"
      Height          =   375
      Left            =   8265
      TabIndex        =   114
      Top             =   3225
      Width           =   1215
   End
   Begin VB.Frame fraPedido 
      Height          =   540
      Left            =   0
      TabIndex        =   19
      Top             =   2070
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
         Height          =   315
         Index           =   27
         Left            =   1005
         TabIndex        =   21
         Top             =   165
         Width           =   1410
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   27
         Left            =   6825
         Picture         =   "frmTHPr.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   165
         Width           =   255
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   18
         Left            =   105
         TabIndex        =   20
         Top             =   180
         Width           =   870
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
         Left            =   2415
         TabIndex        =   22
         Top             =   165
         Width           =   4410
      End
   End
   Begin VB.Frame fraAsiento 
      Height          =   540
      Left            =   0
      TabIndex        =   29
      Top             =   3210
      Width           =   7155
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   26
         Left            =   6795
         Picture         =   "frmTHPr.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   165
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Height          =   280
         Index           =   26
         Left            =   1560
         TabIndex        =   31
         Top             =   165
         Width           =   560
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
         Left            =   2100
         TabIndex        =   32
         Top             =   165
         Width           =   4695
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   17
         Left            =   105
         TabIndex        =   30
         Top             =   180
         Width           =   990
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
      Index           =   25
      Left            =   3840
      TabIndex        =   14
      Top             =   1740
      Width           =   4575
   End
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "&Proveedor"
      Height          =   375
      Left            =   8265
      TabIndex        =   113
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkAfeIndORt 
      Caption         =   "Afecto O&tras Ret."
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   5685
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4425
      Width           =   2640
   End
   Begin VB.CheckBox chkAfeIndIES 
      Caption         =   "Afecto I.&E.S."
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   3000
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4425
      Width           =   1395
   End
   Begin VB.CheckBox chkIndAfeIR4 
      Caption         =   "Afecto I.&R.4ª"
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   60
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4425
      Width           =   1665
   End
   Begin VB.CheckBox chkIndPreGen 
      Caption         =   "Cuentas Re&gistradas"
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   6960
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4770
      Width           =   1815
   End
   Begin VB.CheckBox chkCalcularIR4 
      Caption         =   "C&alcular"
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   1725
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4425
      Width           =   1090
   End
   Begin VB.CheckBox chkCalcularIES 
      Caption         =   "Ca&lcular"
      ForeColor       =   &H80000002&
      Height          =   200
      Left            =   4455
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4425
      Width           =   1090
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   2625
      Width           =   7155
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4935
         Picture         =   "frmTHPr.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   86
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
         Left            =   675
         TabIndex        =   25
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
         TabIndex        =   28
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
         Left            =   1215
         TabIndex        =   26
         Top             =   180
         Width           =   3735
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   9
         Left            =   5265
         TabIndex        =   27
         Top             =   240
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   8
         Left            =   90
         TabIndex        =   24
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
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   120
      Width           =   885
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
         Picture         =   "frmTHPr.frx":06A8
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   60
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
         Left            =   98
         Picture         =   "frmTHPr.frx":0852
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   60
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
         Left            =   98
         Picture         =   "frmTHPr.frx":09FC
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   325
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
         Picture         =   "frmTHPr.frx":0B46
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   880
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
         Picture         =   "frmTHPr.frx":0C48
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1440
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
         Left            =   98
         Picture         =   "frmTHPr.frx":0D4A
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1990
         Width           =   720
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
      TabIndex        =   18
      Top             =   1740
      Width           =   735
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7680
      Picture         =   "frmTHPr.frx":0E94
      Style           =   1  'Graphical
      TabIndex        =   84
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
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox cboTpoMon 
      Height          =   315
      ItemData        =   "frmTHPr.frx":103E
      Left            =   1020
      List            =   "frmTHPr.frx":1040
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1740
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   3060
      TabIndex        =   9
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57737217
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
      TabIndex        =   13
      Top             =   1380
      Width           =   4575
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
      TabIndex        =   11
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
      Index           =   2
      Left            =   1410
      TabIndex        =   5
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
      Index           =   1
      Left            =   900
      TabIndex        =   4
      Top             =   480
      Width           =   525
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   2655
      Left            =   0
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   4755
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   441
      ForeColor       =   8388608
      TabCaption(0)   =   "I&mportes"
      TabPicture(0)   =   "frmTHPr.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(16)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTexto(15)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTexto(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTexto(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTexto(13)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTexto(11)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTexto(14)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDatoDeta(15)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDatoDeta(16)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDatoDeta(17)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDatoDeta(18)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDatoDeta(19)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDatoDeta(20)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDatoDeta(21)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDatoDeta(22)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDatoDeta(23)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDatoDeta(24)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkDesactivar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdDatoAyud(15)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDato(15)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtDato(16)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdDatoAyud(16)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdDatoAyud(17)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDato(17)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdDatoAyud(18)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtDato(18)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdDatoAyud(19)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtDato(19)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdDatoAyud(20)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtDato(20)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdMas(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdMas(2)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdMas(3)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdMas(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdMas(5)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "chkMonedaActiva"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtDato(21)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmdDatoAyud(21)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtDato(22)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdDatoAyud(22)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtDato(23)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdDatoAyud(23)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtDato(24)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdDatoAyud(24)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtDato(10)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtDato(11)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtDato(12)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtDato(13)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtDato(5)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtDato(6)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtDato(7)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtDato(8)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtDato(14)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtDato(9)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).ControlCount=   54
      TabCaption(1)   =   "C&uentas"
      TabPicture(1)   =   "frmTHPr.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dgrDetalle"
      Tab(1).ControlCount=   2
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
         TabIndex        =   64
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
         Index           =   14
         Left            =   1320
         TabIndex        =   65
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
         TabIndex        =   59
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
         TabIndex        =   54
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
         TabIndex        =   49
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
         TabIndex        =   44
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
         Index           =   13
         Left            =   1320
         TabIndex        =   60
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
         Index           =   12
         Left            =   1320
         TabIndex        =   55
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
         Index           =   11
         Left            =   1320
         TabIndex        =   50
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
         Index           =   10
         Left            =   1320
         TabIndex        =   45
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   24
         Left            =   9060
         Picture         =   "frmTHPr.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   111
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
         Index           =   24
         Left            =   6840
         TabIndex        =   68
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   23
         Left            =   9060
         Picture         =   "frmTHPr.frx":1224
         Style           =   1  'Graphical
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   1530
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
         Left            =   6840
         TabIndex        =   63
         Top             =   1500
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   22
         Left            =   9060
         Picture         =   "frmTHPr.frx":13CE
         Style           =   1  'Graphical
         TabIndex        =   107
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
         Index           =   22
         Left            =   6840
         TabIndex        =   58
         Top             =   1200
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   21
         Left            =   9060
         Picture         =   "frmTHPr.frx":1578
         Style           =   1  'Graphical
         TabIndex        =   105
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
         Index           =   21
         Left            =   6840
         TabIndex        =   53
         Top             =   900
         Width           =   675
      End
      Begin VB.CheckBox chkMonedaActiva 
         Caption         =   "M&oneda activa"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   330
         Width           =   1635
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   -70200
         ScaleHeight     =   270
         ScaleWidth      =   1575
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   60
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
            TabIndex        =   103
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
            TabIndex        =   102
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
         Index           =   5
         Left            =   3000
         Picture         =   "frmTHPr.frx":1722
         TabIndex        =   66
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
         Picture         =   "frmTHPr.frx":1824
         TabIndex        =   61
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
         Picture         =   "frmTHPr.frx":1926
         TabIndex        =   56
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
         Picture         =   "frmTHPr.frx":1A28
         TabIndex        =   51
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
         Picture         =   "frmTHPr.frx":1B2A
         TabIndex        =   46
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
         Index           =   20
         Left            =   6840
         TabIndex        =   48
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   20
         Left            =   9060
         Picture         =   "frmTHPr.frx":1C2C
         Style           =   1  'Graphical
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   630
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
         Index           =   19
         Left            =   3300
         TabIndex        =   67
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   19
         Left            =   6540
         Picture         =   "frmTHPr.frx":1DD6
         Style           =   1  'Graphical
         TabIndex        =   97
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
         Index           =   18
         Left            =   3300
         TabIndex        =   62
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   18
         Left            =   6540
         Picture         =   "frmTHPr.frx":1F80
         Style           =   1  'Graphical
         TabIndex        =   95
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
         Index           =   17
         Left            =   3300
         TabIndex        =   57
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   17
         Left            =   6540
         Picture         =   "frmTHPr.frx":212A
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1225
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   16
         Left            =   6540
         Picture         =   "frmTHPr.frx":22D4
         Style           =   1  'Graphical
         TabIndex        =   91
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
         Index           =   16
         Left            =   3300
         TabIndex        =   52
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
         Index           =   15
         Left            =   3300
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   15
         Left            =   6540
         Picture         =   "frmTHPr.frx":247E
         Style           =   1  'Graphical
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   625
         Width           =   255
      End
      Begin VB.CheckBox chkDesactivar 
         Caption         =   "Desactivar Cue&ntas"
         ForeColor       =   &H80000002&
         Height          =   200
         Left            =   4980
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   2160
         Left            =   -74880
         TabIndex        =   83
         Top             =   420
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3810
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
         Left            =   7500
         TabIndex        =   112
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
         Index           =   23
         Left            =   7500
         TabIndex        =   110
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
         Index           =   22
         Left            =   7500
         TabIndex        =   108
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
         Index           =   21
         Left            =   7500
         TabIndex        =   106
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
         Index           =   20
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
         Index           =   19
         Left            =   4260
         TabIndex        =   98
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
         Index           =   18
         Left            =   4260
         TabIndex        =   96
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
         Index           =   17
         Left            =   4260
         TabIndex        =   94
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
         Index           =   16
         Left            =   4260
         TabIndex        =   92
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
         Index           =   15
         Left            =   4260
         TabIndex        =   90
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "IMP. NETO:"
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
         Left            =   195
         TabIndex        =   82
         Top             =   1860
         Width           =   780
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Ret. 4ª Cat.:"
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
         Left            =   180
         TabIndex        =   81
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Otras Ret.:"
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
         Left            =   180
         TabIndex        =   80
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Ret. I.E.S.:"
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
         Left            =   180
         TabIndex        =   79
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Importe Bruto:"
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
         Left            =   180
         TabIndex        =   78
         Top             =   660
         Width           =   1005
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   15
         Left            =   3420
         TabIndex        =   77
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
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   16
         Left            =   6960
         TabIndex        =   76
         Top             =   345
         Width           =   2265
      End
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   3
      Left            =   1020
      TabIndex        =   7
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57737217
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
      Index           =   19
      Left            =   120
      TabIndex        =   33
      Top             =   3960
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
      Index           =   28
      Left            =   2370
      TabIndex        =   35
      Top             =   3930
      Width           =   5625
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   480
      Left            =   0
      Top             =   3840
      Width           =   8385
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   6
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
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   120
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   5
      Left            =   3360
      TabIndex        =   12
      Top             =   1440
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   4
      Left            =   60
      TabIndex        =   10
      Top             =   1440
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   6
      Left            =   60
      TabIndex        =   15
      Top             =   1800
      Width           =   615
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   7
      Left            =   1860
      TabIndex        =   17
      Top             =   1800
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   3
      Left            =   2340
      TabIndex        =   8
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "NºDocum.:"
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
      Top             =   540
      Width           =   765
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
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmTHPr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2016-03-02 controla si tipo de cambio
'0=f.emision 1=otros clientes f.operacion
Private Const pval_tpo_cam As Integer = 0


Private pbNuevo As Boolean
Private pbCorregir As Boolean
Private pbValidada As Boolean
Private pbFecha As Boolean
'[Propio del formulario.
Public unVerMonNac As Byte
Private Const MINIMOINDICEIMPORTEMN As Byte = 5, _
              MINIMOINDICEIMPORTEME As Byte = 10, _
              MINIMOINDICEMAS As Byte = 1, _
              MINIMOINDICECUENTA As Byte = 15, _
              MINIMOINDICECCOSTO As Byte = 20, _
              CANTIDADIMPORTES As Byte = 5
'[Repetir en frmTHPrMasGrd.
Private Const DIFERENCIAMASIMPORTE As Byte = 4, _
              DIFERENCIAMASCUENTA As Byte = 14, _
              DIFERENCIAMASCCOSTO As Byte = 19
Private Const CUENTASCONCCOSTO As Byte = 5
']
'[Repetir en frmTHPrGrd y frmTHPrMasGrd.
Private Const INDMASCTA_INI As Byte = 0, _
              INDMASCTA_MAS As Byte = 1, _
              INDMASCTA_CTA As Byte = 2
']
Private Const ps_OrdenCta As String = "01"
Private s_PedidoValiDsc As String
']

Private Sub cmdPedido_Click()
  frmTPdoGrd.Show vbModal
End Sub

Private Sub Form_Load()
   pbValidada = False
   pbFecha = True
   Me.KeyPreview = True
   
   With frmTHPrGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
      txtLlave(0).MaxLength = .uorstMain!codaux.DefinedSize
      txtLlave(1).MaxLength = .uorstMain!serdoc.DefinedSize
      txtLlave(2).MaxLength = .uorstMain!nrodoc.DefinedSize
    ']
   
    '[Datos                            'Cambiar.
      With cboTpoMon
         .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
         .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
      End With
'      mskDato(0).MaxLength = .uorstMain!Tf1Cta.DefinedSize + 1
      txtDato(0).MaxLength = .uorstMain!coddro.DefinedSize
      txtDato(1).MaxLength = .uorstMain!NroCpb.DefinedSize
      txtDato(2).MaxLength = .uorstMain!RefDoc.DefinedSize
      txtDato(Choose(gsIdioma, 3, 25)).MaxLength = .uorstMain!GloDoc.DefinedSize
      txtDato(Choose(gsIdioma, 25, 3)).MaxLength = .uorstMain!glodocx.DefinedSize
      txtDato(27).MaxLength = .uorstMain!pdocpr.DefinedSize
      txtDato(26).MaxLength = .uorstMain!codasi.DefinedSize
      txtDato(28).MaxLength = .uorstMain!codcon.DefinedSize
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
      txtDato(15).MaxLength = 8
      txtDato(16).MaxLength = 8
      txtDato(17).MaxLength = 8
      txtDato(18).MaxLength = 8
      txtDato(19).MaxLength = 8
      txtDato(20).MaxLength = 5
      txtDato(21).MaxLength = 5
      txtDato(22).MaxLength = 5
      txtDato(23).MaxLength = 5
      txtDato(24).MaxLength = 5
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
  ReDim aLabel(20, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Proveedor:", "NºDocum.:", "F.Operación:", "F.Emisión:", "Referencia :", "Glosa :", "Moneda:", "T.Cambio:", "Diario:", "Comprobante:", "Importe Bruto :", "Ret. 4ta Cat.:", "Ret. I.E.S.:", "Otras Reten. :", "Importe Neto :", "Cuenta Contable", "Centro de Costo", "Asiento Tipo :", "Nro Pedido :", "Ord.Servicio :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Supplier :", "NºDocum.:", "Operti.Date:", "IssueDate:", "Reference :", "Gloss :", "Currency:", "R.Exchange:", "Journal:", "Voucher:", "Gross Amount :", "Withh.4th Class :", "Withh. E.T.S.:", "Others Withh. :", "Net Amount :", "Accountable Account", "Cost Center", "Standar Recorded :", "Nro Order :", "Ord.Service :")
  Next nElemento
  chkIndAfeIR4.Caption = Choose(gsIdioma, "Afecto I.&R. 4ta.", "Affect 4th &Inc.Tax")
  chkCalcularIR4.Caption = Choose(gsIdioma, "Calcular", "Calculate")
  chkAfeIndIES.Caption = Choose(gsIdioma, "Afecto I.&E.S.", "Affect &E.T.S.")
  chkCalcularIES.Caption = Choose(gsIdioma, "Calcular", "Calculate")
  'chkAfeIndORt.Caption = Choose(gsIdioma, "Afecto &Otras Reten.", "Affect Others Withholding")
  chkAfeIndORt.Caption = Choose(gsIdioma, "Afecto &Otras Reten.(AFP/ONP)", "Affect Others Withho.(AFP/ONP)")
  chkDesactivar.Caption = Choose(gsIdioma, "Des&activar Cuentas", "Dis&able Accounts")
  chkIndPreGen.Caption = Choose(gsIdioma, "Cuentas &Registradas", "&Registered Accounts")
  cmdAuxiliar.Caption = Choose(gsIdioma, "Proveedor", "Supplier")
  cmdPedido.Caption = Choose(gsIdioma, "Pedido", "Order")
  sstMain.TabCaption(0) = Choose(gsIdioma, "I&mportes", "A&mounts")
  sstMain.TabCaption(1) = Choose(gsIdioma, "C&uentas", "Acco&unts")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
  ']

   dgrDetalle.MarqueeStyle = dbgHighlightRow
   Set dgrDetalle.DataSource = frmTHPrGrd.uorstCOCpbDet
   
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
      '2015-06-23 valid t.cambio con f.emision teo saco dtpDato(3) dtpDato(3).Tag = dtpDato(3).Value
'ini 2016-02-22 val tipo 0=f.emision 1=f.operacion
      'dtpDato(3).Tag = dtpDato(3).Value
      If pval_tpo_cam = 0 Then
          dtpDato(0).Tag = dtpDato(0).Value
      Else
          dtpDato(3).Tag = dtpDato(3).Value
      End If
'fin 2016-02-22 val tipo 0=f.emision 1=f.operacion
      
      '2015-09-28 para cli exte f.Ope/p f.emision s/teo dtpDato(0).Tag = dtpDato(0).Value
   End If
   txtDato(4).Tag = txtDato(4).Text
 ']
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Call gpTeclasData(KeyCode, Shift, Me, True, True, True, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not frmTHPrGrd.uorstMain.EOF Then
      If frmTHPrGrd.uorstMain.EditMode <> adEditNone Then frmTHPrGrd.uorstMain.CancelUpdate   'Cambiar Formulario de Grid.
   End If
End Sub

Private Sub cmdRetroceder_Click()
   gpTUe_Retroceder frmTHPrGrd.uorstMain, Me 'Cambiar Formulario de Grid.
  'Busca ítem.
   frmTHPrGrd.uorstMain_Grd.MoveFirst
   frmTHPrGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
End Sub
Private Sub cmdAvanzar_Click()
   gpTUe_Avanzar frmTHPrGrd.uorstMain, Me 'Cambiar Formulario de Grid.

  'Busca ítem.
   frmTHPrGrd.uorstMain_Grd.MoveFirst
   frmTHPrGrd.uorstMain_Grd.Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & "'"
End Sub

Public Sub cmdCorregir_Click()
  'Verificación de Mes Cerrado.
  If gbCieHpr Then MsgBox TEXT_9016, vbCritical: Exit Sub
  
  pbCorregir = True
  frmTHPrGrd.uocnnMain.BeginTrans     'Cambiar Formulario de Grid. 'INICIA TRANSACCION.
  
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

'ini 2015-02-23 cambios segun teo
''''ini 2015-02-20 monto bruto excede > 1500
'''   If 1500 > Val(txtDato(5).Text) And chkIndAfeIR4.Value = 1 Then
'''      MsgBox "El importe excede al importe de retencion", vbCritical
'''      txtDato(0).SetFocus
'''      Exit Sub
'''   End If
''''fin 2015-02-20 monto bruto excede > 1500

If CDec(txtDato(5).Text) > 1500 And CDec(txtDato(6).Text) = 0# Then
   If MsgBox(Choose(gsIdioma, "Documento no tiene Retencion de 4ta Categoria Continuar?", "Document Retention has 4th Category Continue?"), vbYesNo) = vbNo Then
      txtDato(6).SetFocus
      Exit Sub
   End If
End If

'fin 2015-02-23 cambios segun teo


 '[Propio del formulario.
   Dim dnSumaMN As Double, _
       dnSumaME As Double
   
'   If txtDato(33).Text = "" Then
'      MsgBox TEXT_6002, vbCritical
'      txtDato(33).SetFocus
'      Exit Sub
'   End If

   If txtDato(0).Text = "" Then
      MsgBox TEXT_6002, vbCritical
      txtDato(0).SetFocus
      Exit Sub
   End If
   
   With frmTHPrGrd.uorstMain
      dnSumaMN = CDec(txtDato(5).Text) - CDec(txtDato(6).Text) - CDec(txtDato(7).Text) - CDec(txtDato(8).Text)
      dnSumaME = CDec(txtDato(10).Text) - CDec(txtDato(11).Text) - CDec(txtDato(12).Text) - CDec(txtDato(13).Text)
      If dnSumaMN <> CDec(txtDato(9).Text) Then
        If (cboTpoMon.ListIndex = TPOMON_EXT_IND And cmdMas(CANTIDADIMPORTES).Tag <> INDMASCTA_MAS And (Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text)) <= 0.02)) Then
          If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text)))) & "." & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
          txtDato(MINIMOINDICEIMPORTEMN + CANTIDADIMPORTES - 1).Text = Format(dnSumaMN, FORMATO_NUM_1)
        Else
         If MsgBox(TEXT_9011 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(9).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
        End If
      ElseIf dnSumaME <> CDec(txtDato(14).Text) Then
        If (cboTpoMon.ListIndex = TPOMON_NAC_IND And cmdMas(CANTIDADIMPORTES).Tag <> INDMASCTA_MAS And (Abs(dnSumaME - CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)) <= 0.02)) Then
          If MsgBox(TEXT_9012 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaME - CDec(txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text)))) & "." & Chr(13) & Choose(gsIdioma, "Cuadre Automático?", "It squares Automatic?"), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
          txtDato(MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1).Text = Format(dnSumaME, FORMATO_NUM_1)
        Else
         If MsgBox(TEXT_9012 & Chr(13) & Choose(gsIdioma, "La diferencia es de ", "The difference is ") & Trim(CStr(Abs(dnSumaME - CDec(txtDato(14).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then Exit Sub
        End If
      End If
   End With
  ' Valido los saldos del pedido
  If txtDato(27).Text <> "" Then
    If Not pfValidaPedido(txtDato(27).Text, "N") Then Exit Sub
  End If
   
  ' Genero las cuentas de acuerdo al asiento tipo
  If txtDato(26).Text <> "" And pbNuevo Then
    ppInsDelCtaCos txtDato(26).Text, INDCCO_INA
    ppInsDelCtaCos txtDato(26).Text, INDCCO_ACT
  End If
   
   ' Valido las Cuentas esten Correctas(llenas para todas los valores)
   If chkIndPreGen.Value = vbChecked Then
    'ini 2014-07-10 inhabilita y activa cuentas registradas
'      chkIndPreGen.Value = IIf(ValidaCtasCCo, 1, 0)
'      If Not (chkIndPreGen.Value = vbChecked) Then
'        If MsgBox(TEXT_9013, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
'          Exit Sub
'        End If
'      End If
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

   With frmTHPrGrd                     'Cambiar Formulario de Grid.
      If pbNuevo And frmTHPrGrd.ubGrabaMas = 0 Then
         .uorstMain.AddNew
      End If
      upDatosDesconectados 0
      With .uorstMain
         If pbNuevo Then
            !UsrCre = gsAbvUsr
            !FyHCre = Now
         Else
            !UsrMdf = gsAbvUsr
'            !FyHMdf = Now
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
         frmTHPrGrd.ubGrabaMas = INDMASCTA_INI

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
   
'ini 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
fEstMayUpd
'fin 2015-06-05 Si Mayorizo o no . Estado Mayorizacion
      
   Exit Sub
Err:
   gpErrores
  
   frmTHPrGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
End Sub

Public Sub cmdDeshacer_Click()
 '[Propio del formulario.
   frmTHPrGrd.uorstCOCpbCab.CancelUpdate
   frmTHPrGrd.uorstCOCpbDet.CancelBatch
   frmTHPrGrd.uocnnMain.RollbackTrans  'RESTAURA TRANSACCION.
   pbCorregir = False
 ']
   
   gpTUe_Deshacer Me
End Sub

Public Sub cmdSalir_Click()
   If pbNuevo Or pbCorregir Then
      pbCorregir = False
      frmTHPrGrd.uocnnMain.RollbackTrans 'RESTAURA TRANSACCION.
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
   Case 0, 26, 27, 28, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
      txtDato(Index).SetFocus
   End Select
   ppAyuBus AYUDAT, Index
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
   If pbValidada Then
'      dtpDato(3).SetFocus 'Cambiar.
      
      txtLlave(0).Enabled = False
      txtLlave(1).Enabled = False
      txtLlave(2).Enabled = False
      lblLlaveDeta(0).Enabled = False
      cmdLlaveAyud(0).Enabled = False
    'ini 2014-07-09 inhabilita y activa cuentas registradas
    chkIndPreGen.Value = 1 'activar chek
    chkIndPreGen.Enabled = False
    'fin 2014-07-09 inhabilita y activa cuentas registradas
   End If
   If pbValidada And dtpDato(3).Enabled Then dtpDato(3).SetFocus 'Cambiar.
End Sub

Private Sub txtllave_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

   Dim dvRegistro As Variant
   
 '[Llena con ceros a la izquierda.     'Cambiar (habilitar/deshabilitar).
   Select Case Index
   Case 1, 2                           'Cambiar (añadir índices).
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
      With frmTHPrGrd                  'Cambiar Formulario de Grid.
         Set .uorstTemporal = .uocnnMain.Execute("SELECT MesPvs FROM COHPrDoc WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' AND CodAux='" & txtLlave(0).Text & "' AND SerDoc='" & txtLlave(1).Text & "' AND NroDoc='" & txtLlave(2).Text & "'")
         If .uorstTemporal.RecordCount > 0 Then
            MsgBox TEXT_8007 & Chr(13) & Choose(gsIdioma, "(mes ", "(Month ") & gfMesLet("01" & .uorstTemporal!mespvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
            Cancel = True
            Exit Sub
         End If
         .uorstTemporal.Close
      End With
    '[Propio del formulario.
      If frmTHPrGrd.ubGrabaMas = 0 Then
         frmTHPrGrd.ubGrabaMas = 1
         With frmTHPrGrd
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
  If KeyCode = vbKeyF2 Then ppAyuBus AYUDAT, Index
End Sub

Private Sub txtDato_LostFocus(Index As Integer) 'Cambiar.
''   Dim doColumna As Field
''
   Select Case Index
''   Case 4
''      If txtDato(Index).Tag <> frmTHPrGrd.uorstCOCta!TpoTCb Or CDec(txtDato(Index).Text) = 0 Then
''         txtDato(Index).Tag = frmTHPrGrd.uorstCOCta!TpoTCb
''         With frmTHPrGrd.uorstTGTCb
''            If .RecordCount <> 0 Then .MoveFirst
''            .Find "FehTCb = '" & dtpDato(3).Value & "'"
'''            uorstMain!ImpTCb = IIf(frmTHPrGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
''         End With
''      End If
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
      If Index = 5 Or Index = 10 Then
         If chkCalcularIR4 Then txtDato(Index + 1).Text = Format(CDec(txtDato(Index).Text) * CDec(gnPctIR4) / 100, FORMATO_NUM_1)
         If chkCalcularIES Then txtDato(Index + 2).Text = Format(CDec(txtDato(Index).Text) * CDec(gnPctIES) / 100, FORMATO_NUM_1)
'ini 2014-09-08 RR.HH afecto afp/onp debe considerar solo el valor sol en calculo
        If chkAfeIndORt Then
            With frmTHPrGrd.uorstCodOnpAfp
                If .RecordCount > 0 Then .MoveFirst
                .Find "CodAux='" & txtLlave(0).Text & "'"
                If Not .EOF Then
                    If !Fecnacimiento >= CDate(Format("01/08/1973")) Then
                        'convertir a soles si es dolares calculo en mn afo opn
                        Dim xOtrImp As Double
                        If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                        xOtrImp = CDec(txtDato(5).Text)
                        Else
                        xOtrImp = Format(CDec(txtDato(10).Text) / CDec(txtDato(4).Text), FORMATO_NUM_1)
                        End If
                    
                        Dim xafpAporte As Double:  Dim xafpComi As Double
                        Dim xafpSeguro As Double:
                        xafpAporte = 0: xafpComi = 0: xafpSeguro = 0
                        If !Factor1 <> 0 Then xafpAporte = Format(xOtrImp * CDec(!Factor1) / 100, FORMATO_NUM_1)
                        'Flagcomision 0=comsion mixta  1=comision  flujo
                        If !Flagcomision = "1" Then
                        If !Factor2 <> 0 Then xafpComi = Format(xOtrImp * CDec(!Factor2) / 100, FORMATO_NUM_1)
                        Else
                        If !Factor3 <> 0 Then xafpComi = Format(xOtrImp * CDec(!Factor3) / 100, FORMATO_NUM_1)
                        End If
                        If !topeseg > xOtrImp Then
                        If !Factor4 <> 0 Then xafpSeguro = Format(xOtrImp * CDec(!Factor4) / 100, FORMATO_NUM_1)
                        Else
                        If !Factor4 <> 0 Then xafpSeguro = Format(!topeseg * CDec(!Factor4) / 100, FORMATO_NUM_1)
                        End If
                        'txtDato(Index + 3).Text = Format(xafpAporte + xafpComi + xafpSeguro, FORMATO_NUM_1)
                        'convierte en moneda original
                        If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
                        txtDato(8).Text = Format(xafpAporte + xafpComi + xafpSeguro, FORMATO_NUM_1)
                        Else
                        txtDato(13).Text = Format((xafpAporte + xafpComi + xafpSeguro) / CDec(txtDato(4).Text), FORMATO_NUM_1)
                        End If
                        
                    End If
                End If
            End With
        End If
'fin 2014-09-08 RR.HH afecto afp/onp debe considerar solo el valor sol en calculo
         
''ini 2014-08-01 RR.HH afecto afp/onp
'        If chkAfeIndORt Then
'            With frmTHPrGrd.uorstCodOnpAfp
'                If .RecordCount > 0 Then .MoveFirst
'                .Find "CodAux='" & txtLlave(0).Text & "'"
'                If Not .EOF Then
'                    If !Fecnacimiento >= CDate(Format("01/08/1973")) Then
'                        Dim xafpAporte As Double:  Dim xafpComi As Double
'                        Dim xafpSeguro As Double:
'                        xafpAporte = 0: xafpComi = 0: xafpSeguro = 0
'                        If !Factor1 <> 0 Then xafpAporte = Format(CDec(txtDato(Index).Text) * CDec(!Factor1) / 100, FORMATO_NUM_1)
'                        'Flagcomision 0=comsion mixta  1=comision  flujo
'                        If !Flagcomision = "1" Then
'                        If !Factor2 <> 0 Then xafpComi = Format(CDec(txtDato(Index).Text) * CDec(!Factor2) / 100, FORMATO_NUM_1)
'                        Else
'                        If !Factor3 <> 0 Then xafpComi = Format(CDec(txtDato(Index).Text) * CDec(!Factor3) / 100, FORMATO_NUM_1)
'                        End If
'                        If !topeseg > CDec(txtDato(Index).Text) Then
'                        If !Factor4 <> 0 Then xafpSeguro = Format(CDec(txtDato(Index).Text) * CDec(!Factor4) / 100, FORMATO_NUM_1)
'                        Else
'                        If !Factor4 <> 0 Then xafpSeguro = Format(!topeseg * CDec(!Factor4) / 100, FORMATO_NUM_1)
'                        End If
'                        txtDato(Index + 3).Text = Format(xafpAporte + xafpComi + xafpSeguro, FORMATO_NUM_1)
'                    End If
'                End If
'            End With
'        End If
''fin 2014-08-01 RR.HH afecto afp/onp
         If chkMonedaActiva.Value = vbChecked Then
            If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
               If CDec(txtDato(Index + 1).Text) > 0 Then txtDato((Index + CANTIDADIMPORTES) + 1).Text = Format(gfRedond(CDec(txtDato(Index + 1).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
               If CDec(txtDato(Index + 2).Text) > 0 Then txtDato((Index + CANTIDADIMPORTES) + 2).Text = Format(gfRedond(CDec(txtDato(Index + 2).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
'2014-08-01 RR.HH afecto afp/onp
               If CDec(txtDato(Index + 3).Text) > 0 Then txtDato((Index + CANTIDADIMPORTES) + 3).Text = Format(gfRedond(CDec(txtDato(Index + 3).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
            Else
               If CDec(txtDato(Index + 1).Text) > 0 Then txtDato((Index - CANTIDADIMPORTES) + 1).Text = Format(gfRedond(CDec(txtDato(Index + 1).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
               If CDec(txtDato(Index + 2).Text) > 0 Then txtDato((Index - CANTIDADIMPORTES) + 2).Text = Format(gfRedond(CDec(txtDato(Index + 2).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
'2014-08-01 RR.HH afecto afp/onp
               If CDec(txtDato(Index + 3).Text) > 0 Then txtDato((Index - CANTIDADIMPORTES) + 3).Text = Format(gfRedond(CDec(txtDato(Index + 3).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
            End If
         End If
      End If
      
     'Cálculo del total.
      If (Index = 9 And txtDato(Index).Text = 0) Or (Index = 14 And txtDato(Index).Text = 0) Then
         If Index = 9 Then
            txtDato(9).Text = Format(CDec(txtDato(5).Text) - CDec(txtDato(6).Text) - CDec(txtDato(7).Text) - CDec(txtDato(8).Text), FORMATO_NUM_1)
         Else
            txtDato(14).Text = Format(CDec(txtDato(10).Text) - CDec(txtDato(11).Text) - CDec(txtDato(12).Text) - CDec(txtDato(13).Text), FORMATO_NUM_1)
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
            If frmTHPrGrd.uorstCOHPrDocCta.RecordCount > 0 Then
               ppAbreCtaCCo
               If frmTHPrGrd.uorstCOHPrDocCta.State = adStateOpen Then
                  frmTHPrGrd.uorstCOHPrDocCta.MoveFirst
                  Do
                     If frmTHPrGrd.uorstCOHPrDocCta!codaux = txtLlave(0).Text And _
                       frmTHPrGrd.uorstCOHPrDocCta!serdoc = txtLlave(1).Text And _
                       frmTHPrGrd.uorstCOHPrDocCta!nrodoc = txtLlave(2).Text And _
                       Trim(frmTHPrGrd.uorstCOHPrDocCta!tpocnc) = Trim(Str(Index - MINIMOINDICECUENTA + 1)) Then
                        frmTHPrGrd.uorstCOHPrDocCta.Delete
                     End If
                     frmTHPrGrd.uorstCOHPrDocCta.MoveNext
                  Loop Until frmTHPrGrd.uorstCOHPrDocCta.EOF
                  frmTHPrGrd.uorstCOHPrDocCta.Requery
               End If
            End If
         End If
         cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI
      ElseIf cmdMas(Index - MINIMOINDICECUENTA + 1).Tag = INDMASCTA_INI Then
        cmdMas(Index - MINIMOINDICECUENTA + 1).Enabled = False
      End If
    Case 26
      If txtDato(Index).Text <> "" And pbNuevo Then
        chkDesactivar.Value = vbChecked
      End If
    Case 27
    ' Inicializo clon de pedido
    txtDato(Index).Tag = txtDato(Index).Text

   End Select
End Sub

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
   On Error GoTo Err

  'Completa con ceros a la izquierda.
   Select Case Index
   Case MINIMOINDICECCOSTO To CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
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
   Case 0, 26, 27, 28, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
    Cancel = ppAyuDet(AYUDAT, Index)
    If Cancel Then Exit Sub
         
    If Index >= MINIMOINDICECUENTA And Index < MINIMOINDICECCOSTO Then
      If frmTHPrGrd.uorstCoCta.RecordCount > 0 Then
        If Not frmTHPrGrd.uorstCoCta.EOF Then
          If frmTHPrGrd.uorstCoCta!indcco = INDCCO_ACT Then
            ' Inicializo el centro de costos
            txtDato(Index + CUENTASCONCCOSTO).Tag = txtDato(Index + CUENTASCONCCOSTO).Text
            txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index).Tag <> txtDato(Index).Text, "", txtDato(Index + CUENTASCONCCOSTO).Text)
            If Not IsNull(frmTHPrGrd.uorstCoCta!codcco_def) Then
              txtDato(Index + CUENTASCONCCOSTO).Text = IIf(txtDato(Index + CUENTASCONCCOSTO).Text = "", frmTHPrGrd.uorstCoCta!codcco_def, txtDato(Index + CUENTASCONCCOSTO).Text)
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
  Dim s_PedidoCco As String
  
  s_PedidoCco = "AND indpdocpr='" & IIf(txtDato(27).Text = "", INDCCO_INA, INDCCO_ACT) & "' "
  If tsTipo = AYULLA Then
    Select Case tnIndex
     Case 0                           'Cambiar (añadir índices).
      modAyuBus.Aux_Det "IndPrv=1", txtLlave(tnIndex).Text, 0, 0, Me.Top + txtLlave(tnIndex).Top + txtLlave(tnIndex).Height, Me.Left + txtLlave(tnIndex).Left
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
     Case 26
      modAyuBus.Asi_Cod "tpoasi='" & TPOGNR_HPR & "' ", txtDato(tnIndex).Text, 0, 0, Me.Top + fraAsiento.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left + fraAsiento.Left + txtDato(tnIndex).Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
     Case 27    ' Pedido de compra                           'Cambiar (añadir índices).
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
      Case 28                             ' orden de servicio
         ' Filtro de seleccion
         If ps_Plataforma = pSrvMySql Then
           cmdDatoAyud(tnIndex).Tag = "a.fehcon<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
         ElseIf ps_Plataforma = pSrvSql Then
           cmdDatoAyud(tnIndex).Tag = "a.fehcon<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
         End If
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
      If txtLlave(tnIndex).Text = "" Then lblLlaveDeta(tnIndex).Caption = "": Exit Function
      With frmTHPrGrd.uorstTGAux
        If .RecordCount > 0 Then .MoveFirst
        .Find "CodAux='" & txtLlave(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblLlaveDeta(tnIndex).Caption = " " & !razAux
        End If
      End With
    End Select
  Else
    Select Case tnIndex                 'Cambiar.
     Case 0
      If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
      With frmTHPrGrd.uorstCODro
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
      If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
      With frmTHPrGrd.uorstCoCta
        If .RecordCount > 0 Then .MoveFirst
        .Find "CodCta='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
            'ini 2015-06-30 correccion tipo mon cta
            If tnIndex = 19 And !tpomon <> Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT) Then
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
      If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
      With frmTHPrGrd.uorstCoCCo
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
     Case 26
      If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
      With frmTHPrGrd.uorstCoAsiTipo
        If .RecordCount > 0 Then .MoveFirst
        .Find "codasi='" & txtDato(tnIndex).Text & "'"
        If .EOF Then
          MsgBox TEXT_8006, vbExclamation
          ppAyuDet = True
        Else
          lblDatoDeta(tnIndex).Caption = " " & IIf(IsNull(!detasi), "", !detasi)
        End If
      End With
     Case 27    ' Pedido de compra
      If txtDato(tnIndex).Text = "" Then lblDatoDeta(tnIndex).Caption = "": Exit Function
      ppAyuDet = Not pfValidaPedido(txtDato(tnIndex), "S")
    End Select
  End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
  Dim dnContador As Byte
   
  On Error GoTo Err


  With frmTHPrGrd.uorstMain           'Cambiar RecordSet.
    If tnFase = 0 Then
      'Llaves.
      If pbNuevo Then
        !codemp = gsCodEmp
        !pdoano = gsAnoAct
        !codaux = txtLlave(0).Text
        !serdoc = txtLlave(1).Text
        !nrodoc = txtLlave(2).Text
        !mespvs = gsMesAct
        !PctIR4 = CDec(gnPctIR4)
        !PctIES = CDec(gnPctIES)
      End If
      
      'Datos.
      !tpomon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
      !indpregen = IIf(chkIndPreGen.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
      '[REVISAR.
      !IndAfeIES = chkAfeIndIES.Value
      !IndAfeIR4 = chkIndAfeIR4.Value
      !IndAfeORt = chkAfeIndORt.Value
      ']REVISAR.
      !fehope = dtpDato(3).Value
      !feedoc = dtpDato(0).Value
      !coddro = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
      !NroCpb = txtDato(1).Text
      !RefDoc = txtDato(2).Text
      !GloDoc = IIf(txtDato(Choose(gsIdioma, 3, 25)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 25)).Text)
      !glodocx = IIf(txtDato(Choose(gsIdioma, 25, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 25, 3)).Text)
      !pdocpr = IIf(txtDato(27).Text = "", Null, txtDato(27).Text)
      !codasi = IIf(txtDato(26).Text = "", Null, txtDato(26).Text)
      !codcon = IIf(txtDato(28).Text = "", Null, txtDato(28).Text)
      !ImpTCb = CDec(txtDato(4).Text)
      !ImpBru_MN = CDec(txtDato(5).Text)
      !ImpIR4_MN = CDec(txtDato(6).Text)
      !ImpIES_MN = CDec(txtDato(7).Text)
      !ImpORt_MN = CDec(txtDato(8).Text)
      !ImpNet_MN = CDec(txtDato(9).Text)
      !ImpBru_ME = CDec(txtDato(10).Text)
      !ImpIR4_ME = CDec(txtDato(11).Text)
      !ImpIES_ME = CDec(txtDato(12).Text)
      !ImpORt_ME = CDec(txtDato(13).Text)
      !ImpNet_ME = CDec(txtDato(14).Text)
      
      '[Actualización de datos por Cuentas y Centros de Costo registrados directamente.
      ppAbreCtaCCo
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
        If cmdMas(dnContador).Tag = INDMASCTA_CTA And Val(txtDato(dnContador + DIFERENCIAMASCUENTA).Text) <> 1 Then
          With frmTHPrGrd.uorstCOHPrDocCta
            .MoveFirst
            .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & ps_OrdenCta & "'"
            If Not .EOF Then
              .Delete
              .Update
              .Requery
              frmTHPrGrd.uorstCOHPrDocCCo.Requery
              upActualizaMas dnContador, INDMASCTA_INI
            End If
          End With
        End If
        
        If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
           cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Then
          With frmTHPrGrd.uorstCOHPrDocCta
            If .RecordCount <> 0 Then .MoveFirst
            .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & ps_OrdenCta & "'"
            If .EOF Then
              .AddNew
              !codemp = gsCodEmp
              !pdoano = gsAnoAct
              !codaux = txtLlave(0).Text
              !serdoc = txtLlave(1).Text
              !nrodoc = txtLlave(2).Text
              !tpocnc = dnContador
              !orden = ps_OrdenCta
              !UsrCre = gsAbvUsr
              !FyHCre = Now
            Else
              !UsrMdf = gsAbvUsr
              '!FyHMdf = Now
            End If
            '2015-08-05 trim de cta y ccto !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
            !CodCta = Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)
            !glodet = txtDato(3)
            !impcta_mn = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE).Text)
            !impcta_me = CDec(txtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text)
            .Update
          End With
          If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Or _
             cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Then
            With frmTHPrGrd.uorstCOHPrDocCCo
              If .RecordCount <> 0 Then .MoveFirst
              .Find "cLlave1='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & ps_OrdenCta & txtDato(dnContador + DIFERENCIAMASCUENTA).Text & "'"
              If .EOF Then
                .AddNew
                !codemp = gsCodEmp
                !pdoano = gsAnoAct
                !codaux = txtLlave(0).Text
                !serdoc = txtLlave(1).Text
                !nrodoc = txtLlave(2).Text
                !tpocnc = dnContador
                !orden = ps_OrdenCta
                '2015-08-05 trim de cta y ccto !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
                !CodCta = Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)
                !UsrCre = gsAbvUsr
                !FyHCre = Now
              Else
                !UsrMdf = gsAbvUsr
                '!FyHMdf = Now
              End If
              '2015-08-05 trim de cta y ccto !codcco = txtDato(dnContador + DIFERENCIAMASCCOSTO).Text
              !codcco = Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)
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
      txtLlave(1).Text = !serdoc
      txtLlave(2).Text = !nrodoc
      
      'Datos.
      cboTpoMon.ListIndex = IIf(!tpomon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      chkIndPreGen.Value = IIf(!indpregen = INDPREGEN_ACT, vbChecked, vbUnchecked)
      '[Revisar.
      chkAfeIndIES.Value = IIf(!IndAfeIES, vbChecked, vbUnchecked)
      chkIndAfeIR4.Value = IIf(!IndAfeIR4, vbChecked, vbUnchecked)
      chkAfeIndORt.Value = IIf(!IndAfeORt, vbChecked, vbUnchecked)
      ']Revisar.
      dtpDato(3).Value = !fehope
      dtpDato(0).Value = !feedoc
      txtDato(0).Text = IIf(IsNull(!coddro), "", !coddro)
      txtDato(1).Text = IIf(IsNull(!NroCpb), "", !NroCpb)
      txtDato(2).Text = IIf(IsNull(!RefDoc), "", !RefDoc)
      txtDato(Choose(gsIdioma, 3, 25)).Text = IIf(IsNull(!GloDoc), "", !GloDoc)
      txtDato(Choose(gsIdioma, 25, 3)).Text = IIf(IsNull(!glodocx), "", !glodocx)
      txtDato(27).Text = IIf(IsNull(!pdocpr), "", !pdocpr)
      txtDato(26).Text = IIf(IsNull(!codasi), "", !codasi)
      txtDato(28).Text = IIf(IsNull(!codcon), "", !codcon)
      txtDato(4).Text = Format(!ImpTCb, FORMATO_NUM_2)
      txtDato(5).Text = Format(!ImpBru_MN, FORMATO_NUM_1)
      txtDato(6).Text = Format(!ImpIR4_MN, FORMATO_NUM_1)
      txtDato(7).Text = Format(!ImpIES_MN, FORMATO_NUM_1)
      txtDato(8).Text = Format(!ImpORt_MN, FORMATO_NUM_1)
      txtDato(9).Text = Format(!ImpNet_MN, FORMATO_NUM_1)
      txtDato(10).Text = Format(!ImpBru_ME, FORMATO_NUM_1)
      txtDato(11).Text = Format(!ImpIR4_ME, FORMATO_NUM_1)
      txtDato(12).Text = Format(!ImpIES_ME, FORMATO_NUM_1)
      txtDato(13).Text = Format(!ImpORt_ME, FORMATO_NUM_1)
      txtDato(14).Text = Format(!ImpNet_ME, FORMATO_NUM_1)
      For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
        txtDato(dnContador).Text = ""
        txtDato(dnContador).Tag = ""
      Next
      
      '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
      s_PedidoValiDsc = "S"
      ppAyuDet AYULLA, 0
      ppAyuDet AYUDAT, 15
      ppAyuDet AYUDAT, 16
      ppAyuDet AYUDAT, 17
      ppAyuDet AYUDAT, 18
      ppAyuDet AYUDAT, 19
      ppAyuDet AYUDAT, 20
      ppAyuDet AYUDAT, 21
      ppAyuDet AYUDAT, 22
      ppAyuDet AYUDAT, 23
      ppAyuDet AYUDAT, 24
      ppAyuDet AYUDAT, 26
      ppAyuDet AYUDAT, 27
      ppAyuDet AYUDAT, 28
      ']
      
      '[Propio del formulario.
      For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
        cmdMas(dnContador).Tag = Choose(dnContador, !IndCta_Bru, !IndCta_IR4, !IndCta_IES, !IndCta_ORt, !IndCta_Net)
      Next
      
      ppAbreCtaCCo
      With frmTHPrGrd.uorstCOHPrDocCta
        If .RecordCount > 0 Then
          For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
            If Val(txtDato(dnContador + DIFERENCIAMASIMPORTE).Text) <> 0 Then
              .MoveFirst
              .Find "TpoCnc = " & dnContador
              If Not .EOF Then
                txtDato(dnContador + DIFERENCIAMASCUENTA).Text = !CodCta
                txtDato(dnContador + DIFERENCIAMASCUENTA).Tag = !CodCta
                ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCUENTA
                With frmTHPrGrd.uorstCOHPrDocCCo
                  If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "cLlave = " & dnContador & ps_OrdenCta & frmTHPrGrd.uorstCOHPrDocCta!CodCta
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
  
  'Datos.
  cboTpoMon.ListIndex = TPOMON_NAC_IND
  chkIndPreGen.Value = vbUnchecked
  dtpDato(3).Value = Date
  dtpDato(0).Value = Date
  '   optTpoMon(1).Value = True
  For dnContador = 0 To 3
    txtDato(dnContador).Text = ""
  Next
  txtDato(4).Text = Format(0, FORMATO_NUM_2)
  txtDato(25).Text = ""
  txtDato(26).Text = ""
  txtDato(27).Text = ""
  txtDato(28).Text = ""
  For dnContador = MINIMOINDICEIMPORTEMN To MINIMOINDICEIMPORTEME + CANTIDADIMPORTES - 1
    txtDato(dnContador).Text = Format(0, FORMATO_NUM_1)
  Next
  For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
    txtDato(dnContador).Text = ""
    txtDato(dnContador).Tag = ""
  Next
  
  '[Propio del formulario.
  For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
    cmdMas(dnContador).Tag = INDMASCTA_INI
  Next
  ']
  
  s_PedidoValiDsc = "N"
  
  'Ayudas.
  lblLlaveDeta(0).Caption = ""
  lblDatoDeta(0).Caption = ""
  lblDatoDeta(26).Caption = ""
  lblDatoDeta(27).Caption = ""
  lblDatoDeta(28).Caption = ""
  For dnContador = MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
    lblDatoDeta(dnContador).Caption = ""
  Next
End Sub

Public Sub upHabilitacion(tbHabilitar As Boolean) 'Cambiar.
  Dim dnContador As Byte
  
  'Datos.
  cboTpoMon.Enabled = tbHabilitar
  chkAfeIndIES.Enabled = tbHabilitar
  chkIndAfeIR4.Enabled = tbHabilitar
  chkAfeIndORt.Enabled = tbHabilitar
  chkCalcularIES.Enabled = (chkAfeIndIES.Enabled And chkAfeIndIES.Value = vbChecked)
  chkCalcularIR4.Enabled = (chkIndAfeIR4.Enabled And chkIndAfeIR4.Value = vbChecked)
  chkDesactivar.Enabled = tbHabilitar
  chkIndPreGen.Enabled = tbHabilitar
  chkMonedaActiva.Enabled = tbHabilitar
  dtpDato(3).Enabled = tbHabilitar
  dtpDato(0).Enabled = tbHabilitar
  With txtDato
    For dnContador = 0 To .Count - 1
      .Item(dnContador).Enabled = tbHabilitar
    Next
  End With
  txtDato(26).Enabled = (tbHabilitar And pbNuevo)
  
  For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
    upHabilitaCuenta False, dnContador
  Next
  
  'Ayudas.
  cmdDatoAyud(0).Enabled = tbHabilitar
  lblDatoDeta(0).Enabled = tbHabilitar
  cmdDatoAyud(26).Enabled = (tbHabilitar And pbNuevo)
  lblDatoDeta(26).Enabled = (tbHabilitar And pbNuevo)
  cmdDatoAyud(27).Enabled = tbHabilitar
  lblDatoDeta(27).Enabled = tbHabilitar
  cmdDatoAyud(28).Enabled = tbHabilitar
  lblDatoDeta(28).Enabled = tbHabilitar

  '[Propio del formulario
  txtDato(1).Enabled = False 'Deshabilitación del Comprobante.
  
  If tbHabilitar Then
    For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      cmdMas(dnContador).Enabled = Not (cmdMas(dnContador).Tag = INDMASCTA_CTA)
      upHabilitaCuenta (Not cmdMas(dnContador).Tag = INDMASCTA_MAS), dnContador
    Next
  Else
    For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      cmdMas(dnContador).Enabled = False
      upHabilitaCuenta False, dnContador
      upHabilitaCCosto False, dnContador
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
    If cmdMas(dnContador).Tag = INDMASCTA_CTA Then
      cmdMas(dnContador).Enabled = cmdMas(dnContador).Enabled = True
    Else
      cmdMas(dnContador).Enabled = IIf(chkDesactivar.Value = vbChecked, False, True)
    End If
  Next
   
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
  txtDato(26).Text = IIf((pbNuevo And chkDesactivar.Value = vbUnchecked), "", txtDato(26).Text)
  lblDatoDeta(26).Caption = IIf((pbNuevo And chkDesactivar.Value = vbUnchecked), "", lblDatoDeta(26).Caption)

End Sub

Private Sub chkAfeIndIES_Click()
   chkCalcularIES.Enabled = (chkAfeIndIES.Value = vbChecked)
End Sub

Private Sub chkIndAfeIR4_Click()
   chkCalcularIR4.Enabled = (chkIndAfeIR4.Value = vbChecked)
End Sub

Private Sub chkMonedaActiva_Click()
   unVerMonNac = IIf(chkMonedaActiva.Value, cboTpoMon.ListIndex, IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_EXT_IND, TPOMON_NAC_IND))
   ppCambioTpoMon
End Sub

Private Sub cmdAuxiliar_Click()
  frmMAuxGrd.Show vbModal
  frmTHPrGrd.uorstTGAux.Requery
End Sub

Private Sub cmdMas_Click(Index As Integer) 'Cambiar Formulario de Grid.
  frmTHPrMasGrd.unIndice = Index
  frmTHPrMasGrd.Show vbModal
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
         MsgBox Choose(gsIdioma, "La fecha emi No Corresponde al Periodo de Operacion", "The date does not correspond with operating period"), vbCritical
         dtpDato(Index).SetFocus
         Cancel = True
         Exit Sub
      End If
      
'ini 2015-09-28 para cli exte f.Ope/p f.emision s/teo
''ini 2015-06-23 valid t.cambio con f.emision teo, saco dtpDato(3)
''ini 2014-08-01 RR.HH afecto afp/onp
'ini 2016-02-22 val tipo 0=f.emision 1=f.operacion
      If pval_tpo_cam = 0 Then
          dtpDato(0).Tag = 0
          If (dtpDato(0).Tag <> dtpDato(0).Value) Then
             dtpDato(0).Tag = dtpDato(0).Value
             With frmTHPrGrd.uorstTGTCb
                If .RecordCount <> 0 Then .MoveFirst
                .Find "FehTCb = '" & dtpDato(0).Value & "'"
                If .EOF Then
                   MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
                   Cancel = True
                   Exit Sub
                Else
    '               uorstMain!ImpTCb = IIf(frmTHPrGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
    ''               frmTHPrGrd.uorstMain!ImpTCb = !ImpTCb_Vta
                   txtDato(4).Text = Format(!ImpTCb_Vta, FORMATO_NUM_2)
                End If
             End With
          End If
      End If
'fin 2016-02-22 val tipo 0=f.emision 1=f.operacion
''fin 2014-08-01 RR.HH afecto afp/onp
''fin 2015-06-23 valid t.cambio con f.emision teo saco dtpDato(3)
'fin 2015-09-28 para cli exte f.Ope/p f.emision s/teo

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
'ini 2015-09-28 para cli exte f.Ope/p f.emision s/teo
'ini 2015-06-23 valid t.cambio con f.emision teo saco dtpDato(3)
'ini 2014-08-01 RR.HH afecto afp/onp
'ini 2016-02-22 val tipo 0=f.emision 1=f.operacion
      If pval_tpo_cam = 1 Then
          dtpDato(3).Tag = 0
          If (dtpDato(3).Tag <> dtpDato(3).Value) Then
             dtpDato(3).Tag = dtpDato(3).Value
             With frmTHPrGrd.uorstTGTCb
                If .RecordCount <> 0 Then .MoveFirst
                .Find "FehTCb = '" & dtpDato(3).Value & "'"
                If .EOF Then
                   MsgBox Choose(gsIdioma, "No se ha ingresado Tipo de Cambio para esta Fecha", "Rate of exchange has not been entered for this date"), vbCritical
                   Cancel = True
                   Exit Sub
                Else
    '               uorstMain!ImpTCb = IIf(frmTHPrGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
    ''               frmTHPrGrd.uorstMain!ImpTCb = !ImpTCb_Vta
                   txtDato(4).Text = Format(!ImpTCb_Vta, FORMATO_NUM_2)
                End If
             End With
          End If
      End If
'fin 2016-02-22 val tipo 0=f.emision 1=f.operacion
'fin 2014-08-01 RR.HH afecto afp/onp
'fin 2015-06-23 valid t.cambio con f.emision teo saco dtpDato(3)
'fin 2015-09-28 para cli exte f.Ope/p f.emision s/teo
      If txtDato(4).Text = 0 Then
         If pbFecha Then
            With dtpDato
            For dnContador = 0 To .Count - 2
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
   dgrDetalle.Caption = "xx"
   'dgrDetalle.SetFocus
   dgrDetalle.SetFocus
   
End Sub

Private Sub ppAbreCtaCCo()
   With frmTHPrGrd.uorstCOHPrDocCCo
      frmTHPrGrd.usConnStrgWher_COHPrDocCCo = "WHERE COHPrDocCCo.codemp='" & frmTHPrGrd.uorstMain!codemp & "' AND COHPrDocCCo.pdoano='" & frmTHPrGrd.uorstMain!pdoano & "' "
      frmTHPrGrd.usConnStrgWher_COHPrDocCCo = frmTHPrGrd.usConnStrgWher_COHPrDocCCo & "AND COHPrDocCCo.CodAux='" & frmTHPrGrd.uorstMain!codaux & "' "
      frmTHPrGrd.usConnStrgWher_COHPrDocCCo = frmTHPrGrd.usConnStrgWher_COHPrDocCCo & "AND COHPrDocCCo.SerDoc='" & frmTHPrGrd.uorstMain!serdoc & "' AND COHPrDocCCo.NroDoc='" & frmTHPrGrd.uorstMain!nrodoc & "' "
      If .State = adStateOpen Then .Close
      .Source = frmTHPrGrd.usConnStrgSele_COHPrDocCCo & frmTHPrGrd.usConnStrgWher_COHPrDocCCo & frmTHPrGrd.usConnStrgOrde_COHPrDocCCo
      .Open
      .Properties("Unique Table").Value = "COHPrDocCCo"
   End With
   With frmTHPrGrd.uorstCOHPrDocCta
      frmTHPrGrd.usConnStrgWher_COHPrDocCta = "WHERE COHPrDocCta.codemp='" & frmTHPrGrd.uorstMain!codemp & "' AND COHPrDocCta.pdoano='" & frmTHPrGrd.uorstMain!pdoano & "' "
      frmTHPrGrd.usConnStrgWher_COHPrDocCta = frmTHPrGrd.usConnStrgWher_COHPrDocCta & "AND COHPrDocCta.CodAux='" & frmTHPrGrd.uorstMain!codaux & "' "
      frmTHPrGrd.usConnStrgWher_COHPrDocCta = frmTHPrGrd.usConnStrgWher_COHPrDocCta & "AND COHPrDocCta.SerDoc='" & frmTHPrGrd.uorstMain!serdoc & "' AND COHPrDocCta.NroDoc='" & frmTHPrGrd.uorstMain!nrodoc & "' "
      If .State = adStateOpen Then .Close
      .Source = frmTHPrGrd.usConnStrgSele_COHPrDocCta & frmTHPrGrd.usConnStrgWher_COHPrDocCta & frmTHPrGrd.usConnStrgOrde_COHPrDocCta
      .Open
      .Properties("Unique Table").Value = "COHPrDocCta"
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
  Set frmTHPrGrd.uorstTemporal = frmTHPrGrd.uocnnMain.Execute(sSentencia)
  If Not (frmTHPrGrd.uorstTemporal.BOF Or frmTHPrGrd.uorstTemporal.EOF) And frmTHPrGrd.uorstTemporal.RecordCount > 0 Then
    While Not frmTHPrGrd.uorstTemporal.EOF
            siexiste = True
            frmTHPrGrd.uorstTemporal.MoveNext
    Wend
  Else
  End If
  frmTHPrGrd.uorstTemporal.Close
  
  cuenta = 0
  sSentencia = "SELECT coddro,nrocpb "
  sSentencia = sSentencia & " FROM cohprdoc  "
  sSentencia = sSentencia & " WHERE codemp='" & gsCodEmp & "' "
  sSentencia = sSentencia & " AND pdoano='" & gsAnoAct & "' "
  sSentencia = sSentencia & " AND mespvs='" & gsMesAct & "'"
  sSentencia = sSentencia & " AND coddro='" & txtDato(0).Text & "'"
  sSentencia = sSentencia & " AND nrocpb='" & txtDato(1).Text & "'"
  Set frmTHPrGrd.uorstTemporal = frmTHPrGrd.uocnnMain.Execute(sSentencia)
  If Not (frmTHPrGrd.uorstTemporal.BOF Or frmTHPrGrd.uorstTemporal.EOF) And frmTHPrGrd.uorstTemporal.RecordCount > 0 Then
    While Not frmTHPrGrd.uorstTemporal.EOF
            cuenta = cuenta + 1
            frmTHPrGrd.uorstTemporal.MoveNext
    Wend
  Else
  End If
  frmTHPrGrd.uorstTemporal.Close
  
  If cuenta >= 2 Then masdedos = True


   If txtDato(1).Text <> "" Then
      ppDatosWhere
       If masdedos = False Then
      With frmTHPrGrd.uorstCOCpbCab
        'Si existe, elimina Comprobante existente.
         If .RecordCount > 0 Then
            .MoveFirst
            .Find "cLlave='" & txtDato(0).Text & txtDato(1).Text & "'"
            If Not .EOF Then .Delete
         End If
      End With
      End If
   End If

   With frmTHPrGrd.uorstCOCpbCab
     'Si no está marcado para generar, marca el documento como no generado.
      If chkIndPreGen.Value = vbUnchecked Then
         frmTHPrGrd.uorstMain!indgen = False
         frmTHPrGrd.uorstMain.Update
         Exit Sub
      End If

      ' Captura del siguiente numero de comprobante
      If txtDato(1).Text = "" Then
        txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
        txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
        txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
        txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
        txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
        frmTHPrGrd.uocnnMain.Execute txtDato(1).Tag
        ' Actualizo numero de comprobante tabla de detalle
        frmTHPrGrd.uorstMain!NroCpb = txtDato(1).Text
        frmTHPrGrd.uorstMain.Update
      Else
    If masdedos = True Then
      txtDato(1).Text = gfNumComprobante(gsAnoAct, gsMesAct, txtDato(0).Text)
      txtDato(1).Tag = "UPDATE CoDro SET Cpb" & gsMesAct & "='" & txtDato(1).Text & "' "
      txtDato(1).Tag = txtDato(1).Tag & "WHERE codemp='" & gsCodEmp & "' "
      txtDato(1).Tag = txtDato(1).Tag & "AND pdoano='" & gsAnoAct & "' "
      txtDato(1).Tag = txtDato(1).Tag & "AND CodDro='" & txtDato(0).Text & "'"
      frmTHPrGrd.uocnnMain.Execute txtDato(1).Tag
      ' Actualizo numero de comprobante tabla de detalle
      frmTHPrGrd.uorstMain!NroCpb = txtDato(1).Text
      frmTHPrGrd.uorstMain.Update
    End If
    End If
    
      ppDatosWhere
   
     'Si no hay cuentas, marca el documento como no generado.
      If frmTHPrGrd.uorstCOHPrDocCta.RecordCount = 0 Then
         frmTHPrGrd.uorstMain!indgen = False
         frmTHPrGrd.uorstMain.Update
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
      !tpognr = TPOGNR_HPR
      !IndNCu = INDNCU_FAL
      !glocpb = IIf(txtDato(Choose(gsIdioma, 3, 25)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 25)).Text)
      !glocpbx = IIf(txtDato(Choose(gsIdioma, 25, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 25, 3)).Text)
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With

'[ Teo, Miguel Ange Refresco los recordset de cuentas y centros de costos
  frmTHPrGrd.uorstCOHPrDocCta.Requery
  frmTHPrGrd.uorstCOHPrDocCCo.Requery
']
   With frmTHPrGrd.uorstCOHPrDocCta
     'Crea ítemes de Comprobante.
      .MoveFirst
      Do
         dbProcesaCuenta = True

        'Itemes con Centro de Costo.
         If !tpocnc <= CUENTASCONCCOSTO Then
            With frmTHPrGrd.uorstCOHPrDocCCo
               If .RecordCount <> 0 Then
                  .MoveFirst
                  .Find "cLlave = " & Trim(frmTHPrGrd.uorstCOHPrDocCta!tpocnc) & frmTHPrGrd.uorstCOHPrDocCta!orden & frmTHPrGrd.uorstCOHPrDocCta!CodCta
                  If Not .EOF Then
                     Do
                        dnNumeroItem = dnNumeroItem + 1
                        ppGenera1 True, dnNumeroItem, IIf(CInt(frmTHPrGrd.uorstCOHPrDocCta!tpocnc) >= 2, "", txtDato(28).Text)
                        .MoveNext
                        If .EOF Then Exit Do
                        If !cLlave <> Trim(frmTHPrGrd.uorstCOHPrDocCta!tpocnc) & frmTHPrGrd.uorstCOHPrDocCta!orden & frmTHPrGrd.uorstCOHPrDocCta!CodCta Then Exit Do
                     Loop
                     dbProcesaCuenta = False
                  End If
               End If
            End With
         End If

        'Itemes sin Centro de Costo.
         If dbProcesaCuenta Then
            dnNumeroItem = dnNumeroItem + 1
            ppGenera1 False, dnNumeroItem, IIf(CInt(frmTHPrGrd.uorstCOHPrDocCta!tpocnc) >= 2, "", txtDato(28).Text)
         End If

         .MoveNext
      Loop Until .EOF
   End With

   frmTHPrGrd.uorstMain!indgen = True
   txtDato(0).Enabled = False
   txtDato(1).Enabled = False
   cmdDatoAyud(0).Enabled = False
   lblDatoDeta(0).Enabled = False

   frmTHPrGrd.uorstCOCpbCab.Update
   frmTHPrGrd.uorstCOCpbDet.UpdateBatch
   frmTHPrGrd.uorstMain.Update
End Sub

Private Sub ppGenera1(tbCCosto As Boolean, tnNumeroItem As Integer, ByVal sContrato As String)
  
  With frmTHPrGrd.uorstCOCpbDet
    .AddNew
    !codemp = gsCodEmp
    !pdoano = gsAnoAct
    !coddro = txtDato(0).Text
    !NroCpb = txtDato(1).Text
    !NroIte = tnNumeroItem
    !mespvs = gsMesAct
    !CodCta = frmTHPrGrd.uorstCOHPrDocCta!CodCta
    !fehope = dtpDato(3).Value
    frmTHPrGrd.uorstCoCta.MoveFirst
    frmTHPrGrd.uorstCoCta.Find "CodCta='" & frmTHPrGrd.uorstCOHPrDocCta!CodCta & "'"
    If frmTHPrGrd.uorstCoCta!indcco = INDCCO_ACT Then If tbCCosto Then !codcco = frmTHPrGrd.uorstCOHPrDocCCo!codcco
    If frmTHPrGrd.uorstCoCta!IndDoc = INDDOC_ACT Then
      !codaux = txtLlave(0).Text
    Else
      If Len(Trim(frmTHPrGrd.uorstCOHPrDocCta!codruc)) > 0 Then
        !codaux = frmTHPrGrd.uorstCOHPrDocCta!codruc
      Else
        !codaux = txtLlave(0).Text
      End If
    End If
    
    !serdoc = txtLlave(1).Text
    
    !codtdc = CODTDC_HPR
    !nrodoc = txtLlave(2).Text
    !feedoc = dtpDato(0).Value
    !fevdoc = dtpDato(0).Value
    !ferdoc = dtpDato(0).Value
    !RefDoc = txtDato(2).Text
    
    !GloIte = frmTHPrGrd.uorstCOHPrDocCta!glodet
    !GloItex = frmTHPrGrd.uorstCOHPrDocCta!glodetx
    
    !pdocpr = IIf(txtDato(27).Text = "", Null, txtDato(27).Text)
    !codcon = IIf(sContrato = "", Null, sContrato)
    
    If tbCCosto Then
      If (frmTHPrGrd.uorstCOHPrDocCCo!impcco_me > 0 Or frmTHPrGrd.uorstCOHPrDocCCo!impcco_mn > 0) Then
        !TpoCtb = IIf(frmTHPrGrd.uorstCOHPrDocCta!tpocnc = TPOCNC_TOT_HPR, IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        !TpoCtb = IIf(frmTHPrGrd.uorstCOHPrDocCta!tpocnc = TPOCNC_TOT_HPR, IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
    Else
      If (frmTHPrGrd.uorstCOHPrDocCta!impcta_me > 0 Or frmTHPrGrd.uorstCOHPrDocCta!impcta_mn > 0) Then
        !TpoCtb = IIf(frmTHPrGrd.uorstCOHPrDocCta!tpocnc = TPOCNC_TOT_HPR, IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      Else
        !TpoCtb = IIf(frmTHPrGrd.uorstCOHPrDocCta!tpocnc = TPOCNC_TOT_HPR, IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB))
      End If
    End If
    
    !tpomon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
    !ImpTCb = CDec(txtDato(4).Text)
    
    If tbCCosto Then
      !ImpMN = CDec(Abs(frmTHPrGrd.uorstCOHPrDocCCo!impcco_mn))
      !ImpME = CDec(Abs(frmTHPrGrd.uorstCOHPrDocCCo!impcco_me))
    Else
      !ImpMN = CDec(Abs(frmTHPrGrd.uorstCOHPrDocCta!impcta_mn))
      !ImpME = CDec(Abs(frmTHPrGrd.uorstCOHPrDocCta!impcta_me))
    End If
    !tpognr = TPOGNR_HPR
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
    Set frmTHPrGrd.uorstTemporal = frmTHPrGrd.uocnnMain.Execute(sSentencia)
    If Not (frmTHPrGrd.uorstTemporal.BOF Or frmTHPrGrd.uorstTemporal.EOF) And frmTHPrGrd.uorstTemporal.RecordCount > 0 Then
      While Not frmTHPrGrd.uorstTemporal.EOF
        ' Inicializo el orden de cuenta
        nOrden = IIf(sTpoCnc = frmTHPrGrd.uorstTemporal!tpocnc, nOrden, 0)
        sTpoCnc = frmTHPrGrd.uorstTemporal!tpocnc
        nImporteMN = CDec(txtDato(Val(sTpoCnc) + 4).Text)
        nImporteME = CDec(txtDato(Val(sTpoCnc) + 9).Text)
        ' Inserto las cuentas por compra
        If nImporteMN <> 0 Or nImporteME <> 0 Then
          nOrden = nOrden + 1
          nImpoCtaMN = Round(nImporteMN * (CDec(frmTHPrGrd.uorstTemporal!pordst) / 100), 2)
          nImpoCtaME = Round(nImporteME * (CDec(frmTHPrGrd.uorstTemporal!pordst) / 100), 2)
          With frmTHPrGrd.uorstCOHPrDocCta    'Cambiar RecordSet.
            .AddNew
            'Llaves.
            !codemp = gsCodEmp
            !pdoano = gsAnoAct
            !codaux = txtLlave(0).Text
            !serdoc = txtLlave(1).Text
            !nrodoc = txtLlave(2).Text
            !tpocnc = sTpoCnc
            !orden = Format(nOrden, "00")
            'Datos.
            !CodCta = frmTHPrGrd.uorstTemporal!CodCta
            !codruc = IIf(frmTHPrGrd.uorstTemporal!IndDoc = INDDOC_ACT, txtLlave(0).Text, Null)
            !glodet = IIf(txtDato(Choose(gsIdioma, 3, 25)).Text = "", Null, txtDato(Choose(gsIdioma, 3, 25)).Text)
            !glodetx = IIf(txtDato(Choose(gsIdioma, 25, 3)).Text = "", Null, txtDato(Choose(gsIdioma, 25, 3)).Text)
            !impcta_mn = CDec(nImpoCtaMN)
            !impcta_me = CDec(nImpoCtaME)
            !UsrCre = gsAbvUsr
            !FyHCre = Now
            .Update
          End With
          cmdMas(Val(sTpoCnc)).Tag = INDMASCTA_MAS
          ' Cuenta incial actualiza texto de costos
          If nOrden = 1 Then
            txtDato(Val(sTpoCnc) + (MINIMOINDICECUENTA - 1)).Text = frmTHPrGrd.uorstTemporal!CodCta
            cmdMas(Val(sTpoCnc)).Tag = INDMASCTA_INI
          End If
          upActualizaMas Val(sTpoCnc), cmdMas(Val(sTpoCnc)).Tag
          
          ' Inserto el centro de costo
          If ((Not IsNull(frmTHPrGrd.uorstTemporal!codcco)) And frmTHPrGrd.uorstTemporal!indcco = INDCCO_ACT) Then
            With frmTHPrGrd.uorstCOHPrDocCCo    'Cambiar RecordSet.
              .AddNew
              'Llaves.
              !codemp = gsCodEmp
              !pdoano = gsAnoAct
              !codaux = txtLlave(0).Text
              !serdoc = txtLlave(1).Text
              !nrodoc = txtLlave(2).Text
              !tpocnc = sTpoCnc
              !orden = Format(nOrden, "00")
              !CodCta = frmTHPrGrd.uorstTemporal!CodCta
              'Datos.
              !codcco = frmTHPrGrd.uorstTemporal!codcco
              !impcco_mn = CDec(nImpoCtaMN)
              !impcco_me = CDec(nImpoCtaME)
              !UsrCre = gsAbvUsr
              !FyHCre = Now
              .Update
            End With
            ' Cuenta incial actualiza texto de costos
            If nOrden = 1 Then
              txtDato(Val(sTpoCnc) + (MINIMOINDICECCOSTO - 1)).Text = frmTHPrGrd.uorstTemporal!codcco
            End If
          End If
        End If
        frmTHPrGrd.uorstTemporal.MoveNext
      Wend
    End If
    frmTHPrGrd.uorstTemporal.Close
  ElseIf n_TipoTran = INDCCO_INA Then
    sSentencia = "DELETE FROM CoHprDocCta "
    sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "AND codaux='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND serdoc='" & txtLlave(1).Text & "' "
    sSentencia = sSentencia & "AND nrodoc='" & txtLlave(2).Text & "'"
    frmTHPrGrd.uocnnMain.Execute sSentencia
  End If
  ppAbreCtaCCo

End Sub

Public Sub upActualizaMas(pnIndice As Byte, pnValor As Byte)
   frmTHPr.cmdMas(pnIndice).Tag = pnValor 'Necesaria la referencia por ser llamado externamente.
   With frmTHPrGrd.uorstMain
      Select Case pnIndice
      Case 1
         !IndCta_Bru = pnValor
      Case 2
         !IndCta_IR4 = pnValor
      Case 3
         !IndCta_IES = pnValor
      Case 4
         !IndCta_ORt = pnValor
      Case 5
         !IndCta_Net = pnValor
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
      frmTHPrGrd.uorstCoCta.MoveFirst
      frmTHPrGrd.uorstCoCta.Find "CodCta='" & txtDato(tnIndice + DIFERENCIAMASCUENTA).Text & "'"
      txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTHPrGrd.uorstCoCta!indcco = INDCCO_ACT)
      lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTHPrGrd.uorstCoCta!indcco = INDCCO_ACT)
      cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTHPrGrd.uorstCoCta!indcco = INDCCO_ACT)
   End If
End Sub

Private Sub ppDatosWhere()             'Cambiar.
  With frmTHPrGrd
    .uorstCOCpbCab.Requery
    .usConnStrgWher_COCpbDet = "WHERE COCpbDet.codemp='" & gsCodEmp & "' AND COCpbDet.pdoano='" & gsAnoAct & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='" & txtDato(0).Text & "' "
    .usConnStrgWher_COCpbDet = .usConnStrgWher_COCpbDet & "AND COCpbDet.NroCpb='" & txtDato(1).Text & "' "
    With .uorstCOCpbDet
      .Close
      .Source = frmTHPrGrd.usConnStrgSele_COCpbDet & frmTHPrGrd.usConnStrgWher_COCpbDet & frmTHPrGrd.usConnStrgOrde_COCpbDet
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
            .Item(dnNum).Caption = Choose(gsIdioma, "C.Costo", "Cost Center")
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
'       If (CDec(txtDato(MINIMOINDICEIMPORTEMN + dncontador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dncontador).Text) <> 0) And _
'        (cmdMas(dncontador + 1).Tag = INDMASCTA_INI And Len(Trim(txtDato(dncontador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
'        cmdMas(dncontador + 1).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dncontador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
       If ((CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) And _
        (cmdMas(dnContador + 1).Tag = INDMASCTA_MAS And Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) <> 0) Then
          ValidaCtasCCo = Not (txtDato(MINIMOINDICECUENTA + dnContador).Text = "")
          If Not ValidaCtasCCo Then Exit Function
          If frmTHPrGrd.ubGrabaMas = INDMASCTA_MAS Then
             With frmTHPrGrd.uorstCOHPrDocCta
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
                         With frmTHPrGrd.uorstCoCta
                            .MoveFirst
                            .Find "CodCta='" & frmTHPrGrd.uorstCOHPrDocCta!CodCta & "'"
                            If Not .EOF Then
                               dnIndCCo = frmTHPrGrd.uorstCoCta!indcco
                            End If
                         End With
                      End If
                      If dnIndCCo = INDCCO_ACT Then
                         With frmTHPrGrd.uorstCOHPrDocCCo
                            If .State = adStateOpen Then .Close
                            frmTHPrGrd.usConnStrgWher_COHPrDocCCo = "WHERE COHPrDocCCo.CodAux='" & frmTHPrGrd.uorstMain!codaux & "' And COHPrDocCCo.SerDoc='" & frmTHPrGrd.uorstMain!serdoc & "' And COHPrDocCCo.NroDoc='" & frmTHPrGrd.uorstMain!nrodoc & "' And COHPrDocCCo.TpoCnc='" & Trim(Str(dnContador + 1)) & "' AND COHPrDocCCo.Orden='" & frmTHPrGrd.uorstCOHPrDocCta!orden & "' AND COHPrDocCCo.CodCta='" & frmTHPrGrd.uorstCOHPrDocCta!CodCta & "' "
                            .Source = frmTHPrGrd.usConnStrgSele_COHPrDocCCo & frmTHPrGrd.usConnStrgWher_COHPrDocCCo & frmTHPrGrd.usConnStrgOrde_COHPrDocCCo
                            .Open
                            If .RecordCount = 0 Then
                               MsgBox Choose(gsIdioma, "Cuenta ", "Account ") & frmTHPrGrd.uorstCOHPrDocCta!CodCta & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
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
          With frmTHPrGrd.uorstCoCta
             .MoveFirst
             .Find "CodCta='" & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & "'"
             If Not .EOF Then
                dnIndCCo = frmTHPrGrd.uorstCoCta!indcco
             End If
          End With
          If dnIndCCo = INDCCO_ACT And txtDato(dnContador + 1 + DIFERENCIAMASCUENTA + CUENTASCONCCOSTO).Text = "" Then
             MsgBox Choose(gsIdioma, "Cuenta ", "Account ") & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & Choose(gsIdioma, " requiere C.Costo", " needs Cost Center"), vbInformation
'             MsgBox "Cuenta " & frmTHPrGrd.uorstCOHPrDocCta!CodCta & " requiere C.Costo", vbInformation
             ValidaCtasCCo = False
             Exit Function
          End If
       ElseIf Len(Trim(txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text)) = 0 And ((CDec(txtDato(MINIMOINDICEIMPORTEMN + dnContador).Text) <> 0) Or (CDec(txtDato(MINIMOINDICEIMPORTEME + dnContador).Text) <> 0)) Then
         ValidaCtasCCo = False
       End If
    Next dnContador
    ' Valido que los detalles sean iguales los importes
''    With frmTHPrGrd.uorstCOHPrDocCta
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

Private Function ValidaPedido(ByVal sPedido As String, ByVal sModificar As String) As Boolean
  Dim pnImporteMN As Double, pnImporteME As Double
  Dim pnImporte As Double, pnImporDiferen As Double
  Dim sCuenta As String, sCenCosto As String, sDetalle As String
  Dim sSentencia As String, sMoneda As String
  Dim sDetaCuenta As String, sDetaCenCosto As String
  Dim sMensage As String
    
  ValidaPedido = True
  sCuenta = "": sCenCosto = "": sDetalle = "": sMoneda = TPOMON_NAC
  sDetaCuenta = "": sDetaCenCosto = ""
  pnImporteMN = 0: pnImporteME = 0
  pnImporte = 0: pnImporDiferen = 0
    
  ' Movimientos por compras y honorarios
  With frmTHPrGrd                  'Cambiar Formulario de Grid.
    sSentencia = "SELECT p.codaux, p.pdocpr, "
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
    sSentencia = sSentencia & "LEFT JOIN cocprdoc a ON p.codemp=a.codemp AND p.codaux=a.codaux AND p.pdocpr=a.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND a.feedoc<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "') "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND a.feedoc<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103)) "
    End If
    sSentencia = sSentencia & "LEFT JOIN TGTDc b ON a.codemp=b.codemp AND a.CodTDc=b.CodTDc) "
    sSentencia = sSentencia & "LEFT JOIN cohprdoc d ON p.codemp=d.codemp AND p.codaux=d.codaux AND p.pdocpr=d.pdocpr "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND Concat(d.pdoano, d.serdoc, d.nrodoc)<>'" & gsAnoAct & txtLlave(1).Text & txtLlave(2).Text & "' "
      sSentencia = sSentencia & "AND d.feedoc<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "') "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND (d.pdoano+d.serdoc+d.nrodoc)<>'" & gsAnoAct & txtLlave(1).Text & txtLlave(2).Text & "' "
      sSentencia = sSentencia & "AND d.feedoc<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103)) "
    End If
    sSentencia = sSentencia & "LEFT JOIN TGTDc e ON d.codemp=e.codemp AND e.CodTDc='" & CODTDC_HPR & "') "
    sSentencia = sSentencia & "WHERE p.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND p.codaux='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND p.pdocpr='" & sPedido & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(p.pdoano, p.mespvs)", "(p.pdoano+p.mespvs)") & "<='" & gsAnoAct & gsMesAct & "' "
    sSentencia = sSentencia & "GROUP BY p.codemp, p.codaux, p.pdocpr "
    ' Obtengo los importes de movimientos
    Set .uorstTemporal = .uocnnMain.Execute(sSentencia)
    If Not (.uorstTemporal.BOF Or .uorstTemporal.EOF) And .uorstTemporal.RecordCount > 0 Then
      pnImporteMN = CDec(.uorstTemporal!impcpr_mn)
      pnImporteME = CDec(.uorstTemporal!impcpr_me)
    End If
    .uorstTemporal.Close
  End With
  
  ' Saldo de pedido
  With frmTHPrGrd                  'Cambiar Formulario de Grid.
    sSentencia = "SELECT a.pdocpr, " & Choose(gsIdioma, "a.detpdo", "a.detpdox") & " AS detpdo, a.tpomon, "
    sSentencia = sSentencia & "a.impmn, a.impme, a.impdife, x.codcta, x.codcco, "
    sSentencia = sSentencia & Choose(gsIdioma, "b.detcta", "b.detctax") & " AS detcta, "
    sSentencia = sSentencia & Choose(gsIdioma, "c.detcco", "c.detccox") & " AS detcco "
    sSentencia = sSentencia & "FROM copdocpr a "
    sSentencia = sSentencia & "INNER JOIN copdocprcta X ON a.codemp=x.codemp and a.pdoano=x.pdoano and a.mespvs=x.mespvs and a.coddpe=x.coddpe and a.pdocpr=x.pdocpr "
    sSentencia = sSentencia & "LEFT JOIN cocta b ON a.codemp=b.codemp AND x.codcta=b.codcta AND b.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "LEFT JOIN cocco c ON a.codemp=c.codemp AND x.codcco=c.codcco AND c.pdoano='" & gsAnoAct & "' "
    sSentencia = sSentencia & "WHERE a.codemp='" & gsCodEmp & "' "
    sSentencia = sSentencia & "AND " & IIf(ps_Plataforma = pSrvMySql, "CONCAT(a.pdoano, a.mespvs)", "(a.pdoano+a.mespvs)") & "<='" & gsAnoAct & gsMesAct & "' "
    sSentencia = sSentencia & "AND a.codaux='" & txtLlave(0).Text & "' "
    sSentencia = sSentencia & "AND a.pdocpr='" & sPedido & "' "
    If ps_Plataforma = pSrvMySql Then
      sSentencia = sSentencia & "AND a.fehpdo<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND a.fehpdo<= CONVERT(smalldatetime, '" & Format(dtpDato(0).Value, "dd/mm/yyyy") & "', 103) "
    End If
    sSentencia = sSentencia & "ORDER BY a.pdocpr"
    ' Obtengo los saldos dle pedido
    If .uorstTemporal.State = adStateOpen Then .uorstTemporal.Close
    Set .uorstTemporal = .uocnnMain.Execute(sSentencia)
    If Not (.uorstTemporal.BOF Or .uorstTemporal.EOF) And .uorstTemporal.RecordCount > 0 Then
      pnImporDiferen = CDec(.uorstTemporal!impdife) * -1
      pnImporteMN = CDec(.uorstTemporal!ImpMN) - pnImporteMN
      pnImporteME = CDec(.uorstTemporal!ImpME) - pnImporteME
      sCuenta = IIf(IsNull(.uorstTemporal!CodCta), "", .uorstTemporal!CodCta)
      sCenCosto = IIf(IsNull(.uorstTemporal!codcco), "", .uorstTemporal!codcco)
      sDetalle = IIf(IsNull(.uorstTemporal!detpdo), "", .uorstTemporal!detpdo)
      sMoneda = IIf(IsNull(.uorstTemporal!tpomon), "", .uorstTemporal!tpomon)
      sDetaCuenta = IIf(IsNull(.uorstTemporal!detcta), "", .uorstTemporal!detcta)
      sDetaCenCosto = IIf(IsNull(.uorstTemporal!detcco), "", .uorstTemporal!detcco)
    End If
    .uorstTemporal.Close
  End With
  ' Asigbo los importes de acuerdo a al moneda
  pnImporte = CDec(txtDato(IIf(sMoneda = TPOMON_NAC, MINIMOINDICEIMPORTEMN, MINIMOINDICEIMPORTEME)).Text)
  pnImporte = Round(IIf(sMoneda = TPOMON_NAC, pnImporteMN, pnImporteME) - pnImporte, 2)
  
  If pnImporte < pnImporDiferen Then
    lblDatoDeta(27).Caption = " " & sDetalle
    sMensage = TEXT_8006 & Choose(gsIdioma, " y/o excede importe de Pedido", " and/or exceeds amount of Order")
    MsgBox sMensage, vbExclamation
    ValidaPedido = False
    Exit Function
  End If
  
  ' Modifico los datos generales
  If sModificar = "S" Then
    lblDatoDeta(27).Caption = " " & sDetalle
    txtDato(MINIMOINDICECUENTA).Text = sCuenta
    txtDato(MINIMOINDICECCOSTO).Text = sCenCosto
    lblDatoDeta(MINIMOINDICECUENTA).Caption = " " & sDetaCuenta
    lblDatoDeta(MINIMOINDICECCOSTO).Caption = " " & sDetaCenCosto
  End If
        
End Function

']

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
    With frmTHPrGrd                  'Cambiar Formulario de Grid.
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
    lblDatoDeta(27).Caption = " " & sDetalle
    Exit Function
  End If
  ' Movimientos por compras
  With frmTHPrGrd                  'Cambiar Formulario de Grid.
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
      sSentencia = sSentencia & "AND Concat(a.pdoano, a.codtdc, a.serdoc, a.nrodoc)<>'" & gsAnoAct & "02" & txtLlave(1).Text & txtLlave(2).Text & "' "
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
  With frmTHPrGrd                  'Cambiar Formulario de Grid.
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
  lblDatoDeta(27).Caption = " " & sDetalle
  If pnImporte < pnImporDiferen Then
    sMensage = TEXT_8006 & Choose(gsIdioma, " y/o excede importe de Pedido", " and/or exceeds amount of Order")
    MsgBox sMensage, vbExclamation
    pfValidaPedido = False
    Exit Function
  End If
  
  ' Modifico los datos generales
  If sModificar = "S" And txtDato(27).Tag <> sPedido Then
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
      sSentencia = sSentencia & "AND Concat(cpr.pdoano, cpr.codtdc, cpr.serdoc, cpr.nrodoc)<>'" & gsAnoAct & "02" & txtLlave(1).Text & txtLlave(2).Text & "' "
      sSentencia = sSentencia & "AND cpr.feedoc<='" & Format(dtpDato(0).Value, "yyyy-mm-dd") & "' "
    ElseIf ps_Plataforma = pSrvSql Then
      sSentencia = sSentencia & "AND (cpr.pdoano+cpr.codtdc+cpr.serdoc+cpr.nrodoc)<>'" & gsAnoAct & "02" & txtLlave(1).Text & txtLlave(2).Text & "' "
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
    With frmTHPrGrd                  'Cambiar Formulario de Grid.
      Set .uorstTemporal = .uocnnMain.Execute(sSentencia)
      ' Elimino las cuentas y centro de costos
      sSentencia = "DELETE FROM cocprdoccta "
      sSentencia = sSentencia & "WHERE codemp='" & gsCodEmp & "' AND pdoano='" & gsAnoAct & "' "
      sSentencia = sSentencia & "AND codaux='" & txtLlave(0).Text & "' AND codtdc='" & "02" & "' "
      sSentencia = sSentencia & "AND serdoc='" & txtLlave(1).Text & "' AND nrodoc='" & txtLlave(2).Text & "' "
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
            sSentencia = "INSERT INTO cohprdoccta(codemp, pdoano, codaux, serdoc, nrodoc, tpocnc, orden, codcta, glodet, glodetx, codruc, impcta_mn, impcta_me, usrcre, fyhcre) "
            sSentencia = sSentencia & " VALUES("
            sSentencia = sSentencia & "'" & gsCodEmp & "', "
            sSentencia = sSentencia & "'" & gsAnoAct & "', "
            sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
            'sSentencia = sSentencia & "'" & "02" & "', "
            sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
            sSentencia = sSentencia & "'" & txtLlave(2).Text & "', "
            nOrden = nOrden + 1
            sSentencia = sSentencia & "'1', '" & Format(nOrden, "00") & "', "
            sSentencia = sSentencia & "'" & .uorstTemporal!CodCta & "', "
            sSentencia = sSentencia & IIf(txtDato(Choose(gsIdioma, 3, 26)).Text = "", "Null", "'" & txtDato(Choose(gsIdioma, 3, 26)).Text & "'") & ", "
            sSentencia = sSentencia & IIf(txtDato(Choose(gsIdioma, 26, 3)).Text = "", "Null", "'" & txtDato(Choose(gsIdioma, 26, 3)).Text & "'") & ", "
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
              sSentencia = "INSERT INTO cohprdoccco(codemp, pdoano, codaux,serdoc, nrodoc, tpocnc, orden, codcta, codcco, impcco_mn, impcco_me, usrcre, fyhcre) "
              sSentencia = sSentencia & " VALUES("
              sSentencia = sSentencia & "'" & gsCodEmp & "', "
              sSentencia = sSentencia & "'" & gsAnoAct & "', "
              sSentencia = sSentencia & "'" & txtLlave(0).Text & "', "
              'sSentencia = sSentencia & "'" & "02" & "', "
              sSentencia = sSentencia & "'" & txtLlave(1).Text & "', "
              sSentencia = sSentencia & "'" & txtLlave(2).Text & "', "
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
      frmTHPrGrd.uorstTemporal.Close
    End With
    txtDato(MINIMOINDICEIMPORTEMN).Text = Format(pnImporteMN, FORMATO_NUM_1)
    txtDato(MINIMOINDICEIMPORTEME).Text = Format(pnImporteME, FORMATO_NUM_1)
    txtDato(MINIMOINDICECUENTA).Text = sCuenta
    txtDato(MINIMOINDICECCOSTO).Text = sCenCosto
    lblDatoDeta(MINIMOINDICECUENTA).Caption = " " & sDetaCuenta
    lblDatoDeta(MINIMOINDICECCOSTO).Caption = " " & sDetaCenCosto
  End If
        
End Function

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

