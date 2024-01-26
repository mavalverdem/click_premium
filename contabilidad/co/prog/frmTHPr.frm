VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTHPr 
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
   Begin VB.CommandButton cmdAuxiliar 
      Caption         =   "&Proveedor"
      Height          =   375
      Left            =   8250
      TabIndex        =   98
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkAfeIndORt 
      Caption         =   "Afecto O&tras Ret."
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   4980
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox chkAfeIndIES 
      Caption         =   "Afecto I.&E.S."
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox chkIndAfeIR4 
      Caption         =   "Afecto I.&R.4ª"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   60
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox chkIndPreGen 
      Caption         =   "Cuentas Re&gistradas"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   6960
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CheckBox chkCalcularIR4 
      Caption         =   "C&alcular"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   1380
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3360
      Width           =   915
   End
   Begin VB.CheckBox chkCalcularIES 
      Caption         =   "Ca&lcular"
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3360
      Width           =   915
   End
   Begin VB.Frame fraDiario 
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   93
      Top             =   2340
      Width           =   7155
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   4860
         Picture         =   "frmTHPr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   94
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
      TabIndex        =   84
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
         Picture         =   "frmTHPr.frx":01AA
         Style           =   1  'Graphical
         TabIndex        =   45
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
         Picture         =   "frmTHPr.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Picture         =   "frmTHPr.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   46
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
         Picture         =   "frmTHPr.frx":0648
         Style           =   1  'Graphical
         TabIndex        =   47
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
         Picture         =   "frmTHPr.frx":074A
         Style           =   1  'Graphical
         TabIndex        =   48
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
         Picture         =   "frmTHPr.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   49
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
      TabIndex        =   8
      Top             =   1740
      Width           =   735
   End
   Begin VB.CommandButton cmdLlaveAyud 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   7680
      Picture         =   "frmTHPr.frx":0996
      Style           =   1  'Graphical
      TabIndex        =   66
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
      ItemData        =   "frmTHPr.frx":0B40
      Left            =   1020
      List            =   "frmTHPr.frx":0B42
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1740
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   0
      Left            =   3060
      TabIndex        =   4
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62062593
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      Left            =   1320
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   480
      Width           =   435
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   2655
      Left            =   0
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3720
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   441
      ForeColor       =   8388608
      TabCaption(0)   =   "I&mportes"
      TabPicture(0)   =   "frmTHPr.frx":0B44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
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
      TabPicture(1)   =   "frmTHPr.frx":0B60
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
         TabIndex        =   39
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
         TabIndex        =   34
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
         TabIndex        =   29
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
         TabIndex        =   24
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
         TabIndex        =   19
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
         Index           =   12
         Left            =   1320
         TabIndex        =   30
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
         TabIndex        =   25
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
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   24
         Left            =   9060
         Picture         =   "frmTHPr.frx":0B7C
         Style           =   1  'Graphical
         TabIndex        =   91
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
         TabIndex        =   43
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   23
         Left            =   9060
         Picture         =   "frmTHPr.frx":0D26
         Style           =   1  'Graphical
         TabIndex        =   89
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
         TabIndex        =   38
         Top             =   1500
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   22
         Left            =   9060
         Picture         =   "frmTHPr.frx":0ED0
         Style           =   1  'Graphical
         TabIndex        =   87
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
         TabIndex        =   33
         Top             =   1200
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   21
         Left            =   9060
         Picture         =   "frmTHPr.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   85
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
         TabIndex        =   28
         Top             =   900
         Width           =   675
      End
      Begin VB.CheckBox chkMonedaActiva 
         Caption         =   "M&oneda activa"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1320
         TabIndex        =   18
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
         TabIndex        =   81
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
            TabIndex        =   83
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
            TabIndex        =   82
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
         Picture         =   "frmTHPr.frx":1224
         TabIndex        =   41
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
         Picture         =   "frmTHPr.frx":1326
         TabIndex        =   36
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
         Picture         =   "frmTHPr.frx":1428
         TabIndex        =   31
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
         Picture         =   "frmTHPr.frx":152A
         TabIndex        =   26
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
         Picture         =   "frmTHPr.frx":162C
         TabIndex        =   21
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
         TabIndex        =   23
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   20
         Left            =   9060
         Picture         =   "frmTHPr.frx":172E
         Style           =   1  'Graphical
         TabIndex        =   78
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
         TabIndex        =   42
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   19
         Left            =   6540
         Picture         =   "frmTHPr.frx":18D8
         Style           =   1  'Graphical
         TabIndex        =   76
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
         TabIndex        =   37
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   18
         Left            =   6540
         Picture         =   "frmTHPr.frx":1A82
         Style           =   1  'Graphical
         TabIndex        =   74
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
         TabIndex        =   32
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   17
         Left            =   6540
         Picture         =   "frmTHPr.frx":1C2C
         Style           =   1  'Graphical
         TabIndex        =   72
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
         Picture         =   "frmTHPr.frx":1DD6
         Style           =   1  'Graphical
         TabIndex        =   70
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
         TabIndex        =   27
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
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   285
         Index           =   15
         Left            =   6540
         Picture         =   "frmTHPr.frx":1F80
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   625
         Width           =   255
      End
      Begin VB.CheckBox chkDesactivar 
         Caption         =   "Desactivar Cue&ntas"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   4980
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgrDetalle 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   65
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
         TabIndex        =   92
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
         TabIndex        =   90
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
         TabIndex        =   88
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
         TabIndex        =   86
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
         TabIndex        =   79
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
         TabIndex        =   77
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
         TabIndex        =   75
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
         TabIndex        =   73
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
         TabIndex        =   71
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
         TabIndex        =   69
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label10 
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
         Left            =   195
         TabIndex        =   64
         Top             =   1860
         Width           =   780
      End
      Begin VB.Label Label7 
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
         Left            =   180
         TabIndex        =   63
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Left            =   180
         TabIndex        =   62
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   61
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label21 
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
         Left            =   180
         TabIndex        =   60
         Top             =   660
         Width           =   1005
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
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   345
         Width           =   2265
      End
   End
   Begin MSComCtl2.DTPicker dtpDato 
      Height          =   315
      Index           =   3
      Left            =   1020
      TabIndex        =   3
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62062593
      CurrentDate     =   37102
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
      TabIndex        =   80
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
      TabIndex        =   67
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
      TabIndex        =   57
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
      TabIndex        =   56
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
      TabIndex        =   55
      Top             =   1800
      Width           =   615
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
      TabIndex        =   54
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
      TabIndex        =   53
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label12 
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
      Left            =   60
      TabIndex        =   52
      Top             =   540
      Width           =   765
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
      TabIndex        =   51
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
']

Private Sub Form_Load()
   pbValidada = False
   pbFecha = True
   Me.KeyPreview = True
   
   With frmTHPrGrd                     'Cambiar Formulario de Grid.
    '[Llaves                              'Cambiar
      txtLlave(0).MaxLength = .uorstMain!CodAux.DefinedSize
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
      dtpDato(3).Tag = dtpDato(3).Value
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
   If gbCieHpr Then
      MsgBox TEXT_9016, vbCritical
      Exit Sub
   End If
   
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
 '[Dato con el foco al corregir.       'Cambiar.
   dtpDato(3).SetFocus
 ']
End Sub

Public Sub cmdGrabar_Click()
   On Error GoTo Err

 '[Propio del formulario.
   Dim dnSumaMN As Double, _
       dnSumaME As Double
   
'   If txtDato(33).Text = "" Then
'      MsgBox TEXT_6002, vbCritical
'      txtDato(33).SetFocus
'      Exit Sub
'   End If
   
   With frmTHPrGrd.uorstMain
      dnSumaMN = CDec(txtDato(5).Text) - CDec(txtDato(6).Text) - CDec(txtDato(7).Text) - CDec(txtDato(8).Text)
      dnSumaME = CDec(txtDato(10).Text) - CDec(txtDato(11).Text) - CDec(txtDato(12).Text) - CDec(txtDato(13).Text)
'      If gfRedond(CDec(TxtDato(5).Text) - CDec(TxtDato(6).Text) - CDec(TxtDato(7).Text) - CDec(TxtDato(8).Text), 2) <> CDec(TxtDato(9).Text) Then
      If dnSumaMN <> CDec(txtDato(9).Text) Then
'         If MsgBox(TEXT_9011, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
         If MsgBox(TEXT_9011 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(dnSumaMN - CDec(txtDato(9).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
            Exit Sub
         End If
'      ElseIf gfRedond(CDec(TxtDato(10).Text) - CDec(TxtDato(11).Text) - CDec(TxtDato(12).Text) - CDec(TxtDato(13).Text), 2) <> CDec(TxtDato(14).Text) Then
      ElseIf dnSumaME <> CDec(txtDato(14).Text) Then
'         If MsgBox(TEXT_9012, vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
         If MsgBox(TEXT_9012 & Chr(13) & "La diferencia es de " & Trim(CStr(Abs(dnSumaME - CDec(txtDato(14).Text)))) & ".", vbOKCancel + vbInformation + vbDefaultButton2) = vbCancel Then
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
   Case 0, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1
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
   If pbValidada Then
'      dtpDato(3).SetFocus 'Cambiar.
      
      txtLlave(0).Enabled = False
      txtLlave(1).Enabled = False
      txtLlave(2).Enabled = False
      lblLlaveDeta(0).Enabled = False
      cmdLlaveAyud(0).Enabled = False
   End If
   If pbValidada And dtpDato(3).Enabled Then dtpDato(3).SetFocus 'Cambiar.
End Sub

Private Sub txtLlave_Validate(Index As Integer, Cancel As Boolean)
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
         Set .uorstTemporal = .uocnnMain.Execute("SELECT MesPvs FROM COHPrDoc WHERE CodAux='" & txtLlave(0).Text & "' AND SerDoc='" & txtLlave(1).Text & "' AND NroDoc='" & txtLlave(2).Text & "'")
         If .uorstTemporal.RecordCount > 0 Then
            MsgBox TEXT_8007 & Chr(13) & "(mes " & gfMesLet("01" & .uorstTemporal!MesPvs & gsAnoAct, 0, "", 1, "", 0) & ")", vbExclamation
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
         MsgBox "No se ha ingresado Tipo de Cambio para Esta Fecha", vbCritical
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
         If chkMonedaActiva.Value = vbChecked Then
            If cboTpoMon.ListIndex = TPOMON_NAC_IND Then
               If CDec(txtDato(Index + 1).Text) > 0 Then txtDato((Index + CANTIDADIMPORTES) + 1).Text = Format(gfRedond(CDec(txtDato(Index + 1).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
               If CDec(txtDato(Index + 2).Text) > 0 Then txtDato((Index + CANTIDADIMPORTES) + 2).Text = Format(gfRedond(CDec(txtDato(Index + 2).Text) / CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
            Else
               If CDec(txtDato(Index + 1).Text) > 0 Then txtDato((Index - CANTIDADIMPORTES) + 1).Text = Format(gfRedond(CDec(txtDato(Index + 1).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
               If CDec(txtDato(Index + 2).Text) > 0 Then txtDato((Index - CANTIDADIMPORTES) + 2).Text = Format(gfRedond(CDec(txtDato(Index + 2).Text) * CDec(txtDato(4).Text), 2), FORMATO_NUM_1)
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
                     If frmTHPrGrd.uorstCOHPrDocCta!CodAux = txtLlave(0).Text And _
                       frmTHPrGrd.uorstCOHPrDocCta!SerDoc = txtLlave(1).Text And _
                       frmTHPrGrd.uorstCOHPrDocCta!NroDoc = txtLlave(2).Text And _
                       frmTHPrGrd.uorstCOHPrDocCta!TpoCnc = Trim(Str(Index - MINIMOINDICECUENTA + 1)) Then
'                        frmTHPrGrd.uorstCOHPrDocCCo.MoveFirst
'                        Do
'                           If frmTHPrGrd.uorstCOHPrDocCCo!CodAux = txtLlave(0).Text And _
                             frmTHPrGrd.uorstCOHPrDocCCo!CodTDc = txtLlave(1).Text And _
                             frmTHPrGrd.uorstCOHPrDocCCo!SerDoc = txtLlave(2).Text And _
                             frmTHPrGrd.uorstCOHPrDocCCo!NroDoc = txtLlave(3).Text And _
                             frmTHPrGrd.uorstCOHPrDocCCo!TpoCnc = Trim(Str(Index - MINIMOINDICECUENTA + 1)) And _
                             frmTHPrGrd.uorstCOHPrDocCCo!CodCta = frmTHPrGrd.uorstCOHPrDocCta!CodCta Then
'                              frmTHPrGrd.uorstCOHPrDocCCo.Delete
'                           End If
'                           frmTHPrGrd.uorstCOHPrDocCCo.MoveNext
'                        Loop Until frmTHPrGrd.uorstCOHPrDocCCo.EOF
'                        frmTHPrGrd.uorstCOHPrDocCCo.Requery
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
'            txtDato(Index + CUENTASCONCCOSTO).Enabled = True
'            cmdDatoAyud(Index + CUENTASCONCCOSTO).Enabled = True
      End If
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
   Case 0, MINIMOINDICECUENTA To MINIMOINDICECCOSTO + CANTIDADIMPORTES - 1 'Cambiar (añadir índices).
      Cancel = ppAyuDet(AYUDAT, Index)
      If Cancel Then Exit Sub
         
      If Index >= MINIMOINDICECUENTA And Index < MINIMOINDICECCOSTO Then
'         If frmTHPrGrd.uorstCOCta.RecordCount > 0 And txtDato(Index + CUENTASCONCCOSTO).Text <> "" Then
         If frmTHPrGrd.uorstCOCta.RecordCount > 0 Then
            If Not frmTHPrGrd.uorstCOCta.EOF Then
                If frmTHPrGrd.uorstCOCta!IndCCo = INDCCO_ACT Then
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
'      If frmTHPrGrd.ubGrabaMas = 0 Then
'         frmTHPrGrd.ubGrabaMas = 1
'         With frmTHPrGrd
'           If pbNuevo Then
'              .uorstMain.AddNew
'           End If
'           upDatosDesconectados 0
'           .uorstMain.Update
'        End With
'     End If
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
         With frmTHPrGrd.uorstTGAux
            If .RecordCount > 0 Then .MoveFirst
            .Find "CodAux='" & txtLlave(tnIndex).Text & "'"
            If .EOF Then
               MsgBox TEXT_8006, vbExclamation
               ppAyuDet = True
            Else
               lblLlaveDeta(tnIndex).Caption = " " & !RazAux
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
         With frmTHPrGrd.uorstCODro
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
         With frmTHPrGrd.uorstCOCta
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
         With frmTHPrGrd.uorstCOCCo
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
      End Select
   End If
End Function

Public Sub upDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

 '[Propio del formulario.
   Dim dnContador As Byte
 ']

   With frmTHPrGrd.uorstMain           'Cambiar RecordSet.
      If tnFase = 0 Then
        'Llaves.
         If pbNuevo Then
            !CodAux = txtLlave(0).Text
            !SerDoc = txtLlave(1).Text
            !NroDoc = txtLlave(2).Text
            !MesPvs = gsMesAct
            !PctIR4 = CDec(gnPctIR4)
            !PctIES = CDec(gnPctIES)
         End If

        'Datos.
         !TpoMon = Choose(cboTpoMon.ListIndex + 1, TPOMON_NAC, TPOMON_EXT)
         !IndPreGen = IIf(chkIndPreGen.Value = vbChecked, INDPREGEN_ACT, INDPREGEN_INA)
'[REVISAR.
!IndAfeIES = chkAfeIndIES.Value
!IndAfeIR4 = chkIndAfeIR4.Value
!IndAfeORt = chkAfeIndORt.Value
']REVISAR.
'         !CodSoc = IIf(dcoSocio.BoundText = "", Null, dcoSocio.BoundText)
         !FehOpe = dtpDato(3).Value
         !FeEDoc = dtpDato(0).Value
'         !Tf1Cta = mskDato(0).Text
'         !CodMon = optTpoMon(1).Value
         !CodDro = IIf(txtDato(0).Text = "", Null, txtDato(0).Text)
         !NroCpb = txtDato(1).Text
         !RefDoc = txtDato(2).Text
         !GloDoc = txtDato(3).Text
         !ImpTCb = txtDato(4).Text
         !ImpBru_MN = txtDato(5).Text
         !ImpIR4_MN = txtDato(6).Text
         !ImpIES_MN = txtDato(7).Text
         !ImpORt_MN = txtDato(8).Text
         !ImpNet_MN = txtDato(9).Text
         !ImpBru_ME = txtDato(10).Text
         !ImpIR4_ME = txtDato(11).Text
         !ImpIES_ME = txtDato(12).Text
         !ImpORt_ME = txtDato(13).Text
         !ImpNet_ME = txtDato(14).Text

       '[Actualización de datos por Cuentas y Centros de Costo registrados directamente.
         ppAbreCtaCCo
         For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
            If cmdMas(dnContador).Tag = INDMASCTA_CTA And Val(txtDato(dnContador + DIFERENCIAMASCUENTA).Text) <> 1 Then
               With frmTHPrGrd.uorstCOHPrDocCta
                  .MoveFirst
                  .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & "'"
                  If Not .EOF Then
                     .Delete
                     .Update
                     .Requery
                     frmTHPrGrd.uorstCOHPrDocCCo.Requery
                     Call upActualizaMas(dnContador, INDMASCTA_INI)
                  End If
               End With
            End If
            
            If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Or _
               cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCUENTA).Text)) <> 0 Then
               With frmTHPrGrd.uorstCOHPrDocCta
                  If .RecordCount <> 0 Then .MoveFirst
                  .Find "cLlave='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & "'"
                  If .EOF Then
                     .AddNew
                     !CodAux = txtLlave(0).Text
                     !SerDoc = txtLlave(1).Text
                     !NroDoc = txtLlave(2).Text
                     !TpoCnc = dnContador
                     !UsrCre = gsAbvUsr
                     !FyHCre = Now
                  Else
                     !UsrMdf = gsAbvUsr
'                     !FyHMdf = Now
                  End If
                  !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
                  !ImpCta_MN = txtDato(dnContador + DIFERENCIAMASIMPORTE).Text
                  !ImpCta_ME = txtDato(dnContador + DIFERENCIAMASIMPORTE + CANTIDADIMPORTES).Text
                  .Update
               End With
               If cmdMas(dnContador).Tag = INDMASCTA_INI And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Or _
                  cmdMas(dnContador).Tag = INDMASCTA_CTA And Len(Trim(txtDato(dnContador + DIFERENCIAMASCCOSTO).Text)) <> 0 Then
                  With frmTHPrGrd.uorstCOHPrDocCCo
                     If .RecordCount <> 0 Then .MoveFirst
                     .Find "cLlave1='" & txtLlave(0).Text & txtLlave(1).Text & txtLlave(2).Text & dnContador & txtDato(dnContador + DIFERENCIAMASCUENTA).Text & "'"
                     If .EOF Then
                        .AddNew
                        !CodAux = txtLlave(0).Text
                        !SerDoc = txtLlave(1).Text
                        !NroDoc = txtLlave(2).Text
                        !TpoCnc = dnContador
                        !CodCta = txtDato(dnContador + DIFERENCIAMASCUENTA).Text
                        !UsrCre = gsAbvUsr
                        !FyHCre = Now
                     Else
                        !UsrMdf = gsAbvUsr
'                        !FyHMdf = Now
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
         txtLlave(0).Text = !CodAux
         txtLlave(1).Text = !SerDoc
         txtLlave(2).Text = !NroDoc

        'Datos.
         cboTpoMon.ListIndex = IIf(!TpoMon = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
         chkIndPreGen.Value = IIf(!IndPreGen = INDPREGEN_ACT, vbChecked, vbUnchecked)
'[Revisar.
chkAfeIndIES.Value = IIf(!IndAfeIES, vbChecked, vbUnchecked)
chkIndAfeIR4.Value = IIf(!IndAfeIR4, vbChecked, vbUnchecked)
chkAfeIndORt.Value = IIf(!IndAfeORt, vbChecked, vbUnchecked)
']Revisar.
'         dcoSocio.BoundText = IIf(IsNull(!CodSoc), "", !CodSoc)
         dtpDato(3).Value = !FehOpe
         dtpDato(0).Value = !FeEDoc
'         optTpoMon(1).Value = uorstMain!CodMon
'         mskDato(0).Text = IIf(IsNull(.uorstMain!Tf1Cta), "", .uorstMain!Tf1Cta)
         txtDato(0).Text = IIf(IsNull(!CodDro), "", !CodDro)
         txtDato(1).Text = IIf(IsNull(!NroCpb), "", !NroCpb)
         txtDato(2).Text = IIf(IsNull(!RefDoc), "", !RefDoc)
         txtDato(3).Text = IIf(IsNull(!GloDoc), "", !GloDoc)
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
         Next
      
       '[Busca detalle de códigos      'Cambiar (habilitar/deshabilitar).
         ppAyuDet AYULLA, 0
      '   ppAyuDet AYULLA, 1
         ppAyuDet AYUDAT, 0
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
                        ppAyuDet AYUDAT, dnContador + DIFERENCIAMASCUENTA
                        With frmTHPrGrd.uorstCOHPrDocCCo
                           If .RecordCount > 0 Then
                              .MoveFirst
                              .Find "cLlave = " & dnContador & frmTHPrGrd.uorstCOHPrDocCta!CodCta
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
   
   For dnContador = MINIMOINDICEMAS To MINIMOINDICEMAS + CANTIDADIMPORTES - 1
      Call upHabilitaCuenta(False, dnContador)
   Next

  'Ayudas.
   cmdDatoAyud(0).Enabled = tbHabilitar
   lblDatoDeta(0).Enabled = tbHabilitar

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
'   frmTHPrGrd.uorstCOCpbDet.Requery
'   DatosGrid
' ']
'End Sub

'Private Sub cmdBorrar_Click()
'   If Not frmTHPrGrd.uorstCOCpbCab.EOF Then
'      frmTHPrGrd.uorstCOCpbCab.Delete
'      uorstMain!IndGen = False
'      frmTHPrGrd.uorstCOCpbCab.Update
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
'   If frmTHPrGrd.ubGrabaMas = 0 Then
'      frmTHPrGrd.ubGrabaMas = 1
'      With frmTHPrGrd
'         If pbNuevo Then
'            .uorstMain.AddNew
'         End If
'         upDatosDesconectados 0
'         .uorstMain.Update
'      End With
'   End If

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
   If Index = 3 Then
      If Month(dtpDato(3).Value) > Val(gsMesAct) And Year(dtpDato(3).Value) >= Val(gsAnoAct) Then
         MsgBox "La fecha No Corresponde al Periodo de Operacion", vbCritical
         dtpDato(Index).SetFocus
         Cancel = True
         Exit Sub
      End If
      If dtpDato(3).Tag <> dtpDato(3).Value Then
         dtpDato(3).Tag = dtpDato(3).Value
         With frmTHPrGrd.uorstTGTCb
            If .RecordCount <> 0 Then .MoveFirst
            .Find "FehTCb = '" & dtpDato(3).Value & "'"
            If .EOF Then
               MsgBox "No se ha ingresado Tipo de Cambio para esta fecha.", vbCritical
               Cancel = True
               Exit Sub
            Else
'               uorstMain!ImpTCb = IIf(frmTHPrGrd.uorstCOCta!TpoTCb = TPOTCB_CPR, !ImpTCb_Cpr, !ImpTCb_Vta)
''               frmTHPrGrd.uorstMain!ImpTCb = !ImpTCb_Vta
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
   With frmTHPrGrd.uorstCOHPrDocCCo
      frmTHPrGrd.usConnStrgWher_COHPrDocCCo = "WHERE COHPrDocCCo.CodAux='" & frmTHPrGrd.uorstMain!CodAux & "' AND COHPrDocCCo.SerDoc='" & frmTHPrGrd.uorstMain!SerDoc & "' AND COHPrDocCCo.NroDoc='" & frmTHPrGrd.uorstMain!NroDoc & "' "
      If .State = adStateOpen Then .Close
      .Source = frmTHPrGrd.usConnStrgSele_COHPrDocCCo & frmTHPrGrd.usConnStrgWher_COHPrDocCCo & frmTHPrGrd.usConnStrgOrde_COHPrDocCCo
      .Open
      .Properties("Unique Table").Value = "COHPrDocCCo"
   End With
   With frmTHPrGrd.uorstCOHPrDocCta
      frmTHPrGrd.usConnStrgWher_COHPrDocCta = "WHERE COHPrDocCta.CodAux='" & frmTHPrGrd.uorstMain!CodAux & "' AND COHPrDocCta.SerDoc='" & frmTHPrGrd.uorstMain!SerDoc & "'  AND COHPrDocCta.NroDoc='" & frmTHPrGrd.uorstMain!NroDoc & "' "
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

   If txtDato(1).Text <> "" Then
      ppDatosWhere

      With frmTHPrGrd.uorstCOCpbCab
        'Si existe, elimina Comprobante existente.
         If .RecordCount > 0 Then
            .MoveFirst
            .Find "cLlave='" & txtDato(0).Text & txtDato(1).Text & "'"
            If Not .EOF Then .Delete
         End If
      End With
   End If

   With frmTHPrGrd.uorstCOCpbCab
     'Si no está marcado para generar, marca el documento como no generado.
      If chkIndPreGen.Value = vbUnchecked Then
         frmTHPrGrd.uorstMain!IndGen = False
         frmTHPrGrd.uorstMain.Update
         Exit Sub
      End If

     'Captura del Siguiente Número.
      If txtDato(1).Text = "" Then
         With frmTHPrGrd.uorstCODro
            If .RecordCount <> 0 Then .MoveFirst
            .Find "CodDro = '" & txtDato(0).Text & "'"
            If IsNull(.Fields(2).Value) Then .Fields(2).Value = gfCeros("", .Fields(2).DefinedSize, 0, "0")
            txtDato(1).Text = gfCeros(.Fields(2).Value, .Fields(2).DefinedSize, 1, "0")
            .Fields(2).Value = txtDato(1).Text
            .Update
         End With
         frmTHPrGrd.uorstMain!NroCpb = txtDato(1).Text
         frmTHPrGrd.uorstMain.Update
      End If
      
      ppDatosWhere
   
     'Si no hay cuentas, marca el documento como no generado.
      If frmTHPrGrd.uorstCOHPrDocCta.RecordCount = 0 Then
         frmTHPrGrd.uorstMain!IndGen = False
         frmTHPrGrd.uorstMain.Update
         Exit Sub
      End If

     'Crea encabezado de Comprobante.
      .AddNew
      !MesPvs = gsMesAct
      !CodDro = txtDato(0).Text
      !NroCpb = txtDato(1).Text
      !FehCpb = dtpDato(3).Value
      !TpoGnr = TPOGNR_HPR
      !IndNCu = INDNCU_FAL
      !GloCpb = txtDato(3).Text
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With

   With frmTHPrGrd.uorstCOHPrDocCta
     'Crea ítemes de Comprobante.
      .MoveFirst
      Do
         dbProcesaCuenta = True

        'Itemes con Centro de Costo.
         If !TpoCnc <= CUENTASCONCCOSTO Then
            With frmTHPrGrd.uorstCOHPrDocCCo
               If .RecordCount <> 0 Then
                  .MoveFirst
                  .Find "cLlave = " & Trim(frmTHPrGrd.uorstCOHPrDocCta!TpoCnc) & frmTHPrGrd.uorstCOHPrDocCta!CodCta
                  If Not .EOF Then
                     Do
                        dnNumeroItem = dnNumeroItem + 1
                        Call ppGenera1(True, dnNumeroItem)
                        .MoveNext
                        If .EOF Then Exit Do
                        If !cLlave <> Trim(frmTHPrGrd.uorstCOHPrDocCta!TpoCnc) & frmTHPrGrd.uorstCOHPrDocCta!CodCta Then Exit Do
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

   frmTHPrGrd.uorstMain!IndGen = True
   txtDato(0).Enabled = False
   txtDato(1).Enabled = False
   cmdDatoAyud(0).Enabled = False
   lblDatoDeta(0).Enabled = False

   frmTHPrGrd.uorstCOCpbCab.Update
   frmTHPrGrd.uorstCOCpbDet.UpdateBatch
   frmTHPrGrd.uorstMain.Update
End Sub

Private Sub ppGenera1(tbCCosto As Boolean, tnNumeroItem As Integer)
   With frmTHPrGrd.uorstCOCpbDet
      .AddNew
      !CodDro = txtDato(0).Text
      !NroCpb = txtDato(1).Text
      !NroIte = tnNumeroItem
      !MesPvs = gsMesAct
      !CodCta = frmTHPrGrd.uorstCOHPrDocCta!CodCta
      !FehOpe = dtpDato(3).Value
      frmTHPrGrd.uorstCOCta.MoveFirst
      frmTHPrGrd.uorstCOCta.Find "CodCta='" & frmTHPrGrd.uorstCOHPrDocCta!CodCta & "'"
      If frmTHPrGrd.uorstCOCta!IndCCo = INDCCO_ACT Then If tbCCosto Then !CodCCo = frmTHPrGrd.uorstCOHPrDocCCo!CodCCo
      If frmTHPrGrd.uorstCOCta!IndDoc = INDDOC_ACT Then
         !CodAux = txtLlave(0).Text
         !SerDoc = txtLlave(1).Text
         !CodTDc = CODTDC_HPR
         !NroDoc = txtLlave(2).Text
         !FeEDoc = dtpDato(0).Value
         !FeVDoc = dtpDato(0).Value
         !FeRDoc = dtpDato(0).Value
         !RefDoc = txtDato(2).Text
      End If
      !GloIte = txtDato(3).Text
      !TpoCtb = IIf(frmTHPrGrd.uorstCOHPrDocCta!TpoCnc = TPOCNC_TOT_HPR, IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_POS, TPOCTB_DEB, TPOCTB_HAB), IIf(frmTHPrGrd.uorstTGTDc!SgnTDc = SGNTDC_NEG, TPOCTB_DEB, TPOCTB_HAB))
      !TpoMon = IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      !ImpTCb = txtDato(4).Text
      If tbCCosto Then
'         !ImpMN = frmTHPrGrd.uorstCOHPrDocCCo!ImpCCo * IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, txtDato(4).Text)
'         !ImpME = frmTHPrGrd.uorstCOHPrDocCCo!ImpCCo / IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, txtDato(4).Text, 1)
         !ImpMN = frmTHPrGrd.uorstCOHPrDocCCo!ImpCCo_MN
         !ImpME = frmTHPrGrd.uorstCOHPrDocCCo!ImpCCo_ME
      Else
'         !ImpMN = frmTHPrGrd.uorstCOHPrDocCta!ImpCta * IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, 1, txtDato(4).Text)
'         !ImpME = frmTHPrGrd.uorstCOHPrDocCta!ImpCta / IIf(cboTpoMon.ListIndex = TPOMON_NAC_IND, txtDato(4).Text, 1)
         !ImpMN = frmTHPrGrd.uorstCOHPrDocCta!ImpCta_MN
         !ImpME = frmTHPrGrd.uorstCOHPrDocCta!ImpCta_ME
      End If
      !TpoGnr = TPOGNR_HPR
      !UsrCre = gsAbvUsr
      !FyHCre = Now
   End With
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
         !IndCta_Bru = pnValor
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
      frmTHPrGrd.uorstCOCta.MoveFirst
      frmTHPrGrd.uorstCOCta.Find "CodCta='" & txtDato(tnIndice + DIFERENCIAMASCUENTA).Text & "'"
      txtDato(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTHPrGrd.uorstCOCta!IndCCo = INDCCO_ACT)
      lblDatoDeta(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTHPrGrd.uorstCOCta!IndCCo = INDCCO_ACT)
      cmdDatoAyud(tnIndice + DIFERENCIAMASCCOSTO).Enabled = (frmTHPrGrd.uorstCOCta!IndCCo = INDCCO_ACT)
   End If
End Sub

Private Sub ppDatosWhere()             'Cambiar.
   With frmTHPrGrd
      .uorstCOCpbCab.Requery
   
      .usConnStrgWher_COCpbDet = "WHERE COCpbDet.MesPvs='" & gsMesAct & "' AND COCpbDet.CodDro='" & txtDato(0).Text & "' AND COCpbDet.NroCpb='" & txtDato(1).Text & "' "
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
                      If !TpoCnc = Trim(Str(dnContador + 1)) Then
                         dnTotalCuentaMN = dnTotalCuentaMN + !ImpCta_MN
                         dnTotalCuentaME = dnTotalCuentaME + !ImpCta_ME
                         With frmTHPrGrd.uorstCOCta
                            .MoveFirst
                            .Find "CodCta='" & frmTHPrGrd.uorstCOHPrDocCta!CodCta & "'"
                            If Not .EOF Then
                               dnIndCCo = frmTHPrGrd.uorstCOCta!IndCCo
                            End If
                         End With
                      End If
                      If dnIndCCo = INDCCO_ACT Then
                         With frmTHPrGrd.uorstCOHPrDocCCo
                            If .State = adStateOpen Then .Close
                            frmTHPrGrd.usConnStrgWher_COHPrDocCCo = "WHERE COHPrDocCCo.CodAux='" & frmTHPrGrd.uorstMain!CodAux & "' And COHPrDocCCo.SerDoc='" & frmTHPrGrd.uorstMain!SerDoc & "' And COHPrDocCCo.NroDoc='" & frmTHPrGrd.uorstMain!NroDoc & "' And COHPrDocCCo.TpoCnc='" & Trim(Str(dnContador + 1)) & "' And COHPrDocCCo.CodCta='" & frmTHPrGrd.uorstCOHPrDocCta!CodCta & "' "
                            .Source = frmTHPrGrd.usConnStrgSele_COHPrDocCCo & frmTHPrGrd.usConnStrgWher_COHPrDocCCo & frmTHPrGrd.usConnStrgOrde_COHPrDocCCo
                            .Open
                            If .RecordCount = 0 Then
                               MsgBox "Cuenta " & frmTHPrGrd.uorstCOHPrDocCta!CodCta & " requiere C.Costo", vbInformation
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
          With frmTHPrGrd.uorstCOCta
             .MoveFirst
             .Find "CodCta='" & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & "'"
             If Not .EOF Then
                dnIndCCo = frmTHPrGrd.uorstCOCta!IndCCo
             End If
          End With
          If dnIndCCo = INDCCO_ACT And txtDato(dnContador + 1 + DIFERENCIAMASCUENTA + CUENTASCONCCOSTO).Text = "" Then
             MsgBox "Cuenta " & txtDato(dnContador + 1 + DIFERENCIAMASCUENTA).Text & " requiere C.Costo", vbInformation
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



