VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOPar 
   Caption         =   "[Entidad]"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   1065
      ScaleHeight     =   690
      ScaleWidth      =   2955
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5100
      Width           =   2955
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
         Left            =   2220
         Picture         =   "frmOPar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   700
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
         Left            =   1485
         Picture         =   "frmOPar.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   700
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
         Left            =   780
         Picture         =   "frmOPar.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   700
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
         Left            =   60
         Picture         =   "frmOPar.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   60
         Width           =   720
      End
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   4850
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   8546
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nivel de Codigos"
      TabPicture(0)   =   "frmOPar.frx":0498
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTexto(20)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmNivel(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmNivel(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkProDestino"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboCodPlCta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Monedas"
      TabPicture(1)   =   "frmOPar.frx":04B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDato(1)"
      Tab(1).Control(1)=   "txtDato(0)"
      Tab(1).Control(2)=   "fraMoneda"
      Tab(1).Control(3)=   "cboMonFnc"
      Tab(1).Control(4)=   "lblTexto(2)"
      Tab(1).Control(5)=   "lblTexto(1)"
      Tab(1).Control(6)=   "lblTexto(0)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Impuestos"
      TabPicture(2)   =   "frmOPar.frx":04D0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtDato(18)"
      Tab(2).Control(1)=   "txtDato(17)"
      Tab(2).Control(2)=   "txtDato(5)"
      Tab(2).Control(3)=   "txtDato(4)"
      Tab(2).Control(4)=   "txtDato(3)"
      Tab(2).Control(5)=   "txtDato(2)"
      Tab(2).Control(6)=   "lblTexto(19)"
      Tab(2).Control(7)=   "lblTexto(18)"
      Tab(2).Control(8)=   "lblTexto(6)"
      Tab(2).Control(9)=   "lblTexto(5)"
      Tab(2).Control(10)=   "lblTexto(4)"
      Tab(2).Control(11)=   "lblTexto(3)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Retención/Percepción"
      TabPicture(3)   =   "frmOPar.frx":04EC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblTexto(12)"
      Tab(3).Control(1)=   "lblTexto(10)"
      Tab(3).Control(2)=   "lblTexto(9)"
      Tab(3).Control(3)=   "lblTexto(8)"
      Tab(3).Control(4)=   "lblTexto(14)"
      Tab(3).Control(5)=   "lblTexto(13)"
      Tab(3).Control(6)=   "lblTexto(7)"
      Tab(3).Control(7)=   "lblTexto(11)"
      Tab(3).Control(8)=   "txtDato(9)"
      Tab(3).Control(9)=   "txtDato(8)"
      Tab(3).Control(10)=   "txtDato(7)"
      Tab(3).Control(11)=   "txtDato(6)"
      Tab(3).Control(12)=   "txtDato(11)"
      Tab(3).Control(13)=   "txtDato(10)"
      Tab(3).Control(14)=   "cmdDatoAyud(6)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "cmdDatoAyud(7)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "cmdDatoAyud(9)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "cmdDatoAyud(10)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "chkAgenteRtc"
      Tab(3).Control(19)=   "chkAgentePcp"
      Tab(3).ControlCount=   20
      TabCaption(4)   =   "Varios"
      TabPicture(4)   =   "frmOPar.frx":0508
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkpedido"
      Tab(4).Control(1)=   "chkEjercicioFran"
      Tab(4).Control(2)=   "fraRetencion"
      Tab(4).Control(3)=   "txtDato(12)"
      Tab(4).Control(4)=   "lblTexto(15)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Caja Bancos"
      TabPicture(5)   =   "frmOPar.frx":0524
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtDato(16)"
      Tab(5).Control(1)=   "cmdDatoAyud(16)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "txtDato(15)"
      Tab(5).Control(3)=   "cmdDatoAyud(15)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "lblTexto(17)"
      Tab(5).Control(5)=   "lblDatoDeta(16)"
      Tab(5).Control(6)=   "lblTexto(16)"
      Tab(5).Control(7)=   "lblDatoDeta(15)"
      Tab(5).ControlCount=   8
      Begin VB.ComboBox cboCodPlCta 
         Height          =   315
         Left            =   60
         TabIndex        =   77
         Top             =   4380
         Width           =   4605
      End
      Begin VB.CheckBox chkProDestino 
         Caption         =   "Destinos en Comprobante"
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   315
         TabIndex        =   76
         Top             =   3705
         Width           =   2760
      End
      Begin VB.CheckBox chkpedido 
         Caption         =   "Filtrar Pedido de Compra"
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   -72480
         TabIndex        =   75
         Top             =   1080
         Width           =   2100
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   18
         Left            =   -71265
         TabIndex        =   15
         Top             =   2025
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   17
         Left            =   -71265
         TabIndex        =   13
         Top             =   1680
         Width           =   555
      End
      Begin VB.CheckBox chkEjercicioFran 
         Caption         =   "Ejercicio Francés"
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   -74880
         TabIndex        =   41
         Top             =   1035
         Width           =   1980
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   16
         Left            =   -74835
         TabIndex        =   53
         Top             =   2220
         Width           =   520
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   16
         Left            =   -70635
         Picture         =   "frmOPar.frx":0540
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2220
         Width           =   300
      End
      Begin VB.TextBox txtDato 
         ForeColor       =   &H80000012&
         Height          =   280
         Index           =   15
         Left            =   -74880
         TabIndex        =   50
         Top             =   1515
         Width           =   520
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   280
         Index           =   15
         Left            =   -70680
         Picture         =   "frmOPar.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1515
         Width           =   300
      End
      Begin VB.Frame fraRetencion 
         Caption         =   "Glosa Documentos Venta "
         ForeColor       =   &H80000002&
         Height          =   2070
         Left            =   -74880
         TabIndex        =   44
         Top             =   1905
         Width           =   4545
         Begin VB.CheckBox chkGlosaRtc 
            Alignment       =   1  'Right Justify
            Caption         =   "Sin Retención"
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   1
            Left            =   110
            TabIndex        =   47
            Top             =   1140
            Width           =   1515
         End
         Begin VB.CheckBox chkGlosaRtc 
            Alignment       =   1  'Right Justify
            Caption         =   "Con Retención"
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   0
            Left            =   110
            TabIndex        =   45
            Top             =   270
            Width           =   1515
         End
         Begin VB.TextBox txtDato 
            Height          =   570
            Index           =   14
            Left            =   110
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Top             =   1380
            Width           =   4300
         End
         Begin VB.TextBox txtDato 
            Height          =   570
            Index           =   13
            Left            =   110
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   525
            Width           =   4300
         End
      End
      Begin VB.Frame frmNivel 
         Caption         =   " Centro de Costo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2505
         Index           =   1
         Left            =   2595
         TabIndex        =   69
         Top             =   1065
         Width           =   1800
         Begin VB.CheckBox chkCenCos 
            Caption         =   "2 Dígitos"
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   72
            Top             =   345
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkCenCos 
            Caption         =   "5 Dígitos"
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   71
            Top             =   945
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkCenCos 
            Caption         =   "3 Dígitos"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   70
            Top             =   645
            Width           =   1215
         End
      End
      Begin VB.Frame frmNivel 
         Caption         =   " Cuenta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2505
         Index           =   0
         Left            =   315
         TabIndex        =   61
         Top             =   1080
         Width           =   1800
         Begin VB.CheckBox chkNiveles 
            Caption         =   "3 Dígitos"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   68
            Top             =   645
            Width           =   1215
         End
         Begin VB.CheckBox chkNiveles 
            Caption         =   "4 Dígitos"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   67
            Top             =   945
            Width           =   1215
         End
         Begin VB.CheckBox chkNiveles 
            Caption         =   "5 Dígitos"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   66
            Top             =   1245
            Width           =   1215
         End
         Begin VB.CheckBox chkNiveles 
            Caption         =   "6 Dígitos"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   65
            Top             =   1545
            Width           =   1215
         End
         Begin VB.CheckBox chkNiveles 
            Caption         =   "7 Dígitos"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   64
            Top             =   1845
            Width           =   1215
         End
         Begin VB.CheckBox chkNiveles 
            Caption         =   "8 Dígitos"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   5
            Left            =   165
            TabIndex        =   63
            Top             =   2145
            Width           =   1215
         End
         Begin VB.CheckBox chkNivel2 
            Caption         =   "2 Dígitos"
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   165
            TabIndex        =   62
            Top             =   345
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkAgentePcp 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -71880
         TabIndex        =   31
         Top             =   2535
         Width           =   255
      End
      Begin VB.CheckBox chkAgenteRtc 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -71880
         TabIndex        =   27
         Top             =   1110
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   10
         Left            =   -70830
         Picture         =   "frmOPar.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3180
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   9
         Left            =   -71490
         Picture         =   "frmOPar.frx":0A3E
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2805
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   7
         Left            =   -70830
         Picture         =   "frmOPar.frx":0BE8
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1725
         Width           =   255
      End
      Begin VB.CommandButton cmdDatoAyud 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   300
         Index           =   6
         Left            =   -71490
         Picture         =   "frmOPar.frx":0D92
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1380
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   12
         Left            =   -73740
         TabIndex        =   43
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   10
         Left            =   -71820
         TabIndex        =   33
         Top             =   3180
         Width           =   975
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   11
         Left            =   -71820
         TabIndex        =   34
         Top             =   3540
         Width           =   615
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   6
         Left            =   -71820
         TabIndex        =   28
         Top             =   1380
         Width           =   315
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         Index           =   7
         Left            =   -71820
         TabIndex        =   29
         Top             =   1725
         Width           =   975
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   8
         Left            =   -71820
         TabIndex        =   30
         Top             =   2085
         Width           =   615
      End
      Begin VB.TextBox txtDato 
         Height          =   300
         HideSelection   =   0   'False
         Index           =   9
         Left            =   -71820
         TabIndex        =   32
         Top             =   2805
         Width           =   315
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   5
         Left            =   -71265
         TabIndex        =   21
         Top             =   3075
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   -71265
         TabIndex        =   19
         Top             =   2730
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   -71265
         TabIndex        =   17
         Top             =   2385
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   -71265
         TabIndex        =   11
         Top             =   1335
         Width           =   555
      End
      Begin VB.TextBox txtDato 
         Height          =   375
         Index           =   1
         Left            =   -72600
         TabIndex        =   9
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox txtDato 
         Height          =   375
         Index           =   0
         Left            =   -72600
         TabIndex        =   8
         Top             =   3000
         Width           =   495
      End
      Begin VB.Frame fraMoneda 
         Caption         =   "Monedas de Trabajo"
         ForeColor       =   &H80000002&
         Height          =   795
         Left            =   -74520
         TabIndex        =   23
         Top             =   1920
         Width           =   3855
         Begin VB.OptionButton optMon2 
            Caption         =   "&2 Monedas"
            ForeColor       =   &H80000001&
            Height          =   195
            Left            =   2220
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optMon1 
            Caption         =   "&1 Moneda"
            ForeColor       =   &H80000001&
            Height          =   195
            Left            =   600
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.ComboBox cboMonFnc 
         Height          =   315
         Left            =   -72720
         TabIndex        =   5
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label lblTexto 
         Caption         =   "Código.Plan Ctas. Utilizado x DeudorTriburio :"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   20
         Left            =   360
         TabIndex        =   78
         Top             =   4080
         Width           =   3915
      End
      Begin VB.Label lblTexto 
         Caption         =   "Impuesto General a las Ventas 2 :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   19
         Left            =   -74445
         TabIndex        =   14
         Top             =   2070
         Width           =   2990
      End
      Begin VB.Label lblTexto 
         Caption         =   "Impuesto General a las Ventas 1 :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   18
         Left            =   -74445
         TabIndex        =   12
         Top             =   1725
         Width           =   2990
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Diario de Egreso :"
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
         Left            =   -74835
         TabIndex        =   52
         Top             =   1965
         Width           =   1275
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
         Index           =   16
         Left            =   -74310
         TabIndex        =   54
         Top             =   2220
         Width           =   3675
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Diario de Ingreso :"
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
         Left            =   -74880
         TabIndex        =   49
         Top             =   1260
         Width           =   1305
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
         Index           =   15
         Left            =   -74355
         TabIndex        =   51
         Top             =   1515
         Width           =   3675
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Agente de Percepción"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   11
         Left            =   -74520
         TabIndex        =   60
         Top             =   2535
         Width           =   1590
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Agente de Retención"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   7
         Left            =   -74520
         TabIndex        =   59
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Importe UIT"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   15
         Left            =   -74880
         TabIndex        =   42
         Top             =   1485
         Width           =   840
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Percepción"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   13
         Left            =   -74520
         TabIndex        =   40
         Top             =   3240
         Width           =   1365
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Percepción"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   14
         Left            =   -74520
         TabIndex        =   39
         Top             =   3600
         Width           =   1620
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento Retención"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   8
         Left            =   -74520
         TabIndex        =   38
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Retención"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   9
         Left            =   -74520
         TabIndex        =   37
         Top             =   1785
         Width           =   1290
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Retención"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   10
         Left            =   -74520
         TabIndex        =   36
         Top             =   2145
         Width           =   1545
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento Percepción"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   12
         Left            =   -74520
         TabIndex        =   35
         Top             =   2865
         Width           =   2040
      End
      Begin VB.Label lblTexto 
         Caption         =   "Impuesto Extraordinario de Solidaridad"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   6
         Left            =   -74445
         TabIndex        =   20
         Top             =   3135
         Width           =   2990
      End
      Begin VB.Label lblTexto 
         Caption         =   "Impuesto a la Renta de 4ª Categoría"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   5
         Left            =   -74445
         TabIndex        =   18
         Top             =   2790
         Width           =   2990
      End
      Begin VB.Label lblTexto 
         Caption         =   "Impuesto Selectivo al Consumo"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   -74445
         TabIndex        =   16
         Top             =   2445
         Width           =   2990
      End
      Begin VB.Label lblTexto 
         Caption         =   "Impuesto General a las Ventas :"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   -74445
         TabIndex        =   10
         Top             =   1380
         Width           =   2990
      End
      Begin VB.Label lblTexto 
         Caption         =   "Moneda Extranjera"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   -74100
         TabIndex        =   26
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblTexto 
         Caption         =   "Moneda Nacional"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   -74100
         TabIndex        =   25
         Top             =   3060
         Width           =   1335
      End
      Begin VB.Label lblTexto 
         Caption         =   "Moneda Funcional de la Empresa"
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   0
         Left            =   -74340
         TabIndex        =   24
         Top             =   1320
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmOPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pocnnMain As ADODB.Connection
Private porstCoCfg As ADODB.Recordset
Private porstTGCfg As ADODB.Recordset
Private porstCodro As ADODB.Recordset
Private pbNuevo As Boolean

Private Sub chkGlosaRtc_Click(Index As Integer)
  If chkGlosaRtc(Index).Value = vbChecked Then
    chkGlosaRtc(Choose(Index + 1, 1, 0)).Value = vbUnchecked
  End If
End Sub

Private Sub cmdDatoAyud_Click(Index As Integer)
   Select Case Index                   'Cambiar. Añadir índices.
   Case 6, 7, 9, 10, 15, 16
      txtDato(Index).SetFocus
   End Select
   ppAyuBus Index
End Sub

Private Sub Form_Load()
   Dim dnContador As Integer

   Me.KeyPreview = True
   
 '[Recordsets                          'Cambiar.
   Set pocnnMain = New ADODB.Connection
   Set porstCoCfg = New ADODB.Recordset
   Set porstTGCfg = New ADODB.Recordset
   Set porstCodro = New ADODB.Recordset
   With pocnnMain
      .CursorLocation = adUseClient
      .ConnectionString = CONNSTRG & gsNomBDS
      .Open
   End With
   With porstCoCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT codemp, pdoano, CodCta_Nv3, CodCta_Nv4, CodCta_Nv5, CodCta_Nv6, CodCta_Nv7, CodCta_Nv8, "
      .Source = .Source & "TpoMon_Fnc, TpoMon_Sgn_MN, TpoMon_Sgn_ME, IndMNE, "
      .Source = .Source & "CodTDc_Pcp, CodCta_Pcp, IndPcp, CodTDc_Rtc, CodCta_Rtc, IndRtc, "
      .Source = .Source & "CodCCo_Nv3, CodCCo_Nv5, TpoGlo_Rtc, GloDocr_Rtc, GloDocn_Rtc, "
      .Source = .Source & "coddro_ing, coddro_egr, ejerfran, indpedido, prodestino "
      .Source = .Source & "FROM CoCfg "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "CoCfg"
   End With
   With porstTGCfg
      .ActiveConnection = pocnnMain
      .Source = "SELECT codemp, pdoano, PctIGV, PctIGV1, PctIGV2, PctISC, PctIR4, PctIES, PctRtc, PctPcp, ImpUIT "
      'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
      .Source = .Source & ",codplacta "
      'fin 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
      .Source = .Source & "FROM TGCfg "
      .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
      .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
'     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
      .Properties("Unique Table").Value = "TGCfg"
   End With
   With porstCodro
     .ActiveConnection = pocnnMain
     .Source = "SELECT CodDro, " & Choose(gsIdioma, "DetDro", "DetDrox") & " AS DetDro "
     .Source = .Source & "FROM CODro "
     .Source = .Source & "WHERE codemp='" & gsCodEmp & "' "
     .Source = .Source & "AND pdoano='" & gsAnoAct & "'"
''     .CursorLocation = adUseClient   'Es el Default.
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
   End With
 ']

 '[Datos                               'Cambiar.
   With cboMonFnc
    .AddItem TPOMON_NAC_TXT_0, TPOMON_NAC_IND
    .AddItem TPOMON_EXT_TXT_0, TPOMON_EXT_IND
   End With
 ']
  Dim nContador As Integer
 'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
    With cboCodPlCta
    For nContador = 1 To UBound(aCodPlCta, 1)
        .AddItem aDetPlCta(nContador), aCodPlCta(nContador)
    Next nContador
    'cboCodPlCta.s
   End With
 'fin 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   ppDatosDesconectados 1
   ppHabilitacion False
  
  '[ Cargo los mensajes de botones
  Dim nElemento As Integer
  ReDim aLabel(18, 2)
  For nElemento = 0 To UBound(aLabel, 1) - 1
    aLabel(nElemento, 0) = Choose(nElemento + 1, "Moneda Funcional de la Empresa :", "Moneda Nacional :", "Moneda Extranjera :", "Impuesto General a las Ventas :", "Impuesto Selectivo al Consumo :", "Impuesto a la Renta de 4ta Categoría :", "Impuesto extraordinario de Solidaridad :", "Agente de Retención :", "Tipo Documento Retención :", "Cuenta Retención :", "Porcentaje Retención :", "Agente de Percepción :", "Tipo Documento Percepción :", "Cuenta Percepción :", "Porcentaje Percepción :", "Importe UIT :", "Diario de Ingreso :", "Diario de Egreso :")
    aLabel(nElemento, 1) = Choose(nElemento + 1, "Functional Currency of the Company :", "National Currency :", "Foreign Currency :", "General Sales Tax :", "Selective Consumption Tax :", "Income Tax of 4th Class :", "Extraordinary Tax of Solidarity :", "Agent of Withholding :", "Type Document Withholding :", "Account Withholding :", "Percentage Withholding :", "Agent of Perception :", "Type Document Perception :", "Account Perception :", "Percentage Perception :", "Amount UTT :", "Entry Diary :", "Discharge Diary  :")
  Next nElemento
  For nElemento = 0 To 4
    If gsIdioma = NvlUsr_Sup Then
      sstMain.TabCaption(nElemento) = Choose(nElemento + 1, "Nivel de Codigos", "Monedas", "Impuestos", "Retención/Percepción", "Varios")
    Else
      sstMain.TabCaption(nElemento) = Choose(nElemento + 1, "Level of Codes", "Currencies", "Taxes", "Withholding / Perception", "Several")
    End If
  Next nElemento
  frmNivel(0).Caption = Choose(gsIdioma, "Cuenta", "Account")
  chkNivel2.Caption = Choose(gsIdioma, "2 Dígitos", "2 Digits")
  chkNiveles(0).Caption = Choose(gsIdioma, "3 Dígitos", "3 Digits")
  chkNiveles(1).Caption = Choose(gsIdioma, "4 Dígitos", "4 Digits")
  chkNiveles(2).Caption = Choose(gsIdioma, "5 Dígitos", "5 Digits")
  chkNiveles(3).Caption = Choose(gsIdioma, "6 Dígitos", "6 Digits")
  chkNiveles(4).Caption = Choose(gsIdioma, "7 Dígitos", "7 Digits")
  chkNiveles(5).Caption = Choose(gsIdioma, "8 Dígitos", "8 Digits")
  frmNivel(1).Caption = Choose(gsIdioma, "Centro de Costo", "Cost Center")
  chkCenCos(0).Caption = Choose(gsIdioma, "2 Dígitos", "2 Digits")
  chkCenCos(1).Caption = Choose(gsIdioma, "3 Dígitos", "3 Digits")
  chkCenCos(2).Caption = Choose(gsIdioma, "5 Dígitos", "5 Digits")
  chkProDestino.Caption = Choose(gsIdioma, "Destinos en comprobante", "Proof destinations")
  fraMoneda.Caption = Choose(gsIdioma, "Monedas de Trabajo", "Work Currencies")
  optMon1.Caption = Choose(gsIdioma, "&1 Moneda", "&1 Currency")
  optMon2.Caption = Choose(gsIdioma, "&2 Monedas", "&2 Currencies")
  chkEjercicioFran.Caption = Choose(gsIdioma, "Ejercicio Francés", "French Fiscal")
  chkpedido.Caption = Choose(gsIdioma, "Filtrar Pedidos de Compra", "Filter Purchase Order")
  fraRetencion.Caption = Choose(gsIdioma, "Glosa Documento de Venta", "Gloss Document of Sale")
  chkGlosaRtc(0).Caption = Choose(gsIdioma, "&Con Retención", "&With Withholding")
  chkGlosaRtc(1).Caption = Choose(gsIdioma, "&Sin Retención", "&Not Withholding")
  CaptionBotones Me, False, False, False, False, False, False, False, False, False, True, True, True, True, aLabel
 ']
 
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'   Call gpTeclasData2(KeyAscii)
'End Sub

Private Sub Form_Unload(Cancel As Integer)   'Cambiar Recordsets.
   porstCoCfg.Close
   porstCodro.Close
   pocnnMain.Close
   Set porstCoCfg = Nothing
   Set porstCodro = Nothing
   Set pocnnMain = Nothing
End Sub

Private Sub cmdCorregir_Click()
   cmdCorregir.Enabled = False
   cmdGrabar.Enabled = True
   cmdDeshacer.Enabled = True
   ppHabilitacion (True)
 
 '[Dato con el foco al corregir.       'Cambiar.
   If sstMain.Tab = 2 Then
      cboMonFnc.SetFocus
   Else
      chkNiveles(0).SetFocus
   End If
 ']
End Sub

Private Sub cmdGrabar_Click()
  On Error GoTo Err

  pocnnMain.BeginTrans                'INICIA TRANSACCION.
  
  ppDatosDesconectados 0
  
  porstCoCfg.Update
  porstTGCfg.Update
  
  pocnnMain.CommitTrans               'CONFIRMA TRANSACCION.
  
  '[Propio del formulario.
  With porstCoCfg
    
    gsNivCta = "2" & IIf(!CodCta_Nv3, "3", "") & IIf(!CodCta_Nv4, "4", "") & IIf(!CodCta_Nv5, "5", "") & IIf(!CodCta_Nv6, "6", "") & IIf(!CodCta_Nv7, "7", "") & IIf(!CodCta_Nv8, "8", "")
    gsNivCCo = "2" & IIf(!CodCCo_Nv3, "3", "") & IIf(!CodCCo_Nv5, "5", "")
    gnProDestino = IIf(IsNull(!prodestino), 0, !prodestino)
    gnIndMNE = IIf(IsNull(!IndMNE), 0, !IndMNE)
    gsTpoMon_Fnc = IIf(IsNull(!TpoMon_Fnc), "", !TpoMon_Fnc)
    gsTpoMon_Sgn_MN = IIf(IsNull(!TpoMon_Sgn_MN), "", !TpoMon_Sgn_MN)
    gsTpoMon_Sgn_ME = IIf(IsNull(!TpoMon_Sgn_ME), "", !TpoMon_Sgn_ME)
    gsCodTDc_Pcp = IIf(IsNull(!COdTDC_Pcp), "", !COdTDC_Pcp)
    gsCodTDc_Rtc = IIf(IsNull(!CodTDc_Rtc), "", !CodTDc_Rtc)
    gsCodCta_Pcp = IIf(IsNull(!COdCta_Pcp), "", !COdCta_Pcp)
    gsCodCta_Pcp = IIf(IsNull(!CodCta_Rtc), "", !CodCta_Rtc)
    gsIndRtc = IIf(IsNull(!IndRtc), "N", !IndRtc)
    gsIndPcp = IIf(IsNull(!IndPcp), "N", !IndPcp)
    gsTpoGlo_Rtc = IIf(IsNull(!TpoGlo_Rtc), TPOGRU1_IND, !TpoGlo_Rtc)
    gsGloDoc_Rtc(0) = ""
    gsGloDoc_Rtc(1) = IIf(IsNull(!GloDocr_Rtc), "", !GloDocr_Rtc)
    gsGloDoc_Rtc(2) = IIf(IsNull(!GloDocn_Rtc), "", !GloDocn_Rtc)
    
    gsCodDro_Ing = IIf(IsNull(!coddro_ing), "", !coddro_ing)
    gsCodDro_Egr = IIf(IsNull(!coddro_egr), "", !coddro_egr)
    
  End With
  With porstTGCfg
 'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
gnCodPlaCata = !CodPlaCta
'fin 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
    gnPctIGV = CDec(!PctIGV)
    gnPctIGV1 = CDec(!PctIGV1)
    gnPctIGV2 = CDec(!PctIGV2)
    gnPctISC = CDec(!PctISC)
    gnPctIR4 = CDec(!PctIR4)
    gnPctIES = CDec(!PctIES)
    gnPctRtc = CDec(!PctRtc)
    gnPctPcp = CDec(!PctPcp)
    gnImpUIT = CDec(!ImpUIT)
    
  End With
 ']
   
  cmdCorregir.Enabled = True
  cmdGrabar.Enabled = False
  cmdDeshacer.Enabled = False
  ppHabilitacion False
      
  Exit Sub
Err:
   gpErrores
End Sub

Private Sub cmdDeshacer_Click()
   On Error GoTo Err

   ppDatosDesconectados 1
   cmdCorregir.Enabled = True
   cmdGrabar.Enabled = False
   cmdDeshacer.Enabled = False
   ppHabilitacion False

   Exit Sub
Err:
   gpErrores
End Sub

Private Sub cmdSalir_Click()
   Unload Me
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
End Sub

Private Sub txtDato_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case Index
   Case 6, 7, 9, 10, 13, 14, 15, 16
    If KeyCode = vbKeyF2 Then
      ppAyuBus Index
    End If
  End Select
End Sub

Private Sub ppDatosDesconectados(tnFase As Byte) 'Cambiar.
'tnFase           Fase del procedimiento (0:Grabar 1:Corregir).
   
   On Error GoTo Err

   If tnFase = 0 Then
     'Datos.
      porstCoCfg!TpoMon_Fnc = IIf(cboMonFnc.ListIndex = TPOMON_NAC_IND, TPOMON_NAC, TPOMON_EXT)
      'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
      'si cbo es =08 codico consecutivo entonces poner "99"
      porstTGCfg!CodPlaCta = IIf(cboCodPlCta.ListIndex = "08", "99", gfCeros(Str(cboCodPlCta.ListIndex), 2, 0, "0"))
      'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
      porstCoCfg!CodCta_Nv3 = IIf(chkNiveles(0).Value = vbChecked, 1, 0)
      porstCoCfg!CodCta_Nv4 = IIf(chkNiveles(1).Value = vbChecked, 1, 0)
      porstCoCfg!CodCta_Nv5 = IIf(chkNiveles(2).Value = vbChecked, 1, 0)
      porstCoCfg!CodCta_Nv6 = IIf(chkNiveles(3).Value = vbChecked, 1, 0)
      porstCoCfg!CodCta_Nv7 = IIf(chkNiveles(4).Value = vbChecked, 1, 0)
      porstCoCfg!CodCta_Nv8 = IIf(chkNiveles(5).Value = vbChecked, 1, 0)
      porstCoCfg!prodestino = IIf(chkProDestino.Value = vbChecked, 1, 0)
'      porstMain!CodSoc = IIf(dcoSocio.BoundText = "", " ", dcoSocio.BoundText)
'      porstMain!FehOpe = dtpFecha.Value
      porstCoCfg!IndMNE = IIf(optMon1.Value, INDMNE_INA, INDMNE_ACT)
      porstCoCfg!TpoMon_Sgn_MN = txtDato(0).Text
      porstCoCfg!TpoMon_Sgn_ME = txtDato(1).Text
      porstCoCfg!COdTDC_Pcp = txtDato(9).Text
      porstCoCfg!CodTDc_Rtc = txtDato(6).Text
      porstCoCfg!COdCta_Pcp = txtDato(10).Text
      porstCoCfg!CodCta_Rtc = txtDato(7).Text
      porstCoCfg!IndRtc = IIf(chkAgenteRtc.Value = vbChecked, "S", "N")
      porstCoCfg!IndPcp = IIf(chkAgentePcp.Value = vbChecked, "S", "N")
      
      porstTGCfg!PctIGV = CDec(txtDato(2).Text)
      porstTGCfg!PctIGV1 = CDec(txtDato(17).Text)
      porstTGCfg!PctIGV2 = CDec(txtDato(18).Text)
      porstTGCfg!PctISC = CDec(txtDato(3).Text)
      porstTGCfg!PctIR4 = CDec(txtDato(4).Text)
      porstTGCfg!PctIES = CDec(txtDato(5).Text)
      porstTGCfg!PctRtc = CDec(txtDato(8).Text)
      porstTGCfg!PctPcp = CDec(txtDato(11).Text)
      
      porstCoCfg!CodCCo_Nv3 = IIf(chkCenCos(1).Value = vbChecked, 1, 0)
      porstCoCfg!CodCCo_Nv5 = IIf(chkCenCos(2).Value = vbChecked, 1, 0)
      
      ' Petana varios
      porstCoCfg!ejerfran = IIf(chkEjercicioFran.Value = vbChecked, 1, 0)
      porstCoCfg!indpedido = IIf(chkpedido.Value = vbChecked, 1, 0)
      
      gnIndPedido = IIf(chkpedido.Value = vbChecked, 1, 0)
      
      porstTGCfg!ImpUIT = CDec(txtDato(12).Text)
      porstCoCfg!TpoGlo_Rtc = IIf(chkGlosaRtc(0).Value = vbChecked, TPOGRU2_IND, IIf(chkGlosaRtc(1).Value = vbChecked, TPOGRU3_IND, TPOGRU1_IND))
      porstCoCfg!GloDocr_Rtc = IIf(Trim(txtDato(13).Text) = "", Null, txtDato(13).Text)
      porstCoCfg!GloDocn_Rtc = IIf(Trim(txtDato(14).Text) = "", Null, txtDato(14).Text)
      
      porstCoCfg!coddro_ing = IIf(Trim(txtDato(15).Text) = "", Null, txtDato(15).Text)
      porstCoCfg!coddro_egr = IIf(Trim(txtDato(16).Text) = "", Null, txtDato(16).Text)
   Else
      cboMonFnc.ListIndex = IIf(porstCoCfg!TpoMon_Fnc = TPOMON_NAC, TPOMON_NAC_IND, TPOMON_EXT_IND)
      'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
      cboCodPlCta.ListIndex = IIf(porstTGCfg!CodPlaCta = "99", "08", porstTGCfg!CodPlaCta)
      'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
      chkNiveles(0).Value = IIf(porstCoCfg!CodCta_Nv3 = 1, vbChecked, vbUnchecked)
      chkNiveles(1).Value = IIf(porstCoCfg!CodCta_Nv4 = 1, vbChecked, vbUnchecked)
      chkNiveles(2).Value = IIf(porstCoCfg!CodCta_Nv5 = 1, vbChecked, vbUnchecked)
      chkNiveles(3).Value = IIf(porstCoCfg!CodCta_Nv6 = 1, vbChecked, vbUnchecked)
      chkNiveles(4).Value = IIf(porstCoCfg!CodCta_Nv7 = 1, vbChecked, vbUnchecked)
      chkNiveles(5).Value = IIf(porstCoCfg!CodCta_Nv8 = 1, vbChecked, vbUnchecked)
      chkProDestino.Value = IIf(porstCoCfg!prodestino = 1, vbChecked, vbUnchecked)
'      dcoSocio.Item(0).BoundText = porstMain!CodSoc
'      dtpFecha.Value = porstMain!FehOpe
      optMon1.Value = IIf(porstCoCfg!IndMNE = INDMNE_INA, True, False)
      optMon2.Value = IIf(porstCoCfg!IndMNE = INDMNE_ACT, True, False)
'      optMoneda(1).Value = porstMain!CodMon
      txtDato(0).Text = IIf(IsNull(porstCoCfg!TpoMon_Sgn_MN), "", porstCoCfg!TpoMon_Sgn_MN)
      txtDato(1).Text = IIf(IsNull(porstCoCfg!TpoMon_Sgn_ME), "", porstCoCfg!TpoMon_Sgn_ME)
      txtDato(9).Text = IIf(IsNull(porstCoCfg!COdTDC_Pcp), "", porstCoCfg!COdTDC_Pcp)
      txtDato(6).Text = IIf(IsNull(porstCoCfg!CodTDc_Rtc), "", porstCoCfg!CodTDc_Rtc)
      txtDato(10).Text = IIf(IsNull(porstCoCfg!COdCta_Pcp), "", porstCoCfg!COdCta_Pcp)
      txtDato(7).Text = IIf(IsNull(porstCoCfg!CodCta_Rtc), "", porstCoCfg!CodCta_Rtc)
      chkAgenteRtc.Value = IIf(porstCoCfg!IndRtc = "S", vbChecked, vbUnchecked)
      chkAgentePcp.Value = IIf(porstCoCfg!IndPcp = "S", vbChecked, vbUnchecked)
      
      txtDato(2).Text = Format(porstTGCfg!PctIGV, FORMATO_NUM_4)
      txtDato(17).Text = Format(porstTGCfg!PctIGV1, FORMATO_NUM_4)
      txtDato(18).Text = Format(porstTGCfg!PctIGV2, FORMATO_NUM_4)
      txtDato(3).Text = Format(porstTGCfg!PctISC, FORMATO_NUM_4)
      txtDato(4).Text = Format(porstTGCfg!PctIR4, FORMATO_NUM_4)
      txtDato(5).Text = Format(porstTGCfg!PctIES, FORMATO_NUM_4)
      txtDato(8).Text = Format(porstTGCfg!PctRtc, FORMATO_NUM_4)
      txtDato(11).Text = Format(porstTGCfg!PctPcp, FORMATO_NUM_4)
      
      chkCenCos(1).Value = IIf(porstCoCfg!CodCCo_Nv3 = 1, vbChecked, vbUnchecked)
      chkCenCos(2).Value = IIf(porstCoCfg!CodCCo_Nv5 = 1, vbChecked, vbUnchecked)
      
      'Pestana varios
      chkEjercicioFran.Value = IIf(porstCoCfg!ejerfran = 1, vbChecked, vbUnchecked)
      chkpedido.Value = IIf(porstCoCfg!indpedido = 1, vbChecked, vbUnchecked)
      txtDato(12).Text = Format(porstTGCfg!ImpUIT, FORMATO_NUM_4)
      chkGlosaRtc(0).Value = IIf(porstCoCfg!TpoGlo_Rtc = TPOGRU2_IND, vbChecked, vbUnchecked)
      chkGlosaRtc(1).Value = IIf(porstCoCfg!TpoGlo_Rtc = TPOGRU3_IND, vbChecked, vbUnchecked)
      txtDato(13).Text = IIf(IsNull(porstCoCfg!GloDocr_Rtc), "", porstCoCfg!GloDocr_Rtc)
      txtDato(14).Text = IIf(IsNull(porstCoCfg!GloDocn_Rtc), "", porstCoCfg!GloDocn_Rtc)
      
      txtDato(15).Text = IIf(IsNull(porstCoCfg!coddro_ing), "", porstCoCfg!coddro_ing)
      txtDato(16).Text = IIf(IsNull(porstCoCfg!coddro_egr), "", porstCoCfg!coddro_egr)
   End If
   ppAyuDet 15
   ppAyuDet 16
   ' Inicializo apertura y cierre
   gsMesApe = IIf(chkEjercicioFran.Value = vbChecked, "09", "01")
   gsMesCie = IIf(chkEjercicioFran.Value = vbChecked, "08", "12")
   gnFrances = IIf(chkEjercicioFran.Value = vbChecked, vbChecked, vbUnchecked)
      
   Exit Sub
Err:
   gpErrores
   
   Resume
End Sub

'Public Sub upDatosPredeterminados()    'Cambiar.
'   chkEstado.Value = vbChecked
'   dcoSocio.BoundText = ""
'   dtpFecha.Value = Date
'   optMoneda(1).Value = True
'End Sub

Private Sub ppHabilitacion(tbHabilitar As Boolean) 'Cambiar.
   Dim dnContador As Integer

  'Datos.
   With txtDato
      For dnContador = 0 To .Count - 1
         .Item(dnContador).Enabled = tbHabilitar
      Next
   End With
   cboMonFnc.Enabled = tbHabilitar
   'ini 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
   cboCodPlCta.Enabled = tbHabilitar
   'fin 2014-05-29 Código del Plan de Cuentas utilizado por el deudor tributario
   For dnContador = 0 To 5
      chkNiveles(dnContador).Enabled = tbHabilitar
   Next
   chkProDestino.Enabled = tbHabilitar
   optMon1.Enabled = tbHabilitar
   optMon2.Enabled = tbHabilitar
   
   chkAgenteRtc.Enabled = tbHabilitar
   chkAgentePcp.Enabled = tbHabilitar

   chkEjercicioFran.Enabled = tbHabilitar
   chkpedido.Enabled = tbHabilitar
   
   chkGlosaRtc(0).Enabled = tbHabilitar
   chkGlosaRtc(1).Enabled = tbHabilitar

  'Ayudas.
   cmdDatoAyud.Item(6).Enabled = tbHabilitar
   cmdDatoAyud.Item(7).Enabled = tbHabilitar
   cmdDatoAyud.Item(9).Enabled = tbHabilitar
   cmdDatoAyud.Item(10).Enabled = tbHabilitar
   cmdDatoAyud.Item(15).Enabled = tbHabilitar
   cmdDatoAyud.Item(16).Enabled = tbHabilitar
   lblDatoDeta(15).Enabled = tbHabilitar
   lblDatoDeta(16).Enabled = tbHabilitar

   chkCenCos(1).Enabled = tbHabilitar
'   lblDatoDeta(0).Enabled = tbHabilitar
End Sub


Private Sub ppAyuBus(tnIndex As Integer)
   Select Case tnIndex
   Case 6, 9                           'Cambiar (añadir índices).
      modAyuBus.TDc_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
   Case 7, 10                           'Cambiar (añadir índices).
      modAyuBus.Cta_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
   Case 15, 16                           'Cambiar (añadir índices).
      modAyuBus.Dro_Cod "", txtDato(tnIndex).Text, 0, 0, Me.Top + txtDato(tnIndex).Top + txtDato(tnIndex).Height, Me.Left
      txtDato(tnIndex).Text = frmOAyuBus.uvDato1
      lblDatoDeta(tnIndex).Caption = " " & frmOAyuBus.uvDato2
   End Select
End Sub
Private Function ppAyuDet(tnIndex As Integer)
  Select Case tnIndex                 'Cambiar.
   Case 15, 16
    If txtDato(tnIndex).Text = "" Then
      lblDatoDeta(tnIndex).Caption = ""
      Exit Function
    End If
    With porstCodro
      .MoveFirst
      .Find "coddro='" & txtDato(tnIndex).Text & "'"
      If .EOF Then
        MsgBox TEXT_8006, vbExclamation
        ppAyuDet = True
      Else
        lblDatoDeta(tnIndex).Caption = " " & porstCodro!DetDro
      End If
    End With
  End Select
End Function

'[Código propio del formulario.

']

Private Sub txtDato_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
   Case 15, 16
    Cancel = ppAyuDet(Index)
    If Cancel Then Exit Sub
  End Select
End Sub
