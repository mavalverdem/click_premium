VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fPrmEmpresa 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6345
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "prmempresa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8910
   Begin TabDlg.SSTab tabRegister 
      Height          =   5115
      Left            =   75
      TabIndex        =   105
      Top             =   615
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9022
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabMaxWidth     =   3052
      BackColor       =   14737632
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Empresa"
      TabPicture(0)   =   "prmempresa.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDetalle(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(50)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkPrnBoletaDir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmCuadro(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmCuadro(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNivelCenCosto"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Planilla"
      TabPicture(1)   =   "prmempresa.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDetalle(1)"
      Tab(1).Control(1)=   "frmCuadro(2)"
      Tab(1).Control(2)=   "chkGratixDia"
      Tab(1).Control(3)=   "frmCuadro(3)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Representación"
      TabPicture(2)   =   "prmempresa.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmCuadro(4)"
      Tab(2).Control(1)=   "chkPrnBoleta"
      Tab(2).Control(2)=   "frmCuadro(5)"
      Tab(2).Control(3)=   "lblDetalle(2)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Imagen"
      TabPicture(3)   =   "prmempresa.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmCuadro(6)"
      Tab(3).Control(1)=   "tabFirma"
      Tab(3).Control(2)=   "chkPrnLiqRazon"
      Tab(3).Control(3)=   "chkPrnLiqLogo"
      Tab(3).Control(4)=   "shpCuadro(3)"
      Tab(3).Control(5)=   "lblDetalle(3)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Correo / Contrato"
      TabPicture(4)   =   "prmempresa.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frmCuadro(7)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "frmCuadro(8)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin Threed.SSFrame frmCuadro 
         Height          =   2445
         Index           =   7
         Left            =   -74760
         TabIndex        =   107
         Top             =   195
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   4313
         _StockProps     =   14
         Caption         =   " Configuración para el envío de Correo "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtPwdCorreoEnvio 
            Height          =   280
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   115
            Top             =   1950
            Width           =   3500
         End
         Begin VB.TextBox txtCtaCorreoEnvio 
            Height          =   280
            Left            =   240
            TabIndex        =   113
            Top             =   1350
            Width           =   3500
         End
         Begin VB.TextBox txtUsuarioEnvio 
            Height          =   280
            Left            =   4290
            TabIndex        =   111
            Top             =   720
            Width           =   3500
         End
         Begin VB.ComboBox cboServerEnvio 
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "prmempresa.frx":0098
            Left            =   240
            List            =   "prmempresa.frx":009A
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   720
            Width           =   2235
         End
         Begin VB.Label lblDato 
            AutoSize        =   -1  'True
            Caption         =   "Contraseña Cuenta de Correo de  Envío :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   31
            Left            =   240
            TabIndex        =   114
            Top             =   1710
            Width           =   2940
         End
         Begin VB.Label lblDato 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta de Correo de Envío :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   30
            Left            =   240
            TabIndex        =   112
            Top             =   1110
            Width           =   2040
         End
         Begin VB.Label lblDato 
            AutoSize        =   -1  'True
            Caption         =   "Usuario Envío :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   29
            Left            =   4290
            TabIndex        =   110
            Top             =   480
            Width           =   1110
         End
         Begin VB.Label lblDato 
            AutoSize        =   -1  'True
            Caption         =   "Servidor de Envío :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   108
            Top             =   480
            Width           =   1380
         End
      End
      Begin VB.TextBox txtNivelCenCosto 
         Height          =   280
         Left            =   7980
         TabIndex        =   2
         Top             =   150
         Width           =   500
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2190
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   690
         Width           =   8385
         _Version        =   65536
         _ExtentX        =   14790
         _ExtentY        =   3863
         _StockProps     =   14
         Caption         =   " Dirección "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtUbigeo 
            Height          =   280
            Index           =   0
            Left            =   150
            TabIndex        =   18
            Top             =   1740
            Width           =   975
         End
         Begin VB.TextBox txtNombreZona 
            Height          =   280
            Left            =   3810
            TabIndex        =   16
            Top             =   1155
            Width           =   3470
         End
         Begin VB.TextBox txtTipoZona 
            Height          =   280
            Left            =   150
            TabIndex        =   13
            Top             =   1155
            Width           =   500
         End
         Begin VB.TextBox txtNumero 
            Height          =   280
            Left            =   7020
            TabIndex        =   11
            Top             =   540
            Width           =   1155
         End
         Begin VB.TextBox txtNombreVia 
            Height          =   280
            Left            =   3810
            TabIndex        =   9
            Top             =   540
            Width           =   3120
         End
         Begin VB.TextBox txtTipoVia 
            Height          =   280
            Left            =   150
            TabIndex        =   6
            Top             =   540
            Width           =   500
         End
         Begin VB.TextBox txtTelefono 
            Height          =   280
            Index           =   0
            Left            =   5895
            TabIndex        =   21
            Top             =   1740
            Width           =   2325
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   280
            Index           =   0
            Left            =   705
            TabIndex        =   94
            Top             =   540
            Width           =   280
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   280
            Index           =   1
            Left            =   705
            TabIndex        =   95
            Top             =   1155
            Width           =   280
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdUbigeo 
            Height          =   280
            Index           =   0
            Left            =   1170
            TabIndex        =   96
            Top             =   1740
            Width           =   280
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblUbigeo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Height          =   195
            Index           =   0
            Left            =   1545
            TabIndex        =   19
            Top             =   1785
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Ubigeo :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   17
            Top             =   1485
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombre de Zona :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   3810
            TabIndex        =   15
            Top             =   900
            Width           =   1500
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Height          =   195
            Index           =   1
            Left            =   1065
            TabIndex        =   14
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Zona :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   12
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   7
            Top             =   585
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Número :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   7020
            TabIndex        =   10
            Top             =   285
            Width           =   765
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombre de Vía :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   3810
            TabIndex        =   8
            Top             =   285
            Width           =   1500
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo de Vía :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Teléfono :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   6
            Left            =   5895
            TabIndex        =   20
            Top             =   1485
            Width           =   1290
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1590
         Index           =   1
         Left            =   150
         TabIndex        =   22
         Top             =   2985
         Width           =   8385
         _Version        =   65536
         _ExtentX        =   14790
         _ExtentY        =   2805
         _StockProps     =   14
         Caption         =   " Datos Empresa "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtEmail 
            Height          =   280
            Left            =   180
            TabIndex        =   28
            Top             =   1170
            Width           =   4800
         End
         Begin VB.TextBox txtComercial 
            Height          =   280
            Left            =   180
            TabIndex        =   24
            Top             =   585
            Width           =   4800
         End
         Begin VB.TextBox txtPatronal 
            Height          =   280
            Left            =   5595
            TabIndex        =   26
            Top             =   585
            Width           =   2145
         End
         Begin VB.Label lblDato 
            Caption         =   "E- mail :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   27
            Top             =   915
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Giro Comercial :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   23
            Top             =   330
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Registro Patronal :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   8
            Left            =   5595
            TabIndex        =   25
            Top             =   330
            Width           =   1335
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2250
         Index           =   3
         Left            =   -74850
         TabIndex        =   41
         Top             =   2310
         Width           =   8400
         _Version        =   65536
         _ExtentX        =   14817
         _ExtentY        =   3969
         _StockProps     =   14
         Caption         =   " Personal de Planilla "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtTelefono 
            Height          =   280
            Index           =   1
            Left            =   4410
            TabIndex        =   49
            Top             =   1110
            Width           =   2250
         End
         Begin VB.TextBox txtCenCosto 
            Height          =   280
            Left            =   255
            TabIndex        =   51
            Top             =   1665
            Width           =   1200
         End
         Begin VB.TextBox txtPersonal 
            Height          =   280
            Index           =   2
            Left            =   255
            TabIndex        =   47
            Top             =   1110
            Width           =   3720
         End
         Begin VB.TextBox txtPersonal 
            Height          =   280
            Index           =   0
            Left            =   240
            MaxLength       =   25
            TabIndex        =   43
            Top             =   540
            Width           =   3720
         End
         Begin VB.TextBox txtPersonal 
            Height          =   280
            Index           =   1
            Left            =   4410
            TabIndex        =   45
            Top             =   540
            Width           =   3720
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   4
            Left            =   1530
            TabIndex        =   100
            Top             =   1665
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblDato 
            Caption         =   "Teléfono :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   16
            Left            =   4410
            TabIndex        =   48
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Centro Costo :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   17
            Left            =   255
            TabIndex        =   50
            Top             =   1425
            Width           =   1335
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Height          =   195
            Index           =   4
            Left            =   1905
            TabIndex        =   52
            Top             =   1710
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Paterno :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   42
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Materno :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   14
            Left            =   4410
            TabIndex        =   44
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombres :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   15
            Left            =   255
            TabIndex        =   46
            Top             =   855
            Width           =   1335
         End
      End
      Begin Threed.SSCheck chkGratixDia 
         Height          =   255
         Left            =   -68865
         TabIndex        =   30
         Top             =   405
         Width           =   2340
         _Version        =   65536
         _ExtentX        =   4128
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "BBSS descuenta ausencias "
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2805
         Index           =   6
         Left            =   -74715
         TabIndex        =   86
         Top             =   1575
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   4948
         _StockProps     =   14
         Caption         =   " Logo  "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin VB.Image imgLogo 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Left            =   225
            Stretch         =   -1  'True
            ToolTipText     =   "Haga doble click para logo empresa"
            Top             =   600
            Width           =   3375
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   2520
            Index           =   0
            Left            =   210
            Shape           =   4  'Rounded Rectangle
            Top             =   285
            Width           =   3420
         End
      End
      Begin TabDlg.SSTab tabFirma 
         Height          =   3195
         Left            =   -70410
         TabIndex        =   87
         Top             =   1350
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   5636
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabMaxWidth     =   3351
         BackColor       =   -2147483644
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Firma Representante"
         TabPicture(0)   =   "prmempresa.frx":009C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "shpCuadro(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "imgFirma(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Firma Gerente"
         TabPicture(1)   =   "prmempresa.frx":00B8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "imgFirma(1)"
         Tab(1).Control(1)=   "shpCuadro(2)"
         Tab(1).ControlCount=   2
         Begin VB.Image imgFirma 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Index           =   1
            Left            =   -74775
            Stretch         =   -1  'True
            ToolTipText     =   "Haga doble click para firma"
            Top             =   750
            Width           =   3360
         End
         Begin VB.Image imgFirma 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Index           =   0
            Left            =   225
            Stretch         =   -1  'True
            ToolTipText     =   "Haga doble click para firma"
            Top             =   750
            Width           =   3360
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   2520
            Index           =   1
            Left            =   195
            Shape           =   4  'Rounded Rectangle
            Top             =   480
            Width           =   3420
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   2520
            Index           =   2
            Left            =   -74805
            Shape           =   4  'Rounded Rectangle
            Top             =   480
            Width           =   3420
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2040
         Index           =   4
         Left            =   -74850
         TabIndex        =   55
         Top             =   495
         Width           =   8430
         _Version        =   65536
         _ExtentX        =   14870
         _ExtentY        =   3598
         _StockProps     =   14
         Caption         =   " Representante Legal "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtDocumento 
            Height          =   280
            Index           =   0
            Left            =   6975
            TabIndex        =   65
            Top             =   1110
            Width           =   1290
         End
         Begin VB.TextBox txtTipoDocu 
            Height          =   280
            Index           =   0
            Left            =   3630
            TabIndex        =   63
            Top             =   1110
            Width           =   500
         End
         Begin VB.TextBox txtRepresentante 
            Height          =   280
            Index           =   2
            Left            =   165
            TabIndex        =   61
            Top             =   1110
            Width           =   3180
         End
         Begin VB.TextBox txtRepresentante 
            Height          =   280
            Index           =   0
            Left            =   150
            MaxLength       =   25
            TabIndex        =   57
            Top             =   540
            Width           =   3180
         End
         Begin VB.TextBox txtRepresentante 
            Height          =   280
            Index           =   1
            Left            =   3630
            TabIndex        =   59
            Top             =   540
            Width           =   3180
         End
         Begin VB.TextBox txtCargo 
            Height          =   280
            Index           =   0
            Left            =   165
            TabIndex        =   67
            Top             =   1665
            Width           =   480
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   5
            Left            =   4185
            TabIndex        =   101
            Top             =   1110
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   6
            Left            =   735
            TabIndex        =   102
            Top             =   1665
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblDato 
            Caption         =   "Cargo :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   22
            Left            =   165
            TabIndex        =   66
            Top             =   1425
            Width           =   1340
         End
         Begin VB.Label lblDato 
            Caption         =   "Documento de Identificación :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   21
            Left            =   3630
            TabIndex        =   62
            Top             =   855
            Width           =   2070
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Height          =   195
            Index           =   5
            Left            =   4575
            TabIndex        =   64
            Top             =   1155
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Paterno :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   18
            Left            =   150
            TabIndex        =   56
            Top             =   285
            Width           =   1340
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Materno :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   19
            Left            =   3630
            TabIndex        =   58
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombres :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   20
            Left            =   165
            TabIndex        =   60
            Top             =   855
            Width           =   1340
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Height          =   195
            Index           =   6
            Left            =   1110
            TabIndex        =   68
            Top             =   1665
            Width           =   195
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1605
         Index           =   2
         Left            =   -74850
         TabIndex        =   31
         Top             =   600
         Width           =   8400
         _Version        =   65536
         _ExtentX        =   14817
         _ExtentY        =   2831
         _StockProps     =   14
         Caption         =   " Concepto Remuneración de Cálculo  "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtComision 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   4395
            TabIndex        =   39
            Top             =   540
            Width           =   980
         End
         Begin VB.TextBox txtRemunera 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   195
            TabIndex        =   33
            Top             =   540
            Width           =   980
         End
         Begin VB.TextBox txtAFamiliar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   280
            Left            =   195
            TabIndex        =   36
            Top             =   1155
            Width           =   980
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   3
            Left            =   1275
            TabIndex        =   98
            Top             =   1155
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   2
            Left            =   1275
            TabIndex        =   97
            Top             =   540
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   9
            Left            =   5475
            TabIndex        =   99
            Top             =   540
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblDato 
            Caption         =   "Comisión "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   12
            Left            =   4395
            TabIndex        =   38
            Top             =   285
            Width           =   1755
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Height          =   195
            Index           =   9
            Left            =   5835
            TabIndex        =   40
            Top             =   585
            Width           =   195
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Height          =   195
            Index           =   2
            Left            =   1635
            TabIndex        =   34
            Top             =   585
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Remuneración  Basica "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   10
            Left            =   195
            TabIndex        =   32
            Top             =   285
            Width           =   1755
         End
         Begin VB.Label lblDato 
            Caption         =   "Asignación  Familiar"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   11
            Left            =   195
            TabIndex        =   35
            Top             =   900
            Width           =   1755
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Height          =   195
            Index           =   3
            Left            =   1635
            TabIndex        =   37
            Top             =   1200
            Width           =   195
         End
      End
      Begin Threed.SSCheck chkPrnBoleta 
         Height          =   255
         Left            =   -69795
         TabIndex        =   54
         Top             =   210
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Impresión del  Representante en la Boleta "
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkPrnBoletaDir 
         Height          =   255
         Left            =   5535
         TabIndex        =   3
         Top             =   450
         Width           =   2940
         _Version        =   65536
         _ExtentX        =   5186
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Impresión de Dirección en la Boleta "
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   2040
         Index           =   5
         Left            =   -74850
         TabIndex        =   69
         Top             =   2640
         Width           =   8430
         _Version        =   65536
         _ExtentX        =   14870
         _ExtentY        =   3598
         _StockProps     =   14
         Caption         =   " Gerente Adjunto "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtCargo 
            Height          =   280
            Index           =   1
            Left            =   165
            TabIndex        =   81
            Top             =   1665
            Width           =   480
         End
         Begin VB.TextBox txtGerente 
            Height          =   280
            Index           =   1
            Left            =   3630
            TabIndex        =   73
            Top             =   540
            Width           =   3180
         End
         Begin VB.TextBox txtGerente 
            Height          =   280
            Index           =   0
            Left            =   150
            MaxLength       =   25
            TabIndex        =   71
            Top             =   540
            Width           =   3180
         End
         Begin VB.TextBox txtGerente 
            Height          =   280
            Index           =   2
            Left            =   165
            TabIndex        =   75
            Top             =   1110
            Width           =   3180
         End
         Begin VB.TextBox txtTipoDocu 
            Height          =   280
            Index           =   1
            Left            =   3630
            TabIndex        =   77
            Top             =   1110
            Width           =   500
         End
         Begin VB.TextBox txtDocumento 
            Height          =   280
            Index           =   1
            Left            =   6975
            TabIndex        =   79
            Top             =   1110
            Width           =   1290
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   7
            Left            =   4185
            TabIndex        =   103
            Top             =   1110
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   8
            Left            =   735
            TabIndex        =   104
            Top             =   1665
            Width           =   285
            _Version        =   65536
            _ExtentX        =   503
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Height          =   195
            Index           =   8
            Left            =   1110
            TabIndex        =   82
            Top             =   1665
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombres :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   25
            Left            =   165
            TabIndex        =   74
            Top             =   855
            Width           =   1340
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Materno :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   24
            Left            =   3630
            TabIndex        =   72
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Paterno :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   23
            Left            =   150
            TabIndex        =   70
            Top             =   285
            Width           =   1340
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Height          =   195
            Index           =   7
            Left            =   4575
            TabIndex        =   78
            Top             =   1155
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Documento de Identificación :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   26
            Left            =   3630
            TabIndex        =   76
            Top             =   855
            Width           =   2070
         End
         Begin VB.Label lblDato 
            Caption         =   "Cargo :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   27
            Left            =   165
            TabIndex        =   80
            Top             =   1425
            Width           =   1340
         End
      End
      Begin Threed.SSCheck chkPrnLiqRazon 
         Height          =   255
         Left            =   -74295
         TabIndex        =   84
         Top             =   720
         Width           =   2805
         _Version        =   65536
         _ExtentX        =   4939
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Razón Social Formatos Liquidación "
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkPrnLiqLogo 
         Height          =   255
         Left            =   -69975
         TabIndex        =   85
         Top             =   720
         Width           =   2805
         _Version        =   65536
         _ExtentX        =   4939
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Logo Formatos de Liquidación "
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1875
         Index           =   8
         Left            =   -74760
         TabIndex        =   116
         Top             =   2730
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   3307
         _StockProps     =   14
         Caption         =   " Configuración documento de Contrato "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtDirContrato 
            Height          =   280
            Index           =   0
            Left            =   240
            TabIndex        =   118
            Top             =   720
            Width           =   5500
         End
         Begin VB.TextBox txtDirContrato 
            Height          =   280
            Index           =   1
            Left            =   240
            TabIndex        =   121
            Top             =   1350
            Width           =   5500
         End
         Begin Threed.SSCommand cmdDirectorio 
            Height          =   360
            Index           =   0
            Left            =   5850
            TabIndex        =   119
            Top             =   690
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   635
            _StockProps     =   78
            ForeColor       =   -2147483631
            BevelWidth      =   0
            Outline         =   0   'False
            AutoSize        =   2
            Picture         =   "prmempresa.frx":00D4
         End
         Begin Threed.SSCommand cmdDirectorio 
            Height          =   360
            Index           =   1
            Left            =   5850
            TabIndex        =   122
            Top             =   1320
            Width           =   390
            _Version        =   65536
            _ExtentX        =   688
            _ExtentY        =   635
            _StockProps     =   78
            ForeColor       =   -2147483631
            BevelWidth      =   0
            Outline         =   0   'False
            AutoSize        =   2
            Picture         =   "prmempresa.frx":00F0
         End
         Begin VB.Label lblDato 
            AutoSize        =   -1  'True
            Caption         =   "Directorio plantilla de contrato :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   117
            Top             =   480
            Width           =   2190
         End
         Begin VB.Label lblDato 
            AutoSize        =   -1  'True
            Caption         =   "Directorio almacenar contrato :"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   33
            Left            =   240
            TabIndex        =   120
            Top             =   1110
            Width           =   2175
         End
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   510
         Index           =   3
         Left            =   -74715
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   8175
      End
      Begin VB.Label lblDato 
         Caption         =   "Nivel Mov. Centro Costo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   50
         Left            =   6045
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label lblDetalle 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   3
         Left            =   -74850
         TabIndex        =   83
         Top             =   150
         Width           =   375
      End
      Begin VB.Label lblDetalle 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   2
         Left            =   -74850
         TabIndex        =   53
         Top             =   150
         Width           =   375
      End
      Begin VB.Label lblDetalle 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   1
         Left            =   -74850
         TabIndex        =   29
         Top             =   150
         Width           =   375
      End
      Begin VB.Label lblDetalle 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ... "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   150
         Width           =   375
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   8910
      _Version        =   65536
      _ExtentX        =   15716
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Begin Threed.SSCommand cmdCancel 
         Height          =   360
         Left            =   7935
         TabIndex        =   90
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prmempresa.frx":010C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   7545
         TabIndex        =   91
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prmempresa.frx":0128
      End
      Begin Threed.SSCommand cmdProceso 
         Height          =   360
         Left            =   345
         TabIndex        =   92
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prmempresa.frx":0144
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1035
         TabIndex        =   89
         Top             =   120
         Width           =   6000
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   93
      Top             =   5835
      Width           =   8910
      _Version        =   65536
      _ExtentX        =   15716
      _ExtentY        =   900
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   106
      Top             =   480
      Visible         =   0   'False
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   4233
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2117"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "fPrmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_Registro As String                            ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_ModoPll As String, s_ModoCfg As String        ' Modo de actualización de los datos
Private s_BaseData As String, s_BaseConta As String     ' Base de datos a procesar, contabilidad
'[
Dim cnn As ADODB.Connection
Sub ShowScreen()
    
  ' Datos de configuración de empresa
  s_Sql = "SELECT cfg.pdoano, cfg.codvia, cfg.direccionvia, cfg.numerodir, cfg.codzona, cfg.direccionzona, cfg.ubigeodir, "
  s_Sql = s_Sql & "cfg.regpatronal, cfg.girocomercial, cfg.telefono, cfg.email, cfg.dirimpbol, cfg.contrato_dot, cfg.contrato_doc, "
  s_Sql = s_Sql & "cfg.psnapepaterno, cfg.psnapematerno, cfg.psnnombres, cfg.psntelefono, cfg.codcco, cfg.gratixasis, "
  s_Sql = s_Sql & "cfg.repapepaterno, cfg.repapematerno, cfg.repnombres, cfg.repcoddci, cfg.repnumdocu, cfg.repcargo, "
  s_Sql = s_Sql & "cfg.gerapepaterno, cfg.gerapematerno, cfg.gernombres, cfg.gercoddci, cfg.gernumdocu, cfg.gercargo, "
  s_Sql = s_Sql & "cfg.repimpbol, cfg.nivelcencosto, cfg.liqprn_razonemp, cfg.liqprn_logoemp, cfg.logo, cfg.firma, cfg.firmanexo, "
  s_Sql = s_Sql & "cfg.server_envio,cfg.usuario_envio,cfg.password_envio,correo_envio "
  s_Sql = s_Sql & "FROM plcfgempresa cfg "
  s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & s_BaseData, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    s_ModoCfg = s_MdoData_Upd
    ' Pestaña Inicial
    gdl_Procedure.EditOptionCheck "AT", chkPrnBoletaDir, (IIf(IsNull(porstRecordset!dirimpbol), 0, porstRecordset!dirimpbol) = s_Estado_Act), s_ModoCfg, True
    gdl_Procedure.EditText "AT", txtNivelCenCosto, FormatNumber(porstRecordset!nivelcencosto, 0), s_ModoCfg, False, 2, vbRightJustify

    gdl_Procedure.EditText "AT", txtTipoVia, gdl_Funcion.aTexto(porstRecordset!codvia), s_ModoCfg, False, porstRecordset!codvia.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreVia, gdl_Funcion.aTexto(porstRecordset!direccionvia), s_ModoCfg, False, porstRecordset!direccionvia.DefinedSize
    gdl_Procedure.EditText "AT", txtnumero, gdl_Funcion.aTexto(porstRecordset!numerodir), s_ModoCfg, False, porstRecordset!numerodir.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoZona, gdl_Funcion.aTexto(porstRecordset!codzona), s_ModoCfg, False, porstRecordset!codzona.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreZona, gdl_Funcion.aTexto(porstRecordset!direccionzona), s_ModoCfg, False, porstRecordset!direccionzona.DefinedSize
    gdl_Procedure.EditText "AT", txtUbigeo(0), gdl_Funcion.aTexto(porstRecordset!ubigeodir), s_ModoCfg, False, porstRecordset!ubigeodir.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(0), gdl_Funcion.aTexto(porstRecordset!telefono), s_ModoCfg, False, porstRecordset!telefono.DefinedSize
    gdl_Procedure.EditText "AT", txtComercial, gdl_Funcion.aTexto(porstRecordset!girocomercial), s_ModoCfg, False, porstRecordset!girocomercial.DefinedSize
    gdl_Procedure.EditText "AT", txtPatronal, gdl_Funcion.aTexto(porstRecordset!regpatronal), s_ModoCfg, False, porstRecordset!regpatronal.DefinedSize
    gdl_Procedure.EditText "AT", txtEmail, gdl_Funcion.aTexto(porstRecordset!Email), s_ModoCfg, False, porstRecordset!Email.DefinedSize
    ' Primera Pestaña
    gdl_Procedure.EditOptionCheck "AT", chkGratixDia, (porstRecordset!gratixasis = s_Estado_Act), s_ModoCfg, True
    gdl_Procedure.EditText "AT", txtPersonal(0), gdl_Funcion.aTexto(porstRecordset!psnapepaterno), s_ModoCfg, False, porstRecordset!psnapepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtPersonal(1), gdl_Funcion.aTexto(porstRecordset!psnapematerno), s_ModoCfg, False, porstRecordset!psnapematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtPersonal(2), gdl_Funcion.aTexto(porstRecordset!psnnombres), s_ModoCfg, False, porstRecordset!psnnombres.DefinedSize
    gdl_Procedure.EditText "AT", txtCenCosto, gdl_Funcion.aTexto(porstRecordset!codcco), s_ModoCfg, False, porstRecordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(1), gdl_Funcion.aTexto(porstRecordset!psntelefono), s_ModoCfg, False, porstRecordset!psntelefono.DefinedSize
    ' Segunda pestaña
    gdl_Procedure.EditOptionCheck "AT", chkPrnBoleta, (porstRecordset!repimpbol = s_Estado_Act), s_ModoCfg, True
    gdl_Procedure.EditText "AT", txtRepresentante(0), gdl_Funcion.aTexto(porstRecordset!repapepaterno), s_ModoCfg, False, porstRecordset!repapepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtRepresentante(1), gdl_Funcion.aTexto(porstRecordset!repapematerno), s_ModoCfg, False, porstRecordset!repapematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtRepresentante(2), gdl_Funcion.aTexto(porstRecordset!repnombres), s_ModoCfg, False, porstRecordset!repnombres.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoDocu(0), gdl_Funcion.aTexto(porstRecordset!repcoddci), s_ModoCfg, False, porstRecordset!repcoddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(0), gdl_Funcion.aTexto(porstRecordset!repnumdocu), s_ModoCfg, False, porstRecordset!repnumdocu.DefinedSize
    gdl_Procedure.EditText "AT", txtCargo(0), gdl_Funcion.aTexto(porstRecordset!repcargo), s_ModoCfg, False, porstRecordset!repcargo.DefinedSize
    
    gdl_Procedure.EditText "AT", txtGerente(0), gdl_Funcion.aTexto(porstRecordset!gerapepaterno), s_ModoCfg, False, porstRecordset!gerapepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtGerente(1), gdl_Funcion.aTexto(porstRecordset!gerapematerno), s_ModoCfg, False, porstRecordset!gerapematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtGerente(2), gdl_Funcion.aTexto(porstRecordset!gernombres), s_ModoCfg, False, porstRecordset!gernombres.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoDocu(1), gdl_Funcion.aTexto(porstRecordset!gercoddci), s_ModoCfg, False, porstRecordset!gercoddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(1), gdl_Funcion.aTexto(porstRecordset!gernumdocu), s_ModoCfg, False, porstRecordset!gernumdocu.DefinedSize
    gdl_Procedure.EditText "AT", txtCargo(1), gdl_Funcion.aTexto(porstRecordset!gercargo), s_ModoCfg, False, porstRecordset!gercargo.DefinedSize
    ' Tercera pestaña
    gdl_Procedure.EditOptionCheck "AT", chkPrnLiqRazon, (porstRecordset!liqprn_razonemp = s_Estado_Act), s_ModoCfg, True
    gdl_Procedure.EditOptionCheck "AT", chkPrnLiqLogo, (porstRecordset!liqprn_logoemp = s_Estado_Act), s_ModoCfg, True
    ReadImagen porstRecordset, imgLogo, "logo"
    ReadImagen porstRecordset, imgFirma(0), "firma"
    ReadImagen porstRecordset, imgFirma(1), "firmanexo"
    
    'Quinta Pestaña
    gdl_Procedure.EditCombo "AT", cboServerEnvio, porstRecordset!server_envio, s_ModoCfg, False
    gdl_Procedure.EditText "AT", txtUsuarioEnvio, gdl_Funcion.aTexto(porstRecordset!usuario_envio), s_ModoCfg, False, porstRecordset!usuario_envio.DefinedSize
    gdl_Procedure.EditText "AT", txtPwdCorreoEnvio, gdl_Funcion.aTexto(gdl_Funcion.Desencripta(IIf(IsNull(porstRecordset!password_envio) = True, " ", porstRecordset!password_envio))), s_ModoCfg, False, porstRecordset!password_envio.DefinedSize
    gdl_Procedure.EditText "AT", txtCtaCorreoEnvio, gdl_Funcion.aTexto(porstRecordset!correo_envio), s_ModoCfg, False, porstRecordset!correo_envio.DefinedSize
    gdl_Procedure.EditText "AT", txtDirContrato(0), gdl_Funcion.aTexto(porstRecordset!contrato_dot), s_ModoCfg, False, porstRecordset!contrato_dot.DefinedSize
    gdl_Procedure.EditText "AT", txtDirContrato(1), gdl_Funcion.aTexto(porstRecordset!contrato_doc), s_ModoCfg, False, porstRecordset!contrato_doc.DefinedSize
  Else
    s_ModoCfg = s_MdoData_Ins
    ' Pestaña inicial
    gdl_Procedure.EditOptionCheck "AT", chkPrnBoletaDir, False, s_ModoCfg, True
    gdl_Procedure.EditText "AT", txtNivelCenCosto, FormatNumber(0, 0), s_ModoCfg, False, 2, vbRightJustify
    
    gdl_Procedure.EditText "AT", txtTipoVia, "", s_ModoCfg, False, porstRecordset!codvia.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreVia, "", s_ModoCfg, False, porstRecordset!direccionvia.DefinedSize
    gdl_Procedure.EditText "AT", txtnumero, "", s_ModoCfg, False, porstRecordset!numerodir.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoZona, "", s_ModoCfg, False, porstRecordset!codzona.DefinedSize
    gdl_Procedure.EditText "AT", txtNombreZona, "", s_ModoCfg, False, porstRecordset!direccionzona.DefinedSize
    gdl_Procedure.EditText "AT", txtUbigeo(0), "", s_ModoCfg, False, porstRecordset!ubigeodir.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(0), "", s_ModoCfg, False, porstRecordset!telefono.DefinedSize
    gdl_Procedure.EditText "AT", txtComercial, "", s_ModoCfg, False, porstRecordset!girocomercial.DefinedSize
    gdl_Procedure.EditText "AT", txtPatronal, gdl_Funcion.aTexto(fEmpresa.dcaRegistro.Recordset!rucemp), s_ModoCfg, False, porstRecordset!regpatronal.DefinedSize
    gdl_Procedure.EditText "AT", txtEmail, "", s_ModoCfg, False, porstRecordset!Email.DefinedSize
    ' Primera pestaña
    gdl_Procedure.EditOptionCheck "AT", chkGratixDia, False, s_ModoCfg, True
    gdl_Procedure.EditText "AT", txtPersonal(0), "", s_ModoCfg, False, porstRecordset!psnapepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtPersonal(1), "", s_ModoCfg, False, porstRecordset!psnapematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtPersonal(2), "", s_ModoCfg, False, porstRecordset!psnnombres.DefinedSize
    gdl_Procedure.EditText "AT", txtCenCosto, "", s_ModoCfg, False, porstRecordset!codcco.DefinedSize
    gdl_Procedure.EditText "AT", txtTelefono(1), "", s_ModoCfg, False, porstRecordset!psntelefono.DefinedSize
    ' Segunda pestaña
    gdl_Procedure.EditOptionCheck "AT", chkPrnBoleta, False, s_ModoCfg, True
    gdl_Procedure.EditText "AT", txtRepresentante(0), "", s_ModoCfg, False, porstRecordset!repapepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtRepresentante(1), "", s_ModoCfg, False, porstRecordset!repapematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtRepresentante(2), "", s_ModoCfg, False, porstRecordset!repnombres.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoDocu(0), "", s_ModoCfg, False, porstRecordset!repcoddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(0), "", s_ModoCfg, False, porstRecordset!repnumdocu.DefinedSize
    gdl_Procedure.EditText "AT", txtCargo(0), "", s_ModoCfg, False, porstRecordset!repcargo.DefinedSize
    
    gdl_Procedure.EditText "AT", txtGerente(0), "", s_ModoCfg, False, porstRecordset!gerapepaterno.DefinedSize
    gdl_Procedure.EditText "AT", txtGerente(1), "", s_ModoCfg, False, porstRecordset!gerapematerno.DefinedSize
    gdl_Procedure.EditText "AT", txtGerente(2), "", s_ModoCfg, False, porstRecordset!gernombres.DefinedSize
    gdl_Procedure.EditText "AT", txtTipoDocu(1), "", s_ModoCfg, False, porstRecordset!gercoddci.DefinedSize
    gdl_Procedure.EditText "AT", txtDocumento(1), "", s_ModoCfg, False, porstRecordset!gernumdocu.DefinedSize
    gdl_Procedure.EditText "AT", txtCargo(1), "", s_ModoCfg, False, porstRecordset!gercargo.DefinedSize
    ' Tercera pestaña
    gdl_Procedure.EditOptionCheck "AT", chkPrnLiqRazon, False, s_ModoCfg, True
    gdl_Procedure.EditOptionCheck "AT", chkPrnLiqLogo, False, s_ModoCfg, True
    imgLogo.Picture = LoadPicture()
    imgLogo.Refresh
    imgFirma(0).Picture = LoadPicture()
    imgFirma(0).Refresh
    imgFirma(1).Picture = LoadPicture()
    imgFirma(1).Refresh
    'Quinta Pestaña
    gdl_Procedure.EditCombo "AT", cboServerEnvio, -1, s_ModoCfg, False
    gdl_Procedure.EditText "AT", txtUsuarioEnvio, "", s_ModoCfg, False, porstRecordset!usuario_envio.DefinedSize
    gdl_Procedure.EditText "AT", txtPwdCorreoEnvio, "", s_ModoCfg, False, porstRecordset!password_envio.DefinedSize
    gdl_Procedure.EditText "AT", txtCtaCorreoEnvio, "", s_ModoCfg, False, porstRecordset!correo_envio.DefinedSize
    gdl_Procedure.EditText "AT", txtDirContrato(0), "", s_ModoCfg, False, porstRecordset!contrato_dot.DefinedSize
    gdl_Procedure.EditText "AT", txtDirContrato(1), "", s_ModoCfg, False, porstRecordset!contrato_doc.DefinedSize
 End If
  lblHelp(0).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtTipoVia.Text, "TV")
  lblHelp(1).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtTipoZona.Text, "TZ")
  lblUbigeo(0).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUbigeo(0).Text, "UG")
  lblHelp(4).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseConta, ps_CodEmpresa, txtCenCosto.Text, "CC")
  lblHelp(5).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtTipoDocu(0).Text, "DI")
  lblHelp(6).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_ClsPlanilla, txtCargo(0).Text, "DC")
  lblHelp(7).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtTipoDocu(1).Text, "DI")
  lblHelp(8).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_ClsPlanilla, txtCargo(1).Text, "DC")

  ' Pestaña de Información de configuración
  s_Sql = "SELECT prm.pdoano, prm.cpcbasico, prm.cpcafamiliar, prm.cpccomisi "
  s_Sql = s_Sql & "FROM plparametroafp prm "
  s_Sql = s_Sql & "WHERE prm.pdoano='" & ps_Anyo & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & s_BaseData, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    s_ModoPll = s_MdoData_Upd
    gdl_Procedure.EditText "AT", txtRemunera, gdl_Funcion.aTexto(porstRecordset!cpcbasico), s_ModoPll, False, porstRecordset!cpcbasico.DefinedSize
    gdl_Procedure.EditText "AT", txtAFamiliar, gdl_Funcion.aTexto(porstRecordset!cpcafamiliar), s_ModoPll, False, porstRecordset!cpccomisi.DefinedSize
    gdl_Procedure.EditText "AT", txtComision, gdl_Funcion.aTexto(porstRecordset!cpccomisi), s_ModoPll, False, porstRecordset!cpccomisi.DefinedSize
  Else
    s_ModoPll = s_MdoData_Ins
    gdl_Procedure.EditText "AT", txtRemunera, "", s_ModoPll, False, porstRecordset!cpcbasico.DefinedSize
    gdl_Procedure.EditText "AT", txtAFamiliar, "", s_ModoPll, False, porstRecordset!cpcafamiliar.DefinedSize
    gdl_Procedure.EditText "AT", txtComision, "", s_ModoPll, False, porstRecordset!cpccomisi.DefinedSize
  End If
  lblHelp(2).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtRemunera.Text, "CP")
  lblHelp(3).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtAFamiliar.Text, "CP")
  lblHelp(9).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtComision.Text, "CP")

End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdDirectorio_Click(Index As Integer)
  Dim sDirectorio As String
  
  sDirectorio = fSeleccionDirectorio()
  txtDirContrato(Index) = sDirectorio & IIf(sDirectorio <> "", "\", "")
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim nTop As Double, nLeft As Double
  
  nTop = (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  nLeft = (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Tipo de via de dirección
    tdbHelp.Columns(0).DataField = "codvia": tdbHelp.Columns(1).DataField = "desvia"
    tdbHelp.Caption = "Tipos de Via"
    nTop = (frmCuadro(0).Top + nTop)
    nLeft = (frmCuadro(0).Left + nLeft)
    s_Sql = gdl_Funcion.HelpTablas("via", "codvia", "", "")
   Case 1     ' Tipo de zona de direccion
    tdbHelp.Columns(0).DataField = "codzona": tdbHelp.Columns(1).DataField = "deszona"
    tdbHelp.Caption = "Tipos de Zona"
    nTop = (frmCuadro(0).Top + nTop)
    nLeft = (frmCuadro(0).Left + nLeft)
    s_Sql = gdl_Funcion.HelpTablas("zon", "codzona", "", "")
   Case 2, 3, 9   ' Concepto de remuneración asegurable,comision, asignacion familiar
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos de Ingresos"
    nTop = (frmCuadro(2).Top + nTop)
    nLeft = (frmCuadro(2).Left + nLeft - IIf(Index = 9, cmdHelp(Index).Left / 2, 0))
    ' Recupera informacion
    s_Sql = gdl_Funcion.HelpTablas("cpc", "codcpc", s_Estado_Ina, "")
   Case 4     ' Centro de costo
    tdbHelp.Columns(0).DataField = "codcco": tdbHelp.Columns(1).DataField = "detcco"
    tdbHelp.Caption = "Centro de Costos"
    nLeft = (frmCuadro(3).Left + nLeft)
    ' Recupera informacion
    s_Sql = gdl_Funcion.HelpTablas("cco", "codcco", pn_NivelCenCosto, "")
   Case 5, 7    ' Tipo de documento de identidad
    tdbHelp.Columns(0).DataField = "coddci": tdbHelp.Columns(1).DataField = "desdci"
    tdbHelp.Caption = "Documentos de Identidad"
    nTop = (frmCuadro(4).Top + nTop)
    nLeft = (3425 + (cmdHelp(Index).Width / 2))
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("dci", "coddci", "", "")
   Case 6, 8    ' Cargo
    tdbHelp.Columns(0).DataField = "codcgo": tdbHelp.Columns(1).DataField = "descgo"
    tdbHelp.Caption = "Cargo de  Personal"
    nTop = (frmCuadro(4).Top + nTop)
    nLeft = (frmCuadro(4).Left + nLeft)
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cgo", "codcgo", ps_ClsPlanilla, "")
   End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & IIf(Index = 4, s_BaseConta, s_BaseData), adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + nTop
  tdbHelp.Left = tabRegister.Left + nLeft
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub

Private Sub cmdProceso_Click()
  
  If MsgBox(vbCrLf & "Esta seguro, se copiaran los datos de Datos de Empresa del Año Anterior al Año Actual", vbQuestion + vbYesNo + vbDefaultButton2, "Sistema de Planilla") = vbYes Then
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, s_BaseData)
    
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    
    ' Elimino la Informacion del Año Actual
    s_Sql = "DELETE FROM plcfgempresa "
    s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
    gdl_Conexion.Execucion s_Sql
    If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
    ' Inserto Informacion del Año Anterior al Año Actual
    s_Sql = "INSERT INTO plcfgempresa ("
    s_Sql = s_Sql & "pdoano, codvia, direccionvia, numerodir, codzona, direccionzona, ubigeodir, regpatronal, girocomercial, telefono, email, repapepaterno, repapematerno, "
    s_Sql = s_Sql & "repnombres, repcargo, repcoddci, repnumdocu, gerapepaterno, gerapematerno, gernombres, gercargo, gercoddci, gernumdocu, psnapepaterno, psnapematerno, "
    s_Sql = s_Sql & "psnnombres, psntelefono, codcco, codcpcrem, codtbluit, codcpc5ta, codcpc5ta_Ing, repimpbol, dirimpbol, contrato_dot, contrato_doc, rembasica, rempromedio, "
    s_Sql = s_Sql & "rempendiente, gratipendiente, remanterior, remganada, codtblretener, codtblpendiente, codtbldividir, gratixasis, gratiliqxdias, remxutiejer1, remxutiejer2, "
    s_Sql = s_Sql & "remxutiejer3, remxutiejer4, rentaxejer_mn, rentaxejer_me, porcepartici, nivelcencosto, logo, firma, firmanexo, usrcre, fyhcre) "
    s_Sql = s_Sql & "SELECT " & ps_Anyo & " as pdoano, codvia, direccionvia, numerodir, codzona, direccionzona, ubigeodir, regpatronal, girocomercial, telefono, email, repapepaterno, repapematerno, "
    s_Sql = s_Sql & "repnombres, repcargo, repcoddci, repnumdocu, gerapepaterno, gerapematerno, gernombres, gercargo, gercoddci, gernumdocu, psnapepaterno, psnapematerno, "
    s_Sql = s_Sql & "psnnombres, psntelefono, codcco, codcpcrem, codtbluit, codcpc5ta, codcpc5ta_Ing, repimpbol, dirimpbol, contrato_dot, contrato_doc, rembasica, rempromedio, "
    s_Sql = s_Sql & "rempendiente, gratipendiente, remanterior, remganada, codtblretener, codtblpendiente, codtbldividir, gratixasis, gratiliqxdias, remxutiejer1, remxutiejer2, "
    s_Sql = s_Sql & "remxutiejer3, remxutiejer4, rentaxejer_mn, rentaxejer_me, porcepartici, nivelcencosto, logo, firma, firmanexo, '" & ps_Usuario & "' AS usrcre, '" & Format(Now, s_FmtFeHoMysql_0) & "' AS fyhcre "
    s_Sql = s_Sql & "FROM plcfgempresa "
    s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo - 1 & "'"
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Error
    
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
    Unload Me
    fPrmEmpresa.Show
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub

Private Sub cmdUbigeo_Click(Index As Integer)
  Set o_SwSelUbica = fPrmEmpresa: n_SwSelUbica = Index
  fSeleccionUbigeo.Show vbModal
  Set o_SwSelUbica = Nothing
  Exit Sub
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  Dim s_PrnRepBoleta As String * 1, s_GratixAsis As String * 1
  Dim s_PrnLiqRazon As String * 1, s_PrnLiqLogo As String * 1
  Dim s_PrnRepBoletaDir As String * 1
  Dim n_Puerto_ServerMail As Integer
  Dim sDirectorioDot As String, sDirectorioDoc As String

  
  ' Primera pestaña (empresa)
  tabRegister.Tab = 0
  If CInt(txtNivelCenCosto.Text) <= 0 Then Beep: MsgBox "Nivel de Centro de Costo no valido; Verificar", vbExclamation: txtNivelCenCosto.SetFocus: Exit Sub
  If txtTipoVia.Text <> "" And lblHelp(0).Caption = "???" Then Beep: MsgBox "Dirección - Tipo de Via no es valido; Verificar", vbExclamation: txtTipoVia.SetFocus: Exit Sub
  If txtTipoZona.Text <> "" And lblHelp(1).Caption = "???" Then Beep: MsgBox "Dirección - Tipo de Zona no es valido; Verificar", vbExclamation: txtTipoZona.SetFocus: Exit Sub
  If txtUbigeo(0).Text <> "" And lblUbigeo(0).Caption = "???" Then Beep: MsgBox "Dirección - Ubicacion Geografica no es valido; Verificar", vbExclamation: txtUbigeo(0).SetFocus: Exit Sub
  If txtTelefono(0).Text = "" Then Beep: MsgBox "Debe Ingresar numero de teléfono de la empresa", vbExclamation:  txtTelefono(0).SetFocus: Exit Sub
  If txtPatronal.Text = "" Then Beep: MsgBox "Debe Ingresar el Registro Patronal", vbExclamation: txtPatronal.SetFocus: Exit Sub
  
  ' Segunda pestaña (planilla)
  tabRegister.Tab = 1
  If txtRemunera.Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Básica", vbExclamation: txtRemunera.SetFocus: Exit Sub
  If lblHelp(2).Caption = "???" Then Beep: MsgBox "Concepto de Remuneración Básica no es valido; Verificar", vbExclamation: txtRemunera.SetFocus: Exit Sub
  If txtAFamiliar.Text = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Asignación Familiar", vbExclamation: txtAFamiliar.SetFocus: Exit Sub
  If lblHelp(3).Caption = "???" Then Beep: MsgBox "Concepto de Asignación Familiar no es valido; Verificar", vbExclamation: txtAFamiliar.SetFocus: Exit Sub
  If txtComision.Text <> "" And lblHelp(9) = "???" Then Beep: MsgBox "Concepto de Comisión no es valido; Verificar", vbExclamation: txtComision.SetFocus: Exit Sub
  If Trim(txtPersonal(0)) = "" And Trim(txtPersonal(1)) = "" And Trim(txtPersonal(2)) = "" Then Beep: MsgBox "Debe Ingresar los nombres del Personal de Planilla", vbExclamation: txtPersonal(0).SetFocus: Exit Sub
  If txtCenCosto = "" Then Beep: MsgBox "Debe Ingresar el Area del Personal", vbExclamation: txtCenCosto.SetFocus: Exit Sub
  If lblHelp(4).Caption = "???" Then Beep: MsgBox "Centro de costo no es valido; Verificar", vbExclamation: txtCenCosto.SetFocus: Exit Sub
  
  ' Tercera pestaña (representación)
  tabRegister.Tab = 2
  If Trim(txtRepresentante(0)) = "" And Trim(txtRepresentante(1)) = "" And Trim(txtRepresentante(2)) = "" Then Beep: MsgBox "Debe Ingresar los nombres del Representante Legal", vbExclamation: txtRepresentante(0).SetFocus: Exit Sub
  If txtTipoDocu(0) = "" Then Beep: MsgBox "Debe Ingresar Tipo documento Identidad Representante Legal", vbExclamation: txtTipoDocu(0).SetFocus: Exit Sub
  If lblHelp(5).Caption = "???" Then Beep: MsgBox "Tipo Documento Identidad Representante Legal no es valido; Verificar", vbExclamation: txtTipoDocu(0).SetFocus: Exit Sub
  If txtDocumento(0) = "" Then Beep: MsgBox "Debe Ingresar documento Identidad Representante Legal", vbExclamation: txtDocumento(0).SetFocus: Exit Sub
  If txtCargo(0) = "" Then Beep: MsgBox "Debe Ingresar Cargo Represntante legal", vbExclamation: txtCargo(0).SetFocus: Exit Sub
  If lblHelp(6).Caption = "???" Then Beep: MsgBox "Cargo Representante Legal no valido; Verificar", vbExclamation: txtCargo(0).SetFocus: Exit Sub
  
  If (Trim(txtGerente(0).Text) <> "" And Trim(txtGerente(1).Text) = "" And Trim(txtGerente(2).Text) = "") Then Beep: MsgBox "Debe Completar los nombres del Gerente Adjunto", vbExclamation: txtGerente(0).SetFocus: Exit Sub
  If (Trim(txtGerente(0).Text) <> "" And txtTipoDocu(1).Text = "") Then Beep: MsgBox "Debe Ingresar Tipo documento Identidad Gerente Adjunto", vbExclamation: txtTipoDocu(1).SetFocus: Exit Sub
  If lblHelp(7).Caption = "???" Then Beep: MsgBox "Tipo Documento Identidad Gerente Adjunto no es valido; Verificar", vbExclamation: txtTipoDocu(1).SetFocus: Exit Sub
  If (txtTipoDocu(1).Text <> "" And txtDocumento(1).Text = "") Then Beep: MsgBox "Debe Ingresar documento Identidad Gerente Adjunto", vbExclamation: txtDocumento(1).SetFocus: Exit Sub
  If (Trim(txtGerente(0).Text) <> "" And txtCargo(1) = "") Then Beep: MsgBox "Debe Ingresar Cargo Gerente Adjunto", vbExclamation: txtCargo(1).SetFocus: Exit Sub
  If lblHelp(8).Caption = "???" Then Beep: MsgBox "Cargo Gerente Adjunto no valido; Verificar", vbExclamation: txtCargo(1).SetFocus: Exit Sub
  tabRegister.Tab = 0
  
  s_PrnRepBoleta = IIf(chkPrnBoleta.Value, s_Estado_Act, s_Estado_Ina)
  s_GratixAsis = IIf(chkGratixDia.Value, s_Estado_Act, s_Estado_Ina)
  s_PrnRepBoletaDir = IIf(chkPrnBoletaDir.Value, s_Estado_Act, s_Estado_Ina)
  
  s_PrnLiqRazon = IIf(chkPrnLiqRazon.Value, s_Estado_Act, s_Estado_Ina)
  s_PrnLiqLogo = IIf(chkPrnLiqLogo.Value, s_Estado_Act, s_Estado_Ina)
  
  'Quinta Pestaña
'  If cboServerEnvio.Text = "" Then Beep: MsgBox "Debe Seleccionar un Servidor de Correo de Envío", vbExclamation: cboServerEnvio.SetFocus: Exit Sub
'  If txtUsuarioEnvio.Text = "" Then Beep: MsgBox "Debe Ingresar el Usuario de Envìo", vbExclamation: txtUsuarioEnvio.SetFocus: Exit Sub
'  If txtCtaCorreoEnvio.Text = "" Then Beep: MsgBox "Debe Ingresar la Cuenta de Correo de Envío", vbExclamation: txtCtaCorreoEnvio.SetFocus: Exit Sub
'  If txtPwdCorreoEnvio.Text = "" Then Beep: MsgBox "Debe Ingresar la Contraseña de Correo de Envío", vbExclamation: txtPwdCorreoEnvio.SetFocus: Exit Sub
  If (txtCtaCorreoEnvio.Text <> "" And gdl_Funcion.ValidaEmail(txtCtaCorreoEnvio.Text) = False) Then Beep: MsgBox "Debe ingresar un Correo de Envío Valido", vbExclamation: txtCtaCorreoEnvio.SetFocus: Exit Sub
  If (txtUsuarioEnvio.Text <> "" And gdl_Funcion.ValidaEmail(txtUsuarioEnvio.Text) = False) Then Beep: MsgBox "Debe ingresar un Usuario de Envío Valido", vbExclamation: txtUsuarioEnvio.SetFocus: Exit Sub
  sDirectorioDot = Trim(txtDirContrato(0).Text)
  If sDirectorioDot <> "" Then
    sDirectorioDot = Left(sDirectorioDot, Len(sDirectorioDot) - 1)
  End If
  If sDirectorioDot <> "" And dir(sDirectorioDot, vbDirectory) = "" Then Beep: MsgBox "Directorio de plantilla de contrato no valido; Verifique", vbExclamation: txtDirContrato(0).SetFocus: Exit Sub
  sDirectorioDoc = Trim(txtDirContrato(1).Text)
  If sDirectorioDoc <> "" Then
    sDirectorioDoc = Left(sDirectorioDoc, Len(sDirectorioDoc) - 1)
  End If
  If sDirectorioDoc <> "" And dir(sDirectorioDoc, vbDirectory) = "" Then Beep: MsgBox "Directorio de almacenar contrato no valido; Verifique", vbExclamation: txtDirContrato(1).SetFocus: Exit Sub
  
  n_Puerto_ServerMail = Choose(cboServerEnvio.ListIndex + 1, 0, 465, 995)
  sDirectorioDot = Replace(Trim(txtDirContrato(0).Text), "\", "\\")
  sDirectorioDoc = Replace(Trim(txtDirContrato(1).Text), "\", "\\")
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
    
  ' Creo los arreglos para la actualización planilla
  a_Campos = Array("pdoano", "codvia", "direccionvia", "numerodir", "codzona", "direccionzona", "ubigeodir", "dirimpbol", "regpatronal", "girocomercial", "telefono", "email", "psnapepaterno", "psnapematerno", "psnnombres", "psntelefono", "codcco", "gratixasis", "repimpbol", "repapepaterno", "repapematerno", "repnombres", "repcoddci", "repnumdocu", "repcargo", "gerapepaterno", "gerapematerno", "gernombres", "gercoddci", "gernumdocu", "gercargo", "nivelcencosto", "liqprn_razonemp", "liqprn_logoemp", "server_envio", "usuario_envio", "password_envio", "correo_envio", "puerto_envio", "contrato_dot", "contrato_doc", IIf(s_ModoCfg = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(s_ModoCfg = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_Anyo, txtTipoVia.Text, txtNombreVia.Text, txtnumero.Text, txtTipoZona.Text, txtNombreZona.Text, txtUbigeo(0).Text, s_PrnRepBoletaDir, txtPatronal.Text, txtComercial.Text, txtTelefono(0).Text, txtEmail.Text, Trim(txtPersonal(0).Text), Trim(txtPersonal(1).Text), Trim(txtPersonal(2).Text), txtTelefono(1).Text, txtCenCosto.Text, s_GratixAsis, s_PrnRepBoleta, Trim(txtRepresentante(0).Text), Trim(txtRepresentante(1).Text), Trim(txtRepresentante(2).Text), txtTipoDocu(0).Text, txtDocumento(0).Text, txtCargo(0).Text, Trim(txtGerente(0).Text), Trim(txtGerente(1).Text), Trim(txtGerente(2).Text), txtTipoDocu(1).Text, txtDocumento(1).Text, txtCargo(1).Text, CInt(txtNivelCenCosto.Text), s_PrnLiqRazon, s_PrnLiqLogo, cboServerEnvio.ListIndex, txtUsuarioEnvio.Text, gdl_Funcion.Encripta(txtPwdCorreoEnvio.Text), txtCtaCorreoEnvio.Text, n_Puerto_ServerMail, sDirectorioDot, sDirectorioDoc, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("pdoano")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, s_BaseData)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If s_ModoCfg = s_MdoData_Ins Then
    If Not Records_Ins("plcfgempresa", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plcfgempresa", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  ' Realizo la grabacion de imagenes - logo, firma
  s_Sql = "SELECT pdoano, logo, firma, firmanexo "
  s_Sql = s_Sql & "FROM plcfgempresa "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenKeyset, adLockOptimistic, adUseClient, s_Sql)
  If Not WriteImagen(porstRecordset, imgLogo, "logo") Then GoTo Error
  If Not WriteImagen(porstRecordset, imgFirma(0), "firma") Then GoTo Error
  If Not WriteImagen(porstRecordset, imgFirma(1), "firmanexo") Then GoTo Error
  porstRecordset.Close
  
  ' Creo los arreglos para la actualización parametros
  a_Campos = Array("pdoano", "cpcbasico", "cpcafamiliar", "cpccomisi", IIf(s_ModoPll = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(s_ModoPll = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_Anyo, txtRemunera.Text, txtAFamiliar.Text, txtComision.Text, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("pdoano")
  ' Realizo el proceso de actualización de los registros
  If s_ModoPll = s_MdoData_Ins Then
    If Not Records_Ins("plparametroafp", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plparametroafp", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Ubico el registro ingresado o actualizado
  ShowScreen
  txtTipoVia.SetFocus
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 6770: Me.Width = 9000
  Me.Left = 2580: Me.Top = 1750
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Parametros de Configuración"
  lblTitle = "Parametros"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Datos de Planilla
  s_BaseData = Trim(fEmpresa.dcaRegistro.Recordset!nombredbemp)
  s_BaseConta = s_BaseData
  If fEmpresa.dcaRegistro.Recordset!siscon = s_Estado_Act Then
    s_BaseConta = "c" & Trim(fEmpresa.dcaRegistro.Recordset!codemp) & ps_Anyo
  End If
  
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & s_BaseConta & ";connection="
  cnn.CursorLocation = adUseClient
  cnn.Open

  s_Registro = Trim(fEmpresa.dcaRegistro.Recordset!codemp)
  For n_Index = 0 To 3
    lblDetalle(n_Index) = " Empresa : " & Trim(fEmpresa.dcaRegistro.Recordset!razemp) & " "
  Next n_Index
  
  ' Adiciono los tipo de conceptos
  For n_Index = 0 To 2
    cboServerEnvio.AddItem Choose(n_Index + 1, "Outlook", "Gmail", "Hotmail")
  Next n_Index
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(1, 2)
  ' Icono y título del formulario
  aElemento(1, 1) = "edit": aElemento(1, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  aElemento(0, 1) = "aceptar"
  aElemento(0, 2) = "Actualizar Información de " & lblTitle
  gdl_Procedure.ViewGrafics Me, cmdUpdate, aElemento
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdProceso, "consolid", "Copia información del Año Anterior " & lblTitle
  gdl_Procedure.LoadGrafics cmdDirectorio(0), "carpetas", "Seleciona ubicación de plantilla de contratos"
  gdl_Procedure.LoadGrafics cmdDirectorio(1), "director", "Seleciona ubicación de contratos"
  cmdCancel.Cancel = True
  
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codcpc", "descpc")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 734.7402, 3465.071)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0)
  Next n_Index
  
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = n_BackColorHelp#: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbHelp, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de Cálculo", 2
  ']
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
  fMenu.cmbejercicio.Enabled = Not Cancel
End Sub

Private Sub imgFirma_DblClick(Index As Integer)
  
  On Error GoTo CancelaDialogo
  fMenu.cdlDialogo.DialogTitle = "Seleccionar Imagen"
  fMenu.cdlDialogo.CancelError = True
  fMenu.cdlDialogo.Flags = cdlOFNHideReadOnly
  fMenu.cdlDialogo.DefaultExt = ".bmp"
  fMenu.cdlDialogo.Filter = "Imagen BMP (*.bmp)|*.bmp|Imagen JPEG(*.jpg)|*.jpg|Imagen GIF (*.gif)|*.gif|Todos los archivos(*.*)|*.*"
  fMenu.cdlDialogo.FilterIndex = 1
  fMenu.cdlDialogo.ShowOpen
  imgFirma(Index).Picture = LoadPicture(fMenu.cdlDialogo.FileName)
  imgFirma(Index).Tag = fMenu.cdlDialogo.FileName
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then
    MsgBox Error(Err.Number)
    Exit Sub
  End If
  On Error GoTo 0

End Sub
Private Sub imgFirma_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Elimino la fotografia
  If Button = vbRightButton And Shift = s_Estado_Ina Then
    If MsgBox("Desea Eliminar Firma del Representante", vbQuestion + vbYesNo) = vbYes Then
      imgFirma(Index).Picture = LoadPicture("")
      imgFirma(Index).Tag = ""
    End If
  End If
End Sub

Private Sub imgLogo_DblClick()
  
  On Error GoTo CancelaDialogo
  fMenu.cdlDialogo.DialogTitle = "Seleccionar Imagen"
  fMenu.cdlDialogo.CancelError = True
  fMenu.cdlDialogo.Flags = cdlOFNHideReadOnly
  fMenu.cdlDialogo.DefaultExt = ".bmp"
  fMenu.cdlDialogo.Filter = "Imagen BMP (*.bmp)|*.bmp|Imagen JPEG(*.jpg)|*.jpg|Imagen GIF (*.gif)|*.gif|Todos los archivos(*.*)|*.*"
  fMenu.cdlDialogo.FilterIndex = 1
  fMenu.cdlDialogo.ShowOpen
  imgLogo.Picture = LoadPicture(fMenu.cdlDialogo.FileName)
  imgLogo.Tag = fMenu.cdlDialogo.FileName
  
CancelaDialogo:
  ' veriofico si existe error y desactivo
  If Not Err.Number = 0 Then
    MsgBox Error(Err.Number)
    Exit Sub
  End If
  On Error GoTo 0

End Sub
Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Elimino la fotografia
  If Button = vbRightButton And Shift = s_Estado_Ina Then
    If MsgBox("Desea Eliminar Logo de la Empresa", vbQuestion + vbYesNo) = vbYes Then
      imgLogo.Picture = LoadPicture("")
      imgLogo.Tag = ""
    End If
  End If
End Sub

Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Tipo de via
    txtTipoVia.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtTipoVia.SetFocus
   Case 1       ' Tipo de zona
    txtTipoZona.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtTipoZona.SetFocus
   Case 2 ' Concepto de remuneración asegurable
    txtRemunera.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtRemunera.SetFocus
   Case 3       ' Asignacion Familiar
    txtAFamiliar.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtAFamiliar.SetFocus
   Case 4       ' Centro de costo
    txtCenCosto.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtCenCosto.SetFocus
   Case 5, 7      ' Tipo de documento identidad
    txtTipoDocu(IIf(n_IndexHelp = 5, 0, 1)).Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtTipoDocu(IIf(n_IndexHelp = 5, 0, 1)).SetFocus
   Case 6, 8      ' Cargo
    txtCargo(IIf(n_IndexHelp = 6, 0, 1)).Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtCargo(IIf(n_IndexHelp = 6, 0, 1)).SetFocus
   Case 9       ' Comision
    txtComision.Text = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp).Caption = tdbHelp.Columns(1).Value
    txtComision.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Tipo de via de dirección
    s_Sql = gdl_Funcion.HelpTablas("via", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 1     ' Tipo de zona de direccion
    s_Sql = gdl_Funcion.HelpTablas("zon", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 2, 3, 9 ' Concepto de remuneración asegurable, afamiliar, comisiones
    s_Sql = gdl_Funcion.HelpTablas("cpc", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina, "")
   Case 4     ' Centro de costo
    s_Sql = gdl_Funcion.HelpTablas("cco", tdbHelp.Columns(ColIndex).DataField, pn_NivelCenCosto, "")
   Case 5, 7    ' Tipo de documento de identidad
    s_Sql = gdl_Funcion.HelpTablas("dci", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 6, 8    ' Cargo
    s_Sql = gdl_Funcion.HelpTablas("cgo", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & IIf(n_IndexHelp = 4, s_BaseConta, s_BaseData), adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
End Sub
Private Sub tdbHelp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or (KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown) Then s_SqlHelp = ""
  If KeyCode = vbKeyF5 Then porstHelp.Requery
End Sub
Private Sub tdbHelp_KeyPress(KeyAscii As Integer)
  Dim porstClone As ADODB.Recordset
  Dim n_Columna As Integer, s_Criterio As String

  If KeyAscii = vbKeyReturn Then
    tdbHelp_DblClick
  ElseIf (UCase$(Chr$(KeyAscii)) >= "A" And UCase$(Chr$(KeyAscii)) <= "Z") Or _
       (Chr$(KeyAscii) >= "0" And Chr$(KeyAscii) <= "9") Or KeyAscii = 32 Or Chr$(KeyAscii) = "." _
       Or Chr$(KeyAscii) = "*" Then
    ' Conformo la cadena de ayuda
    s_SqlHelp = s_SqlHelp & UCase$(Chr$(KeyAscii))
    Set porstClone = porstHelp.Clone()
    
    n_Columna = tdbHelp.Col
    s_Criterio = tdbHelp.Columns(n_Columna).DataField & " >= '" & s_SqlHelp & "'"
    porstClone.Find s_Criterio, 0, adSearchForward, 0
    If Not (porstClone.BOF Or porstClone.EOF) Then
      porstHelp.Bookmark = porstClone.Bookmark
    End If
    porstClone.Close
    Set porstClone = Nothing
  Else
      s_SqlHelp = ""
  End If

End Sub
Private Sub tdbHelp_LostFocus()
  tdbHelp.Visible = False
End Sub
Private Sub txtAFamiliar_GotFocus()
  gdl_Procedure.MarcaGet txtAFamiliar
End Sub
Private Sub txtAFamiliar_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 3
End Sub
Private Sub txtAFamiliar_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtAFamiliar_LostFocus()
  lblHelp(3) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtAFamiliar, "CP")
End Sub
Private Sub txtCenCosto_GotFocus()
  gdl_Procedure.MarcaGet txtCenCosto
End Sub
Private Sub txtCenCosto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 4
End Sub
Private Sub txtCenCosto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCenCosto_LostFocus()
  lblHelp(4) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseConta, ps_CodEmpresa, txtCenCosto, "CC")
End Sub
Private Sub txtCargo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCargo(Index)
End Sub
Private Sub txtCargo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Choose(Index + 1, 6, 8)
End Sub
Private Sub txtCargo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCargo_LostFocus(Index As Integer)
  lblHelp(Choose(Index + 1, 6, 8)) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseConta, ps_ClsPlanilla, txtCargo(Index).Text, "DC")
End Sub
Private Sub txtComercial_GotFocus()
  gdl_Procedure.MarcaGet txtComercial
End Sub
Private Sub txtComercial_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub txtComision_GotFocus()
  gdl_Procedure.MarcaGet txtComision
End Sub
Private Sub txtComision_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 9
End Sub
Private Sub txtComision_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtComision_LostFocus()
  lblHelp(9) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtComision.Text, "CP")
End Sub
Private Sub txtCtaCorreoEnvio_GotFocus()
  gdl_Procedure.MarcaGet txtCtaCorreoEnvio
End Sub
Private Sub txtCtaCorreoEnvio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtDirContratoGotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtDirContrato(Index)
End Sub
Private Sub ttxtDirContrato_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtDocumento_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtDocumento(Index)
End Sub
Private Sub txtDocumento_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtEmail_GotFocus()
  gdl_Procedure.MarcaGet txtEmail
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNivelCenCosto_GotFocus()
  gdl_Procedure.MarcaGet txtNivelCenCosto
End Sub
Private Sub txtNivelCenCosto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNivelCenCosto_Validate(Cancel As Boolean)
  txtNivelCenCosto.Text = IIf(Not IsNumeric(txtNivelCenCosto.Text), 0, txtNivelCenCosto.Text)
  txtNivelCenCosto.Text = FormatNumber(CDec(txtNivelCenCosto.Text), 0)
End Sub
Private Sub txtNombreVia_GotFocus()
  gdl_Procedure.MarcaGet txtNombreVia
End Sub
Private Sub txtNombreVia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNombreZona_GotFocus()
  gdl_Procedure.MarcaGet txtNombreZona
End Sub
Private Sub txtNombreZona_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtnumero_GotFocus()
  gdl_Procedure.MarcaGet txtnumero
End Sub
Private Sub txtnumero_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPatronal_GotFocus()
  gdl_Procedure.MarcaGet txtPatronal
End Sub
Private Sub txtPatronal_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPersonal_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtPersonal(Index)
End Sub
Private Sub txtPersonal_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPwdCorreoEnvio_GotFocus()
  gdl_Procedure.MarcaGet txtPwdCorreoEnvio
End Sub
Private Sub txtPwdCorreoEnvio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_GotFocus()
  gdl_Procedure.MarcaGet txtRemunera
End Sub
Private Sub txtRemunera_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 2
End Sub
Private Sub txtRemunera_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_LostFocus()
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtRemunera, "CP")
End Sub
Private Sub txtRepresentante_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtRepresentante(Index)
End Sub
Private Sub txtRepresentante_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTelefono_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtTelefono(Index)
End Sub
Private Sub txtTelefono_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoDocu_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtTipoDocu(Index)
End Sub
Private Sub txtTipoDocu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Choose(Index + 1, 5, 7)
End Sub
Private Sub txtTipoDocu_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoDocu_LostFocus(Index As Integer)
  lblHelp(Choose(Index + 1, 5, 7)) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtTipoDocu(Index).Text, "DI")
End Sub
Private Sub txtTipoVia_GotFocus()
  gdl_Procedure.MarcaGet txtTipoVia
End Sub
Private Sub txtTipoVia_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtTipoVia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoVia_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtTipoVia, "TV")
End Sub
Private Sub txtTipoZona_GotFocus()
  gdl_Procedure.MarcaGet txtTipoZona
End Sub
Private Sub txtTipoZona_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 1
End Sub
Private Sub txtTipoZona_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoZona_LostFocus()
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & s_BaseData, ps_CodEmpresa, txtTipoZona, "TZ")
End Sub
Private Sub txtUbigeo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtUbigeo(Index)
End Sub
Private Sub txtUbigeo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtUbigeo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtUbigeo_LostFocus(Index As Integer)
  lblUbigeo(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUbigeo(Index), "UG")
End Sub
Private Sub txtUsuarioEnvio_GotFocus()
  gdl_Procedure.MarcaGet txtUsuarioEnvio
End Sub
Private Sub txtUsuarioEnvio_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

