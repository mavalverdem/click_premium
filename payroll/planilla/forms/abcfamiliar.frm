VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fAbcFamiliar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7200
   Icon            =   "abcfamiliar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   7200
   Begin TabDlg.SSTab tabRegister 
      Height          =   4410
      Left            =   75
      TabIndex        =   58
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   7779
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3052
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
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "abcfamiliar.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmCuadro(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmCuadro(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkIncapacidad"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCertificado"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbVinculo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCarta"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmCuadro(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbdocumento"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtpaternidad"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Domicilio"
      TabPicture(1)   =   "abcfamiliar.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkDomicilio"
      Tab(1).Control(1)=   "txtUbigeo(0)"
      Tab(1).Control(2)=   "txtReferencia"
      Tab(1).Control(3)=   "txtNombreZona"
      Tab(1).Control(4)=   "txtTipoZona"
      Tab(1).Control(5)=   "txtNumero(1)"
      Tab(1).Control(6)=   "txtNumero(0)"
      Tab(1).Control(7)=   "txtNombreVia"
      Tab(1).Control(8)=   "txtTipoVia"
      Tab(1).Control(9)=   "cmdHelp(1)"
      Tab(1).Control(10)=   "cmdHelp(2)"
      Tab(1).Control(11)=   "cmdUbigeo(0)"
      Tab(1).Control(12)=   "lblUbigeo(0)"
      Tab(1).Control(13)=   "lblDato(12)"
      Tab(1).Control(14)=   "lblDato(17)"
      Tab(1).Control(15)=   "lblDato(15)"
      Tab(1).Control(16)=   "lblHelp(2)"
      Tab(1).Control(17)=   "lblDato(14)"
      Tab(1).Control(18)=   "lblHelp(1)"
      Tab(1).Control(19)=   "lblDato(13)"
      Tab(1).Control(20)=   "lblDato(19)"
      Tab(1).Control(21)=   "lblDato(20)"
      Tab(1).Control(22)=   "lblDato(21)"
      Tab(1).ControlCount=   23
      Begin VB.TextBox txtpaternidad 
         Height          =   280
         Left            =   120
         TabIndex        =   71
         Top             =   3720
         Width           =   2040
      End
      Begin VB.ComboBox cmbdocumento 
         Height          =   315
         ItemData        =   "abcfamiliar.frx":0044
         Left            =   120
         List            =   "abcfamiliar.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   3360
         Width           =   2100
      End
      Begin Threed.SSCheck chkDomicilio 
         Height          =   195
         Left            =   -70515
         TabIndex        =   42
         Top             =   240
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Domicilio propio"
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
      Begin Threed.SSFrame frmCuadro 
         Height          =   2205
         Index           =   2
         Left            =   4110
         TabIndex        =   23
         Top             =   1800
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3519
         _ExtentY        =   3889
         _StockProps     =   14
         Caption         =   " Situación "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.ComboBox cmbSituacion 
            Height          =   315
            ItemData        =   "abcfamiliar.frx":0048
            Left            =   150
            List            =   "abcfamiliar.frx":004A
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   300
            Width           =   1485
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Fallecimiento"
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
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   26
            Top             =   720
            Width           =   690
            _Version        =   65536
            _ExtentX        =   1217
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Otros"
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
         End
         Begin MSComCtl2.DTPicker dtpalta 
            Height          =   285
            Left            =   120
            TabIndex        =   73
            Top             =   1200
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            Format          =   143589377
            CurrentDate     =   37515
         End
         Begin MSMask.MaskEdBox mskFecBaja 
            Height          =   285
            Left            =   120
            TabIndex        =   74
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Alta "
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Baja "
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   69
            Top             =   1560
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCarta 
         Height          =   280
         Left            =   2400
         TabIndex        =   19
         Top             =   3600
         Width           =   1560
      End
      Begin VB.ComboBox cmbVinculo 
         Height          =   315
         ItemData        =   "abcfamiliar.frx":004C
         Left            =   2445
         List            =   "abcfamiliar.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2040
         Width           =   1485
      End
      Begin VB.TextBox txtCertificado 
         Height          =   280
         Left            =   2400
         TabIndex        =   22
         Top             =   3000
         Width           =   1530
      End
      Begin VB.TextBox txtUbigeo 
         Height          =   280
         Index           =   0
         Left            =   -74685
         TabIndex        =   57
         Top             =   3570
         Width           =   975
      End
      Begin VB.TextBox txtReferencia 
         Height          =   280
         Left            =   -74685
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   2955
         Width           =   5535
      End
      Begin VB.TextBox txtNombreZona 
         Height          =   280
         Left            =   -74685
         TabIndex        =   53
         Top             =   2340
         Width           =   3470
      End
      Begin VB.TextBox txtTipoZona 
         Height          =   280
         Left            =   -74685
         TabIndex        =   51
         Top             =   1725
         Width           =   500
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   1
         Left            =   -69960
         TabIndex        =   49
         Top             =   1110
         Width           =   870
      End
      Begin VB.TextBox txtNumero 
         Height          =   280
         Index           =   0
         Left            =   -71010
         TabIndex        =   47
         Top             =   1110
         Width           =   870
      End
      Begin VB.TextBox txtNombreVia 
         Height          =   280
         Left            =   -74685
         TabIndex        =   45
         Top             =   1110
         Width           =   3470
      End
      Begin VB.TextBox txtTipoVia 
         Height          =   280
         Left            =   -74685
         TabIndex        =   43
         Top             =   495
         Width           =   500
      End
      Begin Threed.SSCheck chkIncapacidad 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   2385
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Incapacidad"
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
      Begin Threed.SSFrame frmCuadro 
         Height          =   1180
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   1700
         Width           =   2190
         _Version        =   65536
         _ExtentX        =   3863
         _ExtentY        =   2081
         _StockProps     =   14
         Caption         =   " Documento de Identidad "
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.TextBox txtDocumento 
            Height          =   280
            Left            =   90
            TabIndex        =   15
            Top             =   800
            Width           =   1440
         End
         Begin VB.TextBox txtTipoDocu 
            Height          =   280
            Left            =   90
            TabIndex        =   14
            Top             =   480
            Width           =   500
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   0
            Left            =   645
            TabIndex        =   59
            Top             =   480
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblDato 
            Caption         =   "Tipo  :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   13
            Top             =   255
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
            Index           =   0
            Left            =   960
            TabIndex        =   61
            Top             =   480
            Width           =   195
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1575
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   90
         Width           =   5985
         _Version        =   65536
         _ExtentX        =   10557
         _ExtentY        =   2778
         _StockProps     =   14
         Caption         =   " Datos de Identificación"
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
         Begin VB.TextBox txtNombres 
            Height          =   280
            Index           =   1
            Left            =   3015
            TabIndex        =   4
            Top             =   480
            Width           =   2700
         End
         Begin VB.TextBox txtNombres 
            Height          =   280
            Index           =   0
            Left            =   150
            TabIndex        =   2
            Top             =   480
            Width           =   2700
         End
         Begin VB.TextBox txtNombres 
            Height          =   280
            Index           =   2
            Left            =   150
            TabIndex        =   6
            Top             =   1020
            Width           =   1260
         End
         Begin VB.ComboBox cmbSexo 
            Height          =   315
            ItemData        =   "abcfamiliar.frx":0050
            Left            =   1560
            List            =   "abcfamiliar.frx":0052
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1020
            Width           =   1260
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   285
            Left            =   3000
            TabIndex        =   10
            Top             =   1080
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            Format          =   143589377
            CurrentDate     =   37515
         End
         Begin VB.Label lblDato 
            Caption         =   "Nombres :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   810
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Materno :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   3015
            TabIndex        =   3
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Apellido Paterno :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label lblDato 
            Caption         =   "Fecha de Nacimiento :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   3000
            TabIndex        =   9
            Top             =   840
            Width           =   1680
         End
         Begin VB.Label lblEdad 
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
            Left            =   4320
            TabIndex        =   11
            Top             =   1080
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Sexo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   1560
            TabIndex        =   7
            Top             =   810
            Width           =   1005
         End
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   1
         Left            =   -74130
         TabIndex        =   62
         Top             =   495
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   2
         Left            =   -74130
         TabIndex        =   64
         Top             =   1725
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin Threed.SSCommand cmdUbigeo 
         Height          =   285
         Index           =   0
         Left            =   -73650
         TabIndex        =   66
         Top             =   3570
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin VB.Label lblDato 
         Caption         =   "Documento que Acredita la Paternidad "
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   9
         Left            =   120
         TabIndex        =   68
         Top             =   2920
         Width           =   2175
      End
      Begin VB.Label lblDato 
         Caption         =   "Certificado Médico"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   2400
         TabIndex        =   21
         Top             =   2760
         Width           =   1530
      End
      Begin VB.Label lblDato 
         Caption         =   "Carta Atención Médica "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   2400
         TabIndex        =   18
         Top             =   3360
         Width           =   1680
      End
      Begin VB.Label lblDato 
         Caption         =   "Vinculo Familiar :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   2445
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
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
         Left            =   -73245
         TabIndex        =   67
         Top             =   3615
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Ubigeo :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   -74685
         TabIndex        =   56
         Top             =   3315
         Width           =   1335
      End
      Begin VB.Label lblDato 
         Caption         =   "Referencia :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   -74685
         TabIndex        =   54
         Top             =   2700
         Width           =   1335
      End
      Begin VB.Label lblDato 
         Caption         =   "Nombre de Zona :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -74685
         TabIndex        =   52
         Top             =   2085
         Width           =   2280
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
         Left            =   -73725
         TabIndex        =   65
         Top             =   1770
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Tipo de Zona :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   -74685
         TabIndex        =   50
         Top             =   1470
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
         Index           =   1
         Left            =   -73725
         TabIndex        =   63
         Top             =   540
         Width           =   195
      End
      Begin VB.Label lblDato 
         Caption         =   "Interior :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   -69960
         TabIndex        =   48
         Top             =   855
         Width           =   870
      End
      Begin VB.Label lblDato 
         Caption         =   "Número :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   -71010
         TabIndex        =   46
         Top             =   855
         Width           =   870
      End
      Begin VB.Label lblDato 
         Caption         =   "Nombre de Vía :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   -74685
         TabIndex        =   44
         Top             =   855
         Width           =   2280
      End
      Begin VB.Label lblDato 
         Caption         =   "Tipo de Vía :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   -74685
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
         Left            =   6225
         TabIndex        =   28
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
         Picture         =   "abcfamiliar.frx":0054
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   5835
         TabIndex        =   29
         Top             =   75
         Visible         =   0   'False
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcfamiliar.frx":0070
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
         Left            =   675
         TabIndex        =   30
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   5070
      Width           =   7200
      _Version        =   65536
      _ExtentX        =   12700
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
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   3
         Left            =   4695
         TabIndex        =   32
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
         Picture         =   "abcfamiliar.frx":008C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   33
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
         Picture         =   "abcfamiliar.frx":00A8
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   34
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcfamiliar.frx":00C4
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   35
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcfamiliar.frx":00E0
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   4500
      Index           =   0
      Left            =   6435
      TabIndex        =   36
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   7937
      _StockProps     =   15
      ForeColor       =   192
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   37
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edición"
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
         Outline         =   -1  'True
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   38
         Tag             =   "0"
         Top             =   840
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   -2147483631
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
         Picture         =   "abcfamiliar.frx":00FC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   39
         Tag             =   "0"
         Top             =   1800
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   -2147483631
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
         Picture         =   "abcfamiliar.frx":0118
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   40
         Tag             =   "0"
         Top             =   2745
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   -2147483631
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
         Picture         =   "abcfamiliar.frx":0134
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   60
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
Attribute VB_Name = "fAbcFamiliar"
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
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  For n_Index = 0 To 2
    cmdHelp(n_Index).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  Next n_Index
  cmdUbigeo(0).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)

End Sub
Sub ShowScreen()
    
  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "AT", txtNombres(0), "", Me.Tag, False, 25
    gdl_Procedure.EditText "AT", txtNombres(1), "", Me.Tag, False, 25
    gdl_Procedure.EditText "AT", txtNombres(2), "", Me.Tag, False, 25
    gdl_Procedure.EditCombo "AT", cmbSexo, -1, Me.Tag, False
    gdl_Procedure.EditDTPicker "AT", dtpFecha, Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    lblEdad = Trim(Year(Date) - Year(dtpFecha)) & " Años"
    gdl_Procedure.EditText "AT", txtTipoDocu, "", Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtDocumento, "", Me.Tag, False, 11
    gdl_Procedure.EditCombo "AT", cmbVinculo, -1, Me.Tag, False
    gdl_Procedure.EditText "AT", txtCarta, "", Me.Tag, False, 20
    gdl_Procedure.EditOptionCheck "AT", chkIncapacidad, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtCertificado, "", Me.Tag, False, 20
    gdl_Procedure.EditCombo "AT", cmbSituacion, 1, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
    
    gdl_Procedure.EditCombo "AT", cmbdocumento, -1, Me.Tag, False
    gdl_Procedure.EditText "AT", txtpaternidad, "", Me.Tag, False, 20
    gdl_Procedure.EditDTPicker "AT", dtpalta, Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditMask "AT", mskFecBaja, "", Me.Tag, True, "##/##/####"
    
    ' Primera pestaña
    gdl_Procedure.EditOptionCheck "AT", chkDomicilio, False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtTipoVia, fAbcPersonal.txtTipoVia.Text, Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtNombreVia, fAbcPersonal.txtNombreVia.Text, Me.Tag, False, 40
    gdl_Procedure.EditText "AT", txtNumero(0), fAbcPersonal.txtNumero(0).Text, Me.Tag, False, 4
    gdl_Procedure.EditText "AT", txtNumero(1), fAbcPersonal.txtNumero(1).Text, Me.Tag, False, 4
    gdl_Procedure.EditText "AT", txtTipoZona, fAbcPersonal.txtTipoZona.Text, Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtNombreZona, fAbcPersonal.txtNombreZona.Text, Me.Tag, False, 40
    gdl_Procedure.EditText "AT", txtReferencia, fAbcPersonal.txtReferencia.Text, Me.Tag, False, 50
    gdl_Procedure.EditText "AT", txtUbigeo(0), fAbcPersonal.txtUbigeo(1).Text, Me.Tag, False, 6
  Else
    gdl_Procedure.EditText "AT", txtNombres(0), fAbcPersonal.tdbFamiliar.Columns(7).Text, Me.Tag, False, 25
    gdl_Procedure.EditText "AT", txtNombres(1), fAbcPersonal.tdbFamiliar.Columns(8).Text, Me.Tag, False, 25
    gdl_Procedure.EditText "AT", txtNombres(2), fAbcPersonal.tdbFamiliar.Columns(9).Text, Me.Tag, False, 25
    gdl_Procedure.EditCombo "AT", cmbSexo, fAbcPersonal.tdbFamiliar.Columns(10).Text, Me.Tag, False
    gdl_Procedure.EditDTPicker "AT", dtpFecha, fAbcPersonal.tdbFamiliar.Columns(1).Text, Me.Tag, True, s_FormatoFecha, dtpShortDate
    lblEdad.Caption = Trim(Year(Date) - Year(dtpFecha.Value)) & " Años"
    gdl_Procedure.EditText "AT", txtTipoDocu, fAbcPersonal.tdbFamiliar.Columns(11).Text, Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtDocumento, fAbcPersonal.tdbFamiliar.Columns(3).Text, Me.Tag, False, 11
    gdl_Procedure.EditCombo "AT", cmbVinculo, CInt(fAbcPersonal.tdbFamiliar.Columns(4).Value), Me.Tag, False
    gdl_Procedure.EditText "AT", txtCarta, fAbcPersonal.tdbFamiliar.Columns(12).Text, Me.Tag, False, 20
    gdl_Procedure.EditOptionCheck "AT", chkIncapacidad, (fAbcPersonal.tdbFamiliar.Columns(22).Text = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtCertificado, fAbcPersonal.tdbFamiliar.Columns(23).Text, Me.Tag, False, 20
    gdl_Procedure.EditCombo "AT", cmbSituacion, CInt(fAbcPersonal.tdbFamiliar.Columns(5).Value), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fAbcPersonal.tdbFamiliar.Columns(24).Text = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fAbcPersonal.tdbFamiliar.Columns(24).Text = s_Estado_Blq), Me.Tag, True
    
    gdl_Procedure.EditCombo "AT", cmbdocumento, fAbcPersonal.tdbFamiliar.Columns(25).Value, Me.Tag, False
    gdl_Procedure.EditText "AT", txtpaternidad, fAbcPersonal.tdbFamiliar.Columns(26).Text, Me.Tag, False, 20
    gdl_Procedure.EditDTPicker "AT", dtpalta, fAbcPersonal.tdbFamiliar.Columns(27).Text, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditMask "AT", mskFecBaja, IIf(IsNull(fAbcPersonal.tdbFamiliar.Columns(28).Text), "", fAbcPersonal.tdbFamiliar.Columns(28).Text), Me.Tag, True, "##/##/####"
    
    ' Primera pestaña
    gdl_Procedure.EditOptionCheck "AT", chkDomicilio, (fAbcPersonal.tdbFamiliar.Columns(13).Text = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditText "AT", txtTipoVia, fAbcPersonal.tdbFamiliar.Columns(14).Text, Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtNombreVia, fAbcPersonal.tdbFamiliar.Columns(15).Text, Me.Tag, False, 40
    gdl_Procedure.EditText "AT", txtNumero(0), fAbcPersonal.tdbFamiliar.Columns(16).Text, Me.Tag, False, 4
    gdl_Procedure.EditText "AT", txtNumero(1), fAbcPersonal.tdbFamiliar.Columns(17).Text, Me.Tag, False, 4
    gdl_Procedure.EditText "AT", txtTipoZona, fAbcPersonal.tdbFamiliar.Columns(18).Text, Me.Tag, False, 2
    gdl_Procedure.EditText "AT", txtNombreZona, fAbcPersonal.tdbFamiliar.Columns(19).Text, Me.Tag, False, 40
    gdl_Procedure.EditText "AT", txtReferencia, fAbcPersonal.tdbFamiliar.Columns(20).Text, Me.Tag, False, 50
    gdl_Procedure.EditText "AT", txtUbigeo(0), fAbcPersonal.tdbFamiliar.Columns(21).Text, Me.Tag, False, 6
  End If
  lblHelp(0).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoDocu.Text, "DI")
  ' Primera pestaña
  lblHelp(1).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoVia.Text, "TV")
  lblHelp(2).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoZona.Text, "TZ")
  lblUbigeo(0).Caption = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, ps_CodEmpresa, txtUbigeo(0).Text, "UG")

End Sub
']
Private Sub cmdAction_Click(Index As Integer)
  Dim n_Secuencia As Integer
  
  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  txtNombres(0).SetFocus
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim(fAbcPersonal.txtCodigo)
    n_Secuencia = CInt(Val(fAbcPersonal.tdbFamiliar.Columns(6).Text))
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpsn", "orden")
    a_Valores = Array(ps_ClsPlanilla, s_Registro, n_Secuencia)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero)
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plfamiliares", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el ado control y la grilla
    fAbcPersonal.RecuperarFamiliares
    ' Verifico si aun existen registros
    l_ExistRecord = ((fAbcPersonal.tdbFamiliar.EOF And fAbcPersonal.tdbFamiliar.BOF) Or fAbcPersonal.tdbFamiliar.VisibleRows = 0)
    If Not l_ExistRecord Then
'      fRemunerExcepcional.dcaRegistro.Recordset.Find ("codcpc >= '" & s_ConceptoPlanilla & "'")
      If fAbcPersonal.tdbFamiliar.EOF Then fAbcPersonal.tdbFamiliar.MoveLast
    Else
      Unload Me
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  If Not l_ExistRecord Then cmdCancel_Click

End Sub
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Tipo de documento de identidad
    tdbHelp.Columns(0).DataField = "coddci": tdbHelp.Columns(1).DataField = "desdci"
    tdbHelp.Caption = "Documentos de Identidad"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("dci", "coddci", "", "")
   Case 1     ' Tipo de via de dirección
    tdbHelp.Columns(0).DataField = "codvia": tdbHelp.Columns(1).DataField = "desvia"
    tdbHelp.Caption = "Tipos de Via"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("via", "codvia", "", "")
   Case 2     ' Tipo de zona de direccion
    tdbHelp.Columns(0).DataField = "codzona": tdbHelp.Columns(1).DataField = "deszona"
    tdbHelp.Caption = "Tipos de Zona"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("zon", "codzona", "", "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (tabRegister.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)))
  tdbHelp.Left = (tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2)))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fAbcPersonal.tdbFamiliar.MoveFirst
   Case 1: If Not fAbcPersonal.tdbFamiliar.BOF Then fAbcPersonal.tdbFamiliar.MovePrevious
           If fAbcPersonal.tdbFamiliar.BOF Then fAbcPersonal.tdbFamiliar.MoveFirst
   Case 2: If Not fAbcPersonal.tdbFamiliar.EOF Then fAbcPersonal.tdbFamiliar.MoveNext
           If fAbcPersonal.tdbFamiliar.EOF Then fAbcPersonal.tdbFamiliar.MoveLast
   Case 3: fAbcPersonal.tdbFamiliar.MoveLast
  End Select

End Sub

Private Sub cmdUbigeo_Click(Index As Integer)
  Set o_SwSelUbica = fAbcFamiliar: n_SwSelUbica = Index
  fSeleccionUbigeo.Show vbModal
  Set o_SwSelUbica = Nothing
  Exit Sub
End Sub
Private Sub cmdUpdate_Click()
  Dim n_Secuencia As Integer, s_Incapacidad As String
  Dim s_domicilio As String, s_Motivo As String
  
  ' Realizo las validaciones de los campos a actualizar
  If Trim(txtNombres(0)) = "" And Trim(txtNombres(1)) = "" And Trim(txtNombres(2)) = "" Then Beep: MsgBox "Debe Ingresar los nombres " & lblTitle, vbExclamation: txtNombres(0).SetFocus: Exit Sub
  If cmbSexo = "" Then Beep: MsgBox "Seleccione el sexo " & lblTitle, vbExclamation: cmbSexo.SetFocus: Exit Sub
  If txtTipoDocu = "" Then Beep: MsgBox "Debe Ingresar el tipo de documento " & lblTitle, vbExclamation: txtTipoDocu.SetFocus: Exit Sub
  If (lblHelp(0) = "" Or lblHelp(0) = "???") Then Beep: MsgBox "Tipo Documento Identidad no es valido; Verificar", vbExclamation: txtTipoDocu.SetFocus: Exit Sub
  If txtDocumento = "" Then Beep: MsgBox "Debe Ingresar el documento de identidad " & lblTitle, vbExclamation: txtDocumento.SetFocus: Exit Sub
  If cmbVinculo = "" Then Beep: MsgBox "Seleccione el vinculo " & lblTitle, vbExclamation: cmbVinculo.SetFocus: Exit Sub
  If chkIncapacidad.Value And txtCertificado = "" Then Beep: MsgBox "Debe Ingresar el certificado médico " & lblTitle, vbExclamation: txtCertificado.SetFocus: Exit Sub
  If cmbVinculo.ListIndex = 4 And txtCarta = "" Then Beep: MsgBox "Debe Ingresar el carta de atención médica " & lblTitle, vbExclamation: txtCarta.SetFocus: Exit Sub
  If cmbSituacion = "" Then Beep: MsgBox "Seleccione la situación " & lblTitle, vbExclamation: cmbSituacion.SetFocus: Exit Sub

  ' Primera pestaña
  If lblHelp(1) = "???" Then Beep: MsgBox "Dirección - Tipo de Via no es valido; Verificar", vbExclamation: txtTipoVia.SetFocus: Exit Sub
  If lblHelp(2) = "???" Then Beep: MsgBox "Dirección - Tipo de Zona no es valido; Verificar", vbExclamation: txtTipoZona.SetFocus: Exit Sub
  If lblUbigeo(0) = "???" Then Beep: MsgBox "Dirección - Ubicacion Geografica no es valido; Verificar", vbExclamation: txtUbigeo(0).SetFocus: Exit Sub

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = Trim(fAbcPersonal.txtCodigo)
  s_Incapacidad = IIf(chkIncapacidad.Value, s_Estado_Act, s_Estado_Ina)
  s_domicilio = IIf(chkDomicilio.Value, s_Estado_Act, s_Estado_Ina)
  s_Motivo = IIf(optEstado(0).Value, s_Estado_Act, IIf(optEstado(1).Value, s_Estado_Blq, s_Estado_Ina))
  
  'txtCertificado = IIf(s_Incapacidad = s_Estado_Ina, "", txtCertificado)
  'txtCarta = IIf(cmbVinculo.ListIndex <> 4, "", txtCarta)
  
  ' Obtengo el orden correlativo
  n_Secuencia = CInt(Val(fAbcPersonal.tdbFamiliar.Columns(6).Text))
  If Me.Tag = s_MdoData_Ins Then
    s_Sql = "SELECT IFNULL(MAX(orden), 0)+1 AS registro "
    s_Sql = s_Sql & "FROM plfamiliares "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codpsn='" & s_Registro & "' "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
    n_Secuencia = CInt(porstRecordset!registro)
  End If

  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpsn", "orden", "apepaterno", "apematerno", "nombres", "fecnacimiento", "sexofam", "coddci", "numdociden", "vinculo", "cartamed", "domicilio", "codvia", "nomviadom", "numerdom", "intedom", "codzona", "nomzonadom", "refedom", "ubigeodom", "incapacidad", "certificadomed", "motivoina", "estadofam", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"), "tipdocpaternidad", "acrepaternidad", "fecalta", "fecbaja")
  a_Valores = Array(ps_ClsPlanilla, s_Registro, n_Secuencia, Trim(txtNombres(0).Text), Trim(txtNombres(1).Text), Trim(txtNombres(2).Text), Format(dtpFecha, s_FmtFechMysql_0), cmbSexo.ListIndex, Trim(txtTipoDocu.Text), Trim(txtDocumento.Text), cmbVinculo.ListIndex, Trim(txtCarta.Text), s_domicilio, Trim(txtTipoVia.Text), Trim(txtNombreVia.Text), Trim(txtNumero(0).Text), Trim(txtNumero(1).Text), Trim(txtTipoZona.Text), Trim(txtNombreZona.Text), Trim(txtReferencia.Text), Trim(txtUbigeo(0).Text), s_Incapacidad, Trim(txtCertificado.Text), s_Motivo, cmbSituacion.ListIndex, ps_Usuario, Format(Now, s_FmtFeHoMysql_0), cmbdocumento.ListIndex, Trim(txtpaternidad), Format(dtpalta, s_FmtFechMysql_0), Format(mskFecBaja, s_FmtFechMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA)
  a_Where = Array("codcls", "codpsn", "orden")

  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)

  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plfamiliares", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plfamiliares", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción

  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  fAbcPersonal.RecuperarFamiliares
  ' Ubico el registro ingresado o actualizado
  'fAbcPersonal.plexpelaboral.f dcaRegistro.Recordset.Find ("codcpc='" & s_ConceptoPlanilla & "'")
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtNombres(0).SetFocus
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
Private Sub dtpFecha_LostFocus()
  lblEdad = Trim(Year(Date) - Year(dtpFecha)) & " Años"
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 6060: Me.Width = 7290
  Me.Left = 3980: Me.Top = 750
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Datos Familiares"
  lblTitle = "Dato Familiar"
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operación del registro
  Me.Tag = fAbcPersonal.tdbFamiliar.Tag
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualización
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fAbcPersonal.tdbFamiliar.EOF Or fAbcPersonal.tdbFamiliar.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fAbcPersonal.tdbFamiliar.Columns(0).Text
  
  ' Configuro los listados, datos adicionales
  For n_Index = 0 To 1: cmbSexo.AddItem Choose(n_Index + 1, "Masculino", "Femenino"): Next n_Index
  For n_Index = 0 To 4: cmbVinculo.AddItem Choose(n_Index + 1, "Otro", "Hijo", "Conyuge", "Concubina(o)", "Gestante"): Next n_Index
  For n_Index = 0 To 1: cmbSituacion.AddItem Choose(n_Index + 1, "Inactivo", "Activo"): Next n_Index
  For n_Index = 0 To 2: cmbdocumento.AddItem Choose(n_Index + 1, "Escritura Publica", "Testamento", "Sentencia de Declaratoria"): Next n_Index
   
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcgo", "descgo")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Cargos de Personal", 2
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' Habilito/desabilito botones inciales
  fAbcPersonal.cmdActionExp(0).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActionExp(1).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActionExp(2).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
End Sub

Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Tipo de documento de identidad
    txtTipoDocu = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipoDocu.SetFocus
   Case 1       ' Tipo de via
    txtTipoVia = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipoVia.SetFocus
   Case 2       ' Tipo de zona
    txtTipoZona = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtTipoZona.SetFocus
  End Select
  
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0  ' Tipo de documento de identidad
    s_Sql = gdl_Funcion.HelpTablas("dci", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 1  ' Tipo de via
    s_Sql = gdl_Funcion.HelpTablas("via", tdbHelp.Columns(ColIndex).DataField, "", "")
   Case 2  ' Tipo de zona
    s_Sql = gdl_Funcion.HelpTablas("zon", tdbHelp.Columns(ColIndex).DataField, "", "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
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
Private Sub txtCarta_GotFocus()
  gdl_Procedure.MarcaGet txtCarta
End Sub
Private Sub txtCarta_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCertificado_GotFocus()
  gdl_Procedure.MarcaGet txtCertificado
End Sub
Private Sub txtCertificado_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtDocumento_GotFocus()
  gdl_Procedure.MarcaGet txtDocumento
End Sub
Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtNombres_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtNombres(Index)
End Sub
Private Sub txtNombres_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
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
Private Sub txtnumero_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtNumero(Index)
End Sub
Private Sub txtnumero_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtReferencia_GotFocus()
  gdl_Procedure.MarcaGet txtReferencia
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoDocu_GotFocus()
  gdl_Procedure.MarcaGet txtTipoDocu
End Sub
Private Sub txtTipoDocu_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtTipoDocu_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoDocu_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoDocu, "DI")
End Sub
Private Sub txtTipoVia_GotFocus()
  gdl_Procedure.MarcaGet txtTipoVia
End Sub
Private Sub txtTipoVia_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 1
End Sub
Private Sub txtTipoVia_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoVia_LostFocus()
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoVia, "TV")
End Sub
Private Sub txtTipoZona_GotFocus()
  gdl_Procedure.MarcaGet txtTipoZona
End Sub
Private Sub txtTipoZona_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 2
End Sub
Private Sub txtTipoZona_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtTipoZona_LostFocus()
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtTipoZona, "TZ")
End Sub
Private Sub txtUbigeo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtUbigeo(Index)
End Sub
Private Sub txtUbigeo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdUbigeo_Click 0
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
