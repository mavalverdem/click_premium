VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcCuentaProvision 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4905
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7200
   Icon            =   "abccuentapvs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   7200
   Begin TabDlg.SSTab tabRegister 
      Height          =   3720
      Left            =   75
      TabIndex        =   44
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   6562
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
      TabCaption(0)   =   "Vacaci�n"
      TabPicture(0)   =   "abccuentapvs.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmCuadro(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Gratificaci�n"
      TabPicture(1)   =   "abccuentapvs.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmCuadro(1)"
      Tab(1).ControlCount=   1
      Begin Threed.SSFrame frmCuadro 
         Height          =   3060
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   195
         Width           =   5865
         _Version        =   65536
         _ExtentX        =   10345
         _ExtentY        =   5397
         _StockProps     =   14
         Caption         =   " Cuentas Contables  "
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
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   2
            Left            =   165
            TabIndex        =   10
            Top             =   2070
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   3
            Left            =   165
            TabIndex        =   13
            Top             =   2625
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   0
            Left            =   165
            TabIndex        =   3
            Top             =   675
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   1
            Left            =   165
            TabIndex        =   6
            Top             =   1230
            Width           =   1200
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   0
            Left            =   1425
            TabIndex        =   45
            Top             =   675
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
            Index           =   1
            Left            =   1425
            TabIndex        =   46
            Top             =   1230
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
            Left            =   1425
            TabIndex        =   47
            Top             =   2070
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
            Index           =   3
            Left            =   1425
            TabIndex        =   48
            Top             =   2625
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin VB.Label lblDato 
            Caption         =   "Moneda Extranjera"
            ForeColor       =   &H00C00000&
            Height          =   165
            Index           =   3
            Left            =   180
            TabIndex        =   8
            Top             =   1620
            Width           =   1650
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Debe :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   180
            TabIndex        =   9
            Top             =   1845
            Width           =   795
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Haber :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   180
            TabIndex        =   12
            Top             =   2385
            Width           =   795
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
            Height          =   180
            Index           =   2
            Left            =   1785
            TabIndex        =   11
            Top             =   2115
            Width           =   180
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
            Height          =   180
            Index           =   3
            Left            =   1785
            TabIndex        =   14
            Top             =   2670
            Width           =   180
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   1170
            Index           =   1
            Left            =   45
            Shape           =   4  'Rounded Rectangle
            Top             =   1815
            Width           =   5790
         End
         Begin VB.Label lblDato 
            Caption         =   "Moneda Nacional"
            ForeColor       =   &H00C00000&
            Height          =   165
            Index           =   0
            Left            =   180
            TabIndex        =   1
            Top             =   225
            Width           =   1650
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Debe :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   2
            Top             =   435
            Width           =   795
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Haber :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   180
            TabIndex        =   5
            Top             =   990
            Width           =   795
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
            Height          =   180
            Index           =   0
            Left            =   1785
            TabIndex        =   4
            Top             =   720
            Width           =   180
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
            Height          =   180
            Index           =   1
            Left            =   1785
            TabIndex        =   7
            Top             =   1275
            Width           =   180
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   1170
            Index           =   0
            Left            =   45
            Shape           =   4  'Rounded Rectangle
            Top             =   420
            Width           =   5790
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   3060
         Index           =   1
         Left            =   -74820
         TabIndex        =   15
         Top             =   195
         Width           =   5865
         _Version        =   65536
         _ExtentX        =   10345
         _ExtentY        =   5397
         _StockProps     =   14
         Caption         =   " Cuentas Contables  "
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
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   5
            Left            =   165
            TabIndex        =   21
            Top             =   1230
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   4
            Left            =   165
            TabIndex        =   18
            Top             =   675
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   7
            Left            =   165
            TabIndex        =   28
            Top             =   2625
            Width           =   1200
         End
         Begin VB.TextBox txtCuenta 
            ForeColor       =   &H00800000&
            Height          =   280
            Index           =   6
            Left            =   165
            TabIndex        =   25
            Top             =   2070
            Width           =   1200
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   4
            Left            =   1425
            TabIndex        =   50
            Top             =   675
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
            Index           =   5
            Left            =   1425
            TabIndex        =   51
            Top             =   1230
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
            Index           =   6
            Left            =   1425
            TabIndex        =   52
            Top             =   2070
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
            Index           =   7
            Left            =   1425
            TabIndex        =   53
            Top             =   2625
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Haber :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   180
            TabIndex        =   20
            Top             =   990
            Width           =   795
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Debe :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   180
            TabIndex        =   17
            Top             =   435
            Width           =   795
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Haber :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   180
            TabIndex        =   27
            Top             =   2385
            Width           =   795
         End
         Begin VB.Label lblDato 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Debe :"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   180
            TabIndex        =   24
            Top             =   1845
            Width           =   795
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
            Height          =   180
            Index           =   5
            Left            =   1785
            TabIndex        =   22
            Top             =   1275
            Width           =   180
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
            Height          =   180
            Index           =   4
            Left            =   1785
            TabIndex        =   19
            Top             =   720
            Width           =   180
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
            Height          =   180
            Index           =   7
            Left            =   1785
            TabIndex        =   29
            Top             =   2670
            Width           =   180
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
            Height          =   180
            Index           =   6
            Left            =   1785
            TabIndex        =   26
            Top             =   2115
            Width           =   180
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   1170
            Index           =   2
            Left            =   45
            Shape           =   4  'Rounded Rectangle
            Top             =   420
            Width           =   5790
         End
         Begin VB.Label lblDato 
            Caption         =   "Moneda Nacional"
            ForeColor       =   &H00C00000&
            Height          =   165
            Index           =   6
            Left            =   180
            TabIndex        =   16
            Top             =   225
            Width           =   1650
         End
         Begin VB.Shape shpCuadro 
            BorderColor     =   &H00C00000&
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   1170
            Index           =   3
            Left            =   45
            Shape           =   4  'Rounded Rectangle
            Top             =   1815
            Width           =   5790
         End
         Begin VB.Label lblDato 
            Caption         =   "Moneda Extranjera"
            ForeColor       =   &H00C00000&
            Height          =   165
            Index           =   9
            Left            =   180
            TabIndex        =   23
            Top             =   1620
            Width           =   1650
         End
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   30
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
         Left            =   6450
         TabIndex        =   31
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
         Picture         =   "abccuentapvs.frx":0044
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   32
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
         Picture         =   "abccuentapvs.frx":0060
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
         TabIndex        =   33
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   34
      Top             =   4395
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
         TabIndex        =   35
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
         Picture         =   "abccuentapvs.frx":007C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   36
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
         Picture         =   "abccuentapvs.frx":0098
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   37
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
         Picture         =   "abccuentapvs.frx":00B4
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   38
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
         Picture         =   "abccuentapvs.frx":00D0
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   3720
      Index           =   0
      Left            =   6435
      TabIndex        =   39
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   6562
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
         TabIndex        =   40
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edici�n"
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
         TabIndex        =   41
         Tag             =   "0"
         Top             =   810
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
         Picture         =   "abccuentapvs.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   42
         Tag             =   "0"
         Top             =   1440
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
         Picture         =   "abccuentapvs.frx":0108
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   43
         Tag             =   "0"
         Top             =   2040
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
         Picture         =   "abccuentapvs.frx":0124
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1695
      TabIndex        =   49
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
Attribute VB_Name = "fAbcCuentaProvision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificaci�n de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_Registro As String                            ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acci�n
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creaci�n", IIf(Me.Tag = s_MdoData_Del, " - Eliminaci�n", IIf(Me.Tag = s_MdoData_Upd, " - Actualizaci�n", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  For n_Index = 0 To 7: cmdHelp(n_Index).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd): Next n_Index

End Sub
Sub ShowScreen()
    
  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "AT", txtCuenta(0), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_debmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_habmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(2), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_debme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(3), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_habme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(4), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_debmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(5), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_habmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(6), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_debme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(7), "", Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_habme.DefinedSize
  Else
    gdl_Procedure.EditText "AT", txtCuenta(0), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctav_debmn), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_debmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctav_habmn), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_habmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(2), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctav_debme), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_debme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(3), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctav_habme), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctav_habme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(4), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctag_debmn), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_debmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(5), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctag_habmn), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_habmn.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(6), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctag_debme), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_debme.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(7), gdl_Funcion.aTexto(fCuentaProvision.dcaRegistro.Recordset!codctag_habme), Me.Tag, False, fCuentaProvision.dcaRegistro.Recordset!codctag_habme.DefinedSize
  End If
  For n_Index = 0 To 7
    lblHelp(n_Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, s_Estado_Act, txtCuenta(n_Index), "CU")
  Next n_Index

End Sub
']
Private Sub cmdAction_Click(Index As Integer)
  Dim n_Secuencia As Integer
  
  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  txtCuenta(0).SetFocus
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("� Est�s Seguro de Eliminar el " & lblTitle & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    n_Secuencia = CInt(fCuentaProvision.dcaRegistro.Recordset!orden)
    s_Registro = Trim(n_Secuencia)
    
    '[ Inicio la conexi�n a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codcco", "orden")
    a_Valores = Array(ps_ClsPlanilla, Trim(fCentroCosto.dcaRegistro.Recordset!codcco), n_Secuencia)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacci�n
    ' Elimino el registro
    If Not Records_Del("plctapvs", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacci�n
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fCuentaProvision.dcaRegistro, fCuentaProvision.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fCuentaProvision.dcaRegistro.Recordset.EOF And fCuentaProvision.dcaRegistro.Recordset.BOF) Or fCuentaProvision.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fCuentaProvision.dcaRegistro.Recordset.Find ("orden >= '" & s_Registro & "'")
      If fCuentaProvision.dcaRegistro.Recordset.EOF Then fCuentaProvision.dcaRegistro.Recordset.MoveLast
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
  '[ Finalizo la conexi�n a la base de datos ]
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
   Case 0 To 7  ' Cuenta contable
    tdbHelp.Columns(0).DataField = "codcta": tdbHelp.Columns(1).DataField = "detcta"
    tdbHelp.Caption = "Cuenta Contable"
    s_Sql = gdl_Funcion.HelpTablas("cta", tdbHelp.Columns(0).DataField, ps_CodEmpresa, "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (tabRegister.Top + (IIf(Index = 0 Or Index = 1 Or Index = 4 Or Index = 5, (cmdHelp(Index).Top + frmCuadro(tabRegister.Tab).Top), 1680) + (cmdHelp(Index).Height / 2)))
  tdbHelp.Left = frmCuadro(tabRegister.Tab).Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fCuentaProvision.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fCuentaProvision.dcaRegistro.Recordset.BOF Then fCuentaProvision.dcaRegistro.Recordset.MovePrevious
           If fCuentaProvision.dcaRegistro.Recordset.BOF Then fCuentaProvision.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fCuentaProvision.dcaRegistro.Recordset.EOF Then fCuentaProvision.dcaRegistro.Recordset.MoveNext
           If fCuentaProvision.dcaRegistro.Recordset.EOF Then fCuentaProvision.dcaRegistro.Recordset.MoveLast
   Case 3: fCuentaProvision.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim n_Secuencia As Integer
  
  ' Realizo las validaciones de los campos a actualizar
  If txtCuenta(0) = "" Then Beep: MsgBox "Ingresar Cuenta contable vacaci�n debe Moneda Nacional", vbExclamation: txtCuenta(0).SetFocus: Exit Sub
  If lblHelp(0) = "???" Then Beep: MsgBox "Cuenta contable vacaci�n debe Moneda Nacional no es valido; Verificar", vbExclamation: txtCuenta(0).SetFocus: Exit Sub
  If txtCuenta(1) = "" Then Beep: MsgBox "Ingresar Cuenta contable vacaci�n haber Moneda Nacional", vbExclamation: txtCuenta(1).SetFocus: Exit Sub
  If lblHelp(1) = "???" Then Beep: MsgBox "Cuenta contable vacaci�n haber Moneda Nacional no es valido; Verificar", vbExclamation: txtCuenta(1).SetFocus: Exit Sub
  If txtCuenta(2) = "" Then Beep: MsgBox "Ingresar Cuenta contable vacaci�n debe Moneda Extranjera", vbExclamation: txtCuenta(2).SetFocus: Exit Sub
  If lblHelp(2) = "???" Then Beep: MsgBox "Cuenta contable vacaci�n debe Moneda Extranjera no es valido; Verificar", vbExclamation: txtCuenta(2).SetFocus: Exit Sub
  If txtCuenta(3) = "" Then Beep: MsgBox "Ingresar Cuenta contable vacaci�n haber Moneda Extranjera", vbExclamation: txtCuenta(3).SetFocus: Exit Sub
  If lblHelp(3) = "???" Then Beep: MsgBox "Cuenta contable vacaci�n haber Moneda Extranjera no es valido; Verificar", vbExclamation: txtCuenta(3).SetFocus: Exit Sub
  If txtCuenta(4) = "" Then Beep: MsgBox "Ingresar Cuenta contable gratificaci�n debe Moneda Nacional", vbExclamation: txtCuenta(4).SetFocus: Exit Sub
  If lblHelp(4) = "???" Then Beep: MsgBox "Cuenta contable gratificaci�n debe Moneda Nacional no es valido; Verificar", vbExclamation: txtCuenta(4).SetFocus: Exit Sub
  If txtCuenta(5) = "" Then Beep: MsgBox "Ingresar Cuenta contable gratificaci�n haber Moneda Nacional", vbExclamation: txtCuenta(5).SetFocus: Exit Sub
  If lblHelp(5) = "???" Then Beep: MsgBox "Cuenta contable gratificaci�n haber Moneda Nacional no es valido; Verificar", vbExclamation: txtCuenta(5).SetFocus: Exit Sub
  If txtCuenta(6) = "" Then Beep: MsgBox "Ingresar Cuenta contable gratificaci�n debe Moneda Extranjera", vbExclamation: txtCuenta(6).SetFocus: Exit Sub
  If lblHelp(6) = "???" Then Beep: MsgBox "Cuenta contable gratificaci�n debe Moneda Extranjera no es valido; Verificar", vbExclamation: txtCuenta(6).SetFocus: Exit Sub
  If txtCuenta(7) = "" Then Beep: MsgBox "Ingresar Cuenta contable gratificaci�n haber Moneda Extranjera", vbExclamation: txtCuenta(7).SetFocus: Exit Sub
  If lblHelp(7) = "???" Then Beep: MsgBox "Cuenta contable gratificaci�n haber Moneda Extranjera no es valido; Verificar", vbExclamation: txtCuenta(7).SetFocus: Exit Sub
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el orden correlativo
  If Me.Tag = s_MdoData_Upd Then
    n_Secuencia = CInt(fCuentaProvision.dcaRegistro.Recordset!orden)
  End If
  If Me.Tag = s_MdoData_Ins Then
    s_Sql = "SELECT IFNULL(MAX(orden), 0)+1 AS registro "
    s_Sql = s_Sql & "FROM plctapvs "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codcco='" & Trim(fCentroCosto.dcaRegistro.Recordset!codcco) & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
    n_Secuencia = CInt(porstRecordset!registro)
  End If
  ' Capturo el registro a actualizar
  s_Registro = Trim(n_Secuencia)
      
  ' Creo los arreglos para la actualizaci�n
  a_Campos = Array("codcls", "codcco", "orden", "codctav_debmn", "codctav_habmn", "codctav_debme", "codctav_habme", "codctag_debmn", "codctag_habmn", "codctag_debme", "codctag_habme", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, Trim(fCentroCosto.dcaRegistro.Recordset!codcco), n_Secuencia, Trim(txtCuenta(0).Text), Trim(txtCuenta(1).Text), Trim(txtCuenta(2).Text), Trim(txtCuenta(3).Text), Trim(txtCuenta(4).Text), Trim(txtCuenta(5).Text), Trim(txtCuenta(6).Text), Trim(txtCuenta(7).Text), ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codcco", "orden")
  
  '[ Inicio la conexi�n a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacci�n
  ' Realizo el proceso de actualizaci�n de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plctapvs", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plctapvs", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacci�n
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fCuentaProvision.dcaRegistro, fCuentaProvision.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fCuentaProvision.dcaRegistro.Recordset.Find ("orden='" & s_Registro & "'")
  ' si es actualizaci�n pasa al modo visualizaci�n
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtCuenta(0).SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexi�n a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminaci�n
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posici�n y titulo del formulario
  Me.Height = 5380: Me.Width = 7290
  Me.Left = 3000: Me.Top = 500
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualizaci�n Cuentas Contables Concepto"
  lblTitle = "Cuentas Contables"
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operaci�n del registro
  Me.Tag = fCuentaProvision.Tag
  
  ' Configuro parametros de visualizaci�n del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y t�tulo del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "A�adir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualizaci�n del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y t�tulo del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualizaci�n
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Informaci�n de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Informaci�n de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fCuentaProvision.dcaRegistro.Recordset.EOF Or fCuentaProvision.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fCuentaProvision.dcaRegistro.Recordset!orden
  
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuraci�n de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "C�digo", "Descripci�n")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de C�lculo", 2
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

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
   Case 0 To 7  ' Cuenta contable
    txtCuenta(n_IndexHelp) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCuenta(n_IndexHelp).SetFocus
  End Select
  
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la informaci�n ordenada
  Select Case n_IndexHelp
   Case 0 To 7  ' Cuenta contable
    s_Sql = gdl_Funcion.HelpTablas("cta", tdbHelp.Columns(ColIndex).DataField, ps_CodEmpresa, "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
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
Private Sub txtCuenta_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCuenta(Index)
End Sub
Private Sub txtCuenta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 2
End Sub
Private Sub txtCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCuenta_LostFocus(Index As Integer)
  lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, s_Estado_Act, txtCuenta(Index), "CU")
End Sub
