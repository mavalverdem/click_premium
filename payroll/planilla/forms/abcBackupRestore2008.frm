VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form FBackupRestore2008 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSFrame sfmProgreso 
      Height          =   480
      Left            =   75
      TabIndex        =   0
      Top             =   4200
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
      _ExtentY        =   847
      _StockProps     =   14
      Caption         =   " Procesando archivo : "
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
      Font3D          =   3
      ShadowStyle     =   1
      Begin MSComctlLib.ProgressBar pgbProgreso 
         Height          =   225
         Left            =   45
         TabIndex        =   1
         Top             =   225
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   3540
      Left            =   75
      TabIndex        =   2
      Top             =   600
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   6244
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
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
      TabCaption(0)   =   "Backup/Restore"
      TabPicture(0)   =   "abcBackupRestore2008.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmCuadro(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmCuadro(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin Threed.SSFrame frmCuadro 
         Height          =   1770
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   3122
         _StockProps     =   14
         Caption         =   " Ubicación "
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
         Begin VB.DriveListBox drive 
            Height          =   315
            Left            =   140
            TabIndex        =   5
            Top             =   495
            Width           =   2240
         End
         Begin VB.DirListBox dir 
            Height          =   1215
            Left            =   2520
            TabIndex        =   4
            Top             =   480
            Width           =   4035
         End
         Begin VB.Label lblDato 
            Caption         =   "Directorio :"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   140
            TabIndex        =   6
            Top             =   250
            Width           =   1005
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1095
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   6810
         _Version        =   65536
         _ExtentX        =   12012
         _ExtentY        =   1931
         _StockProps     =   14
         Caption         =   " Opción "
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
         Begin Threed.SSOption optParametro 
            Height          =   200
            Index           =   0
            Left            =   230
            TabIndex        =   8
            Top             =   285
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Backup"
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
         Begin Threed.SSOption optParametro 
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   9
            Top             =   525
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Restore"
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
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7275
      _Version        =   65536
      _ExtentX        =   12832
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
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   6435
         TabIndex        =   11
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
         Picture         =   "abcBackupRestore2008.frx":001C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6000
         TabIndex        =   12
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcBackupRestore2008.frx":0038
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
         Left            =   450
         TabIndex        =   13
         Top             =   120
         Width           =   5085
      End
   End
End
Attribute VB_Name = "FBackupRestore2008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
