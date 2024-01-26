VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fPrmExpInformacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5190
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   8415
   Icon            =   "prmexpinforma.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8415
   Begin TabDlg.SSTab tabRegister 
      Height          =   3975
      Left            =   75
      TabIndex        =   47
      Top             =   600
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   7011
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
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
      TabCaption(0)   =   "Parametros"
      TabPicture(0)   =   "prmexpinforma.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmCuadro(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmCuadro(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin Threed.SSFrame frmCuadro 
         Height          =   3420
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   135
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   6032
         _StockProps     =   14
         Caption         =   " Remuneración Afecta "
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
            Index           =   4
            Left            =   180
            TabIndex        =   14
            Top             =   3015
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
            Index           =   0
            Left            =   180
            TabIndex        =   2
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
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   1155
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
            Index           =   2
            Left            =   180
            TabIndex        =   8
            Top             =   1785
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
            Index           =   3
            Left            =   180
            TabIndex        =   11
            Top             =   2400
            Width           =   980
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   0
            Left            =   1215
            TabIndex        =   37
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
            Index           =   1
            Left            =   1215
            TabIndex        =   38
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
            Left            =   1215
            TabIndex        =   39
            Top             =   1785
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   3
            Left            =   1215
            TabIndex        =   40
            Top             =   2400
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   4
            Left            =   1215
            TabIndex        =   41
            Top             =   3015
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblDato 
            Caption         =   "5ta Categoría :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   13
            Top             =   2760
            Width           =   2730
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
            Left            =   1575
            TabIndex        =   15
            Top             =   3060
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
            Height          =   180
            Index           =   0
            Left            =   1575
            TabIndex        =   3
            Top             =   585
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "IES :"
            ForeColor       =   &H00C00000&
            Height          =   190
            Index           =   0
            Left            =   180
            TabIndex        =   1
            Top             =   285
            Width           =   2730
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
            Left            =   1575
            TabIndex        =   6
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Pensiones (ONP) :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   4
            Top             =   900
            Width           =   2730
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
            Left            =   1575
            TabIndex        =   9
            Top             =   1830
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Salud (EsSalud) :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   7
            Top             =   1530
            Width           =   2730
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
            Left            =   1575
            TabIndex        =   12
            Top             =   2445
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Fondo Social del Artista :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   10
            Top             =   2145
            Width           =   2730
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   3420
         Index           =   1
         Left            =   4155
         TabIndex        =   16
         Top             =   135
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   6032
         _StockProps     =   14
         Caption         =   " Descuento y/o Aporte "
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
         Begin VB.TextBox txtRetencion 
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
            Index           =   4
            Left            =   180
            TabIndex        =   30
            Top             =   3015
            Width           =   980
         End
         Begin VB.TextBox txtRetencion 
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
            Index           =   3
            Left            =   180
            TabIndex        =   27
            Top             =   2400
            Width           =   980
         End
         Begin VB.TextBox txtRetencion 
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
            Index           =   2
            Left            =   180
            TabIndex        =   24
            Top             =   1785
            Width           =   980
         End
         Begin VB.TextBox txtRetencion 
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
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   540
            Width           =   980
         End
         Begin VB.TextBox txtRetencion 
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
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   1155
            Width           =   980
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   5
            Left            =   1215
            TabIndex        =   42
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
            Index           =   6
            Left            =   1215
            TabIndex        =   43
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
            Index           =   7
            Left            =   1215
            TabIndex        =   44
            Top             =   1785
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   8
            Left            =   1215
            TabIndex        =   45
            Top             =   2400
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
            Left            =   1215
            TabIndex        =   46
            Top             =   3015
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
         End
         Begin VB.Label lblDato 
            Caption         =   "5ta Categoría :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   29
            Top             =   2760
            Width           =   2730
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
            Index           =   9
            Left            =   1575
            TabIndex        =   31
            Top             =   3060
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Fondo Social del Artista :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   26
            Top             =   2145
            Width           =   2730
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
            Index           =   8
            Left            =   1575
            TabIndex        =   28
            Top             =   2445
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Salud (EsSalud) :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   23
            Top             =   1530
            Width           =   2730
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
            Left            =   1575
            TabIndex        =   25
            Top             =   1830
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
            Height          =   180
            Index           =   5
            Left            =   1575
            TabIndex        =   19
            Top             =   585
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "IES :"
            ForeColor       =   &H00C00000&
            Height          =   190
            Index           =   5
            Left            =   180
            TabIndex        =   17
            Top             =   285
            Width           =   2730
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
            Left            =   1575
            TabIndex        =   22
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Pensiones (ONP) :"
            ForeColor       =   &H00C00000&
            Height          =   190
            Index           =   6
            Left            =   180
            TabIndex        =   20
            Top             =   900
            Width           =   2730
         End
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
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
         Left            =   7560
         TabIndex        =   34
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
         Picture         =   "prmexpinforma.frx":0028
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Index           =   0
         Left            =   7170
         TabIndex        =   35
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "prmexpinforma.frx":0044
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
         Left            =   510
         TabIndex        =   33
         Top             =   120
         Width           =   6000
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   36
      Top             =   4680
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
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
      TabIndex        =   48
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
Attribute VB_Name = "fPrmExpInformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_ModoCfg As String                             ' Modo de actualización de los datos
'[
Sub ShowScreen()
    
  ' Pestaña de Información de configuración
  s_Sql = "SELECT cpcremuies, cpcremuonp, cpcremuessalud, cpcremuartista, cpcremuquinta, "
  s_Sql = s_Sql & "cpcies, cpconp, cpcessalud, cpcartista, cpcquinta "
  s_Sql = s_Sql & "FROM plparametroafp "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.BOF) Then
    s_ModoCfg = s_MdoData_Upd
    gdl_Procedure.EditText "AT", txtRemunera(0), gdl_Funcion.aTexto(porstRecordset!cpcremuies), s_ModoCfg, False, porstRecordset!cpcremuies.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(1), gdl_Funcion.aTexto(porstRecordset!cpcremuonp), s_ModoCfg, False, porstRecordset!cpcremuonp.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(2), gdl_Funcion.aTexto(porstRecordset!cpcremuessalud), s_ModoCfg, False, porstRecordset!cpcremuessalud.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(3), gdl_Funcion.aTexto(porstRecordset!cpcremuartista), s_ModoCfg, False, porstRecordset!cpcremuartista.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(4), gdl_Funcion.aTexto(porstRecordset!cpcremuquinta), s_ModoCfg, False, porstRecordset!cpcremuquinta.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(0), gdl_Funcion.aTexto(porstRecordset!cpcies), s_ModoCfg, False, porstRecordset!cpcies.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(1), gdl_Funcion.aTexto(porstRecordset!cpconp), s_ModoCfg, False, porstRecordset!cpconp.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(2), gdl_Funcion.aTexto(porstRecordset!cpcessalud), s_ModoCfg, False, porstRecordset!cpcessalud.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(3), gdl_Funcion.aTexto(porstRecordset!cpcartista), s_ModoCfg, False, porstRecordset!cpcartista.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(4), gdl_Funcion.aTexto(porstRecordset!cpcquinta), s_ModoCfg, False, porstRecordset!cpcquinta.DefinedSize
  Else
    s_ModoCfg = s_MdoData_Ins
    gdl_Procedure.EditText "AT", txtRemunera(0), "", s_ModoCfg, False, porstRecordset!cpcremuies.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(1), "", s_ModoCfg, False, porstRecordset!cpcremuonp.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(2), "", s_ModoCfg, False, porstRecordset!cpcremuessalud.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(3), "", s_ModoCfg, False, porstRecordset!cpcremuartista.DefinedSize
    gdl_Procedure.EditText "AT", txtRemunera(4), "", s_ModoCfg, False, porstRecordset!cpcremuquinta.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(0), "", s_ModoCfg, False, porstRecordset!cpcies.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(1), "", s_ModoCfg, False, porstRecordset!cpconp.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(2), "", s_ModoCfg, False, porstRecordset!cpcessalud.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(3), "", s_ModoCfg, False, porstRecordset!cpcartista.DefinedSize
    gdl_Procedure.EditText "AT", txtRetencion(4), "", s_ModoCfg, False, porstRecordset!cpcquinta.DefinedSize
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(0), "CP")
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(1), "CP")
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(2), "CP")
  lblHelp(3) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(3), "CP")
  lblHelp(4) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(4), "CP")
  lblHelp(5) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(0), "CP")
  lblHelp(6) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(1), "CP")
  lblHelp(7) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(2), "CP")
  lblHelp(8) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(3), "CP")
  lblHelp(9) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(4), "CP")

End Sub
']
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim nTop As Double, nLeft As Double
  
  nTop = (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  nLeft = (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  s_SqlHelp = ""
  Select Case Index
   Case 0, 1, 2, 3, 4 ' Concepto de remuneración afecta
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos de Remuneración"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpc", "codcpc", s_Estado_Ina, "")
    nTop = frmCuadro(0).Top + Choose(Index + 1, nTop, nTop, 1700, 1750, 1800)
    nLeft = frmCuadro(0).Left + nLeft
   Case 5, 6, 9        ' Conceptos de descuento de pensiones
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Descuento"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpc", "codcpc", s_Estado_Act, "")
    nTop = frmCuadro(1).Top + Choose(Index - 4, nTop, nTop, 1700, 1750, 1800)
    nLeft = (frmCuadro(1).Left - 2000) + nLeft
   Case 7, 8          ' Conceptos de aporte de pensiones
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    tdbHelp.Caption = "Conceptos Aporte"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cpc", "codcpc", s_Estado_Blq, "")
    nTop = frmCuadro(1).Top + Choose(Index - 6, 1700, 1750)
    nLeft = (frmCuadro(1).Left - 2000) + nLeft
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = tabRegister.Top + nTop
  tdbHelp.Left = tabRegister.Left + nLeft
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index
  
End Sub
Private Sub cmdUpdate_Click(Index As Integer)
  
  ' Pestaña (configuración)
  If txtRemunera(2) = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Afecta Salud (EsSalud)", vbExclamation: txtRemunera(2).SetFocus: Exit Sub
  If lblHelp(2) = "???" Then Beep: MsgBox "Concepto de Remuneración Afecta Salud (EsSalud) no es valido; Verificar", vbExclamation: txtRemunera(2).SetFocus: Exit Sub
  If txtRetencion(2) = "" Then Beep: MsgBox "Debe Ingresar el Concepto de Aporte Salud (EsSalud)", vbExclamation: txtRetencion(2).SetFocus: Exit Sub
  If lblHelp(7) = "???" Then Beep: MsgBox "Concepto de Aporte Salud (EsSalud) no es valido; Verificar", vbExclamation: txtRetencion(2).SetFocus: Exit Sub
    
  If lblHelp(0) = "???" Then Beep: MsgBox "Concepto de Remuneración Afecta IES no es valido; Verificar", vbExclamation: txtRemunera(0).SetFocus: Exit Sub
  If lblHelp(5) = "???" Then Beep: MsgBox "Concepto de Descuento IES no es valido; Verificar", vbExclamation: txtRetencion(0).SetFocus: Exit Sub
  If (txtRemunera(0) = "" And txtRetencion(0) <> "") Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Afecta IES", vbExclamation: txtRemunera(0).SetFocus: Exit Sub
  If (txtRemunera(0) <> "" And txtRetencion(0) = "") Then Beep: MsgBox "Debe Ingresar el Concepto Descuento IES", vbExclamation: txtRetencion(0).SetFocus: Exit Sub
  
  If lblHelp(1) = "???" Then Beep: MsgBox "Concepto de Remuneración Afecta Pensiones (ONP) no es valido; Verificar", vbExclamation: txtRemunera(1).SetFocus: Exit Sub
  If lblHelp(6) = "???" Then Beep: MsgBox "Concepto de Descuento (ONP) no es valido; Verificar", vbExclamation: txtRetencion(1).SetFocus: Exit Sub
  If (txtRemunera(1) = "" And txtRetencion(1) <> "") Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Afecta Pensiones (ONP)", vbExclamation: txtRemunera(1).SetFocus: Exit Sub
  If (txtRemunera(1) <> "" And txtRetencion(1) = "") Then Beep: MsgBox "Debe Ingresar el Concepto Descuento Pensiones (ONP)", vbExclamation: txtRetencion(1).SetFocus: Exit Sub
  
  If lblHelp(3) = "???" Then Beep: MsgBox "Concepto de Remuneración Afecta Fondo Social Artista no es valido; Verificar", vbExclamation: txtRemunera(3).SetFocus: Exit Sub
  If lblHelp(8) = "???" Then Beep: MsgBox "Concepto de Aporte Fondo Social Artista no es valido; Verificar", vbExclamation: txtRetencion(3).SetFocus: Exit Sub
  If (txtRemunera(3) = "" And txtRetencion(3) <> "") Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Afecta Fondo Social Artista", vbExclamation: txtRemunera(3).SetFocus: Exit Sub
  If (txtRemunera(3) <> "" And txtRetencion(3) = "") Then Beep: MsgBox "Debe Ingresar el Concepto Aporte Fondo Social Artista", vbExclamation: txtRetencion(3).SetFocus: Exit Sub
  
  If lblHelp(4) = "???" Then Beep: MsgBox "Concepto de Remuneración Afecta 5ta Categoría no es valido; Verificar", vbExclamation: txtRemunera(4).SetFocus: Exit Sub
  If lblHelp(9) = "???" Then Beep: MsgBox "Concepto de Descuento 5ta Categoría no es valido; Verificar", vbExclamation: txtRetencion(4).SetFocus: Exit Sub
  If (txtRemunera(4) = "" And txtRetencion(4) <> "") Then Beep: MsgBox "Debe Ingresar el Concepto de Remuneración Afecta 5ta Categoría", vbExclamation: txtRemunera(4).SetFocus: Exit Sub
  If (txtRemunera(4) <> "" And txtRetencion(4) = "") Then Beep: MsgBox "Debe Ingresar el Concepto Descuento 5ta Categoría", vbExclamation: txtRetencion(4).SetFocus: Exit Sub
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
    
  ' Creo los arreglos para la actualización parametros
  a_Campos = Array("pdoano", "cpcremuies", "cpcremuonp", "cpcremuessalud", "cpcremuartista", "cpcremuquinta", "cpcies", "cpconp", "cpcessalud", "cpcartista", "cpcquinta", IIf(s_ModoCfg = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(s_ModoCfg = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_Anyo, txtRemunera(0).Text, txtRemunera(1).Text, txtRemunera(2).Text, txtRemunera(3).Text, txtRemunera(4).Text, txtRetencion(0).Text, txtRetencion(1).Text, txtRetencion(2).Text, txtRetencion(3).Text, txtRetencion(4).Text, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("pdoano")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If s_ModoCfg = s_MdoData_Ins Then
    If Not Records_Ins("plparametroafp", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plparametroafp", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Ubico el registro ingresado o actualizado
  ShowScreen
  txtRemunera(0).SetFocus
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 5670: Me.Width = 8500
  Me.Left = 3080: Me.Top = 2000
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Parametros Exportación Información"
  lblTitle = "Parametros"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
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
End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0, 1, 2, 3, 4   ' Concepto de remuneración afecta
    txtRemunera(n_IndexHelp) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtRemunera(n_IndexHelp).SetFocus
   Case 5, 6, 7, 8, 9   ' Concepto descuento y aporte
    txtRetencion(n_IndexHelp - 5) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtRetencion(n_IndexHelp - 5).SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0, 1, 2, 3, 4     ' Concepto de remuneración afecta
    s_Sql = gdl_Funcion.HelpTablas("cpc", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina, "")
   Case 5, 6, 9           ' Conceptos de descuento de pensiones
    s_Sql = gdl_Funcion.HelpTablas("cpc", tdbHelp.Columns(ColIndex).DataField, s_Estado_Act, "")
   Case 7, 8              ' Conceptos de aporte de pensiones
    s_Sql = gdl_Funcion.HelpTablas("cpc", tdbHelp.Columns(ColIndex).DataField, s_Estado_Blq, "")
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
Private Sub txtRemunera_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtRemunera(Index)
End Sub
Private Sub txtRemunera_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index
End Sub
Private Sub txtRemunera_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRemunera_LostFocus(Index As Integer)
  lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRemunera(Index), "CP")
End Sub
Private Sub txtRetencion_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtRetencion(Index)
End Sub
Private Sub txtRetencion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index + 5
End Sub
Private Sub txtRetencion_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtRetencion_LostFocus(Index As Integer)
  lblHelp(Index + 5) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtRetencion(Index), "CP")
End Sub
