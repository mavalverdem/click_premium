VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fAbcCuentaCorriente 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5835
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7455
   Icon            =   "abccuentacte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   7455
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2250
      TabIndex        =   44
      Top             =   5295
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
   Begin TabDlg.SSTab tabRegister 
      Height          =   4650
      Left            =   75
      TabIndex        =   43
      Top             =   600
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   8202
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
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "abccuentacte.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblNumero"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDato(8)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCronograma"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkGratificacion"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkDolares"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmCuadro(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpFecha"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "frmCuadro(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "frmCuadro(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtMonto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCuota"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbTipo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbDescuento"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.ComboBox cmbDescuento 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "abccuentacte.frx":0028
         Left            =   270
         List            =   "abccuentacte.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2700
         Width           =   1620
      End
      Begin VB.ComboBox cmbTipo 
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
         Height          =   315
         ItemData        =   "abccuentacte.frx":002C
         Left            =   270
         List            =   "abccuentacte.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   1620
      End
      Begin VB.TextBox txtCuota 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3795
         TabIndex        =   15
         Top             =   2055
         Width           =   1425
      End
      Begin VB.TextBox txtMonto 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2150
         TabIndex        =   13
         Top             =   2055
         Width           =   1425
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1695
         Index           =   1
         Left            =   2150
         TabIndex        =   19
         Top             =   2475
         Width           =   4260
         _Version        =   65536
         _ExtentX        =   7514
         _ExtentY        =   2990
         _StockProps     =   14
         Caption         =   "  Documento  "
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
         Begin VB.TextBox txtDocumento 
            Height          =   285
            Left            =   945
            TabIndex        =   23
            Top             =   615
            Width           =   1500
         End
         Begin VB.TextBox txtBanco 
            Height          =   300
            Left            =   135
            MaxLength       =   8
            TabIndex        =   25
            Top             =   1305
            Width           =   630
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   2
            Left            =   870
            TabIndex        =   45
            Top             =   1305
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSOption optTipDocu 
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   20
            Top             =   285
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "&Carta"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optTipDocu 
            Height          =   180
            Index           =   1
            Left            =   1815
            TabIndex        =   21
            Top             =   285
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   317
            _StockProps     =   78
            Caption         =   "C&heque"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblDato 
            Caption         =   "Número :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   22
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblDato 
            Caption         =   "Banco :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   24
            Top             =   1005
            Width           =   1005
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
            Left            =   1230
            TabIndex        =   46
            Top             =   1350
            Width           =   195
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   960
         Index           =   2
         Left            =   270
         TabIndex        =   26
         Top             =   3180
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   1693
         _StockProps     =   14
         Caption         =   " Estado "
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
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   27
            Top             =   285
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Cancelado"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   28
            Top             =   585
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Pendiente"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   270
         TabIndex        =   4
         Top             =   1425
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   113180673
         CurrentDate     =   37515
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1530
         Index           =   0
         Left            =   2150
         TabIndex        =   5
         Top             =   165
         Width           =   4260
         _Version        =   65536
         _ExtentX        =   7514
         _ExtentY        =   2699
         _StockProps     =   14
         Caption         =   " Datos de Descuento "
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
         Begin VB.TextBox txtPeriodo 
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   120
            TabIndex        =   9
            Top             =   1140
            Width           =   980
         End
         Begin VB.TextBox txtConcepto 
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   120
            TabIndex        =   7
            Top             =   540
            Width           =   980
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   0
            Left            =   1170
            TabIndex        =   47
            Top             =   540
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   300
            Index           =   1
            Left            =   1170
            TabIndex        =   49
            Top             =   1140
            Width           =   300
            _Version        =   65536
            _ExtentX        =   529
            _ExtentY        =   529
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin VB.Label lblDato 
            Caption         =   "Periodo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   8
            Top             =   885
            Width           =   1000
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
            Left            =   1530
            TabIndex        =   50
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label lblDato 
            Caption         =   "Concepto :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   285
            Width           =   1000
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
            Left            =   1530
            TabIndex        =   48
            Top             =   600
            Width           =   195
         End
      End
      Begin Threed.SSCheck chkDolares 
         Height          =   285
         Left            =   270
         TabIndex        =   10
         Top             =   1800
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Dólares"
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
         Font3D          =   1
      End
      Begin Threed.SSCheck chkGratificacion 
         Height          =   285
         Left            =   270
         TabIndex        =   11
         Top             =   2115
         Width           =   1770
         _Version        =   65536
         _ExtentX        =   3122
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Incluye Gratificación"
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
         Font3D          =   1
      End
      Begin Threed.SSCommand cmdCronograma 
         Height          =   300
         Left            =   5955
         TabIndex        =   16
         Top             =   2040
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
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
         AutoSize        =   2
         Picture         =   "abccuentacte.frx":0030
      End
      Begin VB.Label lblDato 
         Caption         =   "Tipo Descuento :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   270
         TabIndex        =   17
         Top             =   2445
         Width           =   1665
      End
      Begin VB.Label lblDato 
         Caption         =   "Nº de Cuota :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3795
         TabIndex        =   14
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblNumero 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   270
         TabIndex        =   2
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label lblDato 
         Caption         =   "Monto :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2150
         TabIndex        =   12
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Caption         =   "Fecha :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   3
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Caption         =   "Tipo Transacción :"
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
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   0
         Top             =   200
         Width           =   1665
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
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
         Left            =   6690
         TabIndex        =   30
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
         Picture         =   "abccuentacte.frx":004C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6300
         TabIndex        =   31
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
         Picture         =   "abccuentacte.frx":0068
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
         Left            =   720
         TabIndex        =   32
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   33
      Top             =   5325
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
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
         Left            =   4935
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
         Picture         =   "abccuentacte.frx":0084
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4545
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
         Picture         =   "abccuentacte.frx":00A0
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2835
         TabIndex        =   36
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
         Picture         =   "abccuentacte.frx":00BC
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2445
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
         Picture         =   "abccuentacte.frx":00D8
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   4650
      Index           =   0
      Left            =   6645
      TabIndex        =   38
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8202
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
         TabIndex        =   39
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
         TabIndex        =   40
         Tag             =   "0"
         Top             =   960
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
         Picture         =   "abccuentacte.frx":00F4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   41
         Tag             =   "0"
         Top             =   1770
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
         Picture         =   "abccuentacte.frx":0110
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   42
         Tag             =   "0"
         Top             =   2550
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
         Picture         =   "abccuentacte.frx":012C
      End
   End
End
Attribute VB_Name = "fAbcCuentaCorriente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, parametro de codigo
Private s_Registro As String                     ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins)
  cmdHelp(1).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdHelp(2).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdCronograma.Visible = (Me.Tag = s_MdoData_Ins)

End Sub
Sub ShowScreen()
    
  ' Presenta botones y controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditCombo "PK", cmbtipo, -1, Me.Tag, False
    lblNumero = ""
    gdl_Procedure.EditText "AT", txtConcepto, "", Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!codcpc.DefinedSize
    gdl_Procedure.EditText "AT", txtPeriodo, "", Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!codpdoprv.DefinedSize
    gdl_Procedure.EditDTPicker "AT", dtpFecha, Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", chkDolares, False, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", chkGratificacion, False, Me.Tag, True
    gdl_Procedure.EditCombo "PK", cmbDescuento, 0, Me.Tag, False
    gdl_Procedure.EditText "AT", txtMonto, FormatNumber(0, 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "AT", txtCuota, FormatNumber(0, 0), Me.Tag, False, 15, vbRightJustify
    gdl_Procedure.EditOptionCheck "AT", optTipDocu(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optTipDocu(1), False, Me.Tag, True
    gdl_Procedure.EditText "AT", txtDocumento, "", Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!numchecar.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco, "", Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!codbco.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), False, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), True, Me.Tag, False
  Else
    n_Index = IIf(fCuentaCorriente.dcaRegistro.Recordset!tpoctacte = "P", 0, 1)
    n_Index = IIf(CInt(fCuentaCorriente.dcaRegistro.Recordset!numcuota) = 0, n_Index, 2)
    gdl_Procedure.EditCombo "PK", cmbtipo, n_Index, Me.Tag, False
    lblNumero = gdl_Funcion.aTexto(fCuentaCorriente.dcaRegistro.Recordset!numctacte)
    gdl_Procedure.EditText "PK", txtConcepto, gdl_Funcion.aTexto(fCuentaCorriente.dcaRegistro.Recordset!codcpc), Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!codcpc.DefinedSize
    gdl_Procedure.EditText "AT", txtPeriodo, gdl_Funcion.aTexto(fCuentaCorriente.dcaRegistro.Recordset!codpdoprv), Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!codpdoprv.DefinedSize
    gdl_Procedure.EditDTPicker "PK", dtpFecha, fCuentaCorriente.dcaRegistro.Recordset!fectacte, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "PK", chkDolares, (fCuentaCorriente.dcaRegistro.Recordset!codmon = "E"), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", chkGratificacion, (fCuentaCorriente.dcaRegistro.Recordset!indgratifi = s_Estado_Act), Me.Tag, False
    n_Index = CInt(fCuentaCorriente.dcaRegistro.Recordset!tpodscto)
    gdl_Procedure.EditCombo "PK", cmbDescuento, n_Index, Me.Tag, False
    gdl_Procedure.EditText "AT", txtMonto, FormatNumber((fCuentaCorriente.dcaRegistro.Recordset!Cargo + fCuentaCorriente.dcaRegistro.Recordset!abono), 2), Me.Tag, False, 18, vbRightJustify
    gdl_Procedure.EditText "PK", txtCuota, FormatNumber(fCuentaCorriente.dcaRegistro.Recordset!numcuota, 0), Me.Tag, False, 15, vbRightJustify
    gdl_Procedure.EditOptionCheck "AT", optTipDocu(0), (fCuentaCorriente.dcaRegistro.Recordset!indchecar = "C"), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optTipDocu(1), (fCuentaCorriente.dcaRegistro.Recordset!indchecar = "H"), Me.Tag, True
    gdl_Procedure.EditText "AT", txtDocumento, gdl_Funcion.aTexto(fCuentaCorriente.dcaRegistro.Recordset!numchecar), Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!numchecar.DefinedSize
    gdl_Procedure.EditText "AT", txtBanco, gdl_Funcion.aTexto(fCuentaCorriente.dcaRegistro.Recordset!codbco), Me.Tag, False, fCuentaCorriente.dcaRegistro.Recordset!codbco.DefinedSize
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fCuentaCorriente.dcaRegistro.Recordset!estadoctacte = s_Estado_Act), Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fCuentaCorriente.dcaRegistro.Recordset!estadoctacte = s_Estado_Ina), Me.Tag, False
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtConcepto, "CP")
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo, "PR")
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco, "EB")

End Sub
Private Function ValidaCuentaCte(nInstancia As Integer) As Boolean
  Dim n_Periodos As Integer
  Dim s_FechaFin  As String
  
  If cmbtipo.Text = "" Then Beep: MsgBox "Debe Ingresar el Tipo de Transacción", vbExclamation: cmbtipo.SetFocus: GoTo FinalError
  If optEstado(0).Value Then Beep: MsgBox "Registro no Actualizable", vbExclamation: cmbtipo.SetFocus: GoTo FinalError
  If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el periodo de descuento", vbExclamation: txtPeriodo.SetFocus: GoTo FinalError
  If lblHelp(1) = "???" Then Beep: MsgBox "Periodo de decuento no es valido; Verificar", vbExclamation: txtPeriodo.SetFocus: GoTo FinalError
  If (Trim(dtpFecha.Year) <> ps_Anyo And Left(cmbtipo, 1) <> "C") Then Beep: MsgBox "Fecha debe ser del periodo activo", vbExclamation: dtpFecha.SetFocus: GoTo FinalError
  If cmbDescuento.Text = "" Then Beep: MsgBox "Debe Ingresar Modalidad de Descuento", vbExclamation: cmbDescuento.SetFocus: GoTo FinalError
  If CDec(txtMonto) <= 0 Then Beep: MsgBox "Monto ingresado invalido", vbExclamation: txtMonto.SetFocus: GoTo FinalError
  If (CInt(txtCuota) <= 0 And Left(cmbtipo, 1) <> "C") Then Beep: MsgBox "Número de cuotas es invalido", vbExclamation: txtCuota.SetFocus: GoTo FinalError
  ' Periodos de descuento
  s_Sql = "SELECT codpdo, fechafin FROM plperiodo "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpdo>='" & Trim(txtPeriodo.Text) & "' "
  s_Sql = s_Sql & "AND tpopdo NOT IN('L'" & IIf(chkGratificacion.Value, ", 'G') ", ") ")
  s_Sql = s_Sql & "AND estadopdo<='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "ORDER BY codpdo"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  n_Periodos = CInt(porstRecordset.RecordCount)
  s_FechaFin = Format(porstRecordset!fechafin, s_FormatoFecha)
  porstRecordset.Close
  Set porstRecordset = Nothing
  If (Not (Format(dtpFecha, "yyyymmdd") <= Format(s_FechaFin, "yyyymmdd")) And nInstancia = 1) Then Beep: MsgBox "Fecha debe ser menor o igual que la fecha del descuento", vbExclamation: dtpFecha.SetFocus: GoTo FinalError
  If (n_Periodos <= 0 And Left(cmbtipo.Text, 1) = "C") Then Beep: MsgBox "No existe periodo para la cuota", vbExclamation: txtPeriodo.SetFocus: GoTo FinalError
  If (n_Periodos < CInt(txtCuota) And Left(cmbtipo.Text, 1) = "P") Then Beep: MsgBox "Número de periodos de descuento es menor a las cuotas", vbExclamation: txtCuota.SetFocus: GoTo FinalError
  dtpFecha.Value = IIf(Left(cmbtipo.Text, 1) = "C", s_FechaFin, dtpFecha.Value)
  ValidaCuentaCte = True
  If nInstancia = 1 Then Exit Function
  If txtConcepto = "" Then Beep: MsgBox "Debe Ingresar el concepto de descuento", vbExclamation: txtConcepto.SetFocus: GoTo FinalError
  If lblHelp(0) = "???" Then Beep: MsgBox "Concepto de decuento no es valido; Verificar", vbExclamation: txtConcepto.SetFocus: GoTo FinalError
  If lblHelp(2) = "???" Then Beep: MsgBox "Entidad bancaria no es valido; Verificar", vbExclamation: txtBanco.SetFocus: GoTo FinalError
  
  Exit Function

FinalError:
  ValidaCuentaCte = False

End Function
Private Sub cmdAction_Click(Index As Integer)
  Dim n_Registro As Integer
  Dim s_Personal As String
  
  ' Cargo los datos en la ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    cmbtipo.SetFocus
  ElseIf Index = 2 Then
   txtConcepto.SetFocus
  End If
  If Index <> 1 Then Exit Sub
  ' Realizo las validaciones de la cuenta corriente actualizar
  If Not optEstado(1).Value Then Beep: MsgBox cmbtipo & " se encuentra cancelado; verificar", vbExclamation: GoTo Finalizar
  s_Sql = "SELECT COUNT(*) AS registro "
  s_Sql = s_Sql & "FROM plcuentacte "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpsn='" & Trim(fSelPersonal.dcaRegistro.Recordset!codpsn) & "' "
  s_Sql = s_Sql & "AND numctacte='" & Trim(lblNumero) & "' "
  s_Sql = s_Sql & "AND numcuota>=" & CInt(txtCuota.Text) & " "
  s_Sql = s_Sql & "AND estadoctacte<>'" & s_Estado_Ina & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  n_Registro = CInt(porstRecordset!registro)
  porstRecordset.Close
  If n_Registro > 0 Then Beep: MsgBox cmbtipo & " se encuentra cancelado; verificar", vbExclamation: GoTo Finalizar
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & cmbtipo & " '" & lblNumero & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim$(lblNumero)
    n_Registro = CInt(txtCuota.Text)
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpsn", "numctacte")
    a_Valores = Array(ps_ClsPlanilla, Trim(fSelPersonal.dcaRegistro.Recordset!codpsn), s_Registro)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    If CInt(txtCuota.Text) <> 0 Then
      a_Where = Array("codcls", "codpsn", "numctacte", "numcuota")
      a_Valores = Array(ps_ClsPlanilla, Trim(fSelPersonal.dcaRegistro.Recordset!codpsn), s_Registro, n_Registro)
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero)
    End If
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plcuentacte", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblNumero & "-" & txtCuota, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fCuentaCorriente.dcaRegistro, fCuentaCorriente.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fCuentaCorriente.dcaRegistro.Recordset.EOF And fCuentaCorriente.dcaRegistro.Recordset.BOF) Or fCuentaCorriente.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fCuentaCorriente.dcaRegistro.Recordset.Find ("cPrimaryKey >= '" & s_Registro & n_Registro & "'")
      If fCuentaCorriente.dcaRegistro.Recordset.EOF Then fCuentaCorriente.dcaRegistro.Recordset.MoveLast
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
Private Sub cmdCronograma_Click()
  ' Realizo la validación de los campos
  If Not ValidaCuentaCte(1) Then Exit Sub
  fCronogramaCuotas.Show vbModal
End Sub
Private Sub cmdHelp_Click(Index As Integer)
  Dim s_TablaHelp As String
  
  s_SqlHelp = ""
  If n_IndexHelp = Index And Index <> 1 Then
    tdbHelp.ZOrder 0
    tdbHelp.Visible = True
    Exit Sub
  End If

  Select Case Index
   Case 0     ' Conceptos de planilla
    tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
    s_TablaHelp = "Concepto de Planilla"
    s_Registro = ps_ClsPlanilla & "C" & s_Estado_Act
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", s_Registro, "")
   Case 1     ' Periodo de Pago
    tdbHelp.Columns(0).DataField = "codpdo": tdbHelp.Columns(1).DataField = "despdo"
    s_TablaHelp = "Periodos de Pago"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("pxe", "codpdo", s_Estado_Ina & ps_ClsPlanilla & ps_Anyo, "")
   Case 2     ' Entidad bancaria
    tdbHelp.Columns(0).DataField = "codbco": tdbHelp.Columns(1).DataField = "desbco"
    s_TablaHelp = "Entidad Bancaria"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("bco", "codbco", "", "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = IIf(Index = 2, 1200, frmCuadro(0).Top) + (tabRegister.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)))
  tdbHelp.Left = (frmCuadro(0).Left / 2) + (tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2)))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fCuentaCorriente.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fCuentaCorriente.dcaRegistro.Recordset.BOF Then fCuentaCorriente.dcaRegistro.Recordset.MovePrevious
           If fCuentaCorriente.dcaRegistro.Recordset.BOF Then fCuentaCorriente.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fCuentaCorriente.dcaRegistro.Recordset.EOF Then fCuentaCorriente.dcaRegistro.Recordset.MoveNext
           If fCuentaCorriente.dcaRegistro.Recordset.EOF Then fCuentaCorriente.dcaRegistro.Recordset.MoveLast
   Case 3: fCuentaCorriente.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1, s_TipoCtaCte As String * 1, s_TipoDocu As String * 1
  Dim s_Gratifica As String * 1, s_Descuento As String * 1, s_Moneda As String * 1
  Dim n_CuotaIni As Integer, n_CuotaFin As Integer
  Dim s_Periodo As String, s_Fecha As String, s_Banco As String
  Dim n_Importe As Double, n_Cargomn  As Double
  Dim n_Abonomn  As Double, n_Cargome As Double, n_Abonome As Double
  
  ' Realizo las validaciones de los campos a actualizar
  If Not ValidaCuentaCte(0) Then Exit Sub
  
  s_TipoCtaCte = Left(cmbtipo.Text, 1)
  s_TipoCtaCte = IIf(s_TipoCtaCte = "C", "P", s_TipoCtaCte)
  s_Periodo = Trim(txtPeriodo.Text)
  s_Fecha = Format(dtpFecha, s_FormatoFecha)
  s_Gratifica = IIf(chkGratificacion.Value, s_Estado_Act, s_Estado_Ina)
  s_Moneda = IIf(chkDolares.Value, "E", "N")
  s_Descuento = Trim(cmbDescuento.ListIndex)
  s_TipoDocu = IIf(optTipDocu(0).Value, "C", "H")
  s_Banco = Trim(txtBanco)
  s_Estado = IIf(optEstado(0).Value, s_Estado_Act, s_Estado_Ina)
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = Trim(lblNumero)
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpsn", "numctacte", "numcuota", "tpoctacte", "codcpc", "codpdoprv", "fectacte", "indchecar", "numchecar", "codbco", "indgratifi", "tpodscto", "codmon", "cargo_mn", "abono_mn", "cargo_me", "abono_me", "indprn", "codpdocan", "estadoctacte", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codcls", "codpsn", "numctacte", "numcuota")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    ' Obtengo el numero de cuenta corriente
    s_Sql = "SELECT CONCAT('" & ps_Anyo & "', IFNULL(MAX(RIGHT(numctacte, 2)), '00')) AS snumctacte"
    s_Sql = s_Sql & " FROM plcuentacte"
    s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
    s_Sql = s_Sql & " AND codpsn='" & Trim(fSelPersonal.dcaRegistro.Recordset!codpsn) & "'"
    s_Sql = s_Sql & " AND tpoctacte='" & s_TipoCtaCte & "'"
    s_Sql = s_Sql & " AND SUBSTR(numctacte, 3, 4)='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    lblNumero = s_TipoCtaCte & "-" & Trim(Val(porstRecordset!snumctacte) + 1)
    
    ' Periodos de descuento
    s_Sql = "SELECT codpdo, fechafin FROM plperiodo "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codpdo>='" & Trim(txtPeriodo.Text) & "' "
    s_Sql = s_Sql & "AND tpopdo NOT IN('L'" & IIf(chkGratificacion.Value, ", 'G') ", ") ")
    s_Sql = s_Sql & "AND estadopdo<='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "ORDER BY codpdo "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    n_CuotaFin = CInt(txtCuota.Text)
    n_CuotaIni = IIf(Left(cmbtipo.Text, 1) = "C", n_CuotaFin, 0)
    n_Importe = CDec(txtMonto.Text)
    For n_Index = n_CuotaIni To n_CuotaFin
      s_Periodo = gdl_Funcion.aTexto(porstRecordset!codpdo)
      s_Fecha = Format(IIf(n_Index = 0, dtpFecha, porstRecordset!fechafin), s_FormatoFecha)
      n_Cargomn = FormatNumber(IIf(n_Index = 0, n_Importe, 0), 2)
      n_Abonomn = FormatNumber(IIf(n_Index = 0, 0, CDec(n_Importe / n_CuotaFin)), 2)
      n_Cargome = FormatNumber(0, 2)
      n_Abonome = FormatNumber(0, 2)
      If s_Moneda = "E" Then
        n_Cargome = FormatNumber(IIf(n_Index = 0, n_Importe, 0), 2)
        n_Abonome = FormatNumber(IIf(n_Index = 0, 0, CDec(n_Importe / n_CuotaFin)), 2)
        n_Cargomn = FormatNumber(0, 2)
        n_Abonomn = FormatNumber(0, 2)
      End If
      a_Valores = Array(ps_ClsPlanilla, Trim(fSelPersonal.dcaRegistro.Recordset!codpsn), Trim(lblNumero), n_Index, s_TipoCtaCte, Trim(txtConcepto), s_Periodo, Format(s_Fecha, s_FmtFechMysql_0), s_TipoDocu, Trim(txtDocumento), Trim(txtBanco), s_Gratifica, s_Descuento, s_Moneda, n_Cargomn, n_Abonomn, n_Cargome, n_Abonome, s_Estado_Ina, "", s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
      If Not Records_Ins("plcuentacte", a_Campos, a_Valores, a_Tipos) Then GoTo Error
      If n_Index > 0 Then porstRecordset.MoveNext
    Next n_Index
    porstRecordset.Close
    Set porstRecordset = Nothing
  Else
    n_Cargomn = FormatNumber(0, 2)
    n_Abonomn = FormatNumber(IIf(s_Moneda = "N", CDec(txtMonto), 0), 2)
    n_Cargome = FormatNumber(0, 2)
    n_Abonome = FormatNumber(IIf(s_Moneda = "E", CDec(txtMonto), 0), 2)
    a_Valores = Array(ps_ClsPlanilla, Trim(fSelPersonal.dcaRegistro.Recordset!codpsn), Trim(lblNumero), CInt(txtCuota.Text), s_TipoCtaCte, Trim(txtConcepto), s_Periodo, Format(s_Fecha, s_FmtFechMysql_0), s_TipoDocu, Trim(txtDocumento), Trim(txtBanco), s_Gratifica, s_Descuento, s_Moneda, n_Cargomn, n_Abonomn, n_Cargome, n_Abonome, s_Estado_Ina, "", s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
    If Not Records_Upd("plcuentacte", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  ' Capturo el registro a actualizar
  s_Registro = Trim(lblNumero) & Format(txtCuota.Text, "00")
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fCuentaCorriente.dcaRegistro, fCuentaCorriente.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fCuentaCorriente.dcaRegistro.Recordset.Find ("cPrimaryKey='" & s_Registro & "'")
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
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
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece Posición y Titulo del Formulario
  Me.Height = 6320: Me.Width = 7545
  Me.Left = 2580: Me.Top = 550
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Cuenta Corriente"
  lblTitle = "Cuenta Corriente"
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fCuentaCorriente.Tag

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
  gdl_Procedure.LoadGrafics cmdCronograma, "ajuinfla", "Visualiza cronograma de cuotas"

  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0

  ' Verifico si existen Registros
  l_ExistRecord = (fCuentaCorriente.dcaRegistro.Recordset.EOF Or fCuentaCorriente.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fCuentaCorriente.dcaRegistro.Recordset!numctacte
  ' Configuro los listados, datos adicionales
  For n_Index = 0 To 2: cmbtipo.AddItem Choose(n_Index + 1, "Préstamo", "Adelanto", "Cuota"): Next n_Index
  For n_Index = 0 To 2: cmbDescuento.AddItem Choose(n_Index + 1, "Mensual", "Quincena", "Dualidad"): Next n_Index

  ' Carga los datos en el formulario
  ShowScreen

 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codbco", "desbco")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Entidad Bancaria", 2
  ']

  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
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
   Case 0       ' Concepto de cuenta corriente
    txtConcepto = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtConcepto.SetFocus
   Case 1       ' Periodo de pago
    txtPeriodo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo.SetFocus
   Case 2       ' Entidad bancaria
    txtBanco = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtBanco.SetFocus
  End Select
   
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Conceptos de planillas
    s_Registro = ps_ClsPlanilla & "C" & s_Estado_Ina
    s_Sql = gdl_Funcion.HelpTablas("cxt", tdbHelp.Columns(ColIndex).DataField, s_Registro, "")
   Case 1     ' Periodo de Pago
    s_Sql = gdl_Funcion.HelpTablas("pxe", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla & ps_Anyo, "")
   Case 2     ' Entidad bancaria
    s_Sql = gdl_Funcion.HelpTablas("bco", tdbHelp.Columns(ColIndex).DataField, "", "")
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
Private Sub txtBanco_GotFocus()
  gdl_Procedure.MarcaGet txtBanco
End Sub
Private Sub txtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 2
End Sub
Private Sub txtBanco_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If

End Sub
Private Sub txtBanco_LostFocus()
  lblHelp(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco, "EB")
End Sub
Private Sub txtConcepto_GotFocus()
  gdl_Procedure.MarcaGet txtConcepto
End Sub
Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtPeriodo.SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtConcepto_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtConcepto, "CP")
End Sub
Private Sub txtCuota_GotFocus()
  gdl_Procedure.MarcaGet txtCuota
End Sub
Private Sub txtCuota_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCuota_Validate(Cancel As Boolean)
  txtCuota.Text = IIf(Not IsNumeric(txtCuota.Text), 0, txtCuota.Text)
  txtCuota.Text = FormatNumber(CDec(txtCuota.Text), 0)
End Sub
Private Sub txtDocumento_GotFocus()
  gdl_Procedure.MarcaGet txtDocumento
End Sub
Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    txtBanco.SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtMonto_GotFocus()
  gdl_Procedure.MarcaGet txtMonto
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    If CDec(txtMonto) <= 0 Then
      Beep
      MsgBox "Debe Ingresar el Valor de la " & lblTitle, vbExclamation
      txtMonto.SetFocus
    Else
      SendKeys "{TAB}"
    End If
    KeyAscii = 0
  End If

End Sub
Private Sub txtMonto_Validate(Cancel As Boolean)
  txtMonto.Text = IIf(Not IsNumeric(txtMonto.Text), 0, txtMonto.Text)
  txtMonto.Text = FormatNumber(CDec(txtMonto.Text), 2)
End Sub
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 1
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    txtMonto.SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtPeriodo_LostFocus()
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo, "PR")
End Sub
