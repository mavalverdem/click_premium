VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fAbcEntidadPension 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5835
   ClientLeft      =   2265
   ClientTop       =   375
   ClientWidth     =   7740
   Icon            =   "abcenafp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   7740
   Begin MSAdodcLib.Adodc dcaHelp 
      Height          =   330
      Left            =   510
      Top             =   5325
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   3600
      TabIndex        =   42
      Top             =   5310
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
      TabIndex        =   41
      Top             =   600
      Width           =   6795
      _ExtentX        =   11986
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
      TabPicture(0)   =   "abcenafp.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSunat(11)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmCuadro(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmCuadro(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmCuadro(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescripcion"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboSunat"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.ComboBox cboSunat 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "abcenafp.frx":0028
         Left            =   240
         List            =   "abcenafp.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1340
         MaxLength       =   50
         TabIndex        =   3
         Top             =   615
         Width           =   5265
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1340
         MaxLength       =   8
         TabIndex        =   1
         Top             =   270
         Width           =   825
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   1695
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   1005
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   2990
         _StockProps     =   14
         Caption         =   " Porcentajes "
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
         Begin VB.TextBox txtFactor 
            Height          =   300
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   555
            Width           =   975
         End
         Begin VB.TextBox txtFactor 
            Height          =   300
            Index           =   1
            Left            =   1680
            TabIndex        =   8
            Top             =   555
            Width           =   975
         End
         Begin VB.TextBox txtFactor 
            Height          =   300
            Index           =   3
            Left            =   1680
            TabIndex        =   12
            Top             =   1215
            Width           =   975
         End
         Begin VB.TextBox txtFactor 
            Height          =   300
            Index           =   2
            Left            =   150
            TabIndex        =   10
            Top             =   1215
            Width           =   975
         End
         Begin VB.Label lblDato 
            Caption         =   "Comisión Mixta :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   1680
            TabIndex        =   11
            Top             =   945
            Width           =   1170
         End
         Begin VB.Label lblDato 
            Caption         =   "Inv Sobr GS. :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   9
            Top             =   945
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Comisión % :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   7
            Top             =   285
            Width           =   1005
         End
         Begin VB.Label lblDato 
            Caption         =   "Comisión Fija :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   285
            Width           =   1005
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   3240
         Index           =   1
         Left            =   3375
         TabIndex        =   13
         Top             =   1005
         Width           =   3240
         _Version        =   65536
         _ExtentX        =   5715
         _ExtentY        =   5715
         _StockProps     =   14
         Caption         =   "  Cuentas Corrientes "
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
         Begin VB.TextBox txtDenomina 
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   2835
            Width           =   2970
         End
         Begin VB.TextBox txtCuenta 
            Height          =   300
            Index           =   1
            Left            =   1800
            TabIndex        =   21
            Top             =   2115
            Width           =   1275
         End
         Begin VB.TextBox txtDenomina 
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   1695
            Width           =   2970
         End
         Begin VB.TextBox txtCuenta 
            Height          =   300
            Index           =   0
            Left            =   1560
            TabIndex        =   17
            Top             =   975
            Width           =   1515
         End
         Begin VB.TextBox txtBanco 
            Height          =   300
            Left            =   120
            MaxLength       =   8
            TabIndex        =   15
            Top             =   555
            Width           =   630
         End
         Begin Threed.SSCommand cmdHelp 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   43
            Top             =   555
            Width           =   285
            _Version        =   65536
            _ExtentX        =   494
            _ExtentY        =   494
            _StockProps     =   78
            Caption         =   "..."
            Enabled         =   0   'False
         End
         Begin VB.Label lblDato 
            Caption         =   "Denominación de Cuenta Fondo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   22
            Top             =   2550
            Width           =   2640
         End
         Begin VB.Label lblDato 
            Caption         =   "Nro. Cta. Cte. Fondo :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   20
            Top             =   2160
            Width           =   1560
         End
         Begin VB.Label lblDato 
            Caption         =   "Denominación de Cuenta AFP :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   2640
         End
         Begin VB.Label lblDato 
            Caption         =   "Nro. Cta. Cte. AFP :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   16
            Top             =   1020
            Width           =   1560
         End
         Begin VB.Label lblDato 
            Caption         =   "Banco :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   255
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
            Index           =   0
            Left            =   1200
            TabIndex        =   44
            Top             =   600
            Width           =   195
         End
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   720
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   3480
         Width           =   2850
         _Version        =   65536
         _ExtentX        =   5027
         _ExtentY        =   1270
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
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Activo"
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
            Left            =   1080
            TabIndex        =   26
            Top             =   300
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Inactivo"
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
      Begin VB.Label lblSunat 
         Caption         =   "Codigo Sunat :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   45
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   315
         Width           =   1000
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7740
      _Version        =   65536
      _ExtentX        =   13652
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
         Picture         =   "abcenafp.frx":002C
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6300
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
         Picture         =   "abcenafp.frx":0048
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
      Top             =   5325
      Width           =   7740
      _Version        =   65536
      _ExtentX        =   13652
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
         Picture         =   "abcenafp.frx":0064
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4545
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
         Picture         =   "abcenafp.frx":0080
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2835
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
         Picture         =   "abcenafp.frx":009C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2445
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
         Picture         =   "abcenafp.frx":00B8
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   4650
      Index           =   0
      Left            =   6960
      TabIndex        =   36
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
         Picture         =   "abcenafp.frx":00D4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   39
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
         Picture         =   "abcenafp.frx":00F0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   40
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
         Picture         =   "abcenafp.frx":010C
      End
   End
End
Attribute VB_Name = "fAbcEntidadPension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, parametro de codigo
Private s_EntidadPension As String                      ' Codigo del registro
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)

End Sub
Sub ShowScreen()
    
  ' Presenta botones y controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "PK", txtCodigo, "", Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!codafp.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, "", Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!desafp.DefinedSize
    gdl_Procedure.EditText "AT", txtFactor(0), FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFactor(1), FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFactor(2), FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFactor(3), FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtBanco, "", Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!codbco.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(0), "", Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!ctacteafp.DefinedSize
    gdl_Procedure.EditText "AT", txtDenomina(0), "", Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!desctacteafp.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), "", Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!ctactefondo.DefinedSize
    gdl_Procedure.EditText "AT", txtDenomina(1), "", Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!desctactefondo.DefinedSize
    gdl_Procedure.EditCombo "AT", cboSunat, -1, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
  Else
    gdl_Procedure.EditText "PK", txtCodigo, fEntidadPension.dcaRegistro.Recordset!codafp, Me.Tag, True, fEntidadPension.dcaRegistro.Recordset!codafp.DefinedSize
    gdl_Procedure.EditText "AT", txtDescripcion, gdl_Funcion.aTexto(fEntidadPension.dcaRegistro.Recordset!desafp), Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!desafp.DefinedSize
    gdl_Procedure.EditText "AT", txtFactor(0), FormatNumber(fEntidadPension.dcaRegistro.Recordset!factor1, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFactor(1), FormatNumber(fEntidadPension.dcaRegistro.Recordset!factor2, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFactor(2), FormatNumber(fEntidadPension.dcaRegistro.Recordset!factor3, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtFactor(3), FormatNumber(fEntidadPension.dcaRegistro.Recordset!factor4, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditText "AT", txtBanco, gdl_Funcion.aTexto(fEntidadPension.dcaRegistro.Recordset!codbco), Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!codbco.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(0), gdl_Funcion.aTexto(fEntidadPension.dcaRegistro.Recordset!ctacteafp), Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!ctacteafp.DefinedSize
    gdl_Procedure.EditText "AT", txtDenomina(0), gdl_Funcion.aTexto(fEntidadPension.dcaRegistro.Recordset!desctacteafp), Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!desctacteafp.DefinedSize
    gdl_Procedure.EditText "AT", txtCuenta(1), gdl_Funcion.aTexto(fEntidadPension.dcaRegistro.Recordset!ctactefondo), Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!ctactefondo.DefinedSize
    gdl_Procedure.EditText "AT", txtDenomina(1), gdl_Funcion.aTexto(fEntidadPension.dcaRegistro.Recordset!desctactefondo), Me.Tag, False, fEntidadPension.dcaRegistro.Recordset!desctactefondo.DefinedSize
    For n_Index = 0 To cboSunat.ListCount
      If cboSunat.List(n_Index) = gdl_Funcion.aTexto(fEntidadPension.dcaRegistro.Recordset!codsunat) Then Exit For
    Next n_Index
    n_Index = IIf(n_Index > cboSunat.ListCount, -1, n_Index)
    gdl_Procedure.EditCombo "AT", cboSunat, n_Index, Me.Tag, False
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fEntidadPension.dcaRegistro.Recordset!estadoafp = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fEntidadPension.dcaRegistro.Recordset!estadoafp = s_Estado_Ina), Me.Tag, True
  End If
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco, "EB")

End Sub
Private Sub cmdAction_Click(Index As Integer)

  ' Cargo los datos en la ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtCodigo.SetFocus
  ElseIf Index = 2 Then
   txtDescripcion.SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & " '" & Trim$(txtDescripcion) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_EntidadPension = Trim$(txtCodigo)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codafp")
    a_Valores = Array(s_EntidadPension)
    a_Tipos = Array(TipoDato.Caracter)
      
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plentidadafp", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el Ado control y la grilla
    gdl_Procedure.RefreshAdoControl fEntidadPension.dcaRegistro, fEntidadPension.tdbRegistro, lblTitle
    ' Verifico si aun existen registros
    l_ExistRecord = ((fEntidadPension.dcaRegistro.Recordset.EOF And fEntidadPension.dcaRegistro.Recordset.BOF) Or fEntidadPension.dcaRegistro.Recordset.RecordCount = 0)
    If Not l_ExistRecord Then
      fEntidadPension.dcaRegistro.Recordset.Find ("codafp >= '" & s_EntidadPension & "'")
      If fEntidadPension.dcaRegistro.Recordset.EOF Then fEntidadPension.dcaRegistro.Recordset.MoveLast
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

  If Not cmdHelp(Index).Enabled Then Exit Sub
  s_SqlHelp = ""
  If n_IndexHelp = Index Then
    tdbHelp.ZOrder 0
    tdbHelp.Visible = True
    Exit Sub
  End If
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (tabRegister.Top + frmCuadro(1).Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero inicial, anterior, siguiente o final
  Select Case Index
   Case 0: fEntidadPension.dcaRegistro.Recordset.MoveFirst
   Case 1: If Not fEntidadPension.dcaRegistro.Recordset.BOF Then fEntidadPension.dcaRegistro.Recordset.MovePrevious
           If fEntidadPension.dcaRegistro.Recordset.BOF Then fEntidadPension.dcaRegistro.Recordset.MoveFirst
   Case 2: If Not fEntidadPension.dcaRegistro.Recordset.EOF Then fEntidadPension.dcaRegistro.Recordset.MoveNext
           If fEntidadPension.dcaRegistro.Recordset.EOF Then fEntidadPension.dcaRegistro.Recordset.MoveLast
   Case 3: fEntidadPension.dcaRegistro.Recordset.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim s_Estado As String * 1, s_Sunat As String * 100

  'Realizo las validaciones de los campos a actualizar
  If txtCodigo.Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo " & lblTitle.Caption, vbExclamation: txtCodigo.SetFocus: Exit Sub
  If txtDescripcion.Text = "" Then Beep: MsgBox "Debe Ingresar la Descripción " & lblTitle.Caption, vbExclamation: txtDescripcion.SetFocus: Exit Sub
  If cboSunat.Text = "" Then Beep: MsgBox "Seleccione Codigo Sunat " & lblTitle.Caption, vbExclamation: cboSunat.SetFocus: Exit Sub
  s_Estado = IIf(optEstado(0).Value, s_Estado_Act, s_Estado_Ina)

  If Flag_RestringeSistema = "RESTRINGIR" Then
    If Valida_LicenciaUso(ps_Anyo, ps_Fecha_LimiteProc) = False Then
      MsgBox "Entidad de pensión no puede ser registrado" & Chr(13) & "Se requiere Actualización de Componentes" & Chr(13) & "Por favor comuniquese con el personal de Sistemas.", vbInformation
      Exit Sub
    End If
  End If

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_EntidadPension = txtCodigo
  s_Sunat = cboSunat.Text
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codafp", "desafp", "factor1", "factor2", "factor3", "factor4", "codbco", "ctacteafp", "desctacteafp", "ctactefondo", "desctactefondo", "codsunat", "estadoafp", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(txtCodigo.Text, Trim$(txtDescripcion.Text), CDec(txtFactor(0).Text), CDec(txtFactor(1).Text), CDec(txtFactor(2).Text), CDec(txtFactor(3).Text), Trim$(txtBanco.Text), Trim$(txtCuenta(0).Text), Trim$(txtDenomina(0).Text), Trim$(txtCuenta(1).Text), Trim$(txtDenomina(1).Text), s_Sunat, s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
  a_Where = Array("codafp")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plentidadafp", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plentidadafp", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl fEntidadPension.dcaRegistro, fEntidadPension.tdbRegistro, lblTitle
  ' Ubico el registro ingresado o actualizado
  fEntidadPension.dcaRegistro.Recordset.Find ("codafp='" & s_EntidadPension & "'")
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtCodigo.SetFocus
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
  Me.Height = 6320: Me.Width = 7830
  Me.Left = 1080: Me.Top = 600
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Entidad de Pensión"
  lblTitle = "Entidad de Pensión"
  n_IndexHelp = -1
' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo el modo de operación del registro
  Me.Tag = fEntidadPension.Tag

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
  l_ExistRecord = (fEntidadPension.dcaRegistro.Recordset.EOF Or fEntidadPension.dcaRegistro.Recordset.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fEntidadPension.dcaRegistro.Recordset!codafp

  For n_Index = 0 To 16
    cboSunat.AddItem Choose(n_Index + 1, "02 DECRETO LEY 19990 - SISTEMA NACIONAL DE PENSIONES - ONP", "03 DECRETO LEY 20530 - SISTEMA NACIONAL DE PENSIONES", "09 CAJA DE PESCADOR", "10 CAJA DE PENSIONES MILITAR", "11 CAJA DE PENSIONES POLICIAL", "12 OTROS REGIMENES PENSIONARIOS", "13 REGIMEN DEL SDR", "14 LEY 29903 - SNP - INDEPENDIENTE", "15 REP - TRAB. PESQUEROS", "16 LEY 30003 TDEP", "21 SPP INTEGRA", "22 SPP HORIZONTE", "23 SPP PROFUTURO", "24 SPP PRIMA", "25 SPP HABITAT", "98  PEND ELEC REG PENSIONARIO", "99 SIN REGIMEN PENSIONARIO")
  Next n_Index

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
  ' Asigno el control de datos  ala grilla
  tdbHelp.DataSource = dcaHelp
  
  ' Recupero la información
  s_Sql = gdl_Funcion.HelpTablas("bco", tdbHelp.Columns(0).DataField, "", "")
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaHelp, tdbHelp, s_Sql, adCmdText, adLockReadOnly
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

  If dcaHelp.Recordset.RecordCount = 0 Or (dcaHelp.Recordset.EOF And dcaHelp.Recordset.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  txtBanco.Text = tdbHelp.Columns(0).Value
  lblHelp(0) = tdbHelp.Columns(1).Value
  txtBanco.SetFocus

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  s_Sql = gdl_Funcion.HelpTablas("bco", tdbHelp.Columns(ColIndex).DataField, "", "")
  dcaHelp.RecordSource = s_Sql
  dcaHelp.Refresh

End Sub
Private Sub tdbHelp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or (KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown) Then s_SqlHelp = ""
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaHelp, tdbHelp, ""
End Sub
Private Sub tdbHelp_KeyPress(KeyAscii As Integer)
  Dim n_Columna As Integer
  
  If KeyAscii = vbKeyReturn Then
    tdbHelp_DblClick
  ElseIf (UCase$(Chr$(KeyAscii)) >= "A" And UCase$(Chr$(KeyAscii)) <= "Z") Or _
       (Chr$(KeyAscii) >= "0" And Chr$(KeyAscii) <= "9") Or KeyAscii = 32 Or Chr$(KeyAscii) = "." _
       Or Chr$(KeyAscii) = "*" Then
    If Chr$(KeyAscii) = "*" Then
      s_SqlHelp = ""
    Else
      s_SqlHelp = s_SqlHelp & UCase$(Chr$(KeyAscii))
    End If
    n_Columna = tdbHelp.Col
    s_Sql = gdl_Funcion.HelpTablas("bco", tdbHelp.Columns(n_Columna).DataField, "", s_SqlHelp)
    dcaHelp.RecordSource = s_Sql
    dcaHelp.Refresh
    tdbHelp.Col = n_Columna
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
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtBanco_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    txtCuenta(0).SetFocus
    KeyAscii = 0
  End If
End Sub
Private Sub txtBanco_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco, "EB")
End Sub
Private Sub txtCodigo_GotFocus()
  gdl_Procedure.MarcaGet txtCodigo
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    If txtCodigo = "" Then
      Beep
      MsgBox "Debe Ingresar el Código del " & lblTitle, vbExclamation
      txtCodigo.SetFocus
    Else
      txtDescripcion.SetFocus
      KeyAscii = 0
    End If
  End If

End Sub
Private Sub txtCuenta_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtCuenta(Index)
End Sub
Private Sub txtCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If

End Sub
Private Sub txtDenomina_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtDenomina(Index)
End Sub
Private Sub txtDenomina_KeyPress(Index As Integer, KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If

End Sub
Private Sub txtDescripcion_GotFocus()
  gdl_Procedure.MarcaGet txtDescripcion
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    txtFactor(0).SetFocus
    KeyAscii = 0
  End If

End Sub
Private Sub txtFactor_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtFactor(Index)
End Sub
Private Sub txtFactor_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtFactor_Validate(Index As Integer, Cancel As Boolean)
  txtFactor(Index).Text = IIf(Not IsNumeric(txtFactor(Index).Text), 0, txtFactor(Index).Text)
  If CDec(txtFactor(Index).Text) < 0 Then MsgBox "Factor no puede ser negativo; Verifique", vbInformation: txtFactor(Index).SetFocus: Exit Sub
  txtFactor(Index).Text = FormatNumber(CDec(txtFactor(Index).Text), 2)
End Sub
