VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPersoxEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "selpersoxestado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   7740
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   14
      Top             =   585
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   8546
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2328"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2249"
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
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
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
   Begin MSAdodcLib.Adodc dcaRegistro 
      Height          =   330
      Left            =   45
      Top             =   5490
      Width           =   6840
      _ExtentX        =   12065
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
      Caption         =   "Registro - 00"
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
   Begin Threed.SSPanel panToolBar 
      Height          =   5235
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   585
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   9234
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
         TabIndex        =   13
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Registro"
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
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1635
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
         Picture         =   "selpersoxestado.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   2055
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
         Picture         =   "selpersoxestado.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2760
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
         Picture         =   "selpersoxestado.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   3195
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
         Picture         =   "selpersoxestado.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   4305
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
         Picture         =   "selpersoxestado.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   12
         Tag             =   "0"
         Top             =   4740
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
         Picture         =   "selpersoxestado.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   1200
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
         Picture         =   "selpersoxestado.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   3615
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
         Picture         =   "selpersoxestado.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Tag             =   "0"
         Top             =   495
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
         Picture         =   "selpersoxestado.frx":00EC
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
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
      Begin VB.ComboBox cmbPeriodo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "selpersoxestado.frx":0108
         Left            =   2880
         List            =   "selpersoxestado.frx":010A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   105
         Width           =   2625
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   6525
         TabIndex        =   16
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxestado.frx":010C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   6120
         TabIndex        =   15
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxestado.frx":0128
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   6930
         TabIndex        =   17
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxestado.frx":0144
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   1
         Left            =   795
         TabIndex        =   18
         Top             =   75
         Visible         =   0   'False
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   2
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxestado.frx":0160
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   0
         Left            =   390
         TabIndex        =   19
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   2
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxestado.frx":017C
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   75
         Visible         =   0   'False
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   2
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxestado.frx":0198
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1905
         TabIndex        =   1
         Top             =   180
         Width           =   900
      End
   End
End
Attribute VB_Name = "fSelPersoxEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla
Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(2)                               ' Array de inicio y fin de seleccion de registro
Private s_OptRegistro As String                         ' Instancia del formulario activo
'[
Private Sub RecuperaRegistros(ByVal s_Orden As String)
  
  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, codpsn, apepaterno, apematerno, nombres,"
  s_Sql = s_Sql & " fecnacimiento, ubigeonac, naciextrapsn, sexopsn,"
  s_Sql = s_Sql & " refedirec, codvia, nomviadirec, numerdirec,"
  s_Sql = s_Sql & " intedirec, codzona, nomzondirec, ubigeodir,"
  s_Sql = s_Sql & " estcivilpsn, numhijo, numdepen, coddci, numdociden,"
  s_Sql = s_Sql & " numdocmil, telefono, celular, dctojudicial, pordsctojudi, fotopsn,"
  s_Sql = s_Sql & " fecingreso, codtpt, codcgo, cgoconfianza, codpfs,"
  s_Sql = s_Sql & " codcco, codafp, numeroafp, pagodolar, codbcopago,"
  s_Sql = s_Sql & " cuentapago, ctsdolar, codbcocts, cuentacts, codeps,"
  s_Sql = s_Sql & " regpension, fecingregpen, essvida, cobsctr, afilsindical,"
  s_Sql = s_Sql & " remintegralgrati, remuneta, netocpc, variacpc, imporemuneto,"
  s_Sql = s_Sql & " fecbaja, estadopsn"
  s_Sql = s_Sql & " FROM plpersonal"
  s_Sql = s_Sql & " WHERE codcls='" & ps_ClsPlanilla & "'"
  If Not ribParametro(0).Value Then
    s_Sql = s_Sql & " AND estadopsn" & IIf(ribParametro(1).Value, "<>'I'", "='I'")
  End If
  s_Sql = s_Sql & " ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly
  
  ' Inicializo los rangos de impresion
  as_SelRegistro(0) = "": as_SelRegistro(1) = ""
  If dcaRegistro.Recordset.RecordCount > 0 Then
    dcaRegistro.Recordset.MoveLast: as_SelRegistro(1) = dcaRegistro.Recordset.Bookmark
    dcaRegistro.Recordset.MoveFirst: as_SelRegistro(0) = dcaRegistro.Recordset.Bookmark
  End If

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_FechaHora As String, s_OldMessage As String
  Dim s_Periodo As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Relación de rol de vacaciones
    ' Verifico que existan registros seleccionados
    If cmbPeriodo.Text = "" Then Beep: MsgBox "Debe seleccionar el mes del Analisis", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    s_Periodo = ps_Anyo & Left(cmbPeriodo.Text, 2)
    ' Cambio el Mensaje
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "ROL DE VACACIONES"
    gdl_Procedure.ps_ReportName = "cstrolvaca"
    ReDim aElemento(3, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    aElemento(2, 2) = "Periodo;" & Mid(cmbPeriodo.Text, 6) & " - " & ps_Anyo & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
  ' Genera la información del reporte
    s_Sql = "SELECT psn.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
    s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, CONCAT('" & ps_Anyo & "', '-', DATE_FORMAT(psn.fecingreso,'%m-%d')) AS fecvacacion "
    s_Sql = s_Sql & "FROM plpersonal psn "
    s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND IFNULL(DATE_FORMAT(psn.fecingreso,'%Y'), '')<'" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND IFNULL(DATE_FORMAT(psn.fecingreso,'%m'), '')='" & Right(s_Periodo, 2) & "' "
    s_Sql = s_Sql & "AND IFNULL(DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "')>='" & s_Periodo & "' "
    s_Sql = s_Sql & "ORDER BY fecvacacion, codpsn"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, Index, False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
   Case 1, 2  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index, " ASC", " DESC")
   Case 3 ' Busqueda de registro
    Set go_tdbBusqueda = tdbRegistro
    Set go_dcaBusqueda = dcaRegistro
    gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
    fBusqueda.Show vbModal
   Case 4, 5, 6 ' Selecciono rango de impresión
    gdl_Procedure.MarcaRegistros dcaRegistro, tdbRegistro, as_SelRegistro(0), as_SelRegistro(1), (Index - 4), s_TitleTable
   Case 7, 8  ' Opciones de impresión
    ' Verifico que existan registros seleccionados
    If cmbPeriodo = "" Then Beep: MsgBox "Debe seleccionar el mes del Analisis", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Periodo = ps_Anyo & Left(cmbPeriodo.Text, 2)
        
    ' Cambio el Mensaje
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
      
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = IIf(s_OptRegistro = "anvacacion", "ANALISIS DE VACACIONES", "")
    gdl_Procedure.ps_ReportName = IIf(s_OptRegistro = "anvacacion", "cstvacadet", "cstvacares")
    ReDim aElemento(3, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    aElemento(2, 2) = "Periodo;" & Mid(cmbPeriodo.Text, 6) & " - " & ps_Anyo & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    If s_OptRegistro = "anvacacion" Then
      s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " "
      ' Provisión de vacaciones
      s_Sql = s_Sql & "SELECT pvs.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, MAX(CONCAT(pvs.pdoano,pvs.pdomes)) AS codpdo, MAX(pvs.pdoano) AS anopdo, MAX(pvs.pdomes) AS mespdo, CONCAT(LEFT(pvs.pdopvs, 4), '-', SUBSTRING(pvs.pdopvs, 5, 4)) AS pdovaca, ROUND(SUM(pvs.numerodias),3) AS diaspvs, "
      s_Sql = s_Sql & "MIN(pvs.fechaini) AS fechainivaca, MAX(pvs.fechafin) AS fechafinvaca, ROUND(0, 3) AS diasvaca, "
      s_Sql = s_Sql & "IF(IFNULL(DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "')<'" & s_Periodo & "', DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "') AS pdoproce, "
      s_Sql = s_Sql & "'DIASPROVISIO' AS descri_tipovaca, ROUND(0, 3) AS diastipovp1, ROUND(0, 3) AS diastipovp2, ROUND(0,3) AS diastipovven "
      s_Sql = s_Sql & "FROM plpvsvacaciondet pvs "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON pvs.codcls=psn.codcls AND pvs.codpsn=psn.codpsn "
      s_Sql = s_Sql & "WHERE pvs.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND pvs.fechaini>=psn.fecingreso "
      s_Sql = s_Sql & "AND CONCAT(pvs.pdoano, pvs.pdomes)<='" & s_Periodo & "' "
      s_Sql = s_Sql & "AND pvs.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "GROUP BY pvs.codpsn, pvs.pdopvs "
      s_Sql = s_Sql & "UNION "
      ' Vacaciones fisicas 1
      s_Sql = s_Sql & "SELECT asi.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, asi.codpdo, pdo.anopdo, pdo.mespdo, CONCAT(LEFT(pdovaca1, 4), '-', SUBSTRING(pdovaca1, 5, 4))  AS pdovaca, ROUND(0,3) AS diaspvs, "
      s_Sql = s_Sql & "fechainivaca1 AS fechainivaca, fechafinvaca1 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca1, fechainivaca1) + 1, 0) AS diasvaca, "
      s_Sql = s_Sql & "IF(IFNULL(DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "')<'" & s_Periodo & "', DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "') AS pdoproce, "
      s_Sql = s_Sql & "'V_FISICAS_P1' AS descri_tipovaca, ROUND(IFNULL(DateDiff(fechafinvaca1, fechainivaca1) + 1, 0), 3) AS diastipovp1, ROUND(0, 3) AS diastipovp2, ROUND(0,3) AS diastipovven "
      s_Sql = s_Sql & "FROM plasistencia asi "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & s_Periodo & "' "
      s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechainivaca1) "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechafinvaca1) "
      s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "UNION "
      ' Vacaciones fisicas 2
      s_Sql = s_Sql & "SELECT asi.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, asi.codpdo, pdo.anopdo, pdo.mespdo, CONCAT(LEFT(pdovaca2, 4), '-', SUBSTRING(pdovaca2, 5, 4))  AS pdovaca, ROUND(0,3) AS diaspvs, "
      s_Sql = s_Sql & "fechainivaca2 AS fechainivaca, fechafinvaca2 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca2, fechainivaca2) + 1, 0) AS diasvaca, "
      s_Sql = s_Sql & "IF(IFNULL(DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "')<'" & s_Periodo & "', DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "') AS pdoproce, "
      s_Sql = s_Sql & "'V_FISICAS_P2' AS descri_tipovaca, ROUND(0,3) AS diastipovp1, ROUND(IFNULL(DateDiff(fechafinvaca2, fechainivaca2) + 1, 0), 3) AS diastipovp2, ROUND(0,3) AS diastipovven "
      s_Sql = s_Sql & "FROM plasistencia asi "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & s_Periodo & "' "
      s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechainivaca2) "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechafinvaca2) "
      s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "UNION "
      ' Venta de Vacaciones
      s_Sql = s_Sql & "SELECT asi.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, asi.codpdo, pdo.anopdo, pdo.mespdo, CONCAT(LEFT(pdovaca3, 4), '-', SUBSTRING(pdovaca3, 5, 4))  AS pdovaca, ROUND(0,3) AS diaspvs, "
      s_Sql = s_Sql & "fechainivaca3 AS fechainivaca, fechafinvaca3 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca3, fechainivaca3) + 1, 0) AS diasvaca, "
      s_Sql = s_Sql & "IF(IFNULL(DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "')<'" & s_Periodo & "', DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "') AS pdoproce, "
      s_Sql = s_Sql & "'V_VENDIDAS' AS descri_tipovaca, ROUND(0,3) AS diastipovp1, ROUND(0,3) AS diastipovp2, ROUND(IFNULL(DateDiff(fechafinvaca3, fechainivaca3) + 1, 0), 3) AS diastipovven "
      s_Sql = s_Sql & "FROM plasistencia asi "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & s_Periodo & "' "
      s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechainivaca3) "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechafinvaca3) "
      s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "UNION "
      ' Vacaciones vencidas 1
      s_Sql = s_Sql & "SELECT asi.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, asi.codpdo, pdo.anopdo, pdo.mespdo, CONCAT(LEFT(pdovaca4, 4), '-', SUBSTRING(pdovaca4, 5, 4))  AS pdovaca, 0000.000 AS diaspvs, "
      s_Sql = s_Sql & "fechainivaca4 AS fechainivaca, fechafinvaca4 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca4, fechainivaca4) + 1, 0) AS diasvaca, "
      s_Sql = s_Sql & "IF(IFNULL(DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "')<'" & s_Periodo & "', DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "') AS pdoproce, "
      s_Sql = s_Sql & "'V_VENDIDAS' AS descri_tipovaca, ROUND(0,3) AS diastipovp1, ROUND(0,3) AS diastipovp2, ROUND(IFNULL(DateDiff(fechafinvaca4, fechainivaca4) + 1, 0), 3) AS diastipovven "
      s_Sql = s_Sql & "FROM plasistencia asi "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & s_Periodo & "' "
      s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechainivaca4) "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechafinvaca4) "
      s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "UNION "
      ' Vacaciones vencidas 2
      s_Sql = s_Sql & "SELECT asi.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, asi.codpdo, pdo.anopdo, pdo.mespdo, CONCAT(LEFT(pdovaca5, 4), '-', SUBSTRING(pdovaca5, 5, 4))  AS pdovaca, ROUND(0,3) AS diaspvs, "
      s_Sql = s_Sql & "fechainivaca5 AS fechainivaca, fechafinvaca5 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca5, fechainivaca5) + 1, 0) AS diasvaca, "
      s_Sql = s_Sql & "IF(IFNULL(DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "')<'" & s_Periodo & "', DATE_FORMAT(psn.fecbaja,'%Y%m'), '" & s_Periodo & "') AS pdoproce, "
      s_Sql = s_Sql & "'V_VENDIDAS' AS descri_tipovaca, ROUND(0,3) AS diastipovp1, ROUND(0,3) AS diastipovp2, ROUND(IFNULL(DateDiff(fechafinvaca5, fechainivaca5) + 1, 0), 3) AS diastipovven "
      s_Sql = s_Sql & "FROM plasistencia asi "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & s_Periodo & "' "
      s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechainivaca5) "
      s_Sql = s_Sql & "AND NOT ISNULL(asi.fechafinvaca5) "
      s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "ORDER BY codpsn, LEFT(descri_tipovaca, 2), pdovaca, anopdo, mespdo, codpdo"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    End If
    ' Dias de vacaciones pendientes
    If ribAnalisis(0).Value Then
        
    End If
    ' Genera la información del reporte
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY codpsn, pdovaca, anopdo, mespdo, codpdo"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 7), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
    ' Elimino la tabla temporal
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
  End Select

End Sub
Private Sub Form_Activate()
  ' Bloqueo la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6340: Me.Width = 7830
  Me.Left = 1000: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Trabajador(es)"
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Apellido Paterno", "Apellido Materno", "Nombre(s)", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpsn", "apepaterno", "apematerno", "nombres", "estadopsn")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1080, 1616.33, 1616.33, 1616.33, 300)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0)
  Next n_Index
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(4).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(4).ValueItems.Translate = True
  For n_Index = 0 To 5
    tdbRegistro.Columns(4).ValueItems.Add Item
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "A", "V", "L", "P", "O", "I")
    tdbRegistro.Columns(4).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "estadok", "estadovo", "estadnok", "estadopk", "estadopn", "procenok") & ".bmp")
  Next n_Index
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  tdbRegistro.AllowColMove = False
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(9, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "analmovs", "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Rol de Vacaciones", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
    
    ribAnalisis(n_Index).PictureUp = LoadPicture()
    ribAnalisis(n_Index).ToolTipText = "Análisis " & Choose(n_Index + 1, "Pendiente", "Cancelados", "Historico")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "anctpend", "anctcanc", "ancthist") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribAnalisis(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
 '[ Configuración el control de ayuda
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  ribParametro(0).Value = True
  ribAnalisis(0).Value = True
  
  cmdAction(0).Visible = False
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' Habilito la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = True
End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index
End Sub
Private Sub ribParametro_Click(Index As Integer, Value As Integer)
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
End Sub
Private Sub tdbRegistro_DblClick()
  ' cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub

