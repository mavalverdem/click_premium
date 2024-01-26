VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPersonalCts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "selpersoxcts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8490
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   15
      Top             =   975
      Width           =   7620
      _ExtentX        =   13441
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
      Top             =   5880
      Width           =   7620
      _ExtentX        =   13441
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
      Left            =   7695
      TabIndex        =   4
      Top             =   975
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
         TabIndex        =   14
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
         TabIndex        =   7
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
         Picture         =   "selpersoxcts.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   8
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
         Picture         =   "selpersoxcts.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   9
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
         Picture         =   "selpersoxcts.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   10
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
         Picture         =   "selpersoxcts.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   12
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
         Picture         =   "selpersoxcts.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   13
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
         Picture         =   "selpersoxcts.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   6
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
         Picture         =   "selpersoxcts.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   11
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
         Picture         =   "selpersoxcts.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   5
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
         Picture         =   "selpersoxcts.frx":00EC
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   930
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8490
      _Version        =   65536
      _ExtentX        =   14975
      _ExtentY        =   1640
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
      Begin VB.TextBox txtPeriodo 
         ForeColor       =   &H00800000&
         Height          =   280
         Index           =   0
         Left            =   2910
         TabIndex        =   1
         Top             =   150
         Width           =   810
      End
      Begin VB.TextBox txtPeriodo 
         ForeColor       =   &H00800000&
         Height          =   280
         Index           =   1
         Left            =   2910
         TabIndex        =   3
         Top             =   495
         Width           =   810
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   7410
         TabIndex        =   18
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
         PictureUp       =   "selpersoxcts.frx":0108
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   7005
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
         PictureUp       =   "selpersoxcts.frx":0124
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   7815
         TabIndex        =   19
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
         PictureUp       =   "selpersoxcts.frx":0140
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   1
         Left            =   660
         TabIndex        =   22
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
         PictureUp       =   "selpersoxcts.frx":015C
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   0
         Left            =   255
         TabIndex        =   21
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
         PictureUp       =   "selpersoxcts.frx":0178
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   3780
         TabIndex        =   25
         Top             =   150
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
         Left            =   3780
         TabIndex        =   26
         Top             =   495
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   2
         Left            =   1065
         TabIndex        =   23
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
         PictureUp       =   "selpersoxcts.frx":0194
      End
      Begin Threed.SSRibbon ribFirma 
         Height          =   360
         Left            =   7815
         TabIndex        =   20
         Top             =   495
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   0
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxcts.frx":01B0
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   3
         Left            =   255
         TabIndex        =   24
         Top             =   495
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
         PictureUp       =   "selpersoxcts.frx":01CC
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sub-Periodo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Periodo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Top             =   195
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
         Height          =   180
         Index           =   0
         Left            =   4140
         TabIndex        =   28
         Top             =   195
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
         Left            =   4140
         TabIndex        =   27
         Top             =   540
         Width           =   180
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   0
         Left            =   1725
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   5010
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   3810
      TabIndex        =   29
      Top             =   825
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
Attribute VB_Name = "fSelPersonalCts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(2)                               ' Array de inicio y fin de seleccion de registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_OptRegistro As String                         ' Instancia del formulario activo
'[
Private Sub DepositoCts(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sPersonal As String, sDocIdentidad As String, sEntidadBanco As String
  Dim nRemBasica As Double, nRemPromedio As Double, nRemGratifica As Double
  Dim nRemAfecta  As Double, nImporteCts As Double, sCuenta As String
  Dim s_Moneda As String, sMonRemunera As String, sDescripcionBanco As String
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando Depósito ..."
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo mes
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.pdocts, res.subcts, "
  s_Sql = s_Sql & "res.codpsn, psn.apepaterno, psn.apematerno, psn.nombres, dci.sigladci, psn.numdociden, "
  s_Sql = s_Sql & "psn.fecnacimiento, psn.fecingreso, psn.fecbaja, psn.naciextrapsn, "
  s_Sql = s_Sql & "psn.pagodolar, psn.ctsdolar, psn.cuentacts, psn.codbcocts, bco.desbco, mov.fechacan "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
  s_Sql = s_Sql & "LEFT JOIN plbanco bco ON psn.codbcocts=bco.codbco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdocts='" & Trim(txtPeriodo(0).Text) & "' "
  s_Sql = s_Sql & "AND res.subcts<='" & Trim(txtPeriodo(1).Text) & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "ORDER BY codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal del certificado
  s_Sql = "DROP TABLE IF EXISTS tmpimporte"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimporte ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, "
  s_Sql = s_Sql & "rembasica decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "rempromedio decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remgratifi decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remunects decimal(18, 2) NOT Null Default 0) "
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Inserto las remuneraciones basicas
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, ROUND(SUM(CASE WHEN psn.pagodolar='" & s_Estado_Act & "' THEN res.importe_me ELSE res.importe_mn END), 2) AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remubasicacts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto las remuneraciones promedio
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS rembasica, ROUND(SUM(CASE WHEN psn.pagodolar='" & s_Estado_Act & "' THEN res.importe_me ELSE res.importe_mn END), 2) AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remupromects "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneracion de gratificacion
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "ROUND(SUM(CASE WHEN psn.pagodolar='" & s_Estado_Act & "' THEN res.importe_me ELSE res.importe_mn END), 2) AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remugraticts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneraciones de cts
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, ROUND(SUM(res.importe_" & s_Moneda & "),2) AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remutotalcts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ']
  
  ' Recupero la informacion del deposito
  s_Sql = "SELECT tmp.codpsn, psn.apepaterno, psn.apematerno, psn.nombres, psn.sigladci, psn.numdociden, psn.fecnacimiento, "
  s_Sql = s_Sql & "psn.pdocts, psn.pagodolar, psn.ctsdolar, psn.cuentacts, psn.codbcocts, psn.desbco, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(rembasica, 0)), 2) AS rembasica, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(rempromedio, 0)), 2) AS rempromedio, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(remgratifi, 0)), 2) AS remgratifi, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(remunects, 0)), 2) AS remunects "
  s_Sql = s_Sql & "FROM tmpimporte tmp "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON tmp.codpsn=psn.codpsn "
  s_Sql = s_Sql & "GROUP BY codpsn "
  s_Sql = s_Sql & "ORDER BY psn.codbcocts, tmp.codpsn"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    a_Campos = Array("codbco", "desbco", "codpsn", "sigladci", "numdociden", "apepaterno", "apematerno", "nombres", "fecnacimiento", "monedacts", "cuentacts", "impremunera", "monedarem", "impacumula")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero)
    While Not porstRecordset.EOF
      sPersonal = porstRecordset!codpsn
      sMonRemunera = IIf(porstRecordset!pagodolar = s_Estado_Act, s_Codmon_me_Nom, s_Codmon_mn_Nom)
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      sEntidadBanco = IIf(IsNull(porstRecordset!codbcocts), "XX", porstRecordset!codbcocts)
      sDescripcionBanco = IIf(IsNull(porstRecordset!desbco), "XXXXX", porstRecordset!desbco)
      sCuenta = IIf(IsNull(porstRecordset!cuentacts), "", porstRecordset!cuentacts)
      If CDec(porstRecordset!rembasica) > 0 Then
        ' Obtengo la renta Bruta
        nRemBasica = CDec(porstRecordset!rembasica)
        nRemPromedio = CDec(porstRecordset!rempromedio)
        nRemGratifica = CDec(porstRecordset!remgratifi)
        nRemAfecta = Round((nRemBasica + nRemPromedio + nRemGratifica), 2)
        nImporteCts = CDec(porstRecordset!remunects)
        a_Valores = Array(sEntidadBanco, sDescripcionBanco, sPersonal, gdl_Funcion.aTexto(porstRecordset!sigladci), sDocIdentidad, UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno)), UCase(gdl_Funcion.aTexto(porstRecordset!apematerno)), UCase(gdl_Funcion.aTexto(porstRecordset!nombres)), Format(porstRecordset!fecnacimiento, s_FmtFechMysql_0), s_Moneda, sCuenta, nImporteCts, sMonRemunera, nRemAfecta)
        
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      End If
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub CertificadoCts(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim s_Moneda As String
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sPersonal As String, sSiglaDocu As String, sDocIdentidad As String
  Dim sBancoCts As String, sCuentaCts As String, sDesFechaCan As String, sDesFechas As String
  Dim nImpMonedaCts As Double, nImpCambioCts As Double, nTipoCambio As Double
  Dim nTasaInteres As Double, nInteres As Double, nDias As Long
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando Certificado ..."
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo mes
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.pdocts, res.subcts, res.codpsn, "
  s_Sql = s_Sql & "CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  s_Sql = s_Sql & "dci.sigladci, psn.numdociden, psn.fecingreso, psn.fecbaja, psn.naciextrapsn, "
  s_Sql = s_Sql & "psn.ctsdolar, psn.cuentacts, (CASE WHEN interbankcts='" & s_Estado_Act & "' THEN  bnk.desbco ELSE bco.desbco END) AS desbco, mov.fechaini, mov.fechafin, mov.fechaven, mov.fechacan, "
  s_Sql = s_Sql & "mov.numeroanos, mov.numeromeses, mov.numerodias, mov.porinteres, mov.tipocambio "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
  s_Sql = s_Sql & "LEFT JOIN plbanco bco ON psn.codbcocts=bco.codbco "
  s_Sql = s_Sql & "LEFT JOIN plbanco bnk ON psn.codbnkcts=bnk.codbco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdocts='" & Trim(txtPeriodo(0).Text) & "' "
  s_Sql = s_Sql & "AND res.subcts='" & Trim(txtPeriodo(1).Text) & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Blq & "' "
  s_Sql = s_Sql & "ORDER BY codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal del certificado
  s_Sql = "DROP TABLE IF EXISTS tmpimporte"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimporte ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, "
  s_Sql = s_Sql & "remudiacts decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remumescts decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remuanocts decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remucts_mn decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remucts_me decimal(18, 2) NOT Null Default 0)"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Inserto remuneración cts por dias
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.importe_" & s_Moneda & " AS remudiacts, 0.00 AS remumescts, "
  s_Sql = s_Sql & "0.00 AS remuanocts, 0.00 AS remucts_mn, 0.00 AS remucts_me "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remudiascts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto remuneración cts por meses
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS remudiacts, res.importe_" & s_Moneda & " AS remumescts, "
  s_Sql = s_Sql & "0.00 AS remuanocts, 0.00 AS remucts_mn, 0.00 AS remucts_me "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remumesescts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto remuneración cts por años
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS remudiacts, 0.00 AS remumescts, "
  s_Sql = s_Sql & "res.importe_" & s_Moneda & " AS remuanocts, 0.00 AS remucts_mn, 0.00 AS remucts_me "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remuanoscts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneraciones de cts
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS remudiacts, 0.00 AS remumescts, "
  s_Sql = s_Sql & "0.00 AS remuanocts, res.importe_mn AS remucts_mn, res.importe_me AS remucts_me "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remutotalcts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal importes generales cts
  s_Sql = "DROP TABLE IF EXISTS tmpimportects"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimportects "
  s_Sql = s_Sql & "SELECT codpsn, SUM(IFNULL(remudiacts, 0)) AS remudiacts, "
  s_Sql = s_Sql & "SUM(IFNULL(remumescts, 0)) AS remumescts, SUM(IFNULL(remuanocts, 0)) AS remuanocts, "
  s_Sql = s_Sql & "SUM(IFNULL(remucts_mn, 0)) AS remucts_mn, SUM(IFNULL(remucts_me, 0)) AS remucts_me "
  s_Sql = s_Sql & "FROM tmpimporte "
  s_Sql = s_Sql & "GROUP BY codpsn "
  s_Sql = s_Sql & "ORDER BY codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Recupero la informacion del certificado
  s_Sql = "SELECT tmp.codpsn, psn.nombrespsn, psn.sigladci, psn.numdociden, psn.fecingreso, psn.fecbaja, psn.naciextrapsn, "
  s_Sql = s_Sql & "psn.pdocts, psn.subcts, psn.ctsdolar, psn.cuentacts, psn.desbco, psn.fechaini, psn.fechafin, psn.fechaven, "
  s_Sql = s_Sql & "psn.fechacan, psn.numeroanos, psn.numeromeses, psn.numerodias, psn.porinteres, psn.tipocambio, "
  s_Sql = s_Sql & "tmp.remudiacts, tmp.remumescts, tmp.remuanocts, tmp.remucts_mn, tmp.remucts_me, "
  s_Sql = s_Sql & "res.secuencia, res.codcpc, cpc.descpc, res.importe_" & s_Moneda & " AS importecpc "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN tmpimportects tmp ON res.codpsn=tmp.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN plconceplanilla cxp ON res.codcls=cxp.codcls AND res.codcpc=cxp.codcpc "
  s_Sql = s_Sql & "WHERE res.impbolecpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "OR cxp.defaultcpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "ORDER BY tmp.codpsn, res.secuencia"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    a_Campos = Array("codpsn", "nombrespsn", "sigladci", "numdociden", "monedacts", "moncuentacts", "cuentacts", "desbco", "fecingreso", "fecbaja", "fechacan", "desfechacan", "fechaini", "fechafin", "desfechas", "numeroanos", "numeromeses", "numerodias", "moneda", "secuencia", "codcpc", "descpc", "importecpc", "importeano", "importemes", "importedia", "porinteres", "imporinteres", "impordeposito", "tipocambio", "imporcambio")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    While Not porstRecordset.EOF
      sPersonal = porstRecordset!codpsn
      sSiglaDocu = IIf(IsNull(porstRecordset!sigladci), "", porstRecordset!sigladci)
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      sCuentaCts = IIf(IsNull(porstRecordset!cuentacts), "", porstRecordset!cuentacts)
      sBancoCts = IIf(IsNull(porstRecordset!desbco), "", porstRecordset!desbco)
      sDesFechaCan = Format(porstRecordset!fechacan, "dd") & " de " & gdl_Funcion.NombreMes(Format(porstRecordset!fechacan, "mm")) & " del " & Format(porstRecordset!fechacan, "yyyy")
      sDesFechas = Format(porstRecordset!fechaini, "dd") & " de " & gdl_Funcion.NombreMes(Format(porstRecordset!fechaini, "mm")) & " del " & Format(porstRecordset!fechaini, "yyyy")
      sDesFechas = sDesFechas & " al " & Format(porstRecordset!fechafin, "dd") & " de " & gdl_Funcion.NombreMes(Format(porstRecordset!fechafin, "mm")) & " del " & Format(porstRecordset!fechafin, "yyyy")
        
      ' Obtengo los importes
      nImpMonedaCts = CDec(porstRecordset("remucts_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me")))
      nImpCambioCts = CDec(porstRecordset("remucts_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn")))
      nTipoCambio = CDec(porstRecordset!Tipocambio)
      nDias = gdl_Funcion.NumeroDias360(Format(porstRecordset!fechacan, s_FormatoFecha), Format(DateAdd("d", 1, porstRecordset!fechaven), s_FormatoFecha), Format(porstRecordset!fechacan, s_FormatoFecha))
      'AGREGADO 22/05/2009
      If nDias <= 30 Then
        nDias = DateDiff("d", Format(porstRecordset!fechaven, s_FormatoFecha), Format(porstRecordset!fechacan, s_FormatoFecha))
      End If
      nTasaInteres = (((CDec(porstRecordset!porinteres) / 100) + 1) ^ (nDias / 360)) - 1
      nInteres = CDec(nImpMonedaCts * nTasaInteres)
      ' Importe incluye interes
      nImpMonedaCts = CDec(nImpMonedaCts + CDec(nImpMonedaCts * nTasaInteres))
      nImpCambioCts = CDec(nImpCambioCts + CDec(nImpCambioCts * nTasaInteres))
      a_Valores = Array(sPersonal, UCase(porstRecordset!nombrespsn), sSiglaDocu, sDocIdentidad, Choose(porstRecordset!ctsdolar + 1, s_Codmon_mn_Txt, s_Codmon_me_Txt), Choose(porstRecordset!ctsdolar + 1, s_Codmon_mn_Nom, s_Codmon_me_Nom), sCuentaCts, sBancoCts, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), Format(porstRecordset!fechacan, s_FmtFechMysql_0), sDesFechaCan, Format(porstRecordset!fechaini, s_FmtFechMysql_0), Format(porstRecordset!fechafin, s_FmtFechMysql_0), sDesFechas, CInt(porstRecordset!numeroanos), CInt(porstRecordset!numeromeses), CInt(porstRecordset!numerodias), s_Moneda, CLng(porstRecordset!secuencia), porstRecordset!codcpc, porstRecordset!descpc, CDec(porstRecordset!importecpc), CDec(porstRecordset!remuanocts), CDec(porstRecordset!remumescts), CDec(porstRecordset!remudiacts), CDec(porstRecordset!porinteres), nInteres, nImpMonedaCts, nTipoCambio, nImpCambioCts)
      
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      ' Realizo la actualización de los registros
      If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, codpsn, apepaterno, apematerno, nombres, "
  s_Sql = s_Sql & "CONCAT(IFNULL(apepaterno, ''), ' ', IFNULL(apematerno, ''), ', ', IFNULL(nombres, '')) AS nombrepsn, "
  s_Sql = s_Sql & "fecnacimiento, ubigeonac, naciextrapsn, sexopsn, "
  s_Sql = s_Sql & "refedirec, codvia, nomviadirec, numerdirec, "
  s_Sql = s_Sql & "intedirec, codzona, nomzondirec, ubigeodir, "
  s_Sql = s_Sql & "estcivilpsn, numhijo, numdepen, coddci, numdociden, "
  s_Sql = s_Sql & "numdocmil, telefono, celular, dctojudicial, pordsctojudi, fotopsn, "
  s_Sql = s_Sql & "fecingreso, codtpt, codcgo, cgoconfianza, codpfs, "
  s_Sql = s_Sql & "codcco, codafp, numeroafp, pagodolar, codbcopago, "
  s_Sql = s_Sql & "cuentapago, ctsdolar, codbcocts, cuentacts, codeps, "
  s_Sql = s_Sql & "regpension, fecingregpen, essvida, cobsctr, afilsindical, "
  s_Sql = s_Sql & "remintegralgrati, remuneta, netocpc, variacpc, imporemuneto, "
  s_Sql = s_Sql & "fecbaja, estadopsn "
  s_Sql = s_Sql & "FROM plpersonal "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  If Not ribParametro(0).Value Then
    s_Sql = s_Sql & "AND estadopsn" & IIf(ribParametro(1).Value, "<>'I' ", "='I' ")
  End If
  s_Sql = s_Sql & "ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly
  
  ' Inicializo los rangos de impresion
  as_SelRegistro(0) = "": as_SelRegistro(1) = ""
  If dcaRegistro.Recordset.RecordCount > 0 Then
    dcaRegistro.Recordset.MoveLast: as_SelRegistro(1) = dcaRegistro.Recordset.Bookmark
    dcaRegistro.Recordset.MoveFirst: as_SelRegistro(0) = dcaRegistro.Recordset.Bookmark
  End If

End Sub
Private Sub RemuneraCts(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim s_Moneda As String, s_MonedaCts As String
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sPersonal As String, sSiglaDocu  As String, sDocIdentidad As String
  Dim sCodBancoCts As String, sDesBancoCts As String, sCuentaCts  As String
  Dim nRemBasica As Double, nRemPromedio As Double, nRemGratifica As Double
  Dim nRemAfecta  As Double, nImporteCts As Double
  Dim nRemuneracion(12) As Double
  Dim nSecuencia As Integer
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando Remuneraciones ..."
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo mes
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.pdocts, res.codpsn, "
  s_Sql = s_Sql & "CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  s_Sql = s_Sql & "dci.sigladci, psn.numdociden, psn.fecnacimiento, psn.ctsdolar, psn.cuentacts, "
  s_Sql = s_Sql & "(CASE WHEN interbankcts='" & s_Estado_Act & "' THEN  psn.codbnkcts ELSE psn.codbcocts END) AS codbcocts, "
  s_Sql = s_Sql & "(CASE WHEN interbankcts='" & s_Estado_Act & "' THEN  bnk.desbco ELSE bco.desbco END) AS desbco "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
  s_Sql = s_Sql & "LEFT JOIN plbanco bco ON psn.codbcocts=bco.codbco "
  s_Sql = s_Sql & "LEFT JOIN plbanco bnk ON psn.codbnkcts=bnk.codbco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdocts='" & Trim(txtPeriodo(0).Text) & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "ORDER BY codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal del certificado
  s_Sql = "DROP TABLE IF EXISTS tmpimporte"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimporte ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, "
  s_Sql = s_Sql & "subcts char(2) NOT Null, "
  s_Sql = s_Sql & "rembasica decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "rempromedio decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remgratifi decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remunects decimal(18, 2) NOT Null Default 0) "
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Inserto las remuneraciones basicas
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.subcts, res.importe_" & s_Moneda & " AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remubasicacts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.subcts<='" & Trim(txtPeriodo(1).Text) & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto las remuneraciones promedio
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.subcts, 0.00 AS rembasica, res.importe_" & s_Moneda & " AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remupromects "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.subcts<='" & Trim(txtPeriodo(1).Text) & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneracion de gratificacion
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.subcts, 0.00 AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "res.importe_" & s_Moneda & " AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remugraticts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.subcts<='" & Trim(txtPeriodo(1).Text) & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneraciones de cts
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.subcts, 0.00 AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, res.importe_" & s_Moneda & " AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remutotalcts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.subcts<='" & Trim(txtPeriodo(1).Text) & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ']
  
  ' Recupero la informacion del certificado
  s_Sql = "SELECT tmp.codpsn, psn.nombrespsn, psn.sigladci, psn.numdociden, psn.fecnacimiento, psn.ctsdolar, "
  s_Sql = s_Sql & "psn.cuentacts, psn. codbcocts, psn.desbco, psn.pdocts, tmp.subcts, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(tmp.rembasica, 0)), 2) AS rembasica, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(tmp.rempromedio, 0)), 2) AS rempromedio, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(tmp.remgratifi, 0)), 2) AS remgratifi, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(tmp.remunects, 0)), 2) AS remunects "
  s_Sql = s_Sql & "FROM tmpimporte tmp "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON tmp.codpsn=psn.codpsn "
  s_Sql = s_Sql & "GROUP BY codpsn, psn.pdocts, tmp.subcts "
  s_Sql = s_Sql & "ORDER BY codpsn, psn.pdocts, tmp.subcts"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    a_Campos = Array("codpsn", "nombrespsn", "sigladci", "numdociden", "fecnacimiento", "codbco", "desbco", "monedacts", "cuentacts", "impormes_01", "impormes_02", "impormes_03", "impormes_04", "impormes_05", "impormes_06")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    While Not porstRecordset.EOF
      sPersonal = porstRecordset!codpsn
      sSiglaDocu = IIf(IsNull(porstRecordset!sigladci), "", porstRecordset!sigladci)
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      s_MonedaCts = IIf(porstRecordset!ctsdolar = s_Estado_Act, s_Codmon_me_Txt, s_Codmon_mn_Txt)
      sCodBancoCts = IIf(IsNull(porstRecordset!codbcocts), "", porstRecordset!codbcocts)
      sDesBancoCts = IIf(IsNull(porstRecordset!desbco), "", porstRecordset!desbco)
      sCuentaCts = IIf(IsNull(porstRecordset!cuentacts), "", porstRecordset!cuentacts)
      For nSecuencia = 1 To 12: nRemuneracion(nSecuencia) = 0: Next nSecuencia
      nSecuencia = 0
      Do
        If CDec(porstRecordset!rembasica) > 0 Then
          ' Obtengo la renta Bruta
          nRemBasica = CDec(porstRecordset!rembasica)
          nRemPromedio = CDec(porstRecordset!rempromedio)
          nRemGratifica = CDec(porstRecordset!remgratifi)
          nRemAfecta = Round((nRemBasica + nRemPromedio + nRemGratifica), 2)
          nImporteCts = CDec(porstRecordset!remunects)
          nSecuencia = CInt(porstRecordset!subcts)
          nRemuneracion(nSecuencia) = nRemAfecta
        End If
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        DoEvents
        porstRecordset.MoveNext
        If porstRecordset.EOF Then Exit Do
      Loop Until sPersonal <> porstRecordset!codpsn
      porstRecordset.MovePrevious
      
      a_Valores = Array(sPersonal, UCase(porstRecordset!nombrespsn), sSiglaDocu, sDocIdentidad, Format(porstRecordset!fecnacimiento, s_FmtFechMysql_0), sCodBancoCts, sDesBancoCts, s_MonedaCts, sCuentaCts, Round(nRemuneracion(5) + nRemuneracion(11), 2), Round(nRemuneracion(6) + nRemuneracion(12), 2), Round(nRemuneracion(7) + nRemuneracion(1), 2), Round(nRemuneracion(8) + nRemuneracion(2), 2), Round(nRemuneracion(9) + nRemuneracion(3), 2), Round(nRemuneracion(4) + nRemuneracion(10), 2))
      
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      ' Realizo la actualización de los registros
      If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      
      porstRecordset.MoveNext
    Wend
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub ResumenCts(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim s_Moneda As String
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sPersonal As String, sDocIdentidad As String, sCargo As String
  Dim nRemBasica As Double, nRemPromedio As Double, nRemGratifica As Double
  Dim nRemAfecta  As Double, nImporteCts As Double, sCuenta As String
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando Certificado ..."
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo mes
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.pdocts, res.subcts, res.codpsn, cgo.descgo, "
  s_Sql = s_Sql & "CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  s_Sql = s_Sql & "psn.numdociden, psn.fecingreso, psn.fecbaja, psn.naciextrapsn, "
  s_Sql = s_Sql & "mov.numeroanos, mov.numeromeses, mov.numerodias,psn.codacredor "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND psn.estadopsn<>'I' "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdocts='" & Trim(txtPeriodo(0).Text) & "' "
  s_Sql = s_Sql & "AND res.subcts='" & Trim(txtPeriodo(1).Text) & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "ORDER BY codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal del certificado
  s_Sql = "DROP TABLE IF EXISTS tmpimporte"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimporte ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, "
  s_Sql = s_Sql & "rembasica decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "rempromedio decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remgratifi decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remunects decimal(18, 2) NOT Null Default 0) "
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Inserto las remuneraciones basicas
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.importe_" & s_Moneda & " AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remubasicacts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto las remuneraciones promedio
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS rembasica, res.importe_" & s_Moneda & " AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remupromects "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneracion de gratificacion
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "res.importe_" & s_Moneda & " AS remgratifi, 0.00 AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remugraticts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneraciones de cts
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS rembasica, 0.00 AS rempromedio, "
  s_Sql = s_Sql & "0.00 AS remgratifi, res.importe_" & s_Moneda & " AS remunects "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remutotalcts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ']
  
  ' Recupero la informacion del certificado
  s_Sql = "SELECT tmp.codpsn, psn.nombrespsn, psn.numdociden, psn.fecingreso, psn.fecbaja, psn.naciextrapsn, "
  s_Sql = s_Sql & "psn.pdocts, psn.subcts, psn.descgo, psn.numeroanos, psn.numeromeses, psn.numerodias, "
  s_Sql = s_Sql & "SUM(IFNULL(rembasica, 0)) AS rembasica, "
  s_Sql = s_Sql & "SUM(IFNULL(rempromedio, 0)) AS rempromedio, "
  s_Sql = s_Sql & "SUM(IFNULL(remgratifi, 0)) AS remgratifi, "
  s_Sql = s_Sql & "SUM(IFNULL(remunects, 0)) AS remunects,psn.codacredor "
  s_Sql = s_Sql & "FROM tmpimporte tmp "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON tmp.codpsn=psn.codpsn "
  s_Sql = s_Sql & "GROUP BY codpsn "
  s_Sql = s_Sql & "ORDER BY codpsn"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    a_Campos = Array("codpsn", "nombrespsn", "numdociden", "fecingreso", "fecbaja", "moneda", "descargo", "rembasica", "rempromedio", "remgratifica", "remuneafecta", "remuneracts", "numeroanos", "numeromeses", "numerodias", "codacredor")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter)
    While Not porstRecordset.EOF
      sPersonal = porstRecordset!codpsn
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      sCargo = IIf(IsNull(porstRecordset!descgo), "", porstRecordset!descgo)
      sCuenta = IIf(IsNull(porstRecordset!codacredor), "", porstRecordset!codacredor)
      If CDec(porstRecordset!rembasica) > 0 Then
        ' Obtengo la renta Bruta
        nRemBasica = CDec(porstRecordset!rembasica)
        nRemPromedio = CDec(porstRecordset!rempromedio)
        nRemGratifica = CDec(porstRecordset!remgratifi)
        nRemAfecta = Round((nRemBasica + nRemPromedio + nRemGratifica), 2)
        nImporteCts = CDec(porstRecordset!remunects)
        a_Valores = Array(sPersonal, UCase(porstRecordset!nombrespsn), sDocIdentidad, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), s_Moneda, sCargo, nRemBasica, nRemPromedio, nRemGratifica, nRemAfecta, nImporteCts, CInt(porstRecordset!numeroanos), CInt(porstRecordset!numeromeses), CInt(porstRecordset!numerodias), sCuenta)
        
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      End If
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim sSubTitulo As String, sRepresentante As String
  Dim sDireccion As String, sDepartamento As String
  Dim sProvincia As String, sDistrito As String
  Dim sEmail As String, s_FechaHora As String
  Dim sExpresion As String, sCargoRepresenta As String

  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Select Case Index
   Case 0  ' Actualización de parametros
    If s_OptRegistro = "repcoxtise" Then
      fPrmCertifikCts.Show vbModal
    End If
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
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    If txtPeriodo(0).Text = "" Then Beep: MsgBox "Debe Ingresar el Periodo de Analisis", vbExclamation: txtPeriodo(0).SetFocus: Exit Sub
    If lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Analisis no es valido; Verificar", vbExclamation: txtPeriodo(0).SetFocus: Exit Sub
    If txtPeriodo(1).Text = "" Then Beep: MsgBox "Debe Ingresar el Sub Periodo de Analisis", vbExclamation: txtPeriodo(1).SetFocus: Exit Sub
    If lblHelp(1) = "???" Then Beep: MsgBox "Sub Periodo de Analisis no es valido; Verificar", vbExclamation: txtPeriodo(1).SetFocus: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    
    ' Obtengo y verifico los datos de la empresa
    sDireccion = "": sRepresentante = "": sSubTitulo = ""
    sDepartamento = "": sProvincia = "": sDistrito = ""
    sEmail = ""
    sExpresion = IIf((ribAnalisis(2).Value And ribFirma.Value), "ger", "rep")
    
    s_Sql = "SELECT cfg.codvia, cfg.direccionvia, cfg.numerodir, cfg.codzona, cfg.direccionzona, cfg.ubigeodir, cfg.email, "
    s_Sql = s_Sql & "CONCAT(IFNULL(cfg." & sExpresion & "apepaterno, ''), ' ', IFNULL(cfg." & sExpresion & "apematerno, ''), ', ', IFNULL(cfg." & sExpresion & "nombres, '')) AS representante, "
    s_Sql = s_Sql & "IFNULL(cfg." & sExpresion & "cargo, '') AS repcargo, "
    s_Sql = s_Sql & "cfg." & sExpresion & "numdocu AS repnumdocu, "
    s_Sql = s_Sql & "afp.remubasicacts, afp.remupromects, afp.remutotalcts, afp.remumesescts "
    s_Sql = s_Sql & "FROM plcfgempresa cfg "
    s_Sql = s_Sql & "INNER JOIN plparametroafp afp ON cfg.pdoano=afp.pdoano "
    s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sSubTitulo = gdl_Funcion.aTexto(IIf(ribAnalisis(0).Value, porstRecordset!remubasicacts, porstRecordset!remumesescts))
      sRepresentante = gdl_Funcion.aTexto(porstRecordset!representante)
      sCargoRepresenta = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, gdl_Funcion.aTexto(porstRecordset!repcargo), "DC")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      sDepartamento = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Ina, Left(sDireccion, 2), "UB")
      sProvincia = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Act, Left(sDireccion, 4), "UB")
      sDistrito = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " - " & sDistrito
      sEmail = gdl_Funcion.aTexto(porstRecordset!Email)
    End If
    porstRecordset.Close
    If sSubTitulo = "" Then Beep: MsgBox "Debe configurar los parametros del reporte", vbExclamation: cmdAction(0).SetFocus: Exit Sub

    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Obtengo los datos del sub periodo
    sSubTitulo = "Periodo : "
    s_Sql = "SELECT fechaini, fechafin "
    s_Sql = s_Sql & "FROM plctsperiodosub "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdocts='" & Trim(txtPeriodo(0).Text) & "' "
    s_Sql = s_Sql & "AND subcts='" & Trim(txtPeriodo(1).Text) & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sSubTitulo = UCase(gdl_Funcion.NombreMes(Format(porstRecordset!fechaini, "mm")))
      sSubTitulo = sSubTitulo & " de " & Format(porstRecordset!fechaini, "yyyy")
      sSubTitulo = sSubTitulo & " a " & UCase(gdl_Funcion.NombreMes(Format(porstRecordset!fechafin, "mm")))
      sSubTitulo = sSubTitulo & " de " & Format(porstRecordset!fechafin, "yyyy")
      sSubTitulo = "PERIODO : " & sSubTitulo
    End If
    porstRecordset.Close
   
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = IIf(ribAnalisis(0).Value, "ANALISIS DE C.T.S. ( " & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt) & " ) ", IIf(ribAnalisis(1).Value, "DEPÓSITO DE C.T.S. ( " & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt) & " ) ", IIf(ribAnalisis(2).Value, "CONSTANCIA DE LIQUIDACION DE C.T.S. ", "ANALISIS DE REMUNERACIONES DE C.T.S. ( " & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt) & " ) "))) & Trim(lblHelp(0).Caption)
    gdl_Procedure.ps_ReportName = IIf(ribAnalisis(0).Value, "rptanalisicts", IIf(ribAnalisis(1).Value, "rptdeposicts", IIf(ribAnalisis(2).Value, "rptcerticts", "rptremunects")))
    ReDim aElemento(3, 7): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = "": aElemento(0, 3) = "": aElemento(0, 4) = ""
    aElemento(0, 5) = "": aElemento(0, 6) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    aElemento(1, 3) = "": aElemento(1, 4) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "Direccion;" & sDireccion & ";true"
    aElemento(2, 2) = "Ruc;" & ps_RucEmpresa & ";true"
    aElemento(2, 3) = "Representante;" & sRepresentante & ";true"
    aElemento(2, 4) = "TituloReporte;" & UCase(gdl_Procedure.ps_ReportTitle) & ";true"
    aElemento(2, 5) = "SubTitulo;" & sSubTitulo & ";true"
    aElemento(2, 6) = ""
    If ribAnalisis(2).Value Then
      aElemento(2, 6) = "CargoRepresenta;" & sCargoRepresenta & ";true"
    End If
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""

    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    ' Genera la información del reporte
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If s_OptRegistro = "repcoxtise" Then
      If ribAnalisis(0).Value Then
        s_Sql = s_Sql & "codpsn varchar(11) Not Null, nombrespsn varchar(80) Null, numdociden varchar(11) Null, "
        s_Sql = s_Sql & "fecingreso date Null, fecbaja date Null, moneda char(3) Null, descargo varchar(50) Null, "
        s_Sql = s_Sql & "rembasica decimal(18,2) Null Default '0', rempromedio decimal(18,2) Null Default '0',  "
        s_Sql = s_Sql & "remgratifica decimal(18,2) Null Default '0', numeroanos smallint(2) DEFAULT '0', "
        s_Sql = s_Sql & "numeromeses smallint(2) Null Default '0', numerodias smallint(2) Null Default '0', "
        s_Sql = s_Sql & "remuneafecta decimal(18,2) Null Default '0', remuneracts decimal(18,2) Null Default '0', codacredor varchar(15) Null,"
        s_Sql = s_Sql & "PRIMARY KEY (codpsn)) "
      ElseIf ribAnalisis(1).Value Then
        s_Sql = s_Sql & "codbco char(3) Not Null, desbco varchar(40) Null, codpsn varchar(11) Not Null, sigladci char(3) Null, numdociden varchar(11) Null, "
        s_Sql = s_Sql & "apepaterno varchar(25) Null, apematerno varchar(25) Null, nombres varchar(25) Null, fecnacimiento date Null, "
        s_Sql = s_Sql & "monedacts char(3) Null, cuentacts varchar(20) Null, impremunera decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "monedarem varchar(18) Null, impacumula decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "PRIMARY KEY (codbco, codpsn)) "
      ElseIf ribAnalisis(2).Value Then
        s_Sql = s_Sql & "codpsn varchar(11) Not Null, nombrespsn varchar(80) Null, sigladci char(3) Null, numdociden varchar(11) Null, "
        s_Sql = s_Sql & "monedacts char(3) Null, moncuentacts varchar(20) Null, cuentacts varchar(20) Null, desbco varchar(40) Null, "
        s_Sql = s_Sql & "fecingreso date Null, fecbaja date Null, fechacan date Null, desfechacan varchar(50) Null,"
        s_Sql = s_Sql & "fechaini date Null, fechafin date Null, desfechas varchar(60) Null,"
        s_Sql = s_Sql & "numeroanos smallint(2) Null Default '0', numeromeses smallint(2) Null Default '0', numerodias smallint(2) Null Default '0', "
        s_Sql = s_Sql & "moneda char(3) Null, secuencia int(3) Not Null Default '0', codcpc varchar(4) Not Null, descpc varchar(40) Null, "
        s_Sql = s_Sql & "importecpc decimal(18,2) Null Default '0', importeano decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "importemes decimal(18,2) Null Default '0', importedia decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "porinteres decimal(5,2) Null Default '0', imporinteres decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "impordeposito decimal(18,2) Null Default '0', tipocambio decimal(6,4) Null Default '0', "
        s_Sql = s_Sql & "imporcambio decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "PRIMARY KEY (codpsn, secuencia, codcpc)) "
      ElseIf ribAnalisis(3).Value Then
        s_Sql = s_Sql & "codpsn varchar(11) Not Null, nombrespsn varchar(80) Null, sigladci char(3) Null, numdociden varchar(11) Null, "
        s_Sql = s_Sql & "fecnacimiento date Null, codbco char(3) Null, desbco varchar(40) Null, monedacts char(3) Null, "
        s_Sql = s_Sql & "cuentacts varchar(20) Null, "
        s_Sql = s_Sql & "impormes_01 decimal(18,2) Null Default '0', impormes_02 decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "impormes_03 decimal(18,2) Null Default '0', impormes_04 decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "impormes_05 decimal(18,2) Null Default '0', impormes_06 decimal(18,2) Null Default '0', "
        s_Sql = s_Sql & "PRIMARY KEY (codpsn)) "
        aElemento(2, 3) = "Semestre;" & s_Estado_Blq & ";true"
      End If
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      If ribAnalisis(0).Value Then
        ResumenCts "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
      ElseIf ribAnalisis(1).Value Then
        DepositoCts "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
      ElseIf ribAnalisis(2).Value Then
        CertificadoCts "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
      ElseIf ribAnalisis(3).Value Then
        RemuneraCts "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
      End If
    End If
    ' Obtengo la información del reporte
    s_Sql = "SELECT rpt.*, cfg.logo, cfg.firma "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " rpt, plcfgempresa cfg "
    s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "ORDER BY " & IIf(ribAnalisis(1).Value, "codbco, ", "") & "codpsn"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 7), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  End Select

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0      ' Periodos de cts
    tdbHelp.Columns(0).DataField = "pdocts": tdbHelp.Columns(1).DataField = "descricts"
    tdbHelp.Caption = "Periodo de CTS"
    s_Sql = gdl_Funcion.HelpTablas("ced", "pdocts", s_Estado_Ina & ps_ClsPlanilla, "")
   Case 1       ' Sub periodo de cts
    tdbHelp.Columns(0).DataField = "subcts": tdbHelp.Columns(1).DataField = "descrisub"
    tdbHelp.Caption = "Sub periodo CTS"
    s_Sql = gdl_Funcion.HelpTablas("sed", "subcts", s_Estado_Ina & ps_ClsPlanilla & txtPeriodo(0).Text, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = panToolBar(1).Top + (Choose(Index + 1, cmdHelp(Index).Top, 750, cmdHelp(Index).Top, 750, 850) + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = panToolBar(1).Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub Form_Activate()
  ' Bloqueo la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()

  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6740: Me.Width = 8580
  Me.Left = 500: Me.Top = 150
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro

  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Trabajador(es)"
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Apellido y Nombres", "Fec.Ingreso", "Fec. Cese", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpsn", "nombrepsn", "fecingreso", "fecbaja", "estadopsn")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1000, 3832.66, 950, 950, 300)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", s_FormatoFecha, s_FormatoFecha, "")
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
    aElemento(n_Index, 1) = Choose(n_Index + 1, "promedio", "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Parametros", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 3
    ' Analisis
    ribAnalisis(n_Index).PictureUp = LoadPicture()
    ribAnalisis(n_Index).ToolTipText = Choose(n_Index + 1, "Analisis de Provisión", "Analisis de Depósito", "Boleta de Liquidación", "Analisis de Ingresos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "ancthist", "liquicts", "certifica", "repogene") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribAnalisis(n_Index).PictureUp = LoadPicture(s_Sql)
    ' Filtro
    If n_Index <> 3 Then
      ribParametro(n_Index).PictureUp = LoadPicture()
      ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
      s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
      If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
    End If
  Next n_Index
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  tdbRegistro.DataSource = dcaRegistro
  ribParametro(0).Value = True
  ribAnalisis(0).Value = True
  
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcts", "descricts")
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
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' Habilito la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = Not Cancel
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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0, 1      ' Periodo y sUb periodo de cts
    txtPeriodo(n_IndexHelp) = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo(n_IndexHelp).SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Periodo de cts
    s_Sql = gdl_Funcion.HelpTablas("ced", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla, "")
   Case 1     ' Sub periodo de cts
    s_Sql = gdl_Funcion.HelpTablas("sed", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla & Trim(txtPeriodo(0).Text), "")
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
Private Sub tdbRegistro_DblClick()
  cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
Private Sub txtPeriodo_GotFocus(Index As Integer)
  gdl_Procedure.MarcaGet txtPeriodo(Index)
End Sub
Private Sub txtPeriodo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click Index
End Sub
Private Sub txtPeriodo_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus(Index As Integer)
  If Index = 0 Then
    lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo(Index), "EC")
  Else
    lblHelp(Index) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, Trim(txtPeriodo(0).Text) & "|" & Trim(txtPeriodo(Index).Text), "SC")
  End If
End Sub
Private Sub txtPeriodo_Validate(Index As Integer, Cancel As Boolean)
  
  If Index = 0 Then
    lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, Trim(txtPeriodo(Index).Text) & "|" & Trim(txtPeriodo(1).Text), "SC")
    If Not (lblHelp(1) = "???" Or lblHelp(1) = "") Then Exit Sub
    txtPeriodo(1).Text = "": lblHelp(1) = ""
  End If

End Sub
