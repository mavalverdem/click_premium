VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 01"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "selperiodo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   11115
   Begin MSAdodcLib.Adodc dcaSeleccion 
      Height          =   330
      Index           =   3
      Left            =   45
      Top             =   5475
      Width           =   5340
      _ExtentX        =   9419
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
      Left            =   10320
      TabIndex        =   0
      Top             =   555
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
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "0"
         Top             =   1275
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
         Picture         =   "selperiodo.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1695
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
         Picture         =   "selperiodo.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Tag             =   "0"
         Top             =   2415
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
         Picture         =   "selperiodo.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Tag             =   "0"
         Top             =   2835
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
         Picture         =   "selperiodo.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Tag             =   "0"
         Top             =   3975
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
         Picture         =   "selperiodo.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Tag             =   "0"
         Top             =   4395
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
         Picture         =   "selperiodo.frx":0098
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   0
         Left            =   15
         TabIndex        =   11
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Tag             =   "0"
         Top             =   855
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
         Picture         =   "selperiodo.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Tag             =   "0"
         Top             =   3255
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
         Picture         =   "selperiodo.frx":00D0
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   1
         Left            =   15
         TabIndex        =   12
         Top             =   285
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Otros"
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
         Index           =   8
         Left            =   300
         TabIndex        =   9
         Tag             =   "1"
         Top             =   855
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
         Picture         =   "selperiodo.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   9
         Left            =   300
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1275
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
         Picture         =   "selperiodo.frx":0108
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11115
      _Version        =   65536
      _ExtentX        =   19606
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
      Begin Threed.SSRibbon ribSeccion 
         Height          =   360
         Index           =   0
         Left            =   9315
         TabIndex        =   15
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   2
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selperiodo.frx":0124
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   1110
         TabIndex        =   23
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selperiodo.frx":0140
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   705
         TabIndex        =   22
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selperiodo.frx":015C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   1515
         TabIndex        =   24
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selperiodo.frx":0178
      End
      Begin Threed.SSRibbon ribOrdenar 
         Height          =   360
         Left            =   5460
         TabIndex        =   25
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   0
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selperiodo.frx":0194
      End
      Begin Threed.SSRibbon ribSeccion 
         Height          =   360
         Index           =   1
         Left            =   9720
         TabIndex        =   16
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   2
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selperiodo.frx":01B0
      End
      Begin Threed.SSRibbon ribSeccion 
         Height          =   360
         Index           =   2
         Left            =   10125
         TabIndex        =   17
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupNumber     =   2
         GroupAllowAllUp =   -1  'True
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selperiodo.frx":01CC
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   5235
      Left            =   5460
      TabIndex        =   18
      Top             =   570
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   9234
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   494
      TabMaxWidth     =   2381
      BackColor       =   16777215
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
      TabCaption(0)   =   "Centro Costo"
      TabPicture(0)   =   "selperiodo.frx":01E8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dcaSeleccion(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdbSeleccion(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ubicación"
      TabPicture(1)   =   "selperiodo.frx":0204
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dcaSeleccion(1)"
      Tab(1).Control(1)=   "tdbSeleccion(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Sección"
      TabPicture(2)   =   "selperiodo.frx":0220
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dcaSeleccion(2)"
      Tab(2).Control(1)=   "tdbSeleccion(2)"
      Tab(2).ControlCount=   2
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   4425
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   90
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   7805
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
      Begin MSAdodcLib.Adodc dcaSeleccion 
         Height          =   330
         Index           =   0
         Left            =   60
         Top             =   4545
         Width           =   4665
         _ExtentX        =   8229
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
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   4425
         Index           =   1
         Left            =   -74940
         TabIndex        =   20
         Top             =   90
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   7805
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
      Begin MSAdodcLib.Adodc dcaSeleccion 
         Height          =   330
         Index           =   1
         Left            =   -74940
         Top             =   4545
         Width           =   4665
         _ExtentX        =   8229
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
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   4425
         Index           =   2
         Left            =   -74940
         TabIndex        =   21
         Top             =   90
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   7805
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
      Begin MSAdodcLib.Adodc dcaSeleccion 
         Height          =   330
         Index           =   2
         Left            =   -74940
         Top             =   4545
         Width           =   4665
         _ExtentX        =   8229
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
   End
   Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
      Height          =   4845
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   570
      Width           =   5340
      _ExtentX        =   9419
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
End
Attribute VB_Name = "fSelPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(4, 2)                            ' Array de inicio y fin de seleccion de registro
Private s_OptRegistro As String                         ' Instancia del formulario activo
Dim cnn As ADODB.Connection
'[
Private Sub ppExcelHorizontal(ByVal sTablaTmp As String, ByVal sPeriodo As String, ByVal sTitulo As String)
  Dim sHojaExcel As String, s_OldMessage As String
  Dim nFila As Long, nColumna As Long
  Dim nRegistro As Long, nSecuImporte As Long
  Dim nTipoSeccion As Integer, nColInicio As Integer
  Dim nNumColumna As Integer, nColx As Integer
  
  Dim aCodigoTra As Variant
  Dim poApExcel As Object

  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando Información Planilla de Trabajo Excel ..."
  
  ' Inicializo variables
  sHojaExcel = sPeriodo
  Set poApExcel = CreateObject("Excel.application")
  poApExcel.Visible = False
  poApExcel.Workbooks.Add
  poApExcel.Sheets("Hoja1").Name = Left(sHojaExcel, 20)
  poApExcel.ActiveWindow.Zoom = 75
  poApExcel.Cells(1, 1).Formula = sTitulo & " : " & sHojaExcel
  poApExcel.Cells(1, 1).Font.Size = 18
  
  TipodeProgreso = 1
  ' Información de trabajadores
  nRegistro = 0
  s_Sql = "SELECT codpsn, nombrepsn, detcco, codafp, desafp, descgo, fecingreso, AVG(dias) dias "
  s_Sql = s_Sql & "FROM " & sTablaTmp & " "
  s_Sql = s_Sql & "GROUP BY codpsn "
  s_Sql = s_Sql & "HAVING SUM(importe01+importe02+importe03+importe04+importe05+importe06+importe07+importe08+"
  s_Sql = s_Sql & "importe09+importe10+importe11+importe12+importe13+importe14+importe15+importe16+importe17+importe18+importe19+"
  s_Sql = s_Sql & "importe20+importe21+importe22+importe23+importe24+importe25+importe26+importe27+importe28+importe29+importe30+"
  s_Sql = s_Sql & "importe31+importe32+importe33+importe34+importe35)>=0 "
  s_Sql = s_Sql & "ORDER BY codpsn"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  nNumColumna = 8
  On Error GoTo Finalizar
  IntervalodeTiempo = porstRecordset.RecordCount
  porstRecordset.MoveFirst
  nFila = 1
  ' Cabecera de campos
  For nColumna = 1 To nNumColumna
    poApExcel.Cells(nFila + 3, nColumna).Formula = porstRecordset.Fields(nColumna - 1).Name
  Next nColumna
  ' Copio detalle de trabajadores a excel
  nColInicio = 1
  nRegistro = porstRecordset.RecordCount
  poApExcel.Cells(nFila + 4, nColInicio).CopyFromRecordset porstRecordset
  ' Inicializo arreglo de codigo de trabajadores
  aCodigoTra = porstRecordset.GetRows(nRegistro, 0, 0)
  porstRecordset.Close
  
  nColumna = 1: nColInicio = 8
  For nTipoSeccion = 0 To 2
    ' Inserto informacion ingreso, descuento y aportes no existentes
    s_Sql = "INSERT INTO " & sTablaTmp & "(seccion, codcco, codsec, codpsn, secuencia, nombrepsn, numdociden, detcco, dessec, codafp, "
    s_Sql = s_Sql & "desafp, descgo, fecingreso, dias) "
    s_Sql = s_Sql & "SELECT DISTINCT '" & nTipoSeccion & "' seccion, rxn.codcco, rxn.codsec, rxn.codpsn, rxn.secuencia, rxn.nombrepsn, "
    s_Sql = s_Sql & "rxn.numdociden, rxn.detcco, rxn.dessec, rxn.codafp, rxn.desafp, rxn.descgo, rxn.fecingreso, AVG(rxn.dias) dias "
    s_Sql = s_Sql & "FROM " & sTablaTmp & " rxn "
    s_Sql = s_Sql & "WHERE NOT EXISTS(SELECT * FROM tmprptpreplanilla tmp "
    s_Sql = s_Sql & "WHERE tmp.codcco=rxn.codcco AND tmp.codsec=rxn.codsec "
    s_Sql = s_Sql & "AND tmp.codpsn=rxn.codpsn AND tmp.seccion='" & nTipoSeccion & "') "
    s_Sql = s_Sql & "GROUP BY rxn.codcco, rxn.codsec, rxn.codpsn, rxn.nombrepsn, rxn.numdociden, rxn.detcco, rxn.dessec, rxn.codafp, rxn.desafp, rxn.descgo, rxn.fecingreso "
    s_Sql = s_Sql & "ORDER BY codpsn, secuencia"
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    ' Cabacera de conceptos
    s_Sql = "SELECT DISTINCT alias01, alias02, alias03, alias04, alias05, alias06, alias07, alias08, alias09, alias10, alias11, alias12, "
    s_Sql = s_Sql & "alias13, alias14, alias15, alias16, alias17, alias18, alias19, alias20, alias21, alias22, alias23, alias24, alias25, "
    s_Sql = s_Sql & "alias26, alias27, alias28, alias29, alias30, alias31, alias32, alias33, alias34, alias35 "
    s_Sql = s_Sql & "FROM " & sTablaTmp & " "
    s_Sql = s_Sql & "WHERE seccion='" & nTipoSeccion & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    porstRecordset.MoveFirst
    
    nColInicio = poApExcel.Cells(nFila + 3, nColumna).End(xlToRight).Column
    ' Copio cabecera de conceptos a excel
    poApExcel.Cells(nFila + 3, nColInicio + nColumna).CopyFromRecordset porstRecordset
    porstRecordset.Close
    
    nNumColumna = poApExcel.Cells(nFila + 3, nColInicio).End(xlToRight).Column
    nColx = nNumColumna - nColInicio
    ' Importes de conceptos
    s_Sql = "SELECT importe01, importe02, importe03, importe04, importe05, importe06, importe07, importe08, importe09, importe10, importe11, importe12, importe13, importe14, "
    s_Sql = s_Sql & "importe15, importe16, importe17, importe18, importe19, importe20, importe21, importe22, importe23, importe24, importe25, importe26, importe27, importe28, "
    s_Sql = s_Sql & "importe29, importe30, importe31, importe32, importe33, importe34, importe35 "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "WHERE seccion='" & nTipoSeccion & "' "
    s_Sql = s_Sql & "AND ((importe01+importe02+importe03+importe04+importe05+importe06+importe07+importe08+"
    s_Sql = s_Sql & "importe09+importe10+importe11+importe12+importe13+importe14+importe15+importe16+importe17+importe18+importe19+"
    s_Sql = s_Sql & "importe20+importe21+importe22+importe23+importe24+importe25+importe26+importe27+importe28+importe29+importe30+"
    s_Sql = s_Sql & "importe31+importe32+importe33+importe34+importe35)>=0  or dias=0)"
    s_Sql = s_Sql & "ORDER BY codpsn, secuencia"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    porstRecordset.MoveFirst
    nRegistro = porstRecordset.RecordCount
    nColumna = 1
  ' Copio importes de conceptos a excel
    poApExcel.Cells(nFila + 4, nColInicio + nColumna).CopyFromRecordset porstRecordset, nRegistro, nColx
  Next nTipoSeccion
  ' Porcentaje de avance
  IntervalodeTiempo = nRegistro
  labelprogreso = "Procesando información de planilla de trabajo"
  Progreso.Show vbModal
  
  ' Formato titulos negrita y fondo azul
  nNumColumna = poApExcel.Cells(nFila + 3, nColumna).End(xlToRight).Column
  poApExcel.Range(poApExcel.Cells(nFila + 3, 1), poApExcel.Cells(nFila + 3, nNumColumna)).Select
  poApExcel.Selection.Font.Bold = True
  With poApExcel.Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent5
    .TintAndShade = 0.399975585192419
    .PatternTintAndShade = 0
  End With
  
  nColx = poApExcel.Cells(nFila + 3, nColumna).End(xlDown).Row
  nColInicio = 8
  ' Formato contabilidad importes
  poApExcel.Cells(nFila + 3, nColInicio + nColumna).Select
  poApExcel.Range(poApExcel.Cells(nFila + 4, nColInicio + nColumna), poApExcel.Cells(nColx + 1, nNumColumna)).Select
  poApExcel.Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
  poApExcel.Cells(nFila + 3, nColumna).Select
  'poApExcel.Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
  
  ' Formato de totales
  poApExcel.Cells(nColx + 1, nColInicio + nColumna).FormulaR1C1 = "=SUM(R[-" & nRegistro & "]C:R[-1]C)"
  poApExcel.Cells(nColx + 1, nColInicio + nColumna).Font.Bold = True
  poApExcel.Cells(nColx + 1, nColInicio + nColumna).Copy (poApExcel.Range(poApExcel.Cells(nColx + 1, nColInicio + 2), poApExcel.Cells(nColx + 1, nNumColumna)))
  
  'Elimino detalle adicional
  nColInicio = poApExcel.Cells(nFila + 3, nColumna).End(xlToRight).Column
  nNumColumna = poApExcel.Cells(nFila + 4, nColumna).End(xlToRight).Column
  poApExcel.Range(poApExcel.Cells(nFila + 4, nColInicio + nColumna), poApExcel.Cells(nColx + 1, nNumColumna)).ClearContents
  
  MsgBox ("Proceso de Exportacion a Excel, Finalizado")
  poApExcel.Visible = True

Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub ppExcelVertical(ByVal sPeriodo As String, ByVal sProceso As String, ByVal sFechaHora As String)
  Dim sHojaExcel As String
  Dim ApExcelver As Variant
  Dim ia As Long
    
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroEnEspera
  
  sHojaExcel = Left(sPeriodo, 15)
  s_Sql = "SELECT concat('', res.codpsn) AS Codigo, psn.numdociden AS Documento, concat('', DATE_FORMAT(psn.fecingreso,'%d/%m/%Y')) AS FechaIngreso, "
  s_Sql = s_Sql & "concat(psn.apepaterno,' ', psn.apematerno,' ', psn.nombres) AS Nombre, case res.tipocpc when 0 then '1Ingresos' when 1 then '2Descuentos' else '3Aportaciones' end AS Tipo, "
  s_Sql = s_Sql & "concat(res.codcpc,'-', cpc.descpc) AS Concepto, res.importe_mn AS ImporteMN, res.importe_me AS importeME "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON psn.codcls=res.codcls AND psn.codpsn=res.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo IN (SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & sProceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & sFechaHora & "') "
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' AND res.tipocpc IN(0, 1, 2) "
  s_Sql = s_Sql & "ORDER BY res.codpsn, res.codpdo, res.tipocpc, res.codcpc "
  cols = 8
  
  Set ApExcelver = CreateObject("Excel.application")
  ApExcelver.Visible = False
  ApExcelver.Workbooks.Add
  ApExcelver.Sheets("Hoja1").Name = sHojaExcel
  
  ApExcelver.ActiveWindow.Zoom = 75
  ApExcelver.Cells(1, 1).Formula = "Informacion del Trabajador : " & sPeriodo
  ApExcelver.Cells(1, 1).Font.Size = 18
  ApExcelver.Cells(2, 1).Formula = ""
  '************************************
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  On Error GoTo Finalizar
  'IntervalodeTiempo = porstRecordset.RecordCount
  porstRecordset.MoveFirst
  For ia = 1 To porstRecordset.RecordCount
    If ia = 1 Then
      For j = 1 To cols
        ApExcelver.Cells(ia + 3, j).Formula = porstRecordset.Fields(j - 1).Name
      Next j
    End If
    For j = 1 To cols
      ApExcelver.Cells(ia + 4, j).Formula = porstRecordset(j - 1)
    Next j
    porstRecordset.MoveNext
  Next ia
  TipodeProgreso = 1
  IntervalodeTiempo = 100
  labelprogreso = "Exportando Datos a Excel"
  Progreso.Show vbModal
  
  MsgBox ("Proceso de Exportacion a Excel, terminado")
  ApExcelver.Visible = True

Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
End Sub
Private Sub ppPrePlanilla(nTabIndex As Integer, s_Tabla As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)
  Dim sCamSeccion As String, sSeccion As String, sDesSeccion As String
  Dim sCamRubro As String, sRubro As String, sDesRubro As String
  Dim sSentenciaIni As String, sSentenciaFin As String
  Dim sGrupo As String, sQuiebre As String, sPersonal As String
  Dim s_OldMessage As String
  Dim a_Ingreso(), a_Descuento(), a_Aporte(), a_Registro()
  Dim nNetoPagar As Double
  Dim nRegistro As Long, nRegistros As Long, nSecuencia As Long
  Dim nNivel As Integer, nGrupo As Integer
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando Información Planilla de Trabajo ..."
  
  nGrupo = IIf(Not ribSeccion(0).Value, s_Estado_Act, s_Estado_Ina)
  ' Agrupacion default
  sGrupo = "XXXXXXXXX"
  sCamRubro = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codcco", "codcco", "codcco", "codubica", "codubica", "codsec", "codsec")
  sDesRubro = Choose(nTabIndex + 1, "detcco", "desubica", "dessec", "detcco", "detcco", "detcco", "desubica", "desubica", "dessec", "dessec")
  sCamSeccion = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codcco", "codubica", "codsec", "codcco", "codsec", "codcco", "codubica")
  sDesSeccion = Choose(nTabIndex + 1, "detcco", "desubica", "dessec", "detcco", "desubica", "dessec", "detcco", "dessec", "detcco", "desubica")
  
  ' Genero las cabecera de los conceptos
  s_Sql = "SELECT DISTINCTROW res.tipocpc, res.codcpc, cpc.aliascpc, res.secuencia "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 0, 1, 2, 0, 0, 0, 1, 1, 2, 2) & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  If nTabIndex >= 4 Then
    s_Sql = s_Sql & "AND dxr." & sCamSeccion & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 2, 2, 2, 2, 1, 2, 0, 2, 0, 1) & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.tipocpc, res.codcpc "
  s_Sql = s_Sql & "ORDER BY tipocpc, secuencia, codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    n_Index = 0
    ' Dimensiones de arreglos e inicializo totales
    ReDim a_Ingreso(3, 0), a_Descuento(3, 0), a_Aporte(3, 0)
    ' Ingresos
    a_Ingreso(1, n_Index) = "TING"
    a_Ingreso(2, n_Index) = "T INGRESOS"
    a_Ingreso(3, n_Index) = CDbl(0)
    ' Descuentos
    a_Descuento(1, n_Index) = "TDSC"
    a_Descuento(2, n_Index) = "T DESCTO"
    a_Descuento(3, n_Index) = CDbl(0)
    ' Aportes
    a_Aporte(1, n_Index) = "TAPO"
    a_Aporte(2, n_Index) = "T APORTES"
    a_Aporte(3, n_Index) = CDbl(0)
    While Not porstRecordset.EOF
      ' Redimensiono el arreglo de cabeceras
      If porstRecordset("tipocpc") = "0" Then
        n_Index = UBound(a_Ingreso, 2) + 1
        ReDim Preserve a_Ingreso(3, n_Index)
        a_Ingreso(1, n_Index) = porstRecordset("codcpc")
        a_Ingreso(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Ingreso(3, n_Index) = CDbl(0)
      ElseIf porstRecordset("tipocpc") = "1" Then
        n_Index = UBound(a_Descuento, 2) + 1
        ReDim Preserve a_Descuento(3, n_Index)
        a_Descuento(1, n_Index) = porstRecordset("codcpc")
        a_Descuento(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Descuento(3, n_Index) = CDbl(0)
      ElseIf porstRecordset("tipocpc") = s_Estado_Blq Then
        n_Index = UBound(a_Aporte, 2) + 1
        ReDim Preserve a_Aporte(3, n_Index)
        a_Aporte(1, n_Index) = porstRecordset("codcpc")
        a_Aporte(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Aporte(3, n_Index) = CDbl(0)
      End If
      porstRecordset.MoveNext
    Wend
  End If
  porstRecordset.Close
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' [ Elimino y genero temporal de total ingresos
  s_Sql = "DROP TABLE IF EXISTS tmpingresos"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS tmpingresos "
  s_Sql = s_Sql & "SELECT res.tipocpc, res.codpsn, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)), 2) AS importe "
  sSentenciaIni = "FROM plresultado res "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSentenciaIni = sSentenciaIni & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  sSentenciaIni = sSentenciaIni & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  sSentenciaFin = "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSentenciaFin = sSentenciaFin & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  sSentenciaFin = sSentenciaFin & "WHERE proceso='" & s_Proceso & "' "
  sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
  sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  If nTabIndex <> 3 Then
    sSentenciaFin = sSentenciaFin & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    sSentenciaFin = sSentenciaFin & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 0, 1, 2, 0, 0, 0, 1, 1, 2, 2) & "' "
    sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
    sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  End If
  If nTabIndex >= 4 Then
    sSentenciaFin = sSentenciaFin & "AND dxr." & sCamSeccion & " IN(SELECT valor FROM rangoimpresion "
    sSentenciaFin = sSentenciaFin & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 2, 2, 2, 2, 1, 2, 0, 2, 0, 1) & "' "
    sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
    sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "GROUP BY res.tipocpc, " & Choose(nTabIndex + 1, "dxr.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "", "dxr.codcco, dxr.codubica, ", "dxr.codcco, dxr.codsec, ", "dxr.codubica, dxr.codcco, ", "dxr.codubica, dxr.codsec, ", "dxr.codsec, dxr.codcco, ", "dxr.codsec, dxr.codubica, ") & "res.codpsn "
  s_Sql = s_Sql & "ORDER BY tipocpc, codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Verifico si existen descuentos
  s_Sql = "SELECT IFNULL(COUNT(*), 0) AS registros "
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & "LEFT JOIN tmpingresos tmp ON res.codpsn=tmp.codpsn "
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Act & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  nRegistros = 0
  If Not (porstRecordset.EOF And porstRecordset.BOF) Then
    nRegistros = CLng(porstRecordset!registros)
  End If
  porstRecordset.Close
  
  ' [ Elimino y genero temporal de detalle
  s_Sql = "DROP TABLE IF EXISTS tmpdetalle"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdetalle "
  s_Sql = s_Sql & "SELECT res.tipocpc, res.codpsn, psn.numdociden, "
  s_Sql = s_Sql & "CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', CONCAT_WS(' ', psn.nombres, date_format(fecbaja,'%d/%m/%y'))) AS nombrepsn, "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxr.codcco, cco.detcco, ", "dxc.codcco, cco.detcco, ", "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, ")
  s_Sql = s_Sql & Choose(nTabIndex + 1, "dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, ", "dxr.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxc.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, ", "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ")
  s_Sql = s_Sql & "dxr.codafp, afp.desafp, cgo.descgo, res.codcpc, cpc.aliascpc, MAX(dxr.fecingreso) AS fecingreso, "
  s_Sql = s_Sql & "SUM(IFNULL((asi.diatrabajo+asi.diamediotm+asi.diaparcial)" & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8), "*(dxc.porcentaje/100)", "") & ", 0)) AS dias, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8), "*(dxc.porcentaje/100)", "") & ", 0)), 2) AS importe, "
  s_Sql = s_Sql & "IFNULL(tmp.importe" & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8), "*(dxc.porcentaje/100)", "") & ", 0.00) AS ingresos, res.secuencia "
  sSentenciaIni = "FROM plresultado res "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSentenciaIni = sSentenciaIni & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  If (nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8) Then
    sSentenciaIni = sSentenciaIni & "INNER JOIN plcencospro dxc ON dxc.codcls=dxr.codcls AND dxc.codpdo=dxr.codpdo AND dxc.codpsn=dxr.codpsn "
    sSentenciaFin = Replace(sSentenciaFin, "dxr.codcco", "dxc.codcco")
  End If
  sSentenciaIni = sSentenciaIni & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON " & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8), "dxc", "dxr") & ".codcco=cco.codcco "
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & "LEFT JOIN tmpingresos tmp ON res.codpsn=tmp.codpsn "
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "GROUP BY res.tipocpc, " & Choose(nTabIndex + 1, "dxc.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "", "dxc.codcco, dxr.codubica, ", "dxc.codcco, dxr.codsec, ", "dxr.codubica, dxc.codcco, ", "dxr.codubica, dxr.codsec, ", "dxr.codsec, dxc.codcco, ", "dxr.codsec, dxr.codubica, ") & "res.codpsn, res.codcpc "
  s_Sql = s_Sql & "ORDER BY tipocpc, " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "", "codcco, codubica, ", "codcco, codsec, ", "codubica, codcco, ", "codubica, codsec, ", "codsec, codcco, ", "codsec, codubica, ") & "codpsn, secuencia"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ' Registros que no tienen descuento
  If nRegistros > 0 Then
    s_Sql = "INSERT INTO tmpdetalle "
    s_Sql = s_Sql & "SELECT '" & s_Estado_Act & "' AS tipocpc, psn.codpsn, psn.numdociden, "
    s_Sql = s_Sql & "CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', CONCAT_WS(' ', psn.nombres, date_format(fecbaja,'%d/%m/%y'))) AS nombrepsn, "
    s_Sql = s_Sql & Choose(nTabIndex + 1, "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxr.codcco, cco.detcco, ", "dxc.codcco, cco.detcco, ", "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, ")
    s_Sql = s_Sql & Choose(nTabIndex + 1, "dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, ", "dxr.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxc.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, ", "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ")
    s_Sql = s_Sql & "dxr.codafp, afp.desafp, cgo.descgo, 'dsct' AS codcpc, 'dscto' AS aliascpc, MAX(dxr.fecingreso) AS fecingreso, "
    s_Sql = s_Sql & "SUM(IFNULL((asi.diatrabajo+asi.diamediotm+asi.diaparcial)" & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8), "*(dxc.porcentaje/100)", "") & ", 0)) AS dias, "
    s_Sql = s_Sql & "0 AS importe, IFNULL(tmp.importe" & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8), "*(dxc.porcentaje/100)", "") & ", 0.00) AS ingresos, 9 AS secuencia "
    s_Sql = s_Sql & "FROM pldatoresultado dxr "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON dxr.codcls=psn.codcls AND dxr.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plasistencia asi ON dxr.codcls=asi.codcls AND dxr.codpdo=asi.codpdo AND dxr.codpsn=asi.codpsn "
    s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
    s_Sql = s_Sql & "INNER JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
    s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
    s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
    If (nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8) Then
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=dxr.codcls AND dxc.codpdo=dxr.codpdo AND dxc.codpsn=dxr.codpsn "
    End If
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON " & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5 Or nTabIndex = 6 Or nTabIndex = 8), "dxc", "dxr") & ".codcco=cco.codcco "
    s_Sql = s_Sql & "INNER JOIN tmpingresos tmp ON dxr.codpsn=tmp.codpsn "
    s_Sql = s_Sql & "WHERE dxr.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND dxr.codpdo IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    If nTabIndex <> 3 Then
      s_Sql = s_Sql & "AND " & IIf((nTabIndex = 0 Or nTabIndex = 4 Or nTabIndex = 5), "dxc.", "dxr.") & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 0, 1, 2, 0, 0, 0, 1, 1, 2, 2) & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    End If
    If nTabIndex >= 4 Then
      s_Sql = s_Sql & "AND " & IIf((nTabIndex = 6 Or nTabIndex = 8), "dxc.", "dxr.") & sCamSeccion & " IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 2, 2, 2, 2, 1, 2, 0, 2, 0, 1) & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    End If
    s_Sql = s_Sql & "AND NOT EXISTS (SELECT * FROM plresultado res "
    s_Sql = s_Sql & "WHERE res.codcls=dxr.codcls "
    s_Sql = s_Sql & "AND res.codpdo=dxr.codpdo "
    s_Sql = s_Sql & "AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Act & "') "
    s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "dxc.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "", "dxc.codcco, dxr.codubica, ", "dxc.codcco, dxr.codsec, ", "dxr.codubica, dxc.codcco, ", "dxr.codubica, dxr.codsec, ", "dxr.codsec, dxc.codcco, ", "dxr.codsec, dxr.codubica, ") & "dxr.codpsn "
    s_Sql = s_Sql & "ORDER BY tipocpc, " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "", "codcco, codubica, ", "codcco, codsec, ", "codubica, codcco, ", "codubica, codsec, ", "codsec, codcco, ", "codsec, codubica, ") & "codpsn, secuencia"
    If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  End If
  ' Seleciono la información  detallada
  s_Sql = "SELECT * FROM tmpdetalle "
  s_Sql = s_Sql & "ORDER BY tipocpc, " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "", "codcco, codubica, ", "codcco, codsec, ", "codubica, codcco, ", "codubica, codsec, ", "codsec, codcco, ", "codsec, codubica, ") & "codpsn, secuencia"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Mensaje de proceso
  MuestraMensaje "Procesando Planilla Trabajo ..."
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    
    nNivel = UBound(a_Ingreso, 2)
    nNivel = IIf(UBound(a_Descuento, 2) > nNivel, UBound(a_Descuento, 2), nNivel)
    nNivel = IIf(UBound(a_Aporte, 2) > nNivel, UBound(a_Aporte, 2), nNivel)
    ' Genero los arreglos de la grabación
    a_Campos = Array("seccion", "codcco", "codsec", "codpsn", "secuencia", "nombrepsn", "numdociden", "detcco", "dessec", "codafp", "desafp", "descgo", "fecingreso", "dias", _
     "alias01", "alias02", "alias03", "alias04", "alias05", "alias06", "alias07", "alias08", "alias09", "alias10", "alias11", "alias12", "alias13", "alias14", "alias15", "alias16", "alias17", "alias18", "alias19", "alias20", _
     "alias21", "alias22", "alias23", "alias24", "alias25", "alias26", "alias27", "alias28", "alias29", "alias30", "alias31", "alias32", "alias33", "alias34", "alias35", "importe01", "importe02", "importe03", "importe04", "importe05", "importe06", "importe07", "importe08", "importe09", "importe10", _
     "importe11", "importe12", "importe13", "importe14", "importe15", "importe16", "importe17", "importe18", "importe19", "importe20", "importe21", "importe22", "importe23", "importe24", "importe25", "importe26", "importe27", "importe28", "importe29", "importe30", "importe31", "importe32", "importe33", "importe34", "importe35")
    a_Valores = Array("", "", "", "", "", "", "", "", "", "", "", "", "", CDec(0), _
     "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
     "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), _
     CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Numero, _
     TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
     TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, _
     TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    nSecuencia = 0
    
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    While Not porstRecordset.EOF
      ' Inicialización de variables
      nSecuencia = nSecuencia + 1
      sQuiebre = porstRecordset("tipocpc")
      sRubro = porstRecordset(sCamRubro)
      sSeccion = porstRecordset(sCamSeccion)
      sPersonal = porstRecordset("codpsn")
      ' Inicializo los importes del detalle
      If sQuiebre = "0" Then
        a_Registro = a_Ingreso
      ElseIf sQuiebre = "1" Then
        a_Registro = a_Descuento
      ElseIf sQuiebre = "2" Then
        a_Registro = a_Aporte
      End If
      For n_Index = 0 To UBound(a_Registro, 2): a_Registro(3, n_Index) = CDbl(0): Next n_Index
      Do
        For n_Index = 1 To UBound(a_Registro, 2)
          If a_Registro(1, n_Index) = porstRecordset("codcpc") Then Exit For
        Next n_Index
        n_Index = IIf(n_Index > UBound(a_Registro, 2), UBound(a_Registro, 2), n_Index)
        a_Registro(3, n_Index) = CDec(porstRecordset("importe"))
        ' Totalizo detalle
        a_Registro(3, 0) = a_Registro(3, 0) + a_Registro(3, n_Index)
        
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        porstRecordset.MoveNext
        If porstRecordset.EOF Then Exit Do
      Loop While (sQuiebre = porstRecordset("tipocpc") And sRubro = porstRecordset(sCamRubro) And sSeccion = porstRecordset(sCamSeccion) And sPersonal = porstRecordset("codpsn"))
      porstRecordset.MovePrevious
      
      ' Inicializo valores del archivo temporal
      For n_Index = 1 To 35: a_Valores(13 + n_Index) = "": a_Valores(48 + n_Index) = CDec(0): Next n_Index
      
      ' Valores del archivo temporal
      a_Valores(0) = sQuiebre
      a_Valores(1) = Choose(nGrupo + 1, sGrupo, sRubro)
      a_Valores(2) = Choose(nGrupo + 1, sGrupo, sSeccion)
      a_Valores(3) = sPersonal
      a_Valores(4) = nSecuencia
      a_Valores(5) = gdl_Funcion.aTexto(porstRecordset("nombrepsn"))
      a_Valores(6) = gdl_Funcion.aTexto(porstRecordset("numdociden"))
      a_Valores(7) = Trim(porstRecordset(sDesRubro))
      a_Valores(8) = Trim(porstRecordset(sDesSeccion))
      a_Valores(9) = gdl_Funcion.aTexto(porstRecordset("codafp"))
      a_Valores(10) = gdl_Funcion.aTexto(porstRecordset("desafp"))
      a_Valores(11) = gdl_Funcion.aTexto(porstRecordset("descgo"))
      a_Valores(12) = Format(porstRecordset("fecingreso"), s_FmtFechMysql_0)
      a_Valores(13) = CLng(porstRecordset("dias"))
      ' Conceptos de acuerdo al tipo de concepto
      For n_Index = 1 To UBound(a_Registro, 2)
        a_Valores(13 + n_Index) = a_Registro(2, n_Index)
        a_Valores(48 + n_Index) = CDec(a_Registro(3, n_Index))
      Next n_Index
      ' Total de acuerdo al tipo de concepto
      a_Valores(13 + n_Index) = a_Registro(2, 0)
      a_Valores(48 + n_Index) = CDec(a_Registro(3, 0))
      
      ' Neto a Pagar
      If sQuiebre = "1" Then
        a_Valores(14 + n_Index) = "NETO PAGAR"
        a_Valores(49 + n_Index) = CDec(porstRecordset("ingresos")) - CDec(a_Registro(3, 0))
      End If
      If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      porstRecordset.MoveNext
    Wend
    porstRecordset.Close
    
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' [ Elimino y genero temporal de total ingresos, detalle
  s_Sql = "DROP TABLE IF EXISTS tmpingresos"
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  s_Sql = "DROP TABLE IF EXISTS tmpdetalle"
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub ppPreRemuneracion(nTabIndex As Integer, s_Tabla As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)

  Dim nNivel As Integer, nSecuencia As Long
  Dim a_Ingreso(), a_Descuento(), a_Aporte()
  Dim sQuiebre As String, sPersonal As String
  Dim sCamRubro As String, sRubro As String, sDesRubro As String
  Dim a_Registro(), nNetoPagar As Double
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sGrupo As String, nGrupo As Integer
  Dim sSentenciaIni As String, sSentenciaFin As String

  ' Agrupacion default
  nGrupo = IIf(Not ribSeccion(0).Value, s_Estado_Act, s_Estado_Ina)
  sGrupo = "XXXXX"
  sCamRubro = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codcco")
  sDesRubro = Choose(nTabIndex + 1, "detcco", "desubica", "dessec", "detcco")
  
  ' Genero las cabecera de los conceptos
  s_Sql = "SELECT DISTINCTROW res.tipocpc, res.codcpc, cpc.aliascpc, res.secuencia "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.tipocpc, res.codcpc "
  s_Sql = s_Sql & "ORDER BY tipocpc, secuencia, codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    n_Index = 0
    ' Dimensiones de arreglos e inicializo totales
    ReDim a_Ingreso(3, 0), a_Descuento(3, 0), a_Aporte(3, 0)
    ' Ingresos
    a_Ingreso(1, n_Index) = "TING"
    a_Ingreso(2, n_Index) = "T INGRESOS"
    a_Ingreso(3, n_Index) = CDbl(0)
    ' Descuentos
    a_Descuento(1, n_Index) = "TDSC"
    a_Descuento(2, n_Index) = "T DESCTO"
    a_Descuento(3, n_Index) = CDbl(0)
    ' Aportes
    a_Aporte(1, n_Index) = "TAPO"
    a_Aporte(2, n_Index) = "T APORTES"
    a_Aporte(3, n_Index) = CDbl(0)
    While Not porstRecordset.EOF
      ' Redimensiono el arreglo de cabeceras
      If porstRecordset("tipocpc") = "0" Then
        n_Index = UBound(a_Ingreso, 2) + 1
        ReDim Preserve a_Ingreso(3, n_Index)
        a_Ingreso(1, n_Index) = porstRecordset("codcpc")
        a_Ingreso(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Ingreso(3, n_Index) = CDbl(0)
      ElseIf porstRecordset("tipocpc") = "1" Then
        n_Index = UBound(a_Descuento, 2) + 1
        ReDim Preserve a_Descuento(3, n_Index)
        a_Descuento(1, n_Index) = porstRecordset("codcpc")
        a_Descuento(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Descuento(3, n_Index) = CDbl(0)
      ElseIf porstRecordset("tipocpc") = s_Estado_Blq Then
        n_Index = UBound(a_Aporte, 2) + 1
        ReDim Preserve a_Aporte(3, n_Index)
        a_Aporte(1, n_Index) = porstRecordset("codcpc")
        a_Aporte(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Aporte(3, n_Index) = CDbl(0)
      End If
      porstRecordset.MoveNext
    Wend
  End If
  porstRecordset.Close
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' [ Elimino y genero temporal de total ingresos
  s_Sql = "DROP TABLE IF EXISTS tmpingresos"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS tmpingresos "
  s_Sql = s_Sql & "SELECT res.tipocpc, res.codpsn, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)), 2) AS importe "
  sSentenciaIni = "FROM plresultado res "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSentenciaIni = sSentenciaIni & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  sSentenciaIni = sSentenciaIni & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  sSentenciaFin = "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSentenciaFin = sSentenciaFin & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  sSentenciaFin = sSentenciaFin & "WHERE proceso='" & s_Proceso & "' "
  sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
  sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  If nTabIndex <> 3 Then
    sSentenciaFin = sSentenciaFin & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    sSentenciaFin = sSentenciaFin & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
    sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "GROUP BY res.tipocpc, " & Choose(nTabIndex + 1, "dxr.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "") & "res.codpsn "
  s_Sql = s_Sql & "ORDER BY tipocpc, codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Verifico si existen descuentos
  s_Sql = "SELECT IFNULL(COUNT(*), 0) AS registros "
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & "LEFT JOIN tmpingresos tmp ON res.codpsn=tmp.codpsn "
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Act & "'"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  nRegistros = CLng(porstRecordset!registros)
  
  ' [ Elimino y genero temporal de detalle
  s_Sql = "DROP TABLE IF EXISTS tmpdetalle"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS tmpdetalle "
  s_Sql = s_Sql & "SELECT res.tipocpc, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, ''),' ',ifnull(date_format(fecbaja,'%d/%m/%y'),'01/01/00')) AS nombrepsn, "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxr.codcco, cco.detcco, ")
  s_Sql = s_Sql & "dxr.codafp, afp.desafp, cgo.descgo, res.codcpc, cpc.aliascpc, psn.fecingreso, "
  s_Sql = s_Sql & "SUM(IFNULL(asi.diatrabajo" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)) AS dias, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)), 2) AS importe, "
  s_Sql = s_Sql & "IFNULL(tmp.importe" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0.00) AS ingresos, res.secuencia "
  sSentenciaIni = "FROM plresultado res "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSentenciaIni = sSentenciaIni & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  If nTabIndex = 0 Then
    sSentenciaIni = sSentenciaIni & "INNER JOIN plcencospro dxc ON dxc.codcls=dxr.codcls AND dxc.codpdo=dxr.codpdo AND dxc.codpsn=dxr.codpsn "
    sSentenciaFin = Replace(sSentenciaFin, "dxr.codcco", "dxc.codcco")
  End If
  sSentenciaIni = sSentenciaIni & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON " & IIf(nTabIndex = 0, "dxc", "dxr") & ".codcco=cco.codcco "
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & "LEFT JOIN tmpingresos tmp ON res.codpsn=tmp.codpsn "
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "GROUP BY res.tipocpc, " & Choose(nTabIndex + 1, "dxc.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "") & "res.codpsn, res.codcpc "
  s_Sql = s_Sql & "ORDER BY tipocpc, " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "") & "codpsn, secuencia"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ' Registros que no tienen descuento
  If nRegistros > 0 Then
    s_Sql = "INSERT INTO tmpdetalle "
    s_Sql = s_Sql & "SELECT '" & s_Estado_Act & "' AS tipocpc, psn.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, "
    s_Sql = s_Sql & Choose(nTabIndex + 1, "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "dxr.codcco, cco.detcco, ")
    s_Sql = s_Sql & "dxr.codafp, afp.desafp, cgo.descgo, 'dsct' AS codcpc, 'dscto' AS aliascpc, psn.fecingreso, "
    s_Sql = s_Sql & "SUM(IFNULL(asi.diatrabajo" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)) AS dias, "
    s_Sql = s_Sql & "0 AS importe, IFNULL(tmp.importe" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0.00) AS ingresos, 9 AS secuencia "
    s_Sql = s_Sql & "FROM pldatoresultado dxr "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON dxr.codcls=psn.codcls AND dxr.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plasistencia asi ON dxr.codcls=asi.codcls AND dxr.codpdo=asi.codpdo AND dxr.codpsn=asi.codpsn "
    s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
    s_Sql = s_Sql & "INNER JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
    s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
    s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
    If nTabIndex = 0 Then
      s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=dxr.codcls AND dxc.codpdo=dxr.codpdo AND dxc.codpsn=dxr.codpsn "
    End If
    s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON " & IIf(nTabIndex = 0, "dxc", "dxr") & ".codcco=cco.codcco "
    s_Sql = s_Sql & "INNER JOIN tmpingresos tmp ON dxr.codpsn=tmp.codpsn "
    s_Sql = s_Sql & "WHERE dxr.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND dxr.codpdo IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    If nTabIndex <> 3 Then
      s_Sql = s_Sql & "AND " & IIf(nTabIndex = 0, "dxc.", "dxr.") & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    End If
    s_Sql = s_Sql & "AND NOT EXISTS (SELECT * FROM plresultado res "
    s_Sql = s_Sql & "WHERE res.codcls=dxr.codcls "
    s_Sql = s_Sql & "AND res.codpdo=dxr.codpdo "
    s_Sql = s_Sql & "AND res.codpsn=dxr.codpsn "
    s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Act & "') "
    s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "dxc.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "") & "dxr.codpsn "
    s_Sql = s_Sql & "ORDER BY tipocpc, " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "") & "codpsn, secuencia"
    If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  End If
  ' Seleciono la información  detallada
  s_Sql = "SELECT * FROM tmpdetalle "
  s_Sql = s_Sql & "ORDER BY tipocpc, " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "") & "codpsn, secuencia"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Imprimiendo Planilla Trabajo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    
    nNivel = UBound(a_Ingreso, 2)
    nNivel = IIf(UBound(a_Descuento, 2) > nNivel, UBound(a_Descuento, 2), nNivel)
    nNivel = IIf(UBound(a_Aporte, 2) > nNivel, UBound(a_Aporte, 2), nNivel)
    ' Genero los arreglos de la grabación
    a_Campos = Array("seccion", "codcco", "codpsn", "secuencia", "nombrepsn", "detcco", "codafp", "desafp", "descgo", "fecingreso", "dias", _
     "alias01", "alias02", "alias03", "alias04", "alias05", "alias06", "alias07", "alias08", "alias09", "alias10", "alias11", "alias12", "alias13", "alias14", "alias15", "alias16", "alias17", "alias18", "alias19", "alias20", _
     "alias21", "alias22", "alias23", "alias24", "alias25", "alias26", "alias27", "alias28", "alias29", "alias30", "importe01", "importe02", "importe03", "importe04", "importe05", "importe06", "importe07", "importe08", "importe09", "importe10", _
     "importe11", "importe12", "importe13", "importe14", "importe15", "importe16", "importe17", "importe18", "importe19", "importe20", "importe21", "importe22", "importe23", "importe24", "importe25", "importe26", "importe27", "importe28", "importe29", "importe30")
    a_Valores = Array("", "", "", "", "", "", "", "", "", "", CDec(0), _
     "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
     "", "", "", "", "", "", "", "", "", "", CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), _
     CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Numero, _
     TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
     TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, _
     TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    nSecuencia = 0
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    
    While Not porstRecordset.EOF
      ' Inicialización de variables
      nSecuencia = nSecuencia + 1
      sQuiebre = porstRecordset("tipocpc")
      sRubro = porstRecordset(sCamRubro)
      sPersonal = porstRecordset("codpsn")
      ' Inicializo los importes del detalle
      If sQuiebre = "0" Then
        a_Registro = a_Ingreso
      ElseIf sQuiebre = "1" Then
        a_Registro = a_Descuento
      ElseIf sQuiebre = "2" Then
        a_Registro = a_Aporte
      End If
      For n_Index = 0 To UBound(a_Registro, 2): a_Registro(3, n_Index) = CDbl(0): Next n_Index
      Do
        For n_Index = 1 To UBound(a_Registro, 2)
          If a_Registro(1, n_Index) = porstRecordset("codcpc") Then Exit For
        Next n_Index
        n_Index = IIf(n_Index > UBound(a_Registro, 2), UBound(a_Registro, 2), n_Index)
        a_Registro(3, n_Index) = CDec(porstRecordset("importe"))
        ' Totalizo detalle
        a_Registro(3, 0) = a_Registro(3, 0) + a_Registro(3, n_Index)
        
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        porstRecordset.MoveNext
        If porstRecordset.EOF Then Exit Do
      Loop While (sQuiebre = porstRecordset("tipocpc") And sRubro = porstRecordset(sCamRubro) And sPersonal = porstRecordset("codpsn"))
      porstRecordset.MovePrevious
      ' Inicializo valores del archivo temporal
      For n_Index = 1 To 30: a_Valores(9 + n_Index) = "": a_Valores(39 + n_Index) = CDec(0): Next n_Index
      ' Valores del archivo temporal
      a_Valores(0) = sQuiebre
      a_Valores(1) = Choose(nGrupo + 1, sGrupo, sRubro)
      a_Valores(2) = sPersonal
      a_Valores(3) = nSecuencia
      a_Valores(4) = gdl_Funcion.aTexto(porstRecordset("nombrepsn"))
      a_Valores(5) = Trim(porstRecordset(sDesRubro))
      a_Valores(6) = gdl_Funcion.aTexto(porstRecordset("codafp"))
      a_Valores(7) = gdl_Funcion.aTexto(porstRecordset("desafp"))
      a_Valores(8) = gdl_Funcion.aTexto(porstRecordset("descgo"))
      a_Valores(9) = Format(porstRecordset("fecingreso"), s_FmtFechMysql_0)
      a_Valores(10) = CLng(porstRecordset("dias"))
      ' Conceptos de acuerdo al tipo de concepto
      For n_Index = 1 To UBound(a_Registro, 2)
        a_Valores(10 + n_Index) = a_Registro(2, n_Index)
        a_Valores(40 + n_Index) = CDec(a_Registro(3, n_Index))
      Next n_Index
      ' Total de acuerdo al tipo de concepto
      a_Valores(10 + n_Index) = a_Registro(2, 0)
      a_Valores(40 + n_Index) = CDec(a_Registro(3, 0))
      If sQuiebre = "1" Then
        'Neto a Pagar
        a_Valores(11 + n_Index) = "NETO PAGAR"
        a_Valores(41 + n_Index) = CDec(porstRecordset("ingresos")) - CDec(a_Registro(3, 0))
      End If
      If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      porstRecordset.MoveNext
    Wend
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' [ Elimino y genero temporal de total ingresos, detalle
  s_Sql = "DROP TABLE IF EXISTS tmpingresos"
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  s_Sql = "DROP TABLE IF EXISTS tmpdetalle"
  gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub ppPreResumen(nTabIndex As Integer, s_Tabla As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)
  Dim nNivel As Integer, nSecuencia As Long
  Dim a_Ingreso(), a_Descuento(), a_Aporte()
  Dim sCamRubro As String, sRubro As String, sDesRubro As String
  Dim sCamTabular As String, sTabular As String, sDesTabular As String
  Dim sDescripcion As String, sDescripcionTab As String, sQuiebre As String, nDias As Long
  Dim a_Registro(), nNetoPagar As Double, nIngresos As Double
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sSentenciaIni As String, sSentenciaFin As String

  ' Agrupacion default
  sCamTabular = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codpdo", "codcco", "codcco", "codubica", "codubica", "codsec", "codsec")
  sDesTabular = Choose(nTabIndex + 1, "detcco", "desubica", "dessec", "despdo", "detcco", "detcco", "desubica", "desubica", "dessec", "dessec")
  
  sCamRubro = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codpdo", "codubica", "codsec", "codcco", "codsec", "codcco", "codubica")
  sDesRubro = Choose(nTabIndex + 1, "detcco", "desubica", "dessec", "despdo", "desubica", "dessec", "detcco", "dessec", "detcco", "desubica")
  
  ' Genero las cabecera de los conceptos
  s_Sql = "SELECT DISTINCTROW res.tipocpc, res.codcpc, cpc.aliascpc, res.secuencia "
  sSentenciaIni = "FROM plresultado res "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  sSentenciaIni = sSentenciaIni & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  sSentenciaIni = sSentenciaIni & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  
  sSentenciaFin = "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  sSentenciaFin = sSentenciaFin & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  sSentenciaFin = sSentenciaFin & "WHERE proceso='" & s_Proceso & "' "
  sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
  sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  sSentenciaFin = sSentenciaFin & "AND res.impbolecpc='" & s_Estado_Act & "' "
  If nTabIndex <> 3 Then
   sSentenciaFin = sSentenciaFin & "AND dxr." & sCamTabular & " IN(SELECT valor FROM rangoimpresion "
   sSentenciaFin = sSentenciaFin & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 0, 1, 2, 0, 0, 0, 1, 1, 2, 2) & "' "
   sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
   sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  End If
  
  If nTabIndex >= 4 Then
   sSentenciaFin = sSentenciaFin & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
   sSentenciaFin = sSentenciaFin & "WHERE proceso='" & Left(s_Proceso, 9) & Choose(nTabIndex + 1, 2, 2, 2, 2, 1, 2, 0, 2, 0, 1) & "' "
   sSentenciaFin = sSentenciaFin & "AND usrcre='" & ps_Usuario & "' "
   sSentenciaFin = sSentenciaFin & "AND fyhcre='" & s_FechaHora & "') "
  End If
  
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "GROUP BY res.tipocpc, res.codcpc "
  s_Sql = s_Sql & "ORDER BY tipocpc, secuencia, codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    n_Index = 0
    ' Dimensiones de arreglos e inicializo totales
    ReDim a_Ingreso(3, 0), a_Descuento(3, 0), a_Aporte(3, 0)
    ' Ingresos
    a_Ingreso(1, n_Index) = "TING"
    a_Ingreso(2, n_Index) = "T INGRESOS"
    a_Ingreso(3, n_Index) = CDbl(0)
    ' Descuentos
    a_Descuento(1, n_Index) = "TDSC"
    a_Descuento(2, n_Index) = "T DESCTO"
    a_Descuento(3, n_Index) = CDbl(0)
    ' Aportes
    a_Aporte(1, n_Index) = "TAPO"
    a_Aporte(2, n_Index) = "T APORTES"
    a_Aporte(3, n_Index) = CDbl(0)
    While Not porstRecordset.EOF
      ' Redimensiono el arreglo de cabeceras
      If porstRecordset("tipocpc") = "0" Then
        n_Index = UBound(a_Ingreso, 2) + 1
        ReDim Preserve a_Ingreso(3, n_Index)
        a_Ingreso(1, n_Index) = porstRecordset("codcpc")
        a_Ingreso(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Ingreso(3, n_Index) = CDbl(0)
      ElseIf porstRecordset("tipocpc") = "1" Then
        n_Index = UBound(a_Descuento, 2) + 1
        ReDim Preserve a_Descuento(3, n_Index)
        a_Descuento(1, n_Index) = porstRecordset("codcpc")
        a_Descuento(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Descuento(3, n_Index) = CDbl(0)
      ElseIf porstRecordset("tipocpc") = s_Estado_Blq Then
        n_Index = UBound(a_Aporte, 2) + 1
        ReDim Preserve a_Aporte(3, n_Index)
        a_Aporte(1, n_Index) = porstRecordset("codcpc")
        a_Aporte(2, n_Index) = UCase(porstRecordset("aliascpc"))
        a_Aporte(3, n_Index) = CDbl(0)
      End If
      porstRecordset.MoveNext
    Wend
  End If
  porstRecordset.Close
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' Verifico si existen descuentos
  s_Sql = "SELECT IFNULL(COUNT(*), 0) AS registros "
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "dxr.codcco", "dxr.codubica", "dxr.codsec", "res.codpdo", "dxr.codcco, dxr.codubica", "dxr.codcco, dxr.codsec", "dxr.codubica, dxr.codcco", "dxr.codubica, dxr.codsec", "dxr.codsec, dxr.codcco", "dxr.codsec, dxr.codubica")
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  nRegistros = CLng(porstRecordset!registros)
  
  ' Registros detalle con campos
  s_Sql = "SELECT DISTINCTROW res.tipocpc, "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "res.codpdo, pdo.despdo, ", "dxr.codcco, cco.detcco, dxr.codubica, ubi.desubica, ", "dxr.codcco, cco.detcco, dxr.codsec, sec.dessec, ", "dxr.codubica, ubi.desubica, dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, dxr.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, dxr.codubica, ubi.desubica, ")
  s_Sql = s_Sql & "res.secuencia, res.codcpc, cpc.aliascpc, "
  s_Sql = s_Sql & "SUM(IFNULL((asi.diatrabajo+asi.diamediotm+asi.diaparcial)" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)) AS dias, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)), 2) AS importe "
  sSentenciaIni = "FROM plresultado res "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  sSentenciaIni = sSentenciaIni & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  sSentenciaIni = sSentenciaIni & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  If nTabIndex = 0 Then
    sSentenciaIni = sSentenciaIni & "INNER JOIN plcencospro dxc ON dxc.codcls=dxr.codcls AND dxc.codpdo=dxr.codpdo AND dxc.codpsn=dxr.codpsn "
    sSentenciaFin = Replace(sSentenciaFin, "dxr.codcco", "dxc.codcco")
  End If
  sSentenciaIni = sSentenciaIni & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON " & IIf(nTabIndex = 0, "dxc", "dxr") & ".codcco=cco.codcco "
  s_Sql = s_Sql & sSentenciaIni
  s_Sql = s_Sql & sSentenciaFin
  s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "res.codpdo, pdo.despdo, ", "dxr.codcco, cco.detcco, dxr.codubica, ubi.desubica, ", "dxr.codcco, cco.detcco, dxr.codsec, sec.dessec, ", "dxr.codubica, ubi.desubica, dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, dxr.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, dxr.codubica, ubi.desubica, ") & "res.tipocpc, res.codcpc, cpc.aliascpc "
  If nRegistros > 0 Then
    s_Sql = s_Sql & "UNION "
    s_Sql = s_Sql & "SELECT DISTINCT '" & s_Estado_Act & "' AS tipocpc, "
    s_Sql = s_Sql & Choose(nTabIndex + 1, "dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "res.codpdo, pdo.despdo, ", "dxr.codcco, cco.detcco, dxr.codubica, ubi.desubica, ", "dxr.codcco, cco.detcco, dxr.codsec, sec.dessec, ", "dxr.codubica, ubi.desubica, dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, dxr.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, dxr.codubica, ubi.desubica, ")
    s_Sql = s_Sql & "9 AS secuencia, 'dsct' AS codcpc, 'dscto' AS aliascpc, "
    s_Sql = s_Sql & "SUM(IFNULL((asi.diatrabajo+asi.diamediotm+asi.diaparcial)" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)) AS dias, 0.00 AS importe "
    s_Sql = s_Sql & sSentenciaIni
    s_Sql = s_Sql & sSentenciaFin
    s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "AND NOT EXISTS (SELECT * FROM plresultado dsc "
    s_Sql = s_Sql & "WHERE res.codcls=dsc.codcls "
    s_Sql = s_Sql & "AND res.codpdo=dsc.codpdo "
    s_Sql = s_Sql & "AND res.codpsn=dsc.codpsn "
    s_Sql = s_Sql & "AND dsc.impbolecpc='" & s_Estado_Act & "' "
    s_Sql = s_Sql & "AND dsc.tipocpc='" & s_Estado_Act & "') "
    s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "res.codpdo, pdo.despdo, ", "dxr.codcco, cco.detcco, dxr.codubica, ubi.desubica, ", "dxr.codcco, cco.detcco, dxr.codsec, sec.dessec, ", "dxr.codubica, ubi.desubica, dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, dxr.codsec, sec.dessec, ", "dxr.codsec, sec.dessec, dxr.codcco, cco.detcco, ", "dxr.codsec, sec.dessec, dxr.codubica, ubi.desubica, ") & "res.tipocpc "
  End If
  s_Sql = s_Sql & "ORDER BY " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "codpdo, ", "codcco, codubica, ", "codcco, codsec, ", "codubica, codcco, ", "codubica, codsec, ", "codsec, codcco, ", "codsec, codubica, ") & "tipocpc, secuencia"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Imprimiendo Planilla Trabajo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    
    nNivel = UBound(a_Ingreso, 2)
    nNivel = IIf(UBound(a_Descuento, 2) > nNivel, UBound(a_Descuento, 2), nNivel)
    nNivel = IIf(UBound(a_Aporte, 2) > nNivel, UBound(a_Aporte, 2), nNivel)
    ' Genero los arreglos de la grabación
    a_Campos = Array("seccion", "codtab", "destab", "codrubro", "secuencia", "desrubro", "dias", _
     "alias01", "alias02", "alias03", "alias04", "alias05", "alias06", "alias07", "alias08", "alias09", "alias10", "alias11", "alias12", "alias13", "alias14", "alias15", "alias16", "alias17", "alias18", "alias19", "alias20", _
     "alias21", "alias22", "alias23", "alias24", "alias25", "alias26", "alias27", "alias28", "alias29", "alias30", "alias31", "alias32", "alias33", "alias34", "alias35", "importe01", "importe02", "importe03", "importe04", "importe05", "importe06", "importe07", "importe08", "importe09", "importe10", _
     "importe11", "importe12", "importe13", "importe14", "importe15", "importe16", "importe17", "importe18", "importe19", "importe20", "importe21", "importe22", "importe23", "importe24", "importe25", "importe26", "importe27", "importe28", "importe29", "importe30", "importe31", "importe32", "importe33", "importe34", "importe35")
    a_Valores = Array("", "", "", "", CLng(0), "", CLng(0), _
     "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
     "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), _
     CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, _
     TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, _
     TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, _
     TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    nSecuencia = 0
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    
    
    While Not porstRecordset.EOF
      ' Inicialización de variables
      nSecuencia = nSecuencia + 1
      sQuiebre = porstRecordset("tipocpc")
      sRubro = porstRecordset(sCamRubro)
      sDescripcion = Trim(porstRecordset(sDesRubro))
      sTabular = porstRecordset(sCamTabular)
      sDescripcionTab = porstRecordset(sDesTabular)
      
      nDias = CLng(porstRecordset("dias"))
      ' Inicializo los importes del detalle
      If sQuiebre = s_Estado_Ina Then
        a_Registro = a_Ingreso
        nIngresos = CDec(0)
      ElseIf sQuiebre = s_Estado_Act Then
        a_Registro = a_Descuento
      ElseIf sQuiebre = s_Estado_Blq Then
        a_Registro = a_Aporte
      End If
      For n_Index = 0 To UBound(a_Registro, 2): a_Registro(3, n_Index) = CDbl(0): Next n_Index
      Do
        For n_Index = 1 To UBound(a_Registro, 2)
          If a_Registro(1, n_Index) = porstRecordset("codcpc") Then Exit For
        Next n_Index
        n_Index = IIf(n_Index > UBound(a_Registro, 2), UBound(a_Registro, 2), n_Index)
        a_Registro(3, n_Index) = CDec(porstRecordset("importe"))
        ' Totalizo detalles
        a_Registro(3, 0) = a_Registro(3, 0) + a_Registro(3, n_Index)
        nIngresos = nIngresos + CDec(IIf(sQuiebre = s_Estado_Ina, porstRecordset("importe"), 0))
        
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        porstRecordset.MoveNext
        If porstRecordset.EOF Then Exit Do
      Loop While (sQuiebre = porstRecordset("tipocpc") And sRubro = porstRecordset(sCamRubro) And sTabular = porstRecordset(sCamTabular))
      ' Inicializo valores del archivo temporal
      For n_Index = 1 To 35: a_Valores(6 + n_Index) = "": a_Valores(41 + n_Index) = CDec(0): Next n_Index
      ' Valores del archivo temporal
      a_Valores(0) = sQuiebre
      a_Valores(1) = sTabular
      a_Valores(2) = sDescripcionTab
      a_Valores(3) = sRubro
      a_Valores(4) = nSecuencia
      a_Valores(5) = sDescripcion
      a_Valores(6) = nDias
      ' Conceptos de acuerdo al tipo de concepto
      For n_Index = 1 To UBound(a_Registro, 2)
        a_Valores(6 + n_Index) = a_Registro(2, n_Index)
        a_Valores(41 + n_Index) = CDec(a_Registro(3, n_Index))
      Next n_Index
      ' Total de acuerdo al tipo de concepto
      a_Valores(6 + n_Index) = a_Registro(2, 0)
      a_Valores(41 + n_Index) = CDec(a_Registro(3, 0))
      If sQuiebre = "1" Then
        ' Neto a pagar
        a_Valores(7 + n_Index) = "NETO PAGAR"
        a_Valores(42 + n_Index) = nIngresos - CDec(a_Registro(3, 0))
      End If
      If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
    Wend
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
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
Private Sub ppPreSintesis(nTabIndex As Integer, s_Tabla As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)
  Dim nContador As Integer, nColumna As Integer
  Dim sColumna As String, a_Detalle(), a_Sintesis(9)
  Dim nImporteIng As Double, nImporteDsc As Double, nImportePag As Double
  Dim sCamRubro As String, sRubro As String, sDesRubro As String
  Dim sDescripcion As String, nDias As Long
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  
  ' Inicializo valores
  sCamRubro = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codpdo")
  sDesRubro = Choose(nTabIndex + 1, "detcco", "desubica", "dessec", "despdo")
  
  ' Genero las cabecera d elos conceptos
  s_Sql = "SELECT DISTINCTROW res.codcpc, cpc.descpc, res.tipocpc, cpc.aliascpc, "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "dxc.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "res.codpdo, pdo.despdo, ")
  s_Sql = s_Sql & "SUM(IFNULL(asi.diatrabajo" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)) AS dias, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(IF(res.tipocpc='" & s_Estado_Ina & "', res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)), 2) AS imporingreso, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(IF(res.tipocpc='" & s_Estado_Act & "', res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)), 2) AS impordescto, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(IF(res.tipocpc='" & s_Estado_Blq & "', res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)" & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)), 2) AS imporaporte, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "me", "mn") & IIf(nTabIndex = 0, "*(dxc.porcentaje/100)", "") & ", 0)), 2) AS importecmb "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  If nTabIndex = 0 Then
    s_Sql = s_Sql & "INNER JOIN plcencospro dxc ON dxc.codcls=res.codcls AND dxc.codpdo=res.codpdo AND dxc.codpsn=res.codpsn "
  End If
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON " & IIf(nTabIndex = 0, "dxc", "dxr") & ".codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND " & IIf(nTabIndex = 0, "dxc.", "dxr.") & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "dxc.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "res.codpdo, ") & "res.tipocpc, res.codcpc "
  s_Sql = s_Sql & "ORDER BY " & Choose(nTabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "codpdo, ") & "tipocpc, codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Imprimiendo Planilla Trabajo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    ' Recorro los registros
    While Not porstRecordset.EOF
      sRubro = porstRecordset(sCamRubro)
      nImporteIng = 0: nImporteDsc = 0: nImportePag = 0
      nContador = 0: nColumna = 9
      sDescripcion = Trim(porstRecordset(sDesRubro))
      nDias = CLng(porstRecordset("dias"))
      ReDim a_Detalle(9, 0)
      Do
        ' selecciono el tipo de concepto
        If nColumna <> CInt(porstRecordset("tipocpc")) Then
          nColumna = CInt(porstRecordset("tipocpc"))
          sColumna = Choose(nColumna + 1, "imporingreso", "impordescto", "imporaporte")
          nContador = 0
        End If
        nContador = nContador + 1
        ' Redimensiono e inicializo el arreglo de los detalles
        If nContador > UBound(a_Detalle, 2) Then
          ReDim Preserve a_Detalle(9, nContador)
          a_Detalle(1, nContador) = "": a_Detalle(2, nContador) = "": a_Detalle(3, nContador) = ""
          a_Detalle(4, nContador) = "": a_Detalle(5, nContador) = "": a_Detalle(6, nContador) = ""
          a_Detalle(7, nContador) = CDec(0): a_Detalle(8, nContador) = CDec(0): a_Detalle(9, nContador) = CDec(0)
        End If
        ' Asigno los datos al arreglo
        a_Detalle(nColumna + 1, nContador) = porstRecordset("codcpc")
        a_Detalle(nColumna + 4, nContador) = porstRecordset("descpc")
        a_Detalle(nColumna + 7, nContador) = CDec(porstRecordset(sColumna))
        ' Obtengo ingresos y descuentos otra moneda
        nImporteIng = nImporteIng + CDec(Choose(nColumna + 1, porstRecordset!importecmb, 0, 0))
        nImporteDsc = nImporteDsc + CDec(Choose(nColumna + 1, 0, porstRecordset!importecmb, 0))
        porstRecordset.MoveNext
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        If porstRecordset.EOF Then Exit Do
      Loop While sRubro = porstRecordset(sCamRubro)
      ' Obtengo el importe en otra moneda
      nImportePag = CDec(nImporteIng - nImporteDsc)
    
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      ' Inserto los detalle del analisis
      For n_Index = 1 To UBound(a_Detalle, 2)
        ' Inicializo los datos del detalle
        a_Sintesis(1) = "": a_Sintesis(2) = "": a_Sintesis(3) = ""
        a_Sintesis(4) = "": a_Sintesis(5) = "": a_Sintesis(6) = ""
        a_Sintesis(7) = 0: a_Sintesis(8) = 0: a_Sintesis(9) = 0
        If UBound(a_Detalle, 2) >= n_Index Then
          a_Sintesis(1) = a_Detalle(1, n_Index): a_Sintesis(2) = a_Detalle(2, n_Index)
          a_Sintesis(3) = a_Detalle(3, n_Index): a_Sintesis(4) = a_Detalle(4, n_Index)
          a_Sintesis(5) = a_Detalle(5, n_Index): a_Sintesis(6) = a_Detalle(6, n_Index)
          a_Sintesis(7) = a_Detalle(7, n_Index): a_Sintesis(8) = a_Detalle(8, n_Index)
          a_Sintesis(9) = a_Detalle(9, n_Index)
        End If
        a_Campos = Array("codrubro", "secuencia", "desrubro", "dias", "codcpcing", "descpcing", "impcpcing", "codcpcdsc", "descpcdsc", "impcpcdsc", "codcpcapo", "descpcapo", "impcpcapo", "impornetocmb")
        a_Valores = Array(sRubro, n_Index, sDescripcion, nDias, a_Sintesis(1), a_Sintesis(4), a_Sintesis(7), a_Sintesis(2), a_Sintesis(5), a_Sintesis(8), a_Sintesis(3), a_Sintesis(6), a_Sintesis(9), nImportePag)
        a_Tipos = Array(TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero)
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      Next n_Index
      gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
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
Private Sub RecuperaRegistros(ByVal nIndex As Integer, ByVal s_Orden As String)

  If nIndex = 0 Then
    s_Sql = "SELECT codcco, detcco, estcco "
    s_Sql = s_Sql & "FROM cocco "
    s_Sql = s_Sql & "WHERE LENGTH(codcco)>=" & pn_NivelCenCosto & " "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  ElseIf nIndex = 1 Then
    s_Sql = "SELECT codubica, desubica, estadoubica "
    s_Sql = s_Sql & "FROM plubicacion "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  ElseIf nIndex = 2 Then
    s_Sql = "SELECT codsec, dessec, estadosec "
    s_Sql = s_Sql & "FROM plseccion "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  ElseIf nIndex = 3 Then
    s_Sql = "SELECT codpdo, despdo, fechaini, fechafin, estadopdo "
    s_Sql = s_Sql & "FROM plperiodo "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND estadopdo<>'" & s_Estado_Ina & "' "
    s_Sql = s_Sql & "ORDER BY " & s_Orden
  End If
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & IIf(nIndex = 0, ps_DaBasCon, ps_DataBase), dcaSeleccion(nIndex), tdbSeleccion(nIndex), s_Sql, adCmdText, adLockReadOnly
  
  ' Inicializo los rangos de impresion
  as_SelRegistro(nIndex, 0) = "": as_SelRegistro(nIndex, 1) = ""
  If dcaSeleccion(nIndex).Recordset.RecordCount > 0 Then
    dcaSeleccion(nIndex).Recordset.MoveLast: as_SelRegistro(nIndex, 1) = dcaSeleccion(nIndex).Recordset.Bookmark
    dcaSeleccion(nIndex).Recordset.MoveFirst: as_SelRegistro(nIndex, 0) = dcaSeleccion(nIndex).Recordset.Bookmark
  End If
  
End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim nTabIndex As Integer, nTabIndexSel As Integer, nParametroSel As Integer
  Dim s_Periodo As String, s_Moneda As String, sOrden As String
  Dim s_TituloReporte As String, s_FechaHora As String
  
  nTabIndex = IIf(Not ribSeccion(0).Value, tabRegister.Tab, 3)
  ' Verifico que Existan Registros
  If (dcaSeleccion(3).Recordset.EOF Or dcaSeleccion(3).Recordset.BOF) Or (dcaSeleccion(3).Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & tdbSeleccion(0).Caption, vbExclamation: Exit Sub
  Select Case Index
   Case 0, 1  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros nTabIndex, tdbSeleccion(nTabIndex).Columns(tdbSeleccion(nTabIndex).Col).DataField & Choose(Index + 1, " ASC", " DESC")
   Case 2 ' Busqueda de registro
    If Not (dcaSeleccion(nTabIndex).Recordset.EOF Or dcaSeleccion(nTabIndex).Recordset.BOF) Then
      Set go_tdbBusqueda = tdbSeleccion(nTabIndex)
      Set go_dcaBusqueda = dcaSeleccion(nTabIndex)
      gn_ColBusqueda = (tdbSeleccion(nTabIndex).Columns.Count - 1)
      fBusqueda.Show vbModal
    End If
   Case 3, 4, 5 ' Selecciono rango de impresión
    gdl_Procedure.MarcaRegistros dcaSeleccion(nTabIndex), tdbSeleccion(nTabIndex), as_SelRegistro(nTabIndex, 0), as_SelRegistro(nTabIndex, 1), (Index - 3), tdbSeleccion(nTabIndex).Caption
   Case 6, 7, 8 ' Opciones de impresión
    nTabIndex = IIf(Not ribSeccion(0).Value, tabRegister.Tab, 3)
    nTabIndexSel = nTabIndex
    nParametroSel = IIf(Not ribSeccion(2).Value, nTabIndex, 4)
    If (nParametroSel >= 4 And nParametroSel <= 9) Then
      Select Case nTabIndex
       Case 0: nParametroSel = IIf(tdbSeleccion(1).SelBookmarks.Count > 0, 4, 5): nTabIndexSel = IIf(tdbSeleccion(1).SelBookmarks.Count > 0, 1, 2)
       Case 1: nParametroSel = IIf(tdbSeleccion(0).SelBookmarks.Count > 0, 6, 7): nTabIndexSel = IIf(tdbSeleccion(0).SelBookmarks.Count > 0, 0, 2)
       Case 2: nParametroSel = IIf(tdbSeleccion(0).SelBookmarks.Count > 0, 8, 9): nTabIndexSel = IIf(tdbSeleccion(0).SelBookmarks.Count > 0, 0, 1)
      End Select
    End If
    If (nParametroSel >= 0 And nParametroSel < 3) Then
       nParametroSel = IIf(tdbSeleccion(nTabIndex).SelBookmarks.Count > 0, nTabIndex, "")
    End If
    
    ' Verifico que existan registros seleccionados
    If tdbSeleccion(3).SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(3).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(0).SelBookmarks.Count = 0 And (nParametroSel = 0 Or nParametroSel = 4 Or nParametroSel = 5 Or nParametroSel = 6 Or nParametroSel = 8) Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(0).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(1).SelBookmarks.Count = 0 And (nParametroSel = 1 Or nParametroSel = 4 Or nParametroSel = 6 Or nParametroSel = 7 Or nParametroSel = 9) Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(1).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(2).SelBookmarks.Count = 0 And (nParametroSel = 2 Or nParametroSel = 5 Or nParametroSel = 7 Or nParametroSel = 8 Or nParametroSel = 9) Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(2).Caption & " de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
    s_Periodo = ""
    
    nTabIndex = 3
    '  Barro el arreglo de registros (periodos) marcados (bookmarks)
    For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
      tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      s_Periodo = s_Periodo & " - " & Trim(tdbSeleccion(nTabIndex).Columns(1).Text)
    Next n_Index
    
    nTabIndex = IIf(Not ribSeccion(0).Value, tabRegister.Tab, nTabIndex)
    If nParametroSel <> 3 Then
      ' Barro el arreglo de registros marcadas (bookmarks)
      For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
        tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
        gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, Left(s_OptRegistro, 9) & nTabIndex, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      Next n_Index
     ' segunda seleccion
      If nParametroSel > 3 Then
        For n_Index = 0 To tdbSeleccion(nTabIndexSel).SelBookmarks.Count - 1
          tdbSeleccion(nTabIndexSel).Bookmark = tdbSeleccion(nTabIndexSel).SelBookmarks(n_Index)
          gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, Left(s_OptRegistro, 9) & nTabIndexSel, tdbSeleccion(nTabIndexSel).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
        Next n_Index
      End If
    End If
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "REPORTE DE ANALISIS " & IIf(ribParametro(0).Value, "DETALLE", IIf(ribParametro(1).Value, "RESUMEN", "SINTESIS"))
    gdl_Procedure.ps_ReportName = IIf(ribParametro(0).Value, "rptpreplanilla", IIf(ribParametro(1).Value, "rptpreplaresum", "rptpreplasinte"))
    s_TituloReporte = "PLANILLA DE REMUNERACIONES - " & UCase(tdbSeleccion(nTabIndex).Caption) & IIf(nParametroSel > 3, "/" & UCase(tdbSeleccion(nTabIndexSel).Caption), "")
    s_TituloReporte = s_TituloReporte & " (" & IIf(s_Moneda = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt) & ")"
    
    'ReDim aElemento(3, 4): ReDim aElementos(2)
     ReDim aElemento(3, 5): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbSeleccion(nTabIndex).Columns(0).DataField & " ASC"
    aElemento(0, 2) = "": aElemento(0, 3) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "":  aElemento(1, 2) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & s_TituloReporte & ";true"
    aElemento(2, 2) = "Periodo;" & s_Periodo & ";true"
    aElemento(2, 3) = "Grupo;" & nParametroSel & ";true"
    aElemento(2, 4) = "Sub_Grupo;" & nTabIndexSel & ";true"
    
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If ribParametro(0).Value Then
      s_Sql = s_Sql & "seccion char(1) Not Null, "
      s_Sql = s_Sql & "codcco varchar(10) Not Null, "
      s_Sql = s_Sql & "codsec varchar(10) Not Null, "
      s_Sql = s_Sql & "codpsn varchar(11) Not Null, "
      s_Sql = s_Sql & "secuencia smallint(5) Not Null, "
      s_Sql = s_Sql & "nombrepsn varchar(80) Null, "
      s_Sql = s_Sql & "numdociden varchar(11) Null, "
      s_Sql = s_Sql & "detcco varchar(50) Null, "
      s_Sql = s_Sql & "dessec varchar(50) Null, "
      s_Sql = s_Sql & "codafp char(2) Null, "
      s_Sql = s_Sql & "desafp varchar(50) Null, "
      s_Sql = s_Sql & "descgo varchar(50) Null, "
      s_Sql = s_Sql & "fecingreso date default Null, "
      s_Sql = s_Sql & "dias smallint(3) Null, "
      For n_Index = 1 To 35
        s_Sql = s_Sql & "alias" & Format(n_Index, "00") & " varchar(10) Null, "
        s_Sql = s_Sql & "importe" & Format(n_Index, "00") & " decimal(18,2) Null Default '0', "
      Next n_Index
      s_Sql = s_Sql & "PRIMARY KEY (seccion, codcco, codsec, codpsn, secuencia)) "
      sOrden = "seccion, " & IIf(Not ribSeccion(0).Value, "codcco, codsec, ", "") & IIf(ribOrdenar.Value, "nombrepsn, ", "") & "codpsn, secuencia"
    ElseIf ribParametro(1).Value Then
      s_Sql = s_Sql & "seccion char(1) Not Null, "
      s_Sql = s_Sql & "codtab varchar(10) Not Null, "
      s_Sql = s_Sql & "destab varchar(50) Null, "
      s_Sql = s_Sql & "codrubro varchar(10) Not Null, "
      s_Sql = s_Sql & "secuencia smallint(5) Null, "
      s_Sql = s_Sql & "desrubro varchar(50) Null, "
      s_Sql = s_Sql & "dias int(6) Null, "
      For n_Index = 1 To 35
        s_Sql = s_Sql & "alias" & Format(n_Index, "00") & " varchar(10) Null, "
        s_Sql = s_Sql & "importe" & Format(n_Index, "00") & " decimal(18,2) Null Default '0', "
      Next n_Index
      s_Sql = s_Sql & "PRIMARY KEY (seccion, codrubro, secuencia)) "
      sOrden = "codrubro, secuencia"
    ElseIf ribParametro(2).Value Then
      s_Sql = s_Sql & "codrubro varchar(10) Not Null, "
      s_Sql = s_Sql & "secuencia smallint(5) Not Null, "
      s_Sql = s_Sql & "desrubro varchar(50) Null, "
      s_Sql = s_Sql & "dias int(6) Null, "
      s_Sql = s_Sql & "codcpcing varchar(4) Null, "
      s_Sql = s_Sql & "descpcing varchar(40) Null, "
      s_Sql = s_Sql & "impcpcing decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcdsc varchar(4) Null, "
      s_Sql = s_Sql & "descpcdsc varchar(40) Null, "
      s_Sql = s_Sql & "impcpcdsc decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcapo varchar(4) Null, "
      s_Sql = s_Sql & "descpcapo varchar(40) Null, "
      s_Sql = s_Sql & "impcpcapo decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "impornetocmb decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "PRIMARY KEY (codrubro, secuencia)) "
      sOrden = "codrubro, secuencia"
    End If
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' Genera la información del reporte
    If ribParametro(0).Value Then
      ppPrePlanilla nParametroSel, "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora, s_Moneda
    ElseIf ribParametro(1).Value Then
      ppPreResumen nParametroSel, "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora, s_Moneda
    ElseIf ribParametro(2).Value Then
      ppPreSintesis nTabIndex, "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora, s_Moneda
    Else
      ppPreRemuneracion nTabIndex, "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora, s_Moneda
    End If
    If Index <> 8 Then
      ' Elimino campo cargo
      If ribParametro(0).Value Then
        s_Sql = "ALTER TABLE tmp" & gdl_Procedure.ps_ReportName & " "
        s_Sql = s_Sql & "DROP COLUMN descgo"
        gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      End If
      ' Selecciono información
      s_Sql = "SELECT * "
      s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
      s_Sql = s_Sql & "ORDER BY " & sOrden
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      ' Ejecuto reporte y saco de memoria la información
      gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Else
      ppExcelHorizontal "tmp" & gdl_Procedure.ps_ReportName, s_Periodo, s_TituloReporte
    End If
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, Left(s_OptRegistro, 9) & nTabIndex, "", ps_Usuario, s_FechaHora, "E"
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, Left(s_OptRegistro, 9) & nTabIndexSel, "", ps_Usuario, s_FechaHora, "E"
    ' ]
   Case 9
    nTabIndex = IIf(Not ribSeccion(0).Value, tabRegister.Tab, 3)
    ' Verifico que existan registros seleccionados
    If tdbSeleccion(3).SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(3).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 1 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 2 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
    s_Periodo = "": nTabIndex = 3
    'Barro el arreglo de registros (periodos) marcados (bookmarks)
    For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
      tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      s_Periodo = s_Periodo & " - " & Trim(tdbSeleccion(nTabIndex).Columns(1).Text)
    Next n_Index
    ppExcelVertical s_Periodo, s_OptRegistro, s_FechaHora
  End Select

End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim Item As New ValueItem

  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
  cnn.CursorLocation = adUseClient
  cnn.Open
    
  ' Establece posición del formulario
  Me.Height = 6330: Me.Width = 11200
  Me.Left = 400: Me.Top = 350
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Seleción de Periodos de Pago"
  s_TitleTable = "Periodos de Pago"
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Inicio", "Final", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpdo", "despdo", "fechaini", "fechafin", "estadopdo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 850, 1700, 950, 950, 300)
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
  gdl_Procedure.InicializaGrilla tdbSeleccion(3), aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbSeleccion(3).Columns(4).ValueItems.Presentation = dbgNormal
  tdbSeleccion(3).Columns(4).ValueItems.Translate = True
  For n_Index = 0 To 1
    tdbSeleccion(3).Columns(4).ValueItems.Add Item
    tdbSeleccion(3).Columns(4).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Estado_Act, s_Estado_Blq)
    tdbSeleccion(3).Columns(4).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "proceok", "perioblk") & ".bmp")
  Next n_Index
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbSeleccion(3), s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbSeleccion(3).GroupByCaption = "Arrastrar titulo de columna de agrupación"
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(10, 3)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir", "excellnk", "excelxlsx")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir", "Exporta Excel - Horizontal", "Exporta Excel - Vertical")
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&a", "&d", "&b", "&p", "&f", "&r", "&v", "&i", "&h", "&e")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo graficos botones seccion
  For n_Index = 0 To 2
    ' Tipo de analisis
    ribSeccion(n_Index).PictureUp = LoadPicture()
    ribSeccion(n_Index).ToolTipText = Choose(n_Index + 1, "Periodo de Pago", "Parámetro de Clasificación", "Doble Clasificación")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "saldmes", "dividir", "asoclibr") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribSeccion(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribSeccion(0).Value = True
  
  ribOrdenar.PictureUp = LoadPicture()
  ribOrdenar.ToolTipText = "Reporte Alfabeticamente"
  s_Sql = gdl_Procedure.ps_PathImagen & "ordalfab.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribOrdenar.PictureUp = LoadPicture(s_Sql)
  ribOrdenar.Value = False
  
  ' [ Configuros grillas de selección
  ReDim aElemento(3, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codcco", "detcco", "estcco")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 700, 3066.03, 300)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0)
  Next n_Index
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  For n_Index = 0 To 2
    aElemento(0, 1) = Choose(n_Index + 1, "codcco", "codubica", "codsec")
    aElemento(1, 1) = Choose(n_Index + 1, "detcco", "desubica", "dessec")
    aElemento(2, 1) = Choose(n_Index + 1, "estcco", "estadoubica", "estadosec")
    gdl_Procedure.InicializaGrilla tdbSeleccion(n_Index), aElemento, aElementos
    ' Cambio el formato de la grilla columna de valores
    tdbSeleccion(n_Index).Columns(2).ValueItems.Presentation = dbgNormal
    tdbSeleccion(n_Index).Columns(2).ValueItems.Translate = True
    ' Primera columna
    tdbSeleccion(n_Index).Columns(2).ValueItems.Add Item
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(0).Value = IIf(n_Index = 0, "A", s_Estado_Act)
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(0).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & "estadok.bmp")
    ' Segunda columna
    tdbSeleccion(n_Index).Columns(2).ValueItems.Add Item
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(1).Value = IIf(n_Index = 0, "I", s_Estado_Ina)
    tdbSeleccion(n_Index).Columns(2).ValueItems.Item(1).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & "estadnok.bmp")
    ' Personaliza el estilo de la grilla de TDBGrid
    gdl_Procedure.DefineStyleGrilla tdbSeleccion(n_Index), Choose(n_Index + 1, "Centro Costo", "Ubicación o Localidad", "Sección de Empresa"), 1
    ' Agrupacion de columnas y titulo DataView = dbgGroupView
    tdbSeleccion(n_Index).GroupByCaption = "Arrastrar titulo de columna de agrupación"
  Next n_Index
  ']
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 2
    ' Tipo de analisis
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = Choose(n_Index + 1, "Detallado", "Resumen", "Síntesis")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "analmovs", "resumen", "ajuinfla") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribParametro(0).Value = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  For n_Index = 0 To 3
    tdbSeleccion(n_Index).DataSource = dcaSeleccion(n_Index)
    RecuperaRegistros n_Index, tdbSeleccion(n_Index).Columns(0).DataField & " ASC"
  Next n_Index
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
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

Private Sub tdbSeleccion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaSeleccion(Index), tdbSeleccion(Index), " " & tdbSeleccion(Index).Caption
End Sub


