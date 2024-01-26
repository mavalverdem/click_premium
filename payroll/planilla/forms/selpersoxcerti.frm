VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPersoCertifik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "selpersoxcerti.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   7740
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   11
      Top             =   990
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
      Top             =   5895
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
      TabIndex        =   0
      Top             =   990
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
         TabIndex        =   10
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
         TabIndex        =   3
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
         Picture         =   "selpersoxcerti.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   4
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
         Picture         =   "selpersoxcerti.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   5
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
         Picture         =   "selpersoxcerti.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   6
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
         Picture         =   "selpersoxcerti.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   8
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
         Picture         =   "selpersoxcerti.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   9
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
         Picture         =   "selpersoxcerti.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   2
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
         Picture         =   "selpersoxcerti.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   7
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
         Picture         =   "selpersoxcerti.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   1
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
         Picture         =   "selpersoxcerti.frx":00EC
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   930
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7740
      _Version        =   65536
      _ExtentX        =   13652
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
      Begin VB.ComboBox cmbParametro 
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
         ItemData        =   "selpersoxcerti.frx":0108
         Left            =   2970
         List            =   "selpersoxcerti.frx":010A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   90
         Width           =   2220
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   6525
         TabIndex        =   14
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
         PictureUp       =   "selpersoxcerti.frx":010C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   6090
         TabIndex        =   13
         Top             =   75
         Width           =   420
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   635
         _StockProps     =   65
         BackColor       =   14737632
         GroupAllowAllUp =   0   'False
         PictureDnChange =   2
         Autosize        =   2
         BevelWidth      =   0
         Outline         =   0   'False
         PictureUp       =   "selpersoxcerti.frx":0128
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   6930
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
         PictureUp       =   "selpersoxcerti.frx":0144
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   1
         Left            =   795
         TabIndex        =   16
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
         PictureUp       =   "selpersoxcerti.frx":0160
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   0
         Left            =   390
         TabIndex        =   17
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
         PictureUp       =   "selpersoxcerti.frx":017C
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   2970
         TabIndex        =   21
         Top             =   465
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   141623297
         CurrentDate     =   37515
      End
      Begin Threed.SSRibbon ribFirma 
         Height          =   360
         Left            =   6930
         TabIndex        =   22
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
         PictureUp       =   "selpersoxcerti.frx":0198
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha  :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   1995
         TabIndex        =   20
         Top             =   495
         Width           =   900
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parametro :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1995
         TabIndex        =   18
         Top             =   150
         Width           =   900
      End
   End
End
Attribute VB_Name = "fSelPersoCertifik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(2)                               ' Array de inicio y fin de seleccion de registro
Private s_OptRegistro As String
Private desde As Integer
Private cnn As ADODB.Connection
' Instancia del formulario activo
'[
Private Sub CertificadoPension(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim s_Moneda As String
  Dim nRentaBruta As Double, nImpuestoRenta As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim sConceptoRemun As String, sConceptoReten As String
  Dim sPersonal As String, sDocIdentidad As String, sNumeroAfp As String

  ' Cambio el Mensaje y Muestro la Barra
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  sConceptoRemun = Choose(cmbParametro.ListIndex + 1, "cpcremuonp", "cpcremuessalud", "cpcremuessalud")
  sConceptoReten = Choose(cmbParametro.ListIndex + 1, "cpconp", "cpcessalud", "cpceps")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo mes
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  s_Sql = s_Sql & "psn.numdociden, psn.numeroafp, dxr.fecingreso, asi.fechacese, dxr.naciextrapsn, res.codpdo "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg." & sConceptoReten & " "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.codpdo>'" & s_PeriodoRemAper & "' "
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
  s_Sql = s_Sql & "codcpc varchar(4) NOT Null, "
  s_Sql = s_Sql & "rembruta decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "impreten decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "fecingreso date default Null, "
  s_Sql = s_Sql & "fecbaja date default Null)"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Genero tabla de remuneraciones asegurables
  s_Sql = "DROP TABLE IF EXISTS tmpasegurable"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpasegurable "
  s_Sql = s_Sql & "SELECT res.codcls, res.codpsn, ROUND(SUM(IFNULL(res.importe_" & s_Moneda & ", 0)), 2) AS remasegurable "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.codpdo=psn.codpdo "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg." & sConceptoRemun & " "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto el importe de retencion
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.codcpc, ras.remasegurable AS rembruta, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & s_Moneda & ", 0)), 2) AS impreten, "
  s_Sql = s_Sql & "MAX(psn.fecingreso) AS fecingreso, MAX(psn.fechacese) AS fecbaja "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.codpdo=psn.codpdo "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg." & sConceptoReten & " "
  s_Sql = s_Sql & "INNER JOIN tmpasegurable ras ON res.codcls=ras.codcls AND res.codpsn=ras.codpsn "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Recupero la informacion del certificado
  s_Sql = "SELECT DISTINCTROW tmp.codpsn, psn.nombrespsn, psn.numdociden, "
  s_Sql = s_Sql & "psn.numeroafp, tmp.fecingreso, tmp.fecbaja, psn.naciextrapsn, "
  s_Sql = s_Sql & "tmp.codcpc, cpc.descpc, tmp.rembruta, tmp.impreten "
  s_Sql = s_Sql & "FROM tmpimporte tmp "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON tmp.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON tmp.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "ORDER BY codpsn"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    a_Campos = Array("codpsn", "nombrespsn", "numdociden", "numeroafp", "fecingreso", "fecbaja", "codcpc", "descpc", "moneda", "remunbruta", "retencion")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero)
    While Not porstRecordset.EOF
      sPersonal = porstRecordset!codpsn
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      sNumeroAfp = IIf(IsNull(porstRecordset!numeroafp), "", porstRecordset!numeroafp)
      If CDec(porstRecordset!impreten) > 0 Then
        nRentaBruta = CDec(porstRecordset!rembruta)
        nImpuestoRenta = CDec(porstRecordset!impreten)
        a_Valores = Array(sPersonal, UCase(porstRecordset!nombrespsn), sDocIdentidad, sNumeroAfp, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), Trim(porstRecordset!codcpc), "REMUNERACIONES AFECTAS", s_Moneda, nRentaBruta, nImpuestoRenta)
        
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
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub CertificadoQuinta(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim a_TablaUit(12) As Double
  Dim nRentaBruta As Double, nRentaBruAper As Double, nRentaBruEmp As Double
  Dim nDeduccion As Double, nRentaNeta As Double, nImporteUit As Double
  Dim nImpuestoRenta As Double, nImpuRentaApe As Double, nImpuRentaEmp As Double, nImpuRentaIng As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim s_Moneda As String, sDocIdentidad As String, sDireccion As String
  
  ' Cambio el Mensaje y Muestro la Barra
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  
  '[ Obtengo la UIT
  s_Sql = "SELECT DISTINCTROW tbl.valor01, tbl.valor02, tbl.valor03, tbl.valor04, tbl.valor05, tbl.valor06, "
  s_Sql = s_Sql & "tbl.valor07, tbl.valor08, tbl.valor09, tbl.valor10, tbl.valor11, tbl.valor12 "
  s_Sql = s_Sql & "FROM pltablabase tbl "
  s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON tbl.pdoano=cfg.pdoano AND tbl.codtbl=cfg.codtbluit "
  s_Sql = s_Sql & "WHERE tbl.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND tbl.pdoano='" & ps_Anyo & "' "
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    a_TablaUit(1) = CDec(porstRecordset("valor01"))
    For n_Index = 2 To 12
      a_TablaUit(n_Index) = a_TablaUit(n_Index - 1) + CDec(porstRecordset("valor" & Format(n_Index, "00")))
    Next n_Index
  End If
  porstRecordset.Close
  ']

  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo mes
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  's_Sql = "CREATE TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  s_Sql = s_Sql & "CONCAT(IFNULL(via.abrevia, ''), ' ', IFNULL(psn.nomviadirec, ''), ' ', ' Nº ', IFNULL(psn.numerdirec, ''), ' ', IFNULL(psn.intedirec, ''), ' ', IFNULL(zon.abrezona, ''), ' ', IFNULL(psn.nomzondirec, '')) AS direccionpsn, "
  s_Sql = s_Sql & "IFNULL(psn.ubigeodir, '') AS ubigeodir, IFNULL(dci.sigladci, '') AS sigladci, psn.numdociden, MAX(dxr.fecingreso) AS fecingreso, MAX(asi.fechacese) AS fecbaja, dxr.naciextrapsn, MAX(res.pdomes) AS mesfin, MAX(res.codpdo) AS pdofin "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON pdo.codcls=res.codcls AND pdo.codpdo=res.codpdo AND pdo.tpopdo NOT IN('G', 'O') "
  s_Sql = s_Sql & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
  s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
  s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.codpdo>'" & s_PeriodoRemAper & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "GROUP BY codpsn "
  s_Sql = s_Sql & "ORDER BY codpsn "
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal del certificado
  s_Sql = "DROP TABLE IF EXISTS tmpimporte"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimporte ( "
  's_Sql = "CREATE TABLE tmpimporte ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, "
  s_Sql = s_Sql & "remganado decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remapertura decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "impquinta decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "impapertura decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "impquintaing decimal(18, 2) NOT Null Default 0) "
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Inserto las remuneraciones ganadas
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.importe_" & s_Moneda & " AS remganado, 0.00 AS remapertura, "
  s_Sql = s_Sql & "0.00 AS impquinta, 0.00 AS impapertura, 0.00 AS impquintaing "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdomes=psn.mesfin AND res.codpdo=psn.pdofin "
  s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remganada "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto las remuneraciones de apertura(otra empresa)
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codpsn, 0.00 AS remganado, res.importe_" & s_Moneda & " AS remapertura, "
  s_Sql = s_Sql & "0.00 AS impquinta, 0.00 AS impapertura, 0.00 AS impquintaing "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remanterior "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & s_PeriodoRemAper & "' "
' s_Sql = s_Sql & "AND res.codproce='" & s_ProcesoRemAper & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto el impuesto de la quinta
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS remganado, 0.00 AS remapertura, "
  s_Sql = s_Sql & "res.importe_" & s_Moneda & " AS impquinta, 0.00 AS impapertura, 0.00 AS impquintaing "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.codcpc5ta "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto devolución de impuesto de la quinta
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS remganado, 0.00 AS remapertura, "
  s_Sql = s_Sql & "0.00 AS impquinta, 0.00 AS impapertura, res.importe_" & s_Moneda & " AS impquintaing "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.codcpc5ta_ing "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto el impuesto de quinta apertura(otra empresa)
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codpsn, 0.00 AS remganado, 0.00 AS remapertura, "
  s_Sql = s_Sql & "0.00 AS impquinta, res.importe_" & s_Moneda & " AS impapertura, 0.00 AS impquintaing "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plcfgempresa cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.codcpc5ta "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & s_PeriodoRemAper & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ']
  
  ' Recupero la informacion del certificado
  s_Sql = "SELECT tmp.codpsn, psn.nombrespsn, psn.numdociden,psn.fecingreso,psn.fecbaja, psn.naciextrapsn, psn.mesfin, "
  s_Sql = s_Sql & "psn.direccionpsn,psn.sigladci,psn.ubigeodir,IFNULL(cgo.descgo, '') as descgo, "
  s_Sql = s_Sql & "SUM(IFNULL(remganado, 0)) AS remganado, "
  s_Sql = s_Sql & "SUM(IFNULL(remapertura, 0)) AS remapertura, "
  s_Sql = s_Sql & "SUM(IFNULL(impquinta, 0)) AS impquinta, "
  s_Sql = s_Sql & "SUM(IFNULL(impapertura, 0)) AS impapertura, "
  s_Sql = s_Sql & "SUM(IFNULL(impquintaing, 0)) AS impquintaing "
  s_Sql = s_Sql & "FROM tmpimporte tmp "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON tmp.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON psn.codcls=dxr.codcls AND psn.pdofin=dxr.codpdo AND psn.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "GROUP BY codpsn "
  s_Sql = s_Sql & "ORDER BY codpsn"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    
    a_Campos = Array("codpsn", "nombrespsn", "sigladci", "numdociden", "descgo", "direccionpsn", "fecingreso", "fecbaja", "moneda", "rentabruta", "rentabruaper", "rentabruemp", "deduccion", "rentaneta", "impuestorenta", "impurentaape", "impurentaemp", "impurentaing")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    While Not porstRecordset.EOF
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      sDireccion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, porstRecordset!ubigeodir, "UB")
      sDireccion = porstRecordset!direccionpsn & " - " & Trim(sDireccion)
      If CDec(porstRecordset!remganado) > 0 Then
        ' Obtengo la renta Bruta
        nImporteUit = Round(a_TablaUit(CInt(porstRecordset!mesfin)) / CInt(porstRecordset!mesfin), 2)
        nRentaBruta = CDec(porstRecordset!remganado)
        nRentaBruAper = CDec(porstRecordset!remapertura)
        nRentaBruEmp = Round(nRentaBruta - nRentaBruAper, 2)
        nDeduccion = Round(IIf(porstRecordset!naciextrapsn = s_Estado_Act, 0, nImporteUit) * 7, 2)
        nRentaNeta = Round(nRentaBruta - nDeduccion, 2)
        nRentaNeta = IIf(nRentaNeta > 0, nRentaNeta, 0)
        nImpuestoRenta = CDec(porstRecordset!impquinta)
        nImpuRentaApe = CDec(porstRecordset!impapertura)
        nImpuRentaEmp = Round(nImpuestoRenta - nImpuRentaApe, 2)
        nImpuRentaIng = CDec(porstRecordset!impquintaing) * -1
        a_Valores = Array(porstRecordset!codpsn, UCase(porstRecordset!nombrespsn), porstRecordset!sigladci, sDocIdentidad, UCase(porstRecordset!descgo), sDireccion, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), s_Moneda, nRentaBruta, nRentaBruAper, nRentaBruEmp, nDeduccion, nRentaNeta, nImpuestoRenta, nImpuRentaApe, nImpuRentaEmp, nImpuRentaIng)
        
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
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub CertificadoUtilidad(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim a_Parametro(6) As String
  Dim nRentaBruta As Double, nPorcentaje As Double, nParticipacion As Double
  Dim nPartixCalculo As Double, nPartixDia As Double, nTotalRemunera As Double
  Dim nRemunera As Double, nPartixRemunera As Double, nPartixPsn As Double
  Dim nTotalDias As Long, nDias As Integer, nImporte As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim s_Moneda As String, sDocIdentidad As String
  
  ' Cambio el Mensaje y Muestro la Barra
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  
  ' Recupero los parametros de remuneraciones
  s_Sql = "SELECT remxutiejer1, remxutiejer2, remxutiejer3, remxutiejer4, "
  s_Sql = s_Sql & "rentaxejer_" & s_Moneda & " AS rentaxejerci, porcepartici "
  s_Sql = s_Sql & "FROM plcfgempresa "
  s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros del reporte", vbCritical: Exit Sub
  
  ' Obtengo los conceptos de recuperación
  a_Parametro(1) = gdl_Funcion.aTexto(porstRecordset!remxutiejer1)
  a_Parametro(2) = gdl_Funcion.aTexto(porstRecordset!remxutiejer2)
  a_Parametro(3) = gdl_Funcion.aTexto(porstRecordset!remxutiejer3)
  a_Parametro(4) = gdl_Funcion.aTexto(porstRecordset!remxutiejer4)
  nRentaBruta = CDec(porstRecordset!rentaxejerci)
  nPorcentaje = CDec(porstRecordset!porcepartici)
  nParticipacion = Round(nRentaBruta * (nPorcentaje / 100), 2)
  nPartixCalculo = Round(nParticipacion / 2, 2)
  porstRecordset.Close
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección
  s_Sql = "DROP TABLE IF EXISTS tmpregistro "
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpregistro "
  ' Obtengo la información inicial
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  s_Sql = s_Sql & "psn.numdociden, psn.fecingreso, psn.fecbaja, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & s_Moneda & ", 0)), 2) AS remuneracion "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.codpdo<>'" & s_PeriodoRemAper & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND res.codcpc IN ('" & a_Parametro(1) & "', '" & a_Parametro(2) & "', '" & a_Parametro(3) & "', '" & a_Parametro(4) & "') "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Genero tabla de dias trabajados
  s_Sql = "DROP TABLE IF EXISTS tmpasistencia"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpasistencia "
  s_Sql = s_Sql & "SELECT DISTINCTROW asi.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  s_Sql = s_Sql & "psn.numdociden, psn.fecingreso, psn.fecbaja, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(asi.dialaboral, 0)), 0) AS dias "
  s_Sql = s_Sql & "FROM plasistencia asi "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND asi.codpdo<>'" & s_PeriodoRemAper & "' "
  s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND pdo.anopdo='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND pdo.estadopdo<>'" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "GROUP BY asi.codpsn "
  s_Sql = s_Sql & "ORDER BY asi.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  '***************************************************************
  
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
  cnn.CursorLocation = adUseClient
  cnn.Open
  
  Dim rsclases As New Recordset
  Dim strsql As String
  
  nTotalRemunera = 0
  nTotalDias = 0
  strsql = "select codcls from plclasplan where LEFT(tipo,2)='01' or LEFT(tipo,2)='02' "
  
  rsclases.Open strsql, cnn, adOpenStatic, adLockOptimistic
  rsclases.MoveFirst
  For desde = 1 To rsclases.RecordCount
  
    ' Obtengo los totales generales de calculo - remuneraciones
    s_Sql = "SELECT ROUND(SUM(IFNULL(res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0)), 2) AS sumremunera "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
    s_Sql = s_Sql & "WHERE res.codcls='" & rsclases(0) & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND res.codpdo<>'" & s_PeriodoRemAper & "' "
    s_Sql = s_Sql & "AND res.codcpc IN ('" & a_Parametro(1) & "', '" & a_Parametro(2) & "', '" & a_Parametro(3) & "', '" & a_Parametro(4) & "') "
    
     
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If IsNull(porstRecordset!sumremunera) = True Then
    Else
    nTotalRemunera = CDec(porstRecordset!sumremunera) + nTotalRemunera
    End If
    porstRecordset.Close
    
      
    ' Obtengo los totales generales de calculo - dias
    s_Sql = "SELECT ROUND(SUM(IFNULL(asi.dialaboral, 0)), 0) AS numdias "
    s_Sql = s_Sql & "FROM plasistencia asi "
    s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
    s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo "
    s_Sql = s_Sql & "WHERE asi.codcls='" & rsclases(0) & "' "
    s_Sql = s_Sql & "AND asi.codpdo<>'" & s_PeriodoRemAper & "' "
    s_Sql = s_Sql & "AND pdo.anopdo='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND pdo.estadopdo<>'" & s_Estado_Ina & "' "
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If IsNull(porstRecordset!numdias) = True Then
    Else
    nTotalDias = CLng(porstRecordset!numdias) + nTotalDias
    End If
    porstRecordset.Close
       
    rsclases.MoveNext
  Next
 
  '***************************************************************
      
  ' Genero tabla de trabajadores por dias
  s_Sql = "DROP TABLE IF EXISTS tmpxistenpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpxistenpsn "
  s_Sql = s_Sql & "SELECT DISTINCTROW asi.codpsn, asi.nombrespsn, asi.numdociden, "
  s_Sql = s_Sql & "asi.fecingreso, asi.fecbaja, 0.00 AS remuneracion "
  s_Sql = s_Sql & "FROM tmpasistencia asi "
  s_Sql = s_Sql & "LEFT JOIN tmpregistro tmp ON asi.codpsn=tmp.codpsn "
  s_Sql = s_Sql & "WHERE IFNULL(tmp.codpsn, '')='' "
  s_Sql = s_Sql & "ORDER BY asi.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto los registros de asistencia
  s_Sql = "INSERT INTO tmpregistro "
  s_Sql = s_Sql & "SELECT DISTINCTROW asi.codpsn, asi.nombrespsn, asi.numdociden, "
  s_Sql = s_Sql & "asi.fecingreso, asi.fecbaja, asi.remuneracion "
  s_Sql = s_Sql & "FROM tmpxistenpsn asi "
  s_Sql = s_Sql & "ORDER BY asi.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ']
  
  ' Recupero la informacion del reporte
  s_Sql = "SELECT DISTINCTROW psn.codpsn, psn.nombrespsn, psn.numdociden, psn.fecingreso, "
  s_Sql = s_Sql & "psn.fecbaja, psn.remuneracion, asi.dias "
  s_Sql = s_Sql & "FROM tmpregistro psn "
  s_Sql = s_Sql & "INNER JOIN tmpasistencia asi ON psn.codpsn=asi.codpsn "
  s_Sql = s_Sql & "ORDER BY psn.codpsn"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    a_Campos = Array("codpsn", "nombrespsn", "numdociden", "fecingreso", "fecbaja", "moneda", "rentabruta", "porpartici", "renpartici", "partixcalculo", "totaldias", "diaspsn", "partixdiapsn", "totalremun", "remunpsn", "partixrempsn", "partixpsn")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    While Not porstRecordset.EOF
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      If (CDec(porstRecordset!remuneracion) > 0 Or CDec(porstRecordset!dias) > 0) Then
        ' Calculop por dias (50%)
        nDias = CInt(porstRecordset!dias)
        nImporte = CDec(nPartixCalculo / nTotalDias)
        nPartixDia = Round(nImporte * nDias, 2)
        ' Calculop por remuneraciones (50%)
        nRemunera = CDec(porstRecordset!remuneracion)
        nImporte = CDec(nPartixCalculo / nTotalRemunera)
        nPartixRemunera = Round(nImporte * nRemunera, 2)
        nPartixPsn = Round(nPartixDia + nPartixRemunera, 2)
        a_Valores = Array(porstRecordset!codpsn, UCase(porstRecordset!nombrespsn), sDocIdentidad, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), s_Moneda, nRentaBruta, nPorcentaje, nParticipacion, nPartixCalculo, nTotalDias, nDias, nPartixDia, nTotalRemunera, nRemunera, nPartixRemunera, nPartixPsn)
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
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
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
  Dim sDireccion As String, sRepresentante As String
  Dim sTipoDocum As String, sNumDocumento As String
  Dim sExpresion As String, sCargoRepresenta As String
  Dim sConcepto As String, nParticipacion As Double
  Dim sFechaPrn As String, sDistrito As String
  Dim sPrnLogo As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Actualización de parametros
    Me.Tag = s_MdoData_Vis
    If s_OptRegistro = "certifi5ta" Then
      fPrmCertifik5ta.Show vbModal
    ElseIf s_OptRegistro = "certifisnp" Then
      fPrmCertifikSnp.Show vbModal
    ElseIf s_OptRegistro = "certifiuti" Then
      fPrmCertifikUti.Show vbModal
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
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    sFechaPrn = Format(dtpFecha, "dd") & " de " & gdl_Funcion.NombreMes(Format(dtpFecha, "mm")) & " del " & Format(dtpFecha, "yyyy")
        
    ' Obtengo y verifico los datos de la empresa
    sDireccion = "": sRepresentante = "": sCargoRepresenta = ""
    sConcepto = IIf(s_OptRegistro = "certifi5ta", "cfg.codcpc5ta ", IIf(s_OptRegistro = "certifiuti", "cfg.remxutiejer1", Choose(cmbParametro.ListIndex + 1, "afp.cpconp", "afp.cpcessalud", "afp.cpceps")))
    sExpresion = IIf((ribAnalisis(0).Value And ribFirma.Value), "ger", "rep")
    
    s_Sql = "SELECT cfg.codvia, cfg.direccionvia, cfg.numerodir, cfg.codzona, cfg.direccionzona, cfg.ubigeodir, "
    s_Sql = s_Sql & "CONCAT(IFNULL(cfg." & sExpresion & "apepaterno, ''), ' ', IFNULL(cfg." & sExpresion & "apematerno, ''), ', ', IFNULL(cfg." & sExpresion & "nombres, '')) AS representante, "
    s_Sql = s_Sql & "IFNULL(dci.sigladci, '') AS sigladci, IFNULL(cfg." & sExpresion & "numdocu, '') AS repnumdocu, IFNULL(cfg." & sExpresion & "cargo, '') AS repcargo, "
    s_Sql = s_Sql & sConcepto & " AS cpcimpuesto, cfg.rentaxejer_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS rentaxejerci, cfg.porcepartici, cfg.liqprn_logoemp "
    s_Sql = s_Sql & "FROM plcfgempresa cfg "
    s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON cfg." & sExpresion & "coddci=dci.coddci "
    s_Sql = s_Sql & IIf(s_OptRegistro = "certifisnp", "INNER JOIN plparametroafp afp ON cfg.pdoano=afp.pdoano ", "")
    s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    sConcepto = ""
    sPrnLogo = s_Estado_Ina
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sConcepto = gdl_Funcion.aTexto(porstRecordset!cpcimpuesto)
      sRepresentante = gdl_Funcion.aTexto(porstRecordset!representante)
      sCargoRepresenta = gdl_Funcion.aTexto(porstRecordset!repcargo)
      sCargoRepresenta = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, sCargoRepresenta, "DC")
      sNumDocumento = porstRecordset!sigladci & " " & porstRecordset!repnumdocu
      sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      sDistrito = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " - " & sDistrito
      nParticipacion = CDec(porstRecordset!rentaxejerci)
      nParticipacion = Round(nParticipacion * (CDec(porstRecordset!porcepartici) / 100), 2)
      sPrnLogo = gdl_Funcion.aTexto(porstRecordset!liqprn_logoemp)
    End If
    porstRecordset.Close
    If sConcepto = "" Then Beep: MsgBox "Debe configurar los parametros del reporte", vbExclamation: cmdAction(0).SetFocus: Exit Sub
    If (nParticipacion <= 0 And s_OptRegistro = "certifiuti") Then Beep: MsgBox "Debe configurar los parametros del reporte", vbExclamation: cmdAction(0).SetFocus: Exit Sub
    
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando " & IIf(ribAnalisis(0).Value, "Certificado", "Resumen") & " ..."
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    If ribAnalisis(0).Value Then
      gdl_Procedure.ps_ReportTitle = "CERTIFICADO DE " & IIf(s_OptRegistro = "certifi5ta", "QUINTA", IIf(s_OptRegistro = "certifiuti", "DISTRIBUCION DE UTILIDADES", Trim(cmbParametro.Text)))
      gdl_Procedure.ps_ReportName = IIf(s_OptRegistro = "certifi5ta", "rptcertiquinta", IIf(s_OptRegistro = "certifiuti", "rptcertiuti", "rptcertisnp"))
    Else
      gdl_Procedure.ps_ReportTitle = "RESUMEN " & IIf(s_OptRegistro = "certifi5ta", "DE QUINTA", IIf(s_OptRegistro = "certifiuti", "DE DISTRIBUCION DE UTILIDADES", "ANUAL DE CONTRIBUCIONES A " & Trim(cmbParametro.Text)))
      gdl_Procedure.ps_ReportName = IIf(s_OptRegistro = "certifi5ta", "rptcertiquinta", IIf(s_OptRegistro = "certifiuti", "rptresumuti", "rptresumsnp"))
    End If
    ReDim aElemento(3, 9): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = "": aElemento(0, 3) = "": aElemento(0, 4) = "": aElemento(0, 5) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    aElemento(1, 3) = "": aElemento(1, 4) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "Direccion;" & sDireccion & ";true"
    aElemento(2, 2) = "Ruc;" & ps_RucEmpresa & ";true"
    aElemento(2, 3) = "Representante;" & sRepresentante & ";true"
    aElemento(2, 4) = "Ejercicio;" & ps_Anyo & ";true"
    aElemento(2, 5) = "FechaPrn;" & UCase(sDistrito & ", " & sFechaPrn) & ";true"
    aElemento(2, 6) = "DocRepresentante;" & sNumDocumento & ";true"
    aElemento(2, 7) = "CargoRepresentante;" & sCargoRepresenta & ";true"
    aElemento(2, 8) = ""
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    ' Genera la información del reporte
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If s_OptRegistro = "certifi5ta" Then
      s_Sql = s_Sql & "codpsn varchar(11) Not Null, nombrespsn varchar(80) Null, "
      s_Sql = s_Sql & "sigladci char(3) Null, numdociden varchar(11) Null, descgo varchar(80) Null, direccionpsn varchar(100) Null, "
      s_Sql = s_Sql & "fecingreso date Null, fecbaja date Null, moneda char(3) Null, rentabruta decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "rentabruaper decimal(18,2) Null Default '0', rentabruemp decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "deduccion decimal(18,2) Null Default '0', rentaneta decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "impuestorenta decimal(18,2) Null Default '0', impurentaape decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "impurentaemp decimal(18,2) Null Default '0', impurentaing decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "PRIMARY KEY (codpsn)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      CertificadoQuinta "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
    ElseIf s_OptRegistro = "certifisnp" Then
      If ribAnalisis(0).Value Then
        aElemento(2, 8) = "EntidadPension;" & Trim(cmbParametro.Text) & ";true"
      Else
        aElemento(2, 5) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & " (" & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt) & ")" & ";true"
        aElemento(2, 8) = ""
      End If
      s_Sql = s_Sql & "codpsn varchar(11) Not Null, nombrespsn varchar(80) Null, numdociden varchar(11) Null, "
      s_Sql = s_Sql & "numeroafp varchar(15) Null, fecingreso date Null, fecbaja date Null, "
      s_Sql = s_Sql & "codcpc varchar(4) Not Null, descpc varchar(50) Null, moneda char(3) Null, "
      s_Sql = s_Sql & "remunbruta decimal(18,2) Null Default '0', retencion decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "PRIMARY KEY (codpsn, codcpc)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      CertificadoPension "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
    ElseIf s_OptRegistro = "certifiuti" Then
      If ribAnalisis(1).Value Then
        aElemento(2, 5) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & " (" & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt) & ")" & ";true"
        aElemento(2, 8) = ""
      End If
      s_Sql = s_Sql & "codpsn varchar(11) Not Null, nombrespsn varchar(80) Null, numdociden varchar(11) Null, "
      s_Sql = s_Sql & "fecingreso date Null, fecbaja date Null, moneda char(3) Null, rentabruta decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "porpartici decimal(6,2) Null Default '0', renpartici decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "partixcalculo decimal(18,2) Null Default '0', totaldias int(8) Null Default '0', "
      s_Sql = s_Sql & "diaspsn smallint(3) Null Default '0', partixdiapsn decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "totalremun decimal(18,2) Null Default '0', remunpsn decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "partixrempsn decimal(18,2) Null Default '0', partixpsn decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "PRIMARY KEY (codpsn)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      CertificadoUtilidad "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
    End If
    ' Obtengo la información del reporte
    s_Sql = "SELECT rpt.*, " & IIf(sPrnLogo = s_Estado_Act, "cfg.logo", "Null") & " AS logo, cfg.firma "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " rpt, plcfgempresa cfg "
    s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "ORDER BY codpsn"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 7), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
  
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
  Me.Height = 6730: Me.Width = 7830
  Me.Left = 505: Me.Top = 100
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
    aElemento(n_Index, 1) = Choose(n_Index + 1, "promedio", "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Parametros", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 1
    ribAnalisis(n_Index).PictureUp = LoadPicture()
    ribAnalisis(n_Index).ToolTipText = Choose(n_Index + 1, "Certificado", "Resumen Anual")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "certifica", "resumen") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribAnalisis(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  
  ribFirma.PictureUp = LoadPicture()
  ribFirma.ToolTipText = "Representante Adjunto"
  s_Sql = gdl_Procedure.ps_PathImagen & "dividir.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribFirma.PictureUp = LoadPicture(s_Sql)
  ribFirma.Value = False
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Configuro los parametos adicionales
  For n_Index = 1 To 3: cmbParametro.AddItem Choose(n_Index, "S.N.P", "ESSALUD", "E.P.S"): Next n_Index
  cmbParametro.ListIndex = 0
  gdl_Procedure.EditDTPicker "PK", dtpFecha, Date, s_MdoData_Ins, True, s_FormatoFecha, dtpShortDate
  
  lblDato(0).Visible = (s_OptRegistro = "certifisnp")
  cmbParametro.Visible = (s_OptRegistro = "certifisnp")
  tdbRegistro.DataSource = dcaRegistro
  ribParametro(0).Value = True
  ribAnalisis(0).Value = True
 
     
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
Private Sub tdbRegistro_DblClick()
  cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
