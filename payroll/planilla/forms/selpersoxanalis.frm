VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPersoxAnalisis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "selpersoxanalis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   7740
   Begin MSAdodcLib.Adodc dcaRegistro 
      Height          =   330
      Left            =   45
      Top             =   5910
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
      TabIndex        =   4
      Top             =   1005
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
         Left            =   0
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
         Top             =   1590
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
         Picture         =   "selpersoxanalis.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2010
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
         Picture         =   "selpersoxanalis.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2700
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
         Picture         =   "selpersoxanalis.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   3135
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
         Picture         =   "selpersoxanalis.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   12
         Tag             =   "0"
         Top             =   4230
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
         Picture         =   "selpersoxanalis.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   13
         Tag             =   "0"
         Top             =   4665
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
         Picture         =   "selpersoxanalis.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1155
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
         Picture         =   "selpersoxanalis.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   3555
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
         Picture         =   "selpersoxanalis.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   465
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
         Picture         =   "selpersoxanalis.frx":00EC
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   930
      Index           =   1
      Left            =   0
      TabIndex        =   0
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
      Begin VB.TextBox txtPeriodo 
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
         Height          =   285
         Left            =   1650
         MaxLength       =   8
         TabIndex        =   2
         Top             =   540
         Width           =   1150
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   6690
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
         PictureUp       =   "selpersoxanalis.frx":0108
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   6285
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
         PictureUp       =   "selpersoxanalis.frx":0124
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   7095
         TabIndex        =   20
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
         PictureUp       =   "selpersoxanalis.frx":0140
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   1
         Left            =   645
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
         PictureUp       =   "selpersoxanalis.frx":015C
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   0
         Left            =   240
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
         PictureUp       =   "selpersoxanalis.frx":0178
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   2865
         TabIndex        =   22
         Top             =   540
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
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
         Left            =   3270
         TabIndex        =   3
         Top             =   585
         Width           =   195
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo de Pago :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   1
         Top             =   585
         Width           =   1320
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   7485
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   15
      Top             =   1005
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   0
      TabIndex        =   21
      Top             =   390
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
Attribute VB_Name = "fSelPersoxAnalisis"
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
Private s_OptRegistro As String
Private opcion As Integer
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
Private Sub RentaQuinta(ByVal s_Archivo As String, ByVal s_Periodo As String, s_Proceso As String, s_FechaHora As String)
  Dim psRegistro As String, a_Parametro(12)
  Dim s_Mes As String, s_Trabajador As String
 'CAMBIO ENERO 2015
  ' Dim n_Importe As Double, a_Importe(18) As Double
  Dim n_Importe As Double, a_Importe(20) As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim aEscala() As Double
  Dim n_Escala As Double
  
  ' Recupero los parametros de analisis
  s_Mes = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, s_Periodo, "MP")
  psRegistro = "valor" & Format(s_Mes, "00")
  s_Sql = "SELECT cfg.rembasica, cfg.rempromedio, cfg.rempendiente, cfg.gratipendiente, cfg.remanterior, cfg.remganada, "
  s_Sql = s_Sql & "cfg.codcpc5ta, cfg.codcpc5ta_ing, cfg.codtbluit, cfg.codtblretener, cfg.codtblpendiente, cfg.codtbldividir, "
  s_Sql = s_Sql & "uit." & psRegistro & " AS importeuit, ret." & psRegistro & " AS mesretener, "
  s_Sql = s_Sql & "pen." & psRegistro & " AS mespendiente, diw." & psRegistro & " AS mesdividir "
  s_Sql = s_Sql & "FROM plcfgempresa cfg "
  s_Sql = s_Sql & "INNER JOIN pltablabase uit ON cfg.pdoano=uit.pdoano AND cfg.codtbluit=uit.codtbl AND uit.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "INNER JOIN pltablabase ret ON cfg.pdoano=ret.pdoano AND cfg.codtblretener=ret.codtbl AND ret.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "INNER JOIN pltablabase pen ON cfg.pdoano=pen.pdoano AND cfg.codtblpendiente=pen.codtbl AND pen.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "INNER JOIN pltablabase diw ON cfg.pdoano=diw.pdoano AND cfg.codtbldividir=diw.codtbl AND diw.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
  
  ' Obtengo los conceptos de recuperación
  a_Parametro(1) = gdl_Funcion.aTexto(porstRecordset!rembasica)
  a_Parametro(2) = gdl_Funcion.aTexto(porstRecordset!rempromedio)
  a_Parametro(3) = gdl_Funcion.aTexto(porstRecordset!rempendiente)
  a_Parametro(4) = gdl_Funcion.aTexto(porstRecordset!gratipendiente)
  a_Parametro(5) = gdl_Funcion.aTexto(porstRecordset!remanterior)
  a_Parametro(6) = gdl_Funcion.aTexto(porstRecordset!remganada)
  a_Parametro(7) = gdl_Funcion.aTexto(porstRecordset!codcpc5ta)
  a_Parametro(8) = gdl_Funcion.aTexto(porstRecordset!codcpc5ta_ing)
  a_Parametro(9) = CDec(porstRecordset!importeuit)
  a_Parametro(10) = CInt(porstRecordset!mespendiente)
  a_Parametro(11) = CInt(porstRecordset!mesretener)
  a_Parametro(12) = CInt(porstRecordset!mesdividir)
  porstRecordset.Close
  
  ' Obtengo el escala de factores
  s_Sql = "SELECT orden, numerouit, factor FROM plescalaquinta WHERE pdoanyo='" & ps_Anyo & "'  ORDER BY orden"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  n_Index = 1: n_Importe = 0
  While Not porstRecordset.EOF
    ' Incremento la dimensión del array
    ReDim Preserve aEscala(3, n_Index)
    ' Defino las escala y porcentajes
    aEscala(1, n_Index) = n_Importe
    aEscala(2, n_Index) = CDec(a_Parametro(9) * CInt(porstRecordset!numerouit))
    aEscala(3, n_Index) = CDec(porstRecordset!factor)
    n_Importe = CDec(aEscala(2, n_Index))
    porstRecordset.MoveNext
    n_Index = UBound(aEscala, 2) + 1
  Wend
  porstRecordset.Close
  
  ' Recupero la información a analizar
  s_Sql = "SELECT DISTINCTROW res.codpdo, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
  s_Sql = s_Sql & "dxr.naciextrapsn, res.codcpc, ROUND(SUM(IFNULL(res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0)), 2) AS remuneracion, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(ran.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0)), 2) AS remuneranterior "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "LEFT JOIN plresultado ran ON res.codcls=ran.codcls AND res.codpsn=ran.codpsn AND res.pdoano=ran.pdoano "
  s_Sql = s_Sql & "AND ran.codpdo='" & s_PeriodoRemAper & "' AND ran.codcpc='" & a_Parametro(5) & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & s_Periodo & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND res.codcpc IN ('" & a_Parametro(1) & "', '" & a_Parametro(2) & "', '" & a_Parametro(3) & "', '" & a_Parametro(4) & "', '" & a_Parametro(6) & "', '" & a_Parametro(7) & "', '" & a_Parametro(8) & "') "
  s_Sql = s_Sql & "GROUP BY res.codcls, res.codpsn, res.codcpc "
  s_Sql = s_Sql & "ORDER BY res.codcls, res.codpsn, res.codcpc"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Muestro la Barra de porcentaje
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    'Importe no afecto al renta de quinta
    a_Importe(10) = CDec(a_Parametro(9) * 7)
    ' Genero os arreglos de grabaciones
    a_Campos = Array("codpsn", "nompsn", "basico", "promedio", "sueldomes", "rempendiente", "grapendiente", "remanterior", "remganada", "remtotal", "remanual", "importeuit", "renneta", "renafecta1", "renafecta2", "renafecta3", "renafecta4", "renafecta5", "renafecta", "impretenido", "rentotal", "impquinta", "mespendiente", "mesretener", "mesdividir")
    a_Valores = Array("", "", CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(a_Importe(10)), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CInt(a_Parametro(10)), CInt(a_Parametro(11)), CInt(a_Parametro(12)))
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      s_Trabajador = porstRecordset!codpsn
      psRegistro = porstRecordset!naciextrapsn
      For n_Index = 1 To UBound(a_Importe): a_Importe(n_Index) = 0: Next n_Index
      a_Valores(0) = porstRecordset("codpsn"): a_Valores(1) = porstRecordset("nompsn")
      a_Importe(6) = CDec(porstRecordset("remuneranterior"))
      a_Importe(10) = CDec(a_Parametro(9) * 7)
      Do
        n_Importe = CDec(porstRecordset("remuneracion"))
        ' Remuneración basica
        If a_Parametro(1) = porstRecordset("codcpc") And n_Importe > 0 Then
          a_Importe(1) = a_Importe(1) + n_Importe
        End If
        ' Remuneración promedio
        If a_Parametro(2) = porstRecordset("codcpc") And n_Importe > 0 Then
          a_Importe(2) = a_Importe(2) + n_Importe
        End If
        ' Remuneración pendiente
        If a_Parametro(3) = porstRecordset("codcpc") And n_Importe > 0 Then
          a_Importe(4) = a_Importe(4) + n_Importe
        End If
        ' Gratificación pendiente
        If a_Parametro(4) = porstRecordset("codcpc") And n_Importe > 0 Then
          a_Importe(5) = a_Importe(5) + n_Importe
        End If
        ' Remuneración ganada
        If a_Parametro(6) = porstRecordset("codcpc") And n_Importe > 0 Then
          a_Importe(7) = a_Importe(7) + n_Importe
        End If
        ' Impuesto de quinta
        If a_Parametro(7) = porstRecordset("codcpc") And n_Importe > 0 Then
          a_Importe(20) = a_Importe(20) + n_Importe
        End If
        ' devolución de impuesto de quinta
        If a_Parametro(8) = porstRecordset("codcpc") And n_Importe > 0 Then
          a_Importe(20) = a_Importe(20) + (n_Importe * -1)
        End If
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        DoEvents
        porstRecordset.MoveNext
        ' Fin de archivo
        If porstRecordset.EOF Then Exit Do
      Loop While s_Trabajador = porstRecordset!codpsn
      '[ Realizo las operaciones complementarias
      ' -- Sueldo mensual
      a_Importe(3) = CDec(a_Importe(1) + a_Importe(2))
      ' -- Remuneración pendiente
      a_Importe(4) = CDec(a_Importe(4) - a_Importe(5))
      a_Importe(4) = IIf(psRegistro = s_Estado_Ina, a_Importe(4), 0)
      ' -- Gratificación pendiente
      a_Importe(5) = IIf(psRegistro = s_Estado_Ina, a_Importe(5), 0)
      ' -- Remuneración ganada
      a_Importe(7) = CDec(a_Importe(7) - a_Importe(6))
      ' -- Remuneración total
      a_Importe(8) = CDec(a_Importe(6) + a_Importe(7))
      ' -- Remuneración anual
      a_Importe(9) = CDec(a_Importe(4) + a_Importe(5) + a_Importe(8))
      ' -- Renta neta deducible
      a_Importe(10) = IIf(psRegistro = s_Estado_Ina, a_Importe(10), 0)
      a_Importe(11) = CDec(a_Importe(9) - a_Importe(10))
      ' -- Renta afecta tramo final
      a_Importe(16) = a_Importe(11) * aEscala(3, UBound(aEscala, 2))
      
      If psRegistro = s_Estado_Ina Then
        n_Importe = a_Importe(11)
        a_Importe(12) = CDec(0): a_Importe(13) = CDec(0): a_Importe(14) = CDec(0): a_Importe(15) = CDec(0): a_Importe(16) = CDec(0)
        If n_Importe > aEscala(1, 1) Then
          ' Obtengo la renta neta de acuerdo a la escala
          For n_Index = 1 To UBound(aEscala, 2)
            If a_Importe(11) > aEscala(1, n_Index) And a_Importe(11) <= aEscala(2, n_Index) Then
              a_Importe(11 + n_Index) = CDec(n_Importe * aEscala(3, n_Index))
              Exit For
            Else
              n_Escala = (aEscala(2, n_Index) - IIf(n_Index > 1, aEscala(2, n_Index - 1), 0))
              a_Importe(11 + n_Index) = CDec(n_Escala * aEscala(3, n_Index))
              n_Importe = n_Importe - n_Escala
            End If
          Next n_Index
        End If
      End If
      ' -- Renta afecta total
      a_Importe(17) = CDec(a_Importe(12) + a_Importe(13) + a_Importe(14) + a_Importe(15) + a_Importe(16))
      ' -- Impuesto retenido
      s_Sql = "SELECT IFNULL(SUM(importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & "), 0)  AS renta "
      s_Sql = s_Sql & "FROM plresultado "
      s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND codcpc='" & a_Parametro(7) & "' "
      s_Sql = s_Sql & "AND pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND pdomes<='" & Format(a_Parametro(11), "00") & "'"
      Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      a_Importe(18) = CDec(0)
      If Not ((porstHelp.EOF And porstHelp.BOF) Or porstHelp.RecordCount = 0) Then
        a_Importe(18) = CDec(porstHelp!renta)
      End If
      porstHelp.Close
      ' -- Renta total
      a_Importe(19) = CDec(a_Importe(17) - a_Importe(18))
      ']
      ' Impuesto es diferente de cero
      If (a_Importe(20) <> 0 Or (a_Importe(20) = 0 And ribAnalisis(1).Value)) Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        For n_Index = 1 To UBound(a_Importe): a_Valores(1 + n_Index) = Round(CDec(a_Importe(n_Index)), 2): Next n_Index
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      End If
    Wend
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub

Private Sub cmdAction_Click(Index As Integer)
  Dim s_FechaHora As String, s_OldMessage As String
  Dim s_Representante As String, s_RegisPatronal As String
  Dim s_Distrito  As String, s_FechaReporte As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Visualizar o analizar registro
    If txtPeriodo.Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0).Caption = "" Or lblHelp(0).Caption = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    Me.Tag = s_MdoData_Vis
    fPrmRentaQuinta.Show vbModal
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
    If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Representante = "": s_RegisPatronal = "": s_Distrito = "": s_FechaReporte = ""
    
    ' Cambio el Mensaje
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "ANÁLISIS DE RENTA DE QUINTA"
    gdl_Procedure.ps_ReportName = "cstrentaquinta"
    ReDim aElemento(3, 7): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    aElemento(2, 2) = "Periodo;" & Trim(txtPeriodo.Text) & " - " & Trim(lblHelp(0).Caption) & ";true"
    aElemento(2, 3) = "": aElemento(2, 4) = ""
    aElemento(2, 5) = "": aElemento(2, 6) = ""
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ("
    s_Sql = s_Sql & "codpsn varchar(11) Not Null, nompsn varchar(80) Not Null, "
    s_Sql = s_Sql & "basico decimal(18,2) Null Default '0', promedio decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "sueldomes decimal(18,2) Null Default '0', rempendiente decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "grapendiente decimal(18,2) Null Default '0', remanterior decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "remganada decimal(18,2) Null Default '0', remtotal decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "remanual decimal(18,2) Null Default '0', importeuit decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "renneta decimal(18,2) Null Default '0', renafecta1 decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "renafecta2 decimal(18,2) Null Default '0', renafecta3 decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "renafecta4 decimal(18,2) Null Default '0', renafecta5 decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "renafecta decimal(18,2) Null Default '0', impretenido decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "rentotal decimal(18,2) Null Default '0', impquinta decimal(18,2) Null Default '0', "
    s_Sql = s_Sql & "mespendiente smallint(2) Null Default '0', mesretener smallint(2) Null Default '0', "
    s_Sql = s_Sql & "mesdividir smallint(2) Null Default '0', "
    s_Sql = s_Sql & "PRIMARY KEY (codpsn, nompsn))"
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    ' Proceso la informacion de quinta
    RentaQuinta "tmp" & gdl_Procedure.ps_ReportName, Trim(txtPeriodo.Text), s_OptRegistro, s_FechaHora
    ' Información de reporte
    s_Sql = "SELECT * FROM tmp" & gdl_Procedure.ps_ReportName
    s_Sql = s_Sql & " ORDER BY " & aElemento(0, 1)
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

Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Periodo de Pago
    gdl_Procedure.DefineStyleGrilla tdbHelp, "Periodo de Pago", 2
    tdbHelp.Columns(0).DataField = "codpdo": tdbHelp.Columns(1).DataField = "despdo"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("ped", "codpdo", s_Estado_Ina & ps_ClsPlanilla & ps_Anyo, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = panToolBar(1).Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
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
  Me.Height = 6705: Me.Width = 7830
  Me.Left = 105: Me.Top = 180
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
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Configuración de Parametros", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codbco", "desbco")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 934.7402, 3255.071)
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Periodo de Pago", 2
  ']
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribAnalisis(0).Value = True
  
  ' Cargo los graficos de los botones de analisis
  For n_Index = 0 To 1
    ribAnalisis(n_Index).PictureUp = LoadPicture()
    ribAnalisis(n_Index).ToolTipText = "Análisis " & Choose(n_Index + 1, "Retenciones", "General")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "anctpend", "ancthist") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribAnalisis(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  ribParametro(0).Value = True
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Periodo de pago
    txtPeriodo = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtPeriodo.SetFocus
  End Select
   
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Periodo de Pago
    s_Sql = gdl_Funcion.HelpTablas("ped", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla & ps_Anyo, "")
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
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo, "PR")
End Sub
