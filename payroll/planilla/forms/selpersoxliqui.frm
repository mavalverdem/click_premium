VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPersoxLiquida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   Icon            =   "selpersoxliqui.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   8670
   Begin MSAdodcLib.Adodc dcaRegistro 
      Height          =   330
      Left            =   45
      Top             =   5895
      Width           =   7790
      _ExtentX        =   13732
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
      Left            =   7910
      TabIndex        =   5
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
         Left            =   0
         TabIndex        =   15
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
         TabIndex        =   8
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
         Picture         =   "selpersoxliqui.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   9
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
         Picture         =   "selpersoxliqui.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   10
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
         Picture         =   "selpersoxliqui.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   11
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
         Picture         =   "selpersoxliqui.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   13
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
         Picture         =   "selpersoxliqui.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   14
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
         Picture         =   "selpersoxliqui.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   7
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
         Picture         =   "selpersoxliqui.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   12
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
         Picture         =   "selpersoxliqui.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   6
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
         Picture         =   "selpersoxliqui.frx":00EC
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   930
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8670
      _Version        =   65536
      _ExtentX        =   15293
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
         Height          =   280
         Left            =   3135
         MaxLength       =   8
         TabIndex        =   2
         Top             =   165
         Width           =   1150
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   7605
         TabIndex        =   22
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
         PictureUp       =   "selpersoxliqui.frx":0108
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   7200
         TabIndex        =   21
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
         PictureUp       =   "selpersoxliqui.frx":0124
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   8010
         TabIndex        =   23
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
         PictureUp       =   "selpersoxliqui.frx":0140
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   1
         Left            =   630
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
         PictureUp       =   "selpersoxliqui.frx":015C
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   0
         Left            =   225
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
         PictureUp       =   "selpersoxliqui.frx":0178
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   2
         Left            =   1035
         TabIndex        =   18
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
         PictureUp       =   "selpersoxliqui.frx":0194
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   4350
         TabIndex        =   26
         Top             =   165
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSRibbon ribFirma 
         Height          =   360
         Left            =   8010
         TabIndex        =   24
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
         PictureUp       =   "selpersoxliqui.frx":01B0
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   1740
         TabIndex        =   4
         Top             =   450
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   141754369
         CurrentDate     =   37515
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   3
         Left            =   225
         TabIndex        =   19
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
         PictureUp       =   "selpersoxliqui.frx":01CC
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
         Left            =   3135
         TabIndex        =   3
         Top             =   525
         Width           =   195
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo de Pago :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1770
         TabIndex        =   1
         Top             =   210
         Width           =   1320
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   0
         Left            =   1545
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   5535
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   20
      Top             =   990
      Width           =   7790
      _ExtentX        =   13732
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
      TabIndex        =   25
      Top             =   795
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
Attribute VB_Name = "fSelPersoxLiquida"
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

' Instancia del formulario activo
'[
Private Sub BoletaLiquida(ByVal s_Archivo As String, ByVal s_Periodo As String, s_Proceso As String, s_FechaHora As String)
  Dim porstDetalle As ADODB.Recordset
  Dim s_Ano As String, s_Mes As String, s_Trabajador As String
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  Dim sFechaIngreso As String, sFechaCese As String, sFechaInicio As String
  Dim sServicio As String, sMoneda As String, sMonedaPago As String, sDesMoneda As String
  Dim nDetalle As Long, nMesTrunco As Integer, nDiaTrunco As Integer
  Dim nAnoVacacion As Integer, nMesVacacion As Integer, nDiaVacacion As Integer
  Dim sDetalle As String, sPeriodoCts As String, sColumna As String
  Dim nColumna As Integer, nContador As Integer
  Dim aDetalle()
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  ' Obtengo los datos del personal
  s_Mes = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, s_Periodo, "MP")
  s_Sql = "SELECT DISTINCTROW psn.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nompsn, "
  s_Sql = s_Sql & "IFNULL(dxr.codcgo, '') AS codcgo, IFNULL(cgo.descgo, '') AS descgo, dxr.fecingreso, IFNULL(psn.nroessalud, '') AS nroessalud, "
  s_Sql = s_Sql & "IFNULL(psn.numdociden, '') AS numdociden, IFNULL(psn.numeroafp, '') AS numeroafp, IFNULL(dxr.codafp, '') AS codafp, IFNULL(afp.desafp, '') AS desafp, "
  s_Sql = s_Sql & "asi.fechacese, asi.observacion, asi.dialiquidacion, asi.liquidavacacion, asi.diagratificacion, asi.fechainiliqvaca, asi.fechafinliqvaca, "
  s_Sql = s_Sql & "IF(psn.pagodolar='" & s_Estado_Act & "', '" & s_Codmon_me & "', '" & s_Codmon_mn & "') AS monpago, IFNULL(pdo.despdo, '') AS despdo, pdo.tipocambio, "
  s_Sql = s_Sql & "asi.fechacese, asi.observacion, asi.dialiquidacion, asi.liquidavacacion, asi.diagratificacion, asi.fechainiliqvaca, asi.fechafinliqvaca, "
  s_Sql = s_Sql & "par.remupromeliq, par.remuvacaliq, par.remuvacatrun, par.remuliquiex, par.remuliquisu, par.remubasicacts, par.remupromects, par.remugraticts, "
  s_Sql = s_Sql & "res.pdoano, res.pdomes, IFNULL(cco.detcco, '') detcco "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plparametroafp par ON res.pdoano=par.pdoano "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
  s_Sql = s_Sql & "AND dxr.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "ORDER BY codpsn"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Imprimiendo Liquidación ..."
    fMenu.panPercent.Visible = True
    sMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    sDesMoneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Nom, s_Codmon_me_Nom)
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Ano = porstRecordset!pdoano
    s_Mes = porstRecordset!pdomes
    
    ' Genero os arreglos de grabaciones
    a_Campos = Array("codpsn", "nompsn", "codcgo", "descgo", "fecingreso", "fecbaja", "observacion", "servicio", "numdociden", "nroessalud", "numeroafp", "codafp", "desafp", "despdo", "fechainicts", "fechafincts", "numeromeses", "numerodias", "fechainiliqvaca", "fechafinliqvaca", "anoliquidavaca", "mesliquidavaca", "dialiquidavaca", "fechainivacatru", "fechafinvacatru", "mesvacatrun", "diavacatrun", "dialiquidacion", "detalle", "pdocts", "secuencia", "codcpcing", "descpcing", "impcpcing", "codcpcdsc", "descpcdsc", "impcpcdsc", "codcpcapo", "descpcapo", "impcpcapo", "moneda", "monpago", "importipcmb", "impornetocmb", "detcco", "desmoneda")
    a_Valores = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(0), "", "", CLng(0), "", "", CDec(0), "", "", CDec(0), "", "", CDec(0), "", "", CDec(0), CDec(0), "", "")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      s_Trabajador = porstRecordset!codpsn
      sMonedaPago = IIf(porstRecordset!monpago = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt)
      sFechaIngreso = Format(porstRecordset!fecingreso, s_FormatoFecha)
      sFechaCese = Format(porstRecordset!fechacese, s_FormatoFecha)
      nDetalle = 0: sDetalle = "a": sPeriodoCts = "0000-0"
      
      ' Tiempo de servicio
      nContador = gdl_Funcion.NumeroDias360(CDate(sFechaCese), CDate(sFechaIngreso), CDate(sFechaCese))
      nColumna = nContador \ 360
      sServicio = nColumna & " Año(s) "
      
      nContador = nContador - (nColumna * 360)
      nColumna = nContador \ 30
      sServicio = sServicio & " " & nColumna & " mes(es) "
      
      nColumna = nContador - (nColumna * 30)
      sServicio = sServicio & " y " & nColumna & " día(s) "
      
      ' Primer Ingresos remunarativos
      s_Sql = "SELECT res.secuencia, res.codcpc, cpc.descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plparametroafp prm ON res.pdoano=prm.pdoano AND res.codcpc=prm.cpcbasico "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.secuencia, res.codcpc, cpc.descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plparametroafp prm ON res.pdoano=prm.pdoano AND res.codcpc=prm.cpcafamiliar "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT res.secuencia, res.codcpc, cpc.descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.tipocpc='" & s_Estado_Ina & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.codcpc='" & porstRecordset!remupromeliq & "' "
      s_Sql = s_Sql & "ORDER BY secuencia"
      Set porstDetalle = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstDetalle.BOF And porstDetalle.EOF) Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        While Not porstDetalle.EOF
          nDetalle = nDetalle + 1
          a_Valores = Array(s_Trabajador, porstRecordset!nompsn, porstRecordset!codcgo, porstRecordset!descgo, Format(sFechaIngreso, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!observacion), sServicio, porstRecordset!numdociden, porstRecordset!nroessalud, porstRecordset!numeroafp, porstRecordset!codafp, porstRecordset!desafp, porstRecordset!despdo, "", "", CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(porstRecordset!dialiquidacion), sDetalle, sPeriodoCts, nDetalle, porstDetalle!codcpc, porstDetalle!descpc, CDec(porstDetalle!imporingreso), "", "", CDec(0), "", "", CDec(0), sMoneda, sMonedaPago, CDec(porstRecordset!Tipocambio), CDec(0), porstRecordset!detcco, sDesMoneda)
          If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
          porstDetalle.MoveNext
        Wend
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      End If
      porstDetalle.Close
      nDetalle = 0: sDetalle = "b": sPeriodoCts = "0000-0"
      
      ' Segundo detalle Compensacion tiempo de servicio
      s_Sql = "DROP TABLE IF EXISTS tmpliquidasub"
      If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
      
      s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS tmpliquidasub "
      s_Sql = s_Sql & "SELECT DISTINCTROW mov.codcls, mov.pdocts, MAX(CONCAT(mov.pdoano, mov.subcts)) AS cPrimaryKey, "
      s_Sql = s_Sql & "MAX(mov.pdoano) AS pdoano, MAX(mov.subcts) AS subcts, mov.codpsn "
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
      s_Sql = s_Sql & "INNER JOIN plctsperiodosub sub ON res.codcls=sub.codcls AND res.pdocts=sub.pdocts AND res.subcts=sub.subcts "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.codcpc IN('" & porstRecordset!remubasicacts & "', '" & porstRecordset!remupromects & "', '" & porstRecordset!remugraticts & "') "
      s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "GROUP BY mov.codcls, mov.pdocts "
      s_Sql = s_Sql & "ORDER BY pdocts, subcts"
      If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
      
      s_Sql = "SELECT res.pdocts, res.subcts, res.secuencia, res.codcpc, cpc.descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb, "
      s_Sql = s_Sql & "mov.fechaini, mov.fechafin, mov.porinteres, mov.numeroanos, mov.numeromeses, mov.numerodias "
      s_Sql = s_Sql & "FROM plctsresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.pdoano=mov.pdoano AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
      s_Sql = s_Sql & "INNER JOIN tmpliquidasub sub ON res.codcls=sub.codcls AND res.pdocts=sub.pdocts AND res.pdoano=sub.pdoano AND res.subcts=sub.subcts "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.codcpc IN('" & porstRecordset!remubasicacts & "', '" & porstRecordset!remupromects & "', '" & porstRecordset!remugraticts & "') "
      s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "ORDER BY pdocts, subcts, secuencia"
      Set porstDetalle = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstDetalle.BOF And porstDetalle.EOF) Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        While Not porstDetalle.EOF
          sPeriodoCts = porstDetalle!pdocts
          ' Tiempo de servicio
          nMesTrunco = CInt(porstDetalle!numeromeses)
          nDiaTrunco = CInt(porstDetalle!numerodias)
          nDiaTrunco = gdl_Funcion.NumeroDias360(CDate(sFechaCese), CDate(porstDetalle!fechaini), CDate(sFechaCese))
          nColumna = nDiaTrunco \ 360
          nDiaTrunco = nDiaTrunco - (nColumna * 360)
          nMesTrunco = nDiaTrunco \ 30
          nDiaTrunco = nDiaTrunco - (nMesTrunco * 30)
          
          nDetalle = nDetalle + 1
          a_Valores = Array(s_Trabajador, porstRecordset!nompsn, porstRecordset!codcgo, porstRecordset!descgo, Format(sFechaIngreso, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!observacion), sServicio, porstRecordset!numdociden, porstRecordset!nroessalud, porstRecordset!numeroafp, porstRecordset!codafp, porstRecordset!desafp, porstRecordset!despdo, Format(porstDetalle!fechaini, s_FmtFechMysql_0), Format(porstDetalle!fechafin, s_FmtFechMysql_0), CInt(nMesTrunco), CInt(nDiaTrunco), "", "", CInt(0), CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(porstRecordset!dialiquidacion), sDetalle, sPeriodoCts, nDetalle, porstDetalle!codcpc, porstDetalle!descpc, CDec(porstDetalle!imporingreso), "", "", CDec(0), "", "", CDec(0), sMoneda, sMonedaPago, CDec(porstRecordset!Tipocambio), CDec(0), porstRecordset!detcco, sDesMoneda)
          If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
          porstDetalle.MoveNext
        Wend
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      End If
      porstDetalle.Close
      nDetalle = 0: sDetalle = "c": sPeriodoCts = "0000-0"
      
      ' Tercer detalle Vacaciones pendientes y truncas
      nDiaVacacion = CInt(porstRecordset!liquidavacacion)
      nAnoVacacion = nDiaVacacion \ 360
      nDiaVacacion = nDiaVacacion - (nAnoVacacion * 360)
      nMesVacacion = nDiaVacacion \ 30
      nDiaVacacion = nDiaVacacion - (nMesVacacion * 30)
      
      nDiaTrunco = 0
      sFechaInicio = Left(sFechaIngreso, 6) & IIf((s_Ano = Right(sFechaIngreso, 4)) Or (s_Ano <> Right(sFechaIngreso, 4) And (s_Mes & Left(sFechaCese, 2) > Mid(sFechaIngreso, 4, 2) & Left(sFechaIngreso, 2))), s_Ano + 1, s_Ano)
      sFechaInicio = Format(DateAdd("yyyy", -1, CDate(sFechaInicio)), s_FormatoFecha)
      If sFechaInicio <> "" And sFechaCese <> "" Then
        nDiaTrunco = gdl_Funcion.NumeroDias360(CDate(sFechaCese), CDate(sFechaInicio), CDate(sFechaCese))
      End If
      nMesTrunco = nDiaTrunco \ 30
      nDiaTrunco = nDiaTrunco - (nMesTrunco * 30)
      
      s_Sql = "SELECT 1 AS secuencia, res.codcpc, 'Vacaciones Devengadas' AS descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.codcpc='" & porstRecordset!remuvacaliq & "' "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT 2 AS secuencia, res.codcpc, 'Vacaciones Truncas' AS descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.codcpc='" & porstRecordset!remuvacatrun & "' "
      s_Sql = s_Sql & "ORDER BY secuencia"
      Set porstDetalle = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstDetalle.BOF And porstDetalle.EOF) Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        While Not porstDetalle.EOF
          nDetalle = CInt(porstDetalle!secuencia)
          a_Valores = Array(s_Trabajador, porstRecordset!nompsn, porstRecordset!codcgo, porstRecordset!descgo, Format(sFechaIngreso, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!observacion), sServicio, porstRecordset!numdociden, porstRecordset!nroessalud, porstRecordset!numeroafp, porstRecordset!codafp, porstRecordset!desafp, porstRecordset!despdo, "", "", CInt(0), CInt(0), Format(porstRecordset!fechainiliqvaca, s_FmtFechMysql_0), Format(porstRecordset!fechafinliqvaca, s_FmtFechMysql_0), nAnoVacacion, nMesVacacion, nDiaVacacion, Format(sFechaInicio, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), nMesTrunco, nDiaTrunco, CInt(porstRecordset!dialiquidacion), sDetalle, sPeriodoCts, nDetalle, porstDetalle!codcpc, porstDetalle!descpc, CDec(porstDetalle!imporingreso), "", "", CDec(0), "", "", CDec(0), sMoneda, sMonedaPago, CDec(porstRecordset!Tipocambio), CDec(0), porstRecordset!detcco, sDesMoneda)
          If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
          porstDetalle.MoveNext
        Wend
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      End If
      porstDetalle.Close
      nDetalle = 0: sDetalle = "d": sPeriodoCts = "0000-0"
      
      ' Cuarto detalle Gratificación extraordinaria, suma graciosa
      s_Sql = "SELECT 1 AS secuencia, res.codcpc, cpc.descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.codcpc='" & porstRecordset!remuliquiex & "' "
      s_Sql = s_Sql & "UNION "
      s_Sql = s_Sql & "SELECT 2 AS secuencia, res.codcpc, cpc.descpc, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS imporingreso, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.codcpc='" & porstRecordset!remuliquisu & "' "
      s_Sql = s_Sql & "ORDER BY secuencia"
      Set porstDetalle = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstDetalle.BOF And porstDetalle.EOF) Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        While Not porstDetalle.EOF
          nDetalle = CInt(porstDetalle!secuencia)
          a_Valores = Array(s_Trabajador, porstRecordset!nompsn, porstRecordset!codcgo, porstRecordset!descgo, Format(sFechaIngreso, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!observacion), sServicio, porstRecordset!numdociden, porstRecordset!nroessalud, porstRecordset!numeroafp, porstRecordset!codafp, porstRecordset!desafp, porstRecordset!despdo, "", "", CInt(0), CInt(0), Format(porstRecordset!fechainiliqvaca, s_FmtFechMysql_0), Format(porstRecordset!fechafinliqvaca, s_FmtFechMysql_0), nAnoVacacion, nMesVacacion, nDiaVacacion, Format(sFechaInicio, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), nMesTrunco, nDiaTrunco, CInt(porstRecordset!dialiquidacion), sDetalle, sPeriodoCts, nDetalle, porstDetalle!codcpc, porstDetalle!descpc, CDec(porstDetalle!imporingreso), "", "", CDec(0), "", "", CDec(0), sMoneda, sMonedaPago, CDec(porstRecordset!Tipocambio), CDec(0), porstRecordset!detcco, sDesMoneda)
          If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
          porstDetalle.MoveNext
        Wend
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      End If
      porstDetalle.Close
      nDetalle = 0: sDetalle = "e": sPeriodoCts = "0000-0"
      
      ' Quinto detalle resumen de liquidación
      s_Sql = "SELECT res.secuencia, res.codcpc, cpc.descpc, res.tipocpc, "
      s_Sql = s_Sql & "IF(res.tipocpc='" & s_Estado_Ina & "', res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS imporingreso, "
      s_Sql = s_Sql & "IF(res.tipocpc='" & s_Estado_Act & "', res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS impordescto, "
      s_Sql = s_Sql & "IF(res.tipocpc='" & s_Estado_Blq & "', res.importe_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS imporaporte, "
      s_Sql = s_Sql & "res.importe_" & IIf(fMenu.ribMoneda(0).Value, "me", "mn") & " AS importecmb "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & s_Trabajador & "' "
      s_Sql = s_Sql & "AND res.impbolecpc='" & s_Estado_Act & "' "
      s_Sql = s_Sql & "AND res.codpsn='" & porstRecordset("codpsn") & "' "
      s_Sql = s_Sql & "ORDER BY tipocpc, secuencia"
      Set porstDetalle = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      If Not (porstDetalle.BOF And porstDetalle.EOF) Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        nColumna = 6: nContador = 0
        ReDim aDetalle(6, 0)
        Do While Not porstDetalle.EOF
          ' Selecciono el tipo de concepto
          If nColumna <> CInt(porstDetalle!tipocpc) Then
            nColumna = CInt(porstDetalle!tipocpc)
            sColumna = Choose(nColumna + 1, "imporingreso", "impordescto", "imporaporte")
            nContador = 0
          End If
          ' Actualizo los ingresos de liquidación
          If porstDetalle!tipocpc = s_Estado_Ina Then
            nDetalle = nDetalle + 1
            a_Valores = Array(s_Trabajador, porstRecordset!nompsn, porstRecordset!codcgo, porstRecordset!descgo, Format(sFechaIngreso, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!observacion), sServicio, porstRecordset!numdociden, porstRecordset!nroessalud, porstRecordset!numeroafp, porstRecordset!codafp, porstRecordset!desafp, porstRecordset!despdo, "", "", CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(porstRecordset!dialiquidacion), sDetalle, sPeriodoCts, nDetalle, porstDetalle!codcpc, porstDetalle!descpc, CDec(porstDetalle(sColumna)), "", "", CDec(0), "", "", CDec(0), sMoneda, sMonedaPago, CDec(porstRecordset!Tipocambio), CDec(0), porstRecordset!detcco, sDesMoneda)
            If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
          Else
            nContador = nContador + 1
            ' Redimensiono e inicializo el arreglo de los detalles
            If nContador > UBound(aDetalle, 2) Then
              ReDim Preserve aDetalle(6, nContador)
              aDetalle(1, nContador) = "": aDetalle(2, nContador) = ""
              aDetalle(3, nContador) = "": aDetalle(4, nContador) = ""
              aDetalle(5, nContador) = CDec(0): aDetalle(6, nContador) = CDec(0)
            End If
            ' Asigno los datos al arreglo
            aDetalle(nColumna + 0, nContador) = porstDetalle("codcpc")
            aDetalle(nColumna + 2, nContador) = porstDetalle("descpc")
            aDetalle(nColumna + 4, nContador) = CDec(porstDetalle(sColumna))
          End If
          porstDetalle.MoveNext
        Loop
        porstDetalle.Close
        nDetalle = 0: sDetalle = "f": sPeriodoCts = "0000-0"
      
        ' Actualizo los descuentos y aportes de liquidación
        For n_Index = 1 To UBound(aDetalle, 2)
          nDetalle = nDetalle + 1
          a_Valores = Array(s_Trabajador, porstRecordset!nompsn, porstRecordset!codcgo, porstRecordset!descgo, Format(sFechaIngreso, s_FmtFechMysql_0), Format(sFechaCese, s_FmtFechMysql_0), gdl_Funcion.aTexto(porstRecordset!observacion), sServicio, porstRecordset!numdociden, porstRecordset!nroessalud, porstRecordset!numeroafp, porstRecordset!codafp, porstRecordset!desafp, porstRecordset!despdo, "", "", CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(0), "", "", CInt(0), CInt(0), CInt(porstRecordset!dialiquidacion), sDetalle, sPeriodoCts, nDetalle, "", "", CDec(0), aDetalle(1, n_Index), aDetalle(3, n_Index), aDetalle(5, n_Index), aDetalle(2, n_Index), aDetalle(4, n_Index), aDetalle(6, n_Index), sMoneda, sMonedaPago, CDec(porstRecordset!Tipocambio), CDec(0), porstRecordset!detcco, sDesMoneda)
          If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        Next n_Index
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
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
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
  Dim s_Representante As String, s_CargoRepresentante As String
  Dim s_RegisPatronal As String, s_Expresion As String
  Dim s_Distrito  As String, s_Provincia As String
  Dim s_PrnEmpresa As String, s_PrnLogo As String
  Dim sFechaPrn As String
    
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Visualizar o analizar registro
    If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    Me.Tag = s_MdoData_Vis
    If s_OptRegistro = "repliquida" Then
      fPrmBoleLiquida.Show vbModal
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
    If txtPeriodo.Text = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0).Caption = "" Or lblHelp(0).Caption = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Representante = "": s_RegisPatronal = "": s_Distrito = "": s_Provincia = ""
    s_Expresion = IIf(((ribAnalisis(0).Value Or ribAnalisis(1).Value) And ribFirma.Value), "ger", "rep")
    sFechaPrn = Format(dtpFecha, "dd") & " de " & gdl_Funcion.NombreMes(Format(dtpFecha, "mm")) & " del " & Format(dtpFecha, "yyyy")
    
    ' Verifico que existan parametros de boletas
    s_Sql = "SELECT prm.remupromeliq, prm.remuvacaliq, prm.remuvacatrun, prm.remuliquiex, prm.remuliquisu, "
    s_Sql = s_Sql & "cfg.ubigeodir, IFNULL(cfg.regpatronal, '') AS regpatronal, cfg.repimpbol, "
    s_Sql = s_Sql & "CONCAT(IFNULL(cfg." & s_Expresion & "nombres, '') , ' ', IFNULL(cfg." & s_Expresion & "apepaterno, ''), ' ', IFNULL(cfg." & s_Expresion & "apematerno, '')) AS representante, "
    s_Sql = s_Sql & "cfg." & s_Expresion & "cargo AS repcargo, cfg." & s_Expresion & "coddci AS repcoddci, cfg." & s_Expresion & "numdocu AS repnumdocu, cfg.girocomercial, "
    s_Sql = s_Sql & "cfg.liqprn_razonemp, cfg.liqprn_logoemp "
    s_Sql = s_Sql & "FROM plparametroafp prm, plcfgempresa cfg "
    s_Sql = s_Sql & "WHERE prm.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND cfg.pdoano=prm.pdoano"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de exportación", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!remupromeliq) = "" Then Beep: MsgBox "Debe configurar parametro de Remuneración Liquidación - Promedio", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!remuvacaliq) = "" Then Beep: MsgBox "Debe configurar parametro de Vacaciones de Liquidación", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!remuvacatrun) = "" Then Beep: MsgBox "Debe configurar parametro de Vacaciones Truncas", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!remuliquiex) = "" Then Beep: MsgBox "Debe configurar parametro de Remuneración Liquidación - Extraordinaria", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!remuliquisu) = "" Then Beep: MsgBox "Debe configurar parametro de Remuneración Liquidación - Compensación", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!representante) = ", " Then Beep: MsgBox "Debe configurar el parametro de Representante Legal", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!repnumdocu) = ", " Then Beep: MsgBox "Debe configurar el parametro de Documento Representante Legal", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!regpatronal) = "" Then Beep: MsgBox "Debe configurar el parametro Regimen Patronal", vbCritical: Exit Sub
    If gdl_Funcion.aTexto(porstRecordset!girocomercial) = "" Then Beep: MsgBox "Debe configurar el parametro Giro Comercial", vbCritical: Exit Sub
    
    s_Representante = porstRecordset!representante
    s_CargoRepresentante = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, porstRecordset!repcargo, "DC")
    s_RegisPatronal = porstRecordset!regpatronal
    s_Distrito = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
    s_Provincia = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Act, Left(s_Distrito, 4), "UB")
    s_Distrito = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, s_Distrito, "UB")
    s_PrnEmpresa = gdl_Funcion.aTexto(porstRecordset!liqprn_razonemp)
    s_PrnLogo = gdl_Funcion.aTexto(porstRecordset!liqprn_logoemp)
    
    ' Cambio el Mensaje
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Procesando Información ..."
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = IIf(ribAnalisis(0).Value, "LIQUIDACION DE BENEFICIOS SOCIALES", IIf(ribAnalisis(1).Value, "CERTIFICADO DE TRABAJO", IIf(ribAnalisis(2).Value, "CARTA DE AUTORIZACIÓN DE RETIRO DE CTS", "CONSTANCIA DE TRABAJO")))
    gdl_Procedure.ps_ReportName = IIf(ribAnalisis(0).Value, "rptboliquida", IIf(ribAnalisis(1).Value, "rptcertiliq", IIf(ribAnalisis(2).Value, "rptretirocts", "rptconstatra")))
    ReDim aElemento(3, 9): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    aElemento(2, 2) = "Ruc;" & ps_RucEmpresa & ";true"
    aElemento(2, 3) = "Representante;" & s_Representante & ";true"
    aElemento(2, 4) = "CargoRepresenta;" & s_CargoRepresentante & ";true"
    aElemento(2, 5) = "Provincia;" & UCase(s_Provincia) & ";true"
    aElemento(2, 6) = "Distrito;" & UCase(s_Distrito) & ";true"
    aElemento(2, 7) = "FechaPrn;" & UCase(s_Distrito & ", " & sFechaPrn) & ";true"
    aElemento(2, 8) = "VisualizaEmpresa;" & s_PrnEmpresa & ";true"
    
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    If ribAnalisis(0).Value Then
      aElemento(2, 5) = "RegPatronal;" & s_RegisPatronal & ";true"
      ' [ Generación e impresión de información para el reporte
      s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      
      s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
      s_Sql = s_Sql & "codpsn varchar(11) Not Null, nompsn varchar(80) Null, "
      s_Sql = s_Sql & "codcgo char(3) Null, descgo varchar(80) Null, "
      s_Sql = s_Sql & "fecingreso date Null, fecbaja date Null, "
      s_Sql = s_Sql & "observacion varchar(60) Null, servicio varchar(40) Null, "
      s_Sql = s_Sql & "numdociden varchar(11) Null, nroessalud varchar(15) Null, "
      s_Sql = s_Sql & "numeroafp varchar(15) Null, codafp char(2) Null, desafp varchar(40) Null, "
      s_Sql = s_Sql & "despdo varchar(40) Null, fechainicts date Null, fechafincts date Null, "
      s_Sql = s_Sql & "numeromeses smallint(2) Null Default '0', numerodias smallint(2) Null Default '0', "
      s_Sql = s_Sql & "fechainiliqvaca date Null, fechafinliqvaca date Null, anoliquidavaca smallint(2) Null Default '0', "
      s_Sql = s_Sql & "mesliquidavaca smallint(2) Null Default '0', dialiquidavaca smallint(2) Null Default '0', "
      s_Sql = s_Sql & "fechainivacatru date Null, fechafinvacatru date Null, mesvacatrun smallint(2) Null Default '0', "
      s_Sql = s_Sql & "diavacatrun smallint(2) Null Default '0', dialiquidacion smallint(3) Null Default '0', "
      s_Sql = s_Sql & "detalle char(1) Not Null, pdocts varchar(6) Not Null, secuencia smallint(3) Not Null, "
      s_Sql = s_Sql & "codcpcing varchar(4) Null, descpcing varchar(40) Null, impcpcing decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcdsc varchar(4) Null, descpcdsc varchar(40) Null, impcpcdsc decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "codcpcapo varchar(4) Null, descpcapo varchar(40) Null, impcpcapo decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "moneda char(3) Null, monpago char(3) Null, importipcmb decimal(6,3) Null Default '0', impornetocmb decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "detcco varchar(50) Null, desmoneda varchar(100), "
      s_Sql = s_Sql & "PRIMARY KEY (codpsn, detalle, pdocts, secuencia))"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' Proceso la informacion de liquidacion
      BoletaLiquida "tmp" & gdl_Procedure.ps_ReportName, Trim(txtPeriodo.Text), s_OptRegistro, s_FechaHora
      ' Recupero informacion
      s_Sql = "SELECT liq.*, " & IIf(s_PrnLogo = s_Estado_Act, "cfg.logo", "Null") & " AS logo, cfg.firma "
      s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " liq, plcfgempresa cfg "
      s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "ORDER BY codpsn, detalle, pdocts, secuencia"
    ElseIf (ribAnalisis(1).Value Or ribAnalisis(3).Value) Then
      s_Sql = "SELECT DISTINCT psn.codpsn, CONCAT(IFNULL(psn.apepaterno,''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrespsn, "
      s_Sql = s_Sql & "dci.sigladci, psn.numdociden, psn.fecingreso, psn.fecbaja, cgo.descgo, asi.liqnocalifica, "
      s_Sql = s_Sql & "CONCAT(IFNULL(zon.abrezona, ''), ' ', IFNULL(cfg.direccionzona, '')) AS empzona, "
      s_Sql = s_Sql & "CONCAT(IFNULL(via.abrevia, ''), ' ', IFNULL(cfg.direccionvia, ''), ' ', IFNULL(cfg.numerodir, '')) AS empvia, "
      s_Sql = s_Sql & "cfg.girocomercial, dcr.sigladci AS sigladcirep, cfg." & s_Expresion & "numdocu AS repnumdocu, " & IIf(s_PrnLogo = s_Estado_Act, "cfg.logo", "Null") & " AS logo, "
      s_Sql = s_Sql & "cfg.firma" & IIf(s_Expresion = "rep", "", "nexo") & " AS firma "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plresultado res ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
      s_Sql = s_Sql & "INNER JOIN pldocidentidad dci ON dci.coddci=psn.coddci "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON cgo.codcls=psn.codcls AND cgo.codcgo=psn.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plcfgempresa cfg ON cfg.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON zon.codzona=cfg.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON via.codvia=cfg.codvia "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dcr ON dcr.coddci=cfg." & s_Expresion & "coddci "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion rng "
      s_Sql = s_Sql & "WHERE rng.proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND rng.usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND rng.fyhcre='" & s_FechaHora & "') "
      If ribAnalisis(1).Value Then
        s_Sql = s_Sql & "AND psn.estadopsn='I' "
      End If
      s_Sql = s_Sql & "ORDER BY " & aElemento(0, 1)
    ElseIf ribAnalisis(2).Value Then
      s_Sql = "SELECT DISTINCT psn.codpsn, CONCAT(IFNULL(psn.apepaterno,''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrespsn, "
      s_Sql = s_Sql & "dci.sigladci, psn.numdociden, psn.fecingreso, psn.fecbaja, cgo.descgo, psn.cuentacts, bco.desbco, "
      s_Sql = s_Sql & "CASE WHEN psn.ctsdolar='" & s_Estado_Ina & "' THEN '" & s_Codmon_mn_Txt & "' ELSE '" & s_Codmon_me_Txt & "' END AS moncuentacts, "
      s_Sql = s_Sql & "CONCAT(IFNULL(zon.abrezona, ''), ' ', IFNULL(cfg.direccionzona, '')) AS empzona, "
      s_Sql = s_Sql & "CONCAT(IFNULL(via.abrevia, ''), ' ', IFNULL(cfg.direccionvia, ''), ' ', IFNULL(cfg.numerodir, '')) AS empvia, "
      s_Sql = s_Sql & "cfg.girocomercial, dcr.sigladci AS repsigladci, cfg.repnumdocu, cgg.descgo AS descgoger, "
      s_Sql = s_Sql & "CONCAT(IFNULL(cfg.gernombres, '') , ' ', IFNULL(cfg.gerapepaterno, ''), ' ', IFNULL(cfg.gerapematerno, '')) AS gernombres, "
      s_Sql = s_Sql & "dcg.sigladci AS gersigladci, cfg.gernumdocu, cfg.firma, cfg.firmanexo, " & IIf(s_PrnLogo = s_Estado_Act, "cfg.logo", "Null") & " AS logo "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plresultado res ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
      s_Sql = s_Sql & "INNER JOIN pldocidentidad dci ON dci.coddci=psn.coddci "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON cgo.codcls=psn.codcls AND cgo.codcgo=psn.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plbanco bco ON bco.codbco=psn.codbcocts "
      s_Sql = s_Sql & "LEFT JOIN plcfgempresa cfg ON cfg.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON zon.codzona=cfg.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON via.codvia=cfg.codvia "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dcr ON dcr.coddci=cfg.repcoddci "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dcg ON dcg.coddci=cfg.gercoddci "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgg ON cgg.codcls=psn.codcls AND cgg.codcgo=cfg.gercargo "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND res.codpdo='" & txtPeriodo.Text & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion rng "
      s_Sql = s_Sql & "WHERE rng.proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND rng.usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND rng.fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "AND psn.estadopsn='I' "
      s_Sql = s_Sql & "ORDER BY " & aElemento(0, 1)
    End If
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
    If s_OptRegistro = "repliquida" Then
      s_Sql = gdl_Funcion.HelpTablas("pet", "codpdo", s_Estado_Ina & "L" & ps_ClsPlanilla & ps_Anyo, "")
    Else
      s_Sql = gdl_Funcion.HelpTablas("ped", "codpdo", s_Estado_Blq & ps_ClsPlanilla & ps_Anyo, "")
    End If
   Case 1
    gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos", 2
   tdbHelp.Columns(0).DataField = "codcpc": tdbHelp.Columns(1).DataField = "descpc"
   s_Sql = gdl_Funcion.HelpTablas("cxt", "codcpc", ps_ClsPlanilla & "F" & "0", "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = panToolBar(1).Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = panToolBar(1).Left + ((cmdHelp(Index).Left / 1.5) + (cmdHelp(Index).Width / 2))
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
  Me.Height = 6750: Me.Width = 8760
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
  
  ReDim aElemento(6, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Apellido Paterno", "Apellido Materno", "Nombre(s)", "Fec. Cese", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpsn", "apepaterno", "apematerno", "nombres", "fecbaja", "estadopsn")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1080, 1616.33, 1616.33, 1616.33, 950, 300)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", "", "", s_FormatoFecha, "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0, 0, 0)
  Next n_Index
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(5).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(5).ValueItems.Translate = True
  For n_Index = 0 To 5
    tdbRegistro.Columns(5).ValueItems.Add Item
    tdbRegistro.Columns(5).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "A", "V", "L", "P", "O", "I")
    tdbRegistro.Columns(5).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "estadok", "estadovo", "estadnok", "estadopk", "estadopn", "procenok") & ".bmp")
  Next n_Index
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  tdbRegistro.AllowColMove = False
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(9, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = IIf(s_OptRegistro = "anrenta5ta" Or s_OptRegistro = "repliquida", "reporte", "registro"): aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, IIf(s_OptRegistro = "anrenta5ta" Or s_OptRegistro = "repliquida", "promedio", "seleccio"), "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, IIf(s_OptRegistro = "anrenta5ta" Or s_OptRegistro = "repliquida", "Configuración de Parametros", "Selecciona y Edita Registro"), "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codbco", "desbco")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 864.7402, 3315.071)
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
  
  For n_Index = 0 To 3
    ribAnalisis(n_Index).PictureUp = LoadPicture()
    ribAnalisis(n_Index).ToolTipText = Choose(n_Index + 1, "Boleta Liquidación", "Certificado de Trabajo", "Carta Retiro CTS", "Constancia de Trabajo")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "repogene", "certifica", "liquicts", "constanci") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribAnalisis(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribAnalisis(0).Value = True
  
  ribFirma.PictureUp = LoadPicture()
  ribFirma.ToolTipText = "Representante Adjunto"
  s_Sql = gdl_Procedure.ps_PathImagen & "dividir.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribFirma.PictureUp = LoadPicture(s_Sql)
  ribFirma.Value = False
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  ribParametro(0).Value = True
  gdl_Procedure.EditDTPicker "PK", dtpFecha, Date, s_MdoData_Ins, True, s_FormatoFecha, dtpShortDate
  
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

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then Beep: MsgBox "No existen Registros para Seleccionar", vbExclamation: Exit Sub
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
    s_Sql = gdl_Funcion.HelpTablas("pet", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & "L" & ps_ClsPlanilla & ps_Anyo, "")
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
