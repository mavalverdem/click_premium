VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fReporContabiliza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 01"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "rptcontabil.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   11115
   Begin MSAdodcLib.Adodc dcaSeleccion 
      Height          =   330
      Index           =   3
      Left            =   45
      Top             =   5460
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
      TabIndex        =   1
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
         Left            =   150
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1110
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
         Picture         =   "rptcontabil.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   4
         Tag             =   "0"
         Top             =   1530
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
         Picture         =   "rptcontabil.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   2295
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
         Picture         =   "rptcontabil.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   2730
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
         Picture         =   "rptcontabil.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   3885
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
         Picture         =   "rptcontabil.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   4320
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
         Picture         =   "rptcontabil.frx":0098
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
         Left            =   150
         TabIndex        =   2
         Tag             =   "0"
         Top             =   675
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
         Picture         =   "rptcontabil.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   3150
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
         Picture         =   "rptcontabil.frx":00D0
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   10
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
      Begin VB.ComboBox cmbProceso 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "rptcontabil.frx":00EC
         Left            =   4050
         List            =   "rptcontabil.frx":00EE
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   105
         Width           =   4110
      End
      Begin Threed.SSRibbon ribSeccion 
         Height          =   360
         Left            =   10005
         TabIndex        =   14
         Top             =   75
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
         PictureUp       =   "rptcontabil.frx":00F0
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   1110
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
         PictureUp       =   "rptcontabil.frx":010C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   705
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
         PictureUp       =   "rptcontabil.frx":0128
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proceso :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   2955
         TabIndex        =   12
         Top             =   150
         Width           =   1005
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   0
         Left            =   2625
         Shape           =   4  'Rounded Rectangle
         Top             =   45
         Width           =   6255
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   5235
      Left            =   5460
      TabIndex        =   17
      Top             =   555
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
      TabPicture(0)   =   "rptcontabil.frx":0144
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dcaSeleccion(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdbSeleccion(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ubicación"
      TabPicture(1)   =   "rptcontabil.frx":0160
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dcaSeleccion(1)"
      Tab(1).Control(1)=   "tdbSeleccion(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Sección"
      TabPicture(2)   =   "rptcontabil.frx":017C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dcaSeleccion(2)"
      Tab(2).Control(1)=   "tdbSeleccion(2)"
      Tab(2).ControlCount=   2
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   4425
         Index           =   0
         Left            =   60
         TabIndex        =   18
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
      TabIndex        =   0
      Top             =   555
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
Attribute VB_Name = "fReporContabiliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(4, 2)                            ' Array de inicio y fin de seleccion de registro
Private s_OptRegistro As String                         ' Instancia del formulario activo
'[
Private Sub ConDetalle(nTabIndex As Integer, s_Tabla As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)
  Dim sMoneda As String, sArchivo As String
  Dim sPersonal As String, sNomPersonal As String
  Dim sCencosto As String, sImpuesto As String, sDetalle As String
  Dim nImporte As Double, nTipoCambio As Double
  Dim sCamRubro As String, s_OldMessage As String
  Dim nRegistro As Long, nRegistros As Long

  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Imprimiendo Contabilización ..."
  
  ' Agrupacion default
  sCamRubro = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codpdo")
  sMoneda = IIf(s_Moneda = s_Codmon_mn, "mn", "me")
  sArchivo = "condetalle"
  
  ' Genero la tabla temporal de contabilización
  s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS " & sArchivo & " ( "
  s_Sql = s_Sql & "codcta varchar(15) NOT Null, codpsn varchar(11) Null, "
  s_Sql = s_Sql & "nombrepsn varchar(75) Null, codcco varchar(9) Null, "
  s_Sql = s_Sql & "detalle varchar(60) Null, codmon char(1) Null, clavecon char(2) Null, "
  s_Sql = s_Sql & "tipocambio decimal(6,3) Null Default '0', fechaproceso date Null, "
  s_Sql = s_Sql & "debe decimal(18,2) Null Default '0', haber decimal(18,2) Null Default '0') "
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Primer Paso : Cuentas que no tiene (centro de costo, tercero)
  s_Sql = "INSERT INTO " & sArchivo & " "
  s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, Null AS codpsn, Null AS nombrepsn, Null AS codcco, cpc.descpc, res.codmon, '40' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS debe, 0.00 AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & " "
  s_Sql = s_Sql & "UNION "
  s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, Null AS codpsn, Null AS nombrepsn, Null AS codcco, cpc.descpc, res.codmon, '50' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "0.00 AS debe, ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Segundo Paso : Cuentas que tiene (centro de costo, tercero)
  s_Sql = "INSERT INTO " & sArchivo & " "
  s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, dxr.codcco, cpc.descpc, res.codmon, '11' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS debe, 0.00 AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn, dxr.codcco "
  s_Sql = s_Sql & "UNION "
  s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, dxr.codcco, cpc.descpc, res.codmon, '11' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "0.00 AS debe, ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn, dxr.codcco "
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Tercer Paso : Cuentas que tiene tercero y no centro de costo
  s_Sql = "INSERT INTO " & sArchivo & " "
  s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, Null AS codcco, cpc.descpc, res.codmon, '11' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS debe, 0.00 AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", res.codpsn "
  s_Sql = s_Sql & "UNION "
  s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ', ', IFNULL(psn.nombres, '')) AS nombrepsn, Null AS codcco, cpc.descpc, res.codmon, '11' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "0.00 AS debe, ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Act & "' AND cta.indcco='" & s_Estado_Ina & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", res.codpsn "
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Cuarto Paso : Cuentas que no tiene tercero y tiene centro de costo
  s_Sql = "INSERT INTO " & sArchivo & " "
  s_Sql = s_Sql & "SELECT res.codcta_deb" & sMoneda & " AS codcta, Null AS codpsn, Null AS nombrepsn, dxr.codcco, cpc.descpc, res.codmon, '40' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS debe, 0.00 AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_deb" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_deb" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_deb" & sMoneda & ", dxr.codcco "
  s_Sql = s_Sql & "UNION "
  s_Sql = s_Sql & "SELECT res.codcta_hab" & sMoneda & " AS codcta, Null AS codpsn, Null AS nombrepsn, dxr.codcco, cpc.descpc, res.codmon, '50' AS clavecon, "
  s_Sql = s_Sql & "pdo.tipocambio, pdo.fechaproceso, "
  s_Sql = s_Sql & "0.00 AS debe, ROUND(SUM(IFNULL(res.importe_" & sMoneda & ", 0)), 2) AS haber "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocta cta ON res.codcta_hab" & sMoneda & "=cta.codcta AND cta.tpocta='" & s_Estado_Act & "' AND cta.estcta='" & s_MdoData_Ins & "' "
  s_Sql = s_Sql & "AND cta.inddoc='" & s_Estado_Ina & "' AND cta.indcco='" & s_Estado_Act & "' "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codproce_pdo='" & Right(Trim(cmbProceso), 2) & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND IFNULL(res.codcta_hab" & sMoneda & ", '')<>'' "
  If nTabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & nTabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY res.codcta_hab" & sMoneda & ", dxr.codcco "
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Ultimo-1 Paso : Actualizo los codigos de deudor
  s_Sql = "UPDATE " & sArchivo & " det, plpersonal psn "
  s_Sql = s_Sql & "SET det.codcta=psn.coddeudor "
  s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND det.codpsn=psn.codpsn "
  s_Sql = s_Sql & "AND IFNULL(psn.coddeudor, '')<>'' "
  s_Sql = s_Sql & "AND IFNULL(det.debe, 0)<>0"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Finalizar
  ' Ultimo-2 Paso : Actualizo los codigos de acreedor
  s_Sql = "UPDATE " & sArchivo & " det, plpersonal psn "
  s_Sql = s_Sql & "SET det.codcta=psn.codacredor "
  s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND det.codpsn=psn.codpsn "
  s_Sql = s_Sql & "AND IFNULL(psn.codacredor, '')<>'' "
  s_Sql = s_Sql & "AND IFNULL(det.haber, 0)<>0"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Finalizar
  
  ' Registros de contabilización
  s_Sql = "SELECT codcta, codpsn, nombrepsn, codcco, detalle, codmon, "
  s_Sql = s_Sql & "clavecon, tipocambio, fechaproceso, debe, haber "
  s_Sql = s_Sql & "FROM " & sArchivo & " "
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    ' Muestro la Barra
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    sMoneda = IIf(sMoneda = "mn", "PEN", "USD")
    sDetalle = Trim(Left(Trim(cmbProceso), 50))
    
    ' Genero los arreglos de la grabación
    a_Campos = Array("codproce", "secuencia", "desproce", "sociedad", "fecdocum", "clasedoc", "fecconta", "moneda", "tipcambio", "referencia", "glosacpb", "clavecon", "cuenta", "indcme", "importe", "impuesto", "fecvalida", "fecvenci", "codcco", "asigna", "glosadet", "codpsn", "nombrepsn")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    While Not porstRecordset.EOF
      sPersonal = gdl_Funcion.aTexto(porstRecordset!codpsn)
      sNomPersonal = gdl_Funcion.aTexto(porstRecordset!nombrepsn)
      sCencosto = gdl_Funcion.aTexto(porstRecordset!codcco)
      sImpuesto = IIf(sCencosto <> "", "C0", "")
      sCencosto = IIf(sCencosto = "", "0", sCencosto)
      nImporte = Abs(CDec(porstRecordset!debe) + CDec(porstRecordset!haber))
      nTipoCambio = CDec(IIf(sMoneda = "PEN", 0, porstRecordset!Tipocambio))
      a_Valores = Array(Right(Trim(cmbProceso), 2), nRegistro, Left(Trim(cmbProceso), 50), gdl_Funcion.PadR(Val(ps_CodEmpresa), 4, "0"), Format(porstRecordset!fechaproceso, s_FmtFechMysql_0), "NO", Format(porstRecordset!fechaproceso, s_FmtFechMysql_0), sMoneda, nTipoCambio, Format(porstRecordset!fechaproceso, "yyyymmdd"), Left(sDetalle, 30), porstRecordset!clavecon, porstRecordset!codcta, "", nImporte, sImpuesto, "", Format(porstRecordset!fechaproceso, s_FmtFechMysql_0), sCencosto, Format(porstRecordset!fechaproceso, "ddmmyyyy"), UCase(Left(porstRecordset!detalle, 30)), sPersonal, sNomPersonal)
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      gdl_Conexion.ConfirmaTransaccion  ' Confirma transacción
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
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
Private Sub PreSintesis(nTabIndex As Integer, s_Tabla As String, s_Proceso As String, s_FechaHora As String, s_Moneda As String)
  Dim nContador As Integer, nColumna As Integer
  Dim sColumna As String, a_Detalle(), a_Sintesis(9)
  Dim nImporteIng As Double, nImporteDsc As Double, nImportePag As Double
  Dim sCamRubro As String, sRubro As String, sDesRubro As String
  Dim sDescripcion As String, nDias As Long
  Dim nRegistro As Long, nRegistros As Long, s_OldMessage As String
  
  ' Inicializo valores
  sCamRubro = Choose(nTabIndex + 1, "codcco", "codubica", "codsec", "codpdo")
  sDesRubro = Choose(nTabIndex + 1, "detcco", "desubica", "dessec", "despdo")
  
  ' Genero las cabecera de los conceptos
  s_Sql = "SELECT DISTINCTROW res.codcpc, cpc.descpc, res.tipocpc, cpc.aliascpc, "
  s_Sql = s_Sql & Choose(nTabIndex + 1, "dxr.codcco, cco.detcco, ", "dxr.codubica, ubi.desubica, ", "dxr.codsec, sec.dessec, ", "res.codpdo, pdo.despdo, ")
  s_Sql = s_Sql & "SUM(IFNULL(asi.diatrabajo, 0)) AS dias, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(IF(res.tipocpc='" & s_Estado_Ina & "', res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0), 0)), 2) AS imporingreso, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(IF(res.tipocpc='" & s_Estado_Act & "', res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0), 0)), 2) AS impordescto, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(IF(res.tipocpc='" & s_Estado_Blq & "', res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0), 0)), 2) AS imporaporte, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "me", "mn") & ", 0)), 2) AS importecmb "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON res.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON res.codcls=pdo.codcls AND res.codpdo=pdo.codpdo "
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
  s_Sql = s_Sql & "GROUP BY " & Choose(nTabIndex + 1, "dxr.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "res.codpdo, ") & "res.tipocpc, res.codcpc "
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
    s_Sql = s_Sql & "WHERE LENGTH(codcco)=" & pn_NivelCenCosto & " "
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
  Dim s_Periodo As String, s_Moneda As String, sOrden As String
  Dim s_TituloReporte As String, s_FechaHora As String
  Dim nTabIndex As Integer
  
  nTabIndex = IIf(ribSeccion.Value, tabRegister.Tab, 3)
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
   Case 6, 7  ' Opciones de impresión
    nTabIndex = IIf(ribSeccion.Value, tabRegister.Tab, 3)
    ' Verifico que existan registros seleccionados
    If cmbProceso = "" Then Beep: MsgBox "Debe Seleccionar Proceso Calculo", vbInformation: cmbProceso.SetFocus: Exit Sub
    If tdbSeleccion(3).SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(3).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 0 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 1 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 2 Then Beep: MsgBox "Debe Seleccionar Rango " & tdbSeleccion(nTabIndex).Caption & " de Impresión", vbExclamation: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
    s_Periodo = ""
    nTabIndex = 3
    ' Barro el arreglo de registros (periodos) marcados (bookmarks)
    For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
      tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      s_Periodo = s_Periodo & " - " & Trim(tdbSeleccion(nTabIndex).Columns(1).Text)
    Next n_Index
    nTabIndex = IIf(ribSeccion.Value, tabRegister.Tab, nTabIndex)
    If nTabIndex <> 3 Then
      ' Barro el arreglo de registros marcadas (bookmarks)
      For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
        tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
        gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, Left(s_OptRegistro, 9) & nTabIndex, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      Next n_Index
    End If
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "REPORTE DE ANALISIS " & IIf(ribParametro(0).Value, "DETALLE", "RESUMEN")
    gdl_Procedure.ps_ReportName = IIf(ribParametro(0).Value, "rptcontadeta", "rptcontaresu")
    s_TituloReporte = "ANALISIS CONTABILIZA - " & UCase(tdbSeleccion(nTabIndex).Caption)
    s_TituloReporte = s_TituloReporte & " (" & IIf(s_Moneda = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt) & ")"
    
    ReDim aElemento(3, 4): ReDim aElementos(2)
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
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If ribParametro(0).Value Then
      s_Sql = s_Sql & "codproce char(2) Not Null, "
      s_Sql = s_Sql & "secuencia smallint(5) Not Null, "
      s_Sql = s_Sql & "desproce varchar(50) Null, "
      s_Sql = s_Sql & "sociedad char(4) Null, "
      s_Sql = s_Sql & "fecdocum date default Null, "
      s_Sql = s_Sql & "clasedoc char(2) Null, "
      s_Sql = s_Sql & "fecconta date default Null, "
      s_Sql = s_Sql & "moneda char(3) Null, "
      s_Sql = s_Sql & "tipcambio decimal(9,3) Null Default '0', "
      s_Sql = s_Sql & "referencia varchar(8) Null, "
      s_Sql = s_Sql & "glosacpb varchar(30) Null, "
      s_Sql = s_Sql & "clavecon char(2) Null, "
      s_Sql = s_Sql & "cuenta varchar(15) Null, "
      s_Sql = s_Sql & "indcme char(3) Null, "
      s_Sql = s_Sql & "importe decimal(18,2) Null Default '0', "
      s_Sql = s_Sql & "impuesto char(2) Null, "
      s_Sql = s_Sql & "fecvalida date default Null, "
      s_Sql = s_Sql & "fecvenci date default Null, "
      s_Sql = s_Sql & "codcco varchar(9) NOT Null, "
      s_Sql = s_Sql & "asigna varchar(10) NOT Null, "
      s_Sql = s_Sql & "glosadet varchar(30) Null, "
      s_Sql = s_Sql & "codpsn varchar(11) Null, "
      s_Sql = s_Sql & "nombrepsn varchar(75) Null, "
      s_Sql = s_Sql & "PRIMARY KEY (codproce, secuencia)) "
      sOrden = "codproce, secuencia"
    ElseIf ribParametro(1).Value Then
      s_Sql = s_Sql & "codrubro varchar(8) Not Null, "
      s_Sql = s_Sql & "secuencia smallint(5) Not Null, "
      s_Sql = s_Sql & "desrubro varchar(50) Null, "
      s_Sql = s_Sql & "codsec char(2) Null, "
      s_Sql = s_Sql & "dias decimal(18,2) Null Default '0', "
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
      ConDetalle nTabIndex, "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora, s_Moneda
    Else
      PreSintesis nTabIndex, "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora, s_Moneda
    End If
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY " & sOrden
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
    ' ]
   Case 8       ' Genera archivo de texto
  End Select

End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()
  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6315: Me.Width = 11200
  Me.Left = 400: Me.Top = 350
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Periodos de Pago"
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Inicio", "Final", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpdo", "despdo", "fechaini", "fechafin", "estadopdo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 750, 1800, 950, 950, 300)
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
  ReDim aElemento(8, 3)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&a", "&d", "&b", "&p", "&f", "&r", "&v", "&i")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  ' Cargo el grafico del boton de seccion
  ribSeccion.PictureUp = LoadPicture()
  ribSeccion.ToolTipText = "Parámetro de Clasificación "
  s_Sql = gdl_Procedure.ps_PathImagen & "dividir.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribSeccion.PictureUp = LoadPicture(s_Sql)
  ribSeccion.Value = False
  
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
  For n_Index = 0 To 1
    ' Tipo de analisis
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = Choose(n_Index + 1, "Detallado", "Resumen")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "analmovs", "resumen") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribParametro(0).Value = True
  
  ' Carga los datos en el formulario
  cmbProceso.Clear
  cmbProceso.Locked = False
  If s_OptRegistro = "rpcontapla" Then
    s_Sql = "SELECT codproce, desproce "
    s_Sql = s_Sql & "FROM plproceso "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND estadoproce<>'" & s_EstadoRemAper & "' "
    s_Sql = s_Sql & "ORDER BY codproce"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.EOF And porstRecordset.BOF) Then
      Do While Not porstRecordset.EOF
        cmbProceso.AddItem porstRecordset("desproce") & Space(60) & "|" & Trim(porstRecordset("codproce"))
        porstRecordset.MoveNext
      Loop
      porstRecordset.Close
      cmbProceso.ListIndex = 0
    End If
    cmbProceso.Width = 4110
  ElseIf s_OptRegistro = "rpcontapvs" Then
    For n_Index = 1 To 3
      cmbProceso.AddItem "Provisión de " & Choose(n_Index, "Vacaciones", "Gratificaciones", "C.T.S.")
    Next n_Index
    cmbProceso.ListIndex = 0
    cmbProceso.Width = 3000
  End If
  
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
