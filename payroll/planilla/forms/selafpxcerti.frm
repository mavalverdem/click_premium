VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelEntiAfpCertifik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "selafpxcerti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   6120
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4725
      Left            =   45
      TabIndex        =   11
      Top             =   570
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   8334
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
      Left            =   0
      Top             =   5400
      Width           =   5250
      _ExtentX        =   9260
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
      Left            =   5340
      TabIndex        =   0
      Top             =   570
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
         Picture         =   "selafpxcerti.frx":000C
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
         Picture         =   "selafpxcerti.frx":0028
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
         Picture         =   "selafpxcerti.frx":0044
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
         Picture         =   "selafpxcerti.frx":0060
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
         Picture         =   "selafpxcerti.frx":007C
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
         Picture         =   "selafpxcerti.frx":0098
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
         Picture         =   "selafpxcerti.frx":00B4
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
         Picture         =   "selafpxcerti.frx":00D0
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
         Picture         =   "selafpxcerti.frx":00EC
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6120
      _Version        =   65536
      _ExtentX        =   10795
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
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   1
         Left            =   795
         TabIndex        =   14
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
         PictureUp       =   "selafpxcerti.frx":0108
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   0
         Left            =   390
         TabIndex        =   13
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
         PictureUp       =   "selafpxcerti.frx":0124
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   2655
         TabIndex        =   19
         Top             =   90
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   137822209
         CurrentDate     =   37515
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   4965
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
         PictureUp       =   "selafpxcerti.frx":0140
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   4530
         TabIndex        =   15
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
         PictureUp       =   "selafpxcerti.frx":015C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   5370
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
         PictureUp       =   "selafpxcerti.frx":0178
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha  :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   18
         Top             =   120
         Width           =   900
      End
   End
End
Attribute VB_Name = "fSelEntiAfpCertifik"
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
Private Sub CertificadoPension(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  
  Dim s_Moneda As String, s_OldMessage As String
  Dim nRentaBruta As Double, nImpuestoRenta As Double
  Dim nRegistro As Long, nRegistros As Long
  Dim sConceptoRemun As String, sConceptoReten As String
  Dim sPersonal As String, sDocIdentidad As String, sNumeroAfp As String
  
  ' Cambio el Mensaje y Muestro la Barra
  s_OldMessage = fMenu.panMessage.Caption
  MuestraMensaje "Generando " & IIf(ribAnalisis(0).Value, "Certificado", "Resumen") & " ..."
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  s_Moneda = IIf(fMenu.ribMoneda(0).Value, "mn", "me")
  sConceptoRemun = "cpcremase"
  sConceptoReten = "cpcapobli"
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo mes
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  's_Sql = "CREATE TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.codpsn, CONCAT(IFNULL(psn.apepaterno, ''), ' ', IFNULL(psn.apematerno, ''), ',  ', IFNULL(psn.nombres, '')) AS nombrespsn, "
  'MODIFICACION 08/01/2008
  's_Sql = s_Sql & "psn.numeroafp, psn.numdociden, dxr.fecingreso, asi.fechacese, dxr.naciextrapsn, res.codpdo, dxr.codafp "
  s_Sql = s_Sql & "psn.numeroafp, psn.numdociden, dxr.fecingreso, psn.fecbaja as fechacese, dxr.naciextrapsn, res.codpdo, dxr.codafp "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg." & sConceptoReten & " "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.codpdo>'" & s_PeriodoRemAper & "' "
  s_Sql = s_Sql & "AND dxr.codafp IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  If Not ribParametro(0).Value Then
    s_Sql = s_Sql & "AND psn.estadopsn" & IIf(ribParametro(1).Value, "<>", "=") & "'I' "
  End If
  s_Sql = s_Sql & "ORDER BY codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal del certificado
  s_Sql = "DROP TABLE IF EXISTS tmpimporte"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimporte ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, "
  s_Sql = s_Sql & "codcpc varchar(4) NOT Null, "
  If ribAnalisis(1).Value Then
    s_Sql = s_Sql & "pdomes char(2) NOT Null, "
  End If
  s_Sql = s_Sql & "rembruta decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "impreten decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "fecingreso date Null, "
  s_Sql = s_Sql & "fecbaja date Null)"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Genero tabla de remuneraciones asegurables
  s_Sql = "DROP TABLE IF EXISTS tmpasegurable"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpasegurable "
  s_Sql = s_Sql & "SELECT res.codcls, res.codpsn, " & IIf(ribAnalisis(1).Value, "res.pdomes, ", "")
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & s_Moneda & ", 0)), 2) AS remasegurable "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.codpdo=psn.codpdo "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg." & sConceptoRemun & " "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "GROUP BY res.codpsn" & IIf(ribAnalisis(1).Value, ", res.pdomes ", " ")
  s_Sql = s_Sql & "ORDER BY res.codpsn" & IIf(ribAnalisis(1).Value, ", res.pdomes", "")
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto el importe de aporte de pensiones
  For n_Index = 1 To 3
    sConceptoReten = Choose(n_Index, "cpcapobli", "cpcseguro", "cpcporcen")
    s_Sql = "INSERT INTO tmpimporte "
    s_Sql = s_Sql & "SELECT res.codpsn, res.codcpc, " & IIf(ribAnalisis(1).Value, "res.pdomes, ", "")
    s_Sql = s_Sql & "ras.remasegurable AS rembruta, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_" & s_Moneda & ", 0)), 2) AS impreten, "
    s_Sql = s_Sql & "MAX(psn.fecingreso) AS fecingreso, MAX(psn.fechacese) AS fecbaja  "
    s_Sql = s_Sql & "FROM plresultado res "
    s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.codpdo=psn.codpdo "
    s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg." & sConceptoReten & " "
    s_Sql = s_Sql & "INNER JOIN tmpasegurable ras ON res.codcls=ras.codcls AND res.codpsn=ras.codpsn " & IIf(ribAnalisis(1).Value, "AND res.pdomes=ras.pdomes ", "")
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "GROUP BY res.codpsn" & IIf(ribAnalisis(1).Value, ", res.pdomes ", " ")
    s_Sql = s_Sql & "ORDER BY res.codpsn" & IIf(ribAnalisis(1).Value, ", res.pdomes", "")
    If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  Next n_Index
  
  ' Genero tabla de remuneraciones aportes de pensiones
  s_Sql = "DROP TABLE IF EXISTS tmppensiones"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmppensiones "
  's_Sql = "CREATE TABLE tmppensiones "
  s_Sql = s_Sql & "SELECT tmp.codpsn, tmp.codcpc, " & IIf(ribAnalisis(1).Value, "tmp.pdomes, ", "")
  s_Sql = s_Sql & "tmp.fecingreso, tmp.fecbaja, "
  If ribAnalisis(0).Value Then
    s_Sql = s_Sql & "tmp.rembruta, tmp.impreten "
  Else
    s_Sql = s_Sql & "tmp.rembruta, ROUND(SUM(tmp.impreten), 2) AS impreten "
  End If
  s_Sql = s_Sql & "FROM tmpimporte tmp "
  If ribAnalisis(1).Value Then
    s_Sql = s_Sql & "GROUP BY tmp.codpsn, tmp.pdomes "
  End If
  s_Sql = s_Sql & "ORDER BY tmp.codpsn" & IIf(ribAnalisis(1).Value, ", tmp.pdomes", "")
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Recupero la informacion del certificado
  s_Sql = "SELECT DISTINCTROW tmp.codpsn, psn.nombrespsn, psn.numdociden, tmp.fecingreso, "
  s_Sql = s_Sql & "tmp.fecbaja, psn.naciextrapsn, psn.numeroafp, tmp.codcpc, cpc.descpc, "
  s_Sql = s_Sql & "tmp.rembruta, tmp.impreten" & IIf(ribAnalisis(1).Value, ", tmp.pdomes ", " ")
  s_Sql = s_Sql & "FROM tmppensiones tmp "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON tmp.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON tmp.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "WHERE tmp.impreten<>0.00 "
  s_Sql = s_Sql & "ORDER BY codpsn" & IIf(ribAnalisis(1).Value, ", pdomes", "")
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
   If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt)
    ' Arreglos de grabación
    If ribAnalisis(0).Value Then
      a_Campos = Array("codpsn", "nombrespsn", "numdociden", "numeroafp", "fecingreso", "fecbaja", "codcpc", "descpc", "moneda", "remunbruta", "retencion")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero)
    Else
      a_Campos = Array("codpsn", "nombrespsn", "numdociden", "numeroafp", "fecingreso", "fecbaja", "codcpc", "descpc", "moneda", "remunbruta01", "retencion01", "remunbruta02", "retencion02", "remunbruta03", "retencion03", "remunbruta04", "retencion04", "remunbruta05", "retencion05", "remunbruta06", "retencion06", "remunbruta07", "retencion07", "remunbruta08", "retencion08", "remunbruta09", "retencion09", "remunbruta10", "retencion10", "remunbruta11", "retencion11", "remunbruta12", "retencion12")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero, TipoDato.Numero)
    End If
    While Not porstRecordset.EOF
      ' Datos generales
      sPersonal = porstRecordset!codpsn
      sDocIdentidad = IIf(IsNull(porstRecordset!numdociden), "", porstRecordset!numdociden)
      sNumeroAfp = IIf(IsNull(porstRecordset!numeroafp), "", porstRecordset!numeroafp)
      nRentaBruta = CDec(porstRecordset!rembruta)
      nImpuestoRenta = CDec(porstRecordset!impreten)
      ' Valores de grabación
      
      If ribAnalisis(0).Value Then
        a_Valores = Array(sPersonal, UCase(porstRecordset!nombrespsn), sDocIdentidad, sNumeroAfp, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), Trim(porstRecordset!codcpc), Trim(porstRecordset!descpc), s_Moneda, nRentaBruta, nImpuestoRenta)
      Else
        a_Valores = Array(sPersonal, UCase(porstRecordset!nombrespsn), sDocIdentidad, sNumeroAfp, Format(porstRecordset!fecingreso, s_FmtFechMysql_0), Format(porstRecordset!fecbaja, s_FmtFechMysql_0), Trim(porstRecordset!codcpc), Trim(porstRecordset!descpc), s_Moneda, CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0), CDec(0))
      End If
      ' Analisis de información
      If ribAnalisis(0).Value Then
        ' Incremento el porcentaje
        nRegistro = nRegistro + 1
        fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
        DoEvents
        porstRecordset.MoveNext
      Else
        Do
          If CDec(porstRecordset!impreten) > 0 Then
            nRentaBruta = CDec(porstRecordset!rembruta)
            nImpuestoRenta = CDec(porstRecordset!impreten)
            n_Index = (CInt(porstRecordset!pdomes) * 2)
            a_Valores(n_Index + 7) = nRentaBruta
            a_Valores(n_Index + 8) = nImpuestoRenta
          End If
          ' Incremento el porcentaje
          nRegistro = nRegistro + 1
          fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
          DoEvents
          porstRecordset.MoveNext
          ' Fin de archivo
          If porstRecordset.EOF Then Exit Do
        Loop While sPersonal = porstRecordset!codpsn
      End If
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      ' Realizo la actualización de los registros
      If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    Wend
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False
  MuestraMensaje s_OldMessage
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codafp, desafp, factor1, factor2, factor3, factor4,"
  s_Sql = s_Sql & " codbco, ctacteafp, desctacteafp, ctactefondo, desctactefondo, estadoafp"
  s_Sql = s_Sql & " FROM plentidadafp"
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
  Dim s_FechaHora As String, sFechaPrn As String
  Dim sDireccion As String, sDistrito As String
  Dim sRepresentante As String, sCargoRepresenta As String
  Dim sNumDocumento As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Actualización de parametros
    Me.Tag = s_MdoData_Vis
    fPrmCertifikAfp.Show vbModal
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
        
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Obtengo los datos de la empresa
    sDireccion = "": sRepresentante = ""
    s_Sql = "SELECT codvia, direccionvia, numerodir, codzona, direccionzona, ubigeodir, "
    s_Sql = s_Sql & "CONCAT(repapepaterno, ' ', repapematerno, ', ', repnombres) AS representante, repnumdocu, "
    s_Sql = s_Sql & "IFNULL(dci.sigladci, '') AS sigladci, repcargo "
    s_Sql = s_Sql & "FROM plcfgempresa "
    s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON plcfgempresa.repcoddci=dci.coddci "
    s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sRepresentante = gdl_Funcion.aTexto(porstRecordset!representante)
      sCargoRepresenta = gdl_Funcion.aTexto(porstRecordset!repcargo)
      sCargoRepresenta = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, sCargoRepresenta, "DC")
      sNumDocumento = porstRecordset!sigladci & " " & porstRecordset!repnumdocu
      sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      sDistrito = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " - " & sDistrito
    End If
    porstRecordset.Close
    
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = IIf(ribAnalisis(0).Value, "CERTIFICADO DEL SISTEMA PRIVADO DE PENSIONES AFP", "RESUMEN ANUAL DEL SISTEMA PRIVADO DE PENSIONES AFP")
    gdl_Procedure.ps_ReportName = IIf(ribAnalisis(0).Value, "rptcertisnp", "rptresumafp")
    ReDim aElemento(3, 9): ReDim aElementos(2)
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
    aElemento(2, 4) = "Ejercicio;" & ps_Anyo & ";true"
    aElemento(2, 5) = ""
    aElemento(2, 6) = ""
    aElemento(2, 7) = "DocRepresentante;" & sNumDocumento & ";true"
    aElemento(2, 8) = "CargoRepresentante;" & sCargoRepresenta & ";true"
    
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    ' Genera la información del reporte
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If s_OptRegistro = "certifiafp" Then
      s_Sql = s_Sql & "codpsn varchar(11) Not Null, nombrespsn varchar(80) Null, numdociden varchar(11) Null, "
      s_Sql = s_Sql & "numeroafp varchar(15) Null, fecingreso date Null, fecbaja date Null, "
      s_Sql = s_Sql & "codcpc varchar(4) Not Null, descpc varchar(50) Null, moneda char(3) Null, "
      If ribAnalisis(0).Value Then
        s_Sql = s_Sql & "remunbruta decimal(18,2) Null Default '0', retencion decimal(18,2) Null Default '0', "
        aElemento(2, 5) = "EntidadPension;" & "SISTEMA PRIVADO DE PENSIONES - AFP " & Trim(dcaRegistro.Recordset!desafp) & ";true"
        aElemento(2, 6) = "FechaPrn;" & UCase(sDistrito & ", " & sFechaPrn) & ";true"
      Else
        For n_Index = 1 To 12
          s_Sql = s_Sql & "remunbruta" & Format(n_Index, "00") & " decimal(18,2) Null Default '0', "
          s_Sql = s_Sql & "retencion" & Format(n_Index, "00") & " decimal(18,2) Null Default '0', "
        Next n_Index
        aElemento(2, 5) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & " - " & Trim(dcaRegistro.Recordset!desafp) & " (" & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt) & ")" & ";true"
        aElemento(2, 6) = ""
      End If
      s_Sql = s_Sql & "PRIMARY KEY (codpsn, codcpc)) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      CertificadoPension "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
    End If
    ' Obtengo la información del reporte
    s_Sql = "SELECT rpt.*, cfg.logo, cfg.firma "
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
  End Select

End Sub
Private Sub Form_Load()

  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6330: Me.Width = 6210
  Me.Left = 700: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Entidad Pensión"
  
  ReDim aElemento(3, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codafp", "desafp", "estadoafp")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 800, 3556.03, 300)
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
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(2).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(2).ValueItems.Translate = True
  For n_Index = 0 To 1
    tdbRegistro.Columns(2).ValueItems.Add Item
    tdbRegistro.Columns(2).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, s_Estado_Act, s_Estado_Ina)
    tdbRegistro.Columns(2).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "estadok", "estadnok") & ".bmp")
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
  
  ' Cargo los graficos de los botones de parametro y analisis
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
  ribParametro(0).Value = True

  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
  ribAnalisis(0).Value = True
  ' Configuro los parametos adicionales
  gdl_Procedure.EditDTPicker "PK", dtpFecha, Date, s_MdoData_Ins, True, s_FormatoFecha, dtpShortDate

  ' Bloqueo la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = False
  
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
Private Sub tdbRegistro_DblClick()
  cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
