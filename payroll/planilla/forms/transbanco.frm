VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fTransferBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "transbanco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   6135
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4605
      Left            =   45
      TabIndex        =   16
      Top             =   990
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   8123
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
      Top             =   5625
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
      Height          =   4980
      Index           =   0
      Left            =   5355
      TabIndex        =   7
      Top             =   990
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   8784
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
         Index           =   3
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   2115
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
         Picture         =   "transbanco.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   2535
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
         Picture         =   "transbanco.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   12
         Tag             =   "0"
         Top             =   3210
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
         Picture         =   "transbanco.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   13
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
         Picture         =   "transbanco.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   14
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
         Picture         =   "transbanco.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   1680
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
         Picture         =   "transbanco.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   570
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
         Picture         =   "transbanco.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   22
         Tag             =   "0"
         Top             =   1005
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
         Picture         =   "transbanco.frx":00D0
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   1080
      TabIndex        =   17
      Top             =   915
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
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   930
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
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
         ForeColor       =   &H00000080&
         Height          =   280
         Left            =   2160
         TabIndex        =   2
         Top             =   180
         Width           =   945
      End
      Begin VB.TextBox txtBanco 
         ForeColor       =   &H00000080&
         Height          =   280
         Left            =   2160
         TabIndex        =   5
         Top             =   510
         Width           =   645
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   1
         Left            =   570
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
         PictureUp       =   "transbanco.frx":00EC
      End
      Begin Threed.SSRibbon ribAnalisis 
         Height          =   360
         Index           =   0
         Left            =   165
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
         PictureUp       =   "transbanco.frx":0108
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   510
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   3165
         TabIndex        =   20
         Top             =   180
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   180
         Width           =   885
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   540
         Width           =   885
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
         Left            =   3570
         TabIndex        =   3
         Top             =   210
         Width           =   195
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
         Left            =   3285
         TabIndex        =   6
         Top             =   540
         Width           =   195
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   0
         Left            =   1110
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   4935
      End
   End
End
Attribute VB_Name = "fTransferBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private s_Direccion As String                           ' Dirección de empresa
'[
Private Sub EliminaCarta(ByVal sNroCarta As String)
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Elimino los datos de carta existente
  s_Sql = "DELETE FROM plcartabanco "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpdo='" & txtPeriodo.Text & "' "
  s_Sql = s_Sql & "AND codbco='" & txtBanco.Text & "' "
  s_Sql = s_Sql & "AND nrocarta='" & sNroCarta & "'"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Error
  
  If fTransferBancos.ribAnalisis(1).Value Then
    ' Actualizo los datos de movimientos de C.T.S.
    s_Sql = "UPDATE plctsmovimiento mov, plpersonal psn "
    s_Sql = s_Sql & "SET mov.porinteres=0.00, mov.nrodeposito=NULL "
    s_Sql = s_Sql & "WHERE mov.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND mov.pdocts='" & txtPeriodo.Text & "' "
    s_Sql = s_Sql & "AND mov.nrodeposito='" & sNroCarta & "' "
    s_Sql = s_Sql & "AND mov.estadomov='" & s_Estado_Blq & "' "
    s_Sql = s_Sql & "AND psn.codcls=mov.codcls "
    s_Sql = s_Sql & "AND psn.codpsn=mov.codpsn "
    s_Sql = s_Sql & "AND psn.codbcocts='" & txtBanco.Text & "' "
    If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se Actualizo exitosamente la " & tdbRegistro.Caption, vbInformation
  ' Refresco el ado control y la grilla
  gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, tdbRegistro.Caption
  ' Ubico el registro ingresado o actualizado
  dcaRegistro.Recordset.Find ("nrocarta='" & sNroCarta & "'")
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub ExportaBancos(ByVal s_Archivo As String, ByVal s_NroCarta As String, s_Accion As String)
  Dim pofsoFileExp As FileSystemObject, potxtFileExp As TextStream
  Dim psRegistro As String, s_Caracter As String, s_Contenido As String
  Dim n_PosIni As Integer, n_PosFin As Integer, n_Longitud As Integer
  Dim n_Importe As Double, nImporte_mn As Double, nImporte_me As Double
  Dim nRegistro As Long, nRegistros As Long, nSecuencia As Long
  Dim s_OldMessage As String, sClaveLimite As String, sArchivo As String
  Dim nSumImporte As Double, nImporteLimite As Double
  Dim nArchivo As Integer
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  ' Sumatoria remuneración ganada
  If ribAnalisis(1).Value Then      ' cts
    ' Elimino y creo el archivo temporal de grabacion/restauración de información
    gdl_Conexion.Execucion "DROP TABLE IF EXISTS tmpdepositocts", Elimina
    s_Sql = "CREATE TEMPORARY TABLE tmpdepositocts "
    s_Sql = s_Sql & "SELECT DISTINCTROW mov.codcls, mov.pdocts, mov.codpsn, mov.nrodeposito, "
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.importe_mn, 0)), 2) AS remunera_mn, ROUND(SUM(IFNULL(res.importe_me, 0)), 2) AS remunera_me "
    s_Sql = s_Sql & "FROM plctsresultado res "
    s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON cfg.pdoano='" & ps_Anyo & "' AND cfg.remupercibects=res.codcpc "
    s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON mov.codcls=res.codcls AND mov.pdocts=res.pdocts AND mov.codpsn=res.codpsn "
    s_Sql = s_Sql & "INNER JOIN plcartabanco prm ON prm.codcls=mov.codcls AND prm.codpdo=mov.pdocts AND prm.codpsn=mov.codpsn AND prm.nrocarta=mov.nrodeposito "
    s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND res.pdocts='" & Trim(txtPeriodo.Text) & "' "
    s_Sql = s_Sql & "AND mov.nrodeposito='" & s_NroCarta & "' "
    s_Sql = s_Sql & "AND prm.codbco='" & Trim(txtBanco.Text) & "' "
    s_Sql = s_Sql & "GROUP BY mov.pdocts, mov.codpsn, mov.nrodeposito "
    s_Sql = s_Sql & "ORDER BY mov.codpsn"
    gdl_Conexion.Execucion s_Sql, Inserta
  End If
  
  'Recupero la información para exportar
  s_Sql = "SELECT DISTINCTROW prm.codcpc, cpc.descpc, prm.codpsn, psn.apepaterno, psn.apematerno, psn.nombres, psn.pagodolar, "
  s_Sql = s_Sql & "prm.desmotivo, prm.fechaproce, prm.codmon, psn.coddci, psn.numdociden, dci.sigladci, psn.fecnacimiento, "
  s_Sql = s_Sql & "CONCAT(prm.codpsn, prm.codcpc) AS cPrimaryKey, cfg.psnapepaterno, cfg.psnapematerno, cfg.psnnombres, "
  s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, zona.abrezona, psn.nomzondirec, psn.ubigeodir, "
  s_Sql = s_Sql & "psn.codbco" & IIf(ribAnalisis(0).Value, "pago", "cts") & " AS codbcopago, "
  s_Sql = s_Sql & "psn.cuenta" & IIf(ribAnalisis(0).Value, "pago", "cts") & " AS cuentapago, "
  s_Sql = s_Sql & "psn.interbank" & IIf(ribAnalisis(0).Value, "pago", "cts") & " AS interbankpago, "
  s_Sql = s_Sql & "bco.desbco, bco.codentidad, bco.cuenta" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & " AS cuentabco, "
  s_Sql = s_Sql & "IFNULL(bco.impolimite_" & IIf(fMenu.ribMoneda(0).Value, "mn", "me") & ", 0) AS impolimite, "
  s_Sql = s_Sql & "psn.tippago, bco.formato, bnk.codentidad AS codentidadbnk, "
  s_Sql = s_Sql & "cfg.codvia, cfg.direccionvia, cfg.numerodir, cfg.codzona, cfg.direccionzona, cfg.ubigeodir AS ubigeodir_emp, "
  s_Sql = s_Sql & "ROUND(SUM(IFNULL(prm.importe_mn, 0)), 2) AS importe_mn, ROUND(SUM(IFNULL(prm.importe_me, 0)), 2) AS importe_me, "
  If ribAnalisis(1).Value Then      ' cts
    s_Sql = s_Sql & "ROUND(SUM(IFNULL(res.remunera_mn, 0)), 2) AS remunera_mn, ROUND(SUM(IFNULL(res.remunera_me, 0)), 2) AS remunera_me "
  Else
    s_Sql = s_Sql & "0 AS remunera_mn, 0 AS remunera_me "
  End If
  s_Sql = s_Sql & "FROM plcartabanco prm "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON prm.codcls=psn.codcls AND prm.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plbanco bco ON prm.codbco=bco.codbco "
  s_Sql = s_Sql & "INNER JOIN plconcepto cpc ON prm.codcpc=cpc.codcpc "
  s_Sql = s_Sql & "INNER JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
  s_Sql = s_Sql & "LEFT JOIN plbanco bnk ON psn.codbnk" & IIf(ribAnalisis(0).Value, "pago", "cts") & "=bnk.codbco "
  s_Sql = s_Sql & "LEFT JOIN plcfgempresa cfg ON cfg.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
  s_Sql = s_Sql & "LEFT JOIN pltipozona zona ON psn.codzona=zona.codzona "
  If ribAnalisis(1).Value Then      ' cts
    s_Sql = s_Sql & "INNER JOIN tmpdepositocts res ON res.codcls=prm.codcls AND res.pdocts=prm.codpdo AND res.codpsn=prm.codpsn AND res.nrodeposito=prm.nrocarta "
  End If
  s_Sql = s_Sql & "WHERE prm.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND prm.codbco='" & Trim(txtBanco.Text) & "' "
  s_Sql = s_Sql & "AND prm.codpdo='" & Trim(txtPeriodo.Text) & "' "
  s_Sql = s_Sql & "AND prm.nrocarta='" & s_NroCarta & "' "
  'AGREGADO EL 20 DE JUNIO DEL 2008
  s_Sql = s_Sql & "AND psn." & IIf(ribAnalisis(0).Value, "pagodolar", "ctsdolar") & "='" & IIf(fMenu.ribMoneda(0).Value, 0, 1) & "' "
  s_Sql = s_Sql & "GROUP BY prm.codpsn, prm.codcpc "
  s_Sql = s_Sql & "ORDER BY prm.codpsn, prm.codcpc"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Archivo ..."
    fMenu.panPercent.Visible = True
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    ' reporte bancos
    If s_Accion = "R" Then
      ' Genero os arreglos de grabaciones
      a_Campos = Array("codpsn", "apepaterno", "apematerno", "nombres", "codcpc", "descpc", "desmotivo", "codmon", "codbco", "desbco", "cuentapago", "coddci", "numdociden", "codentidad", "cuentabco", "importemn", "importeme")
      a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero)
    Else
      ' Creo objeto de archivo
      Set pofsoFileExp = CreateObject("Scripting.FileSystemObject")
      nImporteLimite = CDec(porstRecordset!impolimite)
      sArchivo = s_Archivo: s_Caracter = " ": nArchivo = 0
    End If
    ' Secuencia de reporte o archivo
    While Not porstRecordset.EOF
      ' Genero el registro de grabación
      If s_Accion = "R" Then
        gdl_Conexion.IniciaTransaccion    ' Inicia transacción
        nImporte_mn = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, 0))
        nImporte_me = CDec(IIf(porstRecordset!codmon = s_Codmon_me, porstRecordset!importe_me, 0))
        a_Valores = Array(Trim(porstRecordset("codpsn")), UCase(gdl_Funcion.aTexto(porstRecordset("apepaterno"))), UCase(gdl_Funcion.aTexto(porstRecordset("apematerno"))), UCase(gdl_Funcion.aTexto(porstRecordset("nombres"))), gdl_Funcion.aTexto(porstRecordset("codcpc")), UCase(gdl_Funcion.aTexto(porstRecordset("descpc"))), gdl_Funcion.aTexto(porstRecordset("desmotivo")), gdl_Funcion.aTexto(porstRecordset("codmon")), gdl_Funcion.aTexto(porstRecordset("codbcopago")), UCase(gdl_Funcion.aTexto(porstRecordset("desbco"))), gdl_Funcion.aTexto(porstRecordset("cuentapago")), gdl_Funcion.aTexto(porstRecordset("coddci")), gdl_Funcion.aTexto(porstRecordset("numdociden")), gdl_Funcion.aTexto(porstRecordset("codentidad")), gdl_Funcion.aTexto(porstRecordset("cuentabco")), nImporte_mn, nImporte_me)
        ' Realizo la actualización de los registros
        If Not Records_Ins(s_Archivo, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
      Else
        ' inicializo variables
        nSecuencia = 0: nSumImporte = 0
        Do
          ' cabecera de archivo
          If nSecuencia = 0 Then
            ' Creo objeto de archivo
            s_Archivo = Replace(sArchivo, ".txt", "-" & nArchivo & ".txt")
            Set potxtFileExp = pofsoFileExp.CreateTextFile(s_Archivo, True)
            sClaveLimite = gdl_Funcion.aTexto(porstRecordset!cPrimaryKey)
            ' Registro incial del archivo
            If (porstRecordset!Formato = "1" Or porstRecordset!Formato = "5" Or porstRecordset!Formato = "6") Then    ' Formato credito, bbva, bcp
              psRegistro = InicioArchivo(sClaveLimite, porstRecordset!Formato, porstRecordset)
              potxtFileExp.WriteLine psRegistro
            ElseIf porstRecordset!Formato = "9" Then   ' Formato bbvacash
              For n_Index = 1 To 3
                psRegistro = InicioArchivo_BBVACash(sClaveLimite, porstRecordset!Formato, porstRecordset, n_Index)
                potxtFileExp.WriteLine psRegistro
              Next n_Index
            End If
            nArchivo = nArchivo + 1
          End If
          
          psRegistro = ""
          Select Case porstRecordset!Formato
           Case "1"                 ' Formato credito
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              If ribAnalisis(0).Value Then      ' Remuneraciones
                psRegistro = psRegistro & gdl_Funcion.PadR("", 1, s_Caracter)                   ' Espacios en blanco
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Estado_Blq, 1, s_Caracter)                  ' Tipo de registro constante
                psRegistro = psRegistro & Left(gdl_Funcion.aTexto(porstRecordset!cuentapago), 1)      ' Tipo de cuenta
                n_PosIni = 2
                s_Contenido = gdl_Funcion.aTexto(porstRecordset!cuentapago)
                n_PosFin = InStr(n_PosIni, s_Contenido, "-", vbBinaryCompare)
                n_Longitud = n_PosFin - n_PosIni
                s_Contenido = Mid(s_Contenido, n_PosIni, n_Longitud)
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 3, s_Caracter)  ' Cuenta abono - sucursal
                n_PosIni = n_PosFin + 1
                s_Contenido = gdl_Funcion.aTexto(porstRecordset!cuentapago)
                n_PosFin = InStr(n_PosIni, s_Contenido, "-", vbBinaryCompare)
                n_Longitud = n_PosFin - n_PosIni
                s_Contenido = Mid(s_Contenido, n_PosIni, n_Longitud)
                psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, 8, "0")         ' Cuenta abono - numero
                n_PosIni = n_PosFin + 1
                s_Contenido = gdl_Funcion.aTexto(porstRecordset!cuentapago)
                n_PosFin = InStr(n_PosIni, s_Contenido, "-", vbBinaryCompare)
                n_Longitud = n_PosFin - n_PosIni
                s_Contenido = Mid(s_Contenido, n_PosIni, n_Longitud)
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 1, s_Caracter)  ' Cuenta abono - moneda
                n_PosIni = n_PosFin + 1
                s_Contenido = gdl_Funcion.aTexto(porstRecordset!cuentapago)
                s_Contenido = Mid(s_Contenido, n_PosIni)
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 8, s_Caracter)  ' Cuenta abono - control y espacios
                s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
                psRegistro = psRegistro & gdl_Funcion.PadR(Left(s_Contenido, 40), 40, s_Caracter)
                psRegistro = psRegistro & Left(IIf(gdl_Funcion.aTexto(porstRecordset!codmon) = s_Codmon_mn, s_Codmon_mn_Txt, s_Codmon_me_Txt), 2)   ' Moneda de transferencia
                n_Importe = Format(CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me)), "###########0.00") * 100
                psRegistro = psRegistro & gdl_Funcion.PadL(n_Importe, 15, "0")      ' Importe de transferencia
                ' sumatoria abonos
                n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
                nSumImporte = nSumImporte + n_Importe
                psRegistro = psRegistro & gdl_Funcion.PadR(UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo)), 40, s_Caracter)    ' Motivo de transferencia
                psRegistro = psRegistro & gdl_Funcion.PadR("0", 1, s_Caracter)      ' Constante nota de abono
                psRegistro = psRegistro & gdl_Funcion.PadR(gdl_Funcion.aTexto(porstRecordset!sigladci), 3, s_Caracter)    ' Sigla de documento de identidad
                psRegistro = psRegistro & gdl_Funcion.PadR(gdl_Funcion.aTexto(porstRecordset!numdociden), 12, s_Caracter) ' Numero de documento de identidad
                psRegistro = psRegistro & gdl_Funcion.PadR("1", 1, s_Caracter)      ' Constante validar IDC vs cuenta
              Else
                ' 1: Tipo registro - constante
                s_Contenido = "3": n_Longitud = 1
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 2: Espacios - constante
                s_Contenido = "": n_Longitud = 14
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 3: Moneda cuenta abono
                s_Contenido = "M" & gdl_Funcion.aTexto(porstRecordset!codmon): n_Longitud = 2
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 4: Espacios - constante
                s_Contenido = "": n_Longitud = 6
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 5: Ruc de la empresa
                s_Contenido = ps_RucEmpresa: n_Longitud = 11
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 6: Cuenta deposito cts
                s_Contenido = Mid(Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1), IIf(porstRecordset!interbankpago = s_Estado_Act, 1, 2))
                n_Longitud = 11
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 7: Espacios - constante
                s_Contenido = "": n_Longitud = 6
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 8: Importe trasnferencia
                n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
                s_Contenido = Replace(Format(n_Importe, "############0.00"), ".", ""): n_Longitud = 15
                psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
                ' sumatoria abonos
                nSumImporte = nSumImporte + n_Importe
                ' 9: Digitos- constante
                s_Contenido = "00000": n_Longitud = 5
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                 ' 10: Espacios - constante
                s_Contenido = "": n_Longitud = 10
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 11: Tipo documento empresa - constante
                s_Contenido = "R": n_Longitud = 1
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 12: Digito - constante
                s_Contenido = "1": n_Longitud = 1
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 13: Espacios - constante
                s_Contenido = "": n_Longitud = 1
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 14: Nombre trabajador
                s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
                n_Longitud = 40
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 15: Dirección empresa
                s_Contenido = s_Direccion: n_Longitud = 40
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 16: Fecha nacimiento
                s_Contenido = Format(porstRecordset!fecnacimiento, "yyyymmdd"): n_Longitud = 8
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 17: Numero documento empleado
                s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 9
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 18: Tipo documento empleado
                n_Longitud = 1: s_Contenido = Right(gdl_Funcion.aTexto(porstRecordset!coddci), n_Longitud)
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 19: Codigo empresa - constante
                s_Contenido = "": n_Longitud = 6
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 20: Espacios - constante
                s_Contenido = "": n_Longitud = 12
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 21: Moneda ultimas 4 remuneraciones
                s_Contenido = "M" & IIf(gdl_Funcion.aTexto(porstRecordset!pagodolar) = s_Estado_Ina, s_Codmon_mn, s_Codmon_me)
                n_Longitud = 2
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                ' 22: Importe ultimas 4 remuneraciones
                n_Importe = CDec(IIf(porstRecordset!pagodolar = s_Estado_Ina, porstRecordset!remunera_mn, porstRecordset!remunera_me))
                s_Contenido = Replace(Format(n_Importe, "############0.00"), ".", ""): n_Longitud = 15
                psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              End If
              potxtFileExp.WriteLine psRegistro
            End If
           Case "2"                   ' Formato continental
            nSecuencia = nSecuencia + 1
            If ribAnalisis(0).Value Then      ' Remuneraciones
              psRegistro = psRegistro & gdl_Funcion.PadR(gdl_Funcion.aTexto(porstRecordset!numdociden), 10, s_Caracter)   ' Documento de identificación del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 12, s_Caracter)
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apematerno))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 12, s_Caracter)
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!nombres))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 11, s_Caracter)
              psRegistro = psRegistro & gdl_Funcion.PadR("A A A A ", 8, s_Caracter)                                               ' Codigo del servicio - vacio
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "")
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, 20, s_Caracter)                                     ' Cuenta de deposito
              n_Importe = Format(CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me)), "###########0.00") * 100
              psRegistro = psRegistro & gdl_Funcion.PadL(n_Importe, 14, "0")                                              ' Importe de la operación
              ' sumatoria abonos
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              nSumImporte = nSumImporte + n_Importe
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 30, s_Caracter)                                     ' Concepto de la operación
              psRegistro = psRegistro & gdl_Funcion.PadR("", 6, s_Caracter)                                               ' Espacios en blanco
            Else
              psRegistro = psRegistro & IIf(porstRecordset!coddci = "01", "L", "E")
              psRegistro = psRegistro & gdl_Funcion.PadR(gdl_Funcion.aTexto(porstRecordset!numdociden), 12, s_Caracter)   ' Documento de identificación del trabajador
              s_Contenido = gdl_Funcion.PadR(UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres)), 35, s_Caracter)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 35, s_Caracter)                               ' Nombre del trabajador
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me)) * 100
              psRegistro = psRegistro & gdl_Funcion.PadL(n_Importe, 14, "0")                                              ' Importe de la operación
              ' sumatoria abonos
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              nSumImporte = nSumImporte + n_Importe
              
              psRegistro = psRegistro & "0011"
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "")
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, 20, s_Caracter)                                     ' Cuenta de deposito
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo))
            End If
            potxtFileExp.WriteLine psRegistro
           Case "3", "10"                   ' Formato ScotiaBank
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              ' 1: Codigo o documento de identidad empleado
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 8
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: Nombre del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
              n_Longitud = 30
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: Concepto del pago
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo)): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: Fecha de proceso
              s_Contenido = Format(porstRecordset!fechaproce, "yyyymmdd"): n_Longitud = 8
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: Monto a pagar
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = CDec(n_Importe) * 100: n_Longitud = 11
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              nSumImporte = nSumImporte + n_Importe
              ' 6: Forma de pago - constante
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "4", "3"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: Cuenta de abono
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 10
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "", s_Contenido)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: Numero documento empleado
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 8
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 9: constante - vacio
              s_Contenido = "": n_Longitud = 1
              If porstRecordset!Formato <> "10" Then
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              End If
              ' 10: Cuenta interbancaria
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "", s_Contenido)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, IIf(porstRecordset!Formato = "10", "0", s_Caracter))
              potxtFileExp.WriteLine psRegistro
            End If
           Case "4"                   ' Formato interbank
            nSecuencia = nSecuencia + 1
            psRegistro = psRegistro & "02"                                                                                    ' Formato del registro
            psRegistro = psRegistro & gdl_Funcion.PadR(gdl_Funcion.aTexto(porstRecordset!numdociden), 20, s_Caracter)         ' Codigo o documento del empleado
            psRegistro = psRegistro & gdl_Funcion.PadR("", 29, s_Caracter)                                                    ' Espacios en blanco
            psRegistro = psRegistro & IIf(gdl_Funcion.aTexto(porstRecordset!codmon) = s_Codmon_mn, "01", "10")                ' Moneda de cuenta de deposito
            n_Importe = Format(CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me)), "###########0.00") * 100
            psRegistro = psRegistro & gdl_Funcion.PadL(n_Importe, 15, "0")                                                    ' Importe de la operación
            ' sumatoria abonos
            n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
            nSumImporte = nSumImporte + n_Importe
            s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "1", "0"): n_Longitud = 1
            psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)                                   ' Indicador de cuenta de banco u otro banco
            s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "99", "09"): n_Longitud = 2
            psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)                                   ' Tipo de abono cuenta
            psRegistro = psRegistro & gdl_Funcion.PadR(IIf(ribAnalisis(0).Value, "002", "007"), 3, s_Caracter)                ' Producto de pago cuenta de ahorros o cuenta de cts
            psRegistro = psRegistro & IIf(gdl_Funcion.aTexto(porstRecordset!codmon) = s_Codmon_mn, "01", "10")
            s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 3
            s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "", s_Contenido)
            psRegistro = psRegistro & gdl_Funcion.PadR(Left(s_Contenido, 3), n_Longitud, s_Caracter)                          ' Moneda de la cuenta de deposito
            s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
            s_Contenido = Mid(s_Contenido, IIf(porstRecordset!interbankpago = s_Estado_Act, 1, 4))
            psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)                                   ' Oficina de la cuenta de deposito
            psRegistro = psRegistro & gdl_Funcion.PadR("P", 1, s_Caracter)                                                    ' Indicador de Tipo de persona
            s_Contenido = gdl_Funcion.aTexto(porstRecordset!coddci): n_Longitud = 2
            s_Contenido = IIf(s_Contenido = "04", "03", s_Contenido)
            psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)                                   ' Tipo de documento del beneficiario
            psRegistro = psRegistro & gdl_Funcion.PadR(gdl_Funcion.aTexto(porstRecordset!numdociden), 15, s_Caracter)         ' Numero de documento de identidad
            psRegistro = psRegistro & gdl_Funcion.PadR(UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno)), 20, s_Caracter)  ' Apellido paterno
            psRegistro = psRegistro & gdl_Funcion.PadR(UCase(gdl_Funcion.aTexto(porstRecordset!apematerno)), 20, s_Caracter)  ' Apellido materno
            psRegistro = psRegistro & gdl_Funcion.PadR(UCase(gdl_Funcion.aTexto(porstRecordset!nombres)), 20, s_Caracter)     ' Nombres
            psRegistro = psRegistro & gdl_Funcion.PadR("", 23, s_Caracter)                                                    ' Espacios en blanco
            potxtFileExp.WriteLine psRegistro
           Case "5"                   ' Formato bbva
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              '1: Tipo de registro - constante: 002
              s_Contenido = "002": n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '2: Tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "L", IIf(s_Contenido = "CEX", "E", IIf(s_Contenido = "RUC", "R", "P")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 1, s_Caracter)
              '3: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '4: Tipo de abono
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "P", "I"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '5: Numero de cuenta abono
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '6: Nombre del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
              s_Contenido = Left(s_Contenido, 40): n_Longitud = 40
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '7: Importe de transferencia
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = Replace(Replace(FormatNumber(CDec(n_Importe), 2), ",", "", 1), ".", "", 1): n_Longitud = 15
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              nSumImporte = nSumImporte + n_Importe
              '8: Referencia
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo)): n_Longitud = 40
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '9: Indicador de aviso
              s_Contenido = "": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '10: Medio de aviso
              s_Contenido = "": n_Longitud = 50
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '11: Indicador de proceso - vacio
              s_Contenido = "": n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              '12: Descripción - vacio
              s_Contenido = "": n_Longitud = 30
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              
              If ribAnalisis(1).Value Then      ' cts
                '13: Filler - vacio
                s_Contenido = "": n_Longitud = 21
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
                '14: Importe de remuneracion
                n_Importe = Round(CDec(IIf(porstRecordset!pagodolar = s_Estado_Ina, porstRecordset!remunera_mn, porstRecordset!remunera_me)) * 4, 2)
                s_Contenido = Format(CDbl(n_Importe), "###########0.00") * 100: n_Longitud = 15
                psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
                '15: Moneda remuneracion
                s_Contenido = IIf((gdl_Funcion.aTexto(porstRecordset!pagodolar) = s_Estado_Ina Or porstRecordset!interbankpago = s_Estado_Act), "PEN", "USD"): n_Longitud = 3
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              End If
              '16: Filler - vacio
              s_Contenido = "": n_Longitud = 18
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
            End If
           Case "6"                   ' Formato bcp
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              '1: Tipo de registro constante: 2
              s_Contenido = s_Estado_Blq
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 1, s_Caracter)
              If ribAnalisis(0).Value Then      ' Remuneraciones
                '2: Tipo de cuenta abono
                'A: Cuenta de Ahorros C: Cuenta Corriente M: Cuenta Maestra B: Interbancaria (CCI)
                s_Contenido = Left(gdl_Funcion.aTexto(porstRecordset!cuentapago), 1)
                s_Contenido = IIf(IsNumeric(s_Contenido), "A", s_Contenido)
                s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "B", s_Contenido)
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 1, s_Caracter)
              End If
              '3: Numero de cuenta abono
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!cuentapago)
              s_Contenido = Mid(s_Contenido, IIf(IsNumeric(Left(s_Contenido, 1)), 1, 2))
              s_Contenido = Replace(s_Contenido, "-", "", 1)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 20, s_Caracter)
              '4: Tipo documento empleado
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci)
              s_Contenido = IIf(s_Contenido = "DNI", "1", IIf(s_Contenido = "CEX", "3", "4"))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 1, s_Caracter)
              '5: Numero documento empleado
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 12, s_Caracter)
              '6: Correlativo menor edad - espacio blanco
              psRegistro = psRegistro & gdl_Funcion.PadR("", 3, s_Caracter)
              '7: Nombre del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 75, s_Caracter)
              '8: Referencia beneficiario
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 40, s_Caracter)
              '9: Referencia empresa
              s_Contenido = UCase(Left(gdl_Funcion.aTexto(porstRecordset!desmotivo), 20))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 20, s_Caracter)
              '10: moneda del importe
              s_Contenido = IIf(gdl_Funcion.aTexto(porstRecordset!codmon) = s_Codmon_mn, "0001", "1001")
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 4, s_Caracter)
              '11: Importe de trasnferencia
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              psRegistro = psRegistro & gdl_Funcion.PadL(Format(n_Importe, "############0.00"), 17, "0")
              nSumImporte = nSumImporte + n_Importe
              If ribAnalisis(0).Value Then      ' Remuneraciones
                '12: Validar IDC
                psRegistro = psRegistro & gdl_Funcion.PadR("S", 1, s_Caracter)
              Else
                '12: Moneda remuneracion
                s_Contenido = IIf(gdl_Funcion.aTexto(porstRecordset!pagodolar) = s_Estado_Ina, "0001", "1001")
                psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 4, s_Caracter)
                '13: Importe de remuneracion
                n_Importe = Round(CDec(IIf(porstRecordset!pagodolar = s_Estado_Ina, porstRecordset!remunera_mn, porstRecordset!remunera_me)) * 4, 2)
                psRegistro = psRegistro & gdl_Funcion.PadL(Format(n_Importe, "############0.00"), 17, "0")
              End If
              potxtFileExp.WriteLine psRegistro
            End If
           Case "7"                   ' Formato bnp
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              '1: Numero de cuenta abono
              s_Contenido = Mid(Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1), 2)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, 11, s_Caracter)
              '2: Importe de transferencia
              n_Importe = Format(CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me)), "############0.00") * 100
              psRegistro = psRegistro & gdl_Funcion.PadL(n_Importe, 15, "0")
              ' sumatoria abonos
              nSumImporte = nSumImporte + n_Importe
              potxtFileExp.WriteLine psRegistro
            End If
           Case "8"                   ' Formato citibank
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              ' 1: Tipo de registro - constante
              s_Contenido = "PAY": n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: Codigo del pais - constante
              s_Contenido = "604": n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: Codigo del pais - constante
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentabco), "-", ""): n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 4: Fecha de pago
              s_Contenido = Format(porstRecordset!fechaproce, "yymmdd"): n_Longitud = 6
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: Tipo de pago - constante
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "071", IIf(porstRecordset!tippago = "3", "073", "072")): n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: Referencia del cliente
              s_Contenido = "PERIOD " & txtPeriodo.Text: n_Longitud = 15
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: Secuencia
              s_Contenido = nSecuencia: n_Longitud = 8
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 8: Numero documento empleado
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 9: Moneda
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!codmon)
              s_Contenido = IIf(s_Contenido = s_Codmon_mn, "PEN", "USD"): n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 10: Numero documento empleado - beneficiario
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 11: Importe de pago
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = Format(CDbl(n_Importe), "###########0.00") * 100: n_Longitud = 15
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              nSumImporte = nSumImporte + n_Importe
              ' 12: Fecha de vencimiento - constante vacio
              s_Contenido = "": n_Longitud = 6
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 13: Detalle de pago 1
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo)): n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 14: Detalle de pago 2 - constante vacio
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 15: Detalle de pago 3 - constante vacio
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 15: Detalle de pago 4 - constante vacio
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 16: Transacción local - constante
              s_Contenido = IIf(porstRecordset!tippago = "3", "00", "21")
              s_Contenido = IIf(ribAnalisis(0).Value, s_Contenido, "23"): n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 17: Tipo cuenta empresa - constante
              s_Contenido = "01": n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 18: Nombre beneficiario
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
              n_Longitud = 80
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 19: Dirección beneficiario - 1
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!nomviadirec) & " N° " & gdl_Funcion.aTexto(porstRecordset!numerdirec))
              n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 20: Dirección beneficiario - 2
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 21: Distrito
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!ubigeodir): n_Longitud = 15
              s_Contenido = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, s_Contenido, "UB")
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 22: Estado beneficiario - blanco
              s_Contenido = "": n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 23: Codigo postal
              s_Contenido = "": n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 23: Telefono beneficiario
              s_Contenido = "": n_Longitud = 16
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 24: Codigo entidad banco interbancario
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!codentidadbnk): n_Longitud = 3
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, s_Contenido, "")
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 25: Agencia del banco
              s_Contenido = "": n_Longitud = 8
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "00000000", s_Contenido)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 26: Cuenta interbancaria
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!cuentapago): n_Longitud = 35
              s_Contenido = Mid(s_Contenido, IIf(ribAnalisis(0).Value, 2, 1))
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!codentidadbnk) & s_Contenido
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, s_Contenido, "")
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 27: Tipo de cuenta interbancaria
              s_Contenido = Left(gdl_Funcion.aTexto(porstRecordset!cuentapago), 1): n_Longitud = 2
              s_Contenido = IIf((s_Contenido = "A" Or ribAnalisis(1).Value), "02", "01")
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, s_Contenido, ""): n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 28: Dirección del banco constante
              s_Contenido = "": n_Longitud = 30
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 29: Entidad del banco constante
              s_Contenido = "": n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 30: Numero agencia del banco constante
              s_Contenido = "": n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 31: Nombre agencia del banco constante
              s_Contenido = "": n_Longitud = 14
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 32: Numero sucursal (pais) constante
              s_Contenido = "": n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 33: Nombre sucursal del banco constante
              s_Contenido = "": n_Longitud = 19
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 34: Numero fax beneficiario
              s_Contenido = "": n_Longitud = 16
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 35: Persona contacto beneficiario
              s_Contenido = "": n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 36: Departamento contacto beneficiario
              s_Contenido = "": n_Longitud = 15
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 37: Cuenta citibank
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!cuentapago): n_Longitud = 10
              s_Contenido = Mid(s_Contenido, IIf(ribAnalisis(0).Value, 2, 1))
              s_Contenido = IIf(porstRecordset!tippago = "3", "", s_Contenido)
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, s_Contenido, "")
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 38: Tipo cuenta citibank
              s_Contenido = Left(gdl_Funcion.aTexto(porstRecordset!cuentapago), 1): n_Longitud = 2
              s_Contenido = IIf((s_Contenido = "A" Or ribAnalisis(1).Value), "02", "01")
              s_Contenido = IIf(porstRecordset!tippago = "3", "", s_Contenido)
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, s_Contenido, "")
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 39: Sucursal de destino
              s_Contenido = IIf(porstRecordset!tippago = "3", "888", "001"): n_Longitud = 3
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "099", s_Contenido)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 40: Collection ID - constante
              s_Contenido = "": n_Longitud = 50
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 41: Codigo actividad beneficiario - constante
              s_Contenido = "": n_Longitud = 5
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 42: Email beneficiario - constante
              s_Contenido = "": n_Longitud = 50
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 43: Valor maximo pago - constante
              s_Contenido = "100000000": n_Longitud = 15
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 44: Tipo de actualizacion - constante
              s_Contenido = "3": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 45: Numero cheque - constante
              s_Contenido = "": n_Longitud = 11
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 46: Cheque impreso - constante
              s_Contenido = "": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 47: Encaje pago - constante
              s_Contenido = "": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 48: Caracter blanco
              s_Contenido = "": n_Longitud = 254
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
            End If
           Case "9"                   ' Formato bbva cash
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              ' primera linea
              ' 1: tipo de registro - constante
              s_Contenido = "2210": n_Longitud = 4
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: tipo documento empresa - constante
              s_Contenido = "R": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: documento empresa
              s_Contenido = ps_RucEmpresa: n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "L", IIf(s_Contenido = "CEX", "E", IIf(s_Contenido = "RUC", "R", "P")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: Nombre del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
              n_Longitud = 35: s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: Importe de transferencia
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = CDec(n_Importe) * 100: n_Longitud = 14
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              nSumImporte = nSumImporte + n_Importe
              ' 9: Moneda remuneracion
              s_Contenido = IIf(gdl_Funcion.aTexto(porstRecordset!pagodolar) = s_Estado_Ina, "PEN", "USD"): n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 10: libre
              s_Contenido = "": n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 11: codigo devolucion - constante
              s_Contenido = "0000": n_Longitud = 4
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 12: mensaje devolucion - constante
              s_Contenido = "": n_Longitud = 40
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 13: libre
              s_Contenido = "": n_Longitud = 117
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
              
              ' segunda linea
              psRegistro = ""
              ' 1: tipo de registro - constante
              s_Contenido = "2220": n_Longitud = 4
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: tipo documento empresa - constante
              s_Contenido = "R": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: documento empresa
              s_Contenido = ps_RucEmpresa: n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "L", IIf(s_Contenido = "CEX", "E", IIf(s_Contenido = "RUC", "R", "P")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: codigo banco cuenta - constante
              s_Contenido = "0011": n_Longitud = 4
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: Numero de cuenta abono
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: direccion trabajador
              n_Longitud = 35
              s_Contenido = Left(UCase(gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!nomviadirec) & " N° " & gdl_Funcion.aTexto(porstRecordset!numerdirec) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!nomzondirec)), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 9: Tipo de abono
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "P", "I"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 10: libre
              s_Contenido = "": n_Longitud = 24
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 11: tipo cuenta - constante
              s_Contenido = "00": n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 12: libre
              s_Contenido = "": n_Longitud = 139
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
              
              ' tercera linea
              psRegistro = ""
              ' 1: tipo de registro - constante
              s_Contenido = "2230": n_Longitud = 4
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: tipo documento empresa - constante
              s_Contenido = "R": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: documento empresa
              s_Contenido = ps_RucEmpresa: n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "L", IIf(s_Contenido = "CEX", "E", IIf(s_Contenido = "RUC", "R", "P")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: distrito personal
              n_Longitud = 35
              s_Contenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, gdl_Funcion.aTexto(porstRecordset!ubigeodir), "UB"), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: provincia personal
              n_Longitud = 25
              s_Contenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Act, Left(gdl_Funcion.aTexto(porstRecordset!ubigeodir), 4), "UB"), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: departamento personal
              n_Longitud = 25
              s_Contenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Ina, Left(gdl_Funcion.aTexto(porstRecordset!ubigeodir), 2), "UB"), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 9: libre
              s_Contenido = "": n_Longitud = 140
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
              
              ' cuarto linea
              psRegistro = ""
              ' 1: tipo de registro - constante
              s_Contenido = "2240": n_Longitud = 4
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: tipo documento empresa - constante
              s_Contenido = "R": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: documento empresa
              s_Contenido = ps_RucEmpresa: n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "L", IIf(s_Contenido = "CEX", "E", IIf(s_Contenido = "RUC", "R", "P")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: concepto o referencia 1
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo)): n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: concepto o referencia 1 - vacio
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: libre
              s_Contenido = "": n_Longitud = 155
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
            End If
           Case "11"                   ' Formato bif
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              ' 1: secuencia
              s_Contenido = nSecuencia: n_Longitud = 7
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, s_Caracter)
              ' 2: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "1", IIf(s_Contenido = "CEX", "3", IIf(s_Contenido = "RUC", "6", "4")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 11
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: Apellido paterno del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno)): n_Longitud = 20
              s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: Apellido materno del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apematerno)): n_Longitud = 20
              s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: Nombres del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!nombres)): n_Longitud = 44
              s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: direccion trabajador - vacio
              n_Longitud = 60
              s_Contenido = Left(UCase(gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!nomviadirec) & " N° " & gdl_Funcion.aTexto(porstRecordset!numerdirec) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!nomzondirec)), n_Longitud)
              s_Contenido = ""
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: telefono trabajador - vacio
              s_Contenido = "": n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 9: Tipo planilla
              s_Contenido = IIf(ribAnalisis(0).Value, "H", "C"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 10: Codigo banco - constante
              s_Contenido = "038": n_Longitud = 3
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 11: Numero de cuenta abono
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, s_Caracter)
              ' 12: Moneda remuneracion
              s_Contenido = IIf(gdl_Funcion.aTexto(porstRecordset!pagodolar) = s_Estado_Ina, "1", "2"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 13: Importe de transferencia
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = Replace(Replace(FormatNumber(CDec(n_Importe), 2), ",", "", 1), ".", "", 1): n_Longitud = 14
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, s_Caracter)
              nSumImporte = nSumImporte + n_Importe
              ' 14: Motivo deposito
              s_Contenido = IIf(ribAnalisis(0).Value, "5", "0"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
            End If
           Case "12"                   ' Formato wiesse sudameris(FInalizar)
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              ' 1: ruc empresa
              s_Contenido = ps_RucEmpresa: n_Longitud = 11
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: tipo de servicio 20:MN, 21:ME
              s_Contenido = IIf(fMenu.ribMoneda(0).Value, "20", "21"): n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: constante vacio
              s_Contenido = "": n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: codigo empleado
              n_Longitud = 10: s_Contenido = Left(gdl_Funcion.aTexto(porstRecordset!codpsn), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: nombre del trabajador
              n_Longitud = 30
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!nombres)) & " "
              s_Contenido = s_Contenido & UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno)) & " "
              s_Contenido = s_Contenido & UCase(gdl_Funcion.aTexto(porstRecordset!apematerno))
              s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: situacion contsante - activo
              s_Contenido = "1": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: constante vacio
              s_Contenido = "": n_Longitud = 8
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: constante vacio- cero
              s_Contenido = "0.0000": n_Longitud = 6
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, s_Caracter)
              ' 9: Importe de transferencia
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = Replace(FormatNumber(CDec(n_Importe), 2), ",", "", 1): n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              nSumImporte = nSumImporte + n_Importe
              ' 10: importe 2 - cero
              s_Contenido = "0.00": n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 10: importe 3 - cero
              s_Contenido = "0.00": n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 10: importe 4 - cero
              s_Contenido = "0.00": n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 10: importe 5 - cero
              s_Contenido = "0.00": n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 10: importe 6 - cero
              s_Contenido = "0.00": n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' 11: cuenta abono
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 14
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Act, "", s_Contenido)
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, s_Caracter)
              ' 12: constante vacio
              s_Contenido = "": n_Longitud = 16
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 13: modalida de cobro
              s_Contenido = "5": n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 14: cuenta de cargo
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentabco), "-", "", 1): n_Longitud = 14
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 15: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!coddci): n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 16: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 17: nombre responsable
              n_Longitud = 30
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!psnnombres)) & " "
              s_Contenido = s_Contenido & UCase(gdl_Funcion.aTexto(porstRecordset!psnapepaterno)) & " "
              s_Contenido = s_Contenido & UCase(gdl_Funcion.aTexto(porstRecordset!psnapematerno))
              s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 18: Fecha de proceso
              s_Contenido = Format(porstRecordset!fechaproce, "ddmmyyyy"): n_Longitud = 8
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 19: Cuenta interbancaria
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "", s_Contenido)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
            End If
           Case "13"                   ' Formato bbva cash detalle
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              psRegistro = ""
              ' 1: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "L", IIf(s_Contenido = "CEX", "E", IIf(s_Contenido = "RUC", "R", "P")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 12
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: Nombre del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
              n_Longitud = 35: s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: Importe de transferencia
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = CDec(n_Importe) * 100: n_Longitud = 14
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' sumatoria
              nSumImporte = nSumImporte + n_Importe
              ' 5: codigo banco cuenta
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!codentidad): n_Longitud = 4
              If porstRecordset!interbankpago = s_Estado_Act Then
                s_Contenido = gdl_Funcion.aTexto(porstRecordset!codentidadbnk)
              End If
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: Numero de cuenta abono
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: direccion trabajador
              n_Longitud = 35
              s_Contenido = Left(Trim(UCase(gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!nomviadirec) & IIf(gdl_Funcion.aTexto(porstRecordset!numerdirec) = "", "", " N° ") & gdl_Funcion.aTexto(porstRecordset!numerdirec) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!nomzondirec))), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: distrito trabajador
              n_Longitud = 35
              s_Contenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, gdl_Funcion.aTexto(porstRecordset!ubigeodir), "UB"), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 9: provincia trabajador
              n_Longitud = 25
              s_Contenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Act, Left(gdl_Funcion.aTexto(porstRecordset!ubigeodir), 4), "UB"), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 10: departamento trabajador
              n_Longitud = 25
              s_Contenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Ina, Left(gdl_Funcion.aTexto(porstRecordset!ubigeodir), 2), "UB"), n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 11: concepto o referencia 1
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo)): n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 12: concepto o referencia 2 - vacio
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 13: concepto o referencia 3 - vacio
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 14: concepto o referencia 4 - vacio
              s_Contenido = "": n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 15: Forma de pago
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "P", "I"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 16: tipo cuenta
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "00", "02"): n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
            End If
           Case "14"                   ' Formato bbva cash síntesis
            s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
            If porstRecordset!codmon = s_Contenido Then
              nSecuencia = nSecuencia + 1
              psRegistro = ""
              ' 1: Número documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!numdociden): n_Longitud = 10
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 2: Nombre del trabajador
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!apepaterno) & " " & gdl_Funcion.aTexto(porstRecordset!apematerno) & " " & gdl_Funcion.aTexto(porstRecordset!nombres))
              n_Longitud = 35: s_Contenido = Left(s_Contenido, n_Longitud)
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 3: direccion trabajador - constante
              n_Longitud = 2
              s_Contenido = "A"
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 4: distrito trabajador - constante
              n_Longitud = 2
              s_Contenido = "A"
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 5: provincia trabajador - constante
              n_Longitud = 2
              s_Contenido = "A"
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 6: departamento trabajador - constante
              n_Longitud = 2
              s_Contenido = "A"
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 7: Numero de cuenta abono
              s_Contenido = Replace(gdl_Funcion.aTexto(porstRecordset!cuentapago), "-", "", 1): n_Longitud = 20
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 8: Importe de transferencia
              n_Importe = CDec(IIf(porstRecordset!codmon = s_Codmon_mn, porstRecordset!importe_mn, porstRecordset!importe_me))
              s_Contenido = CDec(n_Importe) * 100: n_Longitud = 14
              psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
              ' sumatoria
              nSumImporte = nSumImporte + n_Importe
              ' 9: concepto o referencia 1
              s_Contenido = UCase(gdl_Funcion.aTexto(porstRecordset!desmotivo)): n_Longitud = 35
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 10: codigo entidad banacaria
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!codentidad): n_Longitud = 4
              If porstRecordset!interbankpago = s_Estado_Act Then
                s_Contenido = gdl_Funcion.aTexto(porstRecordset!codentidadbnk)
              End If
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 11: tipo documento identidad
              s_Contenido = gdl_Funcion.aTexto(porstRecordset!sigladci): n_Longitud = 1
              s_Contenido = IIf(s_Contenido = "DNI", "L", IIf(s_Contenido = "CEX", "E", IIf(s_Contenido = "RUC", "R", "P")))
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 12: Forma de pago
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "P", "I"): n_Longitud = 1
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              ' 13: tipo cuenta
              s_Contenido = IIf(porstRecordset!interbankpago = s_Estado_Ina, "00", "02"): n_Longitud = 2
              psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
              potxtFileExp.WriteLine psRegistro
            End If
          End Select
          ' Incremento el porcentaje
          nRegistro = nRegistro + 1
          fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
          DoEvents
          porstRecordset.MoveNext
          If porstRecordset.EOF Then Exit Do
          ' sumatoria limite de archivo
          s_Contenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
          n_Importe = nSumImporte
          If porstRecordset!codmon = s_Contenido Then
            n_Importe = CDec(IIf(fMenu.ribMoneda(0).Value, porstRecordset!importe_mn, porstRecordset!importe_me))
            n_Importe = nSumImporte + n_Importe
          End If
          n_Importe = IIf(nImporteLimite = 0, 0, n_Importe)
        Loop While (nImporteLimite >= n_Importe)
        ' Retorno un registro para datos finales
        porstRecordset.MovePrevious
        nRegistro = nRegistro - 1
        
        ' linea final del archivo
        psRegistro = ""
        ' Sumatoria de depositos
        If porstRecordset!Formato = "8" Then   ' citibank
          ' 1: Tipo de registro - constante
          s_Contenido = "TRL": n_Longitud = 3
          psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
          ' 2: Cantidad registros pago
          s_Contenido = nSecuencia: n_Longitud = 15
          psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
          ' 3: Sumatoria transferencia
          s_Contenido = Format(CDbl(nSumImporte), "###########0.00") * 100: n_Longitud = 15
          psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
          ' 4: Cantidad registros - constante
          s_Contenido = "0": n_Longitud = 15
          psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, "0")
          ' 5: Cantidad registros enviados
          s_Contenido = nSecuencia: n_Longitud = 15
          psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
          ' 48: Caracter blanco
          s_Contenido = "": n_Longitud = 37
          psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
          potxtFileExp.WriteLine psRegistro
        ElseIf porstRecordset!Formato = "9" Then   ' bbva conticash
          psRegistro = ""
          ' 1: tipo de registro - constante
          s_Contenido = "2910": n_Longitud = 4
          psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
          ' 2: tipo documento empresa - constante
          s_Contenido = "R": n_Longitud = 1
          psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
          ' 3: documento empresa
          s_Contenido = ps_RucEmpresa: n_Longitud = 12
          psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
          ' 4: Cantidad registros archivo
          s_Contenido = ((nSecuencia + 1) * 4): n_Longitud = 10
          psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
          ' 5: Cantidad abonos
          s_Contenido = nSecuencia: n_Longitud = 8
          psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
          ' 6: Sumatoria transferencia
          s_Contenido = Format(CDbl(nSumImporte), "###########0.00") * 100: n_Longitud = 14
          psRegistro = psRegistro & gdl_Funcion.PadL(s_Contenido, n_Longitud, "0")
          ' 8: libre
          s_Contenido = "": n_Longitud = 106
          psRegistro = psRegistro & gdl_Funcion.PadR(s_Contenido, n_Longitud, s_Caracter)
          potxtFileExp.WriteLine psRegistro
        End If
        ' Cierro objeto y saco de memoria
        potxtFileExp.Close
      End If
      ' Incremento el porcentaje
      nRegistro = nRegistro + 1
      fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
      DoEvents
      porstRecordset.MoveNext
    Wend
    ' saco objetos de memoria
    Set potxtFileExp = Nothing
    Set pofsoFileExp = Nothing
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Elimino archivo temporal
  gdl_Conexion.Execucion "DROP TABLE IF EXISTS tmpdepositocts", Elimina
  
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
  s_Sql = "SELECT DISTINCTROW codcls, nrocarta, desmotivo, fechaproce, codcpc, porinteres "
  s_Sql = s_Sql & "FROM plcartabanco "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND codpdo='" & Trim(txtPeriodo.Text) & "' "
  s_Sql = s_Sql & "AND codbco='" & Trim(txtBanco.Text) & "' "
  s_Sql = s_Sql & "ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaRegistro, tdbRegistro, s_Sql, adCmdText, adLockReadOnly

End Sub
Private Function InicioArchivo(ByVal sRecordInicial As String, ByVal nFormato As Integer, ByVal o_rstRecordset As ADODB.Recordset) As String
  Dim nPosIni As Integer, nPosFin As Integer, nLongitud As Integer
  Dim nImporte As Double, nTotalImporte As Double, nLimiteArchivo As Double
  Dim sCheckSum As String, nCheckSum As Double
  Dim nRegistros As Long, nTamano As Integer
  Dim sCaracter As String, sContenido As String
  Dim porstClone As ADODB.Recordset
  
  ' Instancio el objeto
  Set porstClone = o_rstRecordset.Clone()
  nTotalImporte = 0: nCheckSum = 0: nRegistros = 0
  sCaracter = " ": sCheckSum = ""
  InicioArchivo = ""
  ' registro inicial
  porstClone.Find ("cprimarykey='" & sRecordInicial & "'")
  If ribAnalisis(0).Value Then      ' Remuneraciones
    If nFormato = "1" Then          ' Formato banco credito
      ' 1: Indicador de planilla nueva
      sContenido = "#": nTamano = nFormato
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 2: Tipo de registro
      sContenido = "1": nTamano = nFormato
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 3: Tipo de pago masivo
      sContenido = "H": nTamano = nFormato
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 4: Tipo de producto
      sContenido = "C": nTamano = nFormato
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 5: Cuenta cargo - sucursal
      nPosIni = 1
      sContenido = gdl_Funcion.aTexto(porstClone!cuentabco)
      nPosFin = InStr(nPosIni, sContenido, "-", vbBinaryCompare)
      nLongitud = nPosFin - nPosIni
      sContenido = Mid(sContenido, nPosIni, nLongitud)
      nTamano = 3
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: Validación de cuenta cargo - numero
      nPosIni = nPosFin + 1
      sContenido = gdl_Funcion.aTexto(porstClone!cuentabco)
      nPosFin = InStr(nPosIni, sContenido, "-", vbBinaryCompare)
      nLongitud = nPosFin - nPosIni
      sContenido = Mid(sContenido, nPosIni, nLongitud)
      nTamano = 8
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      sCheckSum = sContenido
      ' 7: Validación de cuenta cargo - moneda
      nPosIni = nPosFin + 1
      sContenido = gdl_Funcion.aTexto(porstClone!cuentabco)
      nPosFin = InStr(nPosIni, sContenido, "-", vbBinaryCompare)
      nLongitud = nPosFin - nPosIni
      sContenido = Mid(sContenido, nPosIni, nLongitud)
      nTamano = nFormato
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      sCheckSum = sContenido
      ' 8: Validación de cuenta cargo - control y espacios
      nPosIni = nPosFin + 1
      sContenido = gdl_Funcion.aTexto(porstClone!cuentabco)
      sContenido = Mid(sContenido, nPosIni)
      nTamano = 8
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      sCheckSum = sContenido
      ' 9: Moneda de transferencia
      InicioArchivo = InicioArchivo & Left(IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn_Txt, s_Codmon_me_Txt), 2)
      ' Acumulo sumatoria de cuentas
      nCheckSum = nCheckSum + CDbl(sCheckSum)
    ElseIf nFormato = "5" Then      ' Formato bbva
      ' 1: Tipo de registro - constante
      sContenido = "700": nTamano = 3
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 2: Cuenta cargo
      sContenido = Replace(gdl_Funcion.aTexto(porstClone!cuentabco), "-", "", 1)
      nTamano = 20
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 3: Moneda de cuenta cargo
      sContenido = IIf(fMenu.ribMoneda(0).Value, "PEN", "USD")
      nTamano = 3
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    ElseIf nFormato = "6" Then          ' Formato bcp
      ' 4: Subtipo de planilla
      sContenido = "X": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 5: Tipo de cuenta cargo
      sContenido = "C": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: Moneda de cuenta cargo
      sContenido = IIf(fMenu.ribMoneda(0).Value, "0001", "1001")
      nTamano = 4
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 7: Cuenta cargo
      sContenido = Replace(gdl_Funcion.aTexto(porstClone!cuentabco), "-", "", 1)
      nTamano = 20
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' Acumulo sumatoria de cuentas
      sCheckSum = Mid(Replace(gdl_Funcion.aTexto(porstClone!cuentabco), "-", "", 1), 4, 10)
      nCheckSum = nCheckSum + CDbl(sCheckSum)
    End If
  Else                              ' C.T.S.
    If nFormato = "5" Then      ' Formato bbva
      ' 1: Tipo de registro - constante
      sContenido = IIf(fMenu.ribMoneda(0).Value, "600", "610")
      nTamano = 3
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 2: Cuenta cargo
      sContenido = Replace(gdl_Funcion.aTexto(porstClone!cuentabco), "-", "", 1)
      nTamano = 20
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 3: Moneda de cuenta cargo
      sContenido = IIf(fMenu.ribMoneda(0).Value, "PEN", "USD")
      nTamano = 3
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    ElseIf nFormato = "6" Then          ' Formato bcp
      ' 4: Tipo de cuenta cargo
      sContenido = "C": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 5: Moneda de cuenta cargo
      sContenido = IIf(fMenu.ribMoneda(0).Value, "0001", "1001")
      nTamano = 4
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: Cuenta cargo
      sContenido = Replace(gdl_Funcion.aTexto(porstClone!cuentabco), "-", "", 1)
      nTamano = 20
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 7: Tipo documento empresa - constante
      sContenido = "6": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 8: numero documento empresa
      sContenido = ps_RucEmpresa: nTamano = 12
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' Acumulo sumatoria de cuentas
      sCheckSum = Mid(Replace(gdl_Funcion.aTexto(porstClone!cuentabco), "-", "", 1), 4, 10)
      nCheckSum = nCheckSum + CDbl(sCheckSum)
    Else
      ' 1: Tipo de registro - constante
      sContenido = "2": nTamano = 1
      InicioArchivo = gdl_Funcion.PadR(sContenido, nTamano, sCaracter) & InicioArchivo
      ' 2: Espacios - constante
      sContenido = "": nTamano = 14
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 3: Moneda de transferencia - constante
      sContenido = "M" & IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me): nTamano = 2
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 4: Espacios - constante
      sContenido = "": nTamano = 6
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 5: Documento empresa
      sContenido = ps_RucEmpresa: nTamano = 11
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: Codigo trabajadores - constante
      sContenido = "00000000000": nTamano = 11
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 7: Espacios - constante
      sContenido = "": nTamano = 6
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    End If
  End If
  
  ' Validacion de sumatorias
  nLimiteArchivo = CDec(porstClone!impolimite)
  Do
    sContenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
    If porstClone!codmon = sContenido Then
      If (nFormato = "1" Or nFormato = "6") Then          ' Formato bcp
        sCheckSum = ""
        nPosIni = 11: nLongitud = 10
        sContenido = Replace(gdl_Funcion.aTexto(porstClone!cuentapago), "-", "", 1)
        If porstClone!interbankpago <> s_Estado_Act Then
          nPosIni = 4: nLongitud = 11
        End If
        sContenido = Mid(sContenido, nPosIni, nLongitud)
        sCheckSum = sContenido
        nCheckSum = nCheckSum + CDbl(IIf(sCheckSum = "", 0, sCheckSum))
      End If
      ' sumatoria de importe de abonos
      nImporte = CDec(IIf(fMenu.ribMoneda(0).Value, porstClone!importe_mn, porstClone!importe_me))
      nTotalImporte = nTotalImporte + nImporte
      nRegistros = nRegistros + 1
    End If
    porstClone.MoveNext
    If porstClone.EOF Then Exit Do
    ' limite de archivo
    sContenido = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
    nImporte = nTotalImporte
    If porstClone!codmon = sContenido Then
      nImporte = CDec(IIf(fMenu.ribMoneda(0).Value, porstClone!importe_mn, porstClone!importe_me))
      nImporte = nTotalImporte + nImporte
    End If
    nImporte = IIf(nLimiteArchivo = 0, 0, nImporte)
  Loop While (nLimiteArchivo >= nImporte)
  ' Retorno un registro para datos finales
  porstClone.MovePrevious
  
  nImporte = CDec(nTotalImporte) * 100
  If ribAnalisis(0).Value Then      ' Remuneraciones
    If nFormato = "1" Then          ' Formato banco credito
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(Round(nImporte, 2), 15, "0")                                                ' Importe total de transferencia
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(Format(porstClone!fechaproce, "ddmmyyyy"), 8, sCaracter)           ' Fecha de proceso
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(UCase(gdl_Funcion.aTexto(porstClone!desmotivo)), 20, sCaracter)    ' Motivo de transferencia
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(nCheckSum, 15, "0")      ' Sumatoria de los codigos de cuentas
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(nRegistros, 6, "0")      ' Total de registros de transferencia
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR("1", 1, sCaracter)       ' Tipo de pago masivo constante
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR("", 15, sCaracter)       ' Identificador de dividendos contanste
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR("1", 1, sCaracter)       ' Identificador de npoota de cargo constante
    ElseIf nFormato = "5" Then      ' Formato bbva
      ' 4: Importe total de transferencia
      sContenido = Replace(Replace(FormatNumber(CDbl(nTotalImporte), 2), ",", "", 1), ".", "", 1): nTamano = 15
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      ' 5: Tipo de proceso - constante
      sContenido = "A": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: Fecha de proceso
      sContenido = Format(porstClone!fechaproce, "yyyymmdd"): nTamano = 8
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 7: Hora de proceso - constante
      sContenido = "": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 8: Motivo o Referencia de transferencia
      sContenido = UCase(gdl_Funcion.aTexto(porstClone!desmotivo)): nTamano = 25
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 9: Total de registro
      sContenido = nRegistros: nTamano = 6
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      ' 10: Caracter de validación - constante
      sContenido = "S": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 11: Valor control - constante
      sContenido = "": nTamano = 15
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 12: Indicador proceso - constante
      sContenido = "": nTamano = 3
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 13: Descripcion para banco - constante
      sContenido = "": nTamano = 30
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 14: Filter - constante
      sContenido = "": nTamano = 20
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    ElseIf nFormato = "6" Then          ' Formato bcp
      ' 3: Fecha de proceso
      sContenido = Format(porstClone!fechaproce, "yyyymmdd"): nTamano = 8
      InicioArchivo = gdl_Funcion.PadR(sContenido, nTamano, sCaracter) & InicioArchivo
      ' 2: Total de registro
      InicioArchivo = gdl_Funcion.PadL(nRegistros, 6, "0") & InicioArchivo
      ' 1: Tipo de registro
      sContenido = "1": nTamano = 1
      InicioArchivo = gdl_Funcion.PadR(sContenido, nTamano, sCaracter) & InicioArchivo
      ' 8: Importe total de transferencia
      sContenido = Format(CDbl(nTotalImporte), "#############0.00"): nTamano = 17
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      ' 9: Motivo de transferencia
      sContenido = UCase(gdl_Funcion.aTexto(porstClone!desmotivo)): nTamano = 40
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 10: Sumatoria de control
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(nCheckSum, 15, "0")
    End If
  Else
    If nFormato = "5" Then      ' Formato bbva
      ' 4: Importe a cargar
      sContenido = Format(CDbl(nTotalImporte), "#############0.00") * 100: nTamano = 15
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      ' 5: Tipo de proceso - constante inmediata
      sContenido = "A": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: Fecha de proceso - constate - inmediata
      sContenido = "": nTamano = 8
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 7: Hora de proceso - constante inmediata
      sContenido = "": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 8: Motivo de transferencia
      sContenido = UCase(gdl_Funcion.aTexto(porstClone!desmotivo)): nTamano = 25
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 9: Total de registro
      nTamano = 6
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(nRegistros, nTamano, "0")
      ' 10: Validacion de pertenecia
      sContenido = "S": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 11: Valor de control
      sContenido = "": nTamano = 15
      'InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, "0")
      ' 12: Indicador de proceso
      sContenido = "": nTamano = 3
      'InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, "0")
      ' 13: Descripción
      sContenido = "": nTamano = 30
      'InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 14: Filler
      sContenido = "": nTamano = 21
      'InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    ElseIf nFormato = "6" Then          ' Formato bcp
      ' 3: Fecha de proceso
      sContenido = Format(porstClone!fechaproce, "yyyymmdd"): nTamano = 8
      InicioArchivo = gdl_Funcion.PadR(sContenido, nTamano, sCaracter) & InicioArchivo
      ' 2: Total de registro
      InicioArchivo = gdl_Funcion.PadL(nRegistros, 6, "0") & InicioArchivo
      ' 1: Tipo de registro
      sContenido = "1": nTamano = 1
      InicioArchivo = gdl_Funcion.PadR(sContenido, nTamano, sCaracter) & InicioArchivo
      ' 9: Importe total de transferencia
      sContenido = Format(CDbl(nTotalImporte), "#############0.00"): nTamano = 17
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      ' 10: Motivo de transferencia
      sContenido = UCase(gdl_Funcion.aTexto(porstClone!desmotivo)): nTamano = 40
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 11: Sumatoria de control
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(nCheckSum, 15, "0")
    Else
      ' 8: Importe total transferencia
      sContenido = Replace(Format(CDbl(nTotalImporte), "###########0.00"), ".", ""): nTamano = 15
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      ' 9: Numero registros (trabajadores) transferencia
      sContenido = nRegistros: nTamano = 5
      InicioArchivo = InicioArchivo & gdl_Funcion.PadL(sContenido, nTamano, "0")
      ' 10: Espacios - constante
      sContenido = "": nTamano = 10
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 11: Tipo documento empresa - constante
      sContenido = "R": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 12: Identificador - constante
      sContenido = "1": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 13: Espacios - constante
      sContenido = "": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 14: Nombre empresa
      sContenido = ps_NomEmpresa: nTamano = 40
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 15: Dirección empresa
      sContenido = s_Direccion: nTamano = 40
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 17: Espacios - constante
      sContenido = "": nTamano = 18
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 18: Codigo empresa - constante
      sContenido = "": nTamano = 6
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 19: Espacios - constante
      sContenido = "": nTamano = 11
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 20: Marca - constante
      sContenido = "@": nTamano = 1
      InicioArchivo = InicioArchivo & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    End If
  End If
  Set porstClone = Nothing
  
End Function
Private Function InicioArchivo_BBVACash(ByVal sRecordInicial As String, ByVal nFormato As Integer, ByVal o_rstRecordset As ADODB.Recordset, ByVal nSecuencia As Integer) As String
  Dim sCaracter As String, sContenido As String
  Dim porstClone As ADODB.Recordset
  Dim nRegistros As Long, nTamano As Integer
  
  ' Instancio el objeto
  Set porstClone = o_rstRecordset.Clone()
  nRegistros = 0: sCaracter = " "
  InicioArchivo_BBVACash = ""
  ' registro inicial
  porstClone.Find ("cprimarykey='" & sRecordInicial & "'")
  If ribAnalisis(0).Value Then      ' Remuneraciones
    If nSecuencia = 1 Then
      ' 1: codigo registro - constante
      sContenido = "2110": nTamano = 4
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 2: tipo documento empresa - constante
      sContenido = "R": nTamano = 1
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 3: documento empresa
      sContenido = ps_RucEmpresa: nTamano = 12
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 4: fecha creacion
      sContenido = Format(porstClone!fechaproce, "ddmmyyyy"): nTamano = 8
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 5: fecha proceso
      sContenido = Format(porstClone!fechaproce, "ddmmyyyy"): nTamano = 8
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: cuenta corriente
      sContenido = Replace(gdl_Funcion.aTexto(porstClone!cuentabco), "-", "", 1): nTamano = 20
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 7: Moneda de cuenta cargo
      sContenido = IIf(fMenu.ribMoneda(0).Value, "PEN", "USD"): nTamano = 3
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 8: libre - constante
      sContenido = "": nTamano = 12
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 9: validacion - constante
      sContenido = "0": nTamano = 1
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 10: indicador devolucion - constante
      sContenido = "0": nTamano = 1
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 11: nombre archivo
      sContenido = fMenu.cdlDialogo.FileName: nTamano = 20
      sContenido = Mid(sContenido, InStrRev(sContenido, "\", -1, vbBinaryCompare) + 1)
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 12: servicio - constante
      sContenido = "108": nTamano = 3
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 13: libre
      sContenido = "": nTamano = 162
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    ElseIf nSecuencia = 2 Then
      ' 1: codigo registro - constante
      sContenido = "2120": nTamano = 4
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 2: tipo documento empresa - constante
      sContenido = "R": nTamano = 1
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 3: documento empresa
      sContenido = ps_RucEmpresa: nTamano = 12
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 4: razon social empresa
      nTamano = 35: sContenido = Left(ps_NomEmpresa, nTamano)
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 5: domicilio empresa
      nTamano = 35
      sContenido = Left(gdl_Funcion.aTexto(porstClone!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstClone!numerodir), nTamano)
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: libre
      sContenido = "": nTamano = 168
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    ElseIf nSecuencia = 3 Then
      ' 1: codigo registro - constante
      sContenido = "2130": nTamano = 4
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 2: tipo documento empresa - constante
      sContenido = "R": nTamano = 1
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 3: documento empresa
      sContenido = ps_RucEmpresa: nTamano = 12
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 4: distrito
      nTamano = 35
      sContenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, gdl_Funcion.aTexto(porstClone!ubigeodir_emp), "UB"), nTamano)
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 5: provincia
      nTamano = 25
      sContenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Act, Left(gdl_Funcion.aTexto(porstClone!ubigeodir_emp), 4), "UB"), nTamano)
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 6: departamento
      nTamano = 25
      sContenido = Left(gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Ina, Left(gdl_Funcion.aTexto(porstClone!ubigeodir_emp), 2), "UB"), nTamano)
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
      ' 7: libre
      sContenido = "": nTamano = 153
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, nTamano, sCaracter)
    End If
  Else                              ' C.T.S.
    If nFormato = "9" Then          ' Formato bbva - cash
      sContenido = IIf(fMenu.ribMoneda(0).Value, "MN", "ME")
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR("2", 1, sCaracter)                 ' Tipo de registro constante
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR("", 14, sCaracter)                 ' Espacios en blanco
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(sContenido, 2, sCaracter)          ' Moneda de cuenta de transferencia
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR("", 6, sCaracter)                  ' Espacios en blanco
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR(ps_RucEmpresa, 11, sCaracter)      ' Numero de ruc de la empresa
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR("", 11, "0")                       ' Codigo de trabajador
      InicioArchivo_BBVACash = InicioArchivo_BBVACash & gdl_Funcion.PadR("", 6, sCaracter)                  ' Espacios en blanco
    End If
  End If
  Set porstClone = Nothing
  
End Function
Private Sub cmdAction_Click(Index As Integer)
  Dim s_Archivo As String, sDireccion As String
  Dim l_ExistRecord As Boolean
  
  ' Verifico si existen Registros
  l_ExistRecord = (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0)
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0  ' Genera carta de transferencia
    If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If txtBanco = "" Then Beep: MsgBox "Debe Ingresar Entidad Bancaria de transferencia", vbExclamation: txtBanco.SetFocus: Exit Sub
    If lblHelp(1) = "" Or lblHelp(1) = "???" Then Beep: MsgBox "Entidad bancaria no existe; verifique", vbExclamation: txtBanco.SetFocus: Exit Sub
    Me.Tag = IIf(l_ExistRecord, s_MdoData_Ins, s_MdoData_Upd)
    fGeneraCartaBanco.Show vbModal
   Case 1  ' Elimina carta de transferencia
    If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If txtBanco = "" Then Beep: MsgBox "Debe Ingresar Entidad Bancaria de transferencia", vbExclamation: txtBanco.SetFocus: Exit Sub
    If lblHelp(1) = "" Or lblHelp(1) = "???" Then Beep: MsgBox "Entidad bancaria no existe; verifique", vbExclamation: txtBanco.SetFocus: Exit Sub
    Beep
    If MsgBox("¿ Estás Seguro de Eliminar la " & tdbRegistro.Caption & " '" & Trim$(dcaRegistro.Recordset!desmotivo) & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
      EliminaCarta dcaRegistro.Recordset!nrocarta
    End If
   Case 2, 3  ' Ordena registro ascendentemente o descendentemente
    ' Verifico que existan registros
    If l_ExistRecord Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index, " ASC", " DESC")
   Case 4 ' Busqueda de registro
    ' Verifico que existan registros
    If l_ExistRecord Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
    Set go_tdbBusqueda = tdbRegistro
    Set go_dcaBusqueda = dcaRegistro
    gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
    fBusqueda.Show vbModal
   Case 5     ' Genero el archivo de transferencia
    ' Verifico que existan registros
    If l_ExistRecord Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
    If txtPeriodo = "" Then Beep: MsgBox "Debe Ingresar el Codigo del Periodo de Pago", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If lblHelp(0) = "" Or lblHelp(0) = "???" Then Beep: MsgBox "Periodo de Pago no existe; verifique", vbExclamation: txtPeriodo.SetFocus: Exit Sub
    If txtBanco = "" Then Beep: MsgBox "Debe Ingresar Entidad Bancaria de transferencia", vbExclamation: txtBanco.SetFocus: Exit Sub
    If lblHelp(1) = "" Or lblHelp(1) = "???" Then Beep: MsgBox "Entidad bancaria no existe; verifique", vbExclamation: txtBanco.SetFocus: Exit Sub
    
    ' Verifico que existan parametros de planilla
    s_Sql = "SELECT codvia, direccionvia, numerodir, codzona, direccionzona, ubigeodir "
    s_Sql = s_Sql & "FROM plcfgempresa "
    s_Sql = s_Sql & "WHERE pdoano='" & ps_Anyo & "' "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If (porstRecordset.BOF And porstRecordset.EOF) Then Beep: MsgBox "Debe configurar los parametros de la empresa", vbCritical: Exit Sub
    s_Direccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
    s_Direccion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, s_Direccion, "UB")
    s_Direccion = gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " - " & s_Direccion
    porstRecordset.Close
    
    s_Archivo = Trim(txtPeriodo.Text) & Trim(txtBanco.Text) & ".txt"
    On Error GoTo CancelaDialogo
    fMenu.cdlDialogo.DialogTitle = "Grabar Archivo Como"
    fMenu.cdlDialogo.CancelError = True
    fMenu.cdlDialogo.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
    fMenu.cdlDialogo.FileName = s_Archivo
    fMenu.cdlDialogo.DefaultExt = ".txt"
    fMenu.cdlDialogo.Filter = "Archivos de texto(*.txt)|*.txt|Todos los archivos(*.*)|*.*"
    fMenu.cdlDialogo.ShowSave
  
CancelaDialogo:
    ' veriofico si existe error y desactivo
    If Not Err.Number = 0 Then MsgBox Error(Err.Number): Exit Sub
    On Error GoTo 0
    
    ChDir App.path
    If MsgBox("¿ Estás Seguro de Generar Archivo de Transferencia Bancaria? ", vbQuestion + vbYesNo) = vbYes Then
      s_Archivo = fMenu.cdlDialogo.FileName
      ExportaBancos s_Archivo, dcaRegistro.Recordset("nrocarta"), "G"
      MsgBox "Proceso de Exportación Finalizo con Exito", vbInformation
    End If
    ChDrive Left$(App.path, 1)
    ChDir App.path
   Case 6, 7  ' Opciones de impresión
    ' Verifico que existan registros
    If l_ExistRecord Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
        
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "TRANSFERENCIA BANCARIA"
    gdl_Procedure.ps_ReportName = "rpttransbanco"
    ReDim aElemento(3, 6): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = "": aElemento(0, 3) = "": aElemento(0, 4) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    aElemento(1, 3) = "": aElemento(1, 4) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & "TRANSFERENCIA BANCARIA" & ";true"
    aElemento(2, 2) = "NumeroCarta;" & dcaRegistro.Recordset("nrocarta") & ";true"
    aElemento(2, 3) = "Planilla;" & ps_DesClsPlanilla & ";true"
    aElemento(2, 4) = "Periodo;" & Trim(txtPeriodo) & " - " & lblHelp(0) & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
  
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "codpsn varchar(11) Not Null, apepaterno varchar(25) Null, apematerno varchar(25) Null, nombres varchar(25) Null, "
    s_Sql = s_Sql & "codcpc varchar(4) Null, descpc varchar(40) Null, desmotivo varchar(25) Null, "
    s_Sql = s_Sql & "codmon char(1) Null, codbco char(3) Null, desbco varchar(40) Null, cuentapago varchar(20) Null,"
    s_Sql = s_Sql & "coddci char(2) Null, numdociden varchar(15) Null, codentidad varchar(6) Null, "
    s_Sql = s_Sql & "cuentabco varchar(20) Null, importemn decimal(18,2) Null Default '0.00', importeme decimal(18,2) Null Default '0.00', "
    s_Sql = s_Sql & "PRIMARY KEY (codpsn)) "
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    ' Genera la información del reporte
    ExportaBancos "tmp" & gdl_Procedure.ps_ReportName, dcaRegistro.Recordset("nrocarta"), "R"
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY codpsn"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
  End Select

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Periodo de Pago
    If ribAnalisis(0).Value Then
      tdbHelp.Columns(0).DataField = "codpdo": tdbHelp.Columns(1).DataField = "despdo"
      tdbHelp.Caption = "Periodos de Pago"
      s_Sql = gdl_Funcion.HelpTablas("ped", "codpdo", s_Estado_Ina & ps_ClsPlanilla & ps_Anyo, "")
    Else
      tdbHelp.Columns(0).DataField = "pdocts": tdbHelp.Columns(1).DataField = "descricts"
      tdbHelp.Caption = "Periodo de CTS"
      s_Sql = gdl_Funcion.HelpTablas("cxe", "pdocts", s_Estado_Blq & ps_ClsPlanilla, "")
    End If
   Case 1     ' Entidad bancaria
    tdbHelp.Columns(0).DataField = "codbco": tdbHelp.Columns(1).DataField = "desbco"
    tdbHelp.Caption = "Entidad Bancaria"
    s_Sql = gdl_Funcion.HelpTablas("bco", "codbco", "", "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = panToolBar(1).Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2))
  tdbHelp.Left = 1480
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub Form_Load()

  ' Establece posición del formulario
  Me.Height = 6500: Me.Width = 6230
  Me.Left = 520: Me.Top = 300
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = "Transferencia de Depositos a Bancos"
  s_TitleTable = "Carta de Transferencia"
  
  ReDim aElemento(3, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Carta", "Detalle", "Con")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "nrocarta", "desmotivo", "codcpc")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1250, 2956.03, 450)
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
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  tdbRegistro.AllowColMove = False
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(8, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "registro": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "ajustado", "borrar", "ordascen", "orddesce", "busqueda", "genarchi", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Generación Carta de Transferencia", "Elimina Carta de Transferencia", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Generación de Archivo", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 1
    ribAnalisis(n_Index).PictureUp = LoadPicture()
    ribAnalisis(n_Index).ToolTipText = Choose(n_Index + 1, "Planilla de Remuneraciones", "Liquidación de C.T.S.")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "remunera", "liquicts") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribAnalisis(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ribAnalisis(0).Value = True
  
  
 '[ Configuración el control de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codpdo", "despdo")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 834.7402, 3365.071)
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Entidad Pensiones", 2
  ']
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
  
  ' Bloqueo la seleccion de ejercicio
  fMenu.cmbejercicio.Enabled = False
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If porstHelp.State = adStateOpen Then porstHelp.Close
  Set porstHelp = Nothing
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
Private Sub ribAnalisis_Click(Index As Integer, Value As Integer)
  gdl_Procedure.EditText "AT", txtPeriodo, "", s_MdoData_Ins, False, Choose(Index + 1, 8, 6)
  gdl_Procedure.EditText "AT", txtBanco, "", s_MdoData_Ins, False, 3
  lblHelp(0) = ""
  txtBanco_LostFocus
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
   Case 1       ' Entidad de banco
    txtBanco = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtBanco.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Periodo de Pago
    If ribAnalisis(0).Value Then
      s_Sql = gdl_Funcion.HelpTablas("ped", tdbHelp.Columns(ColIndex).DataField, s_Estado_Ina & ps_ClsPlanilla & ps_Anyo, "")
    Else
      s_Sql = gdl_Funcion.HelpTablas("cxe", tdbHelp.Columns(ColIndex).DataField, s_Estado_Blq & ps_ClsPlanilla, "")
    End If
   Case 1     ' Entidad de banco
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
Private Sub tdbRegistro_DblClick()
  cmdAction_Click 0
End Sub
Private Sub tdbRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaRegistro, tdbRegistro, " " & s_TitleTable
End Sub
Private Sub tdbRegistro_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then cmdAction_Click 0
End Sub
Private Sub txtBanco_GotFocus()
  gdl_Procedure.MarcaGet txtBanco
End Sub
Private Sub txtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 1
End Sub
Private Sub txtBanco_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtBanco_LostFocus()
  lblHelp(1) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txtBanco, "EB")
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
End Sub
Private Sub txtPeriodo_GotFocus()
  gdl_Procedure.MarcaGet txtPeriodo
End Sub
Private Sub txtPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    EnviarTecla vbKeyTab: KeyAscii = 0
  End If
End Sub
Private Sub txtPeriodo_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtPeriodo, IIf(ribAnalisis(0).Value, "PR", "EC"))
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
End Sub
