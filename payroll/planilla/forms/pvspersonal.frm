VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fPvsPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "pvspersonal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   8490
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   5085
      Left            =   45
      TabIndex        =   17
      Top             =   555
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8969
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
      Top             =   5670
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
      Height          =   5460
      Index           =   0
      Left            =   7695
      TabIndex        =   0
      Top             =   555
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   9631
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
         Index           =   10
         Left            =   255
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1125
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
         Picture         =   "pvspersonal.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   9
         Left            =   255
         TabIndex        =   10
         Tag             =   "1"
         Top             =   705
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
         Picture         =   "pvspersonal.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Tag             =   "0"
         Top             =   1455
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
         Picture         =   "pvspersonal.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1875
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
         Picture         =   "pvspersonal.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   4275
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
         Picture         =   "pvspersonal.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   4695
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
         Picture         =   "pvspersonal.frx":0098
      End
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
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Tag             =   "0"
         Top             =   705
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
         Picture         =   "pvspersonal.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   4
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
         Picture         =   "pvspersonal.frx":00D0
      End
      Begin Threed.SSPanel panTool 
         Height          =   255
         Index           =   1
         Left            =   15
         TabIndex        =   16
         Top             =   285
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Proceso"
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
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   2865
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
         Picture         =   "pvspersonal.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   3285
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
         Picture         =   "pvspersonal.frx":0108
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   3705
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
         Picture         =   "pvspersonal.frx":0124
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   11
         Left            =   255
         TabIndex        =   12
         Tag             =   "1"
         Top             =   1545
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
         Picture         =   "pvspersonal.frx":0140
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   12
         Left            =   255
         TabIndex        =   13
         Tag             =   "1"
         Top             =   1950
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
         Picture         =   "pvspersonal.frx":015C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   13
         Left            =   255
         TabIndex        =   14
         Tag             =   "1"
         Top             =   2520
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
         Picture         =   "pvspersonal.frx":0178
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8490
      _Version        =   65536
      _ExtentX        =   14975
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7800
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "pvspersonal.frx":0194
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkCesados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cesados"
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cboPeriodo 
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
         ItemData        =   "pvspersonal.frx":0816
         Left            =   3240
         List            =   "pvspersonal.frx":0818
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   90
         Width           =   2625
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   7275
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
         PictureUp       =   "pvspersonal.frx":081A
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   6870
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
         PictureUp       =   "pvspersonal.frx":0836
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   7680
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
         PictureUp       =   "pvspersonal.frx":0852
      End
      Begin Threed.SSCommand cmdGenera 
         Height          =   360
         Left            =   960
         TabIndex        =   19
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         ForeColor       =   -2147483638
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
         Picture         =   "pvspersonal.frx":086E
      End
      Begin MSComctlLib.Toolbar toolbarcanc 
         Height          =   360
         Left            =   120
         TabIndex        =   26
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cancelar Provisiones"
               ImageIndex      =   1
               Style           =   5
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label lblDato 
         BackStyle       =   0  'Transparent
         Caption         =   "Mes :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   24
         Top             =   120
         Width           =   420
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         Height          =   390
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   3540
      End
   End
End
Attribute VB_Name = "fPvsPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private as_SelRegistro(2)                               ' Array de inicio y fin de seleccion de registro
Private s_OptRegistro As String                         ' Instancia del formulario activo
Dim cnn As ADODB.Connection
'[
Private Sub AnalisisCTS(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim nRegistro As Long, nRegistros As Long
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo sub periodo
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW res.codcls, res.codpsn, res.pdocts, res.subcts, cgo.descgo, "
  s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS nombrepsn, "
  s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, psn.codcco, cco.detcco, res.codmon, res.pdoano, res.pdomes, "
  s_Sql = s_Sql & "mov.fechaini, mov.fechafin, mov.numeroanos, mov.numeromeses, mov.numerodias, mov.fechacan "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN plctsmovimiento mov ON res.codcls=mov.codcls AND res.pdocts=mov.pdocts AND res.subcts=mov.subcts AND res.codpsn=mov.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON psn.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
  s_Sql = s_Sql & "AND res.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "GROUP BY res.codpsn "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la tabla temporal del certificado
  s_Sql = "DROP TABLE IF EXISTS tmpimporte"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpimporte ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, "
  s_Sql = s_Sql & "remunermn decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "remunerme decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "importemn decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "importeme decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "provisimn decimal(18, 2) NOT Null Default 0, "
  s_Sql = s_Sql & "provisime decimal(18, 2) NOT Null Default 0)"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  
  ' Inserto las remuneraciones basicas
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.importe_mn AS remunermn, res.importe_me AS remunerme, "
  s_Sql = s_Sql & "0.00 AS importemn, 0.00 AS importeme, 0.00 AS provisimn, 0.00 AS provisime "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remubasicacts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto las remuneraciones promedio
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.importe_mn AS remunermn, res.importe_me AS remunerme, "
  s_Sql = s_Sql & "0.00 AS importemn, 0.00 AS importeme, 0.00 AS provisimn, 0.00 AS provisime "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remupromects "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneracion de gratificacion
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, res.importe_mn AS remunermn, res.importe_me AS remunerme, "
  s_Sql = s_Sql & "0.00 AS importemn, 0.00 AS importeme, 0.00 AS provisimn, 0.00 AS provisime "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remugraticts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "AND res.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la remuneraciones de cts
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS remunermn, 0.00 AS remunerme, res.importe_mn AS importemn, "
  s_Sql = s_Sql & "res.importe_me AS importeme, 0.00 AS provisimn, 0.00 AS provisime "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remutotalcts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Inserto la provisión de cts
  s_Sql = "INSERT INTO tmpimporte "
  s_Sql = s_Sql & "SELECT res.codpsn, 0.00 AS remunermn, 0.00 AS remunerme, 0.00 AS importemn, "
  s_Sql = s_Sql & "0.00 AS importeme, res.importe_mn AS provisimn, res.importe_me AS provisime "
  s_Sql = s_Sql & "FROM plctsresultado res "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn AND res.pdocts=psn.pdocts AND res.subcts=psn.subcts "
  s_Sql = s_Sql & "INNER JOIN plparametroafp cfg ON res.pdoano=cfg.pdoano AND res.codcpc=cfg.remunepvscts "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.pdoano='" & ps_Anyo & "' "
  s_Sql = s_Sql & "ORDER BY res.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  ']
  
  ' Genero la información de provisiones
  s_Sql = "INSERT INTO " & s_Archivo & " "
  s_Sql = s_Sql & "SELECT tmp.codpsn, tmp.nombrepsn, tmp.codcco, tmp.detcco, tmp.pdoano, tmp.pdomes, "
  s_Sql = s_Sql & "tmp.descgo, tmp.fecingreso, tmp.fecbaja, tmp.fechaini, tmp.fechafin, "
  s_Sql = s_Sql & "tmp.numeromeses, tmp.numerodias, tmp.codmon, "
  s_Sql = s_Sql & "SUM(IFNULL(imp.remunermn, 0)) AS remunera_mn, "
  s_Sql = s_Sql & "SUM(IFNULL(imp.remunerme, 0)) AS remunera_me, "
  s_Sql = s_Sql & "SUM(IFNULL(imp.importemn, 0)) AS imporpvsacu_mn, "
  s_Sql = s_Sql & "SUM(IFNULL(imp.importeme, 0)) AS imporpvsacu_me, "
  s_Sql = s_Sql & "SUM(IFNULL(imp.provisimn, 0)) AS importepvs_mn, "
  s_Sql = s_Sql & "SUM(IFNULL(imp.provisime, 0)) AS importepvs_me, "
  s_Sql = s_Sql & "tmp.fechacan "
  s_Sql = s_Sql & "FROM tmpimporte imp "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin tmp ON imp.codpsn=tmp.codpsn "
  s_Sql = s_Sql & "GROUP BY imp.codpsn "
  s_Sql = s_Sql & "ORDER BY tmp.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub AnalisisVacaciones(ByVal s_Archivo As String, s_Proceso As String, s_FechaHora As String)
  Dim nDiaDevengado As Double, nDiaTrunco As Double, nDiaFinal As Double
  Dim nRegistro As Long, nRegistros As Long
  
  ' Cambio el Mensaje y Muestro la Barra
  fMenu.panPercent.Visible = True
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  '[ Genero la tabla temporal de selección ultimo sub periodo
  s_Sql = "DROP TABLE IF EXISTS tmpmesfin"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  
  s_Sql = "CREATE TEMPORARY TABLE tmpmesfin "
  s_Sql = s_Sql & "SELECT DISTINCTROW vac.codcls, vac.codpsn, vac.pdoano, vac.pdomes, MAX(vac.pdopvs) AS pdopvs, "
  s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS nombrepsn, "
  s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, psn.codcco, cco.detcco "
  s_Sql = s_Sql & "FROM plpvsvacaciondet vac "
  s_Sql = s_Sql & "INNER JOIN plpvsperiodovac pdo ON vac.codcls=pdo.codcls AND vac.codpvs=pdo.codpvs "
  s_Sql = s_Sql & "INNER JOIN plpvsvacacion sub ON vac.codcls=sub.codcls AND vac.codpvs=sub.codpvs AND vac.codpsn=sub.codpsn AND vac.pdopvs=sub.pdopvs "
  s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON vac.codcls=psn.codcls AND vac.codpsn=psn.codpsn "
  s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON psn.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE vac.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND vac.pdoano='" & ps_Anyo & "' "
  If cboPeriodo.Text <> "" Then
    s_Sql = s_Sql & "AND vac.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
  End If
  s_Sql = s_Sql & "AND vac.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "GROUP BY vac.codpsn "
  s_Sql = s_Sql & "ORDER BY vac.pdoano, vac.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ']
  
  ' Genero la información de provisiones
  s_Sql = "INSERT INTO " & s_Archivo & " "
  s_Sql = s_Sql & "SELECT vac.codpsn, psn.nombrepsn, "
  s_Sql = s_Sql & "psn.codcco, psn.detcco, vac.pdoano, vac.pdomes, vac.pdopvs, pdo.descripvs, sub.fechaini AS subfechaini, "
  s_Sql = s_Sql & "sub.fechafin AS subfechafin, psn.fecingreso, psn.fecbaja, vac.fechaini, vac.fechafin, "
  s_Sql = s_Sql & "vac.numerodias, 0 AS diadevengado, 0 AS diaperiodo, 0 AS diafisico, vac.codmon, "
  s_Sql = s_Sql & "vac.remunera_mn, vac.remunera_me, vac.imporpvsacu_mn, "
  s_Sql = s_Sql & "vac.imporpvsacu_me, vac.importepvs_mn, vac.importepvs_me, vac.fechacan "
  s_Sql = s_Sql & "FROM plpvsvacaciondet vac "
  s_Sql = s_Sql & "INNER JOIN plpvsperiodovac pdo ON vac.codcls=pdo.codcls AND vac.codpvs=pdo.codpvs "
  s_Sql = s_Sql & "INNER JOIN plpvsvacacion sub ON vac.codcls=sub.codcls AND vac.codpvs=sub.codpvs AND vac.codpsn=sub.codpsn AND vac.pdopvs=sub.pdopvs "
  s_Sql = s_Sql & "INNER JOIN tmpmesfin psn ON vac.codcls=psn.codcls AND vac.codpsn=psn.codpsn AND vac.pdopvs=psn.pdopvs "
  s_Sql = s_Sql & "WHERE vac.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND vac.pdoano='" & ps_Anyo & "' "
  If cboPeriodo.Text <> "" Then
    s_Sql = s_Sql & "AND vac.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
  End If
  s_Sql = s_Sql & "ORDER BY vac.pdoano, vac.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Selecciones los dias fisicos de vacaciones
  s_Sql = "DROP TABLE IF EXISTS pvsvacacion"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE pvsvacacion ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, codpdo varchar(8) NOT Null, "
  s_Sql = s_Sql & "pdoano char(4) Null, pdomes char(2) Null, "
  s_Sql = s_Sql & "pdovaca varchar(8) Null, fechainivaca date Null, "
  s_Sql = s_Sql & "fechafinvaca date Null, diasvaca int(4) DEFAULT 0)"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ' Inserto la informacion de dias fisicos
  s_Sql = "INSERT INTO pvsvacacion "
  s_Sql = s_Sql & "SELECT asi.codpsn, asi.codpdo, "
  s_Sql = s_Sql & "pdo.anopdo AS pdoano, pdo.mespdo AS pdomes, pdovaca1 AS pdovaca, "
  s_Sql = s_Sql & "fechainivaca1 AS fechainivaca, fechafinvaca1 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca1, fechainivaca1) + 1, 0) AS diasvaca "
  s_Sql = s_Sql & "FROM plasistencia asi "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & ps_Anyo & Format(cboPeriodo.ListIndex, "00") & "' "
  s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND NOT ISNULL(fechainivaca1) "
  s_Sql = s_Sql & "AND NOT ISNULL(fechafinvaca1) "
  s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "UNION "
  s_Sql = s_Sql & "SELECT asi.codpsn, asi.codpdo, "
  s_Sql = s_Sql & "pdo.anopdo AS pdoano, pdo.mespdo AS pdomes, pdovaca2 AS pdovaca, "
  s_Sql = s_Sql & "fechainivaca2 AS fechainivaca, fechafinvaca2 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca2, fechainivaca2) + 1, 0) AS diasvaca "
  s_Sql = s_Sql & "FROM plasistencia asi "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & ps_Anyo & Format(cboPeriodo.ListIndex, "00") & "' "
  s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND NOT ISNULL(fechainivaca2) "
  s_Sql = s_Sql & "AND NOT ISNULL(fechafinvaca2) "
  s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "UNION "
  s_Sql = s_Sql & "SELECT asi.codpsn, asi.codpdo, "
  s_Sql = s_Sql & "pdo.anopdo AS pdoano, pdo.mespdo AS pdomes, pdovaca3 AS pdovaca, "
  s_Sql = s_Sql & "fechainivaca3 AS fechainivaca, fechafinvaca3 AS fechafinvaca, IFNULL(DateDiff(fechafinvaca3, fechainivaca3) + 1, 0) AS diasvaca "
  s_Sql = s_Sql & "FROM plasistencia asi "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON asi.codcls=psn.codcls AND asi.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plperiodo pdo ON asi.codcls=pdo.codcls AND asi.codpdo=pdo.codpdo AND CONCAT(pdo.anopdo, pdo.mespdo)<='" & ps_Anyo & Format(cboPeriodo.ListIndex, "00") & "' "
  s_Sql = s_Sql & "WHERE asi.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND NOT ISNULL(fechainivaca3) "
  s_Sql = s_Sql & "AND NOT ISNULL(fechafinvaca3) "
  s_Sql = s_Sql & "AND asi.codpsn IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "ORDER BY codpsn, pdovaca, pdoano, pdomes, codpdo"
  'If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Informacion agrupada de vacaciones
  s_Sql = "DROP TABLE IF EXISTS tmpvacacion"
  If Not gdl_Conexion.Execucion(s_Sql, Elimina) Then GoTo Finalizar
  s_Sql = "CREATE TEMPORARY TABLE tmpvacacion ( "
  s_Sql = s_Sql & "codpsn varchar(11) NOT Null, diasvaca int(4) DEFAULT 0)"
  If Not gdl_Conexion.Execucion(s_Sql, Seleccion) Then GoTo Finalizar
  ' Informacion final de vacaciones fisicas
  s_Sql = "INSERT INTO tmpvacacion "
  s_Sql = s_Sql & "SELECT vac.codpsn, SUM(vac.diasvaca) AS diasvaca "
  s_Sql = s_Sql & "FROM pvsvacacion vac "
  s_Sql = s_Sql & "GROUP BY vac.codpsn "
  s_Sql = s_Sql & "ORDER BY codpsn "
  If Not gdl_Conexion.Execucion(s_Sql, Inserta) Then GoTo Finalizar
  
  ' Actualizo los días físicos de vacaciones
  s_Sql = "UPDATE " & s_Archivo & " pvs, tmpvacacion vac "
  s_Sql = s_Sql & "SET pvs.diafisico=(vac.diasvaca*12) "
  s_Sql = s_Sql & "WHERE vac.codpsn=pvs.codpsn"
  If Not gdl_Conexion.Execucion(s_Sql, Modifica) Then GoTo Finalizar
  
  ' Recupero la informacion de las vacaciones
  s_Sql = "SELECT codpsn, nombrepsn, codcco, detcco, pdoano, pdomes, pdopvs, descripvs, "
  s_Sql = s_Sql & "subfechaini, subfechafin, fecingreso, fecbaja, fechaini, fechafin, "
  s_Sql = s_Sql & "numerodias, diadevengado, diaperiodo, diafisico, codmon, remunera_mn, remunera_me, "
  s_Sql = s_Sql & "imporpvsacu_mn, imporpvsacu_me, importepvs_mn, importepvs_me, fechacan "
  s_Sql = s_Sql & "FROM " & s_Archivo & " tmp "
  s_Sql = s_Sql & "ORDER BY codpsn"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  If Not (porstRecordset.BOF And porstRecordset.EOF) Then
    nRegistros = porstRecordset.RecordCount: nRegistro = 0
    ' Arreglos de grabación
    a_Campos = Array("codpsn", "diadevengado", "diaperiodo")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Numero, TipoDato.Numero)
    a_Where = Array("codpsn")
    While Not porstRecordset.EOF
      nDiaDevengado = 0
      nDiaTrunco = gdl_Funcion.NumeroDias360(porstRecordset!fechafin, porstRecordset!subfechaini, porstRecordset!fechafin)
      nDiaFinal = porstRecordset!numerodias - porstRecordset!diafisico
      If nDiaTrunco >= nDiaFinal Then
        nDiaTrunco = nDiaFinal
        nDiaTrunco = IIf(nDiaFinal > 0 And (nDiaFinal <> nDiaTrunco), nDiaTrunco - nDiaFinal, nDiaFinal)
      Else
        nDiaDevengado = nDiaFinal - nDiaTrunco
      End If
      nDiaTrunco = Round(nDiaTrunco / 12, 2)
      nDiaDevengado = Round(nDiaDevengado / 12, 2)
      ' Actualizo la informacion del reporte
      a_Valores = Array(porstRecordset!codpsn, nDiaDevengado, nDiaTrunco)
        
      gdl_Conexion.IniciaTransaccion    ' Inicia transacción
      ' Realizo la actualización de los registros
      If Not Records_Upd(s_Archivo, a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
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
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub ImportaVacaciones()
  Dim sArchivo As String
  Dim Fila_Actual As Integer, Fila_Ultima As Integer, Columna_Actual As Integer
  Dim CodigoPersona As String
  Dim Dato As Variant
  Dim Obj_Hoja As Object
  Dim objExcel As Excel.Application
  Dim xLibro As Excel.Workbook
  Dim n As Integer
  Dim rsdelete As New Recordset
  Dim rsinsert As New Recordset
  Dim s_sql_delete As String, s_sql_insert As String
  
  On Error GoTo Error
  
  If cboPeriodo.Text = "" Then
    MsgBox "Falta, Seleccionar Mes"
    Exit Sub
  Else
    fMenu.cdlDialogo.CancelError = False
    fMenu.cdlDialogo.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist
    fMenu.cdlDialogo.Filter = "Archivos de Excel(*.xls; *.xlsx)|*.xls;*.xlsx"
    fMenu.cdlDialogo.FilterIndex = 1
    fMenu.cdlDialogo.DialogTitle = "Seleccionar Archivo"
    fMenu.cdlDialogo.InitDir = ps_PathSystem
    fMenu.cdlDialogo.FileName = ""
    fMenu.cdlDialogo.ShowOpen
    ' Capturo archivo seleccionado
    sArchivo = fMenu.cdlDialogo.FileName
    
    If sArchivo = "" Then Exit Sub
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(sArchivo)
    'objExcel.Visible = True
    If Val(objExcel.Application.VERSION) >= 8 Then
      Set Obj_Hoja = objExcel.ActiveSheet
    Else
      Set Obj_Hoja = objExcel
    End If
  
    Fila_Ultima = 1
    Do While True
      If IsEmpty(Obj_Hoja.Cells(Fila_Ultima, 1)) Then Exit Do
      Fila_Ultima = Fila_Ultima + 1
    Loop
  
    For n = 1 To Fila_Ultima - 1
      'Aqui mi filtro por meses
      CodigoPersona = Obj_Hoja.Cells(n + 1, 3)
      s_sql_delete = "delete from plpvsvacaciondet where codcls='" & Obj_Hoja.Cells(n + 1, 1) & "' and codpvs='" & Obj_Hoja.Cells(n + 1, 2) & "' and codpsn='" & Obj_Hoja.Cells(n + 1, 3) & "' and pdopvs='" & Obj_Hoja.Cells(n + 1, 4) & "' and pdoano='" & Obj_Hoja.Cells(n + 1, 5) & "' and pdomes='" & Obj_Hoja.Cells(n + 1, 6) & "'"
      If Left(cboPeriodo.Text, 2) = Obj_Hoja.Cells(n + 1, 6) And ps_Anyo = Obj_Hoja.Cells(n + 1, 5) Then
        rsdelete.Open s_sql_delete, cnn, adOpenStatic, adLockOptimistic
      End If
    Next n
      
    For n = 1 To Fila_Ultima - 1
      'Aqui mi filtro por meses
      CodigoPersona = Obj_Hoja.Cells(n + 1, 3)
      s_sql_insert = " insert into plpvsvacaciondet (codcls,codpvs,codpsn,pdopvs,pdoano,pdomes,fechaini,fechafin,numerodias,codmon,remunera_mn,remunera_me,imporpvsacu_mn,imporpvsacu_me,importepvs_mn,importepvs_me,codcta_debmn,codcta_habmn,codcta_debme,codcta_habme,usrcre,fyhcre)"
      s_sql_insert = s_sql_insert & " values('" & Obj_Hoja.Cells(n + 1, 1) & "','" & Obj_Hoja.Cells(n + 1, 2) & "','" & Obj_Hoja.Cells(n + 1, 3) & "','" & Obj_Hoja.Cells(n + 1, 4) & "','" & Obj_Hoja.Cells(n + 1, 5) & "','" & Obj_Hoja.Cells(n + 1, 6) & "','" & Format(Obj_Hoja.Cells(n + 1, 7), s_FmtFechMysql_0) & "',"
      s_sql_insert = s_sql_insert & "'" & Format(Obj_Hoja.Cells(n + 1, 8), s_FmtFechMysql_0) & "'," & Obj_Hoja.Cells(n + 1, 9) & ",'" & Obj_Hoja.Cells(n + 1, 10) & "'," & Obj_Hoja.Cells(n + 1, 11) & "," & Obj_Hoja.Cells(n + 1, 12) & "," & Obj_Hoja.Cells(n + 1, 13) & "," & Obj_Hoja.Cells(n + 1, 14) & ","
      s_sql_insert = s_sql_insert & "" & Obj_Hoja.Cells(n + 1, 15) & "," & Obj_Hoja.Cells(n + 1, 16) & ",'" & Obj_Hoja.Cells(n + 1, 17) & "','" & Obj_Hoja.Cells(n + 1, 18) & "','" & Obj_Hoja.Cells(n + 1, 19) & "','" & Obj_Hoja.Cells(n + 1, 20) & "',"
      s_sql_insert = s_sql_insert & "'admin','" & Format(Now, s_FmtFechMysql_0) & "')"
      
      If Left(cboPeriodo.Text, 2) = Obj_Hoja.Cells(n + 1, 6) And ps_Anyo = Obj_Hoja.Cells(n + 1, 5) Then
        rsinsert.Open s_sql_insert, cnn, adOpenStatic, adLockOptimistic
      End If
    Next n
  
    objExcel.ActiveWorkbook.Close False
    objExcel.Quit
    Set Obj_Hoja = Nothing
    Set objExcel = Nothing
  End If
  Exit Sub
Error:   MsgBox "Revisar Excel, Codigo :" & CodigoPersona: objExcel.ActiveWorkbook.Close False: objExcel.Quit: Set Obj_Hoja = Nothing: Set objExcel = Nothing: Exit Sub

End Sub
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, codpsn, CONCAT(IFNULL(apepaterno, ''), ' ', IFNULL(apematerno, ''), ', ', IFNULL(nombres, '')) AS nombrepsn, "
  s_Sql = s_Sql & "fecnacimiento, ubigeonac, naciextrapsn, sexopsn, "
  s_Sql = s_Sql & "coddci, numdociden, numdocmil, fecingreso, "
  s_Sql = s_Sql & "codcco, codafp, numeroafp, ctsdeposito, pagodolar, "
  s_Sql = s_Sql & "remintegralgrati, remintegralvaca, remimprecisa, remuneta, netocpc, variacpc, imporemuneto, "
  s_Sql = s_Sql & "fecbaja,estadopsn "
  s_Sql = s_Sql & "FROM plpersonal "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  If Not ribParametro(0).Value Then
    s_Sql = s_Sql & " AND estadopsn" & IIf(ribParametro(1).Value, "<>'I'", "='I'")
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
']
Private Sub cmdAction_Click(Index As Integer)
  Dim s_FechaHora As String, s_OldMessage As String
  Dim s_Reporte As String, s_Periodo As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Me.Tag = ""
  Select Case Index
   Case 0 ' Visualizar o analizar, eliminar registro
    If Not (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Then
      Me.Tag = IIf(Index = 0, s_MdoData_Vis, s_MdoData_Del)
      If s_OptRegistro = "pvsvacacio" Then
        fPvsVacacion.Show
      ElseIf s_OptRegistro = "pvsgratifi" Then
        fPvsGratificacion.Show
      ElseIf s_OptRegistro = "pvscoxtise" Then
        fCtsMovimiento.Show
      End If
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
            
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Información ..."
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    s_Reporte = IIf(cboPeriodo.Text = "", "cst", "rpt")
    s_Periodo = IIf(cboPeriodo.Text = "", "Ejercicio", Mid(cboPeriodo.Text, 6)) & " - " & ps_Anyo
    gdl_Procedure.ps_ReportTitle = UCase(IIf(s_OptRegistro = "pvsvacacio", "Provisión de Vacaciones", IIf(s_OptRegistro = "pvsgratifi", "Provisión de Gratificaciones", "Provisión de CTS")))
    gdl_Procedure.ps_ReportName = IIf(s_OptRegistro = "pvsvacacio", s_Reporte & "pvsvacacion", IIf(s_OptRegistro = "pvsgratifi", s_Reporte & "pvsgratifi", s_Reporte & "provisicts"))
    
    ReDim aElemento(2, 3): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbRegistro.Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "": aElemento(1, 2) = ""
    ' Parametros de campos del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & "; true"
    aElemento(2, 1) = "TituloReporte;" & gdl_Procedure.ps_ReportTitle & ";true"
    aElemento(2, 2) = "Periodo;" & s_Periodo & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    If s_OptRegistro = "pvsvacacio" Then
      If cboPeriodo.Text <> "" Then
        s_Sql = s_Sql & "codpsn varchar(11) NOT Null, nombrepsn varchar(75) Null, codcco varchar(10) Null, "
        s_Sql = s_Sql & "detcco varchar(40) Null, pdoano char(4) Null, pdomes char(2) Null, "
        s_Sql = s_Sql & "pdopvs varchar(8) Null, descripvs varchar(45) Null, subfechaini date Null, "
        s_Sql = s_Sql & "subfechafin date Null, fecingreso date Null, fecbaja date Null, "
        s_Sql = s_Sql & "fechaini date Null, fechafin date Null, numerodias decimal(7,2) DEFAULT 0, "
        s_Sql = s_Sql & "diadevengado decimal(7,2) DEFAULT 0, diaperiodo decimal(7,2) DEFAULT 0, diafisico decimal(7,2) DEFAULT 0, "
        s_Sql = s_Sql & "codmon char(1) Null, remunera_mn decimal(18, 2) DEFAULT 0.00, remunera_me decimal(18, 2) DEFAULT 0.00, "
        s_Sql = s_Sql & "imporpvsacu_mn decimal(18, 2) DEFAULT 0.00, imporpvsacu_me decimal(18, 2) DEFAULT 0.00, "
        s_Sql = s_Sql & "importepvs_mn decimal(18, 2) DEFAULT 0.00, importepvs_me decimal(18, 2) DEFAULT 0.00, "
        s_Sql = s_Sql & "fechacan date NULL) "
        gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
        ' Genero la información
        AnalisisVacaciones "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
        s_Sql = "SELECT * FROM tmp" & gdl_Procedure.ps_ReportName & " "
        s_Sql = s_Sql & "ORDER BY codcco, " & aElemento(0, 1)
      Else
        ' Selección de reporte general
        s_Sql = "SELECT vac.codpsn, vac.pdoano, vac.pdomes, vac.pdopvs, pdo.descripvs, sub.fechaini AS subfechaini, sub.fechafin AS subfechafin, "
        s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS apellidosnombres, "
        s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, "
        s_Sql = s_Sql & "vac.fechaini, vac.fechafin, vac.numerodias, vac.codmon, vac.remunera_mn, vac.remunera_me, vac.imporpvsacu_mn, "
        s_Sql = s_Sql & "vac.imporpvsacu_me, vac.importepvs_mn, vac.importepvs_me, vac.fechacan "
        s_Sql = s_Sql & "FROM plpvsvacaciondet vac "
        s_Sql = s_Sql & "INNER JOIN plpvsperiodovac pdo ON vac.codcls=pdo.codcls AND vac.codpvs=pdo.codpvs "
        s_Sql = s_Sql & "INNER JOIN plpvsvacacion sub ON vac.codcls=sub.codcls AND vac.codpvs=sub.codpvs AND vac.codpsn=sub.codpsn AND vac.pdopvs=sub.pdopvs "
        s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON vac.codcls=psn.codcls AND vac.codpsn=psn.codpsn "
        s_Sql = s_Sql & "WHERE vac.codcls='" & ps_ClsPlanilla & "'"
        s_Sql = s_Sql & "AND vac.pdoano='" & ps_Anyo & "' "
        If cboPeriodo.Text <> "" Then
          s_Sql = s_Sql & "AND vac.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
        End If
        s_Sql = s_Sql & "AND vac.codpsn IN(SELECT valor FROM rangoimpresion "
        s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
        s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
        s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
        s_Sql = s_Sql & "ORDER BY vac.pdoano, " & aElemento(0, 1)
      End If
    ElseIf s_OptRegistro = "pvsgratifi" Then        ' Provisión de gratificaciones
      s_Sql = "SELECT gra.codpsn, gra.pdoano, gra.sempvs, gra.pdomes, pdo.descripvs, pdo.mesini, pdo.mesfin, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, '')), ', ', TRIM(IFNULL(psn.nombres, ''))) AS apellidosnombres, "
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, psn.codcco, cco.detcco, "
      s_Sql = s_Sql & "gra.fechaini, gra.fechafin, gra.numerodias, gra.codmon, gra.remunera_mn, gra.remunera_me, gra.imporpvsacu_mn, "
      s_Sql = s_Sql & "gra.imporpvsacu_me, gra.importepvs_mn, gra.importepvs_me, gra.fechacan "
      s_Sql = s_Sql & "FROM plpvsgratifica gra "
      s_Sql = s_Sql & "INNER JOIN plpvsperiodogra pdo ON pdo.codcls=gra.codcls AND pdo.pdoano=gra.pdoano AND pdo.sempvs=gra.sempvs "
      s_Sql = s_Sql & "LEFT JOIN plpersonal psn ON psn.codcls=gra.codcls AND psn.codpsn=gra.codpsn "
      s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON cco.codcco=psn.codcco "
      s_Sql = s_Sql & "WHERE gra.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND gra.pdoano='" & ps_Anyo & "' "
      If cboPeriodo.Text <> "" Then
        s_Sql = s_Sql & "AND gra.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "' "
      End If
      s_Sql = s_Sql & "AND gra.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      s_Sql = s_Sql & "ORDER BY psn.codcco, gra.sempvs, " & aElemento(0, 1)
    ElseIf s_OptRegistro = "pvscoxtise" Then        ' Provisión de cts
      s_Sql = s_Sql & "codpsn varchar(11) NOT Null, nombrepsn varchar(75) Null, codcco varchar(12) Null, "
      s_Sql = s_Sql & "detcco varchar(40) Null, pdoano char(4) Null, pdomes char(2) Null, "
      s_Sql = s_Sql & "descgo varchar(50), fecingreso date Null, fecbaja date Null, "
      s_Sql = s_Sql & "fechaini date Null, fechafin date Null, "
      s_Sql = s_Sql & "numeromes int(2) DEFAULT 0, numerodia int(2) DEFAULT 0, "
      s_Sql = s_Sql & "codmon char(1) Null, remunera_mn decimal(18, 2) DEFAULT 0.00, remunera_me decimal(18, 2) DEFAULT 0.00, "
      s_Sql = s_Sql & "imporpvsacu_mn decimal(18, 2) DEFAULT 0.00, imporpvsacu_me decimal(18, 2) DEFAULT 0.00, "
      s_Sql = s_Sql & "importepvs_mn decimal(18, 2) DEFAULT 0.00, importepvs_me decimal(18, 2) DEFAULT 0.00, "
      s_Sql = s_Sql & "fechacan date NULL) "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      If cboPeriodo.Text <> "" Then
        ' Genero la información
        AnalisisCTS "tmp" & gdl_Procedure.ps_ReportName, s_OptRegistro, s_FechaHora
      End If
      s_Sql = "SELECT * FROM tmp" & gdl_Procedure.ps_ReportName & " "
      s_Sql = s_Sql & "ORDER BY codcco, " & aElemento(0, 1)
    End If
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 7), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino el rango de impresion
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
    ' Elimino la tabla temporal
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
    ' ]
   Case 9  ' Peridos de provisión de vacaciones, gratificación y cts
    fMenu.Tag = s_OptRegistro
    If s_OptRegistro = "pvsvacacio" Then
      o_PvsVacaPeriodo.Show
    ElseIf s_OptRegistro = "pvsgratifi" Then
      o_PvsGratiPeriod.Show
    ElseIf s_OptRegistro = "pvscoxtise" Then
      fCtsPeriodo.Show
    End If
   Case 10 ' Calculo de provision vacación ,gratificación, cts
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Proceso", vbExclamation: Exit Sub
    fMenu.Tag = s_OptRegistro & "1"
    If s_OptRegistro = "pvsvacacio" Then
      o_PvsVacaCalcul.Show
    ElseIf s_OptRegistro = "pvsgratifi" Then
      o_PvsGratiCalcul.Show
    ElseIf s_OptRegistro = "pvscoxtise" Then
      fCtsCalculo.Show
    End If
   Case 11 ' Depurar calculo vacaciones, gratificación y cts
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Proceso", vbExclamation: Exit Sub
    fMenu.Tag = s_OptRegistro & "2"
    If s_OptRegistro = "pvsvacacio" Then
      o_PvsVacaDepura.Show
    ElseIf s_OptRegistro = "pvsgratifi" Then
      o_PvsGratiDepura.Show
    ElseIf s_OptRegistro = "pvscoxtise" Then
      fCtsDepuracion.Show
    End If
   Case 12 ' Cancelación de vacaciones, gratificación y cts
    If s_OptRegistro = "pvscoxtise" Then
      ' Verifico que existan registros seleccionados
      If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Proceso", vbExclamation: Exit Sub
  
      fMenu.Tag = s_OptRegistro
      s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
      ' Cambio el Mensaje y Muestro la Barra
      s_OldMessage = fMenu.panMessage.Caption
      MuestraMensaje "Generando Información ..."
      fCtsCancelacion.Show
      ' Reinicializo los mensajes
      MuestraMensaje s_OldMessage
    End If
  End Select

End Sub
Private Sub cmdGenera_Click()
  Dim sOldMessage As String, sFecIngreso As String, sFecCese As String
  Dim sPeriodo As String, dFecInicio As Date, dFecFinal As Date
  Dim nContador As Integer, sPvsPeriodo As String
  Dim porstPeriodo As ADODB.Recordset
    
  ' Verifico que existan registros seleccionados
  If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Proceso", vbExclamation: Exit Sub
  ' Verifico que existan periodos de provisión
  s_Sql = "SELECT COUNT(*) AS registro "
  s_Sql = s_Sql & "FROM plpvsperiodovac "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND estadopvs<>'" & s_Estado_Blq & "'"
  Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If porstRecordset("registro") = 0 Then Beep: MsgBox "Debe Registrar los Ejercicios de Provisión", vbExclamation: Exit Sub
    
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  ' Obtengo los ejerccios de de provisión general
  s_Sql = "SELECT codpvs, fechapvs "
  s_Sql = s_Sql & "FROM plpvsperiodovac "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND estadopvs<>'" & s_Estado_Blq & "'"
  Set porstPeriodo = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstPeriodo.BOF And porstPeriodo.EOF) Then
    ' Cambio el Mensaje y Muestro la Barra
    sOldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Periodo ..."
    fMenu.panPercent.Visible = True
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To (tdbRegistro.SelBookmarks.Count - 1)
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      
      ' Personal activo
      If chkCesados.Value = Unchecked Then
        ' Personal acivo, sin remuneración integral
        If (dcaRegistro.Recordset!estadopsn <> "I" And dcaRegistro.Recordset!remintegralvaca = s_Estado_Ina) Then
          sFecIngreso = Format(tdbRegistro.Columns(2).Value, s_FormatoFecha)
          porstPeriodo.MoveFirst
          While Not porstPeriodo.EOF
            sPvsPeriodo = porstPeriodo("codpvs")
            ' Año ingreso mayor periodo provisión
            If (sPvsPeriodo > Format(sFecIngreso, "yyyy")) Then
              ' Si no se encuentra provisionado
              s_Sql = "select count(*) AS registro "
              s_Sql = s_Sql & "from plpvsvacacion "
              s_Sql = s_Sql & "where codcls='" & ps_ClsPlanilla & "' "
              s_Sql = s_Sql & "and codpvs='" & sPvsPeriodo & "' "
              s_Sql = s_Sql & "and codpsn='" & Trim(tdbRegistro.Columns(0).Value) & "'"
              Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
              If porstRecordset("registro") = 0 Then
                dFecInicio = Left(sFecIngreso, 6) & (sPvsPeriodo - 1)
                dFecFinal = DateAdd("yyyy", 1, dFecInicio) - 1
                sPeriodo = Format(dFecInicio, "yyyy") & sPvsPeriodo
                
                ' Creo los arreglos para la actualización
                a_Campos = Array("codcls", "codpvs", "codpsn", "pdopvs", "fechaini", "fechafin", "numerodias", "estadovac", "usrcre", "fyhcre")
                a_Valores = Array(ps_ClsPlanilla, sPvsPeriodo, Trim(tdbRegistro.Columns(0).Value), sPeriodo, Format(dFecInicio, s_FmtFechMysql_0), Format(dFecFinal, s_FmtFechMysql_0), 360, s_Estado_Ina, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
                a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
                
                ' Realizo el proceso de actualización de los registros
                If Not Records_Ins("plpvsvacacion", a_Campos, a_Valores, a_Tipos) Then GoTo Error
              End If
            End If
            porstPeriodo.MoveNext
          Wend
        End If
      Else
        ' Personal inactivo
        sFecIngreso = Format(tdbRegistro.Columns(2).Value, s_FormatoFecha)
        sFecCese = Format(tdbRegistro.Columns(3).Value, s_FormatoFecha)
        porstPeriodo.MoveFirst
        
        While Not porstPeriodo.EOF
          sPvsPeriodo = porstPeriodo("codpvs")
          ' Si no se encuentra provisionado
          s_Sql = "SELECT COUNT(*) AS registro "
          s_Sql = s_Sql & "FROM plpvsvacacion "
          s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
          s_Sql = s_Sql & "AND codpvs='" & sPvsPeriodo & "' "
          s_Sql = s_Sql & "AND codpsn='" & Trim(tdbRegistro.Columns(0).Value) & "'"
          Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
            
          If porstRecordset("registro") = 0 Then
            dFecInicio = Left(sFecIngreso, 6) & (sPvsPeriodo - 1)
            dFecFinal = DateAdd("yyyy", 1, dFecInicio) - 1
            sPeriodo = (sPvsPeriodo - 1) & sPvsPeriodo
               
            If Format(dFecInicio, s_FmtFechMysql_0) >= Format(sFecIngreso, s_FmtFechMysql_0) Then
              If Format(sFecCese, s_FmtFechMysql_0) >= Format(dFecInicio, s_FmtFechMysql_0) Then
                ' Creo los arreglos para la actualización
                a_Campos = Array("codcls", "codpvs", "codpsn", "pdopvs", "fechaini", "fechafin", "numerodias", "estadovac", "usrcre", "fyhcre")
                a_Valores = Array(ps_ClsPlanilla, sPvsPeriodo, Trim(tdbRegistro.Columns(0).Value), sPeriodo, Format(dFecInicio, s_FmtFechMysql_0), Format(dFecFinal, s_FmtFechMysql_0), 360, s_Estado_Ina, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
                a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
                ' Realizo el proceso de actualización de los registros
                If Not Records_Ins("plpvsvacacion", a_Campos, a_Valores, a_Tipos) Then GoTo Error
              End If
            End If
          End If
          porstPeriodo.MoveNext
        Wend
      End If
      
      ' Incremento el porcentaje
      'fMenu.panPercent.FloodPercent = (((n_Index + 1) * 100) \ tdbRegistro.SelBookmarks.Count)
    Next n_Index
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    MsgBox "Generación concluyó satisfactoriamente.", vbInformation + vbOKOnly
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.Visible = False
  fMenu.panPercent.FloodPercent = 0
  MuestraMensaje sOldMessage
  ' cierro objetos
  porstPeriodo.Close
  Set porstPeriodo = Nothing
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing

End Sub
Private Sub dcaRegistro_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  If s_OptRegistro = "pvsvacacio" Then
    If FormVisible("fPvsVacacion") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fPvsVacacion.RecuperaRegistros "codpvs, pdopvs"
      End If
    End If
  ElseIf s_OptRegistro = "pvsgratifi" Then
    If FormVisible("fPvsGratificacion") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fPvsGratificacion.RecuperaRegistros "pdoano, sempvs"
      End If
    End If
  ElseIf s_OptRegistro = "pvscoxtise" Then
    If FormVisible("fCtsMovimiento") Then
      If Not dcaRegistro.Recordset.EOF And Not dcaRegistro.Recordset.BOF Then
        fCtsMovimiento.RecuperaRegistros "pdocts, subcts"
      End If
    End If
  End If

End Sub
Private Sub Form_Activate()
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()

Set cnn = New ADODB.Connection
cnn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & ps_Servidor & ";uid=" & ps_UserId & ";pwd=" & ps_Password & ";database=" & ps_DataBase & ";connection="
cnn.CursorLocation = adUseClient
cnn.Open

  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6510: Me.Width = 8580
  Me.Left = 400: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = s_SwRegistro
  
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
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(14, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "provision": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "seleccio", "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir", "analsald", "genvouch", "borrafor", "pagar", "consolid")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Selecciona " & s_TitleTable, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango", "Presentación Preliminar", "Imprimir", "Periodo de Provisión", "Calculo Provisión", "Depurar Calculo", "Cancelación de Provisión", "Importar Información")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Visualizo adicional pagos cts e importar
  cmdAction(12).Tag = IIf(s_OptRegistro = "pvscoxtise", "1", "3")
  cmdAction(13).Tag = IIf(s_OptRegistro = "pvsvacacio", "1", "3")
  
  ' Configuro los botones de parametro
  gdl_Procedure.LoadGrafics cmdGenera, "actsaldo", "Genera Periodo x Personal"
  
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  For n_Index = 1 To 13: cboPeriodo.AddItem Choose(n_Index, "", "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  cboPeriodo.ListIndex = 0
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Recupero los registros con el control de datos asignado (orden)
  tdbRegistro.DataSource = dcaRegistro
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
  ribParametro(0).Value = True
  cmdGenera.Visible = (s_OptRegistro = "pvsvacacio")
  chkCesados.Visible = (s_OptRegistro = "pvsvacacio")
  
  '*****************************
  Dim Rst As New Recordset
  Dim sql As String
  Dim i As Integer
  toolbarcanc.Visible = False
  
  If s_OptRegistro = "pvsvacacio" Or s_OptRegistro = "pvsgratifi" Then
    toolbarcanc.Visible = True
    If s_OptRegistro = "pvsvacacio" Then
      sql = "SELECT codpvs from plpvsperiodovac "
    Else
      sql = "SELECT concat(pdoano,sempvs) from plpvsperiodogra "
    End If
    sql = sql & "WHERE codcls='" & ps_ClsPlanilla & "'"
    Rst.Open sql, cnn, adOpenStatic, adLockOptimistic
    On Error GoTo Error
    Rst.MoveFirst
    i = 1
    While Not Rst.EOF
      toolbarcanc.Buttons(1).ButtonMenus.Add(i).Key = "A" & i
      toolbarcanc.Buttons(1).ButtonMenus(i).Text = Rst(0)
      i = i + 1
      Rst.MoveNext
    Wend
  End If
'*****************************

Error:
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If s_OptRegistro = "pvsvacacio" Then
    If FormVisible("fPvsVacacion") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fPvsVacacion.Caption, vbExclamation
      Cancel = True
    End If
  ElseIf s_OptRegistro = "pvsgratifi" Then
    If FormVisible("fPvsGratificacion") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fPvsGratificacion.Caption, vbExclamation
      Cancel = True
    End If
  ElseIf s_OptRegistro = "pvscoxtise" Then
    If FormVisible("fCtsMovimiento") Then
      Beep
      MsgBox "Primero debe cerrar la Pantalla de " & fCtsMovimiento.Caption, vbExclamation
      Cancel = True
    End If
  End If
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
Private Sub toolbarcanc_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

Dim s_FechaHora As String, s_OldMessage As String

Dim anno As Integer


Select Case ButtonMenu.Key
Case "A" & ButtonMenu.Index

fMenu.Tag = s_OptRegistro
s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
      
For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
Next n_Index
   
If MsgBox("Desea Cancelar las Provisiones Seleccionadas, estas Seguro? ", vbOKCancel + vbQuestion, "Aviso  " & "") = vbOK Then
        
              
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Información ..."
        
    anno = ButtonMenu.Text
    anno = anno - 1
        
    If s_OptRegistro = "pvsvacacio" Then
    
    s_Sql = "update plpvsvacacion vac set vac.estadovac=2 "
    s_Sql = s_Sql & " WHERE vac.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & " AND vac.codpvs='" & ButtonMenu.Text & "' and vac.pdopvs='" & anno & ButtonMenu.Text & "'"
    s_Sql = s_Sql & " AND vac.codpsn IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & " WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & " AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & " AND fyhcre='" & s_FechaHora & "') "
    
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "update plpvsvacaciondet vac set vac.estadodet=2 "
    s_Sql = s_Sql & " WHERE vac.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & " AND vac.codpvs='" & ButtonMenu.Text & "' and vac.pdopvs='" & anno & ButtonMenu.Text & "'"
    's_Sql = s_Sql & " AND vac.pdoano in ('" & Left(ButtonMenu.Text, 4) & "','" & anno & "') and vac.pdomes='" & Format(cboPeriodo.ListIndex, "00") & "'"
    s_Sql = s_Sql & " AND vac.pdoano in ('" & Left(ButtonMenu.Text, 4) & "','" & anno & "')"
    s_Sql = s_Sql & " AND vac.codpsn IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & " WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & " AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & " AND fyhcre='" & s_FechaHora & "') "
   
    End If
    
    If s_OptRegistro = "pvsgratifi" Then
         
    s_Sql = "update plpvsgratifica gra set gra.estadogra=2 "
    s_Sql = s_Sql & " WHERE gra.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & " AND gra.pdoano='" & Left(ButtonMenu.Text, 4) & "' and gra.sempvs='" & Right(ButtonMenu.Text, 1) & "'"
    s_Sql = s_Sql & " AND gra.codpsn IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & " WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & " AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & " AND fyhcre='" & s_FechaHora & "') "
    
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
    s_Sql = "update plpvsperiodogra gra set gra.estadopvs=2 "
    s_Sql = s_Sql & " WHERE gra.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & " AND gra.pdoano='" & Left(ButtonMenu.Text, 4) & "' and gra.sempvs='" & Right(ButtonMenu.Text, 1) & "'"
       
    End If
    
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
   
    If s_OptRegistro = "pvsvacacio" Then
        MsgBox "Provisiones de Vacaciones Canceladas", vbExclamation, "Sistema de Planillas"
    End If
    If s_OptRegistro = "pvsgratifi" Then
        MsgBox "Provisiones de Gratificaciones Canceladas", vbExclamation, "Sistema de Planillas"
    End If
    
   
    
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
        
End If
   
    
End Select
End Sub
