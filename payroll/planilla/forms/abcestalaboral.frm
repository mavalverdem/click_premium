VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcEstaLaboral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   7170
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   360
      TabIndex        =   0
      Top             =   3960
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
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
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
      Begin Threed.SSCommand cmdCancel 
         Height          =   360
         Left            =   6450
         TabIndex        =   2
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
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
         Picture         =   "abcestalaboral.frx":0000
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   3
         Top             =   75
         Visible         =   0   'False
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcestalaboral.frx":001C
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   675
         TabIndex        =   4
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Height          =   3495
      Index           =   0
      Left            =   6435
      TabIndex        =   5
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   6165
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
         TabIndex        =   6
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edici�n"
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
         TabIndex        =   7
         Tag             =   "0"
         Top             =   600
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
         Picture         =   "abcestalaboral.frx":0038
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1230
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
         Picture         =   "abcestalaboral.frx":0054
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   1830
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
         Picture         =   "abcestalaboral.frx":0070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   4110
      Width           =   7170
      _Version        =   65536
      _ExtentX        =   12647
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
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   3
         Left            =   4680
         TabIndex        =   11
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
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
         Picture         =   "abcestalaboral.frx":008C
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   12
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
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
         Picture         =   "abcestalaboral.frx":00A8
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   13
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcestalaboral.frx":00C4
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   14
         Top             =   75
         Width           =   390
         _Version        =   65536
         _ExtentX        =   688
         _ExtentY        =   635
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "abcestalaboral.frx":00E0
      End
      Begin MSAdodcLib.Adodc dcaHelp 
         Height          =   330
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Adodc1"
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
   Begin TabDlg.SSTab tabRegister 
      Height          =   3510
      Left            =   75
      TabIndex        =   15
      Top             =   600
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   6191
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabMaxWidth     =   3052
      BackColor       =   -2147483644
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
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "abcestalaboral.frx":00FC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblHelp(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdHelp(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtruc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtcodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbMes"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.ComboBox cmbMes 
         ForeColor       =   &H00FF8080&
         Height          =   315
         ItemData        =   "abcestalaboral.frx":0118
         Left            =   720
         List            =   "abcestalaboral.frx":011A
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   840
         Width           =   1590
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Height          =   300
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   5895
         Begin VB.TextBox txttasa 
            Enabled         =   0   'False
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
            Height          =   300
            Left            =   2280
            TabIndex        =   24
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblDato 
            Alignment       =   1  'Right Justify
            Caption         =   "Tasa :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   960
            TabIndex        =   23
            Top             =   840
            Width           =   1200
         End
         Begin VB.Label lblDato 
            Caption         =   "Cod. Establecimiento :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   435
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1920
         End
      End
      Begin VB.TextBox txtruc 
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
         Height          =   300
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Mes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   450
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
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Ruc :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "fAbcEstaLaboral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificaci�n de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_Registro As String                            ' Codigo del registro
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acci�n
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creaci�n", IIf(Me.Tag = s_MdoData_Del, " - Eliminaci�n", IIf(Me.Tag = s_MdoData_Upd, " - Actualizaci�n", " - Consulta")))
  For n_Index = 0 To 3: cmdMove(n_Index).Visible = (Me.Tag = s_MdoData_Vis): Next n_Index
  cmdUpdate.Visible = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)
  cmdAction(0).Enabled = (Me.Tag <> s_MdoData_Ins)
  cmdAction(1).Enabled = (Me.Tag = s_MdoData_Upd Or Me.Tag = s_MdoData_Vis)
  cmdAction(2).Enabled = (Me.Tag = s_MdoData_Del Or Me.Tag = s_MdoData_Vis)
  cmdHelp(0).Enabled = (Me.Tag = s_MdoData_Ins Or Me.Tag = s_MdoData_Upd)

End Sub
Sub ShowScreen()
  ' Presenta Botones y Controles
  EnabledBotons
  ' Presenta datos en pantalla de acuerdo al modo Seleccionado
  If Me.Tag = s_MdoData_Ins Then
    gdl_Procedure.EditText "AT", txtruc, "", Me.Tag, False, 11
    gdl_Procedure.EditText "AT", txtCodigo, "", Me.Tag, False, 4, vbRightJustify
    gdl_Procedure.EditText "AT", txttasa, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    
    gdl_Procedure.EditCombo "AT", cmbmes, 0, Me.Tag, False
    
  Else
    gdl_Procedure.EditText "AT", txtruc, fAbcPersonal.tdbesta.Columns(2).Text, Me.Tag, False, 11
    gdl_Procedure.EditText "AT", txtCodigo, fAbcPersonal.tdbesta.Columns(3).Text, Me.Tag, True, 4, vbRightJustify
    gdl_Procedure.EditText "AT", txttasa, FormatNumber(fAbcPersonal.tdbesta.Columns(5).Text, 2), Me.Tag, True, 6, vbRightJustify
           
    n_Index = fAbcPersonal.tdbesta.Columns(1).Text
    gdl_Procedure.EditCombo "AT", cmbmes, (n_Index - 1), Me.Tag, False
    
  End If
  
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtruc, "RU")

End Sub
']
Private Sub cmdAction_Click(Index As Integer)
  Dim n_Secuencia As Integer
  Dim p_ano As String
  Dim p_mes As String
  
  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtruc.SetFocus
  ElseIf Index = 2 Then
    txtruc.SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("� Est�s Seguro de Eliminar el " & lblTitle & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim(fAbcPersonal.txtCodigo)
    n_Secuencia = CInt(fAbcPersonal.tdbesta.Columns(5).Text)
   p_ano = fAbcPersonal.tdbesta.Columns(0).Text
   p_mes = fAbcPersonal.tdbesta.Columns(1).Text
   
    '[ Inicio la conexi�n a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpsn", "orden", "ano", "mes")
    a_Valores = Array(ps_ClsPlanilla, s_Registro, n_Secuencia, p_ano, p_mes)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter)
    gdl_Conexion.IniciaTransaccion    'Inicia transacci�n
    ' Elimino el registro
    If Not Records_Del("plestalaboral", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacci�n
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el ado control y la grilla
    fAbcPersonal.Recuperarestalaboral
    ' Verifico si aun existen registros
    l_ExistRecord = ((fAbcPersonal.tdbesta.EOF And fAbcPersonal.tdbesta.BOF) Or fAbcPersonal.tdbesta.VisibleRows = 0)
    If Not l_ExistRecord Then
'      fRemunerExcepcional.dcaRegistro.Recordset.Find ("codcpc >= '" & s_ConceptoPlanilla & "'")
      If fAbcPersonal.tdbesta.EOF Then fAbcPersonal.tdbesta.MoveLast
    Else
      Unload Me
    End If
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexi�n a la base de datos ]
  Set gdl_Conexion = Nothing
  If Not l_ExistRecord Then cmdCancel_Click

End Sub
Private Sub cmdCancel_Click()
    
  If Me.Tag = s_MdoData_Vis Or l_ExistRecord Then
    Unload Me
  Else
    Me.Tag = s_MdoData_Vis: ShowScreen
  End If

End Sub
Private Sub cmdHelp_Click(Index As Integer)

  s_SqlHelp = ""
  ' Recupero la informaci�n
  s_Sql = gdl_Funcion.HelpTablas("eqd", "codeqd", ps_RucEmpresa, "")
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (tabRegister.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)))
  tdbHelp.Left = (tabRegister.Left + (cmdHelp(Index).Left + (cmdHelp(Index).Width / 2)))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub
Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fAbcPersonal.tdbesta.MoveFirst
   Case 1: If Not fAbcPersonal.tdbesta.BOF Then fAbcPersonal.tdbesta.MovePrevious
           If fAbcPersonal.tdbesta.BOF Then fAbcPersonal.tdbesta.MoveFirst
   Case 2: If Not fAbcPersonal.tdbesta.EOF Then fAbcPersonal.tdbesta.MoveNext
           If fAbcPersonal.tdbesta.EOF Then fAbcPersonal.tdbesta.MoveLast
   Case 3: fAbcPersonal.tdbesta.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim n_Secuencia As Integer
  
  ' Realizo las validaciones de los campos a actualizar
  If txtruc = "" Then Beep: MsgBox "Debe Ingresar el Ruc " & lblTitle, vbExclamation: txtruc.SetFocus: Exit Sub
  If txtCodigo = "" Then Beep: MsgBox "Debe Ingresar el Codigo de Establecimiento " & lblTitle, vbExclamation: txtCodigo.SetFocus: Exit Sub
  'If lblHelp(0) = "???" Then Beep: MsgBox "Ruc de Empresa no es valido; Verificar", vbExclamation: txtruc.SetFocus: Exit Sub
  If cmbmes = "" Then Beep: MsgBox "Debe Ingresar el Periodo " & lblTitle, vbExclamation: cmbmes.SetFocus: Exit Sub
    
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = Trim(fAbcPersonal.txtCodigo)
  ' Obtengo el orden correlativo
  n_Secuencia = CInt(Val(fAbcPersonal.tdbesta.Columns(5).Text))
  If Me.Tag = s_MdoData_Ins Then
    s_Sql = "SELECT IFNULL(MAX(orden), 0)+1 AS registro "
    s_Sql = s_Sql & "FROM plestalaboral "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codpsn='" & s_Registro & "' "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
    n_Secuencia = CInt(porstRecordset!registro)
  End If
   
  ' Creo los arreglos para la actualizaci�n
  a_Campos = Array("codcls", "codpsn", "orden", "ano", "mes", "ruc", "codest", "tasa", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, s_Registro, n_Secuencia, Year(gs_FechaHora), Left(cmbmes, 2), Trim(txtruc.Text), Trim(txtCodigo.Text), Trim(txttasa.Text), ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA)
  a_Where = Array("codcls", "codpsn", "orden")
  
  '[ Inicio la conexi�n a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacci�n
  ' Realizo el proceso de actualizaci�n de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plestalaboral", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plestalaboral", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacci�n
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  fAbcPersonal.Recuperarestalaboral
  ' Ubico el registro ingresado o actualizado
  'fAbcPersonal.plexpelaboral.f dcaRegistro.Recordset.Find ("codcpc='" & s_ConceptoPlanilla & "'")
  ' si es actualizaci�n pasa al modo visualizaci�n
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txtruc.SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexi�n a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminaci�n
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posici�n y titulo del formulario
  Me.Height = 4500: Me.Width = 7290
  Me.Left = 3980: Me.Top = 1830
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualizaci�n Empresas donde labora el Trabajador (Destaque)"
  lblTitle = "Empresas donde labora el Trabajador (Destaque)"
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operaci�n del registro
  Me.Tag = fAbcPersonal.tdbesta.Tag
  
  ' Configuro parametros de visualizaci�n del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y t�tulo del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "A�adir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualizaci�n del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y t�tulo del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualizaci�n
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Informaci�n de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Informaci�n de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fAbcPersonal.tdbesta.EOF Or fAbcPersonal.tdbesta.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fAbcPersonal.tdbesta.Columns(0).Text
  
' Adiciono los meses
  For n_Index = 1 To 12
    cmbmes.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre")
  Next n_Index
  cmbmes.ListIndex = 1
  
  
  ' Carga los datos en el formulario
  ShowScreen
 
 '[ Configuraci�n de la grilla de ayuda
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Codigo", "RUC", "Descripcion", "Cod. Establecimiento", "Tasa")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codepr", "rucepr", "desepr", "cdgepr", "tasepr")
      aElemento(n_Index, 2) = Choose(n_Index + 1, 900, 1500, 3500, 1000, 1000)
      aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbLeftJustify)
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
      aElementos(n_Index, 1) = n_BackColorHelp#: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbHelp, aElemento, aElementos
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Ruc de Establecimientos de Empresas donde labora el Trabajador (Destaque)", 2
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' Habilito/desabilito botones inciales
  fAbcPersonal.cmdActionEsta(0).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActionEsta(1).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActionEsta(2).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  txtruc.Text = tdbHelp.Columns(1).Value
  lblHelp(0) = tdbHelp.Columns(2).Value
  txtCodigo.Text = tdbHelp.Columns(3).Value
  txttasa.Text = tdbHelp.Columns(4).Value
  txtruc.SetFocus
  
End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la informaci�n ordenada
  s_Sql = gdl_Funcion.HelpTablas("eqd", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
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
Private Sub txtruc_GotFocus()
  gdl_Procedure.MarcaGet txtruc
End Sub
Private Sub txtruc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtruc_LostFocus()
  'lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtruc, "RU")
End Sub
Private Sub txtCodigo_GotFocus()
  gdl_Procedure.MarcaGet txtCodigo
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    If txtCodigo = "" Then
      Beep
      MsgBox "Debe Ingresar el Numero de Codigo de Establecimiento " & lblTitle, vbExclamation
      txtCodigo.SetFocus
    Else
      txttasa.SetFocus
      KeyAscii = 0
    End If
  End If

End Sub
Private Sub txttasa_Validate(Cancel As Boolean)
  txttasa.Text = IIf(Not IsNumeric(txttasa.Text), 0, txttasa.Text)
  If CDec(txttasa.Text) < 0 Then MsgBox "Tasa no puede ser negativo; Verifique", vbInformation: txttasa.SetFocus: Exit Sub
  txttasa.Text = FormatNumber(CDec(txttasa.Text), 2)
End Sub
