VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form fAbcComprobantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   7170
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   1
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
         Picture         =   "abcComprobantes.frx":0000
      End
      Begin Threed.SSCommand cmdUpdate 
         Height          =   360
         Left            =   6060
         TabIndex        =   2
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
         Picture         =   "abcComprobantes.frx":001C
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
         TabIndex        =   3
         Top             =   120
         Width           =   5070
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   2  'Align Bottom
      Height          =   510
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   4140
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
         Left            =   4695
         TabIndex        =   5
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
         Picture         =   "abcComprobantes.frx":0038
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   2
         Left            =   4305
         TabIndex        =   6
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
         Picture         =   "abcComprobantes.frx":0054
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   1
         Left            =   2595
         TabIndex        =   7
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
         Picture         =   "abcComprobantes.frx":0070
      End
      Begin Threed.SSCommand cmdMove 
         Height          =   360
         Index           =   0
         Left            =   2205
         TabIndex        =   8
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
         Picture         =   "abcComprobantes.frx":008C
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2160
      TabIndex        =   9
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
      Height          =   3510
      Index           =   0
      Left            =   6435
      TabIndex        =   10
      Top             =   600
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   6191
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
         TabIndex        =   11
         Top             =   15
         Width           =   720
         _Version        =   65536
         _ExtentX        =   1270
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Edición"
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
         TabIndex        =   12
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
         Picture         =   "abcComprobantes.frx":00A8
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   13
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
         Picture         =   "abcComprobantes.frx":00C4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   14
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
         Picture         =   "abcComprobantes.frx":00E0
      End
   End
   Begin TabDlg.SSTab tabRegister 
      Height          =   3615
      Left            =   0
      TabIndex        =   15
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6376
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
      TabPicture(0)   =   "abcComprobantes.frx":00FC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDato(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDato(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDato(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDato(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDato(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDato(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblHelp(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdHelp(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpFechas(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmCuadro(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpFechas(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtnumero"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtserie"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtimporte"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txttipo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txttipo 
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
         Height          =   300
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   30
         Top             =   120
         Width           =   345
      End
      Begin VB.TextBox txtimporte 
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
         Left            =   1680
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtserie 
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
         Left            =   1680
         TabIndex        =   25
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtnumero 
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
         Left            =   1680
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpFechas 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   17
         Top             =   1560
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   115867649
         CurrentDate     =   37515
      End
      Begin Threed.SSFrame frmCuadro 
         Height          =   645
         Index           =   0
         Left            =   1680
         TabIndex        =   18
         Top             =   2400
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   1138
         _StockProps     =   14
         Caption         =   "Retencion"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   19
            Top             =   285
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&Si"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optEstado 
            Height          =   195
            Index           =   1
            Left            =   1635
            TabIndex        =   20
            Top             =   285
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "&No"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSComCtl2.DTPicker dtpFechas 
         Height          =   300
         Index           =   1
         Left            =   1680
         TabIndex        =   29
         Top             =   1920
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   115867649
         CurrentDate     =   37515
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   31
         Top             =   120
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   494
         _StockProps     =   78
         Caption         =   "..."
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   32
         Top             =   120
         Width           =   195
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Pago :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Importe :"
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
         Left            =   360
         TabIndex        =   26
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Serie :"
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
         Index           =   1
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Número :"
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
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Emision :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo :"
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
         Left            =   360
         TabIndex        =   21
         Top             =   120
         Width           =   1200
      End
   End
End
Attribute VB_Name = "fAbcComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String                         ' Titulo de la ventana
Private n_IndexTool As Integer                          ' Indice de la barra de herramientas
Private l_ExistRecord As Boolean                        ' Flag de Verificación de existencia de Registros
Private n_Index As Integer, s_ParCodigo As String       ' Indice para bucle, y parametro de codigo
Private s_Registro As String                            ' Codigo del registro
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
'[
Private Sub EnabledBotons()

  ' Habilita o inabilita los controles de acuerdo a la acción
  Me.Caption = s_TitleWindow & IIf(Me.Tag = s_MdoData_Ins, " - Creación", IIf(Me.Tag = s_MdoData_Del, " - Eliminación", IIf(Me.Tag = s_MdoData_Upd, " - Actualización", " - Consulta")))
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
    gdl_Procedure.EditText "AT", txttipo, "", Me.Tag, False, 50
    gdl_Procedure.EditText "AT", txtserie, "", Me.Tag, False, 30
    gdl_Procedure.EditText "AT", txtnumero, "", Me.Tag, False, 30
     gdl_Procedure.EditText "AT", txtimporte, FormatNumber(0, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), Date, Me.Tag, True, s_FormatoFecha, dtpShortDate
     gdl_Procedure.EditOptionCheck "AT", optEstado(0), True, Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), False, Me.Tag, True
  Else
    gdl_Procedure.EditText "AT", txttipo, fAbcPersonal.tdbcomprobantes.Columns(1).Text, Me.Tag, False, 50
    gdl_Procedure.EditText "AT", txtserie, fAbcPersonal.tdbcomprobantes.Columns(2).Text, Me.Tag, False, 30
    gdl_Procedure.EditText "AT", txtnumero, fAbcPersonal.tdbcomprobantes.Columns(3).Text, Me.Tag, False, 50
    gdl_Procedure.EditText "AT", txtimporte, FormatNumber(fAbcPersonal.tdbcomprobantes.Columns(4).Text, 2), Me.Tag, False, 6, vbRightJustify
    gdl_Procedure.EditDTPicker "AT", dtpFechas(0), fAbcPersonal.tdbcomprobantes.Columns(5).Text, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditDTPicker "AT", dtpFechas(1), fAbcPersonal.tdbcomprobantes.Columns(6).Text, Me.Tag, True, s_FormatoFecha, dtpShortDate
    gdl_Procedure.EditOptionCheck "AT", optEstado(0), (fAbcPersonal.tdbcomprobantes.Columns(7).Value = s_Estado_Act), Me.Tag, True
    gdl_Procedure.EditOptionCheck "AT", optEstado(1), (fAbcPersonal.tdbcomprobantes.Columns(7).Value = s_Estado_Ina), Me.Tag, True
  End If
   lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txttipo, "CO")
End Sub
']
Private Sub cmdAction_Click(Index As Integer)
  Dim n_Secuencia As Integer
  
  ' Cargo los datos en la Ventana de acuerdo al modo
  Me.Tag = Choose(Index + 1, s_MdoData_Ins, s_MdoData_Del, s_MdoData_Upd)
  ShowScreen
  If Index = 0 Then
    txtserie.SetFocus
  ElseIf Index = 2 Then
   txtimporte.SetFocus
  End If
  If Index <> 1 Then Exit Sub
  Beep
  If MsgBox("¿ Estás Seguro de Eliminar el " & lblTitle & "' ?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    ' Capturo el registro a eliminar
    s_Registro = Trim(fAbcPersonal.txtCodigo)
    n_Secuencia = CInt(fAbcPersonal.tdbcomprobantes.Columns(0).Text)
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    ' Creo los arreglos de eliminacion
    a_Where = Array("codcls", "codpsn", "orden")
    a_Valores = Array(ps_ClsPlanilla, s_Registro, n_Secuencia)
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero)
    gdl_Conexion.IniciaTransaccion    'Inicia transacción
    ' Elimino el registro
    If Not Records_Del("plcomprobantect", a_Where, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion  'Confirma transacción
    
    MsgBox "Se Elimino exitosamente " & lblTitle, vbInformation
    ' Refresco el ado control y la grilla
    fAbcPersonal.Recuperarcomprobantes
    ' Verifico si aun existen registros
    l_ExistRecord = ((fAbcPersonal.tdbcomprobantes.EOF And fAbcPersonal.tdbcomprobantes.BOF) Or fAbcPersonal.tdbcomprobantes.VisibleRows = 0)
    If Not l_ExistRecord Then
'      fRemunerExcepcional.dcaRegistro.Recordset.Find ("codcpc >= '" & s_ConceptoPlanilla & "'")
      If fAbcPersonal.tdbcomprobantes.EOF Then fAbcPersonal.tdbcomprobantes.MoveLast
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
  '[ Finalizo la conexión a la base de datos ]
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

  If Not cmdHelp(Index).Enabled Then Exit Sub
  s_SqlHelp = ""
  
  
  
  If n_IndexHelp = Index Then
    tdbHelp.ZOrder 0
    tdbHelp.Visible = True
    Exit Sub
  End If
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = (tabRegister.Top + (cmdHelp(Index).Top + (cmdHelp(Index).Height / 2)))
  tdbHelp.Height = 2400: tdbHelp.Width = 4500
  tdbHelp.Top = 600
  tdbHelp.ZOrder 0
  tdbHelp.Visible = True
  n_IndexHelp = Index

End Sub

Private Sub cmdMove_Click(Index As Integer)

  ' Mueve el Puntero Inicial, Anterior, Siguiente o Final
  Select Case Index
   Case 0: fAbcPersonal.tdbcomprobantes.MoveFirst
   Case 1: If Not fAbcPersonal.tdbcomprobantes.BOF Then fAbcPersonal.tdbcomprobantes.MovePrevious
           If fAbcPersonal.tdbcomprobantes.BOF Then fAbcPersonal.tdbcomprobantes.MoveFirst
   Case 2: If Not fAbcPersonal.tdbcomprobantes.EOF Then fAbcPersonal.tdbcomprobantes.MoveNext
           If fAbcPersonal.tdbcomprobantes.EOF Then fAbcPersonal.tdbcomprobantes.MoveLast
   Case 3: fAbcPersonal.tdbcomprobantes.MoveLast
  End Select

End Sub
Private Sub cmdUpdate_Click()
  Dim n_Secuencia As Integer
  Dim s_Estado As String * 1
  
  ' Realizo las validaciones de los campos a actualizar
  If txttipo = "" Then Beep: MsgBox "Debe Ingresar el Tipo de Comprobante " & lblTitle, vbExclamation: txttipo.SetFocus: Exit Sub
  If lblHelp(0) = "???" Then Beep: MsgBox "Tipo de Comprobante no valido; Verificar", vbExclamation: txttipo.SetFocus: Exit Sub
  If txtserie = "" Then Beep: MsgBox "Debe Ingresar el Numero de Serie " & lblTitle, vbExclamation: txtserie.SetFocus: Exit Sub
  If txtnumero = "" Then Beep: MsgBox "Debe Ingresar el Numero " & lblTitle, vbExclamation: txtnumero.SetFocus: Exit Sub
  If Not (dtpFechas(1) >= dtpFechas(0)) Then Beep: MsgBox "Fecha de termino debe ser mayor o igual que la fecha Inicial", vbExclamation: dtpFechas(1).SetFocus: Exit Sub
  
 s_Estado = IIf(optEstado(0).Value, s_Estado_Act, s_Estado_Ina)
  
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera
  ' Capturo el registro a actualizar
  s_Registro = Trim(fAbcPersonal.txtCodigo)
  ' Obtengo el orden correlativo
  n_Secuencia = CInt(Val(fAbcPersonal.tdbcomprobantes.Columns(0).Text))
  If Me.Tag = s_MdoData_Ins Then
    s_Sql = "SELECT IFNULL(MAX(orden), 0)+1 AS registro "
    s_Sql = s_Sql & "FROM plcomprobantect "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND codpsn='" & s_Registro & "' "
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenKeyset, adLockReadOnly, adUseClient, s_Sql)
    n_Secuencia = CInt(porstRecordset!registro)
  End If
    
  ' Creo los arreglos para la actualización
  a_Campos = Array("codcls", "codpsn", "orden", "tipo", "serie", "numero", "monto", "fecemision", "fecpago", "retencion", IIf(Me.Tag = s_MdoData_Ins, "usrcre", "usrmdf"), IIf(Me.Tag = s_MdoData_Ins, "fyhcre", "fyhmdf"))
  a_Valores = Array(ps_ClsPlanilla, s_Registro, n_Secuencia, Trim(txttipo.Text), Trim(txtserie.Text), Trim(txtnumero.Text), CDec(txtimporte), Format(dtpFechas(0), s_FmtFechMysql_0), Format(dtpFechas(1), s_FmtFechMysql_0), s_Estado, ps_Usuario, Format(Now, s_FmtFeHoMysql_0))
  a_Tipos = Array(TipoDato.Caracter, TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA, TipoDato.FECHA, TipoDato.Caracter, TipoDato.Caracter, TipoDato.FECHA)
  a_Where = Array("codcls", "codpsn", "orden")
  
  '[ Inicio la conexión a la base de datos ]
  ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
  
  gdl_Conexion.IniciaTransaccion    ' Inicia transacción
  ' Realizo el proceso de actualización de los registros
  If Me.Tag = s_MdoData_Ins Then
    If Not Records_Ins("plcomprobantect", a_Campos, a_Valores, a_Tipos) Then GoTo Error
  Else
    If Not Records_Upd("plcomprobantect", a_Campos, a_Valores, a_Tipos, a_Where) Then GoTo Error
  End If
  gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
    
  MsgBox "Se " & IIf(Me.Tag = s_MdoData_Ins, "Inserto", "Actualizo") & " exitosamente el " & lblTitle, vbInformation
  ' Refresco el ado control y la grilla
  fAbcPersonal.Recuperarcomprobantes
  ' Ubico el registro ingresado o actualizado
  'fAbcPersonal.plexpelaboral.f dcaRegistro.Recordset.Find ("codcpc='" & s_ConceptoPlanilla & "'")
  ' si es actualización pasa al modo visualización
  If Me.Tag = s_MdoData_Upd Then
    cmdCancel_Click
  Else
    ShowScreen
    txttipo.SetFocus
  End If
  GoTo Finalizar
  
Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Coloco el puntero en normal
  gdl_Procedure.PunteroNormal
  '[ Finalizo la conexión a la base de datos ]
  Set gdl_Conexion = Nothing
  
End Sub
Private Sub Form_Activate()
  ' Si es modo de eliminación
  If Me.Tag = s_MdoData_Del Then cmdAction_Click (1)
End Sub
Private Sub Form_Load()

  'Establece posición y titulo del formulario
  Me.Height = 4500: Me.Width = 7290
  Me.Left = 3980: Me.Top = 1830
  
  ' Titulo del formulario y panel
  s_TitleWindow = "Actualización Comprobantes de Cuarta Categoria"
  lblTitle = "Comprobantes de Cuarta Categoria"
  n_IndexHelp = -1
  ' Coloco el puntero en espera
  gdl_Procedure.PunteroEnEspera

  ' Obtengo el modo de operación del registro
  Me.Tag = fAbcPersonal.tdbcomprobantes.Tag
  
  ' Configuro parametros de visualización del formulario y los controles del toolbar
  ReDim aElemento(3, 2)
  ' Icono y título del formulario
  aElemento(3, 1) = "edit": aElemento(3, 2) = s_TitleWindow
  ' Cargo los graficos a los controles del toolbar
  For n_Index = 0 To 2
    aElemento(n_Index, 1) = Choose(n_Index + 1, "anadir", "borrar", "modifica")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Añadir ", "Eliminar ", "Modificar ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Configuro parametros de visualización del formulario y los controles de movimiento
  ReDim aElemento(4, 2)
  ' Icono y título del formulario
  aElemento(4, 1) = "edit": aElemento(4, 2) = s_TitleWindow
  ' Cargo los graficos a los controles de movimiento
  For n_Index = 0 To 3
    aElemento(n_Index, 1) = Choose(n_Index + 1, "primero", "anterior", "siguient", "ultimo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ir al Primero ", "Ir al Anterior ", "Ir al Siguiente ", "Ir al Ultimo ") & lblTitle
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdMove, aElemento
  
  ' Configuro los Controles de actualización
  gdl_Procedure.LoadGrafics cmdUpdate, "aceptar", "Actualizar Información de " & lblTitle
  gdl_Procedure.LoadGrafics cmdCancel, "cancelar", "Cancelar Información de " & lblTitle
  cmdCancel.Cancel = True
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  
  ' Verifico si existen Registros
  l_ExistRecord = (fAbcPersonal.tdbcomprobantes.EOF Or fAbcPersonal.tdbcomprobantes.BOF)
  If Not l_ExistRecord Then s_ParCodigo = fAbcPersonal.tdbcomprobantes.Columns(0).Text
  
  ' Carga los datos en el formulario
  ShowScreen
  
  '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codtic", "destic")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Tipo de Comprobante", 2
  ' Asigno el control de datos  ala grilla
  tdbHelp.DataSource = dcaHelp
  
  ' Recupero la información
  s_Sql = gdl_Funcion.HelpTablas("tic", tdbHelp.Columns(0).DataField, "", "")
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & ps_DataBase, dcaHelp, tdbHelp, s_Sql, adCmdText, adLockReadOnly
  ']
  
  ' Coloco el puntero normal
  gdl_Procedure.PunteroNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' Habilito/desabilito botones inciales
  fAbcPersonal.cmdActioncomprobantes(0).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActioncomprobantes(1).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
  fAbcPersonal.cmdActioncomprobantes(2).Enabled = (fAbcPersonal.Tag = s_MdoData_Upd)
End Sub
Private Sub panTool_Click(Index As Integer)
  Dim n_ToolBar As Byte
  
  n_ToolBar = 0
  ' Ubico los botones en la barra de menu
  gdl_Procedure.panToolPosicion panToolBar(n_ToolBar), panTool, cmdAction, n_IndexTool, Index
  ' Actualiza Indice de Barra Actual
  n_IndexTool = Index

End Sub
Private Sub txtimporte_Validate(Cancel As Boolean)
  txtimporte.Text = IIf(Not IsNumeric(txtimporte.Text), 0, txtimporte.Text)
  If CDec(txtimporte.Text) < 0 Then MsgBox "Factor no puede ser negativo; Verifique", vbInformation: txtimporte.SetFocus: Exit Sub
  txtimporte.Text = FormatNumber(CDec(txtimporte.Text), 2)
End Sub

Private Sub txttipo_GotFocus()
  gdl_Procedure.MarcaGet txttipo
End Sub
Private Sub txttipo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txttipo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txttipo_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_CodEmpresa, txttipo, "CO")
End Sub

Private Sub tdbHelp_DblClick()

  If dcaHelp.Recordset.RecordCount = 0 Or (dcaHelp.Recordset.EOF And dcaHelp.Recordset.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  txttipo.Text = tdbHelp.Columns(0).Value
  lblHelp(0) = tdbHelp.Columns(1).Value
  txttipo.SetFocus

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)

  ' Recupero la información ordenada
  s_Sql = gdl_Funcion.HelpTablas("tic", tdbHelp.Columns(ColIndex).DataField, "", "")
  dcaHelp.RecordSource = s_Sql
  dcaHelp.Refresh

End Sub
Private Sub tdbHelp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or (KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown) Then s_SqlHelp = ""
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaHelp, tdbHelp, ""
End Sub
Private Sub tdbHelp_KeyPress(KeyAscii As Integer)
  Dim n_Columna As Integer
  
  If KeyAscii = vbKeyReturn Then
    tdbHelp_DblClick
  ElseIf (UCase$(Chr$(KeyAscii)) >= "A" And UCase$(Chr$(KeyAscii)) <= "Z") Or _
       (Chr$(KeyAscii) >= "0" And Chr$(KeyAscii) <= "9") Or KeyAscii = 32 Or Chr$(KeyAscii) = "." _
       Or Chr$(KeyAscii) = "*" Then
    If Chr$(KeyAscii) = "*" Then
      s_SqlHelp = ""
    Else
      s_SqlHelp = s_SqlHelp & UCase$(Chr$(KeyAscii))
    End If
    n_Columna = tdbHelp.Col
    s_Sql = gdl_Funcion.HelpTablas("niv", tdbHelp.Columns(n_Columna).DataField, "", s_SqlHelp)
    dcaHelp.RecordSource = s_Sql
    dcaHelp.Refresh
    tdbHelp.Col = n_Columna
  Else
      s_SqlHelp = ""
  End If

End Sub
Private Sub tdbHelp_LostFocus()
  tdbHelp.Visible = False
End Sub



