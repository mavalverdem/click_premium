VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fReporPlanillaGnral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 01"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   Icon            =   "rptplanillagnral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   11490
   Begin MSAdodcLib.Adodc dcaSeleccion 
      Height          =   330
      Index           =   3
      Left            =   45
      Top             =   5460
      Width           =   6000
      _ExtentX        =   10583
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
      Height          =   5190
      Index           =   0
      Left            =   10710
      TabIndex        =   3
      Top             =   585
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   9155
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
         Index           =   6
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   3240
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
         Picture         =   "rptplanillagnral.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Tag             =   "0"
         Top             =   975
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
         Picture         =   "rptplanillagnral.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1410
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
         Picture         =   "rptplanillagnral.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2370
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
         Picture         =   "rptplanillagnral.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2805
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
         Picture         =   "rptplanillagnral.frx":007C
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
         TabIndex        =   4
         Tag             =   "0"
         Top             =   435
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
         Picture         =   "rptplanillagnral.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   7
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
         Picture         =   "rptplanillagnral.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   9
         Left            =   150
         TabIndex        =   21
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
         Picture         =   "rptplanillagnral.frx":00D0
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   22
         Tag             =   "0"
         Top             =   3795
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
         Picture         =   "rptplanillagnral.frx":00EC
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   8
         Left            =   150
         TabIndex        =   23
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
         Picture         =   "rptplanillagnral.frx":0108
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   510
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11490
      _Version        =   65536
      _ExtentX        =   20267
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
      Begin VB.TextBox txtFormato 
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
         Left            =   1485
         MaxLength       =   2
         TabIndex        =   2
         Top             =   105
         Width           =   585
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   300
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Top             =   105
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "..."
      End
      Begin Threed.SSRibbon ribSeccion 
         Height          =   360
         Left            =   10425
         TabIndex        =   12
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
         PictureUp       =   "rptplanillagnral.frx":0124
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
         Left            =   2565
         TabIndex        =   15
         Top             =   150
         Width           =   195
      End
      Begin VB.Label lblDato 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formato Planilla :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1320
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
      Height          =   4845
      Index           =   3
      Left            =   45
      TabIndex        =   13
      Top             =   585
      Width           =   6000
      _ExtentX        =   10583
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
   Begin TabDlg.SSTab tabRegister 
      Height          =   5190
      Left            =   6120
      TabIndex        =   17
      Top             =   585
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   9155
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
      TabPicture(0)   =   "rptplanillagnral.frx":0140
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dcaSeleccion(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdbSeleccion(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ubicación"
      TabPicture(1)   =   "rptplanillagnral.frx":015C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdbSeleccion(1)"
      Tab(1).Control(1)=   "dcaSeleccion(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Sección"
      TabPicture(2)   =   "rptplanillagnral.frx":0178
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tdbSeleccion(2)"
      Tab(2).Control(1)=   "dcaSeleccion(2)"
      Tab(2).ControlCount=   2
      Begin TrueOleDBGrid80.TDBGrid tdbSeleccion 
         Height          =   4365
         Index           =   0
         Left            =   60
         TabIndex        =   18
         Top             =   75
         Width           =   4415
         _ExtentX        =   7779
         _ExtentY        =   7699
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
         Top             =   4470
         Width           =   4415
         _ExtentX        =   7779
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
         Height          =   4365
         Index           =   1
         Left            =   -74940
         TabIndex        =   19
         Top             =   75
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   7699
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
         Top             =   4470
         Width           =   4410
         _ExtentX        =   7779
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
         Height          =   4365
         Index           =   2
         Left            =   -74940
         TabIndex        =   20
         Top             =   75
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   7699
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
         Top             =   4470
         Width           =   4410
         _ExtentX        =   7779
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
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2505
      TabIndex        =   16
      Top             =   210
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
Attribute VB_Name = "fReporPlanillaGnral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                         ' Declarar variable antes de usarla

Private s_TitleWindow As String, s_TitleTable As String ' Titulos de la ventanas y la grilla
Private n_IndexTool As Integer, n_Index As Integer      ' Indice de la barra de herramientas, indice para bucle
Private porstHelp As ADODB.Recordset                    ' Recordset de ayuda
Private n_IndexHelp As Integer, s_SqlHelp As String     ' Indice de la opciones y cadena de ayuda
Private as_SelRegistro(4, 2)                            ' Array de inicio y fin de seleccion de registro
'[
Private Function GeneraCadena(ByVal s_Cadena As String, ByVal s_TituloCelda As String, ByVal s_Expresion As String, ByVal n_Posicion As Integer, ByVal n_Longitud As Integer, ByVal s_Caracter As String, ByVal s_Tipo As String) As String
  Dim nLenCadena As Integer, nLen As Integer
  Dim nPosInicio As Integer
  
  nPosInicio = IIf(Left(s_TituloCelda, 1) = "|" Or Left(s_TituloCelda, 1) = "[", 2, 1)
  If s_Tipo = "D" Then
    s_Expresion = gdl_Funcion.PadR(s_Expresion, n_Longitud, Chr(32))
  Else
    s_Expresion = gdl_Funcion.PadL(s_Expresion, n_Longitud, Chr(32))
  End If
  s_Expresion = Replace(s_Expresion, "[", "", 1, 1, vbTextCompare)
  ' Caracter de separación de campos
  nLenCadena = Len(s_Cadena)
  nLen = nLenCadena - n_Posicion
  If n_Posicion >= nLenCadena Then
    s_Cadena = s_Cadena & String((n_Posicion - nLenCadena), Chr(32))
  Else
    n_Longitud = n_Longitud - nLen
    n_Longitud = IIf(n_Longitud > 0, n_Longitud, 0)
    If s_Tipo = "C" Or s_Tipo = "G" Then
      s_Expresion = Right(s_Expresion, n_Longitud)
    ElseIf s_Tipo = "D" Then
      s_Expresion = Left(s_Expresion, n_Longitud)
    End If
    s_Caracter = IIf(n_Longitud > 0, s_Caracter, "")
  End If
  n_Longitud = n_Longitud - (nPosInicio - 1)
  n_Longitud = IIf(n_Longitud > 0, n_Longitud, 0)
  s_Expresion = Choose(nPosInicio, "", s_Caracter) & IIf(s_Tipo = "D", Left(s_Expresion, n_Longitud), Right(s_Expresion, n_Longitud))
  s_Cadena = s_Cadena & s_Expresion
  GeneraCadena = s_Cadena
    
End Function
Private Sub GeneraPlanilla(ByVal n_Modo As Integer, ByVal n_TabIndex As Integer, ByVal s_Tabla As String, ByVal s_Formato As String, ByVal s_Moneda As String, ByVal a_DatoTitulo, ByVal s_Proceso, ByVal s_FechaHora As String)
  Dim a_Columnas(), a_Totales() As Double
  Dim nNivel As Integer, sValor As String, nValor As Double
  Dim a_sCabecera(7) As String, a_sDetalle(7) As String
  Dim a_Detalle(), sPersonal As String, nSecuencia As Long
  Dim a_Quiebre() As Double, sQuiebre As String, sDesquiebre As String
  Dim sCamRubro As String, sRubro As String, sDesRubro As String
  Dim nGrupo As Integer, sSeccion As String, sCaracter As String
  Dim nRegistros As Long, nRegistro As Long
  
  ' Muestro la Barra
  fMenu.panPercent.Visible = True
  
  sCaracter = Choose(n_Modo, "[", "[", "[")
  ' Agrupacion default
  nGrupo = IIf(ribSeccion.Value, s_Estado_Act, s_Estado_Ina)
  sCamRubro = Choose(n_TabIndex + 1, "codcco", "codubica", "codsec", "codcco")
  sDesRubro = Choose(n_TabIndex + 1, "detcco", "desubica", "dessec", "detcco")
  
  ' Registros detalle con campos
  s_Sql = "SELECT pll.fila, pll.columna, pll.posicion, pll.tipo, pll.alias, pll.descripcion, pll.longitud, "
  s_Sql = s_Sql & "pll.subrayado, pll.sizefont, pll.sizepapel, pll.posipapel,  pll.imprimecab, pll.despll, vfx.valor "
  s_Sql = s_Sql & "FROM plplanilla pll "
  s_Sql = s_Sql & "LEFT JOIN plvarfunc vfx ON pll.tipo=vfx.tipo AND pll.alias=vfx.codigo AND IFNULL(vfx.valor, '')<>'' "
  s_Sql = s_Sql & "WHERE pll.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND pll.codpll='" & s_Formato & "' "
  s_Sql = s_Sql & "ORDER BY pll.fila, pll.columna"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  
  For n_Index = 1 To UBound(a_sCabecera): a_sCabecera(n_Index) = "x": a_sDetalle(n_Index) = "x": Next n_Index
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    n_Index = 0: s_Sql = ""
    nNivel = n_Index
    While Not porstRecordset.EOF
      n_Index = n_Index + 1
      ' Redimensiono el arreglo de las cabeceras
      ReDim Preserve a_Columnas(8, n_Index)
      a_Columnas(1, n_Index) = CInt(porstRecordset("fila"))
      a_Columnas(2, n_Index) = CInt(porstRecordset("posicion"))
      a_Columnas(3, n_Index) = CInt(porstRecordset("longitud"))
      a_Columnas(4, n_Index) = Trim(porstRecordset("tipo"))
      a_Columnas(5, n_Index) = Trim(porstRecordset("alias"))
      a_Columnas(6, n_Index) = UCase(Left(gdl_Funcion.aTexto(porstRecordset("descripcion")), a_Columnas(3, n_Index)))
      a_Columnas(7, n_Index) = Trim(porstRecordset("subrayado"))
      a_Columnas(8, n_Index) = gdl_Funcion.aTexto(porstRecordset("valor"))
      If a_Columnas(8, n_Index) <> "" And a_Columnas(8, n_Index) <> "NULL" Then
        s_Sql = s_Sql & Trim(porstRecordset!valor) & " AS dato_" & Format(n_Index, "00") & ", "
      End If
      
      nNivel = CInt(porstRecordset("fila"))
      a_sCabecera(nNivel) = GeneraCadena(a_sCabecera(nNivel), a_Columnas(6, n_Index), a_Columnas(6, n_Index), a_Columnas(2, n_Index), a_Columnas(3, n_Index), sCaracter, a_Columnas(4, n_Index))
      porstRecordset.MoveNext
    Wend
  End If
  
  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT " & s_Sql
  s_Sql = s_Sql & "pll.codpll, pll.despll, pll.imprimecab, pll.tipo, pll.alias, pll.descripcion, pll.fila, pll.columna, pll.posicion, pll.subrayado, pll.longitud, "
  s_Sql = s_Sql & "res.codpsn, ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)), 2) AS nimporte, "
  s_Sql = s_Sql & Choose(n_TabIndex + 1, "dxr.codcco, cco.detcco ", "dxr.codubica, ubi.desubica ", "dxr.codsec, sec.dessec ", "dxr.codcco, cco.detcco ")
  s_Sql = s_Sql & "FROM plplanilla pll "
  s_Sql = s_Sql & "INNER JOIN plresultado res ON pll.codcls=res.codcls AND pll.alias=res.codcpc AND pll.tipo='C' "
  s_Sql = s_Sql & "INNER JOIN pldatoresultado dxr ON res.codcls=dxr.codcls AND res.codpdo=dxr.codpdo AND res.codpsn=dxr.codpsn "
  s_Sql = s_Sql & "INNER JOIN plasistencia asi ON res.codcls=asi.codcls AND res.codpdo=asi.codpdo AND res.codpsn=asi.codpsn "
  s_Sql = s_Sql & "INNER JOIN plpersonal psn ON res.codcls=psn.codcls AND res.codpsn=psn.codpsn "
  s_Sql = s_Sql & "INNER JOIN plentidadafp afp ON dxr.codafp=afp.codafp "
  s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON dxr.codeps=eps.codeps "
  s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON dxr.codcls=cgo.codcls AND dxr.codcgo=cgo.codcgo "
  s_Sql = s_Sql & "INNER JOIN plubicacion ubi ON dxr.codubica=ubi.codubica "
  s_Sql = s_Sql & "INNER JOIN plseccion sec ON dxr.codsec=sec.codsec "
  s_Sql = s_Sql & "INNER JOIN " & ps_DaBasCon & ".cocco cco ON dxr.codcco=cco.codcco "
  s_Sql = s_Sql & "WHERE pll.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND pll.codpll='" & s_Formato & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  If n_TabIndex <> 3 Then
    s_Sql = s_Sql & "AND dxr." & sCamRubro & " IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & Left(s_Proceso, 9) & n_TabIndex & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  End If
  s_Sql = s_Sql & "GROUP BY " & Choose(n_TabIndex + 1, "dxr.codcco, ", "dxr.codubica, ", "dxr.codsec, ", "") & "res.codpsn, res.codcpc "
  s_Sql = s_Sql & "ORDER BY " & Choose(n_TabIndex + 1, "codcco, ", "codubica, ", "codsec, ", "") & "codpsn, fila, columna"
  Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  nSecuencia = 0
  ' Si hay registros de configuración
  If Not (porstRecordset.EOF And porstRecordset.BOF) Or porstRecordset.RecordCount > 0 Then
    gdl_Conexion.IniciaTransaccion    ' Inicia transacción
    
    nRegistros = porstRecordset.RecordCount
    ' Genero los registros de la tabla de reporte
    ReDim a_Detalle(UBound(a_Columnas, 2))
    ReDim a_Quiebre(UBound(a_Detalle))
    ReDim a_Totales(UBound(a_Quiebre))
    sSeccion = Choose(nGrupo + 1, "codpll", sCamRubro)
    
    a_Campos = Array("seccion", "secuencia", "cabecera1", "cabecera2", "cabecera3", "cabecera4", "cabecera5", "cabecera6", "cabecera7", "detalle1", "detalle2", "detalle3", "detalle4", "detalle5", "detalle6", "detalle7")
    a_Tipos = Array(TipoDato.Caracter, TipoDato.Numero, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter, TipoDato.Caracter)
    While Not porstRecordset.EOF
      sQuiebre = porstRecordset(sSeccion)
      ' Rango - quiebre adicional
      If nGrupo = s_Estado_Act Then
        sDesquiebre = UCase(porstRecordset(sDesRubro))
        For n_Index = 1 To UBound(a_Quiebre): a_Quiebre(n_Index) = 0: Next n_Index
        ' Inserto el titulo del quiebre
        For n_Index = 1 To UBound(a_sDetalle): a_sDetalle(n_Index) = "": Next n_Index
        a_sDetalle(1) = Trim(sQuiebre) & " - " & Trim(sDesquiebre)
        a_Valores = Array("Q", nSecuencia, a_sCabecera(1), a_sCabecera(2), a_sCabecera(3), a_sCabecera(4), a_sCabecera(5), a_sCabecera(6), a_sCabecera(7), a_sDetalle(1), a_sDetalle(2), a_sDetalle(3), a_sDetalle(4), a_sDetalle(5), a_sDetalle(6), a_sDetalle(7))
        If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
      End If
      Do
        ' Inicializo las variables
        nSecuencia = nSecuencia + 1
        sPersonal = porstRecordset("codpsn")
        a_sDetalle(1) = "x" & gdl_Funcion.PadL((nSecuencia), 3, " ")
        For n_Index = 2 To UBound(a_sDetalle): a_sDetalle(n_Index) = "x": Next n_Index
        For n_Index = 1 To UBound(a_Detalle): a_Detalle(n_Index) = 0: Next n_Index
        Do
          a_Detalle(porstRecordset("columna")) = CDec(porstRecordset("nimporte"))
          ' Incremento el porcentaje
          nRegistro = nRegistro + 1
          fMenu.panPercent.FloodPercent = ((nRegistro * 100) \ nRegistros)
          porstRecordset.MoveNext
          If porstRecordset.EOF Then Exit Do
        Loop While sPersonal = porstRecordset("codpsn")
        porstRecordset.MovePrevious
        ' Genero los detalles por columna
        For n_Index = 1 To UBound(a_Columnas, 2)
          nNivel = a_Columnas(1, n_Index)
          ' Registro de detalle
          If a_Columnas(4, n_Index) = "D" Then
            If a_Columnas(8, n_Index) <> "NULL" Then
            
            
              'If n_Index <> 12 Then sValor = gdl_Funcion.PadR(Left(porstRecordset("dato_" & Format(n_Index, "00")), a_Columnas(3, n_Index)), a_Columnas(3, n_Index), Chr(32))
              sValor = gdl_Funcion.PadR(Left(porstRecordset("dato_" & Format(n_Index, "00")), a_Columnas(3, n_Index)), a_Columnas(3, n_Index), Chr(32))
              
            Else
              sValor = gdl_Funcion.PadR("", a_Columnas(3, n_Index), Chr(32))
            End If
          ElseIf a_Columnas(4, n_Index) = "C" Then
            nValor = CDec(a_Detalle(n_Index))
            sValor = IIf(nValor = 0, "", Right(FormatNumber(nValor, 2), a_Columnas(3, n_Index)))
            sValor = gdl_Funcion.PadL(sValor, a_Columnas(3, n_Index), Chr(32))
            a_Totales(n_Index) = a_Totales(n_Index) + CDec(nValor)
            a_Quiebre(n_Index) = a_Quiebre(n_Index) + CDec(nValor)
          ElseIf a_Columnas(4, n_Index) = "G" Then
            nValor = SumDetalle(s_Moneda, s_Formato, s_Proceso, s_FechaHora, sPersonal, a_Columnas(1, n_Index), n_Index)
            sValor = IIf(nValor = 0, "", Right(FormatNumber(nValor, 2), a_Columnas(3, n_Index)))
            sValor = gdl_Funcion.PadL(sValor, a_Columnas(3, n_Index), Chr(32))
            a_Totales(n_Index) = a_Totales(n_Index) + CDec(nValor)
            a_Quiebre(n_Index) = a_Quiebre(n_Index) + CDec(nValor)
          End If
          ' Linea detalle de reporte
          a_sDetalle(nNivel) = GeneraCadena(a_sDetalle(nNivel), a_Columnas(6, n_Index), UCase(sValor), a_Columnas(2, n_Index), a_Columnas(3, n_Index), sCaracter, a_Columnas(4, n_Index))
        Next n_Index
        a_Valores = Array("D", nSecuencia, a_sCabecera(1), a_sCabecera(2), a_sCabecera(3), a_sCabecera(4), a_sCabecera(5), a_sCabecera(6), a_sCabecera(7), a_sDetalle(1), a_sDetalle(2), a_sDetalle(3), a_sDetalle(4), a_sDetalle(5), a_sDetalle(6), a_sDetalle(7))
        If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
        porstRecordset.MoveNext
        ' Fin de archivo
        If porstRecordset.EOF Then Exit Do
      Loop While sQuiebre = porstRecordset(sSeccion)
    Wend
    ' Inserto los totales generales
    a_sDetalle(1) = "xRESUMEN GENERAL :"
    For n_Index = 2 To UBound(a_sDetalle): a_sDetalle(n_Index) = "x": Next n_Index
    For n_Index = 1 To UBound(a_Totales)
      ' Registro de detalle
      nNivel = a_Columnas(1, n_Index)
      If a_Columnas(4, n_Index) = "D" Then
        sValor = gdl_Funcion.PadR("", a_Columnas(3, n_Index), Chr(32))
      Else
        nValor = CDec(a_Totales(n_Index))
        sValor = gdl_Funcion.PadL(Right(FormatNumber(nValor, 2), a_Columnas(3, n_Index)), a_Columnas(3, n_Index), Chr(32))
      End If
      ' Linea detalle de reporte
      a_sDetalle(nNivel) = GeneraCadena(a_sDetalle(nNivel), a_Columnas(6, n_Index), UCase(sValor), a_Columnas(2, n_Index), a_Columnas(3, n_Index), IIf(a_Columnas(4, n_Index) = "D", "", sCaracter), a_Columnas(4, n_Index))
    Next n_Index
    a_Valores = Array("P", nSecuencia + 1, a_sCabecera(1), a_sCabecera(2), a_sCabecera(3), a_sCabecera(4), a_sCabecera(5), a_sCabecera(6), a_sCabecera(7), a_sDetalle(1), a_sDetalle(2), a_sDetalle(3), a_sDetalle(4), a_sDetalle(5), a_sDetalle(6), a_sDetalle(7))
    If Not Records_Ins(s_Tabla, a_Campos, a_Valores, a_Tipos) Then GoTo Error
    gdl_Conexion.ConfirmaTransaccion ' Confirma transacción
  End If
  GoTo Finalizar

Error:
  gdl_Conexion.CancelaTransaccion
Finalizar:
  ' Reinicializo los mensajes
  fMenu.panPercent.FloodPercent = 0
  fMenu.panPercent.Visible = False

End Sub
Private Sub PrinterDetalle(ByVal nSizeFont As Double, ByVal sFont As String, ByVal sBold As String, ByVal sSubrayado As String, ByVal sCursiva As String, ByVal aDetalle, ByRef nFila As Double, ByRef nColumna As Double)
  Dim nDetalle As Integer
  
  Printer.Font = sFont
  Printer.FontSize = CDec(nSizeFont)
  Printer.Font.Bold = (sBold = s_Estado_Act)
  Printer.Font.Underline = (sSubrayado = s_Estado_Act)
  Printer.Font.Italic = (sCursiva = s_Estado_Act)
  Printer.CurrentY = nFila
  Printer.CurrentX = nColumna
  For nDetalle = 1 To UBound(aDetalle)
    If Mid(aDetalle(nDetalle), 2) <> "" Then
      Printer.Print Mid(aDetalle(nDetalle), 2)
    End If
  Next nDetalle

End Sub
Private Sub RecuperaRegistros(ByVal nIndex As Integer, ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  If nIndex = 0 Then
    s_Sql = "SELECT codcco, detcco, estcco "
    s_Sql = s_Sql & "FROM cocco "
    s_Sql = s_Sql & "WHERE LENGTH(codcco)>=" & pn_NivelCenCosto & " "
  ElseIf nIndex = 1 Then
    s_Sql = "SELECT codubica, desubica, estadoubica "
    s_Sql = s_Sql & "FROM plubicacion "
  ElseIf nIndex = 2 Then
    s_Sql = "SELECT codsec, dessec, estadosec "
    s_Sql = s_Sql & "FROM plseccion "
  ElseIf nIndex = 3 Then
    s_Sql = "SELECT codpdo, despdo, fechaini, fechafin, anopdo, mespdo, tpopdo, estadopdo "
    s_Sql = s_Sql & "FROM plperiodo "
    s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND anopdo='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND estadopdo='" & s_Estado_Blq & "' "
  End If
  s_Sql = s_Sql & "ORDER BY " & s_Orden
  gdl_Procedure.SeteaAdoControl ps_StrgConnec & IIf(nIndex = 0, ps_DaBasCon, ps_DataBase), dcaSeleccion(nIndex), tdbSeleccion(nIndex), s_Sql, adCmdText, adLockReadOnly

  ' Inicializo los rangos de impresion
  as_SelRegistro(nIndex, 0) = "": as_SelRegistro(nIndex, 1) = ""
  If dcaSeleccion(nIndex).Recordset.RecordCount > 0 Then
    dcaSeleccion(nIndex).Recordset.MoveLast: as_SelRegistro(nIndex, 1) = dcaSeleccion(nIndex).Recordset.Bookmark
    dcaSeleccion(nIndex).Recordset.MoveFirst: as_SelRegistro(nIndex, 0) = dcaSeleccion(nIndex).Recordset.Bookmark
  End If

End Sub
Private Function SumDetalle(ByVal s_Moneda As String, ByVal s_Formato As String, ByVal s_Proceso As String, ByVal s_FechaHora As String, ByVal s_Personal As String, ByVal n_Fila As Integer, ByVal n_Columna As Integer) As Double
  Dim porstResumen As ADODB.Recordset

  s_Sql = "SELECT res.codpsn, ROUND(SUM(IFNULL(res.importe_" & IIf(s_Moneda = s_Codmon_mn, "mn", "me") & ", 0)), 2) AS importe "
  s_Sql = s_Sql & "FROM plresultado res "
  s_Sql = s_Sql & "INNER JOIN pldetaplanilla det ON res.codcls=det.codcls AND res.codcpc=det.codcpc "
  s_Sql = s_Sql & "WHERE res.codcls='" & ps_ClsPlanilla & "' "
  s_Sql = s_Sql & "AND res.codpsn='" & s_Personal & "' "
  s_Sql = s_Sql & "AND res.codpdo IN(SELECT valor FROM rangoimpresion "
  s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
  s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
  s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
  s_Sql = s_Sql & "AND det.codpll='" & s_Formato & "' "
  s_Sql = s_Sql & "AND det.fila=" & n_Fila & " "
  s_Sql = s_Sql & "AND det.columna=" & n_Columna & " "
  s_Sql = s_Sql & "GROUP BY codpsn"
  Set porstResumen = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  If Not (porstResumen.EOF And porstResumen.BOF) Or porstResumen.RecordCount > 0 Then
    SumDetalle = CDec(porstResumen!importe)
  End If
  ' Elimino el objeto de memeoria
  porstResumen.Close
  Set porstResumen = Nothing

End Function
Private Sub cmdAction_Click(Index As Integer)
  Dim s_Moneda As String, s_Reporte As String
  Dim a_Parametro(5) As String, s_FechaHora As String
  Dim s_Proceso As String, s_OldMessage As String
  Dim nTabIndex As Integer
  
  ' Inicializo el modo de registro o selección
  nTabIndex = IIf(ribSeccion.Value, tabRegister.Tab, 3)
  Me.Tag = ""
  Select Case Index
   Case 0 ' Configuración de formatos
    fPlanillaGnral.Show
   Case 1, 2  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros nTabIndex, tdbSeleccion(nTabIndex).Columns(tdbSeleccion(nTabIndex).Col).DataField & Choose(Index, " ASC", " DESC")
   Case 3 ' Busqueda de registro
    If Not (dcaSeleccion(nTabIndex).Recordset.EOF Or dcaSeleccion(nTabIndex).Recordset.BOF) Then
      Set go_tdbBusqueda = tdbSeleccion(nTabIndex)
      Set go_dcaBusqueda = dcaSeleccion(nTabIndex)
      gn_ColBusqueda = (tdbSeleccion(nTabIndex).Columns.Count - 1)
      fBusqueda.Show vbModal
    End If
   Case 4, 5, 6 ' Selecciono rango de impresión
    gdl_Procedure.MarcaRegistros dcaSeleccion(nTabIndex), tdbSeleccion(nTabIndex), as_SelRegistro(nTabIndex, 0), as_SelRegistro(nTabIndex, 1), (Index - 4), tdbSeleccion(nTabIndex).Caption
   Case 7, 8, 9 ' Opciones de impresión
    ' Verifico que existan registros y seleccionados
    If tdbSeleccion(3).SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión - " & tdbSeleccion(3).Caption, vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión - " & tdbSeleccion(nTabIndex).Caption, vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 1 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión - " & tdbSeleccion(nTabIndex).Caption, vbExclamation: Exit Sub
    If tdbSeleccion(nTabIndex).SelBookmarks.Count = 0 And nTabIndex = 2 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión - " & tdbSeleccion(nTabIndex).Caption, vbExclamation: Exit Sub
    If txtFormato = "" Then Beep: MsgBox "Debe Ingresar el Formato de la planilla", vbExclamation: txtFormato.SetFocus: Exit Sub
    If (lblHelp(0) = "" Or lblHelp(0) = "???") Then Beep: MsgBox "Formato de la planilla no existe; verifique", vbExclamation: txtFormato.SetFocus: Exit Sub
    s_Proceso = "planignral"
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Información ..."
    ' Coloco el puntero en espera
    gdl_Procedure.PunteroEnEspera
    
    '[ Inicio la conexión a la base de datos ]
    ps_StrgConnec = OpenConnection(ps_Servidor, ps_DataBase)
    
    nTabIndex = 3
    s_Reporte = dcaSeleccion(nTabIndex).Recordset!mespdo
    ' Barro el arreglo de registros (periodos) marcados (bookmarks)
    For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
      tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
      If (s_Reporte <> dcaSeleccion(nTabIndex).Recordset!mespdo) Then Beep: MsgBox "Los periodos son de diferentes meses; verifique", vbExclamation: GoTo Finaliza
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    nTabIndex = IIf(ribSeccion.Value, tabRegister.Tab, nTabIndex)
    If nTabIndex <> 3 Then
      ' Barro el arreglo de registros marcadas (bookmarks)
      For n_Index = 0 To tdbSeleccion(nTabIndex).SelBookmarks.Count - 1
        tdbSeleccion(nTabIndex).Bookmark = tdbSeleccion(nTabIndex).SelBookmarks(n_Index)
        gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, Left(s_Proceso, 9) & nTabIndex, tdbSeleccion(nTabIndex).Columns(0).Text, ps_Usuario, s_FechaHora, "A"
      Next n_Index
    End If
    
    ' Obtengo los datos de la empresa
    s_Reporte = "ah1"
     a_Parametro(2) = "":  a_Parametro(1) = ""
     
    'MYSQL 4
    s_Sql = "SELECT DISTINCT via.abrevia, prm.direccionvia, prm.numerodir, zon.abrezona, prm.direccionzona, prm.ubigeodir, prm.regpatronal, "
    s_Sql = s_Sql & "pdo.despdo, pdo.anopdo, pdo.mespdo, pll.imprimecab, pll.sizepapel, pll.posipapel, pll.sizefont "
    s_Sql = s_Sql & "FROM plcfgempresa prm, plperiodo pdo, plplanilla pll "
    s_Sql = s_Sql & "LEFT JOIN pltipovia via ON prm.codvia=via.codvia "
    s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON prm.codzona=zon.codzona "
    s_Sql = s_Sql & "WHERE prm.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND pdo.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdo.codpdo IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    s_Sql = s_Sql & "AND pdo.anopdo=prm.pdoano "
    s_Sql = s_Sql & "AND pll.codcls=pdo.codcls "
    s_Sql = s_Sql & "AND pll.codpll='" & txtFormato.Text & "'"
    
    'MYSQL 5
    s_Sql = "SELECT DISTINCT via.abrevia, prm.direccionvia, prm.numerodir, zon.abrezona, prm.direccionzona, prm.ubigeodir, prm.regpatronal, "
    s_Sql = s_Sql & "pdo.despdo, pdo.anopdo, pdo.mespdo, pll.imprimecab, pll.sizepapel, pll.posipapel, pll.sizefont "
    s_Sql = s_Sql & "FROM plcfgempresa prm, plperiodo pdo, plplanilla pll,pltipovia via,pltipozona zon "
    s_Sql = s_Sql & "WHERE prm.codvia=via.codvia and prm.codzona=zon.codzona and  prm.pdoano='" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND pdo.codcls='" & ps_ClsPlanilla & "' "
    s_Sql = s_Sql & "AND pdo.codpdo IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & s_Proceso & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    s_Sql = s_Sql & "AND pdo.anopdo=prm.pdoano "
    s_Sql = s_Sql & "AND pll.codcls=pdo.codcls "
    s_Sql = s_Sql & "AND pll.codpll='" & txtFormato.Text & "'"
    
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      a_Parametro(1) = gdl_Funcion.aTexto(porstRecordset!regpatronal)
      a_Parametro(2) = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      a_Parametro(2) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, a_Parametro(2), "UB")
      a_Parametro(2) = gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!direccionzona) & " - " & a_Parametro(2)
      'a_Parametro(4) = "MES DE " & UCase(gdl_Funcion.NombreMes(porstRecordset!mespdo)) & " DEL " & ps_Anyo
    
      ' Si seleccion de periodo es unica
      If tdbSeleccion(3).SelBookmarks.Count = 1 Then
        n_Index = InStr(1, porstRecordset!despdo, " ")
        n_Index = IIf(n_Index = 0, Len(Trim(porstRecordset!despdo)), n_Index - 1)
        'a_Parametro(4) = "MES DE " & UCase(Mid(porstRecordset!despdo, 1, n_Index)) & " DE " & ps_Anyo
      End If
      'REVISAR ????? 27/06/2008 JGOMEZ
      'a_Parametro(4) = "SEMANA " & porstRecordset!despdo & " DEL " & ps_Anyo
      a_Parametro(4) = porstRecordset!despdo & " DEL " & ps_Anyo
          
      a_Parametro(5) = gdl_Funcion.aTexto(porstRecordset!imprimecab)
      s_Reporte = Choose(porstRecordset!sizepapel + 1, "a", "b", "c")
      s_Reporte = s_Reporte & Choose(porstRecordset!posipapel, "v", "h")
      s_Reporte = s_Reporte & IIf(porstRecordset!sizefont = 6, "1", IIf(porstRecordset!sizefont = 8, "2", "3"))
    End If
    
    porstRecordset.Close
    
    s_Moneda = IIf(fMenu.ribMoneda(0).Value, s_Codmon_mn, s_Codmon_me)
    a_Parametro(3) = "P L A N I L L A  D E  P A G O  D E  R E M U N E R A C I O N E S"
    a_Parametro(5) = Choose(Index - 6, s_Estado_Act, s_Estado_Act, a_Parametro(5))
    ' Parametros de Impresión
    gdl_Procedure.ps_ReportTitle = "PLANILLA DE REMUNERACIONES"
    gdl_Procedure.ps_ReportName = "rptplagnral" & s_Reporte
    
    ReDim aElemento(3, 8): ReDim aElementos(2)
    ' Parametros del Reporte
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = tdbSeleccion(nTabIndex).Columns(0).DataField & " ASC"
    aElemento(0, 2) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = "":  aElemento(1, 2) = ""
    aElemento(1, 3) = "": aElemento(1, 4) = "":  aElemento(1, 5) = ""
    aElemento(1, 6) = "": aElemento(1, 7) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & a_Parametro(3) & ";true"
    aElemento(2, 2) = "Direccion;" & a_Parametro(2) & ";true"
    'aElemento(2, 3) = "TituloPeriodo;" & Left(a_Parametro(4), 6) & ";true"
    aElemento(2, 3) = "TituloPeriodo;" & a_Parametro(4) & ";true"
    aElemento(2, 4) = "Periodo;" & Mid(a_Parametro(4), 8) & ";true"
    aElemento(2, 5) = "RegPatronal;" & a_Parametro(1) & ";true"
    aElemento(2, 6) = "NroRuc;" & ps_RucEmpresa & ";true"
    
    aElemento(2, 7) = "Imprimecab;" & a_Parametro(5) & ";true"
    
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Conexion.Execucion s_Sql, Elimina
    
    s_Sql = "CREATE TEMPORARY TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ( "
    s_Sql = s_Sql & "secuencia smallint(5) Not Null, "
    s_Sql = s_Sql & "seccion char(1) Not  Null, "
    s_Sql = s_Sql & "cabecera1 varchar(255) Null, "
    s_Sql = s_Sql & "cabecera2 varchar(255) Null, "
    s_Sql = s_Sql & "cabecera3 varchar(255) Null, "
    s_Sql = s_Sql & "cabecera4 varchar(255) Null, "
    s_Sql = s_Sql & "cabecera5 varchar(255) Null, "
    s_Sql = s_Sql & "cabecera6 varchar(255) Null, "
    s_Sql = s_Sql & "cabecera7 varchar(255) Null, "
    s_Sql = s_Sql & "detalle1 varchar(255) Null, "
    s_Sql = s_Sql & "detalle2 varchar(255) Null, "
    s_Sql = s_Sql & "detalle3 varchar(255) Null, "
    s_Sql = s_Sql & "detalle4 varchar(255) Null, "
    s_Sql = s_Sql & "detalle5 varchar(255) Null, "
    s_Sql = s_Sql & "detalle6 varchar(255) Null, "
    s_Sql = s_Sql & "detalle7 varchar(255) Null, "
    s_Sql = s_Sql & "PRIMARY KEY (secuencia, seccion)) "
    gdl_Conexion.Execucion s_Sql, Inserta
    
    ' Genera y selecciono la información del reporte
    GeneraPlanilla (Index - 6), nTabIndex, "tmp" & gdl_Procedure.ps_ReportName, txtFormato.Text, s_Moneda, a_Parametro, s_Proceso, s_FechaHora
    s_Sql = "SELECT * "
    s_Sql = s_Sql & "FROM  tmp" & gdl_Procedure.ps_ReportName & " "
    s_Sql = s_Sql & "ORDER BY secuencia, seccion"
    Set porstRecordset = gdl_Conexion.Recordset(adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, IIf(Index = 5, 1, 0), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino la tabla temporal y el rango de impresion
    s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    gdl_Conexion.Execucion s_Sql, Elimina
Finaliza:
    ' Elimino el rango de impresion
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_Proceso, "", ps_Usuario, s_FechaHora, "E"
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, Left(s_Proceso, 9) & nTabIndex, "", ps_Usuario, s_FechaHora, "E"
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
    '[ Finalizo la conexión a la base de datos ]
    Set gdl_Conexion = Nothing
    ' Coloco el puntero en normal
    gdl_Procedure.PunteroNormal
    ' ]
  End Select

End Sub
Private Sub cmdHelp_Click(Index As Integer)

  s_SqlHelp = ""
  Select Case Index
   Case 0     ' Periodo de pago bloqueados
    tdbHelp.Columns(0).DataField = "codpll": tdbHelp.Columns(1).DataField = "despll"
    tdbHelp.Caption = "Formato Planilla"
    ' Recupero la información
    s_Sql = gdl_Funcion.HelpTablas("pll", "codpll", ps_ClsPlanilla, "")
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
  fMenu.cmbejercicio.Enabled = False
End Sub
Private Sub Form_Load()

  Dim Item As New ValueItem

  ' Establece posición del formulario
  Me.Height = 6330: Me.Width = 11580
  Me.Left = 110: Me.Top = 180
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Titulo del formulario y la Grilla
  s_TitleWindow = Me.Caption
  s_TitleTable = "Periodos de Pago"
  
  ReDim aElemento(5, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Inicio", "Final", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpdo", "despdo", "fechaini", "fechafin", "estadopdo")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 850, 2350, 950, 950, 300)
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
    aElemento(n_Index, 1) = Choose(n_Index + 1, "promedio", "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "imprimir", "imprapid")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Configuración de Formato", "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir", "Imprime Detalle")
    aElemento(n_Index, 3) = Choose(n_Index + 1, "&c", "&a", "&d", "&b", "&p", "&f", "&r", "&v", "&i")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
 '[ Configuración el control de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codpll", "despll")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Entidad Pensiones", 2
  ']
  
  ' [ Configuros grillas de selección
  ReDim aElemento(3, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codcco", "detcco", "estcco")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 700, 2800, 300)
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
  
  ' Cargo el grafico del boton de seccion
  ribSeccion.PictureUp = LoadPicture()
  ribSeccion.ToolTipText = "Seccion de Centro de Costos"
  s_Sql = gdl_Procedure.ps_PathImagen & "dividir.bmp"
  If gdl_Funcion.ExisteArchivo(s_Sql) Then ribSeccion.PictureUp = LoadPicture(s_Sql)
  ribSeccion.Value = False
  
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  ' Recupero los registros con el control de datos asignado (orden)
  For n_Index = 0 To 3
    tdbSeleccion(n_Index).DataSource = dcaSeleccion(n_Index)
    RecuperaRegistros n_Index, tdbSeleccion(n_Index).Columns(0).DataField & " ASC"
  Next n_Index

End Sub
Private Sub Form_Unload(Cancel As Integer)
  
  If FormVisible("fPlanillaGnral") Then
    Beep
    MsgBox "Primero debe cerrar la Pantalla de " & fPlanillaGnral.Caption, vbExclamation
    Cancel = True
  End If
  If Not Cancel Then
    If porstHelp.State = adStateOpen Then porstHelp.Close
    Set porstHelp = Nothing
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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Formato de planilla
    txtFormato = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtFormato.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Formato de planilla
    s_Sql = gdl_Funcion.HelpTablas("pll", tdbHelp.Columns(ColIndex).DataField, ps_ClsPlanilla, "")
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
Private Sub tdbSeleccion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then gdl_Procedure.RefreshAdoControl dcaSeleccion(Index), tdbSeleccion(Index), " " & tdbSeleccion(Index).Caption
End Sub
Private Sub txtFormato_GotFocus()
  gdl_Procedure.MarcaGet txtFormato
End Sub
Private Sub txtFormato_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtFormato_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtFormato_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DataBase, ps_ClsPlanilla, txtFormato, "PL")
End Sub
