VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form fSelPersonalCst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro - 00"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "selpersoxcst.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   7845
   Begin TrueOleDBGrid80.TDBGrid tdbRegistro 
      Height          =   4845
      Left            =   45
      TabIndex        =   15
      Top             =   975
      Width           =   7000
      _ExtentX        =   12356
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
      Top             =   5880
      Width           =   7000
      _ExtentX        =   12356
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
      Left            =   7065
      TabIndex        =   5
      Top             =   975
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
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1140
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
         Picture         =   "selpersoxcst.frx":000C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1560
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
         Picture         =   "selpersoxcst.frx":0028
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   3
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2310
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
         Picture         =   "selpersoxcst.frx":0044
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Tag             =   "0"
         Top             =   2745
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
         Picture         =   "selpersoxcst.frx":0060
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   6
         Left            =   150
         TabIndex        =   12
         Tag             =   "0"
         Top             =   3900
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
         Picture         =   "selpersoxcst.frx":007C
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   7
         Left            =   150
         TabIndex        =   13
         Tag             =   "0"
         Top             =   4335
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
         Picture         =   "selpersoxcst.frx":0098
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   6
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
         Picture         =   "selpersoxcst.frx":00B4
      End
      Begin Threed.SSCommand cmdAction 
         Height          =   360
         Index           =   5
         Left            =   150
         TabIndex        =   11
         Tag             =   "0"
         Top             =   3165
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
         Picture         =   "selpersoxcst.frx":00D0
      End
   End
   Begin Threed.SSPanel panToolBar 
      Align           =   1  'Align Top
      Height          =   930
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7845
      _Version        =   65536
      _ExtentX        =   13838
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Left            =   1890
         TabIndex        =   4
         Top             =   495
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   137625601
         CurrentDate     =   38675
      End
      Begin VB.ComboBox cmbPeriodo 
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
         ItemData        =   "selpersoxcst.frx":00EC
         Left            =   1890
         List            =   "selpersoxcst.frx":00EE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   495
         Width           =   1980
      End
      Begin VB.TextBox txtCenCosto 
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   1890
         TabIndex        =   1
         Top             =   150
         Width           =   810
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   1
         Left            =   6705
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
         PictureUp       =   "selpersoxcst.frx":00F0
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   0
         Left            =   6300
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
         PictureUp       =   "selpersoxcst.frx":010C
      End
      Begin Threed.SSRibbon ribParametro 
         Height          =   360
         Index           =   2
         Left            =   7110
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
         PictureUp       =   "selpersoxcst.frx":0128
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   20
         Top             =   150
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
         Caption         =   "Periodo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   2
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro Costo :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   0
         Top             =   195
         Width           =   1005
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
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   21
         Top             =   195
         Width           =   180
      End
      Begin VB.Shape shpCuadro 
         BorderColor     =   &H00C00000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   0
         Left            =   705
         Shape           =   4  'Rounded Rectangle
         Top             =   75
         Width           =   4935
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdbHelp 
      Height          =   2400
      Left            =   2880
      TabIndex        =   22
      Top             =   825
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
Attribute VB_Name = "fSelPersonalCst"
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
Private s_OptRegistro As String                         ' Instancia del formulario activo
'[
Private Sub RecuperaRegistros(ByVal s_Orden As String)

  ' Cadenas de Texto, Recuperar Información
  s_Sql = "SELECT codcls, codpsn, apepaterno, apematerno, nombres, "
  s_Sql = s_Sql & "CONCAT(IFNULL(apepaterno, ''), ' ', IFNULL(apematerno, ''), ', ', IFNULL(nombres, '')) AS nombrepsn, "
  s_Sql = s_Sql & "fecnacimiento, ubigeonac, nacionalidad, naciextrapsn, sexopsn, "
  s_Sql = s_Sql & "refedirec, codvia, nomviadirec, numerdirec, "
  s_Sql = s_Sql & "intedirec, codzona, nomzondirec, ubigeodir, "
  s_Sql = s_Sql & "estcivilpsn, numhijo, numdepen, coddci, numdociden, "
  s_Sql = s_Sql & "numdocmil, telefono, celular, dctojudicial, pordsctojudi, fotopsn, "
  s_Sql = s_Sql & "fecingreso, codtpt, codcgo, cgoconfianza, codpfs, "
  s_Sql = s_Sql & "codcco, codafp, numeroafp, pagodolar, codbcopago, "
  s_Sql = s_Sql & "cuentapago, ctsdolar, codbcocts, cuentacts, codeps, "
  s_Sql = s_Sql & "regpension, fecingregpen, essvida, cobsctr, afilsindical, "
  s_Sql = s_Sql & "remintegralgrati, remuneta, netocpc, variacpc, imporemuneto, "
  s_Sql = s_Sql & "fecbaja, estadopsn "
  s_Sql = s_Sql & "FROM plpersonal "
  s_Sql = s_Sql & "WHERE codcls='" & ps_ClsPlanilla & "' "
  If Trim(txtCenCosto.Text) <> "" Then
    s_Sql = s_Sql & "AND codcco='" & Trim(txtCenCosto.Text) & "' "
  End If
  If Not ribParametro(0).Value Then
    s_Sql = s_Sql & "AND estadopsn" & IIf(ribParametro(1).Value, "<>'I' ", "='I' ")
  End If
  ' Vacaciones
  If s_OptRegistro = "op3" Then
    s_Sql = s_Sql & "AND DATE_FORMAT(fecingreso, '%Y')<'" & ps_Anyo & "' "
    s_Sql = s_Sql & "AND DATE_FORMAT(fecingreso, '%m')='" & Left(cmbPeriodo, 2) & "' "
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
Private Sub cmbPeriodo_LostFocus()
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
End Sub
Private Sub cmdAction_Click(Index As Integer)
  Dim s_FechaHora As String, s_OldMessage As String
  Dim sDireccion As String, sRegPatronal  As String
  Dim sBasico As String
  Dim sComision As String
  
  ' Verifico que Existan Registros
  If (dcaRegistro.Recordset.EOF Or dcaRegistro.Recordset.BOF) Or (dcaRegistro.Recordset.RecordCount = 0) Then Beep: MsgBox "No Existen " & s_TitleTable, vbExclamation: Exit Sub
  ' Inicializo el modo de registro o selección
  Select Case Index
   Case 0, 1  ' Ordena registro ascendentemente o descendentemente
    RecuperaRegistros tdbRegistro.Columns(tdbRegistro.Col).DataField & Choose(Index, " ASC", " DESC")
   Case 2 ' Busqueda de registro
    Set go_tdbBusqueda = tdbRegistro
    Set go_dcaBusqueda = dcaRegistro
    gn_ColBusqueda = (tdbRegistro.Columns.Count - 1)
    fBusqueda.Show vbModal
   Case 3, 4, 5 ' Selecciono rango de impresión
    gdl_Procedure.MarcaRegistros dcaRegistro, tdbRegistro, as_SelRegistro(0), as_SelRegistro(1), (Index - 4), s_TitleTable
   Case 6, 7  ' Opciones de impresión
    ' Verifico que existan registros seleccionados
    If tdbRegistro.SelBookmarks.Count = 0 Then Beep: MsgBox "Debe Seleccionar Rango de Impresión", vbExclamation: Exit Sub
    If lblHelp(0) = "???" Then Beep: MsgBox "Centro de Costo no es valido; Verificar", vbExclamation: txtCenCosto.SetFocus: Exit Sub
    If cmbPeriodo = "" And s_OptRegistro = "op3" Then Beep: MsgBox "Debe Ingresar el Periodo de Analisis", vbExclamation: cmbPeriodo.SetFocus: Exit Sub
    s_FechaHora = Format(Now, s_FmtFeHoMysql_0)
    
    ' Cambio el Mensaje y Muestro la Barra
    s_OldMessage = fMenu.panMessage.Caption
    MuestraMensaje "Generando Información ..."
    
    ' Obtengo los datos de la empresa
    sDireccion = "": sRegPatronal = ""
    s_Sql = "SELECT via.abrevia, cfg.direccionvia, cfg.numerodir, zon.abrezona, cfg.direccionzona, cfg.ubigeodir, cfg.regpatronal, prm.cpcbasico,prm.cpccomisi "
    s_Sql = s_Sql & "FROM plcfgempresa cfg "
    s_Sql = s_Sql & "LEFT JOIN pltipovia via ON cfg.codvia=via.codvia "
    s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON cfg.codzona=zon.codzona "
    s_Sql = s_Sql & "LEFT JOIN plparametroafp prm ON cfg.pdoano=prm.pdoano "
    s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "'"
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    If Not (porstRecordset.BOF And porstRecordset.BOF) Then
      sRegPatronal = gdl_Funcion.aTexto(porstRecordset!regpatronal)
      sDireccion = gdl_Funcion.aTexto(porstRecordset!ubigeodir)
      sDireccion = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_BDSystems, s_Estado_Blq, sDireccion, "UB")
      sDireccion = gdl_Funcion.aTexto(porstRecordset!abrevia) & " " & gdl_Funcion.aTexto(porstRecordset!direccionvia) & " Nº " & gdl_Funcion.aTexto(porstRecordset!numerodir) & " " & gdl_Funcion.aTexto(porstRecordset!abrezona) & " " & gdl_Funcion.aTexto(porstRecordset!direccionzona) & " - " & sDireccion
      sBasico = gdl_Funcion.aTexto(porstRecordset!cpcbasico)
      sComision = gdl_Funcion.aTexto(porstRecordset!cpccomisi)
    End If
    porstRecordset.Close
    
    ' Barro el arreglo de registros marcadas (bookmarks)
    For n_Index = 0 To tdbRegistro.SelBookmarks.Count - 1
      tdbRegistro.Bookmark = tdbRegistro.SelBookmarks(n_Index)
      gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, tdbRegistro.Columns(0).Text, ps_Usuario, s_FechaHora, "A"
    Next n_Index
    
    ' Parametros de Impresión
    n_Index = Right(s_OptRegistro, 1)
    gdl_Procedure.ps_ReportTitle = Choose(n_Index, "Padrón General de Personal", "Datos de Trabajo", "Rol de Vacaciones - " & Trim(Mid(cmbPeriodo, 5)), "Remuneraciones Default", "Experiencia Laboral", "Estudios Realizados", "Datos Familiares", "Vencimiento de Contratos - " & Trim(dtpFecha), "Ficha de Datos")
    gdl_Procedure.ps_ReportName = Choose(n_Index, "lstpadronpsn", "lstdatrabapsn", "lstrolvacapsn", "lstremunepsn", "lstexplabopsn", "lstestudiopsn", "lstdatofamipsn", "lstcontratopsn", "lstfichapsn")
    ReDim aElemento(3, 5): ReDim aElementos(2)
    ' Parametros del store procedure
    aElemento(0, 0) = ps_CodEmpresa
    aElemento(0, 1) = "": aElemento(0, 2) = ""
    aElemento(0, 3) = "": aElemento(0, 4) = ""
    ' Formulas del Reporte
    aElemento(1, 0) = "": aElemento(1, 1) = ""
    ' Campos de Parametros del Reporte
    aElemento(2, 0) = "NombreEmpresa;" & ps_NomEmpresa & ";true"
    aElemento(2, 1) = "TituloReporte;" & UCase(gdl_Procedure.ps_ReportTitle) & ";true"
    aElemento(2, 2) = "Direccion;" & sDireccion & ";true"
    aElemento(2, 3) = "NroRuc;" & ps_RucEmpresa & ";true"
    aElemento(2, 4) = "RegPatronal;" & sRegPatronal & ";true"
    ' Filtro de Formulas y Grupos del Reporte
    aElementos(0) = "": aElementos(1) = ""
    
    ' [ Generación e impresión de información para el reporte
    s_Sql = "SELECT psn.codpsn, "
    s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS apellidopsn, "
    s_Sql = s_Sql & "IFNULL(psn.nombres, '') AS nombrepsn, "
    If s_OptRegistro = "op1" Then               ' Padron de trabajadores
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, cgo.descgo, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(via.abrevia, '')), ' ', TRIM(IFNULL(psn.nomviadirec, '')), ' ', TRIM(IFNULL(psn.numerdirec, '')), ' - ', TRIM(IFNULL(psn.intedirec, '')), ' ', TRIM(IFNULL(zon.abrezona, '')), ' ', TRIM(IFNULL(psn.nomzondirec, ''))) AS direccion, "
      s_Sql = s_Sql & "TRIM(IFNULL(ubg.desubg, '')) AS distrito, "
      s_Sql = s_Sql & "psn.nacionalidad, psn.fecnacimiento, IF(psn.sexopsn='0', 'M', 'F') AS sexopsn, psn.numdociden, psn.estcivilpsn, "
      s_Sql = s_Sql & "afp.desafp, psn.numeroafp, eps.deseps, psn.nroessalud, psn.estadopsn "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "LEFT JOIN " & ps_BDSystems & ".tgubigeo ubg ON psn.ubigeodir=ubg.codubg AND nivelubg='" & s_Estado_Blq & "' "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    ElseIf s_OptRegistro = "op2" Then           ' Datos de trabajo
      s_Sql = s_Sql & "psn.fecingreso, psn.fecbaja, cgo.descgo, "
      s_Sql = s_Sql & "afp.desafp, psn.numeroafp, eps.deseps, psn.nroessalud, psn.estadopsn, cco.detcco, "
      s_Sql = s_Sql & "IF(psn.pagodolar='" & s_Estado_Act & "', '" & s_Codmon_me_Txt & "', '" & s_Codmon_mn_Txt & "') AS monpago, psn.codbcopago, psn.cuentapago, "
      s_Sql = s_Sql & "IF(psn.ctsdolar='" & s_Estado_Act & "', '" & s_Codmon_me_Txt & "', '" & s_Codmon_mn_Txt & "') AS moncts, psn.codbcocts, psn.cuentacts, psn.ctsdeposito, "
      s_Sql = s_Sql & "psn.regpension, psn.essvida, psn.cobsctr, psn.afilsindical, psn.codtpt, "
      s_Sql = s_Sql & "psn.cgoconfianza, psn.codpfs "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "LEFT JOIN " & ps_DaBasCon & ".cocco cco ON psn.codcco=cco.codcco "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    ElseIf s_OptRegistro = "op3" Then           ' Rol de vacaciones
      s_Sql = s_Sql & "psn.fecingreso, psn.estadopsn, CONVERT(CONCAT('" & ps_Anyo & "-', DATE_FORMAT(fecingreso, '-%m-%d')), date) AS fecvacacion "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND DATE_FORMAT(fecingreso, '%Y')<'" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND DATE_FORMAT(fecingreso, '%m')='" & Left(cmbPeriodo, 2) & "' "
    ElseIf s_OptRegistro = "op4" Then           ' Remuneraciones
      s_Sql = s_Sql & "psn.fecingreso, cgo.descgo, psn.estadopsn, "
      s_Sql = s_Sql & "IF(psn.pagodolar='" & s_Estado_Act & "', '" & s_Codmon_me_Txt & "', '" & s_Codmon_mn_Txt & "') AS monpago, cpc.descpc, "
      s_Sql = s_Sql & "IF(rxd.codmon='" & s_Codmon_mn & "', rxd.imporemune, 0.00) AS importemn, "
      s_Sql = s_Sql & "IF(rxd.codmon='" & s_Codmon_me & "', rxd.imporemune, 0.00) AS importeme "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plconcepto cpc ON rxd.codcpc=cpc.codcpc "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    ElseIf s_OptRegistro = "op5" Then           ' Experiencia laboral
      s_Sql = s_Sql & "exp.orden, exp.empresa, cgo.descgo, exp.fechaini, exp.fechafin, "
      s_Sql = s_Sql & "exp.observacion, psn.estadopsn "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plexpelaboral exp ON psn.codcls=exp.codcls AND psn.codpsn=exp.codpsn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON exp.codcls=cgo.codcls AND exp.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    ElseIf s_OptRegistro = "op6" Then           ' Estudios realizados
      s_Sql = s_Sql & "est.orden, est.institucion, est.grado, est.fechaini, est.fechafin, "
      s_Sql = s_Sql & "est.observacion, psn.estadopsn "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plestudios est ON psn.codcls=est.codcls AND psn.codpsn=est.codpsn "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    ElseIf s_OptRegistro = "op7" Then           ' Datos familiares
      s_Sql = s_Sql & "fam.orden, CONCAT(TRIM(IFNULL(fam.apepaterno, '')), ' ', TRIM(IFNULL(fam.apematerno, '')), ', ', TRIM(IFNULL(fam.nombres, ''))) AS apenombrefam, "
      s_Sql = s_Sql & "fam.fecnacimiento, IF(fam.sexofam='0', 'M', 'F') AS sexofam, fam.numdociden, fam.vinculo, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(via.abrevia, '')), ' ', TRIM(IFNULL(fam.nomviadom, '')), ' ', TRIM(IFNULL(fam.numerdom, '')), ' - ', TRIM(IFNULL(fam.intedom, '')), ' ', TRIM(IFNULL(zon.abrezona, '')), ' ', TRIM(IFNULL(fam.nomzonadom, ''))) AS direccionfam, "
      s_Sql = s_Sql & "TRIM(IFNULL(ubg.desubg, '')) AS distritofam, "
      s_Sql = s_Sql & "fam.incapacidad, fam.certificadomed, fam.cartamed, fam.motivoina, fam.estadofam "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plfamiliares fam ON psn.codcls=fam.codcls AND psn.codpsn=fam.codpsn "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON fam.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON fam.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN " & ps_BDSystems & ".tgubigeo ubg ON fam.ubigeodom=ubg.codubg AND nivelubg='" & s_Estado_Blq & "' "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
    ElseIf s_OptRegistro = "op8" Then           ' Contratos
      s_Sql = s_Sql & "CONCAT(con.numdocumen, '-', con.ano, con.mes, con.dia) AS numcontrato, "
      s_Sql = s_Sql & "con.fechaini, con.fechafin, con.observacion, con.estadocon "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plcontrato con ON psn.codcls=con.codcls AND psn.codpsn=con.codpsn "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND DATE_FORMAT(con.fechafin, '%Y%m%d')<='" & Format(dtpFecha, "yyyymmdd") & "' "
      s_Sql = s_Sql & "AND con.estadocon='" & s_Estado_Act & "' "
    ElseIf s_OptRegistro = "op9" Then           ' Ficha de datos
      ' [ Genero informacion para el listado
      s_Sql = "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      
      s_Sql = "CREATE TABLE IF NOT EXISTS tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn varchar(11) Null, nombrepsn varchar(80) Null, detalle char(1) Null, "
      s_Sql = s_Sql & "fecingreso date Null, codcgo char(3) Null, "
      s_Sql = s_Sql & "descgo varchar(80) Null, codmon char(1) Null, "
      s_Sql = s_Sql & "imporemune decimal(18,2) default '0.00', sigladci char(3) default Null, "
      s_Sql = s_Sql & "numdociden varchar(11) Null, fecnacimiento date Null, "
      s_Sql = s_Sql & "ubigeonac varchar(50) Null, abrezona varchar(4) Null, "
      s_Sql = s_Sql & "nomzondirec varchar(40) Null, abrevia varchar(4) Null, "
      s_Sql = s_Sql & "nomviadirec varchar(40) Null, numerdirec varchar(4) Null, "
      s_Sql = s_Sql & "intedirec varchar(4) Null, ubigeodir varchar(50) Null, "
      s_Sql = s_Sql & "telefono varchar(10) Null, celular varchar(10) Null, "
      s_Sql = s_Sql & "estcivilpsn char(1) Null, numhijo smallint(2) default '0', "
      s_Sql = s_Sql & "desafp varchar(40) Null, numeroafp varchar(15) Null, "
      s_Sql = s_Sql & "nroessalud varchar(15) Null, deseps varchar(40) Null, "
      s_Sql = s_Sql & "empresa varchar(50) Null, exlcargo varchar(40) Null, "
      s_Sql = s_Sql & "expeinicio date Null, expefinal date Null, "
      s_Sql = s_Sql & "expeobserva varchar(100) Null, institucion varchar(50) Null, "
      s_Sql = s_Sql & "estuinicio date Null, estufinal date Null, "
      s_Sql = s_Sql & "grado varchar(100) Null, nombrefam varchar(80) Null, "
      s_Sql = s_Sql & "vinculo char(1) default '0', sigladcifam char(3) Null, "
      s_Sql = s_Sql & "documentofam varchar(11) Null, fecnacifami date Null, "
      s_Sql = s_Sql & "numcontrato varchar(8) Null, coninicio date Null, "
      s_Sql = s_Sql & "confinal date Null, conobserva varchar(50) Null, "
      s_Sql = s_Sql & "estadocon char(1) Null, detcco varchar(50) Null, observacion varchar(100) Null, "
      s_Sql = s_Sql & "annoest varchar(10) Null, mesest varchar(10) Null, establecimiento varchar(100) Null,descpc varchar(100) Null,ximporemune decimal(18,2) default '0.00',xcodmon char(1) Null  )"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' [ Primero : Inserto la informacion experiencia laboral
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, empresa, "
      s_Sql = s_Sql & "exlcargo, expeinicio, expefinal, expeobserva, detcco) "
      s_Sql = s_Sql & "SELECT psn.codpsn, 'a' AS detalle, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.nombres, '')), ' ', TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS nombrepsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.codcgo, cgo.descgo, rxd.codmon, rxd.imporemune, dci.sigladci, "
      s_Sql = s_Sql & "psn.numdociden, psn.fecnacimiento, ubigeonac, zon.abrezona, psn.nomzondirec, "
      s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, ubigeodir, "
      s_Sql = s_Sql & "psn.telefono, psn.celular, psn.estcivilpsn, psn.numhijo, afp.desafp, psn.numeroafp, "
      s_Sql = s_Sql & "psn.nroessalud, eps.deseps, exl.empresa, ecg.descgo AS exlcargo, exl.fechaini AS expeinicio, "
      s_Sql = s_Sql & "exl.fechafin AS expefinal, exl.observacion AS expeobserva, cec.detcco as detcco "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND rxd.codcpc='" & sBasico & "' "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "LEFT JOIN plexpelaboral exl ON psn.codcls=exl.codcls AND psn.codpsn=exl.codpsn "
      s_Sql = s_Sql & "LEFT JOIN plcargo ecg ON exl.codcls=ecg.codcls AND exl.codcgo=ecg.codcgo "
      s_Sql = s_Sql & "LEFT JOIN cocco cec ON psn.codcco=cec.codcco "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      If Trim(txtCenCosto.Text) <> "" Then
        s_Sql = s_Sql & "AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
      End If
      s_Sql = s_Sql & "ORDER BY codpsn"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' [ Segundo : Inserto la informacion formacion academica
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, institucion, "
      s_Sql = s_Sql & "estuinicio, estufinal, grado, observacion ) "
      s_Sql = s_Sql & "SELECT psn.codpsn, 'b' AS detalle, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.nombres, '')), ' ', TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS nombrepsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.codcgo, cgo.descgo, rxd.codmon, rxd.imporemune, dci.sigladci, "
      s_Sql = s_Sql & "psn.numdociden, psn.fecnacimiento, ubigeonac, zon.abrezona, psn.nomzondirec, "
      s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, ubigeodir, "
      s_Sql = s_Sql & "psn.telefono, psn.celular, psn.estcivilpsn, psn.numhijo, afp.desafp, psn.numeroafp, "
      s_Sql = s_Sql & "psn.nroessalud, eps.deseps, est.institucion, est.fechaini AS estuinicio, "
      s_Sql = s_Sql & "est.fechafin AS estufinal, niv.desniv AS grado, est.observacion  "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND rxd.codcpc='" & sBasico & "' "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "INNER JOIN plestudios est ON psn.codcls=est.codcls AND psn.codpsn=est.codpsn "
      s_Sql = s_Sql & "INNER JOIN plniveducativo niv ON est.grado=niv.codniv "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      If Trim(txtCenCosto.Text) <> "" Then
        s_Sql = s_Sql & "AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
      End If
      s_Sql = s_Sql & "ORDER BY codpsn"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' [ Tercero : Inserto la informacion de familiares
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, nombrefam, "
      s_Sql = s_Sql & "vinculo, sigladcifam, documentofam, fecnacifami) "
      s_Sql = s_Sql & "SELECT psn.codpsn, 'c' AS detalle, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.nombres, '')), ' ', TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS nombrepsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.codcgo, cgo.descgo, rxd.codmon, rxd.imporemune, dci.sigladci, "
      s_Sql = s_Sql & "psn.numdociden, psn.fecnacimiento, ubigeonac, zon.abrezona, psn.nomzondirec, "
      s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, ubigeodir, "
      s_Sql = s_Sql & "psn.telefono, psn.celular, psn.estcivilpsn, psn.numhijo, afp.desafp, psn.numeroafp, "
      s_Sql = s_Sql & "psn.nroessalud, eps.deseps, CONCAT(TRIM(IFNULL(fam.nombres, '')), ' ', TRIM(IFNULL(fam.apepaterno, '')), ' ', TRIM(IFNULL(fam.apematerno, ''))) AS nombrefam, "
      s_Sql = s_Sql & "fam.vinculo, dcf.sigladci AS sigladcifam, fam.numdociden AS documentofam, fam.fecnacimiento AS fecnacifami "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND rxd.codcpc='" & sBasico & "' "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "INNER JOIN plfamiliares fam ON psn.codcls=fam.codcls AND psn.codpsn=fam.codpsn "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dcf ON fam.coddci=dcf.coddci "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      If Trim(txtCenCosto.Text) <> "" Then
        s_Sql = s_Sql & "AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
      End If
      s_Sql = s_Sql & "ORDER BY codpsn"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' [ Cuarto: Inserto la informacion de contratos
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, numcontrato, "
      s_Sql = s_Sql & "coninicio, confinal, conobserva, estadocon) "
      s_Sql = s_Sql & "SELECT psn.codpsn, 'd' AS detalle, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.nombres, '')), ' ', TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS nombrepsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.codcgo, cgo.descgo, rxd.codmon, rxd.imporemune, dci.sigladci, "
      s_Sql = s_Sql & "psn.numdociden, psn.fecnacimiento, ubigeonac, zon.abrezona, psn.nomzondirec, "
      s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, ubigeodir, "
      s_Sql = s_Sql & "psn.telefono, psn.celular, psn.estcivilpsn, psn.numhijo, afp.desafp, psn.numeroafp, "
      s_Sql = s_Sql & "psn.nroessalud, eps.deseps, CONCAT(TRIM(IFNULL(con.ano, '')), TRIM(IFNULL(con.mes, '')), TRIM(IFNULL(con.dia, ''))) AS numcontrato, "
      s_Sql = s_Sql & "con.fechaini AS coninicio, con.fechafin AS confinal, tpc.destco AS conobserva, con.estadocon "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "INNER JOIN plcontrato con ON psn.codcls=con.codcls AND psn.codpsn=con.codpsn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND rxd.codcpc='" & sBasico & "' "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "LEFT JOIN pltipcontrato tpc ON con.tipcon=tpc.codtco "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      If Trim(txtCenCosto.Text) <> "" Then
        s_Sql = s_Sql & "AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
      End If
      s_Sql = s_Sql & "ORDER BY codpsn"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      ' [ Quinto: Inserto la informacion de Establecimientos
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, annoest, "
      s_Sql = s_Sql & "mesest, establecimiento) "
      s_Sql = s_Sql & "SELECT psn.codpsn, 'e' AS detalle, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.nombres, '')), ' ', TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS nombrepsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.codcgo, cgo.descgo, rxd.codmon, rxd.imporemune, dci.sigladci, "
      s_Sql = s_Sql & "psn.numdociden, psn.fecnacimiento, ubigeonac, zon.abrezona, psn.nomzondirec, "
      s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, ubigeodir, "
      s_Sql = s_Sql & "psn.telefono, psn.celular, psn.estcivilpsn, psn.numhijo, afp.desafp, psn.numeroafp, "
      s_Sql = s_Sql & "psn.nroessalud, eps.deseps, esta.ano, esta.mes, estp.desepr "
      s_Sql = s_Sql & "FROM plpersonal psn "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND rxd.codcpc='" & sBasico & "' "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "INNER JOIN plestalaboral esta ON psn.codcls=esta.codcls AND psn.codpsn=esta.codpsn "
      s_Sql = s_Sql & "INNER JOIN plestablecimientopropio estp ON esta.codest=estp.cdgepr "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      If Trim(txtCenCosto.Text) <> "" Then
        s_Sql = s_Sql & "AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
      End If
      s_Sql = s_Sql & "ORDER BY codpsn"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      
      ' [ Sexto: Inserto la informacion de Remuneraciones
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, "
      s_Sql = s_Sql & "descpc,ximporemune,xcodmon) "
      s_Sql = s_Sql & "SELECT psn.codpsn, 'f' AS detalle, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.nombres, '')), ' ', TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS nombrepsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.codcgo, cgo.descgo, rxd.codmon, rxd.imporemune, dci.sigladci, "
      s_Sql = s_Sql & "psn.numdociden, psn.fecnacimiento, ubigeonac, zon.abrezona, psn.nomzondirec, "
      s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, ubigeodir, "
      s_Sql = s_Sql & "psn.telefono, psn.celular, psn.estcivilpsn, psn.numhijo, afp.desafp, psn.numeroafp, "
      s_Sql = s_Sql & "psn.nroessalud, eps.deseps,con.descpc,rem.imporemune,rem.codmon  "
      s_Sql = s_Sql & "FROM plremudefa rem "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn on rem.codcls=psn.codcls and rem.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto con on rem.codcpc=con.codcpc and con.tipocpc=0 "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND rxd.codcpc='" & sBasico & "' "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      If Trim(txtCenCosto.Text) <> "" Then
        s_Sql = s_Sql & " AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
      End If
      s_Sql = s_Sql & " ORDER BY codpsn"
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
    
      ' [ Septimo: Inserto la informacion de Comisiones
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, "
      s_Sql = s_Sql & "descpc,ximporemune,xcodmon) "
      s_Sql = s_Sql & "SELECT psn.codpsn, 'f' AS detalle, "
      s_Sql = s_Sql & "CONCAT(TRIM(IFNULL(psn.nombres, '')), ' ', TRIM(IFNULL(psn.apepaterno, '')), ' ', TRIM(IFNULL(psn.apematerno, ''))) AS nombrepsn, "
      s_Sql = s_Sql & "psn.fecingreso, psn.codcgo, cgo.descgo, rxd.codmon, rxd.imporemune, dci.sigladci, "
      s_Sql = s_Sql & "psn.numdociden, psn.fecnacimiento, ubigeonac, zon.abrezona, psn.nomzondirec, "
      s_Sql = s_Sql & "via.abrevia, psn.nomviadirec, psn.numerdirec, psn.intedirec, ubigeodir, "
      s_Sql = s_Sql & "psn.telefono, psn.celular, psn.estcivilpsn, psn.numhijo, afp.desafp, psn.numeroafp, "
      s_Sql = s_Sql & "psn.nroessalud, eps.deseps,con.descpc,case res.codmon when 'N' then res.importe_mn else res.importe_me end,res.codmon  "
      s_Sql = s_Sql & "FROM plresultado res "
      s_Sql = s_Sql & "INNER JOIN plpersonal psn on res.codcls=psn.codcls and res.codpsn=psn.codpsn "
      s_Sql = s_Sql & "INNER JOIN plconcepto con on res.codcpc=con.codcpc and con.tipocpc=0 "
      s_Sql = s_Sql & "LEFT JOIN plcargo cgo ON psn.codcls=cgo.codcls AND psn.codcgo=cgo.codcgo "
      s_Sql = s_Sql & "LEFT JOIN plremudefa rxd ON psn.codcls=rxd.codcls AND psn.codpsn=rxd.codpsn AND rxd.codcpc='" & sBasico & "' "
      s_Sql = s_Sql & "LEFT JOIN pldocidentidad dci ON psn.coddci=dci.coddci "
      s_Sql = s_Sql & "LEFT JOIN pltipozona zon ON psn.codzona=zon.codzona "
      s_Sql = s_Sql & "LEFT JOIN pltipovia via ON psn.codvia=via.codvia "
      s_Sql = s_Sql & "LEFT JOIN plentidadafp afp ON psn.codafp=afp.codafp "
      s_Sql = s_Sql & "LEFT JOIN plentidadeps eps ON psn.codeps=eps.codeps "
      s_Sql = s_Sql & "WHERE psn.codcls='" & ps_ClsPlanilla & "' AND res.codcpc='" & sComision & "' "
      s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
      s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
      s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
      s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
      If Trim(txtCenCosto.Text) <> "" Then
        s_Sql = s_Sql & " AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
      End If
      s_Sql = s_Sql & " ORDER BY res.codpdo desc limit 0,1 "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      
      ' [ Octavo: Inserto la Sumatoria
      s_Sql = "INSERT INTO tmp" & gdl_Procedure.ps_ReportName & " ("
      s_Sql = s_Sql & "codpsn, detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, "
      s_Sql = s_Sql & "descpc,ximporemune,xcodmon) "
      s_Sql = s_Sql & "SELECT codpsn, 'f' AS detalle, nombrepsn, fecingreso, codcgo, descgo, codmon, imporemune, "
      s_Sql = s_Sql & "sigladci, numdociden, fecnacimiento, ubigeonac, abrezona, nomzondirec, "
      s_Sql = s_Sql & "abrevia, nomviadirec, numerdirec, intedirec, ubigeodir, telefono, celular, "
      s_Sql = s_Sql & "estcivilpsn, numhijo, desafp, numeroafp, nroessalud, deseps, "
      s_Sql = s_Sql & "'                 Total' as x,sum(ximporemune),xcodmon "
      s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " where detalle='f'"
      s_Sql = s_Sql & " GROUP BY codpsn "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
           
      ' [ Noveno: Actualizo los Ubigeos Nacimiento y Direccion
      s_Sql = "SELECT distinct ubigeonac "
      s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      While porstRecordset.EOF = False
        s_Sql = " UPDATE tmp" & gdl_Procedure.ps_ReportName & " "
        s_Sql = s_Sql & " SET ubigeonac=( "
        s_Sql = s_Sql & " SELECT CONCAT(dp.desubg, '/', pr.desubg, '/', ds.desubg) AS desubg "
        s_Sql = s_Sql & " FROM ((" & ps_BDSystems & ".tgubigeo ds "
        s_Sql = s_Sql & " LEFT JOIN " & ps_BDSystems & ".tgubigeo pr ON LEFT(ds.codubg, 4)=pr.codubg) "
        s_Sql = s_Sql & " LEFT JOIN " & ps_BDSystems & ".tgubigeo dp ON LEFT(ds.codubg, 2)=dp.codubg) "
        s_Sql = s_Sql & " WHERE ds.nivelubg='2' "
        s_Sql = s_Sql & " AND ds.codubg='" & gdl_Funcion.aTexto(porstRecordset!ubigeonac) & "' "
        s_Sql = s_Sql & " AND pr.nivelubg='1' "
        s_Sql = s_Sql & " AND dp.nivelubg='0' ) "
        s_Sql = s_Sql & " WHERE ubigeonac='" & gdl_Funcion.aTexto(porstRecordset!ubigeonac) & "' "
        gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
        porstRecordset.MoveNext
      Wend
      porstRecordset.Close
      
      s_Sql = "SELECT distinct ubigeodir "
      s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " "
      Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
      While porstRecordset.EOF = False
        s_Sql = " UPDATE tmp" & gdl_Procedure.ps_ReportName & " "
        s_Sql = s_Sql & " SET ubigeodir=( "
        s_Sql = s_Sql & " SELECT CONCAT(dp.desubg, '/', pr.desubg, '/', ds.desubg) AS desubg "
        s_Sql = s_Sql & " FROM ((" & ps_BDSystems & ".tgubigeo ds "
        s_Sql = s_Sql & " LEFT JOIN " & ps_BDSystems & ".tgubigeo pr ON LEFT(ds.codubg, 4)=pr.codubg) "
        s_Sql = s_Sql & " LEFT JOIN " & ps_BDSystems & ".tgubigeo dp ON LEFT(ds.codubg, 2)=dp.codubg) "
        s_Sql = s_Sql & " WHERE ds.nivelubg='2' "
        s_Sql = s_Sql & " AND ds.codubg='" & gdl_Funcion.aTexto(porstRecordset!ubigeodir) & "' "
        s_Sql = s_Sql & " AND pr.nivelubg='1' "
        s_Sql = s_Sql & " AND dp.nivelubg='0' ) "
        s_Sql = s_Sql & " WHERE ubigeodir='" & gdl_Funcion.aTexto(porstRecordset!ubigeodir) & "' "
      gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, s_Sql
      porstRecordset.MoveNext
      Wend
      porstRecordset.Close
      
      ' Obtengo la informacion del reporte
      s_Sql = "SELECT DISTINCTROW lst.codpsn, lst.nombrepsn, lst.detalle, lst.fecingreso, lst.codcgo, lst.descgo, "
      s_Sql = s_Sql & "(CASE lst.codmon WHEN '" & s_Codmon_me & "' THEN '" & s_Codmon_me_Txt & "' ELSE '" & s_Codmon_mn_Txt & "' END) AS codmon_sgn, "
      s_Sql = s_Sql & "lst.imporemune, lst.sigladci, lst.numdociden, lst.fecnacimiento, lst.ubigeonac, lst.abrezona, "
      s_Sql = s_Sql & "lst.nomzondirec, lst.abrevia, lst.nomviadirec, lst.numerdirec, lst.intedirec, lst.ubigeodir, "
      s_Sql = s_Sql & "lst.telefono, lst.celular, "
      s_Sql = s_Sql & "(CASE lst.estcivilpsn WHEN 'S' THEN 'Soltero(a)' WHEN 'C' THEN 'Casado(a)' WHEN 'V' THEN 'Viudo(a)' WHEN 'D' THEN 'Divorciado(a)' WHEN 'O' THEN 'Conviviente' ELSE '' END) AS estcivilpsn, "
      s_Sql = s_Sql & "lst.numhijo, lst.desafp, lst.numeroafp, lst.nroessalud, lst.deseps, lst.empresa, lst.exlcargo, "
      s_Sql = s_Sql & "lst.expeinicio, lst.expefinal, lst.expeobserva, lst.institucion, lst.estuinicio, lst.estufinal, "
      s_Sql = s_Sql & "lst.grado, lst.nombrefam, "
      s_Sql = s_Sql & "(CASE lst.vinculo WHEN '0' THEN 'Otro' WHEN '1' THEN 'Hijo' WHEN '2' THEN 'Conyuge' WHEN '3' THEN 'Concubina(o)' WHEN '4' THEN 'Gestante' ELSE '' END) AS vinculo, "
      s_Sql = s_Sql & "lst.sigladcifam, lst.documentofam, lst.fecnacifami, lst.numcontrato, lst.coninicio, lst.confinal, "
      s_Sql = s_Sql & "lst.conobserva, lst.estadocon, psn.fotopsn, cfg.logo, lst.detcco, lst.observacion , lst.annoest, lst.mesest, lst.establecimiento, lst.descpc,lst.ximporemune,lst.xcodmon "
      s_Sql = s_Sql & "FROM tmp" & gdl_Procedure.ps_ReportName & " lst, plcfgempresa cfg, plpersonal psn "
      s_Sql = s_Sql & "WHERE cfg.pdoano='" & ps_Anyo & "' "
      s_Sql = s_Sql & "AND psn.codcls='" & ps_ClsPlanilla & "' "
      s_Sql = s_Sql & "AND lst.codpsn=psn.codpsn "
    End If
    s_Sql = s_Sql & "AND psn.codpsn IN(SELECT valor FROM rangoimpresion "
    s_Sql = s_Sql & "WHERE proceso='" & s_OptRegistro & "' "
    s_Sql = s_Sql & "AND usrcre='" & ps_Usuario & "' "
    s_Sql = s_Sql & "AND fyhcre='" & s_FechaHora & "') "
    If Trim(txtCenCosto.Text) <> "" Then
      s_Sql = s_Sql & "AND psn.codcco='" & Trim(txtCenCosto.Text) & "' "
    End If
    s_Sql = s_Sql & "ORDER BY codpsn"
      
    Set porstRecordset = OpenRecordset(ps_StrgConnec & ps_DataBase, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
    ' Ejecuto reporte y saco de memoria la información
    gdl_Procedure.ParametersPrinter ps_StrgConnec & ps_DataBase, fMenu.CryReport, (Index - 6), False, True, False, True, True, aElemento, aElementos, porstRecordset
    Set porstRecordset = Nothing
    ' Elimino el rango de impresion
    gdl_Funcion.RangoImpresion ps_StrgConnec & ps_DataBase, s_OptRegistro, "", ps_Usuario, s_FechaHora, "E"
    ' Elimino la tabla temporal
    gdl_Funcion.Execution ps_StrgConnec & ps_DataBase, "DROP TABLE IF EXISTS tmp" & gdl_Procedure.ps_ReportName
    ' Reinicializo los mensajes
    MuestraMensaje s_OldMessage
  End Select

End Sub
Private Sub cmdHelp_Click(Index As Integer)
  
  s_SqlHelp = ""
  Select Case Index
   Case 0      ' Centro de costo
    tdbHelp.Columns(0).DataField = "codcco": tdbHelp.Columns(1).DataField = "detcco"
    tdbHelp.Caption = "Centro de Costos"
    s_Sql = gdl_Funcion.HelpTablas("cco", "codcco", pn_NivelCenCosto, "")
  End Select
  ' Recupera información
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
  tdbHelp.DataSource = porstHelp
  
  ' Muestra la grilla de ayuda
  tdbHelp.Top = panToolBar(1).Top + (Choose(Index + 1, cmdHelp(Index).Top, 750, cmdHelp(Index).Top, 750, 850) + (cmdHelp(Index).Height / 2))
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
  Me.Height = 6740: Me.Width = 7940
  Me.Left = 2900: Me.Top = 100
  ' Recupera parámetro
  gdl_Procedure.pl_RecordSelector = True
  
  ' Caso de instacia del formulario
  s_OptRegistro = fConsultaVarios.Tag

  ' Inicializo los datos de ayuda
  Set porstHelp = New ADODB.Recordset
  n_IndexHelp = -1
  
  ' Configuro datos iniciales
  For n_Index = 1 To 12: cmbPeriodo.AddItem Choose(n_Index, "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", "09 - Setiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"): Next n_Index
  gdl_Procedure.EditText "AT", txtCenCosto, "", s_MdoData_Ins, False, 5
  gdl_Procedure.EditDTPicker "AT", dtpFecha, Date, s_MdoData_Ins, True, s_FormatoFecha, dtpShortDate
  lblDato(1).Caption = IIf(s_OptRegistro = "op3", "Periodo : ", "Fecha : ")
  cmbPeriodo.Visible = (s_OptRegistro = "op3")
  dtpFecha.Visible = (s_OptRegistro = "op8")
  lblDato(1).Visible = (s_OptRegistro = "op3" Or s_OptRegistro = "op8")
  
  ' Titulo del formulario y la Grilla
  n_Index = Right(s_OptRegistro, 1)
  s_TitleWindow = "Listado " & Choose(n_Index, "Padrón de Empleados", "Datos de Trabajos", "Rol de Vacaciones", "Remuneraciones", "Experiencia Laboral", "Estudios Realizados", "Datos Familiares", "Contratos", "Ficha de Datos")
  s_TitleTable = "Trabajador(es)"
  ReDim aElemento(4, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Apellido y Nombres", "Fec.Ingreso", "Ok")
    aElemento(n_Index, 1) = Choose(n_Index + 1, "codpsn", "nombrepsn", "fecingreso", "estadopsn")
    aElemento(n_Index, 2) = Choose(n_Index + 1, 1000, 4162.66, 950, 300)
    aElemento(n_Index, 3) = Choose(n_Index + 1, vbLeftJustify, vbLeftJustify, vbLeftJustify, vbCenter)
    aElemento(n_Index, 4) = Choose(n_Index + 1, "", "", s_FormatoFecha, "")
    aElemento(n_Index, 5) = Choose(n_Index + 1, False, False, False, False)
    aElemento(n_Index, 6) = Choose(n_Index + 1, True, True, True, True)
    aElemento(n_Index, 7) = Choose(n_Index + 1, "", "", "", "")
    aElemento(n_Index, 8) = Choose(n_Index + 1, dbgTop, dbgTop, dbgTop, dbgTop)
    aElemento(n_Index, 9) = Choose(n_Index + 1, 0, 0, 0, 0)
  Next n_Index
  ReDim aElementos(1, 3)
  For n_Index = 0 To (UBound(aElementos, 1) - 1)
    aElementos(n_Index, 0) = ""
    aElementos(n_Index, 1) = 13427690: aElementos(n_Index, 2) = vbBlack
  Next n_Index
  ' Actualizo los campos que se usa en la grilla de TDBGrid
  gdl_Procedure.InicializaGrilla tdbRegistro, aElemento, aElementos
  ' Cambio el formato de la grilla columna de valores
  tdbRegistro.Columns(3).ValueItems.Presentation = dbgNormal
  tdbRegistro.Columns(3).ValueItems.Translate = True
  For n_Index = 0 To 5
    tdbRegistro.Columns(3).ValueItems.Add Item
    tdbRegistro.Columns(3).ValueItems.Item(n_Index).Value = Choose(n_Index + 1, "A", "V", "L", "P", "O", "I")
    tdbRegistro.Columns(3).ValueItems.Item(n_Index).DisplayValue = LoadPicture(gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "estadok", "estadovo", "estadnok", "estadopk", "estadopn", "procenok") & ".bmp")
 Next n_Index
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  
  ' Personaliza el estilo de la grilla de TDBGrid
  gdl_Procedure.DefineStyleGrilla tdbRegistro, s_TitleTable, 1
  ' Agrupacion de columnas y titulo DataView = dbgGroupView
  tdbRegistro.GroupByCaption = "Arrastrar titulo de columna de agrupación"
  tdbRegistro.AllowColMove = False
  
  ' Configuro parametros de visualización del formulario y los controles
  ReDim aElemento(8, 2)
  ' Icono y título del formulario
  aElemento(UBound(aElemento, 1), 1) = "reporte": aElemento(UBound(aElemento, 1), 2) = s_TitleWindow
  ' Cargo los graficos a los controles
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
    aElemento(n_Index, 1) = Choose(n_Index + 1, "ordascen", "orddesce", "busqueda", "selinici", "selfinal", "cancrang", "prelimin", "Imprimir")
    aElemento(n_Index, 2) = Choose(n_Index + 1, "Ordenar Ascendente", "Ordenar Descendente", "Buscar " & s_TitleTable$, "Establece Inicio de Rango", "Establece Fin de Rango", "Inicializa Rango de Impresión", "Presentación Preliminar", "Imprimir")
  Next n_Index
  gdl_Procedure.ViewGrafics Me, cmdAction, aElemento
  
  ' Cargo los graficos de los botones de parametro
  For n_Index = 0 To 2
    ribParametro(n_Index).PictureUp = LoadPicture()
    ribParametro(n_Index).ToolTipText = "Personal " & Choose(n_Index + 1, "Todos", "Activos", "Inactivos")
    s_Sql = gdl_Procedure.ps_PathImagen & Choose(n_Index + 1, "persoall", "filtrook", "filtronok") & ".bmp"
    If gdl_Funcion.ExisteArchivo(s_Sql) Then ribParametro(n_Index).PictureUp = LoadPicture(s_Sql)
  Next n_Index
  ' Presenta Barra de Herramientas
  n_IndexTool = -1: panTool_Click 0
  tdbRegistro.DataSource = dcaRegistro
  ribParametro(0).Value = True
  
 '[ Configuración de la grilla de ayuda
  ReDim aElemento(2, 10)
  For n_Index = 0 To (UBound(aElemento, 1) - 1)
      aElemento(n_Index, 0) = Choose(n_Index + 1, "Código", "Descripción")
      aElemento(n_Index, 1) = Choose(n_Index + 1, "codcts", "descricts")
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
  gdl_Procedure.DefineStyleGrilla tdbHelp, "Conceptos de Cálculo", 2
  ']
  
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
Private Sub tdbHelp_DblClick()

  If porstHelp.RecordCount = 0 Or (porstHelp.EOF And porstHelp.BOF) Then
    Beep
    MsgBox "No existen Registros para Seleccionar", vbExclamation
    Exit Sub
  End If
  Select Case n_IndexHelp
   Case 0       ' Centro de Costo
    txtCenCosto = tdbHelp.Columns(0).Value
    lblHelp(n_IndexHelp) = tdbHelp.Columns(1).Value
    txtCenCosto.SetFocus
  End Select

End Sub
Private Sub tdbHelp_HeadClick(ByVal ColIndex As Integer)
  
  ' Recupero la información ordenada
  Select Case n_IndexHelp
   Case 0     ' Centro de costo
    s_Sql = gdl_Funcion.HelpTablas("cco", tdbHelp.Columns(ColIndex).DataField, pn_NivelCenCosto, "")
  End Select
  Set porstHelp = OpenRecordset(ps_StrgConnec & ps_DaBasCon, adOpenForwardOnly, adLockReadOnly, adUseClient, s_Sql)
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
Private Sub txtCenCosto_GotFocus()
  gdl_Procedure.MarcaGet txtCenCosto
End Sub
Private Sub txtCenCosto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then cmdHelp_Click 0
End Sub
Private Sub txtCenCosto_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub
Private Sub txtCenCosto_LostFocus()
  lblHelp(0) = gdl_Funcion.DameDescripcion(ps_StrgConnec & ps_DaBasCon, ps_CodEmpresa, txtCenCosto, "CC")
  RecuperaRegistros tdbRegistro.Columns(0).DataField & " ASC"
End Sub
